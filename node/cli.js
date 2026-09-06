import { access, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { openDocx } from './docx-document.js';
import { MemoryZip, unzipDocx } from './zip-archive.js';
import { validateDocxPackage } from '../services/standalone-docx-plumbing.js';
import { validateRedlineOoxml } from '../core/redline-validation.js';
import { configureLogger } from '../adapters/logger.js';

const suffixes = { apply: 'redlined', accept: 'accepted', reject: 'rejected', 'delete-comments': 'comments-removed' };
const commandOptions = {
    inspect: new Set(['help', 'search', 'revised', 'table', 'body', 'nonEmpty', 'index', 'indexes', 'range', 'view']),
    extract: new Set(['help', 'search', 'revised', 'table', 'body', 'nonEmpty', 'index', 'indexes', 'range', 'view']),
    preflight: new Set(['help', 'operations', 'author', 'strictTargets']),
    apply: new Set(['help', 'operations', 'author', 'output', 'inPlace', 'force', 'expectedRevision']),
    accept: new Set(['help', 'author', 'allAuthors', 'output', 'inPlace', 'force']),
    reject: new Set(['help', 'author', 'allAuthors', 'output', 'inPlace', 'force']),
    'delete-comments': new Set(['help', 'author', 'allAuthors', 'output', 'inPlace', 'force']),
    validate: new Set(['help', 'baseline'])
};

function cliError(code, message, exitCode = 2, details) { return { status: 'error', error: { code, message, ...(details ? { details } : {}) }, exitCode }; }
const optionAliases = new Map([
    ['operationsFile', 'operations'],
    ['o', 'output'],
    ['a', 'author'],
    ['i', 'inPlace'],
    ['f', 'force'],
    ['h', 'help']
]);
function parseArgs(argv) {
    const positionals = []; const flags = {};
    for (let index = 0; index < argv.length; index++) {
        const token = argv[index];
        if (!token.startsWith('-') || token === '-') { positionals.push(token); continue; }
        const prefixLength = token.startsWith('--') ? 2 : 1;
        const [rawKey, inline] = token.slice(prefixLength).split(/=(.*)/s);
        const normalizedKey = rawKey.replace(/-([a-z])/g, (_, c) => c.toUpperCase());
        const key = optionAliases.get(normalizedKey) || normalizedKey;
        if (inline !== undefined) flags[key] = inline;
        else if (argv[index + 1] && (!argv[index + 1].startsWith('-') || /^-\d/.test(argv[index + 1]))) flags[key] = argv[++index];
        else flags[key] = true;
    }
    return { command: positionals[0], input: positionals[1], extraPositionals: positionals.slice(2), flags };
}
function positiveInteger(value) {
    const text = String(value).trim();
    const parsed = /^\d+$/.test(text) ? Number(text) : null;
    return Number.isSafeInteger(parsed) && parsed > 0 ? parsed : null;
}
function invalidFilter(message) {
    const error = new Error(message);
    error.code = 'INVALID_FILTER';
    return error;
}
function parseIndexes(value) {
    const tokens = String(value).split(',');
    if (!tokens.length || tokens.some(token => positiveInteger(token) == null)) {
        throw invalidFilter('--indexes must be a comma-separated list of positive 1-based integers.');
    }
    return tokens.map(positiveInteger);
}
function parseRange(value) {
    const match = String(value).match(/^\s*(\d+)\s*([:,\-])\s*(\d+)\s*$/);
    if (!match) throw invalidFilter('--range must use START:END with positive 1-based integers.');
    const start = positiveInteger(match[1]);
    const end = positiveInteger(match[3]);
    if (start == null || end == null || end < start) {
        throw invalidFilter('--range must have positive 1-based endpoints with END greater than or equal to START.');
    }
    return { start, end };
}
function validateCommandOptions(command, flags, extraPositionals) {
    if (extraPositionals.length > 0) return cliError('UNEXPECTED_ARGUMENT', `Unexpected argument: ${extraPositionals[0]}`);
    const allowed = commandOptions[command];
    const unknown = Object.keys(flags).find(option => !allowed.has(option));
    return unknown ? cliError('UNKNOWN_OPTION', `Unknown option for ${command}: --${unknown.replace(/[A-Z]/g, letter => `-${letter.toLowerCase()}`)}`) : null;
}
function inspectionOptions(flags) {
    const options = {};
    if (flags.search) options.search = flags.search;
    if (flags.revised) options.revisedOnly = true;
    if (flags.table) options.inTable = true;
    if (flags.body) options.inTable = false;
    if (flags.nonEmpty) options.skipEmpty = true;
    const selectors = ['index', 'indexes', 'range'].filter(name => flags[name] !== undefined);
    if (selectors.length > 1) throw invalidFilter('Use only one of --index, --indexes, or --range.');
    if (flags.index !== undefined) {
        const index = positiveInteger(flags.index);
        if (index == null) throw invalidFilter('--index must be a positive 1-based integer.');
        options.indexes = [index];
    }
    if (flags.indexes !== undefined) options.indexes = parseIndexes(flags.indexes);
    if (flags.range !== undefined) options.range = parseRange(flags.range);
    if (flags.view) {
        if (!['accepted', 'rejected', 'current'].includes(String(flags.view))) {
            throw invalidFilter('--view must be accepted, rejected, or current.');
        }
        options.revisionView = flags.view;
    }
    return options;
}
async function readOperations(file) {
    if (!file) throw Object.assign(new Error('Use --operations <file.json>.'), { code: 'OPERATIONS_REQUIRED' });
    let parsed; try { parsed = JSON.parse(await readFile(file, 'utf8')); } catch (error) { throw Object.assign(new Error(`Could not read operations JSON: ${error.message}`), { code: 'INVALID_OPERATIONS_FILE' }); }
    const operations = Array.isArray(parsed) ? parsed : (parsed?.operations || parsed?.changes);
    if (!Array.isArray(operations)) throw Object.assign(new Error('Operations JSON must be an array or an object with an operations or changes array.'), { code: 'INVALID_OPERATIONS_FILE' });
    return { operations, expectedRevision: parsed?.expectedRevision || null };
}
function outputPath(command, input, flags) {
    if (flags.inPlace) return input;
    if (flags.output) return path.resolve(String(flags.output));
    const parsed = path.parse(input); return path.join(parsed.dir, `${parsed.name}.${suffixes[command]}${parsed.ext || '.docx'}`);
}
async function writeMutation(command, input, flags, result) {
    if (!result.written) return { status: result.status || 'ok', ...result, outputPath: null };
    const destination = outputPath(command, input, flags);
    if (!flags.inPlace && !flags.force) { try { await access(destination); throw Object.assign(new Error(`Output already exists: ${destination}`), { code: 'OUTPUT_EXISTS' }); } catch (error) { if (error.code !== 'ENOENT') throw error; } }
    await writeFile(destination, result.toBuffer());
    return { status: result.status || 'ok', ...result, outputPath: destination };
}
function serializable(value) {
    const { buffer: _buffer, toBuffer: _toBuffer, ...rest } = value || {};
    return rest;
}

async function collectValidationIssues(buffer) {
    const entries = unzipDocx(buffer);
    const documentXml = entries.get('word/document.xml')?.toString('utf8') || '';
    const revision = validateRedlineOoxml(documentXml);
    const issues = revision.issues.map(issue => ({ source: 'word/document.xml', ...issue }));
    try {
        await validateDocxPackage(new MemoryZip(entries));
    } catch (error) {
        issues.push({ source: 'package', code: 'PACKAGE_VALIDATION', severity: 'error', message: error.message });
    }
    return issues;
}

function validationIssueKey(issue) {
    return `${issue.source || ''}:${issue.code}:${issue.message}`;
}

function subtractValidationIssues(issues, baselineIssues) {
    const remainingBaseline = new Map();
    for (const issue of baselineIssues) {
        const key = validationIssueKey(issue);
        remainingBaseline.set(key, (remainingBaseline.get(key) || 0) + 1);
    }
    return issues.filter(issue => {
        const key = validationIssueKey(issue);
        const remaining = remainingBaseline.get(key) || 0;
        if (remaining === 0) return true;
        remainingBaseline.set(key, remaining - 1);
        return false;
    });
}

export async function executeCli(argv) {
    const { command, input: rawInput, extraPositionals, flags } = parseArgs(argv);
    if (command === 'help' || flags.help) return { status: 'ok', command: 'help', usage: 'docx-redline <inspect|extract|preflight|apply|accept|reject|delete-comments|validate> <file.docx> [options]' };
    if (!command) return cliError('COMMAND_REQUIRED', 'A command is required.');
    if (!['inspect','extract','preflight','apply','accept','reject','delete-comments','validate'].includes(command)) return cliError('UNKNOWN_COMMAND', `Unknown command: ${command}`);
    if (!rawInput) return cliError('INPUT_REQUIRED', 'An input .docx path is required.');
    const optionError = validateCommandOptions(command, flags, extraPositionals);
    if (optionError) return optionError;
    let inspectOptions = null;
    if (command === 'inspect' || command === 'extract') {
        try { inspectOptions = inspectionOptions(flags); }
        catch (error) { return cliError(error.code || 'INVALID_FILTER', error.message); }
    }
    const input = path.resolve(rawInput);
    let buffer; try { buffer = await readFile(input); } catch (error) { return cliError('INPUT_READ_FAILED', error.message); }
    try {
        const document = openDocx(buffer);
        if (command === 'inspect') return { ...document.inspect(inspectOptions), command, input, indexBase: 1 };
        if (command === 'extract') {
            const inspected = document.inspect(inspectOptions);
            return { status: inspected.status, command, input, indexBase: 1, paragraphs: inspected.paragraphs.map(({ index, ref, paragraphId, fingerprint, exactText, inTable, list, nearestHeading }) => ({ index, ref, paragraphId, fingerprint, exactText, inTable, list, nearestHeading })), warnings: inspected.warnings };
        }
        if (command === 'validate') {
            const issues = await collectValidationIssues(buffer);
            if (flags.baseline) {
                const baseline = path.resolve(String(flags.baseline));
                let baselineBuffer;
                try { baselineBuffer = await readFile(baseline); }
                catch (error) { return cliError('BASELINE_READ_FAILED', error.message); }
                const baselineIssues = await collectValidationIssues(baselineBuffer);
                const introducedIssues = subtractValidationIssues(issues, baselineIssues);
                const hasIntroducedErrors = introducedIssues.some(issue => issue.severity === 'error');
                return {
                    status: hasIntroducedErrors ? 'error' : 'ok',
                    command,
                    input,
                    baseline,
                    valid: !hasIntroducedErrors,
                    issues,
                    baselineIssues,
                    introducedIssues
                };
            }
            const hasErrors = issues.some(issue => issue.severity === 'error');
            return { status: hasErrors ? 'error' : 'ok', command, input, valid: !hasErrors, issues };
        }
        const opsData = command === 'preflight' || command === 'apply' ? await readOperations(flags.operations) : null;
        const operations = opsData?.operations || null;
        let expectedRevision = opsData?.expectedRevision || null;
        if (flags.expectedRevision) {
            if (typeof flags.expectedRevision === 'string') {
                try {
                    expectedRevision = JSON.parse(flags.expectedRevision);
                } catch {
                    expectedRevision = {
                        algorithm: 'sha256',
                        version: 1,
                        scope: 'package',
                        value: flags.expectedRevision.trim()
                    };
                }
            } else if (typeof flags.expectedRevision === 'object') {
                expectedRevision = flags.expectedRevision;
            }
        }
        if (command === 'preflight') return { ...document.preflight(operations, flags.author, { strictTargets: flags.strictTargets !== 'false' }), command, input };
        if (command === 'apply') {
            if (!flags.author && operations.some(operation => !operation?.author)) return cliError('AUTHOR_REQUIRED', 'Use --author or set author on every operation.');
            const result = await document.applyOperations(operations, {
                author: flags.author,
                atomic: true,
                validate: true,
                strictTargets: true,
                ...(expectedRevision ? { expectedRevision } : {})
            });
            const mutationResult = await writeMutation(command, input, flags, result);
            return {
                command,
                input,
                ...serializable(mutationResult),
                ...(result.status === 'error' || result.error ? { exitCode: 2 } : {})
            };
        }
        const filter = flags.allAuthors ? { allAuthors: true } : flags.author ? { author: String(flags.author) } : null;
        if (!filter) return cliError('AUTHOR_REQUIRED', 'Use --author <name> or --all-authors.');
        const result = command === 'delete-comments' ? await document.deleteComments(filter) : await document.resolveRevisions(command, filter);
        return { command, input, ...serializable(await writeMutation(command, input, flags, result)) };
    } catch (error) { return cliError(error.code || 'CLI_FAILED', error.message); }
}

export async function runCli(argv = process.argv.slice(2), io = process) {
    configureLogger({}, { level: 'silent' });
    const result = await executeCli(argv); io.stdout.write(`${JSON.stringify(serializable(result), null, 2)}\n`);
    return result.status === 'error' ? (result.exitCode || 1) : 0;
}
