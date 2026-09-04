import { access, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { openDocx } from './docx-document.js';
import { MemoryZip, unzipDocx } from './zip-archive.js';
import { validateDocxPackage } from '../services/standalone-docx-plumbing.js';
import { validateRedlineOoxml } from '../core/redline-validation.js';
import { configureLogger } from '../adapters/logger.js';

const suffixes = { apply: 'redlined', accept: 'accepted', reject: 'rejected', 'delete-comments': 'comments-removed' };

function cliError(code, message, exitCode = 2, details) { return { status: 'error', error: { code, message, ...(details ? { details } : {}) }, exitCode }; }
function parseArgs(argv) {
    const positionals = []; const flags = {};
    for (let index = 0; index < argv.length; index++) {
        const token = argv[index];
        if (!token.startsWith('--')) { positionals.push(token); continue; }
        const [rawKey, inline] = token.slice(2).split(/=(.*)/s); const key = rawKey.replace(/-([a-z])/g, (_, c) => c.toUpperCase());
        if (inline !== undefined) flags[key] = inline;
        else if (argv[index + 1] && !argv[index + 1].startsWith('--')) flags[key] = argv[++index];
        else flags[key] = true;
    }
    return { command: positionals[0], input: positionals[1], flags };
}
function inspectionOptions(flags) {
    const options = {};
    if (flags.search) options.search = flags.search;
    if (flags.revised) options.revisedOnly = true;
    if (flags.table) options.inTable = true;
    if (flags.body) options.inTable = false;
    if (flags.nonEmpty) options.skipEmpty = true;
    if (flags.indexes) options.indexes = String(flags.indexes).split(',').map(Number).filter(Number.isInteger);
    if (flags.range) { const [start, end] = String(flags.range).split(':').map(Number); if (Number.isInteger(start) && Number.isInteger(end) && end >= start) options.indexes = Array.from({ length: end - start + 1 }, (_, i) => start + i); }
    if (flags.view) options.revisionView = flags.view;
    return options;
}
async function readOperations(file) {
    if (!file) throw Object.assign(new Error('Use --operations <file.json>.'), { code: 'OPERATIONS_REQUIRED' });
    let parsed; try { parsed = JSON.parse(await readFile(file, 'utf8')); } catch (error) { throw Object.assign(new Error(`Could not read operations JSON: ${error.message}`), { code: 'INVALID_OPERATIONS_FILE' }); }
    const operations = Array.isArray(parsed) ? parsed : parsed?.operations;
    if (!Array.isArray(operations)) throw Object.assign(new Error('Operations JSON must be an array or an object with an operations array.'), { code: 'INVALID_OPERATIONS_FILE' });
    return operations;
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

export async function executeCli(argv) {
    const { command, input: rawInput, flags } = parseArgs(argv);
    if (command === 'help' || flags.help) return { status: 'ok', command: 'help', usage: 'docx-redline <inspect|extract|preflight|apply|accept|reject|delete-comments|validate> <file.docx> [options]' };
    if (!command) return cliError('COMMAND_REQUIRED', 'A command is required.');
    if (!['inspect','extract','preflight','apply','accept','reject','delete-comments','validate'].includes(command)) return cliError('UNKNOWN_COMMAND', `Unknown command: ${command}`);
    if (!rawInput) return cliError('INPUT_REQUIRED', 'An input .docx path is required.');
    const input = path.resolve(rawInput);
    let buffer; try { buffer = await readFile(input); } catch (error) { return cliError('INPUT_READ_FAILED', error.message); }
    try {
        const document = openDocx(buffer);
        if (command === 'inspect') return { ...document.inspect(inspectionOptions(flags)), command, input };
        if (command === 'extract') {
            const inspected = document.inspect(inspectionOptions(flags));
            return { status: inspected.status, command, input, paragraphs: inspected.paragraphs.map(({ index, ref, paragraphId, fingerprint, exactText, inTable, list, nearestHeading }) => ({ index, ref, paragraphId, fingerprint, exactText, inTable, list, nearestHeading })), warnings: inspected.warnings };
        }
        if (command === 'validate') {
            const entries = unzipDocx(buffer); const documentXml = entries.get('word/document.xml')?.toString('utf8') || '';
            const revision = validateRedlineOoxml(documentXml); const issues = [...revision.issues];
            try { await validateDocxPackage(new MemoryZip(entries)); } catch (error) { issues.push({ code: 'PACKAGE_VALIDATION', severity: 'error', message: error.message }); }
            return { status: issues.some(issue => issue.severity === 'error') ? 'error' : 'ok', command, input, valid: !issues.some(issue => issue.severity === 'error'), issues };
        }
        const operations = command === 'preflight' || command === 'apply' ? await readOperations(flags.operations) : null;
        if (command === 'preflight') return { ...document.preflight(operations, flags.author, { strictTargets: flags.strictTargets !== 'false' }), command, input };
        if (command === 'apply') {
            if (!flags.author && operations.some(operation => !operation?.author)) return cliError('AUTHOR_REQUIRED', 'Use --author or set author on every operation.');
            const result = await document.applyOperations(operations, { author: flags.author, atomic: true, validate: true, strictTargets: true });
            return { command, input, ...serializable(await writeMutation(command, input, flags, result)) };
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
