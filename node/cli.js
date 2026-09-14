import { access, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { openDocx } from './docx-document.js';
import { MemoryZip, unzipDocx } from './zip-archive.js';
import { validateDocxPackage } from '../services/standalone-docx-plumbing.js';
import { validateRedlineOoxml } from '../core/redline-validation.js';
import { configureLogger } from '../adapters/logger.js';
import { isExistingRevisionsPolicy } from '../services/document-operation-contract.js';
import { normalizeErrorWithRecovery } from '../services/error-recovery.js';
import { buildCliHelp, CLI_COMMANDS, commandOptionKeys } from './cli-help.js';

const suffixes = { apply: 'redlined', accept: 'accepted', reject: 'rejected', 'delete-comments': 'comments-removed' };
const CLI_CONTRACT_VERSION = 8;
const DEFAULT_INSPECTION_LIMIT = 20;
const INSPECTION_SOFT_BYTE_LIMIT = 48 * 1024;
const CLI_CAPABILITIES = [
    'atomic-batch-results-on-package-failure',
    'baseline-aware-validation',
    'compact-mutation-results',
    'cross-author-revision-slicing',
    'document-scoped-list-revision-ids',
    'batch-start-source-binding',
    'recovery-envelope-v1',
    'require-complete-exit',
    'operations-stdin',
    'agent-safety-profile-v2',
    'command-help-v1',
    'inspection-context-v1',
    'bounded-inspection-v1',
    'human-document-references-v1',
    'localized-replacements-v1',
    'speculative-search-apply-v1',
    'localized-change-summary-v1',
    'deduplicated-cli-receipts',
    'compact-cli-json-v1'
];
const commandOptions = Object.fromEntries(CLI_COMMANDS.map(command => [command, new Set(commandOptionKeys(command))]));

function cliError(code, message, exitCode = 2, details) {
    return {
        status: 'error',
        error: normalizeErrorWithRecovery({
            ...(details && typeof details === 'object' ? details : {}),
            code,
            message
        }),
        exitCode
    };
}
const optionAliases = new Map([
    ['operationsFile', 'operations'],
    ['o', 'output'],
    ['a', 'author'],
    ['i', 'inPlace'],
    ['f', 'force'],
    ['h', 'help'],
    ['no-overwrite', 'noOverwrite'],
    ['no-clobber', 'noClobber'],
    ['no-redlines', 'noRedlines'],
    ['generate-redlines', 'generateRedlines'],
    ['require-complete', 'requireComplete'],
    ['context', 'around'],
    ['C', 'around']
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
        else if (argv[index + 1] !== undefined && (argv[index + 1] === '-' || !argv[index + 1].startsWith('-') || /^-\d/.test(argv[index + 1]))) flags[key] = argv[++index];
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
function boundedPositiveInteger(value, optionName, maximum = Number.MAX_SAFE_INTEGER) {
    const parsed = positiveInteger(value);
    if (parsed == null || parsed > maximum) {
        const upperBound = maximum < Number.MAX_SAFE_INTEGER ? ` no greater than ${maximum}` : '';
        throw invalidFilter(`${optionName} must be a positive integer${upperBound}.`);
    }
    return parsed;
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
    if (flags.around !== undefined) {
        const text = String(flags.around).trim();
        if (!/^\d+$/.test(text) || Number(text) > 20) {
            throw invalidFilter('--around must be an integer from 0 through 20.');
        }
        options.around = Number(text);
        if (!flags.search) throw invalidFilter('--around requires --search.');
    }
    if (flags.limit !== undefined) options.limit = boundedPositiveInteger(flags.limit, '--limit', 200);
    if (flags.after !== undefined) options.after = boundedPositiveInteger(flags.after, '--after');
    if (flags.all && flags.limit !== undefined) throw invalidFilter('Use --all or --limit, not both.');
    if (!flags.all && flags.limit === undefined && selectors.length === 0) options.limit = DEFAULT_INSPECTION_LIMIT;
    if (flags.view) {
        if (!['accepted', 'rejected', 'current'].includes(String(flags.view))) {
            throw invalidFilter('--view must be accepted, rejected, or current.');
        }
        options.revisionView = flags.view;
    }
    return options;
}
async function readUtf8Stream(stream) {
    if (!stream || typeof stream[Symbol.asyncIterator] !== 'function') {
        throw new Error('No readable stdin stream was provided.');
    }
    const chunks = [];
    for await (const chunk of stream) chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(String(chunk)));
    return Buffer.concat(chunks).toString('utf8');
}

function parseContextRange(value) {
    const match = String(value).match(/^\s*(-?\d+)\s*:\s*(-?\d+)\s*$/);
    if (!match) throw invalidFilter('--context-range must use START:END with signed integer offsets from -20 through 20.');
    const start = Number(match[1]);
    const end = Number(match[2]);
    if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < -20 || end > 20 || start > end) {
        throw invalidFilter('--context-range requires START <= END and both offsets from -20 through 20.');
    }
    return { start, end, text: `${start}:${end}` };
}

function speculativeCandidate(paragraph) {
    return {
        excerpt: boundedText(paragraph.exactText, 240),
        index: paragraph.index,
        ref: paragraph.ref,
        paragraphId: paragraph.paragraphId,
        fingerprint: paragraph.fingerprint,
        revisionView: 'accepted',
        inTable: paragraph.inTable,
        humanReference: paragraph.humanReference,
        provision: paragraph.provision,
        nearestHeading: paragraph.nearestHeading
    };
}

function speculativeError(code, message, details = {}) {
    return Object.assign(new Error(message), { code, ...details });
}

function resolveSpeculativePatchTarget(document, flags) {
    if (!document) throw speculativeError('INVALID_OPERATION', 'Speculative localized replacement requires an open document.');
    if (flags.contextRange !== undefined && flags.around !== undefined) {
        throw invalidFilter('Use --context-range or --around, not both.');
    }
    const search = flags.search === undefined ? null : String(flags.search).trim();
    if (flags.search !== undefined && (!search || flags.search === true)) {
        throw invalidFilter('--search must be non-empty text.');
    }
    if ((flags.contextRange !== undefined || flags.around !== undefined) && !search) {
        throw invalidFilter('--context-range and --around require --search.');
    }
    let range = { start: 0, end: 0, text: '0:0' };
    if (flags.contextRange !== undefined) range = parseContextRange(flags.contextRange);
    if (flags.around !== undefined) {
        const text = String(flags.around).trim();
        if (!/^\d+$/.test(text) || Number(text) > 20) {
            throw invalidFilter('--around must be an integer from 0 through 20.');
        }
        const around = Number(text);
        range = { start: -around, end: around, text: `${-around}:${around}` };
    }

    const inspected = document.inspect({ revisionView: 'accepted' });
    const paragraphs = Array.isArray(inspected.paragraphs) ? inspected.paragraphs : [];
    let anchors = [];
    let eligible = paragraphs;
    if (search) {
        const folded = search.toLowerCase();
        anchors = paragraphs.filter(paragraph => paragraph.exactText.toLowerCase().includes(folded));
        if (anchors.length === 0) {
            throw speculativeError('TARGET_NOT_FOUND', `No paragraph matched search query: "${search}".`, {
                context: { search, range: range.text },
                recovery: { action: 'reinspect', requiresReinspection: true, sameArgumentsSafe: false }
            });
        }
        const eligibleIndexes = new Set();
        for (const anchor of anchors) {
            const anchorPosition = anchor.index - 1;
            for (let offset = range.start; offset <= range.end; offset += 1) {
                const position = anchorPosition + offset;
                if (position >= 0 && position < paragraphs.length) eligibleIndexes.add(position);
            }
        }
        eligible = [...eligibleIndexes].sort((a, b) => a - b).map(position => paragraphs[position]);
    }

    const find = String(flags.find);
    const matches = eligible.filter(paragraph => paragraph.exactText.includes(find));
    const context = {
        ...(search ? { search, range: range.text, anchorMatchCount: anchors.length } : {})
    };
    if (matches.length === 0) {
        throw speculativeError('PATCH_SOURCE_NOT_FOUND', search
            ? `No paragraph in the contextual range contains the exact patch source: "${find}".`
            : `No paragraph contains the exact patch source: "${find}".`, {
            context,
            ...(anchors.length ? { candidates: anchors.map(speculativeCandidate) } : {}),
            recovery: { action: 'reinspect', requiresReinspection: true, sameArgumentsSafe: false }
        });
    }
    if (matches.length > 1) {
        throw speculativeError('AMBIGUOUS_TARGET', `The exact patch source matched ${matches.length} eligible paragraphs.`, {
            context,
            candidates: matches.map(speculativeCandidate),
            recovery: { action: 'choose_candidate', requiresReinspection: false, sameArgumentsSafe: false }
        });
    }
    const selected = matches[0];
    return {
        target: {
            exactText: selected.exactText,
            ...(selected.paragraphId ? { paragraphId: selected.paragraphId } : {}),
            fingerprint: selected.fingerprint,
            revisionView: 'accepted',
            inTable: selected.inTable
        },
        context: {
            ...context,
            humanReference: selected.humanReference
        }
    };
}

async function readOperations(file, flags = {}, stdin = process.stdin, document = null) {
    const hasInlineTarget = flags?.target !== undefined
        || flags?.targetId !== undefined
        || flags?.targetRef !== undefined;
    const hasInlinePatch = flags?.find !== undefined
        || flags?.replace !== undefined
        || flags?.occurrence !== undefined;
    const hasSpeculativeScope = flags?.search !== undefined
        || flags?.contextRange !== undefined
        || flags?.around !== undefined;
    if (file && (hasInlinePatch || hasSpeculativeScope)) {
        throw Object.assign(new Error('Use --operations or inline speculative find/replace options, not both.'), {
            code: 'INVALID_OPERATION'
        });
    }
    if (hasInlineTarget && hasSpeculativeScope) {
        throw Object.assign(new Error('Use a strong inline target or --search/--context-range, not both.'), {
            code: 'INVALID_OPERATION'
        });
    }
    if (!file && (hasInlineTarget || hasInlinePatch || hasSpeculativeScope)) {
        if (hasSpeculativeScope && !hasInlinePatch) {
            throw Object.assign(new Error('--search, --context-range, and --around are only valid with inline --find/--replace.'), {
                code: 'INVALID_OPERATION'
            });
        }
        const targetRef = flags.targetRef === undefined ? null : positiveInteger(flags.targetRef);
        if (flags.targetRef !== undefined && targetRef == null) {
            throw Object.assign(new Error('--target-ref must be a positive 1-based integer.'), {
                code: 'INVALID_OPERATION'
            });
        }
        const targetId = flags.targetId === undefined ? null : String(flags.targetId).trim();
        if (flags.targetId !== undefined && (!targetId || flags.targetId === true)) {
            throw Object.assign(new Error('--target-id must be a non-empty paragraphId.'), {
                code: 'INVALID_OPERATION'
            });
        }
        const targetText = flags.target === undefined ? null : String(flags.target);
        let target = targetId
            ? { ...(targetText == null ? {} : { exactText: targetText }), paragraphId: targetId }
            : targetText;
        let speculativeContext = null;
        let op;
        if (flags.comment) {
            if (hasInlinePatch) {
                throw Object.assign(new Error('Inline comments cannot be combined with --find/--replace.'), {
                    code: 'INVALID_OPERATION'
                });
            }
            op = {
                type: 'comment',
                ...(target == null ? {} : { target }),
                commentContent: String(flags.comment),
                ...(flags.textToComment ? { textToComment: String(flags.textToComment) } : {}),
                ...(targetRef == null ? {} : { targetRef }),
                ...(flags.author ? { author: String(flags.author) } : {})
            };
        } else if (hasInlinePatch) {
            if (flags.modified !== undefined) {
                throw Object.assign(new Error('Use --modified or --find/--replace, not both.'), {
                    code: 'INVALID_OPERATION'
                });
            }
            if (
                flags.find === undefined
                || flags.find === true
                || String(flags.find).length === 0
                || flags.replace === undefined
                || flags.replace === true
            ) {
                throw Object.assign(new Error('Inline localized replacement requires non-empty --find and string --replace values.'), {
                    code: 'INVALID_OPERATION'
                });
            }
            const occurrence = flags.occurrence === undefined ? null : positiveInteger(flags.occurrence);
            if (flags.occurrence !== undefined && occurrence == null) {
                throw Object.assign(new Error('--occurrence must be a positive 1-based integer.'), {
                    code: 'INVALID_OPERATION'
                });
            }
            if (!hasInlineTarget) {
                const resolution = resolveSpeculativePatchTarget(document, flags);
                target = resolution.target;
                speculativeContext = resolution.context;
            }
            op = {
                type: 'redline',
                ...(target == null ? {} : { target }),
                ...(targetRef == null ? {} : { targetRef }),
                replacements: [{
                    find: String(flags.find),
                    replace: String(flags.replace),
                    ...(occurrence == null ? {} : { occurrence })
                }],
                ...(speculativeContext ? { _speculativeContext: speculativeContext } : {}),
                ...(flags.author ? { author: String(flags.author) } : {}),
                ...(flags.existingRevisions ? { existingRevisions: String(flags.existingRevisions) } : {})
            };
        } else {
            op = {
                type: 'replace',
                ...(target == null ? {} : { target }),
                modified: flags.modified !== undefined ? String(flags.modified) : '',
                ...(targetRef == null ? {} : { targetRef }),
                ...(flags.author ? { author: String(flags.author) } : {}),
                ...(flags.existingRevisions ? { existingRevisions: String(flags.existingRevisions) } : {})
            };
        }
        return { operations: [op], expectedRevision: null };
    }
    if (!file) throw Object.assign(new Error('Use --operations <file.json> or an inline operation.'), { code: 'OPERATIONS_REQUIRED' });
    let parsed;
    try {
        const source = file === '-' ? await readUtf8Stream(stdin) : await readFile(file, 'utf8');
        parsed = JSON.parse(source);
    } catch (error) {
        const location = file === '-' ? ' from stdin' : '';
        throw Object.assign(new Error(`Could not read operations JSON${location}: ${error.message}`), { code: 'INVALID_OPERATIONS_FILE' });
    }
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
    const protectExisting = Boolean(flags.noOverwrite || flags.noClobber) && !flags.force;
    if (!flags.inPlace && protectExisting) {
        try {
            await access(destination);
            throw Object.assign(new Error(`Output already exists: ${destination}`), { code: 'OUTPUT_EXISTS' });
        } catch (error) {
            if (error.code !== 'ENOENT') throw error;
        }
    }
    await writeFile(destination, result.toBuffer());
    return { status: result.status || 'ok', ...result, outputPath: destination };
}
function serializable(value) {
    const { buffer: _buffer, toBuffer: _toBuffer, ...rest } = value || {};
    return rest;
}

function compactExtractParagraph(paragraph) {
    const {
        humanReference, provision, nearestHeading, index, ref, paragraphId,
        fingerprint, revisionView, exactText, inTable, list, selectionRole,
        contextFor
    } = paragraph;
    return {
        humanReference,
        provision,
        nearestHeading,
        index,
        ref,
        paragraphId,
        fingerprint,
        revisionView,
        exactText,
        inTable,
        list,
        ...(selectionRole ? { selectionRole } : {}),
        ...(Array.isArray(contextFor) ? { contextFor } : {})
    };
}

function inspectionResponseBytes(value) {
    return Buffer.byteLength(JSON.stringify(value, null, 2), 'utf8') + 1;
}

function boundInspectionResponse(value, { bypass = false, broadDetailed = false } = {}) {
    const base = {
        ...value,
        machineReferencesAreNotUserLocations: true,
        ...(broadDetailed ? {
            notes: [
                'This detailed inspection was bounded. Prefer extract --search or an explicit --range for ordinary targeting.'
            ]
        } : {})
    };
    if (bypass) return base;

    let paragraphs = [...(base.paragraphs || [])];
    const sourceSelection = base.selection || {};
    const totalMatches = Number.isInteger(sourceSelection.totalMatches)
        ? sourceSelection.totalMatches
        : paragraphs.filter(item => item.selectionRole !== 'context').length;
    const initialReturnedMatches = Number.isInteger(sourceSelection.returnedMatches)
        ? sourceSelection.returnedMatches
        : paragraphs.filter(item => item.selectionRole !== 'context').length;
    let contextTruncated = false;

    const assemble = (oversizeItem = false) => {
        const directIndexes = new Set(paragraphs
            .filter(item => item.selectionRole !== 'context')
            .map(item => item.index));
        paragraphs = paragraphs
            .map(item => item.selectionRole === 'context'
                ? { ...item, contextFor: (item.contextFor || []).filter(index => directIndexes.has(index)) }
                : item)
            .filter(item => item.selectionRole !== 'context' || item.contextFor.length > 0);
        const direct = paragraphs.filter(item => item.selectionRole !== 'context');
        const lastDirect = direct[direct.length - 1] || null;
        const paragraphIndexes = new Set(paragraphs.map(item => item.index));
        const comments = Array.isArray(base.comments)
            ? base.comments.filter(comment => paragraphIndexes.has(comment.paragraphIndex))
            : base.comments;
        const budgetTruncated = direct.length < initialReturnedMatches;
        return {
            ...base,
            paragraphs,
            ...(Array.isArray(base.comments) ? { comments } : {}),
            selection: {
                ...sourceSelection,
                totalMatches,
                returnedMatches: direct.length,
                returnedParagraphs: paragraphs.length,
                truncated: sourceSelection.truncated === true || budgetTruncated || contextTruncated,
                nextAfter: sourceSelection.truncated === true || budgetTruncated
                    ? (lastDirect?.index ?? sourceSelection.nextAfter ?? null)
                    : null,
                softByteLimit: INSPECTION_SOFT_BYTE_LIMIT,
                oversizeItem,
                ...(contextTruncated ? { contextTruncated: true } : {})
            }
        };
    };

    let result = assemble();
    while (inspectionResponseBytes(result) > INSPECTION_SOFT_BYTE_LIMIT) {
        const directPositions = paragraphs
            .map((item, position) => item.selectionRole !== 'context' ? position : -1)
            .filter(position => position >= 0);
        if (directPositions.length > 1) {
            paragraphs.splice(directPositions[directPositions.length - 1], 1);
            result = assemble();
            continue;
        }
        const contextPosition = paragraphs.findLastIndex(item => item.selectionRole === 'context');
        if (contextPosition >= 0) {
            paragraphs.splice(contextPosition, 1);
            contextTruncated = true;
            result = assemble();
            continue;
        }
        result = assemble(true);
        break;
    }
    return result;
}

function boundedText(value, limit = 512) {
    const text = String(value ?? '');
    return text.length > limit ? `${text.slice(0, limit)}…` : text;
}

function compactError(error) {
    if (!error || typeof error !== 'object') return error;
    const fields = [
        'code', 'stage', 'mismatchOffset', 'expectedExcerpt', 'actualExcerpt',
        'expectedCodePoint', 'actualCodePoint', 'ownerAuthor', 'commentIds',
        'recoveryVersion', 'category', 'field', 'captureRef', 'operationIndexes',
        'consumedByOperation', 'expectedScope', 'actualScope'
    ];
    const compact = {};
    for (const field of fields) {
        if (error[field] !== undefined) compact[field] = error[field];
    }
    if (Array.isArray(error.comments)) {
        compact.comments = error.comments.map(comment => ({
            ...(comment?.id !== undefined ? { id: comment.id } : {}),
            ...(comment?.author !== undefined ? { author: boundedText(comment.author, 160) } : {}),
            ...(comment?.text !== undefined ? { text: boundedText(comment.text, 512) } : {})
        }));
    }
    if (Array.isArray(error.candidates)) {
        compact.candidates = error.candidates.map(candidate => {
            if (!candidate || typeof candidate !== 'object') return candidate;
            const excerpt = boundedText(candidate.excerpt ?? candidate.exactText ?? candidate.text ?? '', 240);
            return { ...compactResolvedTarget(candidate, { preserveMatchDetails: true }), ...(excerpt ? { excerpt } : {}) };
        });
    }
    for (const field of ['recovery', 'issueSummary', 'expectedRevision', 'currentRevision']) {
        if (error[field] !== undefined) compact[field] = error[field];
    }
    if (error.context && typeof error.context === 'object') {
        compact.context = {
            ...error.context,
            ...(error.context.currentTarget ? {
                currentTarget: {
                    ...compactResolvedTarget(error.context.currentTarget, { preserveMatchDetails: true }),
                    excerpt: boundedText(
                        error.context.currentTarget.excerpt
                            ?? error.context.currentTarget.exactText
                            ?? error.context.currentTarget.text
                            ?? '',
                        240
                    )
                }
            } : {})
        };
    }
    if (error.sourceTarget && typeof error.sourceTarget === 'object') {
        compact.sourceTarget = {
            ...compactResolvedTarget(error.sourceTarget, { preserveMatchDetails: true }),
            excerpt: boundedText(error.sourceTarget.text ?? error.sourceTarget.exactText ?? '', 240)
        };
    }
    compact.message = boundedText(error.message || String(error));
    return compact;
}

function compactTargetTextMatch(match, preserveDetails = false) {
    if (!match || typeof match !== 'object') return match;
    if (preserveDetails) return match;
    if (match.mode === 'exact') return undefined;
    const differenceCount = Number.isInteger(match.differenceCount)
        ? match.differenceCount
        : (Array.isArray(match.differences) ? match.differences.length : 0);
    return {
        ...(match.mode ? { mode: match.mode } : {}),
        differenceCount
    };
}

function compactResolvedTarget(target, { preserveMatchDetails = false } = {}) {
    if (!target || typeof target !== 'object') return target;
    const {
        text: _text,
        exactText: _exactText,
        targetTextMatch,
        ...compact
    } = target;
    const compactMatch = compactTargetTextMatch(targetTextMatch, preserveMatchDetails);
    return {
        ...compact,
        ...(compactMatch ? { targetTextMatch: compactMatch } : {})
    };
}

function compactReceipt(receipt) {
    if (!receipt || typeof receipt !== 'object') return receipt;
    return {
        ...receipt,
        affectedTargets: Array.isArray(receipt.affectedTargets)
            ? receipt.affectedTargets.map(compactResolvedTarget)
            : [],
        warnings: Array.isArray(receipt.warnings) ? receipt.warnings.map(warning => boundedText(warning)) : []
    };
}

function compactOperationResult(result) {
    if (!result || typeof result !== 'object') return result;
    const { receipt: _receipt, ...withoutReceipt } = result;
    const preserveMatchDetails = result.status === 'error' || !!result.error;
    return {
        ...withoutReceipt,
        ...(result.resolvedTarget ? { resolvedTarget: compactResolvedTarget(result.resolvedTarget, { preserveMatchDetails }) } : {}),
        ...(result.resolvedAnchor ? { resolvedAnchor: compactResolvedTarget(result.resolvedAnchor, { preserveMatchDetails }) } : {}),
        ...(result.error ? { error: compactError(result.error) } : {}),
        ...(Array.isArray(result.warnings) ? { warnings: result.warnings.map(warning => boundedText(warning)) } : {})
    };
}

function summarizeIssues(issues) {
    const list = Array.isArray(issues) ? issues : [];
    const grouped = new Map();
    for (const issue of list) {
        const code = issue?.code || 'UNKNOWN';
        const source = issue?.source || 'unknown';
        const severity = issue?.severity || 'error';
        const key = `${source}:${severity}:${code}`;
        const current = grouped.get(key) || { source, severity, code, count: 0 };
        current.count++;
        grouped.set(key, current);
    }
    return {
        total: list.length,
        errors: list.filter(issue => issue?.severity === 'error').length,
        warnings: list.filter(issue => issue?.severity === 'warning').length,
        byCode: Array.from(grouped.values()).sort((a, b) => (
            a.source.localeCompare(b.source) || a.code.localeCompare(b.code) || a.severity.localeCompare(b.severity)
        ))
    };
}

function compactMutationResult(value) {
    const serialized = serializable(value);
    const {
        documentXml: _documentXml,
        oxml: _oxml,
        commentsXml: _commentsXml,
        commentsExtendedXml: _commentsExtendedXml,
        numberingXml: _numberingXml,
        numberingXmlParts: _numberingXmlParts,
        inspection: _inspection,
        issues: _issues,
        ...compact
    } = serialized;
    const results = Array.isArray(compact.results) ? compact.results.map(compactOperationResult) : [];
    const status = compact.status || 'ok';
    const localizedChangesVerified = results.every(result => !result?.change || (
        result.change.committed === true
        && result.change.finalDisposition === 'applied'
        && result.change.verification?.acceptedViewMatchesCompiledText === true
    ));
    return {
        ...compact,
        ...(Array.isArray(compact.results) ? { results } : {}),
        ...(Array.isArray(compact.receipts) ? { receipts: compact.receipts.map(compactReceipt) } : {}),
        ...(compact.error ? { error: compactError(compact.error) } : {}),
        ...(Array.isArray(compact.warnings) ? { warnings: compact.warnings.map(warning => boundedText(warning)) } : {}),
        ...(compact.validation ? {
            validation: {
                originalIssues: summarizeIssues(compact.validation.originalIssues),
                generatedIssues: summarizeIssues(compact.validation.generatedIssues)
            }
        } : {}),
        completion: compact.written === true
            && status !== 'error'
            && status !== 'partial'
            && results.every(result => result?.status !== 'error')
            && localizedChangesVerified
    };
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

export async function executeCli(argv, io = process) {
    const { command, input: rawInput, extraPositionals, flags } = parseArgs(argv);
    if (command === 'help' || flags.help) {
        const requestedCommand = command === 'help' ? rawInput : command;
        const help = buildCliHelp(requestedCommand || null);
        return help || cliError('UNKNOWN_COMMAND', `Unknown command: ${requestedCommand}`);
    }
    if (!command) return cliError('COMMAND_REQUIRED', 'A command is required.');
    if (!CLI_COMMANDS.includes(command)) return cliError('UNKNOWN_COMMAND', `Unknown command: ${command}`);
    if (command === 'version') {
        if (rawInput || extraPositionals.length > 0) return cliError('UNEXPECTED_ARGUMENT', `Unexpected argument: ${rawInput || extraPositionals[0]}`);
        const optionError = validateCommandOptions(command, flags, []);
        if (optionError) return optionError;
        return {
            status: 'ok',
            command,
            contractVersion: CLI_CONTRACT_VERSION,
            capabilities: CLI_CAPABILITIES
        };
    }
    if (!rawInput) return cliError('INPUT_REQUIRED', 'An input .docx path is required.');
    const optionError = validateCommandOptions(command, flags, extraPositionals);
    if (optionError) return optionError;
    if (flags.existingRevisions != null && !isExistingRevisionsPolicy(flags.existingRevisions)) {
        return cliError('INVALID_OPERATION', `Unsupported existing-revisions policy: "${String(flags.existingRevisions)}".`);
    }
    const profile = flags.profile == null ? null : String(flags.profile);
    if (profile && profile !== 'agent') {
        return cliError('INVALID_PROFILE', `Unknown execution profile: "${profile}". Supported profiles: agent.`);
    }
    let inspectOptions = null;
    if (command === 'inspect' || command === 'extract') {
        try { inspectOptions = inspectionOptions(flags); }
        catch (error) { return cliError(error.code || 'INVALID_FILTER', error.message); }
    }
    const input = path.resolve(rawInput);
    let buffer; try { buffer = await readFile(input); } catch (error) { return cliError('INPUT_READ_FAILED', error.message); }
    try {
        const document = openDocx(buffer);
        if (command === 'inspect') {
            const inspected = document.inspect(inspectOptions);
            const broadDetailed = !flags.search && flags.index === undefined
                && flags.indexes === undefined && flags.range === undefined;
            return boundInspectionResponse(
                { ...inspected, command, input, indexBase: 1 },
                { bypass: !!flags.all, broadDetailed }
            );
        }
        if (command === 'extract') {
            const inspected = document.inspect(inspectOptions);
            return boundInspectionResponse({
                status: inspected.status,
                command,
                input,
                indexBase: 1,
                paragraphs: inspected.paragraphs.map(compactExtractParagraph),
                ...(inspected.selection ? { selection: inspected.selection } : {}),
                warnings: inspected.warnings
            }, { bypass: !!flags.all });
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
        const opsData = command === 'preflight' || command === 'apply'
            ? await readOperations(flags.operations, flags, io.stdin || process.stdin, document)
            : null;
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
        if (command === 'preflight') return {
            ...document.preflight(operations, flags.author, {
                strictTargets: flags.strictTargets !== 'false',
                ...(flags.existingRevisions ? { existingRevisions: flags.existingRevisions } : {})
            }),
            command,
            input
        };
        if (command === 'apply') {
            const agentProfile = profile === 'agent';
            const author = flags.author || process.env.DOCX_REDLINE_AUTHOR || 'AI Redliner';
            const generateRedlines = flags.generateRedlines !== undefined
                ? (flags.generateRedlines !== 'false' && flags.generateRedlines !== false)
                : (!flags.noRedlines);
            const atomic = flags.atomic !== undefined
                ? (flags.atomic === true || flags.atomic === 'true')
                : false;
            const requireComplete = flags.requireComplete !== undefined
                ? (flags.requireComplete === true || flags.requireComplete === 'true')
                : agentProfile;
            const effectiveOptions = {
                author,
                atomic,
                strictTargets: true,
                validate: true,
                generateRedlines,
                existingRevisions: flags.existingRevisions || 'merge-same-author',
                requireComplete
            };
            const result = await document.applyOperations(operations, {
                author,
                atomic,
                validate: true,
                strictTargets: true,
                generateRedlines,
                ...(flags.existingRevisions ? { existingRevisions: flags.existingRevisions } : {}),
                ...(expectedRevision ? { expectedRevision } : {})
            });
            const localizedVerificationFailed = (result.results || []).some(item => item?.change && (
                item.change.committed !== true
                || item.change.finalDisposition !== 'applied'
                || item.change.verification?.acceptedViewMatchesCompiledText !== true
            ));
            const mutationResult = await writeMutation(command, input, flags, result);
            return compactMutationResult({
                command,
                input,
                ...serializable(mutationResult),
                ...(profile ? { executionProfile: profile, effectiveOptions } : {}),
                ...(result.status === 'error'
                    ? { exitCode: 2 }
                    : ((result.status === 'partial' || localizedVerificationFailed) && requireComplete
                        ? { exitCode: 3 }
                        : {}))
            });
        }
        const filter = flags.allAuthors ? { allAuthors: true } : flags.author ? { author: String(flags.author) } : null;
        if (!filter) return cliError('AUTHOR_REQUIRED', 'Use --author <name> or --all-authors.');
        const result = command === 'delete-comments' ? await document.deleteComments(filter) : await document.resolveRevisions(command, filter);
        return compactMutationResult({ command, input, ...serializable(await writeMutation(command, input, flags, result)) });
    } catch (error) { return cliError(error.code || 'CLI_FAILED', error.message, 2, error); }
}

export async function runCli(argv = process.argv.slice(2), io = process) {
    configureLogger({}, { level: 'silent' });
    const result = await executeCli(argv, io);
    const compactJson = parseArgs(argv).flags.compact === true;
    io.stdout.write(`${JSON.stringify(serializable(result), null, compactJson ? 0 : 2)}\n`);
    return Number.isInteger(result.exitCode) && result.exitCode !== 0
        ? result.exitCode
        : (result.status === 'error' ? 1 : 0);
}
