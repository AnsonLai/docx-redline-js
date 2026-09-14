const option = (key, name, description) => Object.freeze({ key, name, description });

const HELP = option('help', '--help, -h', 'Return machine-readable help without opening a document.');
const AUTHOR = option('author', '--author <name>, -a <name>', 'Reviewer name; falls back to DOCX_REDLINE_AUTHOR, then AI Redliner.');
const OUTPUT = option('output', '--output <file>, -o <file>', 'Write to this destination; the source is not overwritten.');
const IN_PLACE = option('inPlace', '--in-place, -i', 'Explicitly overwrite the source document.');
const FORCE = option('force', '--force, -f', 'Allow replacement of an existing destination.');
const NO_OVERWRITE = option('noOverwrite', '--no-overwrite, --no-clobber', 'Refuse replacement of an existing destination.');
const ALL_AUTHORS = option('allAuthors', '--all-authors', 'Resolve review content for every author; requires explicit user authorization.');
const COMPACT = option('compact', '--compact', 'Emit one-line JSON to reduce provider-visible output bytes.');

const INSPECTION_OPTIONS = Object.freeze([
    HELP,
    option('search', '--search <text>', 'Case-insensitive substring search.'),
    option('revised', '--revised', 'Select paragraphs containing tracked revisions.'),
    option('table', '--table', 'Select paragraphs inside tables.'),
    option('body', '--body', 'Select paragraphs outside tables.'),
    option('nonEmpty', '--non-empty', 'Exclude empty paragraphs; this is not a narrow document scope.'),
    option('index', '--index <N>', 'Select one 1-based machine paragraph index.'),
    option('indexes', '--indexes <N,N,...>', 'Select comma-separated 1-based machine paragraph indexes.'),
    option('range', '--range <START:END>', 'Select an inclusive range of 1-based machine paragraph indexes.'),
    option('view', '--view <accepted|rejected|current>', 'Select the revision view; restore discovery normally uses rejected.'),
    option('around', '--around <N>, --context <N>, -C <N>', 'With --search, include 0-20 physical paragraphs around each returned match.'),
    option('limit', '--limit <N>', 'Return at most N direct matches; surrounding context does not count.'),
    option('after', '--after <INDEX>', 'Continue after this exclusive 1-based source paragraph index.'),
    option('all', '--all', 'Explicitly bypass default result and soft output limits.')
]);

const MUTATION_DESTINATION_OPTIONS = Object.freeze([
    AUTHOR, OUTPUT, IN_PLACE, FORCE, NO_OVERWRITE
]);

const APPLY_EXAMPLES = Object.freeze([
    {
        description: 'Ordinary tracked replacement; modified is the complete desired accepted-view paragraph.',
        operation: {
            type: 'redline',
            target: { exactText: 'Original clause.', paragraphId: '1A2B3C4D' },
            modified: 'Revised clause.'
        }
    },
    {
        description: 'Localized tracked replacement in one strong accepted-view paragraph target.',
        operation: {
            type: 'redline',
            target: { paragraphId: '1A2B3C4D' },
            replacements: [{ find: 'thirty (30) days', replace: 'sixty (60) days' }]
        }
    },
    {
        description: 'Comment the complete target paragraph.',
        operation: {
            type: 'comment',
            target: { exactText: 'Clause to review.', paragraphId: '2A2B3C4D' },
            commentContent: 'Please confirm this language.'
        }
    },
    {
        description: 'Counterpropose a wholly foreign-deleted paragraph found through --view rejected.',
        operation: {
            type: 'restore',
            target: {
                exactText: 'Deleted source paragraph.',
                paragraphId: '3A2B3C4D',
                revisionView: 'rejected'
            },
            modified: 'Restored and revised paragraph.'
        }
    }
]);

export const CLI_COMMAND_HELP = Object.freeze({
    version: {
        summary: 'Report the machine contract version and capabilities.',
        usage: 'docx-redline version',
        options: [HELP],
        notes: ['Wrappers should require only the capabilities they use.'],
        examples: [{ command: 'docx-redline version' }]
    },
    inspect: {
        summary: 'Inspect detailed paragraph, revision, comment, and structure metadata.',
        usage: 'docx-redline inspect <file.docx> [options]',
        options: INSPECTION_OPTIONS,
        notes: [
            'Prefer a focused --search, --index, or --range. Use extract when only exact edit targets are needed.',
            'P<number>, index, and paragraph ordinals are machine references, not user-facing Word locations.'
        ],
        examples: [
            { command: 'docx-redline inspect contract.docx --search "force majeure" --around 3' },
            { command: 'docx-redline inspect contract.docx --range 10:25 --view rejected' }
        ]
    },
    extract: {
        summary: 'Return compact exact targets and human-facing legal locations.',
        usage: 'docx-redline extract <file.docx> [options]',
        options: INSPECTION_OPTIONS,
        notes: [
            'Search is always case-insensitive.',
            'Copy exactText plus paragraphId or fingerprint into operations; do not cite P<number> to users.',
            'Broad results are paginated; follow selection.nextAfter or pass --all deliberately.'
        ],
        examples: [
            { command: 'docx-redline extract contract.docx --search "force majeure" --around 3' },
            { command: 'docx-redline extract contract.docx --range 10:25' }
        ]
    },
    preflight: {
        summary: 'Check an operation batch without mutating or writing a document.',
        usage: 'docx-redline preflight <file.docx> --operations <file.json|-> [options]',
        options: [
            HELP,
            option('operations', '--operations <file.json|->, --operations-file <file>', 'Read an operation array/envelope from a UTF-8 file or stdin.'),
            AUTHOR,
            option('strictTargets', '--strict-targets', 'Require strict target descriptors.'),
            option('target', '--target <text>', 'Inline one-operation target text.'),
            option('modified', '--modified <text>', 'Complete desired accepted-view target content.'),
            option('comment', '--comment <text>', 'Create an inline comment operation.'),
            option('textToComment', '--text-to-comment <text>', 'Anchor an inline comment to an exact subspan.'),
            option('targetRef', '--target-ref <N>', 'Disambiguate an inline target with a 1-based machine index.'),
            option('existingRevisions', '--existing-revisions <policy>', 'Select the explicit existing-revision policy.')
        ],
        notes: ['Normal apply already performs validation; preflight is optional.'],
        examples: [{ command: 'docx-redline preflight contract.docx --operations operations.json --author "Editor"' }]
    },
    apply: {
        summary: 'Apply canonical document operations and write a derived DOCX.',
        usage: 'docx-redline apply <file.docx> (--operations <file.json|->|<inline operation>) [options]',
        options: [
            HELP,
            option('operations', '--operations <file.json|->, --operations-file <file>', 'Read an operation array/envelope from a UTF-8 file or serializer-backed stdin.'),
            ...MUTATION_DESTINATION_OPTIONS,
            option('noClobber', '--no-clobber', 'Alias of --no-overwrite.'),
            option('expectedRevision', '--expected-revision <token|json>', 'Reject a stale package revision.'),
            option('target', '--target <text>', 'Inline one-operation target text.'),
            option('targetId', '--target-id <paragraphId>', 'Strong inline target using a paragraphId from fresh inspection.'),
            option('modified', '--modified <text>', 'Complete desired accepted-view target content.'),
            option('find', '--find <exact text>', 'Exact case-sensitive source span for one localized replacement.'),
            option('replace', '--replace <text>', 'Replacement text for --find; an empty string deletes the span.'),
            option('occurrence', '--occurrence <N>', 'Positive 1-based --find occurrence within the uniquely resolved target paragraph.'),
            option('search', '--search <text>', 'Case-insensitive context anchor for speculative find/replace; prefer a longer, fairly unique phrase because all matches contribute scope.'),
            option('contextRange', '--context-range <START:END>', 'Signed inclusive paragraph offsets (-20..20) from each --search anchor; 1:3 means strictly the next three paragraphs.'),
            option('around', '--around <N>', 'With --search, shorthand for symmetric --context-range -N:N (0-20).'),
            option('comment', '--comment <text>', 'Create an inline comment operation.'),
            option('textToComment', '--text-to-comment <text>', 'Anchor an inline comment to an exact subspan.'),
            option('targetRef', '--target-ref <N>', 'Disambiguate an inline target with a 1-based machine index.'),
            option('existingRevisions', '--existing-revisions <policy>', 'Select revision handling; cross-author slicing must be deliberate.'),
            option('atomic', '--atomic[=true|false]', 'Choose all-or-nothing or progressive batch execution.'),
            option('generateRedlines', '--generate-redlines[=true|false]', 'Control tracked-change generation.'),
            option('noRedlines', '--no-redlines', 'Apply clean text without tracked-change markup.'),
            option('requireComplete', '--require-complete', 'Return exit code 3 for progressive partial completion.'),
            option('profile', '--profile agent', 'Require complete machine execution without selecting atomic/progressive or revision policy; explicit flags compose with it.'),
            COMPACT
        ],
        notes: [
            'modified is complete desired accepted-view content, not only inserted words.',
            'Use either modified or find/replace. A strong accepted-view target is preferred when already known.',
            'Without a strong target, find/replace may resolve globally or within --search plus a directional context range; it fails unless exactly one paragraph is eligible.',
            'Generic or repeated search anchors widen the unioned scope and may cause AMBIGUOUS_TARGET; on success, anchorMatchCount greater than 1 signals that the returned location and excerpts should be confirmed.',
            '--occurrence disambiguates repeated text only within one uniquely selected paragraph, never between paragraphs.',
            'Inline find/replace accepts one replacement; use operations JSON for multiple simultaneous replacements.',
            'Use a structured JSON file or serializer-backed stdin; never interpolate legal text through raw shell quoting.',
            'The source is never overwritten unless --in-place is explicit.',
            'The agent profile keeps progressive execution unless --atomic is explicit and does not change existing-revision policy.',
            'Do not accept/reject foreign revisions or remove comments without user authorization.',
            'Inspect completion, written, outputPath, every result, error.recovery, and retryPlan.'
        ],
        examples: APPLY_EXAMPLES
    },
    accept: {
        summary: 'Accept tracked revisions by one author or all authors.',
        usage: 'docx-redline accept <file.docx> (--author <name>|--all-authors) [options]',
        options: [HELP, ...MUTATION_DESTINATION_OPTIONS, option('allAuthors', '--all-authors', ALL_AUTHORS.description), option('noClobber', '--no-clobber', 'Alias of --no-overwrite.'), COMPACT],
        notes: ['Accepting foreign review content requires explicit user authorization.'],
        examples: [{ command: 'docx-redline accept reviewed.docx --author "Editor"' }]
    },
    reject: {
        summary: 'Reject tracked revisions by one author or all authors.',
        usage: 'docx-redline reject <file.docx> (--author <name>|--all-authors) [options]',
        options: [HELP, ...MUTATION_DESTINATION_OPTIONS, option('allAuthors', '--all-authors', ALL_AUTHORS.description), option('noClobber', '--no-clobber', 'Alias of --no-overwrite.'), COMPACT],
        notes: ['Rejecting foreign review content requires explicit user authorization.'],
        examples: [{ command: 'docx-redline reject reviewed.docx --author "Editor"' }]
    },
    'delete-comments': {
        summary: 'Delete comments by one author or all authors.',
        usage: 'docx-redline delete-comments <file.docx> (--author <name>|--all-authors) [options]',
        options: [HELP, ...MUTATION_DESTINATION_OPTIONS, option('allAuthors', '--all-authors', ALL_AUTHORS.description), option('noClobber', '--no-clobber', 'Alias of --no-overwrite.'), COMPACT],
        notes: ['Removing reviewer comments requires explicit user authorization.'],
        examples: [{ command: 'docx-redline delete-comments reviewed.docx --author "Reviewer"' }]
    },
    validate: {
        summary: 'Validate revision markup and DOCX package wiring.',
        usage: 'docx-redline validate <file.docx> [--baseline <file.docx>]',
        options: [HELP, option('baseline', '--baseline <file.docx>', 'Report only validation issues introduced relative to a baseline package.')],
        notes: ['Apply validates before writing; use this command for an explicit audit.'],
        examples: [{ command: 'docx-redline validate reviewed.docx --baseline contract.docx' }]
    }
});

export const CLI_COMMANDS = Object.freeze(Object.keys(CLI_COMMAND_HELP));

export function commandOptionKeys(command) {
    return (CLI_COMMAND_HELP[command]?.options || []).map(item => item.key);
}

const DOCS_BASE_URL = 'https://github.com/AnsonLai/docx-redline-js/blob/main/';
export const CLI_DOCUMENTATION_URLS = Object.freeze([
    `${DOCS_BASE_URL}AGENTS.md`,
    `${DOCS_BASE_URL}docs/AGENT_FAST_START.md`,
    `${DOCS_BASE_URL}docs/SKILL_AUTHORING.md`,
    `${DOCS_BASE_URL}docs/schemas/document-operations.schema.json`
]);

export function buildCliHelp(command = null) {
    if (!command) {
        return {
            status: 'ok',
            command: 'help',
            usage: 'docx-redline <command> [file.docx] [options]',
            commands: CLI_COMMANDS.map(name => ({ name, summary: CLI_COMMAND_HELP[name].summary })),
            notes: ['Run docx-redline <command> --help for flags, semantics, and bounded examples.'],
            documentation: [...CLI_DOCUMENTATION_URLS]
        };
    }
    const entry = CLI_COMMAND_HELP[command];
    if (!entry) return null;
    return {
        status: 'ok',
        command: 'help',
        forCommand: command,
        summary: entry.summary,
        usage: entry.usage,
        options: entry.options.map(({ key: _key, ...publicOption }) => publicOption),
        notes: entry.notes,
        examples: entry.examples,
        documentation: [...CLI_DOCUMENTATION_URLS]
    };
}
