export const COVERAGE_TASKS = Object.freeze([
    'replace',
    'insert',
    'delete',
    'format',
    'comment',
    'accept-reject',
    'list-change',
    'table-reconciliation',
    'mixed-batch'
]);

export const COVERAGE_STRUCTURES = Object.freeze([
    'plain-paragraph',
    'multi-paragraph',
    'formatted-runs',
    'list',
    'table',
    'bookmark',
    'hyperlink',
    'content-control',
    'tab-break',
    'field',
    'comment',
    'note',
    'header-footer',
    'section-boundary',
    'prior-revisions'
]);

export const COVERAGE_ORACLES = Object.freeze([
    'js-exact-round-trip',
    'runtime-validator',
    'xsd',
    'libreoffice',
    'synthetic-word',
    'real-document-word',
    'ai-word-visual-preflight',
    'human-word-visual-review'
]);

export const MANUAL_REVIEW_STATES = Object.freeze([
    'missing',
    'current',
    'stale'
]);

const vocabularySets = {
    task: new Set(COVERAGE_TASKS),
    structures: new Set(COVERAGE_STRUCTURES),
    oracles: new Set(COVERAGE_ORACLES)
};

export function validateCoverageMetadata(metadata, identity) {
    if (!metadata || typeof metadata !== 'object') {
        throw new Error(`${identity}: coverage metadata is missing`);
    }
    if (!vocabularySets.task.has(metadata.task)) {
        throw new Error(`${identity}: unknown coverage task ${JSON.stringify(metadata.task)}`);
    }
    for (const field of ['structures', 'oracles']) {
        if (!Array.isArray(metadata[field]) || metadata[field].length === 0) {
            throw new Error(`${identity}: ${field} must be a non-empty array`);
        }
        if (new Set(metadata[field]).size !== metadata[field].length) {
            throw new Error(`${identity}: ${field} contains duplicates`);
        }
        for (const label of metadata[field]) {
            if (!vocabularySets[field].has(label)) {
                throw new Error(`${identity}: unknown ${field} label ${JSON.stringify(label)}`);
            }
        }
    }
    if (!MANUAL_REVIEW_STATES.includes(metadata.manualReview?.status)) {
        throw new Error(`${identity}: manualReview.status must be missing, current, or stale`);
    }
    if (metadata.manualReview.status !== 'missing' && !metadata.manualReview.reviewedAt) {
        throw new Error(`${identity}: reviewedAt is required for a ${metadata.manualReview.status} review`);
    }
    return metadata;
}

export function sortCoverageMetadata(metadata) {
    return {
        ...metadata,
        structures: [...metadata.structures].sort(
            (a, b) => COVERAGE_STRUCTURES.indexOf(a) - COVERAGE_STRUCTURES.indexOf(b)
        ),
        oracles: [...metadata.oracles].sort(
            (a, b) => COVERAGE_ORACLES.indexOf(a) - COVERAGE_ORACLES.indexOf(b)
        )
    };
}
