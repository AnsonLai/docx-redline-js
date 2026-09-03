const AUTOMATED = ['js-exact-round-trip', 'runtime-validator', 'synthetic-word'];
const AI_REVIEWED = [...AUTOMATED, 'ai-word-visual-preflight'];
const missing = { status: 'missing' };

function coverage(task, structures, oracles = AUTOMATED) {
    return { task, structures, oracles, manualReview: missing };
}

export const WORD_TASK_COVERAGE = Object.freeze({
    'simple-redline': coverage('replace', ['plain-paragraph']),
    'legal-defined-term-replacement': coverage('replace', ['plain-paragraph']),
    'legal-clause-insertion': coverage('insert', ['plain-paragraph']),
    'legal-sentence-deletion': coverage('delete', ['plain-paragraph']),
    'paragraph-insert': coverage('insert', ['multi-paragraph']),
    'legal-paragraph-deletion': coverage('delete', ['multi-paragraph']),
    'administrative-deadline-change': coverage('replace', ['plain-paragraph']),
    'administrative-procedure-insertion': coverage('insert', ['plain-paragraph']),
    'format-only': coverage('format', ['formatted-runs']),
    'legal-defined-term-italic': coverage('format', ['formatted-runs']),
    'administrative-deadline-underline': coverage('format', ['formatted-runs']),
    'whitespace-heavy': coverage('replace', ['plain-paragraph']),
    'legal-dollar-delimiters-preserved': coverage('insert', ['plain-paragraph']),
    'administrative-literal-escapes-preserved': coverage('replace', ['plain-paragraph']),
    'legal-inline-preface-preserved': coverage('replace', ['plain-paragraph']),
    'administrative-multiline-target': coverage('replace', ['multi-paragraph']),
    'legal-leading-whitespace-preserved': coverage('replace', ['plain-paragraph']),
    'legal-prior-revision-no-op': coverage('accept-reject', ['plain-paragraph', 'prior-revisions']),
    'administrative-atomic-batch-rollback': coverage('mixed-batch', ['plain-paragraph', 'prior-revisions']),
    'legal-hostile-revision-id-clamped': coverage('replace', ['plain-paragraph', 'prior-revisions']),
    'legal-bookmark-adjacent-replacement': coverage('replace', ['plain-paragraph', 'bookmark']),
    'legal-internal-hyperlink-adjacent-replacement': coverage('replace', ['multi-paragraph', 'bookmark', 'hyperlink']),
    'legal-mixed-run-formatting-preserved': coverage('replace', ['formatted-runs']),
    'administrative-content-control-replacement': coverage('replace', ['plain-paragraph', 'content-control']),
    'administrative-table-cell-replacement': coverage('replace', ['table']),
    'administrative-list-change-append-item': coverage('list-change', ['list'], AI_REVIEWED),
    'legal-list-change-append-multiple-items': coverage('list-change', ['list'], AI_REVIEWED),
    'administrative-list-change-nested-child': coverage('list-change', ['list'], AI_REVIEWED),
    'legal-list-change-middle-range-insertion': coverage('list-change', ['list'], AI_REVIEWED),
    'legal-list-change-upper-roman-section': coverage('list-change', ['list'], AI_REVIEWED),
    'legal-list-change-lower-roman-subclause': coverage('list-change', ['list'], AI_REVIEWED),
    'legal-list-change-parenthesized-lower-letter': coverage('list-change', ['list'], AI_REVIEWED),
    'administrative-list-change-upper-letter-agenda': coverage('list-change', ['list'], AI_REVIEWED),
    'administrative-list-change-dash-bullet': coverage('list-change', ['list'], AI_REVIEWED),
    'legal-list-change-symbol-bullet-multiple': coverage('list-change', ['list'], AI_REVIEWED),
    'administrative-table-reconciliation-cell-update': coverage('table-reconciliation', ['table'], AI_REVIEWED),
    'legal-table-reconciliation-row-insertion': coverage('table-reconciliation', ['table'], AI_REVIEWED),
    'administrative-table-reconciliation-row-deletion': coverage('table-reconciliation', ['table'], AI_REVIEWED),
    'legal-table-reconciliation-multi-cell-update': coverage('table-reconciliation', ['table'], AI_REVIEWED),
    'administrative-tab-aligned-status': coverage('replace', ['plain-paragraph', 'tab-break'], AI_REVIEWED),
    'administrative-boundary-tabs-preserved': coverage('replace', ['plain-paragraph', 'tab-break'], AI_REVIEWED),
    'legal-locked-field-adjacent-replacement': coverage('replace', ['plain-paragraph', 'field'], AI_REVIEWED),
    'administrative-comment-anchor-adjacent-replacement': coverage('replace', ['plain-paragraph', 'comment'], AI_REVIEWED),
    'administrative-footnote-adjacent-deadline': coverage('replace', ['plain-paragraph', 'note'], AI_REVIEWED),
    'legal-endnote-adjacent-duration': coverage('replace', ['plain-paragraph', 'note'], AI_REVIEWED),
    'administrative-header-footer-package': coverage('replace', ['plain-paragraph', 'header-footer', 'section-boundary'], AI_REVIEWED),
    'legal-external-hyperlink-adjacent-replacement': coverage('replace', ['plain-paragraph', 'hyperlink'], AI_REVIEWED)
});
