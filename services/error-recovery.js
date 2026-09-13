export const ERROR_RECOVERY_VERSION = 1;

const RULES = Object.freeze({
    INVALID_OPERATION: ['request', 'request_fixable', 'change_request'],
    INVALID_AGENT_REQUEST: ['request', 'request_fixable', 'change_request'],
    INVALID_FILTER: ['request', 'request_fixable', 'change_request'],
    INVALID_OPERATIONS_FILE: ['request', 'request_fixable', 'change_request'],
    OPERATIONS_REQUIRED: ['request', 'request_fixable', 'change_request'],
    UNKNOWN_OPTION: ['request', 'request_fixable', 'change_request'],
    UNEXPECTED_ARGUMENT: ['request', 'request_fixable', 'change_request'],
    INVALID_ACTION: ['request', 'request_fixable', 'change_request'],
    INVALID_PROFILE: ['request', 'request_fixable', 'change_request'],
    INVALID_REVISION_TOKEN: ['revision-check', 'request_fixable', 'change_request'],
    BATCH_OPERATION_FAILED: ['batch-execution', 'request_fixable', 'inspect_failed_operations'],
    REVISION_TOKEN_SCOPE_MISMATCH: ['revision-check', 'request_fixable', 'change_request'],
    REVISION_MISMATCH: ['revision-check', 'target_refresh_required', 'reinspect'],
    TARGET_NOT_FOUND: ['target-resolution', 'target_refresh_required', 'reinspect'],
    SOURCE_TARGET_NOT_FOUND: ['target-resolution', 'target_refresh_required', 'reinspect'],
    TARGET_TEXT_MISMATCH: ['target-resolution', 'target_refresh_required', 'reinspect'],
    TARGET_FINGERPRINT_MISMATCH: ['target-resolution', 'target_refresh_required', 'reinspect'],
    TARGET_INDEX_MISMATCH: ['target-resolution', 'target_refresh_required', 'reinspect'],
    TARGET_OCCURRENCE_MISMATCH: ['target-resolution', 'candidate_selection_required', 'choose_candidate'],
    STALE_TARGET_HANDLE: ['target-resolution', 'target_refresh_required', 'reinspect'],
    TARGET_HANDLE_NOT_FOUND: ['target-resolution', 'target_refresh_required', 'reinspect'],
    AMBIGUOUS_TARGET: ['target-resolution', 'candidate_selection_required', 'choose_candidate'],
    TARGET_RANGE_INVALID: ['target-resolution', 'request_fixable', 'change_target_range'],
    ANCHOR_NOT_FOUND: ['anchor-resolution', 'request_fixable', 'change_anchor'],
    AMBIGUOUS_ANCHOR: ['anchor-resolution', 'candidate_selection_required', 'choose_candidate'],
    PATCH_SOURCE_NOT_FOUND: ['patch-compilation', 'request_fixable', 'change_patch'],
    AMBIGUOUS_PATCH_SOURCE: ['patch-compilation', 'candidate_selection_required', 'choose_occurrence'],
    OVERLAPPING_PATCHES: ['patch-compilation', 'source_conflict', 'combine_patches'],
    CONFLICTING_PATCHES: ['patch-compilation', 'source_conflict', 'combine_patches'],
    PATCH_ROUNDTRIP_MISMATCH: ['patch-validation', 'request_fixable', 'reinspect_and_narrow'],
    DIFF_TOKEN_LIMIT: ['patch-compilation', 'request_fixable', 'narrow_operation'],
    STRUCTURED_CONTENT_INVALID: ['request', 'request_fixable', 'change_request'],
    UNSUPPORTED_INSERTION_AFFINITY: ['request', 'request_fixable', 'change_request'],
    EXISTING_REVISIONS: ['target-safety', 'policy_choice_required', 'set_option'],
    COMMENTED_CONTENT_MERGE: ['target-safety', 'user_authorization_required', 'resolve_comments'],
    COMMENTED_CONTENT_DELETE: ['target-safety', 'user_authorization_required', 'resolve_comments'],
    OVERLAPPING_SOURCE_TARGETS: ['batch-compilation', 'source_conflict', 'consolidate_operations'],
    OVERLAPPING_TEXT_EDITS: ['batch-compilation', 'source_conflict', 'consolidate_operations'],
    REVISION_ORDER_CONFLICT: ['batch-compilation', 'source_conflict', 'split_or_consolidate_operations'],
    TARGET_CONSUMED_BY_OPERATION: ['batch-execution', 'source_conflict', 'use_created_content_dependency'],
    CAPTURE_FANOUT_CONFLICT: ['batch-compilation', 'source_conflict', 'split_or_chain_capture_consumers'],
    DUPLICATE_CAPTURE_KEY: ['batch-compilation', 'request_fixable', 'rename_capture'],
    CAPTURE_NOT_FOUND: ['batch-compilation', 'request_fixable', 'add_or_correct_capture'],
    CAPTURE_DEPENDENCY_CYCLE: ['batch-compilation', 'source_conflict', 'replan_batch'],
    AMBIGUOUS_CAPTURE_SELECTION: ['batch-compilation', 'candidate_selection_required', 'choose_candidate'],
    CAPTURE_STALE: ['batch-execution', 'source_conflict', 'replan_batch'],
    GENERATED_OOXML_INVALID: ['validation', 'library_or_builder_failure', 'report_library_failure'],
    DOCUMENT_SERIALIZATION_FAILED: ['serialization', 'library_or_builder_failure', 'report_library_failure'],
    PACKAGE_OPERATION_FAILED: ['package', 'library_or_builder_failure', 'report_library_failure'],
    PACKAGE_VALIDATION: ['package-validation', 'library_or_builder_failure', 'report_library_failure'],
    MUTATION_RECEIPT_MISMATCH: ['receipt-validation', 'library_or_builder_failure', 'report_library_failure'],
    FOREIGN_PARAGRAPH_MARK_DELETION: ['target-safety', 'policy_choice_required', 'use_restore_or_leave_deleted'],
    RESTORATION_STATE_REQUIRED: ['target-safety', 'target_refresh_required', 'reinspect'],
    RESTORATION_COUNT_MISMATCH: ['request', 'request_fixable', 'change_request'],
    REJECTED_INSERTION_STATE_REQUIRED: ['target-safety', 'target_refresh_required', 'reinspect'],
    UNSAFE_REVISION_NESTING: ['target-safety', 'manual_document_resolution', 'manual_resolution'],
    UNSAFE_REVISION_BOUNDARY: ['target-safety', 'manual_document_resolution', 'manual_resolution'],
    UNSAFE_PARAGRAPH_BOUNDARY: ['target-safety', 'manual_document_resolution', 'manual_resolution'],
    UNSAFE_DELETED_TABLE_ROW: ['target-safety', 'manual_document_resolution', 'manual_resolution'],
    UNSUPPORTED_MOVE_REVISION: ['target-safety', 'manual_document_resolution', 'manual_resolution'],
    SECTION_BREAK_PARAGRAPH: ['target-safety', 'manual_document_resolution', 'manual_resolution'],
    UNSAFE_PARAGRAPH_PLACEMENT: ['target-safety', 'manual_document_resolution', 'manual_resolution'],
    UNSUPPORTED_REVISION_VIEW_MUTATION: ['target-safety', 'manual_document_resolution', 'manual_resolution']
});

function issueSummary(error) {
    const issues = Array.isArray(error?.generatedIssues)
        ? error.generatedIssues
        : (Array.isArray(error?.issues) ? error.issues : null);
    if (!issues) return null;
    const byCode = new Map();
    for (const issue of issues) {
        const code = issue?.code || 'UNKNOWN';
        byCode.set(code, (byCode.get(code) || 0) + 1);
    }
    return {
        total: issues.length,
        byCode: Array.from(byCode, ([code, count]) => ({ code, count }))
    };
}

function ruleFor(code) {
    return RULES[code] || ['operation', 'manual_document_resolution', 'inspect_error'];
}

/** Add stable machine recovery metadata while preserving code-specific details. */
export function normalizeErrorWithRecovery(error, context = {}) {
    const source = error && typeof error === 'object' ? error : {};
    const code = typeof source.code === 'string' && source.code ? source.code : 'OPERATION_ERROR';
    const [defaultStage, category, action] = ruleFor(code);
    const details = {};
    for (const [key, value] of Object.entries(source)) {
        if (key === 'name' || key === 'stack' || key === 'message' || key === 'code') continue;
        details[key] = value;
    }
    const requiresAuthorization = category === 'user_authorization_required';
    const requiresReinspection = category === 'target_refresh_required'
        || code === 'PATCH_ROUNDTRIP_MISMATCH';
    const recovery = {
        action,
        sameArgumentsSafe: false,
        requiresReinspection,
        requiresUserAuthorization: requiresAuthorization,
        ...(code === 'EXISTING_REVISIONS' ? {
            field: 'existingRevisions',
            recommendedValue: 'slice-cross-author'
        } : {}),
        ...(source.recovery && typeof source.recovery === 'object' ? source.recovery : {})
    };
    const summary = issueSummary(source);
    const derivedContext = {
        ...context,
        ...(Array.isArray(source.revisionAuthors) ? { revisionAuthors: source.revisionAuthors } : {}),
        ...(source.currentPolicy ? { currentPolicy: source.currentPolicy } : {})
    };
    if (
        ['TARGET_TEXT_MISMATCH', 'TARGET_FINGERPRINT_MISMATCH'].includes(code)
        && Array.isArray(source.candidates)
        && source.candidates.length === 1
    ) {
        derivedContext.currentTarget = source.candidates[0];
        derivedContext.requiresRecomposeModified = true;
    }
    return {
        recoveryVersion: ERROR_RECOVERY_VERSION,
        code,
        message: source.message || String(error),
        stage: source.stage || defaultStage,
        category: source.category || category,
        ...details,
        ...(Object.keys(derivedContext).length > 0 || source.context ? {
            context: { ...(source.context || {}), ...derivedContext }
        } : {}),
        ...(summary ? { issueSummary: summary } : {}),
        recovery
    };
}

export function createRetryPlan({
    atomic = false,
    rolledBack = false,
    results = [],
    receipts = [],
    operationCount = null
} = {}) {
    const count = Number.isInteger(operationCount)
        ? operationCount
        : Math.max(results.length, receipts.length);
    const attempted = new Set(results.map(result => result?.index).filter(Number.isInteger));
    const failedIndexes = results
        .filter(result => result?.status === 'error')
        .map(result => result.index)
        .filter(Number.isInteger);
    const committedIndexes = receipts
        .filter(receipt => receipt?.committed === true)
        .map(receipt => receipt.operationIndex)
        .filter(Number.isInteger);
    const unattemptedIndexes = [];
    for (let index = 1; index <= count; index += 1) {
        if (!attempted.has(index)) unattemptedIndexes.push(index);
    }
    const base = rolledBack || atomic || committedIndexes.length === 0 ? 'original' : 'output';
    return {
        base,
        committedIndexes,
        failedIndexes,
        unattemptedIndexes,
        replayWholeBatch: base === 'original',
        sameArgumentsSafe: false
    };
}
