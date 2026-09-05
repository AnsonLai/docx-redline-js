/**
 * Read-only preflight for document operation batches.
 */

import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { getDefaultAuthor } from '../adapters/config.js';
import { containsTrackedChanges } from '../core/word-xml.js';
import {
    buildParagraphMetadataIndex,
    createParagraphFingerprint,
    findContainingWordElement,
    getDocumentParagraphNodes,
    getParagraphId,
    getParagraphText,
    normalizeWhitespaceForTargeting,
    resolveTargetParagraph
} from '../core/paragraph-targeting.js';
import { createParagraphTextIndex, resolveTextInParagraphIndex } from './comment-locator.js';
import {
    normalizeDocumentOperation,
    resolveDocumentOperationAuthor,
    validateDocumentOperation
} from './document-operation-contract.js';

function normalizedError(error) {
    return {
        code: typeof error?.code === 'string' && error.code ? error.code : 'OPERATION_ERROR',
        message: error?.message || String(error),
        ...(Array.isArray(error?.candidates) ? { candidates: error.candidates } : {})
    };
}

function operationNeedsNumbering(operation) {
    if (operation.operationKind !== 'redline' || typeof operation.modified !== 'string') return false;
    return operation.modified.split(/\r?\n/).some(line => /^\s*(?:[-+*]|\d+[.)])\s+/.test(line));
}

function targetMetadata(xmlDoc, paragraph, resolvedBy, suppliedText, paragraphMetadataIndex = null, revisionView = 'accepted') {
    const cached = paragraphMetadataIndex?.byParagraph?.get(paragraph) || null;
    const paragraphs = cached ? null : getDocumentParagraphNodes(xmlDoc);
    const actualText = cached?.text ?? extractCanonicalParagraphText(paragraph, { revisionView });
    const normalizedSupplied = normalizeWhitespaceForTargeting(suppliedText || '');
    const normalizedActual = normalizeWhitespaceForTargeting(actualText);
    return {
        resolvedBy,
        resolvedTarget: {
            index: cached?.index ?? paragraphs.indexOf(paragraph) + 1,
            paragraphId: cached?.paragraphId ?? getParagraphId(paragraph),
            text: actualText,
            fingerprint: cached?.fingerprint ?? createParagraphFingerprint(paragraph, { text: actualText, revisionView }),
            inTable: cached?.inTable ?? !!findContainingWordElement(paragraph, 'tbl'),
            revisionView
        },
        matchDiagnostics: {
            exactTextMatch: typeof suppliedText === 'string' && suppliedText === actualText,
            normalizedTextMatch: !!normalizedSupplied && normalizedSupplied === normalizedActual,
            suppliedText: suppliedText || '',
            actualText,
            revisionView
        }
    };
}

function buildConflict(code, message, operationIndexes, target) {
    return { code, message, operationIndexes, target };
}

function getCommentIdsInParagraph(paragraph) {
    const ids = new Set();
    for (const localName of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
        for (const node of Array.from(paragraph?.getElementsByTagNameNS?.('*', localName) || [])) {
            const id = node.getAttribute('w:id') || node.getAttribute('id');
            if (id !== '') ids.add(id);
        }
    }
    return Array.from(ids).sort((a, b) => Number(a) - Number(b) || a.localeCompare(b));
}

export function preflightOperations(documentXml, operations, author, options = {}) {
    const parsed = parseOoxmlSafe(documentXml, 'application/xml');
    if (parsed.error || !parsed.doc) {
        return {
            valid: false,
            status: 'error',
            error: parsed.error,
            results: [],
            conflicts: [],
            authorsUsed: [],
            requiredArtifacts: { comments: false, numbering: false }
        };
    }

    const xmlDoc = parsed.doc;
    const metadataIndices = {
        accepted: buildParagraphMetadataIndex(xmlDoc, { revisionView: 'accepted' }),
        rejected: null
    };
    const sourceOperations = Array.isArray(operations) ? operations : [];
    const strictTargets = options.strictTargets !== false;
    const results = [];
    const authorsUsed = new Set();
    let commentsRequired = false;
    let numberingRequired = false;

    for (let index = 0; index < sourceOperations.length; index++) {
        const sourceOperation = sourceOperations[index];
        const validation = validateDocumentOperation(sourceOperation);
        const fallbackOperation = normalizeDocumentOperation(sourceOperation);
        const operation = validation.operation || fallbackOperation;
        const authorUsed = resolveDocumentOperationAuthor(operation, author, getDefaultAuthor());
        authorsUsed.add(authorUsed);

        if (!validation.valid) {
            results.push({
                index: index + 1,
                type: sourceOperation?.type || 'redline',
                operationType: operation.operationKind,
                status: 'error',
                authorUsed,
                error: validation.error
            });
            continue;
        }

        commentsRequired = commentsRequired || operation.operationKind === 'comment';
        numberingRequired = numberingRequired || operationNeedsNumbering(operation);

        const targetView = operation.targetDescriptor?.revisionView === 'rejected' ? 'rejected' : 'accepted';
        let currentMetadataIndex = targetView === 'rejected'
            ? (metadataIndices.rejected || (metadataIndices.rejected = buildParagraphMetadataIndex(xmlDoc, { revisionView: 'rejected' })))
            : metadataIndices.accepted;

        try {
            const resolved = resolveTargetParagraph(xmlDoc, {
                targetText: operation.target,
                targetRef: operation.targetRef,
                targetDescriptor: operation.targetDescriptor,
                opType: operation.operationKind,
                strictAmbiguity: strictTargets,
                paragraphMetadataIndex: currentMetadataIndex,
                metadataIndices,
                onInfo: options.onInfo,
                onWarn: options.onWarn
            });
            const paragraph = resolved.paragraph;
            const metadata = targetMetadata(xmlDoc, paragraph, resolved.resolvedBy, operation.target, currentMetadataIndex, targetView);
            const paragraphText = metadata.resolvedTarget.text;
            const anchor = operation.operationKind === 'comment'
                ? (operation.textToComment || paragraphText)
                : (operation.operationKind === 'highlight' ? operation.textToHighlight : null);
            const anchorResolution = operation.operationKind === 'comment' && anchor != null
                ? resolveTextInParagraphIndex(createParagraphTextIndex(paragraph, { revisionView: targetView }), anchor)
                : null;
            const anchorFound = anchor == null
                || (anchorResolution ? anchorResolution.found : paragraphText.includes(anchor));
            const hasRevisions = containsTrackedChanges(paragraph);
            const existingPolicy = operation.existingRevisions
                || options.existingRevisions
                || 'reject-input';
            const deletingWholeParagraph = operation.operationKind === 'redline' && operation.modified === '';
            const commentIds = deletingWholeParagraph ? getCommentIdsInParagraph(paragraph) : [];

            let error = null;
            if (!anchorFound) {
                error = anchorResolution?.error || {
                    code: 'ANCHOR_NOT_FOUND',
                    message: `Anchor text was not found in target paragraph: "${anchor}".`
                };
            } else if (commentIds.length > 0) {
                const comments = commentIds
                    .map(id => options._existingCommentDetails?.[id])
                    .filter(Boolean);
                error = {
                    code: 'COMMENTED_CONTENT_DELETE',
                    message: 'Refusing to delete a paragraph with existing comments. Resolve or explicitly remove the comments before deleting the paragraph.',
                    commentIds,
                    ...(comments.length > 0 ? { comments } : {})
                };
            } else if (
                operation.operationKind === 'redline'
                && hasRevisions
                && existingPolicy === 'reject-input'
            ) {
                error = {
                    code: 'EXISTING_REVISIONS',
                    message: 'Target paragraph contains tracked changes and existingRevisions is "reject-input".'
                };
            }

            results.push({
                index: index + 1,
                type: sourceOperation?.type || 'redline',
                operationType: operation.operationKind,
                status: error ? 'error' : 'ready',
                authorUsed,
                ...metadata,
                anchor: anchor == null ? null : {
                    text: anchor,
                    found: anchorFound,
                    ...(anchorResolution?.found ? {
                        resolvedBy: anchorResolution.resolvedBy,
                        start: anchorResolution.start,
                        end: anchorResolution.end
                    } : {}),
                    ...(!anchorResolution?.found && Array.isArray(anchorResolution?.error?.candidates)
                        ? { candidates: anchorResolution.error.candidates }
                        : {})
                },
                hasRevisions,
                existingRevisions: existingPolicy,
                ...(error ? { error } : {})
            });
        } catch (error) {
            results.push({
                index: index + 1,
                type: sourceOperation?.type || 'redline',
                operationType: operation.operationKind,
                status: 'error',
                authorUsed,
                error: normalizedError(error)
            });
        }
    }

    const conflicts = [];
    const byTarget = new Map();
    for (const result of results) {
        const targetIndex = result.resolvedTarget?.index;
        if (!targetIndex) continue;
        if (!byTarget.has(targetIndex)) byTarget.set(targetIndex, []);
        byTarget.get(targetIndex).push(result);
    }

    for (const [targetIndex, targetResults] of byTarget) {
        const redlines = targetResults.filter(result => result.operationType === 'redline');
        const highlights = targetResults.filter(result => result.operationType === 'highlight');
        const target = targetResults[0].resolvedTarget;
        if (redlines.length > 1) {
            conflicts.push(buildConflict(
                'OVERLAPPING_TEXT_EDITS',
                `Multiple text edits target paragraph ${targetIndex}; later operations may use a stale anchor.`,
                redlines.map(result => result.index),
                target
            ));
        }
        if (redlines.length > 0 && highlights.length > 0) {
            conflicts.push(buildConflict(
                'REVISION_ORDER_CONFLICT',
                `A text edit and highlight target paragraph ${targetIndex}; operation order can invalidate the target or existing-revision policy.`,
                [...redlines, ...highlights].map(result => result.index).sort((a, b) => a - b),
                target
            ));
        }
    }

    const hasErrors = results.some(result => result.status === 'error');
    return {
        valid: !hasErrors && conflicts.length === 0,
        status: !hasErrors && conflicts.length === 0 ? 'ok' : 'error',
        results,
        conflicts,
        authorsUsed: Array.from(authorsUsed),
        requiredArtifacts: {
            comments: commentsRequired,
            numbering: numberingRequired
        }
    };
}
