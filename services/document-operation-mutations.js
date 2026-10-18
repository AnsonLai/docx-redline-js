/**
 * OOXML mutation implementations used by the document operation applier.
 */

import { createSerializer, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import {
    resolveListDiagnosticPolicy,
    detectUnboundListMarker,
    buildUnboundListMarkerDiagnostic,
    detectRequestedListFormatMismatch,
    buildListFormatDiagnostic,
    findUnknownParagraphFormatKeys,
    PARAGRAPH_FORMAT_REVISION_CONFLICT,
    PARAGRAPH_NUMBERING_NOT_FOUND,
    buildUnknownParagraphFormatDiagnostic,
    diagnosticWarningText
} from './list-conversion-diagnostics.js';
import { getAllReferencedNumIdElements, getCurrentChild } from '../core/current-properties.js';
import { findReconstructionParagraphRange } from '../engine/reconstruction-mapper.js';
import { createRevisionMetadata } from '../core/types.js';
import { containsTrackedChanges, getTrackedChangeAuthors, hasUnattributedTrackedChange, createWordElement, resolveExistingRevisionsPolicy, withOoxmlSourceType } from '../core/word-xml.js';
import {
    snapshotAndAttachRPrChange,
    snapshotAndAttachPPrChange
} from '../engine/run-builders.js';
import { markParagraphInserted, markParagraphDeleted } from '../engine/paragraph-mark-revisions.js';
import {
    applyCharacterFormatToRPr,
    checkRunPropertiesChanged,
    applyParagraphPropertiesToPPr,
    checkParagraphPropertiesChanged
} from '../engine/rpr-helpers.js';
import { refreshRunPropertyChangeIds } from '../core/revision-cloning.js';
import { getDefaultAuthor } from '../adapters/config.js';
import { applyRedlineToOxml as applyRedlineToOxmlEngine } from '../engine/oxml-engine.js';
import {
    getParagraphRestorationRefusal,
    inspectForeignDeletedParagraphTarget
} from '../core/paragraph-revision-safety.js';
import { extractCanonicalParagraphText } from '../core/paragraph-text.js';
import { validateRedlineOoxml } from '../core/redline-validation.js';
import { subtractValidationIssueMultiset, validationErrors } from '../core/validation-delta.js';
import { clonePropertiesWithoutRevisionHistory } from '../core/revision-cloning.js';
import {
    getRunContentPieces,
    getRunTextLength,
    splitTrackChangeCarrier
} from '../engine/surgical-run-splitting.js';
import { applyHighlightToOoxml } from '../engine/formatting-removal.js';
import { parseTable as parseMarkdownTable } from '../pipeline/pipeline.js';
import { injectCommentsIntoOoxml } from './comment-engine.js';
import {
    getParagraphText as getParagraphTextFromOxml,
    getParagraphId,
    createParagraphFingerprint,
    isMarkdownTableText,
    normalizeWhitespaceForTargeting,
    findContainingWordElement,
    resolveTargetParagraphWithSnapshot as resolveTargetParagraphWithSnapshotShared,
    resolveParagraphRangeByRefs,
    getDocumentParagraphNodes,
    validateParagraphBoundaryMutation
} from '../core/paragraph-targeting.js';
import {
    synthesizeExpandedListScopeEdit,
    planListInsertionOnlyEdit,
    getParagraphListInfo,
    stripRedundantLeadingListMarkers
} from '../core/list-targeting.js';
import {
    synthesizeTableMarkdownFromMultilineCellEdit,
    inferTableReplacementParagraphBlock
} from '../core/table-targeting.js';
import {
    buildSingleLineListStructuralFallbackPlan,
    executeSingleLineListStructuralFallback,
    resolveSingleLineListFallbackNumberingAction,
    recordSingleLineListFallbackExplicitSequence,
    clearSingleLineListFallbackExplicitSequence,
    enforceListBindingOnParagraphNodes,
    stripSingleLineListMarkerPrefix
} from '../orchestration/list-structural-fallback.js';
import {
    reserveNextNumberingIdPair,
    remapNumberingPayloadForDocument,
    overwriteParagraphNumIds,
    extractFirstParagraphNumId,
    buildExplicitDecimalMultilevelNumberingXml
} from './numbering-helpers.js';
import {
    allocateNumberingForReplacement,
    checkNumberingBinding,
    continueBatchListSequence,
    mergeNumberingParts
} from './numbering-allocation.js';
import { convertPlainParagraphsToList, numIdIsLive, sequenceMapOf } from './list-format-mutation.js';
import { changeListParagraphsFormat, joinListParagraphs } from './list-format-change.js';
import { LIST_FORMAT_UNCHANGED, normalizeListFormatRequest } from './list-format.js';
import {
    extractReplacementNodesFromOoxml,
    normalizeBodySectionOrderStandalone
} from './standalone-docx-plumbing.js';
import { prepareRevisionAllocator } from './document-operation-session.js';
import {
    buildExplicitRangeInsertionEntries,
    deriveParagraphThenTableAppend,
    deriveSingleParagraphListAdjacencyInsertion,
    deriveSingleParagraphPlainAdjacencyInsertion
} from './operation-heuristics.js';
import { generateTableOoxml } from './table-reconciliation.js';
import { preprocessMarkdown } from '../pipeline/markdown-processor.js';
import { resolveTargetFromCapture, ensureParagraphIdsOnImportedNode } from './capture-engine.js';
import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml
} from './revision-comment-management.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

function getCommentIdsInElement(element) {
    const ids = new Set();
    for (const localName of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
        for (const node of Array.from(element?.getElementsByTagNameNS?.(NS_W, localName) || [])) {
            const id = node.getAttribute('w:id') || node.getAttribute('id');
            if (id !== '') ids.add(id);
        }
    }
    return Array.from(ids).sort((a, b) => Number(a) - Number(b) || a.localeCompare(b));
}

function commentDetailsForIds(ids, detailsById) {
    if (!detailsById || typeof detailsById !== 'object') return [];
    return ids.map(id => detailsById[id]).filter(Boolean);
}

function resolveMutationDocument(documentXml, options = {}) {
    const operationSession = options?._documentOperationSession || null;
    if (operationSession?.valid && operationSession.document) {
        return {
            xmlDoc: operationSession.document,
            serializer: operationSession.serializer,
            operationSession
        };
    }
    return {
        xmlDoc: parseOoxmlSafe(documentXml, 'application/xml').doc,
        serializer: createSerializer(),
        operationSession: null
    };
}

function completedDocumentXml(xmlDoc, serializer, documentXml, operationSession) {
    return operationSession ? documentXml : serializer.serializeToString(xmlDoc);
}

async function applyRedlineToOxml(oxml, originalText, modifiedText, options = {}) {
    const result = await applyRedlineToOxmlEngine(oxml, originalText, modifiedText, options);
    if (result?.useNativeApi && typeof result?.oxml !== 'string') {
        return withOoxmlSourceType({
            ...result,
            oxml,
            hasChanges: false,
            warnings: [
                ...(Array.isArray(result?.warnings) ? result.warnings : []),
                'Standalone mode cannot execute native Word API fallback for this operation.'
            ]
        });
    }
    return result;
}

async function reconcileMarkdownTableOoxml(oxml, originalText, markdownTable, options = {}) {
    const sourceOoxml = typeof oxml === 'string' ? oxml : '';
    const tableText = typeof markdownTable === 'string' ? markdownTable : String(markdownTable || '');
    let tableData;
    try {
        tableData = parseMarkdownTable(tableText);
    } catch {
        tableData = { headers: [], rows: [] };
    }
    if (!((tableData?.headers?.length || 0) > 0 || (tableData?.rows?.length || 0) > 0)) {
        return {
            oxml: sourceOoxml,
            hasChanges: false,
            isMarkdownTable: false,
            warnings: ['Could not parse Markdown table from input.']
        };
    }
    return {
        ...await applyRedlineToOxml(sourceOoxml, originalText || '', tableText, options),
        isMarkdownTable: true,
        tableData
    };
}

function getParagraphText(paragraph) {
    return getParagraphTextFromOxml(paragraph);
}

function escapeInvisibleText(value, maxLength = 96) {
    const text = String(value ?? '');
    const bounded = text.length > maxLength ? `${text.slice(0, maxLength)}…` : text;
    return bounded
        .replace(/\\/g, '\\\\')
        .replace(/\r/g, '\\r')
        .replace(/\n/g, '\\n')
        .replace(/\t/g, '\\t')
        .replace(/\u00a0/g, '\\u00A0')
        .replace(/\u202f/g, '\\u202F')
        .replace(/\u2060/g, '\\u2060');
}

function codePointLabel(value, offset) {
    if (offset >= value.length) return 'END';
    return `U+${value.codePointAt(offset).toString(16).toUpperCase().padStart(4, '0')}`;
}

function describeTargetTextMatch(actualText, requestedText) {
    const actual = String(actualText ?? '');
    const requested = String(requestedText ?? '');
    if (actual === requested) return { mode: 'exact' };

    const differences = [];
    const limit = Math.min(actual.length, requested.length);
    for (let offset = 0; offset < limit && differences.length < 8; offset++) {
        if (actual[offset] === requested[offset]) continue;
        differences.push({
            offset,
            sourceCodePoint: codePointLabel(actual, offset),
            requestedCodePoint: codePointLabel(requested, offset)
        });
    }
    if (actual.length !== requested.length && differences.length < 8) {
        differences.push({
            offset: limit,
            sourceCodePoint: codePointLabel(actual, limit),
            requestedCodePoint: codePointLabel(requested, limit)
        });
    }

    const spaceEquivalent = actual.length === requested.length
        && actual.replace(/\u00a0/g, ' ') === requested.replace(/\u00a0/g, ' ');
    return {
        mode: spaceEquivalent ? 'space_equivalent' : 'normalized',
        differences,
        sourceExcerpt: escapeInvisibleText(actual),
        requestedExcerpt: escapeInvisibleText(requested)
    };
}

function resolvedTargetDiagnostics(xmlDoc, paragraph, metadata, targetText, options) {
    const revisionView = options?.targetDescriptor?.revisionView === 'rejected'
        ? 'rejected'
        : 'accepted';
    const viewMetadata = metadata?.revisionView === revisionView ? metadata : null;
    const index = viewMetadata?.index ?? getDocumentParagraphNodes(xmlDoc).indexOf(paragraph) + 1;
    const paragraphId = viewMetadata?.paragraphId ?? getParagraphId(paragraph);
    const inTable = viewMetadata?.inTable ?? !!findContainingWordElement(paragraph, 'tbl');
    const text = viewMetadata?.text ?? extractCanonicalParagraphText(paragraph, { revisionView });
    const fingerprint = viewMetadata?.fingerprint ?? createParagraphFingerprint(paragraph, {
        text,
        index,
        paragraphId,
        inTable,
        revisionView
    });
    return {
        index,
        paragraphId,
        text,
        fingerprint,
        revisionView,
        inTable,
        targetTextMatch: describeTargetTextMatch(
            text,
            options?.targetDescriptor?.text ?? options?.targetDescriptor?.exactText ?? targetText
        )
    };
}

function resolveTargetParagraph(xmlDoc, targetText, targetRef, opType, runtimeContext = null, options = {}) {
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => { };
    const onWarn = typeof options?.onWarn === 'function' ? options.onWarn : () => { };
    const session = options?._documentOperationSession || null;
    const paragraphMetadataIndex = session?.getParagraphMetadataIndex?.() || null;

    if (options?._compiledSourceId && options?._sourceTargetRegistry) {
        const resolved = options._sourceTargetRegistry.resolve(options._compiledSourceId, xmlDoc);
        if (resolved?.error || !resolved?.paragraph) return resolved;
        resolved.resolvedBy = options._compiledResolvedBy || resolved.resolvedBy;
        if (options?._resolutionCapture) {
            const paragraph = resolved.paragraph;
            const metadata = paragraphMetadataIndex?.byParagraph?.get(paragraph) || null;
            Object.assign(options._resolutionCapture, {
                resolvedBy: resolved.resolvedBy,
                resolvedTarget: resolvedTargetDiagnostics(xmlDoc, paragraph, metadata, targetText, options)
            });
        }
        return resolved;
    }

    if (options?.targetDescriptor?.captureRef) {
        try {
            const resolved = resolveTargetFromCapture(xmlDoc, session, options.targetDescriptor, opType, options);
            if (options?._resolutionCapture && resolved?.paragraph) {
                const paragraph = resolved.paragraph;
                const metadata = paragraphMetadataIndex?.byParagraph?.get(paragraph) || null;
                Object.assign(options._resolutionCapture, {
                    resolvedBy: resolved.resolvedBy,
                    resolvedTarget: resolvedTargetDiagnostics(xmlDoc, paragraph, metadata, targetText, options)
                });
            }
            return resolved;
        } catch (error) {
            return {
                error: {
                    code: typeof error?.code === 'string' && error.code ? error.code : 'TARGET_NOT_FOUND',
                    message: error?.message || String(error)
                }
            };
        }
    }

    const resolved = resolveTargetParagraphWithSnapshotShared(xmlDoc, {
        targetText,
        targetRef,
        opType,
        targetRefSnapshot: runtimeContext?.targetRefSnapshot || null,
        targetDescriptor: options?.targetDescriptor || null,
        strictAmbiguity: options?.strictTargets === true,
        paragraphMetadataIndex,
        onInfo,
        onWarn
    });
    if (options?._resolutionCapture && resolved?.paragraph) {
        const paragraph = resolved.paragraph;
        const metadata = paragraphMetadataIndex?.byParagraph?.get(paragraph) || null;
        Object.assign(options._resolutionCapture, {
            resolvedBy: resolved.resolvedBy,
            resolvedTarget: resolvedTargetDiagnostics(xmlDoc, paragraph, metadata, targetText, options)
        });
    }
    return resolved;
}

/**
 * Zero-based document-order positions `{ from, to }` of an operation's target (and its targetEnd, when it
 * is a range), resolved against the current DOM without mutating, capturing or warning. Used by the
 * revision fidelity verifier to scope accept-first baselines and the accepted-intent check exactly.
 * Returns null when the target or its end cannot be resolved.
 */
export function resolveOperationSourceRange(xmlDoc, targetText, targetRef, targetEndRef, runtimeContext, options = {}) {
    const quiet = { ...options, _resolutionCapture: null, onInfo: () => { }, onWarn: () => { } };
    const positionOf = paragraph => {
        if (!paragraph) return null;
        const position = getDocumentParagraphNodes(xmlDoc).indexOf(paragraph);
        return position >= 0 ? position : null;
    };
    let start;
    try {
        start = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'redline', runtimeContext, quiet);
    } catch {
        return null;
    }
    const from = positionOf(start?.paragraph);
    if (from === null) return null;
    const descriptor = options.targetEndDescriptor || null;
    let endParagraph = null;
    if (options._compiledSourceEndId && options._sourceTargetRegistry) {
        endParagraph = options._sourceTargetRegistry.resolve(options._compiledSourceEndId, xmlDoc)?.paragraph || null;
    } else if (descriptor) {
        endParagraph = resolveTargetParagraph(xmlDoc, descriptor.text || '', descriptor.index, 'redline', runtimeContext, {
            ...quiet,
            targetDescriptor: descriptor,
            _compiledSourceId: null
        })?.paragraph || null;
    } else if (targetEndRef) {
        endParagraph = resolveTargetParagraphWithSnapshotShared(xmlDoc, {
            targetRef: targetEndRef,
            opType: 'redline',
            targetRefSnapshot: runtimeContext?.targetRefSnapshot || null,
            onInfo: () => { },
            onWarn: () => { }
        })?.paragraph || null;
    } else {
        return { from, to: from };
    }
    const to = positionOf(endParagraph);
    return to !== null && to >= from ? { from, to } : null;
}

function extractReplacementNodes(outputOxml) {
    return extractReplacementNodesFromOoxml(outputOxml);
}

function removeListPackagingSentinel(replacementNodes, warnings) {
    if (!Array.isArray(replacementNodes) || replacementNodes.length === 0) return replacementNodes;
    if (!Array.isArray(warnings) || !warnings.includes('Paragraph expanded to list fragment')) return replacementNodes;

    const last = replacementNodes[replacementNodes.length - 1];
    if (!last || last.localName !== 'p') return replacementNodes;
    const meaningfulContent = Array.from(last.childNodes || []).some(child => (
        child.nodeType === 1 && child.localName !== 'pPr'
    ));
    if (meaningfulContent) return replacementNodes;
    return replacementNodes.slice(0, -1);
}

function normalizeBodySectionOrder(xmlDoc) {
    normalizeBodySectionOrderStandalone(xmlDoc);
}

function removeProofErrNodes(paragraph) {
    for (const node of Array.from(paragraph?.getElementsByTagNameNS?.(NS_W, 'proofErr') || [])) {
        node.parentNode?.removeChild(node);
    }
}

function preprocessRedlineTargetParagraph(targetParagraph) {
    if (!targetParagraph) return;
    removeProofErrNodes(targetParagraph);
}

function getDirectWordChild(element, localName) {
    if (!element) return null;
    return Array.from(element.childNodes || []).find(
        node => node && node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === localName
    ) || null;
}

function computeTableIndexInDocument(xmlDoc, tableElement) {
    if (!xmlDoc || !tableElement) return null;
    const tables = Array.from(xmlDoc.getElementsByTagNameNS(NS_W, 'tbl'));
    const idx = tables.indexOf(tableElement);
    return idx >= 0 ? idx + 1 : null;
}

function normalizeMultilineTableStructuralPayload(text) {
    return String(text || '')
        .replace(/\r\n/g, '\n')
        .split('\n')
        .map(line => line.trim())
        .filter(Boolean)
        .join('\n');
}

function computeTableStructuralDedupeKey(xmlDoc, containingTable, modifiedText) {
    const tableIndex = computeTableIndexInDocument(xmlDoc, containingTable);
    if (!Number.isInteger(tableIndex) || tableIndex < 1) return null;
    const normalizedPayload = normalizeMultilineTableStructuralPayload(modifiedText);
    if (!normalizedPayload) return null;
    return `table:${tableIndex}|payload:${normalizedPayload}`;
}

function ensureListProperties(xmlDoc, paragraph, ilvl, numId) {
    let pPr = getDirectWordChild(paragraph, 'pPr');
    if (!pPr) {
        pPr = createWordElement(xmlDoc, 'w:pPr');
        paragraph.insertBefore(pPr, paragraph.firstChild);
    }

    let numPr = getDirectWordChild(pPr, 'numPr');
    if (!numPr) {
        numPr = createWordElement(xmlDoc, 'w:numPr');
        pPr.appendChild(numPr);
    }

    let ilvlEl = getDirectWordChild(numPr, 'ilvl');
    if (!ilvlEl) {
        ilvlEl = createWordElement(xmlDoc, 'w:ilvl');
        numPr.appendChild(ilvlEl);
    }
    ilvlEl.setAttribute('w:val', String(Math.max(0, Number.parseInt(ilvl, 10) || 0)));

    let numIdEl = getDirectWordChild(numPr, 'numId');
    if (!numIdEl) {
        numIdEl = createWordElement(xmlDoc, 'w:numId');
        numPr.appendChild(numIdEl);
    }
    numIdEl.setAttribute('w:val', String(numId));
}

function buildInsertedListParagraph(xmlDoc, anchorParagraph, entry, revisionMetadata, author, options = {}) {
    const generateRedlines = options.generateRedlines !== false;
    const paragraph = createWordElement(xmlDoc, 'w:p');

    const anchorPPr = getDirectWordChild(anchorParagraph, 'pPr');
    if (anchorPPr) {
        paragraph.appendChild(clonePropertiesWithoutRevisionHistory(anchorPPr));
    }
    ensureListProperties(xmlDoc, paragraph, entry.ilvl, entry.numId);

    // A new list item is a whole inserted paragraph, not merely inserted text.
    // Tracking its paragraph mark lets Word (and our accept/reject helpers)
    // remove the list paragraph itself on Reject All instead of leaving an
    // empty bullet or number behind.
    if (generateRedlines) {
        markParagraphInserted(paragraph, { author });
    }

    const run = createWordElement(xmlDoc, 'w:r');
    const anchorFirstRun = Array.from(anchorParagraph.getElementsByTagNameNS(NS_W, 'r'))[0] || null;
    const anchorRunPr = anchorFirstRun ? getDirectWordChild(anchorFirstRun, 'rPr') : null;
    if (anchorRunPr) {
        run.appendChild(clonePropertiesWithoutRevisionHistory(anchorRunPr));
    }

    const textNode = createWordElement(xmlDoc, 'w:t');
    const safeText = String(entry.text || '').trim();
    if (/^\s|\s$/.test(safeText)) textNode.setAttribute('xml:space', 'preserve');
    textNode.textContent = safeText;
    run.appendChild(textNode);
    if (generateRedlines) {
        const metadata = revisionMetadata || createRevisionMetadata(author, xmlDoc);
        const ins = createWordElement(xmlDoc, 'w:ins');
        ins.setAttribute('w:id', String(metadata.id));
        ins.setAttribute('w:author', metadata.author);
        ins.setAttribute('w:date', metadata.date);
        ins.appendChild(run);
        paragraph.appendChild(ins);
    } else {
        paragraph.appendChild(run);
    }

    return paragraph;
}

function serializeParagraphRangeAsDocument(paragraphs, serializer) {
    const paragraphXml = (paragraphs || [])
        .map(paragraph => serializer.serializeToString(paragraph))
        .join('');
    return `<w:document xmlns:w="${NS_W}"><w:body>${paragraphXml}</w:body></w:document>`;
}
function applyExplicitRangeListInsertions({
    xmlDoc,
    explicitRangeParagraphs,
    insertionEntries,
    generateRedlines,
    author
}) {
    if (!Array.isArray(explicitRangeParagraphs) || explicitRangeParagraphs.length === 0) return false;
    if (!Array.isArray(insertionEntries) || insertionEntries.length === 0) return false;

    const parent = explicitRangeParagraphs[0].parentNode;
    if (!parent || explicitRangeParagraphs.some(paragraph => paragraph.parentNode !== parent)) return false;

    const tailInsertionPoint = explicitRangeParagraphs[explicitRangeParagraphs.length - 1].nextSibling;
    for (const entry of insertionEntries) {
        const referenceParagraph = entry.insertBeforeOriginalIndex != null
            ? explicitRangeParagraphs[entry.insertBeforeOriginalIndex]
            : explicitRangeParagraphs[explicitRangeParagraphs.length - 1];
        if (!referenceParagraph) return false;

        const listParagraph = buildInsertedListParagraph(
            xmlDoc,
            referenceParagraph,
            {
                ilvl: entry.ilvl,
                markerType: entry.markerType,
                numId: entry.numId,
                text: entry.text
            },
            generateRedlines ? createRevisionMetadata(author, xmlDoc) : null,
            author,
            { generateRedlines }
        );

        if (entry.insertBeforeOriginalIndex != null) {
            parent.insertBefore(listParagraph, referenceParagraph);
        } else {
            parent.insertBefore(listParagraph, tailInsertionPoint);
        }
    }

    normalizeBodySectionOrder(xmlDoc);
    return true;
}
function buildFallbackInsertedPlainParagraph(xmlDoc, text, revisionMetadata, author, options = {}) {
    const generateRedlines = options.generateRedlines !== false;
    const paragraph = createWordElement(xmlDoc, 'w:p');
    const run = createWordElement(xmlDoc, 'w:r');
    const textNode = createWordElement(xmlDoc, 'w:t');
    const safeText = String(text || '');
    if (/^\s|\s$/.test(safeText)) textNode.setAttribute('xml:space', 'preserve');
    textNode.textContent = safeText;
    run.appendChild(textNode);

    if (generateRedlines) {
        if (options.trackParagraphMark === true) {
            markParagraphInserted(paragraph, { author });
        }
        const metadata = revisionMetadata || createRevisionMetadata(author, xmlDoc, 'ins');
        const ins = createWordElement(xmlDoc, 'w:ins');
        ins.setAttribute('w:id', String(metadata.id));
        ins.setAttribute('w:author', metadata.author);
        ins.setAttribute('w:date', metadata.date);
        ins.appendChild(run);
        paragraph.appendChild(ins);
    } else {
        paragraph.appendChild(run);
    }

    return paragraph;
}

function buildEmptyParagraphTemplateFromAnchor(xmlDoc, anchorParagraph) {
    const paragraph = createWordElement(xmlDoc, 'w:p');
    const anchorPPr = getDirectWordChild(anchorParagraph, 'pPr');
    if (anchorPPr) paragraph.appendChild(clonePropertiesWithoutRevisionHistory(anchorPPr));

    const run = createWordElement(xmlDoc, 'w:r');
    const anchorFirstRun = Array.from(anchorParagraph.getElementsByTagNameNS(NS_W, 'r'))[0] || null;
    const anchorRunPr = anchorFirstRun ? getDirectWordChild(anchorFirstRun, 'rPr') : null;
    if (anchorRunPr) run.appendChild(clonePropertiesWithoutRevisionHistory(anchorRunPr));

    const textNode = createWordElement(xmlDoc, 'w:t');
    textNode.textContent = '';
    run.appendChild(textNode);
    paragraph.appendChild(run);
    return paragraph;
}

function wrapParagraphContentInInsertion(xmlDoc, paragraph, revisionMetadata, author, options = {}) {
    const wrappedParagraph = createWordElement(xmlDoc, 'w:p');
    const pPr = options.sanitizeParagraphProperties === true
        ? createSanitizedRestorationPPr(xmlDoc, paragraph)
        : clonePropertiesWithoutRevisionHistory(getDirectWordChild(paragraph, 'pPr'));
    if (pPr) wrappedParagraph.appendChild(pPr);
    if (options.paragraphId) wrappedParagraph.setAttributeNS(NS_W14, 'w14:paraId', options.paragraphId);
    if (options.trackParagraphMark === true) {
        markParagraphInserted(wrappedParagraph, { author });
    }

    const ins = createWordElement(xmlDoc, 'w:ins');
    const metadata = revisionMetadata || createRevisionMetadata(author, xmlDoc, 'ins');
    ins.setAttribute('w:id', String(metadata.id));
    ins.setAttribute('w:author', metadata.author);
    ins.setAttribute('w:date', metadata.date);

    for (const child of Array.from(paragraph.childNodes || [])) {
        if (child?.nodeType === 1 && child.namespaceURI === NS_W && child.localName === 'pPr') continue;
        ins.appendChild(child.cloneNode(true));
    }

    wrappedParagraph.appendChild(ins);
    return wrappedParagraph;
}

async function buildInsertedPlainParagraph(xmlDoc, anchorParagraph, text, revisionMetadata, author, options = {}) {
    const generateRedlines = options.generateRedlines !== false;
    const serializer = createSerializer();
    const templateParagraph = buildEmptyParagraphTemplateFromAnchor(xmlDoc, anchorParagraph);
    const templateXml = serializer.serializeToString(templateParagraph);
    const markdownResult = await applyRedlineToOxml(
        templateXml,
        '',
        String(text || ''),
        {
            author,
            generateRedlines: false
        }
    );

    let sourceParagraph = null;
    if (typeof markdownResult?.oxml === 'string') {
        const extracted = extractReplacementNodes(markdownResult.oxml);
        sourceParagraph = (extracted.replacementNodes || []).find(
            node => node && node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === 'p'
        ) || null;
    }

    if (!sourceParagraph) {
        return buildFallbackInsertedPlainParagraph(
            xmlDoc,
            text,
            revisionMetadata,
            author,
            { ...options, generateRedlines }
        );
    }

    if (!generateRedlines) {
        return sourceParagraph;
    }

    return wrapParagraphContentInInsertion(xmlDoc, sourceParagraph, revisionMetadata, author, options);
}

function collectNumberingIdsFromNodes(nodes) {
    const ids = new Set();
    for (const node of nodes || []) {
        // Deliberately historical-inclusive: the numbering-part merge must define every id the content can
        // reference, including ids held only by pPrChange snapshots (Reject restores them).
        for (const numIdNode of getAllReferencedNumIdElements(node)) {
            const val = numIdNode.getAttribute('w:val') || numIdNode.getAttribute('val');
            if (val != null && val !== '') {
                ids.add(String(val));
            }
        }
    }
    return Array.from(ids);
}

function directDeletionCarrierText(carrier) {
    return Array.from(carrier?.childNodes || []).reduce((text, child) => {
        return text + (child?.nodeType === 1 && child.namespaceURI === NS_W && child.localName === 'r'
            ? getRunContentPieces(child).map(piece => piece.text).join('')
            : '');
    }, '');
}

function findOccurrenceOffset(text, needle, occurrence) {
    let from = 0;
    let found = -1;
    for (let index = 0; index < occurrence; index++) {
        found = text.indexOf(needle, from);
        if (found < 0) return -1;
        from = found + Math.max(needle.length, 1);
    }
    return found;
}

function countNonOverlappingOccurrences(text, needle) {
    if (!needle) return 0;
    let count = 0;
    let from = 0;
    while (from <= text.length) {
        const found = text.indexOf(needle, from);
        if (found < 0) break;
        count += 1;
        from = found + needle.length;
    }
    return count;
}

function hasUnsupportedDeletionSplitMarkup(paragraph, carrier) {
    const unsafeParagraphNames = new Set([
        'commentRangeStart', 'commentRangeEnd', 'commentReference',
        'bookmarkStart', 'bookmarkEnd', 'moveFromRangeStart', 'moveFromRangeEnd',
        'moveToRangeStart', 'moveToRangeEnd'
    ]);
    if (Array.from(paragraph.getElementsByTagName?.('*') || []).some(node => unsafeParagraphNames.has(node.localName))) {
        return true;
    }
    for (const child of Array.from(carrier?.childNodes || [])) {
        if (child.nodeType !== 1) continue;
        if (!(child.namespaceURI === NS_W && child.localName === 'r')) return true;
        for (const runChild of Array.from(child.childNodes || [])) {
            if (runChild.nodeType !== 1 || runChild.namespaceURI !== NS_W) continue;
            if (!['rPr', 'delText', 't', 'tab', 'br', 'cr', 'noBreakHyphen', 'softHyphen'].includes(runChild.localName)) return true;
        }
    }
    return false;
}

function insertedRunPropertiesAtDeletionOffset(carrier, localOffset) {
    let offset = 0;
    for (const child of Array.from(carrier?.childNodes || [])) {
        if (!(child?.nodeType === 1 && child.namespaceURI === NS_W && child.localName === 'r')) continue;
        const length = getRunTextLength(getRunContentPieces(child));
        if (localOffset <= offset + length) {
            return clonePropertiesWithoutRevisionHistory(getDirectWordChild(child, 'rPr'));
        }
        offset += length;
    }
    return null;
}

const RESTORATION_PPR_ALLOWLIST = new Set([
    'pStyle', 'numPr', 'ind', 'jc', 'spacing', 'tabs',
    'keepNext', 'keepLines', 'outlineLvl', 'contextualSpacing'
]);

function wordAttribute(node, localName) {
    return node?.getAttribute?.(`w:${localName}`) || node?.getAttribute?.(localName) || '';
}

function normalizedAuthor(author) {
    return String(author || '').trim().toLowerCase();
}

function directWordParagraphSibling(paragraph, direction = 'nextSibling') {
    let cursor = paragraph?.[direction] || null;
    while (cursor) {
        if (cursor.nodeType === 1 && cursor.namespaceURI === NS_W && cursor.localName === 'p') return cursor;
        cursor = cursor[direction] || null;
    }
    return null;
}

function createSanitizedRestorationPPr(xmlDoc, sourceParagraph) {
    const sourcePPr = getDirectWordChild(sourceParagraph, 'pPr');
    const pPr = createWordElement(xmlDoc, 'w:pPr');
    if (!sourcePPr) return pPr;
    for (const child of Array.from(sourcePPr.childNodes || [])) {
        if (child.nodeType !== 1 || child.namespaceURI !== NS_W) continue;
        if (!RESTORATION_PPR_ALLOWLIST.has(child.localName)) continue;
        pPr.appendChild(xmlDoc.importNode(child, true));
    }
    return pPr;
}

function collectRestorationAnchorWarnings(sourceParagraph) {
    const warnings = [];
    const bookmarkNames = new Set();
    for (const node of Array.from(sourceParagraph?.getElementsByTagNameNS?.(NS_W, 'bookmarkStart') || [])) {
        const name = wordAttribute(node, 'name') || '(unnamed)';
        bookmarkNames.add(name);
    }
    for (const name of bookmarkNames) warnings.push(`RESTORATION_DROPPED_BOOKMARK:${name}`);

    const commentIds = new Set();
    for (const localName of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
        for (const node of Array.from(sourceParagraph?.getElementsByTagNameNS?.(NS_W, localName) || [])) {
            const id = wordAttribute(node, 'id');
            if (id !== '') commentIds.add(id);
        }
    }
    for (const id of commentIds) warnings.push(`RESTORATION_DROPPED_COMMENT:${id}`);
    return warnings;
}

function removeRestorationAnchors(paragraph) {
    const anchorNames = [
        'bookmarkStart', 'bookmarkEnd',
        'commentRangeStart', 'commentRangeEnd', 'commentReference'
    ];
    for (const localName of anchorNames) {
        for (const node of Array.from(paragraph?.getElementsByTagNameNS?.(NS_W, localName) || [])) {
            node.parentNode?.removeChild(node);
        }
    }
    for (const run of Array.from(paragraph?.getElementsByTagNameNS?.(NS_W, 'r') || [])) {
        const meaningful = Array.from(run.childNodes || []).some(child => (
            child.nodeType === 1 && child.namespaceURI === NS_W && child.localName !== 'rPr'
        ));
        if (!meaningful) run.parentNode?.removeChild(run);
    }
}

function buildRejectedRestorationTemplate(xmlDoc, sourceParagraph, serializer) {
    const rejected = rejectTrackedChangesInOoxml(serializer.serializeToString(sourceParagraph), { allAuthors: true });
    const parsed = parseOoxmlSafe(rejected.oxml, 'application/xml');
    const rejectedParagraph = parsed.doc?.getElementsByTagNameNS?.(NS_W, 'p')?.[0] || null;
    if (!rejectedParagraph) return null;

    const paragraph = createWordElement(xmlDoc, 'w:p');
    paragraph.appendChild(createSanitizedRestorationPPr(xmlDoc, sourceParagraph));
    for (const child of Array.from(rejectedParagraph.childNodes || [])) {
        if (child.nodeType === 1 && child.namespaceURI === NS_W && child.localName === 'pPr') continue;
        paragraph.appendChild(xmlDoc.importNode(child, true));
    }
    removeRestorationAnchors(paragraph);
    return paragraph;
}

async function editRestorationTemplate(xmlDoc, template, modifiedText, author, serializer) {
    const originalText = extractCanonicalParagraphText(template);
    if (originalText === modifiedText) return template;
    const result = await applyRedlineToOxml(
        serializer.serializeToString(template),
        originalText,
        modifiedText,
        {
            author,
            generateRedlines: false,
            structuredContent: false
        }
    );
    if (result?.status === 'error' || typeof result?.oxml !== 'string') return null;
    const extracted = extractReplacementNodes(result.oxml);
    const paragraph = (extracted.replacementNodes || []).find(node => (
        node?.nodeType === 1 && node.namespaceURI === NS_W && node.localName === 'p'
    ));
    return paragraph ? xmlDoc.importNode(paragraph, true) : null;
}

function allocateFreshParagraphId(xmlDoc, operationSession) {
    const used = new Set(getDocumentParagraphNodes(xmlDoc)
        .map(paragraph => paragraph.getAttributeNS?.(NS_W14, 'paraId') || paragraph.getAttribute?.('w14:paraId') || '')
        .filter(Boolean)
        .map(value => value.toUpperCase()));
    let candidate = null;
    do {
        candidate = operationSession?.generateParagraphId?.()
            || (0x40000000 + used.size + 1).toString(16).toUpperCase();
    } while (used.has(candidate.toUpperCase()));
    return candidate;
}

function trackRestoredParagraph(xmlDoc, paragraph, author, operationSession) {
    const root = xmlDoc.documentElement;
    if (root && !root.getAttribute('xmlns:w14')) root.setAttribute('xmlns:w14', NS_W14);
    return wrapParagraphContentInInsertion(xmlDoc, paragraph, null, author, {
        paragraphId: allocateFreshParagraphId(xmlDoc, operationSession),
        sanitizeParagraphProperties: true,
        trackParagraphMark: true
    });
}

function isInsertedParagraphByAuthor(paragraph, author) {
    const pPr = getDirectWordChild(paragraph, 'pPr');
    const rPr = getDirectWordChild(pPr, 'rPr');
    const marker = getDirectWordChild(rPr, 'ins');
    return !!marker && normalizedAuthor(wordAttribute(marker, 'author')) === normalizedAuthor(author);
}

function followingParagraphBlock(lastSource, count) {
    const paragraphs = [];
    let cursor = lastSource;
    for (let i = 0; i < count; i++) {
        cursor = directWordParagraphSibling(cursor, 'nextSibling');
        if (!cursor) return [];
        paragraphs.push(cursor);
    }
    return paragraphs;
}

function paragraphTextVector(oxml) {
    const parsed = parseOoxmlSafe(oxml, 'application/xml');
    if (parsed.error || !parsed.doc) return null;
    return getDocumentParagraphNodes(parsed.doc).map(paragraph => extractCanonicalParagraphText(paragraph));
}

function sameTextVector(left, right) {
    return Array.isArray(left) && Array.isArray(right)
        && left.length === right.length
        && left.every((value, index) => value === right[index]);
}

function lifecycleTextVector(oxml, action) {
    const result = action === 'accept'
        ? acceptTrackedChangesInOoxml(oxml, { allAuthors: true })
        : rejectTrackedChangesInOoxml(oxml, { allAuthors: true });
    return paragraphTextVector(result.oxml);
}

function buildExpectedRestorationDocument(beforeXml, sourceStartIndex, existingIndexes, templates) {
    const parsed = parseOoxmlSafe(beforeXml, 'application/xml');
    if (parsed.error || !parsed.doc) return null;
    const expectedDoc = parsed.doc;
    const originalParagraphs = getDocumentParagraphNodes(expectedDoc);
    const lastSource = originalParagraphs[sourceStartIndex + templates.length - 1] || null;
    if (!lastSource?.parentNode) return null;

    for (const index of [...existingIndexes].sort((a, b) => b - a)) {
        const paragraph = originalParagraphs[index];
        paragraph?.parentNode?.removeChild(paragraph);
    }
    const insertionPoint = lastSource.nextSibling;
    for (const template of templates) {
        lastSource.parentNode.insertBefore(expectedDoc.importNode(template, true), insertionPoint);
    }
    return createSerializer().serializeToString(expectedDoc);
}

function verifyParagraphRestorationLifecycle(beforeXml, outputXml, sourceStartIndex, existingIndexes, templates, insertedParagraphs = []) {
    const baselineValidation = validateRedlineOoxml(beforeXml);
    const outputValidation = validateRedlineOoxml(outputXml);
    const generatedIssues = subtractValidationIssueMultiset(outputValidation.issues, baselineValidation.issues);
    const envelopeIssues = insertedParagraphs.flatMap(paragraph => (
        validateRedlineOoxml(createSerializer().serializeToString(paragraph)).issues
            .filter(issue => issue.severity === 'error')
    ));
    const generatedErrors = validationErrors(generatedIssues);
    if (generatedErrors.length > 0 || envelopeIssues.length > 0) {
        return {
            valid: false,
            stage: 'validation',
            code: 'GENERATED_OOXML_INVALID',
            generatedIssues,
            envelopeIssues,
            message: [...generatedErrors, ...envelopeIssues].map(issue => issue.message).join(' ')
        };
    }

    const expectedXml = buildExpectedRestorationDocument(beforeXml, sourceStartIndex, existingIndexes, templates);
    if (!expectedXml) return { valid: false, stage: 'expected-document', message: 'Could not construct the restoration lifecycle oracle.' };

    const comparisons = [
        ['current', paragraphTextVector(expectedXml), paragraphTextVector(outputXml)],
        ['accept-all', lifecycleTextVector(expectedXml, 'accept'), lifecycleTextVector(outputXml, 'accept')],
        ['reject-all', lifecycleTextVector(beforeXml, 'reject'), lifecycleTextVector(outputXml, 'reject')]
    ];
    for (const [stage, expected, actual] of comparisons) {
        if (!sameTextVector(expected, actual)) {
            return {
                valid: false,
                stage,
                expected,
                actual,
                message: `Paragraph restoration ${stage} lifecycle text did not match the expected body-scoped paragraph sequence.`
            };
        }
    }
    return { valid: true };
}

/**
 * Inserts a new tracked run at an exact offset inside text visible only in the
 * rejected view of a wholly deleted paragraph. The foreign deletion is split,
 * never rewritten or re-authored.
 */
export async function insertIntoRejectedDeletedText(
    documentXml,
    targetText,
    anchor,
    modified,
    author,
    targetRef = null,
    runtimeContext = null,
    options = {}
) {
    const { serializer, xmlDoc, operationSession } = resolveMutationDocument(documentXml, options);
    if (!xmlDoc) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: { code: 'PARSE_ERROR', message: 'Could not parse document OOXML.' }
        };
    }

    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'insert', runtimeContext, options);
    if (resolved?.error || !resolved?.paragraph) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: resolved?.error || { code: 'TARGET_NOT_FOUND', message: 'Rejected-view insertion target was not found.' }
        };
    }

    const paragraph = resolved.paragraph;
    const state = inspectForeignDeletedParagraphTarget(paragraph, author);
    if (!state.matches) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: {
                code: 'REJECTED_INSERTION_STATE_REQUIRED',
                message: 'Rejected-view insertion requires a wholly deleted paragraph owned by another author.'
            }
        };
    }
    const structuralRefusal = getParagraphRestorationRefusal(paragraph, { requireFollowingParagraph: false });
    if (structuralRefusal) {
        return { documentXml, hasChanges: false, status: 'error', error: structuralRefusal };
    }
    if (/\r|\n/.test(modified)) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: {
                code: 'UNSAFE_PARAGRAPH_BOUNDARY',
                message: 'Rejected-view insertion supports run-level text only; use restore for paragraph boundaries.'
            }
        };
    }

    const rejectedText = extractCanonicalParagraphText(paragraph, { revisionView: 'rejected' });
    if (!anchor.occurrenceExplicit && countNonOverlappingOccurrences(rejectedText, anchor.exactText) > 1) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: {
                code: 'AMBIGUOUS_ANCHOR',
                message: 'The rejected-view insertion anchor is repeated; provide anchor.occurrence explicitly.'
            }
        };
    }
    const anchorStart = findOccurrenceOffset(rejectedText, anchor.exactText, anchor.occurrence);
    if (anchorStart < 0) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: { code: 'ANCHOR_NOT_FOUND', message: 'The rejected-view insertion anchor was not found at the requested occurrence.' }
        };
    }
    const insertionOffset = anchorStart + anchor.offset;
    const carriers = Array.from(paragraph.childNodes || []).filter(
        node => node?.nodeType === 1 && node.namespaceURI === NS_W && node.localName === 'del'
    );
    if (carriers.some(carrier => hasUnsupportedDeletionSplitMarkup(paragraph, carrier))) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: {
                code: 'UNSAFE_REVISION_BOUNDARY',
                message: 'The rejected-view insertion boundary contains comments, bookmarks, fields, hyperlinks, moves, or non-text run markup that cannot be split safely.'
            }
        };
    }
    let carrierStart = 0;
    let targetCarrier = null;
    let carrierOffset = 0;
    for (const carrier of carriers) {
        const length = directDeletionCarrierText(carrier).length;
        if (insertionOffset >= carrierStart && insertionOffset <= carrierStart + length) {
            targetCarrier = carrier;
            carrierOffset = insertionOffset - carrierStart;
            break;
        }
        carrierStart += length;
    }
    if (!targetCarrier || carriers.map(directDeletionCarrierText).join('') !== rejectedText) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: {
                code: 'UNSAFE_REVISION_NESTING',
                message: 'The rejected-view anchor is not contained in a supported direct deletion carrier.'
            }
        };
    }

    const runProperties = insertedRunPropertiesAtDeletionOffset(targetCarrier, carrierOffset);
    const { leftCarrier, rightCarrier } = splitTrackChangeCarrier(
        xmlDoc,
        targetCarrier,
        carrierOffset,
        options._revisionIdAllocator || null
    );
    const insertion = createWordElement(xmlDoc, 'w:ins');
    const metadata = createRevisionMetadata(author, options._revisionIdAllocator || xmlDoc, 'ins');
    insertion.setAttribute('w:id', String(metadata.id));
    insertion.setAttribute('w:author', metadata.author);
    insertion.setAttribute('w:date', metadata.date);
    const run = createWordElement(xmlDoc, 'w:r');
    if (runProperties) run.appendChild(runProperties);
    const textNode = createWordElement(xmlDoc, 'w:t');
    if (/^\s|\s$/.test(modified)) textNode.setAttribute('xml:space', 'preserve');
    textNode.textContent = modified;
    run.appendChild(textNode);
    insertion.appendChild(run);

    const parent = targetCarrier.parentNode;
    if (leftCarrier) parent.insertBefore(leftCarrier, targetCarrier);
    parent.insertBefore(insertion, targetCarrier);
    if (rightCarrier) parent.insertBefore(rightCarrier, targetCarrier);
    options?._mutationRemovedNodes?.push(targetCarrier);
    parent.removeChild(targetCarrier);
    options?._mutationLiveNodes?.push(paragraph, insertion);
    operationSession?.invalidateParagraphMetadata?.();

    const outputXml = serializer.serializeToString(xmlDoc);
    const outputRejected = extractCanonicalParagraphText(paragraph, { revisionView: 'rejected' });
    const outputAccepted = extractCanonicalParagraphText(paragraph, { revisionView: 'accepted' });
    if (outputRejected !== rejectedText || !outputAccepted.includes(modified)) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: {
                code: 'PATCH_ROUNDTRIP_MISMATCH',
                message: 'Rejected-view insertion did not preserve the deleted source and expose the requested inserted text.'
            }
        };
    }
    return { documentXml: outputXml, hasChanges: true, status: 'ok' };
}

/**
 * Materializes an explicit counterproposal for one or more paragraphs wholly
 * deleted (content and paragraph mark) by another author.
 */
export async function restoreDeletedParagraphByExactText(
    documentXml,
    targetText,
    modified,
    author,
    targetRef = null,
    targetEndRef = null,
    runtimeContext = null,
    options = {}
) {
    const { serializer, xmlDoc, operationSession } = resolveMutationDocument(documentXml, options);
    if (!xmlDoc) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: { code: 'PARSE_ERROR', message: 'Could not parse document OOXML.' }
        };
    }

    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'restore', runtimeContext, options);
    if (resolved?.error || !resolved?.paragraph) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: resolved?.error || { code: 'TARGET_NOT_FOUND', message: 'Restoration target paragraph not found.' }
        };
    }

    const firstSource = resolved.paragraph;
    let sourceParagraphs = [firstSource];
    if (options._compiledSourceEndId) {
        const endResolved = resolveTargetParagraph(
            xmlDoc,
            options.targetEndDescriptor?.text || '',
            targetEndRef,
            'restore',
            runtimeContext,
            { ...options, _compiledSourceId: options._compiledSourceEndId }
        );
        const allParagraphs = endResolved?.paragraph ? getDocumentParagraphNodes(xmlDoc) : [];
        const startIndex = allParagraphs.indexOf(firstSource);
        const endIndex = allParagraphs.indexOf(endResolved?.paragraph);
        sourceParagraphs = startIndex >= 0 && endIndex >= startIndex
            ? allParagraphs.slice(startIndex, endIndex + 1)
            : null;
    } else if (targetEndRef) {
        sourceParagraphs = resolveParagraphRangeByRefs(xmlDoc, targetRef, targetEndRef, {
            opType: 'restore',
            targetRefSnapshot: runtimeContext?.targetRefSnapshot || null,
            onInfo: options.onInfo,
            onWarn: options.onWarn
        });
    } else if (options.targetEndDescriptor) {
        const endResolved = resolveTargetParagraph(
            xmlDoc,
            options.targetEndDescriptor.text,
            options.targetEndDescriptor.index,
            'restore',
            runtimeContext,
            {
                ...options,
                targetDescriptor: options.targetEndDescriptor,
                _compiledSourceId: options._compiledSourceEndId || null
            }
        );
        const allParagraphs = endResolved?.paragraph ? getDocumentParagraphNodes(xmlDoc) : [];
        const startIndex = allParagraphs.indexOf(firstSource);
        const endIndex = allParagraphs.indexOf(endResolved?.paragraph);
        sourceParagraphs = startIndex >= 0 && endIndex >= startIndex
            ? allParagraphs.slice(startIndex, endIndex + 1)
            : null;
    }
    if (!Array.isArray(sourceParagraphs) || sourceParagraphs.length === 0 || sourceParagraphs[0] !== firstSource) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: {
                code: 'RESTORATION_COUNT_MISMATCH',
                message: 'The restoration range could not be resolved as a contiguous paragraph block.'
            }
        };
    }

    const replacements = Array.isArray(modified) ? modified.map(String) : [String(modified || '')];
    if (replacements.length !== sourceParagraphs.length || replacements.some(text => text.length === 0)) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: {
                code: 'RESTORATION_COUNT_MISMATCH',
                message: `Restoration requires exactly one non-empty replacement per source paragraph (${sourceParagraphs.length} expected, ${replacements.length} received).`
            }
        };
    }

    const parent = firstSource.parentNode;
    if (!parent || sourceParagraphs.some(paragraph => paragraph.parentNode !== parent)) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: {
                code: 'UNSAFE_PARAGRAPH_PLACEMENT',
                message: 'Restoration source paragraphs must be a contiguous block in one structural container.'
            }
        };
    }

    for (const paragraph of sourceParagraphs) {
        const structuralRefusal = getParagraphRestorationRefusal(paragraph);
        if (structuralRefusal?.code === 'UNSUPPORTED_MOVE_REVISION') {
            return {
                documentXml,
                hasChanges: false,
                status: 'error',
                error: structuralRefusal
            };
        }
        const state = inspectForeignDeletedParagraphTarget(paragraph, author);
        if (!state.matches) {
            return {
                documentXml,
                hasChanges: false,
                status: 'error',
                error: {
                    code: 'RESTORATION_STATE_REQUIRED',
                    message: state.hasParagraphMarkDeletion && !state.foreignParagraphMarkDeletion
                        ? 'Explicit cross-author restoration does not apply to a paragraph deleted by the current author; use merge-same-author.'
                        : 'Explicit restoration requires a foreign paragraph-mark deletion with every pre-existing content node deleted.',
                    ...(state.ownerAuthor ? { ownerAuthor: state.ownerAuthor } : {})
                }
            };
        }
        const refusal = structuralRefusal || getParagraphRestorationRefusal(paragraph);
        if (refusal) {
            return {
                documentXml,
                hasChanges: false,
                status: 'error',
                error: refusal
            };
        }
    }

    const lastSource = sourceParagraphs[sourceParagraphs.length - 1];
    const existingBlock = followingParagraphBlock(lastSource, sourceParagraphs.length);
    const hasExistingRestoration = existingBlock.length === sourceParagraphs.length
        && existingBlock.every(paragraph => isInsertedParagraphByAuthor(paragraph, author));
    if (
        hasExistingRestoration
        && existingBlock.every((paragraph, index) => extractCanonicalParagraphText(paragraph) === replacements[index])
    ) {
        return {
            documentXml,
            hasChanges: false,
            status: 'no-op',
            warnings: ['An identical same-author paragraph restoration already exists.']
        };
    }

    const beforeXml = serializer.serializeToString(xmlDoc);
    const beforeParagraphs = getDocumentParagraphNodes(xmlDoc);
    const sourceStartIndex = beforeParagraphs.indexOf(firstSource);
    const existingIndexes = hasExistingRestoration
        ? existingBlock.map(paragraph => beforeParagraphs.indexOf(paragraph))
        : [];
    const templates = [];
    const warnings = [];
    for (let index = 0; index < sourceParagraphs.length; index++) {
        const source = sourceParagraphs[index];
        warnings.push(...collectRestorationAnchorWarnings(source));
        const template = buildRejectedRestorationTemplate(xmlDoc, source, serializer);
        const edited = template
            ? await editRestorationTemplate(xmlDoc, template, replacements[index], author, serializer)
            : null;
        if (!edited) {
            return {
                documentXml,
                hasChanges: false,
                status: 'error',
                error: {
                    code: 'PATCH_ROUNDTRIP_MISMATCH',
                    message: `Could not reconstruct restoration paragraph ${index + 1} from its rejected revision view.`
                }
            };
        }
        templates.push(edited);
    }

    const insertedParagraphs = templates.map(template => {
        const paragraph = xmlDoc.importNode(template, true);
        return trackRestoredParagraph(xmlDoc, paragraph, author, operationSession);
    });
    const insertionPoint = lastSource.nextSibling;
    for (const paragraph of insertedParagraphs) parent.insertBefore(paragraph, insertionPoint);
    if (hasExistingRestoration) {
        for (const paragraph of existingBlock) {
            options?._mutationRemovedNodes?.push(paragraph);
            paragraph.parentNode?.removeChild(paragraph);
        }
    }
    options?._mutationLiveNodes?.push(...insertedParagraphs);
    normalizeBodySectionOrder(xmlDoc);

    const outputXml = serializer.serializeToString(xmlDoc);
    const oracle = verifyParagraphRestorationLifecycle(
        beforeXml,
        outputXml,
        sourceStartIndex,
        existingIndexes,
        templates,
        insertedParagraphs
    );
    if (!oracle.valid) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: {
                code: oracle.code || 'PATCH_ROUNDTRIP_MISMATCH',
                message: oracle.message,
                stage: oracle.stage,
                ...(oracle.generatedIssues ? { generatedIssues: oracle.generatedIssues } : {}),
                ...(oracle.envelopeIssues ? { envelopeIssues: oracle.envelopeIssues } : {}),
                ...(oracle.expected ? { expected: oracle.expected } : {}),
                ...(oracle.actual ? { actual: oracle.actual } : {})
            }
        };
    }

    return {
        documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
        hasChanges: true,
        status: 'ok',
        numberingXml: null,
        ...(warnings.length > 0 ? { warnings: [...new Set(warnings)] } : {})
    };
}

async function tryExplicitDecimalHeaderListConversion({
    xmlDoc,
    serializer,
    documentXml,
    operationSession,
    targetParagraph,
    currentParagraphText,
    modifiedText,
    author,
    runtimeContext,
    revisionIdAllocator,
    generateRedlines = true,
    onInfo = () => { },
    options = {}
}) {
    if (!targetParagraph) return null;
    const scopedParagraphOxml = serializer.serializeToString(targetParagraph);
    const explicitPlan = buildSingleLineListStructuralFallbackPlan({
        oxml: scopedParagraphOxml,
        originalText: currentParagraphText,
        modifiedText,
        allowExistingList: false
    });
    if (
        !explicitPlan ||
        explicitPlan.numberingKey !== 'numbered:decimal:single' ||
        !Number.isInteger(explicitPlan.startAt) ||
        explicitPlan.startAt < 1
    ) {
        return null;
    }

    const strippedContent = stripSingleLineListMarkerPrefix(explicitPlan.listInput || modifiedText);
    if (!strippedContent) return null;

    onInfo('[List] Applying explicit numeric header conversion with direct list binding.');
    const numberingAction = resolveSingleLineListFallbackNumberingAction(
        explicitPlan,
        runtimeContext?.listFallbackSequenceState || null
    );

    const explicitStart = explicitPlan.startAt;
    const numberingState = runtimeContext?.numberingIdState || null;
    let appliedNumId = null;
    let numberingXml = null;

    if (numberingAction.type === 'explicitReuse' && numberingAction.numId) {
        appliedNumId = String(numberingAction.numId);
        onInfo(`[List] Reusing explicit-start list sequence (${numberingAction.numberingKey} -> numId ${appliedNumId}, next ${explicitStart + 1}).`);
    } else {
        const reservedPair = reserveNextNumberingIdPair(numberingState);
        if (!reservedPair) return null;

        appliedNumId = String(reservedPair.numId);
        numberingXml = buildExplicitDecimalMultilevelNumberingXml(
            reservedPair.numId,
            reservedPair.abstractNumId,
            explicitStart
        );

        if (numberingAction.type === 'explicitStartNew') {
            onInfo(`[List] Started explicit-start list sequence (${numberingAction.numberingKey} -> numId ${appliedNumId}).`);
        }
        onInfo(`[List] Using isolated explicit-start numbering (start ${explicitStart}, numId ${appliedNumId}, abstractNumId ${reservedPair.abstractNumId}).`);
    }

    // Redline only after numbering reservation succeeded so an aborted conversion never records receipt revisions.
    const redlineResult = await applyRedlineToOxml(
        serializer.serializeToString(targetParagraph),
        currentParagraphText,
        strippedContent,
        {
            author,
            generateRedlines,
            _revisionIdAllocator: revisionIdAllocator
        }
    );
    if (!redlineResult?.hasChanges || typeof redlineResult?.oxml !== 'string') return null;

    const extracted = extractReplacementNodes(redlineResult.oxml);
    const replacementNodes = extracted.replacementNodes;

    if (explicitPlan.numberingKey && runtimeContext?.listFallbackSharedNumIdByKey instanceof Map) {
        runtimeContext.listFallbackSharedNumIdByKey.delete(explicitPlan.numberingKey);
    }

    if (numberingAction.type === 'explicitStartNew' || numberingAction.type === 'explicitReuse') {
        recordSingleLineListFallbackExplicitSequence(
            runtimeContext?.listFallbackSequenceState || null,
            numberingAction.numberingKey || explicitPlan.numberingKey,
            appliedNumId,
            explicitStart
        );
    } else {
        clearSingleLineListFallbackExplicitSequence(
            runtimeContext?.listFallbackSequenceState || null,
            numberingAction.numberingKey || explicitPlan.numberingKey
        );
    }

    enforceListBindingOnParagraphNodes(replacementNodes, {
        numId: appliedNumId,
        ilvl: 0,
        clearParagraphPropertyChanges: true,
        removeListPropertyNode: true,
        trackPropertyChange: generateRedlines ? { author, revisionIdAllocator } : null
    });

    const parent = targetParagraph.parentNode;
    if (!parent) return null;
    for (const node of replacementNodes) {
        const imported = xmlDoc.importNode(node, true);
        ensureParagraphIdsOnImportedNode(imported, operationSession);
        const inserted = parent.insertBefore(imported, targetParagraph);
        options?._mutationLiveNodes?.push(inserted);
    }
    options?._mutationRemovedNodes?.push(targetParagraph);
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);
    if (operationSession?.receiptCollector && numberingXml) {
        const numIds = collectNumberingIdsFromNodes(options?._mutationLiveNodes);
        for (const id of numIds) {
            operationSession.receiptCollector.recordNumbering(id);
        }
    }
    return {
        documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
        hasChanges: true,
        numberingXml
    };
}

const FALLBACK_STYLE_TO_LIST_FORMAT = Object.freeze({
    upperAlpha: 'upperLetter',
    lowerAlpha: 'lowerLetter',
    upperRoman: 'upperRoman',
    lowerRoman: 'lowerRoman',
    bullet: 'bullet'
});

/**
 * Text-equal marker conversion ("A. Header" -> "A. Header") of a plain paragraph, done in place through the
 * list-format convert-plain core: the typed marker becomes a tracked deletion and the binding a tracked pPrChange
 * in the same paragraph. This replaces the whole-paragraph delete/insert for letters, roman numerals and bullets,
 * which lost w14:paraId, bookmarks and direct formatting and reviewed as one deleted and one inserted line. Shapes
 * the core refuses (marker inside a hyperlink, field or revision; changed marker punctuation) keep the structural
 * fallback below, which is still reject-exact.
 */
function tryInPlaceMarkerListConversion({
    xmlDoc,
    serializer,
    documentXml,
    operationSession,
    targetParagraph,
    currentParagraphText,
    modifiedText,
    author,
    runtimeContext,
    revisionIdAllocator,
    generateRedlines = true,
    onInfo = () => { },
    options = {}
}) {
    if (!targetParagraph || getParagraphListInfo(targetParagraph)) return null;
    if (normalizeWhitespaceForTargeting(currentParagraphText) !== normalizeWhitespaceForTargeting(modifiedText)) return null;
    const plan = buildSingleLineListStructuralFallbackPlan({
        oxml: serializer.serializeToString(targetParagraph),
        originalText: currentParagraphText,
        modifiedText,
        allowExistingList: false
    });
    const format = plan ? FALLBACK_STYLE_TO_LIST_FORMAT[plan.numberingStyle] : null;
    if (!format) return null;
    const allParagraphs = getDocumentParagraphNodes(xmlDoc);
    const conversion = convertPlainParagraphsToList(xmlDoc, [targetParagraph], normalizeListFormatRequest({
        mode: 'convert-plain',
        format,
        strip: 'marker'
    }), {
        author: author || getDefaultAuthor(),
        generateRedlines,
        revisionIdAllocator,
        runtimeContext,
        stylesXml: runtimeContext?.stylesXml || options.stylesXml || null,
        numberingXml: runtimeContext?.numberingXml || options.numberingXml || null,
        paragraphIndexOf: paragraph => allParagraphs.indexOf(paragraph) + 1
    });
    if (!conversion.ok) {
        onInfo(`[List] In-place marker conversion declined (${conversion.error.code}); using the structural fallback.`);
        return null;
    }
    onInfo(`[List] Converted typed ${format} marker in place (numId ${conversion.items[0].numId}).`);
    ensureParagraphIdsOnImportedNode(targetParagraph, operationSession);
    options?._mutationLiveNodes?.push(targetParagraph);
    for (const numId of conversion.allocatedNumIds) operationSession?.receiptCollector?.recordNumbering(numId);
    normalizeBodySectionOrder(xmlDoc);
    return {
        documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
        hasChanges: true,
        numberingXml: conversion.numberingXml
    };
}

// CT_PPrBase child order (ECMA-376), used to place copied direct properties on the new list paragraph.
const PPR_CHILD_ORDER = [
    'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers',
    'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct',
    'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing',
    'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl',
    'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'
];
const PPR_NOT_COPIED = new Set(['pStyle', 'numPr', 'rPr', 'sectPr', 'pPrChange']);

/**
 * The list structural fallback rebuilds the paragraph as a tracked-deleted source paragraph plus a new list
 * paragraph from text alone. The deleted paragraph must keep the source `pPr` (Reject All restores it, so the
 * Reject All verifier sees no property change), the new paragraph carries the source's direct properties that the
 * list binding does not own (`keepNext`, spacing, alignment...), and the paragraph that survives Accept All
 * keeps the source `w14:paraId` / `w14:textId`.
 */
function restoreSourceIdentityAndPropertiesOnFallback(xmlDoc, replacementNodes, targetParagraph) {
    const paragraphs = replacementNodes.filter(node => node?.nodeType === 1 && node.localName === 'p');
    const markOf = (paragraph, kind) => {
        const rPr = getDirectWordChild(getDirectWordChild(paragraph, 'pPr'), 'rPr');
        return !!rPr && !!getDirectWordChild(rPr, kind);
    };
    const deleted = paragraphs.find(paragraph => markOf(paragraph, 'del'));
    const created = paragraphs.find(paragraph => !markOf(paragraph, 'del'));
    const sourcePPr = getDirectWordChild(targetParagraph, 'pPr');
    if (deleted && sourcePPr) {
        const restored = xmlDoc.importNode(sourcePPr, true);
        const mark = getDirectWordChild(getDirectWordChild(deleted, 'pPr'), 'rPr');
        const delMarker = mark && getDirectWordChild(mark, 'del');
        if (delMarker) {
            let rPr = getDirectWordChild(restored, 'rPr');
            if (!rPr) {
                rPr = createWordElement(xmlDoc, 'w:rPr');
                const successor = getDirectWordChild(restored, 'sectPr') || getDirectWordChild(restored, 'pPrChange');
                restored.insertBefore(rPr, successor || null);
            }
            Array.from(rPr.childNodes).forEach(child => {
                if (child.nodeType === 1 && (child.localName === 'ins' || child.localName === 'del')) rPr.removeChild(child);
            });
            rPr.appendChild(xmlDoc.importNode(delMarker, true));
            deleted.replaceChild(restored, getDirectWordChild(deleted, 'pPr'));
        }
    }
    if (created && sourcePPr) {
        let createdPPr = getDirectWordChild(created, 'pPr');
        if (!createdPPr) {
            createdPPr = createWordElement(xmlDoc, 'w:pPr');
            created.insertBefore(createdPPr, created.firstChild || null);
        }
        for (const child of Array.from(sourcePPr.childNodes)) {
            if (child.nodeType !== 1 || PPR_NOT_COPIED.has(child.localName)) continue;
            if (getDirectWordChild(createdPPr, child.localName)) continue;
            const rank = PPR_CHILD_ORDER.indexOf(child.localName);
            const successor = Array.from(createdPPr.childNodes).find(node => (
                node.nodeType === 1 && PPR_CHILD_ORDER.indexOf(node.localName) > rank
            ));
            createdPPr.insertBefore(xmlDoc.importNode(child, true), successor || null);
        }
    }
    if (created) {
        for (const name of ['paraId', 'textId']) {
            const value = targetParagraph.getAttributeNS?.(NS_W14, name) || targetParagraph.getAttribute?.(`w14:${name}`);
            if (value) created.setAttributeNS(NS_W14, `w14:${name}`, value);
        }
    }
}

async function trySingleParagraphListStructuralFallback({
    xmlDoc,
    serializer,
    documentXml,
    operationSession,
    targetParagraph,
    currentParagraphText,
    modifiedText,
    author,
    runtimeContext,
    revisionIdAllocator,
    generateRedlines = true,
    onInfo = () => { },
    options = {}
}) {
    if (!targetParagraph) return null;

    const scopedParagraphOxml = serializer.serializeToString(targetParagraph);
    const fallbackPlan = buildSingleLineListStructuralFallbackPlan({
        oxml: scopedParagraphOxml,
        originalText: currentParagraphText,
        modifiedText,
        allowExistingList: false
    });
    if (!fallbackPlan) return null;

    onInfo('[List] No textual diff but list marker detected; forcing structural list conversion fallback.');
    const fallbackResult = await executeSingleLineListStructuralFallback(fallbackPlan, {
        author,
        generateRedlines,
        revisionIdAllocator,
        setAbstractStartOverride: false
    });
    if (!fallbackResult?.hasChanges || !fallbackResult?.oxml) {
        onInfo('[List] Structural list fallback produced no valid OOXML payload.');
        return null;
    }

    const extracted = extractReplacementNodes(fallbackResult.oxml);
    let replacementNodes = extracted.replacementNodes;
    let numberingXml = extracted.numberingXml || fallbackResult?.numberingXml || null;
    const hasPositiveStart = Number.isInteger(fallbackPlan?.startAt) && fallbackPlan.startAt > 0;
    // Only decimal starts use the explicit multilevel-decimal definition. Letter and roman starts (`C.`, `iv.`)
    // keep the generated definition, whose level-0 startOverride is already applied by the fallback executor.
    const hasExplicitStartAt = hasPositiveStart && fallbackPlan?.numberingStyle === 'decimal';
    const hasNonDecimalStart = hasPositiveStart && !hasExplicitStartAt;
    const numberingKey = fallbackResult?.listStructuralFallbackKey || fallbackPlan?.numberingKey || null;
    const numberingAction = resolveSingleLineListFallbackNumberingAction(
        fallbackPlan,
        runtimeContext?.listFallbackSequenceState || null
    );
    if (hasExplicitStartAt) {
        const explicitStart = fallbackPlan.startAt;
        let explicitNumIdForBinding = null;
        const numberingState = runtimeContext?.numberingIdState || null;
        if (numberingAction.type === 'explicitReuse' && numberingAction.numId) {
            explicitNumIdForBinding = String(numberingAction.numId);
            numberingXml = null;
            onInfo(`[List] Reusing explicit-start list sequence (${numberingAction.numberingKey} -> numId ${explicitNumIdForBinding}, next ${explicitStart + 1}).`);
        } else if (numberingState) {
            const reservedPair = reserveNextNumberingIdPair(numberingState);
            if (!reservedPair) return null;
            overwriteParagraphNumIds(replacementNodes, reservedPair.numId);
            explicitNumIdForBinding = String(reservedPair.numId);
            numberingXml = buildExplicitDecimalMultilevelNumberingXml(
                reservedPair.numId,
                reservedPair.abstractNumId,
                explicitStart
            );
            if (numberingAction.type === 'explicitStartNew') {
                onInfo(`[List] Started explicit-start list sequence (${numberingAction.numberingKey} -> numId ${reservedPair.numId}).`);
            }
            onInfo(`[List] Using isolated explicit-start numbering (start ${explicitStart}, numId ${reservedPair.numId}, abstractNumId ${reservedPair.abstractNumId}).`);
        } else {
            const generatedNumId = extractFirstParagraphNumId(replacementNodes);
            explicitNumIdForBinding = generatedNumId ? String(generatedNumId) : null;
            onInfo(`[List] Using isolated list numbering with explicit start ${explicitStart}${generatedNumId ? ` (numId ${generatedNumId})` : ''}.`);
        }

        if (numberingKey && runtimeContext?.listFallbackSharedNumIdByKey instanceof Map) {
            runtimeContext.listFallbackSharedNumIdByKey.delete(numberingKey);
        }
        if (numberingAction.type === 'explicitStartNew' || numberingAction.type === 'explicitReuse') {
            recordSingleLineListFallbackExplicitSequence(
                runtimeContext?.listFallbackSequenceState || null,
                numberingAction.numberingKey || numberingKey,
                explicitNumIdForBinding,
                explicitStart
            );
        } else {
            clearSingleLineListFallbackExplicitSequence(
                runtimeContext?.listFallbackSequenceState || null,
                numberingAction.numberingKey || numberingKey
            );
        }

        if (explicitNumIdForBinding) {
            enforceListBindingOnParagraphNodes(replacementNodes, {
                numId: explicitNumIdForBinding,
                ilvl: 0,
                clearParagraphPropertyChanges: true,
                removeListPropertyNode: true,
                trackPropertyChange: generateRedlines ? { author, revisionIdAllocator } : null
            });
        }
    } else {
        if (numberingXml && runtimeContext?.numberingIdState) {
            const normalizedNumbering = remapNumberingPayloadForDocument(numberingXml, replacementNodes, runtimeContext.numberingIdState);
            replacementNodes = normalizedNumbering.replacementNodes;
            numberingXml = normalizedNumbering.numberingXml;
        }
        clearSingleLineListFallbackExplicitSequence(
            runtimeContext?.listFallbackSequenceState || null,
            numberingAction.numberingKey || numberingKey
        );
    }

    if (!hasExplicitStartAt && numberingXml && runtimeContext) {
        // Separate operations of one batch that convert continuing markers (`A.` then `B.`) join one numbering
        // instance, through the same sequence state `list-format convert-plain` uses.
        const batch = continueBatchListSequence({
            numberingXml,
            replacementNodes,
            sequence: sequenceMapOf(runtimeContext),
            isNumIdLive: numId => numIdIsLive(xmlDoc, numId)
        });
        if (batch) {
            replacementNodes = batch.replacementNodes;
            numberingXml = null;
            onInfo(`[List] Continuing batch list numbering (numId ${batch.numId}).`);
        }
    }

    if (!hasExplicitStartAt && !hasNonDecimalStart && runtimeContext?.listFallbackSharedNumIdByKey instanceof Map) {
        const sharedNumId = numberingKey ? runtimeContext.listFallbackSharedNumIdByKey.get(numberingKey) : null;
        if (sharedNumId) {
            overwriteParagraphNumIds(replacementNodes, sharedNumId);
            numberingXml = null;
            onInfo(`[List] Reusing shared list numbering (${numberingKey} -> numId ${sharedNumId}).`);
        } else if (numberingKey) {
            const generatedNumId = extractFirstParagraphNumId(replacementNodes);
            if (generatedNumId) {
                runtimeContext.listFallbackSharedNumIdByKey.set(numberingKey, generatedNumId);
                onInfo(`[List] Captured shared list numbering (${numberingKey} -> numId ${generatedNumId}).`);
            }
        }
    }

    const parent = targetParagraph.parentNode;
    if (!parent) return null;
    if (generateRedlines) restoreSourceIdentityAndPropertiesOnFallback(xmlDoc, replacementNodes, targetParagraph);
    for (const node of replacementNodes) {
        const imported = xmlDoc.importNode(node, true);
        ensureParagraphIdsOnImportedNode(imported, operationSession);
        const inserted = parent.insertBefore(imported, targetParagraph);
        options?._mutationLiveNodes?.push(inserted);
    }
    options?._mutationRemovedNodes?.push(targetParagraph);
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);
    if (operationSession?.receiptCollector && numberingXml) {
        const numIds = collectNumberingIdsFromNodes(options?._mutationLiveNodes);
        for (const id of numIds) {
            operationSession.receiptCollector.recordNumbering(id);
        }
    }
    return {
        documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
        hasChanges: true,
        numberingXml
    };
}

function listDiagnosticFailure(documentXml, diagnostic) {
    const { severity: _severity, ...error } = diagnostic;
    return { documentXml, hasChanges: false, numberingXml: null, status: 'error', error };
}

function listDiagnosticFields(diagnostics) {
    return diagnostics.length > 0
        ? { diagnostics, warnings: diagnostics.map(diagnosticWarningText) }
        : {};
}

/**
 * Compares the markers of a list edit with the numbering definition the source paragraphs keep.
 * Returns a diagnostic (warning or error severity per policy) or null.
 */
function diagnoseListFormatNotApplied({ sourceParagraphs, modifiedText, runtimeContext, options }) {
    const policy = resolveListDiagnosticPolicy(options);
    if (policy === 'off') return null;
    const sourceNumberingXml = runtimeContext?.numberingIdState?.sourceNumberingXml;
    if (!sourceNumberingXml) return null;
    const sourceItems = [];
    for (const paragraph of sourceParagraphs || []) {
        const info = getParagraphListInfo(paragraph);
        if (info) sourceItems.push({ text: getParagraphText(paragraph), numId: info.numId, ilvl: info.ilvl });
    }
    const found = detectRequestedListFormatMismatch({ modifiedText, sourceItems, sourceNumberingXml });
    if (!found) return null;
    return buildListFormatDiagnostic(found.mismatch, {
        policy,
        numId: found.numId,
        level: found.level,
        marker: found.marker,
        targetText: found.targetText,
        operationIndex: options?._operationIndex
    });
}

export async function applyToParagraphByExactText(documentXml, targetText, modifiedText, author, targetRef = null, targetEndRef = null, runtimeContext = null, options = {}) {
    const generateRedlines = options.generateRedlines !== false;
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => { };
    const onWarn = typeof options?.onWarn === 'function' ? options.onWarn : () => { };
    const { serializer, xmlDoc, operationSession } = resolveMutationDocument(documentXml, options);
    if (!xmlDoc) return { documentXml, hasChanges: false, status: 'error', error: { code: 'PARSE_ERROR', message: 'Could not parse document OOXML.' } };
    const revisionIdAllocator = operationSession?.revisionIdAllocator || prepareRevisionAllocator(xmlDoc, options);
    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'redline', runtimeContext, {
        ...options,
        onInfo,
        onWarn
    });
    if (resolved?.error || !resolved?.paragraph) {
        return {
            documentXml,
            hasChanges: false,
            numberingXml: null,
            status: 'error',
            error: resolved?.error || { code: 'TARGET_NOT_FOUND', message: 'Target paragraph not found.' }
        };
    }
    const targetParagraph = resolved.paragraph;
    if (typeof modifiedText === 'string' && modifiedText.length > 0) {
        const resurrectionTarget = inspectForeignDeletedParagraphTarget(targetParagraph, author);
        if (resurrectionTarget.matches) {
            const ownerAuthor = resurrectionTarget.ownerAuthor || 'unattributed';
            return {
                documentXml,
                hasChanges: false,
                numberingXml: null,
                status: 'error',
                error: {
                    code: 'FOREIGN_PARAGRAPH_MARK_DELETION',
                    message: `Refusing to add visible text to a paragraph whose paragraph mark is deleted by another author (${ownerAuthor}). Use explicit paragraph restoration when supported.`,
                    ownerAuthor
                }
            };
        }
    }
    preprocessRedlineTargetParagraph(targetParagraph);
    const currentParagraphText = getParagraphText(targetParagraph);
    if (modifiedText === '') {
        const commentIds = getCommentIdsInElement(targetParagraph);
        if (commentIds.length > 0) {
            const comments = commentDetailsForIds(commentIds, options._existingCommentDetails);
            return {
                documentXml,
                hasChanges: false,
                numberingXml: null,
                status: 'error',
                error: {
                    code: 'COMMENTED_CONTENT_DELETE',
                    message: 'Refusing to delete a paragraph with existing comments. Resolve or explicitly remove the comments before deleting the paragraph.',
                    commentIds,
                    ...(comments.length > 0 ? { comments } : {})
                }
            };
        }
    }
    const boundaryCheck = validateParagraphBoundaryMutation(targetParagraph, modifiedText, options);
    if (!boundaryCheck.valid) {
        return {
            documentXml,
            hasChanges: false,
            numberingXml: null,
            status: 'error',
            error: {
                code: boundaryCheck.code,
                message: boundaryCheck.message
            }
        };
    }
    // 'auto' resolves against the target paragraph here and again against the whole
    // range once a targetEnd range is resolved below.
    const autoExistingRevisions = options.existingRevisions === 'auto';
    if (autoExistingRevisions) {
        options = { ...options, existingRevisions: resolveExistingRevisionsPolicy('auto', targetParagraph, author) };
    }
    const hasRevisions = containsTrackedChanges(targetParagraph);
    const existingPolicy = options.existingRevisions || 'merge-same-author';
    if (
        hasRevisions
        && existingPolicy !== 'accept-all-first'
        && existingPolicy !== 'accept-all-first-keep-normalized'
    ) {
        if (existingPolicy === 'merge-same-author' || existingPolicy === 'slice-cross-author') {
            const authors = getTrackedChangeAuthors(targetParagraph);
            const opAuthor = String(author || '').trim().toLowerCase();
            const allSame = authors.length > 0
                && !hasUnattributedTrackedChange(targetParagraph)
                && authors.every(a => a.trim().toLowerCase() === opAuthor);
            if (!allSame && existingPolicy === 'merge-same-author') {
                return {
                    documentXml,
                    hasChanges: false,
                    numberingXml: null,
                    status: 'error',
                    error: {
                        code: 'EXISTING_REVISIONS',
                        message: `Target paragraph contains tracked changes from another author (${[...authors, ...(hasUnattributedTrackedChange(targetParagraph) ? ['unattributed'] : [])].join(', ')}). Use existingRevisions: "slice-cross-author" for a surgical edit that preserves reviewer history; accepting or rejecting revisions requires separate authorization.`,
                        revisionAuthors: authors,
                        currentPolicy: existingPolicy
                    }
                };
            }
            const mergeCommentIds = allSame ? getCommentIdsInElement(targetParagraph) : [];
            if (mergeCommentIds.length > 0) {
                const comments = commentDetailsForIds(mergeCommentIds, options._existingCommentDetails);
                return {
                    documentXml,
                    hasChanges: false,
                    numberingXml: null,
                    status: 'error',
                    error: {
                        code: 'COMMENTED_CONTENT_MERGE',
                        message: 'Refusing to merge existing revisions in commented content because reverting the prior revisions could remove or orphan comment anchors.',
                        commentIds: mergeCommentIds,
                        ...(comments.length > 0 ? { comments } : {})
                    }
                };
            }
            if (!allSame) {
                const hasMove = targetParagraph.getElementsByTagNameNS(NS_W, 'moveFrom').length > 0
                    || targetParagraph.getElementsByTagNameNS(NS_W, 'moveTo').length > 0;
                if (hasMove) {
                    return {
                        documentXml,
                        hasChanges: false,
                        numberingXml: null,
                        status: 'error',
                        error: {
                            code: 'UNSAFE_REVISION_NESTING',
                            message: 'Cross-author slicing does not support pending move revisions.'
                        }
                    };
                }
            }
        } else {
            const hasDel = targetParagraph.getElementsByTagNameNS(NS_W, 'del').length > 0;
            const hasMove = targetParagraph.getElementsByTagNameNS(NS_W, 'moveFrom').length > 0
                || targetParagraph.getElementsByTagNameNS(NS_W, 'moveTo').length > 0;
            if (hasDel) {
                return {
                    documentXml,
                    hasChanges: false,
                    numberingXml: null,
                    status: 'error',
                    error: {
                        code: 'UNSAFE_REVISION_NESTING',
                        message: 'Refusing to replace content with pending deletions; nesting revisions is unsafe.'
                    }
                };
            }
            if (hasMove) {
                return {
                    documentXml,
                    hasChanges: false,
                    numberingXml: null,
                    status: 'error',
                    error: {
                        code: 'UNSAFE_REVISION_NESTING',
                        message: 'Refusing to mutate content with move revisions until move lifecycle is designed.'
                    }
                };
            }
            return {
                documentXml,
                hasChanges: false,
                numberingXml: null,
                status: 'error',
                error: {
                    code: 'EXISTING_REVISIONS',
                    message: 'Target paragraph contains tracked changes and existingRevisions is "reject-input".'
                }
            };
        }
    }
    if (
        modifiedText === ''
        && currentParagraphText === ''
        && !(hasRevisions && existingPolicy === 'merge-same-author')
    ) {
        if (generateRedlines) {
            markParagraphDeleted(targetParagraph, { author });
            options?._mutationLiveNodes?.push(targetParagraph);
        } else {
            options?._mutationRemovedNodes?.push(targetParagraph);
            targetParagraph.parentNode?.removeChild(targetParagraph);
        }
        normalizeBodySectionOrder(xmlDoc);
        return {
            documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
            hasChanges: true,
            numberingXml: null
        };
    }
    const containingTable = findContainingWordElement(targetParagraph, 'tbl');
    const rawTableStructuralCandidate = !!containingTable
        && !targetEndRef
        && typeof modifiedText === 'string'
        && modifiedText.includes('\n')
        && !isMarkdownTableText(modifiedText);
    const rawTableStructuralDedupeKey = rawTableStructuralCandidate
        ? computeTableStructuralDedupeKey(xmlDoc, containingTable, modifiedText)
        : null;
    const tableStructuralDedupes = runtimeContext?.tableStructuralRedlineKeys instanceof Set
        ? runtimeContext.tableStructuralRedlineKeys
        : null;
    if (rawTableStructuralDedupeKey && tableStructuralDedupes?.has(rawTableStructuralDedupeKey)) {
        onInfo('[Table] Skipping duplicate table-structural redline for the same table/payload in this turn.');
        return {
            documentXml,
            hasChanges: false,
            numberingXml: null,
            warnings: ['Skipped duplicate table-structural redline in the same turn.']
        };
    }
    const synthesizedTableMarkdown = containingTable
        ? synthesizeTableMarkdownFromMultilineCellEdit(targetParagraph, modifiedText, {
            tableElement: containingTable,
            currentParagraphText,
            onInfo,
            onWarn
        })
        : null;
    let effectiveModifiedText = synthesizedTableMarkdown || modifiedText;
    const useTableScope = !!containingTable && isMarkdownTableText(effectiveModifiedText);
    const isTableMarkdownEdit = isMarkdownTableText(effectiveModifiedText);
    // A targetEnd given as a compact handle (paragraphId + fingerprint, no numeric index) has no targetEndRef;
    // its range comes from the batch-compiled end or the end descriptor, never silently dropped.
    let compiledEndRange = null;
    if (!targetEndRef && (options._compiledSourceEndId || options.targetEndDescriptor)) {
        const rangeResolution = resolveOperationParagraphRange(xmlDoc, targetText, targetRef, targetEndRef, 'redline', runtimeContext, {
            ...options,
            _resolutionCapture: null
        });
        if (rangeResolution.error) {
            return { documentXml, hasChanges: false, status: 'error', error: rangeResolution.error };
        }
        if (rangeResolution.paragraphs.length > 1) {
            const container = rangeResolution.paragraphs[0].parentNode;
            if (rangeResolution.paragraphs.some(paragraph => paragraph.parentNode !== container)) {
                return {
                    documentXml,
                    hasChanges: false,
                    status: 'error',
                    error: { code: 'TARGET_RANGE_INVALID', message: 'A redline target range must stay inside one container (body, table cell, header, ...); it spans a table or other container boundary. Edit the paragraphs on each side separately.' }
                };
            }
            compiledEndRange = rangeResolution.paragraphs;
        }
    }
    const explicitRangeParagraphs = compiledEndRange ? compiledEndRange : targetEndRef
        ? resolveParagraphRangeByRefs(xmlDoc, targetRef, targetEndRef, {
            opType: 'redline',
            targetRefSnapshot: runtimeContext?.targetRefSnapshot || null,
            onInfo,
            onWarn
        })
        : (typeof targetText === 'string' && /\r?\n/.test(targetText)
            ? findReconstructionParagraphRange(xmlDoc, targetText)
            : null);
    const hasExplicitRangeScope = Array.isArray(explicitRangeParagraphs) && explicitRangeParagraphs.length > 0;
    if (autoExistingRevisions && hasExplicitRangeScope) {
        options = { ...options, existingRevisions: resolveExistingRevisionsPolicy('auto', explicitRangeParagraphs, author) };
    }
    if (!useTableScope && hasExplicitRangeScope) {
        const insertionEntries = buildExplicitRangeInsertionEntries(explicitRangeParagraphs, effectiveModifiedText);
        if (insertionEntries && insertionEntries.length > 0) {
            onInfo(`[List] Applying explicit-range insertion-only heuristic (${insertionEntries.length} new item(s)).`);
            for (const entry of insertionEntries) {
                onInfo(`[List] Explicit-range insertion: ilvl=${entry.ilvl}, markerType=${entry.markerType}, text="${String(entry.text || '').slice(0, 80)}${String(entry.text || '').length > 80 ? '…' : ''}"`);
            }
            const rangeFormatDiagnostic = diagnoseListFormatNotApplied({
                sourceParagraphs: explicitRangeParagraphs,
                modifiedText: effectiveModifiedText,
                runtimeContext,
                options
            });
            if (rangeFormatDiagnostic?.severity === 'error') return listDiagnosticFailure(documentXml, rangeFormatDiagnostic);
            const applied = applyExplicitRangeListInsertions({
                xmlDoc,
                explicitRangeParagraphs,
                insertionEntries,
                generateRedlines,
                author
            });
            if (applied) {
                return {
                    documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
                    hasChanges: true,
                    numberingXml: null,
                    ...listDiagnosticFields(rangeFormatDiagnostic ? [rangeFormatDiagnostic] : [])
                };
            }
        }
    }
    let inferredTableRangeParagraphs = null;
    if (!explicitRangeParagraphs && !useTableScope && isTableMarkdownEdit) {
        inferredTableRangeParagraphs = inferTableReplacementParagraphBlock(targetParagraph, {
            getParagraphText
        });
        if (inferredTableRangeParagraphs?.length > 1) {
            onInfo(`[Table] Heuristic range expansion selected ${inferredTableRangeParagraphs.length} paragraph(s) for replacement.`);
        }
    }
    const targetListInfo = getParagraphListInfo(targetParagraph);
    if (
        targetListInfo &&
        typeof effectiveModifiedText === 'string' &&
        !effectiveModifiedText.includes('\n')
    ) {
        const strippedListPrefix = stripRedundantLeadingListMarkers(effectiveModifiedText);
        if (strippedListPrefix && strippedListPrefix !== effectiveModifiedText.trim()) {
            onInfo('[List] Stripped redundant manual list marker prefix from single-line list item edit.');
            effectiveModifiedText = strippedListPrefix;
        }
    }
    if (useTableScope) {
        onInfo('[Table] Markdown table edit detected in table cell target; applying reconciliation at table scope.');
    }

    const bypassSingleParagraphHeuristics = options.explicitStructuredContent === true;

    const adjacencyInsertionCandidate = (
        !bypassSingleParagraphHeuristics
        && !useTableScope
        && !hasExplicitRangeScope
        && targetListInfo
    )
        ? deriveSingleParagraphListAdjacencyInsertion(currentParagraphText, effectiveModifiedText)
        : null;
    if (adjacencyInsertionCandidate) {
        onInfo(`[List] Applying single-paragraph list adjacency insertion heuristic (${adjacencyInsertionCandidate.position}).`);
        const parent = targetParagraph.parentNode;
        if (!parent) throw new Error('Target paragraph has no parent for adjacency list insertion');

        const listParagraph = buildInsertedListParagraph(
            xmlDoc,
            targetParagraph,
            {
                ilvl: targetListInfo.ilvl,
                numId: targetListInfo.numId,
                markerType: 'numbered',
                text: adjacencyInsertionCandidate.text
            },
            generateRedlines ? createRevisionMetadata(author, xmlDoc) : null,
            author,
            { generateRedlines }
        );

        const insertionPoint = adjacencyInsertionCandidate.position === 'before'
            ? targetParagraph
            : targetParagraph.nextSibling;
        ensureParagraphIdsOnImportedNode(listParagraph, operationSession);
        parent.insertBefore(listParagraph, insertionPoint);
        options?._mutationLiveNodes?.push(listParagraph);
        normalizeBodySectionOrder(xmlDoc);
        return {
            documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
            hasChanges: true,
            numberingXml: null
        };
    }

    // Retained paragraph followed by a Markdown table: keep the source paragraph in place (so its own
    // content, formatting and revisions are untouched), apply any requested inline formatting to it,
    // and insert the table after it instead of rewriting the paragraph as a deletion plus a copy.
    const tableAppendCandidate = (
        !useTableScope
        && !hasExplicitRangeScope
        && !targetListInfo
        && !containingTable
        && !targetEndRef
    )
        ? deriveParagraphThenTableAppend(effectiveModifiedText)
        : null;
    const tableAppendParagraph = tableAppendCandidate
        ? preprocessMarkdown(tableAppendCandidate.paragraphMarkdown)
        : null;
    if (
        tableAppendCandidate
        && normalizeWhitespaceForTargeting(tableAppendParagraph.cleanText) === normalizeWhitespaceForTargeting(currentParagraphText)
    ) {
        const tableData = parseMarkdownTable(tableAppendCandidate.tableText);
        if (tableData.headers.length > 0 || tableData.rows.length > 0) {
            onInfo('[Table] Applying retained-paragraph table append heuristic.');
            const parent = targetParagraph.parentNode;
            if (!parent) throw new Error('Target paragraph has no parent for table append');
            const followingNode = targetParagraph.nextSibling;
            if (tableAppendParagraph.formatHints.length > 0) {
                const formatted = await applyRedlineToOxml(
                    serializer.serializeToString(targetParagraph),
                    currentParagraphText || targetText,
                    tableAppendCandidate.paragraphMarkdown,
                    {
                        author,
                        generateRedlines,
                        existingRevisions: options.existingRevisions || 'merge-same-author',
                        structuredContent: false,
                        _revisionIdAllocator: revisionIdAllocator
                    }
                );
                if (formatted?.status === 'error') {
                    return {
                        documentXml,
                        hasChanges: false,
                        numberingXml: null,
                        status: 'error',
                        error: formatted.error
                    };
                }
                if (formatted?.hasChanges) {
                    if (typeof formatted.oxml !== 'string') {
                        throw new Error('Reconciliation engine did not return OOXML for a changed paragraph format');
                    }
                    for (const node of extractReplacementNodes(formatted.oxml).replacementNodes) {
                        const imported = xmlDoc.importNode(node, true);
                        ensureParagraphIdsOnImportedNode(imported, operationSession);
                        options?._mutationLiveNodes?.push(parent.insertBefore(imported, targetParagraph));
                    }
                    options?._mutationRemovedNodes?.push(targetParagraph);
                    parent.removeChild(targetParagraph);
                }
            }
            const tableOoxml = generateTableOoxml(tableData, {
                generateRedlines,
                author,
                revisionIdAllocator,
                trackAsBlock: true
            });
            for (const node of extractReplacementNodes(tableOoxml).replacementNodes) {
                options?._mutationLiveNodes?.push(parent.insertBefore(xmlDoc.importNode(node, true), followingNode));
            }
            // Word cannot end the body (or a container) with a table: it would add an untracked paragraph
            // there that Reject All leaves behind. Supply the paragraph ourselves with a tracked-inserted
            // mark so Reject All removes it and the source paragraph count is restored.
            // Skip whitespace/comment nodes of pretty-printed documents when deciding what follows the table.
            let followingElement = followingNode;
            while (followingElement && followingElement.nodeType !== 1) followingElement = followingElement.nextSibling;
            const endsContainer = !followingElement
                || (followingElement.namespaceURI === NS_W && followingElement.localName === 'sectPr');
            if (endsContainer) {
                const trailing = createWordElement(xmlDoc, 'w:p');
                const trailingPPr = clonePropertiesWithoutRevisionHistory(getDirectWordChild(targetParagraph, 'pPr'));
                if (trailingPPr) {
                    for (const child of Array.from(trailingPPr.childNodes)) {
                        if (child.nodeType === 1 && ['sectPr', 'numPr'].includes(child.localName)) trailingPPr.removeChild(child);
                    }
                    trailing.appendChild(trailingPPr);
                }
                if (generateRedlines) markParagraphInserted(trailing, { author });
                ensureParagraphIdsOnImportedNode(trailing, operationSession);
                options?._mutationLiveNodes?.push(parent.insertBefore(trailing, followingNode));
            }
            normalizeBodySectionOrder(xmlDoc);
            return {
                documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
                hasChanges: true,
                numberingXml: null
            };
        }
    }

    const plainAdjacencyInsertionCandidate = (
        !bypassSingleParagraphHeuristics
        && !useTableScope
        && !hasExplicitRangeScope
        && !targetListInfo
    )
        ? deriveSingleParagraphPlainAdjacencyInsertion(currentParagraphText, effectiveModifiedText)
        : null;
    if (plainAdjacencyInsertionCandidate) {
        onInfo(
            `[Text] Applying single-paragraph plain adjacency insertion heuristic `
            + `(${plainAdjacencyInsertionCandidate.position}, count=${plainAdjacencyInsertionCandidate.paragraphs.length}).`
        );
        const parent = targetParagraph.parentNode;
        if (!parent) throw new Error('Target paragraph has no parent for plain adjacency insertion');

        const insertionPoint = plainAdjacencyInsertionCandidate.position === 'before'
            ? targetParagraph
            : targetParagraph.nextSibling;

        for (const paragraphText of plainAdjacencyInsertionCandidate.paragraphs) {
            const plainParagraph = await buildInsertedPlainParagraph(
                xmlDoc,
                targetParagraph,
                paragraphText,
                generateRedlines ? createRevisionMetadata(author, xmlDoc) : null,
                author,
                // A wholly inserted paragraph must track its own paragraph mark; otherwise Word's
                // Reject All removes the text but leaves an empty paragraph behind.
                { generateRedlines, trackParagraphMark: true }
            );
            const imported = xmlDoc.importNode(plainParagraph, true);
            ensureParagraphIdsOnImportedNode(imported, operationSession);
            const inserted = parent.insertBefore(imported, insertionPoint);
            options?._mutationLiveNodes?.push(inserted);
        }

        normalizeBodySectionOrder(xmlDoc);
        return {
            documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
            hasChanges: true,
            numberingXml: null
        };
    }

    const insertionOnlyPlan = (!bypassSingleParagraphHeuristics && !useTableScope && !hasExplicitRangeScope)
        ? planListInsertionOnlyEdit(targetParagraph, effectiveModifiedText, {
            currentParagraphText,
            onInfo,
            onWarn
        })
        : null;
    if (insertionOnlyPlan && insertionOnlyPlan.entries.length > 0) {
        const insertionFormatDiagnostic = diagnoseListFormatNotApplied({
            sourceParagraphs: [targetParagraph],
            modifiedText: effectiveModifiedText,
            runtimeContext,
            options
        });
        if (insertionFormatDiagnostic?.severity === 'error') return listDiagnosticFailure(documentXml, insertionFormatDiagnostic);
        onInfo(`[List] Applying insertion-only list redline heuristic (${insertionOnlyPlan.entries.length} new item(s)).`);
        for (const entry of insertionOnlyPlan.entries) {
            onInfo(`[List] Insertion entry resolved: ilvl=${entry.ilvl}, markerType=${entry.markerType}, text="${String(entry.text || '').slice(0, 80)}${String(entry.text || '').length > 80 ? '…' : ''}"`);
        }
        const parent = targetParagraph.parentNode;
        if (!parent) throw new Error('Target paragraph has no parent for list insertion');
        const insertionPoint = targetParagraph.nextSibling;
        for (const entry of insertionOnlyPlan.entries) {
            const listParagraph = buildInsertedListParagraph(
                xmlDoc,
                targetParagraph,
                { ...entry, numId: insertionOnlyPlan.numId },
                generateRedlines ? createRevisionMetadata(author, xmlDoc) : null,
                author,
                { generateRedlines }
            );
            ensureParagraphIdsOnImportedNode(listParagraph, operationSession);
            parent.insertBefore(listParagraph, insertionPoint);
            options?._mutationLiveNodes?.push(listParagraph);
        }
        normalizeBodySectionOrder(xmlDoc);
        return {
            documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
            hasChanges: true,
            numberingXml: null,
            ...listDiagnosticFields(insertionFormatDiagnostic ? [insertionFormatDiagnostic] : [])
        };
    }

    const listScopeEdit = (!bypassSingleParagraphHeuristics && !useTableScope && !hasExplicitRangeScope)
        ? synthesizeExpandedListScopeEdit(targetParagraph, effectiveModifiedText, {
            currentParagraphText,
            onInfo,
            onWarn
        })
        : null;
    const useListScope = !!listScopeEdit && Array.isArray(listScopeEdit.paragraphs) && listScopeEdit.paragraphs.length > 0;
    if (useListScope) {
        effectiveModifiedText = listScopeEdit.modifiedText;
    }

    let unboundMarkerDiagnostic = null;
    if (!bypassSingleParagraphHeuristics && !useTableScope && !useListScope && !hasExplicitRangeScope) {
        const explicitHeaderListConversion = await tryExplicitDecimalHeaderListConversion({
            xmlDoc,
            serializer,
            documentXml,
            operationSession,
            targetParagraph,
            currentParagraphText,
            modifiedText: effectiveModifiedText,
            author,
            runtimeContext,
            revisionIdAllocator,
            generateRedlines,
            onInfo,
            options
        });
        if (explicitHeaderListConversion) return explicitHeaderListConversion;

        const inPlaceConversion = tryInPlaceMarkerListConversion({
            xmlDoc,
            serializer,
            documentXml,
            operationSession,
            targetParagraph,
            currentParagraphText,
            modifiedText: effectiveModifiedText,
            author,
            runtimeContext,
            revisionIdAllocator,
            generateRedlines,
            onInfo,
            options
        });
        if (inPlaceConversion) return inPlaceConversion;

        const listFallback = await trySingleParagraphListStructuralFallback({
            xmlDoc,
            serializer,
            documentXml,
            operationSession,
            targetParagraph,
            currentParagraphText,
            modifiedText: effectiveModifiedText,
            author,
            runtimeContext,
            revisionIdAllocator,
            generateRedlines,
            onInfo,
            options
        });
        if (listFallback) return listFallback;

        // Both structural fallbacks declined. A plain paragraph whose new text starts a list marker
        // would be written as literal text: report that instead of returning a silent success.
        const unboundPolicy = resolveListDiagnosticPolicy(options);
        if (unboundPolicy !== 'off' && !targetListInfo) {
            const requestedMarker = detectUnboundListMarker({
                originalText: currentParagraphText,
                modifiedText: effectiveModifiedText,
                targetIsListItem: false
            });
            if (requestedMarker) {
                unboundMarkerDiagnostic = buildUnboundListMarkerDiagnostic(requestedMarker, {
                    policy: unboundPolicy,
                    targetText: currentParagraphText,
                    operationIndex: options?._operationIndex
                });
                if (unboundPolicy === 'error') return listDiagnosticFailure(documentXml, unboundMarkerDiagnostic);
            }
        }
    }

    const originalTextForApply = useListScope
        ? listScopeEdit.originalText
        : (
            explicitRangeParagraphs
                ? explicitRangeParagraphs.map(paragraph => getParagraphText(paragraph)).join('\n')
                : (inferredTableRangeParagraphs
                    ? inferredTableRangeParagraphs.map(paragraph => getParagraphText(paragraph)).join('\n')
                    : (currentParagraphText || targetText))
        );
    const scopedXml = useTableScope
        ? serializer.serializeToString(containingTable)
        : (
            useListScope
                ? serializeParagraphRangeAsDocument(listScopeEdit.paragraphs, serializer)
                : (
                    explicitRangeParagraphs
                        ? serializeParagraphRangeAsDocument(explicitRangeParagraphs, serializer)
                        : (inferredTableRangeParagraphs
                            ? serializeParagraphRangeAsDocument(inferredTableRangeParagraphs, serializer)
                            : serializer.serializeToString(targetParagraph))
                )
        );

    const result = isTableMarkdownEdit
        ? await reconcileMarkdownTableOoxml(scopedXml, originalTextForApply, effectiveModifiedText, {
            author,
            generateRedlines,
            existingRevisions: options.existingRevisions || 'merge-same-author',
            structuredContent: options.structuredContent !== false,
            explicitStructuredContent: options.explicitStructuredContent === true,
            _revisionIdAllocator: revisionIdAllocator,
            _isolatedTableCell: useTableScope
        })
        : await applyRedlineToOxml(scopedXml, originalTextForApply, effectiveModifiedText, {
            author,
            generateRedlines,
            existingRevisions: options.existingRevisions || 'merge-same-author',
            structuredContent: options.structuredContent !== false,
            explicitStructuredContent: options.explicitStructuredContent === true,
            pairReplacements: options.pairReplacements === true,
            insertionAffinity: options.insertionAffinity || null,
            literalTextEdit: options.literalTextEdit === true,
            plainTextContent: options.plainTextContent === true,
            _revisionIdAllocator: revisionIdAllocator,
            _isolatedTableCell: useTableScope
        });
    if (!result?.hasChanges) {
        return {
            documentXml,
            hasChanges: false,
            numberingXml: null,
            status: result?.status || 'no-op',
            error: result?.error
        };
    }
    if (result.useNativeApi && !result.oxml) {
        const warning = 'Format-only fallback requires native Word API; browser demo skipped this operation.';
        onWarn(`[WARN] ${warning}`);
        return { documentXml, hasChanges: false, numberingXml: null, warnings: [warning] };
    }
    if (typeof result.oxml !== 'string') {
        throw new Error('Reconciliation engine did not return OOXML for a changed redline operation');
    }
    const extracted = extractReplacementNodes(result.oxml);
    let replacementNodes = removeListPackagingSentinel(extracted.replacementNodes, result.warnings);
    let numberingXml = extracted.numberingXml;
    if (numberingXml && runtimeContext?.numberingIdState) {
        const allocated = allocateNumberingForReplacement({
            numberingXml,
            replacementNodes,
            sourceParagraphs: useListScope ? listScopeEdit.paragraphs : (explicitRangeParagraphs || [targetParagraph]),
            numberingIdState: runtimeContext.numberingIdState,
            allowContinueSource: !useTableScope
        });
        replacementNodes = allocated.replacementNodes;
        numberingXml = allocated.numberingXml;
    }
    const scopeNodes = useTableScope
        ? [containingTable]
        : (
            useListScope
                ? listScopeEdit.paragraphs
                : (
                    explicitRangeParagraphs
                        ? explicitRangeParagraphs
                        : (inferredTableRangeParagraphs || [targetParagraph])
                )
        );
    const anchorNode = scopeNodes[0];
    const parent = anchorNode.parentNode;
    for (const node of replacementNodes) {
        const imported = xmlDoc.importNode(node, true);
        ensureParagraphIdsOnImportedNode(imported, operationSession);
        const inserted = parent.insertBefore(imported, anchorNode);
        options?._mutationLiveNodes?.push(inserted);
    }
    for (const scopeNode of scopeNodes) {
        if (scopeNode && scopeNode.parentNode === parent) {
            options?._mutationRemovedNodes?.push(scopeNode);
            parent.removeChild(scopeNode);
        }
    }
    normalizeBodySectionOrder(xmlDoc);
    if (rawTableStructuralDedupeKey && tableStructuralDedupes && (useTableScope || containingTable)) {
        tableStructuralDedupes.add(rawTableStructuralDedupeKey);
    }
    if (operationSession?.receiptCollector && numberingXml) {
        const numIds = collectNumberingIdsFromNodes(options?._mutationLiveNodes);
        for (const id of numIds) {
            operationSession.receiptCollector.recordNumbering(id);
        }
    }
    return {
        documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
        hasChanges: true,
        numberingXml,
        status: 'ok',
        ...(unboundMarkerDiagnostic ? {
            diagnostics: [unboundMarkerDiagnostic],
            warnings: [...(Array.isArray(result.warnings) ? result.warnings : []), diagnosticWarningText(unboundMarkerDiagnostic)]
        } : (Array.isArray(result.warnings) && result.warnings.length > 0 ? { warnings: result.warnings } : {}))
    };
}

export async function applyHighlightToParagraphByExactText(documentXml, targetText, textToHighlight, color, author, targetRef = null, runtimeContext = null, options = {}) {
    const generateRedlines = options.generateRedlines !== false;
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => { };
    const onWarn = typeof options?.onWarn === 'function' ? options.onWarn : () => { };
    const { serializer, xmlDoc, operationSession } = resolveMutationDocument(documentXml, options);
    if (!xmlDoc) return { documentXml, hasChanges: false, status: 'error', error: { code: 'PARSE_ERROR', message: 'Could not parse document OOXML.' } };
    const revisionIdAllocator = operationSession?.revisionIdAllocator || prepareRevisionAllocator(xmlDoc, options);
    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'highlight', runtimeContext, {
        ...options,
        onInfo,
        onWarn
    });
    if (resolved?.error || !resolved?.paragraph) {
        return { documentXml, hasChanges: false, status: 'error', error: resolved?.error || { code: 'TARGET_NOT_FOUND', message: 'Target paragraph not found.' } };
    }
    const targetParagraph = resolved.paragraph;
    const paragraphXml = serializer.serializeToString(targetParagraph);
    const highlightedXml = applyHighlightToOoxml(paragraphXml, textToHighlight, color, {
        generateRedlines,
        author,
        _revisionIdAllocator: revisionIdAllocator
    });
    if (!highlightedXml || highlightedXml === paragraphXml) return { documentXml, hasChanges: false };
    const { replacementNodes } = extractReplacementNodes(highlightedXml);
    const parent = targetParagraph.parentNode;
    for (const node of replacementNodes) {
        const imported = xmlDoc.importNode(node, true);
        ensureParagraphIdsOnImportedNode(imported, operationSession);
        const inserted = parent.insertBefore(imported, targetParagraph);
        options?._mutationLiveNodes?.push(inserted);
    }
    options?._mutationRemovedNodes?.push(targetParagraph);
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);
    return {
        documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
        hasChanges: true
    };
}

export async function applyCommentToParagraphByExactText(documentXml, targetText, textToComment, commentContent, author, targetRef = null, runtimeContext = null, options = {}) {
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => { };
    const onWarn = typeof options?.onWarn === 'function' ? options.onWarn : () => { };
    const { serializer, xmlDoc, operationSession } = resolveMutationDocument(documentXml, options);
    if (!xmlDoc) return { documentXml, hasChanges: false, commentsXml: null, status: 'error', error: { code: 'PARSE_ERROR', message: 'Could not parse document OOXML.' } };
    if (!operationSession) prepareRevisionAllocator(xmlDoc, options);
    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'comment', runtimeContext, {
        ...options,
        onInfo,
        onWarn
    });
    if (resolved?.error || !resolved?.paragraph) {
        return { documentXml, hasChanges: false, commentsXml: null, status: 'error', error: resolved?.error || { code: 'TARGET_NOT_FOUND', message: 'Target paragraph not found.' } };
    }
    const targetParagraph = resolved.paragraph;
    const paragraphXml = serializer.serializeToString(targetParagraph);
    const commentAnchor = typeof textToComment === 'string' && textToComment.length > 0
        ? textToComment
        : getParagraphText(targetParagraph);
    const commentResult = injectCommentsIntoOoxml(paragraphXml, [{
        paragraphIndex: 1,
        textToFind: commentAnchor,
        commentContent
    }], { author, commentIdAllocator: options.commentIdAllocator });
    if (commentResult.status === 'error' || commentResult.error || !commentResult.commentsApplied) {
        return {
            documentXml,
            hasChanges: false,
            commentsXml: null,
            status: 'error',
            error: commentResult.error || {
                code: 'ANCHOR_INSERTION_FAILED',
                message: 'The comment anchor was not applied to the resolved paragraph.'
            },
            warnings: commentResult.warnings || []
        };
    }
    const { replacementNodes } = extractReplacementNodes(commentResult.oxml);
    const parent = targetParagraph.parentNode;
    for (const node of replacementNodes) {
        const imported = xmlDoc.importNode(node, true);
        ensureParagraphIdsOnImportedNode(imported, operationSession);
        const inserted = parent.insertBefore(imported, targetParagraph);
        options?._mutationLiveNodes?.push(inserted);
    }
    options?._mutationRemovedNodes?.push(targetParagraph);
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);
    if (operationSession?.receiptCollector && Array.isArray(commentResult.placedComments)) {
        for (const c of commentResult.placedComments) {
            operationSession.receiptCollector.recordComment(c.id);
        }
    }
    return {
        documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
        hasChanges: true,
        commentsXml: commentResult.commentsXml || null,
        warnings: commentResult.warnings || [],
        resolvedAnchor: commentResult.resolvedAnchors?.[0] || null
    };
}

function splitRunAtTextOffset(xmlDoc, run, localOffset, revisionIdAllocator) {
    const run1 = run.cloneNode(true);
    const run2 = run.cloneNode(true);
    refreshRunPropertyChangeIds(run2, revisionIdAllocator);

    const isTextPiece = (node) => {
        if (node.nodeType !== 1) return false;
        const ln = node.localName || node.nodeName.replace(/^w:/, '');
        return ['t', 'tab', 'br', 'noBreakHyphen', 'delText'].includes(ln);
    };

    Array.from(run1.childNodes).forEach(c => { if (isTextPiece(c)) run1.removeChild(c); });
    Array.from(run2.childNodes).forEach(c => { if (isTextPiece(c)) run2.removeChild(c); });

    let fullText = '';
    for (const child of Array.from(run.childNodes)) {
        if (child.nodeType !== 1) continue;
        const ln = child.localName || child.nodeName.replace(/^w:/, '');
        if (ln === 't') fullText += child.textContent || '';
        else if (ln === 'tab') fullText += '\t';
        else if (ln === 'br') fullText += '\n';
        else if (ln === 'noBreakHyphen') fullText += '\u2011';
    }

    const text1 = fullText.slice(0, localOffset);
    const text2 = fullText.slice(localOffset);

    const appendPieces = (targetRun, text) => {
        const parts = text.split(/(\t|\n|\u2011)/);
        for (const part of parts) {
            if (!part) continue;
            if (part === '\t') {
                targetRun.appendChild(createWordElement(xmlDoc, 'w:tab'));
            } else if (part === '\n') {
                targetRun.appendChild(createWordElement(xmlDoc, 'w:br'));
            } else if (part === '\u2011') {
                targetRun.appendChild(createWordElement(xmlDoc, 'w:noBreakHyphen'));
            } else {
                const tEl = createWordElement(xmlDoc, 'w:t');
                if (/^\s|\s$/.test(part)) tEl.setAttribute('xml:space', 'preserve');
                tEl.textContent = part;
                targetRun.appendChild(tEl);
            }
        }
    };

    appendPieces(run1, text1);
    appendPieces(run2, text2);

    const parent = run.parentNode;
    parent.insertBefore(run1, run);
    parent.insertBefore(run2, run);
    parent.removeChild(run);

    return [run1, run2];
}

export async function applyFormattingToParagraphByExactText(
    documentXml,
    targetText,
    textToFormat,
    properties = {},
    author,
    targetRef = null,
    runtimeContext = null,
    options = {}
) {
    const generateRedlines = options.generateRedlines !== false;
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => { };
    const onWarn = typeof options?.onWarn === 'function' ? options.onWarn : () => { };
    const { serializer, xmlDoc, operationSession } = resolveMutationDocument(documentXml, options);
    if (!xmlDoc) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: { code: 'PARSE_ERROR', message: 'Could not parse document OOXML.' }
        };
    }
    const revisionIdAllocator = operationSession?.revisionIdAllocator || prepareRevisionAllocator(xmlDoc, options);
    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'format', runtimeContext, {
        ...options,
        onInfo,
        onWarn
    });
    if (resolved?.error || !resolved?.paragraph) {
        return { documentXml, hasChanges: false, status: 'error', error: resolved?.error || { code: 'TARGET_NOT_FOUND', message: 'Target paragraph not found.' } };
    }
    const targetParagraph = resolved.paragraph;

    const collectVisibleRuns = () => {
        const runEntries = [];
        const allRuns = Array.from(targetParagraph.getElementsByTagNameNS(NS_W, 'r'));
        for (const run of allRuns) {
            let parent = run.parentNode;
            let isDel = false;
            while (parent && parent !== targetParagraph) {
                if (parent.nodeType === 1 && (parent.localName === 'del' || parent.nodeName === 'w:del')) {
                    isDel = true;
                    break;
                }
                parent = parent.parentNode;
            }
            if (isDel) continue;

            let runText = '';
            for (const child of Array.from(run.childNodes)) {
                if (child.nodeType !== 1) continue;
                const ln = child.localName || child.nodeName.replace(/^w:/, '');
                if (ln === 't') runText += child.textContent || '';
                else if (ln === 'tab') runText += '\t';
                else if (ln === 'br') runText += '\n';
                else if (ln === 'noBreakHyphen') runText += '\u2011';
            }
            if (runText.length > 0) {
                runEntries.push({ run, text: runText });
            }
        }
        return runEntries;
    };

    let runEntries = collectVisibleRuns();
    let fullParaText = '';
    let spanRanges = [];
    for (const entry of runEntries) {
        const start = fullParaText.length;
        fullParaText += entry.text;
        const end = fullParaText.length;
        spanRanges.push({ ...entry, start, end });
    }

    const targetDescriptor = options.targetDescriptor;
    // `textOccurrence` is the dedicated in-paragraph selector. Legacy calls without it keep
    // reading `target.occurrence` (which also selects the N-th matching paragraph).
    const explicitTextOccurrence = Number.isInteger(options.textOccurrence) && options.textOccurrence > 0
        ? options.textOccurrence
        : null;
    const occurrence = explicitTextOccurrence ?? targetDescriptor?.occurrence ?? null;
    let matchStart = -1;
    if (occurrence != null && occurrence > 0) {
        let count = 0;
        let offset = 0;
        while (offset <= fullParaText.length - textToFormat.length) {
            const idx = fullParaText.indexOf(textToFormat, offset);
            if (idx === -1) break;
            count++;
            if (count === occurrence) {
                matchStart = idx;
                break;
            }
            offset = idx + 1;
        }
    } else {
        matchStart = fullParaText.indexOf(textToFormat);
    }

    if (matchStart === -1 && explicitTextOccurrence != null && fullParaText.includes(textToFormat)) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: {
                code: 'PATCH_SOURCE_NOT_FOUND',
                message: `Occurrence ${explicitTextOccurrence} of text to format was not found in target paragraph: "${textToFormat}".`
            }
        };
    }

    if (matchStart === -1) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: {
                code: 'TARGET_NOT_FOUND',
                message: `Text to format was not found in target paragraph: "${textToFormat}".`
            }
        };
    }
    const matchEnd = matchStart + textToFormat.length;

    // Split runs at matchStart boundary
    let currentRuns = [];
    for (const span of spanRanges) {
        if (span.start < matchStart && span.end > matchStart) {
            const localOffset = matchStart - span.start;
            const [r1, r2] = splitRunAtTextOffset(xmlDoc, span.run, localOffset, revisionIdAllocator);
            currentRuns.push({ run: r1, start: span.start, end: matchStart, text: span.text.slice(0, localOffset) });
            currentRuns.push({ run: r2, start: matchStart, end: span.end, text: span.text.slice(localOffset) });
        } else {
            currentRuns.push(span);
        }
    }

    // Split runs at matchEnd boundary
    const finalRuns = [];
    for (const span of currentRuns) {
        if (span.start < matchEnd && span.end > matchEnd) {
            const localOffset = matchEnd - span.start;
            const [r1, r2] = splitRunAtTextOffset(xmlDoc, span.run, localOffset, revisionIdAllocator);
            finalRuns.push({ run: r1, start: span.start, end: matchEnd, text: span.text.slice(0, localOffset) });
            finalRuns.push({ run: r2, start: matchEnd, end: span.end, text: span.text.slice(localOffset) });
        } else {
            finalRuns.push(span);
        }
    }

    const targetRuns = finalRuns.filter(span => span.start >= matchStart && span.end <= matchEnd);

    let anyChanged = false;
    for (const entry of targetRuns) {
        const rPr = entry.run.getElementsByTagNameNS(NS_W, 'rPr')[0] || null;
        if (checkRunPropertiesChanged(rPr, properties)) {
            anyChanged = true;
            break;
        }
    }

    if (!anyChanged) {
        return { documentXml, hasChanges: false, status: 'no-op' };
    }

    for (const entry of targetRuns) {
        const run = entry.run;
        let rPr = run.getElementsByTagNameNS(NS_W, 'rPr')[0];
        if (!rPr) {
            rPr = createWordElement(xmlDoc, 'w:rPr');
            run.insertBefore(rPr, run.firstChild);
        }

        let insAncestor = run.parentNode;
        let insideSameAuthorIns = false;
        while (insAncestor && insAncestor !== targetParagraph) {
            if (insAncestor.nodeType === 1 && (insAncestor.localName === 'ins' || insAncestor.nodeName === 'w:ins')) {
                const insAuthor = insAncestor.getAttribute('w:author') || insAncestor.getAttributeNS(NS_W, 'author');
                if (insAuthor === (author || getDefaultAuthor())) {
                    insideSameAuthorIns = true;
                }
                break;
            }
            insAncestor = insAncestor.parentNode;
        }

        const coalesce = insideSameAuthorIns && options.formattingRevisionPolicy === 'coalesce-own-insertion';

        if (generateRedlines && !coalesce) {
            snapshotAndAttachRPrChange(xmlDoc, rPr, author || getDefaultAuthor());
        }

        applyCharacterFormatToRPr(xmlDoc, rPr, properties);
    }

    normalizeBodySectionOrder(xmlDoc);
    ensureParagraphIdsOnImportedNode(targetParagraph, operationSession);
    options?._mutationLiveNodes?.push(targetParagraph);

    return {
        documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
        hasChanges: true,
        status: 'ok'
    };
}

export async function applyParagraphFormatToParagraphByExactText(
    documentXml,
    targetText,
    properties = {},
    author,
    targetRef = null,
    runtimeContext = null,
    options = {}
) {
    const generateRedlines = options.generateRedlines !== false;
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => { };
    const onWarn = typeof options?.onWarn === 'function' ? options.onWarn : () => { };
    const { serializer, xmlDoc, operationSession } = resolveMutationDocument(documentXml, options);
    if (!xmlDoc) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: { code: 'PARSE_ERROR', message: 'Could not parse document OOXML.' }
        };
    }
    if (!operationSession) prepareRevisionAllocator(xmlDoc, options);
    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'paragraph-format', runtimeContext, {
        ...options,
        onInfo,
        onWarn
    });
    if (resolved?.error || !resolved?.paragraph) {
        return { documentXml, hasChanges: false, status: 'error', error: resolved?.error || { code: 'TARGET_NOT_FOUND', message: 'Target paragraph not found.' } };
    }
    const targetParagraph = resolved.paragraph;

    let pPr = targetParagraph.getElementsByTagNameNS(NS_W, 'pPr')[0];
    if (!pPr) {
        pPr = createWordElement(xmlDoc, 'w:pPr');
        targetParagraph.insertBefore(pPr, targetParagraph.firstChild);
    }

    const unknownPropertyKeys = findUnknownParagraphFormatKeys(properties);
    const propertyDiagnostics = unknownPropertyKeys.length > 0
        ? [buildUnknownParagraphFormatDiagnostic(unknownPropertyKeys)]
        : [];
    if (properties?.numbering) {
        const binding = checkNumberingBinding(
            runtimeContext?.numberingIdState,
            properties.numbering.numId,
            Number(properties.numbering.level) || 0
        );
        if (!binding.ok) {
            return {
                documentXml,
                hasChanges: false,
                status: 'error',
                error: {
                    code: PARAGRAPH_NUMBERING_NOT_FOUND,
                    message: `${binding.message} paragraph-format.numbering binds to an existing list; it never creates one.`,
                    field: 'properties.numbering',
                    numId: Number(properties.numbering.numId),
                    level: Number(properties.numbering.level) || 0,
                    reason: binding.reason,
                    recovery: {
                        action: 'inspect_numbering',
                        hint: 'Inspect a paragraph already in the target list and pass its list.numId and list.level. To start a new list, use a list-change/redline operation.'
                    }
                }
            };
        }
    }
    const hasChanges = checkParagraphPropertiesChanged(pPr, properties);
    if (!hasChanges) {
        return { documentXml, hasChanges: false, status: 'no-op', ...listDiagnosticFields(propertyDiagnostics) };
    }

    if (generateRedlines) {
        const authorUsed = author || getDefaultAuthor();
        const pendingChange = getCurrentChild(pPr, 'pPrChange');
        const pendingAuthor = pendingChange?.getAttribute('w:author') ?? pendingChange?.getAttribute('author') ?? '';
        if (pendingChange && pendingAuthor !== authorUsed) {
            return {
                documentXml,
                hasChanges: false,
                status: 'error',
                error: {
                    code: PARAGRAPH_FORMAT_REVISION_CONFLICT,
                    message: `Target paragraph already has a pending paragraph-property change by "${pendingAuthor || 'unattributed'}". Overwriting its snapshot would change what Reject All restores.`,
                    pendingAuthor,
                    recovery: {
                        action: 'resolve_pending_revision',
                        hint: 'Ask the user to accept or reject the existing paragraph formatting change first; this tool does not resolve revisions by other reviewers.'
                    }
                }
            };
        }
        // A same-author pending snapshot already records the earliest "before" state; keep it.
        if (!pendingChange) snapshotAndAttachPPrChange(xmlDoc, pPr, authorUsed);
    }

    applyParagraphPropertiesToPPr(xmlDoc, pPr, properties);
    normalizeBodySectionOrder(xmlDoc);
    ensureParagraphIdsOnImportedNode(targetParagraph, operationSession);
    options?._mutationLiveNodes?.push(targetParagraph);

    return {
        documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
        hasChanges: true,
        status: 'ok',
        ...listDiagnosticFields(propertyDiagnostics)
    };
}

/**
 * Resolves an operation's target range (target plus optional targetEnd) to paragraphs in document order, the
 * same way restore does: batch-compiled source ids first, then legacy refs, then an end descriptor.
 */
function resolveOperationParagraphRange(xmlDoc, targetText, targetRef, targetEndRef, opType, runtimeContext, options) {
    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, opType, runtimeContext, options);
    if (resolved?.error || !resolved?.paragraph) {
        return { error: resolved?.error || { code: 'TARGET_NOT_FOUND', message: 'Target paragraph not found.' } };
    }
    const first = resolved.paragraph;
    const hasEnd = options._compiledSourceEndId || targetEndRef != null || options.targetEndDescriptor;
    if (!hasEnd) return { paragraphs: [first] };
    let paragraphs = null;
    if (options._compiledSourceEndId || options.targetEndDescriptor) {
        const endResolved = resolveTargetParagraph(
            xmlDoc,
            options.targetEndDescriptor?.text || '',
            options._compiledSourceEndId ? targetEndRef : (options.targetEndDescriptor?.index ?? targetEndRef),
            opType,
            runtimeContext,
            {
                ...options,
                ...(options.targetEndDescriptor ? { targetDescriptor: options.targetEndDescriptor } : {}),
                _compiledSourceId: options._compiledSourceEndId || null,
                _resolutionCapture: null
            }
        );
        if (endResolved?.error) return { error: endResolved.error };
        const all = endResolved?.paragraph ? getDocumentParagraphNodes(xmlDoc) : [];
        const startIndex = all.indexOf(first);
        const endIndex = all.indexOf(endResolved?.paragraph);
        paragraphs = startIndex >= 0 && endIndex >= startIndex ? all.slice(startIndex, endIndex + 1) : null;
    } else {
        paragraphs = resolveParagraphRangeByRefs(xmlDoc, targetRef, targetEndRef, {
            opType,
            targetRefSnapshot: runtimeContext?.targetRefSnapshot || null,
            onInfo: options.onInfo,
            onWarn: options.onWarn
        });
    }
    if (!Array.isArray(paragraphs) || paragraphs.length === 0 || paragraphs[0] !== first) {
        return { error: { code: 'TARGET_RANGE_INVALID', message: 'targetEnd must resolve to the target paragraph or a paragraph after it.' } };
    }
    return { paragraphs };
}

/**
 * `list-format` (`convert-plain`): binds plain paragraphs to a new list in place; see services/list-format-mutation.js.
 */
export async function applyListFormatToParagraphByExactText(
    documentXml,
    targetText,
    list,
    author,
    targetRef = null,
    targetEndRef = null,
    runtimeContext = null,
    options = {}
) {
    const generateRedlines = options.generateRedlines !== false;
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => { };
    const onWarn = typeof options?.onWarn === 'function' ? options.onWarn : () => { };
    const { serializer, xmlDoc, operationSession } = resolveMutationDocument(documentXml, options);
    if (!xmlDoc) {
        return { documentXml, hasChanges: false, status: 'error', error: { code: 'PARSE_ERROR', message: 'Could not parse document OOXML.' } };
    }
    const revisionIdAllocator = operationSession?.revisionIdAllocator || prepareRevisionAllocator(xmlDoc, options);
    const range = resolveOperationParagraphRange(xmlDoc, targetText, targetRef, targetEndRef, 'list-format', runtimeContext, {
        ...options,
        onInfo,
        onWarn
    });
    if (range.error) return { documentXml, hasChanges: false, status: 'error', error: range.error };
    const paragraphs = range.paragraphs;
    if (paragraphs.some(paragraph => paragraph.parentNode !== paragraphs[0].parentNode)) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: {
                code: 'TARGET_RANGE_INVALID',
                message: 'A list-format range must stay inside one container (the body, or one table cell).'
            }
        };
    }
    const allParagraphs = getDocumentParagraphNodes(xmlDoc);
    const request = normalizeListFormatRequest(list);
    if (request.mode === 'change-format' || request.mode === 'join') {
        return applyListRebinding(xmlDoc, paragraphs, request, {
            documentXml, serializer, operationSession, runtimeContext, options, author, generateRedlines, onInfo, onWarn,
            paragraphIndexOf: paragraph => {
                const index = allParagraphs.indexOf(paragraph);
                return index >= 0 ? index + 1 : null;
            }
        });
    }
    const conversion = convertPlainParagraphsToList(xmlDoc, paragraphs, request, {
        author: author || getDefaultAuthor(),
        generateRedlines,
        revisionIdAllocator,
        runtimeContext,
        stylesXml: runtimeContext?.stylesXml || options.stylesXml || null,
        numberingXml: runtimeContext?.numberingXml || options.numberingXml || null,
        paragraphIndexOf: paragraph => allParagraphs.indexOf(paragraph) + 1
    });
    if (!conversion.ok) return { documentXml, hasChanges: false, status: 'error', error: conversion.error };

    for (const item of conversion.items) {
        ensureParagraphIdsOnImportedNode(item.paragraph, operationSession);
        options?._mutationLiveNodes?.push(item.paragraph);
    }
    for (const numId of conversion.allocatedNumIds) operationSession?.receiptCollector?.recordNumbering(numId);
    normalizeBodySectionOrder(xmlDoc);
    onInfo(`[List] list-format convert-plain bound ${conversion.items.length} paragraph(s) as ${conversion.format} (numId ${[...new Set(conversion.items.map(item => item.numId))].join(', ')}).`);

    return {
        documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
        hasChanges: true,
        status: 'ok',
        numberingXml: conversion.numberingXml,
        list: {
            mode: 'convert-plain',
            format: conversion.format,
            level: conversion.level,
            strip: conversion.strip,
            items: conversion.items.map(item => ({
                paragraphIndex: item.paragraphIndex,
                paragraphId: getParagraphId(item.paragraph),
                marker: item.marker,
                value: item.value,
                numId: item.numId
            })),
            skipped: conversion.skipped.map(item => ({ paragraphIndex: item.paragraphIndex, reason: item.reason }))
        }
    };
}

/**
 * `list-format` `change-format` / `join`: rebinds existing list items (see services/list-format-change.js). The
 * definitions are read from the source numbering part plus the parts earlier operations of this batch allocated,
 * so a later operation can change a list an earlier one created or rebound.
 */
function applyListRebinding(xmlDoc, paragraphs, request, ctx) {
    const { documentXml, serializer, operationSession, runtimeContext, options, author, generateRedlines, onInfo, paragraphIndexOf } = ctx;
    const refused = error => ({ documentXml, hasChanges: false, status: 'error', error });
    const sourceNumberingXml = runtimeContext?.numberingIdState?.sourceNumberingXml
        || runtimeContext?.numberingXml
        || options.numberingXml
        || '';
    const pendingParts = Array.isArray(operationSession?.numberingXmlParts) ? operationSession.numberingXmlParts : [];
    const numberingXml = pendingParts.length > 0 ? mergeNumberingParts(sourceNumberingXml, pendingParts) : sourceNumberingXml;
    const shared = {
        author: author || getDefaultAuthor(),
        generateRedlines,
        runtimeContext,
        stylesXml: runtimeContext?.stylesXml || options.stylesXml || null,
        numberingXml,
        paragraphIndexOf
    };

    let outcome;
    if (request.mode === 'join') {
        const descriptor = options.continueFromDescriptor || null;
        const resolved = descriptor
            ? resolveTargetParagraph(xmlDoc, descriptor.text || '', descriptor.index ?? null, 'list-format', runtimeContext, {
                ...options,
                targetDescriptor: descriptor,
                _compiledSourceId: options._compiledContinueFromId || null,
                _compiledResolvedBy: null,
                _resolutionCapture: null
            })
            : null;
        if (!resolved?.paragraph) {
            return refused({
                code: resolved?.error?.code || 'TARGET_NOT_FOUND',
                message: `list.continueFrom: ${resolved?.error?.message || 'paragraph not found.'}`,
                field: 'list.continueFrom'
            });
        }
        outcome = joinListParagraphs(xmlDoc, paragraphs, resolved.paragraph, request, shared);
    } else {
        outcome = changeListParagraphsFormat(xmlDoc, paragraphs, request, shared);
    }
    if (!outcome.ok) return refused(outcome.error);
    if (outcome.unchanged) {
        const warning = `${LIST_FORMAT_UNCHANGED}: ${outcome.message}`;
        ctx.onWarn?.(warning);
        return { documentXml, hasChanges: false, status: 'ok', warnings: [warning] };
    }

    for (const item of outcome.items) {
        ensureParagraphIdsOnImportedNode(item.paragraph, operationSession);
        options?._mutationLiveNodes?.push(item.paragraph);
    }
    for (const numId of outcome.allocatedNumIds) operationSession?.receiptCollector?.recordNumbering(numId);
    normalizeBodySectionOrder(xmlDoc);
    const outside = outcome.items.filter(item => !item.inTarget).length;
    onInfo(`[List] list-format ${outcome.mode} (${outcome.scope}) rebound ${outcome.items.length} paragraph(s) from numId ${outcome.sourceNumId} to numId ${outcome.numId}${outside ? ` (${outside} outside the target range)` : ''}.`);

    return {
        documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
        hasChanges: true,
        status: 'ok',
        numberingXml: outcome.numberingXml,
        list: {
            mode: outcome.mode,
            scope: outcome.scope,
            sourceNumId: outcome.sourceNumId,
            numId: outcome.numId,
            ...(outcome.mode === 'change-format'
                ? { format: outcome.format, level: outcome.level, lvlText: outcome.lvlText, levelStarts: outcome.levelStarts }
                : { continueFrom: outcome.continueFrom }),
            items: outcome.items.map(item => ({
                paragraphIndex: item.paragraphIndex,
                paragraphId: getParagraphId(item.paragraph),
                level: item.level,
                numId: item.numId,
                sourceNumId: item.sourceNumId,
                inTarget: item.inTarget
            })),
            skipped: outcome.skipped.map(item => ({ paragraphIndex: item.paragraphIndex, reason: item.reason }))
        }
    };
}
