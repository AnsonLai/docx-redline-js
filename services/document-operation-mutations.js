/**
 * OOXML mutation implementations used by the document operation applier.
 */

import { createSerializer, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { findReconstructionParagraphRange } from '../engine/reconstruction-mapper.js';
import { createRevisionMetadata } from '../core/types.js';
import { containsTrackedChanges, getTrackedChangeAuthors, createWordElement, withOoxmlSourceType } from '../core/word-xml.js';
import {
    markParagraphMarkInserted,
    markParagraphMarkDeleted,
    snapshotAndAttachRPrChange,
    snapshotAndAttachPPrChange
} from '../engine/run-builders.js';
import {
    applyCharacterFormatToRPr,
    checkRunPropertiesChanged,
    applyParagraphPropertiesToPPr,
    checkParagraphPropertiesChanged
} from '../engine/rpr-helpers.js';
import { refreshRunPropertyChangeIds } from '../core/revision-cloning.js';
import { getDefaultAuthor } from '../adapters/config.js';
import { applyRedlineToOxml as applyRedlineToOxmlEngine } from '../engine/oxml-engine.js';
import { applyHighlightToOoxml } from '../engine/formatting-removal.js';
import { parseTable as parseMarkdownTable } from '../pipeline/pipeline.js';
import { analyzeStructuredContent } from '../pipeline/structured-content.js';
import { injectCommentsIntoOoxml } from './comment-engine.js';
import {
    getParagraphText as getParagraphTextFromOxml,
    getParagraphId,
    createParagraphFingerprint,
    isMarkdownTableText,
    findContainingWordElement,
    resolveTargetParagraphWithSnapshot as resolveTargetParagraphWithSnapshotShared,
    resolveParagraphRangeByRefs,
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
    extractReplacementNodesFromOoxml,
    normalizeBodySectionOrderStandalone
} from './standalone-docx-plumbing.js';
import { prepareRevisionAllocator } from './document-operation-session.js';
import {
    buildExplicitRangeInsertionEntries,
    deriveSingleParagraphListAdjacencyInsertion,
    deriveSingleParagraphPlainAdjacencyInsertion
} from './operation-heuristics.js';
import { resolveTargetFromCapture, ensureParagraphIdsOnImportedNode } from './capture-engine.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

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

function resolveTargetParagraph(xmlDoc, targetText, targetRef, opType, runtimeContext = null, options = {}) {
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => { };
    const onWarn = typeof options?.onWarn === 'function' ? options.onWarn : () => { };
    const session = options?._documentOperationSession || null;
    const paragraphMetadataIndex = session?.getParagraphMetadataIndex?.() || null;

    if (options?.targetDescriptor?.captureRef) {
        try {
            const resolved = resolveTargetFromCapture(xmlDoc, session, options.targetDescriptor, opType, options);
            if (options?._resolutionCapture && resolved?.paragraph) {
                const paragraph = resolved.paragraph;
                const metadata = paragraphMetadataIndex?.byParagraph?.get(paragraph) || null;
                Object.assign(options._resolutionCapture, {
                    resolvedBy: resolved.resolvedBy,
                    resolvedTarget: {
                        index: metadata?.index ?? Array.from(xmlDoc.getElementsByTagNameNS(NS_W, 'p')).indexOf(paragraph) + 1,
                        paragraphId: metadata?.paragraphId ?? getParagraphId(paragraph),
                        text: metadata?.text ?? getParagraphText(paragraph),
                        fingerprint: metadata?.fingerprint ?? createParagraphFingerprint(paragraph),
                        inTable: metadata?.inTable ?? !!findContainingWordElement(paragraph, 'tbl')
                    }
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
            resolvedTarget: {
                index: metadata?.index ?? Array.from(xmlDoc.getElementsByTagNameNS(NS_W, 'p')).indexOf(paragraph) + 1,
                paragraphId: metadata?.paragraphId ?? getParagraphId(paragraph),
                text: metadata?.text ?? getParagraphText(paragraph),
                fingerprint: metadata?.fingerprint ?? createParagraphFingerprint(paragraph),
                inTable: metadata?.inTable ?? !!findContainingWordElement(paragraph, 'tbl')
            }
        });
    }
    return resolved;
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
        paragraph.appendChild(anchorPPr.cloneNode(true));
    }
    ensureListProperties(xmlDoc, paragraph, entry.ilvl, entry.numId);

    // A new list item is a whole inserted paragraph, not merely inserted text.
    // Tracking its paragraph mark lets Word (and our accept/reject helpers)
    // remove the list paragraph itself on Reject All instead of leaving an
    // empty bullet or number behind.
    if (generateRedlines) {
        markParagraphMarkInserted(xmlDoc, paragraph, author);
    }

    const run = createWordElement(xmlDoc, 'w:r');
    const anchorFirstRun = Array.from(anchorParagraph.getElementsByTagNameNS(NS_W, 'r'))[0] || null;
    const anchorRunPr = anchorFirstRun ? getDirectWordChild(anchorFirstRun, 'rPr') : null;
    if (anchorRunPr) {
        run.appendChild(anchorRunPr.cloneNode(true));
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

function buildEmptyParagraphTemplateFromAnchor(xmlDoc, anchorParagraph) {
    const paragraph = createWordElement(xmlDoc, 'w:p');
    const anchorPPr = getDirectWordChild(anchorParagraph, 'pPr');
    if (anchorPPr) paragraph.appendChild(anchorPPr.cloneNode(true));

    const run = createWordElement(xmlDoc, 'w:r');
    const anchorFirstRun = Array.from(anchorParagraph.getElementsByTagNameNS(NS_W, 'r'))[0] || null;
    const anchorRunPr = anchorFirstRun ? getDirectWordChild(anchorFirstRun, 'rPr') : null;
    if (anchorRunPr) run.appendChild(anchorRunPr.cloneNode(true));

    const textNode = createWordElement(xmlDoc, 'w:t');
    textNode.textContent = '';
    run.appendChild(textNode);
    paragraph.appendChild(run);
    return paragraph;
}

function wrapParagraphContentInInsertion(xmlDoc, paragraph, revisionMetadata, author) {
    const wrappedParagraph = createWordElement(xmlDoc, 'w:p');
    const pPr = getDirectWordChild(paragraph, 'pPr');
    if (pPr) wrappedParagraph.appendChild(pPr.cloneNode(true));

    const ins = createWordElement(xmlDoc, 'w:ins');
    const metadata = revisionMetadata || createRevisionMetadata(author, xmlDoc);
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
            { generateRedlines }
        );
    }

    if (!generateRedlines) {
        return sourceParagraph;
    }

    return wrapParagraphContentInInsertion(xmlDoc, sourceParagraph, revisionMetadata, author);
}

function collectNumberingIdsFromNodes(nodes) {
    const ids = new Set();
    for (const node of nodes || []) {
        const numIdNodes = Array.from(node?.getElementsByTagNameNS?.('*', 'numId') || []);
        for (const numIdNode of numIdNodes) {
            const val = numIdNode.getAttribute('w:val') || numIdNode.getAttribute('val');
            if (val != null && val !== '') {
                ids.add(String(val));
            }
        }
    }
    return Array.from(ids);
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
        removeListPropertyNode: true
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
    const hasExplicitStartAt = Number.isInteger(fallbackPlan?.startAt) && fallbackPlan.startAt > 0;
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
                removeListPropertyNode: true
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

    if (!hasExplicitStartAt && runtimeContext?.listFallbackSharedNumIdByKey instanceof Map) {
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
    const hasRevisions = containsTrackedChanges(targetParagraph);
    const existingPolicy = options.existingRevisions || 'merge-same-author';
    if (
        hasRevisions
        && existingPolicy !== 'accept-all-first'
        && existingPolicy !== 'accept-all-first-keep-normalized'
    ) {
        if (existingPolicy === 'merge-same-author') {
            const authors = getTrackedChangeAuthors(targetParagraph);
            const opAuthor = String(author || '').trim().toLowerCase();
            const allSame = authors.length > 0 && authors.every(a => a.trim().toLowerCase() === opAuthor);
            if (!allSame) {
                return {
                    documentXml,
                    hasChanges: false,
                    numberingXml: null,
                    status: 'error',
                    error: {
                        code: 'EXISTING_REVISIONS',
                        message: `Target paragraph contains tracked changes from another author (${authors.length ? authors.join(', ') : 'unattributed'}). Pass existingRevisions: "accept-all-first" or resolve revisions first.`
                    }
                };
            }
            const mergeCommentIds = getCommentIdsInElement(targetParagraph);
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
            markParagraphMarkDeleted(xmlDoc, targetParagraph, author, createRevisionMetadata(author, xmlDoc));
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
    const explicitRangeParagraphs = targetEndRef
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
    if (!useTableScope && hasExplicitRangeScope) {
        const insertionEntries = buildExplicitRangeInsertionEntries(explicitRangeParagraphs, effectiveModifiedText);
        if (insertionEntries && insertionEntries.length > 0) {
            onInfo(`[List] Applying explicit-range insertion-only heuristic (${insertionEntries.length} new item(s)).`);
            for (const entry of insertionEntries) {
                onInfo(`[List] Explicit-range insertion: ilvl=${entry.ilvl}, markerType=${entry.markerType}, text="${String(entry.text || '').slice(0, 80)}${String(entry.text || '').length > 80 ? '…' : ''}"`);
            }
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
                    numberingXml: null
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
                { generateRedlines }
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
            numberingXml: null
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
        const normalizedNumbering = remapNumberingPayloadForDocument(numberingXml, replacementNodes, runtimeContext.numberingIdState);
        replacementNodes = normalizedNumbering.replacementNodes;
        numberingXml = normalizedNumbering.numberingXml;
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
        ...(Array.isArray(result.warnings) && result.warnings.length > 0 ? { warnings: result.warnings } : {})
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
    const occurrence = targetDescriptor?.occurrence ?? null;
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
    const revisionIdAllocator = operationSession?.revisionIdAllocator || prepareRevisionAllocator(xmlDoc, options);
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

    const hasChanges = checkParagraphPropertiesChanged(pPr, properties);
    if (!hasChanges) {
        return { documentXml, hasChanges: false, status: 'no-op' };
    }

    if (generateRedlines) {
        snapshotAndAttachPPrChange(xmlDoc, pPr, author || getDefaultAuthor());
    }

    applyParagraphPropertiesToPPr(xmlDoc, pPr, properties);
    normalizeBodySectionOrder(xmlDoc);
    ensureParagraphIdsOnImportedNode(targetParagraph, operationSession);
    options?._mutationLiveNodes?.push(targetParagraph);

    return {
        documentXml: completedDocumentXml(xmlDoc, serializer, documentXml, operationSession),
        hasChanges: true,
        status: 'ok'
    };
}
