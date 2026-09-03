/**
 * Standalone document-operation runner for redline/highlight/comment operations.
 *
 * This module centralizes the browser-demo operation bridge so host layers can
 * stay focused on UI + prompt orchestration.
 */

import { createSerializer, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { findReconstructionParagraphRange } from '../engine/reconstruction-mapper.js';
import {
    RevisionIdAllocator,
    createRevisionMetadata,
    seedRevisionIdsFromDocument,
    setRevisionIdAllocatorForDocument
} from '../core/types.js';
import { createWordElement } from '../core/word-xml.js';
import { markParagraphMarkInserted } from '../engine/run-builders.js';
import {
    applyRedlineToOxml,
    reconcileMarkdownTableOoxml,
    applyHighlightToOoxml,
    injectCommentsIntoOoxml,
    getParagraphText as getParagraphTextFromOxml,
    isMarkdownTableText,
    findContainingWordElement,
    buildTargetReferenceSnapshot,
    resolveTargetParagraphWithSnapshot as resolveTargetParagraphWithSnapshotShared,
    buildSingleLineListStructuralFallbackPlan,
    executeSingleLineListStructuralFallback,
    resolveSingleLineListFallbackNumberingAction,
    recordSingleLineListFallbackExplicitSequence,
    clearSingleLineListFallbackExplicitSequence,
    enforceListBindingOnParagraphNodes,
    synthesizeTableMarkdownFromMultilineCellEdit,
    synthesizeExpandedListScopeEdit,
    planListInsertionOnlyEdit,
    getParagraphListInfo,
    stripRedundantLeadingListMarkers,
    stripSingleLineListMarkerPrefix,
    normalizeWhitespaceForTargeting,
    reserveNextNumberingIdPair,
    remapNumberingPayloadForDocument,
    overwriteParagraphNumIds,
    extractFirstParagraphNumId,
    buildExplicitDecimalMultilevelNumberingXml,
    inferTableReplacementParagraphBlock,
    resolveParagraphRangeByRefs,
    extractReplacementNodesFromOoxml,
    normalizeBodySectionOrderStandalone
} from '../index.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function prepareRevisionAllocator(xmlDoc, options = {}) {
    const allocator = options?._revisionIdAllocator instanceof RevisionIdAllocator
        ? options._revisionIdAllocator
        : new RevisionIdAllocator();
    seedRevisionIdsFromDocument(xmlDoc, allocator);
    setRevisionIdAllocatorForDocument(xmlDoc, allocator);
    return allocator;
}

function getParagraphText(paragraph) {
    return getParagraphTextFromOxml(paragraph);
}

function resolveTargetParagraph(xmlDoc, targetText, targetRef, opType, runtimeContext = null, options = {}) {
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => { };
    const onWarn = typeof options?.onWarn === 'function' ? options.onWarn : () => { };
    return resolveTargetParagraphWithSnapshotShared(xmlDoc, {
        targetText,
        targetRef,
        opType,
        targetRefSnapshot: runtimeContext?.targetRefSnapshot || null,
        onInfo,
        onWarn
    });
}

function extractReplacementNodes(outputOxml) {
    return extractReplacementNodesFromOoxml(outputOxml);
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

const LIST_LINE_REGEX = /^(\s*)((?:\d+(?:\.\d+)*\.?|\((?:\d+|[a-zA-Z]|[ivxlcIVXLC]+)\)|[a-zA-Z]\.|[ivxlcIVXLC]+\.|[-*+\u2022]))\s+(.*)$/;
const INLINE_LIST_MARKER_REGEX = /(?:^|\s)(?:\d+(?:\.\d+)*\.?|[A-Za-z]\.|[ivxlcIVXLC]+\.)\s+/g;

function parseOutlineLevelFromMarker(marker) {
    const normalized = String(marker || '').trim();
    if (!/^\d+(?:\.\d+)+\.?$/.test(normalized)) return null;
    const parts = normalized.replace(/\.$/, '').split('.');
    return Math.max(0, parts.length - 1);
}

function parseModifiedListLines(modifiedText) {
    const lines = String(modifiedText || '')
        .split(/\r?\n/g)
        .map(line => line.trimEnd())
        .filter(line => line.trim().length > 0);
    if (lines.length < 2) return null;

    const parsed = [];
    for (const line of lines) {
        const markerMatch = line.match(LIST_LINE_REGEX);
        if (!markerMatch) return null;
        const marker = markerMatch[2];
        const markerType = /^[-*+\u2022]$/.test(marker) ? 'bullet' : 'numbered';
        parsed.push({
            marker,
            markerType,
            level: Math.floor((markerMatch[1] || '').length / 2),
            outlineLevel: markerType === 'numbered' ? parseOutlineLevelFromMarker(marker) : null,
            text: stripRedundantLeadingListMarkers(markerMatch[3])
        });
    }
    return parsed.length >= 2 ? parsed : null;
}

function buildExplicitRangeInsertionEntries(explicitRangeParagraphs, modifiedText) {
    if (!Array.isArray(explicitRangeParagraphs) || explicitRangeParagraphs.length === 0) return null;
    const parsedLines = parseModifiedListLines(modifiedText);
    if (!parsedLines) return null;

    const originalTexts = explicitRangeParagraphs.map(paragraph =>
        normalizeWhitespaceForTargeting(getParagraphText(paragraph))
    );
    const modifiedTexts = parsedLines.map(item => normalizeWhitespaceForTargeting(item.text));
    if (originalTexts.some(text => !text) || modifiedTexts.some(text => !text)) return null;

    const listInfos = explicitRangeParagraphs.map(paragraph => getParagraphListInfo(paragraph));
    if (listInfos.some(info => !info || !info.numId)) return null;
    const baselineNumId = String(listInfos[0].numId);
    if (listInfos.some(info => String(info.numId) !== baselineNumId)) return null;

    const matchedPairs = [];
    let originalIndex = 0;
    for (let modifiedIndex = 0; modifiedIndex < modifiedTexts.length && originalIndex < originalTexts.length; modifiedIndex += 1) {
        if (modifiedTexts[modifiedIndex] === originalTexts[originalIndex]) {
            matchedPairs.push({ originalIndex, modifiedIndex });
            originalIndex += 1;
        }
    }
    if (originalIndex !== originalTexts.length) return null;

    const matchedModifiedIndexes = new Set(matchedPairs.map(pair => pair.modifiedIndex));
    const insertedIndexes = [];
    for (let idx = 0; idx < parsedLines.length; idx += 1) {
        if (!matchedModifiedIndexes.has(idx)) insertedIndexes.push(idx);
    }
    if (insertedIndexes.length === 0) return null;

    const baseIndentLevel = parsedLines[0]?.level || 0;
    return insertedIndexes.map(modifiedIndex => {
        const nextMatch = matchedPairs.find(pair => pair.modifiedIndex > modifiedIndex) || null;
        const prevMatch = [...matchedPairs].reverse().find(pair => pair.modifiedIndex < modifiedIndex) || null;
        const referenceMatch = nextMatch || prevMatch;
        if (!referenceMatch) return null;
        const referenceListInfo = listInfos[referenceMatch.originalIndex] || listInfos[0];
        if (!referenceListInfo) return null;

        const entry = parsedLines[modifiedIndex];
        const relativeLevel = Math.max(0, entry.level - baseIndentLevel);
        const explicitOutlineLevel = Number.isInteger(entry.outlineLevel) ? entry.outlineLevel : null;
        return {
            text: entry.text,
            markerType: entry.markerType,
            ilvl: explicitOutlineLevel != null
                ? explicitOutlineLevel
                : Math.max(0, (referenceListInfo.ilvl || 0) + relativeLevel),
            numId: String(referenceListInfo.numId),
            insertBeforeOriginalIndex: nextMatch ? nextMatch.originalIndex : null
        };
    }).filter(Boolean);
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

function countWords(text) {
    return String(text || '')
        .trim()
        .split(/\s+/)
        .filter(Boolean)
        .length;
}

function hasMultipleInlineListMarkers(text) {
    const source = String(text || '');
    if (!source) return false;

    let count = 0;
    const regex = new RegExp(INLINE_LIST_MARKER_REGEX.source, INLINE_LIST_MARKER_REGEX.flags);
    while (regex.exec(source)) {
        count += 1;
        if (count >= 2) return true;
    }
    return false;
}

function deriveSingleParagraphListAdjacencyInsertion(currentParagraphText, modifiedText) {
    const rawCurrent = String(currentParagraphText || '').trim();
    const rawModified = String(modifiedText || '').trim();
    if (!rawCurrent || !rawModified || rawModified === rawCurrent) return null;
    if (rawModified.includes('\n')) return null;
    const normalizedCurrent = normalizeWhitespaceForTargeting(rawCurrent);

    const minWords = 6;
    const sanitizeCandidate = text => stripRedundantLeadingListMarkers(String(text || '').trim()).trim();
    const buildCandidate = (position, text) => {
        const cleanedText = sanitizeCandidate(text);
        const cleanedNormalized = normalizeWhitespaceForTargeting(cleanedText);
        if (!cleanedText || cleanedText === rawCurrent) return null;
        if (countWords(cleanedText) < minWords) return null;
        if (hasMultipleInlineListMarkers(cleanedText)) return null;
        if (normalizedCurrent && cleanedNormalized.includes(normalizedCurrent)) return null;
        return { position, text: cleanedText };
    };

    if (rawModified.endsWith(rawCurrent)) {
        const prefix = rawModified.slice(0, rawModified.length - rawCurrent.length);
        const candidate = buildCandidate('before', prefix);
        if (candidate) return candidate;
    }

    if (rawModified.startsWith(rawCurrent)) {
        const suffix = rawModified.slice(rawCurrent.length);
        const candidate = buildCandidate('after', suffix);
        if (candidate) return candidate;
    }

    const normalizedModified = normalizeWhitespaceForTargeting(rawModified);
    if (!normalizedCurrent || normalizedCurrent === normalizedModified) return null;

    if (normalizedModified.endsWith(normalizedCurrent)) {
        const prefix = normalizedModified.slice(0, normalizedModified.length - normalizedCurrent.length);
        const candidate = buildCandidate('before', prefix);
        if (candidate) return candidate;
    }

    if (normalizedModified.startsWith(normalizedCurrent)) {
        const suffix = normalizedModified.slice(normalizedCurrent.length);
        const candidate = buildCandidate('after', suffix);
        if (candidate) return candidate;
    }

    return null;
}

function deriveSingleParagraphPlainAdjacencyInsertion(currentParagraphText, modifiedText) {
    const rawCurrent = String(currentParagraphText || '').trim();
    const rawModified = String(modifiedText || '');
    if (!rawCurrent || !rawModified || !rawModified.includes('\n')) return null;

    const lines = rawModified
        .split(/\r?\n/g)
        .map(line => String(line || '').trim())
        .filter(Boolean);
    if (lines.length < 2) return null;

    const normalize = value => normalizeWhitespaceForTargeting(String(value || ''));
    const normalizedCurrent = normalize(rawCurrent);
    const normalizedFirst = normalize(lines[0]);
    const normalizedLast = normalize(lines[lines.length - 1]);

    if (normalizedLast === normalizedCurrent) {
        const paragraphs = lines.slice(0, -1).map(line => String(line || '').trim()).filter(Boolean);
        if (paragraphs.length > 0) {
            return { position: 'before', paragraphs };
        }
    }

    if (normalizedFirst === normalizedCurrent) {
        const paragraphs = lines.slice(1).map(line => String(line || '').trim()).filter(Boolean);
        if (paragraphs.length > 0) {
            return { position: 'after', paragraphs };
        }
    }

    return null;
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

async function tryExplicitDecimalHeaderListConversion({
    xmlDoc,
    serializer,
    targetParagraph,
    currentParagraphText,
    modifiedText,
    author,
    runtimeContext,
    generateRedlines = true,
    onInfo = () => { }
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
            generateRedlines
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
    for (const node of replacementNodes) parent.insertBefore(xmlDoc.importNode(node, true), targetParagraph);
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);
    return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true, numberingXml };
}

async function trySingleParagraphListStructuralFallback({
    xmlDoc,
    serializer,
    targetParagraph,
    currentParagraphText,
    modifiedText,
    author,
    runtimeContext,
    generateRedlines = true,
    onInfo = () => { }
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
    for (const node of replacementNodes) parent.insertBefore(xmlDoc.importNode(node, true), targetParagraph);
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);
    return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true, numberingXml };
}

async function applyToParagraphByExactText(documentXml, targetText, modifiedText, author, targetRef = null, targetEndRef = null, runtimeContext = null, options = {}) {
    const generateRedlines = options.generateRedlines !== false;
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => { };
    const onWarn = typeof options?.onWarn === 'function' ? options.onWarn : () => { };
    const serializer = createSerializer();
    const xmlDoc = parseOoxmlSafe(documentXml, 'application/xml').doc;
    if (!xmlDoc) return { documentXml, hasChanges: false, status: 'error', error: { code: 'PARSE_ERROR', message: 'Could not parse document OOXML.' } };
    const revisionIdAllocator = prepareRevisionAllocator(xmlDoc, options);
    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'redline', runtimeContext, { onInfo, onWarn });
    const targetParagraph = resolved.paragraph;
    preprocessRedlineTargetParagraph(targetParagraph);
    const currentParagraphText = getParagraphText(targetParagraph);
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
                return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true, numberingXml: null };
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

    const adjacencyInsertionCandidate = (!useTableScope && !hasExplicitRangeScope && targetListInfo)
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
        parent.insertBefore(listParagraph, insertionPoint);
        normalizeBodySectionOrder(xmlDoc);
        return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true, numberingXml: null };
    }

    const plainAdjacencyInsertionCandidate = (!useTableScope && !hasExplicitRangeScope && !targetListInfo)
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
            parent.insertBefore(xmlDoc.importNode(plainParagraph, true), insertionPoint);
        }

        normalizeBodySectionOrder(xmlDoc);
        return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true, numberingXml: null };
    }

    const insertionOnlyPlan = (!useTableScope && !hasExplicitRangeScope)
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
            parent.insertBefore(listParagraph, insertionPoint);
        }
        normalizeBodySectionOrder(xmlDoc);
        return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true, numberingXml: null };
    }

    const listScopeEdit = (!useTableScope && !hasExplicitRangeScope)
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

    if (!useTableScope && !useListScope && !hasExplicitRangeScope) {
        const explicitHeaderListConversion = await tryExplicitDecimalHeaderListConversion({
            xmlDoc,
            serializer,
            targetParagraph,
            currentParagraphText,
            modifiedText: effectiveModifiedText,
            author,
            runtimeContext,
            generateRedlines,
            onInfo
        });
        if (explicitHeaderListConversion) return explicitHeaderListConversion;

        const listFallback = await trySingleParagraphListStructuralFallback({
            xmlDoc,
            serializer,
            targetParagraph,
            currentParagraphText,
            modifiedText: effectiveModifiedText,
            author,
            runtimeContext,
            generateRedlines,
            onInfo
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
            existingRevisions: options.existingRevisions,
            _revisionIdAllocator: revisionIdAllocator,
            _isolatedTableCell: useTableScope
        })
        : await applyRedlineToOxml(scopedXml, originalTextForApply, effectiveModifiedText, {
            author,
            generateRedlines,
            existingRevisions: options.existingRevisions,
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
    let replacementNodes = extracted.replacementNodes;
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
    for (const node of replacementNodes) parent.insertBefore(xmlDoc.importNode(node, true), anchorNode);
    for (const scopeNode of scopeNodes) {
        if (scopeNode && scopeNode.parentNode === parent) parent.removeChild(scopeNode);
    }
    normalizeBodySectionOrder(xmlDoc);
    if (rawTableStructuralDedupeKey && tableStructuralDedupes && (useTableScope || containingTable)) {
        tableStructuralDedupes.add(rawTableStructuralDedupeKey);
    }
    return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true, numberingXml, status: 'ok' };
}

async function applyHighlightToParagraphByExactText(documentXml, targetText, textToHighlight, color, author, targetRef = null, runtimeContext = null, options = {}) {
    const generateRedlines = options.generateRedlines !== false;
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => { };
    const onWarn = typeof options?.onWarn === 'function' ? options.onWarn : () => { };
    const serializer = createSerializer();
    const xmlDoc = parseOoxmlSafe(documentXml, 'application/xml').doc;
    if (!xmlDoc) return { documentXml, hasChanges: false, status: 'error', error: { code: 'PARSE_ERROR', message: 'Could not parse document OOXML.' } };
    const revisionIdAllocator = prepareRevisionAllocator(xmlDoc, options);
    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'highlight', runtimeContext, { onInfo, onWarn });
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
    for (const node of replacementNodes) parent.insertBefore(xmlDoc.importNode(node, true), targetParagraph);
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);
    return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true };
}

async function applyCommentToParagraphByExactText(documentXml, targetText, textToComment, commentContent, author, targetRef = null, runtimeContext = null, options = {}) {
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => { };
    const onWarn = typeof options?.onWarn === 'function' ? options.onWarn : () => { };
    const serializer = createSerializer();
    const xmlDoc = parseOoxmlSafe(documentXml, 'application/xml').doc;
    if (!xmlDoc) return { documentXml, hasChanges: false, commentsXml: null, status: 'error', error: { code: 'PARSE_ERROR', message: 'Could not parse document OOXML.' } };
    prepareRevisionAllocator(xmlDoc, options);
    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'comment', runtimeContext, { onInfo, onWarn });
    const targetParagraph = resolved.paragraph;
    const paragraphXml = serializer.serializeToString(targetParagraph);
    const commentResult = injectCommentsIntoOoxml(paragraphXml, [{ paragraphIndex: 1, textToFind: textToComment, commentContent }], { author });
    if (!commentResult.commentsApplied) return { documentXml, hasChanges: false, commentsXml: null, warnings: commentResult.warnings || [] };
    const { replacementNodes } = extractReplacementNodes(commentResult.oxml);
    const parent = targetParagraph.parentNode;
    for (const node of replacementNodes) parent.insertBefore(xmlDoc.importNode(node, true), targetParagraph);
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);
    return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true, commentsXml: commentResult.commentsXml || null, warnings: commentResult.warnings || [] };
}

function operationTargetPriority(op) {
    if (op?.type === 'comment') return 0;
    return 1;
}

/**
 * Returns a stable operation order that resolves anchor-based operations before
 * text edits can mutate their target text. Comments run first; all other
 * operation types retain their original relative order.
 *
 * @param {Object[]} operations - Structured document operations
 * @returns {Object[]} A reordered copy; input objects and input array are not mutated
 */
export function orderOperationsForStableTargets(operations = []) {
    return (Array.isArray(operations) ? operations : [])
        .map((operation, index) => ({ operation, index }))
        .sort((a, b) => operationTargetPriority(a.operation) - operationTargetPriority(b.operation) || a.index - b.index)
        .map(entry => entry.operation);
}

function mergeCommentsXml(existingXml, incomingXml) {
    if (!incomingXml) return existingXml || null;
    if (!existingXml) return incomingXml;

    const serializer = createSerializer();
    const existingDoc = parseOoxmlSafe(existingXml, 'application/xml').doc;
    const incomingDoc = parseOoxmlSafe(incomingXml, 'application/xml').doc;
    if (!existingDoc || !incomingDoc) return existingXml;
    const existingRoot = existingDoc.documentElement;
    const existingIds = new Set(
        Array.from(existingRoot.getElementsByTagNameNS(NS_W, 'comment'))
            .map(comment => comment.getAttribute('w:id') || comment.getAttribute('id'))
            .filter(Boolean)
    );

    for (const comment of Array.from(incomingDoc.getElementsByTagNameNS(NS_W, 'comment'))) {
        const id = comment.getAttribute('w:id') || comment.getAttribute('id');
        if (id && existingIds.has(id)) continue;
        existingRoot.appendChild(existingDoc.importNode(comment, true));
        if (id) existingIds.add(id);
    }
    return serializer.serializeToString(existingDoc);
}

function cloneBatchRuntimeContext(runtimeContext) {
    if (!runtimeContext || typeof runtimeContext !== 'object') return {};

    const context = { ...runtimeContext };
    if (runtimeContext.listFallbackSharedNumIdByKey instanceof Map) {
        context.listFallbackSharedNumIdByKey = new Map(runtimeContext.listFallbackSharedNumIdByKey);
    }
    if (runtimeContext.tableStructuralRedlineKeys instanceof Set) {
        context.tableStructuralRedlineKeys = new Set(runtimeContext.tableStructuralRedlineKeys);
    }
    if (runtimeContext.numberingIdState && typeof runtimeContext.numberingIdState === 'object') {
        context.numberingIdState = {
            ...runtimeContext.numberingIdState,
            usedNumIds: runtimeContext.numberingIdState.usedNumIds instanceof Set
                ? new Set(runtimeContext.numberingIdState.usedNumIds)
                : runtimeContext.numberingIdState.usedNumIds,
            usedAbstractNumIds: runtimeContext.numberingIdState.usedAbstractNumIds instanceof Set
                ? new Set(runtimeContext.numberingIdState.usedAbstractNumIds)
                : runtimeContext.numberingIdState.usedAbstractNumIds
        };
    }
    if (runtimeContext.listFallbackSequenceState && typeof runtimeContext.listFallbackSequenceState === 'object') {
        context.listFallbackSequenceState = {
            ...runtimeContext.listFallbackSequenceState,
            explicitByNumberingKey: runtimeContext.listFallbackSequenceState.explicitByNumberingKey instanceof Map
                ? new Map(runtimeContext.listFallbackSequenceState.explicitByNumberingKey)
                : runtimeContext.listFallbackSequenceState.explicitByNumberingKey
        };
    }
    return context;
}

function commitBatchRuntimeContext(runtimeContext, context) {
    if (!runtimeContext || typeof runtimeContext !== 'object') return;

    if (runtimeContext.listFallbackSharedNumIdByKey instanceof Map && context.listFallbackSharedNumIdByKey instanceof Map) {
        runtimeContext.listFallbackSharedNumIdByKey.clear();
        for (const entry of context.listFallbackSharedNumIdByKey) runtimeContext.listFallbackSharedNumIdByKey.set(...entry);
        context.listFallbackSharedNumIdByKey = runtimeContext.listFallbackSharedNumIdByKey;
    }
    if (runtimeContext.tableStructuralRedlineKeys instanceof Set && context.tableStructuralRedlineKeys instanceof Set) {
        runtimeContext.tableStructuralRedlineKeys.clear();
        for (const value of context.tableStructuralRedlineKeys) runtimeContext.tableStructuralRedlineKeys.add(value);
        context.tableStructuralRedlineKeys = runtimeContext.tableStructuralRedlineKeys;
    }
    if (runtimeContext.numberingIdState && context.numberingIdState) {
        for (const key of ['usedNumIds', 'usedAbstractNumIds']) {
            if (runtimeContext.numberingIdState[key] instanceof Set && context.numberingIdState[key] instanceof Set) {
                runtimeContext.numberingIdState[key].clear();
                for (const value of context.numberingIdState[key]) runtimeContext.numberingIdState[key].add(value);
                context.numberingIdState[key] = runtimeContext.numberingIdState[key];
            }
        }
        Object.assign(runtimeContext.numberingIdState, context.numberingIdState);
        context.numberingIdState = runtimeContext.numberingIdState;
    }
    if (runtimeContext.listFallbackSequenceState && context.listFallbackSequenceState) {
        const originalMap = runtimeContext.listFallbackSequenceState.explicitByNumberingKey;
        const updatedMap = context.listFallbackSequenceState.explicitByNumberingKey;
        if (originalMap instanceof Map && updatedMap instanceof Map) {
            originalMap.clear();
            for (const entry of updatedMap) originalMap.set(...entry);
            context.listFallbackSequenceState.explicitByNumberingKey = originalMap;
        }
        Object.assign(runtimeContext.listFallbackSequenceState, context.listFallbackSequenceState);
        context.listFallbackSequenceState = runtimeContext.listFallbackSequenceState;
    }
    Object.assign(runtimeContext, context);
}

function normalizeOperationError(error) {
    return {
        code: typeof error?.code === 'string' && error.code ? error.code : 'OPERATION_ERROR',
        message: error?.message || String(error)
    };
}

/**
 * Applies one structured operation (`redline`, `highlight`, or `comment`) to
 * a full `word/document.xml` payload.
 *
 * @param {string} documentXml
 * @param {Object} op
 * @param {string} author
 * @param {Object|null} [runtimeContext=null]
 * @param {{
 *   generateRedlines?: boolean,
 *   onInfo?: (message: string) => void,
 *   onWarn?: (message: string) => void
 * }} [options={}]
 * @returns {Promise<{ documentXml: string, hasChanges: boolean, numberingXml?: string|null, commentsXml?: string|null, warnings?: string[] }>}
 */
export async function applyOperationToDocumentXml(documentXml, op, author, runtimeContext = null, options = {}) {
    const parsed = parseOoxmlSafe(documentXml, 'application/xml');
    if (parsed.error || !parsed.doc) {
        return { documentXml, hasChanges: false, status: 'error', error: parsed.error, warnings: parsed.warnings };
    }
    const operationOptions = {
        ...options,
        _revisionIdAllocator: prepareRevisionAllocator(parsed.doc, options)
    };
    if (op?.type === 'highlight') {
        return applyHighlightToParagraphByExactText(
            documentXml,
            op.target,
            op.textToHighlight,
            op.color,
            author,
            op.targetRef,
            runtimeContext,
            operationOptions
        );
    }
    if (op?.type === 'comment') {
        return applyCommentToParagraphByExactText(
            documentXml,
            op.target,
            op.textToComment,
            op.commentContent,
            author,
            op.targetRef,
            runtimeContext,
            operationOptions
        );
    }
    return applyToParagraphByExactText(
        documentXml,
        op?.target,
        op?.modified,
        author,
        op?.targetRef,
        op?.targetEndRef,
        runtimeContext,
        operationOptions
    );
}

/**
 * Applies a batch of operations using stable target ordering. This prevents a
 * later comment from missing original text changed by an earlier replacement
 * in the same batch.
 *
 * Results retain each operation's original 1-based index even though execution
 * is reordered. Numbering payloads are returned as an array so package callers
 * can merge each one with `ensureNumberingArtifactsInZip`.
 *
 * @param {string} documentXml - Full `word/document.xml` payload
 * @param {Object[]} operations - Structured operations
 * @param {string} author - Revision/comment author
 * @param {Object|null} [runtimeContext=null] - Shared turn context
 * @param {{
 *   atomic?: boolean,
 *   continueOnError?: boolean,
 *   generateRedlines?: boolean,
 *   onInfo?: (message: string) => void,
 *   onWarn?: (message: string) => void
 * }} [options={}] - Runner options. `atomic` defaults to `true`, returning the
 * original document and no package artifacts if any operation fails.
 * `continueOnError` defaults to `true`, so all operations are attempted and
 * represented in `results`; `false` stops after the first operation error.
 * @returns {Promise<{
 *   documentXml: string,
 *   hasChanges: boolean,
 *   commentsXml: string|null,
 *   numberingXmlParts: string[],
 *   results: Array<{index:number,type:string,status:string,warnings?:string[],error?:Object}>,
 *   executionOrder: number[],
 *   rolledBack?: boolean
 * }>}
 */
export async function applyOperationsToDocumentXml(documentXml, operations, author, runtimeContext = null, options = {}) {
    const parsed = parseOoxmlSafe(documentXml, 'application/xml');
    if (parsed.error || !parsed.doc) {
        return {
            documentXml,
            hasChanges: false,
            commentsXml: null,
            numberingXmlParts: [],
            results: [],
            executionOrder: [],
            status: 'error',
            error: parsed.error,
            warnings: parsed.warnings
        };
    }
    const sourceOperations = Array.isArray(operations) ? operations : [];
    const scheduled = sourceOperations
        .map((operation, index) => ({ operation, index }))
        .sort((a, b) => operationTargetPriority(a.operation) - operationTargetPriority(b.operation) || a.index - b.index);

    const atomic = options.atomic !== false;
    const continueOnError = options.continueOnError !== false;
    const context = cloneBatchRuntimeContext(runtimeContext);
    if (!(context.targetRefSnapshot instanceof Map)) {
        context.targetRefSnapshot = buildTargetReferenceSnapshot(parsed.doc);
    }

    let currentDocumentXml = documentXml;
    let commentsXml = null;
    let hasChanges = false;
    const numberingXmlParts = [];
    const results = [];
    const executionOrder = [];
    let operationFailed = false;

    for (const entry of scheduled) {
        const { operation, index } = entry;
        executionOrder.push(index + 1);
        try {
            const result = await applyOperationToDocumentXml(
                currentDocumentXml,
                operation,
                author,
                context,
                options
            );
            currentDocumentXml = result.documentXml;
            hasChanges = hasChanges || result.hasChanges === true;
            commentsXml = mergeCommentsXml(commentsXml, result.commentsXml || null);
            if (result.numberingXml) numberingXmlParts.push(result.numberingXml);
            const isError = result.status === 'error' || !!result.error;
            operationFailed = operationFailed || isError;
            results.push({
                index: index + 1,
                type: operation?.type || 'redline',
                status: isError ? 'error' : (result.hasChanges ? 'applied' : 'no_change'),
                ...(Array.isArray(result.warnings) && result.warnings.length > 0 ? { warnings: result.warnings } : {}),
                ...(result.error ? { error: result.error } : {})
            });
            if (isError && !continueOnError) break;
        } catch (error) {
            const normalizedError = normalizeOperationError(error);
            operationFailed = true;
            results.push({
                index: index + 1,
                type: operation?.type || 'redline',
                status: 'error',
                warnings: [normalizedError.message],
                error: normalizedError
            });
            if (!continueOnError) break;
        }
    }

    results.sort((a, b) => a.index - b.index);
    const rolledBack = atomic && operationFailed;
    if (!rolledBack) commitBatchRuntimeContext(runtimeContext, context);

    return {
        documentXml: rolledBack ? documentXml : currentDocumentXml,
        hasChanges: rolledBack ? false : hasChanges,
        commentsXml: rolledBack ? null : commentsXml,
        numberingXmlParts: rolledBack ? [] : numberingXmlParts,
        results,
        executionOrder,
        ...(rolledBack ? {
            rolledBack: true,
            status: 'error',
            error: {
                code: 'BATCH_OPERATION_FAILED',
                message: 'Atomic batch rolled back because one or more operations failed.'
            }
        } : {})
    };
}
