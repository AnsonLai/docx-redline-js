import {
    getParagraphText,
    normalizeWhitespaceForTargeting
} from '../core/paragraph-targeting.js';
import {
    getParagraphListInfo,
    stripRedundantLeadingListMarkers
} from '../core/list-targeting.js';

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

export function buildExplicitRangeInsertionEntries(explicitRangeParagraphs, modifiedText) {
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

function countWords(text) {
    return String(text || '').trim().split(/\s+/).filter(Boolean).length;
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

export function deriveSingleParagraphListAdjacencyInsertion(currentParagraphText, modifiedText) {
    const rawCurrent = String(currentParagraphText || '').trim();
    const rawModified = String(modifiedText || '').trim();
    if (!rawCurrent || !rawModified || rawModified === rawCurrent || rawModified.includes('\n')) return null;
    const normalizedCurrent = normalizeWhitespaceForTargeting(rawCurrent);
    const sanitizeCandidate = text => stripRedundantLeadingListMarkers(String(text || '').trim()).trim();
    const buildCandidate = (position, text) => {
        const cleanedText = sanitizeCandidate(text);
        const cleanedNormalized = normalizeWhitespaceForTargeting(cleanedText);
        if (!cleanedText || cleanedText === rawCurrent || countWords(cleanedText) < 6) return null;
        if (hasMultipleInlineListMarkers(cleanedText)) return null;
        if (normalizedCurrent && cleanedNormalized.includes(normalizedCurrent)) return null;
        return { position, text: cleanedText };
    };

    if (rawModified.endsWith(rawCurrent)) {
        const candidate = buildCandidate('before', rawModified.slice(0, -rawCurrent.length));
        if (candidate) return candidate;
    }
    if (rawModified.startsWith(rawCurrent)) {
        const candidate = buildCandidate('after', rawModified.slice(rawCurrent.length));
        if (candidate) return candidate;
    }

    const normalizedModified = normalizeWhitespaceForTargeting(rawModified);
    if (!normalizedCurrent || normalizedCurrent === normalizedModified) return null;
    if (normalizedModified.endsWith(normalizedCurrent)) {
        const candidate = buildCandidate('before', normalizedModified.slice(0, -normalizedCurrent.length));
        if (candidate) return candidate;
    }
    if (normalizedModified.startsWith(normalizedCurrent)) {
        const candidate = buildCandidate('after', normalizedModified.slice(normalizedCurrent.length));
        if (candidate) return candidate;
    }
    return null;
}

export function deriveSingleParagraphPlainAdjacencyInsertion(currentParagraphText, modifiedText) {
    const rawCurrent = String(currentParagraphText || '').trim();
    const rawModified = String(modifiedText || '');
    if (!rawCurrent || !rawModified || !rawModified.includes('\n')) return null;

    const lines = rawModified.split(/\r?\n/g).map(line => String(line || '').trim()).filter(Boolean);
    if (lines.length < 2) return null;
    const normalize = value => normalizeWhitespaceForTargeting(String(value || ''));
    const normalizedCurrent = normalize(rawCurrent);

    if (normalize(lines[lines.length - 1]) === normalizedCurrent) {
        const paragraphs = lines.slice(0, -1).filter(Boolean);
        if (paragraphs.length > 0) return { position: 'before', paragraphs };
    }
    if (normalize(lines[0]) === normalizedCurrent) {
        const paragraphs = lines.slice(1).filter(Boolean);
        if (paragraphs.length > 0) return { position: 'after', paragraphs };
    }
    return null;
}
