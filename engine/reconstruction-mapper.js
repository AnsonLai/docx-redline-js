/**
 * Reconstruction mapper.
 *
 * Builds paragraph/property/sentinel mappings and indexed lookups used by reconstruction writing.
 */

import { diff_match_patch } from 'diff-match-patch';

import { appendParagraphBoundary } from '../core/paragraph-offset-policy.js';
import { getDocumentParagraphs } from './format-extraction.js';
import { getElementsByTagNSOrTag, getFirstElementByTagNSOrTag } from '../core/xml-query.js';
import { NS_W } from '../core/types.js';
import { isWordElement } from '../core/word-xml.js';

const DMP = new diff_match_patch();

function localNameOf(node) {
    return String(node?.localName || node?.nodeName || '').replace(/^.*:/, '');
}

function wordAttribute(node, localName) {
    return node?.getAttributeNS?.(NS_W, localName)
        || node?.getAttribute?.(`w:${localName}`)
        || node?.getAttribute?.(localName)
        || '';
}

function createRangeCursorLookup(ranges) {
    let cursor = 0;
    return {
        at(index) {
            // Reconstruction normally walks source offsets forwards, but a
            // replacement insertion deliberately looks back to the beginning
            // of its deleted range to inherit that run's formatting. If the
            // deletion crossed a run/hyperlink boundary, the cursor has
            // already advanced past that range and must be rewound.
            if (cursor > 0 && (!ranges[cursor] || index < ranges[cursor].start)) {
                let low = 0;
                let high = cursor - 1;
                while (low <= high) {
                    const middle = Math.floor((low + high) / 2);
                    if (ranges[middle].end <= index) low = middle + 1;
                    else high = middle - 1;
                }
                cursor = low;
            }
            while (cursor < ranges.length && ranges[cursor].end <= index) {
                cursor++;
            }
            const match = ranges[cursor];
            if (!match) return null;
            if (match.start <= index && index < match.end) return match;
            return null;
        }
    };
}

function indexSentinelsByStart(sentinelMap) {
    const sentinelMapByStart = new Map();
    sentinelMap.forEach(sentinel => {
        if (!sentinelMapByStart.has(sentinel.start)) {
            sentinelMapByStart.set(sentinel.start, []);
        }
        sentinelMapByStart.get(sentinel.start).push(sentinel);
    });
    return sentinelMapByStart;
}

/**
 * Builds reconstruction mapping and cursor-based lookup helpers.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} modifiedText - Modified text
 * @returns {{
 *   paragraphs: Element[],
 *   body: Element|Document,
 *   paragraphMap: Array<{start:number,end:number,pPr:Element|null,container:Node}>,
 *   paragraphStarts: Set<number>,
 *   propertyMap: Array<{start:number,end:number,rPr:Element|null,wrapper?:Element}>,
 *   sentinelMap: Array<Object>,
 *   sentinelMapByStart: Map<number, Object[]>,
 *   referenceMap: Map<string, Node>,
 *   tokenToCharMap: Map<string, string>,
 *   containerFragments: Map<Node, DocumentFragment>,
 *   replacementContainers: Map<Node, Node>,
 *   originalFullText: string,
 *   processedModifiedText: string,
 *   getParagraphInfo: (index:number) => {start:number,end:number,pPr:Element|null,container:Node},
 *   getRunProperties: (index:number) => {rPr:Element|null,wrapper?:Element},
 *   getPropertySpanLength: (index:number,maxLength:number) => number,
 *   isParagraphStart: (index:number) => boolean
 * }}
 */
export function buildReconstructionMapping(xmlDoc, modifiedText, selectedParagraphs = null) {
    const rootElement = xmlDoc.documentElement;
    const isBodyRoot = isWordElement(rootElement, 'body') || localNameOf(rootElement) === 'package';
    const paragraphs = selectedParagraphs || getDocumentParagraphs(xmlDoc);

    let body = getFirstElementByTagNSOrTag(xmlDoc, NS_W, 'body');
    if (!body && isBodyRoot) body = rootElement;

    let originalFullText = '';
    const propertyMap = [];
    const paragraphMap = [];
    const sentinelMap = [];
    const referenceMap = new Map();
    const tokenToCharMap = new Map();
    const breakChars = new Set();
    const characterState = { nextCharCode: 0xe000 };
    const uniqueContainers = new Set();

    paragraphs.forEach((paragraph, paragraphIndex) => {
        const paragraphStart = originalFullText.length;

        Array.from(paragraph.childNodes).forEach(child => {
            originalFullText = processChildNode(
                child,
                originalFullText,
                propertyMap,
                sentinelMap,
                referenceMap,
                tokenToCharMap,
                characterState,
                breakChars
            );
        });

        originalFullText = appendParagraphBoundary(originalFullText, paragraphIndex, paragraphs.length);

        const paragraphEnd = originalFullText.length;
        const pPr = getFirstElementByTagNSOrTag(paragraph, NS_W, 'pPr');
        const container = paragraph.parentNode;
        if (container) uniqueContainers.add(container);

        paragraphMap.push({
            start: paragraphStart,
            end: paragraphEnd,
            pPr,
            container: container || body
        });
    });

    let displayOriginalText = '';
    for (let index = 0; index < originalFullText.length; index++) {
        const char = originalFullText[index];
        displayOriginalText += breakChars.has(char) ? '\n' : char;
    }
    let processedModifiedText = preserveZeroWidthSentinels(displayOriginalText, modifiedText, sentinelMap);
    processedModifiedText = preserveStructuralBreaks(displayOriginalText, originalFullText, processedModifiedText, breakChars);
    tokenToCharMap.forEach((char, tokenString) => {
        const escapedToken = tokenString.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
        processedModifiedText = processedModifiedText.replace(new RegExp(escapedToken, 'g'), char);
    });
    processedModifiedText = preserveReferencePlaceholders(originalFullText, processedModifiedText, referenceMap);

    const containerFragments = new Map();
    uniqueContainers.forEach(container => {
        containerFragments.set(container, xmlDoc.createDocumentFragment());
    });
    if (body && !containerFragments.has(body)) {
        containerFragments.set(body, xmlDoc.createDocumentFragment());
    }
    if (!containerFragments.has(xmlDoc)) {
        containerFragments.set(xmlDoc, xmlDoc.createDocumentFragment());
    }

    const replacementContainers = new Map();
    const paragraphStarts = new Set(paragraphMap.map(paragraph => paragraph.start));
    const paragraphLookup = createRangeCursorLookup(paragraphMap);
    const propertyLookup = createRangeCursorLookup(propertyMap);
    const sentinelMapByStart = indexSentinelsByStart(sentinelMap);

    const getParagraphInfo = (index) => {
        const match = paragraphLookup.at(index);
        if (match) return match;
        if (paragraphMap.length > 0) return paragraphMap[paragraphMap.length - 1];
        return { start: 0, end: 0, pPr: null, container: body || xmlDoc };
    };

    const getRunProperties = (index) => {
        const match = propertyLookup.at(index);
        return match ? { rPr: match.rPr, wrapper: match.wrapper } : { rPr: null };
    };

    const getPropertySpanLength = (index, maxLength) => {
        const match = propertyLookup.at(index);
        if (!match) return 1;
        return Math.min(match.end - index, maxLength);
    };

    return {
        paragraphs,
        body: body || xmlDoc,
        paragraphMap,
        paragraphStarts,
        propertyMap,
        sentinelMap,
        sentinelMapByStart,
        referenceMap,
        tokenToCharMap,
        containerFragments,
        replacementContainers,
        originalFullText,
        processedModifiedText,
        getParagraphInfo,
        getRunProperties,
        getPropertySpanLength,
        isParagraphStart: index => paragraphStarts.has(index)
    };
}

/**
 * Finds the contiguous paragraph range named by caller-provided original text.
 * Reconstruction replaces whole paragraphs, so returning null is safer than
 * silently rebuilding unrelated paragraphs around a partial match.
 */
export function findReconstructionParagraphRange(xmlDoc, originalText) {
    const paragraphs = getDocumentParagraphs(xmlDoc);
    if (paragraphs.length === 0) return [];

    const wanted = normalizeComparisonText(originalText);
    const paragraphTexts = paragraphs.map(extractParagraphVisibleText);
    if (!wanted) {
        const emptyIndex = paragraphTexts.findIndex(text => text === '');
        return emptyIndex >= 0 ? [paragraphs[emptyIndex]] : null;
    }
    if (paragraphs.length === 1 && paragraphTexts[0] === '') {
        return paragraphs;
    }

    const comparisons = [
        text => text,
        text => text.trim(),
        text => text.replace(/\s+/g, ' ').trim()
    ];
    for (const compare of comparisons) {
        const expected = compare(wanted);
        for (let start = 0; start < paragraphs.length; start++) {
            let combined = '';
            for (let end = start; end < paragraphs.length; end++) {
                combined += (end === start ? '' : '\n') + paragraphTexts[end];
                const candidate = compare(combined);
                if (candidate === expected) return paragraphs.slice(start, end + 1);
            }
        }
    }

    return null;
}

function normalizeComparisonText(text) {
    return String(text ?? '').replace(/\r\n?/g, '\n').replace(/\u00a0/g, ' ');
}

function extractParagraphVisibleText(paragraph) {
    let text = '';
    const visit = node => {
        for (const child of Array.from(node?.childNodes || [])) {
            if (child.nodeType !== 1) continue;
            if (isWordElement(child, 'pPr') || isWordElement(child, 'del') || isWordElement(child, 'moveFrom')) continue;
            if (isWordElement(child, 't')) text += child.textContent || '';
            else if (isWordElement(child, 'tab')) text += '\t';
            else if (isWordElement(child, 'br') || isWordElement(child, 'cr')) text += '\n';
            else if (isWordElement(child, 'noBreakHyphen')) text += '\u2011';
            else visit(child);
        }
    };
    visit(paragraph);
    return normalizeComparisonText(text);
}

function preserveStructuralBreaks(displayOriginalText, internalOriginalText, modifiedText, breakChars) {
    if (breakChars.size === 0) return modifiedText;

    const diffs = DMP.diff_main(displayOriginalText, modifiedText);
    let originalOffset = 0;
    let result = '';

    for (const [op, text] of diffs) {
        if (op === 0) {
            for (let index = 0; index < text.length; index++) {
                const internalChar = internalOriginalText[originalOffset + index];
                result += breakChars.has(internalChar) ? internalChar : text[index];
            }
            originalOffset += text.length;
        } else if (op === -1) {
            originalOffset += text.length;
        } else {
            result += text;
        }
    }

    return result;
}

function preserveZeroWidthSentinels(displayOriginalText, modifiedText, sentinelMap) {
    const sentinelsByInternalOffset = new Map();
    sentinelMap.forEach(sentinel => {
        if (sentinel.zeroWidth) sentinelsByInternalOffset.set(sentinel.start, sentinel);
    });
    if (sentinelsByInternalOffset.size === 0) return modifiedText;

    let visibleOriginalText = '';
    let visibleOffset = 0;
    const sentinelsByVisibleBoundary = new Map();

    for (let internalOffset = 0; internalOffset < displayOriginalText.length; internalOffset++) {
        const sentinel = sentinelsByInternalOffset.get(internalOffset);
        if (sentinel) {
            if (!sentinelsByVisibleBoundary.has(visibleOffset)) sentinelsByVisibleBoundary.set(visibleOffset, []);
            sentinelsByVisibleBoundary.get(visibleOffset).push({
                char: displayOriginalText[internalOffset],
                affinity: sentinel.affinity || 'right',
                emitted: false
            });
            continue;
        }
        visibleOriginalText += displayOriginalText[internalOffset];
        visibleOffset++;
    }

    const diffs = DMP.diff_main(visibleOriginalText, modifiedText);
    let originalOffset = 0;
    let result = '';

    const emitSentinels = (boundary, affinity) => {
        const sentinels = sentinelsByVisibleBoundary.get(boundary) || [];
        for (const sentinel of sentinels) {
            if (sentinel.emitted || (affinity && sentinel.affinity !== affinity)) continue;
            result += sentinel.char;
            sentinel.emitted = true;
        }
    };

    for (const [op, text] of diffs) {
        if (op === 1) {
            // Closing field markers belong before text inserted immediately
            // after a field; opening markers stay after text inserted before it.
            emitSentinels(originalOffset, 'left');
            result += text;
            continue;
        }

        for (let index = 0; index < text.length; index++) {
            emitSentinels(originalOffset);
            if (op === 0) result += text[index];
            originalOffset++;
        }
    }

    emitSentinels(originalOffset);
    return result;
}

function preserveReferencePlaceholders(originalFullText, modifiedText, referenceMap) {
    let result = modifiedText;

    for (const referenceChar of referenceMap.keys()) {
        if (result.includes(referenceChar)) continue;

        const originalIndex = originalFullText.indexOf(referenceChar);
        if (originalIndex < 0) continue;

        const prefix = originalFullText.slice(0, originalIndex);
        const suffix = originalFullText.slice(originalIndex + referenceChar.length);

        if (prefix && result.startsWith(prefix)) {
            result = `${result.slice(0, prefix.length)}${referenceChar}${result.slice(prefix.length)}`;
            continue;
        }

        if (suffix && result.endsWith(suffix)) {
            const insertAt = result.length - suffix.length;
            result = `${result.slice(0, insertAt)}${referenceChar}${result.slice(insertAt)}`;
        }
    }

    return result;
}

function processChildNode(child, originalFullText, propertyMap, sentinelMap, referenceMap, tokenToCharMap, characterState, breakChars) {
    if (isWordElement(child, 'r')) {
        return processRunForReconstruction(child, originalFullText, propertyMap, sentinelMap, referenceMap, tokenToCharMap, characterState, breakChars);
    }
    if (isWordElement(child, 'hyperlink')) {
        return processHyperlinkForReconstruction(child, originalFullText, propertyMap);
    }
    if (
        isWordElement(child, 'sdt')
        || isWordElement(child, 'oMath')
        || localNameOf(child) === 'oMath'
        || isWordElement(child, 'bookmarkStart')
        || isWordElement(child, 'bookmarkEnd')
    ) {
        sentinelMap.push({ start: originalFullText.length, node: child });
        return originalFullText + '\uFFFC';
    }
    if (isWordElement(child, 'commentRangeStart') || isWordElement(child, 'commentRangeEnd')) {
        sentinelMap.push({ start: originalFullText.length, node: child, isCommentMarker: true });
        return originalFullText;
    }
    return originalFullText;
}

function processRunForReconstruction(runElement, originalFullText, propertyMap, sentinelMap, referenceMap, tokenToCharMap, characterState, breakChars) {
    let fullText = originalFullText;
    const rPr = getFirstElementByTagNSOrTag(runElement, NS_W, 'rPr');

    Array.from(runElement.childNodes).forEach(runChild => {
        if (isWordElement(runChild, 't')) {
            const textContent = runChild.textContent || '';
            if (textContent.length > 0) {
                propertyMap.push({
                    start: fullText.length,
                    end: fullText.length + textContent.length,
                    rPr
                });
                fullText += textContent;
            }
        } else if (isWordElement(runChild, 'br') || isWordElement(runChild, 'cr')) {
            const char = String.fromCharCode(characterState.nextCharCode++);
            referenceMap.set(char, runChild);
            breakChars.add(char);
            fullText += char;
            propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
        } else if (isWordElement(runChild, 'tab')) {
            fullText += '\t';
            propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
        } else if (isWordElement(runChild, 'noBreakHyphen')) {
            fullText += '\u2011';
            propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
        } else if (['drawing', 'pict', 'object', 'fldChar', 'instrText', 'sym'].some(name => isWordElement(runChild, name))) {
            const textBoxContent = getFirstElementByTagNSOrTag(runChild, NS_W, 'txbxContent');
            const hasTextBox = isWordElement(runChild, 'pict') && !!textBoxContent;
            const isFieldStructure = isWordElement(runChild, 'fldChar') || isWordElement(runChild, 'instrText');
            const fieldCharType = isWordElement(runChild, 'fldChar')
                ? (runChild.getAttributeNS?.(NS_W, 'fldCharType') || runChild.getAttribute('w:fldCharType') || runChild.getAttribute('fldCharType'))
                : null;

            sentinelMap.push({
                start: fullText.length,
                node: runChild,
                wrapInRun: true,
                rPr,
                zeroWidth: isFieldStructure,
                affinity: fieldCharType === 'end' ? 'left' : 'right',
                isTextBox: hasTextBox,
                originalContainer: hasTextBox ? textBoxContent : undefined
            });
            fullText += '\uFFFC';
            propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
        } else if (isWordElement(runChild, 'footnoteReference') || isWordElement(runChild, 'endnoteReference')) {
            const id = wordAttribute(runChild, 'id');
            if (id) {
                const type = isWordElement(runChild, 'footnoteReference') ? 'FN' : 'EN';
                const tokenString = `{{__${type}_${id}__}}`;
                const char = String.fromCharCode(characterState.nextCharCode++);
                referenceMap.set(char, runChild);
                tokenToCharMap.set(tokenString, char);
                fullText += char;
                propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
            }
        } else if (isWordElement(runChild, 'commentReference')) {
            sentinelMap.push({ start: fullText.length, node: runChild, isCommentMarker: true });
        }
    });

    return fullText;
}

function processHyperlinkForReconstruction(hyperlinkElement, originalFullText, propertyMap) {
    let fullText = originalFullText;

    Array.from(hyperlinkElement.childNodes).forEach(hyperlinkChild => {
        if (!isWordElement(hyperlinkChild, 'r')) return;

        const rPr = getFirstElementByTagNSOrTag(hyperlinkChild, NS_W, 'rPr');
        const texts = getElementsByTagNSOrTag(hyperlinkChild, NS_W, 't');
        texts.forEach(textNode => {
            const textContent = textNode.textContent || '';
            if (textContent.length === 0) return;

            propertyMap.push({
                start: fullText.length,
                end: fullText.length + textContent.length,
                rPr,
                wrapper: hyperlinkElement
            });
            fullText += textContent;
        });
    });

    return fullText;
}
