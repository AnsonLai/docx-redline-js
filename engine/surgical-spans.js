import { appendParagraphBoundary } from '../core/paragraph-offset-policy.js';
import { getFirstElementByTag } from '../core/xml-query.js';
import { isWordElement } from '../core/word-xml.js';

export function getRunChildText(child) {
    if (isWordElement(child, 't')) return child.textContent || '';
    if (isWordElement(child, 'br') || isWordElement(child, 'cr')) return '\n';
    if (isWordElement(child, 'tab')) return '\t';
    if (isWordElement(child, 'noBreakHyphen')) return '\u2011';
    if (isWordElement(child, 'softHyphen')) return '\u00ad';
    return '';
}

export function isTextLikeRunChild(child) {
    return isWordElement(child, 't')
        || isWordElement(child, 'br')
        || isWordElement(child, 'cr')
        || isWordElement(child, 'tab')
        || isWordElement(child, 'noBreakHyphen')
        || isWordElement(child, 'softHyphen');
}

export function buildSurgicalTextSpans(paragraphs) {
    let fullText = '';
    const textSpans = [];

    paragraphs.forEach((paragraph, paragraphIndex) => {
        const container = paragraph.parentNode;

        for (let child = paragraph.firstChild; child; child = child.nextSibling) {
            if (isWordElement(child, 'r')) {
                fullText += processRunElement(child, paragraph, container, fullText.length, textSpans).text;
            } else if (isWordElement(child, 'hyperlink')) {
                for (let hc = child.firstChild; hc; hc = hc.nextSibling) {
                    if (isWordElement(hc, 'r')) {
                        fullText += processRunElement(hc, paragraph, container, fullText.length, textSpans).text;
                    }
                }
            }
        }

        fullText = appendParagraphBoundary(fullText, paragraphIndex, paragraphs.length);
    });

    return { fullText, textSpans };
}

function processRunElement(run, paragraph, container, currentOffset, textSpans) {
    const rPr = getFirstElementByTag(run, 'w:rPr');
    let localOffset = currentOffset;
    const textParts = [];

    for (let child = run.firstChild; child; child = child.nextSibling) {
        if (isWordElement(child, 't')) {
            const text = child.textContent || '';
            if (text.length === 0) continue;

            textSpans.push({
                charStart: localOffset,
                charEnd: localOffset + text.length,
                textElement: child,
                runElement: run,
                paragraph,
                container,
                rPr
            });
            localOffset += text.length;
            textParts.push(text);
        } else if (isTextLikeRunChild(child)) {
            const text = getRunChildText(child);
            textSpans.push({
                charStart: localOffset,
                charEnd: localOffset + 1,
                textElement: child,
                runElement: run,
                paragraph,
                container,
                rPr
            });
            localOffset += 1;
            textParts.push(text);
        }
    }

    return { text: textParts.join('') };
}

export function buildSpanIndex(textSpans) {
    const spans = textSpans
        .slice()
        .sort((a, b) => a.charStart - b.charStart || a.charEnd - b.charEnd);

    const starts = spans.map(span => span.charStart);
    const ends = spans.map(span => span.charEnd);

    return { spans, starts, ends };
}

export function forEachOverlappingSpan(spanIndex, startPos, endPos, callback) {
    if (endPos <= startPos || spanIndex.spans.length === 0) {
        return;
    }

    let index = upperBound(spanIndex.ends, startPos);
    while (index < spanIndex.spans.length) {
        const span = spanIndex.spans[index];
        if (span.charStart >= endPos) {
            break;
        }
        callback(span);
        index++;
    }
}

export function findContainingSpan(spanIndex, pos) {
    if (spanIndex.spans.length === 0) return null;

    const index = upperBound(spanIndex.starts, pos) - 1;
    if (index < 0) return null;

    const span = spanIndex.spans[index];
    return pos >= span.charStart && pos < span.charEnd ? span : null;
}

export function findFirstSpanEndingAt(spanIndex, pos) {
    const index = lowerBound(spanIndex.ends, pos);
    if (index < spanIndex.spans.length && spanIndex.ends[index] === pos) {
        return spanIndex.spans[index];
    }
    return null;
}

export function findLastSpanEndingBeforeOrAt(spanIndex, pos) {
    const index = upperBound(spanIndex.ends, pos) - 1;
    if (index >= 0) {
        return spanIndex.spans[index];
    }
    return null;
}

function upperBound(values, target) {
    let left = 0;
    let right = values.length;

    while (left < right) {
        const middle = (left + right) >> 1;
        if (values[middle] <= target) {
            left = middle + 1;
        } else {
            right = middle;
        }
    }

    return left;
}

function lowerBound(values, target) {
    let left = 0;
    let right = values.length;

    while (left < right) {
        const middle = (left + right) >> 1;
        if (values[middle] < target) {
            left = middle + 1;
        } else {
            right = middle;
        }
    }

    return left;
}
