/**
 * Strict, dependency-light planning for mixed Markdown document insertions.
 *
 * Agents should run this before replacing one Word paragraph with content that
 * contains headings, paragraphs, lists, or tables. The planner never guesses a
 * malformed table: a Markdown separator row is required so literal pipe text
 * cannot silently reach the document.
 */

import { matchListMarker } from './list-markers.js';

const TABLE_SEPARATOR = /^\s*\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)+\|?\s*$/;

function isBlank(line) {
    return !line || line.trim().length === 0;
}

function isTableLine(line) {
    const trimmed = String(line || '').trim();
    return trimmed.startsWith('|') && trimmed.endsWith('|');
}

function tableCells(line) {
    return String(line || '').trim().split('|').slice(1, -1).map(cell => cell.trim());
}

function classifyLine(line) {
    if (isBlank(line)) return 'blank';
    if (/^\s*#{1,9}\s+\S/.test(line)) return 'heading';
    if (isTableLine(line)) return 'table';
    if (matchListMarker(line)) return 'list';
    return 'paragraph';
}

/**
 * Decomposes mixed Markdown into explicit document blocks and validates the
 * structural syntax required by the OOXML generator.
 *
 * @param {string} markdown
 * @returns {{ valid: boolean, normalizedMarkdown: string, blocks: Array<object>, issues: Array<object>, counts: Record<string, number>, requiresStructuredContent: boolean }}
 */
export function analyzeStructuredContent(markdown) {
    const source = typeof markdown === 'string' ? markdown.replace(/\r\n?/g, '\n') : String(markdown ?? '');
    const lines = source.split('\n');
    const blocks = [];
    const issues = [];
    let index = 0;

    while (index < lines.length) {
        if (isBlank(lines[index])) {
            index++;
            continue;
        }

        const kind = classifyLine(lines[index]);
        const startLine = index + 1;
        if (kind === 'heading') {
            const match = lines[index].match(/^\s*(#{1,9})\s+(.+?)\s*$/);
            blocks.push({ type: 'heading', level: match[1].length, text: match[2], markdown: lines[index].trim() });
            index++;
            continue;
        }

        if (kind === 'table') {
            const tableLines = [];
            while (index < lines.length && isTableLine(lines[index])) {
                tableLines.push(lines[index].trim());
                index++;
            }
            const separatorPresent = tableLines.length > 1 && TABLE_SEPARATOR.test(tableLines[1]);
            const widths = tableLines.filter(line => !TABLE_SEPARATOR.test(line)).map(line => tableCells(line).length);
            if (!separatorPresent) {
                issues.push({
                    severity: 'error',
                    code: 'TABLE_SEPARATOR_REQUIRED',
                    line: startLine,
                    message: 'Markdown tables require a separator row immediately after the header (for example | --- | --- |).'
                });
            }
            if (widths.length < 2) {
                issues.push({
                    severity: 'error',
                    code: 'TABLE_DATA_ROW_REQUIRED',
                    line: startLine,
                    message: 'Markdown tables require a header and at least one data row.'
                });
            } else if (new Set(widths).size > 1) {
                issues.push({
                    severity: 'error',
                    code: 'TABLE_COLUMN_COUNT_MISMATCH',
                    line: startLine,
                    message: `Markdown table rows have inconsistent column counts: ${widths.join(', ')}.`
                });
            }
            blocks.push({
                type: 'table',
                columns: widths[0] || 0,
                rows: Math.max(0, widths.length - 1),
                hasHeader: separatorPresent,
                markdown: tableLines.join('\n')
            });
            continue;
        }

        if (kind === 'list') {
            const listLines = [];
            while (index < lines.length && classifyLine(lines[index]) === 'list') {
                listLines.push(lines[index].trimEnd());
                index++;
            }
            blocks.push({ type: 'list', items: listLines.length, markdown: listLines.join('\n') });
            continue;
        }

        const paragraphLines = [];
        while (index < lines.length && classifyLine(lines[index]) === 'paragraph') {
            paragraphLines.push(lines[index].trim());
            index++;
        }
        const text = paragraphLines.join(' ').trim();
        blocks.push({ type: 'paragraph', text, markdown: text });
    }

    const counts = { heading: 0, paragraph: 0, list: 0, table: 0 };
    for (const block of blocks) counts[block.type] = (counts[block.type] || 0) + 1;
    const normalizedMarkdown = blocks.map(block => block.markdown).join('\n\n');
    return {
        valid: issues.every(issue => issue.severity !== 'error'),
        normalizedMarkdown,
        blocks,
        issues,
        counts,
        requiresStructuredContent: blocks.length > 1 || blocks.some(block => block.type !== 'paragraph')
    };
}

/**
 * Builds one atomic full-document replacement operation from validated mixed
 * Markdown. The operation remains a single target mutation; blocks are not
 * emitted as a fragile sequence of operations that would invalidate the anchor.
 *
 * @param {string|object} target
 * @param {string} markdown
 * @param {{ author?: string, generateRedlines?: boolean, existingRevisions?: string }} [options]
 */
export function planStructuredReplacement(target, markdown, options = {}) {
    const analysis = analyzeStructuredContent(markdown);
    return {
        ...analysis,
        operation: analysis.valid ? {
            type: 'replace',
            target,
            modified: analysis.normalizedMarkdown,
            structuredContent: true,
            ...(options.author ? { author: options.author } : {}),
            ...(typeof options.generateRedlines === 'boolean' ? { generateRedlines: options.generateRedlines } : {}),
            ...(options.existingRevisions ? { existingRevisions: options.existingRevisions } : {})
        } : null
    };
}
