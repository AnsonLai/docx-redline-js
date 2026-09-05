/**
 * Content analysis helpers for paragraph/list/table classification.
 */

import { ContentType } from '../core/types.js';
import { parseListItem } from './list-markers.js';

/**
 * Parses table from markdown-style table text.
 *
 * @param {string} text - Table text
 * @returns {{ headers: string[], rows: string[][], hasHeader: boolean }}
 */
export function parseTable(text) {
    const lines = text
        .split('\n')
        .map(line => line.trim())
        .filter(line => line.startsWith('|'));
    if (lines.length === 0) {
        return { headers: [], rows: [], hasHeader: false };
    }

    const isSeparatorLine = line => {
        const normalized = line.replace(/\s+/g, '');
        return /^\|:?-{3,}:?(\|:?-{3,}:?)+\|?$/.test(normalized);
    };

    const hasHeader = lines.some(isSeparatorLine);
    const dataLines = lines.filter(line => !isSeparatorLine(line));
    const parsedRows = dataLines.map(line =>
        line
            .split('|')
            .slice(1, -1)
            .map(cell => cell.trim())
    );

    if (!hasHeader) {
        return {
            headers: [],
            rows: parsedRows,
            hasHeader: false
        };
    }

    return {
        headers: parsedRows[0] || [],
        rows: parsedRows.slice(1),
        hasHeader: true
    };
}

/**
 * Parses list items from markdown-style list text.
 *
 * @param {string} text - List text
 * @returns {Array<{ line: string, text: string, marker: string, indent: number, level: number, listType: 'bullet'|'numbered' }>}
 */
export function parseListItems(text) {
    const lines = text.split('\n').filter(line => line.trim().length > 0);
    const items = [];

    lines.forEach(line => {
        const parsed = parseListItem(line, { allowZeroSpaceAfterMarker: true, indentSpaces: 2 });
        if (!parsed) return;
        const level = parsed.outlineLevel ?? parsed.level;

        items.push({
            line,
            text: parsed.text,
            marker: parsed.marker,
            indent: parsed.indent,
            level,
            listType: parsed.listType
        });
    });

    return items;
}

/**
 * Detects content type from text.
 *
 * @param {string} text - Text to classify
 * @returns {ContentType}
 */
export function detectContentType(text) {
    const normalized = (text || '').trim();
    if (!normalized) return ContentType.PARAGRAPH;

    const table = parseTable(normalized);
    if (table.headers.length > 0 || table.rows.length > 0) {
        return ContentType.TABLE;
    }

    const listItems = parseListItems(normalized);
    if (listItems.length > 0) {
        const hasBullet = listItems.some(item => item.listType === 'bullet');
        return hasBullet ? ContentType.BULLET_LIST : ContentType.NUMBERED_LIST;
    }

    return ContentType.PARAGRAPH;
}
