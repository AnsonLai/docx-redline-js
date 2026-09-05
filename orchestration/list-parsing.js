/**
 * Shared markdown list parsing for command adapters.
 *
 * Keeps command-layer list parsing aligned with reconciliation marker logic.
 */

import { parseListItem } from '../pipeline/list-markers.js';

/**
 * Parses markdown list-like content into structured items.
 *
 * Output shape is compatible with command-layer expectations from `parseMarkdownList`.
 *
 * @param {string} content - Raw markdown/text content
 * @returns {{ type: 'numbered'|'bullet'|'text', items: Array<{ type: 'numbered'|'bullet'|'text', level: number, text: string, marker?: string }> }|null}
 */
export function parseMarkdownListContent(content) {
    if (!content) return null;

    const normalized = String(content).trim();
    if (!normalized) return null;

    const lines = normalized.split('\n');
    const items = [];

    for (const line of lines) {
        if (!line.trim()) continue;

        const parsed = parseListItem(line, { allowZeroSpaceAfterMarker: false, indentSpaces: 2 });
        if (parsed) {

            items.push({
                type: parsed.markerType,
                level: parsed.level,
                text: parsed.text.trim(),
                marker: parsed.marker
            });
            continue;
        }

        items.push({
            type: 'text',
            level: 0,
            text: line.trim()
        });
    }

    if (items.length === 0) return null;

    const hasNumbered = items.some(item => item.type === 'numbered');
    const hasBullet = items.some(item => item.type === 'bullet');

    return {
        type: hasNumbered ? 'numbered' : (hasBullet ? 'bullet' : 'text'),
        items
    };
}

/**
 * Checks whether parsed list data includes at least one real list item.
 *
 * @param {{ items?: Array<{ type?: string }> }|null} parsedListData - Parsed list data
 * @returns {boolean}
 */
export function hasListItems(parsedListData) {
    if (!parsedListData || !Array.isArray(parsedListData.items)) return false;
    return parsedListData.items.some(item => item?.type === 'numbered' || item?.type === 'bullet');
}
