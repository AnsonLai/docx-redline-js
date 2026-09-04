/**
 * OOXML Formatting Removal Utilities
 * 
 * Provides functions to surgically remove formatting from OOXML runs
 * while preserving the text content and other properties.
 */

import { parseOoxml, serializeOoxml } from './oxml-engine.js';
import { getDefaultAuthor } from '../adapters/config.js';
import {
    RevisionIdAllocator,
    createRevisionIdAllocator,
    createRevisionMetadata,
    seedRevisionIdsFromDocument
} from '../core/types.js';
import { createWordElement } from '../core/word-xml.js';
import { refreshRunPropertyChangeIds } from '../core/revision-cloning.js';

function removeNode(node) {
    if (node?.parentNode) {
        node.parentNode.removeChild(node);
    }
}

/**
 * Removes specific formatting properties from a run properties (w:rPr) element.
 * This allows surgical removal of bold, italic, underline, color, etc. from OOXML.
 * 
 * @param {Element} rPr - The w:rPr element to modify
 * @param {string[]} formatTypes - Array of format types to remove: ['bold', 'italic', 'underline', 'strikethrough', 'color', 'highlight', 'fontSize', 'fontFamily', 'all']
 * @returns {Element|null} Modified rPr element, or null if all formatting removed
 */
export function removeFormattingFromRPr(rPr, formatTypes = ['all']) {
    if (!rPr) return null;

    const rPrClone = rPr.cloneNode(true);

    if (formatTypes.includes('all')) {
        // Remove all character formatting properties
        const toRemove = ['w:b', 'w:i', 'w:u', 'w:strike', 'w:dstrike', 'w:color',
            'w:sz', 'w:szCs', 'w:rFonts', 'w:highlight', 'w:vertAlign',
            'w:spacing', 'w:w', 'w:kern', 'w:position'];
        toRemove.forEach(tag => {
            // Handle both namespaced and non-namespaced versions
            const elements = rPrClone.querySelectorAll(`${tag}, ${tag.replace('w:', '')}`);
            elements.forEach(removeNode);
        });
    } else {
        // Remove specific properties
        const tagMap = {
            'bold': 'w:b',
            'italic': 'w:i',
            'underline': 'w:u',
            'strikethrough': 'w:strike',
            'doubleStrike': 'w:dstrike',
            'color': 'w:color',
            'highlight': 'w:highlight',
            'fontSize': 'w:sz',
            'fontSizeCs': 'w:szCs', // Complex script font size
            'fontFamily': 'w:rFonts',
            'superscript': 'w:vertAlign',
            'subscript': 'w:vertAlign'
        };

        formatTypes.forEach(type => {
            const tag = tagMap[type];
            if (tag) {
                // Handle both namespaced and non-namespaced versions
                const elements = rPrClone.querySelectorAll(`${tag}, ${tag.replace('w:', '')}`);
                elements.forEach(removeNode);
            }
        });
    }

    // Return null if rPr is now empty (no children)
    return rPrClone.children.length > 0 ? rPrClone : null;
}

/**
 * Applies formatting removal to OOXML containing the specified text.
 * Searches for text in runs and removes specified formatting properties.
 * 
 * @param {string} ooxmlString - OOXML string (paragraph or larger structure)
 * @param {string} targetText - Text to find and remove formatting from
 * @param {string[]} formatTypes - Array of format types to remove
 * @returns {string} Modified OOXML string
 */
export function applyFormattingRemovalToOoxml(ooxmlString, targetText, formatTypes) {
    if (!targetText || !ooxmlString) return ooxmlString;

    const doc = parseOoxml(ooxmlString);
    if (!doc) return ooxmlString;
    const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

    // Find all text runs
    const runs = doc.getElementsByTagNameNS(NS_W, 'r');

    // Also handle runs inside w:ins (insertions)
    const insertions = doc.getElementsByTagNameNS(NS_W, 'ins');
    const allRuns = [...Array.from(runs)];

    for (const ins of insertions) {
        const insideRuns = ins.getElementsByTagNameNS(NS_W, 'r');
        allRuns.push(...Array.from(insideRuns));
    }

    for (const run of allRuns) {
        // Extract text from this run
        const textNodes = run.getElementsByTagNameNS(NS_W, 't');
        const runText = Array.from(textNodes).map(t => t.textContent).join('');

        // If this run contains the target text (or equals it)
        if (runText.includes(targetText) || runText === targetText) {
            // Find the rPr element
            const rPrElements = run.getElementsByTagNameNS(NS_W, 'rPr');

            if (rPrElements.length > 0) {
                const rPr = rPrElements[0];
                const newRPr = removeFormattingFromRPr(rPr, formatTypes);

                if (newRPr) {
                    // Replace with modified rPr
                    if (rPr.parentNode) {
                        rPr.parentNode.replaceChild(newRPr, rPr);
                    }
                } else {
                    // Remove entire rPr if empty
                    removeNode(rPr);
                }
            }
        }
    }

    return serializeOoxml(doc);
}

// ==================== HIGHLIGHT INJECTION ====================

/**
 * Word API color names → OOXML w:highlight values
 */
const HIGHLIGHT_COLOR_MAP = {
    'yellow': 'yellow', 'green': 'green', 'cyan': 'cyan',
    'magenta': 'magenta', 'blue': 'blue', 'red': 'red',
    'darkblue': 'darkBlue', 'darkcyan': 'darkCyan',
    'darkgreen': 'darkGreen', 'darkmagenta': 'darkMagenta',
    'darkred': 'darkRed', 'darkyellow': 'darkYellow',
    'gray25': 'lightGray', 'gray50': 'darkGray',
    'black': 'black', 'white': 'white'
};

/**
 * Injects a highlight color into a run properties (w:rPr) element.
 * If rPr is null, creates a new rPr element with the highlight.
 * Supports track changes via w:rPrChange.
 * 
 * @param {Document} doc - The OOXML document (for creating new elements)
 * @param {Element|null} rPr - The w:rPr element to modify (or null to create new)
 * @param {string} color - Highlight color name (default: 'yellow')
 * @param {Object} options - Options { generateRedlines: boolean, author: string }
 * @returns {Element} Modified or new rPr element with highlight
 */
export function injectHighlightIntoRPr(doc, rPr, color = 'yellow', options = {}) {
    const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const ooxmlColor = HIGHLIGHT_COLOR_MAP[color.toLowerCase()] || 'yellow';
    const generateRedlines = options?.generateRedlines ?? false;
    const author = options?.author || getDefaultAuthor();

    let rPrElement = rPr;
    if (!rPrElement) {
        // Create new rPr element
        rPrElement = createWordElement(doc, 'w:rPr');
    } else {
        rPrElement = rPr.cloneNode(true);
    }

    // Capture "previous" state for redlines BEFORE modification
    // Clone the *original* rPr children before we touch them
    let previousRPrState = null;
    if (generateRedlines) {
        previousRPrState = createWordElement(doc, 'w:rPr');
        Array.from(rPrElement.childNodes).forEach(child => {
            // Don't include existing rPrChange in the "previous" state wrapper usually, 
            // but for simplicity we clone children. Word generally handles nested track changes poorly,
            // so best to exclude rPrChange from the inner previous state.
            if (child.nodeName !== 'w:rPrChange') {
                previousRPrState.appendChild(child.cloneNode(true));
            }
        });
    }

    // --- APPLY CHANGE ---
    // Remove any existing highlight
    const existingHighlight = rPrElement.getElementsByTagNameNS(NS_W, 'highlight');
    Array.from(existingHighlight).forEach(removeNode);

    // Create and add new highlight element
    const highlightEl = createWordElement(doc, 'w:highlight');
    highlightEl.setAttributeNS(NS_W, 'w:val', ooxmlColor);
    rPrElement.appendChild(highlightEl);

    // --- WRAP IN REDLINES IF ENABLED ---
    if (generateRedlines && previousRPrState) {
        const rPrChange = createWordElement(doc, 'w:rPrChange');

        // Attributes
        const metadata = createRevisionMetadata(author, doc);
        rPrChange.setAttribute('w:id', String(metadata.id));
        rPrChange.setAttribute('w:author', metadata.author);
        rPrChange.setAttribute('w:date', metadata.date);

        // Format: <w:rPrChange ...> <w:rPr>...previous...</w:rPr> </w:rPrChange>
        rPrChange.appendChild(previousRPrState);

        // Remove any EXISTING rPrChange to avoid duplicates or nested weirdness
        const existingChange = rPrElement.getElementsByTagNameNS(NS_W, 'rPrChange');
        Array.from(existingChange).forEach(removeNode);

        // Append to rPr
        rPrElement.appendChild(rPrChange);
    }

    return rPrElement;
}

/**
 * Applies highlight formatting to OOXML runs containing the specified text.
 * Performs surgical splitting of runs if the text is a substring.
 * 
 * @param {string} ooxmlString - OOXML string (paragraph or package)
 * @param {string} targetText - Text to find and highlight
 * @param {string} color - Highlight color (default: 'yellow')
 * @returns {string} Modified OOXML string with highlights applied
 */
export function applyHighlightToOoxml(ooxmlString, targetText, color = 'yellow', options = {}) {
    if (!targetText || !ooxmlString) return ooxmlString;

    const doc = parseOoxml(ooxmlString);
    if (!doc) return ooxmlString;
    let revisionIdAllocator;
    if (options?._revisionIdAllocator instanceof RevisionIdAllocator) {
        revisionIdAllocator = options._revisionIdAllocator;
        seedRevisionIdsFromDocument(doc, revisionIdAllocator);
    } else {
        revisionIdAllocator = createRevisionIdAllocator(doc);
    }
    const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

    // Helper to get text from a run
    const getRunText = (run) => {
        const textNodes = run.getElementsByTagNameNS(NS_W, 't');
        return Array.from(textNodes).map(t => t.textContent).join('');
    };

    // Find all runs recursively (this already includes runs inside w:ins)
    const allRuns = Array.from(doc.getElementsByTagNameNS(NS_W, 'r'));

    const sourceRunsWithClaimedRevisionIds = new WeakSet();
    const cloneRunWithText = (sourceRun, text, shouldHighlight) => {
        const clonedRun = sourceRun.cloneNode(true);
        if (sourceRunsWithClaimedRevisionIds.has(sourceRun)) {
            refreshRunPropertyChangeIds(clonedRun, revisionIdAllocator);
        } else {
            sourceRunsWithClaimedRevisionIds.add(sourceRun);
        }
        const textNodes = clonedRun.getElementsByTagNameNS(NS_W, 't');
        Array.from(textNodes).forEach(removeNode);

        const newText = createWordElement(doc, 'w:t');
        newText.setAttribute('xml:space', 'preserve');
        newText.textContent = text;
        clonedRun.appendChild(newText);

        if (shouldHighlight) {
            const rPrElements = clonedRun.getElementsByTagNameNS(NS_W, 'rPr');
            const existingRPr = rPrElements.length > 0 ? rPrElements[0] : null;
            const newRPr = injectHighlightIntoRPr(doc, existingRPr, color, options);

            if (existingRPr) {
                clonedRun.replaceChild(newRPr, existingRPr);
            } else {
                clonedRun.insertBefore(newRPr, clonedRun.firstChild);
            }
        }

        return clonedRun;
    };

    for (const run of allRuns) {
        const runText = getRunText(run);

        if (!runText) continue;

        const matchIndexes = [];
        let searchOffset = 0;
        while (searchOffset <= runText.length - targetText.length) {
            const matchIndex = runText.indexOf(targetText, searchOffset);
            if (matchIndex === -1) break;
            matchIndexes.push(matchIndex);
            searchOffset = matchIndex + targetText.length;
        }
        if (matchIndexes.length === 0) continue;

        const parent = run.parentNode;
        if (!parent) {
            console.warn("[Highlight] Run parent is null; skipping. Likely already processed.");
            continue;
        }

        const fragment = doc.createDocumentFragment();
        let cursor = 0;

        for (const matchIndex of matchIndexes) {
            if (matchIndex > cursor) {
                fragment.appendChild(cloneRunWithText(run, runText.slice(cursor, matchIndex), false));
            }
            fragment.appendChild(cloneRunWithText(
                run,
                runText.slice(matchIndex, matchIndex + targetText.length),
                true
            ));
            cursor = matchIndex + targetText.length;
        }

        if (cursor < runText.length) {
            fragment.appendChild(cloneRunWithText(run, runText.slice(cursor), false));
        }

        parent.replaceChild(fragment, run);
    }

    return serializeOoxml(doc);
}
