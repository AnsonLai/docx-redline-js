/**
 * Reconstruction reconciliation mode orchestration.
 */

import { computeWordDiffs } from '../pipeline/diff-engine.js';
import { buildReconstructionMapping, findReconstructionParagraphRange } from './reconstruction-mapper.js';
import { applyReconstructionDiffs } from './reconstruction-writer.js';
import { withOoxmlSourceType } from '../core/word-xml.js';

/**
 * Applies reconstruction mode reconciliation.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} originalText - Original text (kept for signature compatibility)
 * @param {string} modifiedText - Modified text
 * @param {XMLSerializer} serializer - Serializer instance
 * @param {string} author - Author name
 * @param {Array} formatHints - Format hints
 * @param {boolean} [generateRedlines=true] - Track change toggle
 * @param {{ diffTimeoutSeconds?: number }} [diffOptions={}] - Diff configuration
 * @returns {{ oxml: string, hasChanges: boolean }}
 */
export function applyReconstructionMode(xmlDoc, originalText, modifiedText, serializer, author, formatHints, generateRedlines = true, diffOptions = {}, options = {}) {
    const selectedParagraphs = findReconstructionParagraphRange(xmlDoc, originalText);
    if (selectedParagraphs === null) {
        return withOoxmlSourceType({
            oxml: serializer.serializeToString(xmlDoc),
            hasChanges: false,
            status: 'error',
            error: {
                code: 'PARTIAL_TARGET',
                message: 'Original text did not identify a complete contiguous paragraph range for reconstruction.'
            }
        });
    }

    const mapping = buildReconstructionMapping(xmlDoc, modifiedText, selectedParagraphs);
    if (mapping.paragraphs.length === 0) {
        return withOoxmlSourceType({ oxml: serializer.serializeToString(xmlDoc), hasChanges: false });
    }

    const wordDiffs = computeWordDiffs(mapping.originalFullText, mapping.processedModifiedText, {
        ...diffOptions,
        atomicChars: new Set(mapping.referenceMap.keys())
    });
    const diffs = keepUnchangedTextOutsideWrappers(anchorSharedReferences(wordDiffs, mapping.referenceMap), mapping);

    return withOoxmlSourceType(applyReconstructionDiffs(
        xmlDoc,
        diffs,
        mapping,
        serializer,
        author,
        formatHints,
        generateRedlines,
        options
    ));
}

/** Splits diffs into runs of equal parts and change groups (deleted + inserted text). */
function forEachChangeGroup(diffs, onEqual, onGroup) {
    let index = 0;
    while (index < diffs.length) {
        if (diffs[index][0] === 0) {
            onEqual(diffs[index]);
            index++;
            continue;
        }
        const group = [];
        while (index < diffs.length && diffs[index][0] !== 0) group.push(diffs[index++]);
        onGroup(
            group,
            group.filter(([op]) => op === -1).map(([, text]) => text).join(''),
            group.filter(([op]) => op === 1).map(([, text]) => text).join('')
        );
    }
}

function pushChange(result, deleted, inserted) {
    if (deleted === inserted) {
        if (deleted) result.push([0, deleted]);
        return;
    }
    if (deleted) result.push([-1, deleted]);
    if (inserted) result.push([1, inserted]);
}

/**
 * The writer keeps a reference placeholder (manual break, note reference) that
 * appears in both the deleted and inserted text of one change group, but only
 * emits it with the insertion, after all deleted text. Semantic cleanup can
 * merge edits on both sides of a break into such a group, so Reject All would
 * move deleted text across the break. Split the group at shared references so
 * each one stays an unchanged anchor between its own before/after edits.
 */
function anchorSharedReferences(diffs, referenceMap) {
    if (referenceMap.size === 0) return diffs;
    const result = [];
    forEachChangeGroup(diffs, part => result.push(part), (group, deleted, inserted) => {
        const deletedRefs = [...deleted].filter(char => referenceMap.has(char) && inserted.includes(char));
        const insertedRefs = [...inserted].filter(char => deletedRefs.includes(char));
        if (deletedRefs.length === 0 || deletedRefs.join('') !== insertedRefs.join('')) {
            result.push(...group);
            return;
        }
        let deletedStart = 0;
        let insertedStart = 0;
        for (const ref of deletedRefs) {
            const deletedAt = deleted.indexOf(ref, deletedStart);
            const insertedAt = inserted.indexOf(ref, insertedStart);
            pushChange(result, deleted.slice(deletedStart, deletedAt), inserted.slice(insertedStart, insertedAt));
            result.push([0, ref]);
            deletedStart = deletedAt + 1;
            insertedStart = insertedAt + 1;
        }
        pushChange(result, deleted.slice(deletedStart), inserted.slice(insertedStart));
    });
    return result;
}

function commonPrefixLength(left, right) {
    const limit = Math.min(left.length, right.length);
    let length = 0;
    while (length < limit && left[length] === right[length]) length++;
    return length;
}

function commonSuffixLength(left, right, limit) {
    let length = 0;
    while (length < limit && left[left.length - 1 - length] === right[right.length - 1 - length]) length++;
    return length;
}

/**
 * Word tokens can span a hyperlink boundary (`example.org` + plain `.`), so a
 * replaced token would delete and reinsert unchanged text across it and move
 * that text into or out of the link. Where the replacement's shared prefix or
 * suffix reaches a wrapper boundary in the original, keep that part unchanged.
 */
function keepUnchangedTextOutsideWrappers(diffs, mapping) {
    const wrapperAt = index => mapping.getRunProperties(index).wrapper || null;
    const result = [];
    let originalOffset = 0;

    forEachChangeGroup(diffs, part => {
        result.push(part);
        originalOffset += part[1].length;
    }, (group, deleted, inserted) => {
        const start = originalOffset;
        const end = start + deleted.length;
        originalOffset = end;
        // Only a boundary strictly inside the deleted text splits it; an edge
        // of the whole change (`link` -> `hyperlink` inside a link) does not.
        const isWrapperBoundary = index => index > start && index < end
            && wrapperAt(index - 1) !== wrapperAt(index);

        let prefix = 0;
        let suffix = 0;
        if (deleted && inserted) {
            const sharedSuffix = commonSuffixLength(deleted, inserted, Math.min(deleted.length, inserted.length));
            for (let length = sharedSuffix; length > 0; length--) {
                if (isWrapperBoundary(end - length)) { suffix = length; break; }
            }
            const sharedPrefix = Math.min(
                commonPrefixLength(deleted, inserted),
                Math.min(deleted.length, inserted.length) - suffix
            );
            for (let length = sharedPrefix; length > 0; length--) {
                if (isWrapperBoundary(start + length)) { prefix = length; break; }
            }
        }
        if (prefix === 0 && suffix === 0) {
            result.push(...group);
            return;
        }

        if (prefix > 0) result.push([0, deleted.slice(0, prefix)]);
        pushChange(result, deleted.slice(prefix, deleted.length - suffix), inserted.slice(prefix, inserted.length - suffix));
        if (suffix > 0) result.push([0, deleted.slice(deleted.length - suffix)]);
    });

    return result;
}
