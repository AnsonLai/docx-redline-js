import { parseOoxml, serializeOoxml } from '../engine/oxml-engine.js';
import { WORD_MAIN_NS } from '../core/paragraph-targeting.js';
import { getAllReferencedNumIdElements, getCurrentNumIdElements } from '../core/current-properties.js';

function parseIntegerAttribute(element, names) {
    if (!element || !Array.isArray(names)) return null;
    for (const name of names) {
        const raw = element.getAttribute(name);
        if (raw == null || raw === '') continue;
        const parsed = Number.parseInt(String(raw), 10);
        if (Number.isFinite(parsed)) return parsed;
    }
    return null;
}

function nextAvailableId(startId, occupiedIds, maxPreferred = null) {
    let candidate = Number.isInteger(startId) && startId > 0 ? startId : 1;
    const occupied = occupiedIds instanceof Set ? occupiedIds : new Set();

    while (occupied.has(candidate)) {
        candidate += 1;
    }

    if (Number.isInteger(maxPreferred) && maxPreferred > 0 && candidate > maxPreferred) {
        for (let probe = 1; probe <= maxPreferred; probe += 1) {
            if (!occupied.has(probe)) return probe;
        }
    }

    return candidate;
}

/**
 * Builds a dynamic numbering-id state from existing numbering XML.
 *
 * This avoids hardcoded ID floors and keeps IDs deterministic relative to the
 * current document numbering definitions.
 *
 * @param {string} numberingXml - Existing `word/numbering.xml` content
 * @param {{
 *   minId?: number,
 *   maxPreferred?: number
 * }} [options={}] - Optional ID preferences
 * @returns {{
 *   nextNumId: number,
 *   nextAbstractNumId: number,
 *   usedNumIds: Set<number>,
 *   usedAbstractNumIds: Set<number>,
 *   minId: number,
 *   maxPreferred: number
 * }}
 */
export function createDynamicNumberingIdState(numberingXml, options = {}) {
    const minId = Number.isInteger(options?.minId) && options.minId > 0 ? options.minId : 1;
    const maxPreferred = Number.isInteger(options?.maxPreferred) && options.maxPreferred >= minId
        ? options.maxPreferred
        : 32767;

    const usedNumIds = new Set();
    const usedAbstractNumIds = new Set();

    if (String(numberingXml || '').trim()) {
        try {
            const numberingDoc = parseOoxml(numberingXml);
            const abstractNums = Array.from(numberingDoc.getElementsByTagNameNS('*', 'abstractNum'));
            const nums = Array.from(numberingDoc.getElementsByTagNameNS('*', 'num'));

            for (const abstractNum of abstractNums) {
                const id = parseIntegerAttribute(abstractNum, ['w:abstractNumId', 'abstractNumId']);
                if (id != null) usedAbstractNumIds.add(id);
            }
            for (const num of nums) {
                const id = parseIntegerAttribute(num, ['w:numId', 'numId']);
                if (id != null) usedNumIds.add(id);
            }
        } catch {
            // Ignore malformed numbering XML and fall back to empty sets.
        }
    }

    const maxUsedNumId = usedNumIds.size > 0 ? Math.max(...usedNumIds) : 0;
    const maxUsedAbstractNumId = usedAbstractNumIds.size > 0 ? Math.max(...usedAbstractNumIds) : 0;
    const baseNumId = Math.max(minId, maxUsedNumId + 1);
    const baseAbstractNumId = Math.max(minId, maxUsedAbstractNumId + 1);

    return {
        // Kept so list edits can recognise which existing definition a source list uses and continue it.
        sourceNumberingXml: String(numberingXml || ''),
        nextNumId: nextAvailableId(baseNumId, usedNumIds, maxPreferred),
        nextAbstractNumId: nextAvailableId(baseAbstractNumId, usedAbstractNumIds, maxPreferred),
        usedNumIds,
        usedAbstractNumIds,
        minId,
        maxPreferred
    };
}

function normalizeNumberingIdState(state) {
    if (!state || typeof state !== 'object') return null;

    if (!(state.usedNumIds instanceof Set)) state.usedNumIds = new Set();
    if (!(state.usedAbstractNumIds instanceof Set)) state.usedAbstractNumIds = new Set();

    if (!Number.isInteger(state.minId) || state.minId < 1) {
        state.minId = 1;
    }
    if (!Number.isInteger(state.maxPreferred) || state.maxPreferred < state.minId) {
        state.maxPreferred = 32767;
    }

    if (!Number.isInteger(state.nextNumId) || state.nextNumId < state.minId) {
        state.nextNumId = state.minId;
    }
    if (!Number.isInteger(state.nextAbstractNumId) || state.nextAbstractNumId < state.minId) {
        state.nextAbstractNumId = state.minId;
    }

    state.nextNumId = nextAvailableId(state.nextNumId, state.usedNumIds, state.maxPreferred);
    state.nextAbstractNumId = nextAvailableId(state.nextAbstractNumId, state.usedAbstractNumIds, state.maxPreferred);
    return state;
}

/**
 * Reserves the next available ID from a mutable numbering-id state.
 *
 * @param {ReturnType<typeof createDynamicNumberingIdState>} state
 * @param {'num'|'abstract'} [kind='num']
 * @returns {number|null}
 */
export function reserveNextNumberingId(state, kind = 'num') {
    const normalized = normalizeNumberingIdState(state);
    if (!normalized) return null;

    const useAbstract = kind === 'abstract';
    const id = useAbstract ? normalized.nextAbstractNumId : normalized.nextNumId;
    if (!Number.isInteger(id) || id < 1) return null;

    if (useAbstract) {
        normalized.usedAbstractNumIds.add(id);
        normalized.nextAbstractNumId = nextAvailableId(id + 1, normalized.usedAbstractNumIds, normalized.maxPreferred);
    } else {
        normalized.usedNumIds.add(id);
        normalized.nextNumId = nextAvailableId(id + 1, normalized.usedNumIds, normalized.maxPreferred);
    }

    return id;
}

/**
 * Reserves the next available numbering IDs on a mutable numbering-id state.
 *
 * @param {ReturnType<typeof createDynamicNumberingIdState>} state
 * @returns {{ numId: number, abstractNumId: number } | null}
 */
export function reserveNextNumberingIdPair(state) {
    const numId = reserveNextNumberingId(state, 'num');
    const abstractNumId = reserveNextNumberingId(state, 'abstract');
    if (numId == null || abstractNumId == null) return null;

    return { numId, abstractNumId };
}

export function hasXmlParseError(doc) {
    if (!doc || !doc.documentElement) return true;
    if (doc.documentElement.localName === 'parsererror') return true;
    return doc.getElementsByTagName('parsererror').length > 0;
}

export function isDirectWordChild(node, localName) {
    return !!(
        node &&
        node.nodeType === 1 &&
        node.namespaceURI === WORD_MAIN_NS &&
        node.localName === localName
    );
}

function getAttributeFirst(element, names) {
    for (const name of names || []) {
        const value = element?.getAttribute?.(name);
        if (value != null && value !== '') return value;
    }
    return null;
}

export function getElementId(element, names) {
    const raw = getAttributeFirst(element, names);
    const parsed = Number.parseInt(String(raw || ''), 10);
    return Number.isFinite(parsed) ? parsed : null;
}

function setElementId(element, preferredName, idValue) {
    element?.setAttribute?.(preferredName, String(idValue));
}

function setElementVal(element, value) {
    element?.setAttribute?.('w:val', String(value));
}

/**
 * Overwrites all paragraph-level `w:numId` references in a node collection.
 *
 * @param {Node[]|null|undefined} paragraphNodes
 * @param {string|number|null|undefined} targetNumId
 */
export function overwriteParagraphNumIds(paragraphNodes, targetNumId) {
    if (!Array.isArray(paragraphNodes) || targetNumId == null) return;
    for (const node of paragraphNodes) {
        // R1: retained (tracked-deleted) source paragraphs keep the binding Reject All must restore.
        if (isTrackedDeletedParagraph(node)) continue;
        // Current numIds only: a w:pPrChange snapshot keeps its numId so Reject restores the original binding.
        for (const numIdNode of getCurrentNumIdElements(node)) {
            setElementVal(numIdNode, targetNumId);
        }
    }
}

/**
 * Extracts the first `w:numId` value from a node collection.
 *
 * @param {Node[]|null|undefined} paragraphNodes
 * @returns {string|null}
 */
export function extractFirstParagraphNumId(paragraphNodes) {
    for (const node of paragraphNodes || []) {
        for (const numIdNode of getCurrentNumIdElements(node)) {
            const numId = getElementId(numIdNode, ['w:val', 'val']);
            if (numId != null) return String(numId);
        }
    }
    return null;
}

/**
 * Builds multilevel decimal numbering XML for explicit-start header conversion.
 *
 * @param {string|number} numId
 * @param {string|number} abstractNumId
 * @param {number} startAt
 * @returns {string}
 */
export function buildExplicitDecimalMultilevelNumberingXml(numId, abstractNumId, startAt) {
    const safeNumId = String(numId);
    const safeAbstractNumId = String(abstractNumId);
    const safeStartAt = Number.isInteger(startAt) && startAt > 0 ? startAt : 1;
    const levelsXml = Array.from({ length: 9 }, (_, level) => {
        const lvlText = Array.from({ length: level + 1 }, (_, i) => `%${i + 1}`).join('.') + '.';
        const left = 720 * (level + 1);
        return `
        <w:lvl w:ilvl="${level}">
            <w:start w:val="1"/>
            <w:numFmt w:val="decimal"/>
            <w:lvlText w:val="${lvlText}"/>
            <w:lvlJc w:val="left"/>
            <w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr>
        </w:lvl>`;
    }).join('');
    return `
<w:numbering xmlns:w="${WORD_MAIN_NS}">
    <w:abstractNum w:abstractNumId="${safeAbstractNumId}">
        <w:multiLevelType w:val="multilevel"/>
        ${levelsXml}
    </w:abstractNum>
    <w:num w:numId="${safeNumId}">
        <w:abstractNumId w:val="${safeAbstractNumId}"/>
        <w:lvlOverride w:ilvl="0">
            <w:startOverride w:val="${safeStartAt}"/>
        </w:lvlOverride>
    </w:num>
</w:numbering>`.trim();
}

/** A paragraph whose mark is tracked-deleted is retained source content (R1: its numPr is never rebound). */
export function isTrackedDeletedParagraph(node) {
    if (!node || node.nodeType !== 1 || node.localName !== 'p') return false;
    const pPr = Array.from(node.childNodes || []).find(child => isDirectWordChild(child, 'pPr'));
    const rPr = pPr && Array.from(pPr.childNodes || []).find(child => isDirectWordChild(child, 'rPr'));
    return !!rPr && Array.from(rPr.childNodes || []).some(child => isDirectWordChild(child, 'del'));
}

/**
 * Remaps incoming numbering payload IDs to document-safe IDs, and updates the
 * provided replacement nodes to reference the remapped `w:numId` values.
 *
 * @param {string} numberingXml
 * @param {Node[]} replacementNodes
 * @param {ReturnType<typeof createDynamicNumberingIdState>} numberingIdState
 * @returns {{ numberingXml: string, replacementNodes: Node[] }}
 */
export function remapNumberingPayloadForDocument(numberingXml, replacementNodes, numberingIdState) {
    const numberingDoc = parseOoxml(numberingXml);
    if (hasXmlParseError(numberingDoc)) {
        return {
            numberingXml: String(numberingXml || ''),
            replacementNodes: Array.isArray(replacementNodes)
                ? replacementNodes.map(node => node?.cloneNode ? node.cloneNode(true) : node)
                : []
        };
    }

    const abstractNumMap = new Map();
    const numIdMap = new Map();

    const abstractNums = Array.from(numberingDoc.getElementsByTagNameNS('*', 'abstractNum'));
    for (const abstractNum of abstractNums) {
        const oldId = getElementId(abstractNum, ['w:abstractNumId', 'abstractNumId']);
        if (oldId == null) continue;
        const newId = reserveNextNumberingId(numberingIdState, 'abstract');
        if (newId == null) continue;
        abstractNumMap.set(oldId, newId);
        setElementId(abstractNum, 'w:abstractNumId', newId);
    }

    const nums = Array.from(numberingDoc.getElementsByTagNameNS('*', 'num'));
    for (const num of nums) {
        const oldNumId = getElementId(num, ['w:numId', 'numId']);
        if (oldNumId == null) continue;
        const newNumId = reserveNextNumberingId(numberingIdState, 'num');
        if (newNumId == null) continue;
        numIdMap.set(oldNumId, newNumId);
        setElementId(num, 'w:numId', newNumId);

        const abstractNumIdNode = Array.from(num.getElementsByTagNameNS('*', 'abstractNumId'))[0] || null;
        if (abstractNumIdNode) {
            const oldAbsRef = getElementId(abstractNumIdNode, ['w:val', 'val']);
            if (oldAbsRef != null && abstractNumMap.has(oldAbsRef)) {
                setElementVal(abstractNumIdNode, abstractNumMap.get(oldAbsRef));
            }
        }
    }

    const clonedNodes = Array.isArray(replacementNodes)
        ? replacementNodes.map(node => node?.cloneNode ? node.cloneNode(true) : node)
        : [];
    for (const node of clonedNodes) {
        // A paragraph whose mark is tracked-deleted is retained source content. Its numId refers to the
        // document's own definitions, never to this payload, so it must keep pointing at them.
        if (isTrackedDeletedParagraph(node)) continue;
        // Deliberately historical-inclusive: every numId in the payload (including pPrChange snapshots)
        // refers to the payload's numbering part, so all of them must follow the remap or Reject would
        // restore a binding to a definition that no longer exists.
        for (const numIdNode of getAllReferencedNumIdElements(node)) {
            const oldNumRef = getElementId(numIdNode, ['w:val', 'val']);
            if (oldNumRef != null && numIdMap.has(oldNumRef)) {
                setElementVal(numIdNode, numIdMap.get(oldNumRef));
            }
        }
    }

    return {
        numberingXml: serializeOoxml(numberingDoc),
        replacementNodes: clonedNodes
    };
}
