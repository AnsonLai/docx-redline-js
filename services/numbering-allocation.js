/**
 * Numbering allocator: the single place that creates, remaps, continues, merges and validates
 * numbering definitions (`w:abstractNum` / `w:num`).
 *
 * Requirements (docs/plans: canonical list operations WP-2, revision fidelity WP-5a):
 *  - R1: tracked-deleted source paragraphs never change `numPr` (Reject All must restore the source binding).
 *  - R2: the standalone runner and the `openDocx` facade reach numbering merge through this module only
 *    (`mergeNumberingParts` / `mergeNumberingXmlBySchemaOrder`); neither keeps a merge of its own.
 *  - R3: reuse of the source list is allocator policy (`allocateNumberingForReplacement`), not a heuristic
 *    scattered at a call site.
 *
 * Ids are always reserved through the document-scoped state (`createDynamicNumberingIdState`,
 * `reserveNextNumberingIdPair`), so a newly allocated id never collides with a definition the document owns.
 * The old `num` / `abstractNum` are never deleted or overwritten: Reject All and other paragraphs may still
 * reference them. Current properties are read only through `core/current-properties.js`.
 */

import { parseOoxml, serializeOoxml } from '../engine/oxml-engine.js';
import { WORD_MAIN_NS } from '../core/paragraph-targeting.js';
import {
    getCurrentNumIdElements,
    getCurrentParagraphProperties,
    getHistoricalParagraphProperties
} from '../core/current-properties.js';
import {
    getElementId,
    hasXmlParseError,
    isDirectWordChild,
    isTrackedDeletedParagraph,
    remapNumberingPayloadForDocument,
    reserveNextNumberingIdPair
} from './numbering-helpers.js';

const MAX_LEVEL = 8;
// CT_Lvl child order (ECMA-376 17.9.6).
const LVL_CHILD_ORDER = [
    'start', 'numFmt', 'lvlRestart', 'pStyle', 'isLgl', 'suff', 'lvlText', 'lvlPicBulletId', 'legacy', 'lvlJc', 'pPr', 'rPr'
];

const wordChildren = (node, localName) =>
    Array.from(node?.childNodes || []).filter(child => isDirectWordChild(child, localName));
const descendants = (node, localName) => Array.from(node?.getElementsByTagNameNS?.('*', localName) || []);
const valOf = element => {
    const raw = element?.getAttribute?.('w:val') ?? element?.getAttribute?.('val');
    const parsed = Number.parseInt(String(raw ?? ''), 10);
    return Number.isFinite(parsed) ? parsed : null;
};

// ---------------------------------------------------------------------------------------------------------
// Canonical form
// ---------------------------------------------------------------------------------------------------------

function canonicalNode(node, ignoreIds) {
    if (node.nodeType === 3 || node.nodeType === 4) {
        const text = String(node.nodeValue || '').trim();
        return text ? JSON.stringify(text) : '';
    }
    if (node.nodeType !== 1) return '';
    const attrs = Array.from(node.attributes || [])
        .filter(attr => !/^xmlns(:|$)/.test(attr.name))
        .filter(attr => !(ignoreIds && /^(w:)?(abstractNumId|numId)$/.test(attr.name)))
        .map(attr => `${attr.localName || attr.name}=${attr.value}`)
        .sort()
        .join(',');
    if (ignoreIds && node.localName === 'abstractNumId' && node.parentNode?.localName === 'num') {
        return `<abstractNumId ${attrs.replace(/val=[^,]*/, 'val=*')}/>`;
    }
    const children = Array.from(node.childNodes || []).map(child => canonicalNode(child, ignoreIds)).join('');
    return `<${node.localName} ${attrs}>${children}</>`;
}

/**
 * Canonical string of the definition a `numId` resolves to (the `w:num` plus its `w:abstractNum`), with the
 * ids themselves elided so definitions can be compared across independently allocated documents.
 * Returns null when the numId does not resolve.
 *
 * @param {string|Document} numbering
 * @param {string|number} numId
 */
export function canonicalNumberingDefinition(numbering, numId) {
    const doc = typeof numbering === 'string' ? parseOoxml(numbering) : numbering;
    if (!doc || hasXmlParseError(doc)) return null;
    const num = descendants(doc, 'num').find(candidate => getElementId(candidate, ['w:numId', 'numId']) === Number(numId));
    if (!num) return null;
    const abstractId = valOf(wordChildren(num, 'abstractNumId')[0]);
    const abstract = descendants(doc, 'abstractNum')
        .find(candidate => getElementId(candidate, ['w:abstractNumId', 'abstractNumId']) === abstractId);
    return `${canonicalNode(num, true)}|${abstract ? canonicalNode(abstract, true) : 'missing-abstract'}`;
}

// ---------------------------------------------------------------------------------------------------------
// Create new definitions
// ---------------------------------------------------------------------------------------------------------

function findAbstractForNum(numberingDoc, numId) {
    const num = descendants(numberingDoc, 'num').find(candidate => getElementId(candidate, ['w:numId', 'numId']) === Number(numId));
    const abstractId = valOf(wordChildren(num, 'abstractNumId')[0]);
    if (abstractId == null) return null;
    return descendants(numberingDoc, 'abstractNum')
        .find(candidate => getElementId(candidate, ['w:abstractNumId', 'abstractNumId']) === abstractId) || null;
}

function setLvlChildVal(doc, lvl, localName, value) {
    let child = wordChildren(lvl, localName)[0];
    if (!child) {
        child = doc.createElementNS(WORD_MAIN_NS, `w:${localName}`);
        const rank = LVL_CHILD_ORDER.indexOf(localName);
        const anchor = Array.from(lvl.childNodes || []).find(candidate =>
            candidate.nodeType === 1 && LVL_CHILD_ORDER.indexOf(candidate.localName) > rank);
        if (anchor) lvl.insertBefore(child, anchor);
        else lvl.appendChild(child);
    }
    child.setAttribute('w:val', String(value));
}

function defaultLevel(doc, ilvl) {
    const lvl = doc.createElementNS(WORD_MAIN_NS, 'w:lvl');
    lvl.setAttribute('w:ilvl', String(ilvl));
    setLvlChildVal(doc, lvl, 'start', 1);
    setLvlChildVal(doc, lvl, 'numFmt', 'decimal');
    setLvlChildVal(doc, lvl, 'lvlText', `%${ilvl + 1}.`);
    setLvlChildVal(doc, lvl, 'lvlJc', 'left');
    const pPr = doc.createElementNS(WORD_MAIN_NS, 'w:pPr');
    const ind = doc.createElementNS(WORD_MAIN_NS, 'w:ind');
    ind.setAttribute('w:left', String(720 * (ilvl + 1)));
    ind.setAttribute('w:hanging', '360');
    pPr.appendChild(ind);
    lvl.appendChild(pPr);
    return lvl;
}

/**
 * Copies the level geometry (`w:lvlJc`, `w:pPr` indentation, `w:rPr` fonts, `w:suff`) of a source `w:lvl` onto a
 * target `w:lvl`, replacing whatever geometry the target had.
 *
 * @param {Element} sourceLvl
 * @param {Element} targetLvl
 */
export function copyLevelGeometry(sourceLvl, targetLvl) {
    const doc = targetLvl.ownerDocument;
    for (const name of ['suff', 'lvlJc', 'pPr', 'rPr']) {
        for (const existing of wordChildren(targetLvl, name)) targetLvl.removeChild(existing);
        const source = wordChildren(sourceLvl, name)[0];
        if (!source) continue;
        const clone = doc.importNode(source, true);
        const rank = LVL_CHILD_ORDER.indexOf(name);
        const anchor = Array.from(targetLvl.childNodes || []).find(candidate =>
            candidate.nodeType === 1 && LVL_CHILD_ORDER.indexOf(candidate.localName) > rank);
        if (anchor) targetLvl.insertBefore(clone, anchor);
        else targetLvl.appendChild(clone);
    }
}

/**
 * Adds `w:lvlOverride/w:startOverride` entries to a `w:num`. Per-instance restart values live on the `w:num`,
 * never on the shared `w:abstractNum`.
 *
 * @param {Element} numElement
 * @param {Record<number|string, number>} startOverrides - ilvl -> start value
 */
export function applyStartOverridesToNum(numElement, startOverrides) {
    const doc = numElement.ownerDocument;
    const entries = Object.entries(startOverrides || {})
        .map(([ilvl, start]) => [Number(ilvl), Number(start)])
        .filter(([ilvl, start]) => Number.isInteger(ilvl) && ilvl >= 0 && ilvl <= MAX_LEVEL && Number.isInteger(start) && start >= 0)
        .sort((a, b) => a[0] - b[0]);
    for (const [ilvl, start] of entries) {
        for (const existing of wordChildren(numElement, 'lvlOverride')) {
            if (valOfIlvl(existing) === ilvl) numElement.removeChild(existing);
        }
        const override = doc.createElementNS(WORD_MAIN_NS, 'w:lvlOverride');
        override.setAttribute('w:ilvl', String(ilvl));
        const startOverride = doc.createElementNS(WORD_MAIN_NS, 'w:startOverride');
        startOverride.setAttribute('w:val', String(start));
        override.appendChild(startOverride);
        numElement.appendChild(override);
    }
}

function valOfIlvl(element) {
    return getElementId(element, ['w:ilvl', 'ilvl']);
}

/**
 * Creates a NEW `w:abstractNum` and `w:num` with fresh ids reserved from the document-scoped state.
 *
 * Levels not redefined keep the source level untouched; redefined levels (`numFmt`, `lvlText`, `start`) keep the
 * source level's geometry (indentation, fonts, justification). Style links, `w:nsid` and `w:tmpl` are not copied,
 * so the new definition never aliases a style or another list's identity.
 *
 * @param {{
 *   numberingIdState: object,
 *   sourceNumberingXml?: string|null,
 *   sourceNumId?: number|string|null,
 *   levels?: Array<{ ilvl: number, numFmt?: string, lvlText?: string, start?: number }>,
 *   startOverrides?: Record<number|string, number>
 * }} params
 * @returns {{ numId: number, abstractNumId: number, numberingXml: string }|null} `null` when no ids can be reserved
 */
export function allocateNumberingDefinition(params) {
    const { numberingIdState, sourceNumberingXml, sourceNumId, levels = [], startOverrides } = params || {};
    const pair = reserveNextNumberingIdPair(numberingIdState);
    if (!pair) return null;

    const sourceDoc = String(sourceNumberingXml || '').trim() ? parseOoxml(sourceNumberingXml) : null;
    const sourceAbstract = sourceDoc && !hasXmlParseError(sourceDoc) && sourceNumId != null
        ? findAbstractForNum(sourceDoc, sourceNumId)
        : null;
    const sourceLevels = new Map(descendants(sourceAbstract, 'lvl')
        .map(lvl => [valOfIlvl(lvl), lvl]).filter(([ilvl]) => ilvl != null));

    const doc = parseOoxml(`<w:numbering xmlns:w="${WORD_MAIN_NS}"/>`);
    const root = doc.documentElement;
    const abstract = doc.createElementNS(WORD_MAIN_NS, 'w:abstractNum');
    abstract.setAttribute('w:abstractNumId', String(pair.abstractNumId));
    const multiLevelType = doc.createElementNS(WORD_MAIN_NS, 'w:multiLevelType');
    multiLevelType.setAttribute('w:val', wordChildren(sourceAbstract, 'multiLevelType')[0]?.getAttribute('w:val') || 'multilevel');
    abstract.appendChild(multiLevelType);

    const redefined = new Map(levels.filter(level => Number.isInteger(level?.ilvl)).map(level => [level.ilvl, level]));
    const maxLevel = Math.max(MAX_LEVEL, ...sourceLevels.keys(), ...redefined.keys());
    for (let ilvl = 0; ilvl <= maxLevel; ilvl += 1) {
        const source = sourceLevels.get(ilvl);
        const lvl = source ? doc.importNode(source, true) : defaultLevel(doc, ilvl);
        for (const pStyle of wordChildren(lvl, 'pStyle')) lvl.removeChild(pStyle);
        lvl.removeAttribute('w:tplc');
        lvl.removeAttribute('w:tentative');
        const redefinition = redefined.get(ilvl);
        if (redefinition) {
            if (redefinition.numFmt != null) setLvlChildVal(doc, lvl, 'numFmt', redefinition.numFmt);
            if (redefinition.lvlText != null) setLvlChildVal(doc, lvl, 'lvlText', redefinition.lvlText);
            if (redefinition.start != null) setLvlChildVal(doc, lvl, 'start', redefinition.start);
        }
        abstract.appendChild(lvl);
    }

    const num = doc.createElementNS(WORD_MAIN_NS, 'w:num');
    num.setAttribute('w:numId', String(pair.numId));
    const abstractRef = doc.createElementNS(WORD_MAIN_NS, 'w:abstractNumId');
    abstractRef.setAttribute('w:val', String(pair.abstractNumId));
    num.appendChild(abstractRef);
    applyStartOverridesToNum(num, startOverrides);

    root.appendChild(abstract);
    root.appendChild(num);
    return { numId: pair.numId, abstractNumId: pair.abstractNumId, numberingXml: serializeOoxml(doc) };
}

// ---------------------------------------------------------------------------------------------------------
// Replacement policy (R1, R3)
// ---------------------------------------------------------------------------------------------------------

function levelFormats(abstractNum) {
    const formats = new Map();
    for (const lvl of descendants(abstractNum, 'lvl')) {
        const ilvl = valOfIlvl(lvl);
        const fmt = descendants(lvl, 'numFmt')[0]?.getAttribute('w:val') || null;
        if (ilvl != null && fmt) formats.set(ilvl, fmt);
    }
    return formats;
}

function normalizeListFormat(format) {
    return format === 'bullet' || format === 'circle' || format === 'square' ? 'bullet' : format;
}

function currentBindings(node) {
    return getCurrentNumIdElements(node).map(numIdElement => {
        const ilvlElement = wordChildren(numIdElement.parentNode, 'ilvl')[0];
        return { numIdElement, numId: valOf(numIdElement), ilvl: ilvlElement ? (valOf(ilvlElement) ?? 0) : 0 };
    });
}

/**
 * Policy: a same-kind list replacement continues the single list the source paragraphs already use. When every
 * new list paragraph's generated level has the same format as that list's level, the new paragraphs bind to the
 * existing `numId` and the generated payload is dropped. Anything else (format change, several source lists,
 * unknown source definitions) allocates. R1: tracked-deleted paragraphs (retained source content) are never
 * rebound.
 */
function continueSourceList(numberingXml, replacementNodes, sourceParagraphs, numberingIdState) {
    const sourceXml = numberingIdState?.sourceNumberingXml;
    if (!numberingXml || !String(sourceXml || '').trim() || !Array.isArray(replacementNodes)) return null;

    const sourceNumIds = new Set();
    for (const paragraph of sourceParagraphs || []) {
        for (const { numId } of currentBindings(paragraph)) {
            if (numId != null && numId !== 0) sourceNumIds.add(numId);
        }
    }
    if (sourceNumIds.size !== 1) return null;
    const [sourceNumId] = sourceNumIds;

    const sourceDoc = parseOoxml(sourceXml);
    const payloadDoc = parseOoxml(numberingXml);
    if (hasXmlParseError(sourceDoc) || hasXmlParseError(payloadDoc)) return null;
    const sourceAbstract = findAbstractForNum(sourceDoc, sourceNumId);
    if (!sourceAbstract) return null;
    const sourceFormats = levelFormats(sourceAbstract);

    const bindings = [];
    for (const node of replacementNodes) {
        if (isTrackedDeletedParagraph(node)) continue;
        bindings.push(...currentBindings(node));
    }
    if (bindings.length === 0) return null;

    const payloadFormats = new Map();
    for (const { numId, ilvl } of bindings) {
        if (numId == null) return null;
        if (!payloadFormats.has(numId)) {
            const abstract = findAbstractForNum(payloadDoc, numId);
            if (!abstract) return null;
            payloadFormats.set(numId, levelFormats(abstract));
        }
        const generated = payloadFormats.get(numId).get(ilvl);
        const existing = sourceFormats.get(ilvl);
        if (!generated || !existing || normalizeListFormat(generated) !== normalizeListFormat(existing)) return null;
    }

    const clonedNodes = replacementNodes.map(node => (node?.cloneNode ? node.cloneNode(true) : node));
    for (const node of clonedNodes) {
        if (isTrackedDeletedParagraph(node)) continue;
        for (const numIdElement of getCurrentNumIdElements(node)) numIdElement.setAttribute('w:val', String(sourceNumId));
    }
    return { numId: sourceNumId, replacementNodes: clonedNodes };
}

/**
 * Resolves the numbering of generated replacement paragraphs against the document.
 *
 * Policy order: (1) `continue-source` when `allowContinueSource` and the replacement is the same kind of list as
 * the single source list; (2) `allocate`: remap the generated definitions onto fresh ids from the document-scoped
 * state; (3) `passthrough` when there is nothing to allocate against. Tracked-deleted source paragraphs keep their
 * `numPr` in every case (R1).
 *
 * @param {{
 *   numberingXml: string|null,
 *   replacementNodes: Node[],
 *   sourceParagraphs?: Node[],
 *   numberingIdState?: object|null,
 *   allowContinueSource?: boolean
 * }} params
 * @returns {{ policy: 'continue-source'|'allocate'|'passthrough', numberingXml: string|null, replacementNodes: Node[], numId?: number }}
 */
export function allocateNumberingForReplacement(params) {
    const { numberingXml, replacementNodes, sourceParagraphs, numberingIdState, allowContinueSource = true } = params || {};
    if (numberingXml && numberingIdState && allowContinueSource) {
        const continued = continueSourceList(numberingXml, replacementNodes, sourceParagraphs, numberingIdState);
        if (continued) {
            return { policy: 'continue-source', numberingXml: null, replacementNodes: continued.replacementNodes, numId: continued.numId };
        }
    }
    if (numberingXml && numberingIdState) {
        const remapped = remapNumberingPayloadForDocument(numberingXml, replacementNodes, numberingIdState);
        return { policy: 'allocate', numberingXml: remapped.numberingXml, replacementNodes: remapped.replacementNodes };
    }
    return { policy: 'passthrough', numberingXml: numberingXml ?? null, replacementNodes };
}

// ---------------------------------------------------------------------------------------------------------
// Merge (R2): the only numbering merge in the package
// ---------------------------------------------------------------------------------------------------------

function insertNumberingNodeInSchemaOrder(root, node, kind) {
    const directChildren = Array.from(root.childNodes || []).filter(
        child => child && child.nodeType === 1 && child.namespaceURI === WORD_MAIN_NS
    );
    const anchor = kind === 'abstract'
        ? directChildren.find(child => child.localName === 'num' || child.localName === 'numIdMacAtCleanup')
        : directChildren.find(child => child.localName === 'numIdMacAtCleanup');
    if (anchor) root.insertBefore(node, anchor);
    else root.appendChild(node);
}

/**
 * Merges incoming numbering definitions into an existing numbering part, preserving schema order
 * (`abstractNum*` before `num*`). Definitions are added by id; an incoming id that already exists is kept as is
 * when its canonical content is equal, and is reported through `onConflict` (never overwritten) when it differs,
 * because the paragraphs written with the incoming id would silently bind to someone else's list.
 *
 * Every caller (standalone runner consumers, `openDocx`, `ensureNumberingArtifactsInZip`) merges through here.
 *
 * @param {string} existingNumberingXml
 * @param {string} incomingNumberingXml
 * @param {{ onConflict?: (conflict: { kind: 'abstractNum'|'num', id: number }) => void }} [options]
 * @returns {string}
 */
export function mergeNumberingXmlBySchemaOrder(existingNumberingXml, incomingNumberingXml, options = {}) {
    const existingText = String(existingNumberingXml || '');
    const incomingText = String(incomingNumberingXml || '');
    if (!incomingText.trim()) return existingText;
    if (!existingText.trim()) return incomingText;

    try {
        const existingDoc = parseOoxml(existingText);
        const incomingDoc = parseOoxml(incomingText);
        if (hasXmlParseError(existingDoc) || hasXmlParseError(incomingDoc)) return existingText;
        const existingRoot = existingDoc.documentElement;
        const incomingRoot = incomingDoc.documentElement;
        if (!existingRoot || !incomingRoot) return existingText;

        for (const [kind, localName, attr] of [['abstract', 'abstractNum', 'abstractNumId'], ['num', 'num', 'numId']]) {
            const existingById = new Map(wordChildren(existingRoot, localName)
                .map(node => [getElementId(node, [`w:${attr}`, attr]), node]).filter(([id]) => id != null));
            for (const incomingNode of wordChildren(incomingRoot, localName)) {
                const id = getElementId(incomingNode, [`w:${attr}`, attr]);
                if (id == null) continue;
                const present = existingById.get(id);
                if (present) {
                    if (typeof options?.onConflict === 'function' && canonicalNode(present, false) !== canonicalNode(incomingNode, false)) {
                        options.onConflict({ kind: localName, id });
                    }
                    continue;
                }
                const imported = existingDoc.importNode(incomingNode, true);
                insertNumberingNodeInSchemaOrder(existingRoot, imported, kind);
                existingById.set(id, imported);
            }
        }
        return serializeOoxml(existingDoc);
    } catch {
        return existingText;
    }
}

/**
 * Merges every incoming numbering payload into the existing part, in order. Same arguments for every caller.
 *
 * @param {string|null|undefined} existingNumberingXml
 * @param {string|string[]|null|undefined} incomingParts
 * @param {Parameters<typeof mergeNumberingXmlBySchemaOrder>[2]} [options]
 * @returns {string}
 */
export function mergeNumberingParts(existingNumberingXml, incomingParts, options = {}) {
    let merged = String(existingNumberingXml || '');
    for (const part of (Array.isArray(incomingParts) ? incomingParts : [incomingParts]).filter(Boolean)) {
        merged = mergeNumberingXmlBySchemaOrder(merged, part, options);
    }
    return merged;
}

// ---------------------------------------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------------------------------------

/**
 * Checks that a direct paragraph binding `{ numId, level }` resolves in the document's numbering: the `w:num`
 * exists (in the source part or among ids the current batch allocated), it has an `abstractNum`, and the level
 * is defined by that abstract or by the `w:num`'s own `w:lvlOverride/w:lvl`. A binding is a reference, so this
 * never allocates.
 *
 * @param {{ sourceNumberingXml?: string|null, usedNumIds?: Set<number> }|null|undefined} numberingIdState
 * @param {string|number} numId
 * @param {number} [level]
 * @returns {{ ok: true } | { ok: false, reason: 'unknown-num'|'missing-abstract'|'undefined-level', message: string }}
 */
export function checkNumberingBinding(numberingIdState, numId, level = 0) {
    const id = Number(numId);
    const sourceXml = String(numberingIdState?.sourceNumberingXml || '');
    const doc = sourceXml.trim() ? parseOoxml(sourceXml) : null;
    const sourceNum = doc && !hasXmlParseError(doc)
        ? descendants(doc, 'num').find(candidate => getElementId(candidate, ['w:numId', 'numId']) === id)
        : null;
    if (!sourceNum) {
        // Ids allocated earlier in this batch exist only as pending numbering parts.
        if (numberingIdState?.usedNumIds instanceof Set && numberingIdState.usedNumIds.has(id)) return { ok: true };
        return {
            ok: false,
            reason: 'unknown-num',
            message: `numbering.numId ${id} does not exist in the document's numbering definitions.`
        };
    }
    const abstract = findAbstractForNum(doc, id);
    if (!abstract) {
        return { ok: false, reason: 'missing-abstract', message: `numId ${id} points at an abstractNum that does not exist.` };
    }
    const definesLevel = container => wordChildren(container, 'lvl')
        .some(lvl => Number.parseInt(lvl.getAttribute('w:ilvl') ?? lvl.getAttribute('ilvl') ?? '0', 10) === level);
    const overrideDefines = wordChildren(sourceNum, 'lvlOverride').some(override => definesLevel(override));
    if (!definesLevel(abstract) && !overrideDefines) {
        return { ok: false, reason: 'undefined-level', message: `numId ${id} does not define list level ${level}.` };
    }
    return { ok: true };
}

/**
 * Numbering integrity: every body `w:numId` (current and `pPrChange` snapshot) resolves to a `w:num`, every
 * `w:num` resolves to an `abstractNum`, every `w:ilvl` in use exists in that abstract (or in the `w:num`'s own
 * `w:lvlOverride/w:lvl`), and numbering ids are unique. `numId` 0 means "no list" and is always valid.
 *
 * @param {string|Document} documentXml
 * @param {string|Document|null|undefined} numberingXml
 * @returns {Array<{ code: string, severity: 'error', message: string, numId?: number, abstractNumId?: number, ilvl?: number }>}
 */
export function validateNumberingIntegrity(documentXml, numberingXml) {
    const issues = [];
    const seen = new Set();
    const report = (code, message, extra = {}) => {
        const key = `${code}:${extra.numId ?? ''}:${extra.abstractNumId ?? ''}:${extra.ilvl ?? ''}`;
        if (seen.has(key)) return;
        seen.add(key);
        issues.push({ code, severity: 'error', message, ...extra });
    };

    const bodyDoc = typeof documentXml === 'string' ? parseOoxml(documentXml) : documentXml;
    if (!bodyDoc || hasXmlParseError(bodyDoc)) return issues;
    const numberingDoc = typeof numberingXml === 'string'
        ? (numberingXml.trim() ? parseOoxml(numberingXml) : null)
        : (numberingXml || null);
    const numberingOk = numberingDoc && !hasXmlParseError(numberingDoc);

    const abstractById = new Map();
    const numById = new Map();
    if (numberingOk) {
        for (const [localName, attr, map] of [['abstractNum', 'abstractNumId', abstractById], ['num', 'numId', numById]]) {
            for (const node of descendants(numberingDoc, localName)) {
                const id = getElementId(node, [`w:${attr}`, attr]);
                if (id == null) continue;
                if (map.has(id)) report('NUMBERING_DUPLICATE_ID', `Numbering part defines ${localName} ${id} more than once.`, { [attr]: id });
                else map.set(id, node);
            }
        }
        for (const [numId, num] of numById) {
            const abstractId = valOf(wordChildren(num, 'abstractNumId')[0]);
            if (abstractId == null) {
                report('NUMBERING_ABSTRACT_UNDEFINED', `w:num ${numId} has no w:abstractNumId.`, { numId });
            } else if (!abstractById.has(abstractId)) {
                report('NUMBERING_ABSTRACT_UNDEFINED', `w:num ${numId} references missing w:abstractNum ${abstractId}.`, { numId, abstractNumId: abstractId });
            }
        }
    }

    const levelDefined = (numId, ilvl) => {
        const num = numById.get(numId);
        const abstract = abstractById.get(valOf(wordChildren(num, 'abstractNumId')[0]));
        if (!abstract) return true; // already reported as an undefined abstractNum
        if (descendants(abstract, 'lvl').some(lvl => valOfIlvl(lvl) === ilvl)) return true;
        return wordChildren(num, 'lvlOverride').some(override =>
            valOfIlvl(override) === ilvl && wordChildren(override, 'lvl').length > 0);
    };

    for (const paragraph of Array.from(bodyDoc.getElementsByTagNameNS(WORD_MAIN_NS, 'p'))) {
        const bindings = [getCurrentParagraphProperties(paragraph).numPr, getHistoricalParagraphProperties(paragraph)?.numPr];
        for (const numPr of bindings) {
            if (!numPr?.numIdElement) continue;
            const numId = valOf(numPr.numIdElement);
            if (numId == null || numId === 0) continue;
            if (!numberingOk || !numById.has(numId)) {
                report('NUMBERING_NUM_UNDEFINED', `w:numId ${numId} is used in the document but no w:num defines it.`, { numId });
                continue;
            }
            const ilvl = numPr.ilvlElement ? valOf(numPr.ilvlElement) : null;
            if (ilvl != null && !levelDefined(numId, ilvl)) {
                report('NUMBERING_LEVEL_UNDEFINED', `w:ilvl ${ilvl} is used with w:numId ${numId} but its abstractNum does not define that level.`, { numId, ilvl });
            }
        }
    }
    return issues;
}
