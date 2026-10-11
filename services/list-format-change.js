/**
 * DOM core of `list-format` `change-format` and `join` (canonical list operations plan, WP-5).
 *
 * Both modes rebind EXISTING list items; neither edits a shared numbering definition in place (numbering.xml is not
 * revision-tracked, so an in-place edit would be invisible to Reject All and would restyle every other list that uses
 * the definition):
 * - `change-format` allocates a NEW `w:abstractNum` + `w:num` (copying the source definition's levels and geometry,
 *   with only the requested level redefined) and rebinds the paragraphs to it with one tracked `w:pPrChange` each.
 *   `scope: 'list'` rebinds every live paragraph of the source `numId`; `scope: 'range'` rebinds only the target range,
 *   which must run to the end of the list, with per-level start values so the count continues.
 * - `join` rebinds the paragraphs of one list to the `numId` of an earlier list (`continueFrom`), so they continue
 *   its count. Nothing is allocated.
 *
 * Reject All restores each paragraph's previous `w:numPr` from the `w:pPrChange` snapshot, so the original labels
 * return; Accept All keeps the new binding. Tracked-deleted paragraphs never change `numPr` (allocator rule R1).
 *
 * Everything is refused in a planning pass before any mutation when the result is not exactly predictable from the
 * Word numbering model (verified with Word: `w:num` instances sharing one `w:abstractNum` continue each other's count,
 * a level start (`w:start` or `w:startOverride`) applies again on every restart of that level, and an item below an unused level
 * implicitly starts that level). Current properties are read only through `core/current-properties.js`; no
 * paragraph-mark revision is created (a rebinding is a property change).
 */

import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import {
    getCurrentChild,
    getCurrentParagraphProperties,
    getDirectWordChild,
    getDirectWordChildren
} from '../core/current-properties.js';
import { extractCanonicalParagraphText } from '../core/paragraph-text.js';
import { getParagraphListInfo } from '../core/list-targeting.js';
import { createWordElement } from '../core/word-xml.js';
import { snapshotAndAttachPPrChange } from '../engine/run-builders.js';
import { applyParagraphPropertiesToPPr } from '../engine/rpr-helpers.js';
import { allocateNumberingDefinition, mergeNumberingParts } from './numbering-allocation.js';
import { isTrackedDeletedParagraph } from './numbering-helpers.js';
import {
    LIST_FORMAT_DEFINITION_UNSUPPORTED,
    LIST_FORMAT_JOIN_INVALID,
    LIST_FORMAT_MIXED_RANGE,
    LIST_FORMAT_SPLIT_UNSUPPORTED,
    LIST_FORMAT_STYLE_LINKED,
    LIST_FORMAT_STYLE_NUMBERED,
    LIST_FORMAT_TARGET_NOT_LIST,
    validateLevelText
} from './list-format.js';
import { numberingIdStateOf, parseStyles, refusal, styleNumbering } from './list-format-mutation.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PARAGRAPH_FORMAT_REVISION_CONFLICT = 'PARAGRAPH_FORMAT_REVISION_CONFLICT';
const MAX_LEVEL = 8;

const excerpt = (text, limit = 80) => {
    const value = String(text ?? '');
    return value.length > limit ? `${value.slice(0, limit)}...` : value;
};
const readVal = element => element?.getAttribute?.('w:val') ?? element?.getAttribute?.('val') ?? null;
const readInt = (element, name = 'val') => {
    const raw = element?.getAttribute?.(`w:${name}`) ?? element?.getAttribute?.(name);
    const parsed = Number.parseInt(String(raw ?? ''), 10);
    return Number.isFinite(parsed) ? parsed : null;
};
const isBulletFormat = numFmt => numFmt === 'bullet';

// ---------------------------------------------------------------------------------------------------------
// Numbering model
// ---------------------------------------------------------------------------------------------------------

/**
 * Parsed numbering definitions: `nums` (numId -> { element, abstractId, startOverrides, levelOverrides }) and
 * `abstracts` (abstractNumId -> element). Ids are strings.
 */
function readNumbering(numberingXml) {
    const text = String(numberingXml || '');
    const doc = text.trim() ? parseOoxmlSafe(text, 'application/xml')?.doc : null;
    const nums = new Map();
    const abstracts = new Map();
    if (!doc) return { nums, abstracts };
    for (const abstract of Array.from(doc.getElementsByTagNameNS('*', 'abstractNum'))) {
        const id = readInt(abstract, 'abstractNumId');
        if (id != null && !abstracts.has(String(id))) abstracts.set(String(id), abstract);
    }
    for (const num of Array.from(doc.getElementsByTagNameNS('*', 'num'))) {
        const id = readInt(num, 'numId');
        if (id == null || nums.has(String(id))) continue;
        const startOverrides = new Map();
        const levelOverrides = new Set();
        for (const override of getDirectWordChildren(num, 'lvlOverride')) {
            const ilvl = readInt(override, 'ilvl') ?? 0;
            const start = readInt(getDirectWordChild(override, 'startOverride'));
            if (start != null) startOverrides.set(ilvl, start);
            if (getDirectWordChild(override, 'lvl')) levelOverrides.add(ilvl);
        }
        nums.set(String(id), {
            element: num,
            abstractId: String(readInt(getDirectWordChild(num, 'abstractNumId')) ?? ''),
            startOverrides,
            levelOverrides
        });
    }
    return { nums, abstracts };
}

/** The resolved definition of one `numId`, or a refusal reason. */
function resolveDefinition(numbering, numId) {
    const num = numbering.nums.get(String(numId));
    if (!num) return { reason: 'missing' };
    const abstract = numbering.abstracts.get(num.abstractId);
    if (!abstract) return { reason: 'missing' };
    const levels = new Map();
    for (const lvl of getDirectWordChildren(abstract, 'lvl')) {
        const ilvl = readInt(lvl, 'ilvl') ?? 0;
        levels.set(ilvl, {
            numFmt: readVal(getDirectWordChild(lvl, 'numFmt')) || 'decimal',
            lvlText: readVal(getDirectWordChild(lvl, 'lvlText')) ?? '',
            start: readInt(getDirectWordChild(lvl, 'start')) ?? 0,
            restart: !!getDirectWordChild(lvl, 'lvlRestart'),
            pStyle: readVal(getDirectWordChild(lvl, 'pStyle'))
        });
    }
    const styleLink = readVal(getDirectWordChild(abstract, 'numStyleLink')) || readVal(getDirectWordChild(abstract, 'styleLink'));
    const levelStyles = [...levels.values()].map(level => level.pStyle).filter(Boolean);
    return {
        numId: String(numId),
        num,
        abstract,
        abstractId: num.abstractId,
        levels,
        styleLink: styleLink || null,
        levelStyles,
        /** Effective start of a level for this instance: the num's startOverride, else the abstract's start. */
        startOf: ilvl => (num.startOverrides.has(ilvl) ? num.startOverrides.get(ilvl) : (levels.get(ilvl)?.start ?? 0))
    };
}

/** Paragraph styles in styles.xml that bind one of `numIds` through their own `w:numPr`. */
function stylesBindingNumIds(stylesDoc, numIds) {
    if (!stylesDoc) return [];
    const out = [];
    for (const style of Array.from(stylesDoc.getElementsByTagNameNS('*', 'style'))) {
        const numId = readVal(getCurrentChild(getCurrentChild(getDirectWordChild(style, 'pPr'), 'numPr'), 'numId'));
        if (numId != null && numIds.has(String(numId))) out.push(style.getAttribute('w:styleId') ?? style.getAttribute('styleId'));
    }
    return out;
}

/** Word list paragraphs: body paragraphs whose CURRENT `numPr` binds `numId`, in document order. */
function listMembers(xmlDoc, numId) {
    const members = [];
    for (const paragraph of Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'p'))) {
        const info = getParagraphListInfo(paragraph);
        if (!info || info.numId !== String(numId)) continue;
        members.push({
            paragraph,
            ilvl: info.ilvl,
            deleted: isTrackedDeletedParagraph(paragraph),
            inTextbox: hasAncestor(paragraph, 'txbxContent')
        });
    }
    return members;
}

function hasAncestor(node, localName) {
    for (let current = node?.parentNode; current && current.nodeType === 1; current = current.parentNode) {
        if (current.localName === localName) return true;
    }
    return false;
}

/** numIds (other than `numId`) that share its abstractNum and bind at least one live body paragraph. */
function sharedDefinitionUsers(xmlDoc, numbering, definition) {
    const sharing = [...numbering.nums.entries()]
        .filter(([id, num]) => id !== definition.numId && num.abstractId === definition.abstractId)
        .map(([id]) => id);
    return sharing.filter(id => listMembers(xmlDoc, id).some(member => !member.deleted));
}

/**
 * Word's level counters over list items in document order. `startOf(ilvl)` is the instance start; an item at level k
 * implicitly starts every unused level above it, increments level k and restarts every deeper level. Returns, per
 * item, the counter values of levels 0..k (what a multi-level label such as "%1.%2" can show).
 */
function simulateCounters(levels, startOf, initial = []) {
    const counters = Array.from({ length: MAX_LEVEL + 1 }, (_, i) => (initial[i] ?? null));
    const values = [];
    for (const ilvl of levels) {
        for (let j = 0; j < ilvl; j += 1) if (counters[j] == null) counters[j] = startOf(j);
        counters[ilvl] = counters[ilvl] == null ? startOf(ilvl) : counters[ilvl] + 1;
        for (let j = ilvl + 1; j <= MAX_LEVEL; j += 1) counters[j] = null;
        values.push(counters.slice(0, ilvl + 1));
    }
    return { values, counters };
}

// ---------------------------------------------------------------------------------------------------------
// Shared planning
// ---------------------------------------------------------------------------------------------------------

/**
 * Classifies the target range: every live paragraph must be a direct list item of one numId. Returns
 * `{ numId, items, skipped }` or a refusal.
 */
function classifyRange(paragraphs, { stylesDoc, indexOf, mode }) {
    const items = [];
    const skipped = [];
    const nonList = [];
    for (const paragraph of paragraphs) {
        const paragraphIndex = indexOf(paragraph);
        if (isTrackedDeletedParagraph(paragraph)) {
            // R1: a tracked-deleted paragraph keeps its numPr so Reject All restores it.
            skipped.push({ paragraph, paragraphIndex, reason: 'tracked-deleted' });
            continue;
        }
        const info = getParagraphListInfo(paragraph);
        if (info) {
            items.push({ paragraph, paragraphIndex, ilvl: info.ilvl, numId: info.numId });
            continue;
        }
        const current = getCurrentParagraphProperties(paragraph);
        const styled = current.numPr?.numId === '0' ? null : styleNumbering(stylesDoc, current.pStyle);
        if (styled) {
            return refusal(LIST_FORMAT_STYLE_NUMBERED,
                `Paragraph ${paragraphIndex ?? ''} is numbered by its style "${styled.styleId}" (numId ${styled.numId}), not by a direct list binding; ${mode} rebinds direct list items only.`.replace('  ', ' '),
                { paragraphIndex, styleId: styled.styleId, numId: styled.numId, targetExcerpt: excerpt(extractCanonicalParagraphText(paragraph)) },
                { action: 'convert_in_host', hint: 'Style-linked numbering is not rebound by list-format. Change the style\'s numbering in the host.' });
        }
        nonList.push({ paragraph, paragraphIndex });
    }
    if (items.length === 0) {
        return refusal(LIST_FORMAT_TARGET_NOT_LIST,
            `The target range has no list item; ${mode} changes existing lists.`,
            { paragraphIndexes: nonList.map(item => item.paragraphIndex), targetExcerpt: excerpt(nonList[0] ? extractCanonicalParagraphText(nonList[0].paragraph) : '') },
            { action: 'use_list_format_convert_plain', hint: 'To turn plain paragraphs into a list use list-format convert-plain (typed markers) or paragraph-format numbering (an existing list).' });
    }
    if (nonList.length > 0) {
        return refusal(LIST_FORMAT_MIXED_RANGE,
            `The target range mixes list items and paragraphs that are not list items (paragraphs ${nonList.map(item => item.paragraphIndex).join(', ')}).`,
            { reason: 'non-list', paragraphIndexes: nonList.map(item => item.paragraphIndex) },
            { action: 'narrow_range', hint: 'Target only list items: one paragraph of the list is enough with scope "list".' });
    }
    const numIds = [...new Set(items.map(item => item.numId))];
    if (numIds.length > 1) {
        return refusal(LIST_FORMAT_MIXED_RANGE,
            `The target range spans ${numIds.length} lists (numIds ${numIds.join(', ')}); one ${mode} operation changes one list.`,
            { reason: 'several-lists', numIds, paragraphIndexes: items.filter(item => item.numId !== items[0].numId).map(item => item.paragraphIndex) },
            { action: 'narrow_range', hint: 'Send one operation per list.' });
    }
    return { ok: true, numId: numIds[0], items, skipped };
}

/** Refusals that depend on the source definition: missing, level overrides, style links, shared abstracts. */
function checkDefinition(xmlDoc, numbering, numId, { stylesDoc, role }) {
    const definition = resolveDefinition(numbering, numId);
    if (definition.reason === 'missing') {
        return refusal(LIST_FORMAT_DEFINITION_UNSUPPORTED,
            `numId ${numId} (${role}) has no numbering definition in the numbering part, so its levels cannot be read.`,
            { reason: 'missing', numId: String(numId) },
            { action: 'inspect_numbering', hint: 'Bare document.xml callers must pass the package\'s word/numbering.xml as numberingXml.' });
    }
    if (definition.num.levelOverrides.size > 0) {
        return refusal(LIST_FORMAT_DEFINITION_UNSUPPORTED,
            `numId ${numId} (${role}) redefines level ${[...definition.num.levelOverrides].join(', ')} inside its w:num (w:lvlOverride/w:lvl); list-format does not rebind such instance-level definitions.`,
            { reason: 'level-override', numId: String(numId), levels: [...definition.num.levelOverrides] },
            { action: 'convert_in_host', hint: 'Change this list in the host.' });
    }
    const stylesUsingList = stylesBindingNumIds(stylesDoc, new Set([String(numId), ...[...numbering.nums.entries()]
        .filter(([, num]) => num.abstractId === definition.abstractId).map(([id]) => id)]));
    if (definition.styleLink || definition.levelStyles.length > 0 || stylesUsingList.length > 0) {
        const styleIds = [...new Set([definition.styleLink, ...definition.levelStyles, ...stylesUsingList].filter(Boolean))];
        return refusal(LIST_FORMAT_STYLE_LINKED,
            `numId ${numId} (${role}) is style-linked (${styleIds.map(id => `"${id}"`).join(', ')}): paragraphs can join it through a style without a direct binding, so rebinding the direct items would split the list unpredictably.`,
            { numId: String(numId), styleIds },
            { action: 'convert_in_host', hint: 'Style-linked lists (heading numbering, list styles) are not changed by list-format. Change the numbering of the style in the host.' });
    }
    const shared = sharedDefinitionUsers(xmlDoc, numbering, definition);
    if (shared.length > 0) {
        return refusal(LIST_FORMAT_SPLIT_UNSUPPORTED,
            `numId ${numId} (${role}) shares its abstract definition ${definition.abstractId} with numIds ${shared.join(', ')}, which Word numbers as one continuous list; rebinding one of them would renumber the others.`,
            { reason: 'shared-definition', numId: String(numId), sharedNumIds: shared },
            { action: 'convert_in_host', hint: 'Change these lists together in the host, or restart the other list first.' });
    }
    if (definition.num.element && [...definition.levels.keys()].length === 0) {
        return refusal(LIST_FORMAT_DEFINITION_UNSUPPORTED, `numId ${numId} (${role}) has no levels.`,
            { reason: 'undefined-level', numId: String(numId) }, { action: 'convert_in_host' });
    }
    return { ok: true, definition };
}

/**
 * The paragraphs a scope rebinds. `list`: every live member of the source list. `range`: the range items, which must
 * run to the end of the list (otherwise the following items would renumber). Returns `{ rebind, preceding }`.
 */
function planScope(xmlDoc, numId, rangeItems, scope, indexOf, mode) {
    const members = listMembers(xmlDoc, numId);
    const live = members.filter(member => !member.deleted);
    const textbox = live.find(member => member.inTextbox);
    if (textbox) {
        return refusal(LIST_FORMAT_SPLIT_UNSUPPORTED,
            `numId ${numId} has a list item inside a text box (paragraph ${indexOf(textbox.paragraph) ?? '?'}); list-format changes body lists only.`,
            { reason: 'textbox-member', numId: String(numId) },
            { action: 'convert_in_host', hint: 'Change this list in the host.' });
    }
    const inRange = new Set(rangeItems.map(item => item.paragraph));
    if (scope === 'list') {
        return { ok: true, rebind: live.map(member => ({ ...member, paragraphIndex: indexOf(member.paragraph), inTarget: inRange.has(member.paragraph) })), preceding: [] };
    }
    const firstIndex = live.findIndex(member => inRange.has(member.paragraph));
    const following = live.slice(firstIndex).filter(member => !inRange.has(member.paragraph));
    if (following.length > 0) {
        const lastIndex = indexOf(live[live.length - 1].paragraph);
        return refusal(LIST_FORMAT_SPLIT_UNSUPPORTED,
            `scope "range" would split numId ${numId} mid-sequence: ${following.length} later item(s) of the list (paragraphs ${following.map(member => indexOf(member.paragraph)).join(', ')}) would stay on the source list and renumber.`,
            { reason: 'mid-sequence', numId: String(numId), followingParagraphIndexes: following.map(member => indexOf(member.paragraph)), lastListParagraphIndex: lastIndex },
            { action: 'extend_range_or_use_list_scope', hint: `Extend targetEnd to the last item of the list (paragraph ${lastIndex}), or use scope "list" to ${mode === 'join' ? 'join' : 'change'} the whole list.` });
    }
    return {
        ok: true,
        rebind: live.slice(firstIndex).map(member => ({ ...member, paragraphIndex: indexOf(member.paragraph), inTarget: true })),
        preceding: live.slice(0, firstIndex)
    };
}

function checkPendingSnapshots(rebind, { author, generateRedlines }) {
    // An untracked edit has no Reject All to fall back on, so it must stay inside the paragraphs it names.
    const outside = rebind.filter(item => !item.inTarget);
    if (!generateRedlines && outside.length > 0) {
        return refusal('INVALID_OPERATION',
            `An untracked (generateRedlines: false) list-format would rebind ${outside.length} paragraph(s) outside the target (paragraphs ${outside.map(item => item.paragraphIndex).join(', ')}); untracked edits stay inside their target.`,
            { field: 'generateRedlines', paragraphIndexes: outside.map(item => item.paragraphIndex) },
            { action: 'change_request', hint: 'Keep tracking on, or target the whole list with target..targetEnd.' });
    }
    for (const item of rebind) {
        const pendingChange = getCurrentChild(getDirectWordChild(item.paragraph, 'pPr'), 'pPrChange');
        const pendingAuthor = pendingChange?.getAttribute('w:author') ?? pendingChange?.getAttribute('author') ?? '';
        if (generateRedlines && pendingChange && pendingAuthor !== author) {
            return refusal(PARAGRAPH_FORMAT_REVISION_CONFLICT,
                `Paragraph ${item.paragraphIndex ?? ''} of the list already has a pending paragraph-property change by "${pendingAuthor || 'unattributed'}". Overwriting its snapshot would change what Reject All restores.`.replace('  ', ' '),
                { pendingAuthor, paragraphIndex: item.paragraphIndex },
                { action: 'resolve_pending_revision', hint: 'Ask the user to accept or reject the existing paragraph formatting change first; this tool does not resolve revisions by other reviewers.' });
        }
        item.keepSnapshot = !!pendingChange;
    }
    return { ok: true };
}

function rebindParagraphs(xmlDoc, rebind, numId, { author, generateRedlines }) {
    for (const item of rebind) {
        let pPr = getDirectWordChild(item.paragraph, 'pPr');
        if (!pPr) {
            pPr = createWordElement(xmlDoc, 'w:pPr');
            item.paragraph.insertBefore(pPr, item.paragraph.firstChild);
        }
        // A same-author pending snapshot already records the earliest "before" state; keep it.
        if (generateRedlines && !item.keepSnapshot) snapshotAndAttachPPrChange(xmlDoc, pPr, author);
        applyParagraphPropertiesToPPr(xmlDoc, pPr, { numbering: { numId, level: item.ilvl } });
        item.numId = String(numId);
    }
}

const unchanged = (message, extra = {}) => ({ ok: true, unchanged: true, message, ...extra });

// ---------------------------------------------------------------------------------------------------------
// change-format
// ---------------------------------------------------------------------------------------------------------

/**
 * Changes the numbering format of the list the target range belongs to.
 *
 * @param {Document} xmlDoc
 * @param {Element[]} paragraphs - resolved target range, in document order
 * @param {{ format: string|null, kind: string|null, lvlText: string|null, start: number|null, explicitLevel: number|null, scope: 'list'|'range' }} request
 * @param {{
 *   author: string, generateRedlines?: boolean, runtimeContext?: object|null, stylesXml?: string|null,
 *   numberingXml?: string|null, paragraphIndexOf?: (paragraph: Element) => number|null
 * }} context - `numberingXml` is the effective numbering: the source part plus definitions this batch allocated
 */
export function changeListParagraphsFormat(xmlDoc, paragraphs, request, context = {}) {
    const generateRedlines = context.generateRedlines !== false;
    const author = context.author;
    const indexOf = typeof context.paragraphIndexOf === 'function' ? context.paragraphIndexOf : () => null;
    const stylesDoc = parseStyles(context.stylesXml);
    const numbering = readNumbering(context.numberingXml);
    const scope = request.scope === 'range' ? 'range' : 'list';

    const range = classifyRange(paragraphs, { stylesDoc, indexOf, mode: 'change-format' });
    if (!range.ok) return range;
    const checked = checkDefinition(xmlDoc, numbering, range.numId, { stylesDoc, role: 'the target list' });
    if (!checked.ok) return checked;
    const { definition } = checked;

    // Level: explicit, else the single level the target paragraphs use.
    const targetLevels = [...new Set(range.items.map(item => item.ilvl))].sort((a, b) => a - b);
    if (request.explicitLevel == null && targetLevels.length > 1) {
        return refusal(LIST_FORMAT_MIXED_RANGE,
            `The target paragraphs use list levels ${targetLevels.join(', ')}; change-format redefines one level per operation.`,
            { reason: 'levels', levels: targetLevels },
            { action: 'set_list_level', hint: 'Pass list.level, or target paragraphs of one level.', field: 'list.level', allowedValues: targetLevels });
    }
    const level = request.explicitLevel ?? targetLevels[0];
    const source = definition.levels.get(level);
    if (!source) {
        return refusal(LIST_FORMAT_DEFINITION_UNSUPPORTED, `numId ${range.numId} does not define list level ${level}.`,
            { reason: 'undefined-level', numId: range.numId, level }, { action: 'set_list_level', field: 'list.level', allowedValues: [...definition.levels.keys()] });
    }

    // New level definition.
    const numFmt = request.format ?? (request.kind === 'bullet' ? 'bullet' : source.numFmt);
    const kindChange = isBulletFormat(source.numFmt) !== isBulletFormat(numFmt);
    const lvlText = request.lvlText ?? (!kindChange ? source.lvlText : (isBulletFormat(numFmt) ? '•' : `%${level + 1}.`));
    const lvlTextError = validateLevelText(lvlText, { bullet: isBulletFormat(numFmt), level });
    if (lvlTextError) {
        return refusal('INVALID_OPERATION', `${lvlTextError} (level ${level}, format ${numFmt})`, { field: 'list.lvlText' },
            { action: 'change_request', hint: isBulletFormat(numFmt) ? 'Pass the bullet character as list.lvlText.' : `Pass list.lvlText with %${level + 1}, for example "%${level + 1}."` });
    }

    const plan = planScope(xmlDoc, range.numId, range.items, scope, indexOf, 'change-format');
    if (!plan.ok) return plan;

    // Start values for the new instance.
    const abstractStart = ilvl => definition.levels.get(ilvl)?.start ?? 0;
    const startNew = new Map();
    for (let ilvl = 0; ilvl <= MAX_LEVEL; ilvl += 1) startNew.set(ilvl, definition.startOf(ilvl));
    if (scope === 'range' && plan.preceding.length > 0) {
        const usesRestart = [...plan.preceding, ...plan.rebind].some(member => definition.levels.get(member.ilvl)?.restart);
        if (usesRestart) {
            return refusal(LIST_FORMAT_DEFINITION_UNSUPPORTED,
                `numId ${range.numId} uses w:lvlRestart, so the count a split would continue cannot be computed exactly.`,
                { reason: 'level-restart', numId: range.numId },
                { action: 'extend_range_or_use_list_scope', hint: 'Use scope "list".' });
        }
        const before = simulateCounters(plan.preceding.map(member => member.ilvl), definition.startOf);
        const first = plan.rebind[0].ilvl;
        for (let j = 0; j < first; j += 1) startNew.set(j, before.counters[j] ?? definition.startOf(j));
        startNew.set(first, before.counters[first] == null ? definition.startOf(first) : before.counters[first] + 1);
        // What Word shows for the range today (continuing the list) must equal what the new instance shows.
        const expected = simulateCounters(plan.rebind.map(member => member.ilvl), definition.startOf, before.counters).values;
        const actual = simulateCounters(plan.rebind.map(member => member.ilvl), ilvl => startNew.get(ilvl)).values;
        const mismatch = actual.findIndex((values, index) => JSON.stringify(values) !== JSON.stringify(expected[index]));
        if (mismatch >= 0) {
            return refusal(LIST_FORMAT_SPLIT_UNSUPPORTED,
                `scope "range" cannot continue the count of numId ${range.numId} exactly: a level whose start must change restarts inside the range (paragraph ${plan.rebind[mismatch].paragraphIndex ?? '?'}), and Word applies a level start again on every restart.`,
                { reason: 'nested-restart', numId: range.numId, paragraphIndex: plan.rebind[mismatch].paragraphIndex },
                { action: 'extend_range_or_use_list_scope', hint: 'Start the range at a top-level item of the list, or use scope "list".' });
        }
    }
    if (request.start != null) startNew.set(level, request.start);

    if (!plan.rebind.some(member => member.ilvl === level)) {
        return unchanged(`No paragraph of the ${scope === 'list' ? 'list' : 'range'} uses level ${level}; nothing would change.`, { level, sourceNumId: range.numId });
    }
    const startChanged = request.start != null && request.start !== definition.startOf(level);
    if (numFmt === source.numFmt && lvlText === source.lvlText && !startChanged) {
        return unchanged(`numId ${range.numId} level ${level} already has format ${numFmt} with level text "${lvlText}".`, { level, sourceNumId: range.numId });
    }

    const pending = checkPendingSnapshots(plan.rebind, { author, generateRedlines });
    if (!pending.ok) return pending;

    // Start values (the source instance's own w:startOverride values, a split's continuation values, list.start) are
    // written as the w:start of the NEW abstractNum's levels. That abstractNum is used by the new instance only, so
    // nothing shared changes. Verified in Word: after Revisions.AcceptAll() in a session, a level carrying a
    // w:startOverride that resumes after deeper items is renumbered from the override again (labels B., B. instead of
    // C., D.; correct only after reopening), while a level w:start renders C., D. in both.
    const levelStarts = {};
    for (const [ilvl, start] of startNew) {
        if (start !== abstractStart(ilvl)) levelStarts[ilvl] = start;
    }
    const redefinitions = new Map(Object.entries(levelStarts).map(([ilvl, start]) => [Number(ilvl), { ilvl: Number(ilvl), start }]));
    redefinitions.set(level, {
        ...(redefinitions.get(level) || {}),
        ilvl: level,
        numFmt,
        lvlText,
        ...(numFmt !== source.numFmt ? { dropChildren: kindChange && !isBulletFormat(numFmt) ? ['isLgl', 'lvlPicBulletId', 'legacy'] : ['isLgl'] } : {}),
        ...(kindChange && !isBulletFormat(numFmt) ? { dropRunFonts: true } : {})
    });
    const numberingIdState = numberingIdStateOf(context.runtimeContext, context);
    const allocated = allocateNumberingDefinition({
        numberingIdState,
        sourceNumberingXml: context.numberingXml,
        sourceNumId: range.numId,
        levels: [...redefinitions.values()]
    });
    if (!allocated) {
        return refusal('NUMBERING_ID_EXHAUSTED', 'No numbering ids are available for a new list definition.', {}, { action: 'report_library_failure' });
    }

    rebindParagraphs(xmlDoc, plan.rebind, allocated.numId, { author, generateRedlines });
    return {
        ok: true,
        mode: 'change-format',
        scope,
        format: numFmt,
        level,
        lvlText,
        sourceNumId: range.numId,
        numId: String(allocated.numId),
        levelStarts,
        items: plan.rebind.map(item => ({ paragraph: item.paragraph, paragraphIndex: item.paragraphIndex, level: item.ilvl, numId: item.numId, sourceNumId: range.numId, inTarget: item.inTarget })),
        skipped: range.skipped,
        numberingXml: mergeNumberingParts('', [allocated.numberingXml]),
        allocatedNumIds: [allocated.numId]
    };
}

// ---------------------------------------------------------------------------------------------------------
// join
// ---------------------------------------------------------------------------------------------------------

/**
 * Binds the target list (scope `list`) or the target range (scope `range`, running to the end of its list) to the
 * numId of `continueFromParagraph`'s list, so the items continue that list's count. Each item keeps its level.
 */
export function joinListParagraphs(xmlDoc, paragraphs, continueFromParagraph, request, context = {}) {
    const generateRedlines = context.generateRedlines !== false;
    const author = context.author;
    const indexOf = typeof context.paragraphIndexOf === 'function' ? context.paragraphIndexOf : () => null;
    const stylesDoc = parseStyles(context.stylesXml);
    const numbering = readNumbering(context.numberingXml);
    const scope = request.scope === 'range' ? 'range' : 'list';

    const continueIndex = indexOf(continueFromParagraph);
    const continueInfo = getParagraphListInfo(continueFromParagraph);
    if (!continueInfo || isTrackedDeletedParagraph(continueFromParagraph)) {
        return refusal(LIST_FORMAT_JOIN_INVALID,
            `list.continueFrom (paragraph ${continueIndex ?? '?'}) is ${continueInfo ? 'a tracked-deleted paragraph' : 'not a list item'}; join continues an existing list.`,
            { reason: continueInfo ? 'deleted' : 'not-a-list', paragraphIndex: continueIndex, targetExcerpt: excerpt(extractCanonicalParagraphText(continueFromParagraph)) },
            { action: 'change_request', hint: 'Point list.continueFrom at a live item of the list the target should continue.', field: 'list.continueFrom' });
    }
    if (paragraphs.includes(continueFromParagraph)) {
        return refusal(LIST_FORMAT_JOIN_INVALID, 'list.continueFrom is inside the target range; it must be an item of an earlier list.',
            { reason: 'inside-target', paragraphIndex: continueIndex }, { action: 'change_request', field: 'list.continueFrom' });
    }

    const range = classifyRange(paragraphs, { stylesDoc, indexOf, mode: 'join' });
    if (!range.ok) return range;
    if (range.numId === continueInfo.numId) {
        return unchanged(`The target already belongs to the list of list.continueFrom (numId ${range.numId}).`, { sourceNumId: range.numId });
    }
    const checkedTarget = checkDefinition(xmlDoc, numbering, range.numId, { stylesDoc, role: 'the target list' });
    if (!checkedTarget.ok) return checkedTarget;
    const checkedContinue = checkDefinition(xmlDoc, numbering, continueInfo.numId, { stylesDoc, role: 'the list.continueFrom list' });
    if (!checkedContinue.ok) return checkedContinue;

    const plan = planScope(xmlDoc, range.numId, range.items, scope, indexOf, 'join');
    if (!plan.ok) return plan;

    // The continued list must end before the first joined item; otherwise its later items would renumber.
    const all = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'p'));
    const firstJoined = all.indexOf(plan.rebind[0].paragraph);
    const continued = listMembers(xmlDoc, continueInfo.numId).filter(member => !member.deleted);
    const after = continued.filter(member => all.indexOf(member.paragraph) > firstJoined);
    if (after.length > 0) {
        const reason = all.indexOf(continueFromParagraph) > firstJoined ? 'after-target' : 'interleaved';
        return refusal(reason === 'after-target' ? LIST_FORMAT_JOIN_INVALID : LIST_FORMAT_SPLIT_UNSUPPORTED,
            reason === 'after-target'
                ? 'list.continueFrom comes after the target; join continues an earlier list.'
                : `The list of list.continueFrom (numId ${continueInfo.numId}) has items after the joined paragraphs (paragraphs ${after.map(member => indexOf(member.paragraph)).join(', ')}); they would renumber.`,
            { reason, numId: continueInfo.numId, paragraphIndexes: after.map(member => indexOf(member.paragraph)) },
            { action: 'change_request', hint: 'Join a list to the list that immediately precedes it.', field: 'list.continueFrom' });
    }
    const undefinedLevel = plan.rebind.find(member => !checkedContinue.definition.levels.has(member.ilvl));
    if (undefinedLevel) {
        return refusal(LIST_FORMAT_DEFINITION_UNSUPPORTED,
            `numId ${continueInfo.numId} (the list.continueFrom list) does not define level ${undefinedLevel.ilvl}, which paragraph ${undefinedLevel.paragraphIndex ?? '?'} uses.`,
            { reason: 'undefined-level', numId: continueInfo.numId, level: undefinedLevel.ilvl },
            { action: 'convert_in_host' });
    }

    const pending = checkPendingSnapshots(plan.rebind, { author, generateRedlines });
    if (!pending.ok) return pending;
    rebindParagraphs(xmlDoc, plan.rebind, continueInfo.numId, { author, generateRedlines });
    return {
        ok: true,
        mode: 'join',
        scope,
        sourceNumId: range.numId,
        numId: continueInfo.numId,
        continueFrom: { paragraphIndex: continueIndex, numId: continueInfo.numId },
        items: plan.rebind.map(item => ({ paragraph: item.paragraph, paragraphIndex: item.paragraphIndex, level: item.ilvl, numId: item.numId, sourceNumId: range.numId, inTarget: item.inTarget })),
        skipped: range.skipped,
        numberingXml: null,
        allocatedNumIds: []
    };
}
