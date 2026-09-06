/**
 * Run property (w:rPr) helper utilities.
 *
 * This module owns low-level formatting element operations, including
 * schema-order insertion, format extraction, and format add/remove transforms.
 */

import { createWordElement } from '../core/word-xml.js';

/**
 * Canonical OOXML run-property schema ordering.
 * Shared by all rPr synchronizers.
 */
export const RPR_SCHEMA_ORDER = [
    'w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps',
    'w:strike', 'w:dstrike', 'w:outline', 'w:shadow', 'w:emboss', 'w:imprint', 'w:noProof',
    'w:snapToGrid', 'w:vanish', 'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern',
    'w:position', 'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect', 'w:bdr', 'w:shd',
    'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout',
    'w:specVanish', 'w:oMath', 'w:rPrChange'
];

/**
 * Inserts an rPr child node in schema order.
 *
 * @param {Element} rPr - Run properties element
 * @param {Element} el - Child element to insert
 */
export function insertRPrChildInOrder(rPr, el) {
    const myIndex = RPR_SCHEMA_ORDER.indexOf(el.nodeName);
    const myPriority = myIndex === -1 ? 999 : myIndex;

    let inserted = false;
    for (const child of Array.from(rPr.childNodes)) {
        if (child.nodeType !== 1) continue;
        const childIndex = RPR_SCHEMA_ORDER.indexOf(child.nodeName);
        const childPriority = childIndex === -1 ? 999 : childIndex;
        if (childPriority > myPriority) {
            rPr.insertBefore(el, child);
            inserted = true;
            break;
        }
    }
    if (!inserted) rPr.appendChild(el);
}

/**
 * Shared override routine used by add/remove format transforms.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} rPr - Run properties target
 * @param {{bold?: boolean, italic?: boolean, underline?: boolean, strikethrough?: boolean}} formatFlags - Flags to apply
 * @param {'remove'|'add'} mode - Override mode
 */
function _applyOverrides(xmlDoc, rPr, formatFlags, mode) {
    if (!rPr || !formatFlags) return;

    const applyBold = !!formatFlags.bold;
    const applyItalic = !!formatFlags.italic;
    const applyUnderline = !!formatFlags.underline;
    const applyStrike = !!formatFlags.strikethrough;

    const removalSet = new Set();
    if (applyBold) {
        removalSet.add('w:b');
        removalSet.add('w:bCs');
    }
    if (applyItalic) {
        removalSet.add('w:i');
        removalSet.add('w:iCs');
    }
    if (applyUnderline) removalSet.add('w:u');
    if (applyStrike) removalSet.add('w:strike');

    if (removalSet.size > 0) {
        const toRemove = [];
        for (const child of Array.from(rPr.childNodes)) {
            if (removalSet.has(child.nodeName)) {
                toRemove.push(child);
            }
        }
        for (const el of toRemove) {
            rPr.removeChild(el);
        }
    }

    if (applyBold) {
        const b = createWordElement(xmlDoc, 'w:b');
        b.setAttribute('w:val', mode === 'add' ? '1' : '0');
        insertRPrChildInOrder(rPr, b);

        const bCs = createWordElement(xmlDoc, 'w:bCs');
        bCs.setAttribute('w:val', mode === 'add' ? '1' : '0');
        insertRPrChildInOrder(rPr, bCs);
    }
    if (applyItalic) {
        const i = createWordElement(xmlDoc, 'w:i');
        i.setAttribute('w:val', mode === 'add' ? '1' : '0');
        insertRPrChildInOrder(rPr, i);

        const iCs = createWordElement(xmlDoc, 'w:iCs');
        iCs.setAttribute('w:val', mode === 'add' ? '1' : '0');
        insertRPrChildInOrder(rPr, iCs);
    }
    if (applyUnderline) {
        const u = createWordElement(xmlDoc, 'w:u');
        u.setAttribute('w:val', mode === 'add' ? 'single' : 'none');
        insertRPrChildInOrder(rPr, u);
    }
    if (applyStrike) {
        const strike = createWordElement(xmlDoc, 'w:strike');
        strike.setAttribute('w:val', mode === 'add' ? '1' : '0');
        insertRPrChildInOrder(rPr, strike);
    }
}

/**
 * Builds an rPr XML snippet that explicitly removes formatting while preserving other properties.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} originalRun - Source run
 * @param {Object} formatToRemove - Format flags to remove
 * @param {XMLSerializer} serializer - Serializer instance
 * @returns {string}
 */
export function buildOverrideRPrXml(xmlDoc, originalRun, formatToRemove, serializer) {
    const baseRPr = originalRun.getElementsByTagName('w:rPr')[0] || null;
    const rPr = baseRPr ? baseRPr.cloneNode(true) : createWordElement(xmlDoc, 'w:rPr');
    _applyOverrides(xmlDoc, rPr, formatToRemove, 'remove');

    let rPrXml = serializer.serializeToString(rPr);
    rPrXml = rPrXml.replace(/\s+xmlns:[^=]+="[^"]*"/g, '');
    return rPrXml === '<w:rPr/>' ? '' : rPrXml;
}

/**
 * Removes formatting tags and adds explicit off overrides for the specified flags.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} rPr - Run properties
 * @param {Object} formatToRemove - Format flags to remove
 */
export function applyFormatOverridesToRPr(xmlDoc, rPr, formatToRemove) {
    _applyOverrides(xmlDoc, rPr, formatToRemove, 'remove');
}

/**
 * Extracts format flags from a run properties element.
 *
 * @param {Element|null} rPr - Run properties element
 * @returns {{ bold: boolean, italic: boolean, underline: boolean, strikethrough: boolean, hasFormatting: boolean }}
 */
export function extractFormatFromRPr(rPr) {
    const format = { bold: false, italic: false, underline: false, strikethrough: false, hasFormatting: false };
    if (!rPr) return format;

    for (const child of Array.from(rPr.childNodes)) {
        if (child.nodeName === 'w:b') format.bold = isFormattingElementEnabled(child, false);
        if (child.nodeName === 'w:i') format.italic = isFormattingElementEnabled(child, false);
        if (child.nodeName === 'w:u') format.underline = isFormattingElementEnabled(child, true);
        if (child.nodeName === 'w:strike') format.strikethrough = isFormattingElementEnabled(child, false);

        if (child.nodeName === 'w:rStyle') {
            const styleRef = child.getAttribute('w:val');
            if (styleRef) {
                const lowerStyle = styleRef.toLowerCase();
                if (lowerStyle.includes('bold') || lowerStyle.includes('strong')) format.bold = true;
                if (lowerStyle.includes('italic') || lowerStyle.includes('emphasis')) format.italic = true;
                if (lowerStyle.includes('underline')) format.underline = true;
            }
        }
    }

    format.hasFormatting = format.bold || format.italic || format.underline || format.strikethrough;
    return format;
}

/**
 * Determines whether a formatting element is effectively "on".
 *
 * @param {Element} element - Formatting element
 * @param {boolean} isUnderline - Underline semantic handling
 * @returns {boolean}
 */
function isFormattingElementEnabled(element, isUnderline) {
    const rawValue = element.getAttribute('w:val') || element.getAttribute('val') || '';
    const value = rawValue.toLowerCase();

    if (!value) return true;

    if (isUnderline) {
        return value !== 'none' && value !== '0' && value !== 'false' && value !== 'off';
    }

    return value !== '0' && value !== 'false' && value !== 'off';
}

/**
 * Canonical OOXML paragraph-property schema ordering.
 */
export const PPR_SCHEMA_ORDER = [
    'w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr',
    'w:widowControl', 'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd',
    'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap', 'w:overflowPunct',
    'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd',
    'w:snapToGrid', 'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents',
    'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
    'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange'
];

/**
 * Inserts a pPr child node in schema order.
 *
 * @param {Element} pPr - Paragraph properties element
 * @param {Element} el - Child element to insert
 */
export function insertPPrChildInOrder(pPr, el) {
    const myIndex = PPR_SCHEMA_ORDER.indexOf(el.nodeName);
    const myPriority = myIndex === -1 ? 999 : myIndex;

    let inserted = false;
    for (const child of Array.from(pPr.childNodes)) {
        if (child.nodeType !== 1) continue;
        const childIndex = PPR_SCHEMA_ORDER.indexOf(child.nodeName);
        const childPriority = childIndex === -1 ? 999 : childIndex;
        if (childPriority > myPriority) {
            pPr.insertBefore(el, child);
            inserted = true;
            break;
        }
    }
    if (!inserted) pPr.appendChild(el);
}

/**
 * Checks if applying character properties to an rPr will change anything.
 *
 * @param {Element|null} rPr - Run properties
 * @param {Object} properties - Desired properties
 * @returns {boolean}
 */
export function checkRunPropertiesChanged(rPr, properties) {
    if (!properties) return false;
    const current = extractFormatFromRPr(rPr);

    if (properties.bold !== undefined && current.bold !== !!properties.bold) return true;
    if (properties.italic !== undefined && current.italic !== !!properties.italic) return true;
    if (properties.underline !== undefined && current.underline !== !!properties.underline) return true;

    const desiredStrike = properties.strike !== undefined ? properties.strike : properties.strikethrough;
    if (desiredStrike !== undefined && current.strikethrough !== !!desiredStrike) return true;

    if (properties.highlight !== undefined) {
        const hlEl = Array.from(rPr?.childNodes || []).find(c => c.nodeType === 1 && (c.nodeName === 'w:highlight' || c.localName === 'highlight'));
        const currentHl = hlEl ? (hlEl.getAttribute('w:val') || hlEl.getAttribute('val')) : null;
        const desiredHl = properties.highlight && properties.highlight !== 'none' ? String(properties.highlight).toLowerCase() : null;
        if ((currentHl || null) !== desiredHl) return true;
    }

    if (properties.color !== undefined) {
        const colEl = Array.from(rPr?.childNodes || []).find(c => c.nodeType === 1 && (c.nodeName === 'w:color' || c.localName === 'color'));
        const currentCol = colEl ? (colEl.getAttribute('w:val') || colEl.getAttribute('val')) : null;
        const desiredCol = properties.color && properties.color !== 'auto' ? String(properties.color).replace(/^#/, '').toUpperCase() : null;
        if ((currentCol?.toUpperCase() || null) !== desiredCol) return true;
    }

    if (properties.fontSize !== undefined) {
        const szEl = Array.from(rPr?.childNodes || []).find(c => c.nodeType === 1 && (c.nodeName === 'w:sz' || c.localName === 'sz'));
        const currentSz = szEl ? (szEl.getAttribute('w:val') || szEl.getAttribute('val')) : null;
        const desiredSz = properties.fontSize !== null && properties.fontSize !== '' ? String(properties.fontSize) : null;
        if (currentSz !== desiredSz) return true;
    }

    if (properties.fontFamily !== undefined) {
        const fEl = Array.from(rPr?.childNodes || []).find(c => c.nodeType === 1 && (c.nodeName === 'w:rFonts' || c.localName === 'rFonts'));
        const currentFont = fEl ? (fEl.getAttribute('w:ascii') || fEl.getAttribute('ascii')) : null;
        const desiredFont = properties.fontFamily || null;
        if (currentFont !== desiredFont) return true;
    }

    return false;
}

/**
 * Applies explicit character formatting to an rPr in schema order.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} rPr - Run properties
 * @param {Object} properties - Desired properties
 */
export function applyCharacterFormatToRPr(xmlDoc, rPr, properties) {
    if (!rPr || !properties) return;

    const removeTags = (tags) => {
        const toRemove = [];
        for (const child of Array.from(rPr.childNodes)) {
            if (child.nodeType === 1 && tags.includes(child.nodeName)) {
                toRemove.push(child);
            }
        }
        for (const el of toRemove) rPr.removeChild(el);
    };

    if (properties.bold !== undefined) {
        removeTags(['w:b', 'w:bCs']);
        const b = createWordElement(xmlDoc, 'w:b');
        b.setAttribute('w:val', properties.bold ? '1' : '0');
        insertRPrChildInOrder(rPr, b);
        const bCs = createWordElement(xmlDoc, 'w:bCs');
        bCs.setAttribute('w:val', properties.bold ? '1' : '0');
        insertRPrChildInOrder(rPr, bCs);
    }

    if (properties.italic !== undefined) {
        removeTags(['w:i', 'w:iCs']);
        const i = createWordElement(xmlDoc, 'w:i');
        i.setAttribute('w:val', properties.italic ? '1' : '0');
        insertRPrChildInOrder(rPr, i);
        const iCs = createWordElement(xmlDoc, 'w:iCs');
        iCs.setAttribute('w:val', properties.italic ? '1' : '0');
        insertRPrChildInOrder(rPr, iCs);
    }

    if (properties.underline !== undefined) {
        removeTags(['w:u']);
        const u = createWordElement(xmlDoc, 'w:u');
        u.setAttribute('w:val', properties.underline ? 'single' : 'none');
        insertRPrChildInOrder(rPr, u);
    }

    const strikeVal = properties.strike !== undefined ? properties.strike : properties.strikethrough;
    if (strikeVal !== undefined) {
        removeTags(['w:strike', 'w:dstrike']);
        const strike = createWordElement(xmlDoc, 'w:strike');
        strike.setAttribute('w:val', strikeVal ? '1' : '0');
        insertRPrChildInOrder(rPr, strike);
    }

    if (properties.highlight !== undefined) {
        removeTags(['w:highlight']);
        if (properties.highlight && properties.highlight !== 'none') {
            const hl = createWordElement(xmlDoc, 'w:highlight');
            hl.setAttribute('w:val', String(properties.highlight).toLowerCase());
            insertRPrChildInOrder(rPr, hl);
        }
    }

    if (properties.color !== undefined) {
        removeTags(['w:color']);
        if (properties.color && properties.color !== 'auto') {
            const col = createWordElement(xmlDoc, 'w:color');
            col.setAttribute('w:val', String(properties.color).replace(/^#/, ''));
            insertRPrChildInOrder(rPr, col);
        }
    }

    if (properties.fontSize !== undefined) {
        removeTags(['w:sz', 'w:szCs']);
        if (properties.fontSize !== null && properties.fontSize !== '') {
            const szVal = String(properties.fontSize);
            const sz = createWordElement(xmlDoc, 'w:sz');
            sz.setAttribute('w:val', szVal);
            insertRPrChildInOrder(rPr, sz);
            const szCs = createWordElement(xmlDoc, 'w:szCs');
            szCs.setAttribute('w:val', szVal);
            insertRPrChildInOrder(rPr, szCs);
        }
    }

    if (properties.fontFamily !== undefined) {
        removeTags(['w:rFonts']);
        if (properties.fontFamily) {
            const fonts = createWordElement(xmlDoc, 'w:rFonts');
            fonts.setAttribute('w:ascii', properties.fontFamily);
            fonts.setAttribute('w:hAnsi', properties.fontFamily);
            fonts.setAttribute('w:cs', properties.fontFamily);
            insertRPrChildInOrder(rPr, fonts);
        }
    }
}

/**
 * Checks if applying paragraph properties to a pPr will change anything.
 *
 * @param {Element|null} pPr - Paragraph properties
 * @param {Object} properties - Desired properties
 * @returns {boolean}
 */
export function checkParagraphPropertiesChanged(pPr, properties) {
    if (!properties) return false;

    if (properties.alignment !== undefined) {
        const jc = Array.from(pPr?.childNodes || []).find(c => c.nodeType === 1 && (c.nodeName === 'w:jc' || c.localName === 'jc'));
        const currentVal = jc ? (jc.getAttribute('w:val') || jc.getAttribute('val')) : null;
        if (currentVal !== properties.alignment) return true;
    }

    if (properties.keepNext !== undefined) {
        const kn = Array.from(pPr?.childNodes || []).find(c => c.nodeType === 1 && (c.nodeName === 'w:keepNext' || c.localName === 'keepNext'));
        const currentVal = kn ? isFormattingElementEnabled(kn, false) : false;
        if (currentVal !== !!properties.keepNext) return true;
    }

    if (properties.keepLines !== undefined) {
        const kl = Array.from(pPr?.childNodes || []).find(c => c.nodeType === 1 && (c.nodeName === 'w:keepLines' || c.localName === 'keepLines'));
        const currentVal = kl ? isFormattingElementEnabled(kl, false) : false;
        if (currentVal !== !!properties.keepLines) return true;
    }

    if (properties.pageBreakBefore !== undefined) {
        const pbb = Array.from(pPr?.childNodes || []).find(c => c.nodeType === 1 && (c.nodeName === 'w:pageBreakBefore' || c.localName === 'pageBreakBefore'));
        const currentVal = pbb ? isFormattingElementEnabled(pbb, false) : false;
        if (currentVal !== !!properties.pageBreakBefore) return true;
    }

    if (properties.style !== undefined) {
        const pStyle = Array.from(pPr?.childNodes || []).find(c => c.nodeType === 1 && (c.nodeName === 'w:pStyle' || c.localName === 'pStyle'));
        const currentVal = pStyle ? (pStyle.getAttribute('w:val') || pStyle.getAttribute('val')) : null;
        if (currentVal !== properties.style) return true;
    }

    return false;
}

/**
 * Applies explicit paragraph formatting to a pPr in schema order.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} pPr - Paragraph properties
 * @param {Object} properties - Desired properties
 */
export function applyParagraphPropertiesToPPr(xmlDoc, pPr, properties) {
    if (!pPr || !properties) return;

    const removeTags = (tags) => {
        const toRemove = [];
        for (const child of Array.from(pPr.childNodes)) {
            if (child.nodeType === 1 && tags.includes(child.nodeName)) {
                toRemove.push(child);
            }
        }
        for (const el of toRemove) pPr.removeChild(el);
    };

    if (properties.alignment !== undefined) {
        removeTags(['w:jc']);
        if (properties.alignment) {
            const jc = createWordElement(xmlDoc, 'w:jc');
            jc.setAttribute('w:val', properties.alignment);
            insertPPrChildInOrder(pPr, jc);
        }
    }

    if (properties.keepNext !== undefined) {
        removeTags(['w:keepNext']);
        const kn = createWordElement(xmlDoc, 'w:keepNext');
        if (properties.keepNext) {
            insertPPrChildInOrder(pPr, kn);
        } else {
            kn.setAttribute('w:val', '0');
            insertPPrChildInOrder(pPr, kn);
        }
    }

    if (properties.keepLines !== undefined) {
        removeTags(['w:keepLines']);
        const kl = createWordElement(xmlDoc, 'w:keepLines');
        if (properties.keepLines) {
            insertPPrChildInOrder(pPr, kl);
        } else {
            kl.setAttribute('w:val', '0');
            insertPPrChildInOrder(pPr, kl);
        }
    }

    if (properties.pageBreakBefore !== undefined) {
        removeTags(['w:pageBreakBefore']);
        const pbb = createWordElement(xmlDoc, 'w:pageBreakBefore');
        if (properties.pageBreakBefore) {
            insertPPrChildInOrder(pPr, pbb);
        } else {
            pbb.setAttribute('w:val', '0');
            insertPPrChildInOrder(pPr, pbb);
        }
    }

    if (properties.style !== undefined) {
        removeTags(['w:pStyle']);
        if (properties.style) {
            const pStyle = createWordElement(xmlDoc, 'w:pStyle');
            pStyle.setAttribute('w:val', properties.style);
            insertPPrChildInOrder(pPr, pStyle);
        }
    }
}
