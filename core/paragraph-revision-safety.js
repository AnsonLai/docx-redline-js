const NON_CONTENT_CHILDREN = new Set([
    'pPr',
    'bookmarkStart', 'bookmarkEnd',
    'commentRangeStart', 'commentRangeEnd', 'commentReference',
    'customXmlInsRangeStart', 'customXmlInsRangeEnd',
    'customXmlDelRangeStart', 'customXmlDelRangeEnd',
    'moveFromRangeStart', 'moveFromRangeEnd',
    'moveToRangeStart', 'moveToRangeEnd',
    'permStart', 'permEnd', 'proofErr'
]);

function localNameOf(node) {
    return String(node?.localName || node?.nodeName || '').replace(/^.*:/, '');
}

function directElementChildren(node) {
    return Array.from(node?.childNodes || []).filter(child => child.nodeType === 1);
}

function directChild(node, localName) {
    return directElementChildren(node).find(child => localNameOf(child) === localName) || null;
}

function wordAttribute(node, localName) {
    return node?.getAttribute?.(`w:${localName}`)
        || node?.getAttribute?.(localName)
        || '';
}

function normalizedAuthor(author) {
    return String(author || '').trim().toLowerCase();
}

function isAnchorOnlyRun(node) {
    if (localNameOf(node) !== 'r') return false;
    return directElementChildren(node).every(child => [
        'rPr', 'commentReference',
        'bookmarkStart', 'bookmarkEnd',
        'commentRangeStart', 'commentRangeEnd',
        'proofErr'
    ].includes(localNameOf(child)));
}

function isNonContentChild(node) {
    return NON_CONTENT_CHILDREN.has(localNameOf(node)) || isAnchorOnlyRun(node);
}

function isWhollyDeletedContentNode(node) {
    if (localNameOf(node) === 'del') return true;
    if (!['customXml', 'smartTag', 'sdt', 'sdtContent'].includes(localNameOf(node))) return false;
    const contentChildren = directElementChildren(node).filter(child => (
        !isNonContentChild(child) && localNameOf(child) !== 'sdtPr'
    ));
    return contentChildren.every(isWhollyDeletedContentNode);
}

function paragraphFallsWithinMoveFromRange(paragraph) {
    const root = paragraph?.ownerDocument?.documentElement || null;
    if (!root) return false;
    const openIds = new Set();
    for (const node of [root, ...Array.from(root.getElementsByTagName?.('*') || [])]) {
        if (node === paragraph && openIds.size > 0) return true;
        const name = localNameOf(node);
        const id = wordAttribute(node, 'id');
        if (name === 'moveFromRangeStart' && id !== '') openIds.add(id);
        if (name === 'moveFromRangeEnd' && id !== '') openIds.delete(id);
    }
    return false;
}

function paragraphMarkDeletion(paragraph) {
    const pPr = directChild(paragraph, 'pPr');
    const rPr = directChild(pPr, 'rPr');
    return directChild(rPr, 'del');
}

function hasVisibleInsertionContent(insertion) {
    for (const node of Array.from(insertion?.getElementsByTagName?.('*') || [])) {
        const localName = localNameOf(node);
        if (!['t', 'tab', 'br', 'cr', 'noBreakHyphen', 'softHyphen'].includes(localName)) continue;
        let ancestor = node.parentNode;
        let hidden = false;
        while (ancestor && ancestor !== insertion) {
            const ancestorName = localNameOf(ancestor);
            if (ancestorName === 'del' || ancestorName === 'moveFrom') {
                hidden = true;
                break;
            }
            ancestor = ancestor.parentNode;
        }
        if (hidden) continue;
        if (localName !== 't' || (node.textContent || '').length > 0) return true;
    }
    return false;
}

/**
 * Detects the pre-mutation resurrection target defined by WP08: a paragraph
 * mark deleted by another author with no surviving content in that paragraph.
 */
export function inspectForeignDeletedParagraphTarget(paragraph, mutationAuthor) {
    const markDeletion = paragraphMarkDeletion(paragraph);
    if (!markDeletion) {
        return {
            matches: false,
            hasParagraphMarkDeletion: false,
            foreignParagraphMarkDeletion: false,
            allContentDeleted: false,
            ownerAuthor: null,
            markDeletion: null
        };
    }

    const ownerAuthor = wordAttribute(markDeletion, 'author') || null;
    const foreignParagraphMarkDeletion = !ownerAuthor
        || normalizedAuthor(ownerAuthor) !== normalizedAuthor(mutationAuthor);

    const contentChildren = directElementChildren(paragraph)
        .filter(child => !isNonContentChild(child));
    const allContentDeleted = contentChildren.every(isWhollyDeletedContentNode);
    return {
        matches: foreignParagraphMarkDeletion && allContentDeleted,
        hasParagraphMarkDeletion: true,
        foreignParagraphMarkDeletion,
        allContentDeleted,
        ownerAuthor,
        markDeletion
    };
}

export function getParagraphRestorationRefusal(paragraph) {
    const pPr = directChild(paragraph, 'pPr');
    if (directChild(pPr, 'sectPr')) {
        return {
            code: 'SECTION_BREAK_PARAGRAPH',
            message: 'Refusing to restore a deleted paragraph whose paragraph properties contain a section break.'
        };
    }

    let ancestor = paragraph?.parentNode || null;
    while (ancestor) {
        if (localNameOf(ancestor) === 'moveFrom') {
            return {
                code: 'UNSUPPORTED_MOVE_REVISION',
                message: 'Refusing to restore a paragraph that is part of a pending move-from revision.'
            };
        }
        ancestor = ancestor.parentNode;
    }
    if (
        paragraphFallsWithinMoveFromRange(paragraph)
        ||
        paragraph?.getElementsByTagName?.('*')
        && Array.from(paragraph.getElementsByTagName('*')).some(node => ['moveFrom', 'moveFromRangeStart', 'moveFromRangeEnd'].includes(localNameOf(node)))
    ) {
        return {
            code: 'UNSUPPORTED_MOVE_REVISION',
            message: 'Refusing to restore a paragraph that is part of a pending move-from revision.'
        };
    }

    let row = paragraph?.parentNode || null;
    while (row && localNameOf(row) !== 'tr') row = row.parentNode;
    const rowProperties = directChild(row, 'trPr');
    if (rowProperties && directChild(rowProperties, 'del')) {
        return {
            code: 'UNSAFE_DELETED_TABLE_ROW',
            message: 'Refusing to restore a paragraph inside a table row with a pending row deletion.'
        };
    }

    let sibling = paragraph?.nextSibling || null;
    while (sibling && (sibling.nodeType !== 1 || localNameOf(sibling) !== 'p')) sibling = sibling.nextSibling;
    if (!sibling) {
        return {
            code: 'UNSAFE_PARAGRAPH_PLACEMENT',
            message: 'Refusing to restore a deleted paragraph without a following paragraph in the same structural container.'
        };
    }

    return null;
}

/**
 * Finds already-authored same-paragraph resurrection shapes. This is a
 * warning-only validation predicate because standalone validation has no
 * mutation baseline with which to prove when a foreign insertion was added.
 */
export function findForeignDeletedParagraphResurrections(root) {
    const paragraphs = localNameOf(root) === 'p'
        ? [root]
        : Array.from(root?.getElementsByTagName?.('*') || []).filter(node => localNameOf(node) === 'p');
    const matches = [];

    for (const paragraph of paragraphs) {
        const markDeletion = paragraphMarkDeletion(paragraph);
        if (!markDeletion) continue;
        const ownerAuthor = wordAttribute(markDeletion, 'author') || null;
        const contentChildren = directElementChildren(paragraph)
            .filter(child => !isNonContentChild(child));
        const foreignInsertions = contentChildren.filter(child => {
            if (localNameOf(child) !== 'ins' || !hasVisibleInsertionContent(child)) return false;
            return normalizedAuthor(wordAttribute(child, 'author')) !== normalizedAuthor(ownerAuthor);
        });
        const onlyDeletedContentAndForeignInsertions = contentChildren.every(child => {
            return isWhollyDeletedContentNode(child) || foreignInsertions.includes(child);
        });
        if (foreignInsertions.length > 0 && onlyDeletedContentAndForeignInsertions) {
            matches.push({ paragraph, markDeletion, ownerAuthor, foreignInsertions });
        }
    }
    return matches;
}
