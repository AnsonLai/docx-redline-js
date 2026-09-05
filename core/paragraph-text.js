import { NS_W } from './types.js';

function localName(node) {
    return String(node?.localName || node?.nodeName || '').replace(/^.*:/, '');
}

/** Returns whether a node contributes to the selected revision view. */
export function isNodeVisibleInRevisionView(node, boundary = null, revisionView = 'accepted') {
    const view = revisionView === 'current' ? 'accepted' : revisionView;
    let cursor = node;
    while (cursor && cursor !== boundary) {
        const name = localName(cursor);
        if (view === 'accepted' && (name === 'del' || name === 'moveFrom')) return false;
        if (view === 'rejected' && (name === 'ins' || name === 'moveTo')) return false;
        cursor = cursor.parentNode;
    }
    return true;
}

/** Reads exact visible text from a run, including tabs, breaks and non-breaking hyphens. */
export function readCanonicalRunText(run, options = {}) {
    const revisionView = options.revisionView === 'current' ? 'accepted' : (options.revisionView || 'accepted');
    const boundary = options.boundary || null;
    if (!isNodeVisibleInRevisionView(run, boundary, revisionView)) return '';
    let text = '';
    const visit = node => {
        for (const child of Array.from(node?.childNodes || [])) {
            if (child?.nodeType !== 1 || (child.namespaceURI && child.namespaceURI !== NS_W)) continue;
            if (!isNodeVisibleInRevisionView(child, boundary, revisionView)) continue;
            const name = localName(child);
            if (name === 't' || (name === 'delText' && revisionView === 'rejected')) text += child.textContent || '';
            else if (name === 'tab') text += '\t';
            else if (name === 'br' || name === 'cr') text += '\n';
            else if (name === 'noBreakHyphen') text += '\u2011';
            else if (name === 'softHyphen') text += '\u00ad';
            else visit(child);
        }
    };
    visit(run);
    return text;
}

const attr = (node, name) => node?.getAttribute?.(`w:${name}`) || node?.getAttribute?.(name) || '';

/**
 * Extracts revision-view text segments from a paragraph or container node.
 *
 * Emits structural text pieces tracking revision ancestry (ins, del, moveFrom, moveTo),
 * computes independent accepted and rejected UTF-16 cursor offsets, and merges
 * adjacent pieces with compatible carrier containers and revision metadata.
 *
 * @param {Element | null | undefined} paragraph - Paragraph or container element
 * @param {Object} [options={}] - Extraction options
 * @param {boolean} [options.mergeRuns=true] - Merge adjacent runs sharing compatible container and revision metadata
 * @returns {Array<{ text: string, kind: 'baseline' | 'insertion' | 'deletion' | 'move_from' | 'move_to', author?: string, revisionId?: string, acceptedStart: number | null, rejectedStart: number | null }>}
 */
export function extractParagraphRevisionSegments(paragraph, options = {}) {
    if (!paragraph) return [];

    const rawPieces = [];

    function walk(node, currentRevision, currentCarrier) {
        for (const child of Array.from(node?.childNodes || [])) {
            if (child?.nodeType !== 1) continue;
            if (child.namespaceURI && child.namespaceURI !== NS_W) continue;

            const name = localName(child);
            if (name === 'pPr' || name === 'rPr') continue;

            let nextRevision = currentRevision;
            if (name === 'ins') {
                nextRevision = {
                    kind: 'insertion',
                    author: attr(child, 'author') || undefined,
                    revisionId: attr(child, 'id') || undefined
                };
            } else if (name === 'del') {
                nextRevision = {
                    kind: 'deletion',
                    author: attr(child, 'author') || undefined,
                    revisionId: attr(child, 'id') || undefined
                };
            } else if (name === 'moveFrom') {
                nextRevision = {
                    kind: 'move_from',
                    author: attr(child, 'author') || undefined,
                    revisionId: attr(child, 'id') || undefined
                };
            } else if (name === 'moveTo') {
                nextRevision = {
                    kind: 'move_to',
                    author: attr(child, 'author') || undefined,
                    revisionId: attr(child, 'id') || undefined
                };
            }

            let nextCarrier = currentCarrier;
            if (name === 'r') {
                nextCarrier = child;
            }

            let text = null;
            let kind = nextRevision ? nextRevision.kind : 'baseline';
            let author = nextRevision?.author;
            let revisionId = nextRevision?.revisionId;

            if (name === 't') {
                text = child.textContent || '';
            } else if (name === 'delText') {
                text = child.textContent || '';
                if (!nextRevision) {
                    kind = 'deletion';
                }
            } else if (name === 'tab') {
                text = '\t';
            } else if (name === 'br' || name === 'cr') {
                text = '\n';
            } else if (name === 'noBreakHyphen') {
                text = '\u2011';
            } else if (name === 'softHyphen') {
                text = '\u00ad';
            }

            if (text !== null) {
                if (text.length > 0) {
                    rawPieces.push({
                        text,
                        kind,
                        author,
                        revisionId,
                        carrier: nextCarrier || child,
                        carrierContainer: (nextCarrier || child)?.parentNode || null
                    });
                }
            } else {
                walk(child, nextRevision, nextCarrier);
            }
        }
    }

    const initialCarrier = localName(paragraph) === 'r' ? paragraph : null;
    walk(paragraph, null, initialCarrier);

    if (rawPieces.length === 0) return [];

    const mergeRuns = options.mergeRuns !== false;
    const merged = [];
    let current = null;

    for (const piece of rawPieces) {
        if (!current) {
            current = { ...piece };
            continue;
        }

        const sameKind = current.kind === piece.kind;
        const sameAuthor = current.author === piece.author;
        const sameRevisionId = current.revisionId === piece.revisionId;
        const sameCarrier = current.carrier === piece.carrier;
        const compatibleContainer = mergeRuns && current.carrierContainer === piece.carrierContainer;

        if (sameKind && sameAuthor && sameRevisionId && (sameCarrier || compatibleContainer)) {
            current.text += piece.text;
        } else {
            merged.push(current);
            current = { ...piece };
        }
    }
    if (current) merged.push(current);

    let acceptedCursor = 0;
    let rejectedCursor = 0;
    const segments = [];

    for (const item of merged) {
        const isAccepted = item.kind !== 'deletion' && item.kind !== 'move_from';
        const isRejected = item.kind !== 'insertion' && item.kind !== 'move_to';

        const acceptedStart = isAccepted ? acceptedCursor : null;
        const rejectedStart = isRejected ? rejectedCursor : null;

        if (isAccepted) acceptedCursor += item.text.length;
        if (isRejected) rejectedCursor += item.text.length;

        const segment = {
            text: item.text,
            kind: item.kind,
            acceptedStart,
            rejectedStart
        };
        if (item.author !== undefined) segment.author = item.author;
        if (item.revisionId !== undefined) segment.revisionId = item.revisionId;

        segments.push(segment);
    }

    return segments;
}

/** Canonical exact paragraph text used by inspection, targeting and ingestion. */
export function extractCanonicalParagraphText(paragraph, options = {}) {
    if (!paragraph) return '';
    const view = options.revisionView === 'current' ? 'accepted' : (options.revisionView || 'accepted');
    const segments = extractParagraphRevisionSegments(paragraph, options);
    if (view === 'rejected') {
        return segments.filter(s => s.rejectedStart !== null).map(s => s.text).join('');
    }
    return segments.filter(s => s.acceptedStart !== null).map(s => s.text).join('');
}
