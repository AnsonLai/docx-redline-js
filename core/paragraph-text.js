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

/** Canonical exact paragraph text used by inspection, targeting and ingestion. */
export function extractCanonicalParagraphText(paragraph, options = {}) {
    if (!paragraph) return '';
    const runs = Array.from(paragraph.getElementsByTagNameNS?.(NS_W, 'r') || []);
    return runs.map(run => readCanonicalRunText(run, {
        revisionView: options.revisionView || 'accepted',
        boundary: paragraph
    })).join('');
}
