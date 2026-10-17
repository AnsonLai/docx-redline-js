/**
 * Compares two `inspectDocumentParts` results by what a reader and reviewer see:
 * paragraph text, table position, list label, review segments and comments.
 *
 * A host that writes a prepared package uses this to verify what it committed
 * against what was prepared. Revision ids, dates, rsids, paragraph ids and
 * run boundaries are ignored, because hosts regenerate them on write.
 */

const commentIdentity = comment => [comment.author ?? null, comment.text ?? '', comment.anchoredText || ''];

// Adjacent review segments of one kind and author are one change, however the
// host split the runs.
function reviewSegments(segments = []) {
    return segments.filter(segment => segment.kind !== 'baseline').reduce((result, segment) => {
        const next = {
            kind: segment.kind, author: segment.author || null, text: segment.text,
            acceptedStart: segment.acceptedStart ?? null, rejectedStart: segment.rejectedStart ?? null
        };
        const previous = result.at(-1);
        const adjacent = key => (previous[key] == null && next[key] == null)
            || (previous[key] != null && previous[key] + previous.text.length === next[key]);
        if (previous && previous.kind === next.kind && previous.author === next.author
            && adjacent('acceptedStart') && adjacent('rejectedStart')) previous.text += next.text;
        else result.push(next);
        return result;
    }, []);
}

/**
 * The host-independent shape of one inspection.
 * @param {{ paragraphs: Array<object>, comments?: Array<object> }} inspection
 */
export function inspectionReviewShape(inspection) {
    const comments = inspection?.comments || [];
    return {
        paragraphs: (inspection?.paragraphs || []).map(paragraph => ({
            text: paragraph.exactText,
            inTable: paragraph.inTable === true,
            table: paragraph.table ? [paragraph.table.rowIndex, paragraph.table.cellIndex] : null,
            list: paragraph.list ? [paragraph.list.format, paragraph.list.level, paragraph.list.label] : null,
            review: {
                present: paragraph.hasRevisions === true,
                authors: [...(paragraph.revisionAuthors || [])].sort(),
                segments: reviewSegments(paragraph.segments)
            }
        })),
        comments: comments.map(comment => {
            const parent = comment.parentCommentId != null
                ? comments.find(item => String(item.id) === String(comment.parentCommentId))
                : null;
            return [...commentIdentity(comment), comment.done === true,
                parent ? commentIdentity(parent)
                    : (comment.parentCommentId != null || comment.parentParaId ? 'unresolved parent' : null)];
        }).map(item => JSON.stringify(item)).sort()
    };
}

/**
 * @param {{ paragraphs: Array<object>, comments?: Array<object> }} expected - e.g. inspection of the prepared package
 * @param {{ paragraphs: Array<object>, comments?: Array<object> }} actual - e.g. inspection of what the host committed
 * @returns {{ matches: boolean, differences: Array<{ path: string, expected: any, actual: any }> }}
 *   `differences` lists at most the first ten mismatches.
 */
export function compareDocumentInspections(expected, actual) {
    const left = inspectionReviewShape(expected);
    const right = inspectionReviewShape(actual);
    const differences = [];
    const note = (path, a, b) => {
        if (differences.length < 10 && JSON.stringify(a) !== JSON.stringify(b)) differences.push({ path, expected: a, actual: b });
    };
    note('paragraphs.length', left.paragraphs.length, right.paragraphs.length);
    const count = Math.max(left.paragraphs.length, right.paragraphs.length);
    for (let index = 0; index < count; index++) {
        const a = left.paragraphs[index] ?? null;
        const b = right.paragraphs[index] ?? null;
        if (!a || !b) { note(`paragraphs[${index}]`, a, b); continue; }
        for (const key of ['text', 'inTable', 'table', 'list', 'review']) note(`paragraphs[${index}].${key}`, a[key], b[key]);
    }
    note('comments', left.comments, right.comments);
    return { matches: differences.length === 0, differences };
}
