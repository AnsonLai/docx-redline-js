import { projectWordRevisionView } from '../core/word-revision-semantics.js';

/**
 * Post-apply revision fidelity invariant.
 *
 * A tracked operation only adds revisions on top of the source, so Word's Reject All of
 * the output must equal Word's Reject All of the input: same paragraphs in the same
 * order, same text (breaks and tabs included), same properties and numbering, and the
 * same inline wrapper membership (hyperlinks, simple fields, content controls). Both
 * sides reject every author's revisions, which keeps other reviewers' pre-existing
 * revisions comparable without filtering by revision id.
 *
 * Works on live DOMs and never serializes.
 */

export const REVISION_FIDELITY_MODES = Object.freeze(['enforce', 'warn', 'off']);
const MAX_REPORTED_VIOLATIONS = 5;

function preview(value) {
    const text = JSON.stringify(value ?? null);
    return text.length > 120 ? `${text.slice(0, 117)}...` : text;
}

/** Projects the Reject All view of a document DOM; exposed so callers can cache it. */
export function projectRejectedView(documentNode) {
    return projectWordRevisionView(documentNode, 'rejected');
}

/**
 * Compares two Reject All projections.
 *
 * @param {Object} input
 * @param {Array|Document} input.before - savepoint DOM or its cached projection
 * @param {Array|Document} input.after - post-mutation DOM or its projection
 * @param {'rejected'|'accepted'} [input.baselineView='rejected'] - view of `before` that Reject All
 *   of `after` must reproduce; 'accepted' for policies that accept existing revisions first
 * @returns {{ ok: boolean, violations?: Array<{ invariant: string, paragraph: number, expected: string, actual: string }>,
 *   afterProjection: Array }} `afterProjection` lets a session reuse it as the next operation's baseline
 */
export function verifyRevisionFidelity({ before, after, baselineView = 'rejected' }) {
    const expected = Array.isArray(before) ? before : projectWordRevisionView(before, baselineView);
    const actual = Array.isArray(after) ? after : projectRejectedView(after);
    const violations = [];
    const push = (invariant, paragraph, expectedValue, actualValue) => {
        violations.push({ invariant, paragraph, expected: preview(expectedValue), actual: preview(actualValue) });
    };

    if (expected.length !== actual.length) {
        let first = 0;
        while (first < expected.length && first < actual.length
            && expected[first].text === actual[first].text) first++;
        push('REJECT_PARAGRAPH_COUNT', first + 1, `${expected.length} paragraphs`, `${actual.length} paragraphs`);
        return { ok: false, violations, afterProjection: actual };
    }

    for (let i = 0; i < expected.length && violations.length < MAX_REPORTED_VIOLATIONS; i++) {
        const source = expected[i];
        const output = actual[i];
        const paragraph = i + 1;
        if (source.container !== output.container) {
            push('REJECT_PARAGRAPH_COUNT', paragraph, source.container, output.container);
        } else if (source.text !== output.text) {
            push('REJECT_PARAGRAPH_TEXT', paragraph, source.text, output.text);
        } else if ((source.numPr?.numId ?? null) !== (output.numPr?.numId ?? null)
            || (source.numPr?.ilvl ?? null) !== (output.numPr?.ilvl ?? null)) {
            push('REJECT_NUMBERING_REFS', paragraph, source.numPr, output.numPr);
        } else if (source.props !== output.props) {
            push('REJECT_PARAGRAPH_PROPS', paragraph, source.props, output.props);
        } else if (source.wrappers.join('|') !== output.wrappers.join('|')) {
            push('REJECT_WRAPPERS', paragraph, source.wrappers, output.wrappers);
        }
    }
    return violations.length
        ? { ok: false, violations, afterProjection: actual }
        : { ok: true, afterProjection: actual };
}

/** Builds the structured error payload for a violation result. */
export function buildRevisionFidelityError(violations) {
    const first = violations[0];
    return {
        code: 'REVISION_FIDELITY_VIOLATION',
        stage: 'verification',
        message: `Operation would not restore the source on Reject All (${first.invariant} at paragraph ${first.paragraph}: expected ${first.expected}, got ${first.actual}).`,
        violations
    };
}

/** Baseline view for an existing-revisions policy: accept-first policies edit the accepted source. */
export function revisionFidelityBaselineView(existingRevisions) {
    return String(existingRevisions || '').startsWith('accept-all-first') ? 'accepted' : 'rejected';
}

/**
 * Resolves the effective mode. An explicit option wins; otherwise the diagnostic
 * DOCX_REDLINE_REVISION_FIDELITY environment variable (test lanes), then the rollout
 * default 'warn'.
 */
export function resolveRevisionFidelityMode(value) {
    if (REVISION_FIDELITY_MODES.includes(value)) return value;
    const fromEnvironment = globalThis.process?.env?.DOCX_REDLINE_REVISION_FIDELITY;
    return REVISION_FIDELITY_MODES.includes(fromEnvironment) ? fromEnvironment : 'warn';
}
