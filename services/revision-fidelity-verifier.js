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
export const DEFAULT_REVISION_FIDELITY_MODE = 'enforce';
const MAX_REPORTED_VIOLATIONS = 5;
const EXCERPT_LENGTH = 120;

/**
 * Per-invariant explanation and retry guidance. Agents read `explanation` to understand the
 * failure and `retry` to choose a different operation; the same arguments always fail again.
 */
export const REVISION_FIDELITY_INVARIANTS = Object.freeze({
    REJECT_PARAGRAPH_COUNT: {
        explanation: 'Reject All of the output has a different number of paragraphs than the source: a paragraph was created or removed without a tracked paragraph mark.',
        retry: 'Target one paragraph at a time; add new paragraphs with a separate insertion instead of folding them into a replacement; do not merge or split paragraphs in the same operation as a text edit.'
    },
    REJECT_PARAGRAPH_TEXT: {
        explanation: 'Reject All of the output shows different paragraph text than the source: text, a tab or a manual line break changed without being tracked, or was reordered.',
        retry: 'Use replacements (find/replace) for a small literal change, keep manual line breaks and tabs where they are, and retry with the exact current paragraph text from a fresh extraction.'
    },
    REJECT_PARAGRAPH_PROPS: {
        explanation: 'Reject All of the output shows different paragraph properties (style, alignment, spacing, indentation) than the source: a property changed without a tracked w:pPrChange.',
        retry: 'Change paragraph formatting with a separate paragraph-format operation and keep the text edit free of style or list changes.'
    },
    REJECT_NUMBERING_REFS: {
        explanation: 'Reject All of the output puts a source paragraph in a different list or list level (numId/ilvl) than the source, so the restored document would number differently.',
        retry: 'Edit list item text without changing list markers or indentation, and make list membership or level changes in their own operation.'
    },
    ACCEPT_NEIGHBORS: {
        explanation: 'An untracked edit changed a paragraph outside its target, so the change could not be reviewed or undone.',
        retry: 'Re-extract the target and retry against exactly that paragraph; if the edit must span several paragraphs, use a tracked operation with an explicit targetEnd range.'
    },
    REJECT_WRAPPERS: {
        explanation: 'Reject All of the output moves text into or out of a hyperlink, field or content control compared with the source.',
        retry: 'Keep the change entirely inside or entirely outside the hyperlink, field or content control, for example with replacements that do not cross its boundary.'
    }
});

function excerpt(value) {
    if (value == null) return null;
    const text = typeof value === 'string' ? value : JSON.stringify(value);
    return text.length > EXCERPT_LENGTH ? `${text.slice(0, EXCERPT_LENGTH - 3)}...` : text;
}

function describeRecord(record) {
    return record ? {
        sourceParagraphIndex: record.sourceIndex + 1,
        container: record.container,
        text: excerpt(record.text)
    } : null;
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
 * @param {'rejected'|'accepted'} [input.baselineView='rejected'] - 'accepted' for policies that
 *   accept existing revisions in the target before editing: the changed region is compared with
 *   the source's accepted view, everything else with its rejected view
 * @returns {{ ok: boolean, violations?: Array<{ invariant: string, paragraph: number, expected: string|null,
 *   actual: string|null, explanation: string, source: Object|null, output: Object|null }>,
 *   afterProjection: Array }} `afterProjection` lets a session reuse it as the next operation's baseline
 */
export function verifyRevisionFidelity({ before, after, baselineView = 'rejected', acceptedScope = null }) {
    const actual = Array.isArray(after) ? after : projectRejectedView(after);
    const rejectedBefore = Array.isArray(before) ? before : projectRejectedView(before);
    const expected = baselineView === 'accepted' && !Array.isArray(before)
        ? spliceAcceptedRegion(rejectedBefore, projectWordRevisionView(before, 'accepted'), actual, acceptedScope)
        : rejectedBefore;
    const violations = [];
    const push = (invariant, paragraph, expectedValue, actualValue, sourceRecord, outputRecord) => {
        violations.push({
            invariant,
            paragraph,
            expected: excerpt(expectedValue),
            actual: excerpt(actualValue),
            explanation: REVISION_FIDELITY_INVARIANTS[invariant].explanation,
            source: describeRecord(sourceRecord),
            output: describeRecord(outputRecord)
        });
    };

    if (expected.length !== actual.length) {
        let first = 0;
        while (first < expected.length && first < actual.length && sameRecord(expected[first], actual[first])) first++;
        push('REJECT_PARAGRAPH_COUNT', first + 1, `${expected.length} paragraphs`, `${actual.length} paragraphs`,
            expected[first] || expected[expected.length - 1], actual[first] || actual[actual.length - 1]);
        return { ok: false, violations, afterProjection: actual };
    }

    for (let i = 0; i < expected.length && violations.length < MAX_REPORTED_VIOLATIONS; i++) {
        const source = expected[i];
        const output = actual[i];
        const paragraph = i + 1;
        if (source.container !== output.container) {
            push('REJECT_PARAGRAPH_COUNT', paragraph, `paragraph in ${source.container}`, `paragraph in ${output.container}`, source, output);
        } else if (source.text !== output.text) {
            push('REJECT_PARAGRAPH_TEXT', paragraph, source.text, output.text, source, output);
        } else if ((source.numPr?.numId ?? null) !== (output.numPr?.numId ?? null)
            || (source.numPr?.ilvl ?? null) !== (output.numPr?.ilvl ?? null)) {
            push('REJECT_NUMBERING_REFS', paragraph, source.numPr, output.numPr, source, output);
        } else if (source.props !== output.props) {
            push('REJECT_PARAGRAPH_PROPS', paragraph, source.props, output.props, source, output);
        } else if (source.wrappers.join('|') !== output.wrappers.join('|')) {
            push('REJECT_WRAPPERS', paragraph, source.wrappers, output.wrappers, source, output);
        }
    }
    return violations.length
        ? { ok: false, violations, afterProjection: actual }
        : { ok: true, afterProjection: actual };
}

const sameRecord = (a, b) => a.text === b.text && a.props === b.props && a.container === b.container
    && (a.numPr?.numId ?? null) === (b.numPr?.numId ?? null) && (a.numPr?.ilvl ?? null) === (b.numPr?.ilvl ?? null)
    && a.wrappers.join('|') === b.wrappers.join('|');

/**
 * Untracked operations (generateRedlines: false) have no Reject All to protect, so their
 * guarantee is locality: every paragraph before and after the single resolved target is
 * unchanged in both the Accept All and Reject All views.
 *
 * @param {Object} input
 * @param {Document} input.before - savepoint DOM
 * @param {Document} input.after - post-mutation DOM
 * @param {number} input.targetSourceIndex - zero-based body paragraph position of the target in `before`
 */
export function verifyUntrackedLocality({ before, after, targetSourceIndex }) {
    const violations = [];
    for (const view of ['accepted', 'rejected']) {
        const expected = projectWordRevisionView(before, view);
        const actual = projectWordRevisionView(after, view);
        const targetPosition = expected.findIndex(record => record.sourceIndex >= targetSourceIndex);
        if (targetPosition < 0) continue;
        const needSuffix = expected.length - targetPosition - 1;
        let prefix = 0;
        while (prefix < targetPosition && prefix < actual.length && sameRecord(expected[prefix], actual[prefix])) prefix++;
        let suffix = 0;
        while (suffix < needSuffix && suffix < actual.length - prefix
            && sameRecord(expected[expected.length - 1 - suffix], actual[actual.length - 1 - suffix])) suffix++;
        if (prefix < targetPosition) {
            violations.push(neighborViolation(prefix + 1, expected[prefix], actual[prefix], view));
        } else if (suffix < needSuffix) {
            const index = expected.length - 1 - suffix;
            violations.push(neighborViolation(index + 1, expected[index], actual[actual.length - 1 - suffix], view));
        }
        if (violations.length) break;
    }
    return violations.length ? { ok: false, violations } : { ok: true };
}

function neighborViolation(paragraph, sourceRecord, outputRecord, view) {
    return {
        invariant: 'ACCEPT_NEIGHBORS',
        paragraph,
        view,
        expected: excerpt(sourceRecord?.text ?? null),
        actual: excerpt(outputRecord?.text ?? null),
        explanation: REVISION_FIDELITY_INVARIANTS.ACCEPT_NEIGHBORS.explanation,
        source: describeRecord(sourceRecord),
        output: describeRecord(outputRecord)
    };
}

/**
 * Accept-first policies accept only the target's revisions. With a known target scope
 * (`{ from, to }`, zero-based body paragraph positions) those source paragraphs take the accepted
 * view and all others the rejected view. Without one, the unchanged prefix and suffix keep the
 * rejected view and the paragraphs between them take the accepted view (weaker: a lost revision
 * adjacent to the target is absorbed).
 */
function spliceAcceptedRegion(rejected, accepted, actual, scope) {
    if (scope && Number.isInteger(scope.from) && Number.isInteger(scope.to)) {
        const inScope = record => record.sourceIndex >= scope.from && record.sourceIndex <= scope.to;
        return [...rejected.filter(record => !inScope(record)), ...accepted.filter(inScope)]
            .sort((a, b) => a.sourceIndex - b.sourceIndex);
    }
    let prefix = 0;
    while (prefix < rejected.length && prefix < actual.length && sameRecord(rejected[prefix], actual[prefix])) prefix++;
    let suffix = 0;
    while (suffix < rejected.length - prefix && suffix < actual.length - prefix
        && sameRecord(rejected[rejected.length - 1 - suffix], actual[actual.length - 1 - suffix])) suffix++;
    if (prefix + suffix === rejected.length && prefix + suffix === actual.length) return rejected;
    const low = prefix > 0 ? rejected[prefix - 1].sourceIndex : -1;
    const high = suffix > 0 ? rejected[rejected.length - suffix].sourceIndex : Infinity;
    const middle = accepted.filter(record => record.sourceIndex > low && record.sourceIndex < high);
    return [...rejected.slice(0, prefix), ...middle, ...rejected.slice(rejected.length - suffix)];
}

/**
 * Builds the structured error payload for a violation result. The message alone is enough for a
 * log line; `violations` and `recovery` carry the machine-readable detail for a retry.
 */
export function buildRevisionFidelityError(violations, context = {}) {
    const first = violations[0];
    const where = first.source
        ? `source paragraph ${first.source.sourceParagraphIndex}${first.source.text ? ` ("${first.source.text}")` : ''}`
        : `paragraph ${first.paragraph} of the Reject All view`;
    const invariants = [...new Set(violations.map(violation => violation.invariant))];
    const mode = context.mode || DEFAULT_REVISION_FIDELITY_MODE;
    return {
        code: 'REVISION_FIDELITY_VIOLATION',
        stage: 'verification',
        message: `${mode === 'warn' ? 'Applied with a fidelity warning' : 'Refused, document unchanged'}: `
            + (first.invariant === 'ACCEPT_NEIGHBORS'
                ? `this untracked operation changed a paragraph outside its target (${first.invariant} in the ${first.view} view at ${where}: `
                : `Reject All of this operation's output would not restore the source (${first.invariant} at ${where}: `)
            + `expected ${JSON.stringify(first.expected)}, got ${JSON.stringify(first.actual)}). ${first.explanation} `
            + `Retry: ${REVISION_FIDELITY_INVARIANTS[first.invariant].retry}`,
        violations,
        invariants,
        verificationMode: mode,
        ...(context.baselineView === 'accepted' ? { baselineView: 'accepted' } : {}),
        recovery: {
            hint: REVISION_FIDELITY_INVARIANTS[first.invariant].retry,
            retryGuidance: invariants.map(invariant => ({ invariant, retry: REVISION_FIDELITY_INVARIANTS[invariant].retry })),
            reportIfNarrowRetryFails: 'If a narrower operation on the same paragraph is also refused, the shape is unsupported: report the violations with the operation; do not disable verification to force it through.'
        }
    };
}

/** Baseline view for an existing-revisions policy: accept-first policies edit the accepted source. */
export function revisionFidelityBaselineView(existingRevisions) {
    return String(existingRevisions || '').startsWith('accept-all-first') ? 'accepted' : 'rejected';
}

/**
 * Resolves the effective mode. An explicit option wins; otherwise the diagnostic
 * DOCX_REDLINE_REVISION_FIDELITY environment variable (test lanes); otherwise 'enforce'.
 */
export function resolveRevisionFidelityMode(value) {
    if (REVISION_FIDELITY_MODES.includes(value)) return value;
    const fromEnvironment = globalThis.process?.env?.DOCX_REDLINE_REVISION_FIDELITY;
    return REVISION_FIDELITY_MODES.includes(fromEnvironment) ? fromEnvironment : DEFAULT_REVISION_FIDELITY_MODE;
}
