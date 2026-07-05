import assert from 'assert/strict';

import {
    acceptTrackedChangesInOoxml,
    applyRedlineToOxml,
    ingestWordOoxmlToPlainText,
    preprocessMarkdown,
    rejectTrackedChangesInOoxml,
    validateRedlineOoxml
} from '../../index.js';
import {
    assertDelUsesDelText,
    assertNoNestedRevisions,
    assertRevisionMetadata,
    assertSpacePreserved,
    assertUniqueRevisionIds,
    parseXmlFragment
} from './ooxml-assertions.mjs';

export function normalizeVisibleText(text) {
    return String(text || '').replace(/\s+/g, ' ').trim();
}

export function assertRoundTripStructure(xml) {
    assertNoNestedRevisions(xml);
    assertDelUsesDelText(xml);
    assertRevisionMetadata(xml);
    assertUniqueRevisionIds(xml);
    assertSpacePreserved(xml);

    const validation = validateRedlineOoxml(xml);
    const errors = validation.issues.filter(issue => issue.severity === 'error');
    assert.equal(
        errors.length,
        0,
        `validateRedlineOoxml reported errors: ${errors.map(issue => `${issue.code}: ${issue.message}`).join('; ')}`
    );
}

/**
 * Applies a redline, then asserts the accept/reject round-trip invariant.
 *
 * @param {string} oxml - input OOXML (fragment, document, or package scope)
 * @param {string} original - original plain text
 * @param {string} modified - modified text (may contain markdown)
 * @param {object} [options] - options forwarded to applyRedlineToOxml
 * @returns {Promise<{ redlined: object, accepted: object, rejected: object }>}
 */
export async function assertRoundTrip(oxml, original, modified, options = {}) {
    const redlined = await applyRedlineToOxml(oxml, original, modified, {
        generateRedlines: true,
        author: 'RoundTrip',
        ...options
    });

    assert.equal(typeof redlined.oxml, 'string', 'redline result should include OOXML');
    parseXmlFragment(redlined.oxml);
    assertRoundTripStructure(redlined.oxml);

    const accepted = acceptTrackedChangesInOoxml(redlined.oxml, { author: 'RoundTrip' });
    const acceptedText = ingestWordOoxmlToPlainText(accepted.oxml);
    const { cleanText } = preprocessMarkdown(modified);
    assert.equal(
        normalizeVisibleText(acceptedText),
        normalizeVisibleText(cleanText),
        'accepting generated revisions should yield modified text'
    );

    const rejected = rejectTrackedChangesInOoxml(redlined.oxml, { author: 'RoundTrip' });
    const rejectedText = ingestWordOoxmlToPlainText(rejected.oxml);
    assert.equal(
        normalizeVisibleText(rejectedText),
        normalizeVisibleText(original),
        'rejecting generated revisions should yield original text'
    );

    return { redlined, accepted, rejected };
}
