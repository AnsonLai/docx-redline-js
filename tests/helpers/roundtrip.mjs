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
    assertNoNestedParagraphs,
    assertNoNestedRevisions,
    assertRevisionMetadata,
    assertSectPrLast,
    assertSpacePreserved,
    assertUniqueRevisionIds,
    extractExactVisibleText,
    normalizeParagraphBreaks,
    parseXmlFragment
} from './ooxml-assertions.mjs';

export function normalizeVisibleText(text) {
    return String(text || '').replace(/\s+/g, ' ').trim();
}

/**
 * Runs every structural invariant over generated OOXML.
 *
 * @param {string} xml
 */
export function assertRoundTripStructure(xml) {
    assertNoNestedRevisions(xml);
    assertDelUsesDelText(xml);
    assertRevisionMetadata(xml);
    assertUniqueRevisionIds(xml);
    assertSpacePreserved(xml);
    assertNoNestedParagraphs(xml);
    assertSectPrLast(xml);

    const validation = validateRedlineOoxml(xml);
    const errors = validation.issues.filter(issue => issue.severity === 'error');
    assert.equal(
        errors.length,
        0,
        `validateRedlineOoxml reported errors: ${errors.map(issue => `${issue.code}: ${issue.message}`).join('; ')}`
    );
}

/**
 * Compares resolved text against an expectation at the requested fidelity.
 *
 * 'exact' reads the XML with extractExactVisibleText and normalizes only
 * paragraph separators, so spaces, tabs, and breaks are compared byte-exact.
 * 'normalized' reads with the lossy production reader and collapses all
 * whitespace -- it hides whitespace regressions and is only for cases where
 * markdown preprocessing legitimately rewrites whitespace.
 */
function assertResolvedText(xml, expected, fidelity, label) {
    if (fidelity === 'exact') {
        const actualText = normalizeParagraphBreaks(extractExactVisibleText(xml));
        const expectedText = normalizeParagraphBreaks(expected);
        assert.equal(
            actualText,
            expectedText,
            `${label}\n  expected: ${JSON.stringify(expectedText)}\n  actual:   ${JSON.stringify(actualText)}`
        );
        return;
    }

    assert.equal(
        normalizeVisibleText(ingestWordOoxmlToPlainText(xml)),
        normalizeVisibleText(expected),
        label
    );
}

/**
 * Applies a redline, then asserts the accept/reject round-trip invariant.
 *
 * @param {string} oxml - input OOXML (fragment, document, or package scope)
 * @param {string} original - original plain text
 * @param {string} modified - modified text (may contain markdown)
 * @param {object} [options] - options forwarded to applyRedlineToOxml, plus:
 * @param {'exact'|'normalized'} [options.fidelity='exact'] - text comparison strictness.
 *   Use 'normalized' only when markdown preprocessing legitimately changes
 *   whitespace, and say why at the call site.
 * @returns {Promise<{ redlined: object, accepted: object, rejected: object }>}
 */
export async function assertRoundTrip(oxml, original, modified, options = {}) {
    const { fidelity = 'exact', ...redlineOptions } = options;

    const redlined = await applyRedlineToOxml(oxml, original, modified, {
        generateRedlines: true,
        author: 'RoundTrip',
        ...redlineOptions
    });

    assert.equal(typeof redlined.oxml, 'string', 'redline result should include OOXML');
    parseXmlFragment(redlined.oxml);
    assertRoundTripStructure(redlined.oxml);

    const accepted = acceptTrackedChangesInOoxml(redlined.oxml, { author: 'RoundTrip' });
    const { cleanText } = preprocessMarkdown(modified);
    assertResolvedText(
        accepted.oxml,
        cleanText,
        fidelity,
        'accepting generated revisions should yield modified text'
    );

    const rejected = rejectTrackedChangesInOoxml(redlined.oxml, { author: 'RoundTrip' });
    assertResolvedText(
        rejected.oxml,
        original,
        fidelity,
        'rejecting generated revisions should yield original text'
    );

    return { redlined, accepted, rejected };
}
