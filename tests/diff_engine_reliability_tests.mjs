import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import { applyRedlineToOxml } from '../index.js';
import {
    MAX_DIFF_TOKENS,
    TOKEN_CODE_POINT_BASE,
    charsToWords,
    computeWordDiffs,
    computeWordLevelDiffOps,
    createDiffEngine,
    wordsToChars
} from '../pipeline/diff-engine.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function uniqueText(count, prefix = 'token') {
    return Array.from({ length: count }, (_, index) => `${prefix}${index}`).join(' ');
}

function recoverOriginal(diffs) {
    return diffs.filter(([op]) => op !== 1).map(([, text]) => text).join('');
}

function recoverModified(diffs) {
    return diffs.filter(([op]) => op !== -1).map(([, text]) => text).join('');
}

function testLeadingWhitespaceAndOffsets() {
    const original = '  indented text';
    const modified = '  indented copy';
    const diffs = computeWordDiffs(original, modified);
    assert.equal(recoverOriginal(diffs), original);
    assert.equal(recoverModified(diffs), modified);
    assert.equal(diffs[0][1], '  indented ');

    const operations = computeWordLevelDiffOps(original, modified);
    assert.deepEqual(operations[0], {
        type: 'equal',
        startOffset: 0,
        endOffset: '  indented '.length,
        text: '  indented '
    });
}

function testCodePointEncodingBoundaries() {
    for (const count of [1000, 65535, 65537, 200000]) {
        const original = uniqueText(count, `n${count}_`);
        const modified = `${original} replacement_${count}`;
        const encoded = wordsToChars(original, modified);

        assert.equal(encoded.chars1.codePointAt(0), TOKEN_CODE_POINT_BASE);
        assert.equal(charsToWords([[0, encoded.chars1]], encoded.wordArray)[0][1], original);

        const diffs = computeWordDiffs(original, modified);
        assert.equal(recoverOriginal(diffs), original, `${count}: original reconstruction`);
        assert.equal(recoverModified(diffs), modified, `${count}: modified reconstruction`);
    }
}

function testUnmappedCodesThrow() {
    assert.throws(
        () => charsToWords([[0, String.fromCodePoint(TOKEN_CODE_POINT_BASE + 1)]], ['only-token']),
        /has no mapping/
    );
    assert.throws(() => charsToWords([[0, '\uD800']], ['only-token']), /has no mapping/);
}

function testDeterministicDefault() {
    assert.equal(createDiffEngine().Diff_Timeout, 0);
    assert.equal(createDiffEngine({ diffTimeoutSeconds: 0.000001 }).Diff_Timeout, 0.000001);

    const originalTokens = Array.from({ length: 50000 }, (_, index) => `d${index}`);
    const modifiedTokens = [...originalTokens];
    modifiedTokens[25000] = 'changed-middle-token';
    const original = originalTokens.join(' ');
    const modified = modifiedTokens.join(' ');
    assert.deepEqual(computeWordDiffs(original, modified), computeWordDiffs(original, modified));
}

async function testStructuredOverflowError() {
    // Repeated spaces consume one token in addition to the unique words, so
    // MAX_DIFF_TOKENS words cross the configured ceiling by exactly one.
    const original = uniqueText(MAX_DIFF_TOKENS, 'overflow');
    const modified = `${original} finalreplacement`;
    const oxml = `<w:p xmlns:w="${NS_W}"><w:r><w:t xml:space="preserve">${original}</w:t></w:r></w:p>`;

    assert.throws(() => wordsToChars(original, modified), error => error?.code === 'DIFF_TOKEN_LIMIT');

    const result = await applyRedlineToOxml(oxml, original, modified, { author: 'Phase2' });
    assert.equal(result.status, 'error');
    assert.equal(result.error?.code, 'DIFF_TOKEN_LIMIT');
    assert.equal(result.hasChanges, false);
    assert.equal(result.oxml, oxml, 'overflow refusal must return caller bytes unchanged');
}

testLeadingWhitespaceAndOffsets();
testCodePointEncodingBoundaries();
testUnmappedCodesThrow();
testDeterministicDefault();
await testStructuredOverflowError();

console.log('PASS: diff engine capacity, whitespace, mapping, and determinism');
