import './setup-xml-provider.mjs';

import assert from 'assert/strict';
import { applyHighlightToOoxml } from '../engine/formatting-removal.js';
import {
    directChildByLocalName,
    findRunByText,
    parseXml
} from './helpers/ooxml-assertions.mjs';

function assertHighlight(xml, text, expectedColor) {
    const run = findRunByText(xml, text);
    assert(run, `Expected run with text "${text}"`);

    const rPr = directChildByLocalName(run, 'rPr');
    assert(rPr, `Expected highlighted run "${text}" to have run properties`);

    const highlight = directChildByLocalName(rPr, 'highlight');
    assert(highlight, `Expected run "${text}" to have highlight`);
    assert.equal(highlight.getAttribute('w:val') || highlight.getAttribute('val'), expectedColor);
}

function assertNoHighlight(xml, text) {
    const run = findRunByText(xml, text);
    assert(run, `Expected run with text "${text}"`);

    const rPr = directChildByLocalName(run, 'rPr');
    const highlight = rPr ? directChildByLocalName(rPr, 'highlight') : null;
    assert.equal(highlight, null, `Expected run "${text}" not to be highlighted`);
}

function testApplyHighlightSplitsRun() {
    const originalOoxml = `
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:r>
            <w:t>Hello world</w:t>
        </w:r>
    </w:p>`;

    const result = applyHighlightToOoxml(originalOoxml, 'Hello', 'yellow');
    parseXml(result);

    assertHighlight(result, 'Hello', 'yellow');
    assertNoHighlight(result, ' world');
}

function testHighlightPreservesExistingRunProperties() {
    const ooxmlWithRPr = `
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:r>
            <w:rPr><w:b/></w:rPr>
            <w:t>Hello world</w:t>
        </w:r>
    </w:p>`;

    const result = applyHighlightToOoxml(ooxmlWithRPr, 'Hello', 'green');
    const run = findRunByText(result, 'Hello');
    const rPr = directChildByLocalName(run, 'rPr');

    assertHighlight(result, 'Hello', 'green');
    assert(directChildByLocalName(rPr, 'b'), 'Expected bold property to be preserved');
}

function testHighlightNormalizesColorCase() {
    const originalOoxml = `
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:r><w:t>Hello world</w:t></w:r>
    </w:p>`;

    const result = applyHighlightToOoxml(originalOoxml, 'Hello', 'Cyan');
    assertHighlight(result, 'Hello', 'cyan');
}

function testHighlightWithRedlines() {
    const ooxml = `
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:r>
            <w:t>track changes highlight</w:t>
        </w:r>
    </w:p>`;

    const result = applyHighlightToOoxml(ooxml, 'highlight', 'yellow', {
        generateRedlines: true,
        author: 'TestAuthor'
    });
    const run = findRunByText(result, 'highlight');
    const rPr = directChildByLocalName(run, 'rPr');
    const change = directChildByLocalName(rPr, 'rPrChange');

    assertHighlight(result, 'highlight', 'yellow');
    assert(change, 'Expected highlight change to include tracked rPrChange');
    assert.equal(change.getAttribute('w:author') || change.getAttribute('author'), 'TestAuthor');
}

function testHighlightNoTargetReturnsEquivalentXml() {
    const originalOoxml = `
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:r><w:t>Hello world</w:t></w:r>
    </w:p>`;

    const result = applyHighlightToOoxml(originalOoxml, 'missing', 'yellow');
    parseXml(result);

    assertNoHighlight(result, 'Hello world');
}

function testDefaultNamespaceHighlightOutputStaysParseable() {
    const originalOoxml = `
    <p xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <r><t>Hello world</t></r>
    </p>`;

    const result = applyHighlightToOoxml(originalOoxml, 'world', 'yellow', {
        generateRedlines: true,
        author: 'NamespaceTester'
    });

    parseXml(result);
    assertHighlight(result, 'world', 'yellow');
}

testApplyHighlightSplitsRun();
testHighlightPreservesExistingRunProperties();
testHighlightNormalizesColorCase();
testHighlightWithRedlines();
testHighlightNoTargetReturnsEquivalentXml();
testDefaultNamespaceHighlightOutputStaysParseable();

console.log('highlight_tests.mjs ... PASS');
