import './setup-xml-provider.mjs';

import assert from 'assert/strict';
import { applyRedlineToOxml } from '../engine/oxml-engine.js';
import { wordsToChars, charsToWords } from '../pipeline/diff-engine.js';
import { preprocessMarkdown } from '../pipeline/markdown-processor.js';
import { diff_match_patch } from 'diff-match-patch';
import {
    assertRunFormat,
    assertRunFormatDisabled
} from './helpers/ooxml-assertions.mjs';

async function testReproFormatting() {
    const initialXmlPartial = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
        <pkg:xmlData>
          <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
              <w:p>
                <w:r>
                  <w:rPr><w:rStyle w:val="Normal"/></w:rPr>
                  <w:t>Hello World</w:t>
                </w:r>
              </w:p>
            </w:body>
          </w:document>
        </pkg:xmlData>
      </pkg:part>
    </pkg:package>`;

    const partial = await applyRedlineToOxml(initialXmlPartial, 'Hello World', 'Hello **World**');
    assert.equal(partial.hasChanges, true);
    assertRunFormat(partial.oxml, 'World', { bold: true });
    assertRunFormat(partial.oxml, 'Hello ', { bold: false });

    const initialXmlOff = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
        <pkg:xmlData>
          <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
              <w:p>
                <w:r>
                  <w:rPr><w:b w:val="0"/></w:rPr>
                  <w:t>Hello World</w:t>
                </w:r>
              </w:p>
            </w:body>
          </w:document>
        </pkg:xmlData>
      </pkg:part>
    </pkg:package>`;

    const bold = await applyRedlineToOxml(initialXmlOff, 'Hello World', '**Hello World**');
    assert.equal(bold.hasChanges, true);
    assertRunFormat(bold.oxml, 'Hello World', { bold: true });
}

async function testFormattingSubtraction() {
    const originalOxml = `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r>
                <w:rPr><w:b/></w:rPr>
                <w:t>Bold Text</w:t>
            </w:r>
        </w:p>
    `;
    const result = await applyRedlineToOxml(originalOxml, 'Bold Text', 'Bold Text', {
        author: 'Tester',
        generateRedlines: true
    });

    assert.equal(result.hasChanges, true);
    assertRunFormatDisabled(result.oxml, 'Bold Text', ['b']);

    const originalOxml2 = `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r>
                <w:rPr><w:b/><w:i/></w:rPr>
                <w:t>Bold Italic</w:t>
            </w:r>
        </w:p>
    `;
    const result2 = await applyRedlineToOxml(originalOxml2, 'Bold Italic', '*Bold Italic*', {
        author: 'Tester',
        generateRedlines: true
    });

    assert.equal(result2.hasChanges, true);
    assertRunFormat(result2.oxml, 'Bold Italic', { italic: true });
    assertRunFormatDisabled(result2.oxml, 'Bold Italic', ['b']);
}

function testMarkdownProcessor() {
    const strike = preprocessMarkdown('~~NON-DISCLOSURE AGREEMENT~~');
    assert.equal(strike.cleanText, 'NON-DISCLOSURE AGREEMENT');
    assert.equal(strike.formatHints.length, 1);
    assert.equal(strike.formatHints[0].format.strikethrough, true);

    const nested = preprocessMarkdown('++*NON-DISCLOSURE AGREEMENT*++');
    assert.equal(nested.cleanText, 'NON-DISCLOSURE AGREEMENT');
    assert.equal(nested.formatHints.length, 2);
    assert(nested.formatHints.some(hint => hint.format.underline));
    assert(nested.formatHints.some(hint => hint.format.italic));
}

function testDiffGranularity() {
    const dmp = new diff_match_patch();
    const { chars1, chars2, wordArray } = wordsToChars('British Columbia', 'the State of California');
    const charDiffs = dmp.diff_main(chars1, chars2);
    dmp.diff_cleanupSemantic(charDiffs);
    const wordDiffs = charsToWords(charDiffs, wordArray);

    assert(Array.isArray(wordDiffs));
    assert(wordDiffs.length > 0);
    assert(wordDiffs.some(([op, text]) => op === -1 && text.includes('British')));
    assert(wordDiffs.some(([op, text]) => op === 1 && text.includes('California')));
}

async function testMiddleFormat() {
    const originalOxml = `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r>
                <w:t>The quick brown fox jumps.</w:t>
            </w:r>
        </w:p>
    `;

    const result = await applyRedlineToOxml(
        originalOxml,
        'The quick brown fox jumps.',
        'The **quick** *brown* fox jumps.',
        { author: 'Tester', generateRedlines: true }
    );

    assert.equal(result.hasChanges, true);
    assertRunFormat(result.oxml, 'quick', { bold: true, italic: false });
    assertRunFormat(result.oxml, 'brown', { bold: false, italic: true });
}

await testReproFormatting();
await testFormattingSubtraction();
testMarkdownProcessor();
testDiffGranularity();
await testMiddleFormat();

console.log('formatting_tests.mjs ... PASS');
