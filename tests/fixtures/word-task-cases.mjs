/**
 * English legal/administrative task catalogue for independent Word validation.
 * Keep expectations derived from edit intent rather than engine output.
 */
import { createCommentsPart, createHeaderFooterPart, createNotesPart } from './word-package-parts.mjs';
import { WORD_TASK_COVERAGE } from './word-task-coverage.mjs';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const PRIOR_REVISION_NO_OP_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">A </w:t></w:r>
      <w:del w:id="1" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:del>
      <w:ins w:id="2" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins>
      <w:r><w:t xml:space="preserve"> end</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const HIGH_REVISION_ID_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:ins w:id="2147483000" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Old</w:t></w:r></w:ins>
      <w:r><w:t xml:space="preserve"> clause</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const BOOKMARK_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">This </w:t></w:r>
      <w:bookmarkStart w:id="0" w:name="Agreement"/>
      <w:r><w:t>Agreement</w:t></w:r>
      <w:bookmarkEnd w:id="0"/>
      <w:r><w:t xml:space="preserve"> binds the Seller.</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const INTERNAL_HYPERLINK_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="1" w:name="Definitions"/>
      <w:r><w:t>Definitions</w:t></w:r>
      <w:bookmarkEnd w:id="1"/>
    </w:p>
    <w:p>
      <w:r><w:t xml:space="preserve">See </w:t></w:r>
      <w:hyperlink w:anchor="Definitions"><w:r><w:t>Definitions</w:t></w:r></w:hyperlink>
      <w:r><w:t xml:space="preserve"> within ten days.</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const MIXED_FORMATTING_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Agreement</w:t></w:r>
      <w:r><w:t xml:space="preserve"> requires </w:t></w:r>
      <w:r><w:rPr><w:i/></w:rPr><w:t>notice</w:t></w:r>
      <w:r><w:t xml:space="preserve"> within ten calendar days.</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const CONTENT_CONTROL_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:sdt>
      <w:sdtPr><w:tag w:val="filing-status"/><w:alias w:val="Filing status"/></w:sdtPr>
      <w:sdtContent>
        <w:p><w:r><w:t>The application status is pending.</w:t></w:r></w:p>
      </w:sdtContent>
    </w:sdt>
    <w:sectPr/>
  </w:body>
</w:document>`;

const TABLE_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Agency</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Status</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Finance</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Pending</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>`;

const MIXED_ATTACHMENT_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:t>Date</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const NUMBERED_LIST_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>Approve the agenda.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>Review the financial report.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>Confirm the next meeting date.</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const NESTED_LIST_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="11"/></w:numPr></w:pPr><w:r><w:t>Disclosure obligations.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="11"/></w:numPr></w:pPr><w:r><w:t>Notify affected parties.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="11"/></w:numPr></w:pPr><w:r><w:t>Preserve supporting records.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="11"/></w:numPr></w:pPr><w:r><w:t>Remediation obligations.</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const MIXED_STYLE_NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${NS_W}" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w:abstractNum w:abstractNumId="0" w15:restartNumberingAfterBreak="0">
    <w:nsid w:val="71A1B101"/><w:multiLevelType w:val="hybridMultilevel"/><w:tmpl w:val="71A1B101"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%2."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="1440"/></w:tabs><w:ind w:left="1440" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:lvl>
    <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%3)"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="2160"/></w:tabs><w:ind w:left="2160" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:lvl>
    <w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="2880" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="3600" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="5"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="4320" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="6"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="5040" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="7"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="5760" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="8"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="6480" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1" w15:restartNumberingAfterBreak="0">
    <w:nsid w:val="71A1B102"/><w:multiLevelType w:val="hybridMultilevel"/><w:tmpl w:val="71A1B102"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="1440"/></w:tabs><w:ind w:left="1440" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:lvl>
    <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="2880" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="3600" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="5"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="4320" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="6"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="5040" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="7"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="5760" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="8"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="6480" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="2" w15:restartNumberingAfterBreak="0">
    <w:nsid w:val="71A1B103"/><w:multiLevelType w:val="singleLevel"/><w:tmpl w:val="71A1B103"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="5"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="6"><w:abstractNumId w:val="1"/></w:num>
  <w:num w:numId="7"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>`;

const ROMAN_HIERARCHY_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>Policy foundations</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>Initial review</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>Submit written notice</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>Final review</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>Enforcement</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const ALPHA_DASH_LIST_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="6"/></w:numPr></w:pPr><w:r><w:t>Opening business</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="6"/></w:numPr></w:pPr><w:r><w:t>Roll call</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="6"/></w:numPr></w:pPr><w:r><w:t>Approve prior minutes</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="6"/></w:numPr></w:pPr><w:r><w:t>Reports</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const SYMBOL_BULLET_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>Verify identity</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>Store the audit record</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const SUPPRESSED_HEADING_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Level1"/>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>
        <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/><w:u w:val="single"/></w:rPr>
      </w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/></w:rPr><w:t>A.</w:t></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/></w:rPr><w:tab/></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/><w:u w:val="single"/></w:rPr><w:t>PURPOSE</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const RECONCILIATION_TABLE_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/><w:gridCol w:w="1800"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Item</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Owner</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Status</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Budget review</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Finance</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Pending</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Policy update</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Legal</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Draft</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>`;

const TAB_ALIGNED_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:r><w:t>Department</w:t><w:tab/><w:t>Finance</w:t><w:tab/><w:t>Draft notice</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const BOUNDARY_TAB_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:r><w:tab/><w:t>Indented administrative draft</w:t><w:tab/></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const LOCKED_FIELD_INSTRUCTION = ' PAGE ';
const LOCKED_FIELD_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">See page </w:t></w:r>
      <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
      <w:r><w:instrText xml:space="preserve">${LOCKED_FIELD_INSTRUCTION}</w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>1</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      <w:r><w:t xml:space="preserve"> of this notice.</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const COMMENTED_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">The </w:t></w:r>
      <w:commentRangeStart w:id="0"/>
      <w:r><w:t>Agency decision</w:t></w:r>
      <w:commentRangeEnd w:id="0"/>
      <w:r><w:commentReference w:id="0"/><w:t xml:space="preserve"> is preliminary.</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const COMMENTS_XML = createCommentsPart([{
    id: 0,
    author: 'Administrative Reviewer',
    date: '2026-08-30T00:00:00Z',
    text: 'Confirm the final decision before publication.'
}]);

const FOOTNOTE_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:r><w:t>The filing deadline is Friday</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r><w:r><w:t>.</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const FOOTNOTES_XML = createNotesPart('footnote', [{
    id: 1,
    text: 'Deadlines falling on a holiday move to the next business day.'
}]);

const ENDNOTE_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:r><w:t>The indemnity survives for two years</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="1"/></w:r><w:r><w:t>.</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const ENDNOTES_XML = createNotesPart('endnote', [{
    id: 1,
    text: 'The survival period begins on termination.'
}]);

const HEADER_FOOTER_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
  <w:body>
    <w:p><w:r><w:t>The administrative report is a draft.</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdHeader1"/>
      <w:footerReference w:type="default" r:id="rIdFooter1"/>
    </w:sectPr>
  </w:body>
</w:document>`;

const HEADER_XML = createHeaderFooterPart('header', 'Office of Administrative Review');
const FOOTER_XML = createHeaderFooterPart('footer', 'Confidential working copy');

const EXTERNAL_HYPERLINK_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Review the </w:t></w:r>
      <w:hyperlink r:id="rIdExternalPolicy"><w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>filing policy</w:t></w:r></w:hyperlink>
      <w:r><w:t xml:space="preserve"> before submitting within ten days.</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const WORD_TASK_CASE_DEFINITIONS = [
    {
        name: 'simple-redline',
        category: 'legal',
        task: 'replace-term',
        original: 'The old sentence.',
        modified: 'The new sentence.'
    },
    {
        name: 'legal-defined-term-replacement',
        category: 'legal',
        task: 'replace-defined-term',
        original: 'The Supplier shall retain the Records for three years.',
        modified: 'The Contractor shall retain the Records for three years.'
    },
    {
        name: 'legal-clause-insertion',
        category: 'legal',
        task: 'insert-clause-text',
        original: 'The Receiving Party shall protect Confidential Information.',
        modified: 'The Receiving Party shall use reasonable safeguards to protect Confidential Information.'
    },
    {
        name: 'legal-sentence-deletion',
        category: 'legal',
        task: 'delete-sentence',
        original: 'This Agreement begins on the Effective Date. It renews automatically each year.',
        modified: 'This Agreement begins on the Effective Date.'
    },
    {
        name: 'paragraph-insert',
        category: 'administrative',
        task: 'insert-paragraph',
        original: 'The meeting is called to order.',
        modified: 'The meeting is called to order.\nThe Chair confirms that a quorum is present.'
    },
    {
        name: 'legal-paragraph-deletion',
        category: 'legal',
        task: 'delete-paragraph',
        sourceText: 'The parties agree to the following terms.\nThis paragraph is intentionally removed.',
        original: 'This paragraph is intentionally removed.',
        modified: '',
        expectedAcceptedText: 'The parties agree to the following terms.',
        expectedRejectedText: 'The parties agree to the following terms.\nThis paragraph is intentionally removed.'
    },
    {
        name: 'administrative-deadline-change',
        category: 'administrative',
        task: 'replace-deadline',
        original: 'Applications must be received by Friday at 4:00 p.m.',
        modified: 'Applications must be received by Monday at 12:00 p.m.'
    },
    {
        name: 'administrative-procedure-insertion',
        category: 'administrative',
        task: 'insert-procedure',
        original: 'Submit the completed form to the Clerk.',
        modified: 'Sign and date the completed form, then submit it to the Clerk.'
    },
    {
        name: 'format-only',
        category: 'administrative',
        task: 'apply-bold',
        original: 'Make word bold',
        modified: 'Make **word** bold'
    },
    {
        name: 'legal-defined-term-italic',
        category: 'legal',
        task: 'apply-italic',
        original: 'The term Business Day excludes statutory holidays.',
        modified: 'The term *Business Day* excludes statutory holidays.'
    },
    {
        name: 'administrative-deadline-underline',
        category: 'administrative',
        task: 'apply-underline',
        original: 'Response deadline: September 30.',
        modified: 'Response deadline: ++September 30++.'
    },
    {
        name: 'whitespace-heavy',
        category: 'legal',
        task: 'preserve-significant-spacing',
        original: 'Section  1 applies to the Agency.',
        modified: 'Section  1 applies to the Department.'
    },
    {
        name: 'legal-dollar-delimiters-preserved',
        category: 'legal',
        task: 'preserve-dollar-delimiters',
        original: 'The rate is stated in Schedule A.',
        modified: 'The rate is $X$ per unit as defined in Schedule A.'
    },
    {
        name: 'administrative-literal-escapes-preserved',
        category: 'administrative',
        task: 'preserve-literal-escapes',
        original: 'The filing guide describes supported notation.',
        modified: String.raw`The filing guide preserves literal \n and \r\n notation.`
    },
    {
        name: 'legal-inline-preface-preserved',
        category: 'legal',
        task: 'preserve-inline-preface',
        original: 'This clause is part of the Agreement.',
        modified: 'Here is the text: this clause is part of the actual Agreement.'
    },
    {
        name: 'administrative-multiline-target',
        category: 'administrative',
        task: 'replace-multiline-target',
        sourceText: 'The Clerk records the application.\nThe Director reviews the application.',
        original: 'The Clerk records the application.\nThe Director reviews the application.',
        modified: 'The Clerk records the application.\nThe Director approves the application.'
    },
    {
        name: 'legal-leading-whitespace-preserved',
        category: 'legal',
        task: 'preserve-leading-whitespace',
        original: '  Indented covenant applies to the Seller.',
        modified: '  Indented covenant applies to the Purchaser.'
    },
    {
        name: 'legal-prior-revision-no-op',
        category: 'legal',
        task: 'preserve-prior-revisions-on-no-op',
        sourceDocumentXml: PRIOR_REVISION_NO_OP_DOCUMENT,
        original: 'A new end',
        modified: 'A new end',
        operationOptions: { existingRevisions: 'accept-all-first' },
        expectNoOp: true,
        expectedAcceptedText: 'A new end',
        expectedRejectedText: 'A old end'
    },
    {
        name: 'administrative-atomic-batch-rollback',
        category: 'administrative',
        task: 'rollback-failed-batch',
        sourceDocumentXml: PRIOR_REVISION_NO_OP_DOCUMENT,
        original: 'A new end',
        modified: 'A updated end',
        batchOperations: [
            {
                type: 'replace',
                target: 'A new end',
                modified: 'A updated end'
            },
            {
                type: 'replace',
                target: 'Missing administrative target.',
                modified: 'This operation must fail.'
            }
        ],
        operationOptions: { existingRevisions: 'accept-all-first' },
        expectAtomicRollback: true,
        expectedAcceptedText: 'A new end',
        expectedRejectedText: 'A old end'
    },
    {
        name: 'legal-hostile-revision-id-clamped',
        category: 'legal',
        task: 'clamp-hostile-revision-id',
        sourceDocumentXml: HIGH_REVISION_ID_DOCUMENT,
        original: 'Old clause',
        modified: 'Updated clause',
        operationOptions: { existingRevisions: 'accept-all-first' },
        maxRevisionId: 9999,
        expectedAcceptedText: 'Updated clause',
        expectedRejectedText: 'Old clause'
    },
    {
        name: 'legal-bookmark-adjacent-replacement',
        category: 'legal',
        task: 'replace-adjacent-to-bookmark',
        sourceDocumentXml: BOOKMARK_DOCUMENT,
        original: 'This Agreement binds the Seller.',
        modified: 'This Agreement binds the Purchaser.',
        requiredElements: { bookmarkStart: 1, bookmarkEnd: 1 }
    },
    {
        name: 'legal-internal-hyperlink-adjacent-replacement',
        category: 'legal',
        task: 'replace-adjacent-to-hyperlink',
        sourceDocumentXml: INTERNAL_HYPERLINK_DOCUMENT,
        original: 'See Definitions within ten days.',
        modified: 'See Definitions within fifteen days.',
        expectedAcceptedText: 'Definitions\nSee Definitions within fifteen days.',
        expectedRejectedText: 'Definitions\nSee Definitions within ten days.',
        requiredElements: { hyperlink: 1, bookmarkStart: 1, bookmarkEnd: 1 }
    },
    {
        name: 'legal-mixed-run-formatting-preserved',
        category: 'legal',
        task: 'replace-across-formatted-runs',
        sourceDocumentXml: MIXED_FORMATTING_DOCUMENT,
        original: 'Agreement requires notice within ten calendar days.',
        modified: 'Agreement requires notice within ten business days.',
        requiredElements: { b: 1, i: 1 }
    },
    {
        name: 'administrative-content-control-replacement',
        category: 'administrative',
        task: 'replace-in-content-control',
        sourceDocumentXml: CONTENT_CONTROL_DOCUMENT,
        original: 'The application status is pending.',
        modified: 'The application status is approved.',
        requiredElements: { sdt: 1, sdtPr: 1, sdtContent: 1, tag: 1 }
    },
    {
        name: 'administrative-table-cell-replacement',
        category: 'administrative',
        task: 'replace-table-cell',
        sourceDocumentXml: TABLE_DOCUMENT,
        original: 'Pending',
        modified: 'Approved',
        // Word exposes a paragraph boundary for each cell plus an additional
        // row boundary after the first row once cell markers are removed.
        expectedAcceptedText: 'Agency\nStatus\n\nFinance\nApproved',
        expectedRejectedText: 'Agency\nStatus\n\nFinance\nPending',
        requiredElements: { tbl: 1, tr: 2, tc: 4 }
    },
    {
        name: 'administrative-list-change-append-item',
        category: 'administrative',
        task: 'append-numbered-list-item',
        sourceDocumentXml: NUMBERED_LIST_DOCUMENT,
        original: 'Approve the agenda.',
        modified: '1. Approve the agenda.\n2. Record declarations of interest.',
        operation: {
            type: 'list-change',
            target: 'Approve the agenda.',
            modified: '1. Approve the agenda.\n2. Record declarations of interest.'
        },
        assertionMode: 'contains',
        expectedAcceptedContains: ['Approve the agenda.', 'Record declarations of interest.', 'Review the financial report.'],
        expectedAcceptedAbsent: [],
        expectedRejectedContains: ['Approve the agenda.', 'Review the financial report.'],
        expectedRejectedAbsent: ['Record declarations of interest.'],
        requiredElements: { numPr: 4 }
    },
    {
        name: 'legal-list-change-append-multiple-items',
        category: 'legal',
        task: 'append-multiple-numbered-list-items',
        sourceDocumentXml: NUMBERED_LIST_DOCUMENT,
        original: 'Review the financial report.',
        modified: '1. Review the financial report.\n2. Confirm delegated authority.\n3. Record the approval threshold.',
        operation: {
            type: 'list-change',
            target: 'Review the financial report.',
            modified: '1. Review the financial report.\n2. Confirm delegated authority.\n3. Record the approval threshold.'
        },
        assertionMode: 'contains',
        expectedAcceptedContains: ['Review the financial report.', 'Confirm delegated authority.', 'Record the approval threshold.'],
        expectedAcceptedAbsent: [],
        expectedRejectedContains: ['Review the financial report.', 'Confirm the next meeting date.'],
        expectedRejectedAbsent: ['Confirm delegated authority.', 'Record the approval threshold.'],
        requiredElements: { numPr: 5 }
    },
    {
        name: 'administrative-list-change-nested-child',
        category: 'administrative',
        task: 'insert-nested-list-child',
        sourceDocumentXml: NESTED_LIST_DOCUMENT,
        original: 'Disclosure obligations.',
        modified: '1. Disclosure obligations.\n  - Escalate unresolved notifications.',
        operation: {
            type: 'list-change',
            target: 'Disclosure obligations.',
            modified: '1. Disclosure obligations.\n  - Escalate unresolved notifications.'
        },
        assertionMode: 'contains',
        expectedAcceptedContains: ['Disclosure obligations.', 'Escalate unresolved notifications.', 'Notify affected parties.'],
        expectedAcceptedAbsent: [],
        expectedRejectedContains: ['Disclosure obligations.', 'Notify affected parties.'],
        expectedRejectedAbsent: ['Escalate unresolved notifications.'],
        requiredElements: { numPr: 5 }
    },
    {
        name: 'legal-list-change-middle-range-insertion',
        category: 'legal',
        task: 'insert-between-existing-list-items',
        sourceDocumentXml: NESTED_LIST_DOCUMENT,
        original: 'Notify affected parties.\nPreserve supporting records.',
        modified: '- Notify affected parties.\n- Document the notification date.\n- Preserve supporting records.',
        operation: {
            type: 'list-change',
            target: 'Notify affected parties.\nPreserve supporting records.',
            modified: '- Notify affected parties.\n- Document the notification date.\n- Preserve supporting records.'
        },
        assertionMode: 'contains',
        expectedAcceptedContains: ['Notify affected parties.', 'Document the notification date.', 'Preserve supporting records.'],
        expectedAcceptedAbsent: [],
        expectedRejectedContains: ['Notify affected parties.', 'Preserve supporting records.'],
        expectedRejectedAbsent: ['Document the notification date.'],
        requiredElements: { numPr: 5 }
    },
    {
        name: 'legal-list-change-upper-roman-section',
        category: 'legal',
        task: 'append-upper-roman-section',
        sourceDocumentXml: ROMAN_HIERARCHY_DOCUMENT,
        original: 'Policy foundations',
        modified: 'I. Policy foundations\nII. Definitions and scope',
        operation: { type: 'list-change', target: 'Policy foundations', modified: 'I. Policy foundations\nII. Definitions and scope' },
        assertionMode: 'contains',
        expectedAcceptedContains: ['Policy foundations', 'Definitions and scope', 'Enforcement'],
        expectedAcceptedAbsent: [],
        expectedRejectedContains: ['Policy foundations', 'Enforcement'],
        expectedRejectedAbsent: ['Definitions and scope'],
        packageParts: { numberingXml: MIXED_STYLE_NUMBERING_XML },
        requiredNumberingFormats: ['upperRoman'],
        requiredElements: { numPr: 6 }
    },
    {
        name: 'legal-list-change-lower-roman-subclause',
        category: 'legal',
        task: 'append-lower-roman-subclause',
        sourceDocumentXml: ROMAN_HIERARCHY_DOCUMENT,
        original: 'Initial review',
        modified: 'i. Initial review\nii. Record the reviewer decision',
        operation: { type: 'list-change', target: 'Initial review', modified: 'i. Initial review\nii. Record the reviewer decision' },
        assertionMode: 'contains',
        expectedAcceptedContains: ['Initial review', 'Record the reviewer decision', 'Final review'],
        expectedAcceptedAbsent: [],
        expectedRejectedContains: ['Initial review', 'Final review'],
        expectedRejectedAbsent: ['Record the reviewer decision'],
        packageParts: { numberingXml: MIXED_STYLE_NUMBERING_XML },
        requiredNumberingFormats: ['lowerRoman'],
        requiredElements: { numPr: 6 }
    },
    {
        name: 'legal-list-change-parenthesized-lower-letter',
        category: 'legal',
        task: 'append-parenthesized-lower-letter-item',
        sourceDocumentXml: ROMAN_HIERARCHY_DOCUMENT,
        original: 'Submit written notice',
        modified: '(a) Submit written notice\n(b) Preserve proof of delivery',
        operation: { type: 'list-change', target: 'Submit written notice', modified: '(a) Submit written notice\n(b) Preserve proof of delivery' },
        assertionMode: 'contains',
        expectedAcceptedContains: ['Submit written notice', 'Preserve proof of delivery', 'Final review'],
        expectedAcceptedAbsent: [],
        expectedRejectedContains: ['Submit written notice', 'Final review'],
        expectedRejectedAbsent: ['Preserve proof of delivery'],
        packageParts: { numberingXml: MIXED_STYLE_NUMBERING_XML },
        requiredNumberingFormats: ['lowerLetter'],
        requiredElements: { numPr: 6 }
    },
    {
        name: 'administrative-list-change-upper-letter-agenda',
        category: 'administrative',
        task: 'append-upper-letter-agenda-section',
        sourceDocumentXml: ALPHA_DASH_LIST_DOCUMENT,
        original: 'Opening business',
        modified: 'A. Opening business\nB. Public comment',
        operation: { type: 'list-change', target: 'Opening business', modified: 'A. Opening business\nB. Public comment' },
        assertionMode: 'contains',
        expectedAcceptedContains: ['Opening business', 'Public comment', 'Reports'],
        expectedAcceptedAbsent: [],
        expectedRejectedContains: ['Opening business', 'Reports'],
        expectedRejectedAbsent: ['Public comment'],
        packageParts: { numberingXml: MIXED_STYLE_NUMBERING_XML },
        requiredNumberingFormats: ['upperLetter'],
        requiredElements: { numPr: 5 }
    },
    {
        name: 'administrative-list-change-dash-bullet',
        category: 'administrative',
        task: 'append-dash-bullet',
        sourceDocumentXml: ALPHA_DASH_LIST_DOCUMENT,
        original: 'Roll call',
        modified: '- Roll call\n- Confirm quorum',
        operation: { type: 'list-change', target: 'Roll call', modified: '- Roll call\n- Confirm quorum' },
        assertionMode: 'contains',
        expectedAcceptedContains: ['Roll call', 'Confirm quorum', 'Approve prior minutes'],
        expectedAcceptedAbsent: [],
        expectedRejectedContains: ['Roll call', 'Approve prior minutes'],
        expectedRejectedAbsent: ['Confirm quorum'],
        packageParts: { numberingXml: MIXED_STYLE_NUMBERING_XML },
        requiredNumberingFormats: ['bullet'],
        requiredElements: { numPr: 5 }
    },
    {
        name: 'legal-list-change-symbol-bullet-multiple',
        category: 'legal',
        task: 'append-multiple-symbol-bullets',
        sourceDocumentXml: SYMBOL_BULLET_DOCUMENT,
        original: 'Verify identity',
        modified: '- Verify identity\n- Confirm authority\n- Capture the verification date',
        operation: { type: 'list-change', target: 'Verify identity', modified: '- Verify identity\n- Confirm authority\n- Capture the verification date' },
        assertionMode: 'contains',
        expectedAcceptedContains: ['Verify identity', 'Confirm authority', 'Capture the verification date', 'Store the audit record'],
        expectedAcceptedAbsent: [],
        expectedRejectedContains: ['Verify identity', 'Store the audit record'],
        expectedRejectedAbsent: ['Confirm authority', 'Capture the verification date'],
        packageParts: { numberingXml: MIXED_STYLE_NUMBERING_XML },
        requiredNumberingFormats: ['bullet'],
        requiredElements: { numPr: 4 }
    },
    {
        name: 'legal-suppressed-heading-to-bullet-list',
        category: 'legal',
        task: 'replace-manual-heading-with-bullets',
        sourceDocumentXml: SUPPRESSED_HEADING_DOCUMENT,
        original: 'A.\tPURPOSE',
        modified: '* Article A. Purpose and Interagency Alignment\n* Key Focus: Joint Street Outreach & Medical Triage',
        operation: {
            type: 'list-change',
            target: 'A.\tPURPOSE',
            modified: '* Article A. Purpose and Interagency Alignment\n* Key Focus: Joint Street Outreach & Medical Triage'
        },
        expectedAcceptedText: 'Article A. Purpose and Interagency Alignment\nKey Focus: Joint Street Outreach & Medical Triage',
        expectedRejectedText: 'A.\tPURPOSE',
        packageParts: { numberingXml: MIXED_STYLE_NUMBERING_XML },
        requiredNumberingFormats: ['bullet'],
        requiredElements: { numPr: 3, ins: 4, del: 4 }
    },
    {
        name: 'administrative-table-reconciliation-cell-update',
        category: 'administrative',
        task: 'reconcile-table-cell-update',
        sourceDocumentXml: RECONCILIATION_TABLE_DOCUMENT,
        original: 'Budget review',
        modified: '| Item | Owner | Status |\n| --- | --- | --- |\n| Budget review | Finance | Approved |\n| Policy update | Legal | Draft |',
        operation: {
            type: 'table-reconciliation',
            target: 'Budget review',
            modified: '| Item | Owner | Status |\n| --- | --- | --- |\n| Budget review | Finance | Approved |\n| Policy update | Legal | Draft |'
        },
        assertionMode: 'contains',
        expectedAcceptedContains: ['Budget review', 'Approved', 'Policy update'],
        expectedAcceptedAbsent: ['Pending'],
        expectedRejectedContains: ['Budget review', 'Pending', 'Policy update'],
        expectedRejectedAbsent: ['Approved'],
        requiredElements: { tbl: 1, tr: 3, tc: 9 }
    },
    {
        name: 'legal-table-reconciliation-row-insertion',
        category: 'legal',
        task: 'reconcile-table-row-insertion',
        sourceDocumentXml: RECONCILIATION_TABLE_DOCUMENT,
        original: 'Budget review',
        modified: '| Item | Owner | Status |\n| --- | --- | --- |\n| Budget review | Finance | Pending |\n| Compliance filing | Legal | Scheduled |\n| Policy update | Legal | Draft |',
        operation: {
            type: 'table-reconciliation',
            target: 'Budget review',
            modified: '| Item | Owner | Status |\n| --- | --- | --- |\n| Budget review | Finance | Pending |\n| Compliance filing | Legal | Scheduled |\n| Policy update | Legal | Draft |'
        },
        assertionMode: 'contains',
        expectedAcceptedContains: ['Budget review', 'Compliance filing', 'Scheduled', 'Policy update'],
        expectedAcceptedAbsent: [],
        expectedRejectedContains: ['Budget review', 'Policy update'],
        expectedRejectedAbsent: ['Compliance filing', 'Scheduled'],
        requiredElements: { tbl: 1, tr: 4, tc: 12 }
    },
    {
        name: 'administrative-table-reconciliation-row-deletion',
        category: 'administrative',
        task: 'reconcile-table-row-deletion',
        sourceDocumentXml: RECONCILIATION_TABLE_DOCUMENT,
        original: 'Budget review',
        modified: '| Item | Owner | Status |\n| --- | --- | --- |\n| Budget review | Finance | Pending |',
        operation: {
            type: 'table-reconciliation',
            target: 'Budget review',
            modified: '| Item | Owner | Status |\n| --- | --- | --- |\n| Budget review | Finance | Pending |'
        },
        assertionMode: 'contains',
        expectedAcceptedContains: ['Budget review', 'Finance', 'Pending'],
        expectedAcceptedAbsent: ['Policy update', 'Draft'],
        expectedRejectedContains: ['Budget review', 'Policy update', 'Draft'],
        expectedRejectedAbsent: [],
        requiredElements: { tbl: 1, tr: 3, tc: 9 }
    },
    {
        name: 'legal-table-reconciliation-multi-cell-update',
        category: 'legal',
        task: 'reconcile-multiple-table-cells',
        sourceDocumentXml: RECONCILIATION_TABLE_DOCUMENT,
        original: 'Budget review',
        modified: '| Item | Owner | Status |\n| --- | --- | --- |\n| Budget review | Treasury | Complete |\n| Policy update | Compliance | Approved |',
        operation: {
            type: 'table-reconciliation',
            target: 'Budget review',
            modified: '| Item | Owner | Status |\n| --- | --- | --- |\n| Budget review | Treasury | Complete |\n| Policy update | Compliance | Approved |'
        },
        assertionMode: 'contains',
        expectedAcceptedContains: ['Treasury', 'Complete', 'Compliance', 'Approved'],
        expectedAcceptedAbsent: ['Finance', 'Pending', 'Draft'],
        expectedRejectedContains: ['Finance', 'Pending', 'Legal', 'Draft'],
        expectedRejectedAbsent: ['Treasury', 'Complete', 'Compliance', 'Approved'],
        requiredElements: { tbl: 1, tr: 3, tc: 9 }
    },
    {
        name: 'administrative-tab-aligned-status',
        category: 'administrative',
        task: 'replace-adjacent-to-tabs',
        sourceDocumentXml: TAB_ALIGNED_DOCUMENT,
        original: 'Department\tFinance\tDraft notice',
        modified: 'Department\tFinance\tFinal notice',
        requiredElements: { tab: 2 }
    },
    {
        name: 'administrative-boundary-tabs-preserved',
        category: 'administrative',
        task: 'replace-between-boundary-tabs',
        sourceDocumentXml: BOUNDARY_TAB_DOCUMENT,
        original: '\tIndented administrative draft\t',
        modified: '\tIndented administrative final\t',
        requiredElements: { tab: 2 }
    },
    {
        name: 'legal-locked-field-adjacent-replacement',
        category: 'legal',
        task: 'replace-adjacent-to-complex-field',
        sourceDocumentXml: LOCKED_FIELD_DOCUMENT,
        original: 'See page 1 of this notice.',
        modified: 'See page 1 of this amended notice.',
        requiredElements: { fldChar: 3, instrText: 1 },
        requiredElementParents: { fldChar: 'r', instrText: 'r' },
        requiredElementText: { instrText: [LOCKED_FIELD_INSTRUCTION] }
    },
    {
        name: 'administrative-comment-anchor-adjacent-replacement',
        category: 'administrative',
        task: 'replace-adjacent-to-comment',
        sourceDocumentXml: COMMENTED_DOCUMENT,
        original: 'The Agency decision is preliminary.',
        modified: 'The Agency decision is final.',
        requiredElements: { commentRangeStart: 1, commentRangeEnd: 1, commentReference: 1 },
        packageParts: { commentsXml: COMMENTS_XML }
    },
    {
        name: 'administrative-footnote-adjacent-deadline',
        category: 'administrative',
        task: 'replace-adjacent-to-footnote',
        sourceDocumentXml: FOOTNOTE_DOCUMENT,
        original: 'The filing deadline is Friday.',
        modified: 'The filing deadline is Monday.',
        requiredElements: { footnoteReference: 1 },
        packageParts: { footnotesXml: FOOTNOTES_XML }
    },
    {
        name: 'legal-endnote-adjacent-duration',
        category: 'legal',
        task: 'replace-adjacent-to-endnote',
        sourceDocumentXml: ENDNOTE_DOCUMENT,
        original: 'The indemnity survives for two years.',
        modified: 'The indemnity survives for three years.',
        requiredElements: { endnoteReference: 1 },
        packageParts: { endnotesXml: ENDNOTES_XML }
    },
    {
        name: 'administrative-header-footer-package',
        category: 'administrative',
        task: 'replace-with-header-footer',
        sourceDocumentXml: HEADER_FOOTER_DOCUMENT,
        original: 'The administrative report is a draft.',
        modified: 'The administrative report is final.',
        requiredElements: { headerReference: 1, footerReference: 1 },
        packageParts: {
            headers: [{ relationshipId: 'rIdHeader1', partName: 'header1.xml', xml: HEADER_XML }],
            footers: [{ relationshipId: 'rIdFooter1', partName: 'footer1.xml', xml: FOOTER_XML }]
        }
    },
    {
        name: 'legal-external-hyperlink-adjacent-replacement',
        category: 'legal',
        task: 'replace-adjacent-to-external-hyperlink',
        sourceDocumentXml: EXTERNAL_HYPERLINK_DOCUMENT,
        original: 'Review the filing policy before submitting within ten days.',
        modified: 'Review the filing policy before submitting within fifteen days.',
        requiredElements: { hyperlink: 1 },
        packageParts: {
            externalHyperlinks: [{
                relationshipId: 'rIdExternalPolicy',
                target: 'https://example.com/legal/filing-policy'
            }]
        }
    },
    {
        name: 'legal-structured-attachment-mixed-blocks',
        category: 'legal',
        task: 'replace-anchor-with-mixed-attachment',
        sourceDocumentXml: MIXED_ATTACHMENT_DOCUMENT,
        original: 'Date',
        modified: '# ATTACHMENT 4 - JOINT AGENCY PUBLIC ANNOUNCEMENT\n\nFOR IMMEDIATE RELEASE\n\n| Agency | Contact |\n| --- | --- |\n| BCHD | Dr. Jenkins |\n| MOHS | Marcus Vance |\n\n## PROTOCOL FOR PUBLIC STATEMENTS\n\n1. Joint clearance\n2. Rapid escalation\n\nThis attachment forms part of the agreement.',
        operationOptions: { structuredContent: true },
        assertionMode: 'contains',
        expectedAcceptedContains: [
            'ATTACHMENT 4 - JOINT AGENCY PUBLIC ANNOUNCEMENT',
            'BCHD',
            'Dr. Jenkins',
            'PROTOCOL FOR PUBLIC STATEMENTS',
            'Joint clearance',
            'Rapid escalation',
            'This attachment forms part of the agreement.'
        ],
        expectedAcceptedAbsent: ['Date', '| Agency |'],
        expectedRejectedContains: ['Date'],
        expectedRejectedAbsent: ['ATTACHMENT 4 - JOINT AGENCY PUBLIC ANNOUNCEMENT', 'BCHD'],
        requiredElements: { tbl: 1, tr: 3, tc: 6, numPr: 2, pStyle: 2 }
    }
];

export const WORD_TASK_CASES = WORD_TASK_CASE_DEFINITIONS.map(testCase => ({
    ...testCase,
    coverageMetadata: WORD_TASK_COVERAGE[testCase.name]
}));
