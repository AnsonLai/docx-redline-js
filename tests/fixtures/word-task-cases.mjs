/**
 * English legal/administrative task catalogue for independent Word validation.
 * Keep expectations derived from edit intent rather than engine output.
 */
const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

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

export const WORD_TASK_CASES = [
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
    }
];
