/**
 * English legal/administrative task catalogue for independent Word validation.
 * Keep expectations derived from edit intent rather than engine output.
 */
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
        original: 'The parties agree to the following terms.\nThis paragraph is intentionally removed.',
        modified: 'The parties agree to the following terms.'
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
    }
];
