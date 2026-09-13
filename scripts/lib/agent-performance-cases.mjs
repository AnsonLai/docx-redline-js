/**
 * Deterministic legal-edit tasks used by the observational agent-workflow
 * benchmark. These are protocol/fidelity cases, not legal drafting guidance.
 */
export const AGENT_PERFORMANCE_CASES = Object.freeze([
    {
        id: 'terminal-punctuation',
        description: 'Add terminal punctuation to a short list item.',
        indexes: [37],
        buildEdits: ([paragraph]) => [{
            operationId: 'terminal-punctuation',
            paragraph,
            desiredText: `${paragraph.exactText}.`,
            replacements: [{ find: 'Materials', replace: 'Materials.' }]
        }]
    },
    {
        id: 'term-duration',
        description: 'Change a short duration phrase inside a longer provision.',
        indexes: [32],
        buildEdits: ([paragraph]) => [{
            operationId: 'term-duration',
            paragraph,
            desiredText: paragraph.exactText.replace('five (5) years', 'three (3) years'),
            replacements: [{ find: 'five (5) years', replace: 'three (3) years' }]
        }]
    },
    {
        id: 'simple-mutuality',
        description: 'Convert a one-sided remedies provision into a deterministic mutual form.',
        indexes: [46],
        buildEdits: ([paragraph]) => [{
            operationId: 'simple-mutuality',
            paragraph,
            desiredText: paragraph.exactText
                .replaceAll('The Receiving Party', 'Each Party')
                .replaceAll('the Disclosing Party', 'the other Party'),
            replacements: [
                { find: 'The Receiving Party', replace: 'Each Party' },
                { find: 'the Disclosing Party', replace: 'the other Party', occurrence: 1 },
                { find: 'the Disclosing Party', replace: 'the other Party', occurrence: 2 }
            ]
        }]
    },
    {
        id: 'full-clause-rewrite',
        description: 'Replace a complete notices provision.',
        indexes: [53],
        buildEdits: ([paragraph]) => [{
            operationId: 'full-clause-rewrite',
            paragraph,
            desiredText: 'Notices. All notices must be in writing and delivered personally, by recognized courier, or by email to the addresses specified by the Parties.'
        }]
    },
    {
        id: 'mixed-comment-redline',
        description: 'Add a comment and a redline in one batch.',
        indexes: [52, 53],
        buildEdits: ([assignment, notices]) => [
            {
                operationId: 'assignment-comment',
                paragraph: assignment,
                commentContent: 'Confirm whether affiliate assignments should be permitted.'
            },
            {
                operationId: 'notice-email',
                paragraph: notices,
                desiredText: notices.exactText.replace('addresses listed above', 'addresses and email contacts listed above'),
                replacements: [{
                    find: 'addresses listed above',
                    replace: 'addresses and email contacts listed above'
                }]
            }
        ]
    }
]);
