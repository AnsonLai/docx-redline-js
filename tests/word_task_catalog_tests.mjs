import assert from 'assert/strict';

import { WORD_TASK_CASES } from './fixtures/word-task-cases.mjs';

assert.ok(WORD_TASK_CASES.length >= 18, 'Word task catalogue should cover at least eighteen tasks');

const names = new Set();
const categories = new Set();
const tasks = new Set();

for (const testCase of WORD_TASK_CASES) {
    assert.match(testCase.name, /^[a-z0-9-]+$/);
    assert.ok(!names.has(testCase.name), `Duplicate Word task name: ${testCase.name}`);
    names.add(testCase.name);

    assert.ok(['legal', 'administrative'].includes(testCase.category));
    categories.add(testCase.category);
    tasks.add(testCase.task);
    assert.equal(typeof testCase.original, 'string');
    assert.equal(typeof testCase.modified, 'string');
    if (testCase.expectNoOp) {
        assert.equal(testCase.original, testCase.modified);
        assert.equal(typeof testCase.sourceDocumentXml, 'string');
    } else {
        assert.notEqual(testCase.original, testCase.modified);
    }
    for (const [field, value] of Object.entries({
        original: testCase.original,
        modified: testCase.modified,
        sourceText: testCase.sourceText,
        expectedAcceptedText: testCase.expectedAcceptedText,
        expectedRejectedText: testCase.expectedRejectedText,
        sourceDocumentXml: testCase.sourceDocumentXml
    })) {
        if (value !== undefined) {
            assert.ok(/^[\x00-\x7F]*$/.test(value), `${testCase.name} ${field} must be English/ASCII for this lane`);
        }
    }
}

assert.deepEqual([...categories].sort(), ['administrative', 'legal']);
assert.ok(tasks.size >= 8, 'Word task catalogue should exercise at least eight distinct task types');

console.log(`PASS: Word task catalogue (${WORD_TASK_CASES.length} cases, ${tasks.size} task types)`);
