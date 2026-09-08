import assert from 'assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { openDocx } from '../node/index.js';

const testDir = dirname(fileURLToPath(import.meta.url));
const fixturesDir = join(testDir, 'fixtures', 'cross-author-slicing');
const scenarios = [
    'insert-interior',
    'delete-interior',
    'delete-boundary-start',
    'delete-boundary-end',
    'delete-straddle-baseline-insertion'
];

function fixtureBuffer(scenario, state) {
    return readFileSync(join(fixturesDir, `${scenario}-${state}.docx`));
}

function visibleText(document) {
    return document.inspect().paragraphs.map(paragraph => paragraph.text).join('\n').trim();
}

async function resolvedText(buffer, action) {
    const document = openDocx(buffer);
    const result = await document.resolveRevisions(action, { allAuthors: true, validate: true });
    assert.equal(result.status, 'ok', result.error?.message);
    assert.equal(result.written, true);
    assert.deepEqual(document.inspect().revisionAuthors, []);
    return visibleText(document);
}

// PKG-01..05: reproduce each Word Desktop fixture through the strict package facade.
for (const scenario of scenarios) {
    const pendingBuffer = fixtureBuffer(scenario, 'pending');
    const source = openDocx(pendingBuffer);
    const removeAnson = await source.resolveRevisions('reject', {
        author: 'Anson Lai',
        validate: true
    });
    assert.equal(removeAnson.status, 'ok', `${scenario}: ${removeAnson.error?.message || 'source reconstruction failed'}`);
    assert.equal(removeAnson.written, true, `${scenario}: expected an Anson revision in the Word fixture`);

    const sourceInspection = source.inspect();
    const target = sourceInspection.paragraphs.find(paragraph => paragraph.text.trim().length > 0);
    const desiredText = visibleText(openDocx(pendingBuffer));
    assert(target, `${scenario}: source paragraph not found`);

    const applied = await source.applyOperations([{
        type: 'replace',
        target: {
            exactText: target.exactText,
            fingerprint: target.fingerprint
        },
        modified: desiredText,
        author: 'Anson Lai',
        existingRevisions: 'slice-cross-author'
    }], {
        author: 'Anson Lai',
        atomic: true,
        strictTargets: true,
        validate: true
    });

    assert.equal(applied.status, 'ok', `${scenario}: ${applied.error?.message || 'package application failed'}`);
    assert.equal(applied.written, true, `${scenario}: package was not written`);
    assert.deepEqual(applied.inspection.revisionAuthors, ['Anson Lai', 'Barry Plasteras']);
    assert(applied.artifactsChanged.includes('word/document.xml'));

    const output = applied.toBuffer();
    assert.equal(
        await resolvedText(output, 'accept'),
        visibleText(openDocx(fixtureBuffer(scenario, 'accepted'))),
        `${scenario}: package AcceptAll differs from Word Desktop`
    );
    assert.equal(
        await resolvedText(output, 'reject'),
        visibleText(openDocx(fixtureBuffer(scenario, 'rejected'))),
        `${scenario}: package RejectAll differs from Word Desktop`
    );
}

// PKG-06: create the checked-in three-reviewer negotiation through two package rounds.
{
    const pendingBuffer = fixtureBuffer('multi-author-stacked', 'pending');
    const source = openDocx(pendingBuffer);
    for (const author of ['Chris Davis', 'Anson Lai']) {
        const resolved = await source.resolveRevisions('reject', { author, validate: true });
        assert.equal(resolved.status, 'ok', resolved.error?.message);
        assert.equal(resolved.written, true);
    }

    const firstTarget = source.inspect().paragraphs.find(paragraph => paragraph.text.trim().length > 0);
    const anson = await source.applyOperations([{
        type: 'replace',
        target: { exactText: firstTarget.exactText, fingerprint: firstTarget.fingerprint },
        modified: 'Provision first draft with initial metrics.',
        author: 'Anson Lai',
        existingRevisions: 'slice-cross-author'
    }], { author: 'Anson Lai', atomic: true, validate: true });
    assert.equal(anson.status, 'ok', anson.error?.message);
    assert.equal(anson.written, true);

    const chrisSource = openDocx(anson.toBuffer());
    const secondTarget = chrisSource.inspect().paragraphs.find(paragraph => paragraph.text.trim().length > 0);
    const chris = await chrisSource.applyOperations([{
        type: 'replace',
        target: { exactText: secondTarget.exactText, fingerprint: secondTarget.fingerprint },
        modified: 'Provision first draft with metrics.',
        author: 'Chris Davis',
        existingRevisions: 'slice-cross-author'
    }], { author: 'Chris Davis', atomic: true, validate: true });

    assert.equal(chris.status, 'ok', chris.error?.message);
    assert.equal(chris.written, true);
    assert.deepEqual(chris.inspection.revisionAuthors, ['Anson Lai', 'Barry Plasteras', 'Chris Davis']);
    assert.equal(
        await resolvedText(chris.toBuffer(), 'accept'),
        visibleText(openDocx(fixtureBuffer('multi-author-stacked', 'accepted')))
    );
    assert.equal(
        await resolvedText(chris.toBuffer(), 'reject'),
        visibleText(openDocx(fixtureBuffer('multi-author-stacked', 'rejected')))
    );
}

console.log('PASS: cross-author slicing package differential tests');
