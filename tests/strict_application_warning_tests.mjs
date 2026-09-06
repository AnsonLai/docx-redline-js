import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import {
    applyOperationsToDocumentXml,
    applyOperationToDocumentXml,
    preflightOperations
} from '../services/standalone-operation-runner.js';
import { applyToParagraphByExactText } from '../services/document-operation-mutations.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

function buildDocXml(paragraphs) {
    const body = paragraphs.map(({ text, id }) => {
        const idAttr = id ? ` w14:paraId="${id}"` : '';
        return `<w:p${idAttr}><w:r><w:t>${text}</w:t></w:r></w:p>`;
    }).join('');
    return `<w:document xmlns:w="${NS_W}" xmlns:w14="${NS_W14}"><w:body>${body}</w:body></w:document>`;
}

// Test 1: Permissive batch emits AMBIGUOUS_TARGET_HEURISTIC_USED on duplicate target
{
    const input = buildDocXml([
        { id: '11111111', text: 'Repeated paragraph text.' },
        { id: '22222222', text: 'Repeated paragraph text.' }
    ]);

    const result = await applyOperationsToDocumentXml(input, [
        {
            type: 'redline',
            target: 'Repeated paragraph text.',
            modified: 'Modified first candidate.'
        }
    ], 'Agent', null, { generateRedlines: true, strictTargets: false });

    assert.equal(result.hasChanges, true);
    assert.equal(result.results.length, 1);
    const step = result.results[0];
    assert.equal(step.status, 'applied');
    assert.ok(Array.isArray(step.warnings), 'Expected step.warnings array');
    const ambiguousWarn = step.warnings.find(w => w.includes('AMBIGUOUS_TARGET_HEURISTIC_USED'));
    assert.ok(ambiguousWarn, `Expected AMBIGUOUS_TARGET_HEURISTIC_USED warning, got: ${JSON.stringify(step.warnings)}`);
    assert.ok(ambiguousWarn.includes('2 paragraphs') || ambiguousWarn.includes('2 candidates'), 'Warning should include candidate count (2)');
    assert.ok(ambiguousWarn.includes('v1.0.0'), 'Warning should mention v1.0.0 migration');
    assert.ok(ambiguousWarn.includes('strictTargets'), 'Warning should mention strictTargets guidance');

    // Receipt should also carry the warning
    assert.ok(step.receipt, 'Expected receipt');
    const receiptWarn = step.receipt.warnings.find(w => w.includes('AMBIGUOUS_TARGET_HEURISTIC_USED'));
    assert.ok(receiptWarn, 'Expected receipt.warnings to include AMBIGUOUS_TARGET_HEURISTIC_USED');
}

// Test 2: Strict batch refuses duplicate target with AMBIGUOUS_TARGET
{
    const input = buildDocXml([
        { id: '11111111', text: 'Repeated paragraph text.' },
        { id: '22222222', text: 'Repeated paragraph text.' }
    ]);

    const result = await applyOperationsToDocumentXml(input, [
        {
            type: 'redline',
            target: 'Repeated paragraph text.',
            modified: 'Modified first candidate.'
        }
    ], 'Agent', null, { generateRedlines: true, strictTargets: true });

    assert.equal(result.hasChanges, false);
    assert.equal(result.status, 'error');
    assert.equal(result.results[0].error.code, 'AMBIGUOUS_TARGET');
}

// Test 3: Permissive batch does NOT emit AMBIGUOUS_TARGET_HEURISTIC_USED on unique target
{
    const input = buildDocXml([
        { id: '11111111', text: 'Unique first paragraph.' },
        { id: '22222222', text: 'Unique second paragraph.' }
    ]);

    const result = await applyOperationsToDocumentXml(input, [
        {
            type: 'redline',
            target: 'Unique first paragraph.',
            modified: 'Modified unique paragraph.'
        }
    ], 'Agent', null, { generateRedlines: true, strictTargets: false });

    assert.equal(result.hasChanges, true);
    const step = result.results[0];
    assert.equal(step.status, 'applied');
    const hasAmbiguousWarn = (step.warnings || []).some(w => w.includes('AMBIGUOUS_TARGET_HEURISTIC_USED'));
    assert.equal(hasAmbiguousWarn, false, 'Unique target should not emit AMBIGUOUS_TARGET_HEURISTIC_USED');
}

// Test 4: Permissive batch with duplicate target disambiguated by targetRef does NOT emit ambiguous warning
{
    const input = buildDocXml([
        { id: '11111111', text: 'Repeated paragraph text.' },
        { id: '22222222', text: 'Repeated paragraph text.' }
    ]);

    const result = await applyOperationsToDocumentXml(input, [
        {
            type: 'redline',
            target: 'Repeated paragraph text.',
            targetRef: 2,
            modified: 'Modified second candidate.'
        }
    ], 'Agent', null, { generateRedlines: true, strictTargets: false });

    assert.equal(result.hasChanges, true);
    const step = result.results[0];
    assert.equal(step.status, 'applied');
    assert.equal(step.resolvedBy, 'ref');
    const hasAmbiguousWarn = (step.warnings || []).some(w => w.includes('AMBIGUOUS_TARGET_HEURISTIC_USED'));
    assert.equal(hasAmbiguousWarn, false, 'Disambiguated targetRef should not emit AMBIGUOUS_TARGET_HEURISTIC_USED');
}

// Test 5: Single-operation applyOperationToDocumentXml emits warning on duplicate target
{
    const input = buildDocXml([
        { id: '11111111', text: 'Repeated paragraph text.' },
        { id: '22222222', text: 'Repeated paragraph text.' }
    ]);

    const result = await applyOperationToDocumentXml(input, {
        type: 'redline',
        target: 'Repeated paragraph text.',
        modified: 'Modified text.'
    }, 'Agent', null, { generateRedlines: true, strictTargets: false });

    assert.equal(result.hasChanges, true);
    assert.equal(result.status, 'ok');
    assert.ok(Array.isArray(result.warnings), 'Expected warnings array');
    const warn = result.warnings.find(w => w.includes('AMBIGUOUS_TARGET_HEURISTIC_USED'));
    assert.ok(warn, 'Expected AMBIGUOUS_TARGET_HEURISTIC_USED warning in single op result');
    assert.ok(result.receipt.warnings.some(w => w.includes('AMBIGUOUS_TARGET_HEURISTIC_USED')), 'Expected warning in receipt');
}

// Test 6: Preflight with strictTargets: false emits AMBIGUOUS_TARGET_HEURISTIC_USED on item result
{
    const input = buildDocXml([
        { id: '11111111', text: 'Repeated paragraph text.' },
        { id: '22222222', text: 'Repeated paragraph text.' }
    ]);

    const preflight = preflightOperations(input, [
        {
            type: 'redline',
            target: 'Repeated paragraph text.',
            modified: 'Modified text.'
        }
    ], 'Agent', { strictTargets: false });

    assert.equal(preflight.valid, true);
    assert.equal(preflight.status, 'ok');
    const item = preflight.results[0];
    assert.equal(item.status, 'ready');
    assert.ok(Array.isArray(item.warnings), 'Expected item.warnings');
    assert.ok(item.warnings.some(w => w.includes('AMBIGUOUS_TARGET_HEURISTIC_USED')), 'Preflight should report AMBIGUOUS_TARGET_HEURISTIC_USED');
}

// Test 7: Direct low-level applyToParagraphByExactText invokes onWarn callback
{
    const input = buildDocXml([
        { id: '11111111', text: 'Repeated paragraph text.' },
        { id: '22222222', text: 'Repeated paragraph text.' }
    ]);

    const warned = [];
    const result = await applyToParagraphByExactText(
        input,
        'Repeated paragraph text.',
        'Modified direct.',
        'Agent',
        null,
        null,
        null,
        {
            generateRedlines: true,
            strictTargets: false,
            onWarn: (msg) => warned.push(msg)
        }
    );

    assert.equal(result.hasChanges, true);
    assert.ok(warned.some(w => w.includes('AMBIGUOUS_TARGET_HEURISTIC_USED')), 'onWarn callback should receive warning');
}

console.log('All WP-18 strict application warning tests passed successfully!');
