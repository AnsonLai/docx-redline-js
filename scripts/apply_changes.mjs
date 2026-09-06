#!/usr/bin/env node

// Compatibility entrypoint for legacy docx-redline skill invocations.
// All targeting, transaction, validation, and output behavior is delegated to
// the supported CLI instead of being reimplemented here.
import { runCli } from '../node/cli.js';

const legacyArgs = process.argv.slice(2);
const input = legacyArgs.shift();
const operations = legacyArgs.shift();
const output = legacyArgs[0] && !legacyArgs[0].startsWith('--')
    ? legacyArgs.shift()
    : null;

const delegatedArgs = ['apply'];
if (input) delegatedArgs.push(input);
if (operations) delegatedArgs.push('--operations', operations);
if (output) delegatedArgs.push('--output', output);
delegatedArgs.push(...legacyArgs);
if (!legacyArgs.some(argument => argument === '--author' || argument.startsWith('--author='))) {
    delegatedArgs.push('--author', process.env.DOCX_REDLINE_AUTHOR || 'Agent');
}
if (!legacyArgs.some(argument => argument === '--atomic' || argument === '--no-atomic')) {
    delegatedArgs.push('--atomic');
}

process.exitCode = await runCli(delegatedArgs);
