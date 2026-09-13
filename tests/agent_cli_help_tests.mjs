import assert from 'node:assert/strict';
import { executeCli } from '../node/cli.js';
import { CLI_COMMANDS, CLI_COMMAND_HELP, commandOptionKeys } from '../node/cli-help.js';

const globalHelp = await executeCli(['--help']);
assert.equal(globalHelp.status, 'ok');
assert.equal(globalHelp.command, 'help');
assert.deepEqual(globalHelp.commands.map(item => item.name), CLI_COMMANDS);
assert(globalHelp.documentation.includes('docs/AGENT_FAST_START.md'));
assert(Buffer.byteLength(JSON.stringify(globalHelp, null, 2)) < 8000);

const alternateGlobalHelp = await executeCli(['help']);
assert.deepEqual(alternateGlobalHelp, globalHelp);

for (const command of CLI_COMMANDS) {
    const help = await executeCli([command, '--help']);
    assert.equal(help.status, 'ok', command);
    assert.equal(help.forCommand, command);
    assert.match(help.usage, new RegExp(`docx-redline ${command.replace('-', '\\-')}`));
    assert.equal(help.options.length, commandOptionKeys(command).length);
    assert.equal(new Set(commandOptionKeys(command)).size, commandOptionKeys(command).length, `${command} has duplicate option keys`);
    assert.equal(help.options.some(item => !item.name || !item.description), false, command);
    assert(Buffer.byteLength(JSON.stringify(help, null, 2)) < 12000, `${command} help is too large`);
    assert.equal(JSON.stringify(help).includes('restoreDeletedParagraphByExactText'), false);
    assert.equal(CLI_COMMAND_HELP[command].options.length, help.options.length);
}

const applyHelp = await executeCli(['apply', '--help']);
assert.match(applyHelp.notes.join(' '), /modified is complete desired accepted-view content/i);
assert.match(applyHelp.notes.join(' '), /source is never overwritten unless --in-place/i);
assert.match(applyHelp.notes.join(' '), /serializer-backed stdin/i);
assert.deepEqual(applyHelp.examples.map(item => item.operation.type), ['redline', 'comment', 'restore']);
assert.equal(applyHelp.examples[2].operation.target.revisionView, 'rejected');
assert.equal(applyHelp.examples[0].operation.modified, 'Revised clause.');

const extractHelp = await executeCli(['extract', '--help']);
assert.match(extractHelp.notes.join(' '), /case-insensitive/i);
assert(extractHelp.options.some(item => item.name.includes('--around') && item.name.includes('-C')));
assert(extractHelp.options.some(item => item.name.includes('--limit')));
assert(extractHelp.options.some(item => item.name.includes('--all')));
assert.match(extractHelp.examples[0].command, /extract .*--search .*--around/);

const helpCommand = await executeCli(['help', 'apply']);
assert.deepEqual(helpCommand, applyHelp);
const unknown = await executeCli(['help', 'not-a-command']);
assert.equal(unknown.status, 'error');
assert.equal(unknown.error.code, 'UNKNOWN_COMMAND');

console.log('agent CLI help tests passed');
