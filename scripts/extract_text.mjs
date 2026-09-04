#!/usr/bin/env node

// Compatibility entrypoint for legacy docx-redline skill invocations.
// The supported CLI owns selector parsing and emits 1-based paragraph indexes.
import { runCli } from '../node/cli.js';

process.exitCode = await runCli(['extract', ...process.argv.slice(2)]);
