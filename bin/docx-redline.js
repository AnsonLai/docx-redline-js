#!/usr/bin/env node
import { runCli } from '../node/cli.js';
process.exitCode = await runCli();
