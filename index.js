#!/usr/bin/env node
import { parseArgs } from 'node:util';
import { loadConfig } from './src/config.js';
import { run } from './src/run.js';

const USAGE = `Usage:
  node index.js --check              authenticate and verify the required scopes, nothing else
  node index.js --list               fetch the segment and print its members (no writes)
  node index.js [--limit N]          process members; with DRY_RUN=true (default) only prints what would happen
  DRY_RUN=false node index.js [--limit N]
                                     LIVE: create gift cards (Shopify emails them), tag customers`;

async function main(argv) {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        check: { type: 'boolean' },
        list: { type: 'boolean' },
        limit: { type: 'string' },
        help: { type: 'boolean', short: 'h' },
      },
      strict: true,
      allowPositionals: false,
    });
  } catch (err) {
    console.error(err.message);
    console.error(USAGE);
    return 2;
  }
  const { values } = parsed;
  if (values.help) {
    console.log(USAGE);
    return 0;
  }

  let limit = 0;
  if (values.limit !== undefined) {
    limit = Number(values.limit);
    if (!Number.isInteger(limit) || limit < 1) {
      console.error('--limit must be a positive integer');
      return 2;
    }
  }
  const mode = values.check ? 'check' : values.list ? 'list' : 'run';

  let config;
  try {
    config = loadConfig();
  } catch (err) {
    console.error(`Config error: ${err.message}`);
    return 1;
  }

  try {
    const result = await run({ config, mode, limit });
    return result.exitCode;
  } catch (err) {
    console.error(`Error: ${err.message}`);
    return 1;
  }
}

// Set exitCode instead of calling process.exit() so buffered output is never cut off.
process.exitCode = await main(process.argv.slice(2));
