import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { parseCommand, UsageError, USAGE, COMMANDS, PREVIEW_VARIANTS } from '../src/cli.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('cli: every command and its options', () => {
  assert.deepEqual(parseCommand(['check']), { command: 'check' });
  assert.deepEqual(parseCommand(['select']), { command: 'select', refresh: false });
  assert.deepEqual(parseCommand(['select', '--refresh']), { command: 'select', refresh: true });
  assert.deepEqual(parseCommand(['issue']), { command: 'issue', limit: undefined, retryFailed: false, repairOnly: false });
  assert.deepEqual(parseCommand(['issue', '--limit', '20']), { command: 'issue', limit: 20, retryFailed: false, repairOnly: false });
  assert.deepEqual(parseCommand(['issue', '--limit=500', '--retry-failed']), { command: 'issue', limit: 500, retryFailed: true, repairOnly: false });
  assert.deepEqual(parseCommand(['issue', '--repair-only']), { command: 'issue', limit: undefined, retryFailed: false, repairOnly: true });
  assert.deepEqual(parseCommand(['issue', '--repair-only', '--limit', '5']), { command: 'issue', limit: 5, retryFailed: false, repairOnly: true });
  assert.deepEqual(parseCommand(['usage']), { command: 'usage' });
  assert.deepEqual(parseCommand(['verify']), { command: 'verify' });
  assert.deepEqual(parseCommand(['export']), { command: 'export', refresh: false, out: null });
  assert.deepEqual(parseCommand(['export', '--refresh', '--out', '~/Desktop/a.xlsx']), { command: 'export', refresh: true, out: '~/Desktop/a.xlsx' });
  assert.deepEqual(parseCommand(['preview']), { command: 'preview', variant: 'all', seq: null, open: false });
  assert.deepEqual(parseCommand(['preview', 'first', '--seq', '25', '--open']), { command: 'preview', variant: 'first', seq: 25, open: true });
  assert.deepEqual(parseCommand(['preview', 'original']), { command: 'preview', variant: 'original', seq: null, open: false });
  assert.deepEqual(parseCommand([]), { command: 'help' });
  assert.deepEqual(parseCommand(['issue', '--help']), { command: 'help' });
});

test('cli: usage errors', () => {
  const usage = (argv, re, why) => assert.throws(() => parseCommand(argv), (err) => err instanceof UsageError && re.test(err.message), why);
  usage(['bogus'], /未知命令 "bogus"/);
  usage(['issue', '--limit', '0'], /--limit 必须是正整数/);
  usage(['issue', '--limit=-5'], /--limit 必须是正整数/);
  usage(['issue', '--limit', '2.5'], /--limit 必须是正整数/);
  usage(['issue', '--limit', '20abc'], /--limit 必须是正整数/);
  usage(['issue', '--repair-only', '--retry-failed'], /^--repair-only 和 --retry-failed 不能一起用：--repair-only 只补记和补打 tag，不建卡$/);
  usage(['issue', '--retry-failed', '--limit', '3', '--repair-only'], /不能一起用/);
  usage(['remind', '--round', '1'], /未知命令 "remind"/, 'reminders are sent with Shopify Email now');
  usage(['preview', 'reminder'], /preview 的邮件只能是 first、original/);
  usage(['preview', 'remind1'], /preview 的邮件只能是 first、original/, 'the reminder variants are gone');
  usage(['preview', 'remind2'], /preview 的邮件只能是/);
  usage(['preview', 'first', 'original'], /多余的参数/);
  usage(['preview', '--seq', '0'], /--seq/);
  usage(['export', '--out', ' '], /--out/);
  // node:util rejects unknown options, missing values and stray positionals.
  assert.throws(() => parseCommand(['issue', '--limt', '5']), { code: 'ERR_PARSE_ARGS_UNKNOWN_OPTION' });
  assert.throws(() => parseCommand(['issue', '--limit']), { code: 'ERR_PARSE_ARGS_INVALID_OPTION_VALUE' });
  assert.throws(() => parseCommand(['select', 'now']), { code: 'ERR_PARSE_ARGS_UNEXPECTED_POSITIONAL' });
  assert.throws(() => parseCommand(['issue', '--dry-run']), { code: 'ERR_PARSE_ARGS_UNKNOWN_OPTION' });
  assert.throws(() => parseCommand(['issue', '--repair-only=yes']), { code: 'ERR_PARSE_ARGS_INVALID_OPTION_VALUE' });
  assert.throws(() => parseCommand(['issue', '--repair']), { code: 'ERR_PARSE_ARGS_UNKNOWN_OPTION' });
});

test('cli: USAGE lists issue --repair-only, and index.js hands it to runIssue', () => {
  const issueUsage = USAGE.slice(USAGE.indexOf('  issue '), USAGE.indexOf('  usage '));
  assert.match(issueUsage, /^ {2}issue \[--limit N\] \[--retry-failed\] \[--repair-only\]$/m);
  assert.match(issueUsage, /--repair-only 只补记和补打 tag，不建新卡（不需要 --limit）/);
  // index.js cannot be run here (it reads the real .env), so its wiring is checked in the source.
  const index = fs.readFileSync(path.join(ROOT, 'index.js'), 'utf8');
  const call = index.slice(index.indexOf('runIssue({'), index.indexOf('})', index.indexOf('runIssue({')));
  assert.match(call, /repairOnly: args\.repairOnly/);
  assert.match(call, /retryFailed: args\.retryFailed/);
  assert.match(call, /limit: args\.limit/);
});

test('cli: no remind command anywhere; USAGE explains the Shopify Email reminders and lists two preview variants', () => {
  assert.equal('remind' in COMMANDS, false);
  assert.deepEqual(PREVIEW_VARIANTS, ['first', 'original']);
  assert.doesNotMatch(USAGE, /remind/);
  assert.doesNotMatch(USAGE, /REMIND_/);
  assert.match(USAGE, /^ {2}preview \[first\|original\] \[--seq N\] \[--open\]$/m);
  assert.match(USAGE, /不写邮件名 = 两种都生成/);
  assert.match(USAGE, /提醒邮件改用 Shopify Email 发送：发送前先跑 usage（给用过卡的人打 <SENT_TAG>-USED）/);
  assert.match(USAGE, /customer_tags CONTAINS '<SENT_TAG>' AND NOT customer_tags CONTAINS '<SENT_TAG>-USED'/);
  // index.js neither dispatches a remind command nor imports the deleted module.
  const index = fs.readFileSync(path.join(ROOT, 'index.js'), 'utf8');
  assert.doesNotMatch(index, /remind/i);
  assert.equal(fs.existsSync(path.join(ROOT, 'src', 'remind.js')), false);
  assert.equal(fs.existsSync(path.join(ROOT, 'test', 'remind.test.js')), false);
});

// These invocations fail before the configuration is loaded, so the project's .env is never read.
const run = (...args) => spawnSync(process.execPath, [path.join(ROOT, 'index.js'), ...args], { encoding: 'utf8', cwd: ROOT, env: { PATH: process.env.PATH } });

test('cli: exit codes of the entry point for help and usage errors', () => {
  const none = run();
  assert.equal(none.status, 2);
  assert.ok(none.stdout.includes('用法：node index.js <命令>'));
  const help = run('help');
  assert.equal(help.status, 0);
  assert.equal(help.stdout.trim(), USAGE);
  assert.equal(run('issue', '--help').status, 0);
  const bogus = run('bogus');
  assert.equal(bogus.status, 2);
  assert.ok(bogus.stderr.includes('未知命令 "bogus"'));
  const zero = run('issue', '--limit', '0');
  assert.equal(zero.status, 2);
  assert.ok(zero.stderr.includes('--limit 必须是正整数'));
  const remind = run('remind', '--round', '1');
  assert.equal(remind.status, 2, 'remind is an unknown command now');
  assert.ok(remind.stderr.includes('未知命令 "remind"'));
  assert.equal(run('issue', '--retry-faild').status, 2, 'a misspelt flag never reaches the configuration');
  assert.equal(run('preview', 'bogus').status, 2);
  assert.equal(run('preview', 'remind1').status, 2);
  assert.equal(run('select', '--refrsh').status, 2);
  const both = run('issue', '--repair-only', '--retry-failed');
  assert.equal(both.status, 2, 'a contradictory pair never reaches the configuration');
  assert.ok(both.stderr.includes('--repair-only 和 --retry-failed 不能一起用'));
});
