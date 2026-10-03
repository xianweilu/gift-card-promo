import { test, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { installFakeShopify } from './fake-shopify.js';
import { testConfig, selectionFixture, makeCustomer, makeOrder, memoryLog, gid } from './helpers.js';
import { campaignPaths, readJson, acquireRunLock, runningCommand } from '../src/campaign.js';
import { resetClient } from '../src/shopify.js';
import { runExport, resolveOutPath, compareGids, missingSelectionMessage } from '../src/export-command.js';
import { OPEN_IN_EXCEL_WARNING } from '../src/report/excel.js';

const NOW = '2026-10-05T18:00:00.000Z';
const now = () => new Date(NOW);
const noSleep = async () => {};

function customers(count) {
  return Array.from({ length: count }, (_, i) => makeCustomer({ n: i + 1, lastOrder: makeOrder({ n: 100 + i + 1, total: '50.00' }) }));
}

/**
 * writeReport stand-in: records its arguments and behaves like the real one: each
 * warning is logged through args.log when it happens and also returned in `warnings`.
 * `copyFails` simulates a --out copy that could not be saved (a warning, out: null).
 */
function reportStub({ fail = null, warnings = [], onCall = null, copyFails = null } = {}) {
  const calls = [];
  const fn = async (args) => {
    calls.push(args);
    if (onCall) onCall(args);
    if (fail) throw new Error(fail);
    const all = [...warnings];
    if (args.out && copyFails) all.push(`另存 Excel 到 ${args.out} 失败：${copyFails}`);
    for (const w of all) args.log.warn(w);
    return { file: args.paths.excel, out: args.out && !copyFails ? args.out : null, warnings: all };
  };
  fn.calls = calls;
  return fn;
}

function holdRunLock(paths, command) {
  fs.mkdirSync(paths.dir, { recursive: true });
  const info = { pid: process.pid, host: os.hostname(), command, startedAt: '2026-10-05T16:00:00.000Z', token: 'held-by-test' };
  fs.writeFileSync(paths.runLock, JSON.stringify(info));
  return info;
}

let env;
let fake;
let scratch = [];
afterEach(() => {
  fake?.restore();
  fake = null;
  resetClient();
  env?.cleanup();
  env = null;
  for (const dir of scratch) fs.rmSync(dir, { recursive: true, force: true });
  scratch = [];
});

function tempDir() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'gcp-out-'));
  scratch.push(dir);
  return dir;
}

async function withSelection(count = 2) {
  env = testConfig();
  const list = customers(count);
  await selectionFixture(env.config, { customers: list });
  return { config: env.config, paths: campaignPaths(env.config), list };
}

test('no list yet: exit 1 with "请先运行 select" and the hint to copy the folder back, nothing written', async () => {
  env = testConfig();
  const paths = campaignPaths(env.config);
  fake = installFakeShopify();
  const log = memoryLog();
  const writeReport = reportStub();

  const result = await runExport({ config: env.config, log, now, writeReport });

  assert.equal(result.exitCode, 1);
  assert.equal(result.file, null);
  // After issuing started, select refuses (Shopify has the cards): the original folder is the only way back.
  assert.deepEqual(log.lines, ['ERROR 还没有名单，请先运行 select。如果已经开始发放，不要重新运行 select，请把原来的 campaigns/2026-10/ 文件夹拷回来。']);
  assert.equal(writeReport.calls.length, 0);
  assert.equal(fs.existsSync(paths.excel), false);
  assert.equal(fake.state.calls.token.length, 0);
});

test('missingSelectionMessage names the campaign folder of CAMPAIGN_ID (from the folder name when not given)', () => {
  env = testConfig({ CAMPAIGN_ID: '2026-10-test', SENT_TAG: 'OCT26RTPROMO-TEST', TEST_CUSTOMER_IDS: '1' });
  const paths = campaignPaths(env.config);
  const expected = '还没有名单，请先运行 select。如果已经开始发放，不要重新运行 select，请把原来的 campaigns/2026-10-test/ 文件夹拷回来。';
  assert.equal(missingSelectionMessage(paths, env.config.campaignId), expected);
  assert.equal(missingSelectionMessage(paths), expected);
  holdRunLock(paths, 'select');
  assert.equal(missingSelectionMessage(paths, env.config.campaignId), '还没有名单，select 正在运行，请等它跑完');
});

test('no list yet while select is running: tells the user to wait for select', async () => {
  env = testConfig();
  const paths = campaignPaths(env.config);
  fake = installFakeShopify();
  holdRunLock(paths, 'select');
  const log = memoryLog();
  const writeReport = reportStub();

  const result = await runExport({ config: env.config, refresh: true, log, now, writeReport });

  assert.equal(result.exitCode, 1);
  assert.deepEqual(log.lines, ['ERROR 还没有名单，select 正在运行，请等它跑完']);
  assert.equal(writeReport.calls.length, 0);
  assert.equal(runningCommand(paths), 'select', 'the running select keeps its lock');
});

test('exports from local data only: no Shopify call, no journal write, path printed, each warning once', async () => {
  const { config, paths } = await withSelection();
  fake = installFakeShopify();
  const log = memoryLog();
  const writeReport = reportStub({ warnings: ['Excel 正打开此文件，请关闭后重新打开才能看到最新内容'] });

  const result = await runExport({ config, log, now, writeReport });

  assert.equal(result.exitCode, 0);
  assert.equal(result.file, paths.excel);
  assert.deepEqual(result.warnings, ['Excel 正打开此文件，请关闭后重新打开才能看到最新内容']);
  assert.equal(writeReport.calls.length, 1);
  const args = writeReport.calls[0];
  assert.equal(args.config, config);
  assert.deepEqual(args.paths, paths);
  assert.equal(args.out, null);
  assert.equal(args.now, now);
  assert.equal(args.log, log);
  assert.ok(log.lines.includes(`INFO Excel 已导出：${paths.excel}`), log.lines.join('\n'));
  // writeReport printed it; export does not print result.warnings a second time.
  assert.equal(log.lines.filter((l) => l === 'WARN Excel 正打开此文件，请关闭后重新打开才能看到最新内容').length, 1, log.lines.join('\n'));
  assert.equal(fake.state.calls.token.length, 0);
  assert.equal(fake.state.calls.ops.length, 0);
  assert.equal(fs.existsSync(paths.journal), false, 'export never writes the journal');
  assert.equal(fs.existsSync(paths.tags), false, 'no refresh, no tag snapshot');
});

test('--refresh re-reads the tag holders from Shopify (read-only) into tags.json before exporting', async () => {
  const { config, paths, list } = await withSelection(3);
  fake = installFakeShopify({
    customers: [
      { ...list[0], tags: [config.sentTag] },
      { ...list[1], tags: ['VIP'] },
      { ...list[2], tags: ['vip', config.sentTag] },
      makeCustomer({ n: 10, tags: [config.sentTag] }), // not on the list: still part of the snapshot
    ],
  });
  const log = memoryLog();
  const writeReport = reportStub({ onCall: ({ paths: p }) => assert.ok(fs.existsSync(p.tags), 'tags.json is written before the workbook') });

  const result = await runExport({ config, refresh: true, log, now, sleep: noSleep, writeReport });

  assert.equal(result.exitCode, 0, log.lines.join('\n'));
  assert.deepEqual(readJson(paths.tags), {
    fetchedAt: NOW,
    tag: config.sentTag,
    ids: [gid('Customer', 1), gid('Customer', 3), gid('Customer', 10)],
  });
  assert.equal(writeReport.calls.length, 1);
  assert.deepEqual([...new Set(fake.state.calls.ops.map((o) => o.op))], ['TaggedCustomers']);
  assert.equal(fake.state.calls.create.length + fake.state.calls.tag.length + fake.state.calls.notify.length, 0);
  assert.ok(log.lines.some((l) => l.includes(`带 ${config.sentTag} 的客户：3 个`)), log.lines.join('\n'));
  assert.equal(fs.existsSync(paths.journal), false);
});

test('a failed refresh still exports (with the previous tag data) but exits 1', async () => {
  const { config, paths } = await withSelection();
  fake = installFakeShopify({ failures: { TaggedCustomers: [{ kind: 'http', status: 502 }] } });
  const log = memoryLog();
  const writeReport = reportStub();

  const result = await runExport({ config, refresh: true, log, now, sleep: noSleep, writeReport });

  assert.equal(result.exitCode, 1);
  assert.equal(result.file, paths.excel);
  assert.equal(writeReport.calls.length, 1);
  assert.equal(fs.existsSync(paths.tags), false);
  assert.ok(log.lines.some((l) => l.startsWith('ERROR 刷新 tag 失败') && l.includes('502')), log.lines.join('\n'));
  assert.ok(log.lines.some((l) => l.startsWith('WARN Excel 仍会导出')));
});

test('--out is resolved and passed through to writeReport', async () => {
  const { config, paths } = await withSelection();
  fake = installFakeShopify();
  const dir = tempDir();
  const log = memoryLog();

  // An explicit file.
  let writeReport = reportStub();
  const file = path.join(dir, 'copy.xlsx');
  let result = await runExport({ config, out: file, log, now, writeReport });
  assert.equal(result.exitCode, 0);
  assert.equal(writeReport.calls[0].out, file);
  assert.equal(result.out, file);
  assert.ok(log.lines.includes(`INFO 另存一份：${file}`), log.lines.join('\n'));

  // A folder: the workbook's own name is used.
  writeReport = reportStub();
  result = await runExport({ config, out: dir, log: memoryLog(), now, writeReport });
  assert.equal(writeReport.calls[0].out, path.join(dir, path.basename(paths.excel)));

  // No extension: ".xlsx" is added; a relative path is taken from the current directory.
  writeReport = reportStub();
  result = await runExport({ config, out: path.relative(process.cwd(), path.join(dir, 'report')), log: memoryLog(), now, writeReport });
  assert.equal(writeReport.calls[0].out, path.join(dir, 'report.xlsx'));
});

test('--out into a folder that does not exist: exit 2 before doing anything', async () => {
  const { config } = await withSelection();
  fake = installFakeShopify();
  const missing = path.join(tempDir(), 'nope', 'copy.xlsx');
  const log = memoryLog();
  const writeReport = reportStub();

  const result = await runExport({ config, refresh: true, out: missing, log, now, writeReport });

  assert.equal(result.exitCode, 2);
  assert.equal(writeReport.calls.length, 0);
  assert.equal(fake.state.calls.token.length, 0, 'no refresh either');
  assert.ok(log.lines.some((l) => l.startsWith('ERROR --out 指定的文件夹不存在')), log.lines.join('\n'));
});

test('resolveOutPath: "~" expansion, the main workbook itself, empty values', () => {
  const paths = { excel: '/x/campaigns/2026-10/gift-card-promo-2026-10.xlsx' };
  assert.equal(resolveOutPath('~/report.xlsx', paths), path.join(os.homedir(), 'report.xlsx'));
  assert.equal(resolveOutPath('~', paths), path.join(os.homedir(), 'gift-card-promo-2026-10.xlsx'));
  assert.equal(resolveOutPath(null, paths), null);
  assert.equal(resolveOutPath('  ', paths), null);
  const dir = tempDir();
  const main = { excel: path.join(dir, 'gift-card-promo-2026-10.xlsx') };
  assert.equal(resolveOutPath(main.excel, main), null, 'the main workbook is written anyway');
});

test('works while another command holds the run lock, without touching that lock', async () => {
  const { config, paths } = await withSelection();
  fake = installFakeShopify();
  const held = holdRunLock(paths, 'issue');
  const log = memoryLog();
  const writeReport = reportStub({ onCall: () => assert.equal(runningCommand(paths), 'issue') });

  const result = await runExport({ config, log, now, writeReport });

  assert.equal(result.exitCode, 0, log.lines.join('\n'));
  assert.equal(writeReport.calls.length, 1);
  assert.deepEqual(JSON.parse(fs.readFileSync(paths.runLock, 'utf8')), held);
  assert.ok(log.lines.some((l) => l.includes('issue 正在运行：导出的是此刻的进度')), log.lines.join('\n'));
});

test('a real run lock held in this process (as issue would) does not block export', async () => {
  const { config, paths } = await withSelection();
  fake = installFakeShopify();
  const release = acquireRunLock(paths, 'remind');
  try {
    const log = memoryLog();
    const writeReport = reportStub();
    const result = await runExport({ config, log, now, writeReport });
    assert.equal(result.exitCode, 0);
    assert.equal(writeReport.calls.length, 1);
    assert.equal(runningCommand(paths), 'remind');
    assert.ok(log.lines.some((l) => l.includes('remind 正在运行')));
  } finally {
    release();
  }
});

test('writeReport failing: exit 1 with the message', async () => {
  const { config } = await withSelection();
  fake = installFakeShopify();
  const log = memoryLog();

  const result = await runExport({ config, log, now, writeReport: reportStub({ fail: 'disk full' }) });

  assert.equal(result.exitCode, 1);
  assert.equal(result.file, null);
  assert.ok(log.lines.includes('ERROR 导出 Excel 失败：disk full'), log.lines.join('\n'));
});

test('the --out copy could not be saved: exit 1, but the main workbook is reported and its path printed', async () => {
  // Regression: a failed copy used to read as "导出 Excel 失败" with no path, although the main workbook was updated.
  const { config, paths } = await withSelection();
  fake = installFakeShopify();
  const target = path.join(tempDir(), 'copy.xlsx');
  const log = memoryLog();
  const writeReport = reportStub({ copyFails: 'EACCES: permission denied' });

  const result = await runExport({ config, out: target, log, now, writeReport });

  assert.equal(result.exitCode, 1);
  assert.equal(result.file, paths.excel);
  assert.equal(result.out, null);
  assert.equal(writeReport.calls[0].out, target);
  assert.deepEqual(result.warnings, [`另存 Excel 到 ${target} 失败：EACCES: permission denied`]);
  assert.ok(log.lines.includes(`INFO Excel 已导出：${paths.excel}`), log.lines.join('\n'));
  assert.equal(log.lines.filter((l) => l.startsWith('WARN 另存 Excel 到')).length, 1, 'the warning is printed once');
  assert.ok(log.lines.includes(`ERROR 没有另存到 ${target}（原因见上面的提示）；主 Excel 已更新：${paths.excel}`), log.lines.join('\n'));
  assert.equal(log.lines.some((l) => l.startsWith('INFO 另存一份')), false);
  assert.equal(log.lines.some((l) => l.includes('导出 Excel 失败')), false, 'the main workbook did not fail');
});

test('--out naming the main workbook itself is not a failed copy', async () => {
  const { config, paths } = await withSelection();
  fake = installFakeShopify();
  const log = memoryLog();
  const writeReport = reportStub();

  const result = await runExport({ config, out: paths.excel, log, now, writeReport });

  assert.equal(result.exitCode, 0, log.lines.join('\n'));
  assert.equal(writeReport.calls[0].out, null, 'nothing extra to save');
  assert.equal(result.out, null);
  assert.equal(log.lines.some((l) => l.startsWith('ERROR')), false);
});

test('with the real Excel writer, "Excel 正打开此文件" is printed exactly once (export and an --out copy)', async () => {
  const { config, paths } = await withSelection();
  fake = installFakeShopify();
  fs.writeFileSync(path.join(paths.dir, `~$${path.basename(paths.excel)}`), 'owner'); // the workbook is open in Excel
  const log = memoryLog();

  const result = await runExport({ config, log, now }); // no writeReport: src/report/excel.js

  assert.equal(result.exitCode, 0, log.lines.join('\n'));
  assert.equal(result.file, paths.excel);
  assert.ok(fs.existsSync(paths.excel));
  assert.equal(log.lines.filter((l) => l === `WARN ${OPEN_IN_EXCEL_WARNING}`).length, 1, log.lines.join('\n'));
  assert.deepEqual(result.warnings, [OPEN_IN_EXCEL_WARNING]);

  // The copy is fine: exit 0 and both paths printed.
  const dir = tempDir();
  const copyLog = memoryLog();
  const copy = await runExport({ config, out: dir, log: copyLog, now });
  assert.equal(copy.exitCode, 0, copyLog.lines.join('\n'));
  assert.equal(copy.out, path.join(dir, path.basename(paths.excel)));
  assert.ok(fs.existsSync(copy.out));
  assert.ok(copyLog.lines.includes(`INFO 另存一份：${copy.out}`), copyLog.lines.join('\n'));
  assert.equal(copyLog.lines.filter((l) => l === `WARN ${OPEN_IN_EXCEL_WARNING}`).length, 1, copyLog.lines.join('\n'));
});

test('compareGids orders by the numeric id and puts null last', () => {
  const ids = [gid('Customer', 10), null, gid('Customer', 9), gid('Customer', 100), gid('Customer', 9)];
  assert.deepEqual(ids.sort(compareGids), [gid('Customer', 9), gid('Customer', 9), gid('Customer', 10), gid('Customer', 100), null]);
});
