// A whole campaign against the fake store, through the real command functions
// and the real Excel writer: select → issue (date guard, dry run, batches, a
// lost response) → reminders → usage → verify → export. Checks the promises
// that matter most: one card per person, at most one reminder per person per
// round, only unused cards reminded, and an Excel that reflects every step.
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { installFakeShopify } from './fake-shopify.js';
import { makeCustomer, makeOrder, makeActiveOrder, makeGiftCardOrder, testConfig, memoryLog } from './helpers.js';
import { campaignPaths, readJson, readJournal, foldJournal } from '../src/campaign.js';
import { resetClient } from '../src/shopify.js';
import { runSelect } from '../src/select/index.js';
import { runIssue } from '../src/issue.js';
import { runRemind } from '../src/remind.js';
import { runUsage } from '../src/usage.js';
import { runVerify } from '../src/verify.js';
import { runExport } from '../src/export-command.js';

const { config, dir, cleanup } = testConfig({ SENT_TAG: 'OCT26RTPROMO' });
const dryConfig = { ...config, dryRun: true };
const paths = campaignPaths(config);
after(() => {
  resetClient();
  cleanup();
});

const cid = (n) => `gid://shopify/Customer/${n}`;
const ADDR = (address1) => ({ address1, address2: '', city: 'Austin', provinceCode: 'TX', zip: '78701', countryCodeV2: 'US' });

function cellText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && 'result' in v) return cellText(v.result);
  if (typeof v === 'object' && 'richText' in v) return v.richText.map((r) => r.text).join('');
  if (v instanceof Date) return v.toISOString();
  return String(v);
}

/**
 * Data rows of the table whose header row starts with `firstHeader`, as objects
 * keyed by header. The table ends at the first blank row (a later section of
 * the same sheet is not part of it); eachRow skips empty rows, so a gap in the
 * row numbers marks that blank row.
 */
function table(ws, firstHeader) {
  let header = null;
  let previous = 0;
  let ended = false;
  const rows = [];
  ws.eachRow((row, n) => {
    if (ended) return;
    const values = row.values.slice(1).map(cellText);
    if (!header) {
      if (values[0] === firstHeader) {
        header = values;
        previous = n;
      }
      return;
    }
    if (n !== previous + 1 || values.every((v) => v === '')) {
      ended = true;
      return;
    }
    previous = n;
    rows.push(Object.fromEntries(header.map((h, i) => [h, values[i] ?? ''])));
  });
  return { header, rows };
}

async function readWorkbook(file) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  return wb;
}

test('e2e: a whole campaign with the real commands and the real Excel', { timeout: 180_000 }, async () => {
  let nowMs = Date.parse('2026-10-03T19:00:00Z'); // Sat 10/3 noon in Los Angeles
  const now = () => new Date(nowMs);
  const at = (iso) => {
    nowMs = Date.parse(iso);
  };
  const sleep = async (ms) => {
    nowMs += ms; // simulated time passes while a command waits
  };
  const log = memoryLog();
  const ok = (result, what) => assert.equal(result.exitCode, 0, `${what} failed:\n${log.lines.slice(-40).join('\n')}`);

  const busy = makeCustomer({ n: 8, address: ADDR('8 Busy St'), lastOrder: makeOrder({ n: 801, createdAt: '2026-09-15T12:00:00Z', total: '20.00' }) });
  const customers = [
    makeCustomer({ n: 1, first: 'Ann', lastOrder: makeOrder({ n: 101, createdAt: '2026-06-20T12:00:00Z', total: '150.00' }) }), // 10% = $15.00 → $15.33
    makeCustomer({ n: 2, lastOrder: makeOrder({ n: 201, createdAt: '2026-05-20T12:00:00Z', total: '83.40' }) }), // $8.34 → $10.77
    makeCustomer({ n: 3, lastOrder: makeOrder({ n: 301, createdAt: '2026-04-20T12:00:00Z', total: '350.00' }) }), // $35.00 → $19.77
    makeCustomer({ n: 4, createdAt: '2025-03-01T00:00:00Z' }), // never ordered: 10% of the $160.85 average = $16.09 → $19.77
    makeCustomer({ n: 5, address: ADDR('5 Twin St'), lastOrder: makeOrder({ n: 501, createdAt: '2026-01-10T12:00:00Z', total: '40.00' }) }), // same address as 6
    makeCustomer({ n: 6, address: ADDR('5 Twin Street'), lastOrder: makeOrder({ n: 601, createdAt: '2026-03-10T12:00:00Z', total: '60.00' }) }), // kept: newer order
    makeCustomer({ n: 7, marketingState: 'UNSUBSCRIBED' }),
    busy, // ordered on 9/15
  ];
  const fake = installFakeShopify({ customers, orders: [makeActiveOrder({ n: 801, customer: busy, createdAt: '2026-09-15T12:00:00Z' })], now: () => nowMs });
  const creates = () => fake.opsNamed('GiftCardCreate').length;
  const notifies = () => fake.state.calls.notify.length;
  const cardOf = (n) => fake.state.giftCards.filter((g) => g.customer?.id === cid(n));

  try {
    // ---- 10/3 select ----------------------------------------------------------
    ok(await runSelect({ config, log, now, sleep, bulkPollMs: 1 }), 'select');
    const selection = readJson(paths.selection);
    assert.deepEqual(selection.recipients.map((r) => r.customerId), [cid(1), cid(2), cid(3), cid(6), cid(4)]);
    assert.deepEqual(selection.recipients.map((r) => r.amountCents), [1533, 1077, 1977, 1077, 1977]);
    assert.equal(selection.stats.totalCents, 1533 + 1077 + 1977 + 1077 + 1977);
    assert.ok(fs.existsSync(paths.excel), 'select writes the Excel');

    // ---- live issue before LAUNCH_DATE is refused; a dry run is allowed ---------
    assert.equal((await runIssue({ config, limit: 2, log, now, sleep })).exitCode, 2);
    assert.equal(creates(), 0);
    const journalBefore = readJournal(paths.journal).length;
    ok(await runIssue({ config: dryConfig, limit: 2, log, now, sleep }), 'dry-run issue');
    assert.equal(creates(), 0, 'a dry run creates nothing');
    assert.equal(fake.opsNamed('TagsAdd').length, 0, 'a dry run tags nobody');
    assert.deepEqual(readJournal(paths.journal).slice(journalBefore).map((e) => e.op), ['run.start', 'run.end'], 'a dry run only records that it ran');
    assert.ok(fs.existsSync(path.join(paths.previewDir, 'first.html')), 'the dry run renders the first email');

    // ---- 10/5 issue: 2, then the rest with one lost response --------------------
    at('2026-10-05T16:00:00Z');
    ok(await runIssue({ config, limit: 2, log, now, sleep }), 'issue --limit 2');
    assert.equal(creates(), 2);
    assert.deepEqual(fake.state.calls.create.map((c) => c.customerId), [cid(1), cid(2)]);
    assert.deepEqual(fake.state.calls.create.map((c) => c.initialAmount.amount), ['15.33', '10.77']);
    assert.ok(fake.state.calls.create.every((c) => c.note === 'gift-card-promo [campaign:2026-10]' && c.templateSuffix === 'gift-card-promo' && c.expiresOn === '2026-10-19'));

    at('2026-10-05T16:30:00Z');
    fake.state.failures.GiftCardCreate = [{ kind: 'appliedThenLost' }]; // Shopify creates c3's card, the answer is lost
    ok(await runIssue({ config, limit: 10, log, now, sleep }), 'issue --limit 10');
    for (const n of [1, 2, 3, 4, 6]) assert.equal(cardOf(n).length, 1, `customer ${n} has exactly one card`);
    for (const n of [5, 7, 8]) assert.equal(cardOf(n).length, 0, `customer ${n} has no card`);
    const state = foldJournal(readJournal(paths.journal));
    for (const r of selection.recipients) assert.equal(state.customers.get(r.customerId).status, 'done', `${r.customerId} done`);
    assert.equal(state.customers.get(cid(3)).reconciledFrom, 'issue-inline', 'the lost response was reconciled, not re-created');
    for (const n of [1, 2, 3, 4, 6]) assert.ok(fake.state.customers.find((c) => c.id === cid(n)).tags.includes('OCT26RTPROMO'));

    const createsSoFar = creates();
    ok(await runIssue({ config, limit: 10, log, now, sleep }), 'issue with nobody left');
    assert.equal(creates(), createsSoFar, 'a re-run never creates a second card');
    assert.equal((await runSelect({ config, log, now, sleep, bulkPollMs: 1 })).exitCode, 1, 'select is frozen once issuing started');

    // ---- 10/8 Ann uses $5 of her card ------------------------------------------
    const annCard = cardOf(1)[0];
    fake.spendCard(annCard.id, 500);
    fake.state.giftCardOrders.push(makeGiftCardOrder({
      n: 9001,
      customer: { id: cid(1) },
      createdAt: '2026-10-08T18:00:00Z',
      total: '45.00',
      payments: [{ giftCardNumericId: annCard.id.split('/').pop(), amount: '5.00' }],
      lineItems: [{ name: 'Gold Balloon Arch', sku: 'ARCH-1', quantity: 1, amount: '40.00' }],
    }));

    // ---- reminders: the date guard, then round 1 on 10/12 ------------------------
    at('2026-10-11T17:00:00Z');
    assert.equal((await runRemind({ config, round: 1, log, now, sleep })).exitCode, 2, 'round 1 is refused before REMIND_1_DATE');
    assert.equal(notifies(), 0);
    at('2026-10-12T17:00:00Z');
    ok(await runRemind({ config: dryConfig, round: 1, log, now, sleep }), 'dry-run remind');
    assert.equal(notifies(), 0, 'a dry run sends nothing');
    ok(await runRemind({ config, round: 1, log, now, sleep }), 'remind --round 1');
    const unusedIds = [2, 3, 6, 4].map((n) => cardOf(n)[0].id);
    assert.deepEqual(fake.state.calls.notify, unusedIds, 'only unused cards, in list order');
    ok(await runRemind({ config, round: 1, log, now, sleep }), 'remind --round 1 again');
    assert.equal(notifies(), 4, 'round 1 is sent at most once per person');

    // ---- 10/13 usage --------------------------------------------------------------
    at('2026-10-13T17:00:00Z');
    ok(await runUsage({ config, log, now }), 'usage');
    const usage = readJson(paths.usage);
    assert.equal(usage.summary.issuedCards, 5);
    assert.equal(usage.summary.usedCards, 1);
    assert.equal(usage.summary.usedCents, 500);
    assert.equal(usage.summary.orders, 1);
    assert.equal(usage.summary.giftCardCents, 500);
    assert.equal(usage.summary.customerPaidCents, 4000);
    assert.deepEqual(usage.orders[0].lineItems.map((l) => [l.name, l.quantity]), [['Gold Balloon Arch', 1]]);
    assert.ok(fs.existsSync(path.join(paths.usageDir, '2026-10-13.json')));

    // ---- 10/16 round 2: customer 3 has unsubscribed meanwhile ----------------------
    fake.state.customers.find((c) => c.id === cid(3)).defaultEmailAddress.marketingState = 'UNSUBSCRIBED';
    at('2026-10-16T17:00:00Z');
    ok(await runRemind({ config, round: 2, log, now, sleep }), 'remind --round 2');
    assert.deepEqual(fake.state.calls.notify.slice(4), [2, 6, 4].map((n) => cardOf(n)[0].id));
    const reminders = foldJournal(readJournal(paths.journal)).customers;
    assert.equal(reminders.get(cid(1)).reminders['1'].reason, 'used');
    assert.equal(reminders.get(cid(3)).reminders['2'].reason, 'not-subscribed');

    // ---- verify and export ------------------------------------------------------------
    ok(await runVerify({ config, log, now }), 'verify');
    const verify = readJson(paths.verify);
    assert.equal(verify.cardCount, 5);
    assert.deepEqual(verify.issues, [], `verify found problems: ${JSON.stringify(verify.issues)}`);
    fs.mkdirSync(path.join(dir, 'copy'));
    const out = path.join(dir, 'copy', 'report.xlsx');
    ok(await runExport({ config, out, log, now }), 'export');
    assert.ok(fs.existsSync(out));

    // ---- the final Excel ----------------------------------------------------------------
    const wb = await readWorkbook(paths.excel);
    assert.deepEqual(wb.worksheets.map((ws) => ws.name), ['使用报告', '汇总', '发放名单', '未入选', '同地址重复', '使用明细', '核对', '操作日志', '说明']);
    const list = table(wb.getWorksheet('发放名单'), '序号');
    assert.equal(list.rows.length, 5);
    assert.ok(list.header.includes('是否有 OCT26RTPROMO'), 'the tag column is there');
    assert.deepEqual(list.rows.map((r) => r['状态']), ['已完成', '已完成', '已完成', '已完成', '已完成']);
    assert.deepEqual(list.rows.map((r) => r['是否有 OCT26RTPROMO']), ['是', '是', '是', '是', '是']);
    assert.deepEqual(list.rows.map((r) => r['礼品卡金额']), ['15.33', '10.77', '19.77', '10.77', '19.77']);
    assert.equal(list.rows[0]['第 1 次提醒'], '跳过：已用过卡');
    assert.ok(list.rows.slice(1).every((r) => r['第 1 次提醒'].startsWith('已发 ')), list.rows.map((r) => r['第 1 次提醒']).join(' | '));
    assert.equal(list.rows[2]['第 2 次提醒'], '跳过：未订阅营销邮件');
    assert.equal(list.rows[0]['已使用金额'], '5');
    const usedSheet = table(wb.getWorksheet('使用明细'), '订单号');
    assert.equal(usedSheet.rows.length, 1);
    assert.match(usedSheet.rows[0]['买了什么'], /Gold Balloon Arch × 1/);
    assert.equal(fs.readdirSync(paths.dir).filter((f) => f.includes('.tmp')).length, 0, 'no temp files left in the campaign folder');
  } finally {
    fake.restore();
  }
});
