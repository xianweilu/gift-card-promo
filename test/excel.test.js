// Excel report: build real fixtures (selection via the real rules, journal via
// appendJournal), write the workbook, read it back with ExcelJS and check it.
import { test, describe, beforeEach, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';

import { writeReport, statusLabel, remindLabel, adminUrl, effectiveStatus, numericId, OPEN_IN_EXCEL_WARNING } from '../src/report/excel.js';
import { summaryText, detailText, issueSkipText, remindSkipText, errorText, USAGE_SPLIT_NOTE, roundTagName } from '../src/report/labels.js';
import { campaignPaths, appendJournal, acquireRunLock, writeJsonAtomic } from '../src/campaign.js';
import { initClient, resetClient, gql, throwIfUserErrors } from '../src/shopify.js';
import { addTag } from '../src/customers.js';
import { campaignNote } from '../src/giftcards.js';
// A namespace import: the round-tag contract test below names a missing roundTag export itself.
import * as remindModule from '../src/remind.js';
import { installFakeShopify } from './fake-shopify.js';
import { buildTestSelection, selectionParams } from '../src/select/selection.js';
import { parseCustomer } from '../src/select/rules.js';
import { testConfig, selectionFixture, makeCustomer, makeOrder, makeActiveOrder, memoryLog, gid, NOW_ISO } from './helpers.js';
// The workbook's raw XML parts (what Microsoft Excel actually parses): shared with the other report tests.
import { xlsxParts, worksheetPartsByName, worksheetChildren, worksheetOrderProblems, malformedParts, HAS_XMLLINT } from './report-fixture.js';

const TZ = 'America/Los_Angeles';
const FIXED_NOW = () => new Date('2026-10-12T17:00:00.000Z');
const FAST_LOCK = { pollMs: 5 };

const RECIPIENT_HEADERS = [
  '序号', '状态', '批次', '客户 ID', '姓名', '邮箱', '营销状态', '客户类型', '上次有效订单号', '上次下单日期',
  '最近订单渠道', '距今天数', '上次订单总额', '金额算式', '档位', '礼品卡金额', '同地址账户数', '同地址其他账户',
  '是否有 gift-card-sent-2026-10', '礼品卡 ID', '卡号后 4 位', '建卡时间', '打 tag 时间', '第 1 次提醒', '第 2 次提醒',
  '已使用金额', '剩余余额', '使用的订单', '备注/错误', '城市', '州', '邮编', '订单数', '累计消费', '注册日期',
];
const COL = Object.fromEntries(RECIPIENT_HEADERS.map((h, i) => [h, i + 1]));
COL.tag = COL['是否有 gift-card-sent-2026-10'];

const cust = (n) => gid('Customer', n);
const card = (n) => gid('GiftCard', n);
const order = (n) => gid('Order', n);
/** Store-local wall time as the Date ExcelJS reads back (UTC fields = local time). */
const wall = (y, mo, d, h = 0, mi = 0, s = 0) => new Date(Date.UTC(y, mo - 1, d, h, mi, s));

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const SAME_ADDRESS = { address1: '500 Elm Street', address2: '', city: 'Austin', provinceCode: 'TX', zip: '78701', countryCodeV2: 'US' };
const ACTIVE_ADDRESS = { address1: '77 Oak Ave', address2: '', city: 'Austin', provinceCode: 'TX', zip: '78702', countryCodeV2: 'US' };

/** The fixture customers; `patch` overrides makeCustomer fields per customer number (e.g. { 1: { first: 'X' } }). */
function fixtureCustomers(patch = {}) {
  const c12 = makeCustomer({ n: 12, address: ACTIVE_ADDRESS, lastOrder: makeOrder({ n: 120, createdAt: '2026-08-20T12:00:00Z', total: '30.00' }) });
  const c = (n, fields = {}) => makeCustomer({ n, ...fields, ...(patch[n] ?? {}) });
  return {
    customers: [
      c(1, { createdAt: '2024-01-01T00:00:00Z', lastOrder: makeOrder({ n: 101, createdAt: '2026-05-10T18:00:00Z', total: '150.00' }) }),
      c(2, { lastOrder: makeOrder({ n: 102, createdAt: '2026-04-01T18:00:00Z', total: '83.40' }) }),
      c(3, { address: SAME_ADDRESS, lastOrder: makeOrder({ n: 103, createdAt: '2026-03-15T18:00:00Z', total: '60.00' }), numberOfOrders: 3 }),
      c(4, { address: SAME_ADDRESS, lastOrder: makeOrder({ n: 104, createdAt: '2026-02-01T18:00:00Z', total: '40.00' }) }),
      c(5, { createdAt: '2025-06-01T12:00:00Z' }), // never ordered
      c(6, { lastOrder: makeOrder({ n: 60, createdAt: '2026-02-20T18:00:00Z', total: '0.00' }) }), // last order $0
      c(7, { lastOrder: makeOrder({ n: 70, createdAt: '2026-03-01T18:00:00Z', total: '90.00', cancelled: true }) }), // last order cancelled
      c(8, { lastOrder: makeOrder({ n: 80, createdAt: '2026-01-05T18:00:00Z', total: '45.00', cancelled: true }) }), // only cancelled orders
      c(9, { lastOrder: makeOrder({ n: 90, createdAt: '2026-08-15T18:00:00Z', total: '25.00' }) }), // recent order
      c(10, { marketingState: 'NOT_SUBSCRIBED' }), // counted only, not listed
      c(11, { address: ACTIVE_ADDRESS }), // same address as an active customer
      c12,
      c(13, { tags: ['WHS'] }),
      c(14, { lastOrder: makeOrder({ n: 140, createdAt: '2026-01-20T18:00:00Z', total: '40.00' }) }),
      c(15, { lastOrder: makeOrder({ n: 150, createdAt: '2026-06-01T18:00:00Z', total: '200.00' }) }),
    ],
    activeOrders: [makeActiveOrder({ n: 900, customer: c12, createdAt: '2026-08-20T12:00:00Z' })],
    history: {
      [cust(6)]: [makeOrder({ n: 60, createdAt: '2026-02-20T18:00:00Z', total: '0.00' }), makeOrder({ n: 61, createdAt: '2026-01-10T18:00:00Z', total: '23.00' })],
      [cust(7)]: [makeOrder({ n: 70, createdAt: '2026-03-01T18:00:00Z', total: '90.00', cancelled: true }), makeOrder({ n: 71, createdAt: '2025-12-01T18:00:00Z', total: '350.00' })],
      [cust(8)]: [makeOrder({ n: 80, createdAt: '2026-01-05T18:00:00Z', total: '45.00', cancelled: true })],
    },
  };
}

function journalAt(paths, items) {
  for (const [t, entry] of items) appendJournal(paths.journal, entry, { now: () => t });
}

// run.end summaries exactly as the commands write them (src/issue.js emptySummary, src/remind.js newSummary).
const ISSUE_SUMMARY = {
  batch: 1, dryRun: false, attempted: 7, created: 4, tagged: 3, tagFixed: 0, reconciled: 0, skipped: { 'ordered-since-snapshot': 1 },
  failed: 1, rejected: 0, unknown: 1, amountCents: 6000, tagFailed: 1, reconciledNone: 0, stillUnknown: 0, seqFrom: 1, seqTo: 9,
};
const ISSUE_SUMMARY_TEXT = '尝试 7，建卡 4，打 tag 3，跳过（导出后下过单 1），失败 1，结果不明 1，金额 $60.00，打 tag 失败 1，序号 1–9';
const ISSUE_DRY_SUMMARY = {
  batch: 2, dryRun: true, attempted: 1, created: 0, tagged: 0, tagFixed: 1, reconciled: 0, skipped: {},
  failed: 0, rejected: 0, unknown: 0, amountCents: 1977, tagFailed: 0, reconciledNone: 0, stillUnknown: 2, seqFrom: 9, seqTo: 9,
};
const ISSUE_DRY_SUMMARY_TEXT = '将建卡 1，补打 tag 1，金额 $19.77，仍需人工核对 2，序号 9';
const REMIND_SUMMARY = {
  round: 1, dryRun: false, eligible: 3, planned: 3, attempted: 3, sent: 1, failed: 1, rejected: 0, unknown: 1, retriedUnknown: 0,
  skipped: { used: 1 }, alreadySent: 0, previouslyFailed: 0, waitingUnknown: 0, notIssued: 5, stopped: null,
};
const REMIND_SUMMARY_TEXT = '符合条件 3，本次要发 3，尝试 3，已发 1，失败 1，结果不明 1，跳过（已用过卡 1），未建卡 5';

/** Journal covering every issue status and reminder outcome. */
function writeFixtureJournal(paths, amount) {
  const R1 = '20261005160000-1';
  const R2 = '20261005170000-1';
  const R3 = '20261012160000-1';
  const R4 = '20261016160000-1';
  const p2 = (n) => String(n).padStart(2, '0');
  const T = (hh, mm, ss = 0) => `2026-10-05T${p2(hh)}:${p2(mm)}:${p2(ss)}.000Z`;
  journalAt(paths, [
    [T(16, 0), { op: 'run.start', run: R1, command: 'issue', dryRun: false, batch: 1, limit: 20 }],
    // c1: done
    [T(16, 0, 1), { op: 'create.start', cid: cust(1), amountCents: amount(1), batch: 1, run: R1 }],
    [T(16, 0, 2), { op: 'create.ok', cid: cust(1), giftCardId: card(1001), last4: 'x001', amountCents: amount(1), batch: 1, run: R1 }],
    [T(16, 0, 3), { op: 'tag.ok', cid: cust(1), run: R1 }],
    // c2: created, tag failed
    [T(16, 1, 1), { op: 'create.start', cid: cust(2), amountCents: amount(2), batch: 1, run: R1 }],
    [T(16, 1, 2), { op: 'create.ok', cid: cust(2), giftCardId: card(1002), last4: 'x002', amountCents: amount(2), batch: 1, run: R1 }],
    [T(16, 1, 3), { op: 'tag.fail', cid: cust(2), error: 'tagsAdd rejected: boom', run: R1 }],
    // c3: failed
    [T(16, 2, 1), { op: 'create.start', cid: cust(3), amountCents: amount(3), batch: 1, run: R1 }],
    [T(16, 2, 2), { op: 'create.fail', cid: cust(3), error: 'giftCardCreate rejected: input: Customer is invalid [INVALID]', run: R1 }],
    // c5: unknown
    [T(16, 3, 1), { op: 'create.start', cid: cust(5), amountCents: amount(5), batch: 1, run: R1 }],
    [T(16, 3, 9), { op: 'create.unknown', cid: cust(5), error: 'Network error calling Shopify', run: R1 }],
    // c6: skipped before issuing
    [T(16, 4, 0), { op: 'skip', cid: cust(6), reason: 'ordered-since-snapshot', detail: '#7001', batch: 1, run: R1 }],
    // c7: create.start without an outcome (crash) → in_progress
    [T(16, 5, 0), { op: 'create.start', cid: cust(7), amountCents: amount(7), batch: 1, run: R1 }],
    // c14 / c15: done
    [T(16, 6, 1), { op: 'create.start', cid: cust(14), amountCents: amount(14), batch: 1, run: R1 }],
    [T(16, 6, 2), { op: 'create.ok', cid: cust(14), giftCardId: card(1014), last4: 'x014', amountCents: amount(14), batch: 1, run: R1 }],
    [T(16, 6, 3), { op: 'tag.ok', cid: cust(14), run: R1 }],
    [T(16, 7, 1), { op: 'create.start', cid: cust(15), amountCents: amount(15), batch: 1, run: R1 }],
    [T(16, 7, 2), { op: 'create.ok', cid: cust(15), giftCardId: card(1015), last4: 'x015', amountCents: amount(15), batch: 1, run: R1 }],
    [T(16, 7, 3), { op: 'tag.ok', cid: cust(15), run: R1 }],
    // run.end summaries in the shapes src/issue.js and src/remind.js write
    [T(16, 8, 0), { op: 'run.end', run: R1, summary: ISSUE_SUMMARY, exitCode: 1 }],
    // a dry run
    [T(17, 0, 0), { op: 'run.start', run: R2, command: 'issue', dryRun: true, batch: 2, limit: 500 }],
    [T(17, 0, 5), { op: 'run.end', run: R2, summary: ISSUE_DRY_SUMMARY, exitCode: 0 }],
    // reminders, round 1
    ['2026-10-12T16:00:00.000Z', { op: 'run.start', run: R3, command: 'remind', dryRun: false, options: { round: 1 } }],
    ['2026-10-12T16:04:59.000Z', { op: 'remind.start', cid: cust(1), round: 1, giftCardId: card(1001), run: R3 }],
    ['2026-10-12T16:05:00.000Z', { op: 'remind.ok', cid: cust(1), round: 1, run: R3 }],
    ['2026-10-12T16:05:10.000Z', { op: 'remind.skip', cid: cust(2), round: 1, reason: 'used', run: R3 }],
    ['2026-10-12T16:05:20.000Z', { op: 'remind.start', cid: cust(15), round: 1, giftCardId: card(1015), run: R3 }],
    ['2026-10-12T16:05:50.000Z', { op: 'remind.unknown', cid: cust(15), round: 1, error: 'Network error calling Shopify', run: R3 }],
    ['2026-10-12T16:06:00.000Z', { op: 'remind.start', cid: cust(14), round: 1, giftCardId: card(1014), run: R3 }],
    ['2026-10-12T16:06:01.000Z', { op: 'remind.fail', cid: cust(14), round: 1, error: 'giftCardSendNotificationToCustomer rejected: boom', run: R3 }],
    ['2026-10-12T16:07:00.000Z', { op: 'run.end', run: R3, summary: REMIND_SUMMARY, exitCode: 1 }],
    // round 2, interrupted after remind.start (no run.end)
    ['2026-10-16T16:00:00.000Z', { op: 'run.start', run: R4, command: 'remind', dryRun: false, options: { round: 2 } }],
    ['2026-10-16T16:00:05.000Z', { op: 'remind.start', cid: cust(1), round: 2, giftCardId: card(1001), run: R4 }],
  ]);
}

/** A tag snapshot taken after the fixture journal's tags: c1, c14, c15 (tag.ok) and c2 (tag.fail, but tagged in Shopify). */
function writeFixtureTags(paths, config) {
  writeJsonAtomic(paths.tags, { fetchedAt: '2026-10-06T00:00:00.000Z', tag: config.sentTag, ids: [cust(1), cust(2), cust(14), cust(15)] });
}

function writeFixtureVerify(paths) {
  writeJsonAtomic(paths.verify, {
    version: 1,
    verifiedAt: '2026-10-07T17:00:00.000Z',
    cardCount: 4,
    taggedCount: 3,
    counts: { 'tag-missing': 1, 'still-unknown': 1 },
    issues: [
      { type: 'tag-missing', customerId: cust(2), giftCardId: card(1002), journal: '已建卡，未打 tag', shopify: '客户没有 tag', action: '运行 issue 补打 tag' },
      { type: 'still-unknown', customerId: cust(5), giftCardId: null, journal: '建卡结果不明', shopify: '没有找到这个客户的卡', action: '再运行一次 verify' },
    ],
  });
}

function usageFixture(amount) {
  const cards = [
    { giftCardId: card(1001), customerId: cust(1), last4: 'x001', initialCents: amount(1), balanceCents: 0, usedCents: amount(1), enabled: true, expiresOn: '2026-10-19', createdAt: '2026-10-05T16:00:02.000Z' },
    { giftCardId: card(1002), customerId: cust(2), last4: 'x002', initialCents: amount(2), balanceCents: amount(2) - 500, usedCents: 500, enabled: true, expiresOn: '2026-10-19', createdAt: '2026-10-05T16:01:02.000Z' },
    { giftCardId: card(1014), customerId: cust(14), last4: 'x014', initialCents: amount(14), balanceCents: amount(14), usedCents: 0, enabled: true, expiresOn: '2026-10-19', createdAt: '2026-10-05T16:06:02.000Z' },
    { giftCardId: card(1015), customerId: cust(15), last4: 'x015', initialCents: amount(15), balanceCents: amount(15), usedCents: 0, enabled: true, expiresOn: '2026-10-19', createdAt: '2026-10-05T16:07:02.000Z' },
  ];
  const issuedCents = cards.reduce((a, c) => a + c.initialCents, 0);
  const usedCents = amount(1) + 500;
  const giftCardCents = amount(1) + 700 - 200;
  return {
    version: 1,
    fetchedAt: '2026-10-12T15:00:00.000Z',
    timezone: TZ,
    campaignStartIso: '2026-10-01T19:00:00.000Z',
    cards,
    payments: [
      { orderId: order(5002), orderName: '#5002', orderCreatedAt: '2026-10-07T19:30:00.000Z', orderCustomerId: cust(2), giftCardId: card(1002), cardCustomerId: cust(2), kind: 'SALE', amountCents: 700, netCents: 700, processedAt: '2026-10-07T19:30:05.000Z' },
      { orderId: order(5001), orderName: '#5001', orderCreatedAt: '2026-10-06T18:00:00.000Z', orderCustomerId: cust(1), giftCardId: card(1001), cardCustomerId: cust(1), kind: 'SALE', amountCents: amount(1), netCents: amount(1), processedAt: '2026-10-06T18:00:05.000Z' },
      { orderId: order(5002), orderName: '#5002', orderCreatedAt: '2026-10-07T19:30:00.000Z', orderCustomerId: cust(2), giftCardId: card(1002), cardCustomerId: cust(2), kind: 'REFUND', amountCents: 200, netCents: -200, processedAt: '2026-10-08T10:00:00.000Z' },
    ],
    orders: [
      {
        orderId: order(5001), orderName: '#5001', createdAt: '2026-10-06T18:00:00.000Z', customerId: cust(1), totalCents: 6000, campaignCardCents: amount(1), cancelled: false, giftCardIds: [card(1001)],
        lineItems: [{ name: 'Gold Balloon', sku: 'GB-1', quantity: 2, amountCents: 2000 }, { name: 'Helium Tank', sku: 'HT-1', quantity: 1, amountCents: 4000 }],
      },
      {
        orderId: order(5002), orderName: '#5002', createdAt: '2026-10-07T19:30:00.000Z', customerId: cust(2), totalCents: 2500, campaignCardCents: 500, cancelled: true, giftCardIds: [card(1002)],
        lineItems: [{ name: 'Gold Balloon', sku: 'GB-1', quantity: 1, amountCents: 2500 }],
      },
    ],
    unmatched: [{ orderId: order(5003), orderName: '#5003', orderCreatedAt: '2026-10-08T20:00:00.000Z', customerId: cust(9), amountCents: 1000, processedAt: '2026-10-08T20:00:05.000Z', reason: 'no-gift-card-id' }],
    summary: {
      issuedCards: 4,
      issuedCents,
      usedCards: 2,
      usedCents,
      usedRate: 0.5,
      usedCentsRate: usedCents / issuedCents,
      orders: 2,
      ordersTotalCents: 8500,
      avgOrderCents: 4250,
      giftCardCents,
      customerPaidCents: 8500 - giftCardCents,
      byTier: [
        { tier: 0, label: '$10.77', issued: 2, used: 1, rate: 0.5, usedCents: 500 },
        { tier: 1, label: '$15.33', issued: 1, used: 1, rate: 1, usedCents: amount(1) },
        { tier: 2, label: '$19.77', issued: 1, used: 0, rate: 0, usedCents: 0 },
      ],
      byKind: [{ kind: 'ordered', issued: 4, used: 2, rate: 0.5, usedCents }],
      daily: [
        { date: '2026-10-06', newCardsUsed: 1, orders: 1, ordersTotalCents: 6000, cumulativeCardsUsed: 1, cumulativeRate: 0.25 },
        { date: '2026-10-07', newCardsUsed: 1, orders: 1, ordersTotalCents: 2500, cumulativeCardsUsed: 2, cumulativeRate: 0.5 },
      ],
      topProducts: [{ name: 'Gold Balloon', quantity: 3, amountCents: 4500 }, { name: 'Helium Tank', quantity: 1, amountCents: 4000 }],
    },
  };
}

async function openBook(file) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  return wb;
}

/** Plain value of a cell: formula → its cached result. */
function plain(value) {
  if (value && typeof value === 'object' && !(value instanceof Date)) {
    if ('result' in value) return value.result;
    if (value.richText) return value.richText.map((r) => r.text).join('');
  }
  return value;
}

const cellValue = (ws, r, c) => plain(ws.getRow(r).getCell(c).value);

function rowValues(ws, r, count) {
  return Array.from({ length: count }, (_, i) => cellValue(ws, r, i + 1));
}

/** Row number of the first row whose column `col` shows `text` (null when none). */
function findRow(ws, col, text) {
  for (let r = 1; r <= ws.rowCount; r += 1) if (cellValue(ws, r, col) === text) return r;
  return null;
}

function tmpLeftovers(dir) {
  return fs.readdirSync(dir).filter((f) => f.includes('.tmp-'));
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('excel report', () => {
  let env;
  let paths;
  let selection;
  const amount = (n) => selection.recipients.find((r) => r.customerId === cust(n)).amountCents;
  const seqOf = (n) => selection.recipients.find((r) => r.customerId === cust(n)).seq;

  beforeEach(async () => {
    env = testConfig();
    paths = campaignPaths(env.config);
    selection = await selectionFixture(env.config, fixtureCustomers());
  });
  afterEach(() => env.cleanup());

  function writeAllInputs() {
    writeFixtureJournal(paths, amount);
    writeFixtureTags(paths, env.config);
    writeFixtureVerify(paths);
    writeJsonAtomic(paths.usage, usageFixture(amount));
  }

  test('the fixture selection covers the cases the report must show', () => {
    assert.deepEqual(selection.recipients.map((r) => Number(r.numericId)).sort((a, b) => a - b), [1, 2, 3, 5, 6, 7, 8, 14, 15]);
    const byId = Object.fromEntries(selection.recipients.map((r) => [r.numericId, r]));
    assert.equal(byId['6'].basisWhy, 'last-zero');
    assert.equal(byId['7'].basisWhy, 'last-cancelled');
    assert.equal(byId['8'].neverReason, 'only-cancelled-or-zero');
    assert.equal(byId['5'].kind, 'never');
    assert.equal(selection.duplicates.length, 1);
    assert.ok(selection.notSelected.some((n) => n.primaryCode === 'active-address'));
  });

  test('sheet names and order with usage.json and verify.json', async () => {
    writeAllInputs();
    const log = memoryLog();
    const result = await writeReport({ config: env.config, paths, log, now: FIXED_NOW, lockOptions: FAST_LOCK });
    assert.equal(result.file, paths.excel);
    assert.equal(result.out, null);
    assert.deepEqual(result.warnings, []);
    const wb = await openBook(paths.excel);
    assert.deepEqual(wb.worksheets.map((w) => w.name), ['使用报告', '汇总', '发放名单', '未入选', '同地址重复', '使用明细', '核对', '操作日志', '说明']);
    assert.deepEqual(tmpLeftovers(paths.dir), []);
    assert.equal(fs.existsSync(paths.excelLock), false);
  });

  test('sheet order without usage.json and verify.json', async () => {
    writeFixtureJournal(paths, amount);
    await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    const wb = await openBook(paths.excel);
    assert.deepEqual(wb.worksheets.map((w) => w.name), ['汇总', '发放名单', '未入选', '同地址重复', '操作日志', '说明']);
  });

  test('发放名单: headers, one row per recipient in seq order, statuses, tag column, reminders and notes', async () => {
    writeAllInputs();
    await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    const ws = (await openBook(paths.excel)).getWorksheet('发放名单');

    assert.deepEqual(rowValues(ws, 1, RECIPIENT_HEADERS.length), RECIPIENT_HEADERS);
    assert.equal(cellValue(ws, 1, RECIPIENT_HEADERS.length + 1), null);
    assert.equal(ws.actualRowCount, selection.recipients.length + 1);
    assert.deepEqual(Array.from({ length: selection.recipients.length }, (_, i) => cellValue(ws, i + 2, COL['序号'])), selection.recipients.map((r) => r.seq));

    const row = (n) => findRow(ws, COL['客户 ID'], String(n));
    const at = (n, title) => cellValue(ws, row(n), COL[title] ?? title);

    // statuses (c7 is in_progress but issue is not running → 需人工核对)
    assert.equal(at(1, '状态'), '已完成');
    assert.equal(at(2, '状态'), '已建卡未打tag');
    assert.equal(at(3, '状态'), '失败');
    assert.equal(at(5, '状态'), '需人工核对');
    assert.equal(at(6, '状态'), '发放前跳过');
    assert.equal(at(7, '状态'), '需人工核对');
    assert.equal(at(8, '状态'), '待发放');
    assert.equal(at(14, '状态'), '已完成');
    assert.equal(at(1, '批次'), 1);
    assert.equal(at(8, '批次'), null);

    // whole row filled by status colour (including empty cells); pending stays unfilled
    const fillOf = (n, c) => ws.getRow(row(n)).getCell(c).fill?.fgColor?.argb ?? null;
    assert.equal(fillOf(1, 1), 'FFE2EFDA');
    assert.equal(fillOf(1, COL['备注/错误']), 'FFE2EFDA');
    assert.equal(fillOf(2, COL['姓名']), 'FFFFF2CC');
    assert.equal(fillOf(3, COL['城市']), 'FFFFD7D7');
    assert.equal(fillOf(5, COL['礼品卡 ID']), 'FFFCE4D6');
    assert.equal(fillOf(6, COL['档位']), 'FFEDEDED');
    assert.equal(fillOf(8, COL['姓名']), null);

    // tag column: the tag snapshot (taken after every tag.ok) has c1 and c2 (whose tagsAdd failed but who is tagged in Shopify)
    assert.equal(at(1, 'tag'), '是');
    assert.equal(at(2, 'tag'), '是');
    assert.equal(at(3, 'tag'), '否');
    assert.equal(at(8, 'tag'), '否');

    // reminder cells (store-local MM-DD HH:MM)
    assert.equal(at(1, '第 1 次提醒'), '已发 10-12 09:05');
    assert.equal(at(1, '第 2 次提醒'), '结果不明'); // remind.start without outcome, remind not running
    assert.equal(at(2, '第 1 次提醒'), '跳过：已用过卡');
    assert.equal(at(15, '第 1 次提醒'), '结果不明');
    assert.equal(at(14, '第 1 次提醒'), '失败：Shopify 拒绝：boom');
    assert.equal(at(3, '第 1 次提醒'), null);
    assert.equal(ws.getRow(row(14)).getCell(COL['第 1 次提醒']).font?.color?.argb, 'FFC00000');

    // links are HYPERLINK formulas
    const idCell = ws.getRow(row(1)).getCell(COL['客户 ID']).value;
    assert.deepEqual(idCell, { formula: 'HYPERLINK("https://admin.shopify.com/store/teststore/customers/1","1")', result: '1' });
    assert.deepEqual(ws.getRow(row(1)).getCell(COL['上次有效订单号']).value, { formula: 'HYPERLINK("https://admin.shopify.com/store/teststore/orders/101","#101")', result: '#101' });
    assert.deepEqual(ws.getRow(row(1)).getCell(COL['礼品卡 ID']).value, { formula: 'HYPERLINK("https://admin.shopify.com/store/teststore/gift_cards/1001","1001")', result: '1001' });
    assert.equal(ws.getRow(row(5)).getCell(COL['上次有效订单号']).value, null);

    // recipient details
    assert.equal(at(1, '姓名'), 'First1 Last1');
    assert.equal(at(1, '邮箱'), 'c1@example.org');
    assert.equal(at(1, '营销状态'), '已订阅');
    assert.equal(at(1, '客户类型'), '有下单');
    assert.equal(at(5, '客户类型'), '从没下单');
    assert.deepEqual(at(1, '上次下单日期'), wall(2026, 5, 10));
    assert.equal(at(1, '最近订单渠道'), '网店');
    assert.equal(at(1, '距今天数'), (Date.UTC(2026, 9, 1) - Date.UTC(2026, 4, 10)) / 86_400_000);
    assert.equal(at(1, '上次订单总额'), 150);
    assert.equal(at(1, '金额算式'), '10% × $150.00 = $15.00 → 档位 $15.33');
    assert.equal(at(1, '档位'), '$15.33');
    assert.equal(at(1, '礼品卡金额'), 15.33);
    assert.equal(at(7, '档位'), '$19.77');
    assert.equal(ws.getRow(row(1)).getCell(COL['礼品卡金额']).numFmt, '"$"#,##0.00');
    assert.equal(at(1, '卡号后 4 位'), 'x001');
    assert.deepEqual(at(1, '建卡时间'), wall(2026, 10, 5, 9, 0, 2)); // 16:00:02Z = 09:00 PDT
    assert.equal(ws.getRow(row(1)).getCell(COL['建卡时间']).numFmt, 'yyyy-mm-dd hh:mm');
    assert.deepEqual(at(1, '打 tag 时间'), wall(2026, 10, 5, 9, 0, 3));
    assert.equal(at(2, '打 tag 时间'), null);
    assert.deepEqual(at(1, '注册日期'), wall(2023, 12, 31)); // 2024-01-01T00:00Z is still Dec 31 in Los Angeles
    assert.equal(at(1, '邮编'), '70001');

    // duplicate group, kept account
    assert.equal(at(3, '同地址账户数'), 2);
    assert.equal(at(3, '同地址其他账户'), '4');
    assert.equal(at(1, '同地址账户数'), 1);

    // usage columns
    assert.equal(at(1, '已使用金额'), amount(1) / 100);
    assert.equal(at(1, '剩余余额'), 0);
    assert.equal(at(1, '使用的订单'), '#5001');
    assert.equal(at(2, '已使用金额'), 5);
    assert.equal(at(2, '使用的订单'), '#5002');
    assert.equal(at(14, '已使用金额'), 0);
    assert.equal(at(8, '已使用金额'), null);

    // notes
    assert.equal(at(2, '备注/错误'), 'Shopify 拒绝：boom');
    assert.equal(at(3, '备注/错误'), 'Shopify 拒绝：input: Customer is invalid（INVALID）');
    assert.match(at(6, '备注/错误'), /导出后下过单/);
    assert.match(at(6, '备注/错误'), /最近一笔是 \$0，按更早的付费订单计算/);
    assert.match(at(7, '备注/错误'), /最近一笔已取消，按更早的付费订单计算/);
    assert.match(at(7, '备注/错误'), /^建卡请求发出后没有记录结果（运行被中断）/);
    assert.equal(at(8, '备注/错误'), '订单都已取消或都是 $0，按从没下单处理');
    assert.equal(at(1, '备注/错误'), null);
    assert.equal(at(1, '序号'), seqOf(1));
  });

  test('进行中 is shown only while the matching command is running', async () => {
    writeAllInputs();
    let release = acquireRunLock(paths, 'issue');
    try {
      await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    } finally {
      release();
    }
    let ws = (await openBook(paths.excel)).getWorksheet('发放名单');
    let r7 = findRow(ws, COL['客户 ID'], '7');
    assert.equal(cellValue(ws, r7, COL['状态']), '进行中');
    assert.equal(ws.getRow(r7).getCell(1).fill?.fgColor?.argb, 'FFDDEBF7');
    assert.doesNotMatch(cellValue(ws, r7, COL['备注/错误']), /中断/);
    assert.equal(cellValue(ws, findRow(ws, COL['客户 ID'], '1'), COL['第 2 次提醒']), '结果不明');

    release = acquireRunLock(paths, 'remind');
    try {
      await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    } finally {
      release();
    }
    const wb = await openBook(paths.excel);
    ws = wb.getWorksheet('发放名单');
    r7 = findRow(ws, COL['客户 ID'], '7');
    assert.equal(cellValue(ws, r7, COL['状态']), '需人工核对');
    assert.equal(cellValue(ws, findRow(ws, COL['客户 ID'], '1'), COL['第 2 次提醒']), '进行中');
    // the unfinished remind run shows as running in the run history
    const summary = wb.getWorksheet('汇总');
    const runRows = [];
    for (let r = 1; r <= summary.rowCount; r += 1) if (cellValue(summary, r, 2) === 'remind') runRows.push(cellValue(summary, r, 8));
    assert.deepEqual(runRows, [REMIND_SUMMARY_TEXT, '运行中']);
  });

  test('every sheet is protected (filter and column widths allowed), has a frozen top row and its filter', async () => {
    writeAllInputs();
    await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    const wb = await openBook(paths.excel);
    for (const ws of wb.worksheets) {
      assert.equal(ws.sheetProtection?.sheet, true, `${ws.name} protected`);
      assert.equal(ws.sheetProtection?.autoFilter, true, `${ws.name} filter allowed`);
      assert.equal(ws.sheetProtection?.formatColumns, true, `${ws.name} column widths allowed`);
      assert.equal(ws.sheetProtection?.hashValue, undefined, `${ws.name} has no password`);
      assert.equal(ws.views?.[0]?.state, 'frozen', `${ws.name} frozen`);
      assert.ok(ws.views[0].ySplit >= 1, `${ws.name} header frozen`);
    }
    const list = wb.getWorksheet('发放名单');
    assert.equal(list.views[0].xSplit, 3);
    assert.equal(list.views[0].ySplit, 1);
    assert.equal(list.autoFilter, `A1:AI${selection.recipients.length + 1}`);
    assert.equal(list.getRow(1).height, 30);
    assert.equal(list.getRow(1).getCell(1).font?.bold, true);
    assert.equal(list.getRow(1).getCell(1).fill?.fgColor?.argb, 'FF4472C4');
    assert.equal(wb.getWorksheet('未入选').autoFilter, `A1:K${selection.notSelected.length + 1}`);
    assert.equal(wb.getWorksheet('核对').autoFilter, 'A3:F5');
    assert.ok(wb.getWorksheet('操作日志').autoFilter.startsWith('A1:H'));
  });

  test('使用报告 totals and tables', async () => {
    writeAllInputs();
    await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    const ws = (await openBook(paths.excel)).getWorksheet('使用报告');
    const usage = usageFixture(amount);
    assert.equal(cellValue(ws, 1, 1), '使用报告（截至 2026-10-12 08:00，洛杉矶时间）');

    const issued = findRow(ws, 1, '发出礼品卡');
    assert.deepEqual(rowValues(ws, issued, 3), ['发出礼品卡', 4, usage.summary.issuedCents / 100]);
    assert.equal(ws.getRow(issued).getCell(2).numFmt, '#,##0" 张"');
    const used = findRow(ws, 1, '已经使用');
    assert.deepEqual(rowValues(ws, used, 4), ['已经使用', 2, 0.5, usage.summary.usedCents / 100]);
    assert.ok(Math.abs(cellValue(ws, used, 5) - usage.summary.usedCents / usage.summary.issuedCents) < 1e-9);
    const orders = findRow(ws, 1, '带来订单');
    assert.deepEqual(rowValues(ws, orders, 4), ['带来订单', 2, 85, 42.5]);
    assert.deepEqual(rowValues(ws, orders + 1, 4), [null, null, usage.summary.giftCardCents / 100, usage.summary.customerPaidCents / 100]);

    const tiers = findRow(ws, 1, '按档位');
    assert.deepEqual(rowValues(ws, tiers + 1, 5), ['$10.77', 2, 1, 0.5, 5]);
    const kinds = findRow(ws, 1, '按客户类型');
    assert.deepEqual(rowValues(ws, kinds + 1, 3), ['有下单', 4, 2]);
    const daily = findRow(ws, 1, '每日');
    assert.deepEqual(rowValues(ws, daily + 1, 6), [wall(2026, 10, 6), 1, 1, 60, 1, 0.25]);
    const top = findRow(ws, 1, '卖得最多的 10 个商品');
    assert.deepEqual(rowValues(ws, top + 1, 3), ['Gold Balloon', 3, 45]);
    assert.ok(findRow(ws, 1, '截至 2026-10-12 08:00（洛杉矶时间）。每天运行一次 usage 更新本表；卡的余额小于原金额就算已使用。'));
  });

  test('使用明细: one row per gift-card payment joined with its order, then the unmatched payments', async () => {
    writeAllInputs();
    await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    const ws = (await openBook(paths.excel)).getWorksheet('使用明细');
    assert.deepEqual(rowValues(ws, 1, 9), ['订单号', '下单时间', '客户 ID', '卡号后 4 位', '卡面额', '本单用卡金额', '订单总额', '已取消', '买了什么']);
    assert.equal(ws.autoFilter, 'A1:I4');
    // sorted by order time: #5001 first, then #5002 sale and refund
    assert.deepEqual(rowValues(ws, 2, 9), ['#5001', wall(2026, 10, 6, 11, 0), '1', 'x001', amount(1) / 100, amount(1) / 100, 60, null, 'Gold Balloon × 2; Helium Tank × 1']);
    assert.deepEqual(ws.getRow(2).getCell(1).value, { formula: 'HYPERLINK("https://admin.shopify.com/store/teststore/orders/5001","#5001")', result: '#5001' });
    assert.deepEqual(rowValues(ws, 3, 9), ['#5002', wall(2026, 10, 7, 12, 30), '2', 'x002', amount(2) / 100, 7, 25, '是', 'Gold Balloon × 1']);
    assert.equal(cellValue(ws, 4, 6), -2);
    assert.equal(cellValue(ws, 5, 1), null);

    const title = findRow(ws, 1, '需要人工核对的礼品卡付款：回执里没有本活动礼品卡 ID（1 笔）');
    assert.ok(title, 'unmatched section title');
    assert.equal(title, 7); // two blank rows after the last payment
    assert.deepEqual(rowValues(ws, title + 1, 6), ['订单号', '下单时间', '客户 ID', '付款时间', '金额', '原因']);
    assert.deepEqual(rowValues(ws, title + 2, 6), ['#5003', wall(2026, 10, 8, 13, 0), '9', wall(2026, 10, 8, 13, 0, 5), 10, '回执里没有礼品卡 ID']);
  });

  test('核对, 操作日志 and 汇总 contents', async () => {
    writeAllInputs();
    await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    const wb = await openBook(paths.excel);

    const check = wb.getWorksheet('核对');
    assert.match(cellValue(check, 1, 1), /^核对时间 2026-10-07 10:00（洛杉矶时间）；本活动的卡 4 张；带 gift-card-sent-2026-10 的客户 3 个$/);
    assert.equal(cellValue(check, 2, 1), '发现 2 个问题：有卡但没有 tag 1；仍需人工核对 1');
    assert.deepEqual(rowValues(check, 3, 6), ['问题类型', '客户 ID', '礼品卡 ID', '日志记录', 'Shopify 实际', '建议操作']);
    assert.deepEqual(rowValues(check, 4, 6), ['有卡但没有 tag', '2', '1002', '已建卡，未打 tag', '客户没有 tag', '运行 issue 补打 tag']);
    assert.deepEqual(rowValues(check, 5, 3), ['仍需人工核对', '5', null]);

    const log = wb.getWorksheet('操作日志');
    assert.deepEqual(rowValues(log, 1, 8), ['时间', '命令', '批次/轮次', '客户 ID', '动作', '结果/说明', '礼品卡 ID', '错误']);
    const nonRun = fs.readFileSync(paths.journal, 'utf8').trim().split('\n').map((l) => JSON.parse(l)).filter((e) => !e.op.startsWith('run.'));
    assert.equal(log.actualRowCount, nonRun.length + 1);
    assert.deepEqual(rowValues(log, 3, 8), [wall(2026, 10, 5, 9, 0, 2), 'issue', '1', '1', '建卡成功', `金额 ${'$15.33'}，卡号后 4 位 x001`, '1001', null]);
    const skipRow = findRow(log, 5, '发放前跳过');
    assert.equal(cellValue(log, skipRow, 6), '导出后下过单（#7001）');
    const remindRow = findRow(log, 5, '提醒已发');
    assert.deepEqual(rowValues(log, remindRow, 5), [wall(2026, 10, 12, 9, 5), 'remind', '第 1 次提醒', '1', '提醒已发']);
    const failRow = findRow(log, 5, '提醒失败');
    assert.equal(cellValue(log, failRow, 8), 'Shopify 拒绝：boom');

    const sum = wb.getWorksheet('汇总');
    assert.equal(cellValue(sum, 1, 1), '本文件由程序生成，修改无效，每次运行会覆盖');
    assert.equal(cellValue(sum, 2, 1), '活动 2026-10 · 正式活动 · 生成于 2026-10-12 10:00（洛杉矶时间）');
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '全店客户'), 2), ['全店客户', 15]);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '最终入选'), 2), ['最终入选', selection.recipients.length]);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '10. 同地址重复，保留了另一个账户'), 3), ['10. 同地址重复，保留了另一个账户', 1, '逐行列出']);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '3. 未订阅邮件营销'), 3), ['3. 未订阅邮件营销', 1, '只计数']);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '礼品卡到期日'), 2), ['礼品卡到期日', wall(2026, 10, 19)]);
    assert.equal(cellValue(sum, findRow(sum, 1, '档位（按基数）'), 2), '≤ $10.77 → $10.77；$10.78–$15.33 → $15.33；≥ $15.34 → $19.77');
    assert.equal(cellValue(sum, findRow(sum, 1, '排除的订单渠道'), 2), 'Amazon, Walmart, eBay, Etsy, Sellbrite（205641）, CedCommerce Walmart Connector（1456995）, eBay（1775805）');
    // progress per status, with amounts
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '已完成'), 3), ['已完成', 3, (amount(1) + amount(14) + amount(15)) / 100]);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '需人工核对'), 2), ['需人工核对', 2]);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '待发放'), 2), ['待发放', 1]);
    assert.equal(cellValue(sum, findRow(sum, 1, '下一个待发放的序号'), 2), seqOf(8));
    // reminders per round, with the round's tag
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '轮次'), 9), ['轮次', '提醒日期', '已发', '跳过', '失败', '结果不明', '进行中', '还没处理（已建卡的人）', '本轮 tag']);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '第 1 次提醒'), 9), ['第 1 次提醒', wall(2026, 10, 12), 1, 1, 1, 1, 0, 0, 'gift-card-sent-2026-10-R1']);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '第 2 次提醒'), 9), ['第 2 次提醒', wall(2026, 10, 16), 0, 0, 0, 1, 0, 3, 'gift-card-sent-2026-10-R2']);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '已用过卡'), 3), ['已用过卡', 1, 0]);
    // usage block and run history
    assert.equal(cellValue(sum, findRow(sum, 1, '发出礼品卡（张）'), 2), 4);
    const dry = findRow(sum, 3, '预演');
    assert.deepEqual(rowValues(sum, dry, 8), [wall(2026, 10, 5, 10, 0), 'issue', '预演', '2', 500, 0, wall(2026, 10, 5, 10, 0, 5), ISSUE_DRY_SUMMARY_TEXT]);
    const first = findRow(sum, 8, ISSUE_SUMMARY_TEXT);
    assert.deepEqual(rowValues(sum, first, 6), [wall(2026, 10, 5, 9, 0), 'issue', '实际', '1', 20, 1]);
    assert.ok(findRow(sum, 8, '没有正常结束（可能被中断）'), 'unfinished run while nothing runs');
  });

  test('未入选 and 同地址重复', async () => {
    writeFixtureJournal(paths, amount);
    await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    const wb = await openBook(paths.excel);

    const ns = wb.getWorksheet('未入选');
    assert.deepEqual(rowValues(ns, 1, 11), ['客户 ID', '姓名', '邮箱', '营销状态', '主要原因', '全部原因', '上次下单日期', '订单数', '相关账户', '组号', 'tags']);
    assert.equal(ns.actualRowCount, selection.notSelected.length + 1);
    assert.equal(findRow(ns, 1, '10'), null, 'not-subscribed people are counted only');
    const dup = findRow(ns, 1, '4');
    assert.deepEqual(rowValues(ns, dup, 10), ['4', 'First4 Last4', 'c4@example.org', '已订阅', '10. 同地址重复，保留了另一个账户：保留 3', '10. 同地址重复，保留了另一个账户：保留 3', wall(2026, 2, 1), 1, '3', 'G0001']);
    assert.deepEqual(ns.getRow(dup).getCell(9).value, { formula: 'HYPERLINK("https://admin.shopify.com/store/teststore/customers/3","3")', result: '3' });
    assert.equal(cellValue(ns, findRow(ns, 1, '11'), 9), '12');
    assert.equal(cellValue(ns, findRow(ns, 1, '13'), 11), 'WHS');
    assert.equal(cellValue(ns, findRow(ns, 1, '13'), 9), null);

    const ds = wb.getWorksheet('同地址重复');
    assert.deepEqual(rowValues(ds, 1, 11), ['组号', '组内账户数', '规范化地址', '是否保留', '客户 ID', '姓名', '邮箱', '上次付费下单日期', '订单数', '注册日期', '疑似批量注册']);
    assert.deepEqual(rowValues(ds, 2, 11), ['G0001', 2, '500 elm st · 78701 · US', '保留', '3', 'First3 Last3', 'c3@example.org', wall(2026, 3, 15), 3, wall(2023, 12, 31), null]);
    assert.deepEqual(rowValues(ds, 3, 5), ['G0001', 2, '500 elm st · 78701 · US', '不发', '4']);
    assert.equal(ds.autoFilter, 'A1:K3');
    assert.equal(cellValue(ds, 4, 1), null);
    assert.equal(cellValue(ds, 5, 1), null);
    assert.equal(cellValue(ds, 6, 1), '因同地址账户近 3 个月下过单而不发的人（1 人）');
    assert.deepEqual(rowValues(ds, 7, 5), ['客户 ID', '姓名', '邮箱', '活跃账户', '活跃账户下单时间']);
    assert.deepEqual(rowValues(ds, 8, 5), ['11', 'First11 Last11', 'c11@example.org', '12', wall(2026, 8, 20, 5, 0)]);
    assert.deepEqual(ds.getRow(8).getCell(4).value, { formula: 'HYPERLINK("https://admin.shopify.com/store/teststore/customers/12","12")', result: '12' });
  });

  test('说明 explains sheets, columns, statuses and the tier rule', async () => {
    await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    const ws = (await openBook(paths.excel)).getWorksheet('说明');
    for (const title of [...RECIPIENT_HEADERS, '汇总', '未入选', '待发放', '已建卡未打tag', '结果不明', '档位 1', '档位 3']) {
      assert.ok(findRow(ws, 1, title), `说明 mentions ${title}`);
    }
    assert.equal(cellValue(ws, findRow(ws, 1, '档位 2'), 2), '基数 $10.78–$15.33 → $15.33');
    assert.equal(ws.getRow(findRow(ws, 1, '已完成')).getCell(1).fill?.fgColor?.argb, 'FFE2EFDA');
  });

  test('--out copy is written (directory created) and a directory target gets the file name', async () => {
    const outFile = path.join(env.dir, 'exports', 'nested', 'copy.xlsx');
    const result = await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, out: outFile, lockOptions: FAST_LOCK });
    assert.equal(result.out, outFile);
    const wb = await openBook(outFile);
    assert.equal(wb.worksheets[0].name, '汇总');
    assert.deepEqual(tmpLeftovers(path.dirname(outFile)), []);

    const dirTarget = path.join(env.dir, 'desktop');
    fs.mkdirSync(dirTarget);
    const second = await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, out: dirTarget, lockOptions: FAST_LOCK });
    assert.equal(second.out, path.join(dirTarget, path.basename(paths.excel)));
    assert.ok(fs.existsSync(second.out));
  });

  test('warns when Excel has the file open (~$ owner file)', async () => {
    fs.writeFileSync(path.join(paths.dir, `~$${path.basename(paths.excel)}`), 'owner');
    const log = memoryLog();
    const result = await writeReport({ config: env.config, paths, log, now: FIXED_NOW, lockOptions: FAST_LOCK });
    assert.deepEqual(result.warnings, [OPEN_IN_EXCEL_WARNING]);
    assert.equal(OPEN_IN_EXCEL_WARNING, 'Excel 正打开此文件，请关闭后重新打开才能看到最新内容');
    assert.ok(log.lines.includes(`WARN ${OPEN_IN_EXCEL_WARNING}`));
    assert.ok(fs.existsSync(paths.excel), 'the file is still updated');
  });

  test('zero recipients still produces a valid workbook', async () => {
    await selectionFixture(env.config, {
      customers: [makeCustomer({ n: 1, marketingState: 'UNSUBSCRIBED' }), makeCustomer({ n: 2, lastOrder: makeOrder({ n: 9, createdAt: '2026-09-01T00:00:00Z' }) })],
    });
    await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    const wb = await openBook(paths.excel);
    const ws = wb.getWorksheet('发放名单');
    assert.equal(ws.actualRowCount, 1);
    assert.deepEqual(rowValues(ws, 1, 2), ['序号', '状态']);
    assert.equal(ws.autoFilter, 'A1:AI1');
    const sum = wb.getWorksheet('汇总');
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '最终入选'), 2), ['最终入选', 0]);
    assert.equal(cellValue(sum, findRow(sum, 1, '下一个待发放的序号'), 2), '没有待发放的人');
  });

  test('a test campaign shows 测试 as kind and tier', async () => {
    env.cleanup();
    env = testConfig({ CAMPAIGN_ID: '2026-10-test', TEST_CUSTOMER_IDS: '1,2', SENT_TAG: 'OCT26RTPROMO-TEST' });
    paths = campaignPaths(env.config);
    const testSelection = buildTestSelection({
      config: env.config,
      customers: [makeCustomer({ n: 1 }), makeCustomer({ n: 2 })].map(parseCustomer),
      timezone: TZ,
      createdAt: NOW_ISO,
      snapshot: { exportedAt: NOW_ISO, source: 'nodes', count: 2 },
    });
    writeJsonAtomic(paths.selection, testSelection);
    await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    const wb = await openBook(paths.excel);
    const ws = wb.getWorksheet('发放名单');
    assert.equal(cellValue(ws, 1, 19), '是否有 OCT26RTPROMO-TEST');
    assert.deepEqual([cellValue(ws, 2, COL['客户类型']), cellValue(ws, 2, COL['档位']), cellValue(ws, 2, COL['礼品卡金额']), cellValue(ws, 2, COL['金额算式'])], ['测试', '测试', 0.1, '测试固定金额 $0.10']);
    const sum = wb.getWorksheet('汇总');
    assert.equal(cellValue(sum, 2, 1), '活动 2026-10-test · 测试活动 · 生成于 2026-10-12 10:00（洛杉矶时间）');
    assert.equal(cellValue(sum, findRow(sum, 1, '测试固定金额'), 2), 0.1);
  });

  test('without selection.json it refuses with the Chinese message', async () => {
    fs.rmSync(paths.selection);
    await assert.rejects(writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK }), { message: '还没有名单，请先运行 select' });
    assert.equal(fs.existsSync(paths.excel), false);
    assert.equal(fs.existsSync(paths.excelLock), false);
  });

  test('a failed write removes its temp file, keeps nothing half-written and releases the lock', async () => {
    fs.mkdirSync(paths.excel); // renaming the finished temp file over a directory fails (EISDIR)
    await assert.rejects(writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK }), { code: 'EISDIR' });
    assert.deepEqual(tmpLeftovers(paths.dir), []);
    assert.equal(fs.existsSync(paths.excelLock), false);
    fs.rmdirSync(paths.excel);
    await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    assert.ok(fs.statSync(paths.excel).isFile());
  });

  test('a failure in the middle of writing leaves the previous workbook untouched and no temp file', async () => {
    await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    const before = fs.readFileSync(paths.excel);
    writeFixtureJournal(paths, amount);
    const proto = ExcelJS.stream.xlsx.WorkbookWriter.prototype;
    const original = proto.addWorksheet;
    proto.addWorksheet = function patched(name, options) {
      if (name === '未入选') throw new Error('simulated failure');
      return original.call(this, name, options);
    };
    try {
      await assert.rejects(writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK }), { message: 'simulated failure' });
    } finally {
      proto.addWorksheet = original;
    }
    assert.deepEqual(tmpLeftovers(paths.dir), []);
    assert.ok(fs.readFileSync(paths.excel).equals(before), 'the previous workbook is unchanged');
    assert.equal(fs.existsSync(paths.excelLock), false);
  });

  test('malformed optional shapes do not break the report', async () => {
    const broken = structuredClone(selection);
    broken.recipients = [...broken.recipients, null];
    broken.recipients[0].groupOthers = '123';
    broken.funnel.byRule = { not: 'an array' };
    broken.stats.tiers = null;
    broken.duplicates[0].members = null;
    writeJsonAtomic(paths.selection, broken);
    writeJsonAtomic(paths.usage, { version: 1, fetchedAt: 'not a date', cards: null, payments: [null, { orderId: 1 }], orders: 'x', unmatched: {}, summary: 'x' });
    writeJsonAtomic(paths.verify, { version: 1, verifiedAt: null, issues: { a: 1 } });
    const result = await writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    assert.deepEqual(result.warnings, []);
    const wb = await openBook(paths.excel);
    assert.equal(wb.worksheets.length, 9);
    assert.equal(wb.getWorksheet('发放名单').actualRowCount, selection.recipients.length + 1);
    assert.equal(cellValue(wb.getWorksheet('使用报告'), 1, 1), '使用报告（截至 未知时间，洛杉矶时间）');
  });

  test('an unreadable optional file is skipped with a warning; a foreign tag snapshot is ignored', async () => {
    writeFixtureJournal(paths, amount);
    fs.writeFileSync(paths.verify, '{ not json');
    writeJsonAtomic(paths.tags, { fetchedAt: NOW_ISO, tag: 'some-other-tag', ids: [cust(3)] });
    const log = memoryLog();
    const result = await writeReport({ config: env.config, paths, log, now: FIXED_NOW, lockOptions: FAST_LOCK });
    assert.equal(result.warnings.length, 2);
    assert.match(result.warnings[0], /^verify\.json 无法读取/);
    assert.match(result.warnings[1], /tags\.json 记录的是 tag "some-other-tag"/);
    assert.equal(log.lines.filter((l) => l.startsWith('WARN')).length, 2);
    const wb = await openBook(paths.excel);
    assert.equal(wb.getWorksheet('核对'), undefined);
    const ws = wb.getWorksheet('发放名单');
    assert.equal(cellValue(ws, findRow(ws, COL['客户 ID'], '3'), COL.tag), '否');
  });

  test('two concurrent writeReport calls both resolve and leave a complete file', async () => {
    writeAllInputs();
    const [a, b] = await Promise.all([
      writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK }),
      writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK }),
    ]);
    assert.equal(a.file, paths.excel);
    assert.equal(b.file, paths.excel);
    const wb = await openBook(paths.excel);
    assert.equal(wb.worksheets.length, 9);
    assert.equal(wb.getWorksheet('发放名单').actualRowCount, selection.recipients.length + 1);
    assert.deepEqual(tmpLeftovers(paths.dir), []);
    assert.equal(fs.existsSync(paths.excelLock), false);
  });
});

// ---------------------------------------------------------------------------
// Fixes after the adversarial review (fix spec sections G1–G9, D5)
// ---------------------------------------------------------------------------

describe('excel review fixes', () => {
  let env;
  let paths;
  let selection;
  const amount = (n) => selection.recipients.find((r) => r.customerId === cust(n)).amountCents;

  beforeEach(async () => {
    env = testConfig();
    paths = campaignPaths(env.config);
    selection = await selectionFixture(env.config, fixtureCustomers());
  });
  afterEach(() => env.cleanup());

  const write = (options = {}) => writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK, ...options });
  function writeAllInputs() {
    writeFixtureJournal(paths, amount);
    writeFixtureTags(paths, env.config);
    writeFixtureVerify(paths);
    writeJsonAtomic(paths.usage, usageFixture(amount));
  }
  /** Column `col` of the first row whose column 1 is `label` (null when there is none). */
  const valueOf = (ws, label, col = 2) => {
    const r = findRow(ws, 1, label);
    return r === null ? null : cellValue(ws, r, col);
  };

  // ---- G1 ------------------------------------------------------------------

  test('G1: every worksheet part follows the CT_Worksheet element order; <sheetProtection> comes before <autoFilter>', async () => {
    writeAllInputs();
    await write();
    const parts = await xlsxParts(paths.excel);
    const byName = worksheetPartsByName(parts);
    assert.deepEqual(Object.keys(byName), ['使用报告', '汇总', '发放名单', '未入选', '同地址重复', '使用明细', '核对', '操作日志', '说明']);
    const withFilter = [];
    for (const [name, part] of Object.entries(byName)) {
      const { children, problems } = worksheetOrderProblems(parts[part]);
      assert.deepEqual(problems, [], `${name} (${part}): ${children.join(' > ')}`);
      assert.ok(children.includes('sheetProtection'), `${name} is protected`);
      if (children.includes('autoFilter')) {
        withFilter.push(name);
        assert.ok(children.indexOf('sheetProtection') < children.indexOf('autoFilter'), `${name}: ${children.join(' > ')}`);
      }
    }
    assert.deepEqual(withFilter, ['发放名单', '未入选', '同地址重复', '使用明细', '核对', '操作日志']);
    assert.deepEqual(worksheetChildren(parts[byName['发放名单']]), ['sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetProtection', 'autoFilter', 'pageMargins', 'pageSetup']);
  });

  test('G1: every XML part is well-formed (xmllint --noout)', { skip: HAS_XMLLINT ? false : 'xmllint is not installed' }, async () => {
    writeAllInputs();
    await write();
    const parts = await xlsxParts(paths.excel);
    assert.ok(Object.keys(parts).some((p) => p.startsWith('xl/worksheets/sheet')), 'the workbook has worksheet parts');
    assert.deepEqual(malformedParts(parts), []);
  });

  // ---- G2 ------------------------------------------------------------------

  test('G2: U+FFFE / U+FFFF are removed and lone surrogates become U+FFFD in every value, formula and formula result', async () => {
    // c1 is a recipient (发放名单), c13 is listed in 未入选 (excluded tag), c3's name has a real emoji.
    selection = await selectionFixture(env.config, fixtureCustomers({
      1: { first: 'Ann\uFFFF', last: 'Lee\uD800' },
      2: { first: 'Bo\uDC00b' },
      3: { first: '🎈 Party' },
      13: { first: 'Bob\uFFFE', tags: ['WHS', 'vip\uFFFF'] },
    }));
    const sel = structuredClone(selection);
    sel.recipients.find((r) => r.customerId === cust(1)).basis.orderName = '#1\uFFFE01'; // a HYPERLINK() label
    writeJsonAtomic(paths.selection, sel);
    writeAllInputs();
    appendJournal(paths.journal, { op: 'create.fail', cid: cust(3), error: 'rejected: bad\uFFFFname \uDBFF!', run: 'X' }, { now: () => '2026-10-05T16:09:00.000Z' });
    const usage = usageFixture(amount);
    usage.orders[0].lineItems[0].name = 'Gold\uFFFE Balloon';
    writeJsonAtomic(paths.usage, usage);

    await write();
    const parts = await xlsxParts(paths.excel);
    assert.deepEqual(malformedParts(parts), []);
    for (const [name, xml] of Object.entries(parts)) assert.doesNotMatch(xml, /[\uFFFE\uFFFF]/, name);

    const wb = await openBook(paths.excel);
    const list = wb.getWorksheet('发放名单');
    const row = (n) => findRow(list, COL['客户 ID'], String(n));
    assert.equal(cellValue(list, row(1), COL['姓名']), 'Ann Lee\uFFFD');
    assert.equal(cellValue(list, row(2), COL['姓名']), 'Bo\uFFFDb Last2');
    assert.equal(cellValue(list, row(3), COL['姓名']), '🎈 Party Last3', 'a real surrogate pair is kept');
    assert.deepEqual(list.getRow(row(1)).getCell(COL['上次有效订单号']).value, { formula: 'HYPERLINK("https://admin.shopify.com/store/teststore/orders/101","#101")', result: '#101' });
    assert.match(cellValue(list, row(3), COL['备注/错误']), /rejected: badname \uFFFD!/);
    const ns = wb.getWorksheet('未入选');
    const r13 = findRow(ns, 1, '13');
    assert.equal(cellValue(ns, r13, 2), 'Bob Last13');
    assert.equal(cellValue(ns, r13, 11), 'WHS, vip');
    const detail = wb.getWorksheet('使用明细');
    assert.equal(cellValue(detail, 2, 9), 'Gold Balloon × 2; Helium Tank × 1');
    const log = wb.getWorksheet('操作日志');
    assert.ok(findRow(log, 8, 'rejected: badname \uFFFD!'), 'the journal error is cleaned too');
  });

  test('G2: a long text is cut to Excel\'s cell limit without splitting a surrogate pair', async () => {
    selection = await selectionFixture(env.config, fixtureCustomers({ 1: { first: `x${'🎈'.repeat(20_000)}`, last: '' } }));
    await write();
    const parts = await xlsxParts(paths.excel);
    assert.deepEqual(malformedParts(parts), []);
    const ws = (await openBook(paths.excel)).getWorksheet('发放名单');
    const name = cellValue(ws, findRow(ws, COL['客户 ID'], '1'), COL['姓名']);
    assert.ok(name.length <= 32_767, `${name.length} UTF-16 units`);
    assert.ok(name.endsWith('🎈…'), 'the cut keeps the last whole emoji');
    assert.ok(!name.includes('\uFFFD'), 'no half emoji turned into U+FFFD');
  });

  // ---- G3 ------------------------------------------------------------------

  test('G3: run summaries of every command are Chinese, without zero counters', () => {
    const cases = [
      ['select', false, { recipients: 16117, totalCents: 20011065, source: 'bulk', snapshotReused: false }, '入选 16,117，合计 $200,110.65，导出方式 全店导出'],
      ['select', false, { recipients: 5, totalCents: 7641, source: 'paginated', snapshotReused: true }, '入选 5，合计 $76.41，导出方式 普通分页，复用了 24 小时内导出的客户数据'],
      ['select', false, { error: 'Shopify 返回 500' }, '出错：Shopify 返回 500'],
      ['issue', false, ISSUE_SUMMARY, ISSUE_SUMMARY_TEXT],
      ['issue', true, ISSUE_DRY_SUMMARY, ISSUE_DRY_SUMMARY_TEXT],
      ['issue', false,
        { batch: 4, dryRun: false, attempted: 2, created: 1, tagged: 1, tagFixed: 2, reconciled: 3, skipped: { 'not-subscribed': 1, 'address-ordered-since-snapshot': 2 }, failed: 0, rejected: 1, unknown: 0, amountCents: 1077, tagFailed: 0, reconciledNone: 1, stillUnknown: 0, seqFrom: 30, seqTo: 31 },
        '尝试 2，建卡 1，打 tag 1，补打 tag 2，核对补记 3，跳过（未订阅营销邮件 1，同地址账户导出后下过单 2），被拒（可重试） 1，金额 $10.77，确认未建成 1，序号 30–31'],
      ['issue', false, { batch: 5, dryRun: false, attempted: 0, created: 0, tagged: 0, tagFixed: 0, reconciled: 0, skipped: {}, failed: 0, rejected: 0, unknown: 0, amountCents: 0, tagFailed: 0, reconciledNone: 0, stillUnknown: 0, seqFrom: null, seqTo: null }, '各项都是 0'],
      ['remind', false, REMIND_SUMMARY, REMIND_SUMMARY_TEXT],
      ['remind', true,
        { round: 2, dryRun: true, eligible: 3, planned: 2, attempted: 0, sent: 0, failed: 0, rejected: 0, unknown: 0, retriedUnknown: 0, skipped: { 'not-subscribed': 1, 'card-expired': 1 }, alreadySent: 0, previouslyFailed: 0, waitingUnknown: 0, notIssued: 0, stopped: null },
        '符合条件 3，将发 2，跳过（未订阅营销邮件 1，卡已过期 1）'],
      ['remind', false,
        { round: 1, dryRun: false, eligible: 9, planned: 9, attempted: 6, sent: 1, failed: 5, rejected: 0, unknown: 0, retriedUnknown: 2, skipped: {}, alreadySent: 4, previouslyFailed: 1, waitingUnknown: 3, notIssued: 0, stopped: 'too-many-failures' },
        '符合条件 9，本次要发 9，尝试 6，已发 1，失败 5，补发（之前结果不明） 2，之前已发 4，之前失败 1，结果不明待核对 3，已停止：连续多次没有发送成功'],
      ['remind', false, { round: 1, dryRun: false, eligible: 2, planned: 2, attempted: 1, sent: 1, stopped: 'aborted' }, '符合条件 2，本次要发 2，尝试 1，已发 1，已中断（Ctrl+C）'],
      ['remind', false, { round: 1, dryRun: false, eligible: 2, planned: 2, attempted: 1, rejected: 1, stopped: 'rejected' }, '符合条件 2，本次要发 2，尝试 1，被拒（可重试） 1，已停止：Shopify 没有接受发送请求'],
      ['usage', false, { issuedCards: 5, usedCards: 1, usedCents: 500, orders: 1, ordersTotalCents: 4500, giftCardCents: 500, unmatched: 0 }, '发出的卡 5，已用的卡 1，已用金额 $5.00，订单 1，订单总额 $45.00，礼品卡抵扣 $5.00'],
      ['usage', false, { error: 'GraphQL error: Throttled' }, '出错：Shopify 查询错误：Throttled'],
      ['verify', false, { cardCount: 12, taggedCount: 10, issueCount: 4, fixedCount: 0, counts: { 'duplicate-cards': 1, 'amount-mismatch': 3 }, text: '卡 12 张，发现 4 条，补记日志 0 条' }, '卡 12 张，发现 4 条，补记日志 0 条'],
      ['verify', false, { cardCount: 12, taggedCount: 10, issueCount: 4, fixedCount: 0, counts: { 'duplicate-cards': 1, 'amount-mismatch': 3 } }, '卡 12，带 tag 的客户 10，问题 4，分类（同一客户有多张卡 1，金额不一致 3）'],
      ['verify', false, { error: 'boom', text: '失败：boom' }, '失败：boom'],
    ];
    for (const [command, dryRun, summary, expected] of cases) {
      assert.equal(summaryText(summary, { command, dryRun }), expected, `${command} ${JSON.stringify(summary)}`);
      // the command is inferred when the caller does not pass it
      if (command === 'issue' || command === 'remind') assert.equal(summaryText(summary), expected, `${command} (inferred)`);
      // an error message is shown as it came (it may be Shopify's English); everything else is Chinese
      if (!summary.error) assert.doesNotMatch(expected, /[a-z]+[A-Z]|\b[a-z]+(?:-[a-z]+)+\b|\b(?:source|text|bulk|paginated|aborted|rejected|error|true|false)\b/);
    }
    // older shapes still read
    assert.equal(summaryText({ created: 4, failed: 1, unknown: 1, skipped: 1 }), '建卡 4，失败 1，结果不明 1，跳过 1');
    assert.equal(summaryText(null), '');
  });

  test('G3: the 运行记录 shows the Chinese summaries of the real run.end shapes', async () => {
    writeAllInputs();
    const runs = [
      ['20261003190000-1', 'select', false, { recipients: 9, totalCents: 14000, source: 'bulk', snapshotReused: false }],
      ['20261013170000-1', 'usage', false, { issuedCards: 4, usedCards: 2, usedCents: 1500, orders: 2, ordersTotalCents: 8500, giftCardCents: 1500, unmatched: 1 }],
      ['20261020170000-1', 'verify', false, { cardCount: 4, taggedCount: 3, issueCount: 2, fixedCount: 1, counts: { 'tag-missing': 1, 'still-unknown': 1 }, text: '卡 4 张，发现 2 条，补记日志 1 条' }],
    ];
    for (const [run, command, dryRun, summary] of runs) {
      const t = `${run.slice(0, 4)}-${run.slice(4, 6)}-${run.slice(6, 8)}T17:00:00.000Z`;
      appendJournal(paths.journal, { op: 'run.start', run, command, dryRun, batch: null, limit: null, options: {} }, { now: () => t });
      appendJournal(paths.journal, { op: 'run.end', run, summary, exitCode: 0 }, { now: () => t });
    }
    await write();
    const sum = (await openBook(paths.excel)).getWorksheet('汇总');
    const history = [];
    for (let r = findRow(sum, 1, '开始时间') + 1; r <= sum.rowCount; r += 1) {
      if (cellValue(sum, r, 2)) history.push(`${cellValue(sum, r, 2)}: ${cellValue(sum, r, 8)}`);
    }
    assert.deepEqual(history, [
      'select: 入选 9，合计 $140.00，导出方式 全店导出',
      `issue: ${ISSUE_SUMMARY_TEXT}`,
      `issue: ${ISSUE_DRY_SUMMARY_TEXT}`,
      `remind: ${REMIND_SUMMARY_TEXT}`,
      'usage: 发出的卡 4，已用的卡 2，已用金额 $15.00，订单 2，订单总额 $85.00，礼品卡抵扣 $15.00，需人工核对的付款 1',
      'remind: 没有正常结束（可能被中断）',
      'verify: 卡 4 张，发现 2 条，补记日志 1 条',
    ]);
  });

  // ---- G4 / G8 -------------------------------------------------------------

  test('G8: a failed --out copy is a warning (logged once, returned); the main workbook is written and returned', async () => {
    writeAllInputs();
    const blocker = path.join(env.dir, 'not-a-folder');
    fs.writeFileSync(blocker, 'a file where the --out folder would be');
    fs.writeFileSync(path.join(paths.dir, `~$${path.basename(paths.excel)}`), 'owner'); // Excel has the workbook open
    const log = memoryLog();
    const out = path.join(blocker, 'copy.xlsx');
    const result = await write({ log, out });
    const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    assert.equal(result.file, paths.excel);
    assert.equal(result.out, null);
    assert.equal(result.fileEn, paths.excelEn, 'the English workbook is written too');
    assert.equal(result.outEn, null, 'its copy (next to the Chinese one) fails the same way');
    assert.equal(result.warnings.length, 3, result.warnings.join('\n'));
    assert.match(result.warnings[0], new RegExp(`^另存 Excel 到 ${esc(out)} 失败：`));
    assert.match(result.warnings[1], new RegExp(`^另存英文版 Excel 到 ${esc(path.join(blocker, 'copy-en.xlsx'))} 失败：`));
    assert.equal(result.warnings[2], OPEN_IN_EXCEL_WARNING, 'the "Excel is open" check still runs');
    // G4: writeReport is the one place that logs its warnings, each exactly once
    for (const w of result.warnings) assert.equal(log.lines.filter((l) => l === `WARN ${w}`).length, 1, w);
    assert.equal(log.lines.filter((l) => l.startsWith('WARN')).length, 3);
    const wb = await openBook(paths.excel);
    assert.equal(wb.worksheets.length, 9, 'the main workbook is complete');
    assert.deepEqual(fs.readdirSync(env.dir).filter((f) => f.includes('.tmp-')), []);
  });

  // ---- G5 ------------------------------------------------------------------

  test('G5: temp workbooks of killed writes are removed (old and new names); unrelated files stay', async () => {
    const base = path.basename(paths.excel);
    fs.mkdirSync(paths.dir, { recursive: true });
    const leftovers = [`.${base}.tmp-12345-1790968607386`, `${base}.tmp-17136-1790968607386.xlsx`, `${base}.tmp-1-2.xlsx`];
    const unrelated = [`${base}.bak`, `other.xlsx.tmp-1-2.xlsx`, `.other.xlsx.tmp-1-2`, 'notes.txt', `~$${base}x`, `.${base}.tmp`, `${base}.tmp-1-2.xlsm`];
    for (const f of [...leftovers, ...unrelated]) fs.writeFileSync(path.join(paths.dir, f), 'x');
    await write();
    const names = fs.readdirSync(paths.dir);
    for (const f of leftovers) assert.ok(!names.includes(f), `${f} removed`);
    for (const f of unrelated) assert.ok(names.includes(f), `${f} kept`);
    assert.ok(names.includes(base));
  });

  test('G5: the workbook (and the --out copy) is written to a hidden temp file without the .xlsx suffix', async () => {
    const renames = [];
    const original = fs.renameSync;
    fs.renameSync = function spy(from, to) {
      renames.push([String(from), String(to)]);
      return original.call(this, from, to);
    };
    const outDir = path.join(env.dir, 'desktop');
    fs.mkdirSync(outDir);
    try {
      await write({ out: path.join(outDir, 'copy.xlsx') });
    } finally {
      fs.renameSync = original;
    }
    const main = renames.find(([, to]) => to === paths.excel);
    assert.ok(main, 'the workbook is renamed into place');
    assert.equal(path.dirname(main[0]), paths.dir);
    assert.match(path.basename(main[0]), /^\.gift-card-promo-2026-10\.xlsx\.tmp-\d+-\d+$/);
    const copy = renames.find(([, to]) => to === path.join(outDir, 'copy.xlsx'));
    assert.ok(copy, 'the copy is renamed into place');
    assert.match(path.basename(copy[0]), /^\.copy\.xlsx\.tmp-\d+-\d+$/);
    assert.equal(path.dirname(copy[0]), outDir);
  });

  // ---- G6 ------------------------------------------------------------------

  test('G6: 操作日志 shows skip details in store time, customers by number and marketing states in Chinese', async () => {
    const R = '20261005160000-9';
    const T = (mm) => `2026-10-05T16:${String(mm).padStart(2, '0')}:00.000Z`;
    journalAt(paths, [
      [T(0), { op: 'run.start', run: R, command: 'issue', dryRun: false, batch: 1, limit: 20 }],
      [T(1), { op: 'skip', cid: cust(6), reason: 'ordered-since-snapshot', detail: '2026-10-04T18:30:00Z', batch: 1, run: R }],
      [T(2), { op: 'skip', cid: cust(7), reason: 'address-ordered-since-snapshot', detail: 'gid://shopify/Customer/3', batch: 1, run: R }],
      [T(3), { op: 'skip', cid: cust(8), reason: 'not-subscribed', detail: 'UNSUBSCRIBED', batch: 1, run: R }],
      [T(4), { op: 'skip', cid: cust(5), reason: 'no-email', detail: '邮箱格式无效', batch: 1, run: R }],
      [T(5), { op: 'skip', cid: cust(15), reason: 'ordered-since-snapshot', detail: '2026-12-01T20:15:00.000Z', batch: 1, run: R }], // PST
      [T(6), { op: 'reconcile.found', cid: cust(1), giftCardId: card(1001), last4: 'x001', amountCents: 1533, createdAt: T(5), source: 'preflight', run: R }],
      [T(7), { op: 'tag.ok', cid: cust(1), note: 'already tagged in Shopify', run: R }],
      [T(8), { op: 'remind.skip', cid: cust(2), round: 1, reason: 'not-subscribed', detail: 'NOT_SUBSCRIBED', run: 'R2' }],
      [T(9), { op: 'remind.skip', cid: cust(3), round: 1, reason: 'used', detail: '余额 $5.33 / 面额 $15.33', run: 'R2' }],
      [T(10), { op: 'remind.start', cid: cust(14), round: 1, giftCardId: card(1014), retry: true, run: 'R2' }],
    ]);
    await write();
    const log = (await openBook(paths.excel)).getWorksheet('操作日志');
    const results = Array.from({ length: log.actualRowCount - 1 }, (_, i) => cellValue(log, i + 2, 6));
    assert.deepEqual(results, [
      '导出后下过单（2026-10-04 11:30 店铺时间）',
      '同地址账户导出后下过单（客户 3）',
      '未订阅营销邮件（已退订）',
      '邮箱没了（邮箱格式无效）',
      '导出后下过单（2026-12-01 12:15 店铺时间）',
      '发放前复查，金额 $15.33，卡号后 4 位 x001',
      'Shopify 上已带这个 tag',
      '未订阅营销邮件（未订阅）',
      '已用过卡（余额 $5.33 / 面额 $15.33）',
      '重新发送（之前失败或结果不明）',
    ]);
    for (const text of results) assert.doesNotMatch(text, /gid:\/\/|\d{4}-\d{2}-\d{2}T|\b[A-Z]+_[A-Z]+\b|UNSUBSCRIBED/);
  });

  test('G6: detailText and the skip texts', () => {
    const stamp = (iso) => (Date.parse(iso) === Date.parse('2026-10-04T18:30:00Z') ? '2026-10-04 11:30' : '');
    assert.equal(detailText('2026-10-04T18:30:00Z', { stamp }), '2026-10-04 11:30 店铺时间');
    assert.equal(detailText('2026-10-04T18:30:00Z'), '2026-10-04T18:30:00Z', 'without a clock the time is left alone');
    assert.equal(detailText('2026-10-19', { stamp }), '2026-10-19', 'a calendar date is not an instant');
    assert.equal(detailText('gid://shopify/Customer/42'), '客户 42');
    assert.equal(detailText('gid://shopify/Order/7'), '订单 7');
    assert.equal(detailText('SUBSCRIBED'), '已订阅');
    assert.equal(detailText('NONE'), '没有营销状态');
    assert.equal(detailText('#7001'), '#7001');
    assert.equal(detailText(null), '');
    assert.equal(issueSkipText('relay-email', 'mail.codisto.com'), '邮箱域名在排除名单里（mail.codisto.com）');
    assert.equal(issueSkipText('customer-deleted'), '客户已删除');
    assert.equal(remindSkipText('not-subscribed', 'INVALID'), '未订阅营销邮件（无效）');
    // codes that happen to be Object members are shown as they are, never as an inherited function
    assert.equal(detailText('constructor'), 'constructor');
    assert.equal(remindSkipText('toString'), 'toString');
    assert.equal(summaryText(JSON.parse('{"toString": 2, "skipped": {"constructor": 1}}')), 'toString 2，跳过（constructor 1）');
  });

  // ---- G7 ------------------------------------------------------------------

  async function tagColumn() {
    const ws = (await openBook(paths.excel)).getWorksheet('发放名单');
    const value = (n) => cellValue(ws, findRow(ws, COL['客户 ID'], String(n)), COL.tag);
    return Object.fromEntries([1, 2, 3, 14, 15].map((n) => [n, value(n)]));
  }

  test('G7: a tag snapshot taken after the tag.ok decides the column (a tag removed in the admin shows 否)', async () => {
    writeFixtureJournal(paths, amount); // tag.ok: c1 16:00:03Z, c14 16:06:03Z, c15 16:07:03Z; c2 tag.fail
    writeJsonAtomic(paths.tags, { fetchedAt: '2026-10-06T00:00:00.000Z', tag: env.config.sentTag, ids: [cust(14)] });
    await write();
    assert.deepEqual(await tagColumn(), { 1: '否', 2: '否', 3: '否', 14: '是', 15: '否' });
  });

  test('G7: a tag.ok after the snapshot, or less than 10 minutes before it, still shows 是 (journal); people the snapshot has show 是', async () => {
    writeFixtureJournal(paths, amount);
    // taken 27 s after c14's tag (16:06:03Z), before c15's (16:07:03Z): Shopify's search may not show them yet
    writeJsonAtomic(paths.tags, { fetchedAt: '2026-10-05T16:06:30.000Z', tag: env.config.sentTag, ids: [cust(1), cust(2), cust(3)] });
    await write();
    assert.deepEqual(await tagColumn(), { 1: '是', 2: '是', 3: '是', 14: '是', 15: '是' });
    // more than 10 minutes after c14's tag: the snapshot decides, c14 no longer has the tag
    writeJsonAtomic(paths.tags, { fetchedAt: '2026-10-05T16:16:04.000Z', tag: env.config.sentTag, ids: [cust(1), cust(2), cust(3)] });
    await write();
    assert.deepEqual(await tagColumn(), { 1: '是', 2: '是', 3: '是', 14: '否', 15: '是' });
  });

  test('G7: without a usable snapshot the journal (or the snapshot) decides', async () => {
    writeFixtureJournal(paths, amount);
    await write();
    assert.deepEqual(await tagColumn(), { 1: '是', 2: '否', 3: '否', 14: '是', 15: '是' }, 'no tags.json');
    writeJsonAtomic(paths.tags, { fetchedAt: 'not a time', tag: env.config.sentTag, ids: [cust(2)] });
    await write();
    assert.deepEqual(await tagColumn(), { 1: '是', 2: '是', 3: '否', 14: '是', 15: '是' }, 'tags.json without a readable fetchedAt');
    writeJsonAtomic(paths.tags, { fetchedAt: '2026-10-06T00:00:00.000Z', tag: env.config.sentTag, ids: 'broken' });
    const result = await write();
    assert.deepEqual(result.warnings, ['tags.json 里没有客户列表（ids），已忽略']);
    assert.deepEqual(await tagColumn(), { 1: '是', 2: '否', 3: '否', 14: '是', 15: '是' }, 'tags.json without ids is ignored');
  });

  test('G7: 说明 explains that the latest tag refresh decides the column', async () => {
    await write();
    const ws = (await openBook(paths.excel)).getWorksheet('说明');
    const help = valueOf(ws, `是否有 ${env.config.sentTag}`);
    assert.match(help, /以最近一次从 Shopify 刷新 tag（export --refresh 或 verify）的结果为准/);
    assert.match(help, /在后台删掉的 tag 刷新后会显示“否”/);
    assert.match(help, /刷新之后才打上 tag 的人，按本地日志显示“是”/);
  });

  // ---- D5 / G9 ---------------------------------------------------------------

  test('D5: 汇总 and the 提醒 table show the current .env dates, with a note when they changed after select; the expiry stays frozen', async () => {
    writeFixtureJournal(paths, amount);
    const moved = { ...env.config, remind1Date: '2026-10-13', remind2Date: '2026-10-17', giftCardExpiresOn: '2026-10-20' };
    await writeReport({ config: moved, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    const sum = (await openBook(paths.excel)).getWorksheet('汇总');
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '首封邮件日期（正式建卡）'), 3), ['首封邮件日期（正式建卡）', wall(2026, 10, 5), null]);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '第一次提醒'), 3), ['第一次提醒', wall(2026, 10, 13), '生成名单时为 2026-10-12，现在按 .env 为 2026-10-13']);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '第二次提醒'), 3), ['第二次提醒', wall(2026, 10, 17), '生成名单时为 2026-10-16，现在按 .env 为 2026-10-17']);
    assert.equal(sum.getRow(findRow(sum, 1, '第一次提醒')).getCell(3).font?.color?.argb, 'FFC00000');
    const expiry = rowValues(sum, findRow(sum, 1, '礼品卡到期日'), 3);
    assert.deepEqual(expiry.slice(0, 2), ['礼品卡到期日', wall(2026, 10, 19)], 'cards are created with the frozen expiry');
    assert.equal(expiry[2], '到期日当天仍可使用。注意：.env 的 GIFT_CARD_EXPIRES_ON 现在是 2026-10-20，但建卡仍用生成名单时的 2026-10-19');
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '第 1 次提醒'), 2), ['第 1 次提醒', wall(2026, 10, 13)]);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '第 2 次提醒'), 2), ['第 2 次提醒', wall(2026, 10, 17)]);
    assert.ok(findRow(sum, 1, '第 1 次提醒日期：生成名单时为 2026-10-12，现在按 .env 为 2026-10-13（remind 按现在的日期判断哪天能发）'));
    assert.ok(findRow(sum, 1, '第 2 次提醒日期：生成名单时为 2026-10-16，现在按 .env 为 2026-10-17（remind 按现在的日期判断哪天能发）'));
  });

  test('D5: unchanged dates get no notes', async () => {
    await write();
    const sum = (await openBook(paths.excel)).getWorksheet('汇总');
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '第一次提醒'), 3), ['第一次提醒', wall(2026, 10, 12), null]);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '礼品卡到期日'), 3), ['礼品卡到期日', wall(2026, 10, 19), '到期日当天仍可使用']);
    for (let r = 1; r <= sum.rowCount; r += 1) {
      for (let c = 1; c <= 8; c += 1) assert.doesNotMatch(String(cellValue(sum, r, c) ?? ''), /生成名单时为|GIFT_CARD_EXPIRES_ON 现在是/);
    }
  });

  test('G9: new selection shape (rule 12 customer-deleted, last-test basis) is shown in the funnel, 未入选, 金额 and 备注', async () => {
    const sel = structuredClone(selection);
    if (!sel.funnel.byRule.some((r) => r.code === 'customer-deleted')) sel.funnel.byRule.push({ n: 12, code: 'customer-deleted', label: '补查订单时客户已被删除', count: 0 });
    sel.funnel.byRule.find((r) => r.code === 'customer-deleted').count = 1;
    sel.notSelected.push({
      customerId: cust(77), numericId: '77', name: 'Gone Customer', email: 'c77@example.org', marketingState: 'SUBSCRIBED', primaryRule: 12,
      primaryCode: 'customer-deleted', primaryText: '12. 补查订单时客户已被删除', allReasons: '12. 补查订单时客户已被删除', relatedCustomerId: '', relatedAt: '',
      groupId: '', lastOrderAt: '2026-03-01T12:00:00Z', lastOrderSource: 'web', numberOfOrders: 1, tags: '',
    });
    sel.recipients.find((r) => r.customerId === cust(7)).basisWhy = 'last-test';
    sel.stats.basisFromEarlier = { lastCancelled: 0, lastZero: 1, lastTest: 1 };
    writeJsonAtomic(paths.selection, sel);
    await write();
    const wb = await openBook(paths.excel);
    const sum = wb.getWorksheet('汇总');
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '12. 补查订单时客户已被删除'), 3), ['12. 补查订单时客户已被删除', 1, '逐行列出']);
    assert.ok(findRow(sum, 1, '11. 算出的金额为 $0'), 'rule 11 is still there');
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '按更早的付费订单计算'), 3), ['按更早的付费订单计算', '2 人', '最近一笔已取消 0 人，最近一笔是测试订单 1 人，最近一笔是 $0 1 人']);
    const ns = wb.getWorksheet('未入选');
    assert.deepEqual(rowValues(ns, findRow(ns, 1, '77'), 6), ['77', 'Gone Customer', 'c77@example.org', '已订阅', '12. 补查订单时客户已被删除', '12. 补查订单时客户已被删除']);
    const list = wb.getWorksheet('发放名单');
    assert.match(cellValue(list, findRow(list, COL['客户 ID'], '7'), COL['备注/错误']), /最近一笔是测试订单，按更早的付费订单计算/);
  });

  test('G9: old selection shape (11 rules, no lastTest) still reads', async () => {
    const sel = structuredClone(selection);
    sel.funnel.byRule = sel.funnel.byRule.filter((r) => r.n <= 11);
    sel.stats.basisFromEarlier = { lastCancelled: 1, lastZero: 1 };
    writeJsonAtomic(paths.selection, sel);
    const result = await write();
    assert.deepEqual(result.warnings, []);
    const sum = (await openBook(paths.excel)).getWorksheet('汇总');
    assert.equal(findRow(sum, 1, '12. 补查订单时客户已被删除'), null);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '按更早的付费订单计算'), 3), ['按更早的付费订单计算', '2 人', '最近一笔已取消 1 人，最近一笔是 $0 1 人']);
  });
});

// ---------------------------------------------------------------------------
// Round-2 fixes (round-2 fix spec R2-15 – R2-21). Journals are written by hand
// in the shapes issue / remind write.
// ---------------------------------------------------------------------------

/** The message src/shopify.js gql() throws when fetch behaves like `respond` (no network: fetch is replaced). */
async function gqlErrorMessage(respond) {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = async () => respond();
  initClient({ shop: 'teststore', apiVersion: '2026-07', token: 'token', log: memoryLog(), sleep: async () => {} });
  try {
    await gql('query Probe { shop { id } }');
    return null;
  } catch (err) {
    return err.message;
  } finally {
    globalThis.fetch = originalFetch;
    resetClient();
  }
}

/** The message src/shopify.js throwIfUserErrors throws for `userErrors`. */
function userErrorMessage(operation, userErrors) {
  try {
    throwIfUserErrors(operation, userErrors);
  } catch (err) {
    return err.message;
  }
  return null;
}

/** An issue run.end summary as src/issue.js writes it (emptySummary), with `fields` set. */
const issueEnd = (fields = {}) => ({
  batch: null, dryRun: false, attempted: 0, created: 0, tagged: 0, tagFixed: 0, reconciled: 0, skipped: {}, failed: 0, rejected: 0,
  unknown: 0, amountCents: 0, tagFailed: 0, reconciledNone: 0, stillUnknown: 0, seqFrom: null, seqTo: null, ...fields,
});
/** A remind run.end summary as src/remind.js writes it (newSummary, keys in its order, round-tag counters included), with `fields` set. */
const remindEnd = (fields = {}) => ({
  round: 1, dryRun: false, eligible: 0, planned: 0, attempted: 0, sent: 0, failed: 0, rejected: 0, unknown: 0, retriedUnknown: 0,
  retriedFailed: 0, skipped: {}, alreadySent: 0, alreadySentByTag: 0, previouslyFailed: 0, waitingUnknown: 0, notIssued: 0,
  roundTagged: 0, roundTagFailed: 0, roundTagFixed: 0, roundTagMissing: 0, stopped: null, stoppedByDate: null, ...fields,
});

describe('excel round-2 fixes', () => {
  let env;
  let paths;
  let selection;
  const amount = (n) => selection.recipients.find((r) => r.customerId === cust(n)).amountCents;
  const p2 = (n) => String(n).padStart(2, '0');

  beforeEach(async () => {
    env = testConfig();
    paths = campaignPaths(env.config);
    selection = await selectionFixture(env.config, fixtureCustomers());
  });
  afterEach(() => {
    resetClient();
    env.cleanup();
  });

  const write = (options = {}) => writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK, ...options });
  /** Column `col` of the first row whose column 1 is `label` (null when there is none). */
  const valueOf = (ws, label, col = 2) => {
    const r = findRow(ws, 1, label);
    return r === null ? null : cellValue(ws, r, col);
  };
  /** Every non-empty cell of the workbook as "<sheet>!R<row>C<col>: <text>". */
  async function allCells() {
    const out = [];
    for (const ws of (await openBook(paths.excel)).worksheets) {
      ws.eachRow((row, r) => row.eachCell((cell, c) => {
        const v = plain(cell.value);
        if (v !== null && v !== undefined && v !== '') out.push(`${ws.name}!R${r}C${c}: ${v instanceof Date ? v.toISOString() : v}`);
      }));
    }
    return out;
  }
  /** 运行记录 rows of 汇总: [开始时间, 命令, 预演/实际, 批次/轮次, 数量上限, 退出码, 结束时间, 结果摘要]. */
  function runHistory(sum) {
    const rows = [];
    for (let r = findRow(sum, 1, '开始时间') + 1; r <= sum.rowCount; r += 1) {
      if (cellValue(sum, r, 2)) rows.push(rowValues(sum, r, 8));
    }
    return rows;
  }

  // ---- R2-15 / R2-16 ---------------------------------------------------------

  test('R2-15 / R2-16: 说明 explains exit code 2 after the reminder dates, and that 待发放 / 失败 get no new card from REMIND_1_DATE on', async () => {
    await write();
    const help = (await openBook(paths.excel)).getWorksheet('说明');
    assert.equal(
      valueOf(help, '退出码 2'),
      '用法错误：参数不对、缺少 --limit、还没到活动日期等，或已过 REMIND_1_DATE（issue 只补记和补打 tag，不建新卡）、已过 REMIND_2_DATE（不能再发第 1 次提醒）',
    );
    assert.equal(valueOf(help, '待发放'), '还没有处理；或核对确认上次没有建成卡，下次运行 issue 会重新处理。正式活动从 REMIND_1_DATE 起 issue 不再建新卡。');
    assert.equal(valueOf(help, '失败'), 'Shopify 明确拒绝建卡，原因见“备注/错误”。只有 issue --retry-failed 才会重试。正式活动从 REMIND_1_DATE 起 issue 不再建新卡。');
    for (const status of ['进行中', '需人工核对', '已建卡未打tag', '已完成', '发放前跳过']) assert.doesNotMatch(valueOf(help, status), /REMIND_1_DATE/, status);
    assert.deepEqual([0, 1, 130].map((code) => valueOf(help, `退出码 ${code}`)), ['成功', '中途停止或失败（原因见命令输出）', '按 Ctrl+C 中断（做完当前这个人后退出）']);
  });

  // ---- R2-17 -----------------------------------------------------------------

  test('R2-17: run summaries with newCardsRefused, repairOnly and stoppedByDate read as Chinese', () => {
    const cases = [
      // a live issue on/after REMIND_1_DATE: repairs only, no new card (exit 2)
      ['issue', false, issueEnd({ batch: 3, newCardsRefused: 2 }), '没有建新卡（待建卡） 2'],
      ['issue', false, issueEnd({ batch: 3, tagFixed: 1, reconciled: 1, newCardsRefused: 2 }), '补打 tag 1，核对补记 1，没有建新卡（待建卡） 2'],
      // its dry run (R2-2): no simulated creates
      ['issue', true, issueEnd({ batch: 3, dryRun: true, newCardsRefused: 2 }), '没有建新卡（待建卡） 2'],
      // --repair-only: the mode reads first, wherever the command put the key; false is left out
      ['issue', false, issueEnd({ batch: 4, tagFixed: 2, newCardsRefused: 3, repairOnly: true }), '只补记和补打 tag，补打 tag 2，没有建新卡（待建卡） 3'],
      ['issue', false, { repairOnly: true, ...issueEnd({ batch: 4, reconciled: 1 }) }, '只补记和补打 tag，核对补记 1'],
      ['issue', true, issueEnd({ batch: 4, dryRun: true, repairOnly: true }), '只补记和补打 tag'],
      ['issue', false, issueEnd({ batch: 4, tagFixed: 1, repairOnly: false }), '补打 tag 1'],
      // a live run stopped when the store date reached REMIND_1_DATE (R2-1)
      ['issue', false, issueEnd({ batch: 5, attempted: 3, created: 3, tagged: 3, amountCents: 3231, seqFrom: 4, seqTo: 6, stoppedByDate: '2026-10-12' }), '尝试 3，建卡 3，打 tag 3，金额 $32.31，序号 4–6，按日期停止 2026-10-12'],
      // round 1 stopped when the store date reached REMIND_2_DATE (R2-7), with an expired card skipped
      ['remind', false, remindEnd({ eligible: 5, planned: 5, attempted: 2, sent: 2, skipped: { 'card-expired': 1 }, stoppedByDate: '2026-10-16' }), '符合条件 5，本次要发 5，尝试 2，已发 2，跳过（卡已过期 1），按日期停止 2026-10-16'],
      ['remind', false, remindEnd({ round: 2, eligible: 1, planned: 1, skipped: { 'card-expired': 1 } }), '符合条件 1，本次要发 1，跳过（卡已过期 1）'],
    ];
    for (const [command, dryRun, summary, expected] of cases) {
      assert.equal(summaryText(summary, { command, dryRun }), expected, `${command} ${JSON.stringify(summary)}`);
      assert.equal(summaryText(summary), expected, `${command} (inferred) ${JSON.stringify(summary)}`);
      assert.doesNotMatch(expected, /[a-z]+[A-Z]|\b[a-z]+(?:-[a-z]+)+\b|\b(?:true|false|repairOnly|newCardsRefused|stoppedByDate)\b/);
    }
  });

  test('R2-17: every key the commands write in run.end has a Chinese label (keys as in src/*.js today)', () => {
    const ones = (keys) => Object.fromEntries(keys.map((k) => [k, 1]));
    // select/index.js runSummary; issue.js emptySummary (+ R2 keys) and SKIP_REASONS; remind.js newSummary (+ R2 key, + the
    // round-tag counters of the per-round tag spec) and REMIND_SKIP_REASONS; usage.js runEndSummary; verify.js summary and issue types
    const issueSkips = ['customer-deleted', 'no-email', 'relay-email', 'not-subscribed', 'already-tagged', 'ordered-since-snapshot', 'address-ordered-since-snapshot'];
    const remindSkips = ['no-card', 'multiple-cards', 'card-disabled', 'card-expired', 'used', 'customer-deleted', 'no-email', 'not-subscribed'];
    const verifyTypes = ['missing-in-shopify', 'not-in-journal', 'not-in-selection', 'duplicate-cards', 'amount-mismatch', 'card-disabled', 'tag-missing', 'tag-without-card', 'still-unknown', 'resolved-unknown'];
    const summaries = [
      ['select', { recipients: 3, totalCents: 1077, source: 'paginated', snapshotReused: true }],
      ['issue', {
        batch: 1, dryRun: false,
        ...ones(['attempted', 'created', 'tagged', 'tagFixed', 'reconciled', 'failed', 'rejected', 'unknown', 'amountCents', 'tagFailed', 'reconciledNone', 'stillUnknown', 'newCardsRefused']),
        skipped: ones(issueSkips), seqFrom: 1, seqTo: 2, repairOnly: true, stoppedByDate: '2026-10-12',
      }],
      ['remind', {
        round: 1, dryRun: false,
        ...ones(['eligible', 'planned', 'attempted', 'sent', 'failed', 'rejected', 'unknown', 'retriedUnknown', 'retriedFailed', 'alreadySent', 'previouslyFailed', 'waitingUnknown', 'notIssued']),
        ...ones(['alreadySentByTag', 'roundTagged', 'roundTagFailed', 'roundTagFixed', 'roundTagMissing']),
        skipped: ones(remindSkips), stopped: 'too-many-failures', stoppedByDate: '2026-10-16',
      }],
      ['usage', ones(['issuedCards', 'usedCards', 'usedCents', 'orders', 'ordersTotalCents', 'giftCardCents', 'unmatched'])],
      ['verify', { ...ones(['cardCount', 'taggedCount', 'issueCount', 'fixedCount']), counts: ones(verifyTypes) }],
    ];
    for (const [command, summary] of summaries) {
      const text = summaryText(summary, { command, dryRun: false });
      const raw = [...Object.keys(summary), ...Object.keys(summary.skipped ?? {}), ...Object.keys(summary.counts ?? {})].filter((k) => text.includes(k));
      assert.deepEqual(raw, [], `${command}: ${text}`);
      assert.doesNotMatch(text, /[a-z]+[A-Z]|\b[a-z]+(?:-[a-z]+)+\b|\b(?:true|false|paginated)\b/, `${command}: ${text}`);
    }
  });

  test('R2-17: 运行记录 says why an issue run made no new card, marks repair-only runs and runs stopped by the date', async () => {
    const at = (iso) => () => iso;
    const runs = [
      // [run, started (UTC), ended (UTC), run.start fields, summary, exit code]
      ['20261012170000-1', '2026-10-12T17:00:00.000Z', '2026-10-12T17:00:20.000Z', { command: 'issue', dryRun: true, batch: 3, limit: 5, options: { retryFailed: false } }, issueEnd({ batch: 3, dryRun: true, newCardsRefused: 2 }), 0],
      ['20261012170100-1', '2026-10-12T17:01:00.000Z', '2026-10-12T17:01:30.000Z', { command: 'issue', dryRun: false, batch: 3, limit: 5, options: { retryFailed: false } }, issueEnd({ batch: 3, tagFixed: 1, newCardsRefused: 2 }), 2],
      ['20261012170200-1', '2026-10-12T17:02:00.000Z', '2026-10-12T17:02:30.000Z', { command: 'issue', dryRun: false, batch: 4, limit: null, options: { retryFailed: false, repairOnly: true } }, issueEnd({ batch: 4, reconciled: 1, newCardsRefused: 2, repairOnly: true }), 0],
      // started 10/11 23:50 store time, stopped at midnight
      ['20261012065000-1', '2026-10-12T06:50:00.000Z', '2026-10-12T07:00:01.000Z', { command: 'issue', dryRun: false, batch: 2, limit: 500, options: { retryFailed: false } }, issueEnd({ batch: 2, attempted: 2, created: 2, tagged: 2, amountCents: 2154, seqFrom: 4, seqTo: 5, stoppedByDate: '2026-10-12' }), 2],
      // round 1 started 10/15 23:55 store time, stopped at midnight
      ['20261016065500-1', '2026-10-16T06:55:00.000Z', '2026-10-16T07:00:02.000Z', { command: 'remind', dryRun: false, batch: null, limit: null, options: { round: 1 } }, remindEnd({ eligible: 4, planned: 4, attempted: 1, sent: 1, stoppedByDate: '2026-10-16' }), 2],
    ];
    for (const [run, started, ended, start, summary, exitCode] of runs) {
      appendJournal(paths.journal, { op: 'run.start', run, ...start }, { now: at(started) });
      appendJournal(paths.journal, { op: 'run.end', run, summary, exitCode }, { now: at(ended) });
    }
    await write();
    const sum = (await openBook(paths.excel)).getWorksheet('汇总');
    assert.deepEqual(runHistory(sum), [
      [wall(2026, 10, 11, 23, 50), 'issue', '实际', '2', 500, 2, wall(2026, 10, 12, 0, 0, 1), '尝试 2，建卡 2，打 tag 2，金额 $21.54，序号 4–5，按日期停止 2026-10-12'],
      [wall(2026, 10, 12, 10, 0), 'issue', '预演', '3', 5, 0, wall(2026, 10, 12, 10, 0, 20), '没有建新卡（待建卡） 2'],
      [wall(2026, 10, 12, 10, 1), 'issue', '实际', '3', 5, 2, wall(2026, 10, 12, 10, 1, 30), '补打 tag 1，没有建新卡（待建卡） 2'],
      [wall(2026, 10, 12, 10, 2), 'issue', '实际', '4', null, 0, wall(2026, 10, 12, 10, 2, 30), '只补记和补打 tag，核对补记 1，没有建新卡（待建卡） 2'],
      [wall(2026, 10, 15, 23, 55), 'remind', '实际', '第 1 次提醒', null, 2, wall(2026, 10, 16, 0, 0, 2), '符合条件 4，本次要发 4，尝试 1，已发 1，按日期停止 2026-10-16'],
    ]);
    for (const row of runHistory(sum)) assert.notEqual(row[7], '各项都是 0');
  });

  // ---- R2-18 -----------------------------------------------------------------

  test('R2-18: errorText turns the messages src/shopify.js builds into Chinese and leaves everything else alone', async () => {
    // the real messages, made by src/shopify.js itself (fetch is replaced: nothing leaves this process)
    const rejected = userErrorMessage('giftCardCreate', [{ field: ['input', 'customerId'], message: 'Customer is invalid', code: 'INVALID' }]);
    assert.equal(errorText(rejected), 'Shopify 拒绝：input.customerId: Customer is invalid（INVALID）');
    const twoErrors = userErrorMessage('giftCardCreate', [{ field: ['input'], message: 'Customer is invalid', code: 'INVALID' }, { field: ['input', 'expiresOn'], message: 'must be in the future', code: 'GREATER_THAN' }]);
    assert.equal(errorText(twoErrors), 'Shopify 拒绝：input: Customer is invalid（INVALID）; input.expiresOn: must be in the future（GREATER_THAN）');
    assert.equal(errorText(userErrorMessage('tagsAdd', [{ field: null, message: 'Tag limit reached' }])), 'Shopify 拒绝：Tag limit reached');
    assert.equal(errorText(userErrorMessage('giftCardSendNotificationToCustomer', [{ field: ['id'], message: 'Customer has no email', code: 'INVALID' }])), 'Shopify 拒绝：id: Customer has no email（INVALID）');
    const json = (body, status = 200) => new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
    const gqlCases = [
      [() => { throw new TypeError('fetch failed'); }, '网络错误：fetch failed'],
      [() => new Response('Service Unavailable', { status: 503 }), 'Shopify 返回 HTTP 503：Service Unavailable'],
      [() => new Response('', { status: 502 }), 'Shopify 返回 HTTP 502'],
      [() => new Response('', { status: 429, headers: { 'retry-after': '1' } }), 'Shopify 限流：连续 6 次被拒，已放弃'],
      [() => json({ errors: [{ message: 'Throttled', extensions: { code: 'THROTTLED' } }] }), 'Shopify 限流：连续 6 次被拒，已放弃'],
      [() => json({ errors: [{ message: 'Internal error. Looks like something went wrong on our end.' }] }), 'Shopify 查询错误：Internal error. Looks like something went wrong on our end.'],
      [() => new Response('<html><body>Bad gateway</body></html>', { status: 200 }), 'Shopify 返回了无法解析的内容'],
      [() => json({ data: null }), 'Shopify 没有返回数据'],
    ];
    for (const [respond, expected] of gqlCases) {
      const message = await gqlErrorMessage(respond);
      assert.ok(message, `gql() threw for ${expected}`);
      assert.equal(errorText(message), expected, message);
    }
    // a Shopify message embedded in the commands' own Chinese text (issue: create.unknown when the lookup failed too)
    assert.equal(
      errorText('Network error calling Shopify: fetch failed；查卡也失败了：GraphQL error: Internal error'),
      '网络错误：fetch failed；查卡也失败了：Shopify 查询错误：Internal error',
    );
    // unchanged: the commands' Chinese texts, other errors, Shopify's own words
    for (const same of [
      '上次运行在建卡途中中断，Shopify 上暂时查不到这张卡',
      '上次运行在发送这封提醒时中断，不知道是否已发出',
      '开始建卡超过 10 分钟仍没找到这张卡，确认没有建成',
      'Shopify ignored part of a search filter: created_at',
      'rejected: no operation name',
      'Request rejected by proxy',
      'fetch failed',
    ]) assert.equal(errorText(same), same);
    // translating twice changes nothing; empty stays empty
    for (const m of [rejected, twoErrors, 'Network error calling Shopify: fetch failed；查卡也失败了：GraphQL error: x']) assert.equal(errorText(errorText(m)), errorText(m));
    assert.equal(errorText(null), '');
    assert.equal(errorText(undefined), '');
    assert.equal(errorText(''), '');
  });

  test('R2-18: 发放名单 备注/错误, the reminder cells and 操作日志 错误 show Shopify errors in Chinese; the journal keeps the raw text', async () => {
    const R1 = '20261005160000-7';
    const R2 = '20261012160000-7';
    const T = (mm, ss = 0) => `2026-10-05T16:${p2(mm)}:${p2(ss)}.000Z`;
    const U = (mm, ss = 0) => `2026-10-12T16:${p2(mm)}:${p2(ss)}.000Z`;
    const raw = {
      createFail: 'giftCardCreate rejected: input: Customer is invalid [INVALID]',
      tagFail: 'tagsAdd rejected: tags: Tag limit reached [INVALID]',
      throttled: 'Throttled by Shopify 6 times in a row; giving up',
      lost: 'Network error calling Shopify: fetch failed；查卡也失败了：GraphQL error: Internal error',
      remindFail: 'giftCardSendNotificationToCustomer rejected: input: Customer has no email [INVALID]',
      remindLost: 'Shopify HTTP 503: Service Unavailable',
    };
    const created = (n, mm) => [
      [T(mm), { op: 'create.start', cid: cust(n), amountCents: amount(n), batch: 1, run: R1 }],
      [T(mm, 1), { op: 'create.ok', cid: cust(n), giftCardId: card(3000 + n), last4: `x${n}`, amountCents: amount(n), batch: 1, run: R1 }],
    ];
    journalAt(paths, [
      [T(0), { op: 'run.start', run: R1, command: 'issue', dryRun: false, batch: 1, limit: 20 }],
      [T(1), { op: 'create.start', cid: cust(1), amountCents: amount(1), batch: 1, run: R1 }],
      [T(1, 1), { op: 'create.fail', cid: cust(1), error: raw.createFail, run: R1 }],
      ...created(2, 2),
      [T(2, 2), { op: 'tag.fail', cid: cust(2), error: raw.tagFail, run: R1 }],
      [T(3), { op: 'create.start', cid: cust(3), amountCents: amount(3), batch: 1, run: R1 }],
      [T(3, 1), { op: 'create.rejected', cid: cust(3), error: raw.throttled, run: R1 }],
      ...created(14, 4),
      [T(4, 2), { op: 'tag.ok', cid: cust(14), run: R1 }],
      ...created(15, 5),
      [T(5, 2), { op: 'tag.ok', cid: cust(15), run: R1 }],
      [T(6), { op: 'create.start', cid: cust(5), amountCents: amount(5), batch: 1, run: R1 }],
      [T(6, 9), { op: 'create.unknown', cid: cust(5), error: raw.lost, run: R1 }],
      [T(7), { op: 'run.end', run: R1, summary: issueEnd({ batch: 1, attempted: 6, created: 3, tagged: 2, failed: 1, rejected: 1, unknown: 1, tagFailed: 1 }), exitCode: 1 }],
      [U(0), { op: 'run.start', run: R2, command: 'remind', dryRun: false, options: { round: 1 } }],
      [U(1), { op: 'remind.start', cid: cust(14), round: 1, giftCardId: card(3014), run: R2 }],
      [U(1, 1), { op: 'remind.fail', cid: cust(14), round: 1, error: raw.remindFail, run: R2 }],
      [U(2), { op: 'remind.start', cid: cust(15), round: 1, giftCardId: card(3015), run: R2 }],
      [U(2, 30), { op: 'remind.unknown', cid: cust(15), round: 1, error: raw.remindLost, run: R2 }],
      [U(3), { op: 'remind.start', cid: cust(2), round: 1, giftCardId: card(3002), run: R2 }],
      [U(3, 1), { op: 'remind.rejected', cid: cust(2), round: 1, error: 'HTTP 429 from Shopify 6 times in a row; giving up', run: R2 }],
      [U(4), { op: 'run.end', run: R2, summary: remindEnd({ eligible: 3, planned: 3, attempted: 3, failed: 1, unknown: 1, rejected: 1, stopped: 'rejected' }), exitCode: 1 }],
    ]);
    await write();
    const wb = await openBook(paths.excel);

    const list = wb.getWorksheet('发放名单');
    const at = (n, title) => cellValue(list, findRow(list, COL['客户 ID'], String(n)), COL[title]);
    assert.equal(at(1, '备注/错误'), 'Shopify 拒绝：input: Customer is invalid（INVALID）');
    assert.equal(at(2, '备注/错误'), 'Shopify 拒绝：tags: Tag limit reached（INVALID）');
    assert.equal(at(3, '备注/错误'), 'Shopify 限流：连续 6 次被拒，已放弃');
    assert.equal(at(5, '备注/错误'), '网络错误：fetch failed；查卡也失败了：Shopify 查询错误：Internal error');
    assert.equal(at(14, '第 1 次提醒'), '失败：Shopify 拒绝：input: Customer has no email（INVALID）');
    assert.equal(at(15, '第 1 次提醒'), '结果不明');
    assert.equal(at(2, '第 1 次提醒'), null, 'a remind.rejected send can be tried again');

    const log = wb.getWorksheet('操作日志');
    const errorOf = (n, op) => {
      for (let r = 2; r <= log.rowCount; r += 1) if (cellValue(log, r, 4) === String(n) && cellValue(log, r, 5) === op) return cellValue(log, r, 8);
      return undefined;
    };
    assert.equal(errorOf(1, '建卡被拒'), 'Shopify 拒绝：input: Customer is invalid（INVALID）');
    assert.equal(errorOf(2, '打 tag 失败'), 'Shopify 拒绝：tags: Tag limit reached（INVALID）');
    assert.equal(errorOf(3, '建卡未执行（可重试）'), 'Shopify 限流：连续 6 次被拒，已放弃');
    assert.equal(errorOf(5, '建卡结果不明'), '网络错误：fetch failed；查卡也失败了：Shopify 查询错误：Internal error');
    assert.equal(errorOf(14, '提醒失败'), 'Shopify 拒绝：input: Customer has no email（INVALID）');
    assert.equal(errorOf(15, '提醒结果不明'), 'Shopify 返回 HTTP 503：Service Unavailable');
    assert.equal(errorOf(2, '提醒未发出（可重试）'), 'Shopify 限流：连续 6 次被拒，已放弃');
    assert.equal(errorOf(14, '建卡成功'), null);

    // no cell of the workbook shows the internal English of src/shopify.js (the round-2 reviewer's check, widened)
    const internal = (await allCells()).filter((t) => /\b[a-z][A-Za-z]+ rejected\b|Network error calling Shopify|Throttled by Shopify|HTTP 429 from Shopify|GraphQL error|Shopify HTTP \d|\[INVALID\]/.test(t));
    assert.deepEqual(internal, []);
    // the journal is the record: it keeps Shopify's answer exactly
    const journal = fs.readFileSync(paths.journal, 'utf8');
    for (const text of Object.values(raw)) assert.ok(journal.includes(JSON.stringify(text).slice(1, -1)), text);
  });

  // ---- R2-19 -----------------------------------------------------------------

  test('R2-19: 使用报告 (under the split row), 汇总 and 说明 explain why 礼品卡抵扣 and 已用金额 differ', async () => {
    const NOTE_TEXT = '礼品卡抵扣按结账时计算，之后退回卡里的钱不扣；已用金额 = 原金额 − 现在余额，已扣除退回卡里的钱。所以有退款退回卡时两者不同，差额就是退回卡里的金额。';
    assert.equal(USAGE_SPLIT_NOTE, NOTE_TEXT);
    writeFixtureJournal(paths, amount);
    // checkout view (usage F1): #5002 paid $7.00 with c2's card, $2.00 went back to the card later
    const usage = usageFixture(amount);
    usage.summary.giftCardCents = amount(1) + 700;
    usage.summary.customerPaidCents = usage.summary.ordersTotalCents - usage.summary.giftCardCents;
    writeJsonAtomic(paths.usage, usage);
    await write();
    const wb = await openBook(paths.excel);

    const report = wb.getWorksheet('使用报告');
    const used = findRow(report, 1, '已经使用');
    assert.equal(cellValue(report, used, 4), (amount(1) + 500) / 100, '已用: net of the $2.00 refund back to the card');
    const orders = findRow(report, 1, '带来订单');
    assert.deepEqual(rowValues(report, orders + 1, 4), [null, null, (amount(1) + 700) / 100, usage.summary.customerPaidCents / 100], 'the split row');
    assert.deepEqual(rowValues(report, orders + 2, 3), [NOTE_TEXT, null, null], 'the note right under it');
    assert.equal(report.getRow(orders + 2).getCell(1).font?.color?.argb, 'FF7F7F7F', 'shown as a note');
    assert.equal(cellValue(report, orders + 3, 1), null);
    assert.equal(cellValue(report, orders + 4, 1), '按档位');

    const sum = wb.getWorksheet('汇总');
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '礼品卡抵扣'), 3), ['礼品卡抵扣', (amount(1) + 700) / 100, NOTE_TEXT]);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '已用金额'), 2), ['已用金额', (amount(1) + 500) / 100]);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '顾客另外支付'), 3), ['顾客另外支付', usage.summary.customerPaidCents / 100, null]);

    const help = wb.getWorksheet('说明');
    assert.ok(findRow(help, 1, '使用情况的金额（“使用报告”和“汇总”）'), 'its own section');
    assert.equal(valueOf(help, '礼品卡抵扣和已用金额'), NOTE_TEXT);
  });

  test('R2-19: without usage.json there is no usage block, and 说明 still explains the two amounts', async () => {
    await write();
    const wb = await openBook(paths.excel);
    assert.equal(wb.getWorksheet('使用报告'), undefined);
    assert.equal(findRow(wb.getWorksheet('汇总'), 1, '礼品卡抵扣'), null);
    assert.equal(valueOf(wb.getWorksheet('说明'), '礼品卡抵扣和已用金额'), USAGE_SPLIT_NOTE);
  });

  // ---- R2-20 -----------------------------------------------------------------

  test('R2-20: 最近订单渠道 is explained as the most recent valid order, judged by rule 9, empty without one', async () => {
    const sel = structuredClone(selection);
    // select R2-12: no fallback to a cancelled order's channel
    sel.recipients.find((r) => r.customerId === cust(8)).lastOrderSource = '';
    sel.recipients.find((r) => r.customerId === cust(7)).lastOrderSource = 'web'; // last order cancelled, the earlier valid one is from the web store
    writeJsonAtomic(paths.selection, sel);
    await write();
    const wb = await openBook(paths.excel);
    assert.equal(valueOf(wb.getWorksheet('说明'), '最近订单渠道'), '最近一笔有效订单（未取消、非测试，$0 也算）的来源渠道，第 9 条按它判断；没有有效订单时留空。');
    const list = wb.getWorksheet('发放名单');
    const channel = (n) => cellValue(list, findRow(list, COL['客户 ID'], String(n)), COL['最近订单渠道']);
    assert.deepEqual([1, 5, 7, 8].map(channel), ['网店', null, '网店', null]);
  });

  // ---- R2-21 -----------------------------------------------------------------

  test('R2-21: the not-subscribed skip reads 未订阅营销邮件 everywhere, with the exact marketing state in brackets', async () => {
    const states = { UNSUBSCRIBED: '已退订', NOT_SUBSCRIBED: '未订阅', PENDING: '待确认', INVALID: '无效', REDACTED: '已隐去', NONE: '没有营销状态' };
    for (const [state, label] of Object.entries(states)) {
      assert.equal(issueSkipText('not-subscribed', state), `未订阅营销邮件（${label}）`);
      assert.equal(remindSkipText('not-subscribed', state), `未订阅营销邮件（${label}）`);
    }
    assert.equal(issueSkipText('not-subscribed'), '未订阅营销邮件');
    assert.equal(remindLabel({ status: 'skipped', reason: 'not-subscribed', detail: 'PENDING' }, null, TZ), '跳过：未订阅营销邮件');

    const R1 = '20261005160000-8';
    const R2 = '20261012160000-8';
    journalAt(paths, [
      ['2026-10-05T16:00:00.000Z', { op: 'run.start', run: R1, command: 'issue', dryRun: false, batch: 1, limit: 20 }],
      ['2026-10-05T16:00:01.000Z', { op: 'skip', cid: cust(2), reason: 'not-subscribed', detail: 'UNSUBSCRIBED', batch: 1, run: R1 }],
      ['2026-10-05T16:00:02.000Z', { op: 'create.start', cid: cust(1), amountCents: amount(1), batch: 1, run: R1 }],
      ['2026-10-05T16:00:03.000Z', { op: 'create.ok', cid: cust(1), giftCardId: card(1001), last4: 'x001', amountCents: amount(1), batch: 1, run: R1 }],
      ['2026-10-05T16:00:04.000Z', { op: 'tag.ok', cid: cust(1), run: R1 }],
      ['2026-10-05T16:00:05.000Z', { op: 'run.end', run: R1, summary: issueEnd({ batch: 1, attempted: 1, created: 1, tagged: 1, skipped: { 'not-subscribed': 1 } }), exitCode: 0 }],
      ['2026-10-12T16:00:00.000Z', { op: 'run.start', run: R2, command: 'remind', dryRun: false, options: { round: 1 } }],
      ['2026-10-12T16:00:01.000Z', { op: 'remind.skip', cid: cust(1), round: 1, reason: 'not-subscribed', detail: 'PENDING', run: R2 }],
      ['2026-10-12T16:00:02.000Z', { op: 'run.end', run: R2, summary: remindEnd({ skipped: { 'not-subscribed': 1 } }), exitCode: 0 }],
    ]);
    await write();
    const wb = await openBook(paths.excel);

    const list = wb.getWorksheet('发放名单');
    const row = (n) => findRow(list, COL['客户 ID'], String(n));
    assert.equal(cellValue(list, row(2), COL['状态']), '发放前跳过');
    assert.equal(cellValue(list, row(2), COL['备注/错误']), '未订阅营销邮件');
    assert.equal(cellValue(list, row(1), COL['第 1 次提醒']), '跳过：未订阅营销邮件');

    const log = wb.getWorksheet('操作日志');
    assert.equal(cellValue(log, findRow(log, 5, '发放前跳过'), 6), '未订阅营销邮件（已退订）');
    assert.equal(cellValue(log, findRow(log, 5, '提醒跳过'), 6), '未订阅营销邮件（待确认）');

    const sum = wb.getWorksheet('汇总');
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '发放前跳过的原因') + 1, 2), ['未订阅营销邮件', 1]);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '提醒跳过的原因') + 1, 3), ['未订阅营销邮件', 1, 0]);
    assert.deepEqual(runHistory(sum).map((r) => r[7]), ['尝试 1，建卡 1，打 tag 1，跳过（未订阅营销邮件 1）', '跳过（未订阅营销邮件 1）']);

    const help = wb.getWorksheet('说明');
    assert.match(valueOf(help, '提醒跳过的原因'), /、未订阅营销邮件$/);
    assert.match(valueOf(help, '发放前跳过的原因'), /、未订阅营销邮件、/);
    assert.equal(valueOf(help, '跳过：原因'), '这一轮不符合提醒条件，例如已用过卡、未订阅营销邮件、卡已停用或过期。再次运行同一轮会重新判断。');

    assert.deepEqual((await allCells()).filter((t) => t.includes('已退订营销邮件')), [], 'the old label is gone');
  });
});

// ---------------------------------------------------------------------------
// Per-round reminder tags (round-tag spec of 2026-10-02, fixes duplicates-reminders N1):
// after each reminder remind tags the customer with <SENT_TAG>-R<round>; a customer who
// carries the round's tag counts as sent (journal op remind.found → a sent state with
// source 'tag'), and a failed tagsAdd is journalled as remind.tag.fail (→ tagError on the
// sent state). Journals are written by hand in the shapes the spec gives.
// ---------------------------------------------------------------------------

/** The message src/customers.js addTag throws when Shopify answers with `body` (fetch is replaced: nothing leaves this process). */
async function addTagErrorMessage(body) {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = async () => new Response(JSON.stringify(body), { status: 200, headers: { 'content-type': 'application/json' } });
  initClient({ shop: 'teststore', apiVersion: '2026-07', token: 'token', log: memoryLog(), sleep: async () => {} });
  try {
    await addTag(gid('Customer', 1), 'gift-card-sent-2026-10-R1');
    return null;
  } catch (err) {
    return err.message;
  } finally {
    globalThis.fetch = originalFetch;
    resetClient();
  }
}

describe('excel per-round reminder tags', () => {
  let env;
  let paths;
  let selection;
  const amount = (n) => selection.recipients.find((r) => r.customerId === cust(n)).amountCents;
  const p2 = (n) => String(n).padStart(2, '0');
  const T1 = 'gift-card-sent-2026-10-R1';
  const T2 = 'gift-card-sent-2026-10-R2';
  const SENT_BY_TAG = '已发（按 Shopify 上的本轮 tag 补记）';
  const TAG_MISSING = '；本轮 tag 没打上，下次运行会补打';

  beforeEach(async () => {
    env = testConfig();
    paths = campaignPaths(env.config);
    selection = await selectionFixture(env.config, fixtureCustomers());
  });
  afterEach(() => {
    resetClient();
    env.cleanup();
  });

  const write = (options = {}) => writeReport({ config: env.config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK, ...options });
  /** Column `col` of the first row whose column 1 is `label` (null when there is none). */
  const valueOf = (ws, label, col = 2) => {
    const r = findRow(ws, 1, label);
    return r === null ? null : cellValue(ws, r, col);
  };
  /** Every non-empty cell of the workbook as "<sheet>!R<row>C<col>: <text>". */
  async function allCells() {
    const out = [];
    for (const ws of (await openBook(paths.excel)).worksheets) {
      ws.eachRow((row, r) => row.eachCell((cell, c) => {
        const v = plain(cell.value);
        if (v !== null && v !== undefined && v !== '') out.push(`${ws.name}!R${r}C${c}: ${v instanceof Date ? v.toISOString() : v}`);
      }));
    }
    return out;
  }
  /** 运行记录 rows of 汇总: [开始时间, 命令, 预演/实际, 批次/轮次, 数量上限, 退出码, 结束时间, 结果摘要]. */
  function runHistory(sum) {
    const rows = [];
    for (let r = findRow(sum, 1, '开始时间') + 1; r <= sum.rowCount; r += 1) {
      if (cellValue(sum, r, 2)) rows.push(rowValues(sum, r, 8));
    }
    return rows;
  }
  /** The 操作日志 row (8 values) of customer `n`, action label `action` and reminder `round`, or null. */
  function logRow(log, n, action, round) {
    for (let r = 2; r <= log.rowCount; r += 1) {
      if (cellValue(log, r, 4) === String(n) && cellValue(log, r, 5) === action && cellValue(log, r, 3) === (round ? `第 ${round} 次提醒` : null)) return rowValues(log, r, 8);
    }
    return null;
  }

  /**
   * Round 1 went out on 10/12, but the journal lost those lines (an older backup was put back);
   * only c14's send with an unknown outcome survived. On 10/13 round 1 ran again: c1, c2 and c14
   * carry the round-1 tag, so remind journalled remind.found for them and e-mailed none of them;
   * c15 got his first round-1 email, but tagging him failed. On 10/16 round 2 (--limit 1): c2
   * already carries the round-2 tag; c1 was sent and tagged.
   */
  function writeRoundTagJournal() {
    const R1 = '20261005160000-6';
    const R2 = '20261012160000-6';
    const R3 = '20261013170000-6';
    const R4 = '20261016160000-6';
    const issued = (n, mm) => [
      [`2026-10-05T16:${p2(mm)}:01.000Z`, { op: 'create.start', cid: cust(n), amountCents: amount(n), batch: 1, run: R1 }],
      [`2026-10-05T16:${p2(mm)}:02.000Z`, { op: 'create.ok', cid: cust(n), giftCardId: card(1000 + n), last4: `x${n}`, amountCents: amount(n), batch: 1, run: R1 }],
      [`2026-10-05T16:${p2(mm)}:03.000Z`, { op: 'tag.ok', cid: cust(n), run: R1 }],
    ];
    const options = (round) => ({ round, retryUnknown: false, retryFailed: false });
    journalAt(paths, [
      ['2026-10-05T16:00:00.000Z', { op: 'run.start', run: R1, command: 'issue', dryRun: false, batch: 1, limit: 20 }],
      ...issued(1, 1), ...issued(2, 2), ...issued(14, 3), ...issued(15, 4),
      ['2026-10-05T16:05:00.000Z', { op: 'run.end', run: R1, summary: issueEnd({ batch: 1, attempted: 4, created: 4, tagged: 4, seqFrom: 1, seqTo: 4 }), exitCode: 0 }],
      // 10/12: what is left of the first round-1 run (no run.end)
      ['2026-10-12T16:00:00.000Z', { op: 'run.start', run: R2, command: 'remind', dryRun: false, batch: null, limit: null, options: options(1) }],
      ['2026-10-12T16:01:00.000Z', { op: 'remind.start', cid: cust(14), round: 1, giftCardId: card(1014), run: R2 }],
      ['2026-10-12T16:01:30.000Z', { op: 'remind.unknown', cid: cust(14), round: 1, error: 'Network error calling Shopify: fetch failed', run: R2 }],
      // 10/13: round 1 again
      ['2026-10-13T17:00:00.000Z', { op: 'run.start', run: R3, command: 'remind', dryRun: false, batch: null, limit: null, options: options(1) }],
      ['2026-10-13T17:00:05.000Z', { op: 'remind.found', cid: cust(1), round: 1, source: 'tag', run: R3 }],
      ['2026-10-13T17:00:06.000Z', { op: 'remind.found', cid: cust(2), round: 1, source: 'tag', run: R3 }],
      ['2026-10-13T17:00:07.000Z', { op: 'remind.found', cid: cust(14), round: 1, source: 'tag', run: R3 }],
      ['2026-10-13T17:01:00.000Z', { op: 'remind.start', cid: cust(15), round: 1, giftCardId: card(1015), run: R3 }],
      ['2026-10-13T17:01:01.000Z', { op: 'remind.ok', cid: cust(15), round: 1, run: R3 }],
      ['2026-10-13T17:01:02.000Z', { op: 'remind.tag.fail', cid: cust(15), round: 1, error: 'tagsAdd rejected: tags: Tag limit reached [INVALID]', run: R3 }],
      ['2026-10-13T17:02:00.000Z', { op: 'run.end', run: R3, summary: remindEnd({ eligible: 1, planned: 1, attempted: 1, sent: 1, alreadySentByTag: 3, roundTagFailed: 1 }), exitCode: 0 }],
      // 10/16: round 2
      ['2026-10-16T16:00:00.000Z', { op: 'run.start', run: R4, command: 'remind', dryRun: false, batch: null, limit: 1, options: options(2) }],
      ['2026-10-16T16:00:05.000Z', { op: 'remind.found', cid: cust(2), round: 2, source: 'tag', run: R4 }],
      ['2026-10-16T16:01:00.000Z', { op: 'remind.start', cid: cust(1), round: 2, giftCardId: card(1001), run: R4 }],
      ['2026-10-16T16:01:01.000Z', { op: 'remind.ok', cid: cust(1), round: 2, run: R4 }],
      ['2026-10-16T16:02:00.000Z', { op: 'run.end', run: R4, summary: remindEnd({ round: 2, eligible: 3, planned: 1, attempted: 1, sent: 1, alreadySentByTag: 1, roundTagged: 1 }), exitCode: 0 }],
    ]);
  }

  test('remindLabel: a reminder found by its round tag, and a sent reminder whose round tag is missing (spec texts)', () => {
    // the states foldJournal builds for remind.found (source 'tag') and remind.tag.fail (tagError), by hand
    const found = { status: 'sent', at: null, startedAt: null, giftCardId: null, error: null, reason: null, source: 'tag', run: 'R3' };
    assert.equal(remindLabel(found, null, TZ), SENT_BY_TAG);
    assert.equal(SENT_BY_TAG, '已发（按 Shopify 上的本轮 tag 补记）');
    assert.equal(remindLabel(found, 'remind', TZ), SENT_BY_TAG);
    assert.equal(remindLabel({ ...found, at: '2026-10-13T17:00:05.000Z' }, null, TZ), SENT_BY_TAG, 'no send time: the journal never saw that send');
    const sent = { status: 'sent', startedAt: '2026-10-12T16:04:59.000Z', at: '2026-10-12T16:05:00.000Z', giftCardId: card(1001), error: null, reason: null, run: 'R2' };
    const tagError = 'tagsAdd rejected: tags: Tag limit reached';
    assert.equal(remindLabel({ ...sent, tagError }, null, TZ), '已发 10-12 09:05；本轮 tag 没打上，下次运行会补打');
    assert.equal(remindLabel({ ...found, tagError }, null, TZ), `${SENT_BY_TAG}${TAG_MISSING}`);
    assert.equal(remindLabel({ ...sent, at: null, tagError }, null, TZ), `已发${TAG_MISSING}`);
    assert.equal(remindLabel({ ...sent, tagError: '' }, null, TZ), `已发 10-12 09:05${TAG_MISSING}`, 'a failure without a message is still a failure');
    // no tag problem
    assert.equal(remindLabel(sent, null, TZ), '已发 10-12 09:05');
    assert.equal(remindLabel({ ...sent, tagError: null }, null, TZ), '已发 10-12 09:05');
    assert.equal(remindLabel({ ...sent, source: 'verify' }, null, TZ), '已发 10-12 09:05', 'only source "tag" means found by the round tag');
    // a tagError only qualifies a sent reminder
    assert.equal(remindLabel({ status: 'failed', error: 'boom', tagError }, null, TZ), '失败：boom');
    assert.equal(remindLabel({ status: 'unknown', tagError }, null, TZ), '结果不明');
    assert.equal(remindLabel({ status: 'skipped', reason: 'used', tagError }, null, TZ), '跳过：已用过卡');
    assert.equal(remindLabel({ status: 'in_progress', tagError }, 'remind', TZ), '进行中');
  });

  test('roundTagName names the tags like src/remind.js roundTag: <SENT_TAG>-R<round>, one per round and campaign', () => {
    assert.equal(roundTagName('OCT26RTPROMO', 1), 'OCT26RTPROMO-R1');
    assert.equal(roundTagName('OCT26RTPROMO', 2), 'OCT26RTPROMO-R2');
    assert.equal(roundTagName('OCT26RTPROMO-TEST', 1), 'OCT26RTPROMO-TEST-R1');
    assert.equal(roundTagName('OCT26RTPROMO', '2'), 'OCT26RTPROMO-R2');
    assert.equal(typeof remindModule.roundTag, 'function', 'src/remind.js exports roundTag(sentTag, round) (round-tag spec)');
    for (const sentTag of ['OCT26RTPROMO', 'OCT26RTPROMO-TEST', env.config.sentTag]) {
      for (const round of [1, 2]) assert.equal(roundTagName(sentTag, round), remindModule.roundTag(sentTag, round), `${sentTag} round ${round}`);
    }
  });

  test('run summaries with the round-tag counters read as Chinese (live run, tag repair, dry run)', () => {
    const cases = [
      // the journal was lost after round 1: everyone carries the round tag, nobody is e-mailed again
      [false, remindEnd({ alreadySentByTag: 3 }), '按 tag 认定已发 3'],
      // a normal live run: each send tagged, one tagsAdd failed
      [false, remindEnd({ eligible: 3, planned: 3, attempted: 3, sent: 3, roundTagged: 2, roundTagFailed: 1 }), '符合条件 3，本次要发 3，尝试 3，已发 3，打本轮 tag 2，打本轮 tag 失败 1'],
      // the next live run tags the sent customers who are missing the tag
      [false, remindEnd({ alreadySent: 3, roundTagFixed: 1 }), '之前已发 3，补打本轮 tag 1'],
      [false, remindEnd({ alreadySent: 2, roundTagFixed: 1, roundTagFailed: 1 }), '之前已发 2，打本轮 tag 失败 1，补打本轮 tag 1'],
      [false, remindEnd({ eligible: 1, planned: 1, alreadySent: 2, alreadySentByTag: 1, previouslyFailed: 1 }), '符合条件 1，本次要发 1，之前已发 2，按 tag 认定已发 1，之前失败 1'],
      // a dry run reads the tags but only counts
      [true, remindEnd({ dryRun: true, eligible: 2, planned: 2, alreadySentByTag: 1, roundTagMissing: 1 }), '符合条件 2，将发 2，按 tag 认定已发 1，缺本轮 tag 1'],
      [true, remindEnd({ round: 2, dryRun: true, roundTagMissing: 4 }), '缺本轮 tag 4'],
      // all counters at 0
      [false, remindEnd({ round: 2 }), '各项都是 0'],
    ];
    for (const [dryRun, summary, expected] of cases) {
      assert.equal(summaryText(summary, { command: 'remind', dryRun }), expected, JSON.stringify(summary));
      assert.equal(summaryText(summary), expected, `(inferred) ${JSON.stringify(summary)}`);
      assert.doesNotMatch(expected, /[a-z]+[A-Z]|\b(?:alreadySentByTag|roundTag\w*|true|false)\b/);
    }
  });

  test('the errors a failed round tag journals (src/customers.js addTag) are shown in Chinese', async () => {
    const noPayload = await addTagErrorMessage({ data: { tagsAdd: null } });
    assert.equal(noPayload, 'tagsAdd returned no payload');
    assert.equal(errorText(noPayload), 'Shopify 没有返回结果');
    const refused = await addTagErrorMessage({ data: { tagsAdd: { node: null, userErrors: [{ field: ['tags'], message: 'Tag limit reached' }] } } });
    assert.equal(refused, 'tagsAdd rejected: tags: Tag limit reached');
    assert.equal(errorText(refused), 'Shopify 拒绝：tags: Tag limit reached');
    // the other mutations' "no payload" messages read the same; other words are left alone
    for (const op of ['giftCardCreate', 'giftCardSendNotificationToCustomer']) assert.equal(errorText(`${op} returned no payload`), 'Shopify 没有返回结果');
    assert.equal(errorText('the request returned no payload'), 'the request returned no payload');
    assert.equal(errorText(errorText(noPayload)), 'Shopify 没有返回结果');
  });

  test('说明 and 汇总 name the campaign\'s round tags and say a customer carrying one is never reminded again that round', async () => {
    env.cleanup();
    env = testConfig({ SENT_TAG: 'OCT26RTPROMO' });
    paths = campaignPaths(env.config);
    selection = await selectionFixture(env.config, fixtureCustomers());
    await write();
    const wb = await openBook(paths.excel);

    const help = wb.getWorksheet('说明');
    assert.ok(findRow(help, 1, '本轮 tag（每一轮每人最多一封提醒）'), 'its own section');
    assert.equal(valueOf(help, '本轮 tag'), '第 1 次提醒是 OCT26RTPROMO-R1，第 2 次提醒是 OCT26RTPROMO-R2，两轮互不影响。每发出一封提醒，就给客户打上这一轮的 tag。');
    const carrying = valueOf(help, '带本轮 tag 的人');
    assert.match(carrying, /^每次运行 remind（预演也一样）都先查 Shopify 上带本轮 tag 的客户：他们一律算这一轮已发过，不会再发，加 --retry-unknown 或 --retry-failed 也不会。/);
    assert.match(carrying, /所以即使本地日志丢失或被旧备份覆盖，这一轮也不会给同一个人再发一封；真实运行会把他们在本地日志里补记为“已发（按 Shopify 上的本轮 tag 补记）”。$/);
    const repair = valueOf(help, '补打本轮 tag');
    assert.match(repair, /^打本轮 tag 失败不影响已经发出的提醒（本地日志记着已发），“操作日志”里记一行“打本轮 tag 失败”。下次真实运行同一轮时会先补打；/);
    assert.match(repair, /发出满 10 分钟、Shopify 上却查不到本轮 tag 的人（例如 tag 在后台被删掉）也会补打。补打成功，或发现 Shopify 上其实已带这个 tag，记一行“本轮 tag 已打上”。/);
    assert.match(repair, /发提醒后直接打上 tag 的不单独记一行，只在运行记录里计数。$/);
    // the reminder cells
    assert.equal(valueOf(help, SENT_BY_TAG), '本地日志里这一轮没有“已发”的记录（例如日志丢失，或被更早的备份覆盖），但客户在 Shopify 上带着本轮 tag，所以算作这一轮已发，不会再发。发送时间不详。');
    assert.equal(
      valueOf(help, `已发 MM-DD HH:MM${TAG_MISSING}`),
      '提醒已经发出，但给客户打本轮 tag 失败了：本地日志记着已发，这一轮不会再给他发。下次真实运行同一轮时会先补打这个 tag（见“补打本轮 tag”），补打成功后这里只显示“已发 MM-DD HH:MM”。',
    );
    assert.equal(valueOf(help, '已发 MM-DD HH:MM'), '已让 Shopify 重发礼品卡邮件，时间是店铺时间；发出后给客户打上本轮 tag（见“本轮 tag”）。每一轮每人最多一封。');
    assert.equal(valueOf(help, '第 1 次提醒'), '见“提醒状态”。这一轮的 tag 是 OCT26RTPROMO-R1。');
    assert.equal(valueOf(help, '第 2 次提醒'), '见“提醒状态”。这一轮的 tag 是 OCT26RTPROMO-R2。');
    assert.equal(valueOf(help, '操作日志'), '每次写操作的记录（建卡、打 tag、跳过、提醒），按时间顺序。发提醒后直接打上本轮 tag 的不单独记一行，见“补打本轮 tag”。');
    // every item the help points to exists
    for (const item of ['本轮 tag', '补打本轮 tag', '已发 MM-DD HH:MM']) assert.ok(findRow(help, 1, item), item);

    const sum = wb.getWorksheet('汇总');
    assert.deepEqual([1, 2].map((n) => valueOf(sum, `第 ${n} 次提醒`, 9)), ['OCT26RTPROMO-R1', 'OCT26RTPROMO-R2']);
    assert.equal(sum.getColumn(9).width, 24, 'room for the tag name');
    assert.ok(
      findRow(sum, 1, '提醒只发给已建卡、卡还没用过（余额等于原金额）、没停用没过期、仍订阅营销邮件的人；每一轮每人最多一封：发出后给客户打上本轮 tag，带本轮 tag 的人这一轮不会再发，即使本地日志丢失。跳过的人再次运行同一轮会重新判断。'),
      'the note under the 提醒 table',
    );
  });

  test('a test campaign shows its own round tags (OCT26RTPROMO-TEST-R1 / -R2), never the real campaign\'s', async () => {
    env.cleanup();
    env = testConfig({ CAMPAIGN_ID: '2026-10-test', TEST_CUSTOMER_IDS: '1,2', SENT_TAG: 'OCT26RTPROMO-TEST' });
    paths = campaignPaths(env.config);
    writeJsonAtomic(paths.selection, buildTestSelection({
      config: env.config,
      customers: [makeCustomer({ n: 1 }), makeCustomer({ n: 2 })].map(parseCustomer),
      timezone: TZ,
      createdAt: NOW_ISO,
      snapshot: { exportedAt: NOW_ISO, source: 'nodes', count: 2 },
    }));
    const R = '20261012160000-4';
    journalAt(paths, [
      ['2026-10-12T16:00:00.000Z', { op: 'run.start', run: R, command: 'remind', dryRun: false, batch: null, limit: null, options: { round: 1 } }],
      ['2026-10-12T16:00:01.000Z', { op: 'remind.found', cid: cust(1), round: 1, source: 'tag', run: R }],
      ['2026-10-12T16:00:02.000Z', { op: 'remind.tag.fail', cid: cust(2), round: 2, error: 'Network error calling Shopify: fetch failed', run: R }],
      ['2026-10-12T16:00:03.000Z', { op: 'run.end', run: R, summary: remindEnd({ alreadySentByTag: 1 }), exitCode: 0 }],
    ]);
    await write();
    const wb = await openBook(paths.excel);
    const sum = wb.getWorksheet('汇总');
    assert.deepEqual([1, 2].map((n) => valueOf(sum, `第 ${n} 次提醒`, 9)), ['OCT26RTPROMO-TEST-R1', 'OCT26RTPROMO-TEST-R2']);
    const notes = Array.from({ length: sum.rowCount }, (_, i) => cellValue(sum, i + 1, 1)).filter((v) => typeof v === 'string' && v.startsWith('提醒只发给'));
    assert.deepEqual(notes, ['提醒只发给已建卡、卡还没用过（余额等于原金额）、没停用没过期的人（测试活动不看营销订阅状态）；每一轮每人最多一封：发出后给客户打上本轮 tag，带本轮 tag 的人这一轮不会再发，即使本地日志丢失。跳过的人再次运行同一轮会重新判断。']);
    const help = wb.getWorksheet('说明');
    assert.match(valueOf(help, '本轮 tag'), /^第 1 次提醒是 OCT26RTPROMO-TEST-R1，第 2 次提醒是 OCT26RTPROMO-TEST-R2，/);
    assert.equal(valueOf(help, '第 2 次提醒'), '见“提醒状态”。这一轮的 tag 是 OCT26RTPROMO-TEST-R2。');
    const log = wb.getWorksheet('操作日志');
    assert.deepEqual(logRow(log, 1, '按本轮 tag 补记已发', 1).slice(5, 6), ['Shopify 上已带 OCT26RTPROMO-TEST-R1']);
    assert.deepEqual(logRow(log, 2, '打本轮 tag 失败', 2).slice(5), ['OCT26RTPROMO-TEST-R2 没打上，下次运行会补打', null, '网络错误：fetch failed']);
    assert.deepEqual((await allCells()).filter((t) => /OCT26RTPROMO-R[12]/.test(t)), [], 'no cell names the real campaign\'s round tags');
  });

  test('操作日志 shows remind.found and remind.tag.fail (round, tag name, Chinese error); 运行记录 shows the round-tag counters', async () => {
    writeRoundTagJournal();
    // an entry without its run.start and without a round (still shown, as a reminder entry)
    appendJournal(paths.journal, { op: 'remind.tag.fail', cid: cust(15), error: 'tagsAdd returned no payload', run: 'gone' }, { now: () => '2026-10-16T16:03:00.000Z' });
    await write();
    const wb = await openBook(paths.excel);

    const log = wb.getWorksheet('操作日志');
    assert.deepEqual(logRow(log, 1, '按本轮 tag 补记已发', 1), [wall(2026, 10, 13, 10, 0, 5), 'remind', '第 1 次提醒', '1', '按本轮 tag 补记已发', `Shopify 上已带 ${T1}`, null, null]);
    assert.deepEqual(logRow(log, 14, '按本轮 tag 补记已发', 1), [wall(2026, 10, 13, 10, 0, 7), 'remind', '第 1 次提醒', '14', '按本轮 tag 补记已发', `Shopify 上已带 ${T1}`, null, null]);
    assert.deepEqual(logRow(log, 15, '打本轮 tag 失败', 1), [wall(2026, 10, 13, 10, 1, 2), 'remind', '第 1 次提醒', '15', '打本轮 tag 失败', `${T1} 没打上，下次运行会补打`, null, 'Shopify 拒绝：tags: Tag limit reached（INVALID）']);
    assert.deepEqual(logRow(log, 2, '按本轮 tag 补记已发', 2), [wall(2026, 10, 16, 9, 0, 5), 'remind', '第 2 次提醒', '2', '按本轮 tag 补记已发', `Shopify 上已带 ${T2}`, null, null]);
    assert.deepEqual(logRow(log, 15, '打本轮 tag 失败', null), [wall(2026, 10, 16, 9, 3), 'remind', null, '15', '打本轮 tag 失败', '本轮 tag 没打上，下次运行会补打', null, 'Shopify 没有返回结果']);
    const actions = Array.from({ length: log.actualRowCount - 1 }, (_, i) => cellValue(log, i + 2, 5));
    assert.deepEqual(actions.filter((a) => /本轮 tag/.test(a)), ['按本轮 tag 补记已发', '按本轮 tag 补记已发', '按本轮 tag 补记已发', '打本轮 tag 失败', '按本轮 tag 补记已发', '打本轮 tag 失败']);
    assert.ok(actions.every((a) => !/^remind\./.test(a)), `every op has a Chinese label: ${actions.join(', ')}`);

    const sum = wb.getWorksheet('汇总');
    assert.deepEqual(runHistory(sum).map((r) => [r[1], r[3], r[7]]), [
      ['issue', '1', '尝试 4，建卡 4，打 tag 4，序号 1–4'],
      ['remind', '第 1 次提醒', '没有正常结束（可能被中断）'],
      ['remind', '第 1 次提醒', '符合条件 1，本次要发 1，尝试 1，已发 1，按 tag 认定已发 3，打本轮 tag 失败 1'],
      ['remind', '第 2 次提醒', '符合条件 3，本次要发 1，尝试 1，已发 1，按 tag 认定已发 1，打本轮 tag 1'],
    ]);
    assert.deepEqual([1, 2].map((n) => valueOf(sum, `第 ${n} 次提醒`, 9)), [T1, T2]);
  });

  test('发放名单 and the 提醒 table count a reminder found by its round tag as sent, and flag a sent one whose round tag is missing', async () => {
    // needs foldJournal's handling of remind.found (→ sent, source 'tag') and remind.tag.fail (→ tagError)
    writeRoundTagJournal();
    await write();
    const wb = await openBook(paths.excel);

    const list = wb.getWorksheet('发放名单');
    const row = (n) => findRow(list, COL['客户 ID'], String(n));
    const reminders = (n) => [cellValue(list, row(n), COL['第 1 次提醒']), cellValue(list, row(n), COL['第 2 次提醒'])];
    assert.deepEqual(reminders(1), [SENT_BY_TAG, '已发 10-16 09:01'], 'round 1 from the tag; round 2 sent');
    assert.deepEqual(reminders(2), [SENT_BY_TAG, SENT_BY_TAG], 'both rounds from the tags');
    assert.deepEqual(reminders(14), [SENT_BY_TAG, null], 'the unknown send is settled by the tag, not sent again');
    assert.deepEqual(reminders(15), [`已发 10-13 10:01${TAG_MISSING}`, null], 'sent, round tag missing');
    assert.deepEqual(reminders(3), [null, null]);
    // red marks a failed or unknown send only: these are all sent
    for (const n of [1, 2, 14, 15]) assert.notEqual(list.getRow(row(n)).getCell(COL['第 1 次提醒']).font?.color?.argb, 'FFC00000', `customer ${n}`);

    const sum = wb.getWorksheet('汇总');
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '第 1 次提醒'), 9), ['第 1 次提醒', wall(2026, 10, 12), 4, 0, 0, 0, 0, 0, T1]);
    assert.deepEqual(rowValues(sum, findRow(sum, 1, '第 2 次提醒'), 9), ['第 2 次提醒', wall(2026, 10, 16), 2, 0, 0, 0, 0, 2, T2]);
  });

  test('the next live run\'s tag repair (remind.tag.ok) clears the flag; 操作日志 tells a repaired tag from one found in Shopify', async () => {
    const R1 = '20261005160000-3';
    const R2 = '20261012160000-3';
    const R3 = '20261013170000-3';
    journalAt(paths, [
      ['2026-10-05T16:00:00.000Z', { op: 'run.start', run: R1, command: 'issue', dryRun: false, batch: 1, limit: 20 }],
      ...[1, 15].flatMap((n, i) => [
        [`2026-10-05T16:0${i + 1}:01.000Z`, { op: 'create.start', cid: cust(n), amountCents: amount(n), batch: 1, run: R1 }],
        [`2026-10-05T16:0${i + 1}:02.000Z`, { op: 'create.ok', cid: cust(n), giftCardId: card(1000 + n), last4: `x${n}`, amountCents: amount(n), batch: 1, run: R1 }],
        [`2026-10-05T16:0${i + 1}:03.000Z`, { op: 'tag.ok', cid: cust(n), run: R1 }],
      ]),
      ['2026-10-05T16:05:00.000Z', { op: 'run.end', run: R1, summary: issueEnd({ batch: 1, attempted: 2, created: 2, tagged: 2 }), exitCode: 0 }],
      // 10/12: both sent; c1's tagsAdd answer was lost (Shopify had applied it), c15's was refused
      ['2026-10-12T16:00:00.000Z', { op: 'run.start', run: R2, command: 'remind', dryRun: false, batch: null, limit: null, options: { round: 1 } }],
      ['2026-10-12T16:01:00.000Z', { op: 'remind.start', cid: cust(1), round: 1, giftCardId: card(1001), run: R2 }],
      ['2026-10-12T16:01:01.000Z', { op: 'remind.ok', cid: cust(1), round: 1, run: R2 }],
      ['2026-10-12T16:01:02.000Z', { op: 'remind.tag.fail', cid: cust(1), round: 1, error: 'Network error calling Shopify: fetch failed', run: R2 }],
      ['2026-10-12T16:02:00.000Z', { op: 'remind.start', cid: cust(15), round: 1, giftCardId: card(1015), run: R2 }],
      ['2026-10-12T16:02:01.000Z', { op: 'remind.ok', cid: cust(15), round: 1, run: R2 }],
      ['2026-10-12T16:02:02.000Z', { op: 'remind.tag.fail', cid: cust(15), round: 1, error: 'tagsAdd rejected: tags: Tag limit reached [INVALID]', run: R2 }],
      ['2026-10-12T16:03:00.000Z', { op: 'run.end', run: R2, summary: remindEnd({ eligible: 2, planned: 2, attempted: 2, sent: 2, roundTagFailed: 2 }), exitCode: 0 }],
    ]);
    await write();
    const cell = (ws, n) => cellValue(ws, findRow(ws, COL['客户 ID'], String(n)), COL['第 1 次提醒']);
    let list = (await openBook(paths.excel)).getWorksheet('发放名单');
    assert.deepEqual([1, 15].map((n) => cell(list, n)), [`已发 10-12 09:01${TAG_MISSING}`, `已发 10-12 09:02${TAG_MISSING}`]);

    // 10/13: the next live run of round 1 finds c1 tagged after all, and tags c15 again
    journalAt(paths, [
      ['2026-10-13T17:00:00.000Z', { op: 'run.start', run: R3, command: 'remind', dryRun: false, batch: null, limit: null, options: { round: 1 } }],
      ['2026-10-13T17:00:01.000Z', { op: 'remind.tag.ok', cid: cust(1), round: 1, note: 'already tagged in Shopify', run: R3 }],
      ['2026-10-13T17:00:02.000Z', { op: 'remind.tag.ok', cid: cust(15), round: 1, run: R3 }],
      ['2026-10-13T17:01:00.000Z', { op: 'run.end', run: R3, summary: remindEnd({ alreadySent: 2, roundTagFixed: 1 }), exitCode: 0 }],
    ]);
    await write();
    const wb = await openBook(paths.excel);
    list = wb.getWorksheet('发放名单');
    assert.deepEqual([1, 15].map((n) => cell(list, n)), ['已发 10-12 09:01', '已发 10-12 09:02'], 'no longer flagged');

    const log = wb.getWorksheet('操作日志');
    assert.deepEqual(logRow(log, 1, '本轮 tag 已打上', 1), [wall(2026, 10, 13, 10, 0, 1), 'remind', '第 1 次提醒', '1', '本轮 tag 已打上', `Shopify 上已带 ${T1}`, null, null]);
    assert.deepEqual(logRow(log, 15, '本轮 tag 已打上', 1), [wall(2026, 10, 13, 10, 0, 2), 'remind', '第 1 次提醒', '15', '本轮 tag 已打上', `已补打 ${T1}`, null, null]);
    assert.deepEqual(logRow(log, 1, '打本轮 tag 失败', 1).slice(5), [`${T1} 没打上，下次运行会补打`, null, '网络错误：fetch failed']);
    const sum = wb.getWorksheet('汇总');
    assert.deepEqual(runHistory(sum).slice(1).map((r) => r[7]), ['符合条件 2，本次要发 2，尝试 2，已发 2，打本轮 tag 失败 2', '之前已发 2，补打本轮 tag 1']);
  });

  test('what src/remind.js really journals reads right: tag failures, the repair, then a journal put back to before round 1', async () => {
    // Real live remind runs against the in-process fake Shopify (fetch replaced: nothing leaves
    // this process); each run regenerates the workbook itself, as it does in production.
    env.cleanup();
    env = testConfig();
    paths = campaignPaths(env.config);
    const customers = [1, 2, 3].map((n) => makeCustomer({ n, lastOrder: makeOrder({ n: 5000 + n, createdAt: `2026-06-${p2(30 - n)}T12:00:00Z`, total: '80.00' }) }));
    selection = await selectionFixture(env.config, { customers });
    assert.deepEqual(selection.recipients.map((r) => r.customerId), customers.map((c) => c.id), 'seq n is customer n');
    const issuedAt = '2026-10-05T17:00:00.000Z';
    const giftCards = selection.recipients.map((r) => {
      const amountText = (r.amountCents / 100).toFixed(2);
      return {
        id: card(2000 + r.seq), createdAt: issuedAt, note: campaignNote(env.config.giftCardNote, env.config.campaignId), templateSuffix: 'gift-card-promo', enabled: true,
        lastCharacters: `x00${r.seq}`, initialValue: { amount: amountText, currencyCode: 'USD' }, balance: { amount: amountText, currencyCode: 'USD' }, customer: { id: r.customerId }, expiresOn: '2026-10-19',
      };
    });
    const R = '20261005170000-2';
    journalAt(paths, [
      [issuedAt, { op: 'run.start', run: R, command: 'issue', dryRun: false, batch: 1, limit: 3, options: {} }],
      ...selection.recipients.flatMap((r, i) => [
        [issuedAt, { op: 'create.start', cid: r.customerId, amountCents: r.amountCents, batch: 1, run: R }],
        [issuedAt, { op: 'create.ok', cid: r.customerId, giftCardId: giftCards[i].id, last4: giftCards[i].lastCharacters, amountCents: r.amountCents, batch: 1, run: R }],
        [issuedAt, { op: 'tag.ok', cid: r.customerId, run: R }],
      ]),
      [issuedAt, { op: 'run.end', run: R, summary: issueEnd({ batch: 1, attempted: 3, created: 3, tagged: 3 }), exitCode: 0 }],
    ]);
    const beforeRound1 = fs.readFileSync(paths.journal, 'utf8');
    // round 1's tagsAdd calls, in send order: c1 tagged, c2 refused, c3 applied but its answer lost
    const fake = installFakeShopify({ customers, giftCards, failures: { TagsAdd: [null, { kind: 'userError', message: 'Tag limit reached' }, { kind: 'appliedThenLost' }] }, now: () => Date.parse(issuedAt) });
    const run = async (iso) => {
      const log = memoryLog();
      const result = await remindModule.runRemind({ config: env.config, round: 1, log, now: () => new Date(iso), sleep: async () => {} });
      assert.equal(result.exitCode, 0, log.lines.join('\n'));
      const wb = await openBook(paths.excel);
      const list = wb.getWorksheet('发放名单');
      const cells = [1, 2, 3].map((n) => cellValue(list, findRow(list, COL['客户 ID'], String(n)), COL['第 1 次提醒']));
      return { wb, cells };
    };
    try {
      let { wb, cells } = await run('2026-10-12T17:00:00Z');
      assert.deepEqual(cells, ['已发 10-12 10:00', `已发 10-12 10:00${TAG_MISSING}`, `已发 10-12 10:00${TAG_MISSING}`]);
      let log = wb.getWorksheet('操作日志');
      assert.deepEqual(logRow(log, 2, '打本轮 tag 失败', 1).slice(5), [`${T1} 没打上，下次运行会补打`, null, 'Shopify 拒绝：input: Tag limit reached（INVALID）']);
      assert.deepEqual(logRow(log, 3, '打本轮 tag 失败', 1).slice(5), [`${T1} 没打上，下次运行会补打`, null, '网络错误：fetch failed: socket hang up']);

      // 10/13, round 1 again: the repair (c2 tagged again, c3 found tagged after all), nobody e-mailed
      ({ wb, cells } = await run('2026-10-13T17:00:00Z'));
      assert.deepEqual(cells, ['已发 10-12 10:00', '已发 10-12 10:00', '已发 10-12 10:00']);
      log = wb.getWorksheet('操作日志');
      assert.equal(logRow(log, 2, '本轮 tag 已打上', 1)[5], `已补打 ${T1}`);
      assert.equal(logRow(log, 3, '本轮 tag 已打上', 1)[5], `Shopify 上已带 ${T1}`);

      // 10/14: the journal is replaced by a copy from before round 1, and round 1 is run again
      fs.writeFileSync(paths.journal, beforeRound1);
      ({ wb, cells } = await run('2026-10-14T17:00:00Z'));
      assert.deepEqual(cells, [SENT_BY_TAG, SENT_BY_TAG, SENT_BY_TAG]);
      assert.equal(fake.state.calls.notify.length, 3, 'one round-1 email per person in all');
      const sum = wb.getWorksheet('汇总');
      assert.deepEqual(rowValues(sum, findRow(sum, 1, '第 1 次提醒'), 9), ['第 1 次提醒', wall(2026, 10, 12), 3, 0, 0, 0, 0, 0, T1]);
      assert.deepEqual(runHistory(sum).filter((r) => r[1] === 'remind').map((r) => r[7]), ['按 tag 认定已发 3']);
    } finally {
      fake.restore();
    }
  });
});

describe('excel helpers', () => {
  test('statusLabel', () => {
    assert.equal(statusLabel('pending'), '待发放');
    assert.equal(statusLabel(undefined), '待发放');
    assert.equal(statusLabel('done'), '已完成');
    assert.equal(statusLabel('created'), '已建卡未打tag');
    assert.equal(statusLabel('failed'), '失败');
    assert.equal(statusLabel('unknown'), '需人工核对');
    assert.equal(statusLabel('skipped'), '发放前跳过');
    assert.equal(statusLabel('in_progress', 'issue'), '进行中');
    assert.equal(statusLabel('in_progress', 'remind'), '需人工核对');
    assert.equal(statusLabel('in_progress', null), '需人工核对');
    assert.equal(effectiveStatus('in_progress', 'verify'), 'unknown');
  });

  test('remindLabel', () => {
    assert.equal(remindLabel(undefined, null, TZ), '');
    assert.equal(remindLabel({ status: 'sent', at: '2026-10-12T16:05:00.000Z' }, null, TZ), '已发 10-12 09:05');
    assert.equal(remindLabel({ status: 'sent', at: '2026-12-01T17:00:00.000Z' }, null, TZ), '已发 12-01 09:00'); // PST
    const skipped = {
      'no-card': '没有卡',
      'multiple-cards': '有多张卡',
      'card-disabled': '卡已停用',
      'card-expired': '卡已过期',
      used: '已用过卡',
      'customer-deleted': '客户已删除',
      'no-email': '没有邮箱',
      'not-subscribed': '未订阅营销邮件',
    };
    for (const [reason, label] of Object.entries(skipped)) assert.equal(remindLabel({ status: 'skipped', reason }, null, TZ), `跳过：${label}`);
    assert.equal(remindLabel({ status: 'failed', error: 'boom' }, null, TZ), '失败：boom');
    assert.equal(remindLabel({ status: 'unknown' }, 'remind', TZ), '结果不明');
    assert.equal(remindLabel({ status: 'in_progress' }, 'remind', TZ), '进行中');
    assert.equal(remindLabel({ status: 'in_progress' }, 'issue', TZ), '结果不明');
    assert.equal(remindLabel({ status: 'in_progress' }, null, TZ), '结果不明');
  });

  test('store-local times are exact across daylight-saving changes', () => {
    // 2026-11-01: PDT ends at 09:00Z; both instants are 01:30 local
    assert.equal(remindLabel({ status: 'sent', at: '2026-11-01T08:30:00Z' }, null, TZ), '已发 11-01 01:30');
    assert.equal(remindLabel({ status: 'sent', at: '2026-11-01T09:30:00Z' }, null, TZ), '已发 11-01 01:30');
    // 2026-03-08: PDT starts at 10:00Z
    assert.equal(remindLabel({ status: 'sent', at: '2026-03-08T09:30:00Z' }, null, TZ), '已发 03-08 01:30');
    assert.equal(remindLabel({ status: 'sent', at: '2026-03-08T10:30:00Z' }, null, TZ), '已发 03-08 03:30');
    assert.equal(remindLabel({ status: 'sent', at: '2026-10-12T16:05:00Z' }, null, 'Asia/Shanghai'), '已发 10-13 00:05');
  });

  test('adminUrl and numericId', () => {
    const base = 'https://admin.shopify.com/store/teststore';
    assert.equal(adminUrl('teststore', 'customers', 'gid://shopify/Customer/123'), `${base}/customers/123`);
    assert.equal(adminUrl('teststore', 'customer', 'gid://shopify/Customer/123'), `${base}/customers/123`);
    assert.equal(adminUrl('teststore', 'orders', 'gid://shopify/Order/456'), `${base}/orders/456`);
    assert.equal(adminUrl('teststore', 'gift_cards', 'gid://shopify/GiftCard/789'), `${base}/gift_cards/789`);
    assert.equal(adminUrl('teststore', 'gift-card', '789'), `${base}/gift_cards/789`);
    assert.equal(adminUrl('teststore', null, 'gid://shopify/GiftCard/789'), `${base}/gift_cards/789`);
    assert.equal(adminUrl('teststore', 'customers', null), '');
    assert.equal(adminUrl('teststore', 'customers', 'gid://shopify/Customer/'), '');
    assert.equal(numericId('gid://shopify/Order/12?x=1'), '12');
  });
});

describe('excel performance', () => {
  let env;
  afterEach(() => env?.cleanup());

  test('20,000 recipients + 50,000 not selected + a full journal are written in well under a minute', { timeout: 180_000 }, async () => {
    env = testConfig();
    const { config } = env;
    const paths = campaignPaths(config);
    const N = 20_000;
    const M = 50_000;
    const baseMs = Date.parse('2026-06-30T19:00:00.000Z');
    const recipients = Array.from({ length: N }, (_, i) => {
      const kind = i % 4 === 3 ? 'never' : 'ordered';
      return {
        seq: i + 1,
        customerId: gid('Customer', 100_000 + i),
        numericId: String(100_000 + i),
        name: `Customer ${i}`,
        firstName: 'Customer',
        lastName: String(i),
        email: `c${i}@example.org`,
        marketingState: 'SUBSCRIBED',
        kind,
        basis: kind === 'ordered' ? { orderId: gid('Order', 500_000 + i), orderName: `#${500_000 + i}`, createdAt: new Date(baseMs - i * 3_600_000).toISOString(), totalCents: 5_000 + (i % 300) * 100, sourceName: 'web' } : null,
        basisWhy: null,
        neverReason: kind === 'never' ? 'no-orders' : null,
        lastOrderSource: kind === 'ordered' ? 'web' : '',
        amountCents: 1077,
        rawCents: 500,
        tier: 0,
        formula: '10% × $50.00 = $5.00 → 档位 $10.77',
        addressKey: `${i} main st|78701|us`,
        groupId: null,
        groupSize: 1,
        groupOthers: [],
        city: 'Austin',
        provinceCode: 'TX',
        zip: '78701',
        numberOfOrders: 3,
        amountSpentCents: 15_000,
        accountCreatedAt: new Date(Date.parse('2022-01-01T00:00:00Z') + i * 60_000).toISOString(),
      };
    });
    const notSelected = Array.from({ length: M }, (_, i) => ({
      customerId: gid('Customer', 900_000 + i),
      numericId: String(900_000 + i),
      name: `Excluded ${i}`,
      email: `x${i}@example.org`,
      marketingState: 'SUBSCRIBED',
      primaryRule: 7,
      primaryCode: i % 50 === 0 ? 'active-address' : 'recent-order',
      primaryText: '7. 近 3 个月有下单：2026-08-01',
      allReasons: '5. 带排除 tag：WHS；7. 近 3 个月有下单：2026-08-01',
      relatedCustomerId: i % 50 === 0 ? gid('Customer', 100_000) : '',
      relatedAt: i % 50 === 0 ? '2026-08-01T12:00:00Z' : '',
      groupId: '',
      lastOrderAt: '2026-08-01T12:00:00Z',
      lastOrderSource: 'web',
      numberOfOrders: 2,
      tags: 'WHS, VIP',
    }));
    const selection = {
      version: 1,
      campaignId: config.campaignId,
      mode: 'live',
      createdAt: NOW_ISO,
      snapshot: { exportedAt: NOW_ISO, source: 'bulk', count: N + M },
      params: selectionParams(config, { timezone: TZ, cutoffMs: Date.parse('2026-07-01T07:00:00Z'), cutoffIso: '2026-07-01T00:00:00-07:00' }),
      averageCents: 8000,
      neverAmountCents: 1077,
      recipients,
      notSelected,
      duplicates: [],
      funnel: { total: N + M, byRule: [], tags: {}, channels: {}, relayDomains: {}, notSubscribed: {}, listed: M, unlisted: 0, recipients: N },
      stats: { recipients: N, orderedCount: N * 0.75, neverCount: N * 0.25, totalCents: N * 1077, tiers: [] },
    };
    writeJsonAtomic(paths.selection, selection);
    // journal: every recipient issued (start, ok, tag) and half of them reminded
    const lines = [JSON.stringify({ t: '2026-10-05T16:00:00.000Z', op: 'run.start', run: 'R1', command: 'issue', dryRun: false, batch: 1, limit: N })];
    for (const r of recipients) {
      const t = '2026-10-05T16:30:00.000Z';
      lines.push(JSON.stringify({ t, op: 'create.start', cid: r.customerId, amountCents: 1077, batch: 1, run: 'R1' }));
      lines.push(JSON.stringify({ t, op: 'create.ok', cid: r.customerId, giftCardId: gid('GiftCard', 2_000_000 + r.seq), last4: 'x123', amountCents: 1077, batch: 1, run: 'R1' }));
      lines.push(JSON.stringify({ t, op: 'tag.ok', cid: r.customerId, run: 'R1' }));
      if (r.seq % 2) {
        lines.push(JSON.stringify({ t: '2026-10-12T16:00:00.000Z', op: 'remind.start', cid: r.customerId, round: 1, giftCardId: gid('GiftCard', 2_000_000 + r.seq), run: 'R2' }));
        lines.push(JSON.stringify({ t: '2026-10-12T16:00:01.000Z', op: 'remind.ok', cid: r.customerId, round: 1, run: 'R2' }));
      }
    }
    lines.push(JSON.stringify({ t: '2026-10-05T19:00:00.000Z', op: 'run.end', run: 'R1', summary: { created: N }, exitCode: 0 }));
    fs.writeFileSync(paths.journal, `${lines.join('\n')}\n`);

    // Both editions (Chinese, then English) within the budget.
    const started = Date.now();
    const result = await writeReport({ config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK });
    const elapsed = Date.now() - started;
    assert.ok(elapsed < 60_000, `took ${elapsed} ms`);
    assert.deepEqual(result.warnings, []);
    assert.equal(result.fileEn, paths.excelEn);

    // Stream the results back to prove they are complete and well-formed.
    for (const file of [paths.excel, paths.excelEn]) {
      const counts = {};
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(file, { sharedStrings: 'ignore', hyperlinks: 'ignore', styles: 'ignore', worksheets: 'emit' });
      for await (const sheet of reader) {
        let n = 0;
        for await (const row of sheet) if (row.hasValues) n += 1;
        counts[sheet.name ?? sheet.id] = n;
      }
      const values = Object.values(counts);
      const name = path.basename(file);
      assert.ok(values.includes(N + 1), `${name} 发放名单 rows: ${JSON.stringify(counts)}`);
      assert.ok(values.includes(M + 1), `${name} 未入选 rows: ${JSON.stringify(counts)}`);
      assert.ok(values.includes(lines.length - 2 + 1), `${name} 操作日志 rows: ${JSON.stringify(counts)}`);
    }
  });
});

test('run history: a failed run reads in Chinese (errorText), for summary.error and for verify\'s summary.text', () => {
  assert.equal(summaryText({ error: 'Network error calling Shopify: fetch failed' }, { command: 'usage', dryRun: false }), '出错：网络错误：fetch failed');
  assert.equal(summaryText({ error: 'tagsAdd rejected: id: Customer not found' }, { command: 'select', dryRun: false }), '出错：Shopify 拒绝：id: Customer not found');
  assert.equal(summaryText({ error: 'x', text: '失败：Network error calling Shopify: fetch failed' }, { command: 'verify', dryRun: false }), '失败：网络错误：fetch failed');
  assert.equal(summaryText({ text: '卡 5 张，没有发现问题' }, { command: 'verify', dryRun: false }), '卡 5 张，没有发现问题');
});
