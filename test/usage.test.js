import { test, beforeEach, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

import {
  runUsage,
  buildUsage,
  campaignStart,
  cardSearchSince,
  giftCardIdFromReceipt,
  giftCardOrdersQuery,
  NO_RECEIPT_ID_REASON,
} from '../src/usage.js';
import { campaignPaths, appendJournal, readJournal, acquireRunLock, runningCommand, writeJsonAtomic } from '../src/campaign.js';
import { campaignNote } from '../src/giftcards.js';
import { resetClient } from '../src/shopify.js';
import { OPEN_IN_EXCEL_WARNING } from '../src/report/excel.js';
import { installFakeShopify } from './fake-shopify.js';
import { NOW_ISO, gid, makeCustomer, makeOrder, makeGiftCardOrder, testConfig, memoryLog, selectionFixture } from './helpers.js';

// ---------------------------------------------------------------------------
// Fixture: 7 recipients, 7 campaign cards, a handful of gift-card orders
// ---------------------------------------------------------------------------

const T = {
  firstIssue: '2026-10-05T16:00:00.000Z', // 09:00 on 10/05 in Los Angeles
  now: '2026-10-09T05:00:00.000Z', // 22:00 on 10/08 in Los Angeles (already 10/09 in UTC)
};
const CAMPAIGN_NOTE = campaignNote('gift-card-promo', '2026-10');
const sleep = async () => {};

// Last paid order totals → 10% base → tier: $50 → $5.00 → $10.77 (tier 0), $120 → $12.00 → $15.33 (1),
// $200 → $20.00 → $19.77 (2), $40 and $30 → tier 0. Average $88.00 → $8.80 → the never-ordered (6, 7) get $10.77.
const ORDERED_TOTALS = { 1: '50.00', 2: '120.00', 3: '200.00', 4: '40.00', 5: '30.00' };

function fixtureCustomers() {
  return [
    ...Object.entries(ORDERED_TOTALS).map(([n, total]) => makeCustomer({ n: Number(n), lastOrder: makeOrder({ n: 100 + Number(n), total }) })),
    makeCustomer({ n: 6 }),
    makeCustomer({ n: 7 }),
  ];
}

/** A FIND_CAMPAIGN_CARDS-shaped gift card node. */
function card(n, customerN, cents, { createdAt = '2026-10-05T16:30:00Z', note = CAMPAIGN_NOTE, enabled = true } = {}) {
  const amount = (cents / 100).toFixed(2);
  return {
    id: gid('GiftCard', n),
    createdAt,
    note,
    templateSuffix: 'gift-card-promo',
    enabled,
    expiresOn: '2026-10-19',
    lastCharacters: `x${n}`,
    initialValue: { amount, currencyCode: 'USD' },
    balance: { amount, currencyCode: 'USD' },
    customer: customerN ? { id: gid('Customer', customerN) } : null,
  };
}

const CAMPAIGN_CARDS = () => [
  card(501, 1, 1077), card(502, 2, 1533), card(503, 3, 1977), card(504, 4, 1077),
  card(505, 5, 1077), card(506, 6, 1077), card(507, 7, 1077),
];

/**
 * #2001 card 501 (10/06 local, 10/07 UTC) + a FAILURE attempt   #2002 card 502, then a $5 refund back to it
 * #2003 cards 503 + 504 (+ an AUTHORIZATION, ignored)           #2004 a non-campaign card only
 * #2005 recipient 6, receipt without card id → unmatched        #2006 cancelled, card 507 paid then refunded
 * #2007 non-recipient, receipt without card id → ignored
 */
function scenarioOrders() {
  return [
    makeGiftCardOrder({
      n: 2001, customer: gid('Customer', 1), createdAt: '2026-10-07T03:30:00Z', total: '45.00',
      payments: [{ giftCardNumericId: 501, amount: '10.77' }, { giftCardNumericId: 501, amount: '10.77', status: 'FAILURE' }],
      lineItems: [{ name: 'Latex Balloon 12in', sku: 'LB-12', quantity: 10, amount: '8.00' }, { name: 'Helium Tank', sku: 'HT-1', quantity: 1, amount: '30.00' }],
    }),
    makeGiftCardOrder({
      n: 2002, customer: gid('Customer', 2), createdAt: '2026-10-07T07:30:00Z', total: '15.33',
      payments: [{ giftCardNumericId: 502, amount: '15.33' }, { giftCardNumericId: 502, amount: '5.00', kind: 'REFUND', processedAt: '2026-10-08T18:00:00Z' }],
      lineItems: [{ name: 'Foil Star', sku: 'FS-1', quantity: 2, amount: '15.33' }],
    }),
    makeGiftCardOrder({
      n: 2003, customer: gid('Customer', 3), createdAt: '2026-10-07T20:00:00Z', total: '60.00',
      payments: [{ giftCardNumericId: 503, amount: '19.77', kind: 'AUTHORIZATION' }, { giftCardNumericId: 503, amount: '19.77' }, { giftCardNumericId: 504, amount: '10.77' }],
      lineItems: [{ name: 'Latex Balloon 12in', sku: 'LB-12', quantity: 10, amount: '8.00' }, { name: 'Balloon Arch Kit', sku: 'BAK', quantity: 1, amount: '52.00' }],
    }),
    makeGiftCardOrder({
      n: 2004, customer: gid('Customer', 5), createdAt: '2026-10-06T19:00:00Z', total: '50.00',
      payments: [{ giftCardNumericId: 900, amount: '50.00' }],
      lineItems: [{ name: 'Party Hat', quantity: 50, amount: '50.00' }],
    }),
    makeGiftCardOrder({
      n: 2005, customer: gid('Customer', 6), createdAt: '2026-10-08T01:00:00Z', total: '20.00',
      payments: [{ amount: '10.77', noReceiptId: true }],
      lineItems: [{ name: 'Ribbon', quantity: 1, amount: '20.00' }],
    }),
    makeGiftCardOrder({
      n: 2006, customer: gid('Customer', 7), createdAt: '2026-10-06T18:00:00Z', total: '30.00', cancelled: true,
      payments: [{ giftCardNumericId: 507, amount: '10.77' }, { giftCardNumericId: 507, amount: '10.77', kind: 'REFUND', processedAt: '2026-10-06T19:00:00Z' }],
      lineItems: [{ name: 'Confetti', quantity: 100, amount: '30.00' }],
    }),
    makeGiftCardOrder({
      n: 2007, customer: gid('Customer', 99), createdAt: '2026-10-08T02:00:00Z', total: '25.00',
      payments: [{ amount: '25.00', noReceiptId: true }],
      lineItems: [{ name: 'Streamer', quantity: 3, amount: '25.00' }],
    }),
  ];
}

/** Balances that match the scenario's payments (506 was spent by #2005, whose receipt has no card id). */
function spendScenario() {
  fake.spendCard(gid('GiftCard', 501), 1077);
  fake.spendCard(gid('GiftCard', 502), 1033);
  fake.spendCard(gid('GiftCard', 503), 1977);
  fake.spendCard(gid('GiftCard', 504), 1077);
  fake.spendCard(gid('GiftCard', 506), 1077);
  fake.spendCard(gid('GiftCard', 900), 5000);
}

function installScenario() {
  fake = installFakeShopify({
    giftCards: [
      ...CAMPAIGN_CARDS(),
      card(900, 5, 5000, { note: 'bought in store' }),
      card(901, 6, 1000, { note: campaignNote('gift-card-promo', '2026-09') }),
      card(902, 1, 1077, { createdAt: '2026-10-01T18:00:00Z' }), // marked, but before the search window: never listed
    ],
    giftCardOrders: scenarioOrders(),
  });
  spendScenario();
}

/** Journal of the 2026-10 issue runs: a dry run on 10/04 (ignored), then live runs on 10/05. */
function writeIssueRuns(paths) {
  const at = (t) => ({ now: () => t });
  appendJournal(paths.journal, { op: 'run.start', run: 'r0', command: 'issue', dryRun: true, batch: null, limit: 20 }, at('2026-10-04T16:00:00.000Z'));
  appendJournal(paths.journal, { op: 'run.end', run: 'r0', summary: {}, exitCode: 0 }, at('2026-10-04T16:01:00.000Z'));
  appendJournal(paths.journal, { op: 'run.start', run: 'r1', command: 'issue', dryRun: false, batch: 1, limit: 20 }, at(T.firstIssue));
  appendJournal(paths.journal, { op: 'run.end', run: 'r1', summary: {}, exitCode: 0 }, at('2026-10-05T16:05:00.000Z'));
  appendJournal(paths.journal, { op: 'run.start', run: 'r2', command: 'issue', dryRun: false, batch: 2, limit: 500 }, at('2026-10-05T18:00:00.000Z'));
  appendJournal(paths.journal, { op: 'run.end', run: 'r2', summary: {}, exitCode: 0 }, at('2026-10-05T18:10:00.000Z'));
}

// ---------------------------------------------------------------------------
// Per-test state
// ---------------------------------------------------------------------------

let env;
let paths;
let fake;
let log;
let reports;

beforeEach(() => {
  env = testConfig();
  paths = campaignPaths(env.config);
  log = memoryLog();
  reports = [];
  fake = null;
});

afterEach(() => {
  fake?.restore();
  fake = null;
  resetClient();
  env.cleanup();
});

/** writeReport stub: records the call and what the world looked like at that moment. */
async function writeReport(opts) {
  reports.push({ opts, lockHolder: runningCommand(opts.paths), usageWritten: fs.existsSync(opts.paths.usage) });
  return { file: opts.paths.excel, fileEn: opts.paths.excelEn, out: null, outEn: null, warnings: [] };
}

const go = (overrides = {}) => runUsage({ config: env.config, log, now: () => new Date(T.now), sleep, writeReport, ...overrides });
const readUsageFile = (file) => JSON.parse(fs.readFileSync(file, 'utf8'));
const tail = (id) => String(id).split('/').pop();
const output = () => log.lines.join('\n');

// ---------------------------------------------------------------------------
// Pure pieces
// ---------------------------------------------------------------------------

test('giftCardIdFromReceipt: JSON string, object, double-encoded, damaged and missing receipts', () => {
  assert.equal(giftCardIdFromReceipt('{"gift_card_id":123,"gift_card_last_characters":"x000"}'), 'gid://shopify/GiftCard/123');
  assert.equal(giftCardIdFromReceipt({ gift_card_id: 123 }), 'gid://shopify/GiftCard/123');
  assert.equal(giftCardIdFromReceipt({ gift_card_id: '456' }), 'gid://shopify/GiftCard/456');
  assert.equal(giftCardIdFromReceipt({ gift_card_id: 'gid://shopify/GiftCard/42' }), 'gid://shopify/GiftCard/42');
  assert.equal(giftCardIdFromReceipt('{"gift_card_id":"789"}'), 'gid://shopify/GiftCard/789');
  assert.equal(giftCardIdFromReceipt(JSON.stringify(JSON.stringify({ gift_card_id: 321 }))), 'gid://shopify/GiftCard/321');
  // Beyond 2^53: the digits are kept exactly (JSON.parse would round them).
  assert.equal(giftCardIdFromReceipt('{"gift_card_id":9007199254740993}'), 'gid://shopify/GiftCard/9007199254740993');
  // A damaged receipt still yields an id that is plainly there; otherwise nothing.
  assert.equal(giftCardIdFromReceipt('{"gift_card_id": 55, "x": '), 'gid://shopify/GiftCard/55');
  assert.equal(giftCardIdFromReceipt('{not json'), null);
  assert.equal(giftCardIdFromReceipt('{}'), null);
  assert.equal(giftCardIdFromReceipt(''), null);
  assert.equal(giftCardIdFromReceipt(null), null);
  assert.equal(giftCardIdFromReceipt(undefined), null);
  assert.equal(giftCardIdFromReceipt({ gift_card_id: null }), null);
  assert.equal(giftCardIdFromReceipt({ gift_card_id: 0 }), null);
  assert.equal(giftCardIdFromReceipt('{"gift_card_id":12.5}'), null);
  assert.equal(giftCardIdFromReceipt('{"original_gift_card_id":5}'), null);
  assert.equal(giftCardIdFromReceipt({ original_gift_card_id: 5 }), null);
});

test('campaignStart: the orders search always starts 1 hour before the list; the daily table at the first live issue run or the oldest card', () => {
  const selection = { createdAt: NOW_ISO }; // 2026-10-01T20:00:00.000Z
  const SEARCH = '2026-10-01T19:00:00Z'; // whole seconds, as Shopify's search syntax documents
  const runs = [
    { command: 'select', dryRun: false, startedAt: '2026-10-01T19:30:00.000Z' },
    { command: 'issue', dryRun: true, startedAt: '2026-10-04T16:00:00.000Z' },
    { command: 'issue', dryRun: false, startedAt: '2026-10-05T18:00:00.000Z' },
    { command: 'issue', dryRun: false, startedAt: '2026-10-05T16:00:00.500Z' },
    { command: 'remind', dryRun: false, startedAt: '2026-10-03T16:00:00.000Z' },
  ];
  const cards = [{ createdAt: '2026-10-05T16:30:00Z' }, { createdAt: '2026-10-05T16:00:01Z' }, { createdAt: 'not a date' }];
  const noIssue = runs.filter((r) => r.dryRun || r.command !== 'issue');

  // The first live issue run is the earliest: the daily table starts there; dry runs and other commands do not count.
  assert.deepEqual(campaignStart({ runs, cards, selection }), { startMs: Date.parse('2026-10-05T16:00:00.500Z'), campaignStartIso: SEARCH, startFrom: 'issue' });
  // A journal rebuilt after a loss: its first live issue run (10/07) is after the oldest card (10/05).
  const rebuilt = [{ command: 'issue', dryRun: false, startedAt: '2026-10-07T16:00:00.000Z' }];
  assert.deepEqual(campaignStart({ runs: rebuilt, cards, selection }), { startMs: Date.parse('2026-10-05T16:00:01Z'), campaignStartIso: SEARCH, startFrom: 'card' });
  // No live issue run in the journal (lost), but cards exist.
  assert.deepEqual(campaignStart({ runs: noIssue, cards, selection }), { startMs: Date.parse('2026-10-05T16:00:01Z'), campaignStartIso: SEARCH, startFrom: 'card' });
  // Nothing issued yet: the time the list was made.
  assert.deepEqual(campaignStart({ runs: noIssue, cards: [], selection }), { startMs: Date.parse(NOW_ISO), campaignStartIso: SEARCH, startFrom: 'selection' });
  assert.deepEqual(campaignStart({ selection }), { startMs: Date.parse(NOW_ISO), campaignStartIso: SEARCH, startFrom: 'selection' });

  assert.equal(cardSearchSince(selection), SEARCH);
  assert.equal(giftCardOrdersQuery(SEARCH), "gateway:gift_card created_at:>='2026-10-01T19:00:00Z'");
  assert.throws(() => campaignStart({ runs, cards, selection: {} }), /selection\.json 里没有有效的 createdAt/);
});

test('buildUsage: top 10 products by quantity (ties by amount), only campaign orders; a test campaign reports kind "test"', () => {
  const cust = gid('Customer', 1);
  const selection = { mode: 'test', recipients: [{ customerId: cust, kind: 'test', tier: null }], params: { giftTiersCents: [1077, 1533, 1977] } };
  const cards = [{ id: gid('GiftCard', 1), createdAt: '2026-10-03T18:00:00Z', enabled: true, expiresOn: '2026-10-19', last4: 'abcd', amountCents: 10, balanceCents: 0, customerId: cust }];
  const quantities = { A: 2, B: 9, C: 9, D: 1, E: 2, F: 3, G: 4, H: 6, I: 7, J: 8, K: 10, L: 11 };
  const amounts = { B: '1.00', C: '5.00' };
  const item = (name, q = quantities[name]) => ({ name, quantity: q, amount: amounts[name] ?? '1.00' });
  const orderNodes = [
    makeGiftCardOrder({ n: 1, customer: cust, createdAt: '2026-10-03T19:00:00Z', total: '10.00', payments: [{ giftCardNumericId: 1, amount: '0.05' }], lineItems: ['B', 'C', 'D', 'E', 'F', 'G'].map((n) => item(n)).concat([item('A', 2)]) }),
    makeGiftCardOrder({ n: 2, customer: cust, createdAt: '2026-10-04T19:00:00Z', total: '20.00', payments: [{ giftCardNumericId: 1, amount: '0.05' }], lineItems: ['H', 'I', 'J', 'K', 'L'].map((n) => item(n)).concat([item('A', 3)]) }),
    makeGiftCardOrder({ n: 3, customer: gid('Customer', 2), createdAt: '2026-10-04T20:00:00Z', total: '99.00', payments: [{ giftCardNumericId: 77, amount: '99.00' }], lineItems: [item('Z', 500)] }),
  ];
  const { usage, warnings } = buildUsage({
    selection, cards, orderNodes, timezone: 'America/Los_Angeles',
    campaignStartMs: Date.parse('2026-10-03T17:00:00Z'), campaignStartIso: '2026-10-03T16:00:00Z', fetchedAt: '2026-10-04T22:00:00.000Z',
  });
  assert.deepEqual(usage.summary.topProducts.map((p) => `${p.name}${p.quantity}`), ['L11', 'K10', 'C9', 'B9', 'J8', 'I7', 'H6', 'A5', 'G4', 'F3']);
  assert.deepEqual(usage.summary.topProducts[7], { name: 'A', quantity: 5, amountCents: 200 });
  assert.deepEqual(usage.summary.byKind, [{ kind: 'test', issued: 1, used: 1, rate: 1, usedCents: 10 }]);
  assert.deepEqual(usage.summary.byTier.map((t) => [t.label, t.issued]), [['$10.77', 0], ['$15.33', 0], ['$19.77', 0]]);
  assert.deepEqual(usage.summary.daily.map((d) => [d.date, d.newCardsUsed, d.orders]), [['2026-10-03', 1, 1], ['2026-10-04', 0, 1]]);
  assert.deepEqual(warnings, [], 'balances and payments agree');
});

test('buildUsage: a cancelled order that kept its card payment is listed, but its payment is not part of the order money', () => {
  const [c1, c2, c9] = [1, 2, 9].map((n) => gid('Customer', n));
  const selection = {
    mode: 'live',
    recipients: [{ customerId: c1, kind: 'ordered', tier: 0 }, { customerId: c2, kind: 'never', tier: 0 }],
    params: { giftTiersCents: [1077, 1533, 1977] },
  };
  const usedCard = (n, customerId) => ({ id: gid('GiftCard', n), createdAt: '2026-10-05T16:30:00Z', enabled: true, expiresOn: '2026-10-19', last4: `x${n}`, amountCents: 1077, balanceCents: 0, customerId });
  const cards = [usedCard(1, c1), usedCard(2, c2), usedCard(9, c9)];
  const orderNodes = [
    // Cancelled without a refund: card 1's money is gone, the order is not revenue.
    makeGiftCardOrder({ n: 1, customer: c1, createdAt: '2026-10-06T18:00:00Z', total: '10.77', cancelled: true, payments: [{ giftCardNumericId: 1, amount: '10.77' }] }),
    // Cancelled and refunded first, then card 2 really used on 10/08: the 10/08 order is its first use.
    makeGiftCardOrder({ n: 2, customer: c2, createdAt: '2026-10-06T19:00:00Z', total: '5.00', cancelled: true, payments: [{ giftCardNumericId: 2, amount: '5.00' }, { giftCardNumericId: 2, amount: '5.00', kind: 'REFUND' }] }),
    makeGiftCardOrder({
      n: 3, customer: c2, createdAt: '2026-10-08T18:00:00Z', total: '30.00',
      payments: [{ giftCardNumericId: 2, amount: '10.77' }, { amount: '2.00', kind: 'REFUND', noReceiptId: true }],
      lineItems: Array.from({ length: 30 }, (_, i) => ({ name: `Item ${i + 1}`, quantity: 1, amount: '1.00' })),
    }),
    // 19 gift-card payments ($10.59 + 18 × $0.01 = $10.77) + the helper's shopify_payments row = 20 transactions.
    makeGiftCardOrder({
      n: 4, customer: c9, createdAt: '2026-10-07T18:00:00Z', total: '10.77',
      payments: [{ giftCardNumericId: 9, amount: '10.59' }, ...Array.from({ length: 18 }, () => ({ giftCardNumericId: 9, amount: '0.01' }))],
    }),
  ];
  const { usage, warnings } = buildUsage({
    selection, cards, orderNodes, timezone: 'America/Los_Angeles',
    campaignStartMs: Date.parse('2026-10-05T16:00:00Z'), campaignStartIso: '2026-10-05T15:00:00Z', fetchedAt: '2026-10-08T20:00:00.000Z',
  });

  assert.deepEqual(usage.orders.map((o) => [o.orderName, o.cancelled, o.campaignCardCents]), [['#1', true, 1077], ['#2', true, 0], ['#4', false, 1077], ['#3', false, 1077]]);
  const s = usage.summary;
  assert.deepEqual([s.orders, s.ordersTotalCents, s.giftCardCents, s.customerPaidCents, s.usedCents], [2, 4077, 2154, 1923, 3231]);
  // Card 1 has only the cancelled order (10/06); card 2's first real use is 10/08, not the refunded 10/06 order.
  assert.deepEqual(s.daily.map((d) => [d.date, d.newCardsUsed, d.cumulativeCardsUsed]), [['2026-10-05', 0, 0], ['2026-10-06', 1, 1], ['2026-10-07', 1, 2], ['2026-10-08', 1, 3]]);
  assert.deepEqual(usage.unmatched.map((u) => [u.orderName, u.amountCents, u.reason]), [['#3', 200, '回执里没有礼品卡 ID（退款）']]);
  // Card 9's holder is not on the list: counted in the totals, not in the tier / kind rows.
  assert.deepEqual([s.issuedCards, s.byTier[0].issued, s.byKind.map((k) => k.issued)], [3, 2, [1, 1]]);
  assert.deepEqual(warnings, [
    '1 张本活动的卡不属于名单里的客户（例如 gid://shopify/GiftCard/9），不计入按档位和按客户类型的统计；请运行 verify 核对',
    '1 笔订单的商品数达到单次查询上限 30 个，商品明细可能不完整：#3',
    '1 笔订单的交易数达到单次查询上限 20 笔，用卡付款可能没有取全：#4',
  ]);
});

// ---------------------------------------------------------------------------
// The command
// ---------------------------------------------------------------------------

test('no cards and no orders: zero summary, both files written, Excel refreshed', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  fake = installFakeShopify({});
  const result = await go({ now: () => new Date('2026-10-02T17:00:00.000Z') });

  assert.equal(result.exitCode, 0);
  const u = result.usage;
  assert.equal(u.version, 1);
  assert.equal(u.fetchedAt, '2026-10-02T17:00:00.000Z');
  assert.equal(u.timezone, 'America/Los_Angeles');
  assert.equal(u.campaignStartIso, '2026-10-01T19:00:00Z');
  assert.deepEqual([u.cards, u.payments, u.orders, u.unmatched], [[], [], [], []]);
  assert.deepEqual(u.summary, {
    issuedCards: 0, issuedCents: 0, usedCards: 0, usedCents: 0, usedRate: 0, usedCentsRate: 0,
    orders: 0, ordersTotalCents: 0, avgOrderCents: 0, giftCardCents: 0, customerPaidCents: 0,
    byTier: [
      { tier: 0, label: '$10.77', issued: 0, used: 0, rate: 0, usedCents: 0 },
      { tier: 1, label: '$15.33', issued: 0, used: 0, rate: 0, usedCents: 0 },
      { tier: 2, label: '$19.77', issued: 0, used: 0, rate: 0, usedCents: 0 },
    ],
    byKind: [
      { kind: 'ordered', issued: 0, used: 0, rate: 0, usedCents: 0 },
      { kind: 'never', issued: 0, used: 0, rate: 0, usedCents: 0 },
    ],
    daily: [
      { date: '2026-10-01', newCardsUsed: 0, orders: 0, ordersTotalCents: 0, cumulativeCardsUsed: 0, cumulativeRate: 0 },
      { date: '2026-10-02', newCardsUsed: 0, orders: 0, ordersTotalCents: 0, cumulativeCardsUsed: 0, cumulativeRate: 0 },
    ],
    topProducts: [],
  });
  assert.deepEqual(readUsageFile(paths.usage), u);
  assert.deepEqual(readUsageFile(path.join(paths.usageDir, '2026-10-02.json')), u);
  assert.equal(reports.length, 1);
  assert.match(output(), /发出礼品卡：0 张，面额 \$0\.00/);
});

test('cards: used vs unused by balance, usedCents; other notes and cards outside the window are not campaign cards', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  writeIssueRuns(paths);
  installScenario();
  fake.state.giftCards.find((g) => g.id === gid('GiftCard', 505)).enabled = false;
  const { exitCode, usage } = await go();

  assert.equal(exitCode, 0);
  assert.deepEqual(usage.cards.map((c) => tail(c.giftCardId)), ['501', '502', '503', '504', '505', '506', '507']);
  assert.deepEqual(usage.cards[0], {
    giftCardId: gid('GiftCard', 501), customerId: gid('Customer', 1), last4: 'x501', initialCents: 1077, balanceCents: 0,
    usedCents: 1077, enabled: true, expiresOn: '2026-10-19', createdAt: '2026-10-05T16:30:00Z',
  });
  assert.deepEqual(usage.cards.map((c) => [c.balanceCents, c.usedCents]), [[0, 1077], [500, 1033], [0, 1977], [0, 1077], [1077, 0], [0, 1077], [1077, 0]]);
  assert.equal(usage.cards[4].enabled, false);
  const s = usage.summary;
  assert.deepEqual(
    [s.issuedCards, s.issuedCents, s.usedCards, s.usedCents, s.usedRate, s.usedCentsRate],
    [7, 8895, 5, 6241, 0.7143, 0.7016],
  );
  // The campaign-card search starts one hour before the selection was made.
  assert.match(fake.opsNamed('FindCampaignCards')[0].variables.query, /created_at:>='2026-10-01T19:00:00Z'/);
});

test('orders: payments from receipts, refunds, two cards on one order, non-campaign cards ignored, missing ids listed, cancelled orders kept out of revenue', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  writeIssueRuns(paths);
  installScenario();
  const { exitCode, usage } = await go();
  assert.equal(exitCode, 0);

  // Ledger in processing order; FAILURE and AUTHORIZATION attempts and the shopify_payments rows are ignored.
  assert.deepEqual(usage.payments.map((p) => [p.orderName, tail(p.giftCardId), p.kind, p.amountCents, p.netCents]), [
    ['#2006', '507', 'SALE', 1077, 1077],
    ['#2006', '507', 'REFUND', 1077, -1077],
    ['#2001', '501', 'SALE', 1077, 1077],
    ['#2002', '502', 'SALE', 1533, 1533],
    ['#2003', '503', 'SALE', 1977, 1977],
    ['#2003', '504', 'SALE', 1077, 1077],
    ['#2002', '502', 'REFUND', 500, -500],
  ]);
  // Customer 3 paid with their own card and with customer 4's card.
  assert.deepEqual(usage.payments[5], {
    orderId: gid('Order', 2003), orderName: '#2003', orderCreatedAt: '2026-10-07T20:00:00Z', orderCustomerId: gid('Customer', 3),
    giftCardId: gid('GiftCard', 504), cardCustomerId: gid('Customer', 4), kind: 'SALE', amountCents: 1077, netCents: 1077,
    processedAt: '2026-10-07T20:00:00Z',
  });
  assert.equal(usage.payments[6].processedAt, '2026-10-08T18:00:00Z');

  assert.deepEqual(usage.orders.map((o) => [o.orderName, o.totalCents, o.campaignCardCents, o.cancelled, o.giftCardIds.map(tail)]), [
    ['#2006', 3000, 0, true, ['507']],
    ['#2001', 4500, 1077, false, ['501']],
    ['#2002', 1533, 1033, false, ['502']],
    ['#2003', 6000, 3054, false, ['503', '504']],
  ]);
  assert.deepEqual(usage.orders[1], {
    orderId: gid('Order', 2001), orderName: '#2001', createdAt: '2026-10-07T03:30:00Z', customerId: gid('Customer', 1),
    totalCents: 4500, campaignCardCents: 1077, cancelled: false, giftCardIds: [gid('GiftCard', 501)],
    lineItems: [
      { name: 'Latex Balloon 12in', sku: 'LB-12', quantity: 10, amountCents: 800 },
      { name: 'Helium Tank', sku: 'HT-1', quantity: 1, amountCents: 3000 },
    ],
  });
  assert.deepEqual(usage.orders[0].lineItems, [{ name: 'Confetti', sku: '', quantity: 100, amountCents: 3000 }]);

  // Recipient 6's payment has no card id → listed for review; the non-recipient's (#2007) is not.
  assert.deepEqual(usage.unmatched, [{
    orderId: gid('Order', 2005), orderName: '#2005', orderCreatedAt: '2026-10-08T01:00:00Z', customerId: gid('Customer', 6),
    amountCents: 1077, processedAt: '2026-10-08T01:00:00Z', reason: NO_RECEIPT_ID_REASON,
  }]);
  assert.equal(NO_RECEIPT_ID_REASON, '回执里没有礼品卡 ID');

  // Revenue: #2001 + #2002 + #2003 (the cancelled #2006 is listed above but not counted).
  // The split is taken at checkout, like the order totals: #2002 was paid entirely by card 502
  // ($15.33); the later $5 refund went back to the card, so the customer paid nothing besides it.
  // Card side: 1077 + 1533 + 1977 + 1077 = 5664 (the ledger above stays net: #2002 nets 1033).
  const s = usage.summary;
  assert.deepEqual(
    [s.orders, s.ordersTotalCents, s.avgOrderCents, s.giftCardCents, s.customerPaidCents],
    [3, 12033, 4011, 5664, 6369],
  );

  // Card 506 was spent by #2005, whose receipt names no card: flagged on the console, not guessed.
  assert.match(output(), /WARN 1 张卡的已用金额和订单里的用卡付款对不上，例如尾号 x506：按余额已用 \$10\.77，订单付款合计 \$0\.00/);
  assert.match(output(), /WARN {3}1 笔礼品卡付款的回执里没有礼品卡 ID，已单独列在使用明细里，请人工核对/);
  assert.match(output(), /另有 1 笔已取消的订单用过本活动的卡/);
});

test('summary: daily rows by Los Angeles date, top products, by tier and by customer kind', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  writeIssueRuns(paths);
  installScenario();
  const { usage } = await go();
  const s = usage.summary;

  // #2001 is 10/07 in UTC but 10/06 20:30 in Los Angeles; #2002 (07:30 UTC) is 10/07 00:30 there.
  // Used cards by first payment: 501 → 10/06; 502, 503, 504 → 10/07. 506 (no matched payment) and 507
  // (refunded, balance full) are not in the trend. The cancelled #2006 (10/06) is not an order.
  assert.deepEqual(s.daily, [
    { date: '2026-10-05', newCardsUsed: 0, orders: 0, ordersTotalCents: 0, cumulativeCardsUsed: 0, cumulativeRate: 0 },
    { date: '2026-10-06', newCardsUsed: 1, orders: 1, ordersTotalCents: 4500, cumulativeCardsUsed: 1, cumulativeRate: 0.1429 },
    { date: '2026-10-07', newCardsUsed: 3, orders: 2, ordersTotalCents: 7533, cumulativeCardsUsed: 4, cumulativeRate: 0.5714 },
    { date: '2026-10-08', newCardsUsed: 0, orders: 0, ordersTotalCents: 0, cumulativeCardsUsed: 4, cumulativeRate: 0.5714 },
  ]);

  // Non-cancelled campaign orders only: no Party Hat (#2004), Confetti (#2006, cancelled) or Ribbon (#2005).
  assert.deepEqual(s.topProducts, [
    { name: 'Latex Balloon 12in', quantity: 20, amountCents: 1600 },
    { name: 'Foil Star', quantity: 2, amountCents: 1533 },
    { name: 'Balloon Arch Kit', quantity: 1, amountCents: 5200 },
    { name: 'Helium Tank', quantity: 1, amountCents: 3000 },
  ]);

  assert.deepEqual(s.byTier, [
    { tier: 0, label: '$10.77', issued: 5, used: 3, rate: 0.6, usedCents: 3231 },
    { tier: 1, label: '$15.33', issued: 1, used: 1, rate: 1, usedCents: 1033 },
    { tier: 2, label: '$19.77', issued: 1, used: 1, rate: 1, usedCents: 1977 },
  ]);
  assert.deepEqual(s.byKind, [
    { kind: 'ordered', issued: 5, used: 4, rate: 0.8, usedCents: 5164 },
    { kind: 'never', issued: 2, used: 1, rate: 0.5, usedCents: 1077 },
  ]);

  const out = output();
  assert.match(out, /使用情况（截至 2026-10-08 22:00，洛杉矶时间）/);
  assert.match(out, /发出礼品卡：7 张，面额 \$88\.95/);
  assert.match(out, /已经使用：5 张，占 71\.4%；已用 \$62\.41，占面额 70\.2%/);
  assert.match(out, /带来订单：3 笔，订单总额 \$120\.33，平均每单 \$40\.11/);
  assert.match(out, /其中礼品卡抵扣 \$56\.64，顾客另外支付 \$63\.69/);
  assert.match(out, new RegExp(`Excel：${paths.excel.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}`));
  // The English edition's path follows the Chinese one.
  const excelAt = log.lines.indexOf(`INFO Excel：${paths.excel}`);
  assert.ok(excelAt >= 0, out);
  assert.equal(log.lines[excelAt + 1], `INFO 英文版 Excel：${paths.excelEn}`, out);
  assert.equal(/fake-token|secret/.test(out), false, 'never logs credentials');
});

test('the orders search always starts 1 hour before the list was made; the daily table starts at the first live issue run', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  writeIssueRuns(paths);
  installScenario();
  const SEARCH = "gateway:gift_card created_at:>='2026-10-01T19:00:00Z'";
  const withJournal = await go();
  assert.equal(withJournal.usage.campaignStartIso, '2026-10-01T19:00:00Z', 'not the first issue run (10/05)');
  assert.deepEqual(fake.opsNamed('GiftCardOrders').map((o) => o.variables.query), [SEARCH]);
  assert.equal(withJournal.usage.summary.daily[0].date, '2026-10-05');
  assert.match(output(), /正在读取 2026-10-01 12:00（洛杉矶时间）以来用礼品卡付款的订单（从名单生成前 1 小时起，本活动的卡不会比这更早）/);
  fake.restore();
  resetClient();

  // Same store, but no journal (the file was lost): same search; the daily table starts at the oldest card (10/05).
  fs.rmSync(paths.journal);
  installScenario();
  const noJournal = await go();
  assert.equal(noJournal.exitCode, 0);
  assert.equal(noJournal.usage.campaignStartIso, '2026-10-01T19:00:00Z');
  assert.deepEqual(fake.opsNamed('GiftCardOrders').map((o) => o.variables.query), [SEARCH]);
  assert.deepEqual(noJournal.usage.summary.daily.map((d) => d.date), ['2026-10-05', '2026-10-06', '2026-10-07', '2026-10-08']);
  assert.deepEqual(noJournal.usage.payments, withJournal.usage.payments);
  assert.deepEqual(noJournal.usage.summary, withJournal.usage.summary);
});

test('a journal rebuilt after a loss (verify, then a later issue batch) still counts the orders paid before that batch', async () => {
  // Regression: the orders search used to start at the first live issue run in the journal (10/07 here).
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  const at = (t) => ({ now: () => t });
  appendJournal(paths.journal, { op: 'run.start', run: 'v', command: 'verify', dryRun: false, batch: null, limit: null }, at('2026-10-06T20:00:00.000Z'));
  appendJournal(paths.journal, { op: 'reconcile.found', cid: gid('Customer', 1), giftCardId: gid('GiftCard', 501), last4: 'x501', amountCents: 1077, createdAt: '2026-10-05T16:30:00Z', source: 'verify', run: 'v' }, at('2026-10-06T20:01:00.000Z'));
  appendJournal(paths.journal, { op: 'run.end', run: 'v', summary: {}, exitCode: 0 }, at('2026-10-06T20:02:00.000Z'));
  appendJournal(paths.journal, { op: 'run.start', run: 'i', command: 'issue', dryRun: false, batch: 1, limit: 1 }, at('2026-10-07T16:00:00.000Z'));
  appendJournal(paths.journal, { op: 'run.end', run: 'i', summary: {}, exitCode: 0 }, at('2026-10-07T16:05:00.000Z'));
  fake = installFakeShopify({
    giftCards: [card(501, 1, 1077, { createdAt: '2026-10-05T16:30:00Z' }), card(502, 2, 1533, { createdAt: '2026-10-07T16:01:00Z' })],
    giftCardOrders: [makeGiftCardOrder({ n: 2001, customer: gid('Customer', 1), createdAt: '2026-10-05T20:00:00Z', total: '30.00', payments: [{ giftCardNumericId: 501, amount: '10.77' }] })],
  });
  fake.spendCard(gid('GiftCard', 501), 1077);

  const { exitCode, usage } = await go();
  assert.equal(exitCode, 0);
  assert.equal(usage.campaignStartIso, '2026-10-01T19:00:00Z');
  assert.deepEqual([usage.summary.orders, usage.summary.ordersTotalCents, usage.summary.giftCardCents, usage.summary.usedCards], [1, 3000, 1077, 1]);
  assert.deepEqual(usage.orders.map((o) => o.orderName), ['#2001']);
  // The daily table starts with the oldest card (10/05 09:30 in Los Angeles), not the 10/07 batch.
  assert.deepEqual(usage.summary.daily.map((d) => [d.date, d.newCardsUsed, d.orders]), [
    ['2026-10-05', 1, 1], ['2026-10-06', 0, 0], ['2026-10-07', 0, 0], ['2026-10-08', 0, 0],
  ]);
  assert.equal(log.lines.some((l) => l.includes('对不上')), false, 'the balance and the payment agree');
});

test('refunds back to a campaign card: the order split is taken at checkout; the card ledger stays net', async () => {
  // Regression: "顾客另外支付" mixed the order total before refunds with the card amount after refunds.
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  writeIssueRuns(paths);
  fake = installFakeShopify({
    giftCards: [card(501, 1, 1077), card(502, 2, 1533), card(503, 3, 1977), card(504, 4, 1077)],
    giftCardOrders: [
      // $45.00 order: $10.77 by card 501, the rest by credit card; later $5.00 goes back to the card.
      makeGiftCardOrder({
        n: 2001, customer: gid('Customer', 1), createdAt: '2026-10-06T18:00:00Z', total: '45.00',
        payments: [{ giftCardNumericId: 501, amount: '10.77' }, { giftCardNumericId: 501, amount: '5.00', kind: 'REFUND', processedAt: '2026-10-07T18:00:00Z' }],
      }),
      // $30.00 order, not cancelled: card 502 paid $15.33, then all of it went back to the card.
      makeGiftCardOrder({
        n: 2002, customer: gid('Customer', 2), createdAt: '2026-10-06T19:00:00Z', total: '30.00',
        payments: [{ giftCardNumericId: 502, amount: '15.33' }, { giftCardNumericId: 502, amount: '15.33', kind: 'REFUND', processedAt: '2026-10-07T19:00:00Z' }],
      }),
      // $50.00 order paid through an authorization + capture of card 503 (counted once, as the CAPTURE).
      makeGiftCardOrder({
        n: 2003, customer: gid('Customer', 3), createdAt: '2026-10-06T20:00:00Z', total: '50.00',
        payments: [{ giftCardNumericId: 503, amount: '19.77', kind: 'AUTHORIZATION' }, { giftCardNumericId: 503, amount: '19.77', kind: 'CAPTURE' }],
      }),
      // Cancelled without a refund: card 504's $1.00 is gone, but it is not part of the order money.
      makeGiftCardOrder({ n: 2004, customer: gid('Customer', 4), createdAt: '2026-10-06T21:00:00Z', total: '20.00', cancelled: true, payments: [{ giftCardNumericId: 504, amount: '1.00' }] }),
    ],
  });
  fake.spendCard(gid('GiftCard', 501), 577); // $10.77 − $5.00 back
  fake.spendCard(gid('GiftCard', 503), 1977);
  fake.spendCard(gid('GiftCard', 504), 100);
  const { exitCode, usage } = await go();
  assert.equal(exitCode, 0);
  const s = usage.summary;

  // Orders #2001–#2003: $125.00; at checkout the cards paid 1077 + 1533 + 1977 = 4587; the customers 7913.
  assert.deepEqual([s.orders, s.ordersTotalCents, s.giftCardCents, s.customerPaidCents], [3, 12500, 4587, 7913]);
  assert.equal(s.giftCardCents + s.customerPaidCents, s.ordersTotalCents);
  // The ledger is unchanged: net per order and per payment.
  assert.deepEqual(usage.orders.map((o) => [o.orderName, o.totalCents, o.campaignCardCents, o.cancelled]), [
    ['#2001', 4500, 577, false], ['#2002', 3000, 0, false], ['#2003', 5000, 1977, false], ['#2004', 2000, 100, true],
  ]);
  assert.deepEqual(usage.payments.map((p) => [p.orderName, p.kind, p.amountCents, p.netCents]), [
    ['#2001', 'SALE', 1077, 1077], ['#2002', 'SALE', 1533, 1533], ['#2003', 'CAPTURE', 1977, 1977], ['#2004', 'SALE', 100, 100],
    ['#2001', 'REFUND', 500, -500], ['#2002', 'REFUND', 1533, -1533],
  ]);
  // Card 502 is back to full: not used. Balances: 501 used 577, 503 used 1977, 504 used 100.
  assert.deepEqual([s.usedCards, s.usedCents], [3, 2654]);
  assert.equal(log.lines.some((l) => l.includes('对不上')), false, 'balances agree with the net ledger');
  assert.match(output(), /其中礼品卡抵扣 \$45\.87，顾客另外支付 \$79\.13/);
  // The run log records the same checkout-time figure.
  const end = readJournal(paths.journal).at(-1);
  assert.deepEqual([end.op, end.summary.giftCardCents, end.summary.ordersTotalCents], ['run.end', 4587, 12500]);
});

test('pages through every gift-card order (25 per page)', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  writeIssueRuns(paths);
  const giftCardOrders = Array.from({ length: 30 }, (_, i) => makeGiftCardOrder({
    n: 3000 + i, customer: gid('Customer', 3), createdAt: new Date(Date.parse('2026-10-06T16:00:00Z') + i * 60_000).toISOString(), total: '2.00',
    payments: [{ giftCardNumericId: 503, amount: '0.50' }], lineItems: [{ name: 'Sticker', quantity: 1, amount: '2.00' }],
  }));
  fake = installFakeShopify({ giftCards: CAMPAIGN_CARDS(), giftCardOrders });
  fake.spendCard(gid('GiftCard', 503), 1500);
  const { exitCode, usage } = await go();

  assert.equal(exitCode, 0);
  assert.deepEqual(fake.opsNamed('GiftCardOrders').map((o) => o.variables.after), [null, 'cur-25']);
  assert.equal(usage.orders.length, 30);
  assert.deepEqual([usage.summary.orders, usage.summary.giftCardCents, usage.summary.usedCards], [30, 1500, 1]);
  assert.deepEqual(usage.summary.topProducts, [{ name: 'Sticker', quantity: 30, amountCents: 6000 }]);
  assert.equal(log.lines.some((l) => l.includes('对不上')), false, 'balances match the payments');
});

test('usage.json and usage/<store-local date>.json hold the same data; the journal records the run', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  writeIssueRuns(paths);
  installScenario();
  const { exitCode, usage } = await go();
  assert.equal(exitCode, 0);

  // The run is at 05:00 UTC on 10/09 = 22:00 on 10/08 in Los Angeles.
  assert.deepEqual(fs.readdirSync(paths.usageDir), ['2026-10-08.json']);
  assert.deepEqual(readUsageFile(path.join(paths.usageDir, '2026-10-08.json')), usage);
  assert.deepEqual(readUsageFile(paths.usage), usage);
  assert.equal(usage.fetchedAt, T.now);

  const entries = readJournal(paths.journal).filter((e) => e.op.startsWith('run.') && e.command !== 'issue' && !['r0', 'r1', 'r2'].includes(e.run));
  assert.equal(entries.length, 2);
  const [start, end] = entries;
  assert.deepEqual([start.op, start.command, start.dryRun, start.t], ['run.start', 'usage', false, T.now]);
  assert.deepEqual([end.op, end.run, end.exitCode], ['run.end', start.run, 0]);
  assert.deepEqual(end.summary, { issuedCards: 7, usedCards: 5, usedCents: 6241, orders: 3, ordersTotalCents: 12033, giftCardCents: 5664, unmatched: 1 });
  // Read-only: no Shopify writes of any kind.
  assert.deepEqual([fake.state.calls.create.length, fake.state.calls.tag.length, fake.state.calls.notify.length], [0, 0, 0]);
});

test('run lock: held as "usage" while fetching, released before the Excel is written', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  installScenario();
  const inner = globalThis.fetch;
  const holders = [];
  globalThis.fetch = async (url, init) => {
    holders.push(runningCommand(paths));
    return inner(url, init);
  };
  const { exitCode } = await go();

  assert.equal(exitCode, 0);
  assert.ok(holders.length >= 3);
  assert.deepEqual([...new Set(holders)], ['usage']);
  assert.equal(reports.length, 1);
  assert.equal(reports[0].lockHolder, null, 'writeReport runs after the run lock is released');
  assert.equal(reports[0].usageWritten, true, 'and after usage.json is written');
  assert.equal(reports[0].opts.config, env.config);
  assert.equal(reports[0].opts.paths.excel, paths.excel);
  assert.equal(runningCommand(paths), null);
});

test('another command holding the run lock: exit 1, nothing fetched or written', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  installScenario();
  const release = acquireRunLock(paths, 'issue');
  try {
    const { exitCode, usage } = await go();
    assert.equal(exitCode, 1);
    assert.equal(usage, null);
    assert.match(output(), /ERROR issue 正在运行/);
    assert.equal(fake.state.calls.token.length + fake.state.calls.ops.length, 0);
    assert.equal(fs.existsSync(paths.usage), false);
    assert.equal(fs.existsSync(paths.journal), false);
    assert.equal(reports.length, 0, 'the running command refreshes the Excel itself');
  } finally {
    release();
  }
});

test('no selection, or a selection from another campaign: exit 1 before touching Shopify', async () => {
  fake = installFakeShopify({});
  const missing = await go();
  assert.equal(missing.exitCode, 1);
  assert.match(output(), /还没有名单.*请先运行 node index\.js select/);

  // The very first select is still running: say so instead of "run select".
  const release = acquireRunLock(paths, 'select');
  try {
    assert.equal((await go()).exitCode, 1);
    assert.match(output(), /还没有名单：select 正在运行，请等它跑完再运行 usage/);
  } finally {
    release();
  }

  const selection = await selectionFixture(env.config, { customers: fixtureCustomers(), write: false });
  writeJsonAtomic(paths.selection, { ...selection, campaignId: '2026-09' });
  const other = await go();
  assert.equal(other.exitCode, 1);
  assert.match(output(), /名单属于活动 2026-09，和 CAMPAIGN_ID=2026-10 不一致/);

  assert.equal(fake.state.calls.token.length + fake.state.calls.ops.length, 0);
  assert.equal(reports.length, 0);
  assert.equal(fs.existsSync(paths.runLock), false);
});

test('a Shopify failure: exit 1, no usage files, run.end records the error, lock released, Excel still refreshed', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  writeIssueRuns(paths);
  fake = installFakeShopify({
    giftCards: CAMPAIGN_CARDS(),
    giftCardOrders: scenarioOrders(),
    failures: { GiftCardOrders: [{ kind: 'graphqlError', message: 'Internal error', code: 'INTERNAL_SERVER_ERROR' }] },
  });
  const { exitCode, usage } = await go();

  assert.equal(exitCode, 1);
  assert.equal(usage, null);
  assert.match(output(), /ERROR usage 没有完成：GraphQL error: Internal error/);
  assert.equal(fs.existsSync(paths.usage), false);
  assert.equal(fs.existsSync(paths.usageDir), false);
  const end = readJournal(paths.journal).at(-1);
  assert.deepEqual([end.op, end.exitCode], ['run.end', 1]);
  assert.match(end.summary.error, /Internal error/);
  assert.equal(runningCommand(paths), null);
  assert.equal(reports.length, 1);
});

test('a failing Excel refresh is only a warning', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  installScenario();
  const { exitCode, usage } = await go({
    writeReport: async () => {
      throw new Error('disk full');
    },
  });
  assert.equal(exitCode, 0);
  assert.ok(usage);
  assert.match(output(), /WARN Excel 没有更新：disk full/);
  assert.equal(fs.existsSync(paths.usage), true);
});

test('warnings of the Excel writer are printed once: writeReport logs them, usage does not repeat result.warnings', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  installScenario();
  const file = path.join(paths.dir, 'x.xlsx');
  const warning = 'Excel 正打开此文件，请关闭后重新打开才能看到最新内容';
  // Like the real writer: each warning is logged when it happens and also returned.
  const fileEn = path.join(paths.dir, 'x-en.xlsx');
  await go({
    writeReport: async (opts) => {
      opts.log.warn(warning);
      return { file, fileEn, out: null, outEn: null, warnings: [warning] };
    },
  });
  assert.equal(log.lines.filter((l) => l === `WARN ${warning}`).length, 1, output());
  assert.match(output(), new RegExp(`Excel：${file.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}`));
  assert.ok(log.lines.includes(`INFO 英文版 Excel：${fileEn}`), output());
});

test('when the English workbook failed (fileEn null) usage prints no English path and keeps exit 0', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  installScenario();
  const { exitCode } = await go({
    writeReport: async (opts) => {
      opts.log.warn('英文版 Excel 没有生成：disk full');
      return { file: opts.paths.excel, fileEn: null, out: null, outEn: null, warnings: ['英文版 Excel 没有生成：disk full'] };
    },
  });
  assert.equal(exitCode, 0, output());
  assert.ok(log.lines.includes(`INFO Excel：${paths.excel}`), output());
  assert.equal(log.lines.filter((l) => l.includes('英文版 Excel：')).length, 0, output());
  assert.equal(log.lines.filter((l) => l === 'WARN 英文版 Excel 没有生成：disk full').length, 1, 'printed once, by writeReport');
});

test('with the real Excel writer, "Excel 正打开此文件" is printed exactly once', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  writeIssueRuns(paths);
  installScenario();
  fs.writeFileSync(path.join(paths.dir, `~$${path.basename(paths.excel)}`), 'owner'); // the workbook is open in Excel
  const { exitCode } = await go({ writeReport: undefined }); // the default: src/report/excel.js
  assert.equal(exitCode, 0);
  assert.equal(log.lines.filter((l) => l === `WARN ${OPEN_IN_EXCEL_WARNING}`).length, 1, output());
  assert.ok(fs.existsSync(paths.excel), 'the workbook was written');
  // The English edition is written next to it and its path printed right after the Chinese one.
  assert.ok(fs.existsSync(paths.excelEn), `the English workbook was written:\n${output()}`);
  const excelAt = log.lines.indexOf(`INFO Excel：${paths.excel}`);
  assert.ok(excelAt >= 0, output());
  assert.equal(log.lines[excelAt + 1], `INFO 英文版 Excel：${paths.excelEn}`, output());
});

test('a page that claims more orders but gives no cursor stops the run instead of silently truncating', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  fake = installFakeShopify({ giftCards: CAMPAIGN_CARDS() });
  const inner = globalThis.fetch;
  globalThis.fetch = async (url, init) => {
    if (String(url).endsWith('/graphql.json') && JSON.parse(init.body).query.includes('query GiftCardOrders')) {
      const body = { data: { orders: { nodes: [], pageInfo: { hasNextPage: true, endCursor: null } } } };
      return new Response(JSON.stringify(body), { status: 200, headers: { 'content-type': 'application/json' } });
    }
    return inner(url, init);
  };
  const { exitCode } = await go();
  assert.equal(exitCode, 1);
  assert.match(output(), /ERROR usage 没有完成：Shopify 说还有下一页订单，却没有给出分页游标/);
  assert.equal(fs.existsSync(paths.usage), false);
});

test('cards recorded in the journal but missing from Shopify are reported', async () => {
  await selectionFixture(env.config, { customers: fixtureCustomers() });
  writeIssueRuns(paths);
  appendJournal(paths.journal, { op: 'create.start', cid: gid('Customer', 1), amountCents: 1077, batch: 1, run: 'r1' }, { now: () => '2026-10-05T16:01:00.000Z' });
  appendJournal(paths.journal, { op: 'create.ok', cid: gid('Customer', 1), giftCardId: gid('GiftCard', 501), last4: 'x501', amountCents: 1077, batch: 1, run: 'r1' }, { now: () => '2026-10-05T16:01:01.000Z' });
  appendJournal(paths.journal, { op: 'create.start', cid: gid('Customer', 2), amountCents: 1533, batch: 1, run: 'r1' }, { now: () => '2026-10-05T16:01:02.000Z' });
  appendJournal(paths.journal, { op: 'create.ok', cid: gid('Customer', 2), giftCardId: gid('GiftCard', 777), last4: 'x777', amountCents: 1533, batch: 1, run: 'r1' }, { now: () => '2026-10-05T16:01:03.000Z' });
  fake = installFakeShopify({ giftCards: CAMPAIGN_CARDS() });
  const { exitCode } = await go();
  assert.equal(exitCode, 0);
  assert.match(output(), /WARN 本地日志里有 1 张卡没有出现在 Shopify 的本活动卡列表里（例如 gid:\/\/shopify\/GiftCard\/777）/);
});
