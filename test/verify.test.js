import { test, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { installFakeShopify } from './fake-shopify.js';
import { testConfig, selectionFixture, makeCustomer, makeOrder, memoryLog, gid, NOW_ISO } from './helpers.js';
import { appendJournal, campaignPaths, foldJournal, readJournal, readJson, writeJsonAtomic, STATUS } from '../src/campaign.js';
import { resetClient } from '../src/shopify.js';
import { runVerify, analyzeVerify, ISSUE_TYPES } from '../src/verify.js';
import { buildTestSelection } from '../src/select/selection.js';
import { parseCustomer } from '../src/select/rules.js';
import { OPEN_IN_EXCEL_WARNING } from '../src/report/excel.js';

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const RUN_AT = '2026-10-05T18:00:00.000Z'; // "now" of most verify runs (store time 11:00)
const at = (iso) => () => new Date(iso);
const minutesBefore = (iso, m) => new Date(Date.parse(iso) - m * 60_000).toISOString();
const minutesAfter = (iso, m) => new Date(Date.parse(iso) + m * 60_000).toISOString();
const noSleep = async () => {};
const tail = (id) => (id ? String(id).split('/').pop() : null);
/** CJK and full-width characters: the English texts of verify.json must contain none. */
const CJK = /[　-〿㐀-䶿一-鿿豈-﫿＀-￯]/;

// Last-order totals → amounts with the default tiers: $150 → $15.33, $83.40 → $10.77, $350 → $19.77, $50 → $10.77.
const TOTALS = { 1: '150.00', 2: '83.40', 3: '350.00' };
const AMOUNT = { 1: 1533, 2: 1077, 3: 1977 };
const amountOf = (n) => AMOUNT[n] ?? 1077;

function customers(count) {
  return Array.from({ length: count }, (_, i) => makeCustomer({ n: i + 1, lastOrder: makeOrder({ n: 100 + i + 1, total: TOTALS[i + 1] ?? '50.00' }) }));
}

/** The fake's version of the same customers, some carrying the sent tag. */
function shopifyCustomers(list, taggedNs, tag) {
  return list.map((c) => {
    const n = Number(tail(c.id));
    return { ...structuredClone(c), tags: taggedNs.includes(n) ? [tag] : [] };
  });
}

/** A GiftCard node as the fake stores it. */
function card(n, customerN, { cents = 1077, balance = cents, enabled = true, createdAt = '2026-10-05T16:00:00Z', note = 'gift-card-promo [campaign:2026-10]' } = {}) {
  return {
    id: gid('GiftCard', n),
    createdAt,
    note,
    templateSuffix: 'gift-card-promo',
    enabled,
    lastCharacters: `x${String(n).slice(-3)}`,
    initialValue: { amount: (cents / 100).toFixed(2), currencyCode: 'USD' },
    balance: { amount: (balance / 100).toFixed(2), currencyCode: 'USD' },
    customer: customerN ? { id: gid('Customer', customerN) } : null,
    expiresOn: '2026-10-19',
  };
}

function writeJournal(paths, rows) {
  for (const [t, entry] of rows) appendJournal(paths.journal, entry, { now: () => t });
}

/** Journal lines of a completed issue: create.start → create.ok (→ tag.ok). */
function issued(n, cardN, cents, { tagged = true, t = '2026-10-05T16:00:00.000Z' } = {}) {
  const cid = gid('Customer', n);
  return [
    [t, { op: 'create.start', cid, amountCents: cents, batch: 1, run: 'r-issue' }],
    [t, { op: 'create.ok', cid, giftCardId: gid('GiftCard', cardN), last4: `x${String(cardN).slice(-3)}`, amountCents: cents, batch: 1, run: 'r-issue' }],
    ...(tagged ? [[t, { op: 'tag.ok', cid, run: 'r-issue' }]] : []),
  ];
}

/** Journal lines of a create whose answer was lost (optionally marked unknown). */
function lost(n, t, { unknown = true } = {}) {
  const cid = gid('Customer', n);
  return [
    [t, { op: 'create.start', cid, amountCents: amountOf(n), batch: 1, run: 'r-issue' }],
    ...(unknown ? [[t, { op: 'create.unknown', cid, error: 'Network error calling Shopify: fetch failed', run: 'r-issue' }]] : []),
  ];
}

function reportStub({ fail = null, onCall = null } = {}) {
  const calls = [];
  const fn = async (args) => {
    calls.push(args);
    if (onCall) onCall(args);
    if (fail) throw new Error(fail);
    return { file: args.paths.excel, fileEn: args.paths.excelEn, out: null, outEn: null, warnings: [] };
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

const WRITE_OPS = ['GiftCardCreate', 'TagsAdd', 'SendGiftCardNotification', 'RunCustomerExport', 'CancelBulk'];

let env;
let fake;
afterEach(() => {
  fake?.restore();
  fake = null;
  resetClient();
  env?.cleanup();
  env = null;
});

/**
 * One customer per issue type (plus controls), all in one campaign:
 *   c1  issued + tagged, card ok                          → no issue
 *   c2  issued + tagged, card hidden by index lag         → missing-in-shopify
 *   c3  journal empty, Shopify has a card, not tagged     → not-in-journal (+ reconcile.found)
 *   c4  issued + tagged, Shopify has a second card        → duplicate-cards
 *   c5  issued + tagged, card amount $15.33 ≠ list $10.77 → amount-mismatch
 *   c6  issued + tagged, card disabled                    → card-disabled
 *   c7  journal says tagged, Shopify has no tag           → tag-missing
 *   c8  tagged in Shopify, no card anywhere               → tag-without-card
 *   c9  unknown 20 min ago, card exists                   → resolved-unknown (+ reconcile.found)
 *   c10 unknown 20 min ago, no card                       → resolved-unknown (+ reconcile.none)
 *   c11 in_progress 2 min ago, no card                    → still-unknown
 *   c12 pending; has a card WITHOUT the campaign marker   → no issue (card ignored)
 *   c60 not on the list, has a campaign card              → not-in-selection
 */
async function scenario() {
  env = testConfig();
  const { config } = env;
  const paths = campaignPaths(config);
  const list = customers(12);
  const selection = await selectionFixture(config, { customers: list });
  fake = installFakeShopify({
    customers: shopifyCustomers(list, [1, 2, 4, 5, 6, 8], config.sentTag),
    giftCards: [
      card(501, 1, { cents: 1533 }),
      card(502, 2),
      card(503, 3, { cents: 1977 }),
      card(504, 4),
      card(514, 4, { createdAt: '2026-10-05T16:30:00Z' }),
      card(505, 5, { cents: 1533 }),
      card(506, 6, { enabled: false }),
      card(507, 7),
      card(509, 9, { createdAt: minutesBefore(RUN_AT, 20) }),
      card(560, 60),
      card(570, 12, { note: 'some other promotion' }),
    ],
  });
  fake.state.hiddenCardIds.add(gid('GiftCard', 502));
  writeJournal(paths, [
    ...issued(1, 501, 1533),
    ...issued(2, 502, 1077),
    ...issued(4, 504, 1077),
    ...issued(5, 505, 1077),
    ...issued(6, 506, 1077),
    ...issued(7, 507, 1077),
    ...lost(9, minutesBefore(RUN_AT, 20)),
    ...lost(10, minutesBefore(RUN_AT, 20)),
    ...lost(11, minutesBefore(RUN_AT, 2), { unknown: false }),
  ]);
  return { config, paths, selection };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

test('verify finds every issue type, fixes the journal and writes verify.json / tags.json in the spec shape', async () => {
  const { config, paths, selection } = await scenario();
  const log = memoryLog();
  const writeReport = reportStub();

  const { exitCode, result } = await runVerify({ config, log, now: at(RUN_AT), sleep: noSleep, writeReport });

  assert.equal(exitCode, 0, log.lines.join('\n'));
  assert.deepEqual(result.issues.map((i) => [i.type, tail(i.customerId), tail(i.giftCardId)]), [
    ['still-unknown', '11', null],
    ['duplicate-cards', '4', '514'],
    ['amount-mismatch', '5', '505'],
    ['missing-in-shopify', '2', '502'],
    ['not-in-selection', '60', '560'],
    ['tag-missing', '7', '507'],
    ['tag-without-card', '8', null],
    ['card-disabled', '6', '506'],
    ['not-in-journal', '3', '503'],
    ['resolved-unknown', '9', '509'],
    ['resolved-unknown', '10', null],
  ]);

  // verify.json: exactly the documented shape.
  const report = readJson(paths.verify);
  assert.deepEqual(Object.keys(report), ['version', 'verifiedAt', 'cardCount', 'taggedCount', 'counts', 'issues']);
  assert.equal(report.version, 1);
  assert.equal(report.verifiedAt, RUN_AT);
  assert.equal(report.cardCount, 9, 'hidden card and the card without the marker are not counted');
  assert.equal(report.taggedCount, 6);
  assert.deepEqual(report.counts, {
    'still-unknown': 1,
    'duplicate-cards': 1,
    'amount-mismatch': 1,
    'missing-in-shopify': 1,
    'not-in-selection': 1,
    'tag-missing': 1,
    'tag-without-card': 1,
    'card-disabled': 1,
    'not-in-journal': 1,
    'resolved-unknown': 2,
  });
  assert.deepEqual(Object.keys(report.counts), ISSUE_TYPES.map((t) => t.type));
  assert.deepEqual(report, result);
  for (const issue of report.issues) {
    assert.deepEqual(Object.keys(issue), ['type', 'customerId', 'giftCardId', 'journal', 'shopify', 'action', 'journalEn', 'shopifyEn', 'actionEn']);
    for (const key of ['journal', 'shopify', 'action', 'journalEn', 'shopifyEn', 'actionEn']) {
      assert.equal(typeof issue[key], 'string');
      assert.ok(issue[key].length > 0, `${issue.type}.${key} is empty`);
    }
    // The English texts (for the English workbook) hold no CJK or full-width characters.
    for (const key of ['journalEn', 'shopifyEn', 'actionEn']) {
      assert.doesNotMatch(issue[key], CJK, `${issue.type}.${key} contains Chinese: ${issue[key]}`);
    }
  }
  const byType = (type, n) => report.issues.find((i) => i.type === type && tail(i.customerId) === String(n));
  assert.equal(byType('not-in-journal', 3).action, '已按 Shopify 补记到本地日志');
  assert.equal(byType('not-in-journal', 3).shopify, `有本活动的卡：尾号 x503，原金额 $19.77，2026-10-05 09:00 建卡；客户还没有 ${config.sentTag}，下次运行 issue 会补打`);
  assert.equal(byType('resolved-unknown', 9).action, '已按 Shopify 补记到本地日志');
  assert.match(byType('resolved-unknown', 9).shopify, /^找到了这张卡：尾号 x509/);
  assert.match(byType('resolved-unknown', 10).shopify, /超过 10 分钟仍没找到/);
  assert.match(byType('resolved-unknown', 10).action, /改回待发放/);
  assert.match(byType('still-unknown', 11).action, /约 8 分钟后再运行 verify 或 issue/);
  assert.match(byType('missing-in-shopify', 2).shopify, /搜索索引/);
  assert.match(byType('duplicate-cards', 4).action, /停用这张多出来的卡（保留尾号 x504 那张）/);
  assert.match(byType('amount-mismatch', 5).journal, /\$10\.77/);
  assert.match(byType('amount-mismatch', 5).shopify, /\$15\.33/);
  assert.match(byType('tag-missing', 7).shopify, new RegExp(config.sentTag));
  assert.match(byType('tag-without-card', 8).shopify, /没有本活动的卡/);
  assert.match(byType('not-in-selection', 60).shopify, /不在发放名单里/);
  assert.equal(byType('card-disabled', 6).shopify, '这张卡已停用：尾号 x506，原金额 $10.77，2026-10-05 09:00 建卡');
  assert.equal(byType('card-disabled', 6).action, '如果不是有意停用，请在后台重新启用这张卡');
  // Times in the texts are store-local (America/Los_Angeles): 16:00Z = 09:00.
  assert.match(byType('missing-in-shopify', 2).journal, /2026-10-05 09:00/);

  // The same facts in English (journalEn / shopifyEn / actionEn), same times and amounts.
  assert.equal(byType('not-in-journal', 3).actionEn, 'Recorded in the local journal from Shopify');
  assert.equal(byType('not-in-journal', 3).shopifyEn, `Has a card from this campaign: last 4 x503, initial $19.77, created 2026-10-05 09:00 Los Angeles time; the customer does not carry ${config.sentTag} yet, the next issue run will add it`);
  assert.equal(byType('resolved-unknown', 9).actionEn, 'Recorded in the local journal from Shopify');
  assert.match(byType('resolved-unknown', 9).shopifyEn, /^Found the card: last 4 x509/);
  assert.match(byType('resolved-unknown', 10).shopifyEn, /more than 10 minutes after creation started; confirmed not created/);
  assert.match(byType('resolved-unknown', 10).actionEn, /^Set back to Pending in the local journal/);
  assert.match(byType('resolved-unknown', 10).journalEn, /^Needs review: card creation started 2026-10-05 \d\d:\d\d Los Angeles time, outcome unknown$/);
  assert.match(byType('still-unknown', 11).actionEn, /Run verify or issue again in about 8 minutes/);
  assert.match(byType('still-unknown', 11).journalEn, /^In progress: card creation started /);
  assert.match(byType('missing-in-shopify', 2).shopifyEn, /search index/);
  assert.equal(byType('missing-in-shopify', 2).journalEn, 'Journal has this card: last 4 x502, created 2026-10-05 09:00 Los Angeles time');
  assert.equal(byType('missing-in-shopify', 2).actionEn, 'Run verify again in a few minutes; if still missing, look up the gift card by id in the admin');
  assert.equal(byType('duplicate-cards', 4).actionEn, 'Disable this extra card in the admin after checking (keep the one ending x504)');
  assert.match(byType('duplicate-cards', 4).shopifyEn, /^This customer has 2 cards from this campaign; this one is extra: last 4 x514/);
  assert.equal(byType('amount-mismatch', 5).journalEn, 'List amount $10.77');
  assert.match(byType('amount-mismatch', 5).shopifyEn, /^Card initial amount \$15\.33, last 4 x505$/);
  assert.match(byType('tag-missing', 7).shopifyEn, new RegExp(`^Customer does not carry ${config.sentTag} `));
  assert.match(byType('tag-missing', 7).journalEn, /^Journal says tagged 2026-10-05 09:00 Los Angeles time$/);
  assert.match(byType('tag-without-card', 8).shopifyEn, /but has no card from this campaign$/);
  assert.match(byType('not-in-selection', 60).shopifyEn, /is not on the recipient list$/);
  assert.equal(byType('card-disabled', 6).shopifyEn, 'This card is disabled: last 4 x506, initial $10.77, created 2026-10-05 09:00 Los Angeles time');
  assert.equal(byType('card-disabled', 6).actionEn, 'If it was not disabled on purpose, re-enable the card in the admin');
  // Journal states in English: a done row, an empty journal.
  assert.match(byType('duplicate-cards', 4).journalEn, /^The journal holds the one ending x504$/);
  assert.equal(byType('not-in-journal', 3).journalEn, 'No journal record for this customer');

  // tags.json: everyone carrying the tag in Shopify, sorted.
  assert.deepEqual(readJson(paths.tags), {
    fetchedAt: RUN_AT,
    tag: config.sentTag,
    ids: [1, 2, 4, 5, 6, 8].map((n) => gid('Customer', n)),
  });

  // Journal: run.start/run.end around exactly three fix-ups, all tagged with this run.
  const entries = readJournal(paths.journal);
  const start = entries.find((e) => e.op === 'run.start' && e.command === 'verify');
  assert.ok(start);
  assert.equal(start.dryRun, false);
  const mine = entries.filter((e) => e.run === start.run);
  assert.deepEqual(mine.map((e) => [e.op, tail(e.cid) ?? null]), [
    ['run.start', null],
    ['reconcile.found', '3'],
    ['reconcile.found', '9'],
    ['reconcile.none', '10'],
    ['run.end', null],
  ]);
  const found3 = mine[1];
  assert.deepEqual(
    { giftCardId: found3.giftCardId, last4: found3.last4, amountCents: found3.amountCents, createdAt: found3.createdAt, source: found3.source },
    { giftCardId: gid('GiftCard', 503), last4: 'x503', amountCents: 1977, createdAt: '2026-10-05T16:00:00Z', source: 'verify' },
  );
  assert.equal(mine[2].source, 'verify');
  const end = mine.at(-1);
  assert.equal(end.exitCode, 0);
  assert.equal(end.summary.issueCount, 11);
  assert.equal(end.summary.fixedCount, 3);

  const folded = foldJournal(entries).customers;
  assert.equal(folded.get(gid('Customer', 3)).status, STATUS.CREATED);
  assert.equal(folded.get(gid('Customer', 3)).giftCardId, gid('GiftCard', 503));
  assert.equal(folded.get(gid('Customer', 3)).reconciledFrom, 'verify');
  assert.equal(folded.get(gid('Customer', 9)).status, STATUS.CREATED);
  assert.equal(folded.get(gid('Customer', 9)).giftCardId, gid('GiftCard', 509));
  assert.equal(folded.get(gid('Customer', 10)).status, STATUS.PENDING);
  assert.equal(folded.get(gid('Customer', 11)).status, STATUS.IN_PROGRESS, 'too early to decide');

  // Read-only towards Shopify.
  assert.deepEqual([...new Set(fake.state.calls.ops.map((o) => o.op))].sort(), ['FindCampaignCards', 'TaggedCustomers']);
  for (const op of WRITE_OPS) assert.equal(fake.opsNamed(op).length, 0, op);
  assert.equal(fake.state.calls.create.length + fake.state.calls.tag.length + fake.state.calls.notify.length, 0);
  // Search window: one hour before the list was made.
  assert.equal(selection.createdAt, NOW_ISO);
  assert.match(fake.opsNamed('FindCampaignCards')[0].variables.query, /created_at:>='2026-10-01T19:00:00\.000Z'/);

  // Workbook regenerated once; lock released.
  assert.equal(writeReport.calls.length, 1);
  assert.equal(writeReport.calls[0].config, config);
  assert.deepEqual(writeReport.calls[0].paths, paths);
  assert.equal(fs.existsSync(paths.runLock), false);

  // Console summary in Chinese with counts per type.
  const text = log.lines.join('\n');
  assert.match(text, /核对完成：Shopify 上本活动的卡 9 张，带 gift-card-sent-2026-10 的客户 6 个，名单 12 人/);
  assert.match(text, /发现 11 条：8 条需要人工看，3 条已按 Shopify 补记/);
  assert.match(text, /同一客户多张卡：1/);
  assert.match(text, /结果不明已查清（已补记）：2/);
  assert.match(text, /本地日志补记了 3 条/);
  assert.doesNotMatch(text, /fake-token/);
});

test('a second verify run changes nothing more: fixed rows are gone, open issues stay', async () => {
  const { config, paths } = await scenario();
  await runVerify({ config, log: memoryLog(), now: at(RUN_AT), sleep: noSleep, writeReport: reportStub() });
  const before = readJournal(paths.journal).filter((e) => e.op.startsWith('reconcile.')).length;

  const { exitCode, result } = await runVerify({ config, log: memoryLog(), now: at(minutesAfter(RUN_AT, 1)), sleep: noSleep, writeReport: reportStub() });

  assert.equal(exitCode, 0);
  assert.equal(readJournal(paths.journal).filter((e) => e.op.startsWith('reconcile.')).length, before, 'no new fix-ups');
  assert.deepEqual(result.issues.map((i) => [i.type, tail(i.customerId)]), [
    ['still-unknown', '11'],
    ['duplicate-cards', '4'],
    ['amount-mismatch', '5'],
    ['missing-in-shopify', '2'],
    ['not-in-selection', '60'],
    ['tag-missing', '7'],
    ['tag-without-card', '8'],
    ['card-disabled', '6'],
  ]);
  assert.equal(result.counts['not-in-journal'], 0);
  assert.equal(result.counts['resolved-unknown'], 0);
});

test('an unknown row stays still-unknown until the settle time, then reconcile.none puts it back to pending', async () => {
  env = testConfig();
  const { config } = env;
  const paths = campaignPaths(config);
  const list = customers(2);
  await selectionFixture(config, { customers: list });
  fake = installFakeShopify({ customers: list });
  const started = '2026-10-05T16:00:00.000Z';
  writeJournal(paths, lost(1, started));

  const early = await runVerify({ config, log: memoryLog(), now: at(minutesAfter(started, 5)), sleep: noSleep, writeReport: reportStub() });
  assert.equal(early.exitCode, 0);
  assert.deepEqual(early.result.issues.map((i) => i.type), ['still-unknown']);
  assert.match(early.result.issues[0].action, /约 5 分钟后/);
  assert.equal(readJournal(paths.journal).filter((e) => e.op.startsWith('reconcile.')).length, 0, 'nothing decided yet');
  assert.equal(foldJournal(readJournal(paths.journal)).customers.get(gid('Customer', 1)).status, STATUS.UNKNOWN);

  const settled = await runVerify({ config, log: memoryLog(), now: at(minutesAfter(started, 10)), sleep: noSleep, writeReport: reportStub() });
  assert.deepEqual(settled.result.issues.map((i) => i.type), ['resolved-unknown']);
  const none = readJournal(paths.journal).filter((e) => e.op === 'reconcile.none');
  assert.equal(none.length, 1);
  assert.equal(none[0].cid, gid('Customer', 1));
  assert.match(none[0].note, /verify/);
  assert.ok(none[0].run, 'carries the verify run id');
  const s = foldJournal(readJournal(paths.journal)).customers.get(gid('Customer', 1));
  assert.equal(s.status, STATUS.PENDING);
  assert.equal(s.giftCardId, null);
});

test('unknownSettleMs is configurable', async () => {
  env = testConfig();
  const { config } = env;
  const paths = campaignPaths(config);
  const list = customers(1);
  await selectionFixture(config, { customers: list });
  fake = installFakeShopify({ customers: list });
  writeJournal(paths, lost(1, '2026-10-05T16:00:00.000Z', { unknown: false }));

  const { result } = await runVerify({ config, log: memoryLog(), now: at('2026-10-05T16:01:00.000Z'), sleep: noSleep, writeReport: reportStub(), unknownSettleMs: 30_000 });
  assert.deepEqual(result.issues.map((i) => i.type), ['resolved-unknown']);
  assert.match(result.issues[0].shopify, /超过 30 秒/);
  assert.equal(foldJournal(readJournal(paths.journal)).customers.get(gid('Customer', 1)).status, STATUS.PENDING);
});

// The action of a row verify puts back to pending (reconcile.none).
const SETTLED_RETRY = '已在本地日志改回待发放，可以用 issue 重试';
const SETTLED_NO_NEW_CARDS = (date) => `已在本地日志改回待发放；礼品卡到期日（${date}）已到，issue 不再建新卡`;

test('from the cards\' expiry date on (store date) a row put back to pending does not promise an issue retry', async () => {
  env = testConfig(); // the list freezes GIFT_CARD_EXPIRES_ON=2026-10-19
  const { config } = env;
  const paths = campaignPaths(config);
  const list = customers(3);
  await selectionFixture(config, { customers: list });
  fake = installFakeShopify({ customers: list });
  const crashed = '2026-10-05T18:00:00.000Z'; // creates whose answer never came; no card was made

  // 10/18 23:59 in Los Angeles (already 10/19 in UTC): issue still creates cards today.
  writeJournal(paths, lost(1, crashed));
  const eve = await runVerify({ config, log: memoryLog(), now: at('2026-10-19T06:59:00.000Z'), sleep: noSleep, writeReport: reportStub() });
  assert.equal(eve.exitCode, 0);
  assert.deepEqual(eve.result.issues.map((i) => [i.type, tail(i.customerId), i.action]), [['resolved-unknown', '1', SETTLED_RETRY]]);

  // 10/19 00:00 in Los Angeles: from now on issue only repairs.
  writeJournal(paths, lost(2, crashed));
  const day = await runVerify({ config, log: memoryLog(), now: at('2026-10-19T07:00:00.000Z'), sleep: noSleep, writeReport: reportStub() });
  assert.deepEqual(day.result.issues.map((i) => [i.type, tail(i.customerId), i.action]), [['resolved-unknown', '2', SETTLED_NO_NEW_CARDS('2026-10-19')]]);
  assert.equal(readJson(paths.verify).issues[0].action, SETTLED_NO_NEW_CARDS('2026-10-19'));
  assert.equal(readJson(paths.verify).issues[0].actionEn, 'Set back to Pending in the local journal; the cards\' expiry date (2026-10-19) has been reached, so issue creates no new cards');

  // 10/25: same text; the journal is still put back to pending (what Shopify proves does not change).
  writeJournal(paths, lost(3, crashed));
  const late = await runVerify({ config, log: memoryLog(), now: at('2026-10-25T17:00:00.000Z'), sleep: noSleep, writeReport: reportStub() });
  assert.equal(late.result.issues[0].action, SETTLED_NO_NEW_CARDS('2026-10-19'));
  assert.doesNotMatch(late.result.issues[0].action, /可以用 issue 重试/);
  assert.doesNotMatch(late.result.issues[0].action, /REMIND/);
  assert.deepEqual(readJournal(paths.journal).filter((e) => e.op === 'reconcile.none').map((e) => tail(e.cid)), ['1', '2', '3']);
  const folded = foldJournal(readJournal(paths.journal)).customers;
  for (const n of [1, 2, 3]) assert.equal(folded.get(gid('Customer', n)).status, STATUS.PENDING, `customer ${n}`);
});

test('the no-new-cards date is the expiry frozen in the list (what issue checks), not today\'s .env', async () => {
  env = testConfig(); // the list was made with GIFT_CARD_EXPIRES_ON=2026-10-19
  const list = customers(2);
  await selectionFixture(env.config, { customers: list });
  const moved = testConfig({ CAMPAIGNS_DIR: env.dir, GIFT_CARD_EXPIRES_ON: '2026-10-30' }); // .env changed afterwards
  try {
    const paths = campaignPaths(moved.config);
    fake = installFakeShopify({ customers: list });
    writeJournal(paths, lost(1, '2026-10-05T18:00:00.000Z'));
    const first = await runVerify({ config: moved.config, log: memoryLog(), now: at('2026-10-18T17:00:00.000Z'), sleep: noSleep, writeReport: reportStub() });
    assert.equal(first.result.issues[0].action, SETTLED_RETRY, 'on 10/18 issue still creates cards');
    writeJournal(paths, lost(2, '2026-10-05T18:00:00.000Z'));
    const second = await runVerify({ config: moved.config, log: memoryLog(), now: at('2026-10-19T17:00:00.000Z'), sleep: noSleep, writeReport: reportStub() });
    assert.equal(second.result.issues[0].action, SETTLED_NO_NEW_CARDS('2026-10-19'), 'the cards were made with 10/19, whatever .env says now');
  } finally {
    moved.cleanup();
  }
});

test('a test campaign is bound to the expiry date too; the old reminder dates mean nothing', async () => {
  env = testConfig({ CAMPAIGN_ID: '2026-10-test', SENT_TAG: 'OCT26RTPROMO-TEST', TEST_CUSTOMER_IDS: '1,2' });
  const { config } = env;
  const paths = campaignPaths(config);
  const list = customers(2);
  const selection = buildTestSelection({
    config,
    customers: list.map(parseCustomer),
    timezone: 'America/Los_Angeles',
    createdAt: '2026-10-03T18:00:00.000Z',
    snapshot: { exportedAt: '2026-10-03T18:00:00.000Z', source: 'nodes', count: 2 },
  });
  writeJsonAtomic(paths.selection, selection);
  fake = installFakeShopify({ customers: list });
  writeJournal(paths, lost(1, '2026-10-05T18:00:00.000Z'));
  const before = await runVerify({ config, log: memoryLog(), now: at('2026-10-16T17:00:00.000Z'), sleep: noSleep, writeReport: reportStub() });
  assert.equal(before.exitCode, 0);
  assert.equal(selection.mode, 'test');
  assert.deepEqual(before.result.issues.map((i) => [i.type, tail(i.customerId), i.action]), [['resolved-unknown', '1', SETTLED_RETRY]]);
  writeJournal(paths, lost(2, '2026-10-05T18:00:00.000Z'));
  const after = await runVerify({ config, log: memoryLog(), now: at('2026-10-20T17:00:00.000Z'), sleep: noSleep, writeReport: reportStub() });
  assert.deepEqual(after.result.issues.map((i) => [i.type, tail(i.customerId), i.action]), [['resolved-unknown', '2', SETTLED_NO_NEW_CARDS('2026-10-19')]]);
});

test('analyzeVerify: the no-new-cards text needs an expiry date in the list and a store date on or after it', () => {
  const recipients = [{ seq: 1, customerId: gid('Customer', 1), amountCents: 1077 }];
  const unknown = { status: STATUS.UNKNOWN, startedAt: '2026-10-05T18:00:00.000Z', giftCardId: null, taggedAt: null, attempts: 1 };
  const action = ({ mode = 'live', giftCardExpiresOn = '2026-10-19', nowIso, timezone = 'America/Los_Angeles' }) => {
    const { issues, fixes } = analyzeVerify({
      selection: { mode, params: { currency: 'USD', giftCardExpiresOn }, recipients },
      state: { customers: new Map([[gid('Customer', 1), unknown]]) },
      cards: [],
      taggedIds: new Set(),
      nowMs: Date.parse(nowIso),
      tag: 'OCT26RTPROMO',
      timezone,
    });
    assert.deepEqual(fixes.map((f) => f.op), ['reconcile.none'], 'the journal fix is the same either way');
    assert.deepEqual(issues.map((i) => i.type), ['resolved-unknown']);
    return issues[0].action;
  };
  assert.equal(action({ nowIso: '2026-10-11T23:00:00Z' }), SETTLED_RETRY);
  assert.equal(action({ nowIso: '2026-10-12T17:00:00Z' }), SETTLED_RETRY, 'the old REMIND_1_DATE is just another day');
  assert.equal(action({ nowIso: '2026-10-19T06:59:59Z' }), SETTLED_RETRY, '10/18 23:59:59 in Los Angeles');
  assert.equal(action({ nowIso: '2026-10-19T07:00:00Z' }), SETTLED_NO_NEW_CARDS('2026-10-19'), '10/19 00:00 in Los Angeles');
  assert.equal(action({ nowIso: '2026-10-25T17:00:00Z' }), SETTLED_NO_NEW_CARDS('2026-10-19'));
  assert.equal(action({ nowIso: '2026-10-19T05:00:00Z', timezone: 'UTC' }), SETTLED_NO_NEW_CARDS('2026-10-19'), 'the store\'s own calendar decides');
  assert.equal(action({ mode: 'test', nowIso: '2026-10-25T17:00:00Z' }), SETTLED_NO_NEW_CARDS('2026-10-19'), 'test campaigns too');
  assert.equal(action({ giftCardExpiresOn: '', nowIso: '2026-12-20T17:00:00Z' }), SETTLED_RETRY, 'cards that never expire');
  assert.equal(action({ giftCardExpiresOn: null, nowIso: '2026-12-20T17:00:00Z' }), SETTLED_RETRY, 'no expiry in the list');
  // Older callers (no mode, no expiry in params, a remind1Date that is now ignored): unchanged.
  const { issues } = analyzeVerify({
    selection: { params: { currency: 'USD' }, recipients },
    state: { customers: new Map([[gid('Customer', 1), unknown]]) },
    cards: [],
    taggedIds: new Set(),
    nowMs: Date.parse('2026-10-20T17:00:00Z'),
    tag: 'OCT26RTPROMO',
    remind1Date: '2026-10-12',
  });
  assert.equal(issues[0].action, SETTLED_RETRY);
});

test('a card that only shows up after search-index lag resolves the unknown row with reconcile.found', async () => {
  env = testConfig();
  const { config } = env;
  const paths = campaignPaths(config);
  const list = customers(1);
  await selectionFixture(config, { customers: list });
  const started = '2026-10-05T16:00:00.000Z';
  fake = installFakeShopify({ customers: list, giftCards: [card(801, 1, { cents: 1533, createdAt: started })] });
  fake.state.hiddenCardIds.add(gid('GiftCard', 801));
  writeJournal(paths, lost(1, started));

  const first = await runVerify({ config, log: memoryLog(), now: at(minutesAfter(started, 2)), sleep: noSleep, writeReport: reportStub() });
  assert.deepEqual(first.result.issues.map((i) => i.type), ['still-unknown']);

  fake.state.hiddenCardIds.clear(); // the index caught up
  const second = await runVerify({ config, log: memoryLog(), now: at(minutesAfter(started, 3)), sleep: noSleep, writeReport: reportStub() });
  assert.deepEqual(second.result.issues.map((i) => [i.type, tail(i.giftCardId)]), [['resolved-unknown', '801']]);
  const s = foldJournal(readJournal(paths.journal)).customers.get(gid('Customer', 1));
  assert.equal(s.status, STATUS.CREATED);
  assert.equal(s.giftCardId, gid('GiftCard', 801));
  assert.equal(s.last4, 'x801');
  assert.equal(s.reconciledFrom, 'verify');
});

test('duplicates: the oldest card is recorded when the journal has none; actions depend on the extra card', async () => {
  env = testConfig();
  const { config } = env;
  const paths = campaignPaths(config);
  const list = customers(4);
  await selectionFixture(config, { customers: list });
  fake = installFakeShopify({
    customers: shopifyCustomers(list, [2, 3, 4], config.sentTag),
    giftCards: [
      // c1: no journal record, two unused cards → keep the oldest (601), the other is extra.
      card(602, 1, { cents: 1533, createdAt: '2026-10-05T16:05:00Z' }),
      card(601, 1, { cents: 1533, createdAt: '2026-10-05T16:00:00Z' }),
      // c2: journal card 611 plus an extra that was already used.
      card(611, 2),
      card(612, 2, { balance: 500, createdAt: '2026-10-05T16:10:00Z' }),
      // c3: journal card 621 plus an extra that was deactivated.
      card(621, 3, { cents: 1977 }),
      card(622, 3, { cents: 1977, enabled: false, createdAt: '2026-10-05T16:10:00Z' }),
      // c4: journal card 631 is not visible yet, Shopify shows another one.
      card(631, 4),
      card(632, 4, { createdAt: '2026-10-05T16:10:00Z' }),
    ],
  });
  fake.state.hiddenCardIds.add(gid('GiftCard', 631));
  writeJournal(paths, [...issued(2, 611, 1077), ...issued(3, 621, 1977), ...issued(4, 631, 1077)]);

  const { exitCode, result } = await runVerify({ config, log: memoryLog(), now: at(RUN_AT), sleep: noSleep, writeReport: reportStub() });

  assert.equal(exitCode, 0);
  const rows = result.issues.map((i) => [i.type, tail(i.customerId), tail(i.giftCardId)]);
  assert.deepEqual(rows, [
    ['duplicate-cards', '1', '602'],
    ['duplicate-cards', '2', '612'],
    ['duplicate-cards', '3', '622'],
    ['duplicate-cards', '4', '632'],
    ['missing-in-shopify', '4', '631'],
    ['card-disabled', '3', '622'],
    ['not-in-journal', '1', '601'],
  ]);
  const dup = (n) => result.issues.find((i) => i.type === 'duplicate-cards' && tail(i.customerId) === String(n));
  assert.match(dup(1).journal, /补记了尾号 x601/);
  assert.match(dup(1).action, /停用这张多出来的卡（保留尾号 x601 那张）/);
  assert.match(dup(1).shopify, /一共有 2 张/);
  assert.match(dup(2).action, /已经被用过/);
  assert.match(dup(2).shopify, /余额 \$5\.00/);
  assert.equal(dup(3).action, '已停用，不用再处理');
  assert.match(dup(4).journal, /尾号 x631 那张（这次在 Shopify 上没找到）/);
  assert.equal(result.issues.find((i) => i.type === 'card-disabled').action, '这是多出来的卡，已停用，不用再处理');

  const found = readJournal(paths.journal).filter((e) => e.op === 'reconcile.found');
  assert.deepEqual(found.map((e) => [tail(e.cid), tail(e.giftCardId)]), [['1', '601']]);
});

test('a card for a customer without any account link, and a mismatched currency, are reported', () => {
  const selection = {
    params: { currency: 'USD' },
    recipients: [{ seq: 1, customerId: gid('Customer', 1), amountCents: 1077 }],
  };
  const base = { createdAt: '2026-10-05T16:00:00Z', enabled: true, last4: 'x001', balanceCents: 1077, amountCents: 1077 };
  const { issues, fixes, counts } = analyzeVerify({
    selection,
    state: { customers: new Map() },
    cards: [
      { ...base, id: gid('GiftCard', 1), customerId: gid('Customer', 1), currencyCode: 'CAD' },
      { ...base, id: gid('GiftCard', 2), customerId: null, currencyCode: 'USD' },
    ],
    taggedIds: new Set(),
    nowMs: Date.parse(RUN_AT),
    tag: 'OCT26RTPROMO',
    timezone: 'America/Los_Angeles',
  });
  assert.deepEqual(issues.map((i) => [i.type, tail(i.customerId), tail(i.giftCardId)]), [
    ['amount-mismatch', '1', '1'],
    ['not-in-selection', null, '2'],
    ['not-in-journal', '1', '1'],
  ]);
  assert.equal(issues[0].journal, '名单金额 $10.77 USD');
  assert.match(issues[0].shopify, /^卡的原金额 \$10\.77 CAD，币种不一致/);
  assert.match(issues[1].shopify, /没有关联客户/);
  assert.equal(issues[1].journal, '日志里没有这张卡');
  assert.equal(fixes.length, 1);
  assert.equal(counts['not-in-selection'], 1);
  // English twins of the same rows.
  assert.equal(issues[0].journalEn, 'List amount $10.77 USD');
  assert.equal(issues[0].shopifyEn, 'Card initial amount $10.77 CAD, currency differs, last 4 x001');
  assert.equal(issues[1].shopifyEn, 'This card carries the campaign marker but has no customer: last 4 x001, initial $10.77, created 2026-10-05 09:00 Los Angeles time');
  assert.equal(issues[1].journalEn, 'No journal record for this card');
  assert.equal(issues[1].actionEn, 'Check how this card was created; disable it in the admin if it should not have been issued');
  for (const issue of issues) for (const key of ['journalEn', 'shopifyEn', 'actionEn']) assert.doesNotMatch(issue[key], CJK, `${issue.type}.${key}: ${issue[key]}`);
});

test('analyzeVerify: every English text variant (journal states, durations, extra cards, no-new-cards) is ASCII-only', () => {
  const C = (n) => gid('Customer', n);
  const G = (n) => gid('GiftCard', n);
  const selection = {
    mode: 'live',
    params: { currency: 'USD', giftCardExpiresOn: '2026-10-19' },
    recipients: Array.from({ length: 9 }, (_, i) => ({ seq: i + 1, customerId: C(i + 1), amountCents: 1077 })),
  };
  const started = '2026-10-05T16:00:00.000Z'; // 09:00 Los Angeles
  const cardBase = { createdAt: started, enabled: true, amountCents: 1077, balanceCents: 1077, currencyCode: 'USD' };
  const state = {
    customers: new Map([
      // c1: pending after attempts, Shopify has a card → not-in-journal, "Pending: tried 2 times before"
      [C(1), { status: STATUS.PENDING, attempts: 2, giftCardId: null, taggedAt: null }],
      // c2: failed with our own English error, carries the tag in Shopify, no card → tag-without-card
      [C(2), { status: STATUS.FAILED, error: 'Network error calling Shopify: fetch failed', giftCardId: null, taggedAt: null }],
      // c3: skipped (reason code), carries the tag, no card → tag-without-card with a skip label
      [C(3), { status: STATUS.SKIPPED, skipReason: 'relay-email', giftCardId: null, taggedAt: null }],
      // c4: created (not tagged); journal card hidden, Shopify shows another → duplicate (journal keeper missing) + missing-in-shopify
      [C(4), { status: STATUS.CREATED, giftCardId: G(41), last4: 'x041', createdAt: started, taggedAt: null }],
      // c5: done but the tag is gone in Shopify → tag-missing; a used extra card and a disabled extra card
      [C(5), { status: STATUS.DONE, giftCardId: G(51), last4: 'x051', createdAt: started, taggedAt: started }],
      // c6: unknown, settled on/after the cards' expiry date → "no new cards" action
      [C(6), { status: STATUS.UNKNOWN, startedAt: started, giftCardId: null, taggedAt: null, attempts: 1 }],
      // c7: in progress 90 seconds ago → still-unknown "in about 9 minutes" (seconds rounded up)
      [C(7), { status: STATUS.IN_PROGRESS, startedAt: '2026-10-19T16:58:30.000Z', giftCardId: null, taggedAt: null }],
      // c8: skipped with an unknown reason code
      [C(8), { status: STATUS.SKIPPED, skipReason: 'something-new', giftCardId: null, taggedAt: null }],
    ]),
  };
  const cards = [
    { ...cardBase, id: G(11), customerId: C(1), last4: 'x011' },
    { ...cardBase, id: G(42), customerId: C(4), last4: 'x042', createdAt: '2026-10-05T16:10:00.000Z' },
    { ...cardBase, id: G(51), customerId: C(5), last4: 'x051' },
    { ...cardBase, id: G(52), customerId: C(5), last4: 'x052', balanceCents: 500, createdAt: '2026-10-05T16:10:00.000Z' },
    { ...cardBase, id: G(53), customerId: C(5), last4: 'x053', enabled: false, createdAt: '2026-10-05T16:20:00.000Z' },
    // c8's card is disabled and is the only one → "re-enable" action
    { ...cardBase, id: G(81), customerId: C(8), last4: 'x081', enabled: false },
  ];
  const { issues } = analyzeVerify({
    selection,
    state,
    cards,
    taggedIds: new Set([C(2), C(3)]),
    nowMs: Date.parse('2026-10-19T17:00:00.000Z'), // 10/19 10:00 Los Angeles, the expiry date reached
    tag: 'OCT26RTPROMO',
    timezone: 'America/Los_Angeles',
  });

  const find = (type, n, cardN = undefined) => {
    const hit = issues.find((i) => i.type === type && tail(i.customerId) === String(n) && (cardN === undefined || tail(i.giftCardId) === String(cardN)));
    assert.ok(hit, `${type} for customer ${n} (card ${cardN ?? 'any'}) in ${JSON.stringify(issues.map((i) => [i.type, tail(i.customerId), tail(i.giftCardId)]))}`);
    return hit;
  };
  assert.equal(find('not-in-journal', 1).journalEn, 'Pending: tried 2 times before, confirmed not created');
  assert.equal(find('not-in-journal', 1).shopifyEn, 'Has a card from this campaign: last 4 x011, initial $10.77, created 2026-10-05 09:00 Los Angeles time; the customer does not carry OCT26RTPROMO yet, the next issue run will add it');
  assert.equal(find('tag-without-card', 2).journalEn, 'Failed: Network error calling Shopify: fetch failed');
  assert.equal(find('tag-missing', 5).journalEn, 'Journal says tagged 2026-10-05 09:00 Los Angeles time');
  assert.equal(find('tag-missing', 5).shopifyEn, 'Customer does not carry OCT26RTPROMO (a new tag may not be in the search index yet)');
  assert.equal(find('tag-missing', 5).actionEn, 'Run verify again in a few minutes; if still missing, add OCT26RTPROMO to the customer in the admin');
  assert.equal(find('tag-without-card', 3).journalEn, 'Skipped before issuing: email domain is on the exclusion list');
  assert.equal(find('tag-without-card', 3).shopifyEn, 'Customer carries OCT26RTPROMO but has no card from this campaign');
  assert.equal(find('tag-without-card', 3).actionEn, 'Check who added the tag: with this tag, issue skips this customer');
  assert.equal(find('duplicate-cards', 4, 42).journalEn, 'The journal holds the one ending x041 (not found in Shopify this time)');
  assert.equal(find('missing-in-shopify', 4).journalEn, 'Journal has this card: last 4 x041, created 2026-10-05 09:00 Los Angeles time');
  assert.equal(find('missing-in-shopify', 4).shopifyEn, 'Card not found (a new card may not be in the search index yet, or its note was changed)');
  assert.equal(find('duplicate-cards', 5, 52).journalEn, 'The journal holds the one ending x051');
  assert.equal(find('duplicate-cards', 5, 52).shopifyEn, 'This customer has 3 cards from this campaign; this one is extra: last 4 x052, initial $10.77, balance $5.00, created 2026-10-05 09:10 Los Angeles time');
  assert.equal(find('duplicate-cards', 5, 52).actionEn, 'This extra card has already been used; decide manually how to handle it');
  assert.equal(find('duplicate-cards', 5, 53).shopifyEn, 'This customer has 3 cards from this campaign; this one is extra: last 4 x053, initial $10.77, created 2026-10-05 09:20 Los Angeles time, disabled');
  assert.equal(find('duplicate-cards', 5, 53).actionEn, 'Already disabled, nothing to do');
  assert.equal(find('card-disabled', 5, 53).journalEn, 'Done: last 4 x051, tagged 2026-10-05 09:00 Los Angeles time');
  assert.equal(find('card-disabled', 5, 53).shopifyEn, 'This card is disabled: last 4 x053, initial $10.77, created 2026-10-05 09:20 Los Angeles time');
  assert.equal(find('card-disabled', 5, 53).actionEn, 'This is an extra card and already disabled, nothing to do');
  assert.equal(find('resolved-unknown', 6).journalEn, 'Needs review: card creation started 2026-10-05 09:00 Los Angeles time, outcome unknown');
  assert.equal(find('resolved-unknown', 6).shopifyEn, 'Card still not found more than 10 minutes after creation started; confirmed not created');
  assert.equal(find('resolved-unknown', 6).actionEn, 'Set back to Pending in the local journal; the cards\' expiry date (2026-10-19) has been reached, so issue creates no new cards');
  assert.equal(find('resolved-unknown', 6).action, '已在本地日志改回待发放；礼品卡到期日（2026-10-19）已到，issue 不再建新卡');
  assert.equal(find('still-unknown', 7).journalEn, 'In progress: card creation started 2026-10-19 09:58 Los Angeles time, no result recorded');
  assert.equal(find('still-unknown', 7).shopifyEn, 'Card not found yet (a new card may not be in the search index yet)');
  assert.equal(find('still-unknown', 7).actionEn, 'Run verify or issue again in about 9 minutes to check');
  assert.equal(find('card-disabled', 8).journalEn, 'Skipped before issuing: something-new');
  assert.equal(find('card-disabled', 8).actionEn, 'If it was not disabled on purpose, re-enable the card in the admin');
  for (const issue of issues) {
    for (const key of ['journalEn', 'shopifyEn', 'actionEn']) {
      assert.equal(typeof issue[key], 'string');
      assert.ok(issue[key].length > 0, `${issue.type}.${key} is empty`);
      assert.doesNotMatch(issue[key], CJK, `${issue.type}.${key}: ${issue[key]}`);
    }
  }
});

test('a tagged customer gets no "not tagged" note; an unknown row without a start time is never settled automatically', () => {
  const selection = {
    params: { currency: 'USD' },
    recipients: [
      { seq: 1, customerId: gid('Customer', 1), amountCents: 1077 },
      { seq: 2, customerId: gid('Customer', 2), amountCents: 1077 },
    ],
  };
  const unknownWithoutStart = { status: STATUS.UNKNOWN, startedAt: null, giftCardId: null, taggedAt: null, attempts: 1 };
  const { issues, fixes } = analyzeVerify({
    selection,
    state: { customers: new Map([[gid('Customer', 2), unknownWithoutStart]]) },
    cards: [{ id: gid('GiftCard', 1), customerId: gid('Customer', 1), createdAt: '2026-10-05T16:00:00Z', enabled: true, last4: 'x001', amountCents: 1077, balanceCents: 1077, currencyCode: 'USD' }],
    taggedIds: new Set([gid('Customer', 1)]),
    nowMs: Date.parse('2026-12-31T00:00:00Z'), // long after: still not settled without a start time
    tag: 'OCT26RTPROMO',
    timezone: 'America/Los_Angeles',
  });
  assert.deepEqual(issues.map((i) => [i.type, tail(i.customerId)]), [['still-unknown', '2'], ['not-in-journal', '1']]);
  assert.equal(issues[0].action, '在后台按客户查看有没有这张卡');
  assert.equal(issues[1].shopify, '有本活动的卡：尾号 x001，原金额 $10.77，2026-10-05 09:00 建卡');
  assert.deepEqual(fixes.map((f) => [f.op, tail(f.cid)]), [['reconcile.found', '1']]);
});

test('nothing to report: exit 0 and a clean summary', async () => {
  env = testConfig();
  const { config } = env;
  const paths = campaignPaths(config);
  const list = customers(2);
  await selectionFixture(config, { customers: list });
  fake = installFakeShopify({ customers: shopifyCustomers(list, [1], config.sentTag), giftCards: [card(901, 1, { cents: 1533 })] });
  writeJournal(paths, issued(1, 901, 1533));
  const log = memoryLog();

  const { exitCode, result } = await runVerify({ config, log, now: at(RUN_AT), sleep: noSleep, writeReport: reportStub() });

  assert.equal(exitCode, 0);
  assert.deepEqual(result.issues, []);
  assert.equal(result.cardCount, 1);
  assert.ok(Object.values(result.counts).every((n) => n === 0));
  assert.ok(log.lines.some((l) => l.includes('没有发现问题')));
});

test('another command holding the run lock: exit 1 at once, nothing read or written', async () => {
  env = testConfig();
  const { config } = env;
  const paths = campaignPaths(config);
  await selectionFixture(config, { customers: customers(2) });
  fake = installFakeShopify({ customers: customers(2) });
  const held = holdRunLock(paths, 'issue');
  const log = memoryLog();
  const writeReport = reportStub();

  const { exitCode, result } = await runVerify({ config, log, now: at(RUN_AT), sleep: noSleep, writeReport });

  assert.equal(exitCode, 1);
  assert.equal(result, null);
  assert.ok(log.lines.some((l) => l.startsWith('ERROR') && l.includes('issue 正在运行')), log.lines.join('\n'));
  assert.deepEqual(JSON.parse(fs.readFileSync(paths.runLock, 'utf8')), held, 'the other command keeps its lock');
  assert.equal(fake.state.calls.token.length, 0);
  assert.equal(fake.state.calls.ops.length, 0);
  assert.equal(fs.existsSync(paths.journal), false);
  assert.equal(fs.existsSync(paths.verify), false);
  assert.equal(writeReport.calls.length, 0, 'exits immediately; the running command refreshes the Excel itself');
});

test('no list yet: exit 1 with the right message (also while select is running)', async () => {
  env = testConfig();
  const { config } = env;
  const paths = campaignPaths(config);
  fake = installFakeShopify();
  const writeReport = reportStub();

  let log = memoryLog();
  let out = await runVerify({ config, log, now: at(RUN_AT), sleep: noSleep, writeReport });
  assert.equal(out.exitCode, 1);
  // A lost folder must not send the operator back to select (which refuses once cards exist): copy it back.
  assert.deepEqual(log.lines, ['ERROR 还没有名单，请先运行 select。如果已经开始发放，不要重新运行 select，请把原来的 campaigns/2026-10/ 文件夹拷回来。']);
  assert.equal(fs.existsSync(paths.dir), false, 'no stray campaign folder for a wrong CAMPAIGN_ID');

  holdRunLock(paths, 'select');
  log = memoryLog();
  out = await runVerify({ config, log, now: at(RUN_AT), sleep: noSleep, writeReport });
  assert.equal(out.exitCode, 1);
  assert.deepEqual(log.lines, ['ERROR 还没有名单，select 正在运行，请等它跑完']);
  assert.equal(writeReport.calls.length, 0);
  assert.equal(fake.state.calls.token.length, 0);
});

test('a list from another campaign is refused and the lock is released', async () => {
  env = testConfig();
  const { config } = env;
  const paths = campaignPaths(config);
  const selection = await selectionFixture(config, { customers: customers(1), write: false });
  fs.mkdirSync(paths.dir, { recursive: true });
  fs.writeFileSync(paths.selection, JSON.stringify({ ...selection, campaignId: '2026-09' }));
  fake = installFakeShopify();
  const log = memoryLog();

  const { exitCode } = await runVerify({ config, log, now: at(RUN_AT), sleep: noSleep, writeReport: reportStub() });

  assert.equal(exitCode, 1);
  assert.ok(log.lines.some((l) => l.includes('名单属于活动 2026-09')), log.lines.join('\n'));
  assert.equal(fs.existsSync(paths.runLock), false);
  assert.equal(fs.existsSync(paths.journal), false);
  assert.equal(fake.state.calls.token.length, 0);
});

test('a Shopify failure: exit 1, run.end records it, lock released, Excel still regenerated', async () => {
  env = testConfig();
  const { config } = env;
  const paths = campaignPaths(config);
  const list = customers(2);
  await selectionFixture(config, { customers: list });
  fake = installFakeShopify({ customers: list, failures: { FindCampaignCards: [{ kind: 'http', status: 503 }] } });
  writeJournal(paths, lost(1, minutesBefore(RUN_AT, 30)));
  const log = memoryLog();
  const writeReport = reportStub();

  const { exitCode, result } = await runVerify({ config, log, now: at(RUN_AT), sleep: noSleep, writeReport });

  assert.equal(exitCode, 1);
  assert.equal(result, null);
  assert.ok(log.lines.some((l) => l.startsWith('ERROR verify 没有完成') && l.includes('503')), log.lines.join('\n'));
  const entries = readJournal(paths.journal);
  const end = entries.find((e) => e.op === 'run.end');
  assert.equal(end.exitCode, 1);
  assert.match(end.summary.error, /503/);
  assert.equal(entries.filter((e) => e.op.startsWith('reconcile.')).length, 0, 'no decision without data');
  assert.equal(fs.existsSync(paths.verify), false);
  assert.equal(fs.existsSync(paths.tags), false);
  assert.equal(fs.existsSync(paths.runLock), false);
  assert.equal(writeReport.calls.length, 1);
});

test('the Excel is regenerated after run.end and after the lock is released; its failure is only a warning', async () => {
  env = testConfig();
  const { config } = env;
  const paths = campaignPaths(config);
  const list = customers(1);
  await selectionFixture(config, { customers: list });
  fake = installFakeShopify({ customers: list });
  const seen = [];
  const writeReport = reportStub({
    fail: 'disk full',
    onCall: () => seen.push({ locked: fs.existsSync(paths.runLock), last: readJournal(paths.journal).at(-1).op }),
  });
  const log = memoryLog();

  const { exitCode } = await runVerify({ config, log, now: at(RUN_AT), sleep: noSleep, writeReport });

  assert.equal(exitCode, 0, 'the verification itself succeeded');
  assert.deepEqual(seen, [{ locked: false, last: 'run.end' }]);
  assert.ok(log.lines.some((l) => l.startsWith('WARN Excel 没有更新：disk full')), log.lines.join('\n'));
  assert.ok(fs.existsSync(paths.verify));
});

test('Excel warnings are printed once: writeReport logs them, verify does not repeat result.warnings', async () => {
  env = testConfig();
  const { config } = env;
  const paths = campaignPaths(config);
  const list = customers(1);
  await selectionFixture(config, { customers: list });
  fake = installFakeShopify({ customers: list });
  const warning = 'Excel 正打开此文件，请关闭后重新打开才能看到最新内容';
  const log = memoryLog();
  // Like the real writer: logged when it happens, and returned.
  const writeReport = async (args) => {
    args.log.warn(warning);
    return { file: args.paths.excel, fileEn: args.paths.excelEn, out: null, outEn: null, warnings: [warning] };
  };

  const { exitCode } = await runVerify({ config, log, now: at(RUN_AT), sleep: noSleep, writeReport });

  assert.equal(exitCode, 0, log.lines.join('\n'));
  assert.equal(log.lines.filter((l) => l === `WARN ${warning}`).length, 1, log.lines.join('\n'));
  assert.ok(log.lines.includes(`INFO Excel 已更新：${paths.excel}`), log.lines.join('\n'));
  // The English edition's path follows the Chinese one.
  const zhAt = log.lines.indexOf(`INFO Excel 已更新：${paths.excel}`);
  assert.equal(log.lines[zhAt + 1], `INFO 英文版 Excel 已更新：${paths.excelEn}`, log.lines.join('\n'));
});

test('when the English workbook failed (fileEn null) no English path is printed and the exit code is unchanged', async () => {
  env = testConfig();
  const { config } = env;
  const paths = campaignPaths(config);
  const list = customers(1);
  await selectionFixture(config, { customers: list });
  fake = installFakeShopify({ customers: list });
  const log = memoryLog();
  // Like the real writer: the English failure was already logged as a warning, fileEn is null.
  const writeReport = async (args) => {
    args.log.warn('英文版 Excel 没有生成：disk full');
    return { file: args.paths.excel, fileEn: null, out: null, outEn: null, warnings: ['英文版 Excel 没有生成：disk full'] };
  };

  const { exitCode } = await runVerify({ config, log, now: at(RUN_AT), sleep: noSleep, writeReport });

  assert.equal(exitCode, 0, log.lines.join('\n'));
  assert.ok(log.lines.includes(`INFO Excel 已更新：${paths.excel}`), log.lines.join('\n'));
  assert.equal(log.lines.filter((l) => l.includes('英文版 Excel 已更新')).length, 0, log.lines.join('\n'));
  assert.equal(log.lines.filter((l) => l === 'WARN 英文版 Excel 没有生成：disk full').length, 1, 'printed once, by writeReport');
});

test('with the real Excel writer, "Excel 正打开此文件" is printed exactly once', async () => {
  env = testConfig();
  const { config } = env;
  const paths = campaignPaths(config);
  const list = customers(2);
  await selectionFixture(config, { customers: list });
  fake = installFakeShopify({ customers: shopifyCustomers(list, [1], config.sentTag), giftCards: [card(901, 1, { cents: 1533 })] });
  writeJournal(paths, issued(1, 901, 1533));
  fs.writeFileSync(path.join(paths.dir, `~$${path.basename(paths.excel)}`), 'owner'); // the workbook is open in Excel
  const log = memoryLog();

  const { exitCode } = await runVerify({ config, log, now: at(RUN_AT), sleep: noSleep }); // no writeReport: src/report/excel.js

  assert.equal(exitCode, 0, log.lines.join('\n'));
  assert.ok(fs.existsSync(paths.excel), `the workbook was written:\n${log.lines.join('\n')}`);
  assert.equal(log.lines.filter((l) => l === `WARN ${OPEN_IN_EXCEL_WARNING}`).length, 1, log.lines.join('\n'));
  // The English edition is written next to it and its path printed right after the Chinese one.
  assert.ok(fs.existsSync(paths.excelEn), `the English workbook was written:\n${log.lines.join('\n')}`);
  const zhAt = log.lines.indexOf(`INFO Excel 已更新：${paths.excel}`);
  assert.ok(zhAt >= 0, log.lines.join('\n'));
  assert.equal(log.lines[zhAt + 1], `INFO 英文版 Excel 已更新：${paths.excelEn}`, log.lines.join('\n'));
});

test('a changed SENT_TAG since select is pointed out', async () => {
  env = testConfig();
  const list = customers(1);
  await selectionFixture(env.config, { customers: list });
  const other = testConfig({ CAMPAIGNS_DIR: env.dir, SENT_TAG: 'renamed-tag' });
  try {
    fake = installFakeShopify({ customers: list });
    const log = memoryLog();
    const { exitCode } = await runVerify({ config: other.config, log, now: at(RUN_AT), sleep: noSleep, writeReport: reportStub() });
    assert.equal(exitCode, 0);
    assert.ok(log.lines.some((l) => l.startsWith('WARN') && l.includes('gift-card-sent-2026-10') && l.includes('renamed-tag')), log.lines.join('\n'));
    assert.equal(readJson(campaignPaths(other.config).tags).tag, 'renamed-tag');
  } finally {
    other.cleanup();
  }
});
