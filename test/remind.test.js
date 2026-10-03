// Tests for the `remind` command (src/remind.js) against the in-process fake Shopify.
import { test, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { installFakeShopify } from './fake-shopify.js';
import { testConfig, selectionFixture, makeCustomer, makeOrder, memoryLog, gid, NOW_ISO } from './helpers.js';
import { campaignPaths, appendJournal, readJournal, foldJournal, acquireRunLock, writeJsonAtomic } from '../src/campaign.js';
import { campaignNote } from '../src/giftcards.js';
import { buildTestSelection } from '../src/select/selection.js';
import { parseCustomer } from '../src/select/rules.js';
import {
  runRemind, pickCard, cardSkip, expirySkip, customerSkip, templateStageOn, utcEveningWarning, roundTag, STAGE_LABELS, SELECTION_REPLACED, REMIND_SKIP_REASONS,
} from '../src/remind.js';
import { runVerify } from '../src/verify.js';
import { expectedStage, renderPreviewFor } from '../src/preview.js';
import { localDate } from '../src/time.js';

const ROUND1_DAY = new Date('2026-10-12T17:00:00Z'); // 10:00 on REMIND_1_DATE in Los Angeles
const ROUND2_DAY = new Date('2026-10-16T17:00:00Z'); // 10:00 on REMIND_2_DATE in Los Angeles
const ISSUED_AT = '2026-10-05T17:00:00.000Z';
const ISSUE_RUN = '20261005170000-1';

let teardown = [];
afterEach(() => {
  for (const fn of teardown.reverse()) fn();
  teardown = [];
});

const customerGid = (n) => gid('Customer', n);
const cardGid = (n) => gid('GiftCard', n);

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

/** `count` customers who pass every selection rule; seq n is customer n (newest last order first). */
function eligibleCustomers(count) {
  const base = Date.parse('2026-06-30T12:00:00Z');
  return Array.from({ length: count }, (_, i) => {
    const n = i + 1;
    return makeCustomer({ n, lastOrder: makeOrder({ n: 5000 + n, createdAt: new Date(base - n * 3_600_000).toISOString(), total: '80.00' }) });
  });
}

/** A GiftCard node as the fake store keeps it (carries the campaign marker in its note). */
function makeCard(n, customerId, { amount = '10.77', balance = amount, enabled = true, expiresOn = '2026-10-19', campaignId = '2026-10', createdAt = ISSUED_AT } = {}) {
  return {
    id: cardGid(n),
    createdAt,
    note: campaignNote('gift-card-promo', campaignId),
    templateSuffix: 'gift-card-promo',
    enabled,
    lastCharacters: `x${String(n).slice(-3)}`,
    initialValue: { amount, currencyCode: 'USD' },
    balance: { amount: balance, currencyCode: 'USD' },
    customer: customerId ? { id: customerId } : null,
    expiresOn,
  };
}

/** The journal an `issue` run leaves behind. An entry without giftCardId ends as create.unknown. */
function writeIssueJournal(paths, issued) {
  const at = { now: () => ISSUED_AT };
  appendJournal(paths.journal, { op: 'run.start', run: ISSUE_RUN, command: 'issue', dryRun: false, batch: 1, limit: issued.length, options: {} }, at);
  for (const x of issued) {
    appendJournal(paths.journal, { op: 'create.start', cid: x.cid, amountCents: x.amountCents, batch: 1, run: ISSUE_RUN }, at);
    if (x.giftCardId) {
      appendJournal(paths.journal, { op: 'create.ok', cid: x.cid, giftCardId: x.giftCardId, last4: x.last4, amountCents: x.amountCents, batch: 1, run: ISSUE_RUN }, at);
      appendJournal(paths.journal, { op: 'tag.ok', cid: x.cid, run: ISSUE_RUN }, at);
    } else {
      appendJournal(paths.journal, { op: 'create.unknown', cid: x.cid, error: 'socket hang up', run: ISSUE_RUN }, at);
    }
  }
  appendJournal(paths.journal, { op: 'run.end', run: ISSUE_RUN, summary: {}, exitCode: 0 }, at);
}

/**
 * A campaign right after `issue`: selection.json, one unused card per recipient
 * in the fake store (card 2000+seq) and the matching journal. `tweak(ctx)` may
 * edit ctx.cards / ctx.issued before they are installed and journalled.
 */
async function issuedCampaign({ count = 3, env = {}, testCampaign = false, tweak, failures } = {}) {
  const testIds = Array.from({ length: count }, (_, i) => String(i + 1)).join(',');
  const t = testConfig(testCampaign ? { CAMPAIGN_ID: '2026-10-test', SENT_TAG: 'OCT26RTPROMO-TEST', TEST_CUSTOMER_IDS: testIds, ...env } : env);
  teardown.push(t.cleanup);
  const { config } = t;
  const paths = campaignPaths(config);
  let customers;
  let selection;
  if (testCampaign) {
    customers = Array.from({ length: count }, (_, i) => makeCustomer({ n: i + 1 }));
    selection = buildTestSelection({ config, customers: customers.map(parseCustomer), timezone: 'America/Los_Angeles', createdAt: NOW_ISO, snapshot: { exportedAt: NOW_ISO, source: 'nodes', count } });
    writeJsonAtomic(paths.selection, selection);
  } else {
    customers = eligibleCustomers(count);
    selection = await selectionFixture(config, { customers });
  }
  assert.deepEqual(selection.recipients.map((r) => r.customerId), customers.map((c) => c.id), 'fixture: seq n is customer n');

  const cards = selection.recipients.map((r) => makeCard(2000 + r.seq, r.customerId, { amount: (r.amountCents / 100).toFixed(2), campaignId: config.campaignId }));
  const issued = selection.recipients.map((r, i) => ({ cid: r.customerId, giftCardId: cards[i].id, last4: cards[i].lastCharacters, amountCents: r.amountCents }));
  const ctx = { config, paths, selection, customers, cards, issued };
  tweak?.(ctx);
  writeIssueJournal(paths, ctx.issued);
  const fake = installFakeShopify({ customers: ctx.customers, giftCards: ctx.cards, failures, now: () => Date.parse(ISSUED_AT) });
  teardown.push(() => fake.restore());
  ctx.fake = fake;
  ctx.card = (seq) => fake.state.giftCards.find((g) => g.id === cardGid(2000 + seq));
  ctx.customer = (n) => fake.state.customers.find((c) => c.id === customerGid(n));
  return ctx;
}

/** runRemind with test doubles: round 1 on REMIND_1_DATE unless `opts` override it. */
async function remind(ctx, opts = {}) {
  const log = memoryLog();
  const reports = [];
  const previews = [];
  const sleeps = [];
  const result = await runRemind({
    config: ctx.config,
    round: 1,
    log,
    now: () => ROUND1_DAY,
    sleep: async (ms) => {
      sleeps.push(ms);
    },
    writeReport: async (args) => {
      reports.push(args);
      return { file: ctx.paths.excel, out: null, warnings: [] };
    },
    renderPreview: async (args) => {
      previews.push(args);
      return { file: `${args.paths.previewDir}/${args.variant}.html`, subject: 'Reminder', stage: args.variant };
    },
    ...opts,
  });
  return { ...result, log, reports, previews, sleeps };
}

/**
 * Wrap the fake's fetch to watch every SendGiftCardNotification request.
 * `answer(variables)` may return a Response to reply instead of the fake
 * (the fake's own `userError` failure kind does not use this mutation's payload key).
 * Returns the ids of every request seen. fake.restore() removes the wrapper.
 */
function interceptNotify(answer = () => null) {
  const calls = [];
  const inner = globalThis.fetch;
  globalThis.fetch = async (url, init = {}) => {
    if (String(url).endsWith('/graphql.json')) {
      const { query, variables } = JSON.parse(init.body);
      if (/mutation\s+SendGiftCardNotification\b/.test(query)) {
        calls.push(variables.id);
        const response = await answer(variables);
        if (response) return response;
      }
    }
    return inner(url, init);
  };
  return calls;
}

/** What Shopify answers when it refuses giftCardSendNotificationToCustomer. */
function userErrorsResponse(message = 'Gift card is not eligible for a notification') {
  const body = { data: { giftCardSendNotificationToCustomer: { giftCard: null, userErrors: [{ field: ['id'], message, code: 'INVALID' }] } } };
  return new Response(JSON.stringify(body), { status: 200, headers: { 'content-type': 'application/json' } });
}

/**
 * A clock that stands still between sends and moves `stepMs` on with every
 * SendGiftCardNotification request, as a long run does; it starts at `startIso`.
 * Independent of how often the code reads the clock. fake.restore() removes the wrapper.
 */
function clockMovedBySends(startIso, stepMs = 60_000) {
  let ms = Date.parse(startIso);
  interceptNotify(() => {
    ms += stepMs;
    return null;
  });
  return () => new Date(ms);
}

/** Compare only the listed summary fields. */
function assertSummaryOf(summary, expected) {
  assert.deepEqual(Object.fromEntries(Object.keys(expected).map((k) => [k, summary[k]])), expected);
}

const journalOf = (ctx) => readJournal(ctx.paths.journal);
const remindOps = (ctx) => journalOf(ctx).filter((e) => e.op.startsWith('remind.'));
const roundState = (ctx, n, round = 1) => foldJournal(journalOf(ctx)).customers.get(customerGid(n))?.reminders?.[String(round)] ?? null;
const notified = (ctx) => ctx.fake.state.calls.notify;
const logText = (res) => res.log.lines.join('\n');
const throttledOut = () => Array.from({ length: 6 }, () => ({ kind: 'throttle' })); // 1 try + 5 retries

// ---------------------------------------------------------------------------
// Pure rules
// ---------------------------------------------------------------------------

test('pickCard: the journal card when Shopify lists it, else the only card; otherwise no-card / multiple-cards', () => {
  const a = { id: cardGid(1), last4: 'x001' };
  const b = { id: cardGid(2), last4: 'x002' };
  assert.deepEqual(pickCard([a], null), { card: a });
  assert.deepEqual(pickCard([a, b], cardGid(2)), { card: b });
  assert.deepEqual(pickCard([b], cardGid(1)), { card: b }); // exactly one campaign card: no ambiguity
  assert.equal(pickCard([], cardGid(1)).skip.reason, 'no-card');
  assert.match(pickCard([], cardGid(1)).skip.detail, /\b1\b/);
  assert.deepEqual(pickCard([a, b], null).skip, { reason: 'multiple-cards', detail: '2 张卡：x001、x002' });
  assert.equal(pickCard([a, b], cardGid(3)).skip.reason, 'multiple-cards');
});

test('cardSkip and customerSkip apply the reminder rules in order', () => {
  const card = { enabled: true, expiresOn: '2026-10-19', amountCents: 1077, balanceCents: 1077 };
  assert.equal(cardSkip(card, '2026-10-12'), null);
  assert.equal(cardSkip(card, '2026-10-19'), null, 'still usable on its expiry date');
  assert.deepEqual(cardSkip(card, '2026-10-20'), { reason: 'card-expired', detail: '到期日 2026-10-19' });
  assert.equal(cardSkip({ ...card, expiresOn: null }, '2030-01-01'), null);
  assert.equal(cardSkip({ ...card, enabled: false, balanceCents: 0 }, '2026-10-12').reason, 'card-disabled');
  assert.deepEqual(cardSkip({ ...card, balanceCents: 1076 }, '2026-10-12'), { reason: 'used', detail: '余额 $10.76 / 面额 $10.77' });

  assert.equal(customerSkip(makeCustomer({ n: 1 })), null);
  assert.equal(customerSkip(null).reason, 'customer-deleted');
  assert.equal(customerSkip({}).reason, 'customer-deleted');
  assert.equal(customerSkip(makeCustomer({ n: 1, email: null })).reason, 'no-email');
  assert.deepEqual(customerSkip(makeCustomer({ n: 1, validFormat: false })), { reason: 'no-email', detail: 'c1@example.org' });
  assert.deepEqual(customerSkip(makeCustomer({ n: 1, marketingState: 'UNSUBSCRIBED' })), { reason: 'not-subscribed', detail: 'UNSUBSCRIBED' });
  assert.equal(customerSkip(makeCustomer({ n: 1, marketingState: 'PENDING' })).reason, 'not-subscribed');
  assert.deepEqual(Object.keys(REMIND_SKIP_REASONS), ['no-card', 'multiple-cards', 'card-disabled', 'card-expired', 'used', 'customer-deleted', 'no-email', 'not-subscribed']);
});

// ---------------------------------------------------------------------------
// Who gets a reminder
// ---------------------------------------------------------------------------

test('reminds only unused cards of subscribed customers, in seq order, journalling every skip', async () => {
  const ctx = await issuedCampaign({
    count: 13,
    tweak(c) {
      c.issued = c.issued.filter((x) => x.cid !== customerGid(11)); // 11: never got a card
      c.cards = c.cards.filter((g) => g.id !== cardGid(2011));
      c.issued.find((x) => x.cid === customerGid(8)).giftCardId = null; // 8: issue outcome unknown...
      c.cards.push(makeCard(3108, customerGid(8))); // ...and two campaign cards in Shopify
      c.cards.push(makeCard(3112, customerGid(12))); // 12: a duplicate card, but the journal names 2012
      c.cards.push(makeCard(4001, customerGid(1), { campaignId: '2026-09' })); // another campaign's card: ignored
      const manual = makeCard(4009, customerGid(9));
      manual.note = 'birthday gift'; // a card without the campaign marker: ignored
      c.cards.push(manual);
    },
  });
  const { fake } = ctx;
  fake.spendCard(cardGid(2002), 100); // 2: partly used
  ctx.card(3).enabled = false; // 3: disabled
  ctx.card(4).expiresOn = '2026-10-11'; // 4: expired yesterday (store time)
  ctx.customer(5).defaultEmailAddress.marketingState = 'UNSUBSCRIBED'; // 5
  fake.state.customers.splice(fake.state.customers.indexOf(ctx.customer(6)), 1); // 6: deleted
  ctx.card(7).customer = { id: customerGid(999) }; // 7: the journal's card now belongs to someone else
  ctx.customer(10).defaultEmailAddress.validFormat = false; // 10: email no longer valid
  ctx.card(13).expiresOn = '2026-10-12'; // 13: today is its last usable day: still reminded

  const lateJournal = [];
  interceptNotify((variables) => {
    const last = readJournal(ctx.paths.journal).at(-1);
    if (last.op !== 'remind.start' || last.giftCardId !== variables.id) lateJournal.push(variables.id);
  });

  const res = await remind(ctx);
  assert.equal(res.exitCode, 0);
  assert.deepEqual(lateJournal, [], 'remind.start must be on disk before every send');
  assert.deepEqual(notified(ctx), [cardGid(2001), cardGid(2009), cardGid(2012), cardGid(2013)]);
  assert.deepEqual(res.summary.skipped, {
    used: 1, 'card-disabled': 1, 'card-expired': 1, 'not-subscribed': 1, 'customer-deleted': 1, 'no-card': 1, 'multiple-cards': 1, 'no-email': 1,
  });
  assert.equal(res.summary.eligible, 4);
  assert.equal(res.summary.sent, 4);
  assert.equal(res.summary.notIssued, 1);
  assert.equal(res.summary.stopped, null);

  // Only people whose card qualifies are re-read from Shopify.
  assert.deepEqual(fake.opsNamed('RefreshCustomers').flatMap((o) => o.variables.ids), [1, 5, 6, 9, 10, 12, 13].map(customerGid));

  const runId = journalOf(ctx).filter((e) => e.op === 'run.start').at(-1).run;
  const ops = remindOps(ctx);
  for (const e of ops) {
    assert.equal(e.round, 1);
    assert.equal(e.run, runId);
  }
  const byCustomer = (n) => ops.filter((e) => e.cid === customerGid(n)).map((e) => (e.op === 'remind.skip' ? `skip:${e.reason}` : e.op));
  const expected = {
    1: ['remind.start', 'remind.ok'],
    2: ['skip:used'],
    3: ['skip:card-disabled'],
    4: ['skip:card-expired'],
    5: ['skip:not-subscribed'],
    6: ['skip:customer-deleted'],
    7: ['skip:no-card'],
    8: ['skip:multiple-cards'],
    9: ['remind.start', 'remind.ok'],
    10: ['skip:no-email'],
    11: [],
    12: ['remind.start', 'remind.ok'],
    13: ['remind.start', 'remind.ok'],
  };
  for (const [n, want] of Object.entries(expected)) assert.deepEqual(byCustomer(Number(n)), want, `customer ${n}`);
  assert.deepEqual(ops.filter((e) => e.op === 'remind.start').map((e) => e.giftCardId), [2001, 2009, 2012, 2013].map(cardGid));

  const detail = (n) => ops.find((e) => e.cid === customerGid(n) && e.op === 'remind.skip').detail;
  assert.equal(detail(2), '余额 $9.77 / 面额 $10.77');
  assert.equal(detail(4), '到期日 2026-10-11');
  assert.equal(detail(5), 'UNSUBSCRIBED');
  assert.match(detail(7), /2007/);
  assert.equal(detail(8), '2 张卡：x008、x108');
  assert.equal(detail(10), 'c10@example.org');

  const folded = foldJournal(journalOf(ctx)).customers;
  assert.equal(folded.get(customerGid(1)).reminders['1'].status, 'sent');
  assert.equal(folded.get(customerGid(1)).reminders['1'].giftCardId, cardGid(2001));
  assert.equal(folded.get(customerGid(2)).reminders['1'].status, 'skipped');
  assert.equal(folded.get(customerGid(2)).reminders['1'].reason, 'used');
  assert.equal(folded.get(customerGid(11)), undefined);

  const end = journalOf(ctx).at(-1);
  assert.equal(end.op, 'run.end');
  assert.equal(end.exitCode, 0);
  assert.deepEqual(end.summary, res.summary);
  const start = journalOf(ctx).filter((e) => e.op === 'run.start').at(-1);
  assert.equal(start.command, 'remind');
  assert.equal(start.dryRun, false);
  assert.equal(start.limit, null);
  assert.deepEqual(start.options, { round: 1, retryUnknown: false, retryFailed: false });

  const text = logText(res);
  assert.match(text, /WARN 1 张本活动的卡不属于名单里的客户/); // card 2007 now belongs to customer 999
  assert.match(text, /跳过 8 人：/);
  assert.match(text, /名单里还没有建卡：1 人/);
  assert.equal(res.reports.length, 1);
});

test('a second run of the same round sends nothing new; round 2 is independent of round 1', async () => {
  const ctx = await issuedCampaign({ count: 3 });
  const first = await remind(ctx);
  assert.equal(first.exitCode, 0);
  assert.deepEqual(notified(ctx), [2001, 2002, 2003].map(cardGid));

  const again = await remind(ctx);
  assert.equal(again.exitCode, 0);
  assert.equal(notified(ctx).length, 3, 'at most one reminder per customer and round');
  assert.equal(again.summary.alreadySent, 3);
  assert.equal(again.summary.eligible, 0);
  assert.equal(ctx.fake.opsNamed('RefreshCustomers').length, 1, 'people already reminded are not even re-read');

  ctx.fake.spendCard(cardGid(2002), 1077); // customer 2 spends the whole card after round 1
  const second = await remind(ctx, { round: 2, now: () => ROUND2_DAY });
  assert.equal(second.exitCode, 0);
  assert.deepEqual(notified(ctx).slice(3), [cardGid(2001), cardGid(2003)]);
  assert.deepEqual(second.summary.skipped, { used: 1 });
  assert.equal(roundState(ctx, 1, 1).status, 'sent');
  assert.equal(roundState(ctx, 1, 2).status, 'sent');
  assert.equal(roundState(ctx, 2, 1).status, 'sent', 'round 2 does not touch round 1');
  assert.equal(roundState(ctx, 2, 2).reason, 'used');

  const secondAgain = await remind(ctx, { round: 2, now: () => ROUND2_DAY });
  assert.equal(notified(ctx).length, 5);
  assert.equal(secondAgain.summary.alreadySent, 2);
  assert.deepEqual(secondAgain.summary.skipped, { used: 1 }, 'a skip is evaluated again...');
  assert.equal(remindOps(ctx).filter((e) => e.op === 'remind.skip' && e.round === 2).length, 1, '...but an unchanged skip is not journalled twice');
});

test('a skip is not final: a card hidden by search-index lag gets its reminder in a later run', async () => {
  const ctx = await issuedCampaign({ count: 2 });
  ctx.fake.state.hiddenCardIds.add(cardGid(2001));
  const first = await remind(ctx);
  assert.equal(first.exitCode, 0);
  assert.deepEqual(first.summary.skipped, { 'no-card': 1 });
  assert.deepEqual(notified(ctx), [cardGid(2002)]);
  assert.equal(roundState(ctx, 1).reason, 'no-card');

  ctx.fake.state.hiddenCardIds.clear();
  const second = await remind(ctx);
  assert.deepEqual(notified(ctx), [cardGid(2002), cardGid(2001)]);
  assert.equal(roundState(ctx, 1).status, 'sent');
  assert.equal(second.summary.alreadySent, 1);
});

test('customers are re-read in chunks of 250; cards known only to Shopify count; a dry run lists the first 10', async () => {
  const t = testConfig({ DRY_RUN: 'true' });
  teardown.push(t.cleanup);
  const customers = eligibleCustomers(260);
  const selection = await selectionFixture(t.config, { customers });
  const fake = installFakeShopify({ customers, giftCards: selection.recipients.map((r) => makeCard(2000 + r.seq, r.customerId)) });
  teardown.push(() => fake.restore());
  const ctx = { config: t.config, paths: campaignPaths(t.config), fake }; // no journal at all

  const res = await remind(ctx);
  assert.equal(res.exitCode, 0);
  assert.deepEqual(fake.opsNamed('RefreshCustomers').map((o) => o.variables.ids.length), [250, 10]);
  assert.equal(fake.opsNamed('FindCampaignCards').length, 2, 'all card pages are read');
  assert.equal(res.summary.eligible, 260);
  const text = logText(res);
  assert.match(text, /本次会发的前 10 人/);
  assert.match(text, /#10 First10 Last10 <c10@example.org>/);
  assert.doesNotMatch(text, /#11 First11/);
});

test('a refresh too expensive for one query is retried with fewer customers at a time', async () => {
  const t = testConfig({ DRY_RUN: 'true' });
  teardown.push(t.cleanup);
  const customers = eligibleCustomers(260);
  const selection = await selectionFixture(t.config, { customers });
  const fake = installFakeShopify({
    customers,
    giftCards: selection.recipients.map((r) => makeCard(2000 + r.seq, r.customerId)),
    failures: { RefreshCustomers: [{ kind: 'graphqlError', code: 'MAX_COST_EXCEEDED', message: 'Query cost is 1752, which exceeds the single query max cost limit (1000).' }] },
  });
  teardown.push(() => fake.restore());
  const res = await remind({ config: t.config, paths: campaignPaths(t.config), fake });
  assert.equal(res.exitCode, 0);
  assert.deepEqual(fake.opsNamed('RefreshCustomers').map((o) => o.variables.ids.length), [250, 125, 125, 10]);
  assert.equal(res.summary.eligible, 260);
  assert.match(logText(res), /改为每次 125 位/);
});

test('a config without an explicit dryRun: false never sends', async () => {
  const ctx = await issuedCampaign({ count: 2 });
  const { dryRun, ...noMode } = ctx.config;
  assert.equal(dryRun, false);
  const res = await remind(ctx, { config: noMode });
  assert.equal(res.exitCode, 0);
  assert.equal(res.summary.dryRun, true);
  assert.equal(ctx.fake.opsNamed('SendGiftCardNotification').length, 0);
});

// ---------------------------------------------------------------------------
// Date guard
// ---------------------------------------------------------------------------

test('date guard: a live run before REMIND_1_DATE (store time) exits 2 without touching Shopify', async () => {
  const ctx = await issuedCampaign({ count: 2 });
  const before = journalOf(ctx).length;
  const early = await remind(ctx, { now: () => new Date('2026-10-12T06:59:00Z') }); // 23:59 on 10/11 in Los Angeles
  assert.equal(early.exitCode, 2);
  assert.equal(early.summary, null);
  assert.match(logText(early), /ERROR 第 1 次提醒要到 2026-10-12（店铺时间）才能发/);
  assert.equal(ctx.fake.state.calls.token.length, 0);
  assert.equal(ctx.fake.state.calls.ops.length, 0);
  assert.equal(journalOf(ctx).length, before);
  assert.equal(early.reports.length, 0);

  const onTheDay = await remind(ctx, { now: () => new Date('2026-10-12T07:00:00Z') }); // 00:00 on 10/12 in Los Angeles
  assert.equal(onTheDay.exitCode, 0);
  assert.deepEqual(notified(ctx), [cardGid(2001), cardGid(2002)]);
});

test('date guard: round 2 waits for REMIND_2_DATE', async () => {
  const ctx = await issuedCampaign({ count: 1 });
  const early = await remind(ctx, { round: 2, now: () => new Date('2026-10-15T20:00:00Z') });
  assert.equal(early.exitCode, 2);
  assert.match(logText(early), /第 2 次提醒要到 2026-10-16（店铺时间）才能发/);
  assert.equal(ctx.fake.state.calls.ops.length, 0);
  const onTime = await remind(ctx, { round: 2, now: () => ROUND2_DAY });
  assert.equal(onTime.exitCode, 0);
  assert.equal(notified(ctx).length, 1);
});

test('date guard: a missing REMIND date is a usage error for a live run', async () => {
  const ctx = await issuedCampaign({ count: 1, env: { REMIND_1_DATE: '' } });
  const res = await remind(ctx);
  assert.equal(res.exitCode, 2);
  assert.match(logText(res), /请在 \.env 设置 REMIND_1_DATE，格式 YYYY-MM-DD/);
  assert.equal(ctx.fake.state.calls.ops.length, 0);
});

test('date guard: dry runs are always allowed, with a warning before the date', async () => {
  const ctx = await issuedCampaign({ count: 2, env: { DRY_RUN: 'true' } });
  const res = await remind(ctx, { now: () => new Date('2026-10-05T18:00:00Z') });
  assert.equal(res.exitCode, 0);
  assert.equal(res.summary.planned, 2);
  assert.equal(ctx.fake.opsNamed('SendGiftCardNotification').length, 0);
  assert.match(logText(res), /第 1 次提醒要到 2026-10-12（店铺时间）才能真实发送/);
});

test('date guard: test campaigns send on any day, warning that the template still shows the first email', async () => {
  const ctx = await issuedCampaign({ count: 2, testCampaign: true });
  const res = await remind(ctx, { now: () => new Date('2026-10-03T18:00:00Z') });
  assert.equal(res.exitCode, 0);
  assert.deepEqual(notified(ctx), [cardGid(2001), cardGid(2002)]);
  const text = logText(res);
  assert.match(text, /测试活动：不受提醒日期限制/);
  assert.match(text, /WARN 注意：邮件模板按发送当天的日期切换文案，按 \.env 的日期，今天（2026-10-03）发出的是首封文案，不是第一次提醒文案；第一次提醒的文案请先用 preview 查看/);
});

test('date guard: test campaigns send even when the REMIND date is not set', async () => {
  const ctx = await issuedCampaign({ count: 1, testCampaign: true, env: { REMIND_1_DATE: '' } });
  const res = await remind(ctx, { now: () => new Date('2026-10-03T18:00:00Z') });
  assert.equal(res.exitCode, 0);
  assert.deepEqual(notified(ctx), [cardGid(2001)]);
  assert.match(logText(res), /WARN 注意：邮件模板按发送当天的日期切换文案；\.env 没有同时设置 REMIND_1_DATE 和 REMIND_2_DATE，无法判断今天发出的是哪种文案/);
});

test('templateStageOn is preview.expectedStage\'s rule: first, then remind1 from REMIND_1_DATE, remind2 from REMIND_2_DATE', () => {
  const t = testConfig();
  teardown.push(t.cleanup);
  const { config } = t;
  for (const date of ['2026-10-01', '2026-10-11', '2026-10-12', '2026-10-15', '2026-10-16', '2026-10-19', '2026-11-30']) {
    for (const variant of ['first', 'remind1', 'remind2']) {
      assert.equal(templateStageOn(config, date), expectedStage(config, variant, date), `${variant} sent on ${date}`);
    }
  }
  assert.deepEqual(['2026-10-11', '2026-10-12', '2026-10-15', '2026-10-16'].map((d) => templateStageOn(config, d)), ['first', 'remind1', 'remind1', 'remind2']);
  assert.equal(templateStageOn({ ...config, remind1Date: '' }, '2026-10-16'), null, '.env cannot tell without both dates');
  assert.equal(templateStageOn({ ...config, remind2Date: '' }, '2026-10-12'), null);
  assert.deepEqual(STAGE_LABELS, { first: '首封', remind1: '第一次提醒', remind2: '第二次提醒' });
});

test('date guard: a live round 1 on or after REMIND_2_DATE (store time) exits 2 without touching Shopify: that email would show the round-2 copy', async () => {
  const ctx = await issuedCampaign({ count: 2 });
  const before = journalOf(ctx).length;
  const late = await remind(ctx, { now: () => new Date('2026-10-16T07:00:00Z') }); // 00:00 on 10/16 in Los Angeles
  assert.equal(late.exitCode, 2);
  assert.equal(late.summary, null);
  assert.match(logText(late), /ERROR 第 1 次提醒要在 REMIND_2_DATE（2026-10-16）之前发：今天发出的邮件会显示第二次提醒的文案。请改发 --round 2。/);
  assert.equal(ctx.fake.state.calls.token.length, 0);
  assert.equal(ctx.fake.state.calls.ops.length, 0);
  assert.equal(journalOf(ctx).length, before);
  assert.equal(late.reports.length, 0);
  assert.equal((await remind(ctx, { now: () => new Date('2026-10-19T20:00:00Z') })).exitCode, 2, 'and on every later day');
  assert.deepEqual(notified(ctx), []);

  const lastDay = await remind(ctx, { now: () => new Date('2026-10-16T06:59:00Z') }); // 23:59 on 10/15 in Los Angeles
  assert.equal(lastDay.exitCode, 0);
  assert.deepEqual(notified(ctx), [cardGid(2001), cardGid(2002)]);
  const round2 = await remind(ctx, { round: 2, now: () => ROUND2_DAY });
  assert.equal(round2.exitCode, 0, 'round 2 itself goes out on REMIND_2_DATE');
  assert.equal(notified(ctx).length, 4);
});

test('date guard: a dry round 1 on or after REMIND_2_DATE warns, and its preview is the round-2 copy that day really gets', async () => {
  const ctx = await issuedCampaign({ count: 1, env: { DRY_RUN: 'true' } });
  const res = await remind(ctx, { now: () => ROUND2_DAY });
  assert.equal(res.exitCode, 0);
  assert.match(logText(res), /WARN 注意：第 1 次提醒要在 REMIND_2_DATE（2026-10-16）之前发：今天发出的邮件会显示第二次提醒的文案。请改发 --round 2。/);
  assert.equal(res.previews.length, 1);
  assert.equal(res.previews[0].variant, 'remind1');
  assert.equal(res.previews[0].sendDate, '2026-10-16');
  assert.equal(ctx.fake.opsNamed('SendGiftCardNotification').length, 0);

  // The real renderer shows what Shopify would send on that day: the round-2 copy.
  const real = await remind(ctx, { now: () => ROUND2_DAY, renderPreview: (o) => renderPreviewFor({ ...o, log: memoryLog() }) });
  assert.equal(real.exitCode, 0);
  assert.match(logText(real), /邮件预览（按 2026-10-16 发送时的文案）：.*remind1\.html（主题：.*Last chance/);
});

test('dry run: the preview is rendered for the day the batch would really be sent, with the expiry date the cards carry', async () => {
  const ctx = await issuedCampaign({ count: 1, env: { DRY_RUN: 'true' } });
  const config = { ...ctx.config, giftCardExpiresOn: '2026-10-20' }; // .env changed after select; the cards carry the selection's date
  const previewAt = async (iso, round = 1) => {
    const res = await remind(ctx, { config, round, now: () => new Date(iso) });
    assert.equal(res.exitCode, 0);
    return res.previews[0];
  };
  const early = await previewAt('2026-10-05T18:00:00Z');
  assert.equal(early.sendDate, '2026-10-12', 'a live round 1 cannot go out before REMIND_1_DATE');
  assert.equal(early.expiresOn, '2026-10-19');
  assert.equal(ctx.selection.params.giftCardExpiresOn, '2026-10-19');
  assert.equal((await previewAt('2026-10-13T18:00:00Z')).sendDate, '2026-10-13', 'later than the date: sent today');
  assert.equal((await previewAt('2026-10-13T18:00:00Z', 2)).sendDate, '2026-10-16');
  assert.equal((await previewAt('2026-10-17T18:00:00Z', 2)).sendDate, '2026-10-17');
});

test('test campaigns: the preview uses today, and the notice names the copy today\'s email really shows', async () => {
  const ctx = await issuedCampaign({ count: 1, testCampaign: true, env: { DRY_RUN: 'true' } });
  const run = (round, iso) => remind(ctx, { round, now: () => new Date(iso) });

  const early = await run(1, '2026-10-03T18:00:00Z');
  assert.equal(early.previews[0].sendDate, '2026-10-03');
  assert.match(logText(early), /WARN 注意：.*今天（2026-10-03）发出的是首封文案，不是第一次提醒文案/);

  const between = await run(2, '2026-10-13T17:00:00Z');
  assert.equal(between.previews[0].sendDate, '2026-10-13');
  assert.match(logText(between), /WARN 注意：邮件模板按发送当天的日期切换文案，按 \.env 的日期，今天（2026-10-13）发出的是第一次提醒文案，不是第二次提醒文案；第二次提醒的文案请先用 preview 查看/);
  assert.doesNotMatch(logText(between), /首封文案/, 'between the two dates the test inbox gets the round-1 copy');

  const onTime = await run(1, '2026-10-12T17:00:00Z');
  assert.match(logText(onTime), /INFO 按 \.env 的日期，今天（2026-10-12）发出的是第一次提醒文案/);
  assert.doesNotMatch(logText(onTime), /WARN 注意：邮件模板/);

  const late = await run(1, '2026-10-16T17:00:00Z');
  assert.equal(late.exitCode, 0);
  assert.equal(late.previews[0].sendDate, '2026-10-16');
  assert.match(logText(late), /今天（2026-10-16）发出的是第二次提醒文案，不是第一次提醒文案/);
  assert.doesNotMatch(logText(late), /请改发 --round 2/, 'test campaigns are not date-gated');
});

test('test campaigns remind test accounts whatever their marketing state (as issue does); deleted or email-less ones are still skipped', async () => {
  const ctx = await issuedCampaign({ count: 5, testCampaign: true });
  ctx.customer(1).defaultEmailAddress.marketingState = 'NOT_SUBSCRIBED'; // a colleague who never opted in
  ctx.customer(2).defaultEmailAddress.marketingState = 'UNSUBSCRIBED';
  ctx.customer(3).defaultEmailAddress.validFormat = false;
  ctx.customer(4).defaultEmailAddress = null;
  ctx.fake.state.customers.splice(ctx.fake.state.customers.indexOf(ctx.customer(5)), 1);
  const res = await remind(ctx, { now: () => new Date('2026-10-16T17:00:00Z') }); // a live round 1 on REMIND_2_DATE: not date-gated either
  assert.equal(res.exitCode, 0);
  assert.deepEqual(notified(ctx), [cardGid(2001), cardGid(2002)]);
  assert.deepEqual(res.summary.skipped, { 'no-email': 2, 'customer-deleted': 1 });
  const text = logText(res);
  assert.match(text, /测试活动：不受提醒日期限制，也不检查营销邮件订阅状态（和 issue 一致）/);
  assert.match(text, /正在复查 5 位客户的邮箱…/);

  assert.equal(customerSkip(makeCustomer({ n: 1, marketingState: 'UNSUBSCRIBED' }), { requireSubscribed: false }), null);
  assert.equal(customerSkip(makeCustomer({ n: 1, marketingState: 'PENDING' }), { requireSubscribed: false }), null);
  assert.equal(customerSkip(makeCustomer({ n: 1, email: null }), { requireSubscribed: false }).reason, 'no-email');
  assert.equal(customerSkip(makeCustomer({ n: 1, validFormat: false }), { requireSubscribed: false }).reason, 'no-email');
  assert.equal(customerSkip(null, { requireSubscribed: false }).reason, 'customer-deleted');
  assert.equal(customerSkip(makeCustomer({ n: 1, marketingState: 'UNSUBSCRIBED' })).reason, 'not-subscribed', 'the live campaign keeps the rule');
});

// ---------------------------------------------------------------------------
// Dates during a run (a long run crosses midnight) and the UTC evening
// ---------------------------------------------------------------------------

test('a live round 1 that reaches REMIND_2_DATE (store time) during the run stops before the next send: exit 2, nothing undone', async () => {
  const ctx = await issuedCampaign({ count: 4 });
  const now = clockMovedBySends('2026-10-16T06:58:00Z'); // 23:58 on 10/15 in Los Angeles; each send takes a minute
  const res = await remind(ctx, { now });
  assert.equal(res.exitCode, 2);
  assert.deepEqual(notified(ctx), [cardGid(2001), cardGid(2002)], 'the sends checked at 23:58 and 23:59 went out; at 00:00 the run stopped');
  assertSummaryOf(res.summary, { eligible: 4, planned: 4, attempted: 2, sent: 2, failed: 0, unknown: 0, stopped: null, stoppedByDate: '2026-10-16' });

  const text = logText(res);
  assert.match(text, /ERROR 第 1 次提醒要在 REMIND_2_DATE（2026-10-16）之前发：今天发出的邮件会显示第二次提醒的文案。请改发 --round 2。/);
  assert.match(text, /INFO 现在店铺时间是 2026-10-16 00:00，已停止发送；已处理的人都记在本地日志里/);
  assert.match(text, /INFO 还有 2 人符合条件但没有收到第 1 次提醒：店铺日期已是 2026-10-16，不能再发第 1 次提醒/);
  assert.doesNotMatch(text, /再运行一次会继续/, 'a re-run of round 1 would be refused: no promise that it continues');

  // remind.start carries the instant that was checked; the stopped people have no remind.start at all.
  const starts = remindOps(ctx).filter((e) => e.op === 'remind.start');
  assert.deepEqual(starts.map((e) => [e.cid, e.t]), [[customerGid(1), '2026-10-16T06:58:00.000Z'], [customerGid(2), '2026-10-16T06:59:00.000Z']]);
  assert.deepEqual([1, 2].map((n) => roundState(ctx, n).status), ['sent', 'sent']);
  assert.equal(roundState(ctx, 3), null, 'not "in progress", so never "unknown" later');
  assert.equal(roundState(ctx, 4), null);
  const end = journalOf(ctx).at(-1);
  assert.equal(end.op, 'run.end');
  assert.equal(end.exitCode, 2);
  assert.deepEqual(end.summary, res.summary);
  assert.equal(res.reports.length, 1, 'the Excel is still refreshed');
  acquireRunLock(ctx.paths, 'issue')(); // the lock was released

  // That day round 1 is refused at the start; round 2 reaches everyone, those who missed round 1 included.
  assert.equal((await remind(ctx, { now: () => ROUND2_DAY })).exitCode, 2);
  const round2 = await remind(ctx, { round: 2, now: () => ROUND2_DAY });
  assert.equal(round2.exitCode, 0);
  assert.deepEqual(notified(ctx).slice(2), [2001, 2002, 2003, 2004].map(cardGid));
});

test('nothing reads the clock between the date check and the send: with a clock that ticks on every read, no round-1 email goes out on REMIND_2_DATE', async () => {
  const LA = 'America/Los_Angeles';
  // Three starting seconds, so that some check falls exactly on 23:59:59 whatever the number of reads per send.
  for (const offset of [0, 1, 2]) {
    const ctx = await issuedCampaign({ count: 12 });
    let ms = Date.parse('2026-10-16T06:59:45Z') + offset * 1000; // 23:59:45 (+ offset) on 10/15 in Los Angeles
    const sentOn = [];
    interceptNotify(() => {
      sentOn.push(localDate(ms, LA)); // the store date when Shopify gets the request
      return null;
    });
    const res = await remind(ctx, { now: () => new Date((ms += 1000)) });
    assert.equal(res.exitCode, 2, `offset ${offset}`);
    assert.ok(sentOn.length > 0 && sentOn.length < 12, `offset ${offset}: ${sentOn.length} sends`);
    assert.deepEqual([...new Set(sentOn)], ['2026-10-15'], `offset ${offset}`);
    const starts = remindOps(ctx).filter((e) => e.op === 'remind.start');
    assert.equal(starts.length, sentOn.length);
    assert.ok(starts.every((e) => localDate(Date.parse(e.t), LA) === '2026-10-15'), `offset ${offset}: ${starts.map((e) => e.t).join(', ')}`);
    assert.equal(res.summary.stoppedByDate, '2026-10-16');
  }
});

test('a stop by date can come before the first send (the card lookup ran past midnight); a live round 2 is never stopped by date', async () => {
  const ctx = await issuedCampaign({ count: 2 });
  let ms = Date.parse('2026-10-16T06:59:59Z'); // 23:59:59 on 10/15 in Los Angeles
  const inner = globalThis.fetch;
  globalThis.fetch = async (url, init = {}) => {
    if (/FindCampaignCards/.test(String(init.body ?? ''))) ms += 2_000; // reading the cards takes 2 seconds
    return inner(url, init);
  };
  const res = await remind(ctx, { now: () => new Date(ms) });
  assert.equal(res.exitCode, 2);
  assert.deepEqual(notified(ctx), []);
  assert.equal(res.summary.stoppedByDate, '2026-10-16');
  assert.deepEqual(remindOps(ctx), []);

  const round2 = await remind(ctx, { round: 2, now: clockMovedBySends('2026-10-17T06:59:00Z') }); // 10/16 23:59 → 10/17
  assert.equal(round2.exitCode, 0);
  assert.equal(round2.summary.stoppedByDate, null);
  assert.deepEqual(notified(ctx), [cardGid(2001), cardGid(2002)]);
});

test('a card that expires during the run (it crossed midnight after the expiry day) is skipped as card-expired, and the run goes on', async () => {
  const ctx = await issuedCampaign({
    count: 5,
    tweak(c) {
      c.cards[4].expiresOn = '2026-10-21'; // 5: its expiry was extended in the admin
    },
  });
  const now = clockMovedBySends('2026-10-20T06:58:00Z'); // 23:58 on 10/19 (the expiry day) in Los Angeles
  const res = await remind(ctx, { round: 2, now });
  assert.equal(res.exitCode, 0);
  assert.deepEqual(notified(ctx), [cardGid(2001), cardGid(2002), cardGid(2005)], 'at 00:00 on 10/20 cards 2003 and 2004 had expired');
  // The expired people left the eligible / planned counts and are counted as skipped instead.
  assertSummaryOf(res.summary, { eligible: 3, planned: 3, attempted: 3, sent: 3, skipped: { 'card-expired': 2 }, stoppedByDate: null });

  const skips = remindOps(ctx).filter((e) => e.op === 'remind.skip');
  assert.deepEqual(skips.map((e) => [e.cid, e.round, e.reason, e.detail]), [3, 4].map((n) => [customerGid(n), 2, 'card-expired', '到期日 2026-10-19']));
  assert.ok(skips.every((e) => e.t === '2026-10-20T07:00:00.000Z'));
  assert.equal(roundState(ctx, 3, 2).status, 'skipped');
  assert.equal(roundState(ctx, 3, 2).reason, 'card-expired');
  assert.equal(roundState(ctx, 5, 2).status, 'sent');

  const text = logText(res);
  assert.equal(res.log.lines.filter((l) => l.includes('的卡已过期')).length, 1, 'one warning, not one line per person');
  assert.match(text, /WARN 现在店铺日期是 2026-10-20：#3 First3 Last3 <c3@example\.org> 的卡已过期（到期日 2026-10-19），不发提醒；之后卡已过期的人同样跳过/);
  assert.match(text, /符合条件 3 人；本次发出 3 封/);
  assert.match(text, /跳过 2 人：卡已过期 2/);
  assert.doesNotMatch(text, /还有 \d+ 人符合条件/);

  // The next run finds them expired at the start: the same skip is not journalled twice.
  const again = await remind(ctx, { round: 2, now: () => new Date('2026-10-20T17:00:00Z') });
  assert.equal(again.exitCode, 0);
  assert.deepEqual(again.summary.skipped, { 'card-expired': 2 });
  assert.equal(remindOps(ctx).filter((e) => e.op === 'remind.skip').length, 2);
  assert.equal(notified(ctx).length, 3);
});

test('the expiry is checked before every send in round 1 too', async () => {
  const ctx = await issuedCampaign({
    count: 3,
    tweak(c) {
      c.cards[1].expiresOn = '2026-10-13'; // 2: shortened in the admin
    },
  });
  const res = await remind(ctx, { now: clockMovedBySends('2026-10-14T06:59:00Z') }); // 23:59 on 10/13 in Los Angeles
  assert.equal(res.exitCode, 0);
  assert.deepEqual(notified(ctx), [cardGid(2001), cardGid(2003)]);
  assert.deepEqual(res.summary.skipped, { 'card-expired': 1 });
  assert.equal(roundState(ctx, 2).reason, 'card-expired');
  assert.equal(res.summary.stoppedByDate, null);
});

test('a test campaign is not stopped by date during the run (its round 1 goes on into REMIND_2_DATE), but expired cards are still skipped', async () => {
  const ctx = await issuedCampaign({
    count: 3,
    testCampaign: true,
    tweak(c) {
      c.cards[1].expiresOn = '2026-10-15';
    },
  });
  const res = await remind(ctx, { now: clockMovedBySends('2026-10-16T06:59:00Z') }); // 23:59 on 10/15 in Los Angeles
  assert.equal(res.exitCode, 0);
  assert.deepEqual(notified(ctx), [cardGid(2001), cardGid(2003)]);
  assert.equal(res.summary.stoppedByDate, null);
  assert.deepEqual(res.summary.skipped, { 'card-expired': 1 });
  assert.doesNotMatch(logText(res), /UTC/, 'the UTC evening warning is for the real campaign');
});

test('expirySkip is the expiry rule of cardSkip: a card is usable on its expiry date and expired after it', () => {
  const card = { enabled: true, expiresOn: '2026-10-19', amountCents: 1077, balanceCents: 1077 };
  assert.equal(expirySkip(card, '2026-10-19'), null);
  assert.deepEqual(expirySkip(card, '2026-10-20'), { reason: 'card-expired', detail: '到期日 2026-10-19' });
  assert.deepEqual(expirySkip(card, '2026-10-20'), cardSkip(card, '2026-10-20'));
  assert.equal(expirySkip({ ...card, expiresOn: null }, '2030-01-01'), null);
});

const UTC_EVENING_WARNING = '现在店铺时间 2026-10-15 18:30，UTC 已是 2026-10-16：如果 Shopify 按 UTC 日期选邮件文案，今天发出的提醒会显示第二次提醒文案。建议在洛杉矶时间 17:00 之前运行。';

test('a live round 1 of the real campaign started in the Los Angeles evening, when the UTC date shows the round-2 copy, warns once and still sends', async () => {
  const ctx = await issuedCampaign({ count: 3 });
  const res = await remind(ctx, { now: () => new Date('2026-10-16T01:30:00Z') }); // 18:30 on 10/15 in Los Angeles
  assert.equal(res.exitCode, 0);
  assert.equal(notified(ctx).length, 3, 'only a warning');
  assert.deepEqual(res.log.lines.filter((l) => l.includes('UTC')), [`WARN ${UTC_EVENING_WARNING}`], 'once, however many emails go out');
  const lines = res.log.lines;
  assert.ok(lines.findIndex((l) => l.includes('UTC')) < lines.findIndex((l) => l.includes('开始发送')), 'at the start of the run');

  const at17 = await remind(ctx, { now: () => new Date('2026-10-16T00:00:00Z') }); // 17:00 on 10/15: UTC midnight
  assert.match(logText(at17), /WARN 现在店铺时间 2026-10-15 17:00，UTC 已是 2026-10-16：/);

  // No warning: in the afternoon; when the UTC date shows the same copy; in a dry run (real runs only).
  for (const opts of [
    { now: () => new Date('2026-10-15T23:59:00Z') }, // 16:59 on 10/15
    { now: () => new Date('2026-10-13T01:30:00Z') }, // 18:30 on 10/12: round-1 copy in UTC too
    { round: 2, now: () => new Date('2026-10-17T01:30:00Z') }, // 18:30 on 10/16: round-2 copy either way
    { now: () => new Date('2026-10-16T01:30:00Z'), config: { ...ctx.config, dryRun: true } },
  ]) {
    const r = await remind(ctx, opts);
    assert.equal(r.exitCode, 0);
    assert.deepEqual(r.log.lines.filter((l) => l.includes('UTC')), [], opts.now().toISOString());
  }
});

test('a live round of a test campaign gets no UTC evening warning', async () => {
  const ctx = await issuedCampaign({ count: 1, testCampaign: true });
  const res = await remind(ctx, { now: () => new Date('2026-10-16T01:30:00Z') }); // 18:30 on 10/15 in Los Angeles
  assert.equal(res.exitCode, 0);
  assert.equal(notified(ctx).length, 1);
  assert.doesNotMatch(logText(res), /UTC/);
});

test('utcEveningWarning: only when the UTC date is later than the store date and the template would show another copy on it', () => {
  const t = testConfig();
  teardown.push(t.cleanup);
  const { config } = t;
  const LA = 'America/Los_Angeles';
  const at = (iso, tz = LA, cfg = config) => utcEveningWarning(cfg, Date.parse(iso), tz);
  assert.equal(at('2026-10-16T01:30:00Z'), UTC_EVENING_WARNING);
  assert.equal(at('2026-10-15T23:59:59Z'), null, '16:59:59 in Los Angeles: still 10/15 in UTC');
  assert.match(at('2026-10-16T06:59:00Z'), /^现在店铺时间 2026-10-15 23:59，UTC 已是 2026-10-16：/);
  assert.equal(at('2026-10-16T07:00:00Z'), null, 'after midnight in Los Angeles both dates agree');
  assert.equal(at('2026-10-13T01:30:00Z'), null, 'round-1 copy on both dates');
  assert.equal(at('2026-10-17T01:30:00Z'), null, 'round-2 copy on both dates');
  assert.match(at('2026-10-12T01:30:00Z'), /UTC 已是 2026-10-12：.*今天发出的提醒会显示第一次提醒文案。/, 'the evening before REMIND_1_DATE');
  assert.equal(at('2026-10-15T20:00:00Z', 'Asia/Tokyo'), null, 'east of UTC the UTC date is never the later one');
  assert.equal(at('2026-10-16T01:30:00Z', LA, { ...config, remind2Date: '' }), null, '.env cannot tell without both reminder dates');
});

// ---------------------------------------------------------------------------
// Failures
// ---------------------------------------------------------------------------

test('a lost answer is journalled as unknown, never re-sent automatically, and re-sent once with --retry-unknown', async () => {
  const ctx = await issuedCampaign({ count: 3, failures: { SendGiftCardNotification: [{ kind: 'appliedThenLost' }] } });
  const first = await remind(ctx);
  assert.equal(first.exitCode, 0);
  assert.deepEqual(notified(ctx), [2001, 2002, 2003].map(cardGid), 'the first email did go out; its answer was lost');
  assert.equal(first.summary.unknown, 1);
  assert.equal(first.summary.sent, 2);
  assert.equal(roundState(ctx, 1).status, 'unknown');
  assert.match(remindOps(ctx).find((e) => e.op === 'remind.unknown').error, /socket hang up/);
  assert.match(logText(first), /--retry-unknown/);

  const second = await remind(ctx);
  assert.equal(second.exitCode, 0);
  assert.equal(notified(ctx).length, 3, 'unknown is never re-sent without --retry-unknown');
  assert.equal(second.summary.waitingUnknown, 1);
  assert.equal(second.summary.alreadySent, 2);

  const retry = await remind(ctx, { retryUnknown: true });
  assert.equal(retry.exitCode, 0);
  assert.deepEqual(notified(ctx).slice(3), [cardGid(2001)]);
  assert.equal(retry.summary.retriedUnknown, 1);
  const starts = remindOps(ctx).filter((e) => e.op === 'remind.start' && e.cid === customerGid(1));
  assert.equal(starts.length, 2);
  assert.equal(starts[1].retry, true);
  assert.equal(roundState(ctx, 1).status, 'sent');
  assert.deepEqual(journalOf(ctx).filter((e) => e.op === 'run.start').at(-1).options, { round: 1, retryUnknown: true, retryFailed: false });

  await remind(ctx, { retryUnknown: true });
  assert.equal(notified(ctx).length, 4, 'once sent, --retry-unknown does not send again');
});

test('a send cut off by a crash (remind.start without outcome) becomes unknown and is not re-sent', async () => {
  const ctx = await issuedCampaign({ count: 2 });
  appendJournal(ctx.paths.journal, { op: 'remind.start', cid: customerGid(1), round: 1, giftCardId: cardGid(2001), run: 'crashed-run' }, { now: () => '2026-10-12T16:00:00.000Z' });

  const dry = await remind(ctx, { config: { ...ctx.config, dryRun: true } });
  assert.equal(dry.summary.waitingUnknown, 1);
  assert.equal(roundState(ctx, 1).status, 'in_progress', 'a dry run changes nobody');

  const live = await remind(ctx);
  assert.equal(live.exitCode, 0);
  assert.deepEqual(notified(ctx), [cardGid(2002)]);
  const unknown = remindOps(ctx).filter((e) => e.op === 'remind.unknown');
  assert.equal(unknown.length, 1);
  assert.equal(unknown[0].cid, customerGid(1));
  assert.equal(roundState(ctx, 1).status, 'unknown');
  assert.match(logText(live), /1 人上次发送提醒时中断/);

  await remind(ctx, { retryUnknown: true });
  assert.deepEqual(notified(ctx), [cardGid(2002), cardGid(2001)]);
  assert.equal(roundState(ctx, 1).status, 'sent');
});

test('userErrors → remind.fail and the run continues; a failed reminder is not retried automatically', async () => {
  const ctx = await issuedCampaign({ count: 3 });
  let n = 0;
  const calls = interceptNotify(() => (n++ === 0 ? userErrorsResponse('Gift card is not eligible') : null));
  const res = await remind(ctx);
  assert.equal(res.exitCode, 0);
  assert.equal(res.summary.failed, 1);
  assert.equal(res.summary.sent, 2);
  assert.deepEqual(notified(ctx), [cardGid(2002), cardGid(2003)]);
  const fail = remindOps(ctx).find((e) => e.op === 'remind.fail');
  assert.equal(fail.cid, customerGid(1));
  assert.match(fail.error, /Gift card is not eligible/);
  assert.equal(roundState(ctx, 1).status, 'failed');
  assert.match(logText(res), /#1 First1 Last1 <c1@example.org> 发送失败/);

  const again = await remind(ctx);
  assert.equal(again.exitCode, 0);
  assert.equal(again.summary.previouslyFailed, 1);
  assert.equal(calls.length, 3, 'no new send attempt');
});

test('5 sends in a row without success (failed or unknown) stop the run with exit 1', async () => {
  const ctx = await issuedCampaign({ count: 7, failures: { SendGiftCardNotification: [{ kind: 'network' }, { kind: 'http', status: 502 }] } });
  const plan = ['userError', 'pass', 'userError', 'pass', 'userError'];
  const calls = interceptNotify(() => (plan.shift() === 'userError' ? userErrorsResponse() : null));
  const res = await remind(ctx);
  assert.equal(res.exitCode, 1);
  assert.equal(res.summary.stopped, 'too-many-failures');
  assert.equal(res.summary.failed, 3);
  assert.equal(res.summary.unknown, 2);
  assert.equal(calls.length, 5);
  assert.deepEqual(remindOps(ctx).filter((e) => e.op === 'remind.start').map((e) => e.cid), [1, 2, 3, 4, 5].map(customerGid), 'customers 6 and 7 are not attempted');
  assert.match(logText(res), /连续 5 次没有发送成功，已停止/);
  const end = journalOf(ctx).at(-1);
  assert.equal(end.op, 'run.end');
  assert.equal(end.exitCode, 1);
  assert.equal(res.reports.length, 1);
});

test('a successful send resets the consecutive-failure count', async () => {
  const ctx = await issuedCampaign({ count: 10 });
  const plan = ['userError', 'userError', 'userError', 'userError', 'pass', 'userError', 'userError', 'userError', 'userError', 'pass'];
  interceptNotify(() => (plan.shift() === 'userError' ? userErrorsResponse() : null));
  const res = await remind(ctx);
  assert.equal(res.exitCode, 0);
  assert.equal(res.summary.failed, 8);
  assert.equal(res.summary.sent, 2);
  assert.deepEqual(notified(ctx), [cardGid(2005), cardGid(2010)]);
});

test('throttled out → remind.rejected: nothing was sent, the run stops, a later run reminds that customer', async () => {
  const ctx = await issuedCampaign({ count: 3, failures: { SendGiftCardNotification: throttledOut() } });
  const res = await remind(ctx);
  assert.equal(res.exitCode, 1);
  assert.equal(res.summary.rejected, 1);
  assert.equal(res.summary.stopped, 'rejected');
  assert.equal(notified(ctx).length, 0);
  assert.equal(res.sleeps.length, 5, 'the GraphQL client backed off with the injected sleep');
  assert.deepEqual(remindOps(ctx).map((e) => `${e.op}:${e.cid.split('/').pop()}`), ['remind.start:1', 'remind.rejected:1']);
  assert.equal(roundState(ctx, 1), null, 'rejected returns the row to "not attempted"');
  assert.match(logText(res), /这封没有发出。已停止本次运行/);

  const later = await remind(ctx);
  assert.equal(later.exitCode, 0);
  assert.deepEqual(notified(ctx), [2001, 2002, 2003].map(cardGid));
});

test('--retry-failed: after 5 refusals stopped the run and the cause is fixed, the refused people get their reminder once', async () => {
  const ctx = await issuedCampaign({ count: 7 });
  let refusals = 5;
  const calls = interceptNotify(() => (refusals-- > 0 ? userErrorsResponse('Notifications are temporarily unavailable') : null));
  const stopped = await remind(ctx);
  assert.equal(stopped.exitCode, 1);
  assert.equal(stopped.summary.stopped, 'too-many-failures');
  assert.match(logText(stopped), /ERROR 连续 5 次没有发送成功，已停止。请查明原因后再运行.*原因解决后可用 --retry-failed 重发/);
  assert.match(logText(stopped), /WARN 本次发送失败 5 人.*原因解决后可用 --retry-failed 重发/);
  assert.deepEqual(notified(ctx), []);

  // Without the flag (and with --retry-unknown, which is about lost answers) refused rows stay refused.
  const plain = await remind(ctx);
  assert.equal(plain.exitCode, 0);
  assert.deepEqual(notified(ctx), [cardGid(2006), cardGid(2007)]);
  assert.equal(plain.summary.previouslyFailed, 5);
  assert.match(logText(plain), /这一轮之前发送失败：5 人（不会自动重试，原因解决后可用 --retry-failed 重发）/);
  assert.equal((await remind(ctx, { retryUnknown: true })).summary.previouslyFailed, 5);
  assert.equal(notified(ctx).length, 2);

  // Meanwhile 3 used the card and 4 unsubscribed: a retried row still passes every card and customer check.
  ctx.fake.spendCard(cardGid(2003), 100);
  ctx.customer(4).defaultEmailAddress.marketingState = 'UNSUBSCRIBED';
  const dry = await remind(ctx, { retryFailed: true, config: { ...ctx.config, dryRun: true } });
  assert.equal(dry.summary.planned, 3);
  assert.match(logText(dry), /会发 3 人，序号 1–5（其中 3 人是之前发送失败的重发）/);
  assert.match(logText(dry), /#1 First1 Last1 <c1@example.org> .*（之前发送失败，重发）/);
  assert.equal(notified(ctx).length, 2, 'a dry run sends nothing');

  const retry = await remind(ctx, { retryFailed: true });
  assert.equal(retry.exitCode, 0);
  assert.deepEqual(notified(ctx).slice(2), [cardGid(2001), cardGid(2002), cardGid(2005)]);
  assertSummaryOf(retry.summary, { sent: 3, retriedFailed: 3, retriedUnknown: 0, alreadySent: 2, previouslyFailed: 0, failed: 0 });
  assert.deepEqual(retry.summary.skipped, { used: 1, 'not-subscribed': 1 });
  const journal = journalOf(ctx);
  const lastStart = journal.findLastIndex((e) => e.op === 'run.start'); // run ids repeat under the fixed test clock
  assert.deepEqual(journal[lastStart].options, { round: 1, retryUnknown: false, retryFailed: true });
  const starts = journal.slice(lastStart).filter((e) => e.op === 'remind.start');
  assert.deepEqual(starts.map((e) => [e.cid, e.retry]), [1, 2, 5].map((n) => [customerGid(n), true]));
  for (const n of [1, 2, 5]) assert.equal(roundState(ctx, n).status, 'sent');
  assert.match(logText(retry), /这次也会重发：之前发送失败的人（--retry-failed）/);
  assert.match(logText(retry), /本次重发了之前发送失败的 3 人（--retry-failed）/);

  // Once sent, --retry-failed never sends again.
  const again = await remind(ctx, { retryFailed: true });
  assert.equal(again.summary.attempted, 0);
  assert.equal(notified(ctx).length, 5);
  assert.equal(calls.length, 5 + 2 + 3, 'every send request, refused or not');
});

test('--retry-failed re-sends only refused reminders: an unknown outcome still waits for --retry-unknown', async () => {
  const ctx = await issuedCampaign({ count: 3, failures: { SendGiftCardNotification: [{ kind: 'appliedThenLost' }] } });
  let n = 0;
  interceptNotify(() => (n++ === 1 ? userErrorsResponse('Gift card is not eligible') : null));
  const first = await remind(ctx); // 1: answer lost (unknown) · 2: refused (failed) · 3: sent
  assert.deepEqual([1, 2, 3].map((k) => roundState(ctx, k).status), ['unknown', 'failed', 'sent']);

  const retryFailed = await remind(ctx, { retryFailed: true });
  assert.equal(retryFailed.exitCode, 0);
  assert.deepEqual(notified(ctx), [cardGid(2001), cardGid(2003), cardGid(2002)]);
  assertSummaryOf(retryFailed.summary, { retriedFailed: 1, retriedUnknown: 0, waitingUnknown: 1, alreadySent: 1 });

  const retryUnknown = await remind(ctx, { retryUnknown: true });
  assert.deepEqual(notified(ctx).slice(3), [cardGid(2001)]);
  assertSummaryOf(retryUnknown.summary, { retriedFailed: 0, retriedUnknown: 1, alreadySent: 2 });
  assert.equal(first.summary.failed, 1);
});

test('the list is read again under the run lock: if select rewrote it in between, nothing is sent (exit 1)', async () => {
  const ctx = await issuedCampaign({ count: 2 });
  const before = journalOf(ctx).length;
  const v2 = await selectionFixture(ctx.config, { customers: eligibleCustomers(2), createdAt: '2026-10-02T20:00:00.000Z', write: false });
  let clockReads = 0;
  const res = await remind(ctx, {
    now: () => {
      clockReads += 1;
      // The first clock read is the date guard: after the list was read, before the lock is taken.
      if (clockReads === 1) writeJsonAtomic(ctx.paths.selection, v2);
      return ROUND1_DAY;
    },
  });
  assert.equal(res.exitCode, 1);
  assert.equal(res.summary, null);
  assert.match(logText(res), new RegExp(`ERROR ${SELECTION_REPLACED}`));
  assert.equal(SELECTION_REPLACED, '名单刚被 select 改写，请重新运行');
  assert.equal(ctx.fake.state.calls.token.length, 0);
  assert.deepEqual(notified(ctx), []);
  assert.equal(journalOf(ctx).length, before);
  assert.equal(res.reports.length, 0);
  acquireRunLock(ctx.paths, 'issue')(); // the lock was released

  const next = await remind(ctx); // the next run works from the list on disk
  assert.equal(next.exitCode, 0);
  assert.deepEqual(notified(ctx), [cardGid(2001), cardGid(2002)]);
});

test('warnings of the Excel report are printed once: writeReport logs them itself', async () => {
  const ctx = await issuedCampaign({ count: 1 });
  const warning = 'Excel 正打开此文件，请关闭后重新打开才能看到最新内容';
  const res = await remind(ctx, {
    writeReport: async ({ log }) => {
      log.warn(warning);
      return { file: ctx.paths.excel, out: null, warnings: [warning] };
    },
  });
  assert.equal(res.exitCode, 0);
  assert.equal(res.log.lines.filter((l) => l.includes(warning)).length, 1);
  assert.match(logText(res), /Excel 已更新：/);
});

// ---------------------------------------------------------------------------
// Round tags: Shopify's own record of each round (duplicates-reminders N1:
// the local journal alone can be lost or replaced by an older copy)
// ---------------------------------------------------------------------------

const LIVE_TAG = 'OCT26RTPROMO';
const R1 = 'OCT26RTPROMO-R1';
const minutesAfterRound1 = (m) => new Date(ROUND1_DAY.getTime() + m * 60_000);
/** Customers (gids) the fake store lists with `tag`, in store order. */
const taggedWith = (ctx, tag) => ctx.fake.state.customers.filter((c) => c.tags.includes(tag)).map((c) => c.id);
/** Every TagsAdd request as [customer gid, ...tags]. */
const tagCalls = (ctx) => ctx.fake.opsNamed('TagsAdd').map((o) => [o.variables.id, ...o.variables.tags]);
const removeTag = (ctx, n, tag) => {
  const c = ctx.customer(n);
  c.tags = c.tags.filter((t) => t !== tag);
};
/** The listed fields of a round state from the journal. */
const roundFields = (ctx, n, fields, round = 1) => Object.fromEntries(fields.map((f) => [f, roundState(ctx, n, round)?.[f]]));

test('roundTag: SENT_TAG-R<round>, so each round and the test campaign have their own tag', () => {
  assert.equal(roundTag('OCT26RTPROMO', 1), 'OCT26RTPROMO-R1');
  assert.equal(roundTag('OCT26RTPROMO', 2), 'OCT26RTPROMO-R2');
  assert.equal(roundTag('OCT26RTPROMO', '2'), 'OCT26RTPROMO-R2');
  assert.equal(roundTag('OCT26RTPROMO-TEST', 1), 'OCT26RTPROMO-TEST-R1');
});

test('the run first reads who carries the round tag; every reminder sent is followed by the tag, with no journal line for it', async () => {
  const ctx = await issuedCampaign({ count: 3, env: { SENT_TAG: LIVE_TAG } });
  ctx.fake.spendCard(cardGid(2002), 100); // 2: used, so neither reminded nor tagged
  const res = await remind(ctx);
  assert.equal(res.exitCode, 0);
  assert.deepEqual(notified(ctx), [cardGid(2001), cardGid(2003)]);
  assert.deepEqual(ctx.fake.opsNamed('TaggedCustomers').map((o) => o.variables.query), [`tag:"${R1}"`]);
  assert.deepEqual(tagCalls(ctx), [[customerGid(1), R1], [customerGid(3), R1]]);
  assert.deepEqual(taggedWith(ctx, R1), [customerGid(1), customerGid(3)]);
  // The tag set is read before anything is evaluated; each tag right after its send.
  const watched = new Set(['TaggedCustomers', 'FindCampaignCards', 'SendGiftCardNotification', 'TagsAdd']);
  assert.deepEqual(ctx.fake.state.calls.ops.map((o) => o.op).filter((op) => watched.has(op)),
    ['TaggedCustomers', 'FindCampaignCards', 'SendGiftCardNotification', 'TagsAdd', 'SendGiftCardNotification', 'TagsAdd']);
  assertSummaryOf(res.summary, { sent: 2, roundTagged: 2, roundTagFailed: 0, roundTagFixed: 0, roundTagMissing: 0, alreadySentByTag: 0 });
  assert.deepEqual(remindOps(ctx).map((e) => `${e.op}:${e.cid.split('/').pop()}`), ['remind.skip:2', 'remind.start:1', 'remind.ok:1', 'remind.start:3', 'remind.ok:3']);

  const text = logText(res);
  assert.match(text, /INFO 本轮 tag：OCT26RTPROMO-R1（每发出一封提醒就给客户打上；带这个 tag 的人这一轮不会再收到提醒）。Shopify 上现在有 0 位客户带这个 tag/);
  assert.match(text, /INFO 本轮 tag OCT26RTPROMO-R1：按 tag 认定已发 0 人，打 tag 2 人，打 tag 失败 0 人，补打 tag 0 人/);
  assert.match(text, /INFO 开始发送第 1 次提醒：本次 2 人，预计不到 1 分钟/);
  assert.doesNotMatch(text, /already carry tag/, 'the tag lookup\'s own English line is not printed');
  assert.doesNotMatch(text, /WARN .*tag/);
});

test('journal lost after round 1 (verify rebuilds the cards): round 1 again sends nothing, everyone counts as sent by the round tag', async () => {
  const ctx = await issuedCampaign({ count: 3, env: { SENT_TAG: LIVE_TAG } });
  assert.equal((await remind(ctx)).exitCode, 0);
  assert.equal(notified(ctx).length, 3);

  // 10/13: the journal is gone; verify records the cards again from Shopify, but no reminder.
  fs.rmSync(ctx.paths.journal);
  const day2 = new Date('2026-10-13T17:00:00Z');
  const v = await runVerify({ config: ctx.config, log: memoryLog(), now: () => day2, writeReport: async () => ({ file: 'x', warnings: [] }) });
  assert.equal(v.exitCode, 0);
  assert.deepEqual([1, 2, 3].map((n) => roundState(ctx, n)), [null, null, null]);

  const again = await remind(ctx, { now: () => day2 });
  assert.equal(again.exitCode, 0);
  assert.equal(notified(ctx).length, 3, 'no customer gets round 1 twice');
  assertSummaryOf(again.summary, { alreadySentByTag: 3, alreadySent: 0, eligible: 0, attempted: 0, sent: 0, roundTagged: 0, roundTagFixed: 0, skipped: {} });
  assert.equal(ctx.fake.opsNamed('RefreshCustomers').length, 1, 'people carrying the tag are not even re-read');
  assert.equal(ctx.fake.opsNamed('TagsAdd').length, 3, 'only the first run tagged');
  const runId = journalOf(ctx).filter((e) => e.op === 'run.start').at(-1).run;
  assert.deepEqual(remindOps(ctx).map((e) => [e.op, e.cid, e.round, e.source, e.run]), [1, 2, 3].map((n) => ['remind.found', customerGid(n), 1, 'tag', runId]));
  for (const n of [1, 2, 3]) {
    assert.deepEqual(roundFields(ctx, n, ['status', 'source', 'at', 'giftCardId']), { status: 'sent', source: 'tag', at: null, giftCardId: null }, `customer ${n}`);
  }
  const text = logText(again);
  assert.match(text, /Shopify 上现在有 3 位客户带这个 tag/);
  assert.match(text, /WARN 按 tag 认定已发 3 人：他们在 Shopify 上带本轮 tag OCT26RTPROMO-R1，本地日志里却没有这一轮发出的记录（日志丢失、被旧的备份覆盖，或在后台手动打了 tag）；已在本地日志补记为已发，这一轮不会再给他们发/);
  assert.match(text, /INFO 本轮 tag OCT26RTPROMO-R1：按 tag 认定已发 3 人，打 tag 0 人，打 tag 失败 0 人，补打 tag 0 人/);

  // From now on the journal knows: nothing new is recorded, and these rows are never "repaired".
  const third = await remind(ctx, { now: () => new Date('2026-10-14T17:00:00Z') });
  assertSummaryOf(third.summary, { alreadySent: 3, alreadySentByTag: 0, roundTagFixed: 0, roundTagMissing: 0 });
  assert.equal(remindOps(ctx).length, 3);
  assert.equal(notified(ctx).length, 3);
});

test('journal replaced by an older copy taken before round 1: round 1 again sends nothing', async () => {
  const ctx = await issuedCampaign({ count: 4 });
  const backup = fs.readFileSync(ctx.paths.journal); // e.g. the 10/11 backup: issue only
  ctx.fake.spendCard(cardGid(2004), 1077); // 4 used the card: never reminded, never tagged
  const first = await remind(ctx);
  assert.equal(first.exitCode, 0);
  assert.deepEqual(notified(ctx), [2001, 2002, 2003].map(cardGid));

  fs.writeFileSync(ctx.paths.journal, backup); // campaigns/2026-10 restored from that backup
  const again = await remind(ctx, { now: () => minutesAfterRound1(60) });
  assert.equal(again.exitCode, 0);
  assert.equal(notified(ctx).length, 3, 'nobody gets round 1 twice');
  assertSummaryOf(again.summary, { alreadySentByTag: 3, alreadySent: 0, attempted: 0, sent: 0 });
  assert.deepEqual(again.summary.skipped, { used: 1 }, 'customer 4 (no tag) is evaluated as before');
  assert.deepEqual([1, 2, 3, 4].map((n) => roundState(ctx, n)?.status), ['sent', 'sent', 'sent', 'skipped']);
});

test('an unknown, failed or interrupted row whose customer carries the round tag counts as sent: never re-sent, not even with --retry-unknown / --retry-failed', async () => {
  const ctx = await issuedCampaign({ count: 4, failures: { SendGiftCardNotification: [{ kind: 'appliedThenLost' }] } });
  let n = 0;
  const calls = interceptNotify(() => (n++ === 1 ? userErrorsResponse() : null)); // 1: answer lost · 2: refused · 3: sent
  const first = await remind(ctx, { limit: 3 });
  assert.equal(first.exitCode, 0);
  assert.deepEqual([1, 2, 3].map((k) => roundState(ctx, k).status), ['unknown', 'failed', 'sent']);
  const tag = roundTag(ctx.config.sentTag, 1);
  assert.deepEqual(taggedWith(ctx, tag), [customerGid(3)], 'only a confirmed send is tagged');
  assert.match(logText(first), /WARN 结果不明 1 人：不会自动重发。确认对方确实没收到后，可加 --retry-unknown 补发；确认已经收到的，可以在 Shopify 后台给他加上 tag gift-card-sent-2026-10-R1，下次运行会记为已发/);
  // 4: a send cut off by a crash (remind.start without outcome).
  appendJournal(ctx.paths.journal, { op: 'remind.start', cid: customerGid(4), round: 1, giftCardId: cardGid(2004), run: 'crashed-run' }, { now: () => minutesAfterRound1(1).toISOString() });
  // Their reminders were sent after all (a journal that was lost since), so Shopify shows the tag.
  for (const k of [1, 2, 4]) ctx.customer(k).tags.push(tag);

  const retry = await remind(ctx, { retryUnknown: true, retryFailed: true, now: () => minutesAfterRound1(30) });
  assert.equal(retry.exitCode, 0);
  assert.equal(calls.length, 3, 'no new send request');
  assertSummaryOf(retry.summary, { alreadySentByTag: 3, alreadySent: 1, attempted: 0, retriedUnknown: 0, retriedFailed: 0, waitingUnknown: 0, previouslyFailed: 0 });
  for (const k of [1, 2, 4]) assert.deepEqual(roundFields(ctx, k, ['status', 'source']), { status: 'sent', source: 'tag' }, `customer ${k}`);
  assert.equal(remindOps(ctx).filter((e) => e.op === 'remind.unknown' && e.cid === customerGid(4)).length, 0, 'the interrupted send is not marked unknown first');
  assert.equal(ctx.fake.opsNamed('TagsAdd').length, 1, 'nothing to tag again');
});

test('a round tag that fails after its send is journalled (remind.tag.fail) and warned about; the run goes on and never stops for it', async () => {
  const locked = Array.from({ length: 6 }, () => ({ kind: 'userError', message: 'Tags are locked' }));
  const ctx = await issuedCampaign({ count: 7, env: { SENT_TAG: LIVE_TAG }, failures: { TagsAdd: locked } });
  const res = await remind(ctx);
  assert.equal(res.exitCode, 0);
  assert.equal(notified(ctx).length, 7, 'six failed tags in a row do not stop the run: the stop is about sends');
  assertSummaryOf(res.summary, { sent: 7, failed: 0, unknown: 0, roundTagged: 1, roundTagFailed: 6, stopped: null });
  const runId = journalOf(ctx).filter((e) => e.op === 'run.start').at(-1).run;
  const fails = remindOps(ctx).filter((e) => e.op === 'remind.tag.fail');
  assert.deepEqual(fails.map((e) => [e.cid, e.round, e.run]), [1, 2, 3, 4, 5, 6].map((k) => [customerGid(k), 1, runId]));
  assert.equal(fails[0].error, 'tagsAdd rejected: input: Tags are locked [INVALID]');
  for (const k of [1, 2, 3, 4, 5, 6]) {
    assert.deepEqual(roundFields(ctx, k, ['status', 'tagError']), { status: 'sent', tagError: 'tagsAdd rejected: input: Tags are locked [INVALID]' }, `customer ${k}`);
  }
  assert.deepEqual(roundFields(ctx, 7, ['status', 'tagError']), { status: 'sent', tagError: undefined });
  // Each tag.fail follows its own remind.ok.
  const ops = remindOps(ctx).map((e) => `${e.op}:${e.cid.split('/').pop()}`);
  assert.deepEqual(ops.slice(0, 6), ['remind.start:1', 'remind.ok:1', 'remind.tag.fail:1', 'remind.start:2', 'remind.ok:2', 'remind.tag.fail:2']);

  const text = logText(res);
  assert.match(text, /WARN {3}#1 First1 Last1 <c1@example\.org>：提醒已发出，但打本轮 tag OCT26RTPROMO-R1 失败（tagsAdd rejected: input: Tags are locked \[INVALID\]）；本地日志记着已发，不会重发，下次运行会再试，如果仍被拒请在后台查看这位客户/);
  assert.match(text, /INFO 本轮 tag OCT26RTPROMO-R1：按 tag 认定已发 0 人，打 tag 1 人，打 tag 失败 6 人，补打 tag 0 人/);
  assert.match(text, /WARN 打本轮 tag 失败 6 人：提醒已经发出，本地日志记着已发，不会重发；下次运行 remind --round 1 会补打 tag/);
  assert.doesNotMatch(text, /连续 5 次/);
});

test('a failed round tag does not count toward the 5-failure stop: 4 refused sends, a send whose tag fails, 4 refused sends, the run goes on', async () => {
  const ctx = await issuedCampaign({ count: 10, failures: { TagsAdd: [{ kind: 'userError', message: 'locked' }, { kind: 'userError', message: 'locked' }] } });
  const plan = ['userError', 'userError', 'userError', 'userError', 'pass', 'userError', 'userError', 'userError', 'userError', 'pass'];
  interceptNotify(() => (plan.shift() === 'userError' ? userErrorsResponse() : null));
  const res = await remind(ctx);
  assert.equal(res.exitCode, 0, 'counted as a failure, the failed tag of send 5 would have made it 5 in a row with sends 6–9');
  assertSummaryOf(res.summary, { failed: 8, sent: 2, roundTagged: 0, roundTagFailed: 2, stopped: null });
  assert.deepEqual(notified(ctx), [cardGid(2005), cardGid(2010)]);
});

test('a round tag throttled out (or lost on the network) after the send is only a warning: the run goes on', async () => {
  const ctx = await issuedCampaign({ count: 3, failures: { TagsAdd: [...throttledOut(), { kind: 'network' }] } });
  const res = await remind(ctx);
  assert.equal(res.exitCode, 0);
  assert.equal(notified(ctx).length, 3);
  assertSummaryOf(res.summary, { sent: 3, rejected: 0, roundTagged: 1, roundTagFailed: 2, stopped: null });
  assert.equal(res.sleeps.length, 5, 'the throttled tagsAdd was retried by the GraphQL client first');
  assert.match(roundState(ctx, 1).tagError, /Throttled by Shopify 6 times in a row/);
  assert.match(roundState(ctx, 2).tagError, /Network error calling Shopify/);
});

test('the next live run adds missing round tags before it sends: a failed tag at once, a sent row without the tag once its send is 10 minutes old', async () => {
  const ctx = await issuedCampaign({ count: 4, env: { SENT_TAG: LIVE_TAG }, failures: { TagsAdd: [{ kind: 'userError', message: 'Tags are locked' }] } });
  assert.equal((await remind(ctx, { limit: 3 })).exitCode, 0); // 10:00 — 1's tag fails; 2 and 3 are tagged
  assert.deepEqual(taggedWith(ctx, R1), [customerGid(2), customerGid(3)]);
  removeTag(ctx, 3, R1); // 3's tag is not in the search yet, or someone removed it in the admin

  // 10:05: 1's tag is known to have failed, so it is added again at once; 3 was sent only 5 minutes ago.
  const soon = await remind(ctx, { now: () => minutesAfterRound1(5) });
  assert.equal(soon.exitCode, 0);
  assert.deepEqual(tagCalls(ctx).slice(3), [[customerGid(1), R1], [customerGid(4), R1]], 'then customer 4 is sent and tagged as usual');
  assertSummaryOf(soon.summary, { roundTagFixed: 1, roundTagged: 1, roundTagFailed: 0, sent: 1, alreadySent: 3 });
  const watched = new Set(['TagsAdd', 'FindCampaignCards', 'SendGiftCardNotification']);
  const order = ctx.fake.state.calls.ops.map((o) => o.op).filter((op) => watched.has(op));
  assert.deepEqual(order.slice(order.lastIndexOf('FindCampaignCards') - 1), ['TagsAdd', 'FindCampaignCards', 'SendGiftCardNotification', 'TagsAdd'], 'repairs come before the card lookup');
  const fixed = remindOps(ctx).filter((e) => e.op === 'remind.tag.ok');
  assert.deepEqual(fixed.map((e) => [e.cid, e.round, e.note]), [[customerGid(1), 1, undefined]]);
  assert.equal(roundState(ctx, 1).tagError, null, 'the repair clears the recorded failure');
  assert.match(logText(soon), /INFO 补打本轮 tag：1 人这一轮已发（见本地日志），Shopify 上还没有 tag OCT26RTPROMO-R1/);
  assert.match(logText(soon), /补打 tag 1 人/);

  // 10:10: 3's send is 10 minutes old and the search still does not show its tag.
  const later = await remind(ctx, { now: () => minutesAfterRound1(10) });
  assert.deepEqual(tagCalls(ctx).slice(5), [[customerGid(3), R1]]);
  assertSummaryOf(later.summary, { roundTagFixed: 1, sent: 0, alreadySent: 4 });
  assert.deepEqual(taggedWith(ctx, R1), [1, 2, 3, 4].map(customerGid));

  const quiet = await remind(ctx, { now: () => minutesAfterRound1(60) });
  assert.equal(tagCalls(ctx).length, 6, 'nothing is missing any more');
  assertSummaryOf(quiet.summary, { roundTagFixed: 0, roundTagFailed: 0 });
  assert.equal(notified(ctx).length, 4, 'repairs never send');
});

test('a repair that fails again is recorded again; the run still sends, and a later run repairs it', async () => {
  const ctx = await issuedCampaign({ count: 2, failures: { TagsAdd: [{ kind: 'userError', message: 'locked' }, { kind: 'userError', message: 'still locked' }] } });
  await remind(ctx, { limit: 1 }); // 1 sent, tag fails
  const second = await remind(ctx, { now: () => minutesAfterRound1(20) }); // repair fails again; 2 is sent and tagged
  assert.equal(second.exitCode, 0);
  assertSummaryOf(second.summary, { roundTagFixed: 0, roundTagFailed: 1, roundTagged: 1, sent: 1 });
  assert.match(roundState(ctx, 1).tagError, /still locked/);
  assert.match(logText(second), /WARN {3}#1 First1 Last1 <c1@example\.org>：打本轮 tag gift-card-sent-2026-10-R1 失败（tagsAdd rejected: input: still locked \[INVALID\]）；本地日志记着已发，不会重发，下次运行会再试，如果仍被拒请在后台查看这位客户/);
  const third = await remind(ctx, { now: () => minutesAfterRound1(40) });
  assertSummaryOf(third.summary, { roundTagFixed: 1, roundTagFailed: 0 });
  assert.equal(roundState(ctx, 1).tagError, null);
  assert.equal(notified(ctx).length, 2);
});

test('a round tag whose answer was lost although Shopify added it is not added again; the next live run clears the recorded failure', async () => {
  const ctx = await issuedCampaign({ count: 2, failures: { TagsAdd: [{ kind: 'appliedThenLost' }] } });
  await remind(ctx);
  const tag = roundTag(ctx.config.sentTag, 1);
  assert.match(roundState(ctx, 1).tagError, /socket hang up/);
  assert.deepEqual(taggedWith(ctx, tag), [customerGid(1), customerGid(2)], 'Shopify did add it');

  const next = await remind(ctx, { now: () => minutesAfterRound1(1) });
  assert.equal(ctx.fake.opsNamed('TagsAdd').length, 2, 'not added again');
  assertSummaryOf(next.summary, { roundTagFixed: 0, roundTagFailed: 0, alreadySent: 2 });
  const ok = remindOps(ctx).filter((e) => e.op === 'remind.tag.ok');
  assert.deepEqual(ok.map((e) => [e.cid, e.note]), [[customerGid(1), 'already tagged in Shopify']]);
  assert.equal(roundState(ctx, 1).tagError, null);
});

test('Ctrl+C during the tag repair stops before the next tag and before any send (exit 130)', async () => {
  const ctx = await issuedCampaign({ count: 3, failures: { TagsAdd: [{ kind: 'userError', message: 'locked' }, { kind: 'userError', message: 'locked' }] } });
  await remind(ctx, { limit: 2 }); // 1 and 2 sent, both tags fail
  const controller = new AbortController();
  const inner = globalThis.fetch;
  globalThis.fetch = async (url, init = {}) => {
    const response = await inner(url, init);
    if (/mutation TagsAdd/.test(String(init.body ?? ''))) controller.abort(); // Ctrl+C during the first repair
    return response;
  };
  const res = await remind(ctx, { signal: controller.signal, now: () => minutesAfterRound1(20) });
  assert.equal(res.exitCode, 130);
  assert.equal(res.summary.stopped, 'aborted');
  assert.equal(res.summary.roundTagFixed, 1);
  assert.equal(notified(ctx).length, 2, 'customer 3 is not sent in this run');
  assert.equal(ctx.fake.opsNamed('FindCampaignCards').length, 1, 'stopped before the card lookup');
  assert.equal(roundState(ctx, 1).tagError, null);
  assert.match(roundState(ctx, 2).tagError, /locked/);
});

test('without the round tag set nothing is sent: a failed tag lookup stops the run (exit 1)', async () => {
  const ctx = await issuedCampaign({ count: 2, failures: { TaggedCustomers: [{ kind: 'network' }] } });
  const res = await remind(ctx);
  assert.equal(res.exitCode, 1);
  assert.equal(res.summary.stopped, 'error');
  assert.deepEqual(notified(ctx), []);
  assert.equal(ctx.fake.opsNamed('FindCampaignCards').length, 0);
  assert.match(logText(res), /ERROR 出错，已停止：Network error calling Shopify: fetch failed/);
  assert.deepEqual(remindOps(ctx), []);
});

test('dry run: reads the round tag set, adds no tag, journals only run.start/run.end, and tells what a live run would do with the tags', async () => {
  const ctx = await issuedCampaign({ count: 6, env: { SENT_TAG: LIVE_TAG }, failures: { TagsAdd: [{ kind: 'userError', message: 'locked' }] } });
  assert.equal((await remind(ctx, { limit: 4 })).exitCode, 0); // 1–4 sent; 1's tag failed
  ctx.customer(5).tags.push(R1); // 5's reminder was recorded only in a journal that is lost now
  const tagsBefore = ctx.fake.opsNamed('TagsAdd').length;
  const before = journalOf(ctx).length;
  const stateBefore = foldJournal(journalOf(ctx));

  const dry = await remind(ctx, { config: { ...ctx.config, dryRun: true }, now: () => minutesAfterRound1(30) });
  assert.equal(dry.exitCode, 0);
  assert.equal(ctx.fake.opsNamed('TaggedCustomers').length, 2, 'read like a live run');
  assert.equal(ctx.fake.opsNamed('TagsAdd').length, tagsBefore, 'zero TagsAdd');
  assert.equal(ctx.fake.opsNamed('SendGiftCardNotification').length, 4);
  assert.deepEqual(journalOf(ctx).slice(before).map((e) => e.op), ['run.start', 'run.end']);
  assert.deepEqual(foldJournal(journalOf(ctx)).customers, stateBefore.customers, 'nobody\'s state changes');
  assertSummaryOf(dry.summary, { alreadySent: 4, alreadySentByTag: 1, roundTagMissing: 1, roundTagFixed: 0, roundTagged: 0, roundTagFailed: 0, planned: 1 });
  const text = logText(dry);
  assert.match(text, /Shopify 上现在有 4 位客户带这个 tag/);
  assert.match(text, /INFO \[预演\] 缺本轮 tag 1 人：本地日志记着这一轮已发，Shopify 上却没有 tag OCT26RTPROMO-R1；真实运行会先给他们补打 tag/);
  assert.match(text, /WARN \[预演\] 按 tag 认定已发 1 人：他们在 Shopify 上带本轮 tag OCT26RTPROMO-R1，.*真实运行会在本地日志补记为已发，这一轮不会再给他们发/);
  assert.match(text, /INFO \[预演\] 真实发送预计不到 1 分钟（每封约 0\.7 秒：发信、打本轮 tag、写 2 行日志）/);
  assert.match(text, /INFO 本轮 tag OCT26RTPROMO-R1：按 tag 认定已发 1 人；真实运行会打 tag 1 人（每发出一封打一次），补打 tag 1 人/);

  // The live run does exactly that.
  const live = await remind(ctx, { now: () => minutesAfterRound1(31) });
  assertSummaryOf(live.summary, { alreadySent: 4, alreadySentByTag: 1, roundTagFixed: 1, roundTagged: 1, sent: 1 });
  assert.deepEqual(notified(ctx).slice(4), [cardGid(2006)]);
});

test('round 1 and round 2 have their own tags, and so does the test campaign: another round\'s or campaign\'s tag is ignored', async () => {
  const ctx = await issuedCampaign({ count: 3, env: { SENT_TAG: LIVE_TAG } });
  for (const c of ctx.fake.state.customers) c.tags.push(LIVE_TAG); // every recipient carries SENT_TAG itself
  ctx.customer(3).tags.push('OCT26RTPROMO-TEST-R1'); // 3 also took part in the test campaign
  const first = await remind(ctx);
  assert.equal(first.summary.alreadySentByTag, 0, 'neither SENT_TAG nor the test campaign\'s round tag counts');
  assert.deepEqual(notified(ctx), [2001, 2002, 2003].map(cardGid));
  ctx.customer(2).tags.push('OCT26RTPROMO-R2'); // 2's round 2 was recorded only in a lost journal

  const second = await remind(ctx, { round: 2, now: () => ROUND2_DAY });
  assert.equal(second.exitCode, 0);
  assert.deepEqual(notified(ctx).slice(3), [cardGid(2001), cardGid(2003)], 'the round-1 tag does not stop round 2');
  assertSummaryOf(second.summary, { alreadySentByTag: 1, alreadySent: 0, sent: 2, roundTagged: 2 });
  assert.deepEqual(ctx.fake.opsNamed('TaggedCustomers').map((o) => o.variables.query), ['tag:"OCT26RTPROMO-R1"', 'tag:"OCT26RTPROMO-R2"']);
  assert.deepEqual(tagCalls(ctx).slice(3), [[customerGid(1), 'OCT26RTPROMO-R2'], [customerGid(3), 'OCT26RTPROMO-R2']]);
  assert.deepEqual(roundFields(ctx, 2, ['status', 'source'], 2), { status: 'sent', source: 'tag' });
  assert.deepEqual(roundFields(ctx, 2, ['status', 'source'], 1), { status: 'sent', source: undefined }, 'round 1 is untouched');
  assert.match(logText(second), /本轮 tag：OCT26RTPROMO-R2/);

  // The test campaign (SENT_TAG OCT26RTPROMO-TEST) uses OCT26RTPROMO-TEST-R1 and ignores the live campaign's tags.
  const testCtx = await issuedCampaign({ count: 2, testCampaign: true });
  testCtx.customer(1).tags.push('OCT26RTPROMO-R1');
  const testRun = await remind(testCtx);
  assert.equal(testRun.exitCode, 0);
  assert.deepEqual(notified(testCtx), [cardGid(2001), cardGid(2002)]);
  assert.equal(testRun.summary.alreadySentByTag, 0);
  assert.deepEqual(testCtx.fake.opsNamed('TaggedCustomers').map((o) => o.variables.query), ['tag:"OCT26RTPROMO-TEST-R1"']);
  assert.deepEqual(tagCalls(testCtx), [[customerGid(1), 'OCT26RTPROMO-TEST-R1'], [customerGid(2), 'OCT26RTPROMO-TEST-R1']]);
});

// ---------------------------------------------------------------------------
// Dry run, --limit, Ctrl+C, report, lock, inputs
// ---------------------------------------------------------------------------

test('dry run: reads Shopify, sends nothing, journals only run.start/run.end, previews remind1 / remind2', async () => {
  const ctx = await issuedCampaign({ count: 4, env: { DRY_RUN: 'true' } });
  ctx.fake.spendCard(cardGid(2002), 500);
  const before = journalOf(ctx).length;

  const res = await remind(ctx);
  assert.equal(res.exitCode, 0);
  assert.equal(ctx.fake.opsNamed('SendGiftCardNotification').length, 0);
  assert.equal(ctx.fake.opsNamed('RefreshCustomers').length, 1, 'a dry run re-reads customers like a live run');
  assert.equal(ctx.fake.opsNamed('TaggedCustomers').length, 1, '...and reads who carries the round tag like a live run');
  assert.equal(ctx.fake.opsNamed('TagsAdd').length, 0, 'a dry run adds no tag');
  const added = journalOf(ctx).slice(before);
  assert.deepEqual(added.map((e) => e.op), ['run.start', 'run.end']);
  assert.equal(added[0].command, 'remind');
  assert.equal(added[0].dryRun, true);
  assert.deepEqual(added[0].options, { round: 1, retryUnknown: false, retryFailed: false });
  assert.equal(added[1].exitCode, 0);
  assert.equal(added[1].summary.dryRun, true);
  assert.deepEqual(res.summary.skipped, { used: 1 });
  assert.equal(res.summary.eligible, 3);
  assert.equal(res.summary.planned, 3);
  assert.equal(res.summary.sent, 0);
  assert.equal(foldJournal(journalOf(ctx)).customers.get(customerGid(2)).reminders['1'], undefined, 'nobody\'s state changes');

  assert.equal(res.previews.length, 1);
  assert.equal(res.previews[0].variant, 'remind1');
  assert.equal(res.previews[0].recipient.customerId, customerGid(1), 'previewed with the first person of the batch');
  assert.equal(res.previews[0].config, ctx.config);
  assert.equal(res.previews[0].paths.previewDir, ctx.paths.previewDir);
  assert.equal(typeof res.previews[0].log.info, 'function');
  const text = logText(res);
  assert.match(text, /#1 First1 Last1 <c1@example.org>/);
  assert.doesNotMatch(text, /#2 First2/);
  assert.match(text, /邮件预览（按 2026-10-12 发送时的文案）：.*remind1\.html/);
  assert.match(text, /卡已用过 1/);

  const round2 = await remind(ctx, { round: 2, now: () => ROUND2_DAY });
  assert.equal(round2.exitCode, 0);
  assert.equal(round2.previews[0].variant, 'remind2');
  assert.equal(ctx.fake.opsNamed('SendGiftCardNotification').length, 0);
});

test('dry run: a preview failure is only a warning; with nobody eligible the preview uses sample data', async () => {
  const ctx = await issuedCampaign({ count: 1, env: { DRY_RUN: 'true' } });
  const res = await remind(ctx, {
    renderPreview: async () => {
      throw new Error('template missing');
    },
  });
  assert.equal(res.exitCode, 0);
  assert.match(logText(res), /WARN 邮件预览没有生成：template missing/);

  ctx.fake.spendCard(cardGid(2001), 1);
  const none = await remind(ctx);
  assert.equal(none.summary.eligible, 0);
  assert.equal(none.previews[0].recipient, null);
});

test('--limit counts send attempts; the next run continues with the following people', async () => {
  const ctx = await issuedCampaign({ count: 4 });
  const first = await remind(ctx, { limit: 2 });
  assert.equal(first.exitCode, 0);
  assert.deepEqual(notified(ctx), [cardGid(2001), cardGid(2002)]);
  assert.equal(first.summary.eligible, 4);
  assert.equal(first.summary.planned, 2);
  assert.equal(first.summary.attempted, 2);
  assert.equal(journalOf(ctx).filter((e) => e.op === 'run.start').at(-1).limit, 2);
  assert.equal(remindOps(ctx).filter((e) => e.cid === customerGid(3)).length, 0, 'people beyond the limit are left untouched');
  assert.match(logText(first), /还有 2 人符合条件但这次没有发出/);

  const rest = await remind(ctx);
  assert.deepEqual(notified(ctx), [2001, 2002, 2003, 2004].map(cardGid));
  assert.equal(rest.summary.alreadySent, 2);
});

test('invalid --round or --limit is a usage error (exit 2) before anything else', async () => {
  const ctx = await issuedCampaign({ count: 1 });
  for (const round of [0, 3, '3', 'x', '', undefined, null, 1.5]) {
    const res = await remind(ctx, { round });
    assert.equal(res.exitCode, 2, `round ${round}`);
    assert.match(logText(res), /--round 必须是 1 或 2/);
  }
  for (const limit of [-1, 1.5, 'abc', '', Number.NaN]) {
    const res = await remind(ctx, { limit });
    assert.equal(res.exitCode, 2, `limit ${limit}`);
    assert.match(logText(res), /--limit 必须是正整数/);
  }
  assert.equal(ctx.fake.state.calls.token.length, 0);

  const ok = await remind(ctx, { round: '1', limit: '1' });
  assert.equal(ok.exitCode, 0);
  assert.equal(notified(ctx).length, 1);
});

test('Ctrl+C: the current send finishes, the run stops with exit 130 and still refreshes the Excel', async () => {
  const ctx = await issuedCampaign({ count: 3 });
  const controller = new AbortController();
  const inner = globalThis.fetch;
  globalThis.fetch = async (url, init = {}) => {
    const response = await inner(url, init);
    if (/SendGiftCardNotification/.test(String(init.body ?? ''))) controller.abort(); // Ctrl+C during the first send
    return response;
  };
  const res = await remind(ctx, { signal: controller.signal });
  assert.equal(res.exitCode, 130);
  assert.equal(res.summary.stopped, 'aborted');
  assert.deepEqual(notified(ctx), [cardGid(2001)]);
  assert.deepEqual(remindOps(ctx).map((e) => e.op), ['remind.start', 'remind.ok'], 'the send in flight is recorded');
  assert.equal(journalOf(ctx).at(-1).exitCode, 130);
  assert.equal(res.reports.length, 1);
  acquireRunLock(ctx.paths, 'issue')(); // the lock was released

  const next = await remind(ctx);
  assert.equal(next.exitCode, 0);
  assert.deepEqual(notified(ctx), [2001, 2002, 2003].map(cardGid));
});

test('an already aborted signal stops before any card lookup or send', async () => {
  const ctx = await issuedCampaign({ count: 2 });
  const controller = new AbortController();
  controller.abort();
  const res = await remind(ctx, { signal: controller.signal });
  assert.equal(res.exitCode, 130);
  assert.equal(ctx.fake.opsNamed('FindCampaignCards').length, 0);
  assert.equal(notified(ctx).length, 0);
  assert.deepEqual(journalOf(ctx).slice(-2).map((e) => e.op), ['run.start', 'run.end']);
});

test('writeReport runs at the end of every started run; its failure never changes the exit code', async () => {
  const ctx = await issuedCampaign({ count: 2, failures: { SendGiftCardNotification: throttledOut() } });
  const failed = await remind(ctx, {
    writeReport: async () => {
      throw new Error('locked by Excel');
    },
  });
  assert.equal(failed.exitCode, 1, 'the command\'s own failure is kept');
  assert.match(logText(failed), /WARN Excel 没有更新：locked by Excel/);

  const ok = await remind(ctx);
  assert.equal(ok.exitCode, 0);
  assert.equal(ok.reports.length, 1);
  assert.equal(ok.reports[0].config, ctx.config);
  assert.equal(ok.reports[0].paths.journal, ctx.paths.journal);
  assert.equal(ok.reports[0].now().toISOString(), ROUND1_DAY.toISOString());
  assert.match(logText(ok), /Excel 已更新：/);

  const broken = await remind(ctx, {
    writeReport: async () => {
      throw new Error('disk full');
    },
  });
  assert.equal(broken.exitCode, 0);
  assert.match(logText(broken), /WARN Excel 没有更新：disk full/);
});

test('refuses to run while another write command holds the lock, and releases its own lock', async () => {
  const ctx = await issuedCampaign({ count: 1 });
  const release = acquireRunLock(ctx.paths, 'issue');
  try {
    const res = await remind(ctx);
    assert.equal(res.exitCode, 1);
    assert.match(logText(res), /issue 正在运行/);
    assert.equal(ctx.fake.state.calls.token.length, 0);
    assert.equal(res.reports.length, 0);
  } finally {
    release();
  }
  const res = await remind(ctx);
  assert.equal(res.exitCode, 0);
  assert.equal(notified(ctx).length, 1);
  acquireRunLock(ctx.paths, 'verify')();
});

test('without a matching selection.json the command stops with exit 1', async () => {
  const t = testConfig();
  teardown.push(t.cleanup);
  const fake = installFakeShopify({});
  teardown.push(() => fake.restore());
  const ctx = { config: t.config, paths: campaignPaths(t.config), fake };

  const none = await remind(ctx);
  assert.equal(none.exitCode, 1);
  assert.match(logText(none), /还没有名单，请先运行 select/);

  const selection = await selectionFixture(t.config, { customers: eligibleCustomers(1), write: false });
  writeJsonAtomic(ctx.paths.selection, { ...selection, campaignId: '2026-09' });
  const other = await remind(ctx);
  assert.equal(other.exitCode, 1);
  assert.match(logText(other), /名单属于活动 2026-09，和 \.env 的 CAMPAIGN_ID=2026-10 不一致/);

  writeJsonAtomic(ctx.paths.selection, { ...selection, params: { ...selection.params, sentTag: 'OTHER-TAG' } });
  const tag = await remind(ctx);
  assert.equal(tag.exitCode, 1);
  assert.match(logText(tag), /名单的 SENT_TAG 是 OTHER-TAG/);

  assert.equal(fake.state.calls.token.length, 0);
  assert.equal(none.reports.length + other.reports.length + tag.reports.length, 0);
});

// ---- final review round: tags seen by the by-id read; deleted customers ------------------------

/** Make Shopify's customer tag search lag: TaggedCustomers lists nobody (by-id reads still show tags). */
function hideTagSearch() {
  const inner = globalThis.fetch;
  globalThis.fetch = async (url, init = {}) => {
    if (String(url).endsWith('/graphql.json') && /query\s+TaggedCustomers/.test(JSON.parse(init.body).query)) {
      const body = { data: { customers: { edges: [], pageInfo: { hasNextPage: false, endCursor: null } } }, extensions: { cost: { requestedQueryCost: 1, actualQueryCost: 1, throttleStatus: { maximumAvailable: 20000, currentlyAvailable: 19999, restoreRate: 1000 } } } };
      return new Response(JSON.stringify(body), { status: 200, headers: { 'content-type': 'application/json' } });
    }
    return inner(url, init);
  };
}

test('a round tag the search does not show yet (e.g. added by hand minutes ago) still stops the reminder: the by-id re-read sees it', async () => {
  const ctx = await issuedCampaign({ count: 2, env: { SENT_TAG: LIVE_TAG } });
  ctx.customer(1).tags.push(R1); // the operator tagged customer 1 by hand after confirming the email arrived
  hideTagSearch();
  const res = await remind(ctx);
  assert.equal(res.exitCode, 0);
  assert.deepEqual(notified(ctx), [cardGid(2002)], 'customer 1 is never reminded again');
  assert.equal(res.summary.alreadySentByTag, 1);
  assert.deepEqual(roundFields(ctx, 1, ['status', 'source']), { status: 'sent', source: 'tag' });
  assert.match(logText(res), /按 tag 认定已发 1 人/);
  // A dry run tells the same, and writes nothing.
  const dry = await remind(ctx, { config: { ...ctx.config, dryRun: true } });
  assert.equal(dry.summary.alreadySentByTag, 0, 'customer 1 is already recorded as sent now');
});

test('a customer deleted after their reminder: the tag repair settles it once and never retries or promises a repair', async () => {
  const ctx = await issuedCampaign({ count: 2, env: { SENT_TAG: LIVE_TAG } });
  assert.equal((await remind(ctx)).exitCode, 0);
  assert.deepEqual(taggedWith(ctx, R1), [customerGid(1), customerGid(2)]);
  ctx.fake.state.customers = ctx.fake.state.customers.filter((c) => c.id !== customerGid(1)); // deleted in the admin
  const tagsBefore = tagCalls(ctx).length;

  const later = await remind(ctx, { now: () => minutesAfterRound1(20) });
  assert.equal(later.exitCode, 0);
  assert.equal(tagCalls(ctx).length, tagsBefore, 'no tagsAdd for a deleted customer');
  assert.equal(later.summary.roundTagFailed, 0);
  assert.deepEqual(remindOps(ctx).filter((e) => e.cid === customerGid(1) && e.op.startsWith('remind.tag')).map((e) => [e.op, e.note]), [['remind.tag.ok', 'customer deleted']]);
  assert.doesNotMatch(logText(later), /下次运行会补打/);

  const evenLater = await remind(ctx, { now: () => minutesAfterRound1(60) });
  assert.equal(evenLater.exitCode, 0);
  assert.equal(remindOps(ctx).filter((e) => e.cid === customerGid(1) && e.op.startsWith('remind.tag')).length, 1, 'settled for good');
  assert.equal(notified(ctx).length, 2, 'nobody reminded twice');
});
