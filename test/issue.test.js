import { describe, it, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

import { runIssue, SKIP_REASONS, SELECTION_REPLACED, newCardsBlockedOn } from '../src/issue.js';
import { STATUS, acquireRunLock, appendJournal, campaignPaths, foldJournal, readJournal, writeJsonAtomic } from '../src/campaign.js';
import { buildTestSelection } from '../src/select/selection.js';
import { parseCustomer } from '../src/select/rules.js';
import { resetClient } from '../src/shopify.js';
import { localDate } from '../src/time.js';
import { installFakeShopify } from './fake-shopify.js';
import { NOW_ISO, gid, makeActiveOrder, makeCustomer, makeOrder, memoryLog, selectionFixture, testConfig } from './helpers.js';

const MINUTE = 60_000;
const LAUNCH_DAY_MS = Date.parse('2026-10-05T17:00:00Z'); // 10:00 on LAUNCH_DATE (2026-10-05) in Los Angeles
// Last-order totals by customer number: 10% → $5.00 / $15.00 / $20.00 → tiers $10.77 / $15.33 / $19.77.
const TOTALS = ['50.00', '150.00', '200.00'];
const TIER_AMOUNTS = ['10.77', '15.33', '19.77'];
const TAG = 'gift-card-sent-2026-10'; // testConfig() default SENT_TAG
const NOTE = 'gift-card-promo [campaign:2026-10]';
const noop = async () => {};

/** Customers 1..n who all qualify. Same last-order date, so seq order = customer number. */
function qualifyingCustomers(n) {
  return Array.from({ length: n }, (_, i) => makeCustomer({
    n: i + 1,
    lastOrder: makeOrder({ n: 500 + i + 1, createdAt: '2026-03-01T12:00:00Z', total: TOTALS[i % 3] }),
  }));
}

function writeTestSelection(config, customers) {
  const selection = buildTestSelection({
    config,
    customers: customers.map(parseCustomer),
    timezone: 'America/Los_Angeles',
    createdAt: NOW_ISO,
    snapshot: { exportedAt: NOW_ISO, source: 'nodes', count: customers.length },
  });
  writeJsonAtomic(campaignPaths(config).selection, selection);
  return selection;
}

/** A gift card node as Shopify (the fake) stores it, carrying this campaign's marker. */
function campaignCard({ id, customer, amount = '10.77', createdAt }) {
  return {
    id,
    createdAt,
    note: NOTE,
    templateSuffix: 'gift-card-promo',
    enabled: true,
    lastCharacters: 'x777',
    initialValue: { amount, currencyCode: 'USD' },
    balance: { amount, currencyCode: 'USD' },
    customer: { id: gid('Customer', customer) },
    expiresOn: '2026-10-19',
  };
}

/** Journal lines left by an issue run that died right after writing create.start. */
function crashedCreateStart(k, atMs, amountCents = 1077) {
  const t = new Date(atMs).toISOString();
  appendJournal(h.paths.journal, { op: 'run.start', run: 'crashed', command: 'issue', dryRun: false, batch: 1, limit: 1, options: { retryFailed: false } }, { now: () => t });
  appendJournal(h.paths.journal, { op: 'create.start', cid: gid('Customer', k), amountCents, batch: 1, run: 'crashed' }, { now: () => t });
  return t;
}

function assertSummary(summary, expected) {
  const actual = Object.fromEntries(Object.keys(expected).map((k) => [k, summary[k]]));
  assert.deepEqual(actual, expected);
}

let h = null; // the current test's harness

afterEach(() => {
  h?.fake.restore();
  h?.cleanup();
  resetClient();
  h = null;
});

/**
 * Temp campaign dir + selection.json + fake Shopify sharing one controllable clock.
 * h.run(opts) calls runIssue with stubs (no-op sleep, recording writeReport/renderPreview)
 * and moves the clock one minute forward afterwards.
 */
async function setup({ n = 6, customers, env = {}, failures = {}, testCampaign = false } = {}) {
  const cfg = testConfig(env);
  const list = customers ?? qualifyingCustomers(n);
  const selection = testCampaign ? writeTestSelection(cfg.config, list) : await selectionFixture(cfg.config, { customers: list });
  const clock = { ms: LAUNCH_DAY_MS };
  const fake = installFakeShopify({ customers: list, failures, now: () => clock.ms });
  const paths = campaignPaths(cfg.config);
  const log = memoryLog();
  const reports = [];
  const previews = [];
  h = {
    config: cfg.config,
    cleanup: cfg.cleanup,
    paths,
    selection,
    list,
    clock,
    fake,
    log,
    reports,
    previews,
    async run(opts = {}) {
      const result = await runIssue({
        config: h.config,
        log,
        now: () => new Date(clock.ms),
        sleep: noop,
        writeReport: async (o) => {
          reports.push(o);
          return { file: paths.excel, fileEn: paths.excelEn, out: null, outEn: null, warnings: [] };
        },
        renderPreview: async (o) => {
          previews.push(o);
          return { file: path.join(paths.previewDir, `${o.variant}.html`), subject: 'Subject', stage: o.variant };
        },
        ...opts,
      });
      clock.ms += MINUTE;
      return result;
    },
    journal: () => readJournal(paths.journal),
    states: () => foldJournal(readJournal(paths.journal)).customers,
    statusOf: (k) => foldJournal(readJournal(paths.journal)).customers.get(gid('Customer', k))?.status ?? STATUS.PENDING,
    entriesFor: (k, op) => readJournal(paths.journal).filter((e) => e.cid === gid('Customer', k) && (!op || e.op === op)),
    cardsOf: (k) => fake.state.giftCards.filter((g) => g.customer?.id === gid('Customer', k)),
    createdFor: () => fake.state.calls.create.map((input) => Number(input.customerId.split('/').pop())),
    customer: (k) => fake.state.customers.find((c) => c.id === gid('Customer', k)),
    logged: (text) => log.lines.some((line) => line.includes(text)),
  };
  return h;
}

describe('issue: real runs', () => {
  it('--limit 3 creates exactly 3 cards in seq order with the frozen settings, tags them and refreshes the Excel', async () => {
    await setup({ n: 6 });
    const { exitCode, summary } = await h.run({ limit: 3 });
    assert.equal(exitCode, 0);

    assert.deepEqual(h.createdFor(), [1, 2, 3]);
    h.fake.state.calls.create.forEach((input, i) => {
      assert.deepEqual(input, {
        customerId: gid('Customer', i + 1),
        initialAmount: { amount: TIER_AMOUNTS[i], currencyCode: 'USD' },
        note: NOTE,
        expiresOn: '2026-10-19',
        templateSuffix: 'gift-card-promo',
      });
    });
    assert.deepEqual(h.fake.state.calls.tag, [1, 2, 3].map((k) => ({ id: gid('Customer', k), tags: [TAG] })));

    const states = h.states();
    for (const k of [1, 2, 3]) {
      const s = states.get(gid('Customer', k));
      assert.equal(s.status, STATUS.DONE);
      assert.equal(s.batch, 1);
      assert.equal(s.giftCardId, h.cardsOf(k)[0].id);
      assert.equal(s.last4, h.cardsOf(k)[0].lastCharacters);
      assert.equal(s.amountCents, h.selection.recipients[k - 1].amountCents);
      assert.ok(s.taggedAt);
    }
    for (const k of [4, 5, 6]) assert.equal(h.statusOf(k), STATUS.PENDING);

    const journal = h.journal();
    assert.deepEqual(journal.map((e) => e.op), [
      'run.start',
      'create.start', 'create.ok', 'tag.ok',
      'create.start', 'create.ok', 'tag.ok',
      'create.start', 'create.ok', 'tag.ok',
      'run.end',
    ]);
    const [start] = journal;
    assertSummary(start, { command: 'issue', dryRun: false, batch: 1, limit: 3, options: { retryFailed: false, repairOnly: false } });
    assert.ok(journal.every((e) => e.run === start.run), 'every entry carries the run id');
    assert.ok(journal.every((e) => e.t.startsWith('2026-10-05T17:00')), 'timestamps come from the injected clock');
    assert.deepEqual(journal.at(-1).summary, summary);
    assert.equal(journal.at(-1).exitCode, 0);

    assertSummary(summary, {
      batch: 1, dryRun: false, attempted: 3, created: 3, tagged: 3, tagFixed: 0, reconciled: 0,
      skipped: {}, failed: 0, rejected: 0, unknown: 0, amountCents: 1077 + 1533 + 1977, seqFrom: 1, seqTo: 3,
    });

    assert.equal(h.reports.length, 1);
    assert.equal(h.reports[0].config, h.config);
    assert.equal(h.reports[0].paths.journal, h.paths.journal);
    assert.equal(typeof h.reports[0].now, 'function');
    assert.equal(h.previews.length, 0, 'no preview in a real run');

    assert.ok(h.logged('本批序号：1–3'));
    assert.ok(h.logged('建卡：3 张，金额合计 $45.87'));
    assert.ok(h.logged('打 tag：3 人'));
    assert.ok(h.logged('跳过：0 人'));
    assert.ok(h.logged('结果不明：0 人'));
    assert.ok(h.logged('还有 3 人待发放'));
    assert.equal(h.log.lines.some((l) => /fake-token|secret/.test(l)), false, 'never logs credentials');
  });

  it('create.start is already on disk when Shopify receives each giftCardCreate', async () => {
    await setup({ n: 2 });
    const seen = [];
    const inner = globalThis.fetch;
    globalThis.fetch = async (url, init) => {
      const body = String(url).endsWith('/graphql.json') ? JSON.parse(init.body) : null;
      if (body && /mutation GiftCardCreate/.test(body.query)) {
        const last = readJournal(h.paths.journal).at(-1);
        seen.push({ op: last.op, cid: last.cid, input: body.variables.input.customerId });
      }
      return inner(url, init);
    };
    assert.equal((await h.run({ limit: 2 })).exitCode, 0);
    assert.deepEqual(seen, [1, 2].map((k) => ({ op: 'create.start', cid: gid('Customer', k), input: gid('Customer', k) })));
  });

  it('a second run continues at the next seq, with the next batch number', async () => {
    await setup({ n: 5 });
    assert.equal((await h.run({ limit: 2 })).exitCode, 0);
    const second = await h.run({ limit: 2 });
    assert.equal(second.exitCode, 0);
    assert.equal(second.summary.batch, 2);

    assert.deepEqual(h.createdFor(), [1, 2, 3, 4]);
    const states = h.states();
    assert.deepEqual([1, 2, 3, 4].map((k) => states.get(gid('Customer', k)).batch), [1, 1, 2, 2]);
    for (const k of [1, 2, 3, 4]) assert.equal(h.cardsOf(k).length, 1);
    assert.equal(h.statusOf(5), STATUS.PENDING);

    const starts = h.journal().filter((e) => e.op === 'run.start');
    assert.deepEqual(starts.map((e) => e.batch), [1, 2]);
    assert.notEqual(starts[0].run, starts[1].run);
  });

  it('people skipped by the pre-flight do not count toward --limit', async () => {
    await setup({ n: 5 });
    for (const k of [1, 2]) h.customer(k).defaultEmailAddress.marketingState = 'UNSUBSCRIBED';
    const { exitCode, summary } = await h.run({ limit: 2 });
    assert.equal(exitCode, 0);
    assert.deepEqual(h.createdFor(), [3, 4]);
    assert.equal(summary.attempted, 2);
    assert.deepEqual(summary.skipped, { 'not-subscribed': 2 });
    for (const k of [1, 2]) {
      const [skip] = h.entriesFor(k, 'skip');
      assertSummary(skip, { reason: 'not-subscribed', detail: 'UNSUBSCRIBED', batch: 1 });
      assert.equal(h.statusOf(k), STATUS.SKIPPED);
    }
    assert.ok(h.logged('#1 First1 Last1：跳过，未订阅邮件营销（UNSUBSCRIBED）'));
    assert.ok(h.logged('跳过：2 人（未订阅邮件营销 2）'));
  });

  it('pre-flight skips: deleted, no email, relay domain, unsubscribed, ordered since the snapshot, same-address order', async () => {
    const list = qualifyingCustomers(9);
    await setup({ customers: list });
    const st = h.fake.state;
    st.customers = st.customers.filter((c) => c.id !== gid('Customer', 1)); // 1: deleted
    h.customer(2).defaultEmailAddress = null; // 2: email removed
    h.customer(3).defaultEmailAddress.emailAddress = 'c3@relay.walmart.com'; // 3: marketplace relay address
    h.customer(4).defaultEmailAddress.marketingState = 'UNSUBSCRIBED'; // 4: unsubscribed
    st.orders.push(makeActiveOrder({ n: 901, customer: list[4], createdAt: '2026-10-02T12:00:00Z' })); // 5: ordered
    st.orders.push(makeActiveOrder({ // 6: another account at 6's address ordered
      n: 902,
      customer: { id: gid('Customer', 90), defaultAddress: list[5].defaultAddress },
      createdAt: '2026-10-02T13:00:00Z',
    }));
    h.customer(7).lastOrder = makeOrder({ n: 903, createdAt: '2026-10-03T15:00:00Z', total: '20.00' }); // 7: not in the orders search yet
    st.orders.push(makeActiveOrder({ n: 904, customer: list[7], createdAt: '2026-10-02T14:00:00Z', cancelled: true })); // 8: cancelled order only

    const { exitCode, summary } = await h.run({ limit: 2 });
    assert.equal(exitCode, 0);
    const reasons = Object.fromEntries([1, 2, 3, 4, 5, 6, 7].map((k) => [k, h.entriesFor(k, 'skip')[0]?.reason]));
    assert.deepEqual(reasons, {
      1: 'customer-deleted',
      2: 'no-email',
      3: 'relay-email',
      4: 'not-subscribed',
      5: 'ordered-since-snapshot',
      6: 'address-ordered-since-snapshot',
      7: 'ordered-since-snapshot',
    });
    assert.equal(h.entriesFor(3, 'skip')[0].detail, 'relay.walmart.com');
    assert.equal(h.entriesFor(5, 'skip')[0].detail, '2026-10-02T12:00:00Z');
    assert.equal(h.entriesFor(6, 'skip')[0].detail, gid('Customer', 90));
    assert.ok(h.journal().filter((e) => e.op === 'skip').every((e) => e.batch === 1));
    for (const k of [1, 2, 3, 4, 5, 6, 7]) assert.equal(h.statusOf(k), STATUS.SKIPPED);

    assert.deepEqual(h.createdFor(), [8, 9], 'a cancelled order does not count; skips do not use up the limit');
    assert.deepEqual(summary.skipped, {
      'customer-deleted': 1, 'no-email': 1, 'relay-email': 1, 'not-subscribed': 1,
      'ordered-since-snapshot': 2, 'address-ordered-since-snapshot': 1,
    });
    assert.ok(h.logged(`#6 First6 Last6：跳过，${SKIP_REASONS['address-ordered-since-snapshot']}（客户 90）`));
  });

  it('re-reads customers in smaller pieces when Shopify says the query is too expensive', async () => {
    await setup({
      n: 6,
      failures: { RefreshCustomers: [{ kind: 'graphqlError', code: 'MAX_COST_EXCEEDED', message: 'Query cost is 1750, which exceeds the single query max cost limit (1000).' }] },
    });
    const { exitCode } = await h.run({ limit: 6 });
    assert.equal(exitCode, 0);
    assert.deepEqual(h.fake.opsNamed('RefreshCustomers').map((o) => o.variables.ids.length), [6, 3, 3]);
    assert.deepEqual(h.createdFor(), [1, 2, 3, 4, 5, 6]);
  });

  it('logs progress every 50 people', async () => {
    await setup({ n: 51 });
    const { exitCode, summary } = await h.run({ limit: 51 });
    assert.equal(exitCode, 0);
    assert.equal(summary.created, 51);
    assert.ok(h.logged('进度：已处理 50/51 人，建卡 50 张'));
  });
});

describe('issue: dry run', () => {
  it('reads only: no card, no tag, only run.start/run.end in the journal, preview of the first email', async () => {
    await setup({ n: 6, env: { DRY_RUN: 'true' } });
    h.clock.ms = Date.parse('2026-10-03T18:00:00Z'); // before LAUNCH_DATE: dry runs may run any day
    h.customer(2).defaultEmailAddress.marketingState = 'UNSUBSCRIBED';
    const { exitCode, summary } = await h.run({ limit: 3 });
    assert.equal(exitCode, 0);

    assert.equal(h.fake.opsNamed('GiftCardCreate').length, 0);
    assert.equal(h.fake.opsNamed('TagsAdd').length, 0);
    assert.ok(h.fake.opsNamed('RefreshCustomers').length >= 1, 'the read-only pre-flight still runs');
    const journal = h.journal();
    assert.deepEqual(journal.map((e) => e.op), ['run.start', 'run.end']);
    assertSummary(journal[0], { dryRun: true, batch: 1, limit: 3 });
    assert.equal(journal[1].exitCode, 0);

    assertSummary(summary, {
      batch: 1, dryRun: true, attempted: 3, created: 0, tagged: 0, failed: 0, unknown: 0,
      skipped: { 'not-subscribed': 1 }, amountCents: 1077 + 1977 + 1077, seqFrom: 1, seqTo: 4,
    });

    assert.equal(h.previews.length, 1);
    assert.equal(h.previews[0].variant, 'first');
    assert.equal(h.previews[0].recipient.seq, 1);
    assert.equal(h.previews[0].recipient.customerId, gid('Customer', 1));
    assert.equal(h.previews[0].config, h.config);
    assert.equal(h.previews[0].paths.previewDir, h.paths.previewDir);
    assert.equal(h.previews[0].log, h.log);

    assert.ok(h.logged('预演：这一批会给 3 人建卡，序号 1–4'));
    assert.ok(h.logged('档位 $10.77：2 人，合计 $21.54'));
    assert.ok(h.logged('档位 $19.77：1 人，合计 $19.77'));
    assert.ok(h.logged('金额合计：$41.31'));
    assert.ok(h.logged('会被跳过 1 人（未订阅邮件营销 1）'));
    assert.ok(h.logged('#2 First2 Last2：未订阅邮件营销（UNSUBSCRIBED）'));
    assert.ok(h.logged('#1 First1 Last1 <c1@example.org> $10.77'));
    assert.ok(h.logged(path.join(h.paths.previewDir, 'first.html')));
    assert.ok(h.logged('正式活动要到 2026-10-05（店铺时间）才能真实建卡发信'));
    assert.ok(h.logged('预演结果（批次 1，Shopify 上什么都没改）'));
    assert.equal(h.reports.length, 1, 'the Excel gets the dry run in its run log');

    // Nothing changed: a real run afterwards still starts at seq 1 as batch 1.
    h.config = { ...h.config, dryRun: false };
    h.clock.ms = LAUNCH_DAY_MS;
    const live = await h.run({ limit: 1 });
    assert.equal(live.exitCode, 0);
    assert.equal(live.summary.batch, 1);
    assert.deepEqual(h.createdFor(), [1]);
  });

  it('without --limit previews everyone remaining, re-reading 250 customers at a time', async () => {
    await setup({ n: 260, env: { DRY_RUN: 'true' } });
    const { exitCode, summary } = await h.run({});
    assert.equal(exitCode, 0);
    assert.equal(summary.attempted, 260);
    assert.equal(h.journal()[0].limit, null);
    assert.deepEqual(h.fake.opsNamed('RefreshCustomers').map((o) => o.variables.ids.length), [250, 10]);
    assert.ok(h.logged('序号 1–260'));
  });

  it('a failing preview is only a warning', async () => {
    await setup({ n: 2, env: { DRY_RUN: 'true' } });
    const { exitCode } = await h.run({ renderPreview: async () => { throw new Error('template missing'); } });
    assert.equal(exitCode, 0);
    assert.ok(h.logged('WARN 邮件预览没有生成：template missing'));
  });

  it('the preview shows the expiry date the cards carry (the one frozen in the list), on any day a card may still be created', async () => {
    await setup({ n: 2, env: { DRY_RUN: 'true' } });
    h.config = { ...h.config, giftCardExpiresOn: '2026-10-20' }; // .env changed after select: cards keep the selection's date
    const previewAt = async (iso) => {
      h.clock.ms = Date.parse(iso);
      h.previews.length = 0;
      assert.equal((await h.run({ limit: 1 })).exitCode, 0);
      assert.equal(h.previews.length, 1);
      return h.previews[0];
    };
    const early = await previewAt('2026-10-03T18:00:00Z');
    assert.equal(early.expiresOn, '2026-10-19');
    assert.equal(h.selection.params.giftCardExpiresOn, '2026-10-19');
    assert.equal(early.variant, 'first');
    assert.equal('sendDate' in early, false, 'the copy no longer depends on the send date');
    assert.equal((await previewAt('2026-10-05T07:00:00Z')).expiresOn, '2026-10-19');
    assert.equal((await previewAt('2026-10-13T17:00:00Z')).expiresOn, '2026-10-19', 'the old reminder dates mean nothing any more');
    assert.ok(h.logged('首封邮件预览（用这一批第 1 个人 #1 的名字和金额）：'));
    assert.equal((await previewAt('2026-10-19T06:59:00Z')).expiresOn, '2026-10-19', '23:59 on 10/18: the last day for new cards');

    // From the expiry date on a real run creates no card, so there is no first email to preview.
    h.clock.ms = Date.parse('2026-10-19T07:00:00Z');
    h.previews.length = 0;
    assert.equal((await h.run({ limit: 1 })).exitCode, 0);
    assert.equal(h.previews.length, 0);
  });

  it('a test campaign\'s preview is rendered on any day before the expiry date', async () => {
    await setup({
      customers: [makeCustomer({ n: 1 }), makeCustomer({ n: 2 })],
      testCampaign: true,
      env: { CAMPAIGN_ID: '2026-10-test', SENT_TAG: 'OCT26RTPROMO-TEST', TEST_CUSTOMER_IDS: '1,2', DRY_RUN: 'true' },
    });
    h.clock.ms = Date.parse('2026-10-03T18:00:00Z');
    assert.equal((await h.run({ limit: 1 })).exitCode, 0);
    h.clock.ms = Date.parse('2026-10-18T18:00:00Z');
    assert.equal((await h.run({ limit: 1 })).exitCode, 0);
    assert.deepEqual(h.previews.map((p) => p.variant), ['first', 'first']);
    assert.deepEqual(h.previews.map((p) => p.expiresOn), [h.selection.params.giftCardExpiresOn, h.selection.params.giftCardExpiresOn]);
    assert.equal(h.logged('正式活动要'), false, 'test campaigns get no launch-date warnings');
  });

  it('reports leftovers without touching them, and that a real run would stop', async () => {
    await setup({ n: 2, env: { DRY_RUN: 'true' } });
    crashedCreateStart(1, h.clock.ms);
    h.clock.ms += 2 * MINUTE;
    const { exitCode, summary } = await h.run({ limit: 1 });
    assert.equal(exitCode, 1);
    assert.equal(summary.stillUnknown, 1);
    assert.ok(h.logged('预演：真实运行会在这里停止'));
    assert.deepEqual(h.journal().slice(2).map((e) => e.op), ['run.start', 'run.end']);
    assert.equal(h.statusOf(1), STATUS.IN_PROGRESS, 'a dry run never changes anyone');
    assert.equal(h.fake.opsNamed('GiftCardCreate').length, 0);
  });
});

describe('issue: argument and date guards', () => {
  it('a real run without --limit is a usage error and touches nothing', async () => {
    await setup({ n: 2 });
    const { exitCode } = await h.run({});
    assert.equal(exitCode, 2);
    assert.ok(h.logged('实时运行必须给 --limit N'));
    assert.equal(h.fake.state.calls.token.length, 0);
    assert.equal(h.fake.state.calls.ops.length, 0);
    assert.equal(fs.existsSync(h.paths.journal), false);
    assert.equal(h.reports.length, 0);
  });

  it('--limit above ISSUE_MAX_PER_RUN, or not a positive integer, is a usage error', async () => {
    await setup({ n: 2, env: { ISSUE_MAX_PER_RUN: '5' } });
    assert.equal((await h.run({ limit: 6 })).exitCode, 2);
    assert.ok(h.logged('--limit 6 超过每次运行的上限 ISSUE_MAX_PER_RUN=5'));
    assert.equal((await h.run({ limit: 1.5 })).exitCode, 2);
    assert.equal((await h.run({ limit: -3 })).exitCode, 2);
    assert.equal((await h.run({ limit: Number.NaN })).exitCode, 2);
    h.config = { ...h.config, dryRun: true };
    assert.equal((await h.run({ limit: 6 })).exitCode, 2, 'the cap applies to dry runs too');
    assert.equal(h.fake.state.calls.ops.length, 0);
    assert.equal((await h.run({ limit: 5 })).exitCode, 0);
  });

  it('a real run of the live campaign before LAUNCH_DATE (store time) is refused', async () => {
    await setup({ n: 2 });
    h.clock.ms = Date.parse('2026-10-05T06:59:00Z'); // 2026-10-04 23:59 in Los Angeles
    const early = await h.run({ limit: 1 });
    assert.equal(early.exitCode, 2);
    assert.ok(h.logged('正式活动要到 2026-10-05（店铺时间）才能建卡发信'));
    assert.equal(h.fake.state.calls.token.length, 0);
    assert.equal(fs.existsSync(h.paths.journal), false);

    h.clock.ms = Date.parse('2026-10-05T07:00:00Z'); // midnight in Los Angeles
    const onTime = await h.run({ limit: 1 });
    assert.equal(onTime.exitCode, 0);
    assert.deepEqual(h.createdFor(), [1]);
  });

  it('a real run on or after the cards\' expiry date (store time) creates no new card', async () => {
    await setup({ n: 3 });
    h.clock.ms = Date.parse('2026-10-19T07:00:00Z'); // 00:00 on GIFT_CARD_EXPIRES_ON (2026-10-19) in Los Angeles
    const late = await h.run({ limit: 1 });
    assert.equal(late.exitCode, 2, 'people are still waiting for a card');
    assert.ok(h.logged(`ERROR ${D2_TEXT}本次只做了补记和补打 tag，没有建新卡；还有 3 人待建卡。`));
    assert.ok(h.logged('INFO 现在是店铺时间 2026-10-19 00:00'));
    assert.equal(h.fake.opsNamed('GiftCardCreate').length, 0);
    assert.equal(h.fake.opsNamed('TagsAdd').length, 0);
    assert.equal(h.reports.length, 1, 'the Excel is still refreshed');
    assert.equal(late.summary.newCardsRefused, 3, 'the run history says why nothing was created');
    assert.equal(h.journal().filter((e) => e.op === 'run.end').at(-1).summary.newCardsRefused, 3);
    assert.equal(late.summary.attempted, 0);
    assert.equal(h.logged('REMIND'), false);

    h.clock.ms = Date.parse('2026-10-25T17:00:00Z'); // long after
    assert.equal((await h.run({ limit: 1 })).exitCode, 2);
    assert.equal(h.fake.opsNamed('GiftCardCreate').length, 0);

    h.clock.ms = Date.parse('2026-10-19T06:59:00Z'); // 23:59 on 10/18 in Los Angeles: the last day for new cards
    const lastDay = await h.run({ limit: 1 });
    assert.equal(lastDay.exitCode, 0);
    assert.deepEqual(h.createdFor(), [1]);
    assert.ok(h.logged('INFO 真实运行：会建礼品卡'));
  });

  it('the old reminder dates (10/12, 10/16) no longer block anything: cards are created until the day before expiry', async () => {
    await setup({ n: 4 });
    for (const [k, iso] of [[1, '2026-10-12T07:00:00Z'], [2, '2026-10-12T00:30:00Z'], [3, '2026-10-16T17:00:00Z'], [4, '2026-10-18T23:00:00Z']]) {
      h.clock.ms = Date.parse(iso);
      assert.equal((await h.run({ limit: 1 })).exitCode, 0, iso);
      assert.equal(h.createdFor().at(-1), k);
    }
    assert.deepEqual(h.createdFor(), [1, 2, 3, 4]);
    assert.equal(h.log.lines.some((l) => /REMIND|UTC 已是|提醒的文案/.test(l)), false, h.log.lines.join('\n'));
  });

  it('the rule is the expiry frozen in the list, not today\'s .env; a list without an expiry date is never blocked', async () => {
    await setup({ n: 3 });
    h.config = { ...h.config, giftCardExpiresOn: '2026-10-30' }; // .env moved after select: cards are still made with 10/19
    h.clock.ms = Date.parse('2026-10-19T17:00:00Z');
    assert.equal((await h.run({ limit: 1 })).exitCode, 2);
    assert.ok(h.logged('礼品卡到期日（2026-10-19）已到'));
    assert.ok(h.logged('注意：.env 的 GIFT_CARD_EXPIRES_ON（2026-10-30）和生成名单时（2026-10-19）不一样'));
    assert.deepEqual(h.createdFor(), []);
    assert.equal(newCardsBlockedOn(h.selection, '2026-10-18'), false);
    assert.equal(newCardsBlockedOn(h.selection, '2026-10-19'), true);
    assert.equal(newCardsBlockedOn(h.selection, '2027-01-01'), true);

    // The list says the cards never expire: no date ever blocks a card.
    writeJsonAtomic(h.paths.selection, { ...h.selection, params: { ...h.selection.params, giftCardExpiresOn: '' } });
    h.config = { ...h.config, giftCardExpiresOn: '' };
    h.clock.ms = Date.parse('2026-12-01T17:00:00Z');
    assert.equal((await h.run({ limit: 1 })).exitCode, 0);
    assert.deepEqual(h.createdFor(), [1]);
    assert.equal(h.fake.state.calls.create[0].expiresOn, undefined);
    assert.equal(newCardsBlockedOn({ params: { giftCardExpiresOn: '' } }, '2027-01-01'), false);
    assert.equal(newCardsBlockedOn({ params: {} }, '2027-01-01'), false);
  });

  it('on or after the expiry date a real run still adds a missing tag and records cards Shopify already has, without creating cards', async () => {
    await setup({ n: 3, failures: { TagsAdd: [{ kind: 'userError', message: 'temporarily unavailable' }] } });
    h.clock.ms = Date.parse('2026-10-18T17:00:00Z');
    assert.equal((await h.run({ limit: 1 })).exitCode, 0);
    assert.equal(h.statusOf(1), STATUS.CREATED, 'card created, tag failed');

    h.clock.ms = Date.parse('2026-10-20T17:00:00Z');
    const repair = await h.run({ limit: 5 });
    assert.equal(repair.exitCode, 2, 'customers 2 and 3 still have no card');
    assert.equal(h.statusOf(1), STATUS.DONE, 'the missing tag was added');
    assert.deepEqual(h.createdFor(), [1], 'no new card on or after the expiry date');
    assert.ok(h.logged('还有 2 人待建卡'));
  });

  it('a dry run on or after the expiry date follows the same rule: no simulated card, no preview, a warning, exit 0', async () => {
    await setup({ n: 2, env: { DRY_RUN: 'true' } });
    h.clock.ms = Date.parse('2026-10-20T17:00:00Z');
    const { exitCode, summary } = await h.run({ limit: 1 });
    assert.equal(exitCode, 0);
    assertSummary(summary, { attempted: 0, created: 0, amountCents: 0, seqFrom: null, newCardsRefused: 2 });
    assert.ok(h.logged(`WARN 注意：${D2_TEXT}`));
    assert.ok(h.logged('WARN 预演：真实运行只会补记和补打 tag，不会建新卡（礼品卡到期日 2026-10-19 已到）；还有 2 人待建卡。'));
    assert.equal(h.logged('这一批会给'), false, 'no simulated creates');
    assert.equal(h.previews.length, 0, 'no first-email preview');
    assert.ok(h.logged('将建卡：0 张'));
    assert.equal(h.logged('才能真实建卡发信'), false, 'LAUNCH_DATE has passed');
    assert.equal(h.fake.opsNamed('GiftCardCreate').length, 0);
    assert.deepEqual(h.journal().map((e) => e.op), ['run.start', 'run.end']);
    assert.equal(h.journal()[1].summary.newCardsRefused, 2);
  });

  it('a test campaign is bound to the expiry date too (an expired card is of no use), but not to LAUNCH_DATE', async () => {
    await setup({
      customers: [makeCustomer({ n: 1 }), makeCustomer({ n: 2 })],
      testCampaign: true,
      env: { CAMPAIGN_ID: '2026-10-test', SENT_TAG: 'OCT26RTPROMO-TEST', TEST_CUSTOMER_IDS: '1,2' },
    });
    h.clock.ms = Date.parse('2026-10-16T17:00:00Z');
    assert.equal((await h.run({ limit: 1 })).exitCode, 0);
    assert.deepEqual(h.createdFor(), [1]);
    assert.equal(h.fake.state.calls.create[0].expiresOn, '2026-10-19');

    h.clock.ms = Date.parse('2026-10-19T07:00:00Z'); // 00:00 on the expiry date
    const late = await h.run({ limit: 1 });
    assert.equal(late.exitCode, 2);
    assert.deepEqual(h.createdFor(), [1]);
    assert.ok(h.logged(`ERROR ${D2_TEXT}本次只做了补记和补打 tag，没有建新卡；还有 1 人待建卡。`));
    assert.ok(h.logged('已到礼品卡到期日（2026-10-19）：只补记和补打 tag，不建新卡'));
    assert.equal(late.summary.newCardsRefused, 1);
  });

  it('a test campaign may run before LAUNCH_DATE and is not filtered by audience rules', async () => {
    const customers = [makeCustomer({ n: 1, marketingState: 'UNSUBSCRIBED' }), makeCustomer({ n: 2 })];
    await setup({
      customers,
      testCampaign: true,
      env: { CAMPAIGN_ID: '2026-10-test', SENT_TAG: 'OCT26RTPROMO-TEST', TEST_CUSTOMER_IDS: '1,2' },
    });
    h.clock.ms = Date.parse('2026-10-03T18:00:00Z');
    const { exitCode, summary } = await h.run({ limit: 5 });
    assert.equal(exitCode, 0);
    assert.deepEqual(h.createdFor(), [1, 2]);
    assert.equal(h.fake.state.calls.create[0].initialAmount.amount, '0.10');
    assert.equal(h.fake.state.calls.create[0].note, 'gift-card-promo [campaign:2026-10-test]');
    assert.deepEqual(h.fake.state.calls.tag.map((t) => t.tags), [['OCT26RTPROMO-TEST'], ['OCT26RTPROMO-TEST']]);
    assert.deepEqual(summary.skipped, {});
  });

  it('refuses to run without a selection, or with a config that does not match it', async () => {
    await setup({ n: 1 });
    h.config = { ...h.config, sentTag: 'OCT26RTPROMO' };
    assert.equal((await h.run({ limit: 1 })).exitCode, 1);
    assert.ok(h.logged('配置和名单不一致'));

    h.config = { ...h.config, sentTag: TAG };
    fs.rmSync(h.paths.selection);
    assert.equal((await h.run({ limit: 1 })).exitCode, 1);
    assert.ok(h.logged('还没有名单，请先运行 select'));
    assert.equal(h.fake.state.calls.token.length, 0);
  });

  it('refuses to run while another write command holds the lock', async () => {
    await setup({ n: 2 });
    const release = acquireRunLock(h.paths, 'verify');
    try {
      const { exitCode } = await h.run({ limit: 1 });
      assert.equal(exitCode, 1);
      assert.ok(h.logged('verify 正在运行'));
      assert.equal(h.fake.state.calls.token.length, 0);
      assert.equal(fs.existsSync(h.paths.journal), false);
    } finally {
      release();
    }
    assert.equal((await h.run({ limit: 1 })).exitCode, 0, 'runs once the lock is free');
  });

  it('the list is read again under the run lock: if select rewrote it in between, no card is created (exit 1)', async () => {
    await setup({ n: 3 });
    const v2Customers = [4, 5, 6].map((n) => makeCustomer({ n, lastOrder: makeOrder({ n: 600 + n, createdAt: '2026-03-01T12:00:00Z', total: '50.00' }) }));
    h.fake.state.customers.push(...v2Customers.map((c) => structuredClone(c)));
    const v2 = await selectionFixture(h.config, { customers: v2Customers, createdAt: '2026-10-02T20:00:00.000Z', write: false });
    let clockReads = 0;
    const res = await h.run({
      limit: 1,
      now: () => {
        clockReads += 1;
        // The first clock read is the date guard: after the list was read, before the lock is taken.
        if (clockReads === 1) writeJsonAtomic(h.paths.selection, v2);
        return new Date(h.clock.ms);
      },
    });
    assert.equal(res.exitCode, 1);
    assert.ok(h.logged(`ERROR ${SELECTION_REPLACED}`));
    assert.equal(SELECTION_REPLACED, '名单刚被 select 改写，请重新运行');
    assert.equal(h.fake.state.calls.token.length, 0);
    assert.equal(fs.existsSync(h.paths.journal), false);
    assert.deepEqual(h.createdFor(), []);
    assert.equal(h.reports.length, 0);
    acquireRunLock(h.paths, 'verify')(); // the lock was released

    const next = await h.run({ limit: 1 }); // the next run issues from the list on disk
    assert.equal(next.exitCode, 0);
    assert.deepEqual(h.createdFor(), [4]);
  });
});

/** From now on every giftCardCreate takes `ms` of the shared clock (a stand-in for a long batch). */
function slowCreates(ms) {
  const inner = globalThis.fetch; // the fake's; the fake's restore() puts the real one back
  globalThis.fetch = async (url, init) => {
    const res = await inner(url, init);
    if (String(init?.body ?? '').includes('mutation GiftCardCreate')) h.clock.ms += ms;
    return res;
  };
}

// The refusal text on or after the cards' expiry date (testConfig: GIFT_CARD_EXPIRES_ON=2026-10-19).
const D2_TEXT = '礼品卡到期日（2026-10-19）已到，不再建新卡：现在建的卡客户已经用不了。'
  + '确需补发，请改 .env 的 GIFT_CARD_EXPIRES_ON，换一个新的 CAMPAIGN_ID 重新运行 select。';
const NO_NEW_CARDS = '礼品卡到期日（2026-10-19）已到，不再建新卡';
const TZ = 'America/Los_Angeles';

describe('issue: the date is checked again before every card', () => {
  it('a real run that crosses midnight into the expiry date stops before the next card (exit 2); what was done stays', async () => {
    await setup({ n: 6 });
    h.clock.ms = Date.parse('2026-10-19T06:57:30Z'); // 23:57:30 on 10/18 in Los Angeles
    slowCreates(MINUTE); // cards at 23:57:30, 23:58:30, 23:59:30; the 4th would be at 00:00:30 on 10/19
    const { exitCode, summary } = await h.run({ limit: 6 });
    assert.equal(exitCode, 2);
    assert.deepEqual(h.createdFor(), [1, 2, 3]);
    assert.deepEqual(h.fake.state.giftCards.map((g) => localDate(Date.parse(g.createdAt), TZ)), ['2026-10-18', '2026-10-18', '2026-10-18']);
    assert.deepEqual(h.journal().filter((e) => e.op === 'create.start').map((e) => localDate(Date.parse(e.t), TZ)), ['2026-10-18', '2026-10-18', '2026-10-18']);
    for (const k of [1, 2, 3]) assert.equal(h.statusOf(k), STATUS.DONE, 'cards made before midnight are kept and tagged');
    for (const k of [4, 5, 6]) assert.equal(h.statusOf(k), STATUS.PENDING);
    assert.ok(h.logged(`ERROR ${D2_TEXT}已处理的人都记在本地日志里；剩下的 3 人不会再建卡。`));
    assert.ok(h.logged('现在是店铺时间 2026-10-19 00:00'));
    assertSummary(summary, { attempted: 3, created: 3, tagged: 3, stoppedByDate: '2026-10-19' });
    const end = h.journal().at(-1);
    assertSummary(end, { op: 'run.end', exitCode: 2 });
    assert.equal(end.summary.stoppedByDate, '2026-10-19');
    assert.equal(h.reports.length, 1, 'the Excel is still refreshed');

    // The next run that day only repairs (D2), and says how many are still waiting.
    const next = await h.run({ limit: 6 });
    assert.equal(next.exitCode, 2);
    assert.deepEqual(h.createdFor(), [1, 2, 3]);
    assert.equal(next.summary.newCardsRefused, 3);
  });

  it('the re-check counts only the people this run had left; skipped people before the stop stay skipped', async () => {
    await setup({ n: 5 });
    h.customer(2).defaultEmailAddress.marketingState = 'UNSUBSCRIBED';
    h.clock.ms = Date.parse('2026-10-19T06:59:00Z'); // 23:59 on 10/18
    slowCreates(MINUTE); // #1 at 23:59; #2 is skipped (no card); #3 would be created at 00:00 on 10/19
    const { exitCode, summary } = await h.run({ limit: 5 });
    assert.equal(exitCode, 2);
    assert.deepEqual(h.createdFor(), [1]);
    assert.equal(h.statusOf(2), STATUS.SKIPPED);
    assert.deepEqual(summary.skipped, { 'not-subscribed': 1 });
    assert.ok(h.logged('剩下的 3 人不会再建卡。'));
    assert.equal(summary.stoppedByDate, '2026-10-19');
  });

  it('a test campaign is stopped too when its run crosses midnight into the expiry date', async () => {
    await setup({
      customers: [1, 2, 3].map((n) => makeCustomer({ n })),
      testCampaign: true,
      env: { CAMPAIGN_ID: '2026-10-test', SENT_TAG: 'OCT26RTPROMO-TEST', TEST_CUSTOMER_IDS: '1,2,3' },
    });
    h.clock.ms = Date.parse('2026-10-19T06:58:30Z');
    slowCreates(MINUTE); // #1 at 23:58:30, #2 at 23:59:30; #3 would be at 00:00:30 on 10/19
    const { exitCode, summary } = await h.run({ limit: 3 });
    assert.equal(exitCode, 2);
    assert.deepEqual(h.createdFor(), [1, 2]);
    assert.equal(summary.stoppedByDate, '2026-10-19');
  });

  it('crossing midnight into any other day (the old reminder dates included) never stops a run', async () => {
    await setup({ n: 3 });
    h.clock.ms = Date.parse('2026-10-12T06:58:30Z'); // 23:58:30 on 10/11: the 3rd card is made at 00:00:30 on 10/12
    slowCreates(MINUTE);
    const { exitCode, summary } = await h.run({ limit: 3 });
    assert.equal(exitCode, 0);
    assert.deepEqual(h.createdFor(), [1, 2, 3]);
    assert.deepEqual(h.fake.state.giftCards.map((g) => localDate(Date.parse(g.createdAt), TZ)), ['2026-10-11', '2026-10-11', '2026-10-12']);
    assert.equal(summary.stoppedByDate, undefined);
  });

  it('a run that stays on one day is never stopped by the re-check', async () => {
    await setup({ n: 3 });
    h.clock.ms = Date.parse('2026-10-18T17:00:00Z');
    slowCreates(MINUTE);
    const { exitCode, summary } = await h.run({ limit: 3 });
    assert.equal(exitCode, 0);
    assert.deepEqual(h.createdFor(), [1, 2, 3]);
    assert.equal(summary.stoppedByDate, undefined);
  });
});

describe('issue: on or after the expiry date (no new cards)', () => {
  it('the dry run shows what the real run will do: both create nothing and report the same people waiting', async () => {
    await setup({ n: 5 });
    h.clock.ms = Date.parse('2026-10-05T17:00:00Z');
    assert.equal((await h.run({ limit: 3 })).exitCode, 0);
    h.clock.ms = Date.parse('2026-10-19T17:00:00Z');
    const mark = h.log.lines.length;
    const dry = await h.run({ limit: 5, config: { ...h.config, dryRun: true } });
    const live = await h.run({ limit: 5 });
    const lines = h.log.lines.slice(mark);
    assert.equal(dry.exitCode, 0);
    assert.equal(live.exitCode, 2);
    assert.equal(dry.summary.attempted, live.summary.attempted);
    assert.equal(live.summary.created, 0);
    assert.equal(dry.summary.newCardsRefused, 2);
    assert.equal(live.summary.newCardsRefused, 2);
    assert.ok(h.logged('WARN 预演：真实运行只会补记和补打 tag，不会建新卡（礼品卡到期日 2026-10-19 已到）；还有 2 人待建卡。'));
    assert.ok(h.logged(`ERROR ${D2_TEXT}本次只做了补记和补打 tag，没有建新卡；还有 2 人待建卡。`));
    assert.equal(h.logged('这一批会给'), false);
    assert.equal(h.previews.length, 0);
    assert.deepEqual(h.createdFor(), [1, 2, 3]);
    assert.equal(lines.filter((l) => l.includes('已到礼品卡到期日（2026-10-19）：只补记和补打 tag，不建新卡')).length, 2, 'both banners say it from the start');
    assert.ok(lines.includes('INFO 真实运行：只补记 Shopify 上已有的卡、补打 tag gift-card-sent-2026-10；不建新卡，不发邮件'));
    assert.equal(lines.some((l) => l.includes('真实运行：会建礼品卡')), false, 'the live banner does not promise cards');
  });

  it('a dry run still simulates the tag repair and settles nothing for real', async () => {
    await setup({ n: 2, failures: { TagsAdd: [{ kind: 'userError', message: 'locked' }] } });
    h.clock.ms = Date.parse('2026-10-18T17:00:00Z');
    assert.equal((await h.run({ limit: 1 })).exitCode, 0);
    assert.equal(h.statusOf(1), STATUS.CREATED);
    h.clock.ms = Date.parse('2026-10-20T17:00:00Z');
    const dry = await h.run({ config: { ...h.config, dryRun: true } });
    assert.equal(dry.exitCode, 0);
    assertSummary(dry.summary, { tagFixed: 1, attempted: 0, newCardsRefused: 1 });
    assert.ok(h.logged('预演：1 人已建卡但还没打 tag，真实运行会先给他们补打 tag'));
    assert.equal(h.statusOf(1), STATUS.CREATED, 'a dry run changes nobody');
    assert.equal(h.fake.opsNamed('TagsAdd').length, 1);
  });

  it('--retry-failed with nobody failed: says so, and still exits 2 while people are pending', async () => {
    await setup({ n: 3 });
    h.clock.ms = Date.parse('2026-10-05T17:00:00Z');
    assert.equal((await h.run({ limit: 1 })).exitCode, 0);
    h.clock.ms = Date.parse('2026-10-20T17:00:00Z');
    const r = await h.run({ limit: 5, retryFailed: true });
    assert.equal(r.exitCode, 2);
    assert.ok(h.logged('INFO 没有建卡失败、待重试的人'));
    assert.ok(h.logged(`ERROR ${D2_TEXT}本次只做了补记和补打 tag，没有建新卡；还有 2 人待建卡。`));
    assert.equal(h.logged('没有待建卡的人'), false);
    assert.equal(r.summary.newCardsRefused, 2);
    assert.deepEqual(h.createdFor(), [1]);
  });

  it('--retry-failed with failed rows: they get no card either (exit 2)', async () => {
    await setup({ n: 3, failures: { GiftCardCreate: [{ kind: 'userError', message: 'Customer is invalid' }] } });
    h.clock.ms = Date.parse('2026-10-05T17:00:00Z');
    assert.equal((await h.run({ limit: 1 })).exitCode, 0);
    assert.equal(h.statusOf(1), STATUS.FAILED);
    h.clock.ms = Date.parse('2026-10-20T17:00:00Z');
    const r = await h.run({ limit: 5, retryFailed: true });
    assert.equal(r.exitCode, 2);
    assert.ok(h.logged('还有 1 人待建卡。'));
    assert.equal(r.summary.newCardsRefused, 1);
    assert.equal(h.fake.opsNamed('GiftCardCreate').length, 1, 'no new attempt');
    assert.equal(h.statusOf(1), STATUS.FAILED);
  });

  it('nobody left to create: exit 0 with the repairs done; --retry-failed says there is nobody to retry', async () => {
    await setup({ n: 2 });
    h.clock.ms = Date.parse('2026-10-05T17:00:00Z');
    assert.equal((await h.run({ limit: 2 })).exitCode, 0);
    h.clock.ms = Date.parse('2026-10-20T17:00:00Z');
    const normal = await h.run({ limit: 1 });
    assert.equal(normal.exitCode, 0);
    assert.ok(h.logged('INFO 没有待建卡的人；已完成补记和补打 tag（如果有）'));
    assert.equal(normal.summary.newCardsRefused, undefined);
    const retry = await h.run({ limit: 1, retryFailed: true });
    assert.equal(retry.exitCode, 0);
    assert.ok(h.logged('INFO 没有建卡失败、待重试的人；已完成补记和补打 tag（如果有）'));
  });

  it('a tag repair that fails again is reported, never as done: exit 1 when nobody else waits, 2 when people do', async () => {
    const twice = () => [{ kind: 'userError', message: 'locked' }, { kind: 'userError', message: 'still locked' }];
    await setup({ n: 1, failures: { TagsAdd: twice() } });
    h.clock.ms = Date.parse('2026-10-18T17:00:00Z');
    assert.equal((await h.run({ limit: 1 })).exitCode, 0);
    h.clock.ms = Date.parse('2026-10-19T17:00:00Z');
    const alone = await h.run({ limit: 1 });
    assert.equal(alone.exitCode, 1);
    assert.equal(h.statusOf(1), STATUS.CREATED, 'the tag is still missing');
    assert.ok(h.logged('INFO 没有待建卡的人'));
    assert.ok(h.logged('WARN 有 1 人补打 tag 失败，下次运行会再试'));
    assert.equal(h.logged('已完成补记和补打 tag'), false);
    const fixed = await h.run({ limit: 1 }); // the tag works now
    assert.equal(fixed.exitCode, 0);
    assert.equal(h.statusOf(1), STATUS.DONE);
    assert.ok(h.logged('INFO 没有待建卡的人；已完成补记和补打 tag（如果有）'));

    h.fake.restore();
    h.cleanup();
    resetClient();
    await setup({ n: 3, failures: { TagsAdd: twice() } });
    h.clock.ms = Date.parse('2026-10-18T17:00:00Z');
    assert.equal((await h.run({ limit: 1 })).exitCode, 0);
    h.clock.ms = Date.parse('2026-10-19T17:00:00Z');
    const waiting = await h.run({ limit: 1 });
    assert.equal(waiting.exitCode, 2);
    assert.ok(h.logged('WARN 有 1 人补打 tag 失败，下次运行会再试'));
    assert.ok(h.logged('还有 2 人待建卡。'));
    assert.equal(h.logged('已完成补记和补打 tag'), false);
  });

  it('an open outcome on or after the expiry date: the texts say it is settled as not created, never that it is issued again', async () => {
    await setup({ n: 2 });
    h.clock.ms = Date.parse('2026-10-19T17:00:00Z');
    crashedCreateStart(1, h.clock.ms - 2 * MINUTE); // 2 minutes ago: too early to tell
    const first = await h.run({ limit: 1 });
    assert.equal(first.exitCode, 1);
    const stop = h.log.lines.find((l) => l.includes('之后再运行'));
    assert.ok(stop.endsWith(`请在 2026-10-19 10:08（店铺时间）之后再运行：那时仍查不到的，程序会确认没有建成；${NO_NEW_CARDS}。也可以先在 Shopify 后台的礼品卡列表里按客户核对。`), stop);

    h.clock.ms += 15 * MINUTE;
    const second = await h.run({ limit: 1 });
    assert.equal(second.exitCode, 2);
    assert.deepEqual(h.createdFor(), []);
    const [none] = h.entriesFor(1, 'reconcile.none');
    assert.equal(none.note, `超过 10 分钟仍查不到这张卡，确认未建成；${NO_NEW_CARDS}`);
    assert.ok(h.logged(`超过 10 分钟仍查不到这张卡，确认没有建成；${NO_NEW_CARDS}`));
    assert.ok(h.logged(`确认上次没有建成：1 人；${NO_NEW_CARDS}`));
    assert.ok(h.logged('还有 2 人待建卡。'));
    assert.equal(h.log.lines.some((l) => /重新发放|重新排队/.test(l)), false, h.log.lines.join('\n'));
  });

  it('an open outcome just before midnight: the next run settles it on the expiry date, so no re-issue is promised', async () => {
    await setup({ n: 2, failures: { GiftCardCreate: [{ kind: 'network' }] } });
    h.clock.ms = Date.parse('2026-10-19T06:55:00Z'); // 23:55 on 10/18
    const lost = await h.run({ limit: 1, reconcileWaitMs: 0 });
    assert.equal(lost.exitCode, 1);
    assert.ok(h.logged(`请在 2026-10-19 00:05（店铺时间）之后再运行：程序会先核对这张卡；查不到的确认没有建成，${NO_NEW_CARDS}`));
    h.clock.ms = Date.parse('2026-10-19T06:58:00Z'); // 23:58 on 10/18: still too early to settle
    const early = await h.run({ limit: 1 });
    assert.equal(early.exitCode, 1);
    assert.ok(h.logged(`请在 2026-10-19 00:05（店铺时间）之后再运行：那时仍查不到的，程序会确认没有建成；${NO_NEW_CARDS}。`));
    assert.equal(h.log.lines.some((l) => /重新发放|重新排队/.test(l)), false, h.log.lines.join('\n'));
  });
});

describe('issue: no UTC-evening warning any more (the email copy no longer depends on the send date)', () => {
  it('a real run in the Los Angeles evening creates cards without any warning about UTC or the copy', async () => {
    await setup({ n: 3 });
    h.clock.ms = Date.parse('2026-10-12T00:30:00Z'); // 17:30 on 10/11 in Los Angeles, already 10/12 in UTC
    const { exitCode } = await h.run({ limit: 2 });
    assert.equal(exitCode, 0);
    assert.deepEqual(h.createdFor(), [1, 2]);
    assert.equal(h.log.lines.some((l) => l.startsWith('WARN ')), false, h.log.lines.join('\n'));
    // The evening before the expiry date is still a normal day: the store's calendar decides, not UTC's.
    h.clock.ms = Date.parse('2026-10-19T00:30:00Z'); // 17:30 on 10/18
    assert.equal((await h.run({ limit: 1, config: { ...h.config, dryRun: true } })).exitCode, 0);
    assert.equal(h.log.lines.some((l) => l.startsWith('WARN ')), false, h.log.lines.join('\n'));
    assert.equal(h.previews.length, 1, 'the dry run still previews the first email');
  });
});

describe('issue --repair-only', () => {
  it('an open create outcome: the texts never promise that this run issues it again', async () => {
    await setup({ n: 3, failures: { GiftCardCreate: [{ kind: 'network' }] } });
    const t0 = h.clock.ms;
    assert.equal((await h.run({ limit: 2 })).exitCode, 1); // customer 1: outcome unknown
    h.clock.ms = t0 + 5 * MINUTE;
    assert.equal((await h.run({ repairOnly: true })).exitCode, 1, 'too early to settle');
    assert.ok(h.logged('那时仍查不到的，程序会确认没有建成；本次是 --repair-only，不建新卡；之后正常运行 issue 时才会给他建卡。'));
    h.clock.ms = t0 + 25 * MINUTE;
    const settled = await h.run({ repairOnly: true });
    assert.equal(settled.exitCode, 0);
    assert.equal(h.statusOf(1), STATUS.PENDING);
    assert.equal(h.fake.opsNamed('GiftCardCreate').length, 1, 'repair-only created nothing');
    const [none] = h.entriesFor(1, 'reconcile.none');
    assert.match(none.note, /本次是 --repair-only，不建新卡；之后正常运行 issue 时才会给他建卡/);
    assert.equal(h.logged('重新排队发放'), false);
    assert.ok(h.logged('确认上次没有建成：1 人；本次是 --repair-only，不建新卡'));
  });

  it('before the expiry date, without --limit: adds the missing tag, creates no card, exit 0', async () => {
    await setup({ n: 5, failures: { TagsAdd: [{ kind: 'userError', message: 'locked' }] } });
    assert.equal((await h.run({ limit: 2 })).exitCode, 0); // 10/5: #1's tag failed
    assert.equal(h.statusOf(1), STATUS.CREATED);
    h.clock.ms = Date.parse('2026-10-06T17:00:00Z');
    const { exitCode, summary } = await h.run({ repairOnly: true });
    assert.equal(exitCode, 0);
    assert.deepEqual(h.createdFor(), [1, 2], 'no card from the list being given up');
    assert.equal(h.statusOf(1), STATUS.DONE, 'the missing tag was added');
    for (const k of [3, 4, 5]) assert.equal(h.statusOf(k), STATUS.PENDING);
    // The run's last line before the summary block.
    const heading = h.log.lines.indexOf('INFO —— 本次结果（批次 2）——');
    assert.equal(h.log.lines[heading - 1], 'INFO 只做了补记和补打 tag（--repair-only），没有建新卡。');
    assert.ok(h.logged('只补记和补打 tag，不建新卡（--repair-only）'));
    assert.equal(summary.repairOnly, true);
    assertSummary(summary, { attempted: 0, created: 0, tagFixed: 1 });
    assert.equal(summary.newCardsRefused, undefined);
    const start = h.journal().filter((e) => e.op === 'run.start').at(-1);
    assertSummary(start, { dryRun: false, limit: null, options: { retryFailed: false, repairOnly: true } });
    const end = h.journal().at(-1);
    assert.equal(end.summary.repairOnly, true);
    assert.equal(end.exitCode, 0);
    assert.equal(h.previews.length, 0);
  });

  it('records the cards Shopify has but the journal lost, and tags them; nobody else gets a card', async () => {
    await setup({ n: 4 });
    assert.equal((await h.run({ limit: 2 })).exitCode, 0);
    fs.rmSync(h.paths.journal);
    for (const k of [1, 2]) h.customer(k).tags = [];
    h.clock.ms = Date.parse('2026-10-06T17:00:00Z');
    const { exitCode, summary } = await h.run({ repairOnly: true });
    assert.equal(exitCode, 0);
    for (const k of [1, 2]) {
      assert.equal(h.entriesFor(k, 'reconcile.found')[0].source, 'preflight');
      assert.deepEqual(h.customer(k).tags, [TAG]);
      assert.equal(h.statusOf(k), STATUS.DONE);
    }
    assert.deepEqual(h.createdFor(), [1, 2]);
    assertSummary(summary, { reconciled: 2, tagFixed: 2, created: 0 });
  });

  it('exit 1 while a missing tag still cannot be added', async () => {
    await setup({ n: 2, failures: { TagsAdd: [{ kind: 'userError', message: 'locked' }, { kind: 'userError', message: 'still locked' }] } });
    assert.equal((await h.run({ limit: 1 })).exitCode, 0);
    h.clock.ms = Date.parse('2026-10-06T17:00:00Z');
    const { exitCode, summary } = await h.run({ repairOnly: true });
    assert.equal(exitCode, 1);
    assert.equal(summary.tagFailed, 1);
    assert.ok(h.logged('WARN 有 1 人补打 tag 失败，下次运行会再试'));
    assert.ok(h.logged('INFO 只做了补记和补打 tag（--repair-only），没有建新卡。'));
    assert.deepEqual(h.createdFor(), [1]);
  });

  it('a dry run needs no --limit either: reads only, no preview, exit 0', async () => {
    await setup({ n: 3, env: { DRY_RUN: 'true' } });
    const { exitCode, summary } = await h.run({ repairOnly: true });
    assert.equal(exitCode, 0);
    assert.ok(h.logged('INFO 预演：真实运行只会补记和补打 tag（--repair-only），不会建新卡。'));
    assert.equal(h.previews.length, 0);
    assert.equal(h.logged('这一批会给'), false);
    assert.deepEqual(h.journal().map((e) => e.op), ['run.start', 'run.end']);
    assertSummary(summary, { dryRun: true, repairOnly: true, attempted: 0 });
    assert.equal(h.fake.opsNamed('GiftCardCreate').length + h.fake.opsNamed('TagsAdd').length, 0);
  });

  it('runs on any day: before LAUNCH_DATE, and after the expiry date with people waiting (exit 0, not 2)', async () => {
    await setup({ n: 2 });
    h.clock.ms = Date.parse('2026-10-03T18:00:00Z');
    assert.equal((await h.run({ repairOnly: true })).exitCode, 0);
    h.clock.ms = Date.parse('2026-10-20T17:00:00Z');
    const late = await h.run({ repairOnly: true });
    assert.equal(late.exitCode, 0);
    assert.equal(h.logged('ERROR'), false);
    assert.deepEqual(h.createdFor(), []);
  });

  it('settles an old open outcome as not created and still creates nothing', async () => {
    await setup({ n: 2 });
    crashedCreateStart(1, h.clock.ms);
    h.clock.ms += 11 * MINUTE;
    const { exitCode, summary } = await h.run({ repairOnly: true });
    assert.equal(exitCode, 0);
    assert.equal(summary.reconciledNone, 1);
    assert.equal(h.statusOf(1), STATUS.PENDING);
    assert.deepEqual(h.createdFor(), []);
  });

  it('cannot be combined with --retry-failed (exit 2, nothing touched)', async () => {
    await setup({ n: 2 });
    const { exitCode } = await h.run({ repairOnly: true, retryFailed: true });
    assert.equal(exitCode, 2);
    assert.ok(h.logged('ERROR --repair-only 和 --retry-failed 不能一起用：--repair-only 只补记和补打 tag，不建卡'));
    assert.equal(h.fake.state.calls.token.length, 0);
    assert.equal(fs.existsSync(h.paths.journal), false);
    assert.equal(h.reports.length, 0);
  });
});

describe('issue: lost answers and leftovers', () => {
  it('a card created whose answer was lost is found inline: one card, recorded, tagged', async () => {
    await setup({ n: 3, failures: { GiftCardCreate: [{ kind: 'appliedThenLost' }] } });
    const { exitCode, summary } = await h.run({ limit: 1 });
    assert.equal(exitCode, 0);
    assert.equal(h.fake.opsNamed('GiftCardCreate').length, 1, 'never retried blindly');
    assert.equal(h.cardsOf(1).length, 1);
    assert.deepEqual(h.entriesFor(1).map((e) => e.op), ['create.start', 'reconcile.found', 'tag.ok']);
    const [found] = h.entriesFor(1, 'reconcile.found');
    assertSummary(found, { source: 'issue-inline', giftCardId: h.cardsOf(1)[0].id, amountCents: 1077, last4: h.cardsOf(1)[0].lastCharacters });
    assert.equal(h.statusOf(1), STATUS.DONE);
    assertSummary(summary, { created: 1, tagged: 1, reconciled: 0, unknown: 0, amountCents: 1077 });
  });

  it('a network failure that created nothing: unknown and stop; still stopped before 10 minutes; retried after', async () => {
    await setup({ n: 3, failures: { GiftCardCreate: [{ kind: 'network' }] } });
    const t0 = h.clock.ms;
    const sleeps = [];
    const first = await h.run({ limit: 2, sleep: async (ms) => { sleeps.push(ms); } });
    assert.equal(first.exitCode, 1);
    assert.equal(first.summary.unknown, 1);
    assert.deepEqual(sleeps, Array(12).fill(10_000), 'looked for the card for 2 minutes');
    assert.deepEqual(h.entriesFor(1).map((e) => e.op), ['create.start', 'create.unknown']);
    assert.equal(h.statusOf(1), STATUS.UNKNOWN);
    assert.equal(h.fake.opsNamed('GiftCardCreate').length, 1, 'stopped: customer 2 was not attempted');
    assert.ok(h.logged('的建卡结果不明'));

    h.clock.ms = t0 + 5 * MINUTE;
    const second = await h.run({ limit: 2 });
    assert.equal(second.exitCode, 1);
    assert.equal(second.summary.stillUnknown, 1);
    assert.equal(h.fake.opsNamed('GiftCardCreate').length, 1, 'no create while an outcome is open');
    assert.equal(h.statusOf(1), STATUS.UNKNOWN);
    assert.ok(h.logged('还有 1 人的建卡结果需要确认'));
    assert.ok(h.logged('请在 2026-10-05 10:10（店铺时间）之后再运行：那时仍查不到的，程序会确认没有建成并重新发放；也可以先在 Shopify 后台的礼品卡列表里按客户核对。'));
    assert.ok(h.logged('请在 2026-10-05 10:10（店铺时间）之后再运行：程序会先核对这张卡，查不到才会重新发放'), 'before the expiry date a re-issue is promised');

    h.clock.ms = t0 + 11 * MINUTE;
    const third = await h.run({ limit: 2 });
    assert.equal(third.exitCode, 0);
    const [none] = h.entriesFor(1, 'reconcile.none');
    assert.equal(none.note, '超过 10 分钟仍查不到这张卡，确认未建成，可以重试');
    assert.equal(third.summary.reconciledNone, 1);
    assert.ok(h.logged('超过 10 分钟仍查不到这张卡，确认没有建成，重新排队发放'));
    assert.ok(h.logged('确认上次没有建成、重新排队：1 人'));
    assert.deepEqual(h.createdFor(), [1, 2]);
    assert.equal(h.cardsOf(1).length, 1);
    assert.equal(h.statusOf(1), STATUS.DONE);
  });

  it('a card hidden by search lag is found by the next run: recorded, tagged, never created twice', async () => {
    await setup({ n: 3, failures: { GiftCardCreate: [{ kind: 'appliedThenLost' }] } });
    h.fake.state.hiddenCardIds.add('gid://shopify/GiftCard/1001'); // the first card the fake creates
    const first = await h.run({ limit: 1, reconcileWaitMs: 30_000 });
    assert.equal(first.exitCode, 1);
    assert.equal(h.cardsOf(1).length, 1, 'the card does exist in Shopify');
    assert.equal(h.statusOf(1), STATUS.UNKNOWN);

    h.fake.state.hiddenCardIds.clear(); // the search index caught up
    const second = await h.run({ limit: 1 });
    assert.equal(second.exitCode, 0);
    const [found] = h.entriesFor(1, 'reconcile.found');
    assertSummary(found, { source: 'issue-reconcile', giftCardId: 'gid://shopify/GiftCard/1001', amountCents: 1077 });
    assert.equal(h.cardsOf(1).length, 1);
    assert.equal(h.statusOf(1), STATUS.DONE);
    assert.deepEqual(h.createdFor(), [1, 2]);
    assertSummary(second.summary, { reconciled: 1, tagFixed: 1, created: 1 });
  });

  it('never takes another customer\'s card for this one, even if Shopify ignored the customer filter', async () => {
    await setup({ n: 3 });
    const t0 = h.clock.ms;
    assert.equal((await h.run({ limit: 1 })).exitCode, 0); // customer 1's card, created at t0
    const inner = globalThis.fetch;
    globalThis.fetch = async (url, init) => {
      const body = String(url).endsWith('/graphql.json') ? JSON.parse(init.body) : null;
      if (body && /query FindCampaignCards/.test(body.query)) {
        body.variables.query = body.variables.query.replace(/\s*customer_id:\d+/, ''); // filter silently dropped
        return inner(url, { ...init, body: JSON.stringify(body) });
      }
      return inner(url, init);
    };
    h.fake.state.failures.GiftCardCreate = [{ kind: 'network' }];
    h.clock.ms = t0 + 10_000; // customer 1's card lies inside customer 2's lookup window
    const { exitCode } = await h.run({ limit: 1, reconcileWaitMs: 0 });
    assert.equal(exitCode, 1);
    assert.equal(h.entriesFor(2, 'reconcile.found').length, 0);
    assert.equal(h.entriesFor(2, 'create.unknown').length, 1);
    assert.equal(h.statusOf(2), STATUS.UNKNOWN);
  });

  it('a run killed after Shopify created the card: the next run records the card and only tags', async () => {
    await setup({ n: 2 });
    const t = crashedCreateStart(1, h.clock.ms);
    h.fake.state.giftCards.push(campaignCard({ id: 'gid://shopify/GiftCard/777', customer: 1, createdAt: t }));
    h.clock.ms += 3 * MINUTE;
    const { exitCode, summary } = await h.run({ limit: 1 });
    assert.equal(exitCode, 0);
    assert.equal(h.entriesFor(1, 'reconcile.found')[0].source, 'issue-reconcile');
    assert.equal(h.states().get(gid('Customer', 1)).giftCardId, 'gid://shopify/GiftCard/777');
    assert.equal(h.statusOf(1), STATUS.DONE);
    assert.equal(h.cardsOf(1).length, 1);
    assert.deepEqual(h.createdFor(), [2]);
    assert.equal(summary.batch, 2);
  });

  it('a run killed before its card shows up: marked unknown, and no new card is created', async () => {
    await setup({ n: 2 });
    crashedCreateStart(1, h.clock.ms);
    h.clock.ms += 2 * MINUTE;
    const { exitCode } = await h.run({ limit: 1 });
    assert.equal(exitCode, 1);
    assert.equal(h.entriesFor(1, 'create.unknown').length, 1);
    assert.equal(h.statusOf(1), STATUS.UNKNOWN);
    assert.equal(h.fake.opsNamed('GiftCardCreate').length, 0);
  });
});

describe('issue: rejected creates', () => {
  it('userErrors: recorded as failed and the run goes on', async () => {
    await setup({ n: 4, failures: { GiftCardCreate: [{ kind: 'userError', message: 'Customer is invalid' }] } });
    const { exitCode, summary } = await h.run({ limit: 3 });
    assert.equal(exitCode, 0);
    assert.equal(h.statusOf(1), STATUS.FAILED);
    assert.match(h.entriesFor(1, 'create.fail')[0].error, /Customer is invalid/);
    assert.equal(h.cardsOf(1).length, 0);
    assert.deepEqual(h.fake.state.giftCards.map((g) => g.customer.id), [gid('Customer', 2), gid('Customer', 3)]);
    assertSummary(summary, { attempted: 3, failed: 1, created: 2 });
    assert.ok(h.logged('失败：1 人'));
  });

  it('5 failures in a row stop the run; --retry-failed then retries only those rows', async () => {
    const rejections = Array.from({ length: 5 }, () => ({ kind: 'userError', message: 'Gift cards are disabled' }));
    await setup({ n: 7, failures: { GiftCardCreate: rejections } });
    const first = await h.run({ limit: 7 });
    assert.equal(first.exitCode, 1);
    assert.equal(first.summary.failed, 5);
    assert.equal(h.fake.opsNamed('GiftCardCreate').length, 5);
    assert.equal(h.fake.state.giftCards.length, 0);
    for (const k of [1, 2, 3, 4, 5]) assert.equal(h.statusOf(k), STATUS.FAILED);
    for (const k of [6, 7]) assert.equal(h.statusOf(k), STATUS.PENDING);
    assert.ok(h.logged('连续 5 次建卡被 Shopify 拒绝'));

    const pendingRun = await h.run({ limit: 1 });
    assert.deepEqual(h.createdFor().slice(5), [6], 'a normal run does not retry failed rows');
    assert.equal(pendingRun.exitCode, 0);

    const retry = await h.run({ limit: 10, retryFailed: true });
    assert.equal(retry.exitCode, 0);
    assert.equal(retry.summary.created, 5);
    for (const k of [1, 2, 3, 4, 5, 6]) {
      assert.equal(h.statusOf(k), STATUS.DONE);
      assert.equal(h.cardsOf(k).length, 1);
    }
    assert.equal(h.statusOf(7), STATUS.PENDING);
    assert.equal(h.cardsOf(7).length, 0);
    assert.deepEqual(h.journal().filter((e) => e.op === 'run.start').at(-1).options, { retryFailed: true, repairOnly: false });
  });

  it('throttled on every retry: create.rejected, the run stops, the row is pending again', async () => {
    await setup({ n: 3, failures: { GiftCardCreate: Array.from({ length: 6 }, () => ({ kind: 'throttle' })) } });
    const { exitCode, summary } = await h.run({ limit: 3 });
    assert.equal(exitCode, 1);
    assert.equal(summary.rejected, 1);
    assert.equal(h.fake.opsNamed('GiftCardCreate').length, 6, 'first try + 5 retries inside gql');
    assert.deepEqual(h.entriesFor(1).map((e) => e.op), ['create.start', 'create.rejected']);
    assert.equal(h.statusOf(1), STATUS.PENDING);
    assert.equal(h.fake.state.giftCards.length, 0);
    assert.equal(h.entriesFor(2).length, 0, 'stopped before the next customer');

    const next = await h.run({ limit: 1 });
    assert.equal(next.exitCode, 0);
    assert.deepEqual(h.createdFor(), [1]);
  });
});

describe('issue: tags and Shopify-side guards', () => {
  it('a failed tagsAdd leaves the row "created"; the next run only adds the tag', async () => {
    await setup({ n: 1, failures: { TagsAdd: [{ kind: 'userError', message: 'tags are locked' }] } });
    const first = await h.run({ limit: 1 });
    assert.equal(first.exitCode, 0);
    assert.equal(h.statusOf(1), STATUS.CREATED);
    assert.equal(h.entriesFor(1, 'tag.fail').length, 1);
    assert.equal(first.summary.tagFailed, 1);

    const second = await h.run({ limit: 1 });
    assert.equal(second.exitCode, 0);
    assertSummary(second.summary, { tagFixed: 1, created: 0, attempted: 0 });
    assert.equal(h.fake.opsNamed('GiftCardCreate').length, 1);
    assert.equal(h.cardsOf(1).length, 1);
    assert.equal(h.statusOf(1), STATUS.DONE);
    assert.deepEqual(h.customer(1).tags, [TAG]);
  });

  it('journal deleted: customers already carrying SENT_TAG are skipped even when their cards are not visible', async () => {
    await setup({ n: 4 });
    assert.equal((await h.run({ limit: 2 })).exitCode, 0);
    fs.rmSync(h.paths.journal);
    for (const g of h.fake.state.giftCards) h.fake.state.hiddenCardIds.add(g.id);

    const { exitCode, summary } = await h.run({ limit: 2 });
    assert.equal(exitCode, 0);
    assert.deepEqual(summary.skipped, { 'already-tagged': 2 });
    for (const k of [1, 2]) {
      assert.equal(h.entriesFor(k, 'skip')[0].reason, 'already-tagged');
      assert.equal(h.cardsOf(k).length, 1);
    }
    assert.deepEqual(h.createdFor(), [1, 2, 3, 4]);
  });

  it('journal deleted: an existing campaign card is recorded (preflight) with its tag, never a second card', async () => {
    await setup({ n: 4 });
    assert.equal((await h.run({ limit: 2 })).exitCode, 0);
    fs.rmSync(h.paths.journal);
    const tagCallsBefore = h.fake.opsNamed('TagsAdd').length;

    const { exitCode, summary } = await h.run({ limit: 2 });
    assert.equal(exitCode, 0);
    for (const k of [1, 2]) {
      const [found] = h.entriesFor(k, 'reconcile.found');
      assertSummary(found, { source: 'preflight', giftCardId: h.cardsOf(k)[0].id });
      assert.equal(h.entriesFor(k, 'tag.ok')[0].note, 'already tagged in Shopify');
      assert.equal(h.statusOf(k), STATUS.DONE);
      assert.equal(h.cardsOf(k).length, 1);
    }
    assert.equal(h.fake.opsNamed('TagsAdd').length - tagCallsBefore, 2, 'only the 2 new cards were tagged');
    assert.deepEqual(h.createdFor(), [1, 2, 3, 4]);
    assertSummary(summary, { reconciled: 2, tagFixed: 2, created: 2 });
  });

  it('journal deleted and tags removed: the campaign-card guard still prevents a second card', async () => {
    await setup({ n: 4 });
    assert.equal((await h.run({ limit: 2 })).exitCode, 0);
    fs.rmSync(h.paths.journal);
    for (const k of [1, 2]) h.customer(k).tags = [];

    const { exitCode, summary } = await h.run({ limit: 2 });
    assert.equal(exitCode, 0);
    for (const k of [1, 2]) {
      assert.equal(h.entriesFor(k, 'reconcile.found')[0].source, 'preflight');
      assert.equal(h.cardsOf(k).length, 1);
      assert.deepEqual(h.customer(k).tags, [TAG], 'the tag is added back');
      assert.equal(h.statusOf(k), STATUS.DONE);
    }
    assert.deepEqual(h.createdFor(), [1, 2, 3, 4]);
    assertSummary(summary, { reconciled: 2, tagFixed: 2, created: 2 });
  });
});

describe('issue: interruption and the Excel', () => {
  it('Ctrl+C after the first customer: exit 130, the journal has no dangling create.start', async () => {
    await setup({ n: 4 });
    const controller = new AbortController();
    const inner = globalThis.fetch;
    globalThis.fetch = async (url, init) => {
      const res = await inner(url, init);
      if (String(init?.body ?? '').includes('mutation TagsAdd')) controller.abort(); // after customer 1 is tagged
      return res;
    };
    const { exitCode, summary } = await h.run({ limit: 3, signal: controller.signal });
    assert.equal(exitCode, 130);
    assert.deepEqual(h.createdFor(), [1]);
    assert.equal(summary.created, 1);

    const journal = h.journal();
    assert.equal(journal.filter((e) => e.op === 'create.start').length, 1);
    assert.ok([...foldJournal(journal).customers.values()].every((s) => s.status !== STATUS.IN_PROGRESS));
    assert.equal(h.statusOf(1), STATUS.DONE);
    for (const k of [2, 3, 4]) assert.equal(h.statusOf(k), STATUS.PENDING);
    assertSummary(journal.at(-1), { op: 'run.end', exitCode: 130 });
    assert.equal(h.reports.length, 1, 'the Excel is still regenerated');
    assert.ok(h.logged('收到中断信号'));
  });

  it('prints the English workbook path right after the Chinese one; nothing extra when fileEn is null', async () => {
    await setup({ n: 2 });
    const { exitCode } = await h.run({ limit: 1 });
    assert.equal(exitCode, 0);
    const excelAt = h.log.lines.indexOf(`INFO Excel 已更新：${h.paths.excel}`);
    assert.ok(excelAt >= 0, h.log.lines.join('\n'));
    assert.equal(h.log.lines[excelAt + 1], `INFO 英文版 Excel 已更新：${h.paths.excelEn}`, h.log.lines.join('\n'));

    // The English edition failed (writeReport already warned): only the Chinese path, exit code unchanged.
    const mark = h.log.lines.length;
    const englishFailed = async (o) => {
      o.log.warn('英文版 Excel 没有生成：disk full');
      return { file: h.paths.excel, fileEn: null, out: null, outEn: null, warnings: ['英文版 Excel 没有生成：disk full'] };
    };
    assert.equal((await h.run({ limit: 1, writeReport: englishFailed })).exitCode, 0);
    const lines = h.log.lines.slice(mark);
    assert.ok(lines.includes(`INFO Excel 已更新：${h.paths.excel}`), lines.join('\n'));
    assert.equal(lines.filter((l) => l.includes('英文版 Excel 已更新')).length, 0, lines.join('\n'));
    assert.equal(lines.filter((l) => l === 'WARN 英文版 Excel 没有生成：disk full').length, 1, 'printed once, by writeReport');
  });

  it('a failing Excel export never changes the exit code', async () => {
    await setup({ n: 3, failures: { GiftCardCreate: Array.from({ length: 6 }, () => ({ kind: 'throttle' })) } });
    const broken = async () => { throw new Error('excel is locked'); };
    assert.equal((await h.run({ limit: 1, writeReport: broken })).exitCode, 1);
    assert.equal((await h.run({ limit: 1, writeReport: broken })).exitCode, 0);
    assert.ok(h.logged('WARN Excel 没有更新：excel is locked'));
    assert.deepEqual(h.createdFor(), [1]);
  });
});
