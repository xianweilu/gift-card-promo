// The `select` command (src/select/index.js) and its read-only fetchers
// (src/select/export.js), end to end against the in-process fake Shopify.
import { test, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

import { runSelect, SNAPSHOT_MAX_AGE_MS, CARD_LOOKBACK_MS, ISSUING_STARTED_MESSAGE, cardsExistMessage } from '../src/select/index.js';
import { exportCustomers, fetchActiveOrders, fetchOrderHistory, readSnapshot, textLines } from '../src/select/export.js';
import { initClient, resetClient } from '../src/shopify.js';
import { CUSTOMER_EXPORT_BULK } from '../src/queries.js';
import { campaignPaths, acquireRunLock, appendJournal, readJournal, readJson } from '../src/campaign.js';
import { installFakeShopify, BULK_URL } from './fake-shopify.js';
import { NOW_ISO, NOW_MS, CUTOFF_MS, gid, makeOrder, makeCustomer, makeActiveOrder, testConfig, memoryLog } from './helpers.js';

const HOUR = 3_600_000;
const DAY = 24 * HOUR;
const TZ = 'America/Los_Angeles';
const CUTOFF_ISO = '2026-07-01T00:00:00-07:00';
// select looks for this campaign's gift cards created since NOW − 400 days (whole seconds, UTC).
const CARD_SINCE_ISO = '2025-08-27T20:00:00Z';
const CARD_QUERY = `created_at:>='${CARD_SINCE_ISO}' source:api_client`;
const noSleep = async () => {};
const at = (ms) => () => new Date(ms);

/** A gift card as Shopify lists it (FindCampaignCards node); `campaignId: null` = no campaign marker. */
function giftCard({ n, customer, createdAt, campaignId = '2026-10', amount = '10.77' }) {
  return {
    id: gid('GiftCard', n),
    createdAt,
    note: campaignId ? `gift-card-promo [campaign:${campaignId}]` : 'IBC2026 test',
    templateSuffix: 'gift-card-promo',
    enabled: true,
    lastCharacters: `x${String(n).slice(-3)}`,
    initialValue: { amount, currencyCode: 'USD' },
    balance: { amount, currencyCode: 'USD' },
    customer: { id: customer.id },
    expiresOn: '2026-10-19',
  };
}

// ---------------------------------------------------------------------------
// Fixture: one customer per rule the command has to get right
// ---------------------------------------------------------------------------

// Two spellings of one address (same normalised key: street + ZIP5 + country).
const SHARED_ADDRESS = { address1: '1422 Gardena Avenue', address2: '', city: 'Glendale', provinceCode: 'CA', zip: '91204-1234', countryCodeV2: 'US' };
const SHARED_ADDRESS_VARIANT = { address1: '1422 gardena ave.', address2: '', city: 'glendale', provinceCode: 'CA', zip: '91204', countryCodeV2: 'US' };

const ORDERED = makeCustomer({ n: 1, numberOfOrders: 3, amountSpent: '310.00', lastOrder: makeOrder({ n: 101, createdAt: '2026-03-01T12:00:00Z', total: '150.00' }) });
const NOT_SUBSCRIBED = makeCustomer({ n: 2, marketingState: 'NOT_SUBSCRIBED', lastOrder: makeOrder({ n: 102, createdAt: '2026-02-01T12:00:00Z', total: '80.00' }) });
const EXCLUDED_TAG = makeCustomer({ n: 3, tags: ['VIP', 'whs'], lastOrder: makeOrder({ n: 103, createdAt: '2026-01-10T12:00:00Z', total: '90.00' }) });
// Snapshot says May, but ORDERS_SINCE shows an order in August.
const RECENT = makeCustomer({ n: 4, lastOrder: makeOrder({ n: 104, createdAt: '2026-05-01T12:00:00Z', total: '40.00' }) });
const SAME_ADDRESS_AS_RECENT = makeCustomer({ n: 5, address: { ...RECENT.defaultAddress }, lastOrder: makeOrder({ n: 105, createdAt: '2026-01-15T12:00:00Z', total: '60.00' }) });
const DUP_KEPT = makeCustomer({ n: 6, address: SHARED_ADDRESS, lastOrder: makeOrder({ n: 106, createdAt: '2026-05-20T12:00:00Z', total: '350.00' }) });
const DUP_DROPPED = makeCustomer({ n: 7, address: SHARED_ADDRESS_VARIANT, lastOrder: makeOrder({ n: 107, createdAt: '2026-02-10T12:00:00Z', total: '90.00' }) });
// Last order is $0: the amount comes from the order before it (CUSTOMER_ORDERS follow-up).
const ZERO_LAST = makeCustomer({ n: 8, numberOfOrders: 2, lastOrder: makeOrder({ n: 108, createdAt: '2026-04-01T12:00:00Z', total: '0.00' }) });
const NEVER = makeCustomer({ n: 9, createdAt: '2025-06-01T00:00:00Z' });
const CUSTOMERS = [ORDERED, NOT_SUBSCRIBED, EXCLUDED_TAG, RECENT, SAME_ADDRESS_AS_RECENT, DUP_KEPT, DUP_DROPPED, ZERO_LAST, NEVER];

const ACTIVE_ORDERS = [
  makeActiveOrder({ n: 204, customer: RECENT, createdAt: '2026-08-15T12:00:00Z' }),
  makeActiveOrder({ n: 201, customer: ORDERED, createdAt: '2026-09-01T12:00:00Z', cancelled: true }), // cancelled: does not count
  makeActiveOrder({ n: 209, customer: NEVER, createdAt: '2026-09-10T12:00:00Z', test: true }), // test order: does not count
  makeActiveOrder({ n: 206, customer: DUP_KEPT, createdAt: '2026-06-20T12:00:00Z' }), // before the cutoff: the search must leave it out
];
const ZERO_LAST_HISTORY = [
  makeOrder({ n: 108, createdAt: '2026-04-01T12:00:00Z', total: '0.00' }),
  makeOrder({ n: 98, createdAt: '2026-02-01T12:00:00Z', total: '23.00' }),
];
const HISTORY = { [ZERO_LAST.id]: ZERO_LAST_HISTORY };

// Average of the kept "ordered" bases: (350 + 150 + 23) / 3 = $174.33 → 10% = $17.43 → $19.77.
const EXPECTED_RECIPIENTS = [
  { customerId: DUP_KEPT.id, kind: 'ordered', amountCents: 1977, rawCents: 3500, tier: 2 },
  { customerId: ORDERED.id, kind: 'ordered', amountCents: 1533, rawCents: 1500, tier: 1 },
  { customerId: ZERO_LAST.id, kind: 'ordered', amountCents: 1077, rawCents: 230, tier: 0 },
  { customerId: NEVER.id, kind: 'never', amountCents: 1977, rawCents: 1743, tier: 2 },
];
const EXPECTED_TOTAL_CENTS = 1977 + 1533 + 1077 + 1977;
const EXPECTED_NOT_SELECTED = [
  [EXCLUDED_TAG.id, 'excluded-tag'],
  [RECENT.id, 'recent-order'],
  [SAME_ADDRESS_AS_RECENT.id, 'active-address'],
  [DUP_DROPPED.id, 'duplicate-address'],
];

const brief = (selection) => selection.recipients.map((r) => ({ customerId: r.customerId, kind: r.kind, amountCents: r.amountCents, rawCents: r.rawCents, tier: r.tier }));
const withoutSnapshot = ({ snapshot, ...rest }) => rest;

// ---------------------------------------------------------------------------
// Harness
// ---------------------------------------------------------------------------

const cleanups = [];
afterEach(() => {
  while (cleanups.length) cleanups.pop()();
  resetClient();
});

function setup({ env = {}, customers = CUSTOMERS, orders = ACTIVE_ORDERS, orderHistory = HISTORY, giftCards = [], bulk, failures, searchWarnings } = {}) {
  const cfg = testConfig(env);
  cleanups.push(() => cfg.cleanup());
  const fake = installFakeShopify({ customers, orders, orderHistory, giftCards, bulk, failures, searchWarnings, now: () => NOW_MS });
  cleanups.push(() => fake.restore());
  const log = memoryLog();
  const paths = campaignPaths(cfg.config);
  const reports = [];
  const writeReport = async (args) => {
    reports.push({ ...args, runLockHeld: fs.existsSync(paths.runLock), selectionExists: fs.existsSync(paths.selection) });
    return { file: args.paths.excel, fileEn: args.paths.excelEn, out: null, outEn: null, warnings: [] };
  };
  const run = (options = {}) => runSelect({ config: cfg.config, log, now: at(NOW_MS), sleep: noSleep, writeReport, bulkPollMs: 1_000, ...options });
  return { config: cfg.config, fake, log, paths, reports, run };
}

/** Wrap the fake's fetch for one test (undone in afterEach). */
function wrapFetch(wrapper) {
  const inner = globalThis.fetch;
  globalThis.fetch = (url, init = {}) => wrapper(String(url), init, inner);
  cleanups.push(() => {
    globalThis.fetch = inner;
  });
}

const isGraphql = (url) => url.endsWith('/graphql.json');
const operationOf = (init) => /^\s*(?:query|mutation)\s+(\w+)/.exec(JSON.parse(init.body).query)?.[1];
const jsonResponse = (body) => new Response(JSON.stringify(body), { status: 200, headers: { 'content-type': 'application/json' } });
const linesOf = (log, level) => log.lines.filter((l) => l.startsWith(`${level} `));
const hasLine = (log, text, level = null) => log.lines.some((l) => (!level || l.startsWith(`${level} `)) && l.includes(text));
const snapshotIds = (paths) => fs.readFileSync(paths.snapshot, 'utf8').split('\n').filter(Boolean).map((l) => JSON.parse(l).id);

/** Direct client setup for the fetcher unit tests (no token exchange needed). */
function directClient(log) {
  initClient({ shop: 'teststore', apiVersion: '2026-07', token: 'fake-token', log, sleep: noSleep });
}

// ---------------------------------------------------------------------------
// select: bulk path
// ---------------------------------------------------------------------------

test('bulk export: selection.json has the expected recipients; journal, snapshot files and Excel are written', async () => {
  const t = setup({ bulk: { mode: 'ok', pollsBeforeComplete: 2 } });
  const result = await t.run();
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));

  const saved = readJson(t.paths.selection);
  assert.deepEqual(saved, result.selection);
  assert.equal(saved.mode, 'live');
  assert.equal(saved.campaignId, '2026-10');
  assert.equal(saved.createdAt, NOW_ISO);
  assert.deepEqual(saved.snapshot, { exportedAt: NOW_ISO, source: 'bulk', count: CUSTOMERS.length });
  assert.equal(saved.params.timezone, TZ);
  assert.equal(saved.params.cutoffIso, CUTOFF_ISO);
  assert.equal(saved.params.cutoffMs, CUTOFF_MS);
  assert.equal(saved.params.cutoffDate, '2026-07-01');

  assert.deepEqual(brief(saved), EXPECTED_RECIPIENTS);
  assert.deepEqual(saved.recipients.map((r) => r.seq), [1, 2, 3, 4]);
  const zero = saved.recipients[2];
  assert.equal(zero.basisWhy, 'last-zero');
  assert.equal(zero.basis.orderId, gid('Order', 98));
  assert.equal(zero.basis.totalCents, 2300);
  assert.equal(saved.recipients[3].neverReason, 'no-orders');
  assert.equal(saved.recipients[0].groupSize, 2);
  assert.equal(saved.averageCents, 17433); // whole cents: (350 + 150 + 23) / 3 = $174.333…
  assert.equal(saved.neverAmountCents, 1977);
  assert.equal(saved.stats.totalCents, EXPECTED_TOTAL_CENTS);
  assert.deepEqual(saved.notSelected.map((r) => [r.customerId, r.primaryCode]), EXPECTED_NOT_SELECTED);
  assert.equal(saved.notSelected[2].relatedCustomerId, RECENT.id);
  assert.equal(saved.funnel.total, CUSTOMERS.length);
  assert.equal(saved.funnel.unlisted, 1); // the unsubscribed customer: counted, not listed
  assert.deepEqual(
    saved.duplicates.map((d) => [d.keptCustomerId, d.members.map((m) => m.customerId)]),
    [[DUP_KEPT.id, [DUP_KEPT.id, DUP_DROPPED.id]]],
  );

  // What was asked of Shopify
  const submits = t.fake.opsNamed('RunCustomerExport');
  assert.equal(submits.length, 1);
  assert.equal(submits[0].variables.query, CUSTOMER_EXPORT_BULK);
  assert.equal(t.fake.opsNamed('BulkStatus').length, 2);
  assert.equal(t.fake.state.calls.bulkDownloads, 1);
  assert.equal(t.fake.state.bulk.cancelled, 0);
  assert.equal(t.fake.opsNamed('CustomersPage').length, 0);
  assert.deepEqual(t.fake.opsNamed('OrdersSince').map((o) => o.variables.query), [`created_at:>='${CUTOFF_ISO}'`]);
  assert.deepEqual(t.fake.opsNamed('CustomerOrders').map((o) => o.variables.id), [ZERO_LAST.id]);
  assert.equal(t.fake.opsNamed('ShopInfo').length, 0); // TIMEZONE is configured
  // Shopify is asked for this campaign's cards first, before anything is exported.
  assert.deepEqual(t.fake.opsNamed('FindCampaignCards').map((o) => o.variables.query), [CARD_QUERY]);
  const ops = t.fake.state.calls.ops.map((o) => o.op);
  assert.equal(ops[0], 'FindCampaignCards');
  assert.ok(ops.indexOf('FindCampaignCards') < ops.indexOf('RunCustomerExport'));

  // Snapshot files
  assert.deepEqual(readJson(t.paths.snapshotMeta), {
    exportedAt: NOW_ISO,
    source: 'bulk',
    count: CUSTOMERS.length,
    rootObjectCount: CUSTOMERS.length,
    cutoffIso: CUTOFF_ISO,
    cutoffMs: CUTOFF_MS,
    timezone: TZ,
    inactiveMonths: 3,
  });
  assert.deepEqual(snapshotIds(t.paths), CUSTOMERS.map((c) => c.id));
  assert.equal(fs.existsSync(`${t.paths.snapshot}.tmp`), false);
  assert.deepEqual(readJson(t.paths.activeOrders).map((o) => o.id), [gid('Order', 204), gid('Order', 201), gid('Order', 209)]);
  assert.deepEqual(readJson(t.paths.orderHistory), { [ZERO_LAST.id]: { orders: ZERO_LAST_HISTORY, percent: 10 } });

  // Journal: one select run
  const journal = readJournal(t.paths.journal);
  assert.deepEqual(journal.map((e) => e.op), ['run.start', 'run.end']);
  assert.equal(journal[0].command, 'select');
  assert.equal(journal[0].dryRun, false);
  assert.equal(journal[0].t, NOW_ISO);
  assert.equal(journal[1].run, journal[0].run);
  assert.equal(journal[1].exitCode, 0);
  assert.deepEqual(journal[1].summary, { recipients: 4, totalCents: EXPECTED_TOTAL_CENTS, source: 'bulk', snapshotReused: false });

  // Excel regenerated once, after the run lock was released
  assert.equal(t.reports.length, 1);
  assert.equal(t.reports[0].config, t.config);
  assert.equal(t.reports[0].paths.selection, t.paths.selection);
  assert.equal(typeof t.reports[0].now, 'function');
  assert.equal(t.reports[0].runLockHeld, false);
  assert.equal(fs.existsSync(t.paths.runLock), false);

  // Console output
  assert.ok(hasLine(t.log, '全店客户导出中：已完成 100', 'INFO'));
  assert.ok(hasLine(t.log, '全店客户：9'));
  assert.ok(hasLine(t.log, '入选：4 人'));
  assert.ok(hasLine(t.log, '有下单：3 人，合计 $45.87'));
  assert.ok(hasLine(t.log, '从没下单：1 人，合计 $19.77'));
  assert.ok(hasLine(t.log, '$19.77：2 人（有下单 1，从没下单 1），合计 $39.54'));
  assert.ok(hasLine(t.log, '$15.33：1 人'));
  assert.ok(hasLine(t.log, '$10.77：1 人'));
  assert.ok(hasLine(t.log, '总面额：$65.64'));
  assert.ok(hasLine(t.log, '本次新导出'));
  assert.ok(hasLine(t.log, `名单文件：${t.paths.selection}`));
  assert.ok(hasLine(t.log, `Excel：${t.paths.excel}`));
  // The English edition's path follows the Chinese one.
  const excelAt = t.log.lines.findIndex((l) => l.endsWith(`  Excel：${t.paths.excel}`));
  assert.ok(excelAt >= 0, t.log.lines.join('\n'));
  assert.equal(t.log.lines[excelAt + 1], `INFO   英文版 Excel：${t.paths.excelEn}`, t.log.lines.join('\n'));
  assert.deepEqual(linesOf(t.log, 'ERROR'), []);
  assert.deepEqual(linesOf(t.log, 'WARN'), []);
});

test('bulk download sends no Shopify token, and the signed URL never reaches a log line or a file', async () => {
  const t = setup();
  const downloads = [];
  wrapFetch((url, init, inner) => {
    if (url === BULK_URL) downloads.push(new Headers(init.headers ?? {}));
    return inner(url, init);
  });
  const result = await t.run();
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(downloads.length, 1);
  assert.equal(downloads[0].get('x-shopify-access-token'), null);
  assert.equal(downloads[0].get('authorization'), null);
  for (const line of t.log.lines) {
    assert.ok(!line.includes('fake-bulk.test'), `bulk URL logged: ${line}`);
    assert.ok(!line.includes('fake-token'), `token logged: ${line}`);
  }
  for (const file of [t.paths.snapshotMeta, t.paths.selection, t.paths.journal, t.paths.activeOrders]) {
    assert.ok(!fs.readFileSync(file, 'utf8').includes('fake-bulk.test'), `bulk URL stored in ${file}`);
  }
});

// ---------------------------------------------------------------------------
// select: fallbacks to pagination
// ---------------------------------------------------------------------------

test('bulk job FAILED → paginated fallback produces the same selection', async () => {
  const viaBulk = (await setup().run()).selection;

  const t = setup({ bulk: { mode: 'fail' } });
  const result = await t.run();
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.deepEqual(result.selection.snapshot, { exportedAt: NOW_ISO, source: 'paginated', count: CUSTOMERS.length });
  assert.deepEqual(withoutSnapshot(result.selection), withoutSnapshot(viaBulk));
  assert.equal(t.fake.opsNamed('CustomersPage').length, 1);
  assert.equal(t.fake.state.calls.bulkDownloads, 0);
  assert.equal(t.fake.state.bulk.cancelled, 0); // a failed job needs no cancel
  assert.ok(hasLine(t.log, 'FAILED', 'WARN'));
  assert.ok(hasLine(t.log, '改用普通分页导出', 'WARN'));
  const meta = readJson(t.paths.snapshotMeta);
  assert.equal(meta.source, 'paginated');
  assert.equal(meta.rootObjectCount, null);
  assert.deepEqual(snapshotIds(t.paths), CUSTOMERS.map((c) => c.id));
  assert.ok(hasLine(t.log, '普通分页导出', 'INFO'));
});

test('bulk job whose progress stalls is cancelled, then pagination takes over', async () => {
  const t = setup({ bulk: { mode: 'stall' } });
  // Polls at 1, 2, 3, 4 minutes: objectCount moves once (0 → 100), then stays put for 3 minutes.
  const result = await t.run({ bulkPollMs: 60_000, stallMs: 180_000 });
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(result.selection.snapshot.source, 'paginated');
  assert.deepEqual(brief(result.selection), EXPECTED_RECIPIENTS);
  assert.equal(t.fake.opsNamed('BulkStatus').length, 4);
  const cancels = t.fake.opsNamed('CancelBulk');
  assert.equal(cancels.length, 1);
  assert.equal(cancels[0].variables.id, t.fake.opsNamed('BulkStatus')[0].variables.id);
  assert.equal(t.fake.state.calls.bulkDownloads, 0);
  assert.ok(hasLine(t.log, '进度停在 100 已超过 3 分钟', 'WARN'));
});

test('bulk job running past the time limit is cancelled, then pagination takes over', async () => {
  const t = setup({ bulk: { mode: 'ok', pollsBeforeComplete: 1_000 } });
  const result = await t.run({ bulkPollMs: 60_000, bulkTimeoutMs: 300_000, stallMs: 24 * HOUR });
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(result.selection.snapshot.source, 'paginated');
  assert.equal(t.fake.opsNamed('BulkStatus').length, 5);
  assert.equal(t.fake.opsNamed('CancelBulk').length, 1);
  assert.ok(hasLine(t.log, '超过 5 分钟仍未完成', 'WARN'));
  assert.deepEqual(brief(result.selection), EXPECTED_RECIPIENTS);
});

test('bulk output whose line count differs from rootObjectCount → paginated fallback', async () => {
  const t = setup();
  wrapFetch(async (url, init, inner) => {
    const res = await inner(url, init);
    if (!isGraphql(url) || operationOf(init) !== 'BulkStatus') return res;
    const body = await res.json();
    if (body.data.bulkOperation.status === 'COMPLETED') body.data.bulkOperation.rootObjectCount = String(CUSTOMERS.length + 1);
    return jsonResponse(body);
  });
  const result = await t.run();
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(result.selection.snapshot.source, 'paginated');
  assert.equal(t.fake.state.calls.bulkDownloads, 1);
  assert.equal(t.fake.opsNamed('CustomersPage').length, 1);
  assert.ok(hasLine(t.log, '导出文件有 9 个客户，和 Shopify 报告的 10 个不一致', 'WARN'));
  assert.deepEqual(snapshotIds(t.paths), CUSTOMERS.map((c) => c.id));
  assert.equal(fs.existsSync(`${t.paths.snapshot}.tmp`), false);
  assert.deepEqual(brief(result.selection), EXPECTED_RECIPIENTS);
});

test('bulk output with a repeated customer id → paginated fallback', async () => {
  const t = setup();
  wrapFetch(async (url, init, inner) => {
    const res = await inner(url, init);
    if (url !== BULK_URL) return res;
    const lines = (await res.text()).trim().split('\n');
    lines[lines.length - 1] = lines[0]; // same line count, one id twice
    return new Response(`${lines.join('\n')}\n`, { status: 200 });
  });
  const result = await t.run();
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(result.selection.snapshot.source, 'paginated');
  assert.ok(hasLine(t.log, `客户 ${ORDERED.id} 在导出文件里出现了不止一次`, 'WARN'));
  assert.deepEqual(brief(result.selection), EXPECTED_RECIPIENTS);
});

test('bulk job rejected with userErrors → paginated fallback', async () => {
  const t = setup();
  wrapFetch((url, init, inner) => {
    if (isGraphql(url) && operationOf(init) === 'RunCustomerExport') {
      return Promise.resolve(jsonResponse({
        data: { bulkOperationRunQuery: { bulkOperation: null, userErrors: [{ field: null, message: 'A bulk query operation for this app and shop is already in progress', code: 'OPERATION_IN_PROGRESS' }] } },
      }));
    }
    return inner(url, init);
  });
  const result = await t.run();
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(result.selection.snapshot.source, 'paginated');
  assert.equal(t.fake.opsNamed('BulkStatus').length, 0);
  assert.ok(hasLine(t.log, 'already in progress', 'WARN'));
  assert.deepEqual(brief(result.selection), EXPECTED_RECIPIENTS);
});

test('a dropped bulk download is retried (URL redacted from the warning); a refused link falls back at once', async () => {
  const t = setup();
  let attempts = 0;
  wrapFetch((url, init, inner) => {
    if (url === BULK_URL && (attempts += 1) === 1) return Promise.reject(new TypeError(`fetch failed: socket hang up (${BULK_URL}?Signature=abc)`));
    return inner(url, init);
  });
  const slept = [];
  const result = await t.run({ sleep: async (ms) => { slept.push(ms); } });
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(result.selection.snapshot.source, 'bulk');
  assert.equal(attempts, 2);
  assert.ok(slept.includes(5_000));
  assert.ok(hasLine(t.log, '下载导出文件出错', 'WARN'));
  assert.ok(hasLine(t.log, '<链接已隐藏>', 'WARN'));
  assert.ok(!t.log.lines.some((l) => l.includes('fake-bulk.test') || l.includes('Signature=')));
  assert.equal(t.fake.opsNamed('CustomersPage').length, 0);

  const refused = setup();
  let refusedAttempts = 0;
  wrapFetch((url, init, inner) => {
    if (url === BULK_URL) {
      refusedAttempts += 1;
      return Promise.resolve(new Response('AccessDenied', { status: 403 }));
    }
    return inner(url, init);
  });
  const fallback = await refused.run();
  assert.equal(fallback.exitCode, 0, refused.log.lines.join('\n'));
  assert.equal(refusedAttempts, 1); // a refused link is not downloaded again
  assert.equal(fallback.selection.snapshot.source, 'paginated');
  assert.ok(hasLine(refused.log, 'HTTP 403', 'WARN'));
  assert.deepEqual(brief(fallback.selection), EXPECTED_RECIPIENTS);
});

test('a transient error while submitting the bulk job is retried instead of falling back', async () => {
  const t = setup({ failures: { RunCustomerExport: [{ kind: 'http', status: 502 }] } });
  const result = await t.run();
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(result.selection.snapshot.source, 'bulk');
  assert.equal(t.fake.opsNamed('RunCustomerExport').length, 2);
  assert.equal(t.fake.opsNamed('CustomersPage').length, 0);
  assert.ok(hasLine(t.log, '提交导出任务：Shopify 暂时出错', 'WARN'));
});

test('status polls tolerate a transient error; three in a row cancel the job and fall back', async () => {
  const t = setup({ failures: { BulkStatus: [{ kind: 'network' }] } });
  const result = await t.run();
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(result.selection.snapshot.source, 'bulk');
  assert.equal(t.fake.opsNamed('BulkStatus').length, 2);
  assert.equal(t.fake.opsNamed('CancelBulk').length, 0);
  assert.ok(hasLine(t.log, '查询导出进度出错', 'WARN'));

  const broken = setup({ failures: { BulkStatus: [{ kind: 'network' }, { kind: 'http', status: 502 }, { kind: 'network' }] } });
  const fallback = await broken.run();
  assert.equal(fallback.exitCode, 0, broken.log.lines.join('\n'));
  assert.equal(fallback.selection.snapshot.source, 'paginated');
  assert.equal(broken.fake.opsNamed('BulkStatus').length, 3);
  assert.equal(broken.fake.opsNamed('CancelBulk').length, 1);
  assert.ok(hasLine(broken.log, '连续 3 次查询导出进度出错', 'WARN'));
  assert.deepEqual(brief(fallback.selection), EXPECTED_RECIPIENTS);
});

test('bulk job COMPLETED without a URL (no customers) → empty snapshot and an empty selection', async () => {
  const t = setup({ customers: [], orders: [], orderHistory: {} });
  const result = await t.run();
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(t.fake.state.calls.bulkDownloads, 0);
  assert.equal(t.fake.opsNamed('CustomersPage').length, 0);
  assert.equal(fs.readFileSync(t.paths.snapshot, 'utf8'), '');
  assert.deepEqual(result.selection.snapshot, { exportedAt: NOW_ISO, source: 'bulk', count: 0 });
  assert.deepEqual(result.selection.recipients, []);
  assert.equal(result.selection.stats.totalCents, 0);
  assert.ok(hasLine(t.log, '名单是空的', 'WARN'));
});

// ---------------------------------------------------------------------------
// select: text that node:readline would break apart (U+2028 / U+2029)
// ---------------------------------------------------------------------------

const LS = ' '; // LINE SEPARATOR: JSON.stringify leaves it raw inside strings
const PS = ' '; // PARAGRAPH SEPARATOR: same
const PASTED = [
  makeCustomer({
    n: 31,
    first: `Ann${LS}Marie`,
    last: `O${PS}Neil`,
    address: { address1: `12 Main St${LS}Apt 3`, address2: `Rear${PS}Door`, city: 'Austin', provinceCode: 'TX', zip: '78731', countryCodeV2: 'US' },
    tags: [`VIP${LS}Gold`, `Note${PS}x`],
    lastOrder: makeOrder({ n: 3101, createdAt: '2026-03-01T12:00:00Z', total: '150.00' }),
  }),
  makeCustomer({ n: 32, first: `李${LS}😀`, last: `Chen${PS}Wu`, createdAt: '2025-02-01T00:00:00Z' }), // never ordered
  makeCustomer({ n: 33, first: `Bob${PS}`, last: 'Ray', tags: ['WHS'], lastOrder: makeOrder({ n: 3301, createdAt: '2026-02-01T12:00:00Z', total: '80.00' }) }), // excluded tag
  makeCustomer({ n: 34, lastOrder: makeOrder({ n: 3401, createdAt: '2026-01-01T12:00:00Z', total: '60.00' }) }),
];
const pastedNode = (n) => PASTED.find((c) => c.id === gid('Customer', n));

/** The separators survive into the parsed snapshot and the list, character for character. */
async function assertSeparatorsKept(t, result) {
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  const parsed = new Map((await readSnapshot(t.paths)).map((c) => [c.id, c]));
  assert.equal(parsed.size, PASTED.length);
  for (const node of PASTED) {
    const c = parsed.get(node.id);
    assert.equal(c.firstName, node.firstName);
    assert.equal(c.lastName, node.lastName);
    assert.equal(c.displayName, node.displayName);
    assert.deepEqual(c.tags, node.tags);
    assert.deepEqual(c.address, node.defaultAddress);
  }
  const recipients = new Map(result.selection.recipients.map((r) => [r.customerId, r]));
  assert.deepEqual([...recipients.keys()], [gid('Customer', 31), gid('Customer', 34), gid('Customer', 32)]);
  assert.equal(recipients.get(gid('Customer', 31)).name, pastedNode(31).displayName);
  assert.equal(recipients.get(gid('Customer', 31)).firstName, `Ann${LS}Marie`);
  assert.equal(recipients.get(gid('Customer', 32)).name, pastedNode(32).displayName);
  assert.equal(result.selection.notSelected.find((r) => r.customerId === gid('Customer', 33)).name, pastedNode(33).displayName);
}

test('U+2028 / U+2029 in names, addresses and tags survive the bulk export', async () => {
  const t = setup({ customers: PASTED, orders: [], orderHistory: {} });
  const result = await t.run();
  assert.equal(result.selection?.snapshot.source, 'bulk', t.log.lines.join('\n'));
  assert.equal(t.fake.opsNamed('CustomersPage').length, 0, 'the export was not rejected');
  assert.deepEqual(linesOf(t.log, 'WARN'), []);
  assert.ok(fs.readFileSync(t.paths.snapshot, 'utf8').includes(LS), 'the snapshot holds the raw character');
  await assertSeparatorsKept(t, result);
});

test('U+2028 / U+2029 survive the paginated fallback and give the same list as the bulk path', async () => {
  const viaBulk = (await setup({ customers: PASTED, orders: [], orderHistory: {} }).run()).selection;
  const t = setup({ customers: PASTED, orders: [], orderHistory: {}, bulk: { mode: 'fail' } });
  const result = await t.run();
  assert.equal(result.selection?.snapshot.source, 'paginated', t.log.lines.join('\n'));
  assert.ok(fs.readFileSync(t.paths.snapshot, 'utf8').includes(PS), 'JSON.stringify wrote the raw character');
  await assertSeparatorsKept(t, result);
  assert.deepEqual(withoutSnapshot(result.selection), withoutSnapshot(viaBulk));
});

test('a bulk download arriving in tiny chunks (split lines and split UTF-8 characters) is read intact', async () => {
  const t = setup({ customers: PASTED, orders: [], orderHistory: {} });
  wrapFetch(async (url, init, inner) => {
    const res = await inner(url, init);
    if (url !== BULK_URL) return res;
    const bytes = new Uint8Array(await res.arrayBuffer());
    let i = 0;
    const body = new ReadableStream({
      pull(controller) {
        if (i >= bytes.length) return controller.close();
        controller.enqueue(bytes.slice(i, i + 5)); // 5 bytes: cuts 3- and 4-byte characters apart
        i += 5;
        return undefined;
      },
    });
    return new Response(body, { status: 200 });
  });
  const result = await t.run();
  assert.equal(result.selection?.snapshot.source, 'bulk', t.log.lines.join('\n'));
  await assertSeparatorsKept(t, result);
});

test('readSnapshot: raw U+2028 / U+2029, CRLF, blank lines, a BOM and lines across read chunks', async () => {
  const cfg = testConfig();
  cleanups.push(() => cfg.cleanup());
  const paths = campaignPaths(cfg.config);
  fs.mkdirSync(paths.dir, { recursive: true });
  fs.writeFileSync(paths.snapshot, `﻿${JSON.stringify(PASTED[0])}\r\n\r\n${JSON.stringify(PASTED[1])}\n${JSON.stringify(PASTED[2])}`);
  const three = await readSnapshot(paths);
  assert.deepEqual(three.map((c) => [c.id, c.firstName, c.lastName]), PASTED.slice(0, 3).map((n) => [n.id, n.firstName, n.lastName]));

  // Far more than one 64 KiB read chunk of multi-byte text: every line and character intact.
  const many = Array.from({ length: 600 }, (_, i) => makeCustomer({ n: 1_000 + i, first: `李${LS}${i}😀`, last: `${PS}é${'ü'.repeat(i % 50)}` }));
  fs.writeFileSync(paths.snapshot, `${many.map((c) => JSON.stringify(c)).join('\n')}\n`);
  assert.ok(fs.statSync(paths.snapshot).size > 3 * 65_536);
  const read = await readSnapshot(paths);
  assert.deepEqual(read.map((c) => [c.id, c.firstName, c.lastName]), many.map((n) => [n.id, n.firstName, n.lastName]));

  // A damaged line still names the right line number ("\n" lines only).
  fs.writeFileSync(paths.snapshot, `${JSON.stringify(PASTED[0])}\n${JSON.stringify(PASTED[1])}\n{"id": "gid://shopify/Customer/9", ${LS}\n`);
  await assert.rejects(readSnapshot(paths), /customers\.jsonl 第 3 行不是有效的 JSON/);
});

test('textLines splits on "\\n" only, across chunk boundaries, and wants string chunks', async () => {
  const collect = async (chunks) => {
    const out = [];
    for await (const line of textLines(chunks)) out.push(line);
    return out;
  };
  assert.deepEqual(await collect(['a', `b${LS}c\nd`, 'e\r\n', `\n${PS}`, 'f']), [`ab${LS}c`, 'de\r', '', `${PS}f`]);
  assert.deepEqual(await collect(['x\n']), ['x']);
  assert.deepEqual(await collect([]), []);
  await assert.rejects(collect([Buffer.from('a\n')]), /needs a stream that yields strings/);
});

// ---------------------------------------------------------------------------
// select: snapshot reuse
// ---------------------------------------------------------------------------

test('a second run within 24 hours reuses the snapshot; --refresh and a stale snapshot export again', async () => {
  const t = setup();
  const first = await t.run();
  assert.equal(first.exitCode, 0, t.log.lines.join('\n'));

  const laterMs = NOW_MS + 3 * HOUR;
  const second = await t.run({ now: at(laterMs) });
  assert.equal(second.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(t.fake.opsNamed('RunCustomerExport').length, 1);
  assert.equal(t.fake.state.calls.bulkDownloads, 1);
  assert.equal(t.fake.opsNamed('OrdersSince').length, 1);
  assert.equal(t.fake.opsNamed('CustomerOrders').length, 1); // follow-up answer cached with the snapshot
  assert.equal(second.selection.createdAt, new Date(laterMs).toISOString());
  assert.deepEqual(second.selection.snapshot, first.selection.snapshot);
  assert.equal(second.selection.params.cutoffIso, CUTOFF_ISO); // the snapshot's own cutoff
  assert.deepEqual(second.selection.recipients, first.selection.recipients);
  assert.ok(hasLine(t.log, '复用', 'INFO'));
  assert.equal(readJournal(t.paths.journal).filter((e) => e.op === 'run.end').at(-1).summary.snapshotReused, true);

  const refreshedMs = NOW_MS + 4 * HOUR;
  const refreshed = await t.run({ refresh: true, now: at(refreshedMs) });
  assert.equal(refreshed.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(t.fake.opsNamed('RunCustomerExport').length, 2);
  assert.equal(t.fake.opsNamed('OrdersSince').length, 2);
  assert.equal(t.fake.opsNamed('CustomerOrders').length, 2); // the cache belongs to the old snapshot
  assert.equal(refreshed.selection.snapshot.exportedAt, new Date(refreshedMs).toISOString());
  assert.ok(hasLine(t.log, '按 --refresh 重新导出客户数据', 'INFO'));

  const staleMs = refreshedMs + SNAPSHOT_MAX_AGE_MS;
  const stale = await t.run({ now: at(staleMs) });
  assert.equal(stale.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(t.fake.opsNamed('RunCustomerExport').length, 3);
  assert.equal(stale.selection.snapshot.exportedAt, new Date(staleMs).toISOString());
  assert.equal(stale.selection.params.cutoffIso, '2026-07-02T00:00:00-07:00');
  assert.ok(hasLine(t.log, '已超过 24 小时', 'INFO'));
  assert.deepEqual(brief(stale.selection), EXPECTED_RECIPIENTS);
  assert.equal(t.reports.length, 4);
});

test('changing INACTIVE_MONTHS forces a new export with the new cutoff', async () => {
  const t = setup();
  assert.equal((await t.run()).exitCode, 0);
  t.config.inactiveMonths = 4;
  const result = await t.run({ now: at(NOW_MS + HOUR) });
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(t.fake.opsNamed('RunCustomerExport').length, 2);
  assert.ok(hasLine(t.log, 'INACTIVE_MONTHS 从 3 改成了 4', 'INFO'));
  assert.equal(result.selection.params.cutoffIso, '2026-06-01T00:00:00-07:00');
  assert.equal(readJson(t.paths.snapshotMeta).inactiveMonths, 4);
  // The June order now falls inside the window: that account ordered recently, and so did its address.
  const codes = new Map(result.selection.notSelected.map((r) => [r.customerId, r.primaryCode]));
  assert.equal(codes.get(DUP_KEPT.id), 'recent-order');
  assert.equal(codes.get(DUP_DROPPED.id), 'active-address');
  assert.ok(!result.selection.recipients.some((r) => r.customerId === DUP_KEPT.id || r.customerId === DUP_DROPPED.id));
});

test('cached follow-up lookups are reused only for the GIFT_PERCENT they were fetched with', async () => {
  const t = setup();
  assert.equal((await t.run()).exitCode, 0);
  assert.equal(t.fake.opsNamed('CustomerOrders').length, 1);

  t.config.giftPercent = 20;
  const result = await t.run({ now: at(NOW_MS + HOUR) });
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(t.fake.opsNamed('RunCustomerExport').length, 1); // snapshot reused...
  assert.equal(t.fake.opsNamed('CustomerOrders').length, 2); // ...but the lookup is redone
  assert.equal(readJson(t.paths.orderHistory)[ZERO_LAST.id].percent, 20);
  const zero = result.selection.recipients.find((r) => r.customerId === ZERO_LAST.id);
  assert.equal(zero.rawCents, 460); // 20% of $23.00
});

test('a candidate deleted before the follow-up lookup is left out (rule 12); the cached answer keeps them out', async () => {
  const t = setup();
  wrapFetch((url, init, inner) => {
    if (isGraphql(url) && operationOf(init) === 'CustomerOrders' && JSON.parse(init.body).variables.id === ZERO_LAST.id) {
      t.fake.state.customers = t.fake.state.customers.filter((c) => c.id !== ZERO_LAST.id); // deleted after the export
    }
    return inner(url, init);
  });
  const result = await t.run();
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  const s = result.selection;
  assert.ok(!s.recipients.some((r) => r.customerId === ZERO_LAST.id), 'not on the list');
  // (350 + 150) / 2 = $250.00 average → 10% = $25.00 → $19.77 for the never-ordered customer.
  assert.deepEqual(brief(s), [
    { customerId: DUP_KEPT.id, kind: 'ordered', amountCents: 1977, rawCents: 3500, tier: 2 },
    { customerId: ORDERED.id, kind: 'ordered', amountCents: 1533, rawCents: 1500, tier: 1 },
    { customerId: NEVER.id, kind: 'never', amountCents: 1977, rawCents: 2500, tier: 2 },
  ]);
  assert.equal(s.stats.recipients, 3);
  assert.equal(s.stats.totalCents, 1977 + 1533 + 1977);
  const row = s.notSelected.find((r) => r.customerId === ZERO_LAST.id);
  assert.equal(row.primaryCode, 'customer-deleted');
  assert.equal(row.primaryText, '12. 补查订单时客户已被删除');
  assert.deepEqual(s.notSelected.map((r) => r.primaryRule), [5, 7, 8, 10, 12]);
  assert.equal(s.funnel.byRule.find((r) => r.code === 'customer-deleted').count, 1);
  assert.equal(s.funnel.total, s.funnel.unlisted + s.funnel.listed + s.funnel.recipients);
  assert.deepEqual(readJson(t.paths.orderHistory), { [ZERO_LAST.id]: { orders: [], percent: 10, deleted: true } });
  assert.ok(hasLine(t.log, '其中 1 位客户在 Shopify 上已被删除，不进发放名单', 'WARN'));
  assert.ok(hasLine(t.log, '12. 补查订单时客户已被删除：1', 'INFO'));

  // Within 24 hours the snapshot and the cached follow-up are reused: still left out, no new lookup.
  const again = await t.run({ now: at(NOW_MS + HOUR) });
  assert.equal(again.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(t.fake.opsNamed('CustomerOrders').length, 1);
  assert.deepEqual(brief(again.selection), brief(s));
  assert.equal(again.selection.notSelected.find((r) => r.customerId === ZERO_LAST.id).primaryCode, 'customer-deleted');
});

test('a damaged snapshot file is reported instead of silently producing a short list', async () => {
  const t = setup();
  assert.equal((await t.run()).exitCode, 0);
  const before = fs.readFileSync(t.paths.selection, 'utf8');
  const lines = fs.readFileSync(t.paths.snapshot, 'utf8').trim().split('\n');
  fs.writeFileSync(t.paths.snapshot, `${lines.slice(1).join('\n')}\n`); // one customer lost
  const result = await t.run({ now: at(NOW_MS + HOUR) });
  assert.equal(result.exitCode, 1);
  assert.equal(result.selection, null);
  assert.ok(hasLine(t.log, '文件可能被改动过，请运行 select --refresh 重新导出', 'ERROR'));
  assert.equal(fs.readFileSync(t.paths.selection, 'utf8'), before);
});

// ---------------------------------------------------------------------------
// select: refusals
// ---------------------------------------------------------------------------

test('refuses to run once issuing has started; selection.json stays unchanged', async () => {
  const t = setup();
  assert.equal((await t.run()).exitCode, 0);
  const before = fs.readFileSync(t.paths.selection, 'utf8');
  appendJournal(t.paths.journal, { op: 'create.start', cid: ORDERED.id, amountCents: 1533, batch: 1, run: 'r-issue' });
  const tokenCalls = t.fake.state.calls.token.length;
  const opsBefore = t.fake.state.calls.ops.length;
  const journalBefore = readJournal(t.paths.journal).length;

  const result = await t.run({ refresh: true, now: at(NOW_MS + HOUR) });
  assert.equal(result.exitCode, 1);
  assert.equal(result.selection, null);
  // The way to a new campaign never creates cards from the list being abandoned: issue --repair-only, not issue --limit N.
  assert.deepEqual(linesOf(t.log, 'ERROR'), [
    'ERROR 已经开始发放，不能重新运行 select。要另起一个新活动：先运行 DRY_RUN=false node index.js issue --repair-only（只补记和补打 tag，不建新卡），再运行 verify，确认没有“已建卡未打tag”的人，然后换一个 CAMPAIGN_ID 重新运行 select。',
  ]);
  assert.equal(ISSUING_STARTED_MESSAGE.includes('--limit'), false);
  assert.equal(fs.readFileSync(t.paths.selection, 'utf8'), before);
  assert.equal(t.fake.state.calls.token.length, tokenCalls); // never connected
  assert.equal(t.fake.state.calls.ops.length, opsBefore);
  assert.equal(readJournal(t.paths.journal).length, journalBefore); // a refusal is not a run
  assert.equal(fs.existsSync(t.paths.runLock), false);
});

test('issuing started but selection.json is gone: the refusal says to copy the folder back (issue --repair-only cannot run without it)', async () => {
  const t = setup();
  assert.equal((await t.run()).exitCode, 0);
  appendJournal(t.paths.journal, { op: 'create.start', cid: ORDERED.id, amountCents: 1533, batch: 1, run: 'r-issue' });
  fs.rmSync(t.paths.selection);
  const result = await t.run({ refresh: true, now: at(NOW_MS + HOUR) });
  assert.equal(result.exitCode, 1);
  assert.deepEqual(linesOf(t.log, 'ERROR'), [
    'ERROR 本机的发放日志显示已经开始发放，但名单文件 campaigns/2026-10/selection.json 不见了：名单不能重新生成。'
      + '请从原来运行 select 的电脑或备份，把整个 campaigns/2026-10/ 文件夹拷回这里，再运行 verify 核对。',
  ]);
  assert.equal(fs.existsSync(t.paths.selection), false, 'nothing re-created');
});

/** Every file in the campaign folder with its content (to prove a refusal wrote nothing). */
function folderContents(paths) {
  if (!fs.existsSync(paths.dir)) return {};
  return Object.fromEntries(fs.readdirSync(paths.dir).sort().map((f) => [f, fs.readFileSync(`${paths.dir}/${f}`, 'utf8')]));
}

// selection.json is still here (only the journal is lost or damaged).
const CARDS_EXIST_2026_10 = (n) => `ERROR Shopify 上已有 ${n} 张本活动（[campaign:2026-10]）的礼品卡，说明已经开始发放（本地日志可能丢失或被移动过）。`
  + '名单已冻结，不能重新运行 select。请运行 verify 核对；不要删除 campaigns/2026-10/ 里的文件。';
// No selection.json on this computer: verify and every other command need the original folder back.
const FOLDER_LOST = (n, id = '2026-10') => `ERROR 本机没有这个活动的名单（campaigns/${id}/selection.json），但 Shopify 上已有 ${n} 张本活动（[campaign:${id}]）的礼品卡：`
  + `已经开始发放，名单不能重新生成。请从原来运行 select 的电脑或备份，把整个 campaigns/${id}/ 文件夹拷回这里，再运行 verify 核对。`;

test('journal lost after issuing started: the campaign cards in Shopify keep the list frozen; nothing is written', async () => {
  const t = setup();
  assert.equal((await t.run()).exitCode, 0);
  // 10/5: a card was created for ORDERED but its tag failed; then journal.jsonl was lost.
  t.fake.state.giftCards.push(giftCard({ n: 501, customer: ORDERED, createdAt: new Date(NOW_MS + 4 * DAY).toISOString(), amount: '15.33' }));
  fs.rmSync(t.paths.journal);
  const before = folderContents(t.paths);
  const opsBefore = t.fake.state.calls.ops.length;
  const linesBefore = t.log.lines.length;

  const result = await t.run({ refresh: true, now: at(NOW_MS + 4 * DAY + HOUR) });
  assert.equal(result.exitCode, 1);
  assert.equal(result.selection, null);
  assert.deepEqual(linesOf({ lines: t.log.lines.slice(linesBefore) }, 'ERROR'), [CARDS_EXIST_2026_10(1)]);
  assert.deepEqual(folderContents(t.paths), before, 'no journal, snapshot or list was written');
  assert.equal(fs.existsSync(t.paths.journal), false);
  // Only the card lookup went to Shopify: no export, no orders.
  assert.deepEqual(t.fake.state.calls.ops.slice(opsBefore).map((o) => o.op), ['FindCampaignCards']);
  assert.equal(t.fake.opsNamed('RunCustomerExport').length, 1);
  assert.equal(t.reports.length, 2, 'the Excel is still regenerated from the existing list');
  assert.equal(t.reports[1].runLockHeld, false);
  assert.equal(fs.existsSync(t.paths.runLock), false);

  // Many cards over several pages, and without --refresh: refused the same way; the count is formatted.
  for (let i = 0; i < 1_200; i += 1) t.fake.state.giftCards.push(giftCard({ n: 1_000 + i, customer: ORDERED, createdAt: new Date(NOW_MS + 4 * DAY).toISOString() }));
  const again = await t.run({ now: at(NOW_MS + 4 * DAY + 2 * HOUR) });
  assert.equal(again.exitCode, 1);
  assert.ok(hasLine(t.log, CARDS_EXIST_2026_10('1,201'), 'ERROR'));
  assert.equal(t.fake.opsNamed('RunCustomerExport').length, 1);
});

test('a first select in a fresh folder is refused when Shopify already has cards of this campaign', async () => {
  // campaigns/<id>/ was lost or the project was cloned elsewhere after issuing started.
  const t = setup({ giftCards: [giftCard({ n: 1, customer: ORDERED, createdAt: '2026-10-05T17:00:00Z' })] });
  const result = await t.run({ now: at(Date.parse('2026-10-05T21:00:00Z')) });
  assert.equal(result.exitCode, 1);
  assert.equal(result.selection, null);
  // Without selection.json, "请运行 verify" alone would send the operator in a circle (verify needs the list):
  // the message says to copy the original folder back.
  assert.deepEqual(linesOf(t.log, 'ERROR'), [FOLDER_LOST(1)]);
  assert.deepEqual(Object.keys(folderContents(t.paths)), [], 'nothing written');
  assert.equal(t.reports.length, 0, 'no list, so no Excel');
  for (const op of ['RunCustomerExport', 'CustomersPage', 'OrdersSince', 'CustomerOrders']) assert.equal(t.fake.opsNamed(op).length, 0, op);
});

test('the whole campaign folder lost after issuing: the refusal says to copy campaigns/<id>/ back, not to re-select', async () => {
  const t = setup();
  assert.equal((await t.run()).exitCode, 0);
  // 10/5: two cards issued; later the folder is gone (another computer, a fresh clone: campaigns/ is not in git).
  t.fake.state.giftCards.push(
    giftCard({ n: 801, customer: DUP_KEPT, createdAt: new Date(NOW_MS + 4 * DAY).toISOString(), amount: '19.77' }),
    giftCard({ n: 802, customer: ORDERED, createdAt: new Date(NOW_MS + 4 * DAY).toISOString(), amount: '15.33' }),
  );
  fs.rmSync(t.paths.dir, { recursive: true, force: true });
  const opsBefore = t.fake.state.calls.ops.length;
  const linesBefore = t.log.lines.length;

  const result = await t.run({ refresh: true, now: at(NOW_MS + 11 * DAY) });
  assert.equal(result.exitCode, 1);
  assert.equal(result.selection, null);
  const errors = linesOf({ lines: t.log.lines.slice(linesBefore) }, 'ERROR');
  assert.deepEqual(errors, [FOLDER_LOST(2)]);
  assert.match(errors[0], /拷回这里/);
  assert.doesNotMatch(errors[0], /请运行 verify 核对；/, 'not the "list is here" text');
  assert.deepEqual(Object.keys(folderContents(t.paths)), [], 'nothing written: no journal, no snapshot, no list');
  assert.deepEqual(t.fake.state.calls.ops.slice(opsBefore).map((o) => o.op), ['FindCampaignCards']);
  assert.equal(t.reports.length, 1, 'no list, so no Excel this time');

  // Only selection.json missing (other files still there): the same text; the count is formatted.
  fs.mkdirSync(t.paths.dir, { recursive: true });
  fs.writeFileSync(t.paths.snapshotMeta, '{}');
  for (let i = 0; i < 1_000; i += 1) t.fake.state.giftCards.push(giftCard({ n: 2_000 + i, customer: ORDERED, createdAt: new Date(NOW_MS + 4 * DAY).toISOString() }));
  const again = await t.run({ now: at(NOW_MS + 11 * DAY + HOUR) });
  assert.equal(again.exitCode, 1);
  assert.ok(hasLine(t.log, FOLDER_LOST('1,002'), 'ERROR'), t.log.lines.join('\n'));
  assert.deepEqual(Object.keys(folderContents(t.paths)), ['snapshot.json']);
});

test('cardsExistMessage: the "list is here" text by default, the folder-loss text without selection.json', () => {
  assert.equal(`ERROR ${cardsExistMessage('2026-10', 3)}`, CARDS_EXIST_2026_10(3));
  assert.equal(`ERROR ${cardsExistMessage('2026-10', 3, { listExists: true })}`, CARDS_EXIST_2026_10(3));
  assert.equal(`ERROR ${cardsExistMessage('2026-10', 16_117, { listExists: false })}`, FOLDER_LOST('16,117'));
  assert.equal(`ERROR ${cardsExistMessage('2026-10-test', 1, { listExists: false })}`, FOLDER_LOST(1, '2026-10-test'));
});

test('within a reused snapshot (no --refresh) the Shopify card check still runs first', async () => {
  const t = setup();
  assert.equal((await t.run()).exitCode, 0);
  t.fake.state.giftCards.push(giftCard({ n: 7, customer: DUP_KEPT, createdAt: new Date(NOW_MS + HOUR).toISOString() }));
  const linesBefore = t.log.lines.length;
  const result = await t.run({ now: at(NOW_MS + 2 * HOUR) });
  assert.equal(result.exitCode, 1);
  const newLines = t.log.lines.slice(linesBefore);
  assert.ok(!newLines.some((l) => l.includes('复用')), 'refused before the snapshot is reused');
  assert.equal(t.fake.opsNamed('FindCampaignCards').length, 2);
  assert.equal(readJournal(t.paths.journal).filter((e) => e.op === 'run.start').length, 1, 'the refusal is not a run');
});

test('only cards carrying this campaign\'s marker count; cards older than the look-back are not searched', async () => {
  const t = setup({
    giftCards: [
      giftCard({ n: 1, customer: ORDERED, createdAt: '2026-10-01T10:00:00Z', campaignId: '2026-10-test' }), // the test campaign
      giftCard({ n: 2, customer: ORDERED, createdAt: '2026-09-30T18:00:00Z', campaignId: null }), // the 9/30 run: no marker
      giftCard({ n: 3, customer: NEVER, createdAt: '2026-09-15T18:00:00Z', campaignId: '2026-10b' }),
      giftCard({ n: 4, customer: NEVER, createdAt: new Date(NOW_MS - CARD_LOOKBACK_MS - DAY).toISOString() }), // 401 days old
    ],
  });
  const result = await t.run();
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(CARD_LOOKBACK_MS, 400 * DAY);
  assert.deepEqual(t.fake.opsNamed('FindCampaignCards').map((o) => o.variables.query), [CARD_QUERY]);
  assert.deepEqual(brief(result.selection), EXPECTED_RECIPIENTS);
});

test('test campaign: its own cards freeze it; the live campaign\'s cards do not', async () => {
  const env = { CAMPAIGN_ID: '2026-10-test', SENT_TAG: 'OCT26RTPROMO-TEST', TEST_CUSTOMER_IDS: '1,2' };
  const t = setup({ env, giftCards: [giftCard({ n: 1, customer: ORDERED, createdAt: '2026-10-01T10:00:00Z', campaignId: '2026-10' })] });
  assert.equal((await t.run()).exitCode, 0, t.log.lines.join('\n'));
  const before = folderContents(t.paths);
  t.fake.state.giftCards.push(giftCard({ n: 2, customer: ORDERED, createdAt: new Date(NOW_MS + HOUR).toISOString(), campaignId: '2026-10-test' }));
  fs.rmSync(t.paths.journal);
  const result = await t.run({ now: at(NOW_MS + 2 * HOUR) });
  assert.equal(result.exitCode, 1);
  assert.ok(hasLine(t.log, 'Shopify 上已有 1 张本活动（[campaign:2026-10-test]）的礼品卡', 'ERROR'));
  assert.ok(hasLine(t.log, '不要删除 campaigns/2026-10-test/ 里的文件。', 'ERROR'));
  const { 'journal.jsonl': _journal, ...rest } = before;
  assert.deepEqual(folderContents(t.paths), rest);
  assert.equal(t.fake.opsNamed('RefreshCustomers').length, 1, 'the test customers are not read again');
});

test('the card check fails closed: a Shopify error stops select before anything is written; a transient one is retried', async () => {
  const t = setup({ failures: { FindCampaignCards: [{ kind: 'graphqlError', message: 'Access denied for giftCards field', code: 'ACCESS_DENIED' }] } });
  const result = await t.run();
  assert.equal(result.exitCode, 1);
  assert.ok(hasLine(t.log, 'select 失败（检查 Shopify 上已有的本活动礼品卡）：Shopify 请求出错', 'ERROR'));
  assert.ok(hasLine(t.log, '本次没有生成名单', 'ERROR'));
  assert.deepEqual(Object.keys(folderContents(t.paths)), []);
  assert.equal(t.fake.opsNamed('RunCustomerExport').length, 0);

  const flaky = setup({ failures: { FindCampaignCards: [{ kind: 'network' }, { kind: 'http', status: 503 }] } });
  const slept = [];
  const ok = await flaky.run({ sleep: async (ms) => { slept.push(ms); } });
  assert.equal(ok.exitCode, 0, flaky.log.lines.join('\n'));
  assert.equal(flaky.fake.opsNamed('FindCampaignCards').length, 3);
  assert.ok(slept.includes(2_000) && slept.includes(4_000), `sleeps: ${slept.join(',')}`);
  assert.equal(linesOf(flaky.log, 'WARN').filter((l) => l.includes('检查本活动的礼品卡：Shopify 暂时出错')).length, 2);
});

test('refuses while another command holds the run lock', async () => {
  const t = setup();
  const release = acquireRunLock(t.paths, 'issue');
  try {
    const result = await t.run();
    assert.equal(result.exitCode, 1);
    assert.equal(result.selection, null);
    assert.ok(hasLine(t.log, 'issue 正在运行', 'ERROR'));
    assert.equal(t.fake.state.calls.token.length, 0);
    assert.equal(fs.existsSync(t.paths.selection), false);
    assert.equal(t.reports.length, 0);
    assert.equal(fs.existsSync(t.paths.journal), false);
  } finally {
    release();
  }
  const retry = await t.run();
  assert.equal(retry.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(fs.existsSync(t.paths.runLock), false);
});

// ---------------------------------------------------------------------------
// select: test campaigns
// ---------------------------------------------------------------------------

test('test campaign: only TEST_CUSTOMER_IDS, fixed TEST_GIFT_AMOUNT, no bulk export', async () => {
  const noEmail = makeCustomer({ n: 10, email: null });
  const t = setup({
    customers: [...CUSTOMERS, noEmail],
    env: {
      CAMPAIGN_ID: '2026-10-test',
      SENT_TAG: 'OCT26RTPROMO-TEST',
      TEST_CUSTOMER_IDS: `1, ${gid('Customer', 1)}, 2, 10, 404`,
      TEST_GIFT_AMOUNT: '0.10',
    },
  });
  const result = await t.run();
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));

  const saved = readJson(t.paths.selection);
  assert.deepEqual(saved, result.selection);
  assert.equal(saved.mode, 'test');
  assert.equal(saved.campaignId, '2026-10-test');
  assert.equal(saved.createdAt, NOW_ISO);
  assert.deepEqual(saved.snapshot, { exportedAt: NOW_ISO, source: 'nodes', count: 3 });
  assert.equal(saved.params.testGiftAmountCents, 10);
  // Rules are not applied: the unsubscribed customer is a recipient too; one card per customer.
  assert.deepEqual(saved.recipients.map((r) => [r.seq, r.customerId, r.kind, r.amountCents, r.tier]), [
    [1, ORDERED.id, 'test', 10, null],
    [2, NOT_SUBSCRIBED.id, 'test', 10, null],
  ]);
  assert.equal(saved.recipients[0].formula, '测试固定金额 $0.10');
  assert.equal(saved.funnel.byRule.find((r) => r.code === 'no-email').count, 1);

  assert.deepEqual(t.fake.opsNamed('RefreshCustomers').map((o) => o.variables.ids), [[gid('Customer', 1), gid('Customer', 2), gid('Customer', 10), gid('Customer', 404)]]);
  for (const op of ['RunCustomerExport', 'BulkStatus', 'CustomersPage', 'OrdersSince', 'CustomerOrders']) {
    assert.equal(t.fake.opsNamed(op).length, 0, `${op} must not be called in a test campaign`);
  }
  assert.equal(t.fake.state.calls.bulkDownloads, 0);
  assert.equal(fs.existsSync(t.paths.snapshot), false);
  assert.equal(fs.existsSync(t.paths.snapshotMeta), false);
  assert.ok(hasLine(t.log, `测试客户 ${gid('Customer', 404)} 在 Shopify 上找不到，已跳过`, 'WARN'));
  assert.ok(hasLine(t.log, 'TEST_CUSTOMER_IDS 里有重复的客户', 'WARN'));
  assert.ok(hasLine(t.log, '入选：2 人，每人固定 $0.10'));
  assert.ok(hasLine(t.log, '总面额：$0.20'));
  assert.deepEqual(readJournal(t.paths.journal).map((e) => e.op), ['run.start', 'run.end']);
  assert.equal(readJournal(t.paths.journal)[1].summary.source, 'nodes');
  assert.equal(t.reports.length, 1);
});

// ---------------------------------------------------------------------------
// select: failures
// ---------------------------------------------------------------------------

test('a search warning on OrdersSince stops select: exit 1, no selection, snapshot not committed', async () => {
  const t = setup({ searchWarnings: { OrdersSince: true } });
  const result = await t.run();
  assert.equal(result.exitCode, 1);
  assert.equal(result.selection, null);
  assert.equal(fs.existsSync(t.paths.selection), false);
  assert.equal(fs.existsSync(t.paths.snapshotMeta), false); // the next run exports again
  assert.equal(t.fake.opsNamed('OrdersSince').length, 1); // never retried
  assert.ok(hasLine(t.log, 'select 已停止（读取近期订单）：Shopify 忽略了查询里的搜索条件', 'ERROR'));
  assert.ok(hasLine(t.log, '本次没有生成名单', 'ERROR'));
  assert.equal(t.reports.length, 0); // nothing to report on yet
  const journal = readJournal(t.paths.journal);
  assert.deepEqual(journal.map((e) => e.op), ['run.start', 'run.end']);
  assert.equal(journal[1].exitCode, 1);
  assert.match(journal[1].summary.error, /search filter/);
  assert.equal(fs.existsSync(t.paths.runLock), false);
});

test('a Shopify error during the follow-up keeps the previous selection and still refreshes the Excel', async () => {
  const t = setup();
  assert.equal((await t.run()).exitCode, 0);
  const before = fs.readFileSync(t.paths.selection, 'utf8');
  t.fake.state.failures.CustomerOrders = [{ kind: 'graphqlError', message: 'Internal boom', code: 'ACCESS_DENIED' }];

  const result = await t.run({ refresh: true, now: at(NOW_MS + HOUR) });
  assert.equal(result.exitCode, 1);
  assert.equal(result.selection, null);
  assert.equal(fs.readFileSync(t.paths.selection, 'utf8'), before);
  assert.ok(hasLine(t.log, 'select 失败（补查订单）：Shopify 请求出错', 'ERROR'));
  assert.ok(hasLine(t.log, 'selection.json 仍是上一次的结果', 'ERROR'));
  assert.equal(readJournal(t.paths.journal).at(-1).exitCode, 1);
  assert.equal(t.reports.length, 2);
  assert.equal(fs.existsSync(t.paths.runLock), false);
});

test('transient Shopify errors on read queries are retried with a back-off', async () => {
  const t = setup({ failures: { OrdersSince: [{ kind: 'network' }, { kind: 'http', status: 503 }] } });
  const slept = [];
  const result = await t.run({ sleep: async (ms) => { slept.push(ms); } });
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(t.fake.opsNamed('OrdersSince').length, 3);
  assert.ok(slept.includes(2_000) && slept.includes(4_000), `sleeps: ${slept.join(',')}`);
  assert.equal(linesOf(t.log, 'WARN').filter((l) => l.includes('读取近期订单') && l.includes('重试')).length, 2);
  assert.deepEqual(brief(result.selection), EXPECTED_RECIPIENTS);
});

test('a writeReport failure is only a warning', async () => {
  const t = setup();
  const result = await t.run({ writeReport: async () => { throw new Error('disk full'); } });
  assert.equal(result.exitCode, 0);
  assert.ok(hasLine(t.log, 'Excel 没有生成：disk full', 'WARN'));
  assert.ok(fs.existsSync(t.paths.selection));
});

test('uses the shop time zone when TIMEZONE is empty', async () => {
  const t = setup({ env: { TIMEZONE: '' } });
  const result = await t.run();
  assert.equal(result.exitCode, 0, t.log.lines.join('\n'));
  assert.equal(t.fake.opsNamed('ShopInfo').length, 1);
  assert.equal(result.selection.params.timezone, TZ);
  assert.equal(result.selection.params.cutoffIso, CUTOFF_ISO);
  assert.equal(readJson(t.paths.snapshotMeta).timezone, TZ);
});

// ---------------------------------------------------------------------------
// export.js fetchers
// ---------------------------------------------------------------------------

test('exportCustomers: paginated fallback walks every page and replaces the old snapshot atomically', async () => {
  const cfg = testConfig();
  cleanups.push(() => cfg.cleanup());
  const many = Array.from({ length: 600 }, (_, i) => makeCustomer({ n: i + 1 }));
  const fake = installFakeShopify({ customers: many, bulk: { mode: 'fail' } });
  cleanups.push(() => fake.restore());
  const log = memoryLog();
  directClient(log);
  const paths = campaignPaths(cfg.config);
  fs.mkdirSync(paths.dir, { recursive: true });
  fs.writeFileSync(paths.snapshot, '{"id":"gid://shopify/Customer/999999"}\n'); // an older snapshot

  const result = await exportCustomers({ paths, log, sleep: noSleep, bulkPollMs: 1_000 });
  assert.deepEqual(result, { count: 600, source: 'paginated', rootObjectCount: null });
  assert.deepEqual(fake.opsNamed('CustomersPage').map((o) => o.variables.after ?? null), [null, 'cur-250', 'cur-500']);
  assert.equal(fs.existsSync(`${paths.snapshot}.tmp`), false);

  const customers = await readSnapshot(paths);
  assert.equal(customers.length, 600);
  assert.deepEqual(customers.map((c) => c.id), many.map((c) => c.id));
  assert.equal(customers[0].email, 'c1@example.org');
  assert.equal(customers[0].addressKey, '1 main st|70001|us');
  assert.ok(hasLine(log, '全店客户分页导出完成：600 个客户（3 页）', 'INFO'));
});

test('fetchActiveOrders: pages through every order since the cutoff', async () => {
  const buyer = makeCustomer({ n: 1 });
  const orders = Array.from({ length: 300 }, (_, i) => makeActiveOrder({ n: i + 1, customer: buyer, createdAt: new Date(CUTOFF_MS + (i + 1) * HOUR).toISOString() }));
  orders.push(makeActiveOrder({ n: 999, customer: buyer, createdAt: new Date(CUTOFF_MS - 1).toISOString() })); // just before the cutoff
  const fake = installFakeShopify({ orders });
  cleanups.push(() => fake.restore());
  const log = memoryLog();
  directClient(log);

  const nodes = await fetchActiveOrders({ cutoffIso: CUTOFF_ISO, log, sleep: noSleep });
  assert.equal(nodes.length, 300);
  assert.ok(!nodes.some((o) => o.id === gid('Order', 999)));
  const calls = fake.opsNamed('OrdersSince');
  assert.equal(calls.length, 2);
  assert.ok(calls.every((c) => c.variables.query === `created_at:>='${CUTOFF_ISO}'`));
  assert.equal(calls[1].variables.after, 'cur-250');
});

test('fetchOrderHistory: pages until a paid valid order, caps at 20 pages, handles deleted customers', async () => {
  const a = makeCustomer({ n: 1 });
  const b = makeCustomer({ n: 2 });
  const d = makeCustomer({ n: 4 });
  const deletedId = gid('Customer', 3);
  // a: 55 unusable orders (cancelled, test, $0, $0.04 which is 0 cents at 10%), then a paid $12 order.
  const unusable = (n) => {
    const kind = n % 4;
    if (kind === 0) return makeOrder({ n, total: '30.00', cancelled: true });
    if (kind === 1) return makeOrder({ n, total: '30.00', test: true });
    if (kind === 2) return makeOrder({ n, total: '0.00' });
    return makeOrder({ n, total: '0.04' });
  };
  const aOrders = [...Array.from({ length: 55 }, (_, i) => unusable(i + 1)), makeOrder({ n: 56, total: '12.00' }), makeOrder({ n: 57, total: '99.00' })];
  // b: 1,100 orders of $0: the lookup gives up after 20 pages (1,000 orders).
  const bOrders = Array.from({ length: 1_100 }, (_, i) => makeOrder({ n: 1_000 + i, total: '0.00' }));
  const fake = installFakeShopify({ customers: [a, b, d], orderHistory: { [a.id]: aOrders, [b.id]: bOrders } });
  cleanups.push(() => fake.restore());
  const log = memoryLog();
  directClient(log);

  const history = await fetchOrderHistory([a.id, b.id, deletedId, d.id, a.id], { percent: 10, log, sleep: noSleep });
  assert.deepEqual([...history.keys()], [a.id, b.id, deletedId, d.id]);
  assert.equal(history.get(a.id).orders.length, 56);
  assert.equal(history.get(a.id).orders.at(-1).id, gid('Order', 56));
  assert.equal(history.get(b.id).orders.length, 1_000);
  assert.deepEqual(history.get(deletedId), { orders: [], deleted: true });
  assert.deepEqual(history.get(d.id), { orders: [] }); // exists, has no orders: not deleted
  assert.deepEqual(history.get(a.id).deleted, undefined);

  const callsFor = (id) => fake.opsNamed('CustomerOrders').filter((o) => o.variables.id === id).length;
  assert.equal(callsFor(a.id), 2);
  assert.equal(callsFor(b.id), 20);
  assert.equal(callsFor(deletedId), 1);
  assert.equal(callsFor(d.id), 1);
  assert.ok(hasLine(log, '其中 1 位客户在 Shopify 上已被删除，不进发放名单', 'WARN'));
  assert.ok(hasLine(log, '1 位客户查了 20 页', 'WARN'));
});

test('readSnapshot: a damaged line is a clear error', async () => {
  const cfg = testConfig();
  cleanups.push(() => cfg.cleanup());
  const paths = campaignPaths(cfg.config);
  fs.mkdirSync(paths.dir, { recursive: true });
  fs.writeFileSync(paths.snapshot, `${JSON.stringify(ORDERED)}\n{"id": "gid://shopify/Customer/2", "firstName": \n`);
  await assert.rejects(readSnapshot(paths), /customers\.jsonl 第 2 行不是有效的 JSON/);
  fs.rmSync(paths.snapshot);
  await assert.rejects(readSnapshot(paths), /找不到客户快照/);
});
