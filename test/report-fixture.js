// Shared fixture of the Excel report tests (golden comparison, Chinese/English parity).
//
// Three scenarios exercise every sheet and every text the workbook can show:
//   live   a real campaign: every rule (1-12), a same-address group and a bulk group of 10,
//          a journal with every op type and every stored note / error / detail form,
//          tags.json, verify.json with every issue type, usage.json with refunds, unmatched
//          payments and top products, run.end summaries of every command (failures, dry
//          runs, exotic shapes), unfinished runs while `remind` is running, and .env dates
//          that changed after select
//   test   a test campaign (TEST_CUSTOMER_IDS) in another time zone, with unset dates
//   empty  a live campaign without recipients (rule 11), a single tier, no journal,
//          minimal usage.json / verify.json, an invalid time zone and a foreign tags.json
//
// The inputs are frozen in test/fixtures/report-scenarios.json, so later changes to the
// select rules or the commands do not move the golden output. Customer data is ASCII:
// any CJK character in the English workbook comes from our own texts.
//
// Regenerate the frozen inputs (only when the scenarios themselves change):
//   UPDATE_REPORT_SCENARIOS=1 node --test test/excel-golden.test.js
// (then regenerate the golden output too: see test/excel-golden.test.js).

import fs from 'node:fs';
import path from 'node:path';
import { spawnSync, execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';

import { campaignPaths, acquireRunLock, writeJsonAtomic, ensureDir } from '../src/campaign.js';
import { buildTestSelection } from '../src/select/selection.js';
import { parseCustomer } from '../src/select/rules.js';
import { testConfig, selectionFixture, makeCustomer, makeOrder, makeActiveOrder, memoryLog, gid, NOW_ISO } from './helpers.js';

const HERE = path.dirname(fileURLToPath(import.meta.url));
export const FIXTURES_DIR = path.join(HERE, 'fixtures');
export const SCENARIOS_FILE = path.join(FIXTURES_DIR, 'report-scenarios.json');

export const TZ = 'America/Los_Angeles';
/** "Generated at" of every scenario workbook: 2026-10-12 10:00 store time. */
export const FIXED_NOW_ISO = '2026-10-12T17:00:00.000Z';
export const FIXED_NOW = () => new Date(FIXED_NOW_ISO);
export const FAST_LOCK = { pollMs: 5 };

const cust = (n) => gid('Customer', n);
const card = (n) => gid('GiftCard', n);
const order = (n) => gid('Order', n);
const p2 = (n) => String(n).padStart(2, '0');

// ---------------------------------------------------------------------------
// Building the inputs (real select rules; run once, then frozen as JSON)
// ---------------------------------------------------------------------------

const SAME_ADDRESS = { address1: '500 Elm Street', address2: '', city: 'Austin', provinceCode: 'TX', zip: '78701', countryCodeV2: 'US' };
const ACTIVE_ADDRESS = { address1: '77 Oak Ave', address2: '', city: 'Austin', provinceCode: 'TX', zip: '78702', countryCodeV2: 'US' };
const BULK_ADDRESS = { address1: '9 Warehouse Road', address2: 'Unit 4', city: 'Dallas', provinceCode: 'TX', zip: '75201', countryCodeV2: 'US' };
const BULK = Array.from({ length: 10 }, (_, i) => 31 + i); // c31..c40: one address, 10 accounts

function liveCustomers() {
  const c = (n, fields = {}) => makeCustomer({ n, ...fields });
  const o = (n, createdAt, total, extra = {}) => makeOrder({ n, createdAt, total, ...extra });
  const c12 = c(12, { address: ACTIVE_ADDRESS, lastOrder: o(120, '2026-08-20T12:00:00Z', '30.00') });
  const customers = [
    c(1, { createdAt: '2024-01-01T00:00:00Z', lastOrder: o(101, '2026-05-10T18:00:00Z', '150.00'), numberOfOrders: 4, amountSpent: '420.50' }),
    c(2, { lastOrder: o(102, '2026-04-01T18:00:00Z', '83.40', { source: 'checkout_next' }) }),
    c(3, { address: SAME_ADDRESS, lastOrder: o(103, '2026-03-15T18:00:00Z', '60.00'), numberOfOrders: 3 }),
    c(4, { address: SAME_ADDRESS, lastOrder: o(104, '2026-02-01T18:00:00Z', '40.00') }),
    c(5, { createdAt: '2025-06-01T12:00:00Z' }), // never ordered
    c(6, { lastOrder: o(60, '2026-02-20T18:00:00Z', '0.00') }), // last order $0
    c(7, { lastOrder: o(70, '2026-03-01T18:00:00Z', '90.00', { cancelled: true }) }), // last order cancelled
    c(8, { lastOrder: o(80, '2026-01-05T18:00:00Z', '45.00', { cancelled: true }) }), // only cancelled orders
    c(9, { lastOrder: o(90, '2026-08-15T18:00:00Z', '25.00') }), // rule 7
    c(10, { marketingState: 'NOT_SUBSCRIBED' }), // rule 3 (counted only)
    c(11, { address: ACTIVE_ADDRESS }), // rule 8
    c12, // rule 7, the active account of c11
    c(13, { tags: ['WHS'] }), // rule 5
    c(14, { lastOrder: o(140, '2026-01-20T18:00:00Z', '40.00', { source: 'shopify_draft_order' }) }),
    c(15, { lastOrder: o(150, '2026-06-01T18:00:00Z', '200.00', { source: 'pos' }) }),
    c(16, { lastOrder: o(160, '2026-02-10T18:00:00Z', '75.00', { test: true }) }), // last order a test order
    c(17, { email: '' }), // rule 1: no email
    c(18, { validFormat: false }), // rule 1: invalid format
    c(19, { email: 'c19@mail.codisto.com' }), // rule 2
    c(20, { marketingState: 'UNSUBSCRIBED' }),
    c(21, { marketingState: 'PENDING' }),
    c(22, { marketingState: 'INVALID' }),
    c(23, { marketingState: 'REDACTED' }),
    c(24, { marketingState: null }), // no marketing state → NONE
    c(25, { createdAt: '2026-09-28T12:00:00Z' }), // rule 4
    c(26, { tags: ['gift-card-sent-2026-10'] }), // rule 6
    c(27, { lastOrder: o(270, '2026-03-01T18:00:00Z', '50.00', { source: 'amazon' }) }), // rule 9
    c(28, { lastOrder: o(280, '2026-02-01T18:00:00Z', '65.00', { source: '205641' }) }), // rule 9 (Sellbrite)
    c(29, { tags: ['DISC'], lastOrder: o(290, '2026-08-01T18:00:00Z', '20.00') }), // rules 5 + 7
    c(30, { lastOrder: o(300, '2026-03-05T18:00:00Z', '55.00', { cancelled: true }) }), // rule 12 (deleted at the follow-up)
    ...BULK.map((n, i) => c(n, { address: BULK_ADDRESS, createdAt: `2025-0${1 + (i % 9)}-01T12:00:00Z` })),
    c(41, { lastOrder: o(410, '2026-04-20T18:00:00Z', '125.00', { source: 'tiktok' }) }),
    c(42, { lastOrder: o(420, '2026-05-01T18:00:00Z', '107.75', { source: '580111' }) }),
  ];
  return {
    customers,
    activeOrders: [makeActiveOrder({ n: 900, customer: c12, createdAt: '2026-08-20T12:00:00Z' })],
    history: {
      [cust(6)]: [o(60, '2026-02-20T18:00:00Z', '0.00'), o(61, '2026-01-10T18:00:00Z', '23.00')],
      [cust(7)]: [o(70, '2026-03-01T18:00:00Z', '90.00', { cancelled: true }), o(71, '2025-12-01T18:00:00Z', '350.00')],
      [cust(8)]: [o(80, '2026-01-05T18:00:00Z', '45.00', { cancelled: true })],
      [cust(16)]: [o(160, '2026-02-10T18:00:00Z', '75.00', { test: true }), o(161, '2025-11-01T18:00:00Z', '120.00')],
      [cust(30)]: { orders: [], deleted: true },
    },
  };
}

/** Journal entries ([time, entry]) of the live scenario: every op, every stored text form. */
function liveJournal(selection) {
  const amount = (n) => selection.recipients.find((r) => r.customerId === cust(n)).amountCents;
  const bulkKept = selection.duplicates.find((g) => g.size === 10).keptCustomerId;
  const T = (day, hh, mm, ss = 0) => `2026-10-${p2(day)}T${p2(hh)}:${p2(mm)}:${p2(ss)}.000Z`;
  const run = (id, t, fields) => [t, { op: 'run.start', run: id, ...fields }];
  const end = (id, t, summary, exitCode = 0) => [t, { op: 'run.end', run: id, summary, exitCode }];
  const S1 = '20261003190000-1';
  const S2 = '20261003200000-1';
  const I1 = '20261005160000-1';
  const I2 = '20261005170000-1';
  const I3 = '20261005180000-1';
  const I4 = '20261006160000-1';
  const R0 = '20261012150000-1';
  const R1 = '20261012160000-1';
  const R1D = '20261012170000-1';
  const RX = '20261012180000-1';
  const U1 = '20261013170000-1';
  const U2 = '20261013180000-1';
  const V1 = '20261014170000-1';
  const V2 = '20261014180000-1';
  const V3 = '20261014190000-1';
  const X1 = '20261015170000-1';
  const X2 = '20261015170100-1';
  const X3 = '20261015170200-1';
  const X4 = '20261015170300-1';
  const X5 = '20261015170400-1';
  const X6 = '20261015170500-1';
  const R2 = '20261016160000-1';
  const issueEnd = (fields) => ({
    batch: null, dryRun: false, attempted: 0, created: 0, tagged: 0, tagFixed: 0, reconciled: 0, skipped: {}, failed: 0, rejected: 0,
    unknown: 0, amountCents: 0, tagFailed: 0, reconciledNone: 0, stillUnknown: 0, seqFrom: null, seqTo: null, ...fields,
  });
  const remindEnd = (fields) => ({
    round: 1, dryRun: false, eligible: 0, planned: 0, attempted: 0, sent: 0, failed: 0, rejected: 0, unknown: 0, retriedUnknown: 0,
    retriedFailed: 0, skipped: {}, alreadySent: 0, alreadySentByTag: 0, previouslyFailed: 0, waitingUnknown: 0, notIssued: 0,
    roundTagged: 0, roundTagFailed: 0, roundTagFixed: 0, roundTagMissing: 0, stopped: null, stoppedByDate: null, ...fields,
  });
  const bulkAmount = selection.recipients.find((r) => r.customerId === bulkKept).amountCents;
  /** create.start at 16:mm:01 and create.ok at 16:mm:02 on 10/05 (batch 1). */
  const created = (n, giftCard, mm, fields = {}) => [
    [T(5, 16, mm, 1), { op: 'create.start', cid: cust(n), amountCents: amount(n), batch: 1, run: I1 }],
    [T(5, 16, mm, 2), { op: 'create.ok', cid: cust(n), giftCardId: card(giftCard), last4: `x${String(giftCard).slice(-3)}`, amountCents: amount(n), batch: 1, run: I1, ...fields }],
  ];
  return [
    // ---- select: one run, one failure
    run(S1, T(3, 19, 0), { command: 'select', dryRun: false, batch: null, limit: null, options: {} }),
    end(S1, T(3, 19, 3), { recipients: selection.recipients.length, totalCents: selection.stats.totalCents, source: 'bulk', snapshotReused: false }),
    run(S2, T(3, 20, 0), { command: 'select', dryRun: false, batch: null, limit: null, options: { refresh: true } }),
    end(S2, T(3, 20, 1), { error: 'GraphQL error: Throttled' }, 1),

    // ---- issue batch 1 (live): every outcome
    run(I1, T(5, 16, 0), { command: 'issue', dryRun: false, batch: 1, limit: 20, options: { retryFailed: false } }),
    ...created(1, 1001, 0),
    [T(5, 16, 0, 3), { op: 'tag.ok', cid: cust(1), run: I1 }],
    ...created(2, 1002, 1),
    [T(5, 16, 1, 3), { op: 'tag.fail', cid: cust(2), error: 'tagsAdd rejected: tags: Tag limit reached [INVALID]', run: I1 }],
    [T(5, 16, 2, 1), { op: 'create.start', cid: cust(3), amountCents: amount(3), batch: 1, run: I1 }],
    [T(5, 16, 2, 2), { op: 'create.fail', cid: cust(3), error: 'giftCardCreate rejected: input: Customer is invalid [INVALID]; input.expiresOn: must be in the future [GREATER_THAN]', run: I1 }],
    [T(5, 16, 3, 1), { op: 'create.start', cid: cust(5), amountCents: amount(5), batch: 1, run: I1 }],
    [T(5, 16, 3, 9), { op: 'create.unknown', cid: cust(5), error: 'Network error calling Shopify: fetch failed；查卡也失败了：GraphQL error: Internal error', run: I1 }],
    [T(5, 16, 4, 0), { op: 'skip', cid: cust(6), reason: 'ordered-since-snapshot', detail: '2026-10-04T18:30:00Z', batch: 1, run: I1 }],
    [T(5, 16, 5, 0), { op: 'create.start', cid: cust(7), amountCents: amount(7), batch: 1, run: I1 }], // no outcome: interrupted
    [T(5, 16, 5, 30), { op: 'skip', cid: cust(8), batch: 1, run: I1 }], // a skip without a reason
    ...created(14, 1014, 6, { amountCents: amount(14) + 456 }), // the card differs from the list amount
    [T(5, 16, 6, 3), { op: 'tag.ok', cid: cust(14), run: I1 }],
    [T(5, 16, 7, 1), { op: 'create.start', cid: cust(15), amountCents: amount(15), batch: 1, run: I1 }],
    [T(5, 16, 7, 9), { op: 'create.unknown', cid: cust(15), error: 'Network error calling Shopify: fetch failed', run: I1 }],
    [T(5, 16, 8, 1), { op: 'create.start', cid: cust(16), amountCents: amount(16), batch: 1, run: I1 }],
    [T(5, 16, 8, 2), { op: 'create.rejected', cid: cust(16), error: 'Throttled by Shopify 6 times in a row; giving up', run: I1 }],
    [T(5, 16, 9, 1), { op: 'create.start', cid: cust(42), amountCents: amount(42), batch: 1, run: I1 }],
    [T(5, 16, 9, 2), { op: 'create.unknown', cid: cust(42), error: 'Shopify HTTP 502', run: I1 }],
    [T(5, 16, 9, 30), { op: 'reconcile.found', cid: cust(42), giftCardId: card(1042), last4: 'x042', amountCents: amount(42), createdAt: T(5, 16, 9, 2), source: 'issue-inline', run: I1 }],
    [T(5, 16, 9, 31), { op: 'tag.ok', cid: cust(42), run: I1 }],
    [T(5, 16, 10, 0), { op: 'reconcile.found', cid: cust(41), giftCardId: card(1041), last4: 'x041', amountCents: amount(41), createdAt: T(4, 16, 0, 0), source: 'preflight', run: I1 }],
    [T(5, 16, 10, 1), { op: 'tag.ok', cid: cust(41), run: I1 }],
    // pre-flight skips of people outside the list: every reason and detail form
    [T(5, 16, 11, 0), { op: 'skip', cid: cust(900), reason: 'no-email', detail: '邮箱格式无效', batch: 1, run: I1 }],
    [T(5, 16, 11, 1), { op: 'skip', cid: cust(901), reason: 'relay-email', detail: 'mail.codisto.com', batch: 1, run: I1 }],
    [T(5, 16, 11, 2), { op: 'skip', cid: cust(902), reason: 'not-subscribed', detail: 'UNSUBSCRIBED', batch: 1, run: I1 }],
    [T(5, 16, 11, 3), { op: 'skip', cid: cust(903), reason: 'already-tagged', detail: 'gift-card-sent-2026-10', batch: 1, run: I1 }],
    [T(5, 16, 11, 4), { op: 'skip', cid: cust(904), reason: 'address-ordered-since-snapshot', detail: 'gid://shopify/Customer/12', batch: 1, run: I1 }],
    [T(5, 16, 11, 5), { op: 'skip', cid: cust(905), reason: 'customer-deleted', batch: 1, run: I1 }],
    [T(5, 16, 11, 6), { op: 'skip', cid: cust(906), reason: 'ordered-since-snapshot', detail: '#7001', batch: 1, run: I1 }],
    [T(5, 16, 11, 7), { op: 'skip', cid: cust(907), reason: 'ordered-since-snapshot', detail: '2026-12-01T20:15:00.000Z', batch: 1, run: I1 }],
    [T(5, 16, 11, 8), { op: 'skip', cid: cust(908), reason: 'not-subscribed', detail: 'NONE', batch: 1, run: I1 }],
    [T(5, 16, 11, 9), { op: 'skip', cid: cust(909), reason: 'future-reason', detail: 'gid://shopify/Product/5', batch: 1, run: I1 }],
    [T(5, 16, 11, 10), { op: 'skip', cid: cust(910), reason: 'address-ordered-since-snapshot', detail: 'gid://shopify/Order/7', batch: 1, run: I1 }],
    end(I1, T(5, 16, 12), issueEnd({ batch: 1, attempted: 9, created: 3, tagged: 4, reconciled: 2, skipped: { 'ordered-since-snapshot': 3, 'no-email': 1 }, failed: 1, rejected: 1, unknown: 3, amountCents: 6000, tagFailed: 1, seqFrom: 1, seqTo: 13 }), 1),

    // ---- issue batch 2 (dry run)
    run(I2, T(5, 17, 0), { command: 'issue', dryRun: true, batch: 2, limit: 500, options: { retryFailed: false } }),
    end(I2, T(5, 17, 0, 5), issueEnd({ batch: 2, dryRun: true, attempted: 1, tagFixed: 1, amountCents: 1977, stillUnknown: 2, seqFrom: 9, seqTo: 9 })),

    // ---- issue batch 3 (live, --repair-only): leftovers settled
    run(I3, T(5, 18, 0), { command: 'issue', dryRun: false, batch: 3, limit: null, options: { retryFailed: false, repairOnly: true } }),
    [T(5, 18, 0, 1), { op: 'reconcile.found', cid: cust(15), giftCardId: card(1015), last4: 'x015', amountCents: amount(15), createdAt: T(5, 16, 7, 5), source: 'issue-reconcile', run: I3 }],
    [T(5, 18, 0, 2), { op: 'tag.ok', cid: cust(15), run: I3 }],
    [T(5, 18, 0, 3), { op: 'reconcile.found', cid: bulkKept, giftCardId: card(1031), last4: 'x031', amountCents: bulkAmount, createdAt: T(5, 16, 30), source: 'verify', run: I3 }],
    [T(5, 18, 0, 4), { op: 'tag.ok', cid: bulkKept, note: 'already tagged in Shopify', run: I3 }],
    [T(5, 18, 0, 5), { op: 'reconcile.none', cid: cust(931), note: '超过 10 分钟仍查不到这张卡，确认未建成，可以重试', run: I3 }],
    [T(5, 18, 0, 6), { op: 'reconcile.none', cid: cust(932), note: '超过 10 分钟仍查不到这张卡，确认未建成；正式活动从 REMIND_1_DATE（2026-10-12）起不再建新卡', run: I3 }],
    [T(5, 18, 0, 7), { op: 'reconcile.none', cid: cust(933), note: '超过 90 秒仍查不到这张卡，确认未建成；本次是 --repair-only，不建新卡；之后正常运行 issue 时才会给他建卡', run: I3 }],
    [T(5, 18, 0, 8), { op: 'reconcile.found', cid: cust(934), giftCardId: card(1934), last4: 'x934', amountCents: 1533, source: 'manual', run: I3 }],
    [T(5, 18, 0, 9), { op: 'create.unknown', cid: cust(935), error: '上次运行在建卡途中中断，Shopify 上暂时查不到这张卡', run: I3 }],
    [T(5, 18, 0, 10), { op: 'tag.fail', cid: cust(936), error: 'tagsAdd returned no payload', run: I3 }],
    [T(5, 18, 0, 11), { op: 'create.fail', cid: cust(937), error: 'Shopify returned non-JSON (text/html)', run: I3 }],
    [T(5, 18, 0, 12), { op: 'create.rejected', cid: cust(938), error: 'HTTP 429 from Shopify 6 times in a row; giving up', run: I3 }],
    [T(5, 18, 0, 13), { op: 'create.unknown', cid: cust(939), error: 'Shopify response contained no data', run: I3 }],
    [T(5, 18, 0, 14), { op: 'create.fail', cid: cust(940), error: 'giftCardCreate returned no payload', run: I3 }],
    end(I3, T(5, 18, 1), issueEnd({ batch: 3, tagFixed: 2, reconciled: 2, reconciledNone: 3, newCardsRefused: 2, repairOnly: true, stoppedByDate: '2026-10-12' }), 2),

    // ---- issue batch 4: started, never ended (killed)
    run(I4, T(6, 16, 0), { command: 'issue', dryRun: false, batch: 4, limit: 5, options: { retryFailed: true } }),
    [T(6, 16, 0, 5), { op: 'reconcile.none', cid: cust(16), note: 'verify：开始建卡 10 分钟后仍没在 Shopify 找到这张卡，确认没有建成', run: I4 }],

    // ---- reminders round 1: an older run that never ended, the real run, a dry run
    run(R0, T(12, 15, 0), { command: 'remind', dryRun: false, batch: null, limit: null, options: { round: 1, retryUnknown: false, retryFailed: false } }),
    [T(12, 15, 0, 5), { op: 'remind.start', cid: cust(41), round: 1, giftCardId: card(1041), run: R0 }],
    [T(12, 15, 0, 6), { op: 'remind.unknown', cid: cust(41), round: 1, error: '上次运行在发送这封提醒时中断，不知道是否已发出', run: R0 }],
    run(R1, T(12, 16, 0), { command: 'remind', dryRun: false, batch: null, limit: null, options: { round: 1, retryUnknown: true, retryFailed: false } }),
    [T(12, 16, 4, 59), { op: 'remind.start', cid: cust(1), round: 1, giftCardId: card(1001), run: R1 }],
    [T(12, 16, 5, 0), { op: 'remind.ok', cid: cust(1), round: 1, run: R1 }],
    [T(12, 16, 5, 10), { op: 'remind.skip', cid: cust(2), round: 1, reason: 'used', detail: '余额 $5.33 / 面额 $15.33', run: R1 }],
    [T(12, 16, 5, 20), { op: 'remind.start', cid: cust(15), round: 1, giftCardId: card(1015), run: R1 }],
    [T(12, 16, 5, 50), { op: 'remind.unknown', cid: cust(15), round: 1, error: 'Shopify HTTP 503: Service Unavailable', run: R1 }],
    [T(12, 16, 6, 0), { op: 'remind.start', cid: cust(14), round: 1, giftCardId: card(1014), run: R1 }],
    [T(12, 16, 6, 1), { op: 'remind.fail', cid: cust(14), round: 1, error: 'giftCardSendNotificationToCustomer rejected: input: Customer has no email [INVALID]', run: R1 }],
    [T(12, 16, 6, 30), { op: 'remind.start', cid: cust(41), round: 1, giftCardId: card(1041), retry: true, run: R1 }],
    [T(12, 16, 6, 31), { op: 'remind.ok', cid: cust(41), round: 1, run: R1 }],
    [T(12, 16, 6, 32), { op: 'remind.tag.fail', cid: cust(41), round: 1, error: 'Network error calling Shopify: fetch failed', run: R1 }],
    [T(12, 16, 7, 0), { op: 'remind.found', cid: bulkKept, round: 1, source: 'tag', run: R1 }],
    [T(12, 16, 7, 10), { op: 'remind.skip', cid: cust(42), round: 1, reason: 'multiple-cards', detail: '2 张卡：x042、1043', run: R1 }],
    [T(12, 16, 7, 20), { op: 'remind.skip', cid: cust(950), round: 1, reason: 'no-card', detail: 'Shopify 上这位客户名下没有日志记录的卡 1950', run: R1 }],
    [T(12, 16, 7, 21), { op: 'remind.skip', cid: cust(951), round: 1, reason: 'card-disabled', run: R1 }],
    [T(12, 16, 7, 22), { op: 'remind.skip', cid: cust(952), round: 1, reason: 'customer-deleted', run: R1 }],
    [T(12, 16, 7, 23), { op: 'remind.skip', cid: cust(953), round: 1, reason: 'no-email', detail: 'c953@example.org', run: R1 }],
    [T(12, 16, 7, 24), { op: 'remind.skip', cid: cust(954), round: 1, reason: 'not-subscribed', detail: 'NOT_SUBSCRIBED', run: R1 }],
    [T(12, 16, 7, 25), { op: 'remind.skip', cid: cust(955), round: 1, reason: 'card-expired', detail: '到期日 2026-10-11', run: R1 }],
    [T(12, 16, 7, 26), { op: 'remind.skip', cid: cust(956), round: 1, reason: 'no-card', run: R1 }],
    [T(12, 16, 7, 27), { op: 'remind.start', cid: cust(957), round: 1, giftCardId: card(1957), run: R1 }],
    [T(12, 16, 7, 28), { op: 'remind.rejected', cid: cust(957), round: 1, error: 'HTTP 429 from Shopify 6 times in a row; giving up', run: R1 }],
    end(R1, T(12, 16, 8), remindEnd({
      eligible: 6, planned: 6, attempted: 5, sent: 2, failed: 1, unknown: 1, rejected: 1, retriedUnknown: 1, skipped: { used: 1, 'multiple-cards': 1, 'no-card': 2 },
      alreadySentByTag: 1, notIssued: 5, roundTagged: 1, roundTagFailed: 1, stopped: 'too-many-failures',
    }), 1),
    run(R1D, T(12, 17, 0), { command: 'remind', dryRun: true, batch: null, limit: 3, options: { round: 1 } }),
    end(R1D, T(12, 17, 0, 9), remindEnd({ dryRun: true, eligible: 2, planned: 2, alreadySent: 3, waitingUnknown: 1, previouslyFailed: 1, roundTagMissing: 1 })),
    // a reminder run without a round in its options; entries without a round
    run(RX, T(12, 18, 0), { command: 'remind', dryRun: true, batch: null, limit: null, options: {} }),
    [T(12, 18, 0, 1), { op: 'remind.found', cid: cust(960), source: 'tag', run: RX }],
    [T(12, 18, 0, 2), { op: 'remind.tag.fail', cid: cust(961), error: 'tagsAdd returned no payload', run: RX }],
    [T(12, 18, 0, 3), { op: 'remind.tag.ok', cid: cust(962), run: RX }],
    end(RX, T(12, 18, 1), remindEnd({ round: undefined, dryRun: true, stopped: 'aborted' }), 130),

    // ---- usage: one run, one failure, exotic summaries
    run(U1, T(13, 17, 0), { command: 'usage', dryRun: false, batch: null, limit: null, options: {} }),
    end(U1, T(13, 17, 1), { issuedCards: 9, usedCards: 2, usedCents: 2033, orders: 2, ordersTotalCents: 8500, giftCardCents: 1833, customerPaidCents: 6667, unmatched: 3 }),
    run(U2, T(13, 18, 0), { command: 'usage', dryRun: false, batch: null, limit: null, options: {} }),
    end(U2, T(13, 18, 1), { error: 'Network error calling Shopify: fetch failed' }, 1),

    // ---- verify: with text, failed, without text
    run(V1, T(14, 17, 0), { command: 'verify', dryRun: false, batch: null, limit: null, options: {} }),
    [T(14, 17, 0, 30), { op: 'reconcile.none', cid: cust(970), note: 'verify：开始建卡 10 分钟后仍没在 Shopify 找到这张卡，确认没有建成', run: V1 }],
    [T(14, 17, 0, 31), { op: 'reconcile.found', cid: cust(971), giftCardId: card(1971), last4: 'x971', amountCents: 1977, createdAt: T(5, 16, 40), source: 'verify', run: V1 }],
    end(V1, T(14, 17, 1), { cardCount: 9, taggedCount: 7, issueCount: 11, fixedCount: 2, counts: { 'tag-missing': 1, 'still-unknown': 1 }, text: '卡 9 张，发现 11 条，补记日志 2 条' }),
    run(V2, T(14, 18, 0), { command: 'verify', dryRun: false, batch: null, limit: null, options: {} }),
    end(V2, T(14, 18, 1), { error: 'Network error calling Shopify: fetch failed', text: '失败：Network error calling Shopify: fetch failed' }, 1),
    run(V3, T(14, 19, 0), { command: 'verify', dryRun: false, batch: null, limit: null, options: {} }),
    end(V3, T(14, 19, 1), { cardCount: 9, taggedCount: 7, issueCount: 2, fixedCount: 0, counts: { 'duplicate-cards': 1, 'amount-mismatch': 1, 'future-type': 1 } }),

    // ---- summary shapes the run history must still read
    run(X1, T(15, 17, 0), { command: 'export', dryRun: false, batch: null, limit: null }),
    end(X1, T(15, 17, 0, 1), {
      cards: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12], flags: [true, false, 'yes'], deep: { level1: { level2: { level3: 1 } } },
      unknownCounter: 3, interrupted: true, stopped: 'weird-stop', message: 'see the console output', snapshotReused: false, total: 0, pending: 2, file: 'report.xlsx',
    }),
    run(X2, T(15, 17, 1), { command: 'export', dryRun: false, batch: null, limit: null }),
    end(X2, T(15, 17, 1, 1), 'done'),
    run(X3, T(15, 17, 2), { command: 'export', dryRun: false, batch: null, limit: null }),
    end(X3, T(15, 17, 2, 1), 42),
    run(X4, T(15, 17, 3), { command: 'export', dryRun: false, batch: null, limit: null }),
    end(X4, T(15, 17, 3, 1), {}),
    run(X5, T(15, 17, 4), { command: 'issue', dryRun: false, batch: 5, limit: 1 }),
    end(X5, T(15, 17, 4, 1), issueEnd({ batch: 5 })),
    run(X6, T(15, 17, 5), { command: 'usage', dryRun: false, batch: null, limit: null }),
    end(X6, T(15, 17, 5, 1), { message: 'x'.repeat(600) }),

    // ---- reminders round 2: the run in progress (remind holds the run lock in this scenario)
    run(R2, T(16, 16, 0), { command: 'remind', dryRun: false, batch: null, limit: null, options: { round: 2, retryUnknown: false, retryFailed: true } }),
    [T(16, 16, 0, 5), { op: 'remind.start', cid: cust(1), round: 2, giftCardId: card(1001), run: R2 }],
    [T(16, 16, 0, 10), { op: 'remind.start', cid: cust(2), round: 2, giftCardId: card(1002), run: R2 }],
    [T(16, 16, 0, 11), { op: 'remind.fail', cid: cust(2), round: 2, run: R2 }], // failed without a message
    [T(16, 16, 0, 20), { op: 'remind.skip', cid: cust(3), round: 2, run: R2 }], // skipped without a reason
    [T(16, 16, 0, 30), { op: 'remind.skip', cid: cust(14), round: 2, reason: 'not-subscribed', detail: 'UNSUBSCRIBED', run: R2 }],
    [T(16, 16, 0, 40), { op: 'remind.start', cid: bulkKept, round: 2, giftCardId: card(1031), run: R2 }],
    [T(16, 16, 0, 41), { op: 'remind.ok', cid: bulkKept, round: 2, run: R2 }],
    [T(16, 16, 0, 42), { op: 'remind.tag.fail', cid: bulkKept, round: 2, error: 'tagsAdd rejected: tags: Tag limit reached [INVALID]', run: R2 }],
    [T(16, 16, 0, 50), { op: 'remind.start', cid: cust(41), round: 2, giftCardId: card(1041), run: R2 }],
    [T(16, 16, 0, 51), { op: 'remind.unknown', cid: cust(41), round: 2, error: '上次运行在发送这封提醒时中断，不知道是否已发出', run: R2 }],
    [T(16, 16, 1, 0), { op: 'remind.tag.ok', cid: cust(41), round: 1, run: R2 }], // the round-1 tag repaired
    [T(16, 16, 1, 1), { op: 'remind.tag.ok', cid: cust(963), round: 1, note: 'already tagged in Shopify', run: R2 }],
    [T(16, 16, 1, 2), { op: 'remind.tag.ok', cid: cust(964), round: 1, note: 'customer deleted', run: R2 }],
    [T(16, 16, 1, 3), { op: 'remind.skip', cid: cust(42), round: 2, reason: 'card-expired', detail: '到期日 2026-10-15', run: R2 }],
    [T(16, 16, 1, 4), { op: 'remind.start', cid: cust(15), round: 2, giftCardId: card(1015), retry: true, run: R2 }],
    [T(16, 16, 1, 5), { op: 'remind.rejected', cid: cust(15), round: 2, error: 'HTTP 429 from Shopify 6 times in a row; giving up', run: R2 }],

    // ---- entries whose run.start is missing, and an op this version does not know
    [T(16, 17, 0), { op: 'tag.ok', cid: cust(980), run: 'gone' }],
    [T(16, 17, 0, 1), { op: 'reconcile.found', cid: cust(981), giftCardId: card(1981), source: 'verify', run: 'gone' }],
    [T(16, 17, 0, 2), { op: 'future.op', cid: cust(982), note: 'something new', run: 'gone' }],
    [T(16, 17, 0, 3), { op: 'remind.ok', cid: cust(983), round: 2, run: 'gone' }],
  ];
}

function liveTags() {
  return { fetchedAt: '2026-10-06T00:00:00.000Z', tag: 'gift-card-sent-2026-10', ids: [cust(1), cust(2), cust(14), cust(41)] };
}

function liveVerify(selection) {
  const bulkKept = selection.duplicates.find((g) => g.size === 10).keptCustomerId;
  const issue = (type, customerId, giftCardId, journal, shopify, action, en) => ({
    type, customerId, giftCardId, journal, shopify, action,
    ...(en ? { journalEn: en[0], shopifyEn: en[1], actionEn: en[2] } : {}),
  });
  const issues = [
    issue('missing-in-shopify', cust(14), card(1014), '已完成：卡尾号 x014，2026-10-05 09:06 打 tag', '本活动的卡里没有这张', '在后台按客户查看；确认没有卡时可以用 issue 重新发放', ['Done: card x014, tagged 2026-10-05 09:06', 'Not among this campaign\'s cards', 'Check the customer in the admin']),
    issue('not-in-journal', cust(41), card(1041), '日志里没有这位客户的记录', '尾号 x041，原金额 $15.33，2026-10-04 09:00 建卡', '已按 Shopify 补记到本地日志', ['No journal record for this customer', 'Card x041, $15.33, created 2026-10-04 09:00', 'Recorded in the local journal from Shopify']),
    issue('not-in-selection', cust(990), card(1990), '日志里没有这位客户的记录', '尾号 x990，原金额 $10.77，2026-10-05 09:30 建卡', '这位客户不在名单里，请人工核对', ['No journal record for this customer', 'Card x990, $10.77, created 2026-10-05 09:30', 'Not on the list: check by hand']),
    issue('duplicate-cards', cust(1), card(1101), '已完成：卡尾号 x001，2026-10-05 09:00 打 tag', '尾号 x101，原金额 $15.33，2026-10-05 09:02 建卡', '核实后在后台停用这张多出来的卡（保留尾号 x001 那张）', ['Done: card x001, tagged 2026-10-05 09:00', 'Card x101, $15.33, created 2026-10-05 09:02', 'Disable this extra card in the admin (keep x001)']),
    issue('amount-mismatch', cust(14), card(1014), '名单金额 $10.77', '尾号 x014，原金额 $15.33，2026-10-05 09:06 建卡', '请人工核对金额', ['List amount $10.77', 'Card x014, $15.33, created 2026-10-05 09:06', 'Check the amount by hand']),
    issue('card-disabled', cust(2), card(1002), '已建卡未打tag：卡尾号 x002，2026-10-05 09:01 建卡', '尾号 x002，原金额 $10.77，2026-10-05 09:01 建卡，已停用', '卡已停用，确认是否需要重新启用', ['Card created, not tagged: card x002, created 2026-10-05 09:01', 'Card x002, $10.77, created 2026-10-05 09:01, disabled', 'Card disabled: check whether to enable it again']),
    issue('tag-missing', cust(2), card(1002), '已建卡未打tag：卡尾号 x002，2026-10-05 09:01 建卡', '客户没有 gift-card-sent-2026-10', '运行 issue 补打 tag', ['Card created, not tagged: card x002, created 2026-10-05 09:01', 'Customer does not carry gift-card-sent-2026-10', 'Run issue to add the tag']),
    issue('tag-without-card', cust(26), null, '日志里没有这位客户的记录', '客户带 gift-card-sent-2026-10，但本活动没有他的卡', '请人工核对', ['No journal record for this customer', 'Customer carries gift-card-sent-2026-10 but has no campaign card', 'Check by hand']),
    issue('still-unknown', cust(5), null, '需人工核对：2026-10-05 09:03 开始建卡，结果不明', '还没找到这张卡（刚建的卡可能还没进搜索索引）', '约 4 分钟后再运行 verify 或 issue 复查', ['Needs review: card creation started 2026-10-05 09:03, outcome unknown', 'Card not found yet (a new card may not be in the search index yet)', 'Run verify or issue again in about 4 minutes']),
    issue('resolved-unknown', bulkKept, card(1031), '需人工核对：2026-10-05 09:30 开始建卡，结果不明', '找到了这张卡：尾号 x031，原金额 $10.77，2026-10-05 09:30 建卡', '已按 Shopify 补记到本地日志', ['Needs review: card creation started 2026-10-05 09:30, outcome unknown', 'Found the card: x031, $10.77, created 2026-10-05 09:30', 'Recorded in the local journal from Shopify']),
    issue('future-type', null, null, '', '', '', null),
  ];
  return {
    version: 1,
    verifiedAt: '2026-10-14T17:00:00.000Z',
    cardCount: 9,
    taggedCount: 7,
    counts: { 'missing-in-shopify': 1, 'not-in-journal': 1, 'not-in-selection': 1, 'duplicate-cards': 1, 'amount-mismatch': 1, 'card-disabled': 1, 'tag-missing': 1, 'tag-without-card': 1, 'still-unknown': 1, 'resolved-unknown': 1, 'future-type': 1, 'zero-type': 0 },
    issues,
  };
}

function liveUsage(selection) {
  const amount = (n) => selection.recipients.find((r) => r.customerId === cust(n)).amountCents;
  const cards = [
    { giftCardId: card(1001), customerId: cust(1), last4: 'x001', initialCents: amount(1), balanceCents: 0, usedCents: amount(1), enabled: true, expiresOn: '2026-10-19', createdAt: '2026-10-05T16:00:02.000Z' },
    { giftCardId: card(1002), customerId: cust(2), last4: 'x002', initialCents: amount(2), balanceCents: amount(2) - 500, usedCents: 500, enabled: false, expiresOn: '2026-10-19', createdAt: '2026-10-05T16:01:02.000Z' },
    { giftCardId: card(1014), customerId: cust(14), last4: 'x014', initialCents: amount(14), balanceCents: amount(14), usedCents: 0, enabled: true, expiresOn: '2026-10-19', createdAt: '2026-10-05T16:06:02.000Z' },
    { giftCardId: card(1015), customerId: cust(15), last4: 'x015', initialCents: amount(15), balanceCents: amount(15), usedCents: 0, enabled: true, expiresOn: '2026-10-19', createdAt: '2026-10-05T16:07:05.000Z' },
    { giftCardId: card(1999), customerId: cust(8), last4: 'x999', initialCents: 1077, balanceCents: 1077, usedCents: 0, enabled: true, expiresOn: '2026-10-19', createdAt: '2026-10-05T16:20:00.000Z' },
  ];
  const issuedCents = cards.reduce((a, c) => a + c.initialCents, 0);
  const usedCents = amount(1) + 500;
  const giftCardCents = amount(1) + 700 - 200;
  const products = ['Gold Balloon', 'Helium Tank', '（无名称）', 'Silver Arch', 'Party Pack', 'Confetti', 'Ribbon', 'Weights', 'Pump', 'Banner', 'Candles', 'Cups'];
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
      { orderId: order(5004), orderName: '', orderCreatedAt: '2026-10-09T19:30:00.000Z', orderCustomerId: null, giftCardId: card(1999), cardCustomerId: cust(8), kind: 'SALE', amountCents: 300, processedAt: '2026-10-09T19:30:05.000Z' },
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
      {
        orderId: order(5004), orderName: '#5004', createdAt: '2026-10-09T19:30:00.000Z', customerId: null, totalCents: 1200, campaignCardCents: 300, cancelled: false, giftCardIds: [card(1999)],
        lineItems: [{ name: '（无名称）', sku: null, quantity: 1, amountCents: 1200 }, { sku: 'X', amountCents: 0 }],
      },
    ],
    unmatched: [
      { orderId: order(5003), orderName: '#5003', orderCreatedAt: '2026-10-08T20:00:00.000Z', customerId: cust(9), amountCents: 1000, processedAt: '2026-10-08T20:00:05.000Z', reason: '回执里没有礼品卡 ID' },
      { orderId: order(5003), orderName: '#5003', orderCreatedAt: '2026-10-08T20:00:00.000Z', customerId: cust(9), amountCents: 400, processedAt: '2026-10-09T08:00:00.000Z', reason: '回执里没有礼品卡 ID（退款）' },
      { orderId: order(5005), orderName: '#5005', orderCreatedAt: '2026-10-10T20:00:00.000Z', customerId: null, amountCents: 250, processedAt: '2026-10-10T20:00:05.000Z', reason: 'no-receipt' },
    ],
    summary: {
      issuedCards: cards.length,
      issuedCents,
      usedCards: 2,
      usedCents,
      usedRate: 0.4,
      usedCentsRate: usedCents / issuedCents,
      orders: 3,
      ordersTotalCents: 9700,
      avgOrderCents: 3233,
      giftCardCents,
      customerPaidCents: 9700 - giftCardCents,
      byTier: [
        { tier: 0, label: '$10.77', issued: 3, used: 1, rate: 0.333, usedCents: 500 },
        { tier: 1, issued: 1, used: 1, rate: 1, usedCents: amount(1) },
        { tier: 2, label: '$19.77', issued: 1, used: 0, rate: 0, usedCents: 0 },
        { tier: 7, issued: 0, used: 0, rate: 0.5, usedCents: 0 },
      ],
      byKind: [
        { kind: 'ordered', issued: 4, used: 2, rate: 0.5, usedCents },
        { kind: 'never', issued: 1, used: 0, rate: 0, usedCents: 0 },
        { kind: 'test', issued: 0, used: 0, rate: 0, usedCents: 0 },
        { kind: 'mystery', issued: 0, used: 0, rate: 7.5, usedCents: 0 },
      ],
      daily: [
        { date: '2026-10-06', newCardsUsed: 1, orders: 1, ordersTotalCents: 6000, cumulativeCardsUsed: 1, cumulativeRate: 0.2 },
        { date: '2026-10-07', newCardsUsed: 1, orders: 1, ordersTotalCents: 2500, cumulativeCardsUsed: 2, cumulativeRate: 0.4 },
        { date: 'not-a-date', newCardsUsed: 0, orders: 1, ordersTotalCents: 1200, cumulativeRate: 0.4 },
      ],
      topProducts: products.map((name, i) => ({ name, quantity: 20 - i, amountCents: 10_000 - i * 500 })),
    },
  };
}

function testCustomers() {
  return [
    makeCustomer({ n: 1, lastOrder: makeOrder({ n: 11, createdAt: '2026-09-20T18:00:00Z', total: '40.00', source: 'shopify_draft_order' }) }),
    makeCustomer({ n: 2 }),
    makeCustomer({ n: 3, email: '' }), // no email: excluded, counted only
    makeCustomer({ n: 4, tags: ['OCT26RTPROMO-TEST'] }), // already sent: excluded, listed
  ];
}

function testJournal(selection) {
  const amount = (n) => selection.recipients.find((r) => r.customerId === cust(n)).amountCents;
  const R1 = '20261003170000-2';
  const R2 = '20261003180000-2';
  return [
    ['2026-10-03T17:00:00.000Z', { op: 'run.start', run: R1, command: 'issue', dryRun: false, batch: 1, limit: 2, options: {} }],
    ['2026-10-03T17:00:01.000Z', { op: 'create.start', cid: cust(1), amountCents: amount(1), batch: 1, run: R1 }],
    ['2026-10-03T17:00:02.000Z', { op: 'create.ok', cid: cust(1), giftCardId: card(7001), last4: 'x701', amountCents: amount(1), batch: 1, run: R1 }],
    ['2026-10-03T17:00:03.000Z', { op: 'tag.ok', cid: cust(1), run: R1 }],
    ['2026-10-03T17:00:04.000Z', { op: 'skip', cid: cust(2), reason: 'not-subscribed', detail: 'PENDING', batch: 1, run: R1 }],
    ['2026-10-03T17:00:05.000Z', { op: 'run.end', run: R1, summary: { batch: 1, dryRun: false, attempted: 1, created: 1, tagged: 1, skipped: { 'not-subscribed': 1 }, amountCents: amount(1), seqFrom: 1, seqTo: 2 }, exitCode: 0 }],
    ['2026-10-03T18:00:00.000Z', { op: 'run.start', run: R2, command: 'remind', dryRun: false, batch: null, limit: null, options: { round: 1 } }],
    ['2026-10-03T18:00:01.000Z', { op: 'remind.start', cid: cust(1), round: 1, giftCardId: card(7001), run: R2 }],
    ['2026-10-03T18:00:02.000Z', { op: 'remind.ok', cid: cust(1), round: 1, run: R2 }],
    ['2026-10-03T18:00:03.000Z', { op: 'remind.found', cid: cust(2), run: R2 }],
    ['2026-10-03T18:00:04.000Z', { op: 'remind.tag.ok', cid: cust(2), run: R2 }],
    ['2026-10-03T18:00:05.000Z', { op: 'run.end', run: R2, summary: { round: 1, dryRun: false, eligible: 1, planned: 1, attempted: 1, sent: 1, alreadySentByTag: 1 }, exitCode: 0 }],
  ];
}

const toEntries = (items) => items.map(([t, entry]) => ({ t, ...entry }));

/**
 * The three scenarios with every input file, built with the real select rules.
 * Each scenario: { name, env (testConfig overrides), configPatch, running, selection,
 * journal (entries, each with its `t`), tags, verify, usage } (null = no such file).
 */
export async function buildScenarioInputs() {
  const scenarios = [];

  // ---- live
  {
    const env = testConfig();
    try {
      const data = liveCustomers();
      const selection = await selectionFixture(env.config, { ...data, write: false });
      // A Chinese channel label as an older select wrote it, and an empty key: both must read.
      selection.funnel.channels = { ...selection.funnel.channels, 网店: 2, 新版结账: 1 };
      selection.funnel.relayDomains = { ...selection.funnel.relayDomains, '': 1 };
      scenarios.push({
        name: 'live',
        env: {},
        configPatch: { remind1Date: '2026-10-13', remind2Date: '2026-10-17', giftCardExpiresOn: '2026-10-20' },
        running: 'remind',
        selection,
        journal: toEntries(liveJournal(selection)),
        tags: liveTags(),
        verify: liveVerify(selection),
        usage: liveUsage(selection),
      });
    } finally {
      env.cleanup();
    }
  }

  // ---- test
  {
    const envVars = { CAMPAIGN_ID: '2026-10-test', TEST_CUSTOMER_IDS: '1,2,3,4', SENT_TAG: 'OCT26RTPROMO-TEST' };
    const env = testConfig(envVars);
    try {
      const selection = buildTestSelection({
        config: env.config,
        customers: testCustomers().map(parseCustomer),
        timezone: TZ,
        createdAt: NOW_ISO,
        snapshot: { exportedAt: NOW_ISO, count: 4 },
      });
      selection.params.timezone = 'Asia/Shanghai';
      selection.params.giftCardExpiresOn = '';
      scenarios.push({
        name: 'test',
        env: envVars,
        configPatch: { launchDate: '', giftCardExpiresOn: '' },
        running: null,
        selection,
        journal: toEntries(testJournal(selection)),
        tags: null,
        verify: null,
        usage: null,
      });
    } finally {
      env.cleanup();
    }
  }

  // ---- empty
  {
    const envVars = { GIFT_TIERS: '10.00' };
    const env = testConfig(envVars);
    try {
      const customers = [makeCustomer({ n: 1, createdAt: '2025-02-01T00:00:00Z' }), makeCustomer({ n: 2, marketingState: 'UNSUBSCRIBED' })];
      const selection = await selectionFixture(env.config, { customers, write: false });
      selection.params.timezone = 'Mars/Olympus_Mons';
      scenarios.push({
        name: 'empty',
        env: envVars,
        configPatch: {},
        running: null,
        selection,
        journal: [],
        tags: { fetchedAt: '2026-10-06T00:00:00.000Z', tag: 'some-other-tag', ids: [cust(1)] },
        verify: { version: 1, cardCount: 0, taggedCount: 0, counts: {}, issues: [] },
        usage: { version: 1, cards: [], payments: [], orders: [], unmatched: [], summary: {} },
      });
    } finally {
      env.cleanup();
    }
  }
  return scenarios;
}

// ---------------------------------------------------------------------------
// Using the frozen inputs
// ---------------------------------------------------------------------------

/** The frozen scenarios (test/fixtures/report-scenarios.json). */
export function loadScenarios() {
  return JSON.parse(fs.readFileSync(SCENARIOS_FILE, 'utf8')).scenarios;
}

/**
 * A fresh campaign folder holding `scenario`'s input files. `write(options)` regenerates the
 * workbook(s) with writeReport (extra options are passed through), holding the scenario's run
 * lock while it does. Call `cleanup()` at the end.
 */
export function setupScenario(scenario, writeReport) {
  const env = testConfig(scenario.env);
  const config = { ...env.config, ...scenario.configPatch };
  const paths = campaignPaths(config);
  ensureDir(paths.dir);
  writeJsonAtomic(paths.selection, scenario.selection);
  if (scenario.journal.length) fs.writeFileSync(paths.journal, `${scenario.journal.map((e) => JSON.stringify(e)).join('\n')}\n`);
  for (const key of ['tags', 'verify', 'usage']) if (scenario[key]) writeJsonAtomic(paths[key], scenario[key]);
  const write = async (options = {}) => {
    const release = scenario.running ? acquireRunLock(paths, scenario.running) : null;
    try {
      return await writeReport({ config, paths, log: memoryLog(), now: FIXED_NOW, lockOptions: FAST_LOCK, ...options });
    } finally {
      release?.();
    }
  };
  return { env, config, paths, write, cleanup: () => env.cleanup() };
}

// ---------------------------------------------------------------------------
// Workbook snapshots: every cell's value, formula and style, per sheet
// ---------------------------------------------------------------------------

function cellValueSnapshot(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return { d: v.toISOString() };
  if (typeof v === 'object') {
    if (typeof v.formula === 'string') return { f: v.formula, r: v.result ?? null };
    if (Array.isArray(v.richText)) return { rt: v.richText.map((r) => r.text).join('') };
    if ('error' in v) return { e: String(v.error) };
    return { o: JSON.stringify(v) };
  }
  return v;
}

/** A short text naming everything visible about a cell's style ('' = default). */
function styleSnapshot(cell) {
  const font = cell.font ?? {};
  const fill = cell.fill?.type === 'pattern' ? cell.fill.fgColor?.argb ?? '' : '';
  const align = cell.alignment ?? {};
  const border = cell.border && Object.keys(cell.border).length ? 'grid' : '';
  const parts = [
    cell.numFmt ?? '',
    font.bold ? 'bold' : '',
    font.size ? String(font.size) : '',
    font.color?.argb ?? '',
    font.underline ? 'u' : '',
    fill,
    align.wrapText ? 'wrap' : '',
    align.horizontal ?? '',
    align.vertical ?? '',
    border,
  ];
  return parts.some(Boolean) ? parts.join('|') : '';
}

/**
 * Every sheet of a workbook: { name, columns (widths), views, autoFilter, protection,
 * rows: [rowNumber, height|null, [[col, value, style], ...]] } with rows and cells in order
 * (cells that hold neither a value nor a style are left out).
 */
export async function snapshotWorkbook(file) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  return wb.worksheets.map((ws) => {
    const rows = [];
    for (let r = 1; r <= ws.rowCount; r += 1) {
      const row = ws.getRow(r);
      const cells = [];
      for (let c = 1; c <= row.cellCount; c += 1) {
        const cell = row.getCell(c);
        const value = cellValueSnapshot(cell.value);
        const style = styleSnapshot(cell);
        if (value !== null || style) cells.push([c, value, style]);
      }
      rows.push([r, row.height ?? null, cells]);
    }
    const columns = [];
    for (let c = 1; c <= ws.columnCount; c += 1) columns.push(ws.getColumn(c).width ?? null);
    const view = ws.views?.[0] ?? {};
    return {
      name: ws.name,
      columns,
      views: { state: view.state ?? null, xSplit: view.xSplit ?? null, ySplit: view.ySplit ?? null },
      autoFilter: ws.autoFilter ?? null,
      protection: ws.sheetProtection ? { sheet: !!ws.sheetProtection.sheet, autoFilter: !!ws.sheetProtection.autoFilter, formatColumns: !!ws.sheetProtection.formatColumns } : null,
      rows,
    };
  });
}

/**
 * What both language editions must share: every sheet's columns, views, filter and protection,
 * and per row its height and per cell the position, the style and the value, except that a text
 * is only "text" and the words inside a number format ('#,##0" 张"') are left out.
 * Numbers, dates, HYPERLINK formulas (target and label) and empty cells must be identical.
 */
export function structureOf(sheets) {
  const neutral = (style) => style.replace(/"[^"]*"/g, '""');
  return sheets.map((s) => ({
    columns: s.columns,
    views: s.views,
    autoFilter: s.autoFilter,
    protection: s.protection,
    rows: s.rows.map(([r, height, cells]) => [r, height, cells.map(([c, v, style]) => [c, typeof v === 'string' ? 'text' : v, neutral(style)])]),
  }));
}

/**
 * The cells whose text has a CJK or full-width character, as "<sheet>!R<row>C<col>: <text>"
 * (the English workbook may show them only inside customer data and Shopify's own words).
 */
export const CJK = /[　-〿㐀-䶿一-鿿豈-﫿＀-￯]/;
export function cjkCells(sheets) {
  const out = [];
  for (const s of sheets) {
    for (const [r, , cells] of s.rows) {
      for (const [c, v] of cells) {
        const text = typeof v === 'string' ? v : v && typeof v === 'object' ? [v.f, v.r, v.rt].filter((x) => typeof x === 'string').join(' ') : '';
        if (CJK.test(text)) out.push(`${s.name}!R${r}C${c}: ${text}`);
      }
    }
    if (CJK.test(s.name)) out.push(`${s.name} (sheet name)`);
  }
  return out;
}

// ---------------------------------------------------------------------------
// The workbook's raw XML parts (what Microsoft Excel actually parses)
// ---------------------------------------------------------------------------

/** JSZip as installed for ExcelJS (which reads .xlsx files with it); null when it cannot be resolved. */
function loadJSZip() {
  try {
    const requireHere = createRequire(import.meta.url);
    return createRequire(requireHere.resolve('exceljs/package.json'))('jszip');
  } catch {
    return null;
  }
}

/** { partName: text } of every file in an .xlsx (JSZip, else the `unzip` command). */
export async function xlsxParts(file) {
  const JSZip = loadJSZip();
  const parts = {};
  if (JSZip) {
    const zip = await JSZip.loadAsync(fs.readFileSync(file));
    for (const [name, entry] of Object.entries(zip.files)) if (!entry.dir) parts[name] = await entry.async('string');
    return parts;
  }
  const names = execFileSync('unzip', ['-Z1', file], { encoding: 'utf8' }).split('\n').filter((n) => n && !n.endsWith('/'));
  for (const name of names) parts[name] = execFileSync('unzip', ['-p', file, name], { encoding: 'utf8', maxBuffer: 512 * 1024 * 1024 });
  return parts;
}

const attr = (tag, name) => new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1] ?? null;

/** Sheet name → its worksheet part name (e.g. '发放名单' → 'xl/worksheets/sheet3.xml'). */
export function worksheetPartsByName(parts) {
  const rels = new Map([...parts['xl/_rels/workbook.xml.rels'].matchAll(/<Relationship\b[^>]*>/g)].map((m) => [attr(m[0], 'Id'), attr(m[0], 'Target')]));
  const out = {};
  for (const m of parts['xl/workbook.xml'].matchAll(/<sheet\b[^>]*>/g)) {
    const target = rels.get(attr(m[0], 'r:id'));
    const name = attr(m[0], 'name').replace(/&amp;/g, '&');
    out[name] = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
  }
  return out;
}

// CT_Worksheet in ECMA-376 / ISO-IEC 29500 sml.xsd: the children of <worksheet>, in this order.
const CT_WORKSHEET_ORDER = [
  'sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection',
  'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells',
  'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup',
  'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing',
  'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst',
];
const REPEATABLE = new Set(['cols', 'conditionalFormatting']);

/** Names of the direct children of <worksheet>, in document order. */
export function worksheetChildren(xml) {
  const body = xml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, '<sheetData/>');
  const names = [];
  let depth = 0;
  for (const m of body.slice(body.indexOf('<worksheet')).matchAll(/<(\/?)([A-Za-z][\w:.-]*)\b[^>]*?(\/?)>/g)) {
    const [, close, name, selfClosing] = m;
    if (close) {
      depth -= 1;
      continue;
    }
    if (depth === 1) names.push(name);
    if (!selfClosing) depth += 1;
  }
  return names;
}

/** Problems with a worksheet part's child elements against CT_Worksheet ([] = valid order). */
export function worksheetOrderProblems(xml) {
  const problems = [];
  const children = worksheetChildren(xml);
  let last = -1;
  let lastName = null;
  for (const name of children) {
    const i = CT_WORKSHEET_ORDER.indexOf(name);
    if (i === -1) problems.push(`<${name}> is not a child of CT_Worksheet`);
    else if (i < last || (i === last && !REPEATABLE.has(name))) problems.push(`<${name}> after <${lastName}>`);
    else {
      last = i;
      lastName = name;
    }
  }
  if (children.filter((n) => n === 'sheetData').length !== 1) problems.push('needs exactly one <sheetData>');
  return { children, problems };
}

// XML 1.0 Char: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
const NOT_XML_CHAR = /[^\t\n\r\x20-퟿-�\u{10000}-\u{10FFFF}]/u;

export const HAS_XMLLINT = spawnSync('xmllint', ['--version'], { encoding: 'utf8' }).error === undefined;

/** xmllint --noout on one XML text: null when well-formed, else the parser's message. */
function xmllintError(xml) {
  const r = spawnSync('xmllint', ['--noout', '-'], { input: xml, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
  return r.status === 0 ? null : (r.stderr || r.stdout || `exit ${r.status}`).trim().split('\n').slice(0, 3).join(' | ');
}

/** Every XML part of the workbook that is not well-formed (xmllint when installed, plus a check of the characters). */
export function malformedParts(parts) {
  const bad = [];
  for (const [name, xml] of Object.entries(parts)) {
    if (!/\.(xml|rels)$/.test(name)) continue;
    const ch = NOT_XML_CHAR.exec(xml);
    if (ch) bad.push(`${name}: character U+${ch[0].codePointAt(0).toString(16).toUpperCase().padStart(4, '0')} is not allowed in XML`);
    if (HAS_XMLLINT) {
      const err = xmllintError(xml);
      if (err) bad.push(`${name}: ${err}`);
    }
  }
  return bad;
}

/**
 * JSON with one sheet row per line: readable diffs for the golden files.
 * Objects and arrays down to `depth` are spread over lines, anything deeper is inline;
 * an array of plain values always stays on one line.
 */
export function formatJson(value, depth = 4, indent = '') {
  if (depth <= 0 || value === null || typeof value !== 'object') return JSON.stringify(value);
  const inner = `${indent}  `;
  if (Array.isArray(value)) {
    if (!value.length || value.every((v) => v === null || typeof v !== 'object')) return JSON.stringify(value);
    return `[\n${value.map((v) => `${inner}${formatJson(v, depth - 1, inner)}`).join(',\n')}\n${indent}]`;
  }
  const keys = Object.keys(value).filter((k) => value[k] !== undefined);
  if (!keys.length) return '{}';
  return `{\n${keys.map((k) => `${inner}${JSON.stringify(k)}: ${formatJson(value[k], depth - 1, inner)}`).join(',\n')}\n${indent}}`;
}
