// The audience rules, same-address dedupe and amounts, end to end through
// buildSelection (the same code path select uses), with Shopify-shaped nodes.
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { buildSelection, buildTestSelection, selectionParams, normalizeHistory, recipientIds } from '../src/select/selection.js';
import { parseCustomer, buildActivity, screen, summarize, channelLabel, reasonText, reasonDay, RULES, UNLISTED_RULES } from '../src/select/rules.js';
import { percentOf } from '../src/select/amount.js';
import { localDate } from '../src/time.js';
import { makeCustomer, makeOrder, makeActiveOrder, testConfig, NOW_MS, NOW_ISO, CUTOFF_MS } from './helpers.js';

const { config, cleanup } = testConfig({ SENT_TAG: 'OCT26RTPROMO' });
after(cleanup);

const DAY = 86_400_000;
const ADDR = (address1, zip = '78701') => ({ address1, address2: '', city: 'Austin', provinceCode: 'TX', zip, countryCodeV2: 'US' });

function select(nodes, { activeOrders = [], history = {}, cfg = config } = {}) {
  return buildSelection({
    config: cfg,
    customers: nodes.map(parseCustomer),
    activeOrders,
    history,
    nowMs: NOW_MS,
    cutoffMs: CUTOFF_MS,
    cutoffIso: '2026-07-01T00:00:00-07:00',
    timezone: 'America/Los_Angeles',
    createdAt: NOW_ISO,
    snapshot: { exportedAt: NOW_ISO, source: 'bulk', count: nodes.length },
  });
}
const byId = (list) => new Map(list.map((x) => [x.customerId, x]));
const cid = (n) => `gid://shopify/Customer/${n}`;
const paid = (n, createdAt, total, extra = {}) => makeOrder({ n, createdAt, total, ...extra });

test('rules: each exclusion rule, in order, with the first failing rule as the main reason', () => {
  const active = makeCustomer({ n: 11, address: ADDR('11 Busy St'), lastOrder: paid(1101, '2026-08-15T12:00:00Z', '40.00') });
  const nodes = [
    makeCustomer({ n: 1, lastOrder: paid(101, '2026-03-01T12:00:00Z', '150.00') }), // selected
    makeCustomer({ n: 2, email: null }), // 1 no email
    makeCustomer({ n: 3, validFormat: false }), // 1 invalid format
    makeCustomer({ n: 4, email: 'x@mail.codisto.com' }), // 2 relay domain
    makeCustomer({ n: 5, email: 'Buyer@JokerPartySupply.com' }), // 2 company domain, any case
    makeCustomer({ n: 6, marketingState: 'UNSUBSCRIBED' }), // 3
    makeCustomer({ n: 7, marketingState: 'NOT_SUBSCRIBED', email: 'y@mail.codisto.com' }), // 2 first, 3 also listed
    makeCustomer({ n: 8, createdAt: new Date(NOW_MS - 3 * DAY).toISOString(), tags: ['WHS'] }), // 4 first, 5 also
    makeCustomer({ n: 9, tags: ['whs'] }), // 5, case-insensitive
    makeCustomer({ n: 10, tags: ['former-fake-identified'], lastOrder: paid(1001, '2026-02-01T12:00:00Z', '80.00') }), // selected
    active, // 7: valid order inside the window
    makeCustomer({ n: 12, tags: ['OCT26RTPROMO'] }), // 6 already sent
    makeCustomer({ n: 13, address: ADDR('11 Busy Street') }), // 8 same address as 11
    makeCustomer({ n: 14, lastOrder: paid(1401, '2026-01-10T12:00:00Z', '90.00', { source: 'amazon' }) }), // 9
    makeCustomer({ n: 15, lastOrder: paid(1501, '2026-01-11T12:00:00Z', '90.00', { source: '205641' }) }), // 9 Sellbrite
    makeCustomer({ n: 16, lastOrder: paid(1601, '2026-01-12T12:00:00Z', '90.00') }), // 7 via ORDERS_SINCE only
    makeCustomer({ n: 17, lastOrder: paid(1701, '2026-08-20T12:00:00Z', '90.00', { cancelled: true }) }), // cancelled recent order: not "recent"
    makeCustomer({ n: 18, createdAt: new Date(NOW_MS - 7 * DAY).toISOString() }), // exactly 7 days: allowed
  ];
  const activeOrders = [
    makeActiveOrder({ n: 1101, customer: active, createdAt: '2026-08-15T12:00:00Z' }),
    makeActiveOrder({ n: 1602, customer: { id: cid(16), defaultAddress: nodes[15].defaultAddress }, createdAt: '2026-09-01T12:00:00Z' }),
    makeActiveOrder({ n: 9001, customer: { id: cid(99), defaultAddress: ADDR('500 Other Rd') }, createdAt: '2026-09-02T12:00:00Z', cancelled: true }),
  ];
  const history = { [cid(17)]: [paid(1701, '2026-08-20T12:00:00Z', '90.00', { cancelled: true }), paid(1702, '2026-01-02T12:00:00Z', '60.00')] };
  const s = select(nodes, { activeOrders, history });

  assert.deepEqual(recipientIds(s).sort(), [cid(1), cid(10), cid(17), cid(18)].sort());
  const ns = byId(s.notSelected);
  const primary = (n) => ns.get(cid(n))?.primaryCode;
  // Rules 1-3 are counted but not listed row by row.
  for (const n of [2, 3, 4, 5, 6, 7]) assert.equal(ns.has(cid(n)), false, `customer ${n} is not listed`);
  assert.equal(primary(8), 'too-new');
  assert.match(ns.get(cid(8)).allReasons, /4\. 注册不满 7 天：.*；5\. 带排除 tag：WHS/);
  assert.equal(primary(9), 'excluded-tag');
  assert.equal(ns.get(cid(9)).primaryText, '5. 带排除 tag：WHS');
  assert.equal(primary(11), 'recent-order');
  assert.equal(primary(12), 'already-sent');
  assert.equal(primary(13), 'active-address');
  assert.equal(ns.get(cid(13)).relatedCustomerId, cid(11));
  assert.equal(ns.get(cid(13)).relatedAt, '2026-08-15T12:00:00Z');
  assert.equal(primary(14), 'marketplace-order');
  assert.equal(ns.get(cid(14)).primaryText, '9. 最近订单来自平台渠道：Amazon');
  assert.equal(primary(15), 'marketplace-order');
  assert.equal(ns.get(cid(15)).primaryText, '9. 最近订单来自平台渠道：Sellbrite');
  assert.equal(primary(16), 'recent-order', 'an order found by the ORDERS_SINCE query counts even if lastOrder is older');
  // The reason shows the order that made them active, not the snapshot's older lastOrder.
  assert.equal(ns.get(cid(16)).primaryText, '7. 近 3 个月有下单：2026-09-01');
  assert.equal(ns.get(cid(11)).primaryText, '7. 近 3 个月有下单：2026-08-15');

  // The reasons as data (for the English workbook), in the order of the texts, with the raw details.
  assert.deepEqual(ns.get(cid(8)).reasons, [
    { n: 4, code: 'too-new', detail: nodes[7].createdAt },
    { n: 5, code: 'excluded-tag', detail: 'WHS' },
  ]);
  assert.deepEqual(ns.get(cid(13)).reasons, [{ n: 8, code: 'active-address', detail: cid(11) }]);
  assert.deepEqual(ns.get(cid(14)).reasons, [{ n: 9, code: 'marketplace-order', detail: 'amazon' }]);
  assert.deepEqual(ns.get(cid(16)).reasons, [{ n: 7, code: 'recent-order', detail: '2026-09-01T12:00:00Z' }]);
  for (const row of s.notSelected) {
    assert.equal(row.reasons[0].code, row.primaryCode, 'the first reason is the primary one');
    assert.equal(row.reasons.length, row.allReasons.split('；').length, 'one reason per text segment');
  }

  const count = Object.fromEntries(s.funnel.byRule.map((r) => [r.code, r.count]));
  assert.deepEqual(count, {
    'no-email': 2, 'relay-email': 3, 'not-subscribed': 1, 'too-new': 1, 'excluded-tag': 1, 'already-sent': 1,
    'recent-order': 2, 'active-address': 1, 'marketplace-order': 2, 'duplicate-address': 0, 'zero-amount': 0,
    'customer-deleted': 0,
  });
  assert.equal(s.funnel.total, nodes.length);
  assert.equal(s.funnel.unlisted, 6);
  assert.equal(s.funnel.listed, s.notSelected.length);
  assert.equal(s.funnel.recipients, 4);
  assert.equal(s.funnel.total, s.funnel.unlisted + s.funnel.listed + s.funnel.recipients, 'everyone is accounted for exactly once');
  assert.deepEqual(s.funnel.relayDomains, { 'mail.codisto.com': 2, 'jokerpartysupply.com': 1 });
  assert.deepEqual(s.funnel.notSubscribed, { UNSUBSCRIBED: 1 });
  assert.deepEqual(s.funnel.tags, { WHS: 1 });
  assert.deepEqual(s.funnel.channels, { Amazon: 1, Sellbrite: 1 });
  assert.deepEqual(s.notSelected.map((x) => x.primaryRule), [...s.notSelected.map((x) => x.primaryRule)].sort((a, b) => a - b), 'sorted by rule');
});

test('rules: the activity check ignores cancelled and test orders; $0 orders count', () => {
  const a = makeCustomer({ n: 1, address: ADDR('1 A St') });
  const b = makeCustomer({ n: 2, address: ADDR('2 B St') });
  const c = makeCustomer({ n: 3, address: ADDR('3 C St') });
  const activity = buildActivity([
    makeActiveOrder({ n: 1, customer: a, cancelled: true }),
    makeActiveOrder({ n: 2, customer: b, test: true }),
    makeActiveOrder({ n: 3, customer: c }), // ORDERS_SINCE has no totals: a $0 order is an order
    makeActiveOrder({ n: 4, customer: null }),
  ]);
  assert.deepEqual([...activity.activeCustomerIds], [cid(3)]);
  assert.deepEqual([...activity.activeAddresses.keys()], ['3 c st|78701|us']);
  // The most recent order wins for an address shared by several buyers.
  const d = makeCustomer({ n: 4, address: ADDR('3 C Street') });
  const act2 = buildActivity([makeActiveOrder({ n: 5, customer: c, createdAt: '2026-08-01T00:00:00Z' }), makeActiveOrder({ n: 6, customer: d, createdAt: '2026-09-01T00:00:00Z' })]);
  assert.deepEqual(act2.activeAddresses.get('3 c st|78701|us'), { customerId: cid(4), createdAt: '2026-09-01T00:00:00Z' });
});

test('rules: activeAt holds each customer\'s newest valid order; the recent-order reason shows that date', () => {
  const a = makeCustomer({ n: 1, lastOrder: paid(11, '2026-01-12T12:00:00Z', '90.00') }); // stale snapshot lastOrder
  const activity = buildActivity([
    makeActiveOrder({ n: 12, customer: a, createdAt: '2026-08-01T12:00:00Z' }),
    makeActiveOrder({ n: 13, customer: a, createdAt: '2026-09-20T12:00:00Z', cancelled: true }), // does not count
    makeActiveOrder({ n: 14, customer: a, createdAt: '2026-09-25T12:00:00Z', test: true }), // does not count
    makeActiveOrder({ n: 15, customer: a, createdAt: '2026-09-01T12:00:00Z' }), // newest valid
    makeActiveOrder({ n: 16, customer: a, createdAt: '2026-07-10T12:00:00Z' }),
  ]);
  assert.deepEqual([...activity.activeAt], [[cid(1), '2026-09-01T12:00:00Z']]);
  const s = select([a], { activeOrders: [makeActiveOrder({ n: 15, customer: a, createdAt: '2026-09-01T12:00:00Z' })] });
  assert.equal(byId(s.notSelected).get(cid(1)).primaryText, '7. 近 3 个月有下单：2026-09-01');

  // A lastOrder that is cancelled: the reason shows the valid order, not the cancelled one.
  const b = makeCustomer({ n: 2, lastOrder: paid(21, '2026-09-28T12:00:00Z', '30.00', { cancelled: true }) });
  const sb = select([b], { activeOrders: [makeActiveOrder({ n: 22, customer: b, createdAt: '2026-08-05T12:00:00Z' })], history: { [cid(2)]: [] } });
  assert.equal(byId(sb.notSelected).get(cid(2)).primaryText, '7. 近 3 个月有下单：2026-08-05');

  // A valid in-window lastOrder that the orders query missed is recorded too (and newer wins).
  const c = makeCustomer({ n: 3, lastOrder: paid(31, '2026-09-15T12:00:00Z', '30.00') });
  const sc = select([c], { activeOrders: [makeActiveOrder({ n: 32, customer: c, createdAt: '2026-07-20T12:00:00Z' })] });
  assert.equal(byId(sc.notSelected).get(cid(3)).primaryText, '7. 近 3 个月有下单：2026-09-15');
  const d = makeCustomer({ n: 4, lastOrder: paid(41, '2026-08-15T12:00:00Z', '30.00') });
  assert.equal(byId(select([d]).notSelected).get(cid(4)).primaryText, '7. 近 3 个月有下单：2026-08-15');

  // An activity without activeAt (older callers) falls back to the snapshot's lastOrder date.
  const e = parseCustomer(makeCustomer({ n: 5, lastOrder: paid(51, '2026-02-01T12:00:00Z', '30.00') }));
  const screened = screen([e], { config, nowMs: NOW_MS, cutoffMs: CUTOFF_MS, activity: { activeCustomerIds: new Set([e.id]), activeAddresses: new Map() } });
  assert.deepEqual(screened.excluded[0].reasons, [{ code: 'recent-order', detail: '2026-02-01T12:00:00Z' }]);
});

test('rule 9: the most recent valid order ($0 included) is checked as well as the order the amount comes from', () => {
  const nodes = [
    // Latest valid order is a $0 Sellbrite order; the paid order before it is web → excluded (Sellbrite).
    makeCustomer({ n: 1, lastOrder: paid(12, '2026-05-01T12:00:00Z', '0.00', { source: '205641' }) }),
    // Last order cancelled on Amazon; the latest valid one is web and paid → a recipient (cancelled orders do not count).
    makeCustomer({ n: 2, lastOrder: paid(22, '2026-05-01T12:00:00Z', '50.00', { cancelled: true, source: 'amazon' }) }),
    // Last order cancelled; latest valid order is a $0 web order, the paid basis came through Etsy → excluded (Etsy).
    makeCustomer({ n: 3, lastOrder: paid(32, '2026-05-01T12:00:00Z', '50.00', { cancelled: true }) }),
    // Last order a test order from eBay; the latest valid one is web → a recipient (test orders do not count).
    makeCustomer({ n: 4, lastOrder: paid(42, '2026-05-01T12:00:00Z', '50.00', { test: true, source: 'ebay' }) }),
    // Only a $0 Walmart order and cancelled ones: no basis, but the latest valid order is Walmart → excluded.
    makeCustomer({ n: 5, lastOrder: paid(52, '2026-05-01T12:00:00Z', '0.00', { source: 'walmart' }) }),
    // Latest valid order is a $0 web order, the paid basis is Amazon → excluded (Amazon), as before.
    makeCustomer({ n: 6, lastOrder: paid(62, '2026-05-01T12:00:00Z', '0.00') }),
    // Paid web last order: a recipient.
    makeCustomer({ n: 7, lastOrder: paid(72, '2026-05-01T12:00:00Z', '50.00') }),
  ];
  const history = {
    [cid(1)]: [paid(12, '2026-05-01T12:00:00Z', '0.00', { source: '205641' }), paid(11, '2026-03-01T12:00:00Z', '60.00')],
    [cid(2)]: [paid(22, '2026-05-01T12:00:00Z', '50.00', { cancelled: true, source: 'amazon' }), paid(21, '2026-03-01T12:00:00Z', '60.00')],
    [cid(3)]: [paid(32, '2026-05-01T12:00:00Z', '50.00', { cancelled: true }), paid(31, '2026-04-01T12:00:00Z', '0.00'), paid(30, '2026-03-01T12:00:00Z', '60.00', { source: 'etsy' })],
    [cid(4)]: [paid(42, '2026-05-01T12:00:00Z', '50.00', { test: true, source: 'ebay' }), paid(41, '2026-03-01T12:00:00Z', '60.00')],
    [cid(5)]: [paid(52, '2026-05-01T12:00:00Z', '0.00', { source: 'walmart' }), paid(51, '2026-03-01T12:00:00Z', '60.00', { cancelled: true })],
    [cid(6)]: [paid(62, '2026-05-01T12:00:00Z', '0.00'), paid(61, '2026-03-01T12:00:00Z', '60.00', { source: 'amazon' })],
  };
  const s = select(nodes, { history });
  assert.deepEqual(recipientIds(s).sort(), [cid(2), cid(4), cid(7)].sort());
  const ns = byId(s.notSelected);
  assert.equal(ns.get(cid(1)).primaryText, '9. 最近订单来自平台渠道：Sellbrite');
  assert.equal(ns.get(cid(3)).primaryText, '9. 最近订单来自平台渠道：Etsy');
  assert.equal(ns.get(cid(5)).primaryText, '9. 最近订单来自平台渠道：Walmart');
  assert.equal(ns.get(cid(6)).primaryText, '9. 最近订单来自平台渠道：Amazon');
  assert.deepEqual(s.funnel.channels, { Sellbrite: 1, Etsy: 1, Walmart: 1, Amazon: 1 });
  assert.equal(byId(s.recipients).get(cid(2)).basis.sourceName, 'web');
});

test('rule 12: a candidate Shopify no longer has at the follow-up lookup is left out, listed and counted', () => {
  const nodes = [
    makeCustomer({ n: 1, lastOrder: paid(11, '2026-02-01T12:00:00Z', '80.00', { cancelled: true }) }), // deleted after the export
    makeCustomer({ n: 2, lastOrder: paid(21, '2026-02-01T12:00:00Z', '0.00') }), // exists, no orders left
    makeCustomer({ n: 3, lastOrder: paid(31, '2026-03-01T12:00:00Z', '100.00') }),
  ];
  const history = { [cid(1)]: { orders: [], deleted: true }, [cid(2)]: { orders: [] } };
  const s = select(nodes, { history });
  assert.deepEqual(recipientIds(s), [cid(3), cid(2)]);
  assert.equal(byId(s.recipients).get(cid(2)).neverReason, 'only-cancelled-or-zero', 'an existing customer without orders is still "never ordered"');
  const row = byId(s.notSelected).get(cid(1));
  assert.equal(row.primaryCode, 'customer-deleted');
  assert.equal(row.primaryRule, 12);
  assert.equal(row.primaryText, '12. 补查订单时客户已被删除');
  assert.equal(row.allReasons, '12. 补查订单时客户已被删除');
  assert.deepEqual(row.reasons, [{ n: 12, code: 'customer-deleted' }], 'a reason without a detail has no detail field');
  assert.equal(s.funnel.byRule.find((r) => r.code === 'customer-deleted').count, 1);
  assert.equal(s.funnel.byRule.at(-1).label, '补查订单时客户已被删除');
  assert.equal(UNLISTED_RULES.has('customer-deleted'), false, 'listed row by row in 未入选');
  assert.equal(s.stats.recipients, 2);
  assert.equal(s.stats.totalCents, 1077 + 1077);
  assert.equal(s.averageCents, 10000, 'the deleted customer is not part of the average');
  assert.equal(s.funnel.total, s.funnel.unlisted + s.funnel.listed + s.funnel.recipients);
  // The flag survives normalizeHistory, from a Map or a plain object.
  assert.equal(normalizeHistory(new Map([[cid(1), { orders: [], deleted: true }]])).get(cid(1)).deleted, true);
  assert.equal(normalizeHistory({ [cid(1)]: { orders: [], percent: 10, deleted: true } }).get(cid(1)).deleted, true);
  assert.equal(normalizeHistory({ [cid(2)]: [] }).get(cid(2)).deleted, false);
});

test('amounts: 10% of the last paid order mapped to a tier; never-ordered get the tier of 10% of the average', () => {
  const nodes = [
    makeCustomer({ n: 1, lastOrder: paid(1, '2026-06-01T12:00:00Z', '107.74') }), // $10.77 → $10.77
    makeCustomer({ n: 2, lastOrder: paid(2, '2026-05-01T12:00:00Z', '107.75') }), // $10.78 → $15.33
    makeCustomer({ n: 3, lastOrder: paid(3, '2026-04-01T12:00:00Z', '153.40') }), // $15.34 → $19.77
    makeCustomer({ n: 4, lastOrder: paid(4, '2026-03-01T12:00:00Z', '31.11') }), // $3.11 → $10.77
    makeCustomer({ n: 5, createdAt: '2025-01-01T00:00:00Z' }), // never ordered
  ];
  const s = select(nodes);
  const r = byId(s.recipients);
  assert.deepEqual([1, 2, 3, 4, 5].map((n) => r.get(cid(n)).amountCents), [1077, 1533, 1977, 1077, 1077]);
  assert.deepEqual([1, 2, 3, 4].map((n) => r.get(cid(n)).tier), [0, 1, 2, 0]);
  assert.equal(r.get(cid(2)).formula, '10% × $107.75 = $10.78 → 档位 $15.33');
  // Average of the ordered recipients' last paid totals: (107.74 + 107.75 + 153.40 + 31.11) / 4 = $100.00
  assert.equal(s.averageCents, 10000);
  assert.equal(s.neverAmountCents, 1077);
  const never = r.get(cid(5));
  assert.equal(never.kind, 'never');
  assert.equal(never.neverReason, 'no-orders');
  assert.equal(never.basis, null);
  assert.equal(never.formula, '10% × 平均 $100.00 = $10.00 → 档位 $10.77');
  assert.equal(s.stats.totalCents, 1077 + 1533 + 1977 + 1077 + 1077);
  assert.deepEqual(s.stats.tiers.map((t) => [t.label, t.ordered, t.never, t.count, t.cents]), [['$10.77', 2, 1, 3, 3231], ['$15.33', 1, 0, 1, 1533], ['$19.77', 1, 0, 1, 1977]]);
  assert.equal(s.stats.orderedCount, 4);
  assert.equal(s.stats.neverCount, 1);
  assert.equal(s.stats.medianOrderedRawCents, 1078);
});

test('amounts: a large average lifts never-ordered customers into a higher tier', () => {
  const s = select([makeCustomer({ n: 1, lastOrder: paid(1, '2026-06-01T12:00:00Z', '400.00') }), makeCustomer({ n: 2 })]);
  const never = byId(s.recipients).get(cid(2));
  assert.equal(never.amountCents, 1977);
  assert.equal(never.tier, 2);
});

test('amounts: cancelled, test, $0 and sub-5-cent last orders look further back', () => {
  const nodes = [
    makeCustomer({ n: 1, lastOrder: paid(11, '2026-02-01T12:00:00Z', '80.00', { cancelled: true }) }),
    makeCustomer({ n: 2, lastOrder: paid(21, '2026-02-10T12:00:00Z', '0.00') }),
    makeCustomer({ n: 3, lastOrder: paid(31, '2026-02-11T12:00:00Z', '50.00', { cancelled: true }) }),
    makeCustomer({ n: 4, lastOrder: paid(41, '2026-02-12T12:00:00Z', '0.04') }),
    makeCustomer({ n: 5, lastOrder: paid(51, '2026-08-01T12:00:00Z', '30.00', { cancelled: true }) }),
    makeCustomer({ n: 6, lastOrder: paid(61, '2026-08-02T12:00:00Z', '30.00', { test: true }) }),
    makeCustomer({ n: 7, lastOrder: paid(71, '2026-03-01T12:00:00Z', '100.00') }),
  ];
  const history = {
    [cid(1)]: [paid(11, '2026-02-01T12:00:00Z', '80.00', { cancelled: true }), paid(12, '2026-01-01T12:00:00Z', '23.00')],
    [cid(2)]: [paid(21, '2026-02-10T12:00:00Z', '0.00'), paid(22, '2025-12-01T12:00:00Z', '200.00')],
    [cid(3)]: [paid(31, '2026-02-11T12:00:00Z', '50.00', { cancelled: true }), paid(32, '2026-01-11T12:00:00Z', '0.00')],
    [cid(4)]: [paid(41, '2026-02-12T12:00:00Z', '0.04')],
    [cid(5)]: [paid(51, '2026-08-01T12:00:00Z', '30.00', { cancelled: true }), paid(52, '2026-07-15T12:00:00Z', '30.00')],
    [cid(6)]: [paid(61, '2026-08-02T12:00:00Z', '30.00', { test: true }), paid(62, '2026-01-05T12:00:00Z', '120.00')],
  };
  const s = select(nodes, { history });
  const r = byId(s.recipients);
  const one = r.get(cid(1));
  assert.equal(one.kind, 'ordered');
  assert.equal(one.basisWhy, 'last-cancelled');
  assert.equal(one.basis.orderName, '#12');
  assert.equal(one.basis.totalCents, 2300);
  assert.equal(one.amountCents, 1077);
  const two = r.get(cid(2));
  assert.equal(two.basisWhy, 'last-zero');
  assert.equal(two.basis.totalCents, 20000);
  assert.equal(two.amountCents, 1977);
  for (const n of [3, 4]) {
    assert.equal(r.get(cid(n)).kind, 'never', `customer ${n}`);
    assert.equal(r.get(cid(n)).neverReason, 'only-cancelled-or-zero');
  }
  assert.equal(byId(s.notSelected).get(cid(5)).primaryCode, 'recent-order', 'a valid order inside the window found in the history');
  assert.equal(byId(s.notSelected).get(cid(5)).primaryText, '7. 近 3 个月有下单：2026-07-15');
  assert.equal(r.get(cid(6)).basis.totalCents, 12000, 'a test order is skipped like a cancelled one');
  assert.equal(r.get(cid(6)).basisWhy, 'last-test', 'a test (not cancelled) last order is not reported as cancelled');
  assert.deepEqual(s.stats.basisFromEarlier, { lastCancelled: 1, lastZero: 1, lastTest: 1 });
  assert.deepEqual(s.stats.neverReasons, { noOrders: 0, onlyCancelledOrZero: 2 });
  // Average uses the orders actually used as the basis: (23 + 200 + 120 + 100) / 4.
  assert.equal(s.averageCents, 11075);
});

test('amounts: a last order that is both cancelled and a test counts as cancelled', () => {
  const nodes = [makeCustomer({ n: 1, lastOrder: paid(11, '2026-03-01T12:00:00Z', '40.00', { cancelled: true, test: true }) })];
  const history = { [cid(1)]: [paid(11, '2026-03-01T12:00:00Z', '40.00', { cancelled: true, test: true }), paid(12, '2026-02-01T12:00:00Z', '60.00')] };
  const s = select(nodes, { history });
  assert.equal(s.recipients[0].basisWhy, 'last-cancelled');
  assert.deepEqual(s.stats.basisFromEarlier, { lastCancelled: 1, lastZero: 0, lastTest: 0 });
});

test('amounts: the never-ordered average is whole cents before the percentage, so the shown formula can be redone', () => {
  // $107.74 and $107.75 average $107.745: rounded to $107.75 first, 10% is $10.775 → $10.78 → tier $15.33.
  const nodes = [
    makeCustomer({ n: 1, lastOrder: paid(1, '2026-03-01T12:00:00Z', '107.74') }),
    makeCustomer({ n: 2, lastOrder: paid(2, '2026-03-02T12:00:00Z', '107.75') }),
    makeCustomer({ n: 3 }),
  ];
  const s = select(nodes);
  assert.equal(s.averageCents, 10775);
  assert.equal(s.stats.averageCents, 10775);
  const never = byId(s.recipients).get(cid(3));
  assert.equal(never.rawCents, percentOf(s.averageCents, 10));
  assert.equal(never.rawCents, 1078);
  assert.equal(never.amountCents, 1533);
  assert.equal(never.tier, 1);
  assert.equal(never.formula, '10% × 平均 $107.75 = $10.78 → 档位 $15.33');
  assert.equal(s.neverAmountCents, 1533);
  // A third order makes the average a repeating fraction: still whole cents.
  const third = select([...nodes, makeCustomer({ n: 4, lastOrder: paid(4, '2026-03-03T12:00:00Z', '10.00') })]);
  assert.equal(third.averageCents, Math.round((10774 + 10775 + 1000) / 3));
  assert.ok(Number.isInteger(third.averageCents));
});

test('amounts: missing order history is reported instead of guessed', () => {
  const nodes = [
    makeCustomer({ n: 1, lastOrder: paid(1, '2026-02-01T12:00:00Z', '80.00', { cancelled: true }) }),
    makeCustomer({ n: 2, lastOrder: paid(2, '2026-02-01T12:00:00Z', '0.00') }),
    makeCustomer({ n: 3, lastOrder: paid(3, '2026-02-01T12:00:00Z', '80.00') }),
    makeCustomer({ n: 4, lastOrder: paid(4, '2026-02-01T12:00:00Z', '0.00'), marketingState: 'UNSUBSCRIBED' }), // excluded anyway: no lookup needed
  ];
  assert.throws(() => select(nodes), (err) => /order history missing for 2 customer/.test(err.message) && err.missingHistory.join() === [cid(1), cid(2)].join());
  // Raw nodes or parsed orders, as a Map or a plain object.
  const history = new Map([[cid(1), { orders: [paid(1, '2026-02-01T12:00:00Z', '80.00', { cancelled: true })] }], [cid(2), []]]);
  const s = select(nodes, { history });
  assert.equal(s.recipients.length, 3);
  assert.equal(normalizeHistory(null).size, 0);
  assert.equal(normalizeHistory({ a: [paid(1, '2026-01-01T00:00:00Z', '5.00')] }).get('a').orders[0].totalCents, 500);
});

test('dedupe: one account per address, keeper by last paid order, orders, account age, id', () => {
  const same = (n, address1, extra = {}) => makeCustomer({ n, address: ADDR(address1), ...extra });
  const nodes = [
    // Group 1: most recent paid order wins.
    same(1, '1 Pine St', { lastOrder: paid(1, '2026-04-01T12:00:00Z', '50.00') }),
    same(2, '1 Pine Street', { lastOrder: paid(2, '2026-05-01T12:00:00Z', '20.00') }),
    same(3, '1 pine st.'),
    // Group 2: same order date → more orders wins.
    same(4, '2 Oak Ave', { lastOrder: paid(4, '2026-04-01T12:00:00Z', '50.00'), numberOfOrders: 2 }),
    same(5, '2 Oak Avenue', { lastOrder: paid(5, '2026-04-01T12:00:00Z', '50.00'), numberOfOrders: 5 }),
    // Group 3: never ordered → older account wins.
    same(6, '3 Elm Rd', { createdAt: '2024-05-01T00:00:00Z' }),
    same(7, '3 Elm Road', { createdAt: '2023-05-01T00:00:00Z' }),
    // Group 4: all equal → smaller id wins.
    same(9, '4 Ash Ln'),
    same(8, '4 Ash Lane'),
    // No comparable address: never grouped.
    makeCustomer({ n: 10, address: null }),
    makeCustomer({ n: 11, address: null }),
  ];
  const s = select(nodes);
  assert.deepEqual(recipientIds(s).sort(), [cid(2), cid(5), cid(7), cid(8), cid(10), cid(11)].sort());
  const ns = byId(s.notSelected);
  for (const [loser, keeper] of [[1, 2], [3, 2], [4, 5], [6, 7], [9, 8]]) {
    assert.equal(ns.get(cid(loser)).primaryCode, 'duplicate-address', `customer ${loser}`);
    assert.equal(ns.get(cid(loser)).relatedCustomerId, cid(keeper), `customer ${loser} lost to ${keeper}`);
    assert.match(ns.get(cid(loser)).groupId, /^G\d{4}$/);
  }
  assert.equal(s.duplicates.length, 4);
  const g1 = s.duplicates.find((d) => d.keptCustomerId === cid(2));
  assert.equal(g1.size, 3);
  assert.equal(g1.address, '1 pine st · 78701 · US');
  assert.equal(g1.flaggedBulk, false);
  assert.deepEqual(g1.members.map((m) => [m.numericId, m.kept]), [['2', true], ['1', false], ['3', false]]);
  const kept = byId(s.recipients).get(cid(2));
  assert.equal(kept.groupSize, 3);
  assert.deepEqual(kept.groupOthers.sort(), ['1', '3']);
  assert.equal(byId(s.recipients).get(cid(10)).groupId, null);
  assert.deepEqual(s.stats.duplicates, { groups: 4, accounts: 9, removed: 5, sizeHistogram: { 2: 3, 3: 1 }, largest: [3, 2, 2, 2], flaggedBulk: 0 });
  assert.equal(s.stats.noAddressRecipients, 2);
  // The kept account's amount is its own; the average only counts kept people.
  assert.equal(kept.amountCents, 1077);
  assert.equal(s.averageCents, (2000 + 5000) / 2);
});

test('dedupe: ten or more accounts at one address are flagged as a likely bulk registration', () => {
  const nodes = Array.from({ length: 10 }, (_, i) => makeCustomer({ n: i + 1, address: ADDR('77 Bulk Blvd') }));
  const s = select([...nodes, makeCustomer({ n: 50, lastOrder: paid(50, '2026-05-01T12:00:00Z', '90.00') })]);
  assert.equal(s.duplicates.length, 1);
  assert.equal(s.duplicates[0].flaggedBulk, true);
  assert.equal(s.duplicates[0].size, 10);
  assert.equal(s.stats.duplicates.flaggedBulk, 1);
  assert.equal(s.recipients.length, 2, 'still one card for the address');
});

test('amounts: with nobody who ordered there is no average, so never-ordered people are held back', () => {
  const s = select([makeCustomer({ n: 1 }), makeCustomer({ n: 2 })]);
  assert.equal(s.recipients.length, 0);
  assert.equal(s.averageCents, null);
  assert.deepEqual(s.notSelected.map((x) => x.primaryCode), ['zero-amount', 'zero-amount']);
  assert.match(s.notSelected[0].primaryText, /^11\. 算出的金额为 \$0：没有可用来算平均值的有下单客户$/);
});

test('order: ordered customers by last order (newest first), then never-ordered by sign-up (newest first)', () => {
  const nodes = [
    makeCustomer({ n: 1, createdAt: '2024-01-01T00:00:00Z' }),
    makeCustomer({ n: 2, lastOrder: paid(2, '2026-01-01T12:00:00Z', '50.00') }),
    makeCustomer({ n: 3, createdAt: '2025-06-01T00:00:00Z' }),
    makeCustomer({ n: 4, lastOrder: paid(4, '2026-06-01T12:00:00Z', '50.00') }),
    makeCustomer({ n: 5, lastOrder: paid(5, '2026-06-01T12:00:00Z', '50.00') }),
  ];
  const s = select(nodes);
  assert.deepEqual(s.recipients.map((r) => [r.seq, r.numericId]), [[1, '4'], [2, '5'], [3, '2'], [4, '3'], [5, '1']]);
});

test('selection: parameters and recipient fields are frozen into the file', () => {
  const s = select([makeCustomer({ n: 1, first: 'Ann', last: 'Lee', lastOrder: paid(1, '2026-06-01T12:00:00Z', '50.00', { source: 'web' }), amountSpent: '321.09', numberOfOrders: 4 })]);
  assert.equal(s.version, 1);
  assert.equal(s.mode, 'live');
  assert.equal(s.campaignId, '2026-10');
  assert.deepEqual(s.snapshot, { exportedAt: NOW_ISO, source: 'bulk', count: 1 });
  assert.deepEqual(s.params, selectionParams(config, { timezone: 'America/Los_Angeles', cutoffMs: CUTOFF_MS, cutoffIso: '2026-07-01T00:00:00-07:00' }));
  assert.equal(s.params.cutoffDate, '2026-07-01');
  assert.equal(s.params.sentTag, 'OCT26RTPROMO');
  assert.deepEqual(s.params.giftTiersCents, [1077, 1533, 1977]);
  assert.equal(s.params.giftCardExpiresOn, '2026-10-19');
  assert.equal(s.params.testGiftAmountCents, null);
  const r = s.recipients[0];
  assert.equal(r.name, 'Ann Lee');
  assert.equal(r.firstName, 'Ann');
  assert.equal(r.email, 'c1@example.org');
  assert.equal(r.marketingState, 'SUBSCRIBED');
  assert.equal(r.lastOrderSource, 'web');
  assert.equal(r.amountSpentCents, 32109);
  assert.equal(r.numberOfOrders, 4);
  assert.equal(r.city, 'Austin');
  assert.equal(r.zip, '70001');
  assert.deepEqual(r.basis, { orderId: 'gid://shopify/Order/1', orderName: '#1', createdAt: '2026-06-01T12:00:00Z', totalCents: 5000, sourceName: 'web' });
  assert.deepEqual(s.stats.recipientChannels, { 网店: 1 });
});

test('test campaigns: exactly the listed people, a fixed small amount, no audience rules', () => {
  const t = testConfig({ CAMPAIGN_ID: '2026-10-test', SENT_TAG: 'OCT26RTPROMO-TEST', TEST_CUSTOMER_IDS: '1,2,3,4', TEST_GIFT_AMOUNT: '0.10' });
  try {
    const customers = [
      makeCustomer({ n: 1, marketingState: 'UNSUBSCRIBED', tags: ['WHS'] }), // rules do not apply
      makeCustomer({ n: 2, email: null }),
      makeCustomer({ n: 3, tags: ['oct26rtpromo-test'] }),
      makeCustomer({ n: 4, createdAt: new Date(NOW_MS - DAY).toISOString() }),
    ].map(parseCustomer);
    const s = buildTestSelection({ config: t.config, customers, timezone: 'America/Los_Angeles', createdAt: NOW_ISO, snapshot: { exportedAt: NOW_ISO, source: 'nodes', count: 4 } });
    assert.equal(s.mode, 'test');
    assert.deepEqual(s.recipients.map((r) => [r.seq, r.numericId, r.kind, r.amountCents, r.tier]), [[1, '1', 'test', 10, null], [2, '4', 'test', 10, null]]);
    assert.equal(s.recipients[0].formula, '测试固定金额 $0.10');
    assert.equal(s.params.testGiftAmountCents, 10);
    assert.deepEqual(s.funnel.byRule.filter((r) => r.count).map((r) => [r.code, r.count]), [['no-email', 1], ['already-sent', 1]]);
    assert.deepEqual(s.notSelected.map((x) => x.numericId), ['3']);
  } finally {
    t.cleanup();
  }
});

test('selection: 最近订单渠道 is the channel of the most recent valid order, not of a cancelled one', () => {
  const nodes = [makeCustomer({ n: 1, lastOrder: paid(11, '2026-03-01T12:00:00Z', '80.00', { cancelled: true, source: 'amazon' }) })];
  const history = { [cid(1)]: [paid(11, '2026-03-01T12:00:00Z', '80.00', { cancelled: true, source: 'amazon' }), paid(12, '2026-02-01T12:00:00Z', '60.00', { source: 'web' })] };
  const s = select(nodes, { history });
  assert.equal(s.recipients.length, 1, 'a cancelled Amazon order does not exclude anyone');
  assert.equal(s.recipients[0].lastOrderSource, 'web');
  assert.deepEqual(s.stats.recipientChannels, { 网店: 1 });
});

test('selection: without any valid order 最近订单渠道 is empty, never the channel of a cancelled marketplace order', () => {
  const amazonCancelled = paid(11, '2026-04-01T12:00:00Z', '80.00', { cancelled: true, source: 'amazon' });
  const sellbriteCancelled = paid(21, '2026-04-02T12:00:00Z', '60.00', { cancelled: true, source: '205641' });
  const ebayTest = paid(31, '2026-04-03T12:00:00Z', '60.00', { test: true, source: 'ebay' });
  const zeroWeb = paid(41, '2026-04-04T12:00:00Z', '0.00', { source: 'web' });
  const nodes = [
    makeCustomer({ n: 1, lastOrder: amazonCancelled }), // only a cancelled Amazon order: rule 9 lets them through
    makeCustomer({ n: 2, lastOrder: sellbriteCancelled }), // only a cancelled Sellbrite order
    makeCustomer({ n: 3, lastOrder: ebayTest }), // only a test order
    makeCustomer({ n: 4, lastOrder: zeroWeb }), // a $0 web order is a valid order: its channel counts
    makeCustomer({ n: 5 }), // never ordered
    makeCustomer({ n: 6, lastOrder: paid(61, '2026-03-01T12:00:00Z', '100.00', { source: 'pos' }) }),
  ];
  const history = { [cid(1)]: [amazonCancelled], [cid(2)]: [sellbriteCancelled], [cid(3)]: [ebayTest], [cid(4)]: [zeroWeb] };
  const s = select(nodes, { history });
  assert.deepEqual(recipientIds(s).sort(), [1, 2, 3, 4, 5, 6].map(cid).sort(), 'cancelled and test marketplace orders exclude nobody');
  const r = byId(s.recipients);
  assert.deepEqual([1, 2, 3, 4, 5, 6].map((n) => r.get(cid(n)).lastOrderSource), ['', '', '', 'web', '', 'pos']);
  assert.deepEqual([1, 2, 3, 4, 5].map((n) => r.get(cid(n)).kind), ['never', 'never', 'never', 'never', 'never']);
  assert.deepEqual(s.stats.recipientChannels, { 没有有效订单: 4, 网店: 1, POS: 1 });
  assert.ok(!s.recipients.some((x) => ['amazon', '205641', 'ebay'].includes(x.lastOrderSource)), 'no marketplace channel on the final list');
});

test('test campaigns: 最近订单渠道 is the snapshot lastOrder\'s channel only when that order is valid', () => {
  const t = testConfig({ CAMPAIGN_ID: '2026-10-test', SENT_TAG: 'OCT26RTPROMO-TEST', TEST_CUSTOMER_IDS: '1,2,3,4', TEST_GIFT_AMOUNT: '0.10' });
  try {
    const customers = [
      makeCustomer({ n: 1, lastOrder: paid(11, '2026-09-01T12:00:00Z', '30.00', { source: 'web' }) }),
      makeCustomer({ n: 2, lastOrder: paid(21, '2026-09-01T12:00:00Z', '30.00', { cancelled: true, source: 'amazon' }) }),
      makeCustomer({ n: 3, lastOrder: paid(31, '2026-09-01T12:00:00Z', '0.00', { source: 'pos' }) }),
      makeCustomer({ n: 4 }),
    ].map(parseCustomer);
    const s = buildTestSelection({ config: t.config, customers, timezone: 'America/Los_Angeles', createdAt: NOW_ISO, snapshot: { exportedAt: NOW_ISO, source: 'nodes', count: 4 } });
    assert.deepEqual(s.recipients.map((x) => [x.numericId, x.lastOrderSource]), [['1', 'web'], ['2', ''], ['3', 'pos'], ['4', '']]);
    assert.deepEqual(s.stats.recipientChannels, { 网店: 1, 没有有效订单: 2, POS: 1 });
  } finally {
    t.cleanup();
  }
});

test('reasons: the dates of rules 4 and 7 are the store\'s calendar date (America/Los_Angeles), like the date columns', () => {
  const evening = '2026-09-01T05:30:00Z'; // 2026-08-31 22:30 in Los Angeles
  const active = makeCustomer({ n: 1, lastOrder: paid(11, evening, '50.00') });
  const fresh = makeCustomer({ n: 2, createdAt: '2026-09-28T04:00:00Z', tags: ['WHS'] }); // 2026-09-27 21:00 in Los Angeles
  const fromHistory = makeCustomer({ n: 3, lastOrder: paid(31, '2026-09-10T06:00:00Z', '30.00', { cancelled: true }) });
  const history = { [cid(3)]: [paid(31, '2026-09-10T06:00:00Z', '30.00', { cancelled: true }), paid(32, '2026-08-02T03:15:00Z', '40.00')] }; // 08-01 20:15 LA
  const s = select([active, fresh, fromHistory, makeCustomer({ n: 4, lastOrder: paid(41, '2026-03-01T12:00:00Z', '90.00') })], {
    activeOrders: [makeActiveOrder({ n: 11, customer: active, createdAt: evening })],
    history,
  });
  const ns = byId(s.notSelected);
  assert.equal(ns.get(cid(1)).primaryText, '7. 近 3 个月有下单：2026-08-31');
  assert.equal(ns.get(cid(2)).primaryText, '4. 注册不满 7 天：2026-09-27');
  assert.equal(ns.get(cid(2)).allReasons, '4. 注册不满 7 天：2026-09-27；5. 带排除 tag：WHS');
  assert.equal(ns.get(cid(3)).primaryText, '7. 近 3 个月有下单：2026-08-01', 'a recent order found in the follow-up lookup');
  // Same instants as the store-local dates the Excel shows for 上次下单日期 / 注册日期.
  assert.equal(localDate(Date.parse(evening), 'America/Los_Angeles'), '2026-08-31');

  // A daytime order is the same date in both calendars; no time zone keeps the old UTC date.
  assert.equal(reasonText({ code: 'recent-order', detail: '2026-08-15T12:00:00Z' }, config, 'America/Los_Angeles'), '7. 近 3 个月有下单：2026-08-15');
  assert.equal(reasonText({ code: 'recent-order', detail: evening }, config), '7. 近 3 个月有下单：2026-09-01');
  assert.equal(reasonText({ code: 'too-new', detail: '2026-09-28T04:00:00Z' }, config, 'America/Los_Angeles'), '4. 注册不满 7 天：2026-09-27');
  assert.equal(reasonText({ code: 'too-new', detail: '2026-09-28T04:00:00Z' }, config, 'Asia/Tokyo'), '4. 注册不满 7 天：2026-09-28');
  // Other details are untouched by the time zone.
  assert.equal(reasonText({ code: 'excluded-tag', detail: '2026-09-01T05:30:00Z' }, config, 'America/Los_Angeles'), '5. 带排除 tag：2026-09-01T05:30:00Z');
  assert.equal(reasonText({ code: 'recent-order', detail: '' }, config, 'America/Los_Angeles'), '7. 近 3 个月有下单');
});

test('reasons: summarize without a time zone (older callers) keeps the UTC date of the ISO time', () => {
  const c = parseCustomer(makeCustomer({ n: 1, lastOrder: paid(11, '2026-09-01T05:30:00Z', '50.00') }));
  const excluded = [{ customer: c, reasons: [{ code: 'recent-order', detail: '2026-09-01T05:30:00Z' }] }];
  const base = { total: 1, recipients: [], excluded, duplicates: [], averageCents: null, neverAmountCents: null, config };
  assert.equal(summarize(base).notSelected[0].primaryText, '7. 近 3 个月有下单：2026-09-01');
  assert.equal(summarize({ ...base, timezone: 'America/Los_Angeles' }).notSelected[0].primaryText, '7. 近 3 个月有下单：2026-08-31');
  assert.equal(summarize({ ...base, timezone: 'America/Los_Angeles' }).notSelected[0].allReasons, '7. 近 3 个月有下单：2026-08-31');
});

test('reasonDay: the same calendar date as time.js localDate, across both DST changes and in other zones', () => {
  const zones = ['America/Los_Angeles', 'America/New_York', 'UTC', 'Asia/Kolkata', 'Pacific/Chatham', 'Pacific/Kiritimati'];
  const from = Date.parse('2026-03-07T00:00:00Z'); // US DST starts 03-08, ends 11-01
  const to = Date.parse('2026-11-03T00:00:00Z');
  const same = (ms, tz) => {
    const iso = new Date(ms).toISOString();
    assert.equal(reasonDay(iso, tz), localDate(ms, tz), `${iso} in ${tz}`);
  };
  let checked = 0;
  for (const tz of zones) {
    // 131-minute steps drift through every minute of the day over the months.
    for (let ms = from; ms < to; ms += 131 * 60_000) {
      same(ms, tz);
      checked += 1;
    }
  }
  // Every minute of the six hours around each Los Angeles DST change (local midnight is nearby).
  for (const change of ['2026-03-08T10:00:00Z', '2026-11-01T09:00:00Z']) {
    for (let m = -180; m <= 180; m += 1) {
      same(Date.parse(change) + m * 60_000, 'America/Los_Angeles');
      checked += 1;
    }
  }
  assert.ok(checked > 15_000);
  // Offsets in the ISO text are honoured; a value that is not a time, or a bare date, is kept as before.
  assert.equal(reasonDay('2026-09-01T00:30:00+02:00', 'America/Los_Angeles'), '2026-08-31');
  assert.equal(reasonDay('2026-09-01', 'America/Los_Angeles'), '2026-09-01');
  assert.equal(reasonDay('not a time', 'America/Los_Angeles'), 'not a time');
  assert.equal(reasonDay('2026-09-01T05:30:00Z', ''), '2026-09-01');
});

test('labels: channels and reason texts', () => {
  assert.equal(channelLabel('web'), '网店');
  assert.equal(channelLabel('1456995'), 'CedCommerce Walmart Connector');
  assert.equal(channelLabel('1775805'), 'eBay');
  assert.equal(channelLabel('2329312'), 'Facebook & Instagram');
  assert.equal(channelLabel('SomethingNew'), 'SomethingNew');
  assert.equal(channelLabel(''), '');
  assert.equal(reasonText({ code: 'recent-order', detail: '2026-08-15T12:00:00Z' }, config), '7. 近 3 个月有下单：2026-08-15');
  assert.equal(reasonText({ code: 'duplicate-address', detail: cid(42) }, config), '10. 同地址重复，保留了另一个账户：保留 42');
  assert.equal(reasonText({ code: 'active-address', detail: cid(43) }, config), '8. 同地址账户近 3 个月有下单：下单账户 43');
  assert.equal(reasonText({ code: 'no-email' }, config), '1. 没有邮箱或邮箱格式无效');
  assert.equal(reasonText({ code: 'customer-deleted' }, config), '12. 补查订单时客户已被删除');
  assert.deepEqual(RULES.map((r) => r.n), [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]);
  assert.deepEqual(RULES.at(-1), { n: 12, code: 'customer-deleted', label: '补查订单时客户已被删除' });
});
