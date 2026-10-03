// Assembles the frozen selection (selection.json) from the pure rules.
// Used by the select command and by test fixtures, so both produce the same shape.

import { parseOrder, buildActivity, screen, finalize, summarize, buildTestRecipients } from './rules.js';
import { localDate } from '../time.js';

export const SELECTION_VERSION = 1;

export function selectionParams(config, { timezone, cutoffMs, cutoffIso }) {
  return {
    timezone,
    cutoffIso,
    cutoffMs,
    cutoffDate: cutoffMs ? localDate(cutoffMs, timezone) : null,
    inactiveMonths: config.inactiveMonths,
    minAccountAgeDays: config.minAccountAgeDays,
    requireEmailSubscribed: config.requireEmailSubscribed,
    excludeTags: [...config.excludeTags],
    excludeEmailDomains: [...config.excludeEmailDomains],
    excludeOrderSources: [...config.excludeOrderSources],
    giftPercent: config.giftPercent,
    giftTiersCents: [...config.giftTiersCents],
    currency: config.giftCardCurrency,
    sentTag: config.sentTag,
    giftCardNote: config.giftCardNote,
    giftCardTemplateSuffix: config.giftCardTemplateSuffix,
    giftCardExpiresOn: config.giftCardExpiresOn,
    launchDate: config.launchDate,
    remind1Date: config.remind1Date,
    remind2Date: config.remind2Date,
    testGiftAmountCents: config.testCustomerIds.length ? config.testGiftAmountCents : null,
  };
}

/**
 * Pure end-to-end selection for a live campaign.
 * @param {object} a
 * @param {object} a.config
 * @param {object[]} a.customers parsed customers (rules.parseCustomer)
 * @param {object[]} a.activeOrders ORDERS_SINCE nodes since the cutoff
 * @param {Map<string, {orders: object[], deleted?: boolean}>|object} a.history raw or parsed newest-first
 *        orders per customer id (see normalizeHistory); must cover every id that screen() reports in needsHistory
 * @param {number} a.nowMs time the data was taken (snapshot time)
 * @param {number} a.cutoffMs start of the "ordered recently" window
 * @param {string} a.cutoffIso same instant with the store's offset
 * @param {string} a.timezone
 * @param {string} a.createdAt ISO time the selection is made
 * @param {object} a.snapshot { exportedAt, source, count }
 */
export function buildSelection({ config, customers, activeOrders, history, nowMs, cutoffMs, cutoffIso, timezone, createdAt, snapshot }) {
  const ctx = { config, nowMs, cutoffMs, activity: buildActivity(activeOrders) };
  const screened = screen(customers, ctx);
  const historyMap = normalizeHistory(history);
  const missing = screened.needsHistory.filter((id) => !historyMap.has(id));
  if (missing.length) {
    const err = new Error(`order history missing for ${missing.length} customer(s)`);
    err.missingHistory = missing;
    throw err;
  }
  const result = finalize(screened, historyMap, ctx);
  const summary = summarize({ total: customers.length, ...result, config, timezone });
  return {
    version: SELECTION_VERSION,
    campaignId: config.campaignId,
    mode: 'live',
    createdAt,
    snapshot,
    params: selectionParams(config, { timezone, cutoffMs, cutoffIso }),
    averageCents: result.averageCents,
    neverAmountCents: result.neverAmountCents,
    recipients: result.recipients,
    notSelected: summary.notSelected,
    duplicates: result.duplicates,
    funnel: summary.funnel,
    stats: summary.stats,
  };
}

/** A test campaign: exactly the given customers, fixed amount, no audience rules. */
export function buildTestSelection({ config, customers, timezone, createdAt, snapshot }) {
  const { recipients, excluded } = buildTestRecipients(customers, config);
  const summary = summarize({ total: customers.length, recipients, excluded, duplicates: [], averageCents: null, neverAmountCents: null, config, timezone });
  return {
    version: SELECTION_VERSION,
    campaignId: config.campaignId,
    mode: 'test',
    createdAt,
    snapshot,
    params: selectionParams(config, { timezone, cutoffMs: null, cutoffIso: null }),
    averageCents: null,
    neverAmountCents: null,
    recipients,
    notSelected: summary.notSelected,
    duplicates: [],
    funnel: summary.funnel,
    stats: summary.stats,
  };
}

/**
 * Accepts a Map or a plain object; values are order lists or { orders, deleted? }
 * (fetchOrderHistory / order-history.json), holding raw Order nodes or parsed
 * orders. Returns Map<id, { orders: parsed[], deleted: boolean }>; `deleted`
 * means Shopify no longer had the customer at the follow-up lookup.
 */
export function normalizeHistory(history) {
  const map = new Map();
  if (!history) return map;
  const entries = history instanceof Map ? history.entries() : Object.entries(history);
  for (const [id, value] of entries) {
    const orders = (value?.orders ?? value ?? []).map((o) => (o && 'createdAtMs' in o ? o : parseOrder(o)));
    map.set(id, { orders, deleted: value?.deleted === true });
  }
  return map;
}

/** Ids of the selection's recipients, in issue order. */
export function recipientIds(selection) {
  return selection.recipients.map((r) => r.customerId);
}
