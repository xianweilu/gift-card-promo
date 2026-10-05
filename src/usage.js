// The `usage` command: a read-only daily report of how the campaign's gift
// cards are being used.
//
//   1. List every campaign card (note carries [campaign:<id>]) with its initial
//      value and current balance: balance < initial means the card was used.
//   2. List the orders paid with a gift card since one hour before the list was
//      made (no campaign card can be older; see campaignStart). Each gift-card
//      transaction's receiptJson names the card (gift_card_id), which ties the
//      payment to a campaign card. SALE/CAPTURE take money from the card,
//      REFUND puts it back.
//   3. Write usage.json (latest) plus usage/<YYYY-MM-DD>.json (the store-local
//      date of the run) for the daily trend, print a short Chinese summary and
//      regenerate the Excel (src/report/excel.js reads usage.json for the
//      使用报告 / 使用明细 sheets).
//
// Nothing is written to Shopify. Rates in usage.json are fractions (0.25 = 25%),
// rounded to 4 decimal places; money is integer cents.
//
// Two views of the money, on purpose:
//   - the cards' ledger (payments[].netCents, orders[].campaignCardCents) is net
//     of refunds, like the card balances it is checked against;
//   - the "其中礼品卡抵扣 / 顾客另外支付" split of the order totals
//     (summary.giftCardCents / customerPaidCents) is taken at checkout, like the
//     order totals themselves (totalPriceSet: before any refund). See summarize().

import path from 'node:path';
import {
  campaignPaths,
  readJson,
  writeJsonAtomic,
  acquireRunLock,
  runningCommand,
  LockError,
  appendJournal,
  readJournal,
  foldJournal,
  newRunId,
} from './campaign.js';
import { connect } from './connect.js';
import { gql } from './shopify.js';
import { GIFT_CARD_ORDERS } from './queries.js';
import { findCampaignCards } from './giftcards.js';
import { localDate, localDateTime } from './time.js';
import { formatUsd, tierLabel } from './select/amount.js';

export const USAGE_VERSION = 1;
export const NO_RECEIPT_ID_REASON = '回执里没有礼品卡 ID';

const HOUR_MS = 3_600_000;
const DAY_MS = 86_400_000;
const GIFT_CARD_GID = 'gid://shopify/GiftCard/';
// Gift-card transactions that move money. AUTHORIZATION (always followed by a
// CAPTURE), VOID and failed attempts move nothing and are ignored.
const MONEY_KINDS = new Set(['SALE', 'CAPTURE', 'REFUND']);
// The ones that take money from the card at checkout (a SALE is authorize + capture in one).
const CHARGE_KINDS = new Set(['SALE', 'CAPTURE']);
const KIND_ORDER = ['ordered', 'never', 'test'];
const TOP_PRODUCTS = 10;
// Sizes of the nested lists in GIFT_CARD_ORDERS (src/queries.js). An order that
// reaches one of them may have more entries than Shopify returned.
const LINE_ITEMS_LIMIT = 30;
const TRANSACTIONS_LIMIT = 20;
const MAX_EXAMPLES = 5;

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

/** ISO-8601 UTC without milliseconds, the form Shopify's search syntax documents. */
function searchIso(ms) {
  return new Date(ms).toISOString().replace(/\.\d{3}Z$/, 'Z');
}

function timeOf(iso) {
  const ms = Date.parse(iso);
  return Number.isFinite(ms) ? ms : 0;
}

/** Shopify decimal string → integer cents; a missing or malformed amount counts as 0. */
function moneyCents(amount) {
  const n = Number(amount);
  return amount === null || amount === undefined || !Number.isFinite(n) ? 0 : Math.round(n * 100);
}

function sum(list, pick) {
  return list.reduce((total, x) => total + pick(x), 0);
}

/** part / whole as a fraction rounded to 4 decimals (0 when whole is 0). */
function ratio(part, whole) {
  return whole ? Math.round((part / whole) * 10_000) / 10_000 : 0;
}

/** Numeric order of Shopify gids, so ".../GiftCard/9" sorts before ".../GiftCard/10". */
function compareGid(a, b) {
  const x = String(a ?? '').split('/').pop();
  const y = String(b ?? '').split('/').pop();
  if (/^\d+$/.test(x) && /^\d+$/.test(y)) {
    const bx = BigInt(x);
    const by = BigInt(y);
    return bx < by ? -1 : bx > by ? 1 : 0;
  }
  return x < y ? -1 : x > y ? 1 : 0;
}

function count(n) {
  return Number(n).toLocaleString('en-US');
}

function percentText(part, whole) {
  return `${(whole ? (part / whole) * 100 : 0).toFixed(1)}%`;
}

function zoneLabel(timezone) {
  return timezone === 'America/Los_Angeles' ? '洛杉矶时间' : `${timezone} 时间`;
}

/** "#1、#2、#3" (at most MAX_EXAMPLES, then "等") for warnings. */
function examples(names) {
  const shown = names.slice(0, MAX_EXAMPLES).join('、');
  return names.length > MAX_EXAMPLES ? `${shown} 等` : shown;
}

/** Every calendar date from `from` to `to` inclusive ('YYYY-MM-DD'). */
function calendarDays(from, to) {
  const [y, m, d] = from.split('-').map(Number);
  const days = [];
  for (let ms = Date.UTC(y, m - 1, d); ; ms += DAY_MS) {
    const date = new Date(ms).toISOString().slice(0, 10);
    if (date > to) break;
    days.push(date);
  }
  return days;
}

// ---------------------------------------------------------------------------
// Campaign window
// ---------------------------------------------------------------------------

/** The campaign-card search window everywhere: selection.createdAt − 1 hour. */
export function cardSearchSince(selection) {
  const ms = Date.parse(selection?.createdAt);
  if (!Number.isFinite(ms)) throw new Error('selection.json 里没有有效的 createdAt');
  return searchIso(ms - HOUR_MS);
}

function earliest(times) {
  const valid = times.filter(Number.isFinite);
  return valid.length ? valid.reduce((a, b) => Math.min(a, b)) : null;
}

/**
 * The two starts of the report.
 *   campaignStartIso  start of the gift-card orders search (and usage.json's campaignStartIso):
 *                     always cardSearchSince(selection), one hour before the list was made.
 *                     No campaign card can be older than its list, so no campaign payment is
 *                     missed. The journal is not used for this on purpose: one rebuilt after a
 *                     loss (verify's reconcile.found, then a later issue batch) starts after
 *                     the first cards were created and used.
 *   startMs           first day of the daily table: the earliest of the first live `issue`
 *                     run in the journal and the oldest campaign card; the time the list was
 *                     made when there is neither yet.
 *   startFrom         where startMs came from: 'issue' | 'card' | 'selection'.
 */
export function campaignStart({ runs = [], cards = [], selection }) {
  const campaignStartIso = cardSearchSince(selection);
  const firstIssueMs = earliest(runs.filter((r) => r.command === 'issue' && !r.dryRun).map((r) => Date.parse(r.startedAt)));
  const firstCardMs = earliest(cards.map((c) => Date.parse(c.createdAt)));
  if (firstIssueMs !== null && (firstCardMs === null || firstIssueMs <= firstCardMs)) {
    return { startMs: firstIssueMs, campaignStartIso, startFrom: 'issue' };
  }
  if (firstCardMs !== null) return { startMs: firstCardMs, campaignStartIso, startFrom: 'card' };
  return { startMs: Date.parse(selection.createdAt), campaignStartIso, startFrom: 'selection' };
}

// ---------------------------------------------------------------------------
// Receipts
// ---------------------------------------------------------------------------

function toGiftCardGid(raw) {
  if (typeof raw === 'number') return Number.isSafeInteger(raw) && raw > 0 ? `${GIFT_CARD_GID}${raw}` : null;
  if (typeof raw !== 'string') return null;
  const v = raw.trim();
  if (/^[1-9]\d*$/.test(v)) return `${GIFT_CARD_GID}${v}`;
  if (/^gid:\/\/shopify\/GiftCard\/[1-9]\d*$/.test(v)) return v;
  return null;
}

/**
 * The gid ("gid://shopify/GiftCard/<n>") of the card a gift-card transaction
 * used, read from its receiptJson; null when the receipt does not name one.
 * Shopify sends the receipt as a JSON string; an already-parsed object, a
 * double-encoded string and a damaged string are tolerated. Digits are taken
 * straight from the text so a very large id cannot lose precision in JSON.parse.
 */
export function giftCardIdFromReceipt(receipt) {
  let value = receipt;
  for (let depth = 0; typeof value === 'string' && depth < 3; depth += 1) {
    const m = /"gift_card_id"\s*:\s*"?([1-9]\d*)"?(?=\s*[,}\]])/.exec(value);
    if (m) return `${GIFT_CARD_GID}${m[1]}`;
    try {
      value = JSON.parse(value);
    } catch {
      return null;
    }
  }
  if (!value || typeof value !== 'object') return null;
  return toGiftCardGid(value.gift_card_id);
}

// ---------------------------------------------------------------------------
// Fetching (read-only)
// ---------------------------------------------------------------------------

export function giftCardOrdersQuery(sinceIso) {
  return `gateway:gift_card created_at:>='${sinceIso}'`;
}

/** Every order paid (partly) with a gift card since `sinceIso`, as raw GIFT_CARD_ORDERS nodes. */
export async function fetchGiftCardOrders(sinceIso, { log } = {}) {
  const query = giftCardOrdersQuery(sinceIso);
  const nodes = [];
  const seen = new Set();
  let after = null;
  let pages = 0;
  do {
    const data = await gql(GIFT_CARD_ORDERS, { query, after });
    const connection = data?.orders;
    if (!connection) throw new Error('Shopify 的订单查询没有返回结果');
    for (const node of connection.nodes ?? []) {
      if (!node?.id || seen.has(node.id)) continue; // defensive: never count an order twice
      seen.add(node.id);
      nodes.push(node);
    }
    pages += 1;
    if (log && pages % 20 === 0) log.info(`  已读取 ${count(nodes.length)} 笔订单…`);
    const info = connection.pageInfo ?? {};
    if (info.hasNextPage && !info.endCursor) throw new Error('Shopify 说还有下一页订单，却没有给出分页游标');
    const next = info.hasNextPage ? info.endCursor : null;
    if (next && next === after) throw new Error('Shopify 的订单查询两次返回同一个分页游标，已停止以免死循环');
    after = next;
  } while (after);
  return nodes;
}

// ---------------------------------------------------------------------------
// Building usage.json (pure)
// ---------------------------------------------------------------------------

function toUsageCard(c) {
  const initialCents = c.amountCents;
  const balanceCents = c.balanceCents;
  return {
    giftCardId: c.id,
    customerId: c.customerId ?? null,
    last4: c.last4 ?? '',
    initialCents,
    balanceCents,
    usedCents: Math.max(0, initialCents - balanceCents),
    enabled: c.enabled !== false,
    expiresOn: c.expiresOn ?? null,
    createdAt: c.createdAt,
  };
}

const isUsed = (card) => card.balanceCents < card.initialCents;

function transactionsOf(node) {
  if (Array.isArray(node.transactions)) return node.transactions;
  return node.transactions?.nodes ?? [];
}

function isMoneyGiftCardTransaction(t) {
  return String(t?.gateway ?? '').toLowerCase() === 'gift_card'
    && String(t?.status ?? '').toUpperCase() === 'SUCCESS'
    && MONEY_KINDS.has(String(t?.kind ?? '').toUpperCase());
}

function lineItemsOf(node) {
  return (node.lineItems?.nodes ?? []).map((li) => ({
    name: String(li?.name ?? '').trim(),
    sku: li?.sku ? String(li.sku) : '',
    quantity: Number(li?.quantity) || 0,
    amountCents: moneyCents(li?.discountedTotalSet?.shopMoney?.amount),
  }));
}

/** Issued / used / rate / used money of a group of cards. */
function cardGroup(cards) {
  const used = cards.filter(isUsed).length;
  return { issued: cards.length, used, rate: ratio(used, cards.length), usedCents: sum(cards, (c) => c.usedCents) };
}

function reportedKinds(selection) {
  const present = new Set((selection.recipients ?? []).map((r) => r.kind));
  if (selection.mode !== 'test') {
    present.add('ordered');
    present.add('never');
  }
  return KIND_ORDER.filter((k) => present.has(k));
}

/**
 * Store-local date of each used card's first campaign payment (SALE/CAPTURE),
 * counted per date. A payment on a cancelled order is used only when the card
 * has no other one. Cards that are not used (balance back to full) are left out.
 */
function firstUsesPerDate(cards, payments, cancelledOrderIds, timezone) {
  const firstLive = new Map();
  const firstAny = new Map();
  const keepMin = (map, key, ms) => {
    if (!map.has(key) || ms < map.get(key)) map.set(key, ms);
  };
  for (const p of payments) {
    if (p.kind === 'REFUND') continue;
    const ms = Date.parse(p.processedAt);
    if (!Number.isFinite(ms)) continue;
    keepMin(firstAny, p.giftCardId, ms);
    if (!cancelledOrderIds.has(p.orderId)) keepMin(firstLive, p.giftCardId, ms);
  }
  const perDate = new Map();
  for (const c of cards) {
    if (!isUsed(c)) continue;
    const ms = firstLive.get(c.giftCardId) ?? firstAny.get(c.giftCardId);
    if (ms === undefined) continue;
    const date = localDate(ms, timezone);
    perDate.set(date, (perDate.get(date) ?? 0) + 1);
  }
  return perDate;
}

function dailyRows({ cards, payments, orders, liveOrders, timezone, campaignStartMs, nowMs }) {
  const cancelledOrderIds = new Set(orders.filter((o) => o.cancelled).map((o) => o.orderId));
  const newUsed = firstUsesPerDate(cards, payments, cancelledOrderIds, timezone);
  const perDay = new Map();
  for (const o of liveOrders) {
    const ms = Date.parse(o.createdAt);
    if (!Number.isFinite(ms)) continue;
    const date = localDate(ms, timezone);
    const d = perDay.get(date) ?? { orders: 0, cents: 0 };
    d.orders += 1;
    d.cents += o.totalCents;
    perDay.set(date, d);
  }
  // From the campaign's first day to today, widened if any activity falls outside.
  const bounds = [localDate(campaignStartMs, timezone), localDate(nowMs, timezone), ...newUsed.keys(), ...perDay.keys()].sort();
  let cumulative = 0;
  return calendarDays(bounds[0], bounds[bounds.length - 1]).map((date) => {
    const newCardsUsed = newUsed.get(date) ?? 0;
    cumulative += newCardsUsed;
    const day = perDay.get(date);
    return {
      date,
      newCardsUsed,
      orders: day?.orders ?? 0,
      ordersTotalCents: day?.cents ?? 0,
      cumulativeCardsUsed: cumulative,
      cumulativeRate: ratio(cumulative, cards.length),
    };
  });
}

function topProductsOf(liveOrders) {
  const products = new Map();
  for (const o of liveOrders) {
    for (const li of o.lineItems) {
      const name = li.name || '（无名称）';
      const p = products.get(name) ?? { name, quantity: 0, amountCents: 0 };
      p.quantity += li.quantity;
      p.amountCents += li.amountCents;
      products.set(name, p);
    }
  }
  return [...products.values()]
    .sort((a, b) => b.quantity - a.quantity || b.amountCents - a.amountCents || (a.name < b.name ? -1 : a.name > b.name ? 1 : 0))
    .slice(0, TOP_PRODUCTS);
}

function summarize({ selection, recipients, cards, payments, orders, timezone, campaignStartMs, nowMs }) {
  // Cancelled orders stay listed in orders[] but are not revenue.
  const liveOrders = orders.filter((o) => !o.cancelled);
  const liveOrderIds = new Set(liveOrders.map((o) => o.orderId));
  const used = cardGroup(cards);
  const ordersTotalCents = sum(liveOrders, (o) => o.totalCents);
  // "其中礼品卡抵扣 / 顾客另外支付" splits ordersTotalCents, which is the order totals at
  // checkout (totalPriceSet: before any refund), so the campaign-card side is taken at
  // checkout too: the SALE/CAPTURE amounts of campaign cards on the same non-cancelled
  // orders, refunds NOT subtracted. A refund that later goes back to a campaign card never
  // reached the customer, so it must not show up as "顾客另外支付"; mixing a net card amount
  // with a gross order total did exactly that. payments[] (netCents) and
  // orders[].campaignCardCents stay net of refunds: they are the cards' ledger, which is
  // checked against the balances (consistencyWarnings).
  const giftCardCents = sum(payments, (p) => (CHARGE_KINDS.has(p.kind) && liveOrderIds.has(p.orderId) ? p.amountCents : 0));
  const issuedCents = sum(cards, (c) => c.initialCents);
  const tiers = selection.params?.giftTiersCents ?? [];
  const recipientOf = (card) => recipients.get(card.customerId);
  return {
    issuedCards: cards.length,
    issuedCents,
    usedCards: used.used,
    usedCents: used.usedCents,
    usedRate: used.rate,
    usedCentsRate: ratio(used.usedCents, issuedCents),
    orders: liveOrders.length,
    ordersTotalCents,
    avgOrderCents: liveOrders.length ? Math.round(ordersTotalCents / liveOrders.length) : 0,
    giftCardCents,
    customerPaidCents: ordersTotalCents - giftCardCents,
    byTier: tiers.map((_, tier) => ({ tier, label: tierLabel(tier, tiers), ...cardGroup(cards.filter((c) => recipientOf(c)?.tier === tier)) })),
    byKind: reportedKinds(selection).map((kind) => ({ kind, ...cardGroup(cards.filter((c) => recipientOf(c)?.kind === kind)) })),
    daily: dailyRows({ cards, payments, orders, liveOrders, timezone, campaignStartMs, nowMs }),
    topProducts: topProductsOf(liveOrders),
  };
}

/** Console warnings about data that does not add up (never part of usage.json). */
function consistencyWarnings({ cards, payments, recipients, truncatedItems, truncatedTransactions }) {
  const warnings = [];
  const paidByCard = new Map();
  for (const p of payments) paidByCard.set(p.giftCardId, (paidByCard.get(p.giftCardId) ?? 0) + p.netCents);
  const mismatched = cards.filter((c) => (paidByCard.get(c.giftCardId) ?? 0) !== c.usedCents);
  if (mismatched.length) {
    const ex = mismatched[0];
    warnings.push(`${count(mismatched.length)} 张卡的已用金额和订单里的用卡付款对不上，例如尾号 ${ex.last4}：按余额已用 ${formatUsd(ex.usedCents)}，`
      + `订单付款合计 ${formatUsd(paidByCard.get(ex.giftCardId) ?? 0)}。可能是付款回执里没有礼品卡 ID，或在后台调整过余额，请到 Shopify 后台核对`);
  }
  const outside = cards.filter((c) => !recipients.has(c.customerId));
  if (outside.length) {
    warnings.push(`${count(outside.length)} 张本活动的卡不属于名单里的客户（例如 ${outside[0].giftCardId}），不计入按档位和按客户类型的统计；请运行 verify 核对`);
  }
  if (truncatedItems.length) {
    warnings.push(`${count(truncatedItems.length)} 笔订单的商品数达到单次查询上限 ${LINE_ITEMS_LIMIT} 个，商品明细可能不完整：${examples(truncatedItems)}`);
  }
  if (truncatedTransactions.length) {
    warnings.push(`${count(truncatedTransactions.length)} 笔订单的交易数达到单次查询上限 ${TRANSACTIONS_LIMIT} 笔，用卡付款可能没有取全：${examples(truncatedTransactions)}`);
  }
  return warnings;
}

/**
 * Pure: the usage.json document from what Shopify returned.
 * @param {object} a
 * @param {object} a.selection selection.json (recipients give each card's tier and kind)
 * @param {object[]} a.cards findCampaignCards() result
 * @param {object[]} a.orderNodes GIFT_CARD_ORDERS nodes
 * @param {string} a.timezone store time zone (selection.params.timezone)
 * @param {number} a.campaignStartMs first day of `daily` (campaignStart().startMs)
 * @param {string} a.campaignStartIso start of the orders search window (campaignStart().campaignStartIso)
 * @param {string} a.fetchedAt ISO time the data was read ("today" for `daily`)
 * @returns {{ usage: object, warnings: string[] }}
 */
export function buildUsage({ selection, cards, orderNodes, timezone, campaignStartMs, campaignStartIso, fetchedAt }) {
  const recipients = new Map((selection.recipients ?? []).map((r) => [r.customerId, r]));
  const usageCards = cards.map(toUsageCard).sort((a, b) => timeOf(a.createdAt) - timeOf(b.createdAt) || compareGid(a.giftCardId, b.giftCardId));
  const cardById = new Map(usageCards.map((c) => [c.giftCardId, c]));
  // Whose gift-card payment without a card id is worth a manual look.
  const campaignCustomers = new Set([...recipients.keys(), ...usageCards.map((c) => c.customerId).filter(Boolean)]);

  const payments = [];
  const orders = [];
  const unmatched = [];
  const truncatedItems = [];
  const truncatedTransactions = [];
  const nodes = [...orderNodes].sort((a, b) => timeOf(a.createdAt) - timeOf(b.createdAt) || compareGid(a.id, b.id));
  for (const node of nodes) {
    const orderCustomerId = node.customer?.id ?? null;
    const base = { orderId: node.id, orderName: node.name ?? '', orderCreatedAt: node.createdAt };
    const transactions = transactionsOf(node);
    if (transactions.length >= TRANSACTIONS_LIMIT) truncatedTransactions.push(base.orderName || base.orderId);
    const own = [];
    for (const t of transactions) {
      if (!isMoneyGiftCardTransaction(t)) continue;
      const kind = String(t.kind).toUpperCase();
      const amountCents = Math.abs(moneyCents(t.amountSet?.shopMoney?.amount));
      const processedAt = t.processedAt ?? node.createdAt;
      const giftCardId = giftCardIdFromReceipt(t.receiptJson);
      if (!giftCardId) {
        if (orderCustomerId && campaignCustomers.has(orderCustomerId)) {
          const reason = kind === 'REFUND' ? `${NO_RECEIPT_ID_REASON}（退款）` : NO_RECEIPT_ID_REASON;
          unmatched.push({ ...base, customerId: orderCustomerId, amountCents, processedAt, reason });
        }
        continue;
      }
      const card = cardById.get(giftCardId);
      if (!card) continue; // a gift card from outside this campaign (bought, or another campaign)
      own.push({
        ...base,
        orderCustomerId,
        giftCardId,
        cardCustomerId: card.customerId,
        kind,
        amountCents,
        netCents: kind === 'REFUND' ? -amountCents : amountCents,
        processedAt,
      });
    }
    if (!own.length) continue;
    const lineItems = lineItemsOf(node);
    if (lineItems.length >= LINE_ITEMS_LIMIT) truncatedItems.push(base.orderName || base.orderId);
    payments.push(...own);
    orders.push({
      orderId: node.id,
      orderName: base.orderName,
      createdAt: node.createdAt,
      customerId: orderCustomerId,
      totalCents: moneyCents(node.totalPriceSet?.shopMoney?.amount),
      campaignCardCents: sum(own, (p) => p.netCents),
      cancelled: !!node.cancelledAt,
      giftCardIds: [...new Set(own.map((p) => p.giftCardId))],
      lineItems,
    });
  }
  // A ledger in processing order (Array#sort is stable: same instant keeps order/transaction order).
  payments.sort((a, b) => timeOf(a.processedAt) - timeOf(b.processedAt));
  unmatched.sort((a, b) => timeOf(a.processedAt) - timeOf(b.processedAt));

  const summary = summarize({ selection, recipients, cards: usageCards, payments, orders, timezone, campaignStartMs, nowMs: Date.parse(fetchedAt) });
  const usage = {
    version: USAGE_VERSION,
    fetchedAt,
    timezone,
    campaignStartIso,
    cards: usageCards,
    payments,
    orders,
    unmatched,
    summary,
  };
  return { usage, warnings: consistencyWarnings({ cards: usageCards, payments, recipients, truncatedItems, truncatedTransactions }) };
}

// ---------------------------------------------------------------------------
// The command
// ---------------------------------------------------------------------------

/** The real Excel writer, loaded lazily so this module does not depend on the workbook code at import time. */
async function defaultWriteReport(options) {
  const { writeReport } = await import('./report/excel.js');
  return writeReport(options);
}

/** Cards the journal recorded as created but Shopify's campaign-card list does not show. */
function journalWarnings(state, usage) {
  const listed = new Set(usage.cards.map((c) => c.giftCardId));
  const missing = [...state.customers.values()].filter((s) => s.giftCardId && !listed.has(s.giftCardId));
  if (!missing.length) return [];
  return [`本地日志里有 ${count(missing.length)} 张卡没有出现在 Shopify 的本活动卡列表里（例如 ${missing[0].giftCardId}），这份报告没有计入它们；请运行 verify 核对`];
}

function printSummary(usage, log) {
  const s = usage.summary;
  log.info(`使用情况（截至 ${localDateTime(Date.parse(usage.fetchedAt), usage.timezone)}，${zoneLabel(usage.timezone)}）`);
  log.info(`  发出礼品卡：${count(s.issuedCards)} 张，面额 ${formatUsd(s.issuedCents)}`);
  log.info(`  已经使用：${count(s.usedCards)} 张，占 ${percentText(s.usedCards, s.issuedCards)}；已用 ${formatUsd(s.usedCents)}，占面额 ${percentText(s.usedCents, s.issuedCents)}`);
  log.info(`  带来订单：${count(s.orders)} 笔，订单总额 ${formatUsd(s.ordersTotalCents)}，平均每单 ${formatUsd(s.avgOrderCents)}`);
  log.info(`    其中礼品卡抵扣 ${formatUsd(s.giftCardCents)}，顾客另外支付 ${formatUsd(s.customerPaidCents)}`);
  const cancelled = usage.orders.filter((o) => o.cancelled).length;
  if (cancelled) log.info(`  另有 ${count(cancelled)} 笔已取消的订单用过本活动的卡：列在使用明细里，不计入订单数和金额`);
  if (usage.unmatched.length) {
    log.warn(`  ${count(usage.unmatched.length)} 笔礼品卡付款的${NO_RECEIPT_ID_REASON}，已单独列在使用明细里，请人工核对`);
  }
}

/** What run.end records for the Excel's run log. */
function runEndSummary(usage) {
  const s = usage.summary;
  return {
    issuedCards: s.issuedCards,
    usedCards: s.usedCards,
    usedCents: s.usedCents,
    orders: s.orders,
    ordersTotalCents: s.ordersTotalCents,
    giftCardCents: s.giftCardCents,
    unmatched: usage.unmatched.length,
  };
}

/**
 * Regenerate the Excel. A failure is only a warning: it never replaces the command's own result.
 * writeReport logs each of its warnings itself (result.warnings is a copy), so they are not printed again.
 */
async function refreshExcel(writeReport, { config, paths, log, now }) {
  try {
    const result = await writeReport({ config, paths, log, now });
    log.info(`Excel：${result?.file ?? paths.excel}`);
    // fileEn is null when the English edition failed: writeReport already logged that warning.
    if (result?.fileEn) log.info(`英文版 Excel：${result.fileEn}`);
  } catch (err) {
    log.warn(`Excel 没有更新：${err.message}。可以稍后运行 node index.js export 重新生成`);
  }
}

/**
 * node index.js usage — read-only daily usage report.
 * Exit codes: 0 done; 1 no selection, another command running, or Shopify/file error.
 * Returns { exitCode, usage } (usage is null unless exitCode is 0).
 */
export async function runUsage({ config, log = console, now = () => new Date(), sleep, writeReport = defaultWriteReport } = {}) {
  const paths = campaignPaths(config);

  // The frozen list gives each card's tier and customer kind.
  let selection;
  try {
    selection = readJson(paths.selection, null);
  } catch (err) {
    log.error(`无法读取名单：${err.message}`);
    return { exitCode: 1, usage: null };
  }
  if (!selection) {
    log.error(runningCommand(paths) === 'select'
      ? '还没有名单：select 正在运行，请等它跑完再运行 usage'
      : `还没有名单（${paths.selection}），请先运行 node index.js select`);
    return { exitCode: 1, usage: null };
  }
  if (selection.campaignId !== config.campaignId) {
    log.error(`名单属于活动 ${selection.campaignId}，和 CAMPAIGN_ID=${config.campaignId} 不一致，已停止`);
    return { exitCode: 1, usage: null };
  }
  const timezone = selection.params?.timezone || config.timezone;
  if (!timezone) {
    log.error('名单里没有店铺时区，请重新运行 select');
    return { exitCode: 1, usage: null };
  }

  // select / issue / remind / verify / usage never run at the same time.
  let release;
  try {
    release = acquireRunLock(paths, 'usage');
  } catch (err) {
    log.error(err instanceof LockError ? err.message : `无法取得运行锁：${err.message}`);
    return { exitCode: 1, usage: null };
  }

  const nowIso = () => now().toISOString();
  const run = newRunId(now());
  let exitCode = 1;
  let usage = null;
  let runSummary = null;
  let runStarted = false;
  try {
    const state = foldJournal(readJournal(paths.journal));
    appendJournal(paths.journal, { op: 'run.start', run, command: 'usage', dryRun: false, batch: null, limit: null, options: {} }, { now: nowIso });
    runStarted = true;

    await connect(config, { log, sleep });
    const fetchedAt = nowIso();

    log.info('正在读取本活动的礼品卡和余额…');
    const cards = await findCampaignCards({ campaignId: config.campaignId, sinceIso: cardSearchSince(selection) });
    log.info(`  本活动的卡 ${count(cards.length)} 张`);
    // Needs the cards: the daily table starts at the earlier of the first live issue run and the oldest card.
    const start = campaignStart({ runs: state.runs, cards, selection });
    const since = `${localDateTime(Date.parse(start.campaignStartIso), timezone)}（${zoneLabel(timezone)}）`;
    log.info(`正在读取 ${since}以来用礼品卡付款的订单（从名单生成前 1 小时起，本活动的卡不会比这更早）…`);
    const orderNodes = await fetchGiftCardOrders(start.campaignStartIso, { log });
    log.info(`  用礼品卡付款的订单 ${count(orderNodes.length)} 笔`);

    const built = buildUsage({
      selection,
      cards,
      orderNodes,
      timezone,
      campaignStartMs: start.startMs,
      campaignStartIso: start.campaignStartIso,
      fetchedAt,
    });
    writeJsonAtomic(paths.usage, built.usage);
    const snapshotFile = path.join(paths.usageDir, `${localDate(Date.parse(fetchedAt), timezone)}.json`);
    writeJsonAtomic(snapshotFile, built.usage);
    usage = built.usage;

    printSummary(usage, log);
    for (const w of [...built.warnings, ...journalWarnings(state, usage)]) log.warn(w);
    log.info(`已保存 ${paths.usage}，当天快照 ${snapshotFile}`);
    runSummary = runEndSummary(usage);
    exitCode = 0;
  } catch (err) {
    log.error(`usage 没有完成：${err.message}`);
    runSummary = { error: err.message };
    usage = null;
    exitCode = 1;
  } finally {
    if (runStarted) {
      try {
        appendJournal(paths.journal, { op: 'run.end', run, summary: runSummary, exitCode }, { now: nowIso });
      } catch (err) {
        log.warn(`运行结束记录没有写进日志：${err.message}`);
      }
    }
    release();
    await refreshExcel(writeReport, { config, paths, log, now });
  }
  return { exitCode, usage };
}
