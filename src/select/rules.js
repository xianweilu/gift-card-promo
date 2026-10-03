// Pure selection logic: who gets a card and how much. No I/O here; the
// fetchers in ./export.js supply the data, src/select/index.js glues it.

import { addressKey, describeKey } from './address.js';
import { toCents, giftFor, isGiftableTotal, formulaText, formatUsd } from './amount.js';

/**
 * Rules in evaluation order. `n` is the number shown to people (Excel, plan).
 * Rule 12 was added after 1-11 were published, so it keeps the next number:
 * it is checked during the order follow-up (finalize), when Shopify answers
 * that a candidate no longer exists.
 */
export const RULES = [
  { n: 1, code: 'no-email', label: '没有邮箱或邮箱格式无效' },
  { n: 2, code: 'relay-email', label: '邮箱域名在排除名单里' },
  { n: 3, code: 'not-subscribed', label: '未订阅邮件营销' },
  { n: 4, code: 'too-new', label: '注册不满 7 天' },
  { n: 5, code: 'excluded-tag', label: '带排除 tag' },
  { n: 6, code: 'already-sent', label: '已经发过（已有发放 tag）' },
  { n: 7, code: 'recent-order', label: '近 3 个月有下单' },
  { n: 8, code: 'active-address', label: '同地址账户近 3 个月有下单' },
  { n: 9, code: 'marketplace-order', label: '最近订单来自平台渠道' },
  { n: 10, code: 'duplicate-address', label: '同地址重复，保留了另一个账户' },
  { n: 11, code: 'zero-amount', label: '算出的金额为 $0' },
  { n: 12, code: 'customer-deleted', label: '补查订单时客户已被删除' },
];
export const RULE = Object.fromEntries(RULES.map((r) => [r.code, r]));

/** Rule label with the configured numbers filled in (e.g. "注册不满 7 天"). */
export function ruleLabel(code, config) {
  if (code === 'too-new') return `注册不满 ${config.minAccountAgeDays} 天`;
  if (code === 'recent-order') return `近 ${config.inactiveMonths} 个月有下单`;
  if (code === 'active-address') return `同地址账户近 ${config.inactiveMonths} 个月有下单`;
  return RULE[code].label;
}

/**
 * One reason as shown in the Excel, e.g. "5. 带排除 tag：WHS".
 * @param {{ code: string, detail?: string }} reason
 * @param {object} config
 * @param {string} [timezone] store time zone: the dates of rules 4 and 7 (ISO times from
 *   Shopify) are shown as the store's calendar date, like every other date in the Excel;
 *   without it, the UTC date (the first 10 characters) as before
 */
export function reasonText(reason, config, timezone) {
  const r = RULE[reason.code];
  const d = reason.detail;
  let suffix = '';
  if (d) {
    if (reason.code === 'marketplace-order') suffix = channelLabel(d);
    else if (reason.code === 'active-address' || reason.code === 'duplicate-address') suffix = `${reason.code === 'duplicate-address' ? '保留 ' : '下单账户 '}${String(d).split('/').pop()}`;
    else if (reason.code === 'recent-order' || reason.code === 'too-new') suffix = reasonDay(d, timezone);
    else suffix = String(d);
  }
  return `${r.n}. ${ruleLabel(reason.code, config)}${suffix ? `：${suffix}` : ''}`;
}

/** One reused formatter per time zone: creating an Intl.DateTimeFormat costs ~0.1 ms, and select formats tens of thousands of reasons. */
const DAY_FORMATS = new Map();

/**
 * "YYYY-MM-DD" of an ISO time in `timeZone` (the same date as time.js localDate). Without a
 * time zone, or for a value that is not a time, its first 10 characters. A bare
 * "YYYY-MM-DD" is already a calendar date and is kept (Date.parse would read it as UTC midnight).
 */
export function reasonDay(value, timeZone) {
  const ms = Date.parse(value);
  if (!timeZone || !Number.isFinite(ms) || /^\d{4}-\d{2}-\d{2}$/.test(String(value))) return String(value).slice(0, 10);
  let fmt = DAY_FORMATS.get(timeZone);
  if (!fmt) {
    fmt = new Intl.DateTimeFormat('en-CA', { timeZone, year: 'numeric', month: '2-digit', day: '2-digit' });
    DAY_FORMATS.set(timeZone, fmt);
  }
  const p = Object.fromEntries(fmt.formatToParts(ms).map((x) => [x.type, x.value]));
  return `${p.year}-${p.month}-${p.day}`;
}

/** Exclusions that only mean "cannot or may not be emailed": counted, not listed row by row. */
export const UNLISTED_RULES = new Set(['no-email', 'relay-email', 'not-subscribed']);

const CHANNEL_LABELS = {
  web: '网店',
  checkout_next: '新版结账',
  shopify_draft_order: '草稿订单',
  pos: 'POS',
  iphone: 'iPhone App',
  android: 'Android App',
  '580111': 'Online Store',
  '2329312': 'Facebook & Instagram',
  '205641': 'Sellbrite',
  '1456995': 'CedCommerce Walmart Connector',
  '1775805': 'eBay',
  amazon: 'Amazon',
  walmart: 'Walmart',
  ebay: 'eBay',
  etsy: 'Etsy',
  woocommerce: 'WooCommerce',
};
export function channelLabel(sourceName) {
  if (!sourceName) return '';
  return CHANNEL_LABELS[String(sourceName).toLowerCase()] ?? String(sourceName);
}

// ---------------------------------------------------------------------------
// Parsing raw Shopify nodes
// ---------------------------------------------------------------------------

export function parseOrder(o) {
  if (!o) return null;
  return {
    id: o.id,
    name: o.name ?? '',
    createdAt: o.createdAt,
    createdAtMs: Date.parse(o.createdAt),
    cancelled: !!o.cancelledAt,
    test: !!o.test,
    sourceName: o.sourceName ?? '',
    totalCents: o.totalPriceSet?.shopMoney ? toCents(o.totalPriceSet.shopMoney.amount) : 0,
    currencyCode: o.totalPriceSet?.shopMoney?.currencyCode ?? '',
  };
}

/** A customer node (bulk line, page node or nodes() result) → compact record. */
export function parseCustomer(node) {
  const email = node.defaultEmailAddress?.emailAddress ?? '';
  const a = node.defaultAddress;
  return {
    id: node.id,
    numericId: String(node.id).split('/').pop(),
    firstName: node.firstName ?? '',
    lastName: node.lastName ?? '',
    displayName: node.displayName ?? '',
    createdAt: node.createdAt,
    createdAtMs: Date.parse(node.createdAt),
    numberOfOrders: Number(node.numberOfOrders ?? 0),
    amountSpentCents: node.amountSpent ? toCents(node.amountSpent.amount) : 0,
    tags: node.tags ?? [],
    email,
    emailDomain: email.includes('@') ? email.split('@').pop().toLowerCase() : '',
    emailValid: !!node.defaultEmailAddress?.validFormat,
    marketingState: node.defaultEmailAddress?.marketingState ?? null,
    address: a ? { address1: a.address1 ?? '', address2: a.address2 ?? '', city: a.city ?? '', provinceCode: a.provinceCode ?? '', zip: a.zip ?? '', countryCodeV2: a.countryCodeV2 ?? '' } : null,
    addressKey: addressKey(a),
    lastOrder: parseOrder(node.lastOrder),
  };
}

export function isValidOrder(o) {
  return !!o && !o.cancelled && !o.test;
}

/** Instant of an ISO time for comparisons; an unparseable time sorts first. */
function timeMs(iso) {
  const ms = Date.parse(iso);
  return Number.isFinite(ms) ? ms : -Infinity;
}

function displayName(c) {
  return c.displayName || `${c.firstName} ${c.lastName}`.trim() || c.email || c.numericId;
}

function compareNumericId(a, b) {
  const x = BigInt(a);
  const y = BigInt(b);
  return x < y ? -1 : x > y ? 1 : 0;
}

// ---------------------------------------------------------------------------
// Activity: who ordered since the cutoff
// ---------------------------------------------------------------------------

/**
 * From ORDERS_SINCE nodes. Cancelled and test orders do not count; $0 orders do.
 * Returns { activeCustomerIds: Set, activeAddresses: Map<addressKey, { customerId, createdAt }>,
 * activeAt: Map<customerId, createdAt of their newest valid order> }.
 */
export function buildActivity(orderNodes) {
  const activeCustomerIds = new Set();
  const activeAddresses = new Map();
  const activeAt = new Map();
  for (const o of orderNodes) {
    if (o.cancelledAt || o.test || !o.customer?.id) continue;
    activeCustomerIds.add(o.customer.id);
    if (!activeAt.has(o.customer.id) || timeMs(o.createdAt) > timeMs(activeAt.get(o.customer.id))) {
      activeAt.set(o.customer.id, o.createdAt);
    }
    const key = addressKey(o.customer.defaultAddress);
    if (!key) continue;
    const prev = activeAddresses.get(key);
    if (!prev || Date.parse(o.createdAt) > Date.parse(prev.createdAt)) {
      activeAddresses.set(key, { customerId: o.customer.id, createdAt: o.createdAt });
    }
  }
  return { activeCustomerIds, activeAddresses, activeAt };
}

// ---------------------------------------------------------------------------
// Pass 1: rules 1-8 (no extra data needed)
// ---------------------------------------------------------------------------

/**
 * @param {Iterable} customers parsed customers (parseCustomer)
 * @param {{ config, nowMs, cutoffMs, activity }} ctx
 * @returns {{ candidates, excluded: [{ customer, reasons }], needsHistory: string[] }}
 */
export function screen(customers, ctx) {
  const { config, nowMs, cutoffMs } = ctx;
  const all = Array.isArray(customers) ? customers : [...customers];
  const activeIds = new Set(ctx.activity.activeCustomerIds);
  const activeAddresses = new Map(ctx.activity.activeAddresses);
  const activeAt = new Map(ctx.activity.activeAt ?? []);
  // A valid last order inside the window also counts, in case the orders query missed it.
  for (const c of all) {
    if (isValidOrder(c.lastOrder) && c.lastOrder.createdAtMs >= cutoffMs) {
      activeIds.add(c.id);
      if (!activeAt.has(c.id) || c.lastOrder.createdAtMs > timeMs(activeAt.get(c.id))) {
        activeAt.set(c.id, c.lastOrder.createdAt);
      }
      if (c.addressKey && !activeAddresses.has(c.addressKey)) {
        activeAddresses.set(c.addressKey, { customerId: c.id, createdAt: c.lastOrder.createdAt });
      }
    }
  }

  const excludeTags = config.excludeTags.map((t) => ({ lower: t.toLowerCase(), original: t }));
  const sentTag = config.sentTag.toLowerCase();
  const minAgeMs = config.minAccountAgeDays * 86_400_000;
  const candidates = [];
  const excluded = [];

  for (const c of all) {
    const reasons = [];
    const tags = c.tags.map((t) => t.toLowerCase());
    if (!c.email || !c.emailValid) reasons.push({ code: 'no-email' });
    else if (config.excludeEmailDomains.includes(c.emailDomain)) reasons.push({ code: 'relay-email', detail: c.emailDomain });
    if (config.requireEmailSubscribed && c.marketingState !== 'SUBSCRIBED') reasons.push({ code: 'not-subscribed', detail: c.marketingState ?? 'NONE' });
    if (nowMs - c.createdAtMs < minAgeMs) reasons.push({ code: 'too-new', detail: c.createdAt });
    const tag = excludeTags.find((t) => tags.includes(t.lower));
    if (tag) reasons.push({ code: 'excluded-tag', detail: tag.original });
    if (tags.includes(sentTag)) reasons.push({ code: 'already-sent', detail: config.sentTag });
    if (activeIds.has(c.id)) {
      // The date of the order that makes them active, not the snapshot's (possibly older or cancelled) lastOrder.
      reasons.push({ code: 'recent-order', detail: activeAt.get(c.id) ?? c.lastOrder?.createdAt ?? '' });
    } else if (c.addressKey && activeAddresses.has(c.addressKey)) {
      const a = activeAddresses.get(c.addressKey);
      if (a.customerId !== c.id) reasons.push({ code: 'active-address', detail: a.customerId, at: a.createdAt });
    }
    if (reasons.length) excluded.push({ customer: c, reasons });
    else candidates.push(c);
  }

  const needsHistory = candidates
    .filter((c) => c.lastOrder && !(isValidOrder(c.lastOrder) && isGiftableTotal(c.lastOrder.totalCents, config.giftPercent)))
    .map((c) => c.id);
  return { candidates, excluded, needsHistory };
}

// ---------------------------------------------------------------------------
// Pass 2: rules 9-12 (rule 12 and the history part of rule 7 use the follow-up lookups), dedupe, amounts
// ---------------------------------------------------------------------------

function keeperOrder(a, b) {
  const ad = a.basis?.createdAtMs ?? -Infinity;
  const bd = b.basis?.createdAtMs ?? -Infinity;
  if (ad !== bd) return bd > ad ? 1 : -1;
  if (a.c.numberOfOrders !== b.c.numberOfOrders) return b.c.numberOfOrders - a.c.numberOfOrders;
  if (a.c.createdAtMs !== b.c.createdAtMs) return a.c.createdAtMs - b.c.createdAtMs;
  return compareNumericId(a.c.numericId, b.c.numericId);
}

function recipientOrder(a, b) {
  const ak = a.kind === 'ordered' ? 0 : 1;
  const bk = b.kind === 'ordered' ? 0 : 1;
  if (ak !== bk) return ak - bk;
  const ad = a.kind === 'ordered' ? Date.parse(a.basis.createdAt) : Date.parse(a.accountCreatedAt);
  const bd = b.kind === 'ordered' ? Date.parse(b.basis.createdAt) : Date.parse(b.accountCreatedAt);
  if (ad !== bd) return bd - ad; // newest first
  return compareNumericId(a.numericId, b.numericId);
}

/**
 * @param {{ candidates, excluded }} screened output of screen()
 * @param {Map<string, { orders: object[], deleted?: boolean }>} history newest-first parsed orders for
 *   every id in needsHistory; `deleted` = Shopify no longer has the customer (rule 12)
 * @param {{ config, nowMs, cutoffMs }} ctx
 * @returns {{ recipients, excluded, duplicates, averageCents, neverAmountCents }}
 */
export function finalize(screened, history, ctx) {
  const { config, cutoffMs } = ctx;
  const percent = config.giftPercent;
  const tiersCents = config.giftTiersCents;
  const excluded = [...screened.excluded];
  const sources = new Set(config.excludeOrderSources.map((s) => s.toLowerCase()));
  const fromMarketplace = (o) => !!o && sources.has(String(o.sourceName).toLowerCase());
  const stage = [];

  for (const c of screened.candidates) {
    const lo = c.lastOrder;
    let basis = null;
    let basisWhy = null;
    let neverReason = null;
    if (lo && isValidOrder(lo) && isGiftableTotal(lo.totalCents, percent)) {
      basis = lo;
    } else if (lo) {
      const h = history.get(c.id);
      if (!h) throw new Error(`order history missing for ${c.id}; it must be fetched before finalize()`);
      if (h.deleted) {
        excluded.push({ customer: c, reasons: [{ code: 'customer-deleted' }] });
        continue;
      }
      const valid = h.orders.filter(isValidOrder);
      const recent = valid.find((o) => o.createdAtMs >= cutoffMs);
      if (recent) {
        excluded.push({ customer: c, reasons: [{ code: 'recent-order', detail: recent.createdAt }] });
        continue;
      }
      basis = valid.find((o) => isGiftableTotal(o.totalCents, percent)) ?? null;
      basisWhy = lo.cancelled ? 'last-cancelled' : lo.test ? 'last-test' : 'last-zero';
      if (!basis) neverReason = 'only-cancelled-or-zero';
    } else {
      neverReason = 'no-orders';
    }
    // Rule 9 looks at the most recent valid order (not cancelled, not a test; $0 included)
    // as well as the order the amount comes from.
    const latestValid = isValidOrder(lo) ? lo : (history.get(c.id)?.orders.find(isValidOrder) ?? null);
    const marketplace = [latestValid, basis].find(fromMarketplace);
    if (marketplace) {
      excluded.push({ customer: c, reasons: [{ code: 'marketplace-order', detail: marketplace.sourceName }] });
      continue;
    }
    stage.push({ c, basis, basisWhy, neverReason, latestValid, groupId: null, groupSize: 1, groupOthers: [] });
  }

  // Same-address dedupe: identical keys are one person; keep one account.
  const byKey = new Map();
  const kept = [];
  for (const s of stage) {
    if (!s.c.addressKey) {
      kept.push(s);
      continue;
    }
    if (!byKey.has(s.c.addressKey)) byKey.set(s.c.addressKey, []);
    byKey.get(s.c.addressKey).push(s);
  }
  const duplicates = [];
  let groupNo = 0;
  for (const [key, members] of byKey) {
    members.sort(keeperOrder);
    kept.push(members[0]);
    if (members.length === 1) continue;
    groupNo += 1;
    const groupId = `G${String(groupNo).padStart(4, '0')}`;
    for (const m of members) {
      m.groupId = groupId;
      m.groupSize = members.length;
      m.groupOthers = members.filter((x) => x !== m).map((x) => x.c.numericId);
    }
    duplicates.push({
      groupId,
      addressKey: key,
      address: describeKey(key),
      size: members.length,
      flaggedBulk: members.length >= 10,
      keptCustomerId: members[0].c.id,
      members: members.map((m, i) => ({
        customerId: m.c.id,
        numericId: m.c.numericId,
        name: displayName(m.c),
        email: m.c.email,
        kept: i === 0,
        lastPaidOrderAt: m.basis?.createdAt ?? null,
        numberOfOrders: m.c.numberOfOrders,
        accountCreatedAt: m.c.createdAt,
      })),
    });
    for (const m of members.slice(1)) {
      excluded.push({ customer: m.c, reasons: [{ code: 'duplicate-address', detail: members[0].c.id }], groupId });
    }
  }

  // Average of the kept "ordered" people's last paid order totals (raw order totals, before tiers),
  // rounded to whole cents BEFORE the percentage, so "10% × 平均 $X" in the Excel can be redone by hand.
  const ordered = kept.filter((s) => s.basis);
  const averageCents = ordered.length ? Math.round(ordered.reduce((sum, s) => sum + s.basis.totalCents, 0) / ordered.length) : null;
  const neverGift = averageCents === null ? null : giftFor(averageCents, { percent, tiersCents });

  const recipients = [];
  for (const s of kept) {
    const kind = s.basis ? 'ordered' : 'never';
    const g = s.basis ? giftFor(s.basis.totalCents, { percent, tiersCents }) : neverGift;
    if (!g || g.amountCents < 1) {
      excluded.push({ customer: s.c, reasons: [{ code: 'zero-amount', detail: g ? '' : '没有可用来算平均值的有下单客户' }] });
      continue;
    }
    recipients.push(makeRecipient(s, kind, g, s.basis ? s.basis.totalCents : averageCents, config));
  }
  recipients.sort(recipientOrder);
  recipients.forEach((r, i) => {
    r.seq = i + 1;
  });

  return { recipients, excluded, duplicates, averageCents, neverAmountCents: neverGift?.amountCents ?? null };
}

function makeRecipient(s, kind, g, baseCents, config) {
  const c = s.c;
  return {
    seq: 0,
    customerId: c.id,
    numericId: c.numericId,
    name: displayName(c),
    firstName: c.firstName,
    lastName: c.lastName,
    email: c.email,
    marketingState: c.marketingState,
    kind,
    basis: s.basis ? { orderId: s.basis.id, orderName: s.basis.name, createdAt: s.basis.createdAt, totalCents: s.basis.totalCents, sourceName: s.basis.sourceName } : null,
    basisWhy: s.basisWhy,
    neverReason: s.neverReason,
    // Channel of the most recent valid order (the one rule 9 judged); '' when there is none.
    // Never the snapshot's lastOrder: that may be a cancelled order from a marketplace.
    lastOrderSource: s.latestValid?.sourceName ?? '',
    amountCents: g.amountCents,
    rawCents: g.rawCents,
    tier: g.tier, // 0-based index into config.giftTiersCents; null for test campaigns
    formula: formulaText({ kind, baseCents, percent: config.giftPercent, ...g }),
    addressKey: c.addressKey,
    groupId: s.groupId,
    groupSize: s.groupSize,
    groupOthers: s.groupOthers,
    city: c.address?.city ?? '',
    provinceCode: c.address?.provinceCode ?? '',
    zip: c.address?.zip ?? '',
    numberOfOrders: c.numberOfOrders,
    amountSpentCents: c.amountSpentCents,
    accountCreatedAt: c.createdAt,
  };
}

// ---------------------------------------------------------------------------
// Statistics for the summary sheet
// ---------------------------------------------------------------------------

function inc(obj, key, by = 1) {
  obj[key] = (obj[key] ?? 0) + by;
}

/**
 * Exclusion records for the "未入选" sheet plus every count the summary needs.
 * `timezone` (the store's) is used for the dates inside the reason texts.
 */
export function summarize({ total, recipients, excluded, duplicates, averageCents, neverAmountCents, config, timezone }) {
  const byRule = Object.fromEntries(RULES.map((r) => [r.code, 0]));
  const tags = {};
  const channels = {};
  const relayDomains = {};
  const notSubscribed = {};
  const notSelected = [];
  let unlisted = 0;

  for (const e of excluded) {
    const primary = e.reasons[0];
    inc(byRule, primary.code);
    if (primary.code === 'excluded-tag') inc(tags, primary.detail);
    if (primary.code === 'marketplace-order') inc(channels, channelLabel(primary.detail));
    if (primary.code === 'relay-email') inc(relayDomains, primary.detail);
    if (primary.code === 'not-subscribed') inc(notSubscribed, primary.detail);
    if (UNLISTED_RULES.has(primary.code)) {
      unlisted += 1;
      continue;
    }
    const c = e.customer;
    notSelected.push({
      customerId: c.id,
      numericId: c.numericId,
      name: displayName(c),
      email: c.email,
      marketingState: c.marketingState,
      primaryRule: RULE[primary.code].n,
      primaryCode: primary.code,
      primaryText: reasonText(primary, config, timezone),
      allReasons: e.reasons.map((r) => reasonText(r, config, timezone)).join('；'),
      relatedCustomerId: ['active-address', 'duplicate-address'].includes(primary.code) ? primary.detail : '',
      relatedAt: primary.at ?? '',
      groupId: e.groupId ?? '',
      lastOrderAt: c.lastOrder?.createdAt ?? '',
      lastOrderSource: c.lastOrder?.sourceName ?? '',
      numberOfOrders: c.numberOfOrders,
      tags: c.tags.join(', '),
    });
  }
  notSelected.sort((a, b) => a.primaryRule - b.primaryRule || compareNumericId(a.numericId, b.numericId));

  const ordered = recipients.filter((r) => r.kind === 'ordered');
  const never = recipients.filter((r) => r.kind === 'never');
  const raws = ordered.map((r) => r.rawCents).sort((a, b) => a - b); // 10% bases before tiering
  // People and money per tier, split by kind ("ordered" / "never").
  const tiers = (config.giftTiersCents ?? []).map((amountCents, i) => ({
    tier: i,
    label: formatUsd(amountCents),
    amountCents,
    ordered: recipients.filter((r) => r.kind === 'ordered' && r.tier === i).length,
    never: recipients.filter((r) => r.kind === 'never' && r.tier === i).length,
  }));
  for (const t of tiers) {
    t.count = t.ordered + t.never;
    t.cents = t.count * t.amountCents;
  }
  const recipientChannels = {};
  for (const r of recipients) inc(recipientChannels, r.lastOrderSource ? channelLabel(r.lastOrderSource) : '没有有效订单');
  const sizes = duplicates.map((d) => d.size);
  const sizeHistogram = {};
  for (const n of sizes) inc(sizeHistogram, n === 2 ? '2' : n === 3 ? '3' : n <= 5 ? '4-5' : n <= 9 ? '6-9' : '10+');

  return {
    notSelected,
    funnel: {
      total,
      byRule: RULES.map((r) => ({ n: r.n, code: r.code, label: ruleLabel(r.code, config), count: byRule[r.code] })),
      tags,
      channels,
      relayDomains,
      notSubscribed,
      listed: notSelected.length,
      unlisted,
      recipients: recipients.length,
    },
    stats: {
      recipients: recipients.length,
      orderedCount: ordered.length,
      orderedCents: ordered.reduce((a, r) => a + r.amountCents, 0),
      neverCount: never.length,
      neverCents: never.reduce((a, r) => a + r.amountCents, 0),
      totalCents: recipients.reduce((a, r) => a + r.amountCents, 0),
      averageCents,
      neverAmountCents,
      medianOrderedRawCents: raws.length ? raws[Math.floor(raws.length / 2)] : null,
      tiers,
      basisFromEarlier: {
        lastCancelled: ordered.filter((r) => r.basisWhy === 'last-cancelled').length,
        lastZero: ordered.filter((r) => r.basisWhy === 'last-zero').length,
        lastTest: ordered.filter((r) => r.basisWhy === 'last-test').length,
      },
      neverReasons: { noOrders: never.filter((r) => r.neverReason === 'no-orders').length, onlyCancelledOrZero: never.filter((r) => r.neverReason === 'only-cancelled-or-zero').length },
      recipientChannels,
      duplicates: {
        groups: duplicates.length,
        accounts: sizes.reduce((a, b) => a + b, 0),
        removed: sizes.reduce((a, b) => a + b - 1, 0),
        sizeHistogram,
        largest: [...sizes].sort((a, b) => b - a).slice(0, 10),
        flaggedBulk: duplicates.filter((d) => d.flaggedBulk).length,
      },
      noAddressRecipients: recipients.filter((r) => !r.addressKey).length,
    },
  };
}

// ---------------------------------------------------------------------------
// Test campaigns: exactly the listed customers, fixed small amount
// ---------------------------------------------------------------------------

export function buildTestRecipients(customers, config) {
  const recipients = [];
  const excluded = [];
  const sentTag = config.sentTag.toLowerCase();
  for (const c of customers) {
    if (!c.email || !c.emailValid) {
      excluded.push({ customer: c, reasons: [{ code: 'no-email' }] });
      continue;
    }
    if (c.tags.map((t) => t.toLowerCase()).includes(sentTag)) {
      excluded.push({ customer: c, reasons: [{ code: 'already-sent', detail: config.sentTag }] });
      continue;
    }
    const g = { rawCents: config.testGiftAmountCents, amountCents: config.testGiftAmountCents, tier: null };
    // No follow-up lookup here: the snapshot's lastOrder is the latest valid order only when it is valid itself.
    const latestValid = isValidOrder(c.lastOrder) ? c.lastOrder : null;
    recipients.push(makeRecipient({ c, basis: null, basisWhy: null, neverReason: null, latestValid, groupId: null, groupSize: 1, groupOthers: [] }, 'test', g, g.amountCents, config));
  }
  recipients.forEach((r, i) => {
    r.seq = i + 1;
  });
  return { recipients, excluded };
}
