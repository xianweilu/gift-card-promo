// English renderings of the texts our commands store in Chinese in the data files
// (selection.json, journal.jsonl, usage.json, run summaries), for TEXT.en.stored of
// ./text.js. Every form our own code writes is listed here (src/select/rules.js,
// src/select/amount.js, src/issue.js, src/verify.js, src/usage.js);
// anything else is returned unchanged, so an unknown text never breaks the report.
// No I/O here.

import { formatUsd } from '../select/amount.js';
import { RULES, channelLabel, reasonDay } from '../select/rules.js';

// ---------------------------------------------------------------------------
// Selection rules (selection.json funnel.byRule and notSelected)
// ---------------------------------------------------------------------------

const RULE_BY_N = new Map(RULES.map((r) => [r.n, r]));
const RULE_BY_CODE = new Map(RULES.map((r) => [r.code, r]));

/** English names of the rules; too-new / recent-order / active-address take the configured numbers. */
const RULE_LABELS_EN = {
  'no-email': 'No email or invalid email format',
  'relay-email': 'Email domain on the exclusion list',
  'not-subscribed': 'Not subscribed to marketing emails',
  'too-new': (days) => `Account younger than ${days} days`,
  'excluded-tag': 'Carries an excluded tag',
  'already-sent': 'Already sent (has the sent tag)',
  'recent-order': (months) => `Ordered in the last ${months} months`,
  'active-address': (months) => `Same-address account ordered in the last ${months} months`,
  'marketplace-order': 'Latest order from a marketplace channel',
  'duplicate-address': 'Same address, another account kept',
  'zero-amount': 'Computed amount is $0',
  'customer-deleted': 'Customer deleted at the order follow-up',
};

/** The first number inside a stored Chinese label ("注册不满 7 天" → 7), or null. */
function numberIn(text) {
  const m = /\d+/.exec(String(text ?? ''));
  return m ? Number(m[0]) : null;
}

/**
 * English name of rule `code`. `params` = selection.params (minAccountAgeDays, inactiveMonths);
 * `storedLabel` = the Chinese label as stored, whose number is used when params lack it.
 * An unknown code gets the stored label (or the code) unchanged.
 */
export function ruleLabelEn(code, params = {}, storedLabel = '') {
  const label = RULE_LABELS_EN[code];
  if (typeof label === 'string') return label;
  if (typeof label === 'function') {
    const n = code === 'too-new' ? params?.minAccountAgeDays : params?.inactiveMonths;
    const value = Number.isFinite(Number(n)) && n !== null && n !== undefined && n !== '' ? Number(n) : numberIn(storedLabel) ?? (code === 'too-new' ? 7 : 3);
    return label(value);
  }
  return storedLabel || String(code ?? '');
}

/** Channel names as select stores them (Chinese for the store's own channels) → English. */
const CHANNEL_LABELS_EN = {
  web: 'Online Store (web)',
  checkout_next: 'New checkout',
  shopify_draft_order: 'Draft order',
  网店: 'Online Store (web)',
  新版结账: 'New checkout',
  草稿订单: 'Draft order',
};

/** A channel code or a stored channel label in English; '' for nothing, unknown codes as channelLabel shows them. */
export function channelLabelEn(value) {
  if (value === null || value === undefined || value === '') return '';
  const s = String(value);
  return CHANNEL_LABELS_EN[s.toLowerCase()] ?? CHANNEL_LABELS_EN[s] ?? channelLabel(s);
}

const NO_AVERAGE_ZH = '没有可用来算平均值的有下单客户';
const NO_AVERAGE_EN = 'no ordering customers to compute the average from';

/** The detail of one reason in English (`code` known, `detail` as stored by rules.js). */
function reasonDetailEn(code, detail, timeZone) {
  if (detail === null || detail === undefined || detail === '') return '';
  const d = String(detail);
  switch (code) {
    case 'marketplace-order':
      return channelLabelEn(d);
    case 'active-address':
      return `ordering account ${d.split('/').pop()}`;
    case 'duplicate-address':
      return `kept account ${d.split('/').pop()}`;
    case 'recent-order':
    case 'too-new':
      return reasonDay(d, timeZone);
    case 'zero-amount':
      return d === NO_AVERAGE_ZH ? NO_AVERAGE_EN : d;
    default:
      return d;
  }
}

/** "7. Ordered in the last 3 months: 2026-08-15" from a stored { n, code, detail } reason. */
export function reasonTextEn(reason, params, timeZone) {
  const rule = RULE_BY_CODE.get(reason.code) ?? (Number.isInteger(reason.n) ? RULE_BY_N.get(reason.n) : null);
  const code = rule?.code ?? reason.code;
  const n = reason.n ?? rule?.n ?? '?';
  const label = ruleLabelEn(code, params);
  const detail = reasonDetailEn(code, reason.detail, timeZone);
  return detail ? `${n}. ${label}: ${detail}` : `${n}. ${label}`;
}

// A stored reason segment: "N. label" or "N. label：detail" (rules.js reasonText).
const REASON_SEGMENT = /^(\d+)\. ([^：]*?)(?:：([\s\S]*))?$/;

/** One stored Chinese reason text ("5. 带排除 tag：WHS") → English; unknown forms unchanged. */
function parseReasonSegment(segment, params) {
  const m = REASON_SEGMENT.exec(segment.trim());
  if (!m) return segment;
  const n = Number(m[1]);
  const rule = RULE_BY_N.get(n);
  if (!rule) return segment;
  const label = ruleLabelEn(rule.code, params, m[2]);
  const d = m[3] ?? '';
  let detail = d;
  if (rule.code === 'marketplace-order') detail = channelLabelEn(d);
  else if (rule.code === 'active-address') detail = d.replace(/^下单账户\s*/, 'ordering account ');
  else if (rule.code === 'duplicate-address') detail = d.replace(/^保留\s*/, 'kept account ');
  else if (rule.code === 'zero-amount') detail = d === NO_AVERAGE_ZH ? NO_AVERAGE_EN : d;
  return detail ? `${n}. ${label}: ${detail}` : `${n}. ${label}`;
}

/**
 * The reasons of a selection.json notSelected[] entry in English, as one text.
 * `which` = 'primary' (the first reason) or 'all'. Newer lists carry `reasons: [{ n, code, detail }]`;
 * older ones only the Chinese primaryText / allReasons, which are parsed.
 * `info` = { params, timeZone } of the list.
 */
export function notSelectedReasonsEn(entry, which, info = {}) {
  const params = info.params ?? {};
  if (Array.isArray(entry.reasons) && entry.reasons.length && entry.reasons.every((r) => r && typeof r === 'object')) {
    const reasons = which === 'primary' ? entry.reasons.slice(0, 1) : entry.reasons;
    return reasons.map((r) => reasonTextEn(r, params, info.timeZone)).join('; ');
  }
  const stored = which === 'primary' ? entry.primaryText : entry.allReasons;
  if (stored === null || stored === undefined || stored === '') return stored ?? '';
  return String(stored)
    .split('；')
    .map((segment) => parseReasonSegment(segment, params))
    .join('; ');
}

// ---------------------------------------------------------------------------
// Amount formula (selection.json recipients[].formula, src/select/amount.js formulaText)
// ---------------------------------------------------------------------------

/** "10%" / "7.5%", as amount.js shows the percentage. */
function pctText(percent) {
  const p = Number(percent);
  return `${Number.isInteger(p) ? p : p.toFixed(2).replace(/0+$/, '').replace(/\.$/, '')}%`;
}

/** The stored Chinese formula, word by word, when the numbers to rebuild it are missing. */
function translateFormulaText(text) {
  return String(text)
    .replace(/^测试固定金额\s*/, 'Fixed test amount ')
    .replace(/×\s*平均\s*/, '× average ')
    .replace(/→\s*档位\s*/, '→ tier ');
}

/**
 * The amount formula of a recipient in English, rebuilt from the stored numbers:
 *   ordered  "10% × $150.00 = $15.00 → tier $15.33"
 *   never    "10% × average $105.58 = $10.56 → tier $10.77"
 *   test     "Fixed test amount $0.10"
 * `selection` supplies params.giftPercent and averageCents. Without the numbers, the
 * stored text is translated word by word; without a stored text, ''.
 */
export function formulaEn(recipient, selection = {}) {
  const r = recipient ?? {};
  const amount = Number.isFinite(r.amountCents) ? r.amountCents : null;
  if (r.kind === 'test' && amount !== null) return `Fixed test amount ${formatUsd(amount)}`;
  const percent = selection?.params?.giftPercent;
  const raw = Number.isFinite(r.rawCents) ? r.rawCents : null;
  if ((r.kind === 'ordered' || r.kind === 'never') && amount !== null && raw !== null && Number.isFinite(Number(percent))) {
    const baseCents = r.kind === 'ordered' ? r.basis?.totalCents : selection?.averageCents;
    if (Number.isFinite(baseCents)) {
      const base = r.kind === 'never' ? `average ${formatUsd(Math.round(baseCents))}` : formatUsd(baseCents);
      return `${pctText(percent)} × ${base} = ${formatUsd(raw)} → tier ${formatUsd(amount)}`;
    }
  }
  return r.formula ? translateFormulaText(r.formula) : '';
}

// ---------------------------------------------------------------------------
// Free texts of journal.jsonl, run summaries and usage.json
// ---------------------------------------------------------------------------

/** "10 分钟" / "90 秒" (issue.js and verify.js durationText) → "10 minutes" / "90 seconds". */
function durationEn(text) {
  const m = /^(\d+)\s*(分钟|秒)$/.exec(String(text).trim());
  if (!m) return text;
  const n = Number(m[1]);
  const unit = m[2] === '分钟' ? 'minute' : 'second';
  return `${n} ${unit}${n === 1 ? '' : 's'}`;
}

// issue.js noNewCardsText: from the cards' expiry date on, no new card is created
const EXPIRY_NO_NEW_CARDS = /^礼品卡到期日（([^）]*)）已到，不再建新卡$/;
const expiryNoNewCardsEn = (m) => `the gift card expiry date (${m[1]}) has arrived; no new cards are created`;
const REPAIR_ONLY_NO_REISSUE = '本次是 --repair-only，不建新卡；之后正常运行 issue 时才会给他建卡';

/** The tail of an issue reconcile.none note after "确认未建成；" (why no new card follows). */
function noReissueEn(tail) {
  const m = EXPIRY_NO_NEW_CARDS.exec(tail);
  if (m) return expiryNoNewCardsEn(m);
  if (tail === REPAIR_ONLY_NO_REISSUE) return 'this run is --repair-only and creates no new cards; a later normal issue run will create it';
  return null;
}

/** [exact Chinese text, English] of the fixed texts. */
const EXACT = new Map([
  // issue.js skip detail ('no-email' with an address that is not valid)
  ['邮箱格式无效', 'invalid email format'],
  // issue.js create.unknown: a create interrupted by a crash, the card not visible yet
  ['上次运行在建卡途中中断，Shopify 上暂时查不到这张卡', 'The previous run was interrupted while creating this card; Shopify does not show it yet'],
  // issue.js create.unknown: the error's message was missing, and the part before the lookup error
  ['未知错误', 'unknown error'],
  ['查卡也失败了', 'the card lookup failed too'],
  // usage.js unmatched[].reason and the product-name stand-in
  ['回执里没有礼品卡 ID', 'No gift card ID in the receipt'],
  ['回执里没有礼品卡 ID（退款）', 'No gift card ID in the receipt (refund)'],
  ['（无名称）', '(no name)'],
]);

/** [regex, English from its match] of the templated texts. */
const PATTERNS = [
  // issue.js reconcile.none: "超过 10 分钟仍查不到这张卡，确认未建成，可以重试"
  [/^超过 (\d+ (?:分钟|秒))仍查不到这张卡，确认未建成，可以重试$/, (m) => `Card still not found after ${durationEn(m[1])}: confirmed not created, can be retried`],
  // issue.js reconcile.none: "...，确认未建成；<why no new card>"
  [/^超过 (\d+ (?:分钟|秒))仍查不到这张卡，确认未建成；([\s\S]+)$/, (m) => {
    const why = noReissueEn(m[2]);
    return why ? `Card still not found after ${durationEn(m[1])}: confirmed not created; ${why}` : null;
  }],
  // verify.js reconcile.none: "verify：开始建卡 10 分钟后仍没在 Shopify 找到这张卡，确认没有建成"
  [/^verify：开始建卡 (\d+ (?:分钟|秒))后仍没在 Shopify 找到这张卡，确认没有建成$/, (m) => `verify: card not found in Shopify ${durationEn(m[1])} after creation started; confirmed not created`],
  // issue.js noNewCardsText on its own (a part of an error or note)
  [EXPIRY_NO_NEW_CARDS, expiryNoNewCardsEn],
  // verify.js summary.text: the English line is built from the counters (stored.summaryText → null), but
  // the `text` key itself still reaches the line through labels.js summaryValue, so it reads English too.
  [/^卡 (\d+) 张，发现 (\d+) 条，补记日志 (\d+) 条$/, (m) => `${m[1]} cards, ${m[2]} ${m[2] === "1" ? "issue" : "issues"} found, ${m[3]} journal ${m[3] === "1" ? "entry" : "entries"} added`],
  [/^卡 (\d+) 张，没有发现问题$/, (m) => `${m[1]} cards, no issues found`],
  [/^失败：([\s\S]+)$/, (m) => `Failed: ${m[1]}`],
  // durations alone
  [/^\d+ (?:分钟|秒)$/, (m) => durationEn(m[0])],
];

/**
 * A text our own code wrote into the journal or usage.json, in English; a text that is not one
 * of the known forms (customer data, Shopify's words, a newer command's text) comes back unchanged.
 */
export function storedTextEn(text) {
  if (text === null || text === undefined) return '';
  const s = String(text);
  const t = s.trim();
  if (!t) return s;
  const exact = EXACT.get(t);
  if (exact !== undefined) return exact;
  if (t === REPAIR_ONLY_NO_REISSUE) return 'this run is --repair-only and creates no new cards; a later normal issue run will create it';
  for (const [form, render] of PATTERNS) {
    const m = form.exec(t);
    if (m) {
      const out = render(m);
      if (out !== null && out !== undefined) return out;
    }
  }
  return s;
}
