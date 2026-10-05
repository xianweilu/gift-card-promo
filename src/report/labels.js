// Labels and the small pure formatters that turn journal / summary values into
// one-line texts for the Excel report, in either language of src/report/text.js.
// No I/O here.
//
// Every formatter takes an optional last argument `lang`: 'zh' (the default, so
// every existing caller keeps the Chinese texts), 'en', or a text pack itself.
// The dictionaries exported under their old names are the Chinese ones;
// labelsFor(lang) returns the pack of a language, which has the same names.
//
// Status codes come from src/campaign.js (STATUS); skip reasons and verify
// issue types from the interface spec. An unknown code is shown as-is, so a
// newer (or older: the removed remind command's) journal never breaks the report.

import { formatUsd } from '../select/amount.js';
import { TEXT, textFor } from './text.js';

const ZH = TEXT.zh;

/** The text pack of `lang` ('zh' | 'en' | a pack): the dictionaries below under the same names, and every other workbook text. */
export function labelsFor(lang = 'zh') {
  return textFor(lang);
}

/** Issue status of a recipient (Excel "状态" column). */
export const STATUS_LABELS = ZH.STATUS_LABELS;

/** Display order of the statuses in the summary and the explanation sheet. */
export const STATUS_ORDER = Object.freeze(['pending', 'in_progress', 'unknown', 'created', 'done', 'failed', 'skipped']);

/** Row fill per status (ARGB). Pending stays unfilled (white) so the grid lines remain visible. */
export const STATUS_FILLS = Object.freeze({
  pending: null,
  in_progress: 'FFDDEBF7', // light blue
  unknown: 'FFFCE4D6', // light orange
  created: 'FFFFF2CC', // light yellow
  done: 'FFE2EFDA', // light green
  failed: 'FFFFD7D7', // light red
  skipped: 'FFEDEDED', // light grey
});

/** What each status means (sheet "说明"). */
export const STATUS_HELP = ZH.STATUS_HELP;
/** Pre-flight skip reasons written by issue (journal op "skip"). */
export const ISSUE_SKIP_LABELS = ZH.ISSUE_SKIP_LABELS;

/**
 * The tag usage puts on the customer of every used card: "<SENT_TAG>-USED", e.g.
 * OCT26RTPROMO-USED (test campaign OCT26RTPROMO-TEST2: OCT26RTPROMO-TEST2-USED). The same
 * name as usedTagName() in src/usage.js; test/excel.test.js checks that the two agree.
 */
export function usedTagName(sentTag) {
  return `${sentTag}-USED`;
}

/** The Shopify Email recipient condition of the reminders (the same text in both editions). */
export function reminderSegmentCondition(sentTag) {
  return `customer_tags CONTAINS '${sentTag}' AND NOT customer_tags CONTAINS '${usedTagName(sentTag)}'`;
}

/**
 * How the reminders work now (Shopify Email + the used-card tag), on the "说明" sheet, with
 * this campaign's tag names (`sentTag` = the campaign's SENT_TAG): [item, text] pairs.
 */
export function reminderHelp(sentTag, lang = 'zh') {
  return textFor(lang).reminderHelp(sentTag, usedTagName(sentTag), reminderSegmentCondition(sentTag));
}

/** Customer kind of a recipient. */
export const KIND_LABELS = ZH.KIND_LABELS;
/** Shopify email marketing states ('NONE' = the commands' stand-in for a missing state). */
export const MARKETING_LABELS = ZH.MARKETING_LABELS;
/** Notes for recipients whose amount is not based on their very last order. */
export const BASIS_NOTES = ZH.BASIS_NOTES;
export const NEVER_NOTES = ZH.NEVER_NOTES;
/** Journal ops (sheet "操作日志"). */
export const OP_LABELS = ZH.OP_LABELS;
/** Where a reconcile.found entry came from. */
export const RECONCILE_SOURCE_LABELS = ZH.RECONCILE_SOURCE_LABELS;
/** verify.json issue types (sheet "核对"). */
export const VERIFY_TYPE_LABELS = ZH.VERIFY_TYPE_LABELS;
/** selection.snapshot.source. */
export const SNAPSHOT_SOURCE_LABELS = ZH.SNAPSHOT_SOURCE_LABELS;
/** Fixed English notes some journal entries carry (anything else is shown as-is). */
export const NOTE_LABELS = ZH.NOTE_LABELS;
/** usage.json unmatched[].reason codes that may appear (anything else is shown as-is). */
export const UNMATCHED_REASON_LABELS = ZH.UNMATCHED_REASON_LABELS;
/** Command exit codes (sheet "说明"). */
export const EXIT_CODE_HELP = ZH.EXIT_CODE_HELP;
/**
 * Why 礼品卡抵扣 and 已用金额 can differ (使用报告, 汇总, 说明): the first is
 * taken at checkout (usage.js F1), the second from the card balances.
 */
export const USAGE_SPLIT_NOTE = ZH.USAGE_SPLIT_NOTE;

// ---------------------------------------------------------------------------
// Small lookups
// ---------------------------------------------------------------------------

/** dict[key] when it is the dictionary's own entry (never an inherited Object member such as "constructor"). */
function own(dict, key) {
  return Object.hasOwn(dict, key) ? dict[key] : undefined;
}

/** label from a dictionary, or the raw code (empty for null/undefined). */
export function labelOf(dict, code) {
  if (code === null || code === undefined || code === '') return '';
  return own(dict, code) ?? String(code);
}

/**
 * Like labelOf, for values that are either a code or a text our commands stored
 * (usage.json unmatched[].reason): anything not in `dict` goes through the pack's
 * stored.text (Chinese: the raw value, exactly as labelOf).
 */
export function storedLabelOf(dict, code, lang = 'zh') {
  if (code === null || code === undefined || code === '') return '';
  return own(dict, code) ?? textFor(lang).stored.text(String(code));
}

// A journal detail that is an instant: Shopify's ISO timestamps (always with a zone).
const ISO_INSTANT = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})$/;
const GID = /^gid:\/\/shopify\/(\w+)\/(\d+)(?:\?.*)?$/;

/**
 * A skip detail from the journal as people read it: an ISO instant becomes
 * store-local "YYYY-MM-DD HH:MM 店铺时间" (via `stamp`, the workbook's clock;
 * no brackets, because the detail itself is shown in brackets), a customer
 * gid "客户 <number>", a marketing state its label.
 * Anything else (a text our commands wrote, an order name, a domain, ...) goes
 * through the pack's stored.text (Chinese: unchanged).
 * @param {*} detail
 * @param {{ stamp?: (iso: string) => string }} [options]
 * @param {string|object} [lang]
 */
export function detailText(detail, { stamp } = {}, lang = 'zh') {
  if (detail === null || detail === undefined) return '';
  const s = String(detail).trim();
  if (!s) return '';
  const T = textFor(lang);
  if (stamp && ISO_INSTANT.test(s) && Number.isFinite(Date.parse(s))) {
    const local = stamp(s);
    if (local) return T.storeTime(local);
  }
  const g = GID.exec(s);
  if (g) return own(T.GID_TYPE_LABELS, g[1]) ? `${T.GID_TYPE_LABELS[g[1]]} ${g[2]}` : g[2];
  return own(T.MARKETING_LABELS, s) ?? T.stored.text(s);
}

/** Issue skip reason with its optional detail, e.g. "导出后下过单（#1234）"; `options` as for detailText. */
export function issueSkipText(reason, detail, options, lang = 'zh') {
  const T = textFor(lang);
  const label = labelOf(T.ISSUE_SKIP_LABELS, reason);
  const d = detailText(detail, options, T);
  return d ? T.withDetail(label, d) : label;
}

/** A fixed English journal note in the pack's language; other notes through stored.text (Chinese: unchanged). */
export function noteText(note, lang = 'zh') {
  if (note === null || note === undefined) return '';
  const T = textFor(lang);
  return own(T.NOTE_LABELS, note) ?? T.stored.text(String(note));
}

// ---------------------------------------------------------------------------
// Shopify error messages (built by src/shopify.js) → the pack's language
// ---------------------------------------------------------------------------

/** The " [CODE]" that throwIfUserErrors puts after each userError message ("…; " between them). */
const USER_ERROR_CODE = / \[([A-Z][A-Z0-9_]*)\](?=; |$)/g;

// [message form, text in the pack's words]. Only our own English around Shopify's
// words is translated: Shopify's message text stays as it is.
const SHOPIFY_ERROR_FORMS = [
  // throwIfUserErrors: "<mutation> rejected: <field>: <message> [<CODE>]; ..."
  [/^[a-z][A-Za-z0-9]* rejected: ([\s\S]*)$/, (m, E) => E.rejected(m[1].replace(USER_ERROR_CODE, (_, code) => E.userErrorCode(code)))],
  [/^Network error calling Shopify(?:: ([\s\S]*))?$/, (m, E) => E.network(m[1])],
  [/^Shopify HTTP (\d+)(?:: ([\s\S]*))?$/, (m, E) => E.http(m[1], m[2])],
  [/^(?:Throttled by Shopify|HTTP 429 from Shopify) (\d+) times in a row; giving up$/, (m, E) => E.throttled(m[1])],
  [/^GraphQL error(?:: ([\s\S]*))?$/, (m, E) => E.graphql(m[1])],
  [/^Shopify returned non-JSON\b/, (m, E) => E.nonJson],
  [/^Shopify response contained no data$/, (m, E) => E.noData],
  // src/customers.js addTag (issue's sent tag, usage's used tag) and src/giftcards.js: a mutation answer without its payload
  [/^[a-z][A-Za-z0-9]* returned no payload$/, (m, E) => E.noPayload],
];

/** One Shopify error message in the pack's words, or null when it is not one of the forms above. */
function shopifyErrorText(message, T) {
  for (const [form, text] of SHOPIFY_ERROR_FORMS) {
    const m = form.exec(message);
    if (m) return text(m, T.shopifyErrors);
  }
  return null;
}

/**
 * A journal error as people read it in the workbook (发放名单 备注/错误, 操作日志
 * 错误), e.g.
 *   "giftCardCreate rejected: input: Customer is invalid [INVALID]" → "Shopify 拒绝：input: Customer is invalid（INVALID）"
 *   "Network error calling Shopify: fetch failed" → "网络错误：fetch failed"
 * Anything else (a text our commands wrote, another error) goes through the
 * pack's stored.text (Chinese: unchanged). The commands also embed a Shopify
 * message in their own Chinese text, after "：" and joined with "；" (issue:
 * "<error>；查卡也失败了：<error>"); each of those is translated too.
 * @param {*} message
 * @param {string|object} [lang]
 */
export function errorText(message, lang = 'zh') {
  if (message === null || message === undefined) return '';
  const T = textFor(lang);
  // '；' and '：' are the separators of the commands' stored (Chinese) texts.
  return String(message)
    .split('；')
    .map((part) => {
      const whole = shopifyErrorText(part, T);
      if (whole !== null) return whole;
      const colon = part.indexOf('：');
      const tail = colon >= 0 ? shopifyErrorText(part.slice(colon + 1), T) : null;
      return tail === null ? T.stored.text(part) : `${T.stored.text(part.slice(0, colon))}${T.sep.colon}${tail}`;
    })
    .join(T.sep.list);
}

// ---------------------------------------------------------------------------
// run.end summaries → one line
// ---------------------------------------------------------------------------

// Already shown in their own columns of the run history (批次, 预演/实际); `round` is the
// removed remind command's (older journals).
const SUMMARY_SKIP_KEYS = new Set(['batch', 'round', 'dryRun']);

// A run's mode reads first ("只补记和补打 tag，补打 tag 2"), wherever the command put the key.
const SUMMARY_LEAD_KEYS = new Set(['repairOnly']);

const MAX_SUMMARY_LENGTH = 500;

// verify's own summary text when it failed: "失败：<message>" (the message may be Shopify's).
const STORED_FAILED_TEXT = /^(失败：)(.+)$/s;

/** The command a summary came from when the caller does not say (only issue summaries need it). */
function summaryCommand(summary) {
  if ('tagFixed' in summary || 'attempted' in summary || 'seqFrom' in summary) return 'issue';
  return null;
}

function keyLabel(key, ctx) {
  const S = ctx.T.runSummary;
  if (key === 'attempted' && ctx.dryRun && ctx.command === 'issue') return S.attemptedDry;
  return own(ctx.T.SUMMARY_KEY_LABELS, key) ?? key;
}

/** Label of a key inside a nested object: `skipped` → issue skip reasons, `counts` → verify issue types. */
function innerLabel(parent, key, ctx) {
  if (parent === 'skipped') return labelOf(ctx.T.ISSUE_SKIP_LABELS, key);
  if (parent === 'counts') return labelOf(ctx.T.VERIFY_TYPE_LABELS, key);
  return keyLabel(key, ctx);
}

function summaryValue(key, value, depth, ctx) {
  const { T } = ctx;
  if (value === null || value === undefined || value === '') return '';
  // false flags and zero counters mean "nothing happened": left out
  if (typeof value === 'boolean') return value ? T.runSummary.yes : '';
  if (typeof value === 'number') {
    if (!Number.isFinite(value) || value === 0) return '';
    return /cents$/i.test(key) ? formatUsd(value) : value.toLocaleString('en-US');
  }
  if (typeof value === 'string') return key === 'source' ? labelOf(T.SNAPSHOT_SOURCE_LABELS, value) : T.stored.text(value);
  if (Array.isArray(value)) {
    const items = value.slice(0, 10).map((v) => summaryValue('', v, depth + 1, ctx)).filter(Boolean);
    return items.length ? `${items.join(T.sep.enum)}${value.length > 10 ? T.runSummary.more(value.length) : ''}` : '';
  }
  if (typeof value === 'object') {
    if (depth >= 2) return '…';
    const inner = summaryParts(value, depth + 1, ctx, key);
    return inner.length ? T.runSummary.nested(inner.join(T.sep.comma)) : '';
  }
  return String(value);
}

function summaryParts(obj, depth, ctx, parent = null) {
  const { T } = ctx;
  const parts = [];
  const top = depth === 0;
  let seqShown = false;
  const entries = Object.entries(obj);
  // stable sort: lead keys first, everything else in the command's order
  if (top) entries.sort(([a], [b]) => Number(SUMMARY_LEAD_KEYS.has(b)) - Number(SUMMARY_LEAD_KEYS.has(a)));
  for (const [key, value] of entries) {
    if (top && SUMMARY_SKIP_KEYS.has(key)) continue;
    // A summary's own `text` was already judged by stored.summaryText (shown instead of the
    // counters, or ignored); never repeat it as a counter.
    if (top && key === 'text') continue;
    if (top && (key === 'seqFrom' || key === 'seqTo')) {
      // one "序号 3–5" (at the first of the two keys) instead of two counters
      if (seqShown) continue;
      seqShown = true;
      const from = Number.isFinite(obj.seqFrom) ? obj.seqFrom : null;
      const to = Number.isFinite(obj.seqTo) ? obj.seqTo : null;
      if (from !== null && to !== null) parts.push(from === to ? T.runSummary.seqOne(from) : T.runSummary.seqRange(from, to));
      else if (from !== null || to !== null) parts.push(T.runSummary.seqOne(from ?? to));
      continue;
    }
    if (top && key === 'stopped' && typeof value === 'string' && value) {
      parts.push(T.runSummary.stopped(value));
      continue;
    }
    if (top && key === 'error' && typeof value === 'string' && value) {
      parts.push(T.runSummary.error(errorText(value, T)));
      continue;
    }
    const text = summaryValue(key, value, depth, ctx);
    if (!text) continue;
    const label = parent ? innerLabel(parent, key, ctx) : keyLabel(key, ctx);
    if (typeof value === 'boolean') parts.push(label);
    else parts.push(T.runSummary.labelled(label, text));
  }
  return parts;
}

/**
 * run.end summary (any shape) → one line for the run history, e.g.
 * "尝试 20，建卡 20，打 tag 20，金额 $215.40，序号 1–20". A summary's own `text`
 * is shown as the pack's stored.summaryText decides (Chinese: as-is); zero counters are left out.
 * @param {*} summary
 * @param {{ command?: string|null, dryRun?: boolean }} [run] the run it belongs to
 *   (command and dry-run flag of its run.start); inferred from the summary when missing
 * @param {string|object} [lang]
 */
export function summaryText(summary, run = {}, lang = 'zh') {
  if (summary === null || summary === undefined) return '';
  const T = textFor(lang);
  let text;
  if (typeof summary === 'object' && !Array.isArray(summary)) {
    const storedText = typeof summary.text === 'string' && summary.text.trim() ? T.stored.summaryText(summary.text.trim()) : null;
    if (storedText !== null && storedText !== undefined) {
      // verify writes e.g. "失败：Network error calling Shopify: …": translate an embedded Shopify error.
      text = storedText.replace(STORED_FAILED_TEXT, (_, head, rest) => `${head}${errorText(rest, T)}`);
    } else {
      const ctx = {
        T,
        command: run.command ?? summaryCommand(summary),
        dryRun: typeof run.dryRun === 'boolean' ? run.dryRun : summary.dryRun === true,
      };
      const parts = summaryParts(summary, 0, ctx);
      text = parts.length ? parts.join(T.sep.comma) : Object.keys(summary).length ? T.runSummary.allZero : '';
    }
  } else {
    text = summaryValue('', summary, 0, { T, command: run.command ?? null, dryRun: !!run.dryRun });
  }
  return text.length > MAX_SUMMARY_LENGTH ? `${text.slice(0, MAX_SUMMARY_LENGTH - 1)}…` : text;
}
