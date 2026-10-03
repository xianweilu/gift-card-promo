// `remind --round 1|2`: re-send the campaign's gift card email to recipients
// whose card is still unused, on REMIND_1_DATE / REMIND_2_DATE.
//
// Shopify's giftCardSendNotificationToCustomer re-sends the same "Gift card
// created" notification; the template switches its copy by the date it is sent
// on, so a live campaign may not send round n before REMIND_<n>_DATE (store
// time), nor round 1 on or after REMIND_2_DATE (that email would carry the
// round-2 copy). Test campaigns and dry runs are not date-gated; they are told
// which copy today's email shows instead.
//
// A long live run can cross midnight (store time), so the store date is
// computed again before every send: a round 1 of the real campaign stops
// (exit 2) once it reaches REMIND_2_DATE, and a card that has expired in the
// meantime is skipped like one found expired at the start. Shopify does not
// document whether the template's date is store time or UTC; a live run of the
// real campaign started when the UTC date would already show another copy is
// warned about once (Los Angeles daytime sends are right either way).
//
// At most one reminder per customer and round:
//   - remind.start is journalled (fsync) BEFORE each send, the outcome after it;
//   - right after each send (remind.ok) the customer gets the round tag
//     `${SENT_TAG}-R${round}` (OCT26RTPROMO-R1 / OCT26RTPROMO-R2; a test campaign
//     has its own, e.g. OCT26RTPROMO-TEST-R1): Shopify's own record of the round,
//     as SENT_TAG is for issue. Every run, dry runs too, first reads who carries
//     it. Such a customer counts as reminded this round whatever the journal says
//     (lost, or replaced by an older copy): never sent this round again, not even
//     with --retry-unknown / --retry-failed, and live runs journal remind.found.
//     A failed tagsAdd is journalled (remind.tag.fail) and only warned about (the
//     journal's remind.ok is still the first guard); the next live run adds the
//     tag again, as it does for a sent row whose tag Shopify does not show at
//     least 10 minutes after the send (its customer search lags behind);
//   - a lost answer (timeout, network, 5xx) leaves the round "unknown", which is
//     NEVER sent again automatically: only --retry-unknown does, after a human check;
//   - a userErrors refusal (remind.fail: Shopify sent nothing) is not retried
//     automatically either; --retry-failed re-sends once the cause is fixed;
//   - a definite rejection (remind.rejected: nothing was sent, e.g. throttled out)
//     stops the run, and the next run tries that customer again;
//   - only one write command runs at a time (run lock), and the list is checked
//     again under the lock (select may have rewritten it just before).
// Dry runs (the default) read Shopify exactly like a live run but journal only
// run.start/run.end and add no tag, so nobody's state changes. Their email
// preview is rendered for the day the batch would really be sent.
//
// Timing: one send is the notification (~0.35 s), the round tag's tagsAdd
// (~0.27 s) and two fsynced journal lines, about 0.7 s in all: about 2 hours
// per 10,000 reminders.

import {
  campaignPaths,
  readJson,
  acquireRunLock,
  appendJournal,
  readJournal,
  foldJournal,
  newRunId,
  REMIND_STATUS,
} from './campaign.js';
import { connect } from './connect.js';
import { gql } from './shopify.js';
import { REFRESH_CUSTOMERS } from './queries.js';
import { findCampaignCards, sendGiftCardNotification } from './giftcards.js';
import { addTag, fetchTaggedCustomerIds } from './customers.js';
import { localDate, localDateTime } from './time.js';
import { formatUsd } from './select/amount.js';

const REFRESH_CHUNK = 250; // ids per RefreshCustomers call (halved if Shopify reports MAX_COST_EXCEEDED)
const MAX_CONSECUTIVE_FAILURES = 5; // failed or unknown sends in a row before the run stops (a failed tag never counts)
// customers(query:) reads a search index: a tag added less than this long before the tag set was
// read may not be listed yet, so only older sends whose tag is missing are tagged again.
const TAG_SEARCH_LAG_MS = 10 * 60 * 1000;
const SECONDS_PER_SEND = 0.7; // notification ~0.35 s + round tag ~0.27 s + 2 journal lines (fsync)
const PROGRESS_EVERY = 100; // send attempts between progress lines
const LIST_FIRST = 10; // people a dry run lists by name
const CARD_WINDOW_MS = 60 * 60 * 1000; // card search starts 1 hour before the selection was made
const INTERRUPTED_ERROR = '上次运行在发送这封提醒时中断，不知道是否已发出';
export const SELECTION_REPLACED = '名单刚被 select 改写，请重新运行';

/** Why a recipient gets no reminder this round, in evaluation order, with the label shown to people. */
export const REMIND_SKIP_REASONS = Object.freeze({
  'no-card': '在 Shopify 上找不到这张卡',
  'multiple-cards': '有多张本活动的卡，无法确定发哪一张',
  'card-disabled': '卡已停用',
  'card-expired': '卡已过期',
  used: '卡已用过',
  'customer-deleted': '客户已删除',
  'no-email': '没有有效邮箱',
  'not-subscribed': '已退订或未订阅营销邮件',
});

// ---------------------------------------------------------------------------
// Pure eligibility rules
// ---------------------------------------------------------------------------

const numericId = (id) => String(id ?? '').split('/').pop();

/**
 * The tag a customer gets once reminder `round` was sent to them: `${sentTag}-R${round}`,
 * e.g. OCT26RTPROMO-R1 / OCT26RTPROMO-R2 (a test campaign's SENT_TAG gives OCT26RTPROMO-TEST-R1).
 * Whoever carries it is never reminded again in that round.
 */
export function roundTag(sentTag, round) {
  return `${sentTag}-R${round}`;
}

/** Whether a customer node (RefreshCustomers; null = deleted) carries `tag`. Shopify tags compare case-insensitively. */
export function hasTag(node, tag) {
  const wanted = String(tag).toLowerCase();
  return Array.isArray(node?.tags) && node.tags.some((t) => String(t).toLowerCase() === wanted);
}

/**
 * The campaign card a reminder is about: the one the journal recorded, when
 * Shopify lists it under this customer; otherwise the customer's only card.
 * `cards` are this customer's campaign cards (findCampaignCards rows).
 * Returns { card } or { skip: { reason, detail } }.
 */
export function pickCard(cards, journalCardId = null) {
  if (!cards.length) {
    return { skip: { reason: 'no-card', detail: journalCardId ? `Shopify 上这位客户名下没有日志记录的卡 ${numericId(journalCardId)}` : null } };
  }
  const recorded = journalCardId ? cards.find((c) => c.id === journalCardId) : null;
  if (recorded) return { card: recorded };
  if (cards.length === 1) return { card: cards[0] };
  return { skip: { reason: 'multiple-cards', detail: `${cards.length} 张卡：${cards.map((c) => c.last4 || numericId(c.id)).join('、')}` } };
}

/**
 * The 'card-expired' skip when `card` has expired by the store-local date `today` (YYYY-MM-DD), or null.
 * Shopify still accepts a card on its expiry date: it is expired only after it.
 * Used when the run starts (cardSkip) and again right before each send (a run can cross midnight).
 */
export function expirySkip(card, today) {
  return card.expiresOn && card.expiresOn < today ? { reason: 'card-expired', detail: `到期日 ${card.expiresOn}` } : null;
}

/** Why `card` gets no reminder on the store-local date `today` (YYYY-MM-DD), or null. */
export function cardSkip(card, today) {
  if (!card.enabled) return { reason: 'card-disabled', detail: null };
  const expired = expirySkip(card, today);
  if (expired) return expired;
  // Any spend lowers the balance; a reminder goes only to untouched cards.
  if (card.balanceCents < card.amountCents) {
    return { reason: 'used', detail: `余额 ${formatUsd(card.balanceCents)} / 面额 ${formatUsd(card.amountCents)}` };
  }
  return null;
}

/**
 * Why a freshly re-read customer (a RefreshCustomers node; null = deleted) gets no reminder, or null.
 * `requireSubscribed: false` (test campaigns) leaves out the marketing-state check, as issue does;
 * a deleted customer or a missing / invalid email still gets no reminder.
 */
export function customerSkip(node, { requireSubscribed = true } = {}) {
  if (!node?.id) return { reason: 'customer-deleted', detail: null };
  const email = node.defaultEmailAddress;
  if (!email?.emailAddress || !email.validFormat) return { reason: 'no-email', detail: email?.emailAddress || null };
  if (requireSubscribed && email.marketingState !== 'SUBSCRIBED') return { reason: 'not-subscribed', detail: email.marketingState ?? 'NONE' };
  return null;
}

/** Chinese names of the template's copies (the same words as src/preview.js uses). */
export const STAGE_LABELS = Object.freeze({ first: '首封', remind1: '第一次提醒', remind2: '第二次提醒' });

/**
 * The copy the template shows in an email sent on the store-local `date`
 * (YYYY-MM-DD), by .env's reminder dates: 'first' | 'remind1' | 'remind2'.
 * The same rule as preview.expectedStage (kept here so that remind does not
 * load the template engine). null when a reminder date is missing from .env:
 * then .env cannot tell which copy the pasted template shows.
 */
export function templateStageOn(config, date) {
  if (!config.remind1Date || !config.remind2Date) return null;
  if (date >= config.remind2Date) return 'remind2';
  if (date >= config.remind1Date) return 'remind1';
  return 'first';
}

/** Refusal (live) / warning (dry run) for a round 1 on or after REMIND_2_DATE; also the mid-run stop. */
function round1TooLateText(remind2Date) {
  return `第 1 次提醒要在 REMIND_2_DATE（${remind2Date}）之前发：今天发出的邮件会显示第二次提醒的文案。请改发 --round 2。`;
}

/**
 * The warning for a live run of the real campaign that starts at instant `ms` when the UTC date
 * is already later than the store date (the store's evening; in Los Angeles from 17:00 PDT) and
 * the template would show another copy on that UTC date than on the store date (the same rule as
 * templateStageOn / preview.expectedStage). Shopify does not document whether the template's
 * 'now' is store time or UTC. null = nothing to warn about (also when .env lacks a reminder date).
 */
export function utcEveningWarning(config, ms, timeZone) {
  const storeDate = localDate(ms, timeZone);
  const utcDate = new Date(ms).toISOString().slice(0, 10);
  if (utcDate <= storeDate) return null;
  const utcStage = templateStageOn(config, utcDate);
  if (!utcStage || utcStage === templateStageOn(config, storeDate)) return null;
  return `现在店铺时间 ${localDateTime(ms, timeZone)}，UTC 已是 ${utcDate}：如果 Shopify 按 UTC 日期选邮件文案，`
    + `今天发出的提醒会显示${STAGE_LABELS[utcStage]}文案。建议在洛杉矶时间 17:00 之前运行。`;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const fmt = (n) => Number(n).toLocaleString('en-US');

function parseRound(round) {
  const n = typeof round === 'string' ? Number(round.trim() || NaN) : round;
  return n === 1 || n === 2 ? n : null;
}

/** 0 = no limit; a non-negative integer (or its string form); null = invalid. */
function parseLimit(limit) {
  if (limit === undefined || limit === null) return 0;
  const n = typeof limit === 'string' ? Number(limit.trim() || NaN) : limit;
  return Number.isInteger(n) && n >= 0 ? n : null;
}

function who(r) {
  return `#${r.seq} ${r.name || '(无姓名)'} <${r.email || '无邮箱'}>`;
}

/** The later of two YYYY-MM-DD dates ('' / null count as "not set"). */
function laterDate(a, b) {
  if (!a) return b;
  if (!b) return a;
  return a > b ? a : b;
}

/**
 * True when selection.json is no longer the list read before the run lock was
 * taken: select rewrote it (a new createdAt) or it disappeared / became unreadable.
 */
function selectionReplaced(paths, selection) {
  try {
    return readJson(paths.selection, null)?.createdAt !== selection.createdAt;
  } catch {
    return true;
  }
}

/** selection.json checked against the config: { selection } or { error } (a Chinese message). */
function loadSelection(paths, config) {
  let selection;
  try {
    selection = readJson(paths.selection, null);
  } catch (err) {
    return { error: err.message };
  }
  if (!selection) return { error: '还没有名单，请先运行 select' };
  if (!Array.isArray(selection.recipients) || !selection.params || !Number.isFinite(Date.parse(selection.createdAt))) {
    return { error: `名单文件 ${paths.selection} 不完整（缺少 recipients、params 或 createdAt）` };
  }
  if (selection.campaignId !== config.campaignId) {
    return { error: `名单属于活动 ${selection.campaignId}，和 .env 的 CAMPAIGN_ID=${config.campaignId} 不一致` };
  }
  if (selection.params.sentTag !== config.sentTag) {
    return { error: `名单的 SENT_TAG 是 ${selection.params.sentTag}，和 .env 的 SENT_TAG=${config.sentTag} 不一致` };
  }
  const tz = selection.params.timezone;
  try {
    if (!tz) throw new RangeError('missing');
    new Intl.DateTimeFormat('en-US', { timeZone: tz });
  } catch {
    return { error: `名单文件里的店铺时区无效：${tz || '（空）'}` };
  }
  return { selection };
}

function newSummary(round, dryRun) {
  return {
    round,
    dryRun,
    // passed every check this run (before --limit); a card found expired right before its send
    // (the run crossed midnight) leaves eligible and planned and is counted in skipped instead
    eligible: 0,
    planned: 0, // of those, how many this run sends (or, in a dry run, would send)
    attempted: 0, // sends actually tried
    sent: 0,
    failed: 0,
    rejected: 0,
    unknown: 0,
    retriedUnknown: 0, // sends that were --retry-unknown re-sends
    retriedFailed: 0, // sends that were --retry-failed re-sends
    skipped: {}, // reason → count, this run
    alreadySent: 0, // reminded in an earlier run of this round (the journal says so)
    // carry the round tag in Shopify while the journal shows no sent reminder this round (it was
    // lost or rolled back): counted as sent, not evaluated (live runs journal remind.found)
    alreadySentByTag: 0,
    previouslyFailed: 0, // userErrors in an earlier run of this round (not retried: no --retry-failed)
    waitingUnknown: 0, // outcome unknown, waiting for a human check and --retry-unknown
    notIssued: 0, // on the list but no card was issued to them
    roundTagged: 0, // round tags added right after this run's sends
    roundTagFailed: 0, // round tags this run could not add (after a send or in the repair): remind.tag.fail
    roundTagFixed: 0, // round tags added again at the start of a live run (sent earlier, tag missing in Shopify)
    roundTagMissing: 0, // dry runs only: people whose round tag a live run would add again
    stopped: null, // null | 'rejected' | 'too-many-failures' | 'aborted' | 'error'
    // null, or the store date (YYYY-MM-DD) a live round 1 of the real campaign reached
    // REMIND_2_DATE on during the run: it stopped there (exit 2)
    stoppedByDate: null,
  };
}

function interrupted(summary, log) {
  summary.stopped = 'aborted';
  log.warn('已中断：停止发送。已处理的人都记在日志里，重新运行会从还没处理的人继续');
  return 130;
}

async function defaultWriteReport(options) {
  const { writeReport } = await import('./report/excel.js');
  return writeReport(options);
}

async function defaultRenderPreview(options) {
  const { renderPreviewFor } = await import('./preview.js');
  return renderPreviewFor(options);
}

/**
 * Regenerate the Excel; a failure is only a warning, it never changes the command's result.
 * writeReport logs its own warnings (it is their source), so result.warnings is not printed again.
 */
async function refreshExcel(write, options, log) {
  try {
    const result = await write(options);
    if (result?.file) log.info(`Excel 已更新：${result.file}`);
  } catch (err) {
    log.warn(`Excel 没有更新：${err.message}（可以稍后运行 node index.js export 重新生成）`);
  }
}

/** About how long `sends` live sends take (SECONDS_PER_SEND each), in Chinese. */
function durationText(sends) {
  const seconds = sends * SECONDS_PER_SEND;
  if (seconds < 60) return '不到 1 分钟';
  const minutes = Math.round(seconds / 60);
  return minutes < 90 ? `约 ${minutes} 分钟` : `约 ${(seconds / 3600).toFixed(1)} 小时`;
}

/** The round-tag line of the final summary; a dry run tells what a live run would do. */
function roundTagSummary(s, tag) {
  if (s.dryRun) {
    return `本轮 tag ${tag}：按 tag 认定已发 ${fmt(s.alreadySentByTag)} 人；真实运行会打 tag ${fmt(s.planned)} 人（每发出一封打一次），补打 tag ${fmt(s.roundTagMissing)} 人`;
  }
  return `本轮 tag ${tag}：按 tag 认定已发 ${fmt(s.alreadySentByTag)} 人，打 tag ${fmt(s.roundTagged)} 人，打 tag 失败 ${fmt(s.roundTagFailed)} 人，补打 tag ${fmt(s.roundTagFixed)} 人`;
}

function printSummary(s, tag, log) {
  log.info(`—— 第 ${s.round} 次提醒${s.dryRun ? '（预演）' : ''}结果 ——`);
  if (s.dryRun) {
    log.info(`符合条件 ${fmt(s.eligible)} 人，本次会发 ${fmt(s.planned)} 人`);
  } else {
    log.info(`符合条件 ${fmt(s.eligible)} 人；本次发出 ${fmt(s.sent)} 封，失败 ${fmt(s.failed)}，结果不明 ${fmt(s.unknown)}，被拒未发 ${fmt(s.rejected)}`);
  }
  const skippedTotal = Object.values(s.skipped).reduce((a, b) => a + b, 0);
  if (skippedTotal) {
    const parts = Object.keys(REMIND_SKIP_REASONS).filter((k) => s.skipped[k]).map((k) => `${REMIND_SKIP_REASONS[k]} ${fmt(s.skipped[k])}`);
    log.info(`跳过 ${fmt(skippedTotal)} 人：${parts.join('，')}`);
  }
  if (s.alreadySent) log.info(`这一轮之前已经发过：${fmt(s.alreadySent)} 人（不会再发）`);
  log.info(roundTagSummary(s, tag));
  if (s.roundTagFailed) {
    log.warn(`打本轮 tag 失败 ${fmt(s.roundTagFailed)} 人：提醒已经发出，本地日志记着已发，不会重发；下次运行 remind --round ${s.round} 会补打 tag`);
  }
  if (s.retriedFailed) log.info(`本次重发了之前发送失败的 ${fmt(s.retriedFailed)} 人（--retry-failed）`);
  if (s.failed) log.warn(`本次发送失败 ${fmt(s.failed)} 人（Shopify 拒绝，邮件没有发出）：不会自动重试，原因解决后可用 --retry-failed 重发`);
  if (s.previouslyFailed) log.info(`这一轮之前发送失败：${fmt(s.previouslyFailed)} 人（不会自动重试，原因解决后可用 --retry-failed 重发）`);
  const unknownTotal = s.waitingUnknown + s.unknown;
  if (unknownTotal) {
    // A tag added by hand is read like one this tool added: the next live run records the row as sent.
    log.warn(`结果不明 ${fmt(unknownTotal)} 人：不会自动重发。确认对方确实没收到后，可加 --retry-unknown 补发；`
      + `确认已经收到的，可以在 Shopify 后台给他加上 tag ${tag}，下次运行会记为已发`);
  }
  if (s.notIssued) log.info(`名单里还没有建卡：${fmt(s.notIssued)} 人（不在提醒范围内）`);
  const remaining = s.eligible - s.sent - s.failed - s.unknown;
  if (!s.dryRun && remaining > 0) {
    // After a stop by date a re-run of this round is refused: do not promise that it continues.
    log.info(s.stoppedByDate
      ? `还有 ${fmt(remaining)} 人符合条件但没有收到第 ${s.round} 次提醒：店铺日期已是 ${s.stoppedByDate}，不能再发第 ${s.round} 次提醒`
      : `还有 ${fmt(remaining)} 人符合条件但这次没有发出，再运行一次会继续`);
  }
}

// ---------------------------------------------------------------------------
// The command
// ---------------------------------------------------------------------------

/**
 * Send (or, in a dry run, preview) reminder `round` to every recipient whose
 * campaign card is still unused.
 *
 * @param {object} a
 * @param {object} a.config loadConfig() result; config.dryRun selects the mode
 * @param {1|2|'1'|'2'} a.round
 * @param {number} [a.limit] maximum send attempts this run; 0 = no limit
 * @param {boolean} [a.retryUnknown] also re-send reminders whose outcome is unknown
 * @param {boolean} [a.retryFailed] also re-send reminders Shopify refused with userErrors (nothing was sent);
 *   they still pass every card and customer check
 * @param {{info: Function, warn: Function, error: Function}} [a.log]
 * @param {() => Date} [a.now]
 * @param {(ms: number) => Promise<void>} [a.sleep] used by the GraphQL client's throttle back-off
 * @param {Function} [a.writeReport] defaults to src/report/excel.js writeReport
 * @param {Function} [a.renderPreview] defaults to src/preview.js renderPreviewFor (dry runs only)
 * @param {AbortSignal} [a.signal] aborted by Ctrl+C: the current send finishes, then the run stops (exit 130)
 * @returns {Promise<{ exitCode: number, summary: object|null }>}
 *   exitCode 0 done · 1 stopped or failed · 2 usage error / date guard (also a live round 1 of the
 *   real campaign that reached REMIND_2_DATE during the run: summary.stoppedByDate) · 130 interrupted
 */
export async function runRemind({
  config,
  round,
  limit = 0,
  retryUnknown = false,
  retryFailed = false,
  log = console,
  now = () => new Date(),
  sleep,
  writeReport,
  renderPreview,
  signal,
} = {}) {
  const clock = () => {
    const value = now();
    return value instanceof Date ? value : new Date(value);
  };

  // 1. Arguments.
  const roundNo = parseRound(round);
  if (!roundNo) {
    log.error('--round 必须是 1 或 2');
    return { exitCode: 2, summary: null };
  }
  const max = parseLimit(limit);
  if (max === null) {
    log.error('--limit 必须是正整数');
    return { exitCode: 2, summary: null };
  }

  // 2. The frozen list.
  const paths = campaignPaths(config);
  const loaded = loadSelection(paths, config);
  if (loaded.error) {
    log.error(loaded.error);
    return { exitCode: 1, summary: null };
  }
  const { selection } = loaded;
  const dryRun = config.dryRun !== false; // like DRY_RUN itself: only an explicit false sends
  const testCampaign = selection.mode === 'test';
  const startMs = clock().getTime(); // one reading for every start-of-run date check
  const today = localDate(startMs, selection.params.timezone);

  // 3. Date guard: the template shows round n's copy from REMIND_<n>_DATE on,
  //    and from REMIND_2_DATE on round 1 would carry the round-2 copy.
  //    sendBatch checks the store date again before every send.
  const dateKey = `REMIND_${roundNo}_DATE`;
  const date = roundNo === 1 ? config.remind1Date : config.remind2Date;
  const round1TooLate = roundNo === 1 && Boolean(config.remind2Date) && today >= config.remind2Date;
  let utcWarning = null; // logged once, with the run's other opening lines
  if (!testCampaign && !dryRun) {
    if (!date) {
      log.error(`请在 .env 设置 ${dateKey}，格式 YYYY-MM-DD`);
      return { exitCode: 2, summary: null };
    }
    if (today < date) {
      log.error(`第 ${roundNo} 次提醒要到 ${date}（店铺时间）才能发`);
      log.info(`现在店铺时间是 ${today}。在那之前可以先预演（不加 DRY_RUN=false）`);
      return { exitCode: 2, summary: null };
    }
    if (round1TooLate) {
      log.error(round1TooLateText(config.remind2Date));
      log.info(`现在店铺时间是 ${today}`);
      return { exitCode: 2, summary: null };
    }
    // The store's evening: if Shopify picks the copy by the UTC date, today's emails may show the next one.
    utcWarning = utcEveningWarning(config, startMs, selection.params.timezone);
  } else if (!testCampaign) {
    if (!date) log.warn(`还没有设置 ${dateKey}：预演可以运行，真实发送前请在 .env 设置，格式 YYYY-MM-DD`);
    else if (today < date) log.warn(`注意：第 ${roundNo} 次提醒要到 ${date}（店铺时间）才能真实发送，现在是 ${today}`);
    if (round1TooLate) log.warn(`注意：${round1TooLateText(config.remind2Date)}`);
  } else {
    log.info('测试活动：不受提醒日期限制，也不检查营销邮件订阅状态（和 issue 一致）');
    // Which copy the test inbox really gets today (the template switches by the send date).
    const own = `remind${roundNo}`;
    const shown = templateStageOn(config, today);
    if (!shown) {
      log.warn('注意：邮件模板按发送当天的日期切换文案；.env 没有同时设置 REMIND_1_DATE 和 REMIND_2_DATE，无法判断今天发出的是哪种文案，请先用 preview 查看');
    } else if (shown === own) {
      log.info(`按 .env 的日期，今天（${today}）发出的是${STAGE_LABELS[shown]}文案`);
    } else {
      log.warn(`注意：邮件模板按发送当天的日期切换文案，按 .env 的日期，今天（${today}）发出的是${STAGE_LABELS[shown]}文案，不是${STAGE_LABELS[own]}文案；${STAGE_LABELS[own]}的文案请先用 preview 查看`);
    }
  }

  // 4. One write command at a time; then make sure the list read above is still the one on disk.
  let release;
  try {
    release = acquireRunLock(paths, 'remind');
  } catch (err) {
    log.error(err.message);
    return { exitCode: 1, summary: null };
  }
  if (selectionReplaced(paths, selection)) {
    release();
    log.error(SELECTION_REPLACED);
    return { exitCode: 1, summary: null };
  }

  const summary = newSummary(roundNo, dryRun);
  const tag = roundTag(config.sentTag, roundNo);
  // `at` (a Date) stamps the entry with an instant already read, e.g. the one a send was checked at.
  const journal = (entry, at = null) => appendJournal(paths.journal, entry, { now: () => (at ?? clock()).toISOString() });
  let exitCode = 0;
  let runId = null;
  try {
    await connect(config, { log, sleep });
    const state = foldJournal(readJournal(paths.journal));
    runId = newRunId(clock());
    journal({
      op: 'run.start',
      run: runId,
      command: 'remind',
      dryRun,
      batch: null,
      limit: max || null,
      options: { round: roundNo, retryUnknown: Boolean(retryUnknown), retryFailed: Boolean(retryFailed) },
    });
    log.info(`第 ${roundNo} 次提醒 · 活动 ${config.campaignId}${testCampaign ? '（测试活动）' : ''} · 店铺日期 ${today}`);
    log.info(dryRun
      ? '预演：只读取 Shopify，不发邮件，也不改任何人的状态（真实发送请在命令前加 DRY_RUN=false）'
      : '真实发送：Shopify 会给每位符合条件的客户重发一次礼品卡邮件');
    if (utcWarning) log.warn(utcWarning);
    const retrying = [retryUnknown && '结果不明的人（--retry-unknown）', retryFailed && '之前发送失败的人（--retry-failed）'].filter(Boolean);
    if (retrying.length) log.info(`这次也会重发：${retrying.join('、')}；他们同样要先通过卡和客户的检查`);
    exitCode = await remindRound({
      config, paths, selection, state, round: roundNo, tag, today, dryRun, retryUnknown: Boolean(retryUnknown), retryFailed: Boolean(retryFailed), max,
      log, signal, summary, journal, runId, clock, renderPreview: renderPreview ?? defaultRenderPreview,
    });
  } catch (err) {
    exitCode = 1;
    summary.stopped = summary.stopped ?? 'error';
    log.error(`出错，已停止：${err.message}`);
  } finally {
    if (runId) {
      try {
        journal({ op: 'run.end', run: runId, summary, exitCode });
      } catch (err) {
        log.error(`写运行记录失败：${err.message}`);
        if (exitCode === 0) exitCode = 1;
      }
    }
    release();
  }

  if (runId) {
    printSummary(summary, tag, log);
    await refreshExcel(writeReport ?? defaultWriteReport, { config, paths, log, now: clock }, log);
  }
  return { exitCode, summary };
}

/**
 * Steps 5–9: read who carries the round tag, add missing round tags, find the cards, decide
 * who is eligible, then send (or preview). Returns the exit code.
 */
async function remindRound(ctx) {
  const { config, selection, state, round, tag, today, dryRun, retryUnknown, retryFailed, max, log, signal, summary, journal, runId, clock } = ctx;
  const roundKey = String(round);
  // Test campaigns skip the audience rules (as issue does): no marketing-state check.
  const requireSubscribed = selection.mode !== 'test';
  if (signal?.aborted) return interrupted(summary, log);
  const recipients = [...selection.recipients].sort((a, b) => a.seq - b.seq);

  // 5. Shopify's record of this round, which survives a lost or rolled-back journal: the round
  //    tag goes on every customer right after their reminder is sent.
  const tagReadMs = clock().getTime();
  const roundTagged = await fetchTaggedCustomerIds(tag); // no log: its own line is English
  log.info(`本轮 tag：${tag}（每发出一封提醒就给客户打上；带这个 tag 的人这一轮不会再收到提醒）。Shopify 上现在有 ${fmt(roundTagged.size)} 位客户带这个 tag`);

  /**
   * Add the round tag to recipient `r`. A failure is journalled (remind.tag.fail) and warned about,
   * never fatal and never counted toward the consecutive-failure stop: the journal's remind.ok is
   * still the first guard, and the next live run adds the tag again. Returns true once it is added.
   */
  const tagRound = async (r, afterSend) => {
    try {
      await addTag(r.customerId, tag);
      return true;
    } catch (err) {
      const error = err?.message ?? String(err);
      journal({ op: 'remind.tag.fail', cid: r.customerId, round, error, run: runId });
      summary.roundTagFailed += 1;
      // userErrors (e.g. the customer's tag limit) will not go away by themselves: promise no repair.
      const next = err?.code === 'USER_ERROR' ? '下次运行会再试，如果仍被拒请在后台查看这位客户' : '下次运行会补打 tag';
      log.warn(`  ${who(r)}：${afterSend ? '提醒已发出，但' : ''}打本轮 tag ${tag} 失败（${error}）；本地日志记着已发，不会重发，${next}`);
      return false;
    }
  };

  // 6. Sent earlier, but Shopify does not show the tag: a live run adds it again.
  const repairStop = await repairRoundTags({ ...ctx, recipients, roundTagged, tagReadMs, tagRound });
  if (repairStop !== null) return repairStop;

  // 7. Every campaign card with its current balance, grouped by customer.
  log.info('正在查询本活动的礼品卡和余额…');
  const sinceIso = new Date(Date.parse(selection.createdAt) - CARD_WINDOW_MS).toISOString();
  const cards = await findCampaignCards({ campaignId: config.campaignId, sinceIso });
  const cardsByCustomer = new Map();
  for (const card of cards) {
    if (!card.customerId) continue;
    if (!cardsByCustomer.has(card.customerId)) cardsByCustomer.set(card.customerId, []);
    cardsByCustomer.get(card.customerId).push(card);
  }
  const onList = new Set(recipients.map((r) => r.customerId));
  const foreign = cards.filter((c) => !c.customerId || !onList.has(c.customerId)).length;
  log.info(`Shopify 上有 ${fmt(cards.length)} 张本活动的卡`);
  if (foreign) log.warn(`${fmt(foreign)} 张本活动的卡不属于名单里的客户，不发提醒（可运行 verify 核对）`);

  // A skip is not final: a later run of this round evaluates the person again.
  // Live runs journal it, unless the previous run recorded exactly the same skip.
  const skip = (r, s, prev) => {
    summary.skipped[s.reason] = (summary.skipped[s.reason] ?? 0) + 1;
    if (dryRun) return;
    const detail = s.detail ?? null;
    if (prev?.status === REMIND_STATUS.SKIPPED && prev.reason === s.reason && (prev.detail ?? null) === detail) return;
    journal({ op: 'remind.skip', cid: r.customerId, round, reason: s.reason, ...(detail ? { detail } : {}), run: runId });
  };

  // 8a. Who is still due a reminder this round, and does their card qualify?
  const toCheck = [];
  let interruptedSends = 0;
  for (const r of recipients) {
    const cid = r.customerId;
    const cs = state.customers.get(cid);
    const prev = cs?.reminders?.[roundKey] ?? null;
    if (prev?.status === REMIND_STATUS.SENT) {
      summary.alreadySent += 1;
      continue;
    }
    // The round tag says this round was sent, whatever the journal holds (nothing: it was lost or
    // replaced by an older copy; or an unknown, failed, skipped or interrupted row). Never sent
    // again, with --retry-unknown / --retry-failed neither, and not even evaluated.
    if (roundTagged.has(cid)) {
      summary.alreadySentByTag += 1;
      if (!dryRun) journal({ op: 'remind.found', cid, round, source: 'tag', run: runId });
      continue;
    }
    // userErrors: Shopify sent nothing, so a re-send cannot duplicate an email; only --retry-failed does it.
    const failedBefore = prev?.status === REMIND_STATUS.FAILED;
    if (failedBefore && !retryFailed) {
      summary.previouslyFailed += 1;
      continue;
    }
    const unresolved = prev?.status === REMIND_STATUS.UNKNOWN || prev?.status === REMIND_STATUS.IN_PROGRESS;
    if (prev?.status === REMIND_STATUS.IN_PROGRESS && !dryRun) {
      // We hold the run lock, so no send is in flight: the run that wrote
      // remind.start died before learning the outcome. Record it as unknown.
      journal({ op: 'remind.unknown', cid, round, error: INTERRUPTED_ERROR, run: runId });
      interruptedSends += 1;
    }
    if (unresolved && !retryUnknown) {
      summary.waitingUnknown += 1;
      continue;
    }
    const shopifyCards = cardsByCustomer.get(cid) ?? [];
    const journalCardId = cs?.giftCardId ?? null;
    if (!journalCardId && !shopifyCards.length && !prev) {
      summary.notIssued += 1; // never got a card: not part of the reminder round
      continue;
    }
    const picked = pickCard(shopifyCards, journalCardId);
    const problem = picked.skip ?? cardSkip(picked.card, today);
    if (problem) {
      skip(r, problem, prev);
      continue;
    }
    // retry: null (a first attempt this round) | 'unknown' (--retry-unknown) | 'failed' (--retry-failed)
    toCheck.push({ recipient: r, card: picked.card, retry: unresolved ? 'unknown' : failedBefore ? 'failed' : null, prev });
  }
  if (interruptedSends) {
    log.warn(`${fmt(interruptedSends)} 人上次发送提醒时中断，不知道是否已发出，已标为“结果不明”`);
  }

  // 8b. Re-read those customers: deleted, email gone or invalid, no longer subscribed.
  const eligible = [];
  if (toCheck.length) log.info(`正在复查 ${fmt(toCheck.length)} 位客户的邮箱${requireSubscribed ? '和订阅状态' : ''}…`);
  let chunkSize = REFRESH_CHUNK;
  for (let i = 0; i < toCheck.length;) {
    if (signal?.aborted) return interrupted(summary, log);
    const chunk = toCheck.slice(i, i + chunkSize);
    const ids = chunk.map((x) => x.recipient.customerId);
    let data;
    try {
      data = await gql(REFRESH_CUSTOMERS, { ids });
    } catch (err) {
      // A read, so it is safe to repeat: if Shopify finds the request too
      // expensive for a single query, ask for fewer customers at a time.
      if (err?.code === 'MAX_COST_EXCEEDED' && chunkSize > 1) {
        chunkSize = Math.floor(chunkSize / 2);
        log.info(`  一次复查的客户超过了 Shopify 单次查询的上限，改为每次 ${chunkSize} 位`);
        continue;
      }
      throw err;
    }
    i += chunk.length;
    const nodes = data?.nodes;
    if (!Array.isArray(nodes) || nodes.length !== ids.length) {
      throw new Error(`复查客户时 Shopify 返回了 ${Array.isArray(nodes) ? nodes.length : 0} 条结果，应为 ${ids.length} 条`);
    }
    chunk.forEach((x, j) => {
      const node = nodes[j];
      if (node?.id && node.id !== ids[j]) throw new Error(`复查客户时 Shopify 返回的顺序不对（${ids[j]}）`);
      // The by-id read shows the customer's current tags even while the tag search lags
      // (a tag added minutes ago, e.g. by hand in the admin): the round tag means sent.
      if (hasTag(node, tag)) {
        summary.alreadySentByTag += 1;
        if (!dryRun) journal({ op: 'remind.found', cid: x.recipient.customerId, round, source: 'tag', run: runId });
        return;
      }
      const problem = customerSkip(node, { requireSubscribed });
      if (problem) skip(x.recipient, problem, x.prev);
      else eligible.push(x);
    });
  }
  if (summary.alreadySentByTag) {
    log.warn(`${dryRun ? '[预演] ' : ''}按 tag 认定已发 ${fmt(summary.alreadySentByTag)} 人：他们在 Shopify 上带本轮 tag ${tag}，`
      + '本地日志里却没有这一轮发出的记录（日志丢失、被旧的备份覆盖，或在后台手动打了 tag）；'
      + `${dryRun ? '真实运行会在本地日志补记为已发' : '已在本地日志补记为已发'}，这一轮不会再给他们发`);
  }

  summary.eligible = eligible.length;
  const batch = max ? eligible.slice(0, max) : eligible;
  summary.planned = batch.length;
  if (signal?.aborted) return interrupted(summary, log);

  if (dryRun) {
    await previewBatch({ ...ctx, batch });
    return 0;
  }
  return sendBatch({ ...ctx, batch, skip, tagRound });
}

/**
 * 6. People the journal shows as reminded this round by this tool (remind.ok) whom Shopify does
 * not list with the round tag: their tagsAdd failed, or the tag was removed. A live run adds it
 * again: a row whose tag is known to have failed (tagError) at once, any other row once its send
 * is at least TAG_SEARCH_LAG_MS older than the tag set (until then the search may simply not show
 * the tag yet). A row recorded as failed whose tag Shopify does list after all (an answer lost
 * after Shopify applied it) is journalled remind.tag.ok, so the Excel stops calling it missing.
 * Rows found by the tag (remind.found) are not repaired. Dry runs only count (roundTagMissing).
 * Returns an exit code when interrupted, otherwise null.
 */
async function repairRoundTags({ recipients, state, round, tag, roundTagged, tagReadMs, tagRound, dryRun, log, signal, summary, journal, runId }) {
  const key = String(round);
  const missing = [];
  const confirmed = [];
  for (const r of recipients) {
    const prev = state.customers.get(r.customerId)?.reminders?.[key];
    if (prev?.status !== REMIND_STATUS.SENT || prev.source === 'tag' || prev.tagSettled) continue;
    if (roundTagged.has(r.customerId)) {
      if (prev.tagError) confirmed.push(r);
      continue;
    }
    const sentMs = Date.parse(prev.at ?? '');
    if (prev.tagError || (Number.isFinite(sentMs) && tagReadMs - sentMs >= TAG_SEARCH_LAG_MS)) missing.push(r);
  }
  if (dryRun) {
    summary.roundTagMissing = missing.length;
    if (missing.length) {
      log.info(`[预演] 缺本轮 tag ${fmt(missing.length)} 人：本地日志记着这一轮已发，Shopify 上却没有 tag ${tag}；真实运行会先给他们补打 tag`);
    }
    return null;
  }
  for (const r of confirmed) journal({ op: 'remind.tag.ok', cid: r.customerId, round, note: 'already tagged in Shopify', run: runId });
  if (!missing.length) return null;

  // The tag search may lag or never list someone (a deleted customer): re-read them by id first.
  const nodes = await readCustomersById(missing.map((r) => r.customerId));
  const toTag = [];
  for (const r of missing) {
    const node = nodes.get(r.customerId) ?? null;
    if (!node) {
      // Deleted after the reminder: the tag can never be added; settle it once instead of retrying forever.
      journal({ op: 'remind.tag.ok', cid: r.customerId, round, note: 'customer deleted', run: runId });
    } else if (hasTag(node, tag)) {
      journal({ op: 'remind.tag.ok', cid: r.customerId, round, note: 'already tagged in Shopify', run: runId });
    } else {
      toTag.push(r);
    }
  }
  if (!toTag.length) return null;
  log.info(`补打本轮 tag：${fmt(toTag.length)} 人这一轮已发（见本地日志），Shopify 上还没有 tag ${tag}`);
  for (const r of toTag) {
    if (signal?.aborted) return interrupted(summary, log);
    if (await tagRound(r, false)) {
      journal({ op: 'remind.tag.ok', cid: r.customerId, round, run: runId });
      summary.roundTagFixed += 1;
    }
  }
  return null;
}

/** RefreshCustomers for any number of ids, REFRESH_CHUNK at a time → Map<gid, node>; deleted customers are absent. */
async function readCustomersById(ids) {
  const byId = new Map();
  for (let i = 0; i < ids.length; i += REFRESH_CHUNK) {
    const data = await gql(REFRESH_CUSTOMERS, { ids: ids.slice(i, i + REFRESH_CHUNK) });
    for (const node of data?.nodes ?? []) if (node?.id) byId.set(node.id, node);
  }
  return byId;
}

const RETRY_NOTES = Object.freeze({ unknown: '（结果不明，补发）', failed: '（之前发送失败，重发）' });

/** 9. Dry run: who this run would remind, plus a local preview of the email. */
async function previewBatch({ config, paths, selection, round, today, max, log, summary, batch, renderPreview }) {
  const retriedUnknown = batch.filter((x) => x.retry === 'unknown').length;
  const retriedFailed = batch.filter((x) => x.retry === 'failed').length;
  const retries = [
    retriedUnknown && `${fmt(retriedUnknown)} 人是结果不明的补发`,
    retriedFailed && `${fmt(retriedFailed)} 人是之前发送失败的重发`,
  ].filter(Boolean);
  const range = batch.length ? `，序号 ${batch[0].recipient.seq}–${batch[batch.length - 1].recipient.seq}` : '';
  log.info(`[预演] 符合条件 ${fmt(summary.eligible)} 人；本次${max ? `（--limit ${max}）` : ''}会发 ${fmt(batch.length)} 人${range}`
    + `${retries.length ? `（其中 ${retries.join('，')}）` : ''}`);
  if (batch.length) {
    log.info(batch.length > LIST_FIRST ? `[预演] 本次会发的前 ${LIST_FIRST} 人：` : `[预演] 本次会发的 ${batch.length} 人：`);
    for (const { recipient: r, card, retry } of batch.slice(0, LIST_FIRST)) {
      log.info(`  ${who(r)}  卡尾号 ${card.last4 || numericId(card.id)}  面额 ${formatUsd(card.amountCents)}${retry ? `  ${RETRY_NOTES[retry]}` : ''}`);
    }
    log.info(`[预演] 真实发送预计${durationText(batch.length)}（每封约 ${SECONDS_PER_SEND} 秒：发信、打本轮 tag、写 2 行日志）`);
  }
  const variant = round === 1 ? 'remind1' : 'remind2';
  // The email the batch would really get: a live send happens today, and never
  // before the round's date (test campaigns send today, whatever the date).
  const roundDate = round === 1 ? config.remind1Date : config.remind2Date;
  const sendDate = selection.mode === 'test' ? today : laterDate(today, roundDate);
  try {
    const preview = await renderPreview({
      config, paths, variant, recipient: batch[0]?.recipient ?? null, log,
      sendDate,
      expiresOn: selection.params.giftCardExpiresOn, // the cards carry the selection's expiry, not .env's
    });
    if (preview?.file) log.info(`邮件预览（按 ${sendDate} 发送时的文案）：${preview.file}${preview.subject ? `（主题：${preview.subject}）` : ''}`);
  } catch (err) {
    log.warn(`邮件预览没有生成：${err.message}`);
  }
}

/**
 * 9. Live run: one send at a time, journalled before and after, each followed by the round tag.
 * Returns the exit code.
 * `skip(recipient, { reason, detail }, prev)` counts and journals a skip (remindRound's rule);
 * `tagRound(recipient, afterSend)` adds the round tag (remindRound's helper; a failure is only a warning).
 */
async function sendBatch({ config, selection, round, max, log, signal, summary, journal, runId, batch, clock, skip, tagRound }) {
  if (!batch.length) return 0;
  const tz = selection.params.timezone;
  // A long run can cross midnight (store time) and the template picks its copy by the day each
  // email goes out, so the store date is computed again before every send. Round 1 of the real
  // campaign stops on REMIND_2_DATE (that email would show the round-2 copy); test campaigns are
  // not date-gated. In every round a card that has expired meanwhile is skipped.
  const round1Until = round === 1 && selection.mode !== 'test' ? config.remind2Date || null : null;
  log.info(`开始发送第 ${round} 次提醒：本次 ${fmt(batch.length)} 人${max ? `（--limit ${max}）` : ''}，预计${durationText(batch.length)}`);
  let consecutive = 0;
  let expiredSeen = false;
  for (const { recipient: r, card, retry, prev } of batch) {
    if (signal?.aborted) return interrupted(summary, log);
    const cid = r.customerId;
    // One clock reading for the check and for remind.start: nothing reads the clock between
    // the check and the send, and the journal shows the very instant that was checked.
    const at = clock();
    const day = localDate(at.getTime(), tz);
    if (round1Until && day >= round1Until) {
      // Nothing is undone: everyone handled so far is in the journal; this person has no remind.start.
      summary.stoppedByDate = day;
      log.error(round1TooLateText(round1Until));
      log.info(`现在店铺时间是 ${localDateTime(at.getTime(), tz)}，已停止发送；已处理的人都记在本地日志里`);
      return 2;
    }
    const expired = expirySkip(card, day);
    if (expired) {
      // As if found expired at the start: no longer eligible this run, counted as skipped.
      summary.eligible -= 1;
      summary.planned -= 1;
      if (!expiredSeen) {
        expiredSeen = true;
        log.warn(`现在店铺日期是 ${day}：${who(r)} 的卡已过期（到期日 ${card.expiresOn}），不发提醒；之后卡已过期的人同样跳过`);
      }
      skip(r, expired, prev);
      continue;
    }
    // On disk BEFORE the send: if we die mid-call, the next run knows a reminder may have gone out.
    journal({ op: 'remind.start', cid, round, giftCardId: card.id, ...(retry ? { retry: true } : {}), run: runId }, at);
    summary.attempted += 1;
    if (retry === 'unknown') summary.retriedUnknown += 1;
    else if (retry === 'failed') summary.retriedFailed += 1;

    let failure = null;
    try {
      await sendGiftCardNotification(card.id);
    } catch (err) {
      failure = err;
    }

    if (!failure) {
      journal({ op: 'remind.ok', cid, round, run: runId });
      summary.sent += 1;
      consecutive = 0; // a failed tag below never counts toward the stop
      // Shopify's own record of this send; no journal line when it works (remind.ok is the record).
      if (await tagRound(r, true)) summary.roundTagged += 1;
    } else {
      const error = failure?.message ?? String(failure);
      if (failure?.code === 'USER_ERROR') {
        // Shopify refused this send (userErrors): not retried automatically; --retry-failed re-sends.
        journal({ op: 'remind.fail', cid, round, error, run: runId });
        summary.failed += 1;
        log.warn(`  ${who(r)} 发送失败：${error}`);
      } else if (failure?.outcomeKnown === true) {
        // Definitely not sent (e.g. throttled after every retry): the next run may try again.
        journal({ op: 'remind.rejected', cid, round, error, run: runId });
        summary.rejected += 1;
        summary.stopped = 'rejected';
        log.error(`  ${who(r)}：Shopify 没有接受这次发送（${error}），这封没有发出。已停止本次运行，稍后重新运行会再给这个人发`);
        return 1;
      } else {
        // The email may or may not have gone out. Never re-sent automatically.
        journal({ op: 'remind.unknown', cid, round, error, run: runId });
        summary.unknown += 1;
        log.warn(`  ${who(r)} 结果不明（${error}）：不会自动重发；确认对方没收到后可用 --retry-unknown 补发`);
      }
      consecutive += 1;
      if (consecutive >= MAX_CONSECUTIVE_FAILURES) {
        summary.stopped = 'too-many-failures';
        log.error(`连续 ${MAX_CONSECUTIVE_FAILURES} 次没有发送成功，已停止。请查明原因后再运行：被 Shopify 拒绝（失败）的人，原因解决后可用 --retry-failed 重发`);
        return 1;
      }
    }

    if (summary.attempted % PROGRESS_EVERY === 0) {
      log.info(`  进度 ${fmt(summary.attempted)}/${fmt(batch.length)}：已发 ${fmt(summary.sent)}，失败 ${fmt(summary.failed)}，结果不明 ${fmt(summary.unknown)}`
        + `${summary.roundTagFailed ? `，打本轮 tag 失败 ${fmt(summary.roundTagFailed)}` : ''}`);
    }
  }
  return 0;
}
