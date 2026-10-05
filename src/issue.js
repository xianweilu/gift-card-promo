// `issue` — step 2 of the campaign: create one promotional gift card per
// selected customer, in the frozen selection order, then add SENT_TAG.
//
// Creating the card IS sending the first email: the card carries customerId,
// so Shopify emails its "Gift card created" notification as part of
// giftCardCreate (see src/giftcards.js). There is no separate send step.
//
// Never two cards for one customer. Every create is guarded by
//   1. the journal: `create.start` is fsynced BEFORE giftCardCreate, so a
//      customer with any record is not "pending" again until proven safe;
//   2. SENT_TAG in Shopify, fetched once before the batch;
//   3. this campaign's cards in Shopify (their note carries [campaign:<id>]);
//   4. a fresh read of each customer right before their card is created;
//   5. the run lock (one write command at a time) and the frozen selection.
// When Shopify's answer is lost (network error, 5xx) the request is never
// retried blindly: the card is looked up by customer and time instead. If it
// cannot be found the run stops; a later run gives the row back to the queue
// only once UNKNOWN_SETTLE (10 minutes) has passed without the card showing up
// (Shopify's search index can lag by seconds to minutes).
//
// Dry runs (the default) read everything a real run reads, but write nothing
// to Shopify and only run.start/run.end to the journal; every state change is
// simulated in memory so the preview shows exactly what a real run would do.
//
// Dates: the notification template picks its copy by the day it is sent, so a
// real run of the real campaign creates cards from LAUNCH_DATE until the day
// before REMIND_1_DATE (from then on the creation email would show reminder
// copy). The store's date is checked when the run starts and again right
// before every card, so a run that crosses midnight into REMIND_1_DATE stops
// there. From REMIND_1_DATE on, a run (and its dry run, which follows the same
// rule) only does the repairs that send nothing: settling open outcomes,
// adding missing tags, recording cards Shopify already has. Test campaigns are
// not date-gated; a dry run before LAUNCH_DATE only warns.
//
// --repair-only does those repairs on any day and never creates a card (no
// --limit needed): for a campaign that is being given up.

import {
  STATUS,
  LockError,
  acquireRunLock,
  appendJournal,
  campaignPaths,
  foldJournal,
  newRunId,
  readJournal,
  readJson,
} from './campaign.js';
import { connect } from './connect.js';
import { addTag, fetchTaggedCustomerIds } from './customers.js';
import { campaignNote, createGiftCard, findCampaignCards } from './giftcards.js';
import { ORDERS_SINCE, REFRESH_CUSTOMERS } from './queries.js';
import { formatUsd } from './select/amount.js';
import { buildActivity, isValidOrder, parseCustomer } from './select/rules.js';
import { defaultSleep, gql } from './shopify.js';
import { localDate, localDateTime } from './time.js';

const CHUNK_SIZE = 250; // nodes(ids:) accepts at most 250 ids per request
const MAX_CONSECUTIVE_FAILURES = 5;
const PROGRESS_EVERY = 50;
const CLOCK_SKEW_MS = 60_000; // tolerance between our clock (journal `t`) and Shopify's createdAt
const SELECTION_MARGIN_MS = 3_600_000; // campaign cards are searched from 1 h before the selection was made
const DRY_RUN_SKIP_LIST_MAX = 100;

/** Pre-flight skip reasons (journal `skip.reason`) and how they are shown to people. */
export const SKIP_REASONS = Object.freeze({
  'customer-deleted': '客户已被删除',
  'no-email': '没有邮箱或邮箱格式无效',
  'relay-email': '邮箱域名在排除名单里',
  'not-subscribed': '未订阅邮件营销',
  'already-tagged': '已带发放 tag',
  'ordered-since-snapshot': '名单导出后下过单',
  'address-ordered-since-snapshot': '同地址的账户在名单导出后下过单',
});

export const SELECTION_REPLACED = '名单刚被 select 改写，请重新运行';

/** Refusal (real run) / warning (dry run) on or after REMIND_1_DATE. */
function tooLateText(remind1Date) {
  return `正式活动要在 REMIND_1_DATE（${remind1Date}）之前建卡：今天建卡时 Shopify 发出的首封邮件会显示提醒的文案。`
    + '确需补发，请先改 .env 的提醒日期、运行 node index.js preview，并把模板重新贴到 Shopify 后台。';
}

/** Said instead of "will be issued again" when a row settled as not created can no longer get a card. */
function noNewCardsText(remind1Date) {
  return `正式活动从 REMIND_1_DATE（${remind1Date}）起不再建新卡`;
}

/** Said instead of "will be issued again" in an `issue --repair-only` run (a later normal run does issue it). */
const REPAIR_ONLY_NO_REISSUE = '本次是 --repair-only，不建新卡；之后正常运行 issue 时才会给他建卡';

/** Last line of an `issue --repair-only` run. */
const REPAIR_ONLY_TEXT = '只做了补记和补打 tag（--repair-only），没有建新卡。';
const REPAIR_ONLY_DRY_TEXT = '预演：真实运行只会补记和补打 tag（--repair-only），不会建新卡。';
const REPAIR_ONLY_CONFLICT = '--repair-only 和 --retry-failed 不能一起用：--repair-only 只补记和补打 tag，不建卡';

function tagRetryText(n) {
  return `有 ${n} 人补打 tag 失败，下次运行会再试`;
}

/**
 * True when the real campaign creates no new card on the store-local `date`
 * (YYYY-MM-DD): from REMIND_1_DATE on, the creation email would show reminder copy.
 * Test campaigns are never date-gated.
 */
function newCardsBlockedOn(config, selection, date) {
  return selection.mode !== 'test' && Boolean(config.remind1Date) && date >= config.remind1Date;
}

/** Chinese names of the template's copies (the same words as src/preview.js uses). */
const STAGE_LABELS = Object.freeze({ first: '首封', remind1: '第一次提醒', remind2: '第二次提醒' });

/**
 * The copy the "Gift card created" template shows in a first email sent on
 * `date` (YYYY-MM-DD), by .env's reminder dates: 'first' | 'remind1' | 'remind2'.
 * The rule of preview.expectedStage(config, 'first', date), kept here so that
 * issue does not load the template engine (without both reminder dates the
 * first email's own copy is assumed, as there).
 */
export function firstEmailStageOn(config, date) {
  if (!config.remind1Date || !config.remind2Date) return 'first';
  if (date >= config.remind2Date) return 'remind2';
  if (date >= config.remind1Date) return 'remind1';
  return 'first';
}

// Lazy imports: exceljs is heavy, and preview.js is optional for this command.
const defaultWriteReport = (options) => import('./report/excel.js').then((m) => m.writeReport(options));
const defaultRenderPreview = (options) => import('./preview.js').then((m) => m.renderPreviewFor(options));

/** Unwinds the run with a specific exit code; the reason has already been logged. */
class Stop extends Error {
  constructor(exitCode) {
    super(`issue stopped with exit code ${exitCode}`);
    this.name = 'Stop';
    this.exitCode = exitCode;
  }
}

/**
 * Run the `issue` command.
 *
 * @param {object} o
 * @param {object} o.config loadConfig() result; `dryRun` false means live
 * @param {number} [o.limit] cards to attempt this run (required live, except with repairOnly; ≤ ISSUE_MAX_PER_RUN). Skipped people do not count.
 * @param {boolean} [o.retryFailed] only retry rows whose create was rejected with userErrors
 * @param {boolean} [o.repairOnly] only the repairs (open outcomes, missing tags, cards Shopify already has); never creates a card
 * @param {AbortSignal} [o.signal] Ctrl+C: the current customer is finished, then the run stops (exit 130)
 * @param {number} [o.reconcileWaitMs] how long to look for a card whose create answer was lost
 * @param {number} [o.reconcilePollMs] pause between those lookups
 * @param {number} [o.unknownSettleMs] after this long without the card showing up it is taken as not created
 * @returns {Promise<{ exitCode: number, summary: object }>}
 */
export async function runIssue({
  config,
  limit,
  retryFailed = false,
  repairOnly = false,
  log = console,
  now = () => new Date(),
  sleep = defaultSleep,
  writeReport = defaultWriteReport,
  renderPreview = defaultRenderPreview,
  signal,
  reconcileWaitMs = 120_000,
  reconcilePollMs = 10_000,
  unknownSettleMs = 600_000,
} = {}) {
  const dryRun = config.dryRun !== false; // only an explicit false is a real run
  repairOnly = !!repairOnly;
  const summary = emptySummary(dryRun);
  if (repairOnly) summary.repairOnly = true;
  const done = (exitCode) => ({ exitCode, summary });

  // 1. Arguments, before anything else. 0 is treated like "not given" (a CLI default).
  const limitGiven = limit !== undefined && limit !== null && limit !== 0;
  if (limitGiven && !(Number.isInteger(limit) && limit >= 1)) {
    log.error(`--limit 必须是正整数，例如 --limit 20（收到 ${String(limit)}）`);
    return done(2);
  }
  if (repairOnly && retryFailed) {
    log.error(REPAIR_ONLY_CONFLICT);
    return done(2);
  }
  // A repair-only run creates no card, so it needs no --limit.
  if (!dryRun && !limitGiven && !repairOnly) {
    log.error('实时运行必须给 --limit N，例如 DRY_RUN=false node index.js issue --limit 20');
    return done(2);
  }
  if (limitGiven && limit > config.issueMaxPerRun) {
    log.error(`--limit ${limit} 超过每次运行的上限 ISSUE_MAX_PER_RUN=${config.issueMaxPerRun}`);
    return done(2);
  }

  // 2. The frozen selection, and that it belongs to this configuration.
  const paths = campaignPaths(config);
  let selection;
  try {
    selection = readJson(paths.selection, null);
  } catch (err) {
    log.error(err.message);
    return done(1);
  }
  if (!selection) {
    log.error('还没有名单，请先运行 select');
    return done(1);
  }
  const problem = selectionProblem(selection, config);
  if (problem) {
    log.error(problem);
    return done(1);
  }
  const tz = selection.params.timezone;

  // 3. Date guards, for the real campaign only: test campaigns may run any day.
  //    LAUNCH_DATE refuses real runs that may create cards; dry runs and repair-only runs (which
  //    send nothing) may run before it. From REMIND_1_DATE on, the creation email Shopify sends
  //    would show the reminder copy, so no NEW card is created, in a real run or its dry run;
  //    repairs that send nothing (settling leftovers, adding missing tags, recording cards Shopify
  //    already has) still run. walk() checks the date again before every card.
  const startMs = now().getTime();
  const today = localDate(startMs, tz);
  if (!dryRun && !repairOnly && selection.mode !== 'test' && config.launchDate && today < config.launchDate) {
    log.error(`正式活动要到 ${config.launchDate}（店铺时间）才能建卡发信。现在是店铺时间 ${localDateTime(startMs, tz)}；测试活动不受这个限制`);
    return done(2);
  }
  const dateBlocked = newCardsBlockedOn(config, selection, today);
  warnConfigDrift(config, selection, log);

  // 4. One write command at a time (dry runs too: they must see a journal nobody else is changing).
  let release;
  try {
    release = acquireRunLock(paths, 'issue');
  } catch (err) {
    log.error(err instanceof LockError ? err.message : `拿不到运行锁：${err.message}`);
    return done(1);
  }
  // The list was read before the lock: select may have rewritten it in between. Never issue from a stale copy.
  if (selectionReplaced(paths, selection)) {
    release();
    log.error(SELECTION_REPLACED);
    return done(1);
  }

  const isoNow = () => now().toISOString();
  const run = new IssueRun({
    config, paths, selection, tz, today, startMs, log, now, sleep, signal, renderPreview, dryRun, summary,
    dateBlocked, repairOnly,
    retryFailed: !!retryFailed,
    maxCreates: limitGiven ? limit : Infinity,
    reconcileWaitMs, reconcilePollMs, unknownSettleMs,
  });
  let exitCode = 1;
  let started = false;
  try {
    await connect(config, { log, sleep });
    run.journal = foldJournal(readJournal(paths.journal));
    run.batch = run.journal.lastBatch + 1;
    run.runId = uniqueRunId(newRunId(now()), run.journal.runs);
    summary.batch = run.batch;
    appendJournal(paths.journal, {
      op: 'run.start', run: run.runId, command: 'issue', dryRun, batch: run.batch,
      limit: limitGiven ? limit : null, options: { retryFailed: !!retryFailed, repairOnly },
    }, { now: isoNow });
    started = true;
    exitCode = await run.execute();
  } catch (err) {
    if (err instanceof Stop) {
      exitCode = err.exitCode;
    } else {
      exitCode = 1;
      log.error(`已停止：${err.message}`);
    }
  } finally {
    if (started) {
      try {
        appendJournal(paths.journal, { op: 'run.end', run: run.runId, summary, exitCode }, { now: isoNow });
      } catch (err) {
        log.error(`写本地日志（run.end）失败：${err.message}`);
      }
    }
    release();
    if (started) {
      try {
        run.printSummary();
      } catch (err) {
        log.warn(`打印本次结果失败：${err.message}`);
      }
    }
    // Every command ends by regenerating the Excel; its failure never changes the exit code.
    try {
      const report = await writeReport({ config, paths, log, now });
      if (report?.file) log.info(`Excel 已更新：${report.file}`);
      // fileEn is null when the English edition failed: writeReport already logged that warning.
      if (report?.fileEn) log.info(`英文版 Excel 已更新：${report.fileEn}`);
    } catch (err) {
      log.warn(`Excel 没有更新：${err.message}（可以稍后运行 export 重新生成）`);
    }
  }
  return done(exitCode);
}

// ---------------------------------------------------------------------------
// One run
// ---------------------------------------------------------------------------

class IssueRun {
  constructor(o) {
    Object.assign(this, o);
    this.recipients = [...o.selection.recipients].sort((a, b) => a.seq - b.seq);
    this.byId = new Map(this.recipients.map((r) => [r.customerId, r]));
    // "Campaign card search window": the selection is made before any card exists.
    this.sinceIso = isoSeconds(Date.parse(o.selection.createdAt) - SELECTION_MARGIN_MS);
    const exportedMs = Date.parse(o.selection.snapshot?.exportedAt);
    this.snapshotMs = Number.isFinite(exportedMs) ? exportedMs : Date.parse(o.selection.createdAt);
    this.sentTagLower = o.config.sentTag.toLowerCase();
    // Test campaigns skip the audience rules at select time; the pre-flight does too.
    this.audienceRules = o.selection.mode !== 'test';
    // No new card in this run: REMIND_1_DATE has come (dateBlocked), or --repair-only.
    this.noNewCards = !!(o.dateBlocked || o.repairOnly);
    this.local = new Map(); // customer gid → status set (or, in a dry run, simulated) by this run
    this.wouldCreate = []; // dry run: who would get a card, in order
    this.skippedRows = []; // { r, reason, detail }
    this.refreshSize = CHUNK_SIZE;
    this.journal = null;
    this.batch = null;
    this.runId = null;
  }

  async execute() {
    this.banner();
    const blocked = await this.reconcileLeftovers(); // step 6
    await this.fixTags(); // step 7
    if (blocked.length) return this.stopForUnknown(blocked);
    const pre = await this.preflight(); // step 8
    const wanted = this.retryFailed ? STATUS.FAILED : STATUS.PENDING;
    const candidates = this.recipients.filter((r) => this.status(r.customerId) === wanted);
    const queue = await this.applyGuards(candidates, pre); // records cards Shopify already has; creates nothing
    if (this.noNewCards) return this.refuseNewCards(queue);
    const exitCode = await this.walk(queue, pre); // steps 9-10
    if (this.dryRun) await this.reportDryRun(); // step 11
    return exitCode;
  }

  /**
   * No new card in this run; the repairs above are done.
   *  - --repair-only: exit 0, or 1 when a missing tag could not be added.
   *  - On or after REMIND_1_DATE (the creation email would show the reminder copy): a real run
   *    exits 2 while someone is still waiting for a card, 1 when only tag repairs failed;
   *    its dry run says the same as a warning and exits 0.
   * People still waiting = this run's queue; with --retry-failed and no failed row, the
   * people still pending (a normal run would have created theirs).
   */
  refuseNewCards(queue) {
    const { log, config, summary: s, dryRun } = this;
    if (this.repairOnly) {
      if (s.tagFailed) log.warn(tagRetryText(s.tagFailed));
      log.info(dryRun ? REPAIR_ONLY_DRY_TEXT : REPAIR_ONLY_TEXT);
      return s.tagFailed ? 1 : 0;
    }
    const p = dryRun ? '预演：' : '';
    const waiting = queue.length || (this.retryFailed ? this.remainingPending() : 0);
    if (!queue.length) {
      const nobody = this.retryFailed ? '没有建卡失败、待重试的人' : '没有待建卡的人';
      // "Repairs done" only when nothing is left: no one waiting, no tag still missing.
      const repaired = waiting || s.tagFailed ? '' : `；${dryRun ? '真实运行只会补记和补打 tag（如果有）' : '已完成补记和补打 tag（如果有）'}`;
      log.info(`${p}${nobody}${repaired}`);
    }
    if (s.tagFailed) log.warn(tagRetryText(s.tagFailed));
    if (!waiting) return s.tagFailed ? 1 : 0;
    s.newCardsRefused = waiting;
    if (dryRun) {
      log.warn(`预演：真实运行今天只会补记和补打 tag，不会建新卡；还有 ${waiting} 人待建卡。`);
      return 0;
    }
    log.error(`${tooLateText(config.remind1Date)}本次只做了补记和补打 tag，没有建新卡；还有 ${waiting} 人待建卡。`);
    log.info(`现在是店铺时间 ${this.localTime(this.now().getTime())}；测试活动不受这个限制`);
    return 2;
  }

  /**
   * The store's date reached REMIND_1_DATE during a real run (it crossed midnight): no more
   * cards. Everything done so far stays as it is; `left` people of this run's queue get none.
   */
  stopByDate(date, ms, left) {
    this.summary.stoppedByDate = date;
    this.log.error(`${tooLateText(this.config.remind1Date)}已处理的人都记在本地日志里；剩下的 ${left} 人今天不会再建卡。`);
    this.log.info(`现在是店铺时间 ${this.localTime(ms)}`);
    return 2;
  }

  // -- state ----------------------------------------------------------------

  status(cid) {
    return this.local.get(cid) ?? this.journal.customers.get(cid)?.status ?? STATUS.PENDING;
  }

  setStatus(cid, status) {
    this.local.set(cid, status);
  }

  /** Append (and fsync) a journal entry. Real runs only; `t` pins the timestamp. */
  record(entry, t) {
    if (this.dryRun) return;
    appendJournal(this.paths.journal, { ...entry, run: this.runId }, { now: () => t ?? this.now().toISOString() });
  }

  /** Called before starting each new customer: Ctrl+C finishes the current one, then stops here. */
  checkAbort() {
    if (this.signal?.aborted) {
      this.log.warn('收到中断信号（Ctrl+C）：当前这个人已经处理完，已停止。没处理的人仍是待发放，下次运行接着发');
      throw new Stop(130);
    }
  }

  localTime(ms) {
    return Number.isFinite(ms) ? localDateTime(ms, this.tz) : '?';
  }

  // -- step 6: rows whose create outcome is still open -------------------------

  /** Looks up every in_progress/unknown row. Returns the rows that must still wait. */
  async reconcileLeftovers() {
    const { log } = this;
    const leftovers = [...this.journal.customers.entries()]
      .filter(([, s]) => s.status === STATUS.IN_PROGRESS || s.status === STATUS.UNKNOWN)
      .map(([cid, s]) => ({ cid, s, r: this.byId.get(cid) ?? { seq: '?', name: numericId(cid), customerId: cid } }))
      .sort((a, b) => (Number(a.r.seq) || Infinity) - (Number(b.r.seq) || Infinity));
    if (!leftovers.length) return [];

    const p = this.dryRun ? '预演：' : '';
    log.info(`先核对上次建卡结果不明的 ${leftovers.length} 人：按客户和时间在 Shopify 上找这张卡`);
    const blocked = [];
    const nowMs = this.now().getTime();
    for (const { cid, s, r } of leftovers) {
      this.checkAbort();
      const startedMs = Date.parse(s.startedAt);
      const fromMs = (Number.isFinite(startedMs) ? startedMs : Date.parse(this.sinceIso)) - CLOCK_SKEW_MS;
      const cards = await this.cardsOf(cid, isoSeconds(fromMs));
      if (cards.length) {
        const card = newest(cards);
        this.record({ op: 'reconcile.found', cid, giftCardId: card.id, last4: card.last4, amountCents: card.amountCents, createdAt: card.createdAt, source: 'issue-reconcile' });
        this.setStatus(cid, s.taggedAt ? STATUS.DONE : STATUS.CREATED);
        this.summary.reconciled += 1;
        log.info(`${p}${label(r)}：在 Shopify 上找到了上次那张卡（…${card.last4}，${formatUsd(card.amountCents)}），${this.dryRun ? '真实运行会补记' : '已补记'}，不会再建第二张`);
        this.warnCardDetails(r, card, cards.length);
      } else if (Number.isFinite(startedMs) && nowMs - startedMs >= this.unknownSettleMs) {
        // Back to pending; from REMIND_1_DATE on the real campaign will not issue it again, so do not promise that.
        const settled = `超过 ${durationText(this.unknownSettleMs)}仍查不到这张卡`;
        const noReissue = this.dateBlocked ? noNewCardsText(this.config.remind1Date) : this.repairOnly ? REPAIR_ONLY_NO_REISSUE : null;
        this.record({ op: 'reconcile.none', cid, note: noReissue ? `${settled}，确认未建成；${noReissue}` : `${settled}，确认未建成，可以重试` });
        this.setStatus(cid, STATUS.PENDING);
        this.summary.reconciledNone += 1;
        log.info(`${p}${label(r)}：上次建卡开始于 ${this.localTime(startedMs)}，${settled}，确认没有建成`
          + (noReissue ? `；${noReissue}` : `，${this.dryRun ? '真实运行会' : ''}重新排队发放`));
      } else {
        // Too early to tell: Shopify's search may simply not show the card yet.
        if (s.status === STATUS.IN_PROGRESS) {
          this.record({ op: 'create.unknown', cid, error: '上次运行在建卡途中中断，Shopify 上暂时查不到这张卡' });
        }
        this.setStatus(cid, STATUS.UNKNOWN);
        blocked.push({ r, startedMs });
      }
    }
    return blocked;
  }

  stopForUnknown(blocked) {
    const { log } = this;
    this.summary.stillUnknown = blocked.length;
    for (const { r, startedMs } of blocked) {
      log.error(`${label(r)}：上次建卡开始于 ${this.localTime(startedMs)}（店铺时间），结果不明，Shopify 上暂时还查不到这张卡（Shopify 的搜索可能有几分钟延迟）`);
    }
    const latest = Math.max(...blocked.map((b) => b.startedMs).filter(Number.isFinite));
    const retryAt = Number.isFinite(latest) ? Math.ceil((latest + this.unknownSettleMs) / 60_000) * 60_000 : NaN;
    // A row still not found then goes back to pending, but from REMIND_1_DATE on (today, or by the
    // time the next run may settle it) the real campaign creates no new card: no promise of a re-issue.
    const noReissue = this.dateBlocked || (Number.isFinite(retryAt) && this.blocksNewCardsAt(retryAt));
    const then = noReissue
      ? `那时仍查不到的，程序会确认没有建成；${noNewCardsText(this.config.remind1Date)}。`
      : this.repairOnly
        ? `那时仍查不到的，程序会确认没有建成；${REPAIR_ONLY_NO_REISSUE}。`
        : '那时仍查不到的，程序会确认没有建成并重新发放；';
    log.error(`${this.dryRun ? '预演：真实运行会在这里停止' : '已停止'}：还有 ${blocked.length} 人的建卡结果需要确认，为了不重复建卡，这次不建新卡。`
      + `请在 ${this.localTime(retryAt)}（店铺时间）之后再运行：${then}也可以先在 Shopify 后台的礼品卡列表里按客户核对。`);
    return 1;
  }

  /** True when, at instant `ms`, the real campaign may no longer create cards (REMIND_1_DATE in store time). */
  blocksNewCardsAt(ms) {
    return newCardsBlockedOn(this.config, this.selection, localDate(ms, this.tz));
  }

  // -- step 7: cards that exist but whose tag is missing ------------------------

  async fixTags() {
    const rows = this.recipients.filter((r) => this.status(r.customerId) === STATUS.CREATED);
    if (!rows.length) return;
    if (this.dryRun) {
      for (const r of rows) this.setStatus(r.customerId, STATUS.DONE);
      this.summary.tagFixed += rows.length;
      this.log.info(`预演：${rows.length} 人已建卡但还没打 tag，真实运行会先给他们补打 tag ${this.config.sentTag}`);
      return;
    }
    this.log.info(`补打 tag：${rows.length} 人已建卡但还没打 tag ${this.config.sentTag}`);
    for (const r of rows) {
      this.checkAbort();
      if (await this.tag(r)) {
        this.summary.tagFixed += 1;
        this.log.info(`${label(r)}：已补打 tag`);
      }
    }
  }

  /** tagsAdd + journal. A failure is recorded and the run goes on: the next run tags again. */
  async tag(r) {
    const cid = r.customerId;
    try {
      await addTag(cid, this.config.sentTag);
    } catch (err) {
      this.record({ op: 'tag.fail', cid, error: err.message });
      this.summary.tagFailed += 1;
      this.log.warn(`${label(r)}：打 tag 失败（${err.message}）；卡已经建好，下次运行会补打 tag`);
      return false;
    }
    this.record({ op: 'tag.ok', cid });
    this.setStatus(cid, STATUS.DONE);
    return true;
  }

  // -- step 8: read-only pre-flight -------------------------------------------

  async preflight() {
    const { config, log } = this;
    const taggedIds = await fetchTaggedCustomerIds(config.sentTag); // no log: its line is English
    const cards = await findCampaignCards({ campaignId: config.campaignId, sinceIso: this.sinceIso });
    const cardsByCustomer = new Map();
    for (const card of cards) {
      if (!card.customerId) continue;
      if (!cardsByCustomer.has(card.customerId)) cardsByCustomer.set(card.customerId, []);
      cardsByCustomer.get(card.customerId).push(card);
    }
    const orders = await this.ordersSince(isoSeconds(this.snapshotMs));
    const activity = buildActivity(orders);
    const orderedAt = new Map(); // customer gid → newest valid order time since the snapshot
    let valid = 0;
    for (const o of orders) {
      if (o.cancelledAt || o.test || !o.customer?.id) continue;
      valid += 1;
      const prev = orderedAt.get(o.customer.id);
      if (!prev || Date.parse(o.createdAt) > Date.parse(prev)) orderedAt.set(o.customer.id, o.createdAt);
    }
    log.info(`发放前复查：Shopify 上有 ${taggedIds.size} 个客户已带 tag ${config.sentTag}，本活动已有 ${cards.length} 张卡；`
      + `名单导出（${this.localTime(this.snapshotMs)}）之后有 ${valid} 笔新订单`);
    return { taggedIds, cardsByCustomer, activity, orderedAt };
  }

  /** Shopify-side guards for this run's candidates: an existing campaign card, or SENT_TAG. */
  async applyGuards(candidates, pre) {
    const { log, config } = this;
    const p = this.dryRun ? '预演：' : '';
    const queue = [];
    for (const r of candidates) {
      const cid = r.customerId;
      const cards = pre.cardsByCustomer.get(cid);
      if (cards?.length) {
        // The journal does not know this card (deleted journal, or a card made outside this run's records).
        this.checkAbort();
        const card = newest(cards);
        this.record({ op: 'reconcile.found', cid, giftCardId: card.id, last4: card.last4, amountCents: card.amountCents, createdAt: card.createdAt, source: 'preflight' });
        this.setStatus(cid, STATUS.CREATED);
        this.summary.reconciled += 1;
        log.warn(`${p}${label(r)}：Shopify 上已有本活动的卡（…${card.last4}，${formatUsd(card.amountCents)}），本地日志里没有记录；${this.dryRun ? '真实运行会补记' : '已补记'}，不会再建卡`);
        this.warnCardDetails(r, card, cards.length);
        if (pre.taggedIds.has(cid)) {
          this.record({ op: 'tag.ok', cid, note: 'already tagged in Shopify' });
          this.setStatus(cid, STATUS.DONE);
          this.summary.tagFixed += 1;
        } else if (this.dryRun) {
          this.setStatus(cid, STATUS.DONE);
          this.summary.tagFixed += 1;
        } else if (await this.tag(r)) {
          this.summary.tagFixed += 1;
        }
        continue;
      }
      if (pre.taggedIds.has(cid)) {
        this.checkAbort();
        this.skip(r, 'already-tagged', config.sentTag);
        continue;
      }
      queue.push(r);
    }
    return queue;
  }

  // -- steps 9-10: walk the candidates -------------------------------------------

  async walk(queue, pre) {
    const { log } = this;
    let consecutiveFailures = 0;
    for (let i = 0; i < queue.length && this.summary.attempted < this.maxCreates; i += CHUNK_SIZE) {
      this.checkAbort();
      const chunk = queue.slice(i, i + CHUNK_SIZE);
      const nodes = await this.refreshCustomers(chunk.map((r) => r.customerId));
      for (let j = 0; j < chunk.length; j += 1) {
        const r = chunk[j];
        if (this.summary.attempted >= this.maxCreates) break;
        this.checkAbort();
        const node = nodes.get(r.customerId);
        const why = this.skipReason(r, node ? parseCustomer(node) : null, pre);
        if (why) {
          this.skip(r, why.reason, why.detail); // skips do not count toward --limit
          continue;
        }
        if (this.dryRun) {
          this.simulateCreate(r);
          continue;
        }
        // The store's date again, right before this card: a long run may have crossed midnight
        // into REMIND_1_DATE. The same instant is the card's create.start time.
        const startedMs = this.now().getTime();
        const date = localDate(startedMs, this.tz);
        if (newCardsBlockedOn(this.config, this.selection, date)) return this.stopByDate(date, startedMs, queue.length - (i + j));
        const outcome = await this.issueOne(r, new Date(startedMs).toISOString());
        if (outcome === 'created') {
          consecutiveFailures = 0;
        } else if (outcome === 'failed') {
          consecutiveFailures += 1;
          if (consecutiveFailures >= MAX_CONSECUTIVE_FAILURES) {
            log.error(`已停止：连续 ${MAX_CONSECUTIVE_FAILURES} 次建卡被 Shopify 拒绝。请在 Excel 的"备注/错误"列查看原因，解决后用 --retry-failed --limit N 重试`);
            return 1;
          }
        } else {
          return 1; // rejected / unknown: explained by issueOne
        }
        if (this.summary.attempted % PROGRESS_EVERY === 0) this.logProgress();
      }
    }
    if (Number.isFinite(this.maxCreates) && this.summary.attempted < this.maxCreates) {
      log.info(this.retryFailed ? '没有更多建卡失败、待重试的人了' : '名单里已经没有更多待发放的人了');
    }
    return 0;
  }

  /** The first reason this customer must not get a card now, or null. `c` is the fresh read (null = deleted). */
  skipReason(r, c, pre) {
    if (!c) return { reason: 'customer-deleted' };
    if (!c.email || !c.emailValid) return { reason: 'no-email', detail: c.email ? '邮箱格式无效' : undefined };
    if (this.audienceRules) {
      if ((this.config.excludeEmailDomains ?? []).includes(c.emailDomain)) return { reason: 'relay-email', detail: c.emailDomain };
      if (this.selection.params.requireEmailSubscribed && c.marketingState !== 'SUBSCRIBED') {
        return { reason: 'not-subscribed', detail: c.marketingState ?? 'NONE' };
      }
    }
    if (c.tags.some((t) => String(t).toLowerCase() === this.sentTagLower)) return { reason: 'already-tagged', detail: this.config.sentTag };
    if (this.audienceRules) {
      if (pre.activity.activeCustomerIds.has(c.id)) return { reason: 'ordered-since-snapshot', detail: pre.orderedAt.get(c.id) };
      if (isValidOrder(c.lastOrder) && c.lastOrder.createdAtMs > this.snapshotMs) {
        return { reason: 'ordered-since-snapshot', detail: c.lastOrder.createdAt };
      }
      // The current default address and the one the selection was made with.
      for (const key of new Set([c.addressKey, r.addressKey].filter(Boolean))) {
        const a = pre.activity.activeAddresses.get(key);
        if (a && a.customerId !== c.id) return { reason: 'address-ordered-since-snapshot', detail: a.customerId };
      }
    }
    return null;
  }

  skip(r, reason, detail) {
    const cid = r.customerId;
    this.record({ op: 'skip', cid, reason, ...(detail ? { detail } : {}), batch: this.batch });
    this.setStatus(cid, STATUS.SKIPPED);
    this.summary.skipped[reason] = (this.summary.skipped[reason] ?? 0) + 1;
    this.skippedRows.push({ r, reason, detail });
    if (!this.dryRun) this.log.info(`${label(r)}：跳过，${this.reasonText(reason, detail)}`);
  }

  reasonText(reason, detail) {
    const text = SKIP_REASONS[reason] ?? reason;
    if (!detail) return text;
    let d = String(detail);
    if (d.startsWith('gid://shopify/Customer/')) d = `客户 ${numericId(d)}`;
    else if (/^\d{4}-\d{2}-\d{2}T/.test(d)) d = `${this.localTime(Date.parse(d))} 店铺时间`;
    return `${text}（${d}）`;
  }

  simulateCreate(r) {
    this.wouldCreate.push(r);
    this.summary.attempted += 1;
    this.summary.amountCents += r.amountCents;
    this.noteSeq(r);
    this.setStatus(r.customerId, STATUS.DONE);
  }

  /**
   * create.start → giftCardCreate → outcome → tag. Returns 'created' | 'failed' | 'rejected' | 'unknown'.
   * `startedIso` is the instant walk() checked the date at.
   */
  async issueOne(r, startedIso = this.now().toISOString()) {
    const { log } = this;
    const cid = r.customerId;
    // Write-ahead: if the process dies after Shopify creates the card, this line makes the next run look for it.
    this.record({ op: 'create.start', cid, amountCents: r.amountCents, batch: this.batch }, startedIso);
    this.setStatus(cid, STATUS.IN_PROGRESS);
    this.summary.attempted += 1;
    this.noteSeq(r);

    let created = null;
    let error = null;
    try {
      created = await createGiftCard(cid, this.createOptions(r));
    } catch (err) {
      error = err;
    }

    let card;
    let how = '';
    if (created) {
      card = { id: created.id, last4: created.last4, amountCents: r.amountCents };
      this.record({ op: 'create.ok', cid, giftCardId: card.id, last4: card.last4, amountCents: card.amountCents, batch: this.batch });
    } else if (error?.code === 'USER_ERROR') {
      // Permanent rejection: only --retry-failed tries this row again.
      this.record({ op: 'create.fail', cid, error: error.message });
      this.setStatus(cid, STATUS.FAILED);
      this.summary.failed += 1;
      log.warn(`${label(r)}：建卡被 Shopify 拒绝（${error.message}），记为失败，继续下一个人`);
      return 'failed';
    } else if (error?.outcomeKnown === true) {
      // Definitely not applied (e.g. throttled after every retry): back to pending, stop for now.
      this.record({ op: 'create.rejected', cid, error: error.message });
      this.setStatus(cid, STATUS.PENDING);
      this.summary.rejected += 1;
      log.error(`已停止：Shopify 没有接受给 ${label(r)} 建卡的请求（${error.message}），这张卡没有建成。这个人仍是待发放，下次运行会重试`);
      return 'rejected';
    } else {
      // The answer was lost: the card may exist. Look for it instead of retrying.
      log.warn(`${label(r)}：没有收到 Shopify 的回应（${error?.message ?? '未知错误'}），正在按客户和时间查这张卡是否已建成，最多等 ${durationText(this.reconcileWaitMs)}`);
      const found = await this.findCardWithin(cid, startedIso);
      if (!found.card) {
        const lookup = found.lastError ? `；查卡也失败了：${found.lastError.message}` : '';
        this.record({ op: 'create.unknown', cid, error: `${error?.message ?? '未知错误'}${lookup}` });
        this.setStatus(cid, STATUS.UNKNOWN);
        this.summary.unknown += 1;
        const retryAt = Math.ceil((Date.parse(startedIso) + this.unknownSettleMs) / 60_000) * 60_000;
        // Settled after midnight into REMIND_1_DATE, a card that was not made is not made later either.
        const then = this.blocksNewCardsAt(retryAt)
          ? `程序会先核对这张卡；查不到的确认没有建成，${noNewCardsText(this.config.remind1Date)}`
          : '程序会先核对这张卡，查不到才会重新发放';
        log.error(`已停止：${label(r)} 的建卡结果不明，等了 ${durationText(this.reconcileWaitMs)}在 Shopify 上仍查不到这张卡${lookup}。为了不重复建卡，本次停止。`
          + `请在 ${this.localTime(retryAt)}（店铺时间）之后再运行：${then}`);
        return 'unknown';
      }
      card = found.card;
      how = '（Shopify 没有回应，但已确认卡建成了）';
      this.record({ op: 'reconcile.found', cid, giftCardId: card.id, last4: card.last4, amountCents: card.amountCents, createdAt: card.createdAt, source: 'issue-inline' });
      this.warnCardDetails(r, card, found.count);
    }

    this.setStatus(cid, STATUS.CREATED);
    this.summary.created += 1;
    this.summary.amountCents += card.amountCents;
    const tagged = await this.tag(r);
    if (tagged) this.summary.tagged += 1;
    log.info(`${label(r)}：已建卡 ${formatUsd(card.amountCents)}（…${card.last4}）${how}，Shopify 已自动发出首封邮件${tagged ? '；已打 tag' : ''}`);
    return 'created';
  }

  createOptions(r) {
    const p = this.selection.params;
    return {
      amountCents: r.amountCents,
      currencyCode: p.currency,
      note: campaignNote(p.giftCardNote, this.config.campaignId),
      expiresOn: p.giftCardExpiresOn || undefined,
      templateSuffix: p.giftCardTemplateSuffix || undefined,
    };
  }

  /**
   * Poll for the customer's campaign card created since the attempt began.
   * Waiting is counted in slept time (not the clock) so an injected clock cannot stall it.
   * A failed lookup counts as "not found yet".
   */
  async findCardWithin(cid, startedIso) {
    const sinceIso = isoSeconds(Date.parse(startedIso) - CLOCK_SKEW_MS);
    let waited = 0;
    let lastError = null;
    for (;;) {
      try {
        const cards = await this.cardsOf(cid, sinceIso);
        if (cards.length) return { card: newest(cards), count: cards.length, lastError: null };
        lastError = null;
      } catch (err) {
        lastError = err;
      }
      if (waited >= this.reconcileWaitMs) return { card: null, count: 0, lastError };
      const step = Math.max(1, Math.min(this.reconcilePollMs, this.reconcileWaitMs - waited));
      await this.sleep(step);
      waited += step;
    }
  }

  warnCardDetails(r, card, count) {
    if (count > 1) this.log.warn(`${label(r)}：Shopify 上有 ${count} 张本活动的卡，日志记下最新的一张；请运行 verify 核对`);
    if (Number.isInteger(r.amountCents) && card.amountCents !== r.amountCents) {
      this.log.warn(`${label(r)}：这张卡的金额 ${formatUsd(card.amountCents)} 和名单里的 ${formatUsd(r.amountCents)} 不一样，请核对`);
    }
  }

  // -- Shopify reads ----------------------------------------------------------

  /**
   * This campaign's cards of one customer created since `sinceIso`. The customer is
   * also checked here, not only by the search filter: inside a lookup window the
   * previous customer's card (made a second earlier) must never pass for this one's.
   */
  async cardsOf(cid, sinceIso) {
    const cards = await findCampaignCards({ campaignId: this.config.campaignId, sinceIso, customerId: cid });
    return cards.filter((card) => card.customerId === cid);
  }

  /**
   * nodes(ids:) for up to CHUNK_SIZE ids → Map<gid, node>; deleted customers are absent.
   * If Shopify rejects the request as too expensive, the batch is halved and retried.
   */
  async refreshCustomers(ids) {
    const byId = new Map();
    let i = 0;
    while (i < ids.length) {
      const part = ids.slice(i, i + this.refreshSize);
      let data;
      try {
        data = await gql(REFRESH_CUSTOMERS, { ids: part });
      } catch (err) {
        if (err?.code === 'MAX_COST_EXCEEDED' && part.length > 1) {
          this.refreshSize = Math.max(1, Math.floor(part.length / 2));
          continue;
        }
        throw err;
      }
      for (const node of data.nodes ?? []) if (node?.id) byId.set(node.id, node);
      i += part.length;
    }
    return byId;
  }

  /** Every order created since `sinceIso` (ORDERS_SINCE nodes, all pages). */
  async ordersSince(sinceIso) {
    const nodes = [];
    let after = null;
    do {
      const data = await gql(ORDERS_SINCE, { query: `created_at:>='${sinceIso}'`, after });
      const connection = data.orders;
      if (!connection) throw new Error('orders query returned nothing');
      nodes.push(...(connection.nodes ?? []));
      const next = connection.pageInfo?.hasNextPage ? connection.pageInfo.endCursor : null;
      if (next && next === after) throw new Error('orders query returned the same cursor twice');
      after = next;
    } while (after);
    return nodes;
  }

  // -- output -----------------------------------------------------------------

  banner() {
    const { config, selection, log } = this;
    const parts = [`批次 ${this.batch}`, `活动 ${config.campaignId}${selection.mode === 'test' ? '（测试活动）' : ''}`];
    if (this.repairOnly) parts.push('只补记和补打 tag，不建新卡（--repair-only）');
    else if (this.dateBlocked) parts.push(`已到 REMIND_1_DATE（${config.remind1Date}）：只补记和补打 tag，不建新卡`);
    else parts.push(Number.isFinite(this.maxCreates) ? `本次最多建卡 ${this.maxCreates} 张` : '不限张数（预演全部剩下的人）');
    if (this.retryFailed) parts.push('只重试之前建卡失败的人');
    log.info(parts.join('｜'));
    if (this.dryRun) {
      log.info('预演（DRY_RUN）：只从 Shopify 读数据，不建卡、不发邮件、不打 tag；本地日志只记一条预演记录。真实运行要在命令前加 DRY_RUN=false');
      if (selection.mode !== 'test' && !this.repairOnly) {
        if (config.launchDate && this.today < config.launchDate) {
          log.warn(`注意：正式活动要到 ${config.launchDate}（店铺时间）才能真实建卡发信；预演不受这个限制`);
        }
        if (this.dateBlocked) log.warn(`注意：${tooLateText(config.remind1Date)}`);
      }
    } else if (this.noNewCards) {
      log.info(`真实运行：只补记 Shopify 上已有的卡、补打 tag ${config.sentTag}；不建新卡，不发邮件`);
    } else {
      log.info(`真实运行：会建礼品卡（Shopify 建卡时自动给客户发首封邮件），然后给客户打 tag ${config.sentTag}`);
      this.warnUtcEvening();
    }
  }

  /**
   * Real runs of the real campaign that create cards: Shopify's docs do not say whether the
   * template's 'now' is the store's date or UTC's. In the Los Angeles evening UTC is already the
   * next day; when that day's copy differs, say so once (only daytime sends are right under both).
   */
  warnUtcEvening() {
    if (this.selection.mode === 'test') return;
    const utcDate = new Date(this.startMs).toISOString().slice(0, 10);
    if (utcDate <= this.today) return;
    const utcStage = firstEmailStageOn(this.config, utcDate);
    if (utcStage === firstEmailStageOn(this.config, this.today)) return;
    this.log.warn(`现在店铺时间 ${this.localTime(this.startMs)}，UTC 已是 ${utcDate}：如果 Shopify 按 UTC 日期选邮件文案，`
      + `今天建卡发出的首封邮件会显示${STAGE_LABELS[utcStage]}文案。建议在洛杉矶时间 17:00 之前运行。`);
  }

  logProgress() {
    const s = this.summary;
    const skipped = Object.values(s.skipped).reduce((a, b) => a + b, 0);
    const of = Number.isFinite(this.maxCreates) ? `/${this.maxCreates}` : '';
    this.log.info(`进度：已处理 ${s.attempted}${of} 人，建卡 ${s.created} 张（${formatUsd(s.amountCents)}），跳过 ${skipped}，失败 ${s.failed}`);
  }

  noteSeq(r) {
    const s = this.summary;
    s.seqFrom = s.seqFrom === null ? r.seq : Math.min(s.seqFrom, r.seq);
    s.seqTo = s.seqTo === null ? r.seq : Math.max(s.seqTo, r.seq);
  }

  async reportDryRun() {
    const { log, summary: s } = this;
    const rows = this.wouldCreate;
    if (rows.length) {
      log.info(`预演：这一批会给 ${rows.length} 人建卡，序号 ${s.seqFrom}–${s.seqTo}；Shopify 会给每人自动发首封邮件`);
      const byAmount = new Map();
      for (const r of rows) byAmount.set(r.amountCents, (byAmount.get(r.amountCents) ?? 0) + 1);
      const kind = this.selection.mode === 'test' ? '测试金额' : '档位';
      for (const [amount, n] of [...byAmount].sort((a, b) => a[0] - b[0])) {
        log.info(`  ${kind} ${formatUsd(amount)}：${n} 人，合计 ${formatUsd(amount * n)}`);
      }
      log.info(`  金额合计：${formatUsd(s.amountCents)}`);
    } else {
      log.info('预演：这一批没有人会建卡');
    }
    if (this.skippedRows.length) {
      const counts = Object.entries(s.skipped).map(([k, n]) => `${SKIP_REASONS[k] ?? k} ${n}`).join('，');
      log.info(`预演：会被跳过 ${this.skippedRows.length} 人（${counts}）：`);
      for (const { r, reason, detail } of this.skippedRows.slice(0, DRY_RUN_SKIP_LIST_MAX)) {
        log.info(`  ${label(r)}：${this.reasonText(reason, detail)}`);
      }
      if (this.skippedRows.length > DRY_RUN_SKIP_LIST_MAX) {
        log.info(`  …另外还有 ${this.skippedRows.length - DRY_RUN_SKIP_LIST_MAX} 人，原因见上面的统计`);
      }
    }
    if (!rows.length) return;
    log.info(`前 ${Math.min(10, rows.length)} 人：`);
    for (const r of rows.slice(0, 10)) log.info(`  #${r.seq} ${r.name} <${r.email}> ${formatUsd(r.amountCents)}`);
    // The email this batch would really get: Shopify sends it the day the card is created, which
    // is today, and for the real campaign never before LAUNCH_DATE. The cards carry the
    // selection's expiry date (createOptions), so the preview shows that one, not .env's.
    const sendDate = this.selection.mode === 'test' ? this.today : laterDate(this.today, this.config.launchDate || this.today);
    try {
      const preview = await this.renderPreview({
        config: this.config,
        paths: this.paths,
        variant: 'first',
        recipient: rows[0],
        log,
        sendDate,
        expiresOn: this.selection.params.giftCardExpiresOn,
      });
      if (preview?.file) {
        log.info(`首封邮件预览（用这一批第 1 个人 #${rows[0].seq} 的名字和金额，按 ${sendDate} 发送时的文案）：${preview.file}`
          + `${preview.subject ? `（主题：${preview.subject}）` : ''}`);
      }
    } catch (err) {
      log.warn(`邮件预览没有生成：${err.message}`);
    }
  }

  remainingPending() {
    return this.recipients.filter((r) => this.status(r.customerId) === STATUS.PENDING).length;
  }

  printSummary() {
    const { log, summary: s, dryRun } = this;
    const skippedTotal = Object.values(s.skipped).reduce((a, b) => a + b, 0);
    const reasons = Object.entries(s.skipped).map(([k, n]) => `${SKIP_REASONS[k] ?? k} ${n}`).join('，');
    const will = dryRun ? '将' : '';
    log.info(dryRun ? `—— 预演结果（批次 ${s.batch}，Shopify 上什么都没改）——` : `—— 本次结果（批次 ${s.batch}）——`);
    if (s.seqFrom !== null) log.info(`本批序号：${s.seqFrom}–${s.seqTo}`);
    log.info(`${will}建卡：${dryRun ? s.attempted : s.created} 张，金额合计 ${formatUsd(s.amountCents)}`);
    log.info(`${will}打 tag：${dryRun ? s.attempted : s.tagged} 人`);
    if (s.tagFixed) log.info(`${will}补打 tag（之前已建卡）：${s.tagFixed} 人`);
    if (s.tagFailed) log.info(`打 tag 失败：${s.tagFailed} 人（卡已建好，下次运行会补打）`);
    if (s.reconciled) log.info(`${will}补记 Shopify 上已有的卡：${s.reconciled} 张`);
    if (s.reconciledNone) {
      log.info(this.dateBlocked
        ? `${will}确认上次没有建成：${s.reconciledNone} 人；${noNewCardsText(this.config.remind1Date)}`
        : this.repairOnly
          ? `${will}确认上次没有建成：${s.reconciledNone} 人；${REPAIR_ONLY_NO_REISSUE}`
          : `${will}确认上次没有建成、重新排队：${s.reconciledNone} 人`);
    }
    log.info(`${will}跳过：${skippedTotal} 人${reasons ? `（${reasons}）` : ''}`);
    log.info(`失败：${s.failed} 人${s.failed ? '（原因解决后用 --retry-failed --limit N 重试）' : ''}`);
    if (s.rejected) log.info(`Shopify 没有接受、没有建卡：${s.rejected} 人（仍是待发放，下次运行会重试）`);
    log.info(`结果不明：${s.unknown + s.stillUnknown} 人`);
    log.info(`名单共 ${this.recipients.length} 人，${dryRun ? '这一批之后' : '现在'}还有 ${this.remainingPending()} 人待发放`);
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function emptySummary(dryRun) {
  return {
    batch: null,
    dryRun,
    attempted: 0, // creates attempted (a dry run: cards that would be created)
    created: 0, // new cards, including ones confirmed after a lost answer
    tagged: 0, // new cards tagged in this run
    tagFixed: 0, // tags added (or found) for cards made earlier
    reconciled: 0, // existing cards recorded from Shopify (leftovers + pre-flight)
    skipped: {}, // reason → count
    failed: 0, // userErrors
    rejected: 0, // definitely not applied; back to pending
    unknown: 0, // outcome still unknown after the inline lookup
    amountCents: 0, // value of the new cards (a dry run: of the cards that would be created)
    tagFailed: 0,
    reconciledNone: 0, // leftovers confirmed not created, queued again
    stillUnknown: 0, // leftovers that still have to wait
    seqFrom: null,
    seqTo: null,
  };
}

/** Why this selection cannot be issued from with this config, or null. */
function selectionProblem(selection, config) {
  if (typeof selection !== 'object' || !Array.isArray(selection.recipients) || !selection.params) {
    return 'selection.json 的格式不对（缺少 recipients 或 params），请重新运行 select';
  }
  if (selection.campaignId !== config.campaignId || selection.params.sentTag !== config.sentTag) {
    return `配置和名单不一致：名单是 CAMPAIGN_ID=${selection.campaignId}、SENT_TAG=${selection.params.sentTag}，`
      + `现在的配置是 CAMPAIGN_ID=${config.campaignId}、SENT_TAG=${config.sentTag}。请把 .env 改回生成名单时的设置（还没开始发放的话，也可以重新运行 select）`;
  }
  if (!selection.params.timezone) return '名单里没有店铺时区（params.timezone），请重新运行 select';
  if (!/^[A-Z]{3}$/.test(String(selection.params.currency ?? ''))) return `名单里的币种不对：${selection.params.currency}，请重新运行 select`;
  if (!Number.isFinite(Date.parse(selection.createdAt))) return '名单里没有生成时间（createdAt），请重新运行 select';
  const seen = new Set();
  for (const r of selection.recipients) {
    if (!r || !/^gid:\/\/shopify\/Customer\/\d+$/.test(String(r.customerId)) || !Number.isInteger(r.seq)) {
      return `名单里有一行的客户 ID 或序号不对（${JSON.stringify(r?.customerId)}），请重新运行 select`;
    }
    if (!Number.isInteger(r.amountCents) || r.amountCents <= 0) return `名单里 #${r.seq} 的金额不对（${r.amountCents}），请重新运行 select`;
    if (seen.has(r.customerId)) return `名单里客户 ${r.customerId} 出现了不止一次，请重新运行 select`;
    seen.add(r.customerId);
  }
  return null;
}

/** Cards are created with the selection's frozen settings; say so when .env has changed since. */
function warnConfigDrift(config, selection, log) {
  const p = selection.params;
  const pairs = [
    ['GIFT_CARD_NOTE', config.giftCardNote, p.giftCardNote],
    ['GIFT_CARD_TEMPLATE_SUFFIX', config.giftCardTemplateSuffix, p.giftCardTemplateSuffix],
    ['GIFT_CARD_EXPIRES_ON', config.giftCardExpiresOn, p.giftCardExpiresOn],
    ['GIFT_CARD_CURRENCY', config.giftCardCurrency, p.currency],
  ];
  for (const [key, current, frozen] of pairs) {
    if ((current ?? '') !== (frozen ?? '')) {
      log.warn(`注意：.env 的 ${key}（${current || '空'}）和生成名单时（${frozen || '空'}）不一样；建卡用名单里的设置。要改只能在开始发放前重新运行 select`);
    }
  }
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

/** The later of two YYYY-MM-DD dates. */
function laterDate(a, b) {
  return a > b ? a : b;
}

/** newRunId() is per second and pid; make sure it is new to this journal. */
function uniqueRunId(base, runs) {
  const used = new Set(runs.map((r) => r.run));
  if (!used.has(base)) return base;
  for (let i = 2; ; i += 1) {
    if (!used.has(`${base}-${i}`)) return `${base}-${i}`;
  }
}

/** ISO-8601 truncated to whole seconds, for Shopify search. Rounding down only widens a ">=" window. */
function isoSeconds(ms) {
  return new Date(Math.floor(ms / 1000) * 1000).toISOString().replace('.000Z', 'Z');
}

function newest(cards) {
  return cards.reduce((a, b) => (Date.parse(b.createdAt) > Date.parse(a.createdAt) ? b : a));
}

function numericId(id) {
  return String(id ?? '').split('/').pop();
}

function label(r) {
  return `#${r.seq} ${r.name || numericId(r.customerId)}`;
}

/** "10 分钟" / "90 秒". */
function durationText(ms) {
  return ms >= 60_000 && ms % 60_000 === 0 ? `${ms / 60_000} 分钟` : `${Math.round(ms / 1000)} 秒`;
}
