// `select` (step 1, read-only on Shopify): export every customer, apply the
// audience rules, dedupe by address, compute the tiered amounts and freeze the
// result in campaigns/<CAMPAIGN_ID>/selection.json. It may be re-run until
// issuing starts; every run ends by regenerating the Excel. "Issuing started"
// is checked twice: in the local journal, and in Shopify itself (any gift card
// carrying this campaign's marker), so a lost or moved journal can not unfreeze
// the list. A re-made list would move the campaign-card search window
// (selection.createdAt − 1 h) past the cards already issued.
//
// Snapshot files in the campaign directory:
//   customers.jsonl     raw customer nodes, one per line (bulk export or pagination)
//   active-orders.json  ORDERS_SINCE nodes since the cutoff, fetched right after the export
//   order-history.json  follow-up lookups { [customerId]: { orders, percent, deleted? } }
//                       (deleted: true = Shopify no longer had the customer)
//   snapshot.json       { exportedAt, source, count, rootObjectCount, cutoffIso, cutoffMs, timezone, inactiveMonths }
// snapshot.json is removed before a new export starts and written last, so it
// exists only for a complete snapshot. A snapshot younger than 24 hours (same
// INACTIVE_MONTHS and time zone) is reused unless --refresh is given; the
// rules and amounts are always recomputed from it.

import fs from 'node:fs';

import { connect } from '../connect.js';
import { defaultSleep, ShopifyError } from '../shopify.js';
import { SHOP_INFO, REFRESH_CUSTOMERS } from '../queries.js';
import { findCampaignCards, campaignMarker } from '../giftcards.js';
import {
  campaignPaths,
  ensureDir,
  acquireRunLock,
  LockError,
  readJournal,
  foldJournal,
  appendJournal,
  newRunId,
  readJson,
  writeJsonAtomic,
} from '../campaign.js';
import { monthsAgoMidnight, isoWithOffset, localDate, localDateTime } from '../time.js';
import { parseCustomer } from './rules.js';
import { buildSelection, buildTestSelection } from './selection.js';
import { formatUsd, percentOf } from './amount.js';
import { exportCustomers, fetchActiveOrders, fetchOrderHistory, readSnapshot, readQuery, retryRead } from './export.js';

/** A snapshot younger than this is reused by a plain `select` (no --refresh). */
export const SNAPSHOT_MAX_AGE_MS = 24 * 60 * 60 * 1000;
/** How far back select looks for gift cards of this campaign: longer than any campaign lasts. */
export const CARD_LOOKBACK_MS = 400 * 24 * 60 * 60 * 1000;
/** nodes(ids:) accepts at most 250 ids per call. */
const NODES_BATCH = 250;

const fmt = (n) => Number(n ?? 0).toLocaleString('en-US');

/**
 * The local journal shows that issuing started (a live issue run or any card operation).
 * `issue --repair-only` settles leftovers and adds missing tags without creating a card,
 * so following these steps never issues more cards from the list being abandoned.
 */
export const ISSUING_STARTED_MESSAGE = '已经开始发放，不能重新运行 select。要另起一个新活动：先运行 DRY_RUN=false node index.js issue --repair-only（只补记和补打 tag，不建新卡），再运行 verify，确认没有“已建卡未打tag”的人，然后换一个 CAMPAIGN_ID 重新运行 select。';

/**
 * Shopify already holds `count` gift cards of this campaign although the journal does not say so.
 * `listExists` = campaigns/<id>/selection.json is on this computer. When it is not (another
 * computer, a fresh clone: campaigns/ is not in git), verify and every other command need that
 * folder too, so the only way forward is to copy the original folder back.
 */
export function cardsExistMessage(campaignId, count, { listExists = true } = {}) {
  if (!listExists) {
    return `本机没有这个活动的名单（campaigns/${campaignId}/selection.json），但 Shopify 上已有 ${fmt(count)} 张本活动（${campaignMarker(campaignId)}）的礼品卡：`
      + `已经开始发放，名单不能重新生成。请从原来运行 select 的电脑或备份，把整个 campaigns/${campaignId}/ 文件夹拷回这里，再运行 verify 核对。`;
  }
  return `Shopify 上已有 ${fmt(count)} 张本活动（${campaignMarker(campaignId)}）的礼品卡，说明已经开始发放（本地日志可能丢失或被移动过）。`
    + `名单已冻结，不能重新运行 select。请运行 verify 核对；不要删除 campaigns/${campaignId}/ 里的文件。`;
}

/** The journal shows issuing started but selection.json is gone: only restoring the folder helps. */
export function listLostMessage(campaignId) {
  return `本机的发放日志显示已经开始发放，但名单文件 campaigns/${campaignId}/selection.json 不见了：名单不能重新生成。`
    + `请从原来运行 select 的电脑或备份，把整个 campaigns/${campaignId}/ 文件夹拷回这里，再运行 verify 核对。`;
}

/** ISO-8601 truncated to whole seconds, for Shopify search. Rounding down only widens a ">=" window. */
function isoSeconds(ms) {
  return new Date(Math.floor(ms / 1000) * 1000).toISOString().replace('.000Z', 'Z');
}

/** The real Excel writer, imported lazily so this module loads without it. */
async function defaultWriteReport(options) {
  const { writeReport } = await import('../report/excel.js');
  return writeReport(options);
}

/**
 * Run `select`.
 * @returns {Promise<{ exitCode: number, selection: object|null }>} selection is
 *   the object written to selection.json (null when nothing was written).
 */
export async function runSelect({
  config,
  refresh = false,
  log = console,
  now = () => new Date(),
  sleep,
  writeReport,
  bulkTimeoutMs = 900_000,
  bulkPollMs = 5_000,
  stallMs = 180_000,
} = {}) {
  let paths;
  let release;
  try {
    paths = campaignPaths(config);
    ensureDir(paths.dir);
    release = acquireRunLock(paths, 'select');
  } catch (err) {
    // Nothing ran: the command holding the lock regenerates the Excel itself.
    log.error(err instanceof LockError ? err.message : `select 无法开始：${err.message}`);
    return { exitCode: 1, selection: null };
  }

  const ctx = {
    config,
    paths,
    refresh: !!refresh,
    log,
    now,
    sleep: sleep ?? defaultSleep,
    bulk: { bulkTimeoutMs, bulkPollMs, stallMs },
    stage: '读取发放日志',
  };
  const journalClock = { now: () => now().toISOString() };
  let exitCode = 1;
  let selection = null;
  let runId = null;
  let runSummary = null;

  try {
    // Refusals (list frozen) write nothing: no run record, no snapshot, no list.
    const ready = await prepare(ctx);
    if (ready) {
      const { timezone } = ready;
      ctx.stage = '写入运行记录';
      runId = newRunId(now());
      appendJournal(paths.journal, {
        op: 'run.start',
        run: runId,
        command: 'select',
        dryRun: false,
        batch: null,
        limit: null,
        options: { refresh: ctx.refresh, test: config.testCustomerIds.length > 0 },
      }, journalClock);

      const result = await makeSelection(ctx, timezone);
      ctx.stage = '保存名单';
      writeJsonAtomic(paths.selection, result.selection);
      selection = result.selection;
      exitCode = 0;
      runSummary = {
        recipients: selection.stats.recipients,
        totalCents: selection.stats.totalCents,
        source: selection.snapshot.source,
        snapshotReused: result.reused,
      };
      try {
        logSummary(selection, { paths, log, reused: result.reused });
      } catch (err) {
        log.warn(`名单已保存，但打印汇总时出错：${err.message}`);
      }
    }
  } catch (err) {
    exitCode = 1;
    selection = null;
    runSummary = { error: err?.message ?? String(err) };
    log.error(describeFailure(err, ctx.stage));
    log.error(fs.existsSync(paths.selection) ? '本次没有写入新名单，selection.json 仍是上一次的结果' : '本次没有生成名单');
  } finally {
    if (runId) {
      try {
        appendJournal(paths.journal, { op: 'run.end', run: runId, summary: runSummary, exitCode }, journalClock);
      } catch (err) {
        log.warn(`写入运行记录失败：${err.message}`);
      }
    }
    release();
    // Regenerate the Excel even after a failure (it then shows the previous
    // list). A report problem never changes the command's own result.
    if (fs.existsSync(paths.selection)) {
      try {
        await (writeReport ?? defaultWriteReport)({ config, paths, log, now });
      } catch (err) {
        log.warn(`Excel 没有生成：${err.message}。名单不受影响，可以稍后运行 export 重新生成`);
      }
    }
  }
  return { exitCode, selection };
}

// ---------------------------------------------------------------------------
// Before anything is written: is the list still allowed to change?
// ---------------------------------------------------------------------------

/**
 * Checks that the list is not frozen, first in the local journal, then in
 * Shopify (after connecting and reading the time zone). Returns { timezone }
 * when select may run, or null after logging why it must not. Writes nothing.
 * Live and test campaigns alike; the Shopify check also runs when the
 * customer snapshot would be reused.
 */
async function prepare(ctx) {
  const { config, paths, log } = ctx;
  if (foldJournal(readJournal(paths.journal)).issuingStarted) {
    // Amounts and order must not change mid-campaign. Without selection.json the A3 steps
    // (issue --repair-only, verify) cannot run either: the folder must be restored.
    log.error(fs.existsSync(paths.selection) ? ISSUING_STARTED_MESSAGE : listLostMessage(config.campaignId));
    return null;
  }
  ctx.stage = '连接 Shopify';
  await connect(config, { log, sleep: ctx.sleep });
  ctx.stage = '读取店铺时区';
  const timezone = config.timezone || (await shopTimezone(ctx));
  ctx.stage = '检查 Shopify 上已有的本活动礼品卡';
  const cards = await campaignCardsInShopify(ctx);
  if (cards.length) {
    // The journal may be lost, damaged or in another folder: the cards prove issuing started.
    log.error(cardsExistMessage(config.campaignId, cards.length, { listExists: fs.existsSync(paths.selection) }));
    return null;
  }
  return { timezone };
}

/**
 * The gift cards Shopify has with this campaign's marker, created within
 * CARD_LOOKBACK_MS (read-only; transient errors retried).
 */
function campaignCardsInShopify(ctx) {
  const sinceIso = isoSeconds(ctx.now().getTime() - CARD_LOOKBACK_MS);
  return retryRead(
    () => findCampaignCards({ campaignId: ctx.config.campaignId, sinceIso }),
    { sleep: ctx.sleep, log: ctx.log, label: '检查本活动的礼品卡' },
  );
}

// ---------------------------------------------------------------------------
// Building the selection
// ---------------------------------------------------------------------------

async function makeSelection(ctx, timezone) {
  if (ctx.config.testCustomerIds.length) return { selection: await selectTestCustomers(ctx, timezone), reused: false };
  return selectLive(ctx, timezone);
}

async function shopTimezone(ctx) {
  const data = await readQuery(SHOP_INFO, {}, { sleep: ctx.sleep, log: ctx.log, label: '读取店铺时区' });
  const timezone = data.shop?.ianaTimezone;
  if (!timezone) throw new Error('Shopify 没有返回店铺时区；请在 .env 里设置 TIMEZONE（例如 America/Los_Angeles）');
  return timezone;
}

/** Test campaign: exactly the TEST_CUSTOMER_IDS customers, fixed amount, no audience rules. */
async function selectTestCustomers(ctx, timezone) {
  const { config, log, now } = ctx;
  ctx.stage = '读取测试客户';
  const ids = [...new Set(config.testCustomerIds)];
  if (ids.length < config.testCustomerIds.length) log.warn(`TEST_CUSTOMER_IDS 里有重复的客户，已去掉重复，剩 ${ids.length} 个`);
  log.info(`测试活动：只取 TEST_CUSTOMER_IDS 里的 ${ids.length} 个客户，不套用筛选规则，每人固定 ${formatUsd(config.testGiftAmountCents)}`);

  const nodes = [];
  for (let i = 0; i < ids.length; i += NODES_BATCH) {
    const chunk = ids.slice(i, i + NODES_BATCH);
    const data = await readQuery(REFRESH_CUSTOMERS, { ids: chunk }, { sleep: ctx.sleep, log, label: '读取测试客户' });
    const byId = new Map((data.nodes ?? []).filter((n) => n?.id).map((n) => [n.id, n]));
    for (const id of chunk) {
      if (byId.has(id)) nodes.push(byId.get(id));
      else log.warn(`  测试客户 ${id} 在 Shopify 上找不到，已跳过`);
    }
  }

  ctx.stage = '计算名单';
  const customers = nodes.map(parseCustomer);
  const createdAt = now().toISOString();
  return buildTestSelection({
    config,
    customers,
    timezone,
    createdAt,
    snapshot: { exportedAt: createdAt, source: 'nodes', count: customers.length },
  });
}

/** Live campaign: whole-store snapshot (reused when fresh) → rules → selection. */
async function selectLive(ctx, timezone) {
  const { config, paths, log, now } = ctx;
  ctx.stage = '检查已导出的客户数据';
  const existing = readSnapshotMeta(paths, log);
  const decision = snapshotDecision(existing, { config, paths, timezone, nowMs: now().getTime(), refresh: ctx.refresh });
  let meta;
  if (decision.reuse) {
    meta = existing;
    log.info(`复用 ${localDateTime(Date.parse(meta.exportedAt), timezone)}（${timezone}）导出的客户数据：24 小时内的导出直接复用，要重新导出请加 --refresh`);
  } else {
    log.info(decision.why);
    meta = await takeSnapshot(ctx, timezone);
  }

  ctx.stage = '读取客户快照';
  const customers = await readSnapshot(paths);
  if (customers.length !== meta.count) {
    throw new Error(`customers.jsonl 里有 ${fmt(customers.length)} 个客户，导出记录却是 ${fmt(meta.count)} 个；文件可能被改动过，请运行 select --refresh 重新导出`);
  }
  const activeOrders = readJson(paths.activeOrders, null);
  if (!Array.isArray(activeOrders)) throw new Error('active-orders.json 缺失或格式不对，请运行 select --refresh 重新导出');
  const history = loadHistoryCache(paths, config, log);

  ctx.stage = '计算名单';
  log.info(`按规则筛选 ${fmt(customers.length)} 个客户（近 ${config.inactiveMonths} 个月从 ${localDate(meta.cutoffMs, timezone)} 起算）`);
  const input = {
    config,
    customers,
    activeOrders,
    nowMs: Date.parse(meta.exportedAt),
    cutoffMs: meta.cutoffMs,
    cutoffIso: meta.cutoffIso,
    timezone,
    createdAt: now().toISOString(),
    snapshot: { exportedAt: meta.exportedAt, source: meta.source, count: meta.count },
  };
  try {
    return { selection: buildSelection({ ...input, history }), reused: decision.reuse };
  } catch (err) {
    if (!err.missingHistory) throw err;
    // Candidates whose last order is cancelled / a test / $0: look further back.
    ctx.stage = '补查订单';
    const fetched = await fetchOrderHistory(err.missingHistory, { percent: config.giftPercent, log, sleep: ctx.sleep });
    for (const [id, value] of fetched) {
      // `deleted`: Shopify no longer has the customer (rule 12); kept so a reused snapshot excludes them too.
      history[id] = value.deleted
        ? { orders: value.orders, percent: config.giftPercent, deleted: true }
        : { orders: value.orders, percent: config.giftPercent };
    }
    writeJsonAtomic(paths.orderHistory, history);
    ctx.stage = '计算名单';
    return { selection: buildSelection({ ...input, history }), reused: decision.reuse };
  }
}

/** Export customers and recent orders, then commit the snapshot by writing snapshot.json last. */
async function takeSnapshot(ctx, timezone) {
  const { config, paths, log } = ctx;
  const exportedAt = ctx.now();
  const cutoffMs = monthsAgoMidnight(exportedAt, config.inactiveMonths, timezone);
  const cutoffIso = isoWithOffset(cutoffMs, timezone);
  // Without snapshot.json the data files are never reused, so an interrupted
  // export can not be mistaken for a complete one.
  fs.rmSync(paths.snapshotMeta, { force: true });

  ctx.stage = '导出全店客户';
  const exported = await exportCustomers({ paths, log, sleep: ctx.sleep, ...ctx.bulk });
  ctx.stage = '读取近期订单';
  const activeOrders = await fetchActiveOrders({ cutoffIso, log, sleep: ctx.sleep });
  writeJsonAtomic(paths.activeOrders, activeOrders);
  writeJsonAtomic(paths.orderHistory, {}); // follow-up lookups belong to one snapshot
  const meta = {
    exportedAt: exportedAt.toISOString(),
    source: exported.source,
    count: exported.count,
    rootObjectCount: exported.rootObjectCount,
    cutoffIso,
    cutoffMs,
    timezone,
    inactiveMonths: config.inactiveMonths,
  };
  writeJsonAtomic(paths.snapshotMeta, meta);
  return meta;
}

function readSnapshotMeta(paths, log) {
  try {
    return readJson(paths.snapshotMeta, null);
  } catch (err) {
    log.warn(`snapshot.json 无法读取（${err.message}），将重新导出`);
    return null;
  }
}

/** Whether the previous snapshot can be reused; `why` is the log line when it is not. */
function snapshotDecision(meta, { config, paths, timezone, nowMs, refresh }) {
  const again = (reason) => ({ reuse: false, why: `${reason}，重新导出客户数据` });
  if (refresh) return { reuse: false, why: '按 --refresh 重新导出客户数据' };
  if (!meta) return { reuse: false, why: '还没有导出过客户数据，现在导出' };
  const exportedMs = Date.parse(meta.exportedAt);
  if (!Number.isFinite(exportedMs) || !Number.isFinite(meta.cutoffMs) || !meta.cutoffIso || !Number.isInteger(meta.count) || !meta.source) {
    return again('上次导出的记录不完整');
  }
  if (!fs.existsSync(paths.snapshot) || !fs.existsSync(paths.activeOrders)) return again('上次导出的文件不完整');
  const ageMs = nowMs - exportedMs;
  if (ageMs < 0) return again('上次导出的时间晚于现在（电脑时间可能被调整过）');
  if (ageMs >= SNAPSHOT_MAX_AGE_MS) return again(`上次导出（${localDateTime(exportedMs, timezone)}）已超过 24 小时`);
  if (meta.inactiveMonths !== config.inactiveMonths) return again(`INACTIVE_MONTHS 从 ${meta.inactiveMonths} 改成了 ${config.inactiveMonths}`);
  if (meta.timezone !== timezone) return again(`店铺时区从 ${meta.timezone} 变成了 ${timezone}`);
  return { reuse: true, why: '' };
}

/**
 * Cached follow-up lookups of the current snapshot. An entry only counts for
 * the GIFT_PERCENT it was fetched with (the lookup stops at the first order
 * that is giftable at that percentage). A damaged cache is simply refetched.
 */
function loadHistoryCache(paths, config, log) {
  let cached;
  try {
    cached = readJson(paths.orderHistory, {});
  } catch (err) {
    log.warn(`order-history.json 无法读取（${err.message}），重新补查`);
    cached = {};
  }
  const history = {};
  if (cached && typeof cached === 'object' && !Array.isArray(cached)) {
    for (const [id, value] of Object.entries(cached)) {
      if (value && Array.isArray(value.orders) && value.percent === config.giftPercent) history[id] = value;
    }
  }
  return history;
}

// ---------------------------------------------------------------------------
// Output
// ---------------------------------------------------------------------------

function describeFailure(err, stage) {
  const message = err?.message ?? String(err);
  if (err instanceof ShopifyError && err.code === 'SEARCH_WARNING') {
    return `select 已停止（${stage}）：Shopify 忽略了查询里的搜索条件，查到的数据不可信，不能用来生成名单。${message}`;
  }
  if (err instanceof ShopifyError) return `select 失败（${stage}）：Shopify 请求出错：${message}`;
  return `select 失败（${stage}）：${message}`;
}

function logSummary(selection, { paths, log, reused }) {
  const s = selection.stats;
  const p = selection.params;
  const test = selection.mode === 'test';
  log.info('');
  log.info(`select 完成（${test ? '测试活动' : '正式活动'} ${selection.campaignId}）`);
  if (test) {
    log.info(`  测试客户：${fmt(selection.funnel.total)} 个`);
    log.info(`  入选：${fmt(s.recipients)} 人，每人固定 ${formatUsd(p.testGiftAmountCents)}`);
    const skipped = selection.funnel.byRule.filter((r) => r.count);
    if (skipped.length) log.info(`  未入选：${skipped.map((r) => `${r.label} ${fmt(r.count)} 人`).join('；')}`);
  } else {
    const source = selection.snapshot.source === 'bulk' ? '全店批量导出' : '普通分页导出';
    const exportedAt = localDateTime(Date.parse(selection.snapshot.exportedAt), p.timezone);
    log.info(`  客户数据：${reused ? '复用 24 小时内的导出' : '本次新导出'}（${source}，导出于 ${exportedAt}，${p.timezone}）`);
    log.info(`  近 ${p.inactiveMonths} 个月从 ${p.cutoffDate} 起算`);
    log.info(`  全店客户：${fmt(selection.funnel.total)}`);
    log.info('  未入选（按顺序判断，记第一个不满足的条件）：');
    for (const r of selection.funnel.byRule) log.info(`    ${r.n}. ${r.label}：${fmt(r.count)}`);
    log.info(`  入选：${fmt(s.recipients)} 人`);
    log.info(`    有下单：${fmt(s.orderedCount)} 人，合计 ${formatUsd(s.orderedCents)}`);
    const average = selection.averageCents === null || selection.averageCents === undefined
      ? ''
      : `；平均上次订单 ${formatUsd(Math.round(selection.averageCents))}，${p.giftPercent}% = ${formatUsd(percentOf(selection.averageCents, p.giftPercent))} → ${formatUsd(selection.neverAmountCents)}`;
    log.info(`    从没下单：${fmt(s.neverCount)} 人，合计 ${formatUsd(s.neverCents)}${average}`);
    log.info('  按档位：');
    for (const t of s.tiers) {
      log.info(`    ${t.label}：${fmt(t.count)} 人（有下单 ${fmt(t.ordered)}，从没下单 ${fmt(t.never)}），合计 ${formatUsd(t.cents)}`);
    }
    if (s.duplicates.groups) {
      const bulk = s.duplicates.flaggedBulk ? `；${fmt(s.duplicates.flaggedBulk)} 组有 10 个以上账户，Excel 里标"疑似批量注册"` : '';
      log.info(`  同地址去重：${fmt(s.duplicates.groups)} 组、${fmt(s.duplicates.accounts)} 个账户，少发 ${fmt(s.duplicates.removed)} 人${bulk}`);
    }
  }
  log.info(`  总面额：${formatUsd(s.totalCents)}`);
  if (!s.recipients) log.warn('  名单是空的：没有任何客户入选');
  log.info(`  名单文件：${paths.selection}`);
  log.info(`  Excel：${paths.excel}`);
}
