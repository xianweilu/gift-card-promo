// `export`: rebuild the campaign workbook at any time, also while another
// command is running (plan: "其他命令运行时跑 export 会怎样").
//
// It only reads local files (selection.json, journal.jsonl, tags.json,
// verify.json, usage.json — through src/report/excel.js). It never takes the
// run lock and never writes the journal, so it cannot wait for, interrupt or
// slow down issue/remind; the report module serialises writers of the .xlsx
// with its own short Excel lock. `--refresh` first re-reads from Shopify
// (read-only) which customers carry SENT_TAG and stores that in tags.json.
//
// The helpers exported next to runExport (tag snapshot, end-of-command
// workbook refresh, missing-list message) are shared with src/verify.js.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { campaignPaths, runningCommand, writeJsonAtomic } from './campaign.js';
import { connect } from './connect.js';
import { fetchTaggedCustomerIds } from './customers.js';

/** What an export taken while `command` holds the run lock shows (wording from the plan's table). */
const WHILE_RUNNING = {
  select: 'select 正在重新运行：导出的是上一次的名单；select 跑完后会自动生成新版本',
  issue: 'issue 正在运行：导出的是此刻的进度，正在处理的那一个人显示"进行中"；issue 结束时会再写一次最终版',
  remind: 'remind 正在运行：导出的是此刻的进度；remind 结束时会再写一次最终版',
  verify: 'verify 正在运行：导出的是它开始前的状态；verify 跑完后会自动生成新版本',
  usage: 'usage 正在运行：导出的是它开始前的状态；usage 跑完后会自动生成新版本',
};

// ---------------------------------------------------------------------------
// Shared helpers (also used by src/verify.js)
// ---------------------------------------------------------------------------

/**
 * Order Shopify GIDs by their numeric tail (gid://shopify/Customer/9 before
 * …/10); null/undefined sort last. Used to keep written files stable.
 */
export function compareGids(a, b) {
  if (a === b) return 0;
  if (a === null || a === undefined) return 1;
  if (b === null || b === undefined) return -1;
  const x = String(a).split('/').pop();
  const y = String(b).split('/').pop();
  if (/^\d+$/.test(x) && /^\d+$/.test(y) && x !== y) {
    return x.length - y.length || (x < y ? -1 : 1);
  }
  return String(a) < String(b) ? -1 : String(a) > String(b) ? 1 : 0;
}

/**
 * The message shown when selection.json does not exist (yet). A list that is missing
 * after issuing started (another computer, a fresh clone: campaigns/ is not in git) can
 * not be made again: select refuses once Shopify has this campaign's cards, so the
 * original campaign folder has to be copied back instead.
 * @param {object} paths campaignPaths(config)
 * @param {string} [campaignId] CAMPAIGN_ID; defaults to the campaign folder's name
 */
export function missingSelectionMessage(paths, campaignId = path.basename(paths.dir)) {
  if (runningCommand(paths) === 'select') return '还没有名单，select 正在运行，请等它跑完';
  return `还没有名单，请先运行 select。如果已经开始发放，不要重新运行 select，请把原来的 campaigns/${campaignId}/ 文件夹拷回来。`;
}

/**
 * Read from Shopify which customers carry SENT_TAG and store the snapshot in
 * tags.json ({ fetchedAt, tag, ids }), the source of the workbook's tag column.
 * The client must already be connected. `fetchedAt` is taken before the first
 * page, so the snapshot is at least that fresh.
 * @returns {Promise<{ fetchedAt: string, tag: string, ids: string[], idSet: Set<string> }>}
 */
export async function fetchTagSnapshot({ config, paths, now = () => new Date() }) {
  const fetchedAt = now().toISOString();
  // No logger on purpose: the shared helper's message ("... will be skipped") is meant for issue.
  const idSet = await fetchTaggedCustomerIds(config.sentTag);
  const snapshot = { fetchedAt, tag: config.sentTag, ids: [...idSet].sort(compareGids) };
  writeJsonAtomic(paths.tags, snapshot);
  return { ...snapshot, idSet };
}

/** The injected writeReport, or the real one from src/report/excel.js (loaded only when needed). */
export async function resolveWriteReport(writeReport) {
  if (typeof writeReport === 'function') return writeReport;
  const mod = await import('./report/excel.js');
  return mod.writeReport;
}

/**
 * Print what writeReport produced: the workbook path and the extra copy.
 * Its warnings are not printed here: writeReport logs each one itself when it
 * happens (result.warnings is only a copy for callers and tests).
 */
export function logReportResult(result, log, verb = '已更新') {
  if (result?.file) log.info(`Excel ${verb}：${result.file}`);
  if (result?.out) log.info(`另存一份：${result.out}`);
}

/**
 * End-of-command workbook refresh for commands whose own outcome matters more
 * than the Excel (verify). Never throws: a failure becomes a warning, so it
 * can never hide or replace the command's own result.
 */
export async function refreshReport({ config, paths, log = console, now = () => new Date(), writeReport }) {
  try {
    const write = await resolveWriteReport(writeReport);
    const result = await write({ config, paths, log, now });
    logReportResult(result, log);
    return result ?? null;
  } catch (err) {
    log.warn(`Excel 没有更新：${err.message}。可以稍后运行 node index.js export 重新生成`);
    return null;
  }
}

// ---------------------------------------------------------------------------
// --out
// ---------------------------------------------------------------------------

/**
 * Turn the `--out` value into the absolute path of the extra copy:
 *   "~/..." expands to the home directory; a relative path is taken from the
 *   current directory; an existing directory (or a path ending in "/") gets the
 *   workbook's own file name; a missing ".xlsx" extension is added.
 * Returns null when there is nothing extra to save (no value, or the main
 * workbook itself). Throws when the target folder does not exist.
 */
export function resolveOutPath(out, paths) {
  if (out === null || out === undefined) return null;
  let p = String(out).trim();
  if (!p) return null;
  if (p === '~') p = os.homedir();
  else if (p.startsWith('~/')) p = path.join(os.homedir(), p.slice(2));
  const endsWithSlash = /[\\/]$/.test(p);
  p = path.resolve(p);

  let isDir = false;
  try {
    isDir = fs.statSync(p).isDirectory();
  } catch {
    /* does not exist yet: it is the file to create */
  }
  if (isDir || endsWithSlash) p = path.join(p, path.basename(paths.excel));
  else if (path.extname(p).toLowerCase() !== '.xlsx') p = `${p}.xlsx`;

  const dir = path.dirname(p);
  let dirOk = false;
  try {
    dirOk = fs.statSync(dir).isDirectory();
  } catch {
    dirOk = false;
  }
  if (!dirOk) throw new Error(`--out 指定的文件夹不存在：${dir}`);
  if (p === path.resolve(paths.excel)) return null; // the main workbook is written anyway
  return p;
}

// ---------------------------------------------------------------------------
// The command
// ---------------------------------------------------------------------------

/**
 * `node index.js export [--refresh] [--out <path>]`
 *
 * @param {object} options
 * @param {object} options.config loadConfig() result
 * @param {boolean} [options.refresh] re-read SENT_TAG holders from Shopify first (read-only)
 * @param {string|null} [options.out] also save a copy here
 * @param {{info: Function, warn: Function, error: Function}} [options.log]
 * @param {() => Date} [options.now]
 * @param {(ms: number) => Promise<void>} [options.sleep] passed to the Shopify client (rate-limit waits)
 * @param {Function} [options.writeReport] defaults to src/report/excel.js
 * @returns {Promise<{ exitCode: number, file: string|null, out: string|null, warnings: string[] }>}
 *   exitCode 0 = exported; 1 = no list yet, the refresh failed (the workbook is
 *   still exported with the previous tag data), the --out copy could not be
 *   saved (the main workbook is still exported and its path printed) or the
 *   workbook could not be written; 2 = bad --out.
 */
export async function runExport({ config, refresh = false, out = null, log = console, now = () => new Date(), sleep, writeReport } = {}) {
  const paths = campaignPaths(config);

  if (!fs.existsSync(paths.selection)) {
    log.error(missingSelectionMessage(paths, config.campaignId));
    return { exitCode: 1, file: null, out: null, warnings: [] };
  }

  let outPath = null;
  try {
    outPath = resolveOutPath(out, paths);
  } catch (err) {
    log.error(err.message);
    return { exitCode: 2, file: null, out: null, warnings: [] };
  }

  log.info(`导出 Excel：只读本地的名单和日志${refresh ? '，先从 Shopify 只读刷新 tag' : ''}，不会改 Shopify 上的任何东西`);
  const running = runningCommand(paths);
  if (running && WHILE_RUNNING[running]) log.info(`注意：${WHILE_RUNNING[running]}`);

  let exitCode = 0;
  if (refresh) {
    log.info(`从 Shopify 查询带 ${config.sentTag} 的客户…`);
    try {
      await connect(config, { log, sleep });
      const snapshot = await fetchTagSnapshot({ config, paths, now });
      log.info(`带 ${config.sentTag} 的客户：${snapshot.ids.length} 个（已写入 ${paths.tags}）`);
    } catch (err) {
      exitCode = 1;
      log.error(`刷新 tag 失败：${err.message}`);
      log.warn(`Excel 仍会导出，但"是否有 ${config.sentTag}"一列没有刷新`);
    }
  }

  let result;
  try {
    const write = await resolveWriteReport(writeReport);
    result = await write({ config, paths, log, now, out: outPath });
  } catch (err) {
    log.error(`导出 Excel 失败：${err.message}`);
    return { exitCode: 1, file: null, out: null, warnings: [] };
  }
  const file = result?.file ?? paths.excel;
  logReportResult({ ...result, file }, log, '已导出');
  // writeReport turns a failed --out copy into a warning (already printed) and out: null;
  // the main workbook above is fine, but the copy that was asked for does not exist.
  if (outPath && !result?.out) {
    exitCode = 1;
    log.error(`没有另存到 ${outPath}（原因见上面的提示）；主 Excel 已更新：${file}`);
  }
  return { exitCode, file, out: result?.out ?? null, warnings: result?.warnings ?? [] };
}
