// Read-only data fetchers for `select` (step 1):
//   exportCustomers    every customer → campaigns/<id>/customers.jsonl
//                      (bulk operation; cursor pagination when the bulk job does not work out)
//   fetchActiveOrders  orders created since the cutoff (who ordered recently, and at which address)
//   fetchOrderHistory  newest orders of the few candidates whose last order cannot be the basis
//   readSnapshot       customers.jsonl → parsed customers
//
// Nothing here changes store data: the only mutations are the bookkeeping
// calls that start or cancel the export job. The download URL of a finished
// bulk job is a signed link to customer data, so it is never logged or stored.

import fs from 'node:fs';
import path from 'node:path';
import { Readable } from 'node:stream';

import { gql, defaultSleep, ShopifyError } from '../shopify.js';
import {
  CUSTOMER_EXPORT_BULK,
  RUN_BULK_QUERY,
  BULK_STATUS,
  BULK_CANCEL,
  CUSTOMERS_PAGE,
  ORDERS_SINCE,
  CUSTOMER_ORDERS,
} from '../queries.js';
import { ensureDir } from '../campaign.js';
import { parseCustomer, parseOrder, isValidOrder } from './rules.js';
import { isGiftableTotal } from './amount.js';

/** Follow-up lookups stop after this many pages (50 orders each) per customer. */
export const HISTORY_MAX_PAGES = 20;
/** Tries for a read-only query whose failure may be transient (network, 5xx, empty response). */
const READ_ATTEMPTS = 4;
/** Consecutive failed status polls tolerated before the bulk job is abandoned. */
const POLL_ERROR_LIMIT = 3;
/** Tries for downloading the finished export when the transfer itself fails. */
const DOWNLOAD_ATTEMPTS = 3;
/** Ceiling for downloading the export file (about 111 MB for 198k customers). */
const DOWNLOAD_TIMEOUT_MS = 30 * 60_000;
/** Snapshot lines are written to disk in chunks of roughly this many characters. */
const WRITE_CHUNK_CHARS = 1 << 20;
/** Bulk job states that are still worth waiting for; anything else but COMPLETED is a failure. */
const BULK_PENDING = new Set(['CREATED', 'RUNNING']);
const CUSTOMER_GID = /^gid:\/\/shopify\/Customer\/\d+$/;

const fmt = (n) => Number(n ?? 0).toLocaleString('en-US');

function duration(ms) {
  if (ms >= 60_000) return `${Number((ms / 60_000).toFixed(1))} 分钟`;
  if (ms >= 1_000) return `${Number((ms / 1_000).toFixed(1))} 秒`;
  return `${ms} 毫秒`;
}

/**
 * Removes the signed export URL (and any other link, with its query string)
 * from a message before it can be logged. Whole links go first, so a variant
 * of the URL can not leave its signature behind.
 */
function redact(message, url) {
  let text = String(message ?? '').replace(/https?:\/\/\S+/g, '<链接已隐藏>');
  if (url) text = text.split(url).join('<链接已隐藏>');
  return text;
}

/**
 * The export cannot be used as downloaded (a bad line, a repeated id, or the
 * link was refused): unlike a dropped transfer, downloading again will not help.
 */
class UnusableExportError extends Error {
  constructor(message) {
    super(message);
    this.name = 'UnusableExportError';
  }
}

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

/**
 * Runs a read-only Shopify call `fn`, retried a few times when the failure may
 * be transient (network error, 5xx, missing or non-JSON response; i.e.
 * `outcomeKnown === false`). Retrying a read is always safe. gql() already
 * waits out throttling itself; search warnings and other definite errors are
 * never retried. `fn` must not change store data.
 */
export async function retryRead(fn, { sleep = defaultSleep, log = console, label = 'Shopify 查询' } = {}) {
  for (let attempt = 1; ; attempt += 1) {
    try {
      return await fn();
    } catch (err) {
      const transient = err instanceof ShopifyError && err.outcomeKnown === false;
      if (!transient || attempt >= READ_ATTEMPTS) throw err;
      const waitMs = 2_000 * 2 ** (attempt - 1);
      log.warn(`  ${label}：Shopify 暂时出错（${err.message}），${waitMs / 1000} 秒后重试（第 ${attempt}/${READ_ATTEMPTS - 1} 次重试）`);
      await sleep(waitMs);
    }
  }
}

/** gql() for a read-only query, with retryRead()'s retries. */
export function readQuery(query, variables, options = {}) {
  return retryRead(() => gql(query, variables), options);
}

/**
 * The lines of a text stream (its chunks must be strings: set an encoding),
 * split on "\n" only. node:readline also ends a line at U+2028 / U+2029, which
 * JSON.stringify leaves unescaped inside strings, so one customer whose name,
 * address or tag holds such a character (common in text pasted from PDFs or
 * Word) would be cut in two and the whole snapshot rejected. Lines keep any
 * "\r" and surrounding blanks: callers trim them.
 */
export async function* textLines(input) {
  let rest = '';
  for await (const chunk of input) {
    if (typeof chunk !== 'string') throw new TypeError('textLines needs a stream that yields strings (set an encoding)');
    if (!chunk.includes('\n')) {
      rest += chunk; // a line longer than one chunk: keep collecting
      continue;
    }
    const parts = (rest + chunk).split('\n');
    rest = parts.pop();
    yield* parts;
  }
  if (rest) yield rest;
}

/** The cursor of the next page, or null on the last page. Guards against endless loops. */
function nextCursor(pageInfo, after, what) {
  if (!pageInfo?.hasNextPage) return null;
  const next = pageInfo.endCursor;
  if (!next) throw new Error(`${what}：Shopify 表示还有下一页，却没有给出分页游标`);
  if (next === after) throw new Error(`${what}：Shopify 连续两次返回同一个分页游标，已停止以免死循环`);
  return next;
}

/**
 * Buffered line writer for `<target>.tmp`. commit() fsyncs and renames the
 * temp file over `target` (readers never see half a snapshot); abort() deletes it.
 */
function openLineWriter(target) {
  ensureDir(path.dirname(target));
  const tmp = `${target}.tmp`;
  const fd = fs.openSync(tmp, 'w');
  let chunks = [];
  let size = 0;
  let open = true;
  const flush = () => {
    if (!chunks.length) return;
    fs.writeFileSync(fd, chunks.join('')); // writeFileSync loops over partial writes
    chunks = [];
    size = 0;
  };
  return {
    write(line) {
      chunks.push(line, '\n');
      size += line.length + 1;
      if (size >= WRITE_CHUNK_CHARS) flush();
    },
    commit() {
      flush();
      fs.fsyncSync(fd);
      fs.closeSync(fd);
      open = false;
      fs.renameSync(tmp, target);
    },
    abort() {
      if (open) {
        open = false;
        try {
          fs.closeSync(fd);
        } catch {
          /* already closed */
        }
      }
      fs.rmSync(tmp, { force: true });
    },
  };
}

// ---------------------------------------------------------------------------
// Whole-store customer export
// ---------------------------------------------------------------------------

/**
 * Export every customer into paths.snapshot, one customer node (CUSTOMER
 * fields of src/queries.js) per JSON line. Tries a bulk operation first and
 * falls back to cursor pagination when the job is rejected, fails, stalls,
 * runs too long, or its output does not validate. The previous snapshot file
 * is replaced only by a complete, validated one.
 *
 * @param {object} options
 * @param {object} options.paths campaignPaths(config)
 * @param {number} [options.bulkTimeoutMs] give up on the bulk job after this long
 * @param {number} [options.bulkPollMs] status poll interval
 * @param {number} [options.stallMs] give up when objectCount stays unchanged this long
 * @returns {Promise<{ count: number, source: 'bulk'|'paginated', rootObjectCount: number|null }>}
 *   rootObjectCount is what Shopify reported for the bulk job (null after pagination).
 */
export async function exportCustomers({
  paths,
  log = console,
  sleep = defaultSleep,
  bulkTimeoutMs = 900_000,
  bulkPollMs = 5_000,
  stallMs = 180_000,
  downloadTimeoutMs = DOWNLOAD_TIMEOUT_MS,
} = {}) {
  if (!paths?.snapshot) throw new Error('exportCustomers needs campaign paths');
  log.info('导出全店客户：提交 Shopify 批量导出任务');
  const bulk = await tryBulkExport({ paths, log, sleep, bulkTimeoutMs, bulkPollMs, stallMs, downloadTimeoutMs });
  if (bulk.ok) {
    log.info(`全店客户导出完成：${fmt(bulk.count)} 个客户`);
    return { count: bulk.count, source: 'bulk', rootObjectCount: bulk.rootObjectCount };
  }
  log.warn(`全店批量导出没有成功：${bulk.reason}。改用普通分页导出（每页 250 个客户，20 万客户约需 26 分钟）`);
  return paginateCustomers({ paths, log, sleep });
}

const failed = (reason) => ({ ok: false, reason });

/**
 * One bulk-export attempt. Never throws: every problem becomes
 * { ok: false, reason } so that the caller can fall back to pagination.
 */
async function tryBulkExport({ paths, log, sleep, bulkTimeoutMs, bulkPollMs, stallMs, downloadTimeoutMs }) {
  let id;
  try {
    // Starting an export changes no store data, so a transient failure is
    // retried like a read: at worst an orphaned job runs to completion unused.
    const data = await readQuery(RUN_BULK_QUERY, { query: CUSTOMER_EXPORT_BULK }, { sleep, log, label: '提交导出任务' });
    const payload = data.bulkOperationRunQuery;
    const userErrors = payload?.userErrors ?? [];
    if (userErrors.length) return failed(`Shopify 拒绝了导出任务（${userErrors.map((e) => e.message).join('；')}）`);
    id = payload?.bulkOperation?.id;
    if (!id) return failed('Shopify 没有返回导出任务的 ID');
  } catch (err) {
    return failed(`提交导出任务出错（${err.message}）`);
  }
  log.info(`  导出任务已提交，每 ${duration(bulkPollMs)}查询一次进度`);

  // Elapsed time is the larger of the wall clock and the total time slept, so
  // the limits hold in production and are deterministic with an injected sleep.
  const startedAt = Date.now();
  let sleptMs = 0;
  let lastCount = 0;
  let lastProgressMs = 0;
  let pollErrors = 0;
  for (;;) {
    await sleep(bulkPollMs);
    sleptMs += bulkPollMs;
    const elapsedMs = Math.max(Date.now() - startedAt, sleptMs);

    let op = null;
    let pollError = null;
    try {
      op = (await gql(BULK_STATUS, { id })).bulkOperation ?? null;
    } catch (err) {
      pollError = err;
    }

    if (pollError) {
      // A status poll is a plain read: tolerate a few transient failures.
      pollErrors += 1;
      if (pollErrors >= POLL_ERROR_LIMIT) {
        await cancelBulk(id, log);
        return failed(`连续 ${pollErrors} 次查询导出进度出错（${pollError.message}）`);
      }
      log.warn(`  查询导出进度出错（${pollError.message}），稍后再查`);
    } else {
      pollErrors = 0;
      if (!op) return failed('Shopify 找不到这个导出任务');
      if (op.status === 'COMPLETED') return finishBulk(op, { paths, log, sleep, downloadTimeoutMs });
      if (!BULK_PENDING.has(op.status)) {
        return failed(`导出任务状态为 ${op.status}${op.errorCode ? `（${op.errorCode}）` : ''}`);
      }
      const count = Number(op.objectCount) || 0;
      if (count !== lastCount) {
        lastCount = count;
        lastProgressMs = elapsedMs;
        log.info(`全店客户导出中：已完成 ${fmt(count)}`);
      }
    }

    if (elapsedMs >= bulkTimeoutMs) {
      await cancelBulk(id, log);
      return failed(`超过 ${duration(bulkTimeoutMs)}仍未完成`);
    }
    if (elapsedMs - lastProgressMs >= stallMs) {
      await cancelBulk(id, log);
      return failed(`进度停在 ${fmt(lastCount)} 已超过 ${duration(stallMs)}`);
    }
  }
}

/** Best-effort cancel of an abandoned job; a job left running finishes on its own and changes nothing. */
async function cancelBulk(id, log) {
  try {
    const payload = (await gql(BULK_CANCEL, { id })).bulkOperationCancel;
    const userErrors = payload?.userErrors ?? [];
    if (userErrors.length) {
      log.warn(`  取消导出任务没有成功（${userErrors.map((e) => e.message).join('；')}）；它会在 Shopify 上自行结束，不影响结果`);
    } else {
      log.info('  已取消 Shopify 上的导出任务');
    }
  } catch (err) {
    log.warn(`  取消导出任务出错（${err.message}）；它会在 Shopify 上自行结束，不影响结果`);
  }
}

/** A non-negative integer count from Shopify's UnsignedInt64 string, or null. */
function toCount(value) {
  if (value === null || value === undefined || value === '') return null;
  const n = Number(value);
  return Number.isSafeInteger(n) && n >= 0 ? n : null;
}

/**
 * Download and validate a COMPLETED job's output into the snapshot. Never
 * throws. A failed transfer is retried (the signed link stays valid for days);
 * a file that is wrong (bad line, repeated id, wrong count) is not.
 */
async function finishBulk(op, { paths, log, sleep, downloadTimeoutMs }) {
  const expected = toCount(op.rootObjectCount);
  if (expected === null) return failed('Shopify 没有返回导出的客户数（rootObjectCount）');
  if (!op.url) {
    if (expected !== 0) return failed(`Shopify 报告导出了 ${fmt(expected)} 个客户，却没有给出下载链接`);
    try {
      openLineWriter(paths.snapshot).commit(); // an empty snapshot: the store has no customers
    } catch (err) {
      return failed(`写入客户快照出错（${err.message}）`);
    }
    log.info('  Shopify 导出完成：店铺里没有客户');
    return { ok: true, count: 0, rootObjectCount: 0 };
  }

  log.info(`  Shopify 导出完成（${fmt(expected)} 个客户），正在下载导出文件`);
  for (let attempt = 1; ; attempt += 1) {
    let writer = null;
    try {
      writer = openLineWriter(paths.snapshot);
      const count = await downloadLines(op.url, writer, { log, timeoutMs: downloadTimeoutMs });
      if (count !== expected) {
        writer.abort();
        return failed(`导出文件有 ${fmt(count)} 个客户，和 Shopify 报告的 ${fmt(expected)} 个不一致`);
      }
      writer.commit();
      return { ok: true, count, rootObjectCount: expected };
    } catch (err) {
      writer?.abort();
      const reason = redact(err.message, op.url);
      if (err instanceof UnusableExportError || attempt >= DOWNLOAD_ATTEMPTS) return failed(`下载或检查导出文件出错（${reason}）`);
      const waitMs = 5_000 * attempt;
      log.warn(`  下载导出文件出错（${reason}），${waitMs / 1000} 秒后重新下载（第 ${attempt}/${DOWNLOAD_ATTEMPTS - 1} 次重试）`);
      await sleep(waitMs);
    }
  }
}

/**
 * Stream the JSONL export line by line into `writer`. Every line must be one
 * customer (no child rows) and no id may repeat. Returns the number of lines.
 */
async function downloadLines(url, writer, { log, timeoutMs }) {
  // The URL is itself the credential (a signed, short-lived link): no Shopify
  // token is sent with it, and it never appears in a log line or an error.
  const res = await fetch(url, { signal: AbortSignal.timeout(timeoutMs) });
  if (!res.ok) {
    await res.body?.cancel().catch(() => {});
    const message = `下载失败（HTTP ${res.status}）`;
    throw res.status < 500 ? new UnusableExportError(message) : new Error(message);
  }
  if (!res.body) return 0;

  // The encoding makes the stream decode UTF-8 across chunk boundaries.
  const input = Readable.fromWeb(res.body, { encoding: 'utf8' });
  const ids = new Set();
  let count = 0;
  try {
    for await (const raw of textLines(input)) {
      const line = raw.trim(); // also drops the "\r" of a CRLF line end and a leading BOM
      if (!line) continue;
      count += 1;
      let node;
      try {
        node = JSON.parse(line);
      } catch {
        throw new UnusableExportError(`第 ${fmt(count)} 行不是有效的 JSON`);
      }
      if (!node || typeof node !== 'object' || !CUSTOMER_GID.test(node.id ?? '') || node.__parentId) {
        throw new UnusableExportError(`第 ${fmt(count)} 行不是客户记录`);
      }
      if (ids.has(node.id)) throw new UnusableExportError(`客户 ${node.id} 在导出文件里出现了不止一次`);
      ids.add(node.id);
      writer.write(line);
      if (count % 50_000 === 0) log.info(`  下载中：已写入 ${fmt(count)} 个客户`);
    }
  } finally {
    input.destroy(); // stops the download when validation failed midway
  }
  return count;
}

/** Fallback: the same customer fields, 250 per page, sorted by id. */
async function paginateCustomers({ paths, log, sleep }) {
  const writer = openLineWriter(paths.snapshot);
  const ids = new Set();
  let duplicates = 0;
  let pages = 0;
  let after = null;
  try {
    do {
      const data = await readQuery(CUSTOMERS_PAGE, { after }, { sleep, log, label: '分页导出客户' });
      const connection = data.customers;
      if (!connection) throw new Error('分页导出客户：Shopify 没有返回客户列表');
      pages += 1;
      for (const node of connection.nodes ?? []) {
        if (!node?.id) continue;
        if (ids.has(node.id)) {
          duplicates += 1;
          continue;
        }
        ids.add(node.id);
        writer.write(JSON.stringify(node));
      }
      after = nextCursor(connection.pageInfo, after, '分页导出客户');
      if (after && pages % 20 === 0) log.info(`全店客户分页导出中：已完成 ${fmt(ids.size)}`);
    } while (after);
    writer.commit();
  } catch (err) {
    writer.abort();
    throw err;
  }
  if (duplicates) log.warn(`  分页导出时有 ${fmt(duplicates)} 个客户重复出现，只保留了一次`);
  log.info(`全店客户分页导出完成：${fmt(ids.size)} 个客户（${fmt(pages)} 页）`);
  return { count: ids.size, source: 'paginated', rootObjectCount: null };
}

// ---------------------------------------------------------------------------
// Orders
// ---------------------------------------------------------------------------

/**
 * Every order created since `cutoffIso` (ORDERS_SINCE nodes, oldest first),
 * including cancelled and test orders: rules.buildActivity decides what counts.
 * A Shopify search warning (an ignored filter) is fatal and propagates.
 */
export async function fetchActiveOrders({ cutoffIso, log = console, sleep = defaultSleep } = {}) {
  if (!cutoffIso) throw new Error('fetchActiveOrders needs cutoffIso');
  const query = `created_at:>='${cutoffIso}'`;
  log.info(`读取 ${cutoffIso} 以来的订单（判断谁近期下过单）`);
  const nodes = [];
  let pages = 0;
  let after = null;
  do {
    const data = await readQuery(ORDERS_SINCE, { query, after }, { sleep, log, label: '读取近期订单' });
    const connection = data.orders;
    if (!connection) throw new Error('读取近期订单：Shopify 没有返回订单列表');
    for (const node of connection.nodes ?? []) nodes.push(node);
    pages += 1;
    after = nextCursor(connection.pageInfo, after, '读取近期订单');
    if (after && pages % 10 === 0) log.info(`  近期订单读取中：已读取 ${fmt(nodes.length)} 单`);
  } while (after);
  log.info(`近期订单：共 ${fmt(nodes.length)} 单`);
  return nodes;
}

/**
 * Newest orders of each customer, for candidates whose last order is
 * cancelled, a test or too small to be the basis. Pages CUSTOMER_ORDERS
 * (newest first) until a valid order with a giftable total is found, the
 * orders run out, or `maxPages` pages were read. The list ends with that
 * basis order. A customer that no longer exists (deleted after the export)
 * gets `{ orders: [], deleted: true }`; rules.finalize leaves them out.
 *
 * @returns {Promise<Map<string, { orders: object[], deleted?: true }>>} raw Order nodes, newest first
 */
export async function fetchOrderHistory(customerIds, { percent, log = console, sleep = defaultSleep, maxPages = HISTORY_MAX_PAGES } = {}) {
  if (!(Number(percent) > 0)) throw new Error('fetchOrderHistory needs the gift percent');
  const ids = [...new Set(customerIds)];
  const history = new Map();
  if (!ids.length) return history;
  log.info(`补查 ${fmt(ids.length)} 位候选人的订单（最近一笔已取消、是测试单或金额为 $0，要往前找付过钱的订单）`);
  let missing = 0;
  let capped = 0;
  for (let i = 0; i < ids.length; i += 1) {
    const result = await customerOrders(ids[i], { percent, log, sleep, maxPages });
    history.set(ids[i], result.missing ? { orders: result.orders, deleted: true } : { orders: result.orders });
    if (result.missing) missing += 1;
    if (result.capped) capped += 1;
    if ((i + 1) % 25 === 0 && i + 1 < ids.length) log.info(`  补查订单：已完成 ${fmt(i + 1)}/${fmt(ids.length)}`);
  }
  log.info(`补查订单完成：${fmt(ids.length)} 位`);
  if (missing) log.warn(`  其中 ${fmt(missing)} 位客户在 Shopify 上已被删除，不进发放名单`);
  if (capped) log.warn(`  其中 ${fmt(capped)} 位客户查了 ${maxPages} 页（${fmt(maxPages * 50)} 笔订单）仍没找到付过钱的有效订单，按已查到的订单处理`);
  return history;
}

async function customerOrders(id, { percent, log, sleep, maxPages }) {
  const orders = [];
  let after = null;
  for (let page = 0; page < maxPages; page += 1) {
    const data = await readQuery(CUSTOMER_ORDERS, { id, after }, { sleep, log, label: '补查订单' });
    if (!data.customer) return { orders: [], missing: true, capped: false };
    const connection = data.customer.orders;
    for (const node of connection?.nodes ?? []) {
      orders.push(node);
      const order = parseOrder(node);
      if (isValidOrder(order) && isGiftableTotal(order.totalCents, percent)) return { orders, missing: false, capped: false };
    }
    after = nextCursor(connection?.pageInfo, after, '补查订单');
    if (!after) return { orders, missing: false, capped: false };
  }
  return { orders, missing: false, capped: true };
}

// ---------------------------------------------------------------------------
// Snapshot file
// ---------------------------------------------------------------------------

/**
 * Every customer in paths.snapshot, parsed with rules.parseCustomer, read as a
 * stream. Lines end at "\n" only (see textLines); line numbers in errors count
 * those lines.
 */
export async function readSnapshot(paths) {
  const file = paths.snapshot;
  const name = path.basename(file);
  if (!fs.existsSync(file)) throw new Error(`找不到客户快照 ${file}，请运行 select --refresh 重新导出`);
  const input = fs.createReadStream(file, { encoding: 'utf8' });
  const customers = [];
  let lineNo = 0;
  try {
    for await (const raw of textLines(input)) {
      lineNo += 1;
      const line = raw.trim(); // also drops the "\r" of a CRLF line end
      if (!line) continue;
      let node;
      try {
        node = JSON.parse(line);
      } catch (err) {
        throw new Error(`客户快照 ${name} 第 ${lineNo} 行不是有效的 JSON（${err.message}），请运行 select --refresh 重新导出`);
      }
      if (!node || typeof node.id !== 'string') {
        throw new Error(`客户快照 ${name} 第 ${lineNo} 行不是客户记录，请运行 select --refresh 重新导出`);
      }
      customers.push(parseCustomer(node));
    }
  } finally {
    input.destroy();
  }
  return customers;
}
