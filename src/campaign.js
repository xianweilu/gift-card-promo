import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

// ---------------------------------------------------------------------------
// Paths
// ---------------------------------------------------------------------------

/** Every file a campaign uses lives in campaigns/<CAMPAIGN_ID>/ (git-ignored: it holds customer PII). */
export function campaignPaths(config) {
  const dir = path.join(config.campaignsDir, config.campaignId);
  return {
    dir,
    snapshot: path.join(dir, 'customers.jsonl'), // raw export, one customer per line
    snapshotMeta: path.join(dir, 'snapshot.json'), // { exportedAt, source, count, ... }
    activeOrders: path.join(dir, 'active-orders.json'), // orders since the cutoff, taken with the snapshot
    orderHistory: path.join(dir, 'order-history.json'), // follow-up lookups, keyed by customer id
    selection: path.join(dir, 'selection.json'), // frozen list + amounts (written by select)
    journal: path.join(dir, 'journal.jsonl'), // append-only record of every write and run
    tags: path.join(dir, 'tags.json'), // latest SENT_TAG snapshot (export --refresh / verify)
    verify: path.join(dir, 'verify.json'), // latest verify result
    usage: path.join(dir, 'usage.json'), // latest usage report data
    usageDir: path.join(dir, 'usage'), // one snapshot per day: usage/<YYYY-MM-DD>.json
    previewDir: path.join(dir, 'preview'), // rendered email previews
    runLock: path.join(dir, 'run.lock'),
    excelLock: path.join(dir, 'excel.lock'),
    excel: path.join(dir, `gift-card-promo-${config.campaignId}.xlsx`),
  };
}

export function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

// ---------------------------------------------------------------------------
// Atomic files
// ---------------------------------------------------------------------------

/** temp file → fsync → rename over the target, so readers never see half a file. */
export function writeFileAtomic(file, data) {
  ensureDir(path.dirname(file));
  const tmp = `${file}.${process.pid}.${Date.now()}.tmp`;
  const fd = fs.openSync(tmp, 'w');
  try {
    fs.writeFileSync(fd, data);
    fs.fsyncSync(fd);
  } finally {
    fs.closeSync(fd);
  }
  fs.renameSync(tmp, file);
}

export function writeJsonAtomic(file, value) {
  writeFileAtomic(file, `${JSON.stringify(value, null, 2)}\n`);
}

/** Parsed JSON, or `fallback` when the file does not exist. A corrupt file is an error. */
export function readJson(file, fallback = undefined) {
  let text;
  try {
    text = fs.readFileSync(file, 'utf8');
  } catch (err) {
    if (err.code === 'ENOENT') return fallback;
    throw err;
  }
  try {
    return JSON.parse(text);
  } catch (err) {
    throw new Error(`${file} is not valid JSON (${err.message}); refusing to continue`);
  }
}

// ---------------------------------------------------------------------------
// Locks
// ---------------------------------------------------------------------------
//
// A lock is a small JSON file { pid, host, command, startedAt, token } (run.lock, excel.lock).
//
// Held or stale. A lock is held while its pid is alive on this machine. The host name is only
// informational: campaigns/ is on the local disk, and a Mac's host name changes with the network.
// A file whose content is not a lock payload (empty, cut short, damaged: older versions of this
// tool, a crash, a hand edit, or for an instant this tool on a file system without hard links)
// counts as held while its mtime is less than 10 s away and as stale after that. A file that
// cannot be read at all (any error other than "it does not exist") counts as held: nothing is
// ever removed on the strength of a failed read.
//
// Taking a lock. The payload is written to a private temp file <lock>.<token>.tmp, which is then
// hard-linked to the lock name and unlinked. link() fails with EEXIST while the name is taken,
// so a lock file appears atomically together with its content: nobody ever sees an empty lock
// file that is "being created".
//
// Removing a stale lock without ever removing (or even moving) a newer one. Several processes
// can judge the same file stale at once; if one of them removes it and takes the lock, the
// others must not touch the new lock. So a stale file is removed only by the one process that
// creates its claim <lock>.<id>.reap first (exclusive create; <id> is derived from the stale
// file's inode, size, mtime and content, so a claim names exactly one file). The claimant looks
// again: if the lock name still holds that file, it keeps holding it until the claimant moves
// it, because nobody else removes a claimed file and no lock can be created over an existing
// name. The claimant then moves it aside with rename() to <lock>.<token>-<n>.stale, checks that the
// moved file is the stale one and deletes it (anything else is linked back, never deleted),
// removes its claim and tries to take the lock again. Processes that find the claim wait a
// moment and look again. (Rename-and-verify alone is not enough: the slower process would
// move the winner's new lock aside, and a third process could take the free name before it is
// put back.)
//
// A process killed in the microseconds while it holds a claim leaves the claim behind, and that
// one stale lock then stays until it is deleted by hand, which the LockError message explains.

export class LockError extends Error {
  constructor(message, holder, file = null) {
    super(message);
    this.name = 'LockError';
    this.holder = holder;
    this.file = file;
  }
}

export function pidAlive(pid) {
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch (err) {
    return err.code === 'EPERM';
  }
}

/** A lock file whose content is not a lock payload counts as held for this long after its mtime. */
const UNPARSEABLE_HELD_MS = 10_000;
/**
 * Attempts per tryLock (each one: create, or judge and maybe remove a stale lock). While another
 * process holds the claim on a stale lock, the waits between attempts double from 1 ms up to
 * 32 ms (about 0.1 s in all): a claim is normally held for well under a millisecond.
 */
const LOCK_ATTEMPTS = 8;
const MAX_CLAIM_WAIT_MS = 32;
/** link() errors that mean "this file system has no hard links": create the lock exclusively instead. */
const NO_HARD_LINKS = new Set(['ENOTSUP', 'EOPNOTSUPP', 'EPERM', 'ENOSYS', 'EXDEV']);

const pauseCell = new Int32Array(new SharedArrayBuffer(4));
/** Synchronous sleep (lock attempts are synchronous). */
function pause(ms) {
  Atomics.wait(pauseCell, 0, 0, ms);
}

function unlinkQuiet(file) {
  try {
    fs.unlinkSync(file);
  } catch {
    /* already gone */
  }
}

/** The lock payload in `text`, or null when it is not one (it must at least carry a pid). */
function parseLock(text) {
  let value;
  try {
    value = JSON.parse(text);
  } catch {
    return null;
  }
  return value !== null && typeof value === 'object' && Number.isInteger(value.pid) && value.pid > 0 ? value : null;
}

/**
 * One consistent look at a lock file (open, fstat and read through the same descriptor):
 *   { state: 'missing' } | { state: 'error', error } | { state: 'present', st, text, info }
 * `st` holds bigint stats; `info` is the parsed payload or null.
 */
function readLock(file) {
  let fd;
  try {
    fd = fs.openSync(file, 'r');
  } catch (err) {
    return err.code === 'ENOENT' ? { state: 'missing' } : { state: 'error', error: err };
  }
  try {
    const st = fs.fstatSync(fd, { bigint: true });
    const text = fs.readFileSync(fd, 'utf8');
    return { state: 'present', st, text, info: parseLock(text) };
  } catch (err) {
    return { state: 'error', error: err };
  } finally {
    fs.closeSync(fd);
  }
}

/** The same file with the same content as when it was judged (two 'present' readLock() results). */
function sameFile(a, b) {
  return a.st.dev === b.st.dev && a.st.ino === b.st.ino && a.st.size === b.st.size
    && a.st.mtimeNs === b.st.mtimeNs && a.text === b.text;
}

/** A short name for one particular lock file (inode, size, mtime, content). */
function fileId(seen) {
  const key = `${seen.st.dev}:${seen.st.ino}:${seen.st.size}:${seen.st.mtimeNs}:${seen.text}`;
  return crypto.createHash('sha256').update(key).digest('hex').slice(0, 16);
}

/** Judge a readLock() result: { state: 'missing' } | { state: 'held', holder } | { state: 'stale', seen }. */
function judge(seen, nowMs = Date.now()) {
  if (seen.state === 'missing') return { state: 'missing' };
  if (seen.state === 'error') return { state: 'held', holder: { unreadable: true } };
  if (seen.info) return pidAlive(seen.info.pid) ? { state: 'held', holder: seen.info } : { state: 'stale', seen };
  // Not a payload: held while fresh. A time far in the future (clock change) does not keep it held.
  const ageMs = nowMs - Number(seen.st.mtimeMs);
  return Math.abs(ageMs) < UNPARSEABLE_HELD_MS ? { state: 'held', holder: { unreadable: true } } : { state: 'stale', seen };
}

/** The live holder of a lock file (its payload), or null when it is missing or stale. */
export function lockHolder(file) {
  const verdict = judge(readLock(file));
  return verdict.state === 'held' ? verdict.holder : null;
}

/** Create `file` exclusively with `text`; returns its { dev, ino }. */
function writeNewFile(file, text) {
  const fd = fs.openSync(file, 'wx');
  try {
    fs.writeFileSync(fd, text);
    const st = fs.fstatSync(fd, { bigint: true });
    return { dev: st.dev, ino: st.ino };
  } finally {
    fs.closeSync(fd);
  }
}

/**
 * Make the prepared temp file the lock. Returns the lock's { dev, ino }, or null when the name is
 * taken (by a live lock or a stale one).
 */
function createLock(tmp, file, text, tmpId) {
  try {
    fs.linkSync(tmp, file);
    return tmpId;
  } catch (err) {
    if (err.code === 'EEXIST') return null;
    if (!NO_HARD_LINKS.has(err.code)) throw err;
  }
  // No hard links on this file system: create the name exclusively, then write it. It is empty
  // for that moment, which other processes treat as held (see judge()).
  let fd;
  try {
    fd = fs.openSync(file, 'wx');
  } catch (err) {
    if (err.code === 'EEXIST') return null;
    throw err;
  }
  let created = null;
  try {
    fs.writeFileSync(fd, text);
    const st = fs.fstatSync(fd, { bigint: true });
    created = { dev: st.dev, ino: st.ino };
    return created;
  } finally {
    fs.closeSync(fd);
    if (!created) unlinkQuiet(file); // never leave our own half-written lock behind
  }
}

/**
 * Remove the stale lock file `stale` (a 'present' readLock() result judged stale), following the
 * claim protocol described above. Returns 'removed', 'busy' (another process is removing it right
 * now) or 'changed' (the lock name no longer holds that file).
 */
function removeStale(file, stale, tag) {
  const claim = `${file}.${fileId(stale)}.reap`;
  let fd;
  try {
    fd = fs.openSync(claim, 'wx');
  } catch (err) {
    if (err.code === 'EEXIST') return 'busy';
    throw err;
  }
  try {
    try {
      fs.writeSync(fd, JSON.stringify({ pid: process.pid, host: os.hostname(), at: new Date().toISOString() }));
    } catch {
      /* the claim's content is only informational */
    } finally {
      fs.closeSync(fd);
    }
    // Only the holder of this claim removes this file: if the name holds it now, it holds it until we move it.
    const current = readLock(file);
    if (current.state !== 'present' || !sameFile(current, stale)) return 'changed';
    const aside = `${file}.${tag}.stale`;
    try {
      fs.renameSync(file, aside);
    } catch (err) {
      if (err.code === 'ENOENT') return 'changed';
      throw err;
    }
    const moved = readLock(aside);
    if (moved.state === 'present' && sameFile(moved, stale)) {
      unlinkQuiet(aside);
      return 'removed';
    }
    // Not the stale file after all (only possible when something outside this protocol replaced
    // it): put it back; never delete a file that might be someone's lock.
    try {
      fs.linkSync(aside, file);
      unlinkQuiet(aside);
    } catch {
      /* the name was taken meanwhile: leave the moved file where it is */
    }
    return 'changed';
  } finally {
    unlinkQuiet(claim);
  }
}

/** A release function that removes the lock only while the name still holds our own file and token. */
function releaser(file, token, mine) {
  let released = false;
  return function release() {
    if (released) return;
    released = true;
    const seen = readLock(file);
    if (seen.state === 'present' && seen.st.dev === mine.dev && seen.st.ino === mine.ino && seen.info?.token === token) {
      unlinkQuiet(file);
    }
  };
}

/** Try to take a lock (without waiting for a live holder). Returns { release } or { holder }. */
function tryLock(file, payload) {
  ensureDir(path.dirname(file));
  const text = JSON.stringify(payload);
  const tmp = `${file}.${payload.token}.tmp`;
  try {
    const tmpId = writeNewFile(tmp, text);
    let last = {};
    let waitMs = 1;
    for (let attempt = 0; attempt < LOCK_ATTEMPTS; attempt += 1) {
      const mine = createLock(tmp, file, text, tmpId);
      if (mine) return { release: releaser(file, payload.token, mine) };
      const verdict = judge(readLock(file));
      if (verdict.state === 'missing') continue; // released just now: try again
      if (verdict.state === 'held') return { holder: verdict.holder };
      last = verdict.seen.info ?? {};
      if (removeStale(file, verdict.seen, `${payload.token}-${attempt}`) === 'busy') {
        pause(waitMs);
        waitMs = Math.min(waitMs * 2, MAX_CLAIM_WAIT_MS);
      }
    }
    // Still not settled after every attempt (heavy contention, or a stale lock whose removal was
    // interrupted for good): report it as held, naming the live holder when there is one by now.
    const final = judge(readLock(file));
    return { holder: final.state === 'held' ? final.holder : last };
  } finally {
    unlinkQuiet(tmp);
  }
}

function lockPayload(command) {
  return {
    pid: process.pid,
    host: os.hostname(), // informational only
    command,
    startedAt: new Date().toISOString(),
    token: `${process.pid}-${Date.now()}-${crypto.randomBytes(6).toString('hex')}`,
  };
}

/** Where the lock is and how to recover by hand (user-facing). */
function lockFileHint(file) {
  return `锁文件：${file}。如果确定没有命令在运行（比如电脑重启过或进程已被强制关闭），删除这个文件后重试。`;
}

/**
 * select / issue / remind / verify / usage are mutually exclusive. Throws LockError naming the
 * running command; a lock left behind by a dead process is removed silently.
 */
export function acquireRunLock(paths, command) {
  const result = tryLock(paths.runLock, lockPayload(command));
  if (result.holder) {
    const h = result.holder;
    throw new LockError(`${h.command ?? '另一条命令'} 正在运行（PID ${h.pid ?? '?'}，开始于 ${h.startedAt ?? '?'}），请等它结束再运行。${lockFileHint(paths.runLock)}`, h, paths.runLock);
  }
  return result.release;
}

/** The command currently holding the run lock (e.g. 'issue'), or null. */
export function runningCommand(paths) {
  return lockHolder(paths.runLock)?.command ?? null;
}

/**
 * Serialise writers of the Excel file. Waits (polling) for a live holder;
 * removes a stale lock. `fn` runs while the lock is held.
 */
export async function withExcelLock(paths, fn, { timeoutMs = 180_000, pollMs = 200, sleep } = {}) {
  const wait = sleep ?? ((ms) => new Promise((r) => setTimeout(r, ms)));
  const deadline = Date.now() + timeoutMs;
  for (;;) {
    const result = tryLock(paths.excelLock, lockPayload('excel'));
    if (result.release) {
      try {
        return await fn();
      } finally {
        result.release();
      }
    }
    if (Date.now() >= deadline) {
      const h = result.holder ?? {};
      throw new LockError(`等待写 Excel 超时：${h.command ?? '另一个进程'}（PID ${h.pid ?? '?'}）一直占着 Excel 锁。${lockFileHint(paths.excelLock)}`, h, paths.excelLock);
    }
    await wait(pollMs);
  }
}

// ---------------------------------------------------------------------------
// Journal
// ---------------------------------------------------------------------------

/**
 * Append one entry and fsync before returning, so a crash right after a
 * Shopify call can never lose the record of having made it.
 */
export function appendJournal(file, entry, { now = () => new Date().toISOString() } = {}) {
  ensureDir(path.dirname(file));
  const line = `${JSON.stringify({ t: now(), ...entry })}\n`;
  const fd = fs.openSync(file, 'a+');
  try {
    repairTail(fd);
    fs.writeSync(fd, line);
    fs.fsyncSync(fd);
  } finally {
    fs.closeSync(fd);
  }
}

/**
 * Make the journal end with a newline before the next append. A crash in the
 * middle of an earlier append can leave a last line without its "\n": a line
 * that is complete JSON only gets the missing newline; a torn one (which
 * readJournal already ignores) is cut off. Otherwise the new entry would be
 * glued onto it, lost, and every later read would fail on the merged line.
 */
function repairTail(fd) {
  const size = fs.fstatSync(fd).size;
  if (!size) return;
  const lastByte = Buffer.alloc(1);
  fs.readSync(fd, lastByte, 0, 1, size - 1);
  if (lastByte[0] === 0x0a) return;
  // Start of the last line: just after the previous newline (0 when there is none).
  const CHUNK = 64 * 1024;
  let start = 0;
  for (let end = size; end > 0; end -= CHUNK) {
    const from = Math.max(0, end - CHUNK);
    const buf = Buffer.alloc(end - from);
    fs.readSync(fd, buf, 0, buf.length, from);
    const i = buf.lastIndexOf(0x0a);
    if (i !== -1) {
      start = from + i + 1;
      break;
    }
  }
  const tail = Buffer.alloc(size - start);
  fs.readSync(fd, tail, 0, tail.length, start);
  let complete = false;
  try {
    JSON.parse(tail.toString('utf8'));
    complete = true;
  } catch {
    complete = false;
  }
  if (complete) fs.writeSync(fd, '\n');
  else fs.ftruncateSync(fd, start);
}

/**
 * All journal entries. A torn last line (crash mid-write) is ignored; a bad
 * line anywhere else means the file was edited or damaged, which is fatal.
 */
export function readJournal(file) {
  let text;
  try {
    text = fs.readFileSync(file, 'utf8');
  } catch (err) {
    if (err.code === 'ENOENT') return [];
    throw err;
  }
  const lines = text.split('\n');
  const entries = [];
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i].trim();
    if (!line) continue;
    try {
      entries.push(JSON.parse(line));
    } catch (err) {
      const isLast = lines.slice(i + 1).every((l) => !l.trim());
      if (isLast) break;
      throw new Error(`${file} line ${i + 1} is not valid JSON (${err.message}); the journal must not be edited`);
    }
  }
  return entries;
}

/**
 * Customer status values (Excel labels in src/report/excel.js):
 *   pending      待发放          nothing attempted yet, or a reconcile proved no card exists
 *   in_progress  进行中          create.start written, no outcome yet (live run, or a crash)
 *   unknown      需人工核对      Shopify's answer was lost and the card was not found yet
 *   created      已建卡未打tag    the card exists, the tag is still missing
 *   done         已完成          card created and customer tagged
 *   failed       失败            Shopify rejected the create (userErrors); retried only with --retry-failed
 *   skipped      发放前跳过       pre-flight found the customer no longer qualifies
 */
/**
 * Reminder status per customer and round ('1' | '2'):
 *   in_progress  进行中       remind.start written, no outcome yet
 *   sent         已发          Shopify accepted the re-send (remind.ok), or the customer carries the
 *                              round tag in Shopify while the journal had no send (remind.found:
 *                              source 'tag', at null — the journal was lost or rolled back)
 *   failed       失败          userErrors (not retried automatically)
 *   unknown      结果不明      the answer was lost: NEVER re-sent automatically (--retry-unknown)
 *   skipped      跳过          not eligible this round (used card, unsubscribed, ...)
 *   (absent)     not attempted yet; remind.rejected also returns a row to this state
 * A sent row may carry `tagError`: adding the round tag after the send failed (remind.tag.fail);
 * remind.tag.ok (the tag repaired, or found in Shopify after all) clears it.
 */
export const REMIND_STATUS = Object.freeze({
  IN_PROGRESS: 'in_progress',
  SENT: 'sent',
  FAILED: 'failed',
  UNKNOWN: 'unknown',
  SKIPPED: 'skipped',
});

export const STATUS = Object.freeze({
  PENDING: 'pending',
  IN_PROGRESS: 'in_progress',
  UNKNOWN: 'unknown',
  CREATED: 'created',
  DONE: 'done',
  FAILED: 'failed',
  SKIPPED: 'skipped',
});

function blankState() {
  return {
    status: STATUS.PENDING,
    attempts: 0,
    batch: null,
    run: null,
    startedAt: null,
    giftCardId: null,
    last4: null,
    amountCents: null,
    createdAt: null,
    taggedAt: null,
    error: null,
    skipReason: null,
    reconciledFrom: null,
    reminders: {}, // { '1': { status, startedAt, at, giftCardId, error, reason, detail, source?, tagError? }, '2': {...} }
  };
}

/**
 * Journal ops about issuing a card to a customer (written by issue, and by verify's reconcile).
 * Any one of them means issuing has started, so the list is frozen.
 */
const ISSUE_OPS = new Set([
  'create.start', 'create.ok', 'create.fail', 'create.rejected', 'create.unknown',
  'reconcile.found', 'reconcile.none', 'tag.ok', 'tag.fail', 'skip',
]);

/**
 * Replay the journal. Returns
 *   { customers: Map<customerGid, state>, runs: [run], lastBatch, issuingStarted }
 * where `issuingStarted` (select refuses to rebuild the list once it is set) is true as soon as
 * a live (not dry-run) issue run has started, even one that stopped before its first card, or the
 * journal holds any issue op (ISSUE_OPS: a card attempted, found, tagged, or a customer skipped).
 */
export function foldJournal(entries) {
  const customers = new Map();
  const runsById = new Map();
  let lastBatch = 0; // highest batch number any live run took (the next run takes lastBatch + 1)
  let lastIssuedBatch = 0; // highest batch that attempted a card (repair-only / date-blocked runs create none)
  let issuingStarted = false;
  const get = (cid) => {
    if (!customers.has(cid)) customers.set(cid, blankState());
    return customers.get(cid);
  };

  for (const e of entries) {
    if (ISSUE_OPS.has(e.op)) issuingStarted = true;
    switch (e.op) {
      case 'run.start':
        runsById.set(e.run, { run: e.run, command: e.command, dryRun: !!e.dryRun, batch: e.batch ?? null, limit: e.limit ?? null, startedAt: e.t, endedAt: null, summary: null, options: e.options ?? null });
        if (!e.dryRun && Number.isInteger(e.batch)) lastBatch = Math.max(lastBatch, e.batch);
        // An `issue --repair-only` run never creates a card: its start alone does not freeze the
        // list (anything it records — reconcile.found, tag.ok, skip — still does, via ISSUE_OPS).
        if (e.command === 'issue' && !e.dryRun && !e.options?.repairOnly) issuingStarted = true;
        break;
      case 'run.end': {
        const r = runsById.get(e.run);
        if (r) {
          r.endedAt = e.t;
          r.summary = e.summary ?? null;
          r.exitCode = e.exitCode ?? null;
        }
        break;
      }
      case 'create.start': {
        const s = get(e.cid);
        s.status = STATUS.IN_PROGRESS;
        s.attempts += 1;
        s.batch = e.batch ?? s.batch;
        if (Number.isInteger(e.batch)) lastIssuedBatch = Math.max(lastIssuedBatch, e.batch);
        s.run = e.run ?? s.run;
        s.startedAt = e.t;
        s.amountCents = e.amountCents ?? s.amountCents;
        s.error = null;
        break;
      }
      case 'create.ok':
      case 'reconcile.found': {
        const s = get(e.cid);
        s.status = s.taggedAt ? STATUS.DONE : STATUS.CREATED;
        s.giftCardId = e.giftCardId ?? s.giftCardId;
        s.last4 = e.last4 ?? s.last4;
        s.createdAt = e.createdAt ?? e.t;
        if (e.amountCents) s.amountCents = e.amountCents;
        if (e.batch) s.batch = e.batch;
        s.error = null;
        if (e.op === 'reconcile.found') s.reconciledFrom = e.source ?? 'reconcile';
        break;
      }
      case 'create.fail': {
        const s = get(e.cid);
        s.status = STATUS.FAILED;
        s.error = e.error ?? 'rejected';
        break;
      }
      case 'create.rejected': {
        // Shopify definitely did not apply it (e.g. throttled after all retries): safe to try again.
        const s = get(e.cid);
        s.status = STATUS.PENDING;
        s.error = e.error ?? null;
        break;
      }
      case 'create.unknown': {
        const s = get(e.cid);
        s.status = STATUS.UNKNOWN;
        s.error = e.error ?? 'outcome unknown';
        break;
      }
      case 'reconcile.none': {
        const s = get(e.cid);
        s.status = STATUS.PENDING;
        s.error = e.note ?? null;
        break;
      }
      case 'tag.ok': {
        const s = get(e.cid);
        s.taggedAt = e.t;
        if (s.giftCardId) s.status = STATUS.DONE;
        s.error = null;
        break;
      }
      case 'tag.fail': {
        const s = get(e.cid);
        s.error = e.error ?? 'tag failed';
        break;
      }
      case 'remind.start': {
        const s = get(e.cid);
        s.reminders[String(e.round)] = { status: REMIND_STATUS.IN_PROGRESS, startedAt: e.t, at: null, giftCardId: e.giftCardId ?? null, error: null, reason: null, run: e.run ?? null };
        break;
      }
      case 'remind.ok':
      case 'remind.fail':
      case 'remind.unknown': {
        const s = get(e.cid);
        const r = s.reminders[String(e.round)] ?? { startedAt: null, giftCardId: null };
        r.status = e.op === 'remind.ok' ? REMIND_STATUS.SENT : e.op === 'remind.fail' ? REMIND_STATUS.FAILED : REMIND_STATUS.UNKNOWN;
        r.at = e.t;
        r.error = e.op === 'remind.ok' ? null : e.error ?? null;
        s.reminders[String(e.round)] = r;
        break;
      }
      case 'remind.rejected': {
        // Shopify definitely did not send it (e.g. throttled out): the round may be tried again.
        const s = get(e.cid);
        delete s.reminders[String(e.round)];
        break;
      }
      case 'remind.found': {
        // The customer carries the round tag in Shopify, so this round was sent, though the journal
        // had no remind.ok (lost or rolled back): sent from now on. When and with which card is unknown.
        const s = get(e.cid);
        if (s.reminders[String(e.round)]?.status !== REMIND_STATUS.SENT) {
          s.reminders[String(e.round)] = { status: REMIND_STATUS.SENT, at: null, startedAt: null, giftCardId: null, error: null, reason: null, source: 'tag', run: e.run ?? null };
        }
        break;
      }
      case 'remind.tag.fail':
      case 'remind.tag.ok': {
        // Adding the round tag after a send failed (the status stays: the email went out), or a later
        // run added it / found it. Only for a round state that exists; nothing is created.
        const r = customers.get(e.cid)?.reminders?.[String(e.round)];
        if (r) {
          r.tagError = e.op === 'remind.tag.fail' ? e.error ?? 'tag failed' : null;
          // A deleted customer can never be tagged: stop the repair for good.
          if (e.op === 'remind.tag.ok' && e.note === 'customer deleted') r.tagSettled = true;
        }
        break;
      }
      case 'remind.skip': {
        const s = get(e.cid);
        const prev = s.reminders[String(e.round)];
        if (!prev || prev.status === REMIND_STATUS.SKIPPED) {
          s.reminders[String(e.round)] = { status: REMIND_STATUS.SKIPPED, startedAt: null, at: e.t, giftCardId: null, error: null, reason: e.reason ?? null, detail: e.detail ?? null };
        }
        break;
      }
      case 'skip': {
        const s = get(e.cid);
        if (s.status === STATUS.PENDING || s.status === STATUS.FAILED) {
          s.status = STATUS.SKIPPED;
          s.skipReason = e.reason ?? null;
          s.batch = e.batch ?? s.batch;
        }
        break;
      }
      default:
        // unknown ops are ignored so that newer journals stay readable
        break;
    }
  }
  return { customers, runs: [...runsById.values()], lastBatch, lastIssuedBatch, issuingStarted };
}

/** A short unique id for a run, sortable by time. */
export function newRunId(now = new Date()) {
  return `${now.toISOString().replace(/[-:.TZ]/g, '').slice(0, 14)}-${process.pid}`;
}
