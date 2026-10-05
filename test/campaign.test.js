import { test, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import { once } from 'node:events';
import { fork } from 'node:child_process';
import { Worker } from 'node:worker_threads';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  campaignPaths, writeJsonAtomic, readJson, writeFileAtomic, acquireRunLock, runningCommand, withExcelLock, LockError,
  lockHolder, appendJournal, readJournal, foldJournal, newRunId, STATUS, USED_TAG_STATUS,
} from '../src/campaign.js';

const CAMPAIGN_URL = new URL('../src/campaign.js', import.meta.url).href;
const DEAD_PID = 99_999_999; // above every platform's pid limit, so never alive
const HINT = '如果确定没有命令在运行（比如电脑重启过或进程已被强制关闭），删除这个文件后重试。';
const dirs = [];
function tmpDir(prefix = 'gcp-campaign-') {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), prefix));
  dirs.push(dir);
  return dir;
}
function tmpPaths(campaignId = '2026-10') {
  const paths = campaignPaths({ campaignsDir: tmpDir(), campaignId });
  fs.mkdirSync(paths.dir, { recursive: true });
  return paths;
}
afterEach(() => {
  while (dirs.length) fs.rmSync(dirs.pop(), { recursive: true, force: true });
});

/** What a crashed (or double-Ctrl+C'd) command leaves behind: a lock whose pid is gone. */
function writeStaleLock(file, extra = {}) {
  fs.writeFileSync(file, JSON.stringify({ pid: DEAD_PID, host: os.hostname(), command: 'issue', startedAt: '2026-10-05T16:00:00.000Z', token: 'stale', ...extra }));
}
function setAge(file, ms) {
  const t = new Date(Date.now() - ms);
  fs.utimesSync(file, t, t);
}
/** Files the lock protocol uses next to a lock: temp payloads, moved-aside locks and claims. */
function protocolFiles(dir) {
  return fs.readdirSync(dir).filter((f) => /\.(tmp|stale|reap)$/.test(f));
}

test('campaign: every file lives in campaigns/<id>/', () => {
  const p = campaignPaths({ campaignsDir: '/x/campaigns', campaignId: '2026-10' });
  assert.equal(p.dir, '/x/campaigns/2026-10');
  assert.equal(p.excel, '/x/campaigns/2026-10/gift-card-promo-2026-10.xlsx');
  assert.equal(p.journal, '/x/campaigns/2026-10/journal.jsonl');
  assert.equal(p.selection, '/x/campaigns/2026-10/selection.json');
  assert.equal(p.usageDir, '/x/campaigns/2026-10/usage');
  assert.equal(p.previewDir, '/x/campaigns/2026-10/preview');
  for (const v of Object.values(p)) assert.ok(v.startsWith('/x/campaigns/2026-10'), v);
});

test('campaign: atomic JSON files; missing → fallback; corrupt → refuse', () => {
  const p = tmpPaths();
  assert.equal(readJson(p.selection, null), null);
  assert.deepEqual(readJson(p.selection, { a: 1 }), { a: 1 });
  writeJsonAtomic(p.selection, { recipients: [1, 2] });
  assert.deepEqual(readJson(p.selection), { recipients: [1, 2] });
  assert.ok(fs.readFileSync(p.selection, 'utf8').endsWith('\n'));
  assert.deepEqual(fs.readdirSync(p.dir).filter((f) => f.endsWith('.tmp')), [], 'no temp file left behind');
  fs.writeFileSync(p.selection, '{ broken');
  assert.throws(() => readJson(p.selection), /not valid JSON/);
  writeFileAtomic(path.join(p.dir, 'deep', 'x.txt'), 'hi');
  assert.equal(fs.readFileSync(path.join(p.dir, 'deep', 'x.txt'), 'utf8'), 'hi');
});

// ---------------------------------------------------------------------------
// Locks: normal behaviour
// ---------------------------------------------------------------------------

test('campaign: only one write command at a time; the second one is told who is running and where the lock file is', () => {
  const p = tmpPaths();
  const release = acquireRunLock(p, 'issue');
  assert.equal(runningCommand(p), 'issue');
  assert.deepEqual(protocolFiles(p.dir), [], 'the temp payload file is gone once the lock is taken');
  const message = /^issue 正在运行（PID \d+，开始于 .+），请等它结束再运行。锁文件：(.+)。如果确定没有命令在运行（比如电脑重启过或进程已被强制关闭），删除这个文件后重试。$/;
  assert.throws(() => acquireRunLock(p, 'select'), (err) => err instanceof LockError && message.exec(err.message)?.[1] === p.runLock
    && err.holder.command === 'issue' && err.holder.pid === process.pid && err.file === p.runLock);
  release();
  assert.equal(runningCommand(p), null);
  assert.equal(fs.existsSync(p.runLock), false);
  const again = acquireRunLock(p, 'select');
  assert.equal(runningCommand(p), 'select');
  again();
  assert.deepEqual(fs.readdirSync(p.dir), []);
});

test('campaign: a lock left by a dead process is cleaned up automatically, whatever host name it recorded', () => {
  const p = tmpPaths();
  writeStaleLock(p.runLock);
  assert.equal(lockHolder(p.runLock), null);
  assert.equal(runningCommand(p), null);
  const release = acquireRunLock(p, 'verify');
  assert.equal(runningCommand(p), 'verify');
  release();

  // The host name is informational only (a Mac's host name changes with the network): a dead pid is stale ...
  writeStaleLock(p.runLock, { host: 'some-other-mac', command: 'export' });
  assert.equal(runningCommand(p), null);
  acquireRunLock(p, 'issue')();
  writeStaleLock(p.runLock, { host: undefined, command: 'export' });
  acquireRunLock(p, 'issue')();
  // ... and a live pid is held, whatever host it names.
  fs.writeFileSync(p.runLock, JSON.stringify({ pid: process.pid, host: 'some-other-mac', command: 'export', startedAt: 'x', token: 't' }));
  assert.equal(runningCommand(p), 'export');
  assert.throws(() => acquireRunLock(p, 'issue'), (err) => err instanceof LockError && err.message.startsWith('export 正在运行'));
  fs.unlinkSync(p.runLock);

  // The Excel lock follows the same rules.
  writeStaleLock(p.excelLock, { host: 'old-name.local', command: 'excel' });
  assert.equal(lockHolder(p.excelLock), null);
  assert.deepEqual(protocolFiles(p.dir), []);
  assert.deepEqual(fs.readdirSync(p.dir).sort(), ['excel.lock']);
});

test('campaign: a lock file that is not a payload counts as held for 10 s after its mtime, stale after that', () => {
  const p = tmpPaths();
  for (const junk of ['', 'not json', '{"pid":', '{"command":"issue"}', '{"pid":"12"}', 'null']) {
    fs.writeFileSync(p.runLock, junk);
    // Fresh: another process may be between creating and writing it (older versions of this tool).
    assert.deepEqual(lockHolder(p.runLock), { unreadable: true }, JSON.stringify(junk));
    assert.equal(runningCommand(p), null);
    assert.throws(() => acquireRunLock(p, 'issue'), (err) => err instanceof LockError && err.message.startsWith('另一条命令 正在运行（PID ?，开始于 ?）') && err.message.endsWith(HINT));
    assert.equal(fs.readFileSync(p.runLock, 'utf8'), junk, 'a fresh unparseable lock is never removed');
    setAge(p.runLock, 5_000);
    assert.throws(() => acquireRunLock(p, 'issue'), LockError, 'still held at 5 s');
    setAge(p.runLock, 15_000);
    assert.equal(lockHolder(p.runLock), null, 'stale at 15 s');
    const release = acquireRunLock(p, 'issue');
    assert.equal(runningCommand(p), 'issue');
    release();
  }
  // A time far in the future (clock change, copied file) does not keep it held forever.
  fs.writeFileSync(p.runLock, '');
  setAge(p.runLock, -3_600_000);
  acquireRunLock(p, 'issue')();
  assert.deepEqual(fs.readdirSync(p.dir), []);
});

test('campaign: a lock file that cannot be read counts as held and is never removed, however old', { skip: process.getuid?.() === 0 ? 'root can read anything' : false }, () => {
  const p = tmpPaths();
  writeStaleLock(p.runLock);
  setAge(p.runLock, 3_600_000);
  fs.chmodSync(p.runLock, 0o000);
  try {
    assert.ok(lockHolder(p.runLock), 'held');
    assert.equal(runningCommand(p), null);
    assert.throws(() => acquireRunLock(p, 'issue'), (err) => err instanceof LockError && err.message.endsWith(HINT));
    assert.equal(fs.existsSync(p.runLock), true);
    assert.deepEqual(protocolFiles(p.dir), []);
  } finally {
    fs.chmodSync(p.runLock, 0o644);
  }
});

test('campaign: release only removes its own lock', () => {
  const p = tmpPaths();
  // Same file, rewritten by someone else.
  let release = acquireRunLock(p, 'issue');
  fs.writeFileSync(p.runLock, JSON.stringify({ pid: process.pid, host: os.hostname(), command: 'select', token: 'someone-else' }));
  release();
  assert.equal(runningCommand(p), 'select', 'a lock taken over by someone else is left alone');
  fs.unlinkSync(p.runLock);

  // Deleted by hand while the command was running, then taken by another command.
  release = acquireRunLock(p, 'issue');
  fs.unlinkSync(p.runLock);
  const other = acquireRunLock(p, 'verify');
  release();
  assert.equal(runningCommand(p), 'verify', 'the first command does not remove the second one\'s lock');
  release();
  assert.equal(runningCommand(p), 'verify', 'calling release twice is harmless');
  other();
  assert.equal(fs.existsSync(p.runLock), false);

  // A different file that carries the same token (a copy) is not ours either.
  release = acquireRunLock(p, 'issue');
  const copy = fs.readFileSync(p.runLock);
  fs.unlinkSync(p.runLock);
  fs.writeFileSync(p.runLock, copy);
  release();
  assert.equal(fs.existsSync(p.runLock), true, 'only the very file this process created is removed');
  fs.unlinkSync(p.runLock);

  // Released, then a later lock by someone else: a second release() does not touch it.
  release = acquireRunLock(p, 'issue');
  release();
  const later = acquireRunLock(p, 'usage');
  release();
  assert.equal(runningCommand(p), 'usage');
  later();
  assert.deepEqual(fs.readdirSync(p.dir), []);
});

test('campaign: Excel writers are serialised; stale Excel locks are cleaned; waiting times out', async () => {
  const p = tmpPaths();
  const events = [];
  const writer = (name, ms) => withExcelLock(p, async () => {
    events.push(`${name}:start`);
    await new Promise((r) => setTimeout(r, ms));
    events.push(`${name}:end`);
    return name;
  }, { pollMs: 5 });
  const results = await Promise.all([writer('a', 30), writer('b', 5)]);
  assert.deepEqual(results, ['a', 'b']);
  assert.deepEqual(events, ['a:start', 'a:end', 'b:start', 'b:end']);
  assert.equal(fs.existsSync(p.excelLock), false);

  writeStaleLock(p.excelLock, { command: 'excel' });
  assert.equal(await withExcelLock(p, async () => 'ok', { pollMs: 5 }), 'ok');
  writeStaleLock(p.excelLock, { command: 'excel', host: 'old-name.local' });
  assert.equal(await withExcelLock(p, async () => 'ok', { pollMs: 5 }), 'ok', 'a host name change does not keep a dead lock');

  fs.writeFileSync(p.excelLock, JSON.stringify({ pid: process.pid, host: os.hostname(), command: 'excel', token: 'held' }));
  await assert.rejects(withExcelLock(p, async () => 'never', { timeoutMs: 30, pollMs: 5 }), (err) => err instanceof LockError
    && err.message.startsWith(`等待写 Excel 超时：excel（PID ${process.pid}）一直占着 Excel 锁。锁文件：${p.excelLock}。`)
    && err.message.endsWith(HINT) && err.file === p.excelLock);
  // The lock is released even when the work throws.
  fs.unlinkSync(p.excelLock);
  await assert.rejects(withExcelLock(p, async () => { throw new Error('boom'); }), /boom/);
  assert.equal(fs.existsSync(p.excelLock), false);
  assert.deepEqual(fs.readdirSync(p.dir), []);
});

// ---------------------------------------------------------------------------
// Locks: the two races (remind#3, contracts#8, usage-verify#2, issue-duplicates#3)
// ---------------------------------------------------------------------------

/** One process trying to take the run lock. Never throws: unexpected errors are collected. */
function attempt(paths, command, errors) {
  try {
    return { command, release: acquireRunLock(paths, command) };
  } catch (err) {
    if (err instanceof LockError) return { command, refused: err };
    errors.push(err);
    return { command, error: err };
  }
}

/** Lock attempts wait briefly while another process is removing a stale lock; skip those waits in-process. */
function withoutPauses(fn) {
  const realWait = Atomics.wait;
  Atomics.wait = () => 'timed-out';
  try {
    return fn();
  } finally {
    Atomics.wait = realWait;
  }
}

const FS_CALLS = ['openSync', 'closeSync', 'readFileSync', 'writeFileSync', 'writeSync', 'fstatSync', 'statSync', 'linkSync', 'renameSync', 'unlinkSync', 'mkdirSync'];

/**
 * Deterministic interleaving: runs `outer` (one process) and, just before (or just after) its fs
 * call number `at`, runs `inner` (another process acting at that instant, to completion, while the
 * first one is not scheduled). `inner` may call interleave() itself to bring in a third process.
 * Returns how many fs calls `outer` made.
 */
function interleave(outer, inner, moment = {}) {
  return interleaveAll(outer, [{ ...moment, run: inner }]);
}

/** Like interleave(), with several processes [{ at, after, run }], each acting once at its own moment. */
function interleaveAll(outer, injections) {
  const real = {};
  let count = 0;
  let nested = 0;
  const fire = (inj) => {
    nested += 1;
    try {
      inj.run();
    } finally {
      nested -= 1;
    }
  };
  const due = (n, after) => injections.filter((inj) => (inj.at ?? Infinity) === n && !!inj.after === after);
  for (const name of FS_CALLS) {
    real[name] = fs[name];
    fs[name] = function hooked(...args) {
      if (nested) return real[name].apply(this, args);
      const n = count;
      count += 1;
      for (const inj of due(n, false)) fire(inj);
      try {
        return real[name].apply(this, args);
      } finally {
        for (const inj of due(n, true)) fire(inj);
      }
    };
  }
  try {
    outer();
  } finally {
    for (const name of FS_CALLS) fs[name] = real[name];
  }
  return count;
}

/**
 * Patches fs[name] during `run()`: `before(...args)` runs before each real call and
 * `after(result, ...args)` after it. Calls made inside the hooks go straight to fs.
 */
function hookFs(name, { before = () => {}, after = () => {} }, run) {
  const real = fs[name];
  function hooked(...args) {
    fs[name] = real;
    try {
      before(...args);
    } finally {
      fs[name] = hooked;
    }
    const result = real.apply(this, args);
    fs[name] = real;
    try {
      after(result, ...args);
    } finally {
      fs[name] = hooked;
    }
    return result;
  }
  fs[name] = hooked;
  try {
    return run();
  } finally {
    fs[name] = real;
  }
}

const STARTS = {
  'no lock file': { setup: () => {}, someoneGetsIt: true },
  'a stale lock (dead pid)': { setup: (p) => writeStaleLock(p.runLock), someoneGetsIt: true },
  'an old empty lock file': { setup: (p) => { fs.writeFileSync(p.runLock, ''); setAge(p.runLock, 3_600_000); }, someoneGetsIt: true },
  'a fresh empty lock file (being created by an older version)': { setup: (p) => fs.writeFileSync(p.runLock, ''), someoneGetsIt: false },
  'a live lock': { setup: (p) => fs.writeFileSync(p.runLock, JSON.stringify({ pid: process.pid, command: 'usage', token: 'live' })), someoneGetsIt: false },
};

function resetDir(p, setup) {
  for (const f of fs.readdirSync(p.dir)) fs.rmSync(path.join(p.dir, f), { recursive: true, force: true });
  setup(p);
  return fs.existsSync(p.runLock) ? fs.readFileSync(p.runLock) : null;
}

/** The invariants after the contenders of one scenario are done; releases the holder. */
function checkOutcome(p, parties, errors, { someoneGetsIt, before }, label) {
  assert.deepEqual(errors, [], `${label}: unexpected errors`);
  const holders = parties.filter((x) => x?.release);
  if (someoneGetsIt) {
    assert.equal(holders.length, 1, `${label}: exactly one holder (got ${holders.map((h) => h.command).join(', ') || 'none'})`);
    assert.equal(lockHolder(p.runLock)?.command, holders[0].command, `${label}: the lock file names the holder`);
  } else {
    assert.equal(holders.length, 0, `${label}: nobody may take a lock that is held`);
    assert.deepEqual(fs.readFileSync(p.runLock), before, `${label}: the held lock is untouched`);
  }
  for (const x of parties) if (x?.refused) assert.ok(x.refused.message.endsWith(HINT), label);
  assert.deepEqual(protocolFiles(p.dir), [], `${label}: no temp, moved-aside or claim file is left`);
  for (const h of holders) h.release();
  if (someoneGetsIt) assert.deepEqual(fs.readdirSync(p.dir), [], `${label}: after release nothing is left`);
}

test('locks: a lock file is never visible without its content (it is created by linking a finished temp file)', () => {
  const p = tmpPaths();
  writeStaleLock(p.runLock);
  const seen = [];
  const realLink = fs.linkSync;
  const realOpen = fs.openSync;
  fs.linkSync = function link(from, to) {
    if (to === p.runLock) seen.push({ link: path.basename(from), content: fs.readFileSync(from, 'utf8') });
    return realLink.call(this, from, to);
  };
  fs.openSync = function open(file, flags, ...rest) {
    if (file === p.runLock && flags !== 'r') seen.push({ openForWriting: flags });
    return realOpen.call(this, file, flags, ...rest);
  };
  let release;
  try {
    release = acquireRunLock(p, 'issue'); // first link fails (stale lock), second one succeeds
  } finally {
    fs.linkSync = realLink;
    fs.openSync = realOpen;
  }
  assert.equal(seen.length, 2, JSON.stringify(seen));
  for (const s of seen) {
    assert.match(s.link, /^run\.lock\.\d+-\d+-[0-9a-f]{12}\.tmp$/);
    const payload = JSON.parse(s.content);
    assert.equal(payload.pid, process.pid);
    assert.equal(payload.command, 'issue');
    assert.match(payload.token, /^\d+-\d+-[0-9a-f]{12}$/);
  }
  assert.equal(fs.readFileSync(p.runLock, 'utf8'), seen[1].content);
  assert.deepEqual(protocolFiles(p.dir), []);
  release();
});

test('locks: a lock file another process has just created but not yet written (older versions) is not taken over', () => {
  const p = tmpPaths();
  const fdA = fs.openSync(p.runLock, 'wx'); // process A between open('wx') and write
  assert.throws(() => acquireRunLock(p, 'issue'), LockError);
  fs.writeFileSync(fdA, JSON.stringify({ pid: process.pid, command: 'verify', token: 'A' }));
  fs.closeSync(fdA);
  assert.equal(runningCommand(p), 'verify', 'A still holds the lock it created');
});

test('locks: B takes the stale lock while A is between judging it stale and removing it: A neither removes nor moves B\'s lock', () => {
  const isClaim = (file) => String(file).endsWith('.reap');
  const points = {
    // A judged the lock stale and is about to create its claim (the reviewers' race: A would unlink B's new lock).
    'before A claims it': (fire) => ({ name: 'openSync', hooks: { before: (file) => { if (isClaim(file)) fire(); } } }),
    'right after A claimed it': (fire) => ({ name: 'openSync', hooks: { after: (fd, file) => { if (isClaim(file)) fire(); } } }),
    // A holds its claim and is about to look at the lock again / move it aside.
    'before A looks again': (fire) => {
      let claimed = false;
      return { name: 'openSync', hooks: { before: (file, flags) => { if (isClaim(file)) claimed = true; else if (claimed && file.endsWith('run.lock') && flags === 'r') fire(); } } };
    },
    'before A moves it aside': (fire) => ({ name: 'renameSync', hooks: { before: (from) => { if (from.endsWith('run.lock')) fire(); } } }),
    'right after A moved it aside': (fire) => ({ name: 'renameSync', hooks: { after: (r, from) => { if (from.endsWith('run.lock')) fire(); } } }),
  };
  for (const [point, makeHook] of Object.entries(points)) {
    const p = tmpPaths();
    writeStaleLock(p.runLock);
    const errors = [];
    let b = null;
    let fired = false;
    const { name, hooks } = makeHook(() => {
      if (fired) return;
      fired = true;
      b = attempt(p, 'verify', errors);
    });
    const a = withoutPauses(() => hookFs(name, hooks, () => attempt(p, 'issue', errors)));
    assert.ok(fired, point);
    checkOutcome(p, [a, b], errors, { someoneGetsIt: true }, `B ${point}`);
  }
});

test('locks: three processes after a crash (B checks, A wins, B moves aside, C arrives) never yield two holders', () => {
  // The interleaving that defeats "rename aside, verify, put back" on its own: A completes between
  // B's check and B's rename; C arrives while the name is free. Here B only renames under its claim.
  const p = tmpPaths();
  writeStaleLock(p.runLock);
  const errors = [];
  let a = null;
  let c = null;
  const realRename = fs.renameSync;
  let fired = false;
  fs.renameSync = function rename(from, to) {
    if (fired || from !== p.runLock) return realRename.call(this, from, to);
    fired = true;
    fs.renameSync = realRename;
    a = attempt(p, 'issue', errors);
    realRename.call(this, from, to);
    c = attempt(p, 'export', errors);
    return undefined;
  };
  let b;
  try {
    b = withoutPauses(() => attempt(p, 'verify', errors));
  } finally {
    fs.renameSync = realRename;
  }
  assert.ok(fired);
  checkOutcome(p, [a, b, c], errors, { someoneGetsIt: true }, 'A/B/C');
});

/**
 * The moments at which another process can act while `calls` fs calls are made: before each call,
 * and after the last one. ("Just after call n" is the same moment as "just before call n + 1":
 * only JavaScript runs in between.) Each is an interleave() option.
 */
function moments(calls) {
  const list = Array.from({ length: calls }, (_, at) => ({ at, label: `before call #${at}` }));
  if (calls > 0) list.push({ at: calls - 1, after: true, label: 'after the last call' });
  return list;
}

test('locks: every interleaving of two processes, at every file-system step, leaves at most one holder (exactly one when the lock is free or stale)', () => {
  let scenarios = 0;
  for (const [name, start] of Object.entries(STARTS)) {
    const p = tmpPaths();
    const run = (moment) => {
      const before = resetDir(p, start.setup);
      const errors = [];
      let a = null;
      let b = null;
      const calls = withoutPauses(() => interleave(() => { a = attempt(p, 'issue', errors); }, () => { b = attempt(p, 'verify', errors); }, moment));
      return { calls, parties: [a, b], errors, before };
    };
    const alone = run({ at: Infinity }); // A on its own: how many fs calls it makes
    for (const x of alone.parties) x?.release?.();
    for (const moment of moments(alone.calls)) {
      const r = run(moment);
      assert.ok(r.parties[1], `${name}: B ran ${moment.label}`);
      scenarios += 1;
      checkOutcome(p, r.parties, r.errors, { ...start, before: r.before }, `${name}: B ${moment.label} of A`);
    }
  }
  assert.ok(scenarios > 70, `${scenarios} scenarios`);
});

test('locks: every interleaving of three processes after a crash leaves exactly one holder', { timeout: 300_000 }, () => {
  // Both shapes of three processes with two preemptions, at every moment:
  //   nested      B acts at a moment of A, and C at a moment of B;
  //   one by one  B acts at a moment of A, then C at the same or a later moment of A.
  // "One by one" is what defeats "rename aside, verify, put back" on its own: A checks the stale
  // lock; B removes it and takes the lock; A moves B's lock aside; C takes the free name.
  // Without a stale lock there is nothing but the atomic link(), and an old unparseable lock is
  // removed by the same steps as a dead one (only the judgement differs): the two-process test and
  // the real-concurrency tests cover those, which keeps this one quick.
  const name = 'a stale lock (dead pid)';
  const start = STARTS[name];
  const p = tmpPaths();
  const scenario = (build) => {
    const before = resetDir(p, start.setup);
    const errors = [];
    const r = { a: null, b: null, c: null, callsB: 0 };
    const callsA = withoutPauses(() => build(r, errors));
    return { callsA, callsB: r.callsB, parties: [r.a, r.b, r.c], errors, before };
  };
  const nested = (momentA, momentB) => scenario((r, errors) => interleave(
    () => { r.a = attempt(p, 'issue', errors); },
    () => { r.callsB = interleave(() => { r.b = attempt(p, 'verify', errors); }, () => { r.c = attempt(p, 'export', errors); }, momentB); },
    momentA,
  ));
  const oneByOne = (momentB, momentC) => scenario((r, errors) => interleaveAll(() => { r.a = attempt(p, 'issue', errors); }, [
    { ...momentB, run: () => { r.b = attempt(p, 'verify', errors); } },
    { ...momentC, run: () => { r.c = attempt(p, 'export', errors); } },
  ]));
  const releaseAll = (r) => { for (const x of r.parties) x?.release?.(); };
  const order = (m) => m.at * 2 + (m.after ? 1 : 0);
  let scenarios = 0;

  const alone = nested({ at: Infinity }, { at: Infinity });
  releaseAll(alone);
  for (const momentA of moments(alone.callsA)) {
    // The calls a process makes before the next one comes in do not depend on that next one:
    // count them without it.
    const withB = nested(momentA, { at: Infinity });
    releaseAll(withB);
    for (const momentB of moments(withB.callsB)) {
      const r = nested(momentA, momentB);
      assert.ok(r.parties[2], 'C ran');
      scenarios += 1;
      checkOutcome(p, r.parties, r.errors, { ...start, before: r.before }, `nested: B ${momentA.label} of A, C ${momentB.label} of B`);
    }
    for (const momentC of moments(withB.callsA)) {
      if (order(momentC) < order(momentA)) continue;
      const r = oneByOne(momentA, momentC);
      assert.ok(r.parties[1] && r.parties[2], 'B and C ran');
      scenarios += 1;
      checkOutcome(p, r.parties, r.errors, { ...start, before: r.before }, `one by one: B ${momentA.label} of A, C ${momentC.label} of A`);
    }
  }
  assert.ok(scenarios > 600, `${scenarios} scenarios`);
});

test('locks: a stale lock whose removal was interrupted for good is reported as held, with the way out', () => {
  const p = tmpPaths();
  writeStaleLock(p.runLock);
  // Simulate a process killed right after creating its claim (no clean-up runs).
  const realOpen = fs.openSync;
  fs.openSync = function open(file, flags, ...rest) {
    const fd = realOpen.call(this, file, flags, ...rest);
    if (String(file).endsWith('.reap')) {
      fs.closeSync(fd);
      throw Object.assign(new Error('killed'), { code: 'KILLED' });
    }
    return fd;
  };
  try {
    assert.throws(() => acquireRunLock(p, 'issue'), /killed/);
  } finally {
    fs.openSync = realOpen;
  }
  assert.equal(protocolFiles(p.dir).filter((f) => f.endsWith('.reap')).length, 1, 'the claim was left behind');
  assert.throws(() => withoutPauses(() => acquireRunLock(p, 'verify')), (err) => err instanceof LockError && err.file === p.runLock
    && err.message.includes(`锁文件：${p.runLock}。`) && err.message.endsWith(HINT));
  assert.equal(fs.existsSync(p.runLock), true, 'nothing else touches a claimed stale lock');
  fs.unlinkSync(p.runLock); // what the message tells the user to do
  const release = acquireRunLock(p, 'verify');
  assert.equal(runningCommand(p), 'verify');
  release();
});

test('locks: without hard links (some file systems) the lock is created exclusively and everything still works', () => {
  const p = tmpPaths();
  const realLink = fs.linkSync;
  fs.linkSync = function link(from, to) {
    if (to === p.runLock) throw Object.assign(new Error('ENOTSUP: operation not supported'), { code: 'ENOTSUP' });
    return realLink.call(this, from, to);
  };
  try {
    const release = acquireRunLock(p, 'issue');
    assert.equal(runningCommand(p), 'issue');
    assert.throws(() => acquireRunLock(p, 'verify'), (err) => err instanceof LockError && err.message.startsWith('issue 正在运行'));
    release();
    assert.equal(fs.existsSync(p.runLock), false);
    writeStaleLock(p.runLock);
    const again = acquireRunLock(p, 'verify');
    assert.equal(runningCommand(p), 'verify', 'a stale lock is still removed');
    again();
    // A failure while writing the lock never leaves our own half-written lock behind.
    let lockFd = null;
    const failure = () => hookFs('openSync', { after: (fd, file, flags) => { if (file === p.runLock && flags === 'wx') lockFd = fd; } }, () => hookFs('writeFileSync', {
      before: (target) => { if (target === lockFd) throw Object.assign(new Error('ENOSPC: no space left on device'), { code: 'ENOSPC' }); },
    }, () => acquireRunLock(p, 'issue')));
    assert.throws(failure, /ENOSPC/);
    assert.notEqual(lockFd, null, 'the lock itself was being written');
  } finally {
    fs.linkSync = realLink;
  }
  assert.deepEqual(fs.readdirSync(p.dir), []);
});

// ---------------------------------------------------------------------------
// Locks: real concurrency (separate threads and processes, real system calls)
// ---------------------------------------------------------------------------

const WORKER_SOURCE = `
import { parentPort, workerData } from 'node:worker_threads';
const { acquireRunLock } = await import(workerData.campaignUrl);
const go = new Int32Array(workerData.sab);
let release = null;
parentPort.on('message', (m) => {
  if (m.release) {
    release?.();
    release = null;
    parentPort.postMessage('released');
    return;
  }
  parentPort.postMessage('armed');
  while (Atomics.load(go, 0) !== m.trial) { /* spin, so that every worker starts at the same instant */ }
  try {
    release = acquireRunLock({ runLock: m.lock }, 'issue');
    parentPort.postMessage({ got: true });
  } catch (err) {
    parentPort.postMessage({ got: false, unexpected: err.name === 'LockError' ? null : String(err.stack) });
  }
});
`;

test('locks: 8 threads × 200 trials started together never hold the lock twice (with and without a stale lock)', { timeout: 300_000 }, async () => {
  const dir = tmpDir('gcp-lock-workers-');
  const script = path.join(dir, 'worker.mjs');
  fs.writeFileSync(script, WORKER_SOURCE);
  const lockDir = path.join(dir, 'campaign');
  fs.mkdirSync(lockDir);
  const lock = path.join(lockDir, 'run.lock');
  const sab = new SharedArrayBuffer(4);
  const go = new Int32Array(sab);
  const workers = Array.from({ length: 8 }, () => new Worker(script, { workerData: { campaignUrl: CAMPAIGN_URL, sab } }));
  const next = (w) => once(w, 'message').then(([m]) => m);
  const starts = {
    'stale lock': () => writeStaleLock(lock),
    'old empty lock': () => { fs.writeFileSync(lock, ''); setAge(lock, 3_600_000); },
    'no lock': () => {},
  };
  let trial = 0;
  try {
    for (const [name, setup] of Object.entries(starts)) {
      const histogram = {};
      for (let t = 0; t < 200; t += 1) {
        trial += 1;
        setup();
        const armed = workers.map(next);
        for (const w of workers) w.postMessage({ trial, lock });
        await Promise.all(armed);
        const answers = workers.map(next);
        Atomics.store(go, 0, trial);
        const results = await Promise.all(answers);
        assert.deepEqual(results.filter((r) => r.unexpected), [], name);
        const got = results.filter((r) => r.got).length;
        histogram[got] = (histogram[got] ?? 0) + 1;
        const released = workers.map(next);
        for (const w of workers) w.postMessage({ release: true });
        await Promise.all(released);
        assert.deepEqual(fs.readdirSync(lockDir), [], `${name}, trial ${t}: nothing left after release`);
      }
      assert.deepEqual(histogram, { 1: 200 }, `${name}: holders per trial`);
    }
  } finally {
    await Promise.all(workers.map((w) => w.terminate()));
  }
});

const CHILD_SOURCE = `
const { acquireRunLock } = await import(process.argv[2]);
let release = null;
process.on('message', (m) => {
  if (m.release) {
    release?.();
    release = null;
    process.send('released');
    return;
  }
  if (m.hold) {
    release = acquireRunLock({ runLock: m.lock }, 'issue');
    process.send('holding');
    return;
  }
  while (Date.now() < m.startAt) { /* spin, so that every process starts at the same instant */ }
  try {
    release = acquireRunLock({ runLock: m.lock }, 'export');
    process.send({ got: true });
  } catch (err) {
    process.send({ got: false, unexpected: err.name === 'LockError' ? null : String(err.stack) });
  }
});
process.send('ready');
`;

function spawnChild(scriptDir) {
  const script = path.join(scriptDir, 'child.mjs');
  if (!fs.existsSync(script)) fs.writeFileSync(script, CHILD_SOURCE);
  const child = fork(script, [CAMPAIGN_URL], { execArgv: [], stdio: ['ignore', 'inherit', 'inherit', 'ipc'] });
  const messages = [];
  const waiters = [];
  let gone = null;
  child.on('message', (m) => (waiters.length ? waiters.shift().resolve(m) : messages.push(m)));
  child.on('exit', (code, signal) => {
    gone = new Error(`child ${child.pid} exited (${signal ?? code})`);
    while (waiters.length) waiters.shift().reject(gone);
  });
  child.next = () => {
    if (messages.length) return Promise.resolve(messages.shift());
    if (gone) return Promise.reject(gone);
    return new Promise((resolve, reject) => waiters.push({ resolve, reject }));
  };
  return child;
}

test('locks: separate processes started together after a crash never hold the lock twice', { timeout: 300_000 }, async () => {
  const dir = tmpDir('gcp-lock-procs-');
  const lockDir = path.join(dir, 'campaign');
  fs.mkdirSync(lockDir);
  const lock = path.join(lockDir, 'run.lock');
  const children = Array.from({ length: 6 }, () => spawnChild(dir));
  try {
    await Promise.all(children.map((c) => c.next()));
    const histogram = {};
    for (let t = 0; t < 40; t += 1) {
      if (t % 4 !== 3) writeStaleLock(lock); // mostly a stale lock (dead pid); sometimes none
      const startAt = Date.now() + 30;
      for (const c of children) c.send({ lock, startAt });
      const results = await Promise.all(children.map((c) => c.next()));
      assert.deepEqual(results.filter((r) => r.unexpected), []);
      const got = results.filter((r) => r.got).length;
      histogram[got] = (histogram[got] ?? 0) + 1;
      for (const c of children) c.send({ release: true });
      await Promise.all(children.map((c) => c.next()));
      assert.deepEqual(fs.readdirSync(lockDir), [], `trial ${t}: nothing left after release`);
    }
    assert.deepEqual(histogram, { 1: 40 }, 'holders per trial');
  } finally {
    for (const c of children) c.kill('SIGKILL');
    await Promise.all(children.map((c) => (c.exitCode === null && c.signalCode === null ? once(c, 'exit') : null)));
  }
});

test('locks: the lock of a process that was killed is cleaned up by the next command', { timeout: 60_000 }, async () => {
  const p = tmpPaths();
  const child = spawnChild(tmpDir('gcp-lock-child-'));
  const exited = once(child, 'exit');
  try {
    assert.equal(await child.next(), 'ready');
    child.send({ hold: true, lock: p.runLock });
    assert.equal(await child.next(), 'holding');
    assert.equal(runningCommand(p), 'issue');
    assert.throws(() => acquireRunLock(p, 'verify'), (err) => err instanceof LockError && err.message.startsWith(`issue 正在运行（PID ${child.pid}，`));
  } finally {
    child.kill('SIGKILL'); // e.g. the second Ctrl+C, a closed terminal, a crash: no clean-up runs
    await exited;
  }
  assert.equal(JSON.parse(fs.readFileSync(p.runLock, 'utf8')).pid, child.pid, 'the lock was left behind');
  assert.equal(runningCommand(p), null);
  const release = acquireRunLock(p, 'verify');
  assert.equal(runningCommand(p), 'verify');
  release();
  assert.deepEqual(fs.readdirSync(p.dir), []);
});

// ---------------------------------------------------------------------------
// Journal
// ---------------------------------------------------------------------------

test('campaign: journal lines are appended with a time; a torn last line is ignored, a bad middle line is fatal', () => {
  const p = tmpPaths();
  assert.deepEqual(readJournal(p.journal), []);
  appendJournal(p.journal, { op: 'run.start', run: 'r1' }, { now: () => '2026-10-05T16:00:00.000Z' });
  appendJournal(p.journal, { op: 'run.end', run: 'r1' });
  const entries = readJournal(p.journal);
  assert.equal(entries.length, 2);
  assert.equal(entries[0].t, '2026-10-05T16:00:00.000Z');
  assert.match(entries[1].t, /^\d{4}-\d{2}-\d{2}T/);
  fs.appendFileSync(p.journal, '{"op":"create.st');
  assert.equal(readJournal(p.journal).length, 2, 'torn last line from a crash is ignored');
  fs.appendFileSync(p.journal, '\n{"op":"run.start","run":"r2"}\n');
  assert.throws(() => readJournal(p.journal), /line 3 is not valid JSON/);
});

test('campaign: an append after a crash mid-write neither loses the new entry nor breaks later reads', () => {
  const p = tmpPaths();
  appendJournal(p.journal, { op: 'a' });
  fs.appendFileSync(p.journal, '{"op":"create.st'); // torn line left by a crash
  appendJournal(p.journal, { op: 'b' });
  appendJournal(p.journal, { op: 'c' });
  assert.deepEqual(readJournal(p.journal).map((e) => e.op), ['a', 'b', 'c']);

  // A complete last line that only lacks its newline is kept.
  fs.appendFileSync(p.journal, '{"op":"create.ok","cid":"x"}');
  appendJournal(p.journal, { op: 'd' });
  assert.deepEqual(readJournal(p.journal).map((e) => e.op), ['a', 'b', 'c', 'create.ok', 'd']);

  // A torn multi-byte character, a very long torn line, and a file that is only a torn line.
  const bytes = Buffer.from('{"op":"skip","detail":"已退订营', 'utf8');
  fs.appendFileSync(p.journal, bytes.subarray(0, bytes.length - 1));
  appendJournal(p.journal, { op: 'e' });
  fs.appendFileSync(p.journal, `{"op":"y","pad":"${'x'.repeat(200_000)}`);
  appendJournal(p.journal, { op: 'f' });
  assert.deepEqual(readJournal(p.journal).map((e) => e.op), ['a', 'b', 'c', 'create.ok', 'd', 'e', 'f']);
  const only = path.join(p.dir, 'only.jsonl');
  fs.writeFileSync(only, '{"op":"x');
  appendJournal(only, { op: 'g' });
  assert.deepEqual(readJournal(only).map((e) => e.op), ['g']);
});

const CID = 'gid://shopify/Customer/1';
const fold = (entries) => foldJournal(entries.map((e, i) => ({ t: `2026-10-05T16:00:${String(i).padStart(2, '0')}.000Z`, ...e })));

test('fold: the issue life cycle of one customer', () => {
  let s = fold([{ op: 'create.start', cid: CID, amountCents: 1077, batch: 1, run: 'r1' }]);
  assert.equal(s.customers.get(CID).status, STATUS.IN_PROGRESS);
  assert.equal(s.issuingStarted, true);
  s = fold([{ op: 'create.start', cid: CID, amountCents: 1077, batch: 1 }, { op: 'create.ok', cid: CID, giftCardId: 'gid://shopify/GiftCard/9', last4: 'x009', amountCents: 1077, batch: 1 }]);
  assert.equal(s.customers.get(CID).status, STATUS.CREATED);
  assert.equal(s.customers.get(CID).giftCardId, 'gid://shopify/GiftCard/9');
  s = fold([{ op: 'create.start', cid: CID, batch: 1 }, { op: 'create.ok', cid: CID, giftCardId: 'g', batch: 1 }, { op: 'tag.ok', cid: CID }]);
  const done = s.customers.get(CID);
  assert.equal(done.status, STATUS.DONE);
  assert.equal(done.taggedAt, '2026-10-05T16:00:02.000Z');
  assert.equal(done.attempts, 1);
  assert.equal(done.batch, 1);
});

test('fold: failures, rejections, unknown outcomes and reconciliation', () => {
  const st = (entries) => fold(entries).customers.get(CID);
  assert.equal(st([{ op: 'create.start', cid: CID }, { op: 'create.fail', cid: CID, error: 'bad' }]).status, STATUS.FAILED);
  assert.equal(st([{ op: 'create.start', cid: CID }, { op: 'create.rejected', cid: CID, error: 'throttled' }]).status, STATUS.PENDING);
  assert.equal(st([{ op: 'create.start', cid: CID }, { op: 'create.unknown', cid: CID, error: 'lost' }]).status, STATUS.UNKNOWN);
  assert.equal(st([{ op: 'create.start', cid: CID }, { op: 'create.unknown', cid: CID }, { op: 'reconcile.none', cid: CID, note: 'not found' }]).status, STATUS.PENDING);
  const found = st([{ op: 'create.start', cid: CID }, { op: 'create.unknown', cid: CID }, { op: 'reconcile.found', cid: CID, giftCardId: 'g1', createdAt: '2026-10-05T16:00:00Z', source: 'issue-reconcile' }]);
  assert.equal(found.status, STATUS.CREATED);
  assert.equal(found.reconciledFrom, 'issue-reconcile');
  assert.equal(found.createdAt, '2026-10-05T16:00:00Z');
  assert.equal(found.error, null);
  // A tag found before the card (preflight with the tag already present) still ends as done.
  assert.equal(st([{ op: 'tag.ok', cid: CID }, { op: 'reconcile.found', cid: CID, giftCardId: 'g1', source: 'preflight' }]).status, STATUS.DONE);
  const tagFail = st([{ op: 'create.start', cid: CID }, { op: 'create.ok', cid: CID, giftCardId: 'g' }, { op: 'tag.fail', cid: CID, error: 'nope' }]);
  assert.equal(tagFail.status, STATUS.CREATED);
  assert.equal(tagFail.error, 'nope');
  // Retrying a failed customer counts attempts.
  assert.equal(st([{ op: 'create.start', cid: CID }, { op: 'create.fail', cid: CID }, { op: 'create.start', cid: CID }, { op: 'create.ok', cid: CID, giftCardId: 'g' }]).attempts, 2);
});

test('fold: a pre-flight skip only applies to people not yet issued', () => {
  const st = (entries) => fold(entries).customers.get(CID);
  const skipped = st([{ op: 'skip', cid: CID, reason: 'not-subscribed', batch: 2 }]);
  assert.equal(skipped.status, STATUS.SKIPPED);
  assert.equal(skipped.skipReason, 'not-subscribed');
  assert.equal(skipped.batch, 2);
  assert.equal(st([{ op: 'create.start', cid: CID }, { op: 'create.ok', cid: CID, giftCardId: 'g' }, { op: 'skip', cid: CID, reason: 'already-tagged' }]).status, STATUS.CREATED);
  assert.equal(st([{ op: 'create.start', cid: CID }, { op: 'create.fail', cid: CID }, { op: 'skip', cid: CID, reason: 'no-email' }]).status, STATUS.SKIPPED);
  assert.equal(fold([{ op: 'skip', cid: CID, reason: 'x' }]).issuingStarted, true, 'a skip is written by a live issue run: it freezes the list too');
});

test('fold: issuing has started (the list is frozen) once a live issue run started or the journal has any issue op', () => {
  const started = (entries) => fold(entries).issuingStarted;
  assert.equal(started([]), false);
  // A live issue run freezes the list even when it stopped before its first card (issue-duplicates#2).
  assert.equal(started([{ op: 'run.start', run: 'r1', command: 'issue', dryRun: false, batch: 1 }]), true);
  assert.equal(started([{ op: 'run.start', run: 'r1', command: 'issue', dryRun: false, batch: 1 }, { op: 'skip', cid: CID, reason: 'not-subscribed', batch: 1 }, { op: 'run.end', run: 'r1', exitCode: 130 }]), true);
  assert.equal(started([{ op: 'run.start', run: 'r1', command: 'issue', batch: 1 }]), true, 'without a dryRun flag a run counts as live (as in runs[])');
  assert.equal(started([{ op: 'run.start', run: 'r1', command: 'issue', dryRun: true, batch: 1 }, { op: 'run.end', run: 'r1', exitCode: 0 }]), false, 'dry runs do not');
  for (const command of ['select', 'verify', 'usage', 'export']) {
    assert.equal(started([{ op: 'run.start', run: 'r1', command, dryRun: false }, { op: 'run.end', run: 'r1', exitCode: 0 }]), false, command);
  }
  // Any issue op, even without a create.start (e.g. the journal lost its lines and verify re-recorded the card: selection#1).
  for (const op of ['create.start', 'create.ok', 'create.fail', 'create.rejected', 'create.unknown', 'reconcile.found', 'reconcile.none', 'tag.ok', 'tag.fail', 'skip']) {
    assert.equal(started([{ op, cid: CID, giftCardId: 'gid://shopify/GiftCard/5', reason: 'x', source: 'verify' }]), true, op);
  }
  assert.equal(started([
    { op: 'run.start', run: 'v1', command: 'verify', dryRun: false },
    { op: 'reconcile.found', cid: CID, giftCardId: 'gid://shopify/GiftCard/5', source: 'verify', run: 'v1' },
    { op: 'run.end', run: 'v1', exitCode: 0 },
  ]), true);
  // Used-tag ops, the old reminder ops (journals written before remind was removed) and unknown ops do not (only issue ops do).
  for (const op of ['used.tag.ok', 'used.tag.fail', 'remind.start', 'remind.ok', 'remind.fail', 'remind.unknown', 'remind.rejected', 'remind.skip', 'remind.found', 'remind.tag.fail', 'remind.tag.ok', 'some.future.op']) {
    assert.equal(started([{ op, cid: CID, round: 1 }]), false, op);
  }
  assert.equal(started([{ op: 'used.tag.ok', cid: CID, giftCardId: 'g', run: 'u1' }, { op: 'used.tag.fail', cid: CID, giftCardId: 'g', error: 'x', run: 'u2' }]), false);
  // `issue --repair-only` never creates a card: its start alone does not freeze the list; what it records does.
  assert.equal(started([{ op: 'run.start', run: 'r1', command: 'issue', dryRun: false, batch: 1, options: { repairOnly: true } }, { op: 'run.end', run: 'r1', exitCode: 0 }]), false);
  assert.equal(started([{ op: 'run.start', run: 'r1', command: 'issue', dryRun: false, batch: 1, options: { repairOnly: true } }, { op: 'tag.ok', cid: CID, run: 'r1' }]), true);
});

test('fold: lastIssuedBatch counts only batches that attempted a card (repair-only / date-blocked runs create none)', () => {
  const s = fold([
    { op: 'run.start', run: 'r1', command: 'issue', dryRun: false, batch: 1 },
    { op: 'create.start', cid: CID, amountCents: 1077, batch: 1, run: 'r1' },
    { op: 'run.end', run: 'r1', exitCode: 0 },
    { op: 'run.start', run: 'r2', command: 'issue', dryRun: false, batch: 2, options: { repairOnly: true } },
    { op: 'run.end', run: 'r2', exitCode: 0 },
    { op: 'run.start', run: 'r3', command: 'issue', dryRun: false, batch: 3 },
    { op: 'run.end', run: 'r3', exitCode: 2, summary: { newCardsRefused: 4 } },
  ]);
  assert.equal(s.lastBatch, 3, 'the next run still takes a new number');
  assert.equal(s.lastIssuedBatch, 1);
  assert.equal(fold([]).lastIssuedBatch, 0);
});

test('fold: runs, batches and dry runs', () => {
  const s = fold([
    { op: 'run.start', run: 'r1', command: 'issue', dryRun: true, batch: 1, limit: 20 },
    { op: 'run.end', run: 'r1', summary: { attempted: 0 }, exitCode: 0 },
    { op: 'run.start', run: 'r2', command: 'issue', dryRun: false, batch: 1, limit: 20 },
    { op: 'run.end', run: 'r2', summary: { created: 20 }, exitCode: 0 },
    { op: 'run.start', run: 'r3', command: 'export', dryRun: false, options: { refresh: true } },
    { op: 'some.future.op', cid: CID },
  ]);
  assert.equal(s.lastBatch, 1, 'dry runs do not use up a batch number');
  assert.equal(s.runs.length, 3);
  assert.deepEqual(s.runs.map((r) => [r.run, r.command, r.dryRun]), [['r1', 'issue', true], ['r2', 'issue', false], ['r3', 'export', false]]);
  assert.equal(s.runs[1].exitCode, 0);
  assert.deepEqual(s.runs[1].summary, { created: 20 });
  assert.equal(s.runs[2].endedAt, null, 'a run that never ended (crash) stays open');
  assert.deepEqual(s.runs[2].options, { refresh: true });
  assert.equal(s.customers.size, 0, 'unknown ops are ignored, so newer journals stay readable');
  assert.equal(s.issuingStarted, true, 'r2 is a live issue run');
});

test('fold: used.tag.ok / used.tag.fail record the used-card tag per customer; the latest outcome wins', () => {
  const u = (entries) => fold(entries).customers.get(CID)?.usedTag;
  assert.equal(fold([{ op: 'create.start', cid: CID }]).customers.get(CID).usedTag, null, 'nothing attempted yet');
  // Tagged: the time, the card the tag was added for, no error.
  const ok = { op: 'used.tag.ok', cid: CID, giftCardId: 'gid://shopify/GiftCard/9', run: 'u1' };
  assert.deepEqual(u([ok]), { status: USED_TAG_STATUS.TAGGED, at: '2026-10-05T16:00:00.000Z', giftCardId: 'gid://shopify/GiftCard/9', error: null });
  // Failed: the error is kept (a failure without a message still counts); the next usage run retries.
  const fail = { op: 'used.tag.fail', cid: CID, giftCardId: 'gid://shopify/GiftCard/9', error: 'tagsAdd rejected: id: Customer not found', run: 'u1' };
  assert.deepEqual(u([fail]), { status: USED_TAG_STATUS.FAILED, at: '2026-10-05T16:00:00.000Z', giftCardId: 'gid://shopify/GiftCard/9', error: 'tagsAdd rejected: id: Customer not found' });
  assert.equal(u([{ op: 'used.tag.fail', cid: CID }]).error, 'tag failed');
  assert.equal(u([{ op: 'used.tag.fail', cid: CID }]).giftCardId, null);
  // A later ok replaces a failure (and vice versa: only the latest outcome is kept).
  assert.deepEqual(u([fail, ok]), { status: USED_TAG_STATUS.TAGGED, at: '2026-10-05T16:00:01.000Z', giftCardId: 'gid://shopify/GiftCard/9', error: null });
  assert.equal(u([ok, fail]).status, USED_TAG_STATUS.FAILED);
  // The issue state is left alone: a customer only known from a used.tag op stays pending, and a done one stays done.
  const s = fold([ok]);
  assert.equal(s.customers.get(CID).status, STATUS.PENDING);
  assert.equal(s.customers.get(CID).attempts, 0);
  assert.equal(s.issuingStarted, false);
  const done = fold([{ op: 'create.start', cid: CID, batch: 1 }, { op: 'create.ok', cid: CID, giftCardId: 'g', batch: 1 }, { op: 'tag.ok', cid: CID }, ok]).customers.get(CID);
  assert.equal(done.status, STATUS.DONE);
  assert.equal(done.usedTag.status, USED_TAG_STATUS.TAGGED);
  assert.equal(done.taggedAt, '2026-10-05T16:00:02.000Z', 'the sent tag time is not the used tag time');
});

test('fold: reminder ops of journals written before remind was removed are ignored', () => {
  const s = fold([
    { op: 'create.start', cid: CID, batch: 1 },
    { op: 'create.ok', cid: CID, giftCardId: 'g', batch: 1 },
    { op: 'tag.ok', cid: CID },
    { op: 'run.start', run: 'r3', command: 'remind', dryRun: false, options: { round: 1 } },
    { op: 'remind.start', cid: CID, round: 1, giftCardId: 'g', run: 'r3' },
    { op: 'remind.ok', cid: CID, round: 1, run: 'r3' },
    { op: 'remind.tag.fail', cid: CID, round: 1, error: 'x', run: 'r3' },
    { op: 'remind.found', cid: CID, round: 2, source: 'tag', run: 'r3' },
    { op: 'remind.skip', cid: 'gid://shopify/Customer/2', round: 1, reason: 'used' },
    { op: 'run.end', run: 'r3', exitCode: 0, summary: { sent: 1 } },
  ]);
  const c = s.customers.get(CID);
  assert.equal(c.status, STATUS.DONE);
  assert.equal(c.usedTag, null);
  assert.equal('reminders' in c, false, 'no reminder state exists any more');
  assert.equal(s.customers.has('gid://shopify/Customer/2'), false, 'a remind.skip creates no customer state');
  // The run itself still shows in the run history (its command name as written).
  assert.deepEqual(s.runs.map((r) => [r.run, r.command, r.exitCode]), [['r3', 'remind', 0]]);
});

test('campaign: run ids sort by time', () => {
  assert.match(newRunId(new Date('2026-10-05T16:04:05.678Z')), new RegExp(`^20261005160405-${process.pid}$`));
  assert.ok(newRunId(new Date('2026-10-05T16:04:05Z')) < newRunId(new Date('2026-10-05T16:04:06Z')));
});
