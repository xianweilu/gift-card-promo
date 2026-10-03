// Builds the campaign workbook (campaigns/<id>/gift-card-promo-<id>.xlsx) from
// local state only: selection.json + journal.jsonl, plus tags.json, verify.json
// and usage.json when they exist. The program never reads the workbook back;
// every command simply regenerates it.
//
// Design notes
// - Streaming (ExcelJS WorkbookWriter): each row is styled, committed and
//   released, so 20k recipients + 50k "未入选" rows stay within a flat memory
//   budget. A committed row can no longer be touched, so cells are styled first.
// - Excel has no time zones: times are written as the STORE-LOCAL wall clock
//   (selection.params.timezone), i.e. Dates whose UTC fields are the local time.
// - Links are HYPERLINK() formulas, because Excel caps real hyperlinks at ~65k
//   per sheet.
// - The file is written to a temp file in the same directory and renamed over
//   the target while holding the Excel lock, so nobody sees half a file and a
//   later writer always includes the newest journal entries.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';

import { readJson, readJournal, foldJournal, runningCommand, withExcelLock, ensureDir } from '../campaign.js';
import { campaignNote } from '../giftcards.js';
import { formatUsd, tierLabel, describeTiers } from '../select/amount.js';
import { channelLabel, UNLISTED_RULES } from '../select/rules.js';
import {
  STATUS_LABELS,
  STATUS_ORDER,
  STATUS_FILLS,
  STATUS_HELP,
  ISSUE_SKIP_LABELS,
  REMIND_SKIP_LABELS,
  REMIND_HELP,
  REMIND_SENT_BY_TAG,
  REMIND_TAG_MISSING_SUFFIX,
  KIND_LABELS,
  MARKETING_LABELS,
  BASIS_NOTES,
  NEVER_NOTES,
  OP_LABELS,
  RECONCILE_SOURCE_LABELS,
  VERIFY_TYPE_LABELS,
  SNAPSHOT_SOURCE_LABELS,
  UNMATCHED_REASON_LABELS,
  EXIT_CODE_HELP,
  USAGE_SPLIT_NOTE,
  labelOf,
  issueSkipText,
  remindSkipText,
  noteText,
  errorText,
  summaryText,
  roundTagName,
  roundTagHelp,
} from './labels.js';

export const DEFAULT_TIMEZONE = 'America/Los_Angeles';
export const OPEN_IN_EXCEL_WARNING = 'Excel 正打开此文件，请关闭后重新打开才能看到最新内容';
const NO_SELECTION_MESSAGE = '还没有名单，请先运行 select';

const DAY_MS = 86_400_000;
// A tags.json snapshot overrides a journal tag.ok only when taken this long after it
// (Shopify's customer search may not show a just-added tag yet).
const TAG_INDEX_GRACE_MS = 10 * 60_000;
const EXCEL_MAX_ROWS = 1_048_576;
const MAX_CELL_TEXT = 32_767; // Excel's per-cell character limit
// Hand control back to the event loop after this many cells so the zip/deflate
// pipeline can drain; without it ExcelJS buffers whole sheets in memory.
const YIELD_EVERY_CELLS = 500;

// ---------------------------------------------------------------------------
// Exported helpers (also used by tests)
// ---------------------------------------------------------------------------

/**
 * Status as shown in the Excel. A row left "in_progress" is only really in
 * progress while issue is running; otherwise its outcome is unknown.
 */
export function effectiveStatus(status, running) {
  const s = status || 'pending';
  return s === 'in_progress' && running !== 'issue' ? 'unknown' : s;
}

/** Chinese label of an issue status, e.g. "已完成"; see effectiveStatus for in_progress. */
export function statusLabel(status, running) {
  const s = effectiveStatus(status, running);
  return STATUS_LABELS[s] ?? String(s);
}

/** Reminder status as shown: a stale in_progress (remind not running) is unknown. */
export function effectiveRemindStatus(status, running) {
  return status === 'in_progress' && running !== 'remind' ? 'unknown' : status;
}

/**
 * Text of a "第 N 次提醒" cell for one reminder state from foldJournal
 * ({ status, at, reason, error, source, tagError, ... } or undefined), times in store-local `tz`.
 * A sent state with source 'tag' was found by its round tag in Shopify (remind.found: the
 * journal had no send, so there is no time); a tagError means the round tag is still missing.
 */
export function remindLabel(reminder, running, tz = DEFAULT_TIMEZONE) {
  if (!reminder || !reminder.status) return '';
  switch (effectiveRemindStatus(reminder.status, running)) {
    case 'sent': {
      let text;
      if (reminder.source === 'tag') {
        text = REMIND_SENT_BY_TAG;
      } else {
        const at = clockFor(tz).stamp(reminder.at);
        text = at ? `已发 ${at.slice(5)}` : '已发';
      }
      return reminder.tagError !== undefined && reminder.tagError !== null ? `${text}${REMIND_TAG_MISSING_SUFFIX}` : text;
    }
    case 'skipped':
      return reminder.reason ? `跳过：${labelOf(REMIND_SKIP_LABELS, reminder.reason)}` : '跳过';
    case 'failed': {
      const error = errorText(reminder.error);
      return error ? `失败：${error}` : '失败';
    }
    case 'unknown':
      return '结果不明';
    case 'in_progress':
      return '进行中';
    default:
      return String(reminder.status);
  }
}

/** Numeric part of a Shopify gid ("gid://shopify/Customer/123" → "123"); '' when there is none. */
export function numericId(gid) {
  if (gid === null || gid === undefined) return '';
  const m = /(\d+)(?:\?[^/]*)?$/.exec(String(gid).trim());
  return m ? m[1] : '';
}

const ADMIN_SEGMENTS = {
  customer: 'customers',
  customers: 'customers',
  order: 'orders',
  orders: 'orders',
  giftcard: 'gift_cards',
  giftcards: 'gift_cards',
  gift_card: 'gift_cards',
  gift_cards: 'gift_cards',
};

/**
 * Shopify admin page of a customer, order or gift card:
 * https://admin.shopify.com/store/<shop>/<customers|orders|gift_cards>/<numeric id>.
 * `type` may be singular, plural or omitted (then taken from the gid). '' when no id.
 */
export function adminUrl(shop, type, gid) {
  const id = numericId(gid);
  if (!shop || !id) return '';
  const key = (v) => String(v ?? '').toLowerCase().replace(/[\s-]/g, '_');
  const gidType = /^gid:\/\/shopify\/(\w+)\//.exec(String(gid))?.[1];
  const segment = ADMIN_SEGMENTS[key(type)] ?? ADMIN_SEGMENTS[key(gidType)];
  if (!segment) return '';
  return `https://admin.shopify.com/store/${encodeURIComponent(shop)}/${segment}/${id}`;
}

// ---------------------------------------------------------------------------
// Store-local time
// ---------------------------------------------------------------------------

const zones = new Map();

/**
 * Fast wall clock for one IANA zone. src/time.js builds a new Intl formatter
 * per call (~0.1 ms), too slow for hundreds of thousands of cells; here one
 * formatter is reused and the UTC offset is cached per UTC day (a day whose
 * offset changes, i.e. a DST switch, is computed exactly per instant).
 */
function zoneFor(timeZone) {
  let zone = zones.get(timeZone);
  if (zone) return zone;
  const fmt = new Intl.DateTimeFormat('en-US', {
    timeZone,
    hourCycle: 'h23',
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
    hour: 'numeric',
    minute: 'numeric',
    second: 'numeric',
  });
  const exactOffset = (ms) => {
    const p = {};
    for (const part of fmt.formatToParts(ms)) p[part.type] = part.value;
    const asUtc = Date.UTC(Number(p.year), Number(p.month) - 1, Number(p.day), Number(p.hour), Number(p.minute), Number(p.second));
    return asUtc - Math.floor(ms / 1000) * 1000;
  };
  const byDay = new Map(); // UTC day number → offset ms, or null when it changes that day
  const offset = (ms) => {
    const day = Math.floor(ms / DAY_MS);
    let off = byDay.get(day);
    if (off === undefined) {
      const start = exactOffset(day * DAY_MS);
      off = start === exactOffset(day * DAY_MS + DAY_MS - 1000) ? start : null;
      byDay.set(day, off);
    }
    return off ?? exactOffset(ms);
  };
  zone = { timeZone, wall: (ms) => ms + offset(ms) };
  zones.set(timeZone, zone);
  return zone;
}

function toMs(v) {
  if (v === null || v === undefined || v === '') return null;
  const ms = v instanceof Date ? v.getTime() : typeof v === 'number' ? v : Date.parse(String(v));
  return Number.isFinite(ms) ? ms : null;
}

const pad2 = (n) => String(n).padStart(2, '0');

const clocks = new Map();

/** Converters from instants (ISO strings, ms or Dates) to store-local Excel values. */
function clockFor(timeZone) {
  let clock = clocks.get(timeZone);
  if (clock) return clock;
  const zone = zoneFor(timeZone);
  const wall = (v) => {
    const ms = toMs(v);
    return ms === null ? null : zone.wall(ms);
  };
  clock = {
    timeZone,
    /** Local date + time (to the second) as an Excel datetime value. */
    time(v) {
      const w = wall(v);
      return w === null ? null : new Date(Math.floor(w / 1000) * 1000);
    },
    /** Local calendar date (midnight) as an Excel date value. */
    day(v) {
      const w = wall(v);
      return w === null ? null : new Date(Math.floor(w / DAY_MS) * DAY_MS);
    },
    /** Local calendar day number, for "days between" maths. */
    dayNumber(v) {
      const w = wall(v);
      return w === null ? null : Math.floor(w / DAY_MS);
    },
    /** "YYYY-MM-DD HH:MM" local, for texts. '' when missing. */
    stamp(v) {
      const w = wall(v);
      if (w === null) return '';
      const d = new Date(w);
      return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())} ${pad2(d.getUTCHours())}:${pad2(d.getUTCMinutes())}`;
    },
  };
  clocks.set(timeZone, clock);
  return clock;
}

/** A calendar date string "YYYY-MM-DD" (already store-local) → Excel date value. */
function ymd(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s ?? ''));
  return m ? new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]))) : null;
}

function validTimeZone(tz) {
  try {
    new Intl.DateTimeFormat('en-US', { timeZone: tz });
    return true;
  } catch {
    return false;
  }
}

function timeZoneLabel(tz) {
  return tz === DEFAULT_TIMEZONE ? '洛杉矶时间' : `${tz} 时间`;
}

// ---------------------------------------------------------------------------
// Cell values
// ---------------------------------------------------------------------------

/** The first `n` UTF-16 units of `s`, one less when the cut would split a surrogate pair. */
function headUnits(s, n) {
  const code = s.charCodeAt(n - 1);
  return s.slice(0, code >= 0xd800 && code <= 0xdbff ? n - 1 : n);
}

/** Text cell value: null for empty, truncated to Excel's cell limit. */
function txt(v) {
  if (v === null || v === undefined) return null;
  const s = typeof v === 'string' ? v : String(v);
  if (!s) return null;
  return s.length > MAX_CELL_TEXT ? `${headUnits(s, MAX_CELL_TEXT - 1)}…` : s;
}

// XML 1.0 has no U+FFFE / U+FFFF and no unpaired surrogates. ExcelJS drops C0
// control characters itself, but writes these through, and a single one makes
// the whole sheet part malformed (Excel then "repairs" the file and usually
// loses that sheet's cells). Customer names, tags and product titles are free text.
const XML_SUSPECT = /[\uD800-\uDFFF\uFFFE\uFFFF]/;
const XML_NONCHARACTERS = /[\uFFFE\uFFFF]/g;
const LONE_SURROGATE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;
const wellFormed = typeof String.prototype.toWellFormed === 'function' ? (s) => s.toWellFormed() : (s) => s.replace(LONE_SURROGATE, '\uFFFD');

/** `s` safe for XML: unpaired surrogates → U+FFFD, U+FFFE/U+FFFF removed. */
function xmlSafe(s) {
  return XML_SUSPECT.test(s) ? wellFormed(s).replace(XML_NONCHARACTERS, '') : s;
}

/** A cell value safe for XML: strings, and a formula's text and cached result. */
function xmlSafeValue(v) {
  if (typeof v === 'string') return xmlSafe(v);
  if (v !== null && typeof v === 'object' && typeof v.formula === 'string') {
    const formula = xmlSafe(v.formula);
    const result = typeof v.result === 'string' ? xmlSafe(v.result) : v.result;
    return formula === v.formula && result === v.result ? v : { ...v, formula, result };
  }
  return v;
}

/** Numeric cell value, or null when not a finite number. */
function num(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (typeof v === 'string' && v.trim() !== '') {
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

/** Integer cents → dollars for a money-formatted cell. */
function money(cents) {
  return typeof cents === 'number' && Number.isFinite(cents) ? cents / 100 : null;
}

/** a / b as a fraction for '0.0%' cells; null when undefined. */
function ratio(a, b) {
  return Number.isFinite(a) && Number.isFinite(b) && b > 0 ? a / b : null;
}

/** A rate given by another module: a fraction (0.077) or, defensively, a percentage (7.7). */
function rateValue(v) {
  if (typeof v !== 'number' || !Number.isFinite(v)) return null;
  return v > 1 ? v / 100 : v;
}

function fmtCount(n) {
  return typeof n === 'number' && Number.isFinite(n) ? n.toLocaleString('en-US') : '0';
}

/** The object entries of an input array (anything else → []), so a malformed file cannot crash the report. */
function objects(v) {
  return Array.isArray(v) ? v.filter((x) => x !== null && typeof x === 'object') : [];
}

// Excel rejects string literals longer than 255 characters inside a formula.
const MAX_FORMULA_TEXT = 255;

// ---------------------------------------------------------------------------
// Styles. Every cell gets one of a few shared style objects: ExcelJS caches
// the style id per object, so sharing them keeps 1M+ cells fast. Never mutate
// a style object after creation (cells point at them).
// ---------------------------------------------------------------------------

const COLOR = {
  headerFill: 'FF4472C4',
  white: 'FFFFFFFF',
  link: 'FF0563C1',
  grid: 'FFD9D9D9',
  title: 'FF1F4E78',
  warn: 'FFC00000',
  muted: 'FF7F7F7F',
  band: 'FFEEF3FA',
};

const NUMFMT = {
  money: '"$"#,##0.00',
  int: '0',
  count: '#,##0',
  date: 'yyyy-mm-dd',
  datetime: 'yyyy-mm-dd hh:mm',
  pct: '0.0%',
};

const THIN = { style: 'thin', color: { argb: COLOR.grid } };
const GRID = { top: THIN, left: THIN, bottom: THIN, right: THIN };

const KINDS = {
  text: {},
  bold: { font: { bold: true } },
  warn: { font: { color: { argb: COLOR.warn } } },
  link: { font: { color: { argb: COLOR.link }, underline: true } },
  int: { numFmt: NUMFMT.int },
  count: { numFmt: NUMFMT.count },
  money: { numFmt: NUMFMT.money },
  date: { numFmt: NUMFMT.date },
  datetime: { numFmt: NUMFMT.datetime },
  pct: { numFmt: NUMFMT.pct },
};

const fills = new Map();
function solidFill(argb) {
  if (!argb) return null;
  let f = fills.get(argb);
  if (!f) {
    f = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
    fills.set(argb, f);
  }
  return f;
}

const styles = new Map();
/** Shared style for a table cell of `kind`, optionally filled, with grid borders unless `grid` is false. */
function cellStyle(kind, fillArgb = null, grid = true) {
  const key = `${kind}|${fillArgb ?? ''}|${grid ? 1 : 0}`;
  let s = styles.get(key);
  if (!s) {
    const k = KINDS[kind] ?? KINDS.text;
    s = {};
    if (k.numFmt) s.numFmt = k.numFmt;
    if (k.font) s.font = k.font;
    if (grid) s.border = GRID;
    const fill = solidFill(fillArgb);
    if (fill) s.fill = fill;
    styles.set(key, s);
  }
  return s;
}

/** Shared style with a custom number format (no border), e.g. '#,##0" 张"'. */
function formatStyle(numFmt) {
  const key = `fmt|${numFmt}`;
  let s = styles.get(key);
  if (!s) {
    s = { numFmt };
    styles.set(key, s);
  }
  return s;
}

/** Money with a text prefix, negative-safe: "已用 $1.00" / "已用 -$1.00". */
function prefixedMoney(prefix) {
  return formatStyle(`"${prefix}$"#,##0.00;"${prefix}-$"#,##0.00`);
}

const HEADER = {
  font: { bold: true, color: { argb: COLOR.white } },
  fill: solidFill(COLOR.headerFill),
  alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
  border: GRID,
};
const TITLE = { font: { bold: true, size: 14, color: { argb: COLOR.title } } };
const SECTION = { font: { bold: true, size: 12, color: { argb: COLOR.title } } };
const NOTE = { font: { color: { argb: COLOR.muted } } };
const NOTICE = { font: { bold: true, color: { argb: COLOR.warn } } };
const CHANGED = { font: { color: { argb: COLOR.warn } } }; // a setting changed in .env after select
const LABEL = {};

// ---------------------------------------------------------------------------
// Streaming workbook / sheet wrappers
// ---------------------------------------------------------------------------

const yieldToEventLoop = () => new Promise((resolve) => setImmediate(resolve));

function colLetter(n) {
  let s = '';
  let x = n;
  while (x > 0) {
    const m = (x - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

/**
 * ExcelJS 4.4.0's streaming WorksheetWriter.commit() writes <autoFilter>
 * right after </sheetData> and <sheetProtection> several elements later, but
 * the OOXML schema (CT_Worksheet in sml.xsd) requires sheetProtection BEFORE
 * autoFilter. Excel then reports unreadable content and offers a repair that
 * may drop the protection or the filter. For this one sheet, write the
 * protection at the filter's spot, just before it, and nothing at the later
 * spot. (exceljs is pinned to 4.4.0; test/excel.test.js checks the element
 * order of every worksheet part.)
 */
function protectionBeforeFilter(ws) {
  const writeFilter = ws._writeAutoFilter;
  const writeProtection = ws._writeSheetProtection;
  if (typeof writeFilter !== 'function' || typeof writeProtection !== 'function') return;
  ws._writeAutoFilter = function writeProtectionThenFilter() {
    writeProtection.call(this);
    writeFilter.call(this);
  };
  ws._writeSheetProtection = function writtenBeforeTheFilter() {};
}

class Book {
  constructor(file, { generatedAt, warnings, log }) {
    this.warnings = warnings;
    this.log = log;
    this.error = null;
    this.wb = new ExcelJS.stream.xlsx.WorkbookWriter({
      filename: file,
      useStyles: true,
      useSharedStrings: false,
      creator: 'gift-card-promo',
      lastModifiedBy: 'gift-card-promo',
      created: generatedAt,
      modified: generatedAt,
    });
    // ExcelJS only listens for errors at the very end of commit(); an earlier
    // write error (disk full, ...) would otherwise be an uncaught 'error' event,
    // or leave commit() waiting forever for a 'finish' that never comes.
    this.rejectCommit = null;
    const capture = (err) => {
      this.error ??= err;
      this.rejectCommit?.(err);
    };
    this.wb.stream.on('error', capture);
    this.wb.zip.on('error', capture);
  }

  check() {
    if (this.error) throw this.error;
  }

  sheet(name, options) {
    this.check();
    return new Sheet(this, name, options);
  }

  warn(message) {
    this.warnings.push(message);
    this.log.warn(message);
  }

  async commit() {
    this.check();
    await new Promise((resolve, reject) => {
      this.rejectCommit = reject;
      this.wb.commit().then(resolve, reject);
    });
    this.check();
  }

  /** Stop writing and close the temp file after a failure (best effort). */
  discard() {
    try {
      this.wb.zip.abort();
    } catch {
      /* already finished or never started */
    }
    try {
      this.wb.stream.destroy();
    } catch {
      /* already closed */
    }
  }
}

class Sheet {
  /**
   * @param {Book} book
   * @param {string} name
   * @param {{ widths: number[], freezeRows?: number, freezeCols?: number }} options
   */
  constructor(book, name, { widths, freezeRows = 1, freezeCols = 0 }) {
    this.book = book;
    this.name = name;
    const view = { state: 'frozen', ySplit: freezeRows };
    if (freezeCols) view.xSplit = freezeCols;
    // Views must be given here: the streaming writer emits them immediately.
    this.ws = book.wb.addWorksheet(name, { views: [view] });
    this.ws.columns = widths.map((width) => ({ width })); // written with the first row
    this.rowNo = 0; // last row number written
    this.cells = 0; // cells since the last yield
    this.filter = null;
    this.truncated = false;
  }

  /**
   * Append one row. `rowStyles` is one style per value (array) or one style
   * for every value. Styles are applied before the row is committed.
   */
  async add(rawValues, rowStyles = null, { height } = {}) {
    if (this.rowNo >= EXCEL_MAX_ROWS) {
      this.truncated = true;
      return null;
    }
    const values = rawValues.map(xmlSafeValue);
    const row = this.ws.addRow(values);
    let hasFormula = false;
    for (let i = 0; i < values.length; i += 1) {
      const style = Array.isArray(rowStyles) ? rowStyles[i] : rowStyles;
      if (style) row.getCell(i + 1).style = style;
      const v = values[i];
      if (v !== null && typeof v === 'object' && v.formula) hasFormula = true;
    }
    if (height) row.height = height;
    this.rowNo = row.number;
    row.commit();
    // ExcelJS keeps every formula cell in worksheet._formulae for shared-formula
    // lookups and never frees them. We never use shared formulas, so drop the
    // registry once the row is written; otherwise ~60k links cost ~100 MB.
    if (hasFormula && this.ws._formulae) this.ws._formulae = {};
    this.cells += values.length || 1;
    if (this.cells >= YIELD_EVERY_CELLS) {
      this.cells = 0;
      await yieldToEventLoop();
      this.book.check();
    }
    return this.rowNo;
  }

  /** An empty spacer row. */
  async blank() {
    return this.add([]);
  }

  /** A table header row (bold white on blue, wrapped, taller). Returns its row number. */
  async header(titles) {
    return this.add(titles, HEADER, { height: 30 });
  }

  /** Remember the auto-filter range of the sheet's main table (one per sheet in Excel). */
  setFilter(headerRow, columnCount, lastRow = this.rowNo) {
    this.filter = `A${headerRow}:${colLetter(columnCount)}${Math.max(headerRow, lastRow)}`;
  }

  async finish() {
    if (this.filter) this.ws.autoFilter = this.filter;
    // No password: the protection only prevents accidental edits; filtering
    // and resizing columns stay allowed.
    await this.ws.protect('', { selectLockedCells: true, selectUnlockedCells: true, autoFilter: true, formatColumns: true });
    protectionBeforeFilter(this.ws);
    this.ws.commit();
    if (this.truncated) this.book.warn(`“${this.name}”超过 Excel 的行数上限（${EXCEL_MAX_ROWS.toLocaleString('en-US')} 行），后面的行没有写入`);
    this.book.check();
  }
}

// ---------------------------------------------------------------------------
// writeReport
// ---------------------------------------------------------------------------

/**
 * Rebuild the campaign workbook from selection.json + journal.jsonl (+ tags.json,
 * verify.json, usage.json when present). Holds the Excel lock while reading
 * the state and replacing the file, so concurrent calls never interleave.
 *
 * @param {object} a
 * @param {object} a.config loadConfig() result
 * @param {object} a.paths campaignPaths(config)
 * @param {{info: Function, warn: Function, error: Function}} [a.log]
 * @param {() => Date} [a.now] "generated at" time shown in the workbook
 * @param {string|null} [a.out] also save a copy here (a file, or a directory to put it in)
 * @param {object} [a.lockOptions] passed to withExcelLock ({ timeoutMs, pollMs, sleep }); for tests
 * @returns {Promise<{ file: string, out: string|null, warnings: string[] }>}
 */
export async function writeReport({ config, paths, log = console, now = () => new Date(), out = null, lockOptions = {} } = {}) {
  if (!config || !paths) throw new Error('writeReport needs config and paths');
  return withExcelLock(paths, () => buildAndPublish({ config, paths, log, now, out }), lockOptions);
}

async function buildAndPublish({ config, paths, log, now, out }) {
  // Every warning is logged here, once, and also returned (callers must not print them again).
  const warnings = [];
  const warn = (message) => {
    warnings.push(message);
    log.warn(message);
  };

  // We hold the Excel lock, so no other writer of this workbook is running:
  // any temp workbook still in the folder was left by a killed write.
  removeLeftoverTemps(paths.excel);

  const selection = readJson(paths.selection, null);
  if (!selection) throw new Error(NO_SELECTION_MESSAGE);
  const entries = readJournal(paths.journal);
  const journal = foldJournal(entries);
  const tags = readOptionalJson(paths.tags, 'tags.json', warn);
  const verify = readOptionalJson(paths.verify, 'verify.json', warn);
  const usage = readOptionalJson(paths.usage, 'usage.json', warn);
  const running = runningCommand(paths);
  const generatedAt = toDate(now());

  const ctx = buildContext({ config, selection, entries, journal, tags, verify, usage, running, generatedAt, warn });

  ensureDir(path.dirname(paths.excel));
  const tmp = tempPathFor(paths.excel);
  const book = new Book(tmp, { generatedAt, warnings, log });
  try {
    await writeSheets(book, ctx);
    await book.commit();
    replaceFile(tmp, paths.excel);
  } catch (err) {
    book.discard();
    removeQuietly(tmp);
    throw err;
  }

  // The main workbook is in place now: a failed extra copy must not look like a failed export.
  let outFile = null;
  if (out) {
    try {
      outFile = copyOut(paths.excel, out);
    } catch (err) {
      warn(err.message);
    }
  }

  for (const file of new Set([paths.excel, outFile].filter(Boolean))) {
    if (openInExcel(file)) warn(file === paths.excel ? OPEN_IN_EXCEL_WARNING : `${OPEN_IN_EXCEL_WARNING}：${file}`);
  }
  return { file: paths.excel, out: outFile, warnings };
}

/**
 * Where a workbook is written before it is renamed over `target`: a hidden
 * file in the same folder, without the .xlsx suffix, so a leftover of a killed
 * write is neither mistaken for a report nor opened by Excel.
 */
function tempPathFor(target) {
  return path.join(path.dirname(target), `.${path.basename(target)}.tmp-${process.pid}-${Date.now()}`);
}

const escapeRegExp = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * Delete temp workbooks of `target` left by writes that were killed (second
 * Ctrl+C, SIGKILL, closed terminal): the current ".<name>.tmp-<pid>-<ms>" and
 * the earlier "<name>.tmp-<pid>-<ms>.xlsx". Only call it while holding the
 * Excel lock, when no other write of `target` can be in progress.
 */
function removeLeftoverTemps(target) {
  const dir = path.dirname(target);
  const base = escapeRegExp(path.basename(target));
  const patterns = [new RegExp(`^\\.${base}\\.tmp-.+$`), new RegExp(`^${base}\\.tmp-.+\\.xlsx$`)];
  let names;
  try {
    names = fs.readdirSync(dir);
  } catch {
    return; // no folder yet
  }
  for (const name of names) {
    if (patterns.some((p) => p.test(name))) removeQuietly(path.join(dir, name));
  }
}

function toDate(v) {
  const ms = toMs(v);
  return ms === null ? new Date() : new Date(ms);
}

/** An optional input file: missing → null; unreadable → null plus a warning (the report still gets written). */
function readOptionalJson(file, label, warn) {
  try {
    const value = readJson(file, null);
    return value && typeof value === 'object' ? value : null;
  } catch (err) {
    warn(`${label} 无法读取，本次 Excel 不包含它的内容：${err.message}`);
    return null;
  }
}

function replaceFile(tmp, target) {
  try {
    fs.renameSync(tmp, target);
  } catch (err) {
    if (['EBUSY', 'EPERM', 'EACCES'].includes(err.code)) {
      throw new Error(`无法替换 ${target}（${err.code}）：文件可能被 Excel 锁住，请关闭后重新运行 export`);
    }
    throw err;
  }
}

function removeQuietly(file) {
  try {
    fs.rmSync(file, { force: true });
  } catch {
    /* nothing to clean up */
  }
}

/** Excel (and Word-style long names) create "~$<name>" next to a file while it is open. */
function openInExcel(file) {
  const dir = path.dirname(file);
  const base = path.basename(file);
  return [`~$${base}`, `~$${base.slice(2)}`].some((name) => fs.existsSync(path.join(dir, name)));
}

/** Resolve --out: "~" expansion, relative to cwd, a directory (existing or ending in "/") gets the file name. */
function resolveOut(out, excelFile) {
  let p = String(out).trim();
  const isDirHint = /[\\/]$/.test(p);
  if (p === '~') p = os.homedir();
  else if (p.startsWith('~/')) p = path.join(os.homedir(), p.slice(2));
  p = path.resolve(p);
  let isDir = isDirHint;
  try {
    isDir = isDir || fs.statSync(p).isDirectory();
  } catch {
    /* does not exist yet */
  }
  return isDir ? path.join(p, path.basename(excelFile)) : p;
}

/**
 * Copy the finished workbook to `out` (via a temp file + rename, so a reader
 * never sees half a copy). Throws `另存 Excel 到 … 失败：…` on any failure.
 */
function copyOut(excelFile, out) {
  const target = resolveOut(out, excelFile);
  if (target === path.resolve(excelFile)) return target;
  const tmp = tempPathFor(target);
  try {
    ensureDir(path.dirname(target));
    fs.copyFileSync(excelFile, tmp);
    fs.renameSync(tmp, target);
  } catch (err) {
    removeQuietly(tmp);
    throw new Error(`另存 Excel 到 ${target} 失败：${err.message}`);
  }
  return target;
}

// ---------------------------------------------------------------------------
// Context: everything the sheets need, computed once
// ---------------------------------------------------------------------------

function buildContext({ config, selection, entries, journal, tags, verify, usage, running, generatedAt, warn }) {
  const params = selection.params ?? {};
  let timeZone = params.timezone || config.timezone || DEFAULT_TIMEZONE;
  if (!validTimeZone(timeZone)) {
    warn(`时区 "${timeZone}" 无效，Excel 里的时间改用 ${DEFAULT_TIMEZONE}`);
    timeZone = DEFAULT_TIMEZONE;
  }
  const clock = clockFor(timeZone);
  const shop = config.shop;
  const sentTag = params.sentTag || config.sentTag;
  const tiersCents = Array.isArray(params.giftTiersCents) ? params.giftTiersCents : config.giftTiersCents ?? [];
  const isTest = selection.mode === 'test';
  const exportedDay = clock.dayNumber(selection.snapshot?.exportedAt ?? selection.createdAt);

  // tags.json only counts when it is about this campaign's tag.
  let taggedIds = new Set();
  let tagsFetchedMs = null; // when the tag snapshot was taken; null = no usable snapshot
  if (tags) {
    if (String(tags.tag ?? '').toLowerCase() !== String(sentTag).toLowerCase()) {
      warn(`tags.json 记录的是 tag "${tags.tag}"，不是本活动的 "${sentTag}"，已忽略`);
    } else if (!Array.isArray(tags.ids)) {
      warn('tags.json 里没有客户列表（ids），已忽略');
    } else {
      taggedIds = new Set(tags.ids);
      tagsFetchedMs = toMs(tags.fetchedAt);
    }
  }
  /**
   * "是否有 <SENT_TAG>": a tag snapshot taken well after the journal's tag.ok
   * (or with no tag.ok at all) is the truth, so a tag removed in the admin shows
   * "否" after export --refresh / verify; otherwise the journal or the snapshot.
   * "Well after" = TAG_INDEX_GRACE_MS: Shopify's customer search can take a
   * while to show a tag that was just added, and a fresh tag must not read "否".
   */
  const hasTag = (state, customerId) => {
    const inSnapshot = taggedIds.has(customerId);
    const taggedMs = toMs(state?.taggedAt);
    if (tagsFetchedMs !== null && (taggedMs === null || tagsFetchedMs > taggedMs + TAG_INDEX_GRACE_MS)) return inSnapshot;
    return !!state?.taggedAt || inSnapshot;
  };

  // usage.json indexes
  const cards = objects(usage?.cards);
  const payments = objects(usage?.payments);
  const cardsById = new Map(cards.map((c) => [c.giftCardId, c]));
  const cardsByCustomer = new Map();
  for (const c of cards) {
    if (!c.customerId) continue;
    if (!cardsByCustomer.has(c.customerId)) cardsByCustomer.set(c.customerId, []);
    cardsByCustomer.get(c.customerId).push(c);
  }
  const ordersById = new Map(objects(usage?.orders).map((o) => [o.orderId, o]));
  const orderNamesByCard = new Map();
  for (const p of payments) {
    if (!p.giftCardId) continue;
    if (!orderNamesByCard.has(p.giftCardId)) orderNamesByCard.set(p.giftCardId, new Set());
    orderNamesByCard.get(p.giftCardId).add(p.orderName || ordersById.get(p.orderId)?.orderName || numericId(p.orderId));
  }

  // Campaign dates. LAUNCH_DATE / REMIND_n_DATE are shown as the commands
  // check them now (.env); the expiry as frozen in the list, because every card
  // is created with that one. A difference to the list gets a note.
  const dateSetting = (current, frozen) => {
    const now = current || '';
    const then = frozen ?? null; // null: a list written without this parameter, nothing to compare
    const changed = then !== null && (then || '') !== now;
    return { current: now, frozen: then, changed, note: changed ? `生成名单时为 ${then || '未设置'}，现在按 .env 为 ${now || '未设置'}` : null };
  };
  // issue creates every card with the list's expiry ('' = none), whatever .env says now.
  const expiryFrozen = params.giftCardExpiresOn || '';
  const dates = {
    launch: dateSetting(config.launchDate, params.launchDate),
    remind1: dateSetting(config.remind1Date, params.remind1Date),
    remind2: dateSetting(config.remind2Date, params.remind2Date),
    expiry: { frozen: expiryFrozen, env: config.giftCardExpiresOn || '', changed: expiryFrozen !== (config.giftCardExpiresOn || '') },
  };

  const duplicates = objects(selection.duplicates);
  const groupsById = new Map(duplicates.map((g) => [g.groupId, g]));

  const recipients = objects(selection.recipients).sort((a, b) => (a.seq ?? 0) - (b.seq ?? 0));
  const rows = recipients.map((r) => {
    const state = journal.customers.get(r.customerId) ?? null;
    const status = effectiveStatus(state?.status, running);
    const cardCents = state?.giftCardId && Number.isInteger(state.amountCents) ? state.amountCents : r.amountCents;
    let card = null;
    if (usage) {
      if (state?.giftCardId) card = cardsById.get(state.giftCardId) ?? null;
      else if (cardsByCustomer.get(r.customerId)?.length === 1) card = cardsByCustomer.get(r.customerId)[0];
    }
    return { r, state, status, cardCents, card };
  });

  const ctx = {
    config,
    selection,
    params,
    entries,
    journal,
    verify,
    usage,
    running,
    generatedAt,
    timeZone,
    tzLabel: timeZoneLabel(timeZone),
    shop,
    sentTag,
    tiersCents,
    isTest,
    campaignId: selection.campaignId || config.campaignId,
    inactiveMonths: params.inactiveMonths ?? config.inactiveMonths,
    dates,
    hasTag,
    cardsById,
    ordersById,
    orderNamesByCard,
    groupsById,
    rows,
    runsById: new Map(journal.runs.map((run) => [run.run, run])),
    notSelected: objects(selection.notSelected),
    duplicates,
    // conversions
    time: clock.time,
    day: clock.day,
    stamp: clock.stamp,
    ymd,
    /** A HYPERLINK() cell to the admin page (plain text when no link can be built). */
    link(type, gid, text) {
      const id = numericId(gid);
      const url = adminUrl(shop, type, gid);
      let label = txt(text) ?? (id || txt(gid));
      if (!url || !label) return label;
      if (label.length > MAX_FORMULA_TEXT) label = `${headUnits(label, MAX_FORMULA_TEXT - 1)}…`;
      return { formula: `HYPERLINK("${url}","${label.replace(/"/g, '""')}")`, result: label };
    },
    daysToExport(iso) {
      const d = clock.dayNumber(iso);
      return d === null || exportedDay === null ? null : exportedDay - d;
    },
    tierText(r) {
      if (r.kind === 'test') return '测试';
      return Number.isInteger(r.tier) ? txt(tierLabel(r.tier, tiersCents)) : null;
    },
    remindText(state, round) {
      return txt(remindLabel(state?.reminders?.[round], running, timeZone));
    },
    /** This campaign's tag for reminder `round` (1 / '1' / 2 / '2'), e.g. OCT26RTPROMO-R1; null for any other round. */
    roundTag(round) {
      const n = Number(round);
      return n === 1 || n === 2 ? roundTagName(sentTag, n) : null;
    },
    remindWarn(state, round) {
      const st = state?.reminders?.[round]?.status;
      if (!st) return false;
      const eff = effectiveRemindStatus(st, running);
      return eff === 'failed' || eff === 'unknown';
    },
    ordersText(card) {
      const names = card ? orderNamesByCard.get(card.giftCardId) : null;
      return names?.size ? txt([...names].filter(Boolean).join(', ')) : null;
    },
  };
  ctx.notesFor = (x) => recipientNotes(x, ctx);
  ctx.progress = computeProgress(rows);
  ctx.reminders = computeReminders(rows, running);
  ctx.usageSummary = usage ? normalizeUsageSummary(usage, cards) : null;
  return ctx;
}

/** "备注/错误": why a row is special, joined with "；". */
function recipientNotes({ r, state, status }, ctx) {
  const notes = [];
  if (state) {
    if (status === 'skipped' && state.skipReason) notes.push(issueSkipText(state.skipReason));
    if (state.status === 'in_progress' && status !== 'in_progress') {
      notes.push('建卡请求发出后没有记录结果（运行被中断）；下次运行 issue 会先去 Shopify 查找这张卡');
    }
    if (state.error) notes.push(errorText(state.error));
    if (state.reconciledFrom) notes.push(`卡由${labelOf(RECONCILE_SOURCE_LABELS, state.reconciledFrom)}在 Shopify 找到后补记`);
    if (state.giftCardId && Number.isInteger(state.amountCents) && Number.isInteger(r.amountCents) && state.amountCents !== r.amountCents) {
      notes.push(`卡的金额 ${formatUsd(state.amountCents)} 与名单金额 ${formatUsd(r.amountCents)} 不同`);
    }
  }
  // basisWhy is also set when no earlier paid order was found (then neverReason
  // explains it), so the "earlier order" note only applies with a basis order.
  if (r.basis && BASIS_NOTES[r.basisWhy]) notes.push(BASIS_NOTES[r.basisWhy]);
  if (NEVER_NOTES[r.neverReason]) notes.push(NEVER_NOTES[r.neverReason]);
  const group = r.groupId ? ctx.groupsById.get(r.groupId) : null;
  if (group?.flaggedBulk) notes.push(`同地址 ${group.size} 个账户，疑似批量注册`);
  return txt(notes.join('；'));
}

function computeProgress(rows) {
  const byStatus = new Map(STATUS_ORDER.map((s) => [s, { count: 0, cents: 0 }]));
  const skipReasons = new Map();
  let nextPendingSeq = null;
  for (const row of rows) {
    if (!byStatus.has(row.status)) byStatus.set(row.status, { count: 0, cents: 0 });
    const b = byStatus.get(row.status);
    b.count += 1;
    b.cents += Number.isFinite(row.cardCents) ? row.cardCents : 0;
    if (row.status === 'skipped') {
      const reason = row.state?.skipReason ?? '';
      skipReasons.set(reason, (skipReasons.get(reason) ?? 0) + 1);
    }
    if (row.status === 'pending' && nextPendingSeq === null) nextPendingSeq = row.r.seq ?? null;
  }
  let count = 0;
  let cents = 0;
  for (const b of byStatus.values()) {
    count += b.count;
    cents += b.cents;
  }
  return { byStatus, total: { count, cents }, skipReasons, nextPendingSeq };
}

function computeReminders(rows, running) {
  const result = {};
  for (const round of ['1', '2']) {
    const c = { sent: 0, skipped: 0, failed: 0, unknown: 0, inProgress: 0, notYet: 0, reasons: new Map() };
    for (const row of rows) {
      const rs = row.state?.reminders?.[round];
      if (!rs || !rs.status) {
        if (row.state?.giftCardId) c.notYet += 1;
        continue;
      }
      switch (effectiveRemindStatus(rs.status, running)) {
        case 'sent':
          c.sent += 1;
          break;
        case 'skipped':
          c.skipped += 1;
          c.reasons.set(rs.reason ?? '', (c.reasons.get(rs.reason ?? '') ?? 0) + 1);
          break;
        case 'failed':
          c.failed += 1;
          break;
        case 'in_progress':
          c.inProgress += 1;
          break;
        default:
          c.unknown += 1;
      }
    }
    result[round] = c;
  }
  return result;
}

/** usage.json summary with rates recomputed from the counts (robust to fraction vs percent). */
function normalizeUsageSummary(usage, cards) {
  const s = usage.summary && typeof usage.summary === 'object' ? usage.summary : {};
  const n = (v, fallback = null) => (typeof v === 'number' && Number.isFinite(v) ? v : fallback);
  const issuedCards = n(s.issuedCards, cards.length);
  const issuedCents = n(s.issuedCents, cards.reduce((a, c) => a + (n(c.initialCents) ?? 0), 0));
  const usedCards = n(s.usedCards, cards.filter((c) => n(c.balanceCents) !== null && c.balanceCents < c.initialCents).length);
  const usedCents = n(s.usedCents, cards.reduce((a, c) => a + (n(c.usedCents) ?? 0), 0));
  const orders = n(s.orders, objects(usage.orders).length);
  const ordersTotalCents = n(s.ordersTotalCents);
  const giftCardCents = n(s.giftCardCents);
  return {
    issuedCards,
    issuedCents,
    usedCards,
    usedCents,
    usedRate: ratio(usedCards, issuedCards) ?? rateValue(s.usedRate),
    usedCentsRate: ratio(usedCents, issuedCents) ?? rateValue(s.usedCentsRate),
    orders,
    ordersTotalCents,
    avgOrderCents: n(s.avgOrderCents) ?? (ordersTotalCents !== null && orders > 0 ? ordersTotalCents / orders : null),
    giftCardCents,
    customerPaidCents: n(s.customerPaidCents) ?? (ordersTotalCents !== null && giftCardCents !== null ? ordersTotalCents - giftCardCents : null),
    byTier: objects(s.byTier),
    byKind: objects(s.byKind),
    daily: objects(s.daily),
    topProducts: objects(s.topProducts),
  };
}

// ---------------------------------------------------------------------------
// Sheets, in workbook order
// ---------------------------------------------------------------------------

async function writeSheets(book, ctx) {
  if (ctx.usage) await writeUsageReportSheet(book, ctx);
  await writeSummarySheet(book, ctx);
  await writeRecipientsSheet(book, ctx);
  await writeNotSelectedSheet(book, ctx);
  await writeDuplicatesSheet(book, ctx);
  if (ctx.usage) await writeUsageDetailSheet(book, ctx);
  if (ctx.verify) await writeVerifySheet(book, ctx);
  await writeJournalSheet(book, ctx);
  await writeHelpSheet(book, ctx);
}

// ---- 使用报告 ---------------------------------------------------------------

async function writeUsageReportSheet(book, ctx) {
  const u = ctx.usageSummary;
  const asOf = ctx.stamp(ctx.usage.fetchedAt) || '未知时间';
  const s = book.sheet('使用报告', { widths: [36, 16, 28, 28, 16, 14] });
  const c = (kind) => cellStyle(kind);
  const label = cellStyle('bold', null, false);

  await s.add([`使用报告（截至 ${asOf}，${ctx.tzLabel}）`], TITLE, { height: 24 });
  await s.blank();
  // Totals, laid out like the plan's sample; the units live in the number formats
  // so every figure stays a real number.
  await s.add(['发出礼品卡', u.issuedCards, money(u.issuedCents)], [label, formatStyle('#,##0" 张"'), prefixedMoney('面额 ')]);
  await s.add(['已经使用', u.usedCards, u.usedRate, money(u.usedCents), u.usedCentsRate], [
    label,
    formatStyle('#,##0" 张"'),
    formatStyle('"占 "0.0%'),
    prefixedMoney('已用 '),
    formatStyle('"占面额 "0.0%'),
  ]);
  await s.add(['带来订单', u.orders, money(u.ordersTotalCents), money(u.avgOrderCents)], [label, formatStyle('#,##0" 笔"'), prefixedMoney('订单总额 '), prefixedMoney('平均每单 ')]);
  await s.add([null, null, money(u.giftCardCents), money(u.customerPaidCents)], [null, null, prefixedMoney('其中礼品卡抵扣 '), prefixedMoney('顾客另外支付 ')]);
  // 礼品卡抵扣 (at checkout) and 已用 (card balances) differ after a refund back to a card.
  await s.add([USAGE_SPLIT_NOTE], NOTE);

  await s.blank();
  await s.header(['按档位', '发出', '已使用', '使用率', '已用金额']);
  for (const t of u.byTier) {
    const name = t.label || (Number.isInteger(t.tier) ? tierLabel(t.tier, ctx.tiersCents) : '') || String(t.tier ?? '');
    await s.add([txt(name), num(t.issued), num(t.used), ratio(t.used, t.issued) ?? rateValue(t.rate), money(t.usedCents)], [c('text'), c('count'), c('count'), c('pct'), c('money')]);
  }

  await s.blank();
  await s.header(['按客户类型', '发出', '已使用', '使用率', '已用金额']);
  for (const k of u.byKind) {
    await s.add([labelOf(KIND_LABELS, k.kind) || null, num(k.issued), num(k.used), ratio(k.used, k.issued) ?? rateValue(k.rate), money(k.usedCents)], [c('text'), c('count'), c('count'), c('pct'), c('money')]);
  }

  await s.blank();
  await s.header(['每日', '当天新用的卡', '当天订单', '当天订单金额', '累计使用的卡', '累计使用率']);
  for (const d of u.daily) {
    const cumulative = num(d.cumulativeCardsUsed);
    await s.add([ymd(d.date) ?? txt(d.date), num(d.newCardsUsed), num(d.orders), money(d.ordersTotalCents), cumulative, ratio(cumulative, u.issuedCards) ?? rateValue(d.cumulativeRate)], [
      c('date'),
      c('count'),
      c('count'),
      c('money'),
      c('count'),
      c('pct'),
    ]);
  }

  await s.blank();
  await s.header(['卖得最多的 10 个商品', '数量', '金额']);
  for (const p of u.topProducts.slice(0, 10)) {
    await s.add([txt(p.name), num(p.quantity), money(p.amountCents)], [c('text'), c('count'), c('money')]);
  }

  await s.blank();
  await s.add([`截至 ${asOf}（${ctx.tzLabel}）。每天运行一次 usage 更新本表；卡的余额小于原金额就算已使用。`], NOTE);
  await s.finish();
}

// ---- 汇总 -------------------------------------------------------------------

async function writeSummarySheet(book, ctx) {
  const { selection, params } = ctx;
  const stats = selection.stats ?? {};
  const funnel = selection.funnel ?? null;
  const snapshot = selection.snapshot ?? {};
  // The 9th column is only used by the 提醒 table (本轮 tag).
  const s = book.sheet('汇总', { widths: [40, 18, 14, 14, 14, 12, 18, 70, 24] });
  const plain = (kind) => cellStyle(kind, null, false);
  const kv = (label, value, kind = 'text', note = null, noteStyle = NOTE) => s.add([label, value ?? null, note ?? null], [LABEL, plain(kind), noteStyle]);
  const section = async (title) => {
    await s.blank();
    await s.add([title], SECTION, { height: 22 });
  };
  const trow = (values, kinds) => s.add(values, kinds.map((k) => (typeof k === 'string' ? cellStyle(k) : k)));
  const breakdown = async (title, counts, keyLabel = (k) => k) => {
    const items = Object.entries(counts ?? {}).filter(([, n]) => n > 0).sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])));
    if (!items.length) return;
    await s.blank();
    await s.header([title, '人数']);
    for (const [k, n] of items) await trow([txt(keyLabel(k)) ?? '（空）', n], ['text', 'count']);
  };

  await s.add(['本文件由程序生成，修改无效，每次运行会覆盖'], NOTICE, { height: 20 });
  await s.add([`活动 ${ctx.campaignId} · ${ctx.isTest ? '测试活动' : '正式活动'} · 生成于 ${ctx.stamp(ctx.generatedAt)}（${ctx.tzLabel}）`], NOTE);

  // ---- campaign parameters
  await section('活动参数');
  await kv('活动 ID（CAMPAIGN_ID）', ctx.campaignId);
  await kv('名单类型', ctx.isTest ? '测试活动：只取 TEST_CUSTOMER_IDS 里的客户，不套用筛选规则，金额固定' : '正式活动：按筛选规则选出');
  await kv('名单生成时间', ctx.time(selection.createdAt), 'datetime');
  await kv('客户数据导出时间', ctx.time(snapshot.exportedAt), 'datetime', `${labelOf(SNAPSHOT_SOURCE_LABELS, snapshot.source) || '来源未知'}，共 ${fmtCount(snapshot.count)} 个客户`);
  if (ctx.isTest) {
    await kv('测试固定金额', money(params.testGiftAmountCents ?? ctx.rows[0]?.r.amountCents), 'money');
  } else {
    await kv(`近 ${ctx.inactiveMonths} 个月下单的起算日（cutoff）`, ctx.ymd(params.cutoffDate), 'date', '这一天 0 点（店铺时间）以后有有效订单的人不发');
    await kv('注册满', `${params.minAccountAgeDays ?? ctx.config.minAccountAgeDays} 天`);
    await kv('要求已订阅营销邮件', (params.requireEmailSubscribed ?? ctx.config.requireEmailSubscribed) ? '是' : '否');
    await kv('金额比例', `${params.giftPercent ?? ctx.config.giftPercent}%`, 'text', '基数 = 上次有效订单总额 × 比例，四舍五入到分，再对档位');
    await kv('档位（按基数）', ctx.tiersCents.length ? describeTiers(ctx.tiersCents).join('；') : null);
    await kv('有下单入选者上次订单的平均值', money(selection.averageCents), 'money', '从没下单的人用“平均值 × 比例”对档位');
    await kv('从没下单的人发放金额', money(selection.neverAmountCents), 'money');
  }
  // The expiry is frozen in the list (every card is created with it); the
  // launch and reminder dates are the current .env values the commands check.
  const { expiry, launch, remind1, remind2 } = ctx.dates;
  await kv(
    '礼品卡到期日',
    expiry.frozen ? ctx.ymd(expiry.frozen) : '不设到期日',
    'date',
    expiry.changed ? `到期日当天仍可使用。注意：.env 的 GIFT_CARD_EXPIRES_ON 现在是 ${expiry.env || '未设置'}，但建卡仍用生成名单时的 ${expiry.frozen || '不设到期日'}` : '到期日当天仍可使用',
    expiry.changed ? CHANGED : NOTE,
  );
  for (const [label, date] of [['首封邮件日期（正式建卡）', launch], ['第一次提醒', remind1], ['第二次提醒', remind2]]) {
    await kv(label, ctx.ymd(date.current) ?? '未设置', 'date', date.changed ? date.note : null, CHANGED);
  }
  await kv('发放 tag（SENT_TAG）', ctx.sentTag);
  if (!ctx.isTest) {
    const list = (v, f = String) => (Array.isArray(v) ? txt(v.map(f).join(', ')) : txt(v));
    await kv('排除的 tag（不分大小写）', list(params.excludeTags));
    await kv('排除的邮箱域名', list(params.excludeEmailDomains));
    await kv('排除的订单渠道', list(params.excludeOrderSources, sourceText));
  }
  await kv('礼品卡币种', params.currency ?? ctx.config.giftCardCurrency);
  await kv('礼品卡内部备注（note）', campaignNote(params.giftCardNote ?? ctx.config.giftCardNote, ctx.campaignId));
  await kv('邮件和礼品卡页面模板后缀', txt(params.giftCardTemplateSuffix ?? ctx.config.giftCardTemplateSuffix));

  // ---- funnel
  if (funnel) {
    await section('筛选漏斗');
    await s.header(['条件（按顺序判断，第一个不满足的记为排除原因）', '排除人数', '“未入选”表']);
    await trow(['全店客户', num(funnel.total), null], ['bold', 'count', 'text']);
    for (const r of objects(funnel.byRule)) {
      await trow([`${r.n}. ${r.label}`, num(r.count), UNLISTED_RULES.has(r.code) ? '只计数' : '逐行列出'], ['text', 'count', 'text']);
    }
    await trow(['最终入选', num(funnel.recipients), null], ['bold', 'count', 'text']);
    await s.add([`“未入选”表逐行列出 ${fmtCount(funnel.listed)} 人；只因第 1–3 条（没邮箱、平台中转或占位邮箱、未订阅营销邮件）被排除的 ${fmtCount(funnel.unlisted)} 人只计数，不逐行列出。`], NOTE);
    await breakdown('第 2 条：按邮箱域名', funnel.relayDomains);
    await breakdown('第 3 条：按营销状态', funnel.notSubscribed, (k) => labelOf(MARKETING_LABELS, k));
    await breakdown('第 5 条：按 tag', funnel.tags);
    await breakdown('第 9 条：按渠道', funnel.channels);
  }

  // ---- same-address dedupe
  const d = stats.duplicates;
  if (d && !ctx.isTest) {
    await section('同地址去重');
    await kv('重复组数', num(d.groups), 'count');
    await kv('涉及的候选账户', num(d.accounts), 'count');
    await kv('少发人数（每组只保留一个账户）', num(d.removed), 'count');
    await kv('疑似批量注册（一个地址 10 个以上账户）', num(d.flaggedBulk), 'count', '仍按规则每组发一张，详见“同地址重复”表');
    if (Array.isArray(d.largest) && d.largest.length) await kv('最大的几组（账户数）', d.largest.join('、'));
    await kv('没有可比对地址、不参与去重的入选者', num(stats.noAddressRecipients), 'count');
  }

  // ---- amounts
  await section('金额');
  await s.header(['档位', '有下单', '从没下单', '人数', '合计金额']);
  if (ctx.isTest) {
    await trow([`测试固定金额 ${formatUsd(params.testGiftAmountCents ?? ctx.rows[0]?.r.amountCents)}`, null, null, num(stats.recipients ?? ctx.rows.length), money(stats.totalCents)], ['text', 'count', 'count', 'count', 'money']);
  } else {
    for (const t of objects(stats.tiers)) {
      await trow([txt(t.label ?? tierLabel(t.tier, ctx.tiersCents)), num(t.ordered), num(t.never), num(t.count), money(t.cents)], ['text', 'count', 'count', 'count', 'money']);
    }
    await trow(['合计', num(stats.orderedCount), num(stats.neverCount), num(stats.recipients), money(stats.totalCents)], ['bold', 'count', 'count', 'count', 'money']);
    const bf = stats.basisFromEarlier ?? {};
    const count = (v) => (typeof v === 'number' && Number.isFinite(v) ? v : 0);
    const why = [`最近一笔已取消 ${fmtCount(count(bf.lastCancelled))} 人`];
    if (count(bf.lastTest) > 0) why.push(`最近一笔是测试订单 ${fmtCount(bf.lastTest)} 人`); // older lists have no lastTest
    why.push(`最近一笔是 $0 ${fmtCount(count(bf.lastZero))} 人`);
    await kv('有下单的人合计', money(stats.orderedCents), 'money', `${fmtCount(stats.orderedCount)} 人`);
    await kv('从没下单的人合计', money(stats.neverCents), 'money', `${fmtCount(stats.neverCount)} 人`);
    await kv('按更早的付费订单计算', `${fmtCount(count(bf.lastCancelled) + count(bf.lastTest) + count(bf.lastZero))} 人`, 'text', why.join('，'));
    await kv('订单都已取消或都是 $0，按从没下单处理', `${fmtCount(stats.neverReasons?.onlyCancelledOrZero)} 人`);
    await kv('有下单的人基数的中位数', money(stats.medianOrderedRawCents), 'money');
  }

  // ---- issuing progress
  const p = ctx.progress;
  await section('发放进度');
  await s.header(['状态', '人数', '金额']);
  for (const [status, b] of p.byStatus) {
    await s.add([STATUS_LABELS[status] ?? status, b.count, money(b.cents)], [cellStyle('text', STATUS_FILLS[status] ?? null), cellStyle('count'), cellStyle('money')]);
  }
  await trow(['合计', p.total.count, money(p.total.cents)], ['bold', 'count', 'money']);
  // Only batches that created cards count (an issue --repair-only or date-blocked run creates none).
  await kv('已发放到第几批', ctx.journal.lastIssuedBatch ?? ctx.journal.lastBatch ?? null, 'int');
  await kv('下一个待发放的序号', p.nextPendingSeq ?? '没有待发放的人', 'int');
  if (p.skipReasons.size) {
    await s.blank();
    await s.header(['发放前跳过的原因', '人数']);
    for (const [reason, n] of [...p.skipReasons].sort((a, b) => b[1] - a[1])) await trow([issueSkipText(reason) || '原因未记录', n], ['text', 'count']);
  }
  await s.add(['金额：已建卡的按卡的金额，其余按名单金额。“进行中”只在 issue 运行时出现；issue 不在运行时，这些人算作“需人工核对”。'], NOTE);

  // ---- reminders
  const rem = ctx.reminders;
  await section('提醒');
  await s.header(['轮次', '提醒日期', '已发', '跳过', '失败', '结果不明', '进行中', '还没处理（已建卡的人）', '本轮 tag']);
  for (const round of ['1', '2']) {
    const c = rem[round];
    const date = round === '1' ? remind1 : remind2;
    await trow([`第 ${round} 次提醒`, ctx.ymd(date.current) ?? '未设置', c.sent, c.skipped, c.failed, c.unknown, c.inProgress, c.notYet, ctx.roundTag(round)], [
      'text',
      'date',
      'count',
      'count',
      'count',
      'count',
      'count',
      'count',
      'text',
    ]);
  }
  const reasons = [...new Set([...rem['1'].reasons.keys(), ...rem['2'].reasons.keys()])];
  if (reasons.length) {
    await s.blank();
    await s.header(['提醒跳过的原因', '第 1 次', '第 2 次']);
    for (const reason of reasons) {
      await trow([remindSkipText(reason) || '原因未记录', rem['1'].reasons.get(reason) ?? 0, rem['2'].reasons.get(reason) ?? 0], ['text', 'count', 'count']);
    }
  }
  for (const [round, date] of [['1', remind1], ['2', remind2]]) {
    if (date.changed) await s.add([`第 ${round} 次提醒日期：${date.note}（remind 按现在的日期判断哪天能发）`], CHANGED);
  }
  const audience = ctx.isTest
    ? '已建卡、卡还没用过（余额等于原金额）、没停用没过期的人（测试活动不看营销订阅状态）'
    : '已建卡、卡还没用过（余额等于原金额）、没停用没过期、仍订阅营销邮件的人';
  await s.add([
    `提醒只发给${audience}；每一轮每人最多一封：发出后给客户打上本轮 tag，带本轮 tag 的人这一轮不会再发，即使本地日志丢失。跳过的人再次运行同一轮会重新判断。`,
  ], NOTE);

  // ---- usage
  if (ctx.usageSummary) {
    const u = ctx.usageSummary;
    await section('使用情况');
    await kv('截至', ctx.time(ctx.usage.fetchedAt), 'datetime');
    await kv('发出礼品卡（张）', u.issuedCards, 'count');
    await kv('发出面额', money(u.issuedCents), 'money');
    await kv('已经使用（张）', u.usedCards, 'count', u.usedRate, formatStyle('"占 "0.0%'));
    await kv('已用金额', money(u.usedCents), 'money', u.usedCentsRate, formatStyle('"占面额 "0.0%'));
    await kv('带来订单（笔）', u.orders, 'count');
    await kv('订单总额', money(u.ordersTotalCents), 'money');
    await kv('平均每单', money(u.avgOrderCents), 'money');
    await kv('礼品卡抵扣', money(u.giftCardCents), 'money', USAGE_SPLIT_NOTE);
    await kv('顾客另外支付', money(u.customerPaidCents), 'money');
    if (u.topProducts.length) await kv('卖得最多的商品', txt(u.topProducts.slice(0, 3).map((x) => `${x.name} × ${x.quantity}`).join('；')));
    await kv('每日趋势', '见“使用报告”表');
  }

  // ---- run history
  await section('运行记录');
  await s.header(['开始时间', '命令', '预演/实际', '批次/轮次', '数量上限', '退出码', '结束时间', '结果摘要']);
  const runs = [...ctx.journal.runs].sort((a, b) => String(a.startedAt).localeCompare(String(b.startedAt)));
  const lastOpen = new Map(); // command → its newest run without run.end
  for (const run of runs) if (!run.endedAt) lastOpen.set(run.command, run);
  for (const run of runs) {
    let summary = summaryText(run.summary, { command: run.command, dryRun: run.dryRun });
    if (!run.endedAt) summary = ctx.running === run.command && lastOpen.get(run.command) === run ? '运行中' : '没有正常结束（可能被中断）';
    await trow([ctx.time(run.startedAt), txt(run.command), run.dryRun ? '预演' : '实际', runBatchText(run), num(run.limit), num(run.exitCode), ctx.time(run.endedAt), txt(summary)], [
      'datetime',
      'text',
      'text',
      'text',
      'int',
      'int',
      'datetime',
      'text',
    ]);
  }
  if (!runs.length) await s.add(['还没有运行记录'], NOTE);
  await s.finish();
}

/** "amazon" → "Amazon", "205641" → "Sellbrite（205641）". */
function sourceText(source) {
  const label = channelLabel(source);
  if (!label) return String(source ?? '');
  return label.toLowerCase() === String(source).toLowerCase() ? label : `${label}（${source}）`;
}

function runBatchText(run) {
  if (run.command === 'remind') {
    const round = run.options?.round ?? run.round;
    return round ? `第 ${round} 次提醒` : null;
  }
  return Number.isInteger(run.batch) ? String(run.batch) : null;
}

// ---- 发放名单 -----------------------------------------------------------------

/** Column definitions of 发放名单: header, width, cell kind, help text (说明 sheet) and value. */
function recipientColumns(ctx) {
  return [
    { title: '序号', width: 7, kind: 'int', help: '发放顺序：有下单的人按上次下单日期从近到远，后面是从没下单的人，按注册时间从新到旧。名单生成后不再改变。', value: (x) => num(x.r.seq) },
    { title: '状态', width: 13, kind: 'text', help: '发放状态，见“发放状态”。整行的颜色和状态对应。', value: (x) => STATUS_LABELS[x.status] ?? x.status },
    { title: '批次', width: 6, kind: 'int', help: '第几批 issue 发放（或跳过）的这个人。', value: (x) => num(x.state?.batch) },
    { title: '客户 ID', width: 16, kind: 'link', help: '点击打开 Shopify 后台的客户页面。', value: (x, c) => c.link('customers', x.r.customerId, x.r.numericId) },
    { title: '姓名', width: 18, kind: 'text', help: '导出时的客户姓名。', value: (x) => txt(x.r.name) },
    { title: '邮箱', width: 28, kind: 'text', help: '导出时的默认邮箱。', value: (x) => txt(x.r.email) },
    { title: '营销状态', width: 10, kind: 'text', help: '导出时的邮件营销状态。', value: (x) => labelOf(MARKETING_LABELS, x.r.marketingState) || null },
    { title: '客户类型', width: 10, kind: 'text', help: '有下单 / 从没下单 / 测试。', value: (x) => labelOf(KIND_LABELS, x.r.kind) || null },
    {
      title: '上次有效订单号',
      width: 13,
      kind: 'link',
      help: '金额依据的订单：最近一笔未取消、非测试、付过钱的订单。点击打开订单。',
      value: (x, c) => (x.r.basis ? c.link('orders', x.r.basis.orderId, x.r.basis.orderName) : null),
    },
    { title: '上次下单日期', width: 12, kind: 'date', help: '这笔订单的下单日期（店铺时间）。', value: (x, c) => (x.r.basis ? c.day(x.r.basis.createdAt) : null) },
    {
      title: '最近订单渠道',
      width: 13,
      kind: 'text',
      help: '最近一笔有效订单（未取消、非测试，$0 也算）的来源渠道，第 9 条按它判断；没有有效订单时留空。',
      value: (x) => txt(channelLabel(x.r.lastOrderSource)),
    },
    { title: '距今天数', width: 8, kind: 'int', help: '上次有效订单的下单日期到客户数据导出那天相隔的天数。', value: (x, c) => (x.r.basis ? c.daysToExport(x.r.basis.createdAt) : null) },
    { title: '上次订单总额', width: 12, kind: 'money', help: '下单时的总价，含运费和税，不扣后来的退款。', value: (x) => (x.r.basis ? money(x.r.basis.totalCents) : null) },
    { title: '金额算式', width: 42, kind: 'text', help: '基数 = 上次订单总额（从没下单的人用平均值）× 比例，四舍五入到分，再对档位。', value: (x) => txt(x.r.formula) },
    { title: '档位', width: 9, kind: 'text', help: '发放金额所在的档位；测试活动显示“测试”。', value: (x, c) => c.tierText(x.r) },
    { title: '礼品卡金额', width: 11, kind: 'money', help: '卡的面额：已建卡的按卡的金额，其余按名单金额。', value: (x) => money(x.cardCents) },
    { title: '同地址账户数', width: 8, kind: 'int', help: '同一地址上参与去重的候选账户数（含本人）。没有可比对地址的人留空。', value: (x) => (x.r.addressKey ? num(x.r.groupSize ?? 1) : null) },
    { title: '同地址其他账户', width: 18, kind: 'text', help: '同地址其他候选账户的客户 ID，这些账户不发。', value: (x) => (Array.isArray(x.r.groupOthers) ? txt(x.r.groupOthers.join(', ')) : txt(x.r.groupOthers)) },
    {
      title: `是否有 ${ctx.sentTag}`,
      width: 14,
      kind: 'text',
      help:
        `以最近一次从 Shopify 刷新 tag（export --refresh 或 verify）的结果为准：刷新时客户带 ${ctx.sentTag} 显示“是”，不带显示“否”，`
        + '所以在后台删掉的 tag 刷新后会显示“否”。刷新之后才打上 tag 的人，按本地日志显示“是”；从没刷新过时只看本地日志。'
        + '刷新读的是 Shopify 的搜索结果，刚打上的 tag 可能要过几分钟才查得到，所以打 tag 后 10 分钟内的刷新不会把本地日志的“是”改成“否”。',
      value: (x, c) => (c.hasTag(x.state, x.r.customerId) ? '是' : '否'),
    },
    { title: '礼品卡 ID', width: 16, kind: 'link', help: '点击打开 Shopify 后台的礼品卡页面。', value: (x, c) => (x.state?.giftCardId ? c.link('gift_cards', x.state.giftCardId) : null) },
    { title: '卡号后 4 位', width: 9, kind: 'text', help: '卡号的最后 4 位，用来和后台核对。程序从不读取完整卡号。', value: (x) => txt(x.state?.last4) },
    { title: '建卡时间', width: 16, kind: 'datetime', help: '店铺时间。', value: (x, c) => c.time(x.state?.createdAt) },
    { title: '打 tag 时间', width: 16, kind: 'datetime', help: '店铺时间。', value: (x, c) => c.time(x.state?.taggedAt) },
    { title: '第 1 次提醒', width: 20, kind: 'text', help: `见“提醒状态”。这一轮的 tag 是 ${ctx.roundTag(1)}。`, value: (x, c) => c.remindText(x.state, '1'), warn: (x, c) => c.remindWarn(x.state, '1') },
    { title: '第 2 次提醒', width: 20, kind: 'text', help: `见“提醒状态”。这一轮的 tag 是 ${ctx.roundTag(2)}。`, value: (x, c) => c.remindText(x.state, '2'), warn: (x, c) => c.remindWarn(x.state, '2') },
    { title: '已使用金额', width: 11, kind: 'money', help: '最近一次运行 usage 时这张卡已用掉的金额。', value: (x) => (x.card ? money(x.card.usedCents) : null) },
    { title: '剩余余额', width: 11, kind: 'money', help: '最近一次运行 usage 时这张卡的余额。', value: (x) => (x.card ? money(x.card.balanceCents) : null) },
    { title: '使用的订单', width: 16, kind: 'text', help: '用这张卡付过款的订单号。', value: (x, c) => c.ordersText(x.card) },
    { title: '备注/错误', width: 44, kind: 'text', help: '跳过或失败的原因、Shopify 的错误信息，以及金额依据的特殊情况。', value: (x, c) => c.notesFor(x) },
    { title: '城市', width: 14, kind: 'text', help: '客户默认地址的城市。', value: (x) => txt(x.r.city) },
    { title: '州', width: 6, kind: 'text', help: '客户默认地址的州。', value: (x) => txt(x.r.provinceCode) },
    { title: '邮编', width: 9, kind: 'text', help: '客户默认地址的邮编。', value: (x) => txt(x.r.zip) },
    { title: '订单数', width: 7, kind: 'int', help: '导出时 Shopify 记录的订单数。', value: (x) => num(x.r.numberOfOrders) },
    { title: '累计消费', width: 12, kind: 'money', help: '导出时 Shopify 记录的累计消费。', value: (x) => money(x.r.amountSpentCents) },
    { title: '注册日期', width: 12, kind: 'date', help: '客户账户的创建日期。', value: (x, c) => c.day(x.r.accountCreatedAt) },
  ];
}

async function writeRecipientsSheet(book, ctx) {
  const columns = recipientColumns(ctx);
  const s = book.sheet('发放名单', { widths: columns.map((c) => c.width), freezeRows: 1, freezeCols: 3 });
  const header = await s.header(columns.map((c) => c.title));
  const byStatus = new Map(); // status → [style per column]
  const stylesFor = (status) => {
    if (!byStatus.has(status)) byStatus.set(status, columns.map((c) => cellStyle(c.kind, STATUS_FILLS[status] ?? null)));
    return byStatus.get(status);
  };
  for (const x of ctx.rows) {
    const values = new Array(columns.length);
    let rowStyles = stylesFor(x.status);
    for (let i = 0; i < columns.length; i += 1) {
      const col = columns[i];
      values[i] = col.value(x, ctx) ?? null;
      if (col.warn && col.warn(x, ctx)) {
        if (rowStyles === byStatus.get(x.status)) rowStyles = [...rowStyles];
        rowStyles[i] = cellStyle('warn', STATUS_FILLS[x.status] ?? null);
      }
    }
    await s.add(values, rowStyles);
  }
  s.setFilter(header, columns.length);
  await s.finish();
}

// ---- 未入选 ------------------------------------------------------------------

async function writeNotSelectedSheet(book, ctx) {
  const titles = ['客户 ID', '姓名', '邮箱', '营销状态', '主要原因', '全部原因', '上次下单日期', '订单数', '相关账户', '组号', 'tags'];
  const s = book.sheet('未入选', { widths: [16, 18, 28, 10, 38, 56, 12, 7, 16, 8, 30] });
  const header = await s.header(titles);
  const rowStyles = ['link', 'text', 'text', 'text', 'text', 'text', 'date', 'int', 'link', 'text', 'text'].map((k) => cellStyle(k));
  for (const n of ctx.notSelected) {
    await s.add(
      [
        ctx.link('customers', n.customerId, n.numericId),
        txt(n.name),
        txt(n.email),
        labelOf(MARKETING_LABELS, n.marketingState) || null,
        txt(n.primaryText),
        txt(n.allReasons),
        ctx.day(n.lastOrderAt),
        num(n.numberOfOrders),
        n.relatedCustomerId ? ctx.link('customers', n.relatedCustomerId) : null,
        txt(n.groupId),
        txt(n.tags),
      ],
      rowStyles,
    );
  }
  s.setFilter(header, titles.length);
  await s.finish();
}

// ---- 同地址重复 ----------------------------------------------------------------

async function writeDuplicatesSheet(book, ctx) {
  const titles = ['组号', '组内账户数', '规范化地址', '是否保留', '客户 ID', '姓名', '邮箱', '上次付费下单日期', '订单数', '注册日期', '疑似批量注册'];
  const kinds = ['text', 'int', 'text', 'text', 'link', 'text', 'text', 'date', 'int', 'date', 'text'];
  const s = book.sheet('同地址重复', { widths: [15, 12, 40, 16, 18, 18, 28, 14, 8, 12, 12] });
  const header = await s.header(titles);
  const banded = [kinds.map((k) => cellStyle(k)), kinds.map((k) => cellStyle(k, COLOR.band))];
  let group = 0;
  for (const g of ctx.duplicates) {
    const rowStyles = banded[group % 2];
    group += 1;
    for (const m of objects(g.members)) {
      await s.add(
        [
          txt(g.groupId),
          num(g.size),
          txt(g.address),
          m.kept ? '保留' : '不发',
          ctx.link('customers', m.customerId, m.numericId),
          txt(m.name),
          txt(m.email),
          ctx.day(m.lastPaidOrderAt),
          num(m.numberOfOrders),
          ctx.day(m.accountCreatedAt),
          g.flaggedBulk ? '是' : null,
        ],
        rowStyles,
      );
    }
  }
  s.setFilter(header, titles.length);

  // People not sent because another account at the same address ordered recently.
  const active = ctx.notSelected.filter((n) => n.primaryCode === 'active-address');
  await s.blank();
  await s.blank();
  await s.add([`因同地址账户近 ${ctx.inactiveMonths} 个月下过单而不发的人（${fmtCount(active.length)} 人）`], SECTION, { height: 22 });
  await s.header(['客户 ID', '姓名', '邮箱', '活跃账户', '活跃账户下单时间']);
  const activeStyles = ['link', 'text', 'text', 'link', 'datetime'].map((k) => cellStyle(k));
  for (const n of active) {
    await s.add(
      [
        ctx.link('customers', n.customerId, n.numericId),
        txt(n.name),
        txt(n.email),
        n.relatedCustomerId ? ctx.link('customers', n.relatedCustomerId) : null,
        ctx.time(n.relatedAt),
      ],
      activeStyles,
    );
  }
  await s.finish();
}

// ---- 使用明细 ------------------------------------------------------------------

function itemsText(lineItems) {
  const items = objects(lineItems);
  if (!items.length) return null;
  return txt(items.map((l) => `${l.name ?? ''} × ${l.quantity ?? 1}`).join('; '));
}

async function writeUsageDetailSheet(book, ctx) {
  const u = ctx.usage;
  const s = book.sheet('使用明细', { widths: [12, 17, 16, 17, 12, 14, 12, 8, 60] });
  const titles = ['订单号', '下单时间', '客户 ID', '卡号后 4 位', '卡面额', '本单用卡金额', '订单总额', '已取消', '买了什么'];
  const header = await s.header(titles);
  const rowStyles = ['link', 'datetime', 'link', 'text', 'money', 'money', 'money', 'text', 'text'].map((k) => cellStyle(k));
  const payments = objects(u.payments)
    .map((p, i) => ({ p, i, ms: toMs(p.orderCreatedAt) ?? 0 }))
    .sort((a, b) => a.ms - b.ms || a.i - b.i)
    .map((x) => x.p);
  for (const p of payments) {
    const order = ctx.ordersById.get(p.orderId) ?? null;
    const card = ctx.cardsById.get(p.giftCardId) ?? null;
    await s.add(
      [
        ctx.link('orders', p.orderId, p.orderName || order?.orderName),
        ctx.time(p.orderCreatedAt ?? order?.createdAt),
        ctx.link('customers', p.orderCustomerId ?? order?.customerId ?? p.cardCustomerId),
        txt(card?.last4),
        money(card?.initialCents),
        money(p.netCents ?? p.amountCents),
        money(order?.totalCents),
        order?.cancelled ? '是' : null,
        itemsText(order?.lineItems),
      ],
      rowStyles,
    );
  }
  s.setFilter(header, titles.length);

  const unmatched = objects(u.unmatched);
  await s.blank();
  await s.blank();
  if (!unmatched.length) {
    await s.add(['没有需要人工核对的礼品卡付款'], NOTE);
  } else {
    await s.add([`需要人工核对的礼品卡付款：回执里没有本活动礼品卡 ID（${fmtCount(unmatched.length)} 笔）`], SECTION, { height: 22 });
    await s.header(['订单号', '下单时间', '客户 ID', '付款时间', '金额', '原因']);
    const unmatchedStyles = ['link', 'datetime', 'link', 'datetime', 'money', 'text'].map((k) => cellStyle(k));
    for (const m of unmatched) {
      await s.add(
        [
          ctx.link('orders', m.orderId, m.orderName),
          ctx.time(m.orderCreatedAt),
          ctx.link('customers', m.customerId),
          ctx.time(m.processedAt),
          money(m.amountCents),
          txt(labelOf(UNMATCHED_REASON_LABELS, m.reason)),
        ],
        unmatchedStyles,
      );
    }
  }
  await s.finish();
}

// ---- 核对 ---------------------------------------------------------------------

async function writeVerifySheet(book, ctx) {
  const v = ctx.verify;
  const issues = objects(v.issues);
  const s = book.sheet('核对', { widths: [30, 16, 16, 36, 36, 46], freezeRows: 3 });
  await s.add([`核对时间 ${ctx.stamp(v.verifiedAt) || '未知'}（${ctx.tzLabel}）；本活动的卡 ${fmtCount(v.cardCount)} 张；带 ${ctx.sentTag} 的客户 ${fmtCount(v.taggedCount)} 个`], SECTION, { height: 22 });
  const counts = Object.entries(v.counts ?? {})
    .filter(([, n]) => n > 0)
    .map(([type, n]) => `${labelOf(VERIFY_TYPE_LABELS, type)} ${fmtCount(n)}`)
    .join('；');
  await s.add([issues.length ? `发现 ${fmtCount(issues.length)} 个问题${counts ? `：${counts}` : ''}` : '没有发现问题'], NOTE);
  const titles = ['问题类型', '客户 ID', '礼品卡 ID', '日志记录', 'Shopify 实际', '建议操作'];
  const header = await s.header(titles);
  const rowStyles = ['text', 'link', 'link', 'text', 'text', 'text'].map((k) => cellStyle(k));
  for (const i of issues) {
    await s.add(
      [
        txt(labelOf(VERIFY_TYPE_LABELS, i.type)),
        i.customerId ? ctx.link('customers', i.customerId) : null,
        i.giftCardId ? ctx.link('gift_cards', i.giftCardId) : null,
        txt(i.journal),
        txt(i.shopify),
        txt(i.action),
      ],
      rowStyles,
    );
  }
  s.setFilter(header, titles.length);
  await s.finish();
}

// ---- 操作日志 -----------------------------------------------------------------

/**
 * "结果/说明" of one journal entry. Skip details are made readable: times in
 * store time, customers by number, marketing states in Chinese (ctx.stamp is
 * the workbook's store-local clock).
 */
function entryResult(e, ctx) {
  const amount = Number.isInteger(e.amountCents) ? `金额 ${formatUsd(e.amountCents)}` : '';
  const last4 = e.last4 ? `卡号后 4 位 ${e.last4}` : '';
  const detail = { stamp: ctx.stamp };
  switch (e.op) {
    case 'create.start':
      return amount;
    case 'create.ok':
      return [amount, last4].filter(Boolean).join('，');
    case 'reconcile.found':
      return [labelOf(RECONCILE_SOURCE_LABELS, e.source), amount, last4].filter(Boolean).join('，');
    case 'skip':
      return issueSkipText(e.reason, e.detail, detail);
    case 'remind.skip':
      return remindSkipText(e.reason, e.detail, detail);
    case 'remind.start':
      return e.retry ? '重新发送（之前失败或结果不明）' : noteText(e.note);
    case 'remind.found':
      return `Shopify 上已带 ${ctx.roundTag(e.round) ?? '本轮 tag'}`;
    case 'remind.tag.fail':
      return `${ctx.roundTag(e.round) ?? '本轮 tag'} 没打上，下次运行会补打`;
    case 'remind.tag.ok': {
      // the repair: the tag added again, or (note) already in Shopify although its tagsAdd looked failed
      const tag = ctx.roundTag(e.round) ?? '本轮 tag';
      if (e.note === 'already tagged in Shopify') return `Shopify 上已带 ${tag}`;
      return e.note ? noteText(e.note) : `已补打 ${tag}`;
    }
    default:
      return noteText(e.note);
  }
}

async function writeJournalSheet(book, ctx) {
  const titles = ['时间', '命令', '批次/轮次', '客户 ID', '动作', '结果/说明', '礼品卡 ID', '错误'];
  const s = book.sheet('操作日志', { widths: [16, 9, 12, 16, 18, 44, 16, 44] });
  const header = await s.header(titles);
  const rowStyles = ['datetime', 'text', 'text', 'text', 'text', 'text', 'text', 'warn'].map((k) => cellStyle(k));
  for (const e of ctx.entries) {
    if (!e || typeof e.op !== 'string' || e.op.startsWith('run.')) continue;
    const run = ctx.runsById.get(e.run);
    const isRemind = e.op.startsWith('remind.');
    let batch = null;
    if (isRemind) batch = e.round !== undefined && e.round !== null ? `第 ${e.round} 次提醒` : null;
    else if (Number.isInteger(e.batch ?? run?.batch)) batch = String(e.batch ?? run.batch);
    // Entries whose run.start is missing: the op itself tells remind and issue
    // apart; reconcile.* can also come from verify, so it stays blank.
    const command = run?.command ?? (isRemind ? 'remind' : /^(create|tag)\.|^skip$/.test(e.op) ? 'issue' : null);
    await s.add(
      [
        ctx.time(e.t),
        txt(command),
        batch,
        txt(numericId(e.cid) || e.cid),
        labelOf(OP_LABELS, e.op),
        txt(entryResult(e, ctx)),
        txt(numericId(e.giftCardId)),
        txt(errorText(e.error)),
      ],
      rowStyles,
    );
  }
  s.setFilter(header, titles.length);
  await s.finish();
}

// ---- 说明 ---------------------------------------------------------------------

async function writeHelpSheet(book, ctx) {
  const s = book.sheet('说明', { widths: [30, 120] });
  const item = cellStyle('text');
  const section = async (title) => {
    await s.blank();
    await s.add([title], SECTION, { height: 22 });
  };
  const rows = async (pairs) => {
    for (const [k, v] of pairs) await s.add([txt(k), txt(v)], [item, item]);
  };
  await s.header(['项目', '说明']);

  await section('各表说明');
  await rows([
    ['使用报告', '运行 usage 后出现，排在第一张：发出多少卡、用了多少、带来多少订单和金额、按档位和客户类型的使用率、每日趋势、卖得最多的商品。'],
    ['汇总', '活动参数、筛选漏斗、金额、发放进度、提醒、使用情况和运行记录。'],
    ['发放名单', '每个入选者一行，按发放顺序（序号）。整行颜色表示发放状态。前 3 列和表头冻结。'],
    ['未入选', '被排除的人和原因。只因第 1–3 条（没邮箱、平台中转或占位邮箱、未订阅营销邮件）被排除的人不逐行列出，只在“汇总”里计数。'],
    ['同地址重复', '同一地址的多个候选账户只保留一个（按组着色）；下面另列因同地址账户近期下过单而不发的人。'],
    ['使用明细', '运行 usage 后出现：每笔用本活动礼品卡付款一行；下面另列回执里没有礼品卡 ID、需要人工核对的付款。'],
    ['核对', '运行 verify 后出现：本地日志和 Shopify 对不上的地方，以及建议操作。'],
    ['操作日志', '每次写操作的记录（建卡、打 tag、跳过、提醒），按时间顺序。发提醒后直接打上本轮 tag 的不单独记一行，见“补打本轮 tag”。'],
    ['说明', '本页。'],
  ]);

  await section('发放名单各列');
  await rows(recipientColumns(ctx).map((c) => [c.title, c.help]));

  await section('发放状态（发放名单“状态”列和整行颜色）');
  for (const status of STATUS_ORDER) {
    await s.add([STATUS_LABELS[status], STATUS_HELP[status]], [cellStyle('text', STATUS_FILLS[status] ?? null), item]);
  }

  await section('提醒状态（“第 1 次提醒”“第 2 次提醒”两列）');
  await rows(REMIND_HELP);
  await rows([
    ['提醒跳过的原因', Object.values(REMIND_SKIP_LABELS).join('、')],
    ['发放前跳过的原因', Object.values(ISSUE_SKIP_LABELS).join('、')],
  ]);

  await section('本轮 tag（每一轮每人最多一封提醒）');
  await rows(roundTagHelp(ctx.sentTag));

  await section('金额规则');
  const percent = ctx.params.giftPercent ?? ctx.config.giftPercent;
  if (ctx.isTest) {
    await rows([['测试活动', `每人固定 ${formatUsd(ctx.params.testGiftAmountCents ?? ctx.rows[0]?.r.amountCents)}，不按档位。`]]);
  } else {
    await rows([
      ['基数', `有下单的人：上次有效订单总额 × ${percent}%，四舍五入到分。上次有效订单是最近一笔未取消、非测试、付过钱的订单；最近一笔已取消、是测试订单或是 $0，就继续往前找。`],
      ['从没下单的人', `平均值 × ${percent}%：平均值是本次所有“有下单”入选者上次订单总额的平均数（先四舍五入到分）。订单都已取消或都是 $0 的人也按从没下单处理。`],
      ...(ctx.tiersCents.length ? describeTiers(ctx.tiersCents).map((t, i) => [`档位 ${i + 1}`, `基数 ${t}`]) : []),
      ['先四舍五入再对档位', '例：上次订单 $107.74 → 基数 $10.77 → 档位 $10.77；$107.75 → 基数 $10.78 → 档位 $15.33。'],
    ]);
  }

  await section('使用情况的金额（“使用报告”和“汇总”）');
  await rows([['礼品卡抵扣和已用金额', USAGE_SPLIT_NOTE]]);

  await section('其他');
  await rows([
    ['时间', `所有时间都是店铺时区（${ctx.timeZone}，${ctx.tzLabel}）。`],
    ['链接', '客户 ID、订单号、礼品卡 ID 可以点击，打开 Shopify 后台对应的页面。'],
    ['只读', '本文件由程序生成，每次运行命令都会重新生成；手动修改会被覆盖，程序也从不读取这份 Excel。'],
    ['筛选', '表格已加保护（无密码），可以使用筛选和调整列宽。'],
    ...EXIT_CODE_HELP.map(([code, text]) => [`退出码 ${code}`, text]),
  ]);
  await s.finish();
}
