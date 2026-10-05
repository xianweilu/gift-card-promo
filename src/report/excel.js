// Builds the campaign workbooks from local state only: selection.json +
// journal.jsonl, plus tags.json, verify.json and usage.json when they exist.
// Two editions with the same content and layout:
//   campaigns/<id>/gift-card-promo-<id>.xlsx      Chinese (TEXT.zh of ./text.js)
//   campaigns/<id>/gift-card-promo-<id>-en.xlsx   English (TEXT.en)
// The program never reads a workbook back; every command simply regenerates them.
//
// Design notes
// - Streaming (ExcelJS WorkbookWriter): each row is styled, committed and
//   released, so 20k recipients + 50k "未入选" rows stay within a flat memory
//   budget. A committed row can no longer be touched, so cells are styled first.
// - Excel has no time zones: times are written as the STORE-LOCAL wall clock
//   (selection.params.timezone), i.e. Dates whose UTC fields are the local time.
// - Links are HYPERLINK() formulas, because Excel caps real hyperlinks at ~65k
//   per sheet.
// - Each file is written to a temp file in the same directory and renamed over
//   the target while holding the Excel lock, so nobody sees half a file and a
//   later writer always includes the newest journal entries.
// - Every visible text comes from the active text pack (ctx.T, ctx.lang); the
//   data is read and computed once (buildContext) and shown by both editions
//   (localize). Console messages stay Chinese (CONSOLE).

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';

import { readJson, readJournal, foldJournal, runningCommand, withExcelLock, ensureDir } from '../campaign.js';
import { campaignNote } from '../giftcards.js';
import { formatUsd, tierLabel } from '../select/amount.js';
import { UNLISTED_RULES } from '../select/rules.js';
import { TEXT, textFor, CONSOLE } from './text.js';
import {
  STATUS_ORDER,
  STATUS_FILLS,
  labelOf,
  storedLabelOf,
  issueSkipText,
  noteText,
  errorText,
  summaryText,
  usedTagName,
  reminderHelp,
} from './labels.js';

export const DEFAULT_TIMEZONE = 'America/Los_Angeles';
export const OPEN_IN_EXCEL_WARNING = CONSOLE.openInExcel;
/** The same warning about the English workbook (gift-card-promo-<id>-en.xlsx). */
export const OPEN_IN_EXCEL_WARNING_EN = CONSOLE.openInExcelEn;

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

/** Label of an issue status, e.g. "已完成" ('zh', the default) or the English one; see effectiveStatus for in_progress. */
export function statusLabel(status, running, lang = 'zh') {
  const s = effectiveStatus(status, running);
  return textFor(lang).STATUS_LABELS[s] ?? String(s);
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

/**
 * The English edition's file next to a workbook file: "<name>.xlsx" → "<name>-en.xlsx"
 * (a name without ".xlsx" gets "-en" at the end). campaignPaths().excelEn is this of .excel.
 */
export function englishWorkbookPath(file) {
  return path.join(path.dirname(file), path.basename(file).replace(/(\.xlsx)?$/i, (ext) => `-en${ext}`));
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

function timeZoneLabel(tz, T) {
  return tz === DEFAULT_TIMEZONE ? T.timeZone.default : T.timeZone.other(tz);
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
const XML_SUSPECT = /[\uD800-\uDFFF￾￿]/;
const XML_NONCHARACTERS = /[￾￿]/g;
const LONE_SURROGATE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;
const wellFormed = typeof String.prototype.toWellFormed === 'function' ? (s) => s.toWellFormed() : (s) => s.replace(LONE_SURROGATE, '�');

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
// the style id per object (per workbook), so sharing them keeps 1M+ cells fast.
// Never mutate a style object after creation (cells point at them).
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
  /**
   * @param {string} file the (temp) file to write
   * @param {{ generatedAt: Date, warn: (message: string) => void }} options `warn` logs and collects a warning
   */
  constructor(file, { generatedAt, warn }) {
    this.warn = warn;
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
    if (this.truncated) this.book.warn(CONSOLE.rowLimit(this.name, EXCEL_MAX_ROWS.toLocaleString('en-US')));
    this.book.check();
  }
}

// ---------------------------------------------------------------------------
// writeReport
// ---------------------------------------------------------------------------

/**
 * Rebuild both campaign workbooks from selection.json + journal.jsonl (+ tags.json,
 * verify.json, usage.json when present): the Chinese one (paths.excel) and the English
 * one (paths.excelEn), with the same content. Holds the Excel lock while reading the
 * state and replacing the files, so concurrent calls never interleave.
 *
 * The Chinese workbook is written first, exactly as before: if it fails, writeReport
 * rejects. A failure of the English workbook is only a warning
 * ("英文版 Excel 没有生成：<reason>") with fileEn null; the previous English file, if
 * any, is left as it was.
 *
 * @param {object} a
 * @param {object} a.config loadConfig() result
 * @param {object} a.paths campaignPaths(config) (paths.excelEn defaults to englishWorkbookPath(paths.excel))
 * @param {{info: Function, warn: Function, error: Function}} [a.log]
 * @param {() => Date} [a.now] "generated at" time shown in the workbooks
 * @param {string|null} [a.out] also save a copy here (a file, or a directory to put it in);
 *   the English workbook is copied next to it as "<name>-en.xlsx"
 * @param {object} [a.lockOptions] passed to withExcelLock ({ timeoutMs, pollMs, sleep }); for tests
 * @param {{ zh?: object, en?: object }} [a.packs] text packs to use instead of TEXT.zh / TEXT.en; for tests
 * @returns {Promise<{ file: string, fileEn: string|null, out: string|null, outEn: string|null, warnings: string[] }>}
 *   every warning is logged once by writeReport itself; `warnings` is a copy for callers and tests
 */
export async function writeReport({ config, paths, log = console, now = () => new Date(), out = null, lockOptions = {}, packs = null } = {}) {
  if (!config || !paths) throw new Error('writeReport needs config and paths');
  const zh = packs?.zh ?? TEXT.zh;
  const en = packs?.en ?? TEXT.en;
  return withExcelLock(paths, () => buildAndPublish({ config, paths, log, now, out, zh, en }), lockOptions);
}

async function buildAndPublish({ config, paths, log, now, out, zh, en }) {
  // Every warning is logged here, once, and also returned (callers must not print them again).
  const warnings = [];
  const warn = (message) => {
    warnings.push(message);
    log.warn(message);
  };
  const excelEn = paths.excelEn ?? englishWorkbookPath(paths.excel);

  // We hold the Excel lock, so no other writer of these workbooks is running:
  // any temp workbook still in the folder was left by a killed write.
  removeLeftoverTemps(paths.excel);
  removeLeftoverTemps(excelEn);

  // Read once: both editions show the same state.
  const selection = readJson(paths.selection, null);
  if (!selection) throw new Error(CONSOLE.noSelection);
  const entries = readJournal(paths.journal);
  const journal = foldJournal(entries);
  const tags = readOptionalJson(paths.tags, 'tags.json', warn);
  const verify = readOptionalJson(paths.verify, 'verify.json', warn);
  const usage = readOptionalJson(paths.usage, 'usage.json', warn);
  const running = runningCommand(paths);
  const generatedAt = toDate(now());

  const base = buildContext({ config, selection, entries, journal, tags, verify, usage, running, generatedAt, warn });

  ensureDir(path.dirname(paths.excel));
  await writeWorkbook(paths.excel, localize(base, zh), { generatedAt, warn });

  let fileEn = null;
  try {
    ensureDir(path.dirname(excelEn));
    await writeWorkbook(excelEn, localize(base, en), { generatedAt, warn: (message) => warn(`${CONSOLE.englishPrefix}${message}`) });
    fileEn = excelEn;
  } catch (err) {
    warn(CONSOLE.englishFailed(err?.message ?? String(err)));
  }

  // The main workbooks are in place now: a failed extra copy must not look like a failed export.
  let outFile = null;
  let outEn = null;
  if (out) {
    let target = null;
    try {
      target = resolveOut(out, paths.excel);
      outFile = copyOut(paths.excel, target, { other: excelEn, failed: CONSOLE.copyFailed });
    } catch (err) {
      warn(err.message);
    }
    // The English copy goes next to the Chinese one ("<name>-en.xlsx"); a target that is the
    // English workbook itself was refused above and gets no "-en-en" copy.
    if (fileEn && target && target !== path.resolve(excelEn)) {
      try {
        outEn = copyOut(fileEn, englishWorkbookPath(target), { other: paths.excel, failed: CONSOLE.copyFailedEn });
      } catch (err) {
        warn(err.message);
      }
    }
  }

  for (const file of new Set([paths.excel, outFile].filter(Boolean))) {
    if (openInExcel(file)) warn(file === paths.excel ? OPEN_IN_EXCEL_WARNING : CONSOLE.openInExcelAt(file));
  }
  for (const file of new Set([fileEn, outEn].filter(Boolean))) {
    if (openInExcel(file)) warn(file === fileEn ? OPEN_IN_EXCEL_WARNING_EN : CONSOLE.openInExcelEnAt(file));
  }
  return { file: paths.excel, fileEn, out: outFile, outEn, warnings };
}

/** Write one edition: a temp file in the target's folder, renamed over `file` when complete. */
async function writeWorkbook(file, ctx, { generatedAt, warn }) {
  const tmp = tempPathFor(file);
  const book = new Book(tmp, { generatedAt, warn });
  try {
    await writeSheets(book, ctx);
    await book.commit();
    replaceFile(tmp, file);
  } catch (err) {
    book.discard();
    removeQuietly(tmp);
    throw err;
  }
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
    warn(CONSOLE.unreadable(label, err.message));
    return null;
  }
}

function replaceFile(tmp, target) {
  try {
    fs.renameSync(tmp, target);
  } catch (err) {
    if (['EBUSY', 'EPERM', 'EACCES'].includes(err.code)) {
      throw new Error(CONSOLE.replaceFailed(target, err.code));
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
 * Copy a finished workbook to `target` (via a temp file + rename, so a reader
 * never sees half a copy). The workbook itself is not copied onto itself, and a
 * copy never replaces the other edition's workbook (`other`). Throws
 * `failed(target, reason)` (e.g. "另存 Excel 到 … 失败：…") on any failure.
 */
function copyOut(source, target, { other, failed }) {
  if (target === path.resolve(source)) return target;
  if (target === path.resolve(other)) throw new Error(failed(target, CONSOLE.copyOntoOther));
  const tmp = tempPathFor(target);
  try {
    ensureDir(path.dirname(target));
    fs.copyFileSync(source, tmp);
    fs.renameSync(tmp, target);
  } catch (err) {
    removeQuietly(tmp);
    throw new Error(failed(target, err.message));
  }
  return target;
}

// ---------------------------------------------------------------------------
// Context: everything the sheets need, computed once for both editions
// ---------------------------------------------------------------------------

function buildContext({ config, selection, entries, journal, tags, verify, usage, running, generatedAt, warn }) {
  const params = selection.params ?? {};
  let timeZone = params.timezone || config.timezone || DEFAULT_TIMEZONE;
  if (!validTimeZone(timeZone)) {
    warn(CONSOLE.badTimeZone(timeZone, DEFAULT_TIMEZONE));
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
      warn(CONSOLE.foreignTags(tags.tag, sentTag));
    } else if (!Array.isArray(tags.ids)) {
      warn(CONSOLE.tagsWithoutIds);
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

  /**
   * "用卡 tag": the customer carries "<SENT_TAG>-USED" when the journal recorded usage adding it
   * (used.tag.ok, the latest outcome) or the latest usage run read the customer from Shopify's
   * tag set (usage.json usedTag.taggedCustomerIds). A usage.json about another campaign's tag
   * is ignored with a warning.
   */
  const usedTag = usedTagName(sentTag);
  let usedTaggedIds = new Set();
  if (usage?.usedTag && typeof usage.usedTag === 'object') {
    if (String(usage.usedTag.tag ?? '').toLowerCase() !== usedTag.toLowerCase()) {
      warn(CONSOLE.foreignUsedTag(usage.usedTag.tag, usedTag));
    } else if (Array.isArray(usage.usedTag.taggedCustomerIds)) {
      usedTaggedIds = new Set(usage.usedTag.taggedCustomerIds);
    }
  }
  const hasUsedTag = (state, customerId) => state?.usedTag?.status === 'tagged' || usedTaggedIds.has(customerId);

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

  // Campaign dates. LAUNCH_DATE is shown as issue checks it now (.env); the
  // expiry as frozen in the list, because every card is created with that one.
  // A difference to the list gets a note (summary sheet).
  const dateSetting = (current, frozen) => {
    const now = current || '';
    const then = frozen ?? null; // null: a list written without this parameter, nothing to compare
    return { current: now, frozen: then, changed: then !== null && (then || '') !== now };
  };
  // issue creates every card with the list's expiry ('' = none), whatever .env says now.
  const expiryFrozen = params.giftCardExpiresOn || '';
  const dates = {
    launch: dateSetting(config.launchDate, params.launchDate),
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

  const base = {
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
    shop,
    sentTag,
    tiersCents,
    isTest,
    campaignId: selection.campaignId || config.campaignId,
    inactiveMonths: params.inactiveMonths ?? config.inactiveMonths,
    dates,
    hasTag,
    usedTag,
    hasUsedTag,
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
    ordersText(card) {
      const names = card ? orderNamesByCard.get(card.giftCardId) : null;
      return names?.size ? txt([...names].filter(Boolean).join(', ')) : null;
    },
  };
  base.progress = computeProgress(rows);
  base.usageSummary = usage ? normalizeUsageSummary(usage, cards) : null;
  base.usedTagStats = usage ? usedTagStats(usage, usedTaggedIds, journal.runs) : null;
  return base;
}

/** The context of one edition: the shared data plus the texts of pack `T` (ctx.T, ctx.lang). */
function localize(base, T) {
  const { timeZone, tiersCents } = base;
  const ctx = {
    ...base,
    T,
    lang: T.lang,
    tzLabel: timeZoneLabel(timeZone, T),
    tierText(r) {
      if (r.kind === 'test') return T.recipients.testTier;
      return Number.isInteger(r.tier) ? txt(tierLabel(r.tier, tiersCents)) : null;
    },
  };
  ctx.notesFor = (x) => recipientNotes(x, ctx);
  return ctx;
}

/** "备注/错误": why a row is special, joined with the pack's list separator. */
function recipientNotes({ r, state, status }, ctx) {
  const { T } = ctx;
  const N = T.recipients.notes;
  const notes = [];
  if (state) {
    if (status === 'skipped' && state.skipReason) notes.push(issueSkipText(state.skipReason, undefined, undefined, T));
    if (state.status === 'in_progress' && status !== 'in_progress') notes.push(N.interrupted);
    if (state.error) notes.push(errorText(state.error, T));
    if (state.reconciledFrom) notes.push(N.reconciledFrom(labelOf(T.RECONCILE_SOURCE_LABELS, state.reconciledFrom)));
    if (state.giftCardId && Number.isInteger(state.amountCents) && Number.isInteger(r.amountCents) && state.amountCents !== r.amountCents) {
      notes.push(N.amountDiffers(formatUsd(state.amountCents), formatUsd(r.amountCents)));
    }
  }
  // basisWhy is also set when no earlier paid order was found (then neverReason
  // explains it), so the "earlier order" note only applies with a basis order.
  if (r.basis && T.BASIS_NOTES[r.basisWhy]) notes.push(T.BASIS_NOTES[r.basisWhy]);
  if (T.NEVER_NOTES[r.neverReason]) notes.push(T.NEVER_NOTES[r.neverReason]);
  const group = r.groupId ? ctx.groupsById.get(r.groupId) : null;
  if (group?.flaggedBulk) notes.push(N.bulk(group.size));
  return txt(notes.join(T.sep.list));
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

/**
 * The used-card tag figures of the 使用报告 / 汇总: how many customers carry the tag after the latest
 * usage run (usage.json usedTag.taggedCustomerIds; older files: the latest usage run's counters),
 * and how many that run added / failed to add (its run.end summary usedTagAdded / usedTagFailed).
 */
function usedTagStats(usage, usedTaggedIds, runs) {
  const latest = runs
    .filter((r) => r.command === 'usage' && r.summary && typeof r.summary === 'object' && !Array.isArray(r.summary))
    .sort((a, b) => String(a.startedAt).localeCompare(String(b.startedAt)))
    .at(-1)?.summary ?? {};
  const n = (v) => (typeof v === 'number' && Number.isFinite(v) ? v : 0);
  const fromFile = Array.isArray(usage.usedTag?.taggedCustomerIds) ? usedTaggedIds.size : null;
  return {
    tagged: fromFile ?? n(latest.usedTagged) + n(latest.usedTagAdded),
    added: n(latest.usedTagAdded),
    failed: n(latest.usedTagFailed),
  };
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

// ---- 使用报告 (usage report) -------------------------------------------------

async function writeUsageReportSheet(book, ctx) {
  const { T } = ctx;
  const U = T.usageReport;
  const P = T.moneyPrefix;
  const u = ctx.usageSummary;
  const asOf = ctx.stamp(ctx.usage.fetchedAt) || U.unknownTime;
  const s = book.sheet(T.sheets.usageReport, { widths: [36, 16, 28, 28, 16, 14] });
  const c = (kind) => cellStyle(kind);
  const label = cellStyle('bold', null, false);

  await s.add([U.title(asOf, ctx.tzLabel)], TITLE, { height: 24 });
  await s.blank();
  // Totals, laid out like the plan's sample; the units live in the number formats
  // so every figure stays a real number.
  await s.add([U.issued, u.issuedCards, money(u.issuedCents)], [label, formatStyle(T.fmt.cards), prefixedMoney(P.faceValue)]);
  await s.add([U.used, u.usedCards, u.usedRate, money(u.usedCents), u.usedCentsRate], [
    label,
    formatStyle(T.fmt.cards),
    formatStyle(T.fmt.share),
    prefixedMoney(P.used),
    formatStyle(T.fmt.shareOfValue),
  ]);
  await s.add([U.orders, u.orders, money(u.ordersTotalCents), money(u.avgOrderCents)], [label, formatStyle(T.fmt.orders), prefixedMoney(P.ordersTotal), prefixedMoney(P.perOrder)]);
  await s.add([null, null, money(u.giftCardCents), money(u.customerPaidCents)], [null, null, prefixedMoney(P.giftCardPaid), prefixedMoney(P.customerPaid)]);
  // The gift-card part at checkout and the used amount (card balances) differ after a refund back to a card.
  await s.add([T.USAGE_SPLIT_NOTE], NOTE);
  const ut = ctx.usedTagStats;
  await s.add([U.usedTag(fmtCount(ut.tagged), fmtCount(ut.added), fmtCount(ut.failed))], label);

  await s.blank();
  await s.header(U.byTier);
  for (const t of u.byTier) {
    const name = t.label || (Number.isInteger(t.tier) ? tierLabel(t.tier, ctx.tiersCents) : '') || String(t.tier ?? '');
    await s.add([txt(name), num(t.issued), num(t.used), ratio(t.used, t.issued) ?? rateValue(t.rate), money(t.usedCents)], [c('text'), c('count'), c('count'), c('pct'), c('money')]);
  }

  await s.blank();
  await s.header(U.byKind);
  for (const k of u.byKind) {
    await s.add([labelOf(T.KIND_LABELS, k.kind) || null, num(k.issued), num(k.used), ratio(k.used, k.issued) ?? rateValue(k.rate), money(k.usedCents)], [c('text'), c('count'), c('count'), c('pct'), c('money')]);
  }

  await s.blank();
  await s.header(U.daily);
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
  await s.header(U.topProducts);
  for (const p of u.topProducts.slice(0, 10)) {
    await s.add([txt(T.stored.productName(p.name)), num(p.quantity), money(p.amountCents)], [c('text'), c('count'), c('money')]);
  }

  await s.blank();
  await s.add([U.footer(asOf, ctx.tzLabel)], NOTE);
  await s.finish();
}

// ---- 汇总 (summary) ----------------------------------------------------------

async function writeSummarySheet(book, ctx) {
  const { selection, params, T } = ctx;
  const S = T.summary;
  const stats = selection.stats ?? {};
  const funnel = selection.funnel ?? null;
  const snapshot = selection.snapshot ?? {};
  const s = book.sheet(T.sheets.summary, { widths: [40, 18, 14, 14, 14, 12, 18, 70] });
  const plain = (kind) => cellStyle(kind, null, false);
  const kv = (label, value, kind = 'text', note = null, noteStyle = NOTE) => s.add([label, value ?? null, note ?? null], [LABEL, plain(kind), noteStyle]);
  const section = async (title) => {
    await s.blank();
    await s.add([title], SECTION, { height: 22 });
  };
  const trow = (values, kinds) => s.add(values, kinds.map((k) => (typeof k === 'string' ? cellStyle(k) : k)));
  // Rows sorted by count, then by the stored key: both editions list them in the same order.
  const breakdown = async (title, counts, keyLabel = (k) => k) => {
    const items = Object.entries(counts ?? {}).filter(([, n]) => n > 0).sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])));
    if (!items.length) return;
    await s.blank();
    await s.header([title, S.people]);
    for (const [k, n] of items) await trow([txt(keyLabel(k)) ?? S.emptyKey, n], ['text', 'count']);
  };
  /** The note of a .env date that differs from the list's. */
  const changedNote = (date) => S.dateChanged(date.frozen || T.notSet, date.current || T.notSet);

  await s.add([S.notice], NOTICE, { height: 20 });
  await s.add([S.subtitle(ctx.campaignId, ctx.isTest, ctx.stamp(ctx.generatedAt), ctx.tzLabel)], NOTE);

  // ---- campaign parameters
  await section(S.params);
  await kv(S.campaignId, ctx.campaignId);
  await kv(S.listType, ctx.isTest ? S.listTypeTest : S.listTypeLive);
  await kv(S.createdAt, ctx.time(selection.createdAt), 'datetime');
  await kv(S.exportedAt, ctx.time(snapshot.exportedAt), 'datetime', S.exportedNote(labelOf(T.SNAPSHOT_SOURCE_LABELS, snapshot.source) || S.sourceUnknown, fmtCount(snapshot.count)));
  if (ctx.isTest) {
    await kv(S.testAmount, money(params.testGiftAmountCents ?? ctx.rows[0]?.r.amountCents), 'money');
  } else {
    await kv(S.cutoff(ctx.inactiveMonths), ctx.ymd(params.cutoffDate), 'date', S.cutoffNote);
    await kv(S.minAge, S.days(params.minAccountAgeDays ?? ctx.config.minAccountAgeDays));
    await kv(S.requireSubscribed, (params.requireEmailSubscribed ?? ctx.config.requireEmailSubscribed) ? T.yes : T.no);
    await kv(S.percent, S.percentValue(params.giftPercent ?? ctx.config.giftPercent), 'text', S.percentNote);
    await kv(S.tiers, ctx.tiersCents.length ? T.describeTiers(ctx.tiersCents).join(T.sep.list) : null);
    await kv(S.average, money(selection.averageCents), 'money', S.averageNote);
    await kv(S.neverAmount, money(selection.neverAmountCents), 'money');
  }
  // The expiry is frozen in the list (every card is created with it); the
  // launch date is the current .env value issue checks.
  const { expiry, launch } = ctx.dates;
  await kv(
    S.expiry,
    expiry.frozen ? ctx.ymd(expiry.frozen) : S.noExpiry,
    'date',
    expiry.changed ? S.expiryChanged(expiry.env || T.notSet, expiry.frozen || S.noExpiry) : S.expiryNote,
    expiry.changed ? CHANGED : NOTE,
  );
  await kv(S.launchDate, ctx.ymd(launch.current) ?? T.notSet, 'date', launch.changed ? changedNote(launch) : null, CHANGED);
  await kv(S.sentTag, ctx.sentTag);
  if (!ctx.isTest) {
    const list = (v, f = String) => (Array.isArray(v) ? txt(v.map(f).join(', ')) : txt(v));
    await kv(S.excludeTags, list(params.excludeTags));
    await kv(S.excludeDomains, list(params.excludeEmailDomains));
    await kv(S.excludeSources, list(params.excludeOrderSources, (source) => sourceText(source, T)));
  }
  await kv(S.currency, params.currency ?? ctx.config.giftCardCurrency);
  await kv(S.note, campaignNote(params.giftCardNote ?? ctx.config.giftCardNote, ctx.campaignId));
  await kv(S.templateSuffix, txt(params.giftCardTemplateSuffix ?? ctx.config.giftCardTemplateSuffix));

  // ---- funnel
  if (funnel) {
    await section(S.funnel);
    await s.header(S.funnelHeader);
    await trow([S.allCustomers, num(funnel.total), null], ['bold', 'count', 'text']);
    for (const r of objects(funnel.byRule)) {
      await trow([S.ruleRow(r.n, T.stored.ruleLabel(r, params)), num(r.count), UNLISTED_RULES.has(r.code) ? S.countedOnly : S.listed], ['text', 'count', 'text']);
    }
    await trow([S.finalRecipients, num(funnel.recipients), null], ['bold', 'count', 'text']);
    await s.add([S.funnelNote(fmtCount(funnel.listed), fmtCount(funnel.unlisted))], NOTE);
    await breakdown(S.byDomain, funnel.relayDomains);
    await breakdown(S.byMarketing, funnel.notSubscribed, (k) => labelOf(T.MARKETING_LABELS, k));
    await breakdown(S.byTag, funnel.tags);
    await breakdown(S.byChannel, funnel.channels, (k) => T.stored.channel(k));
  }

  // ---- same-address dedupe
  const d = stats.duplicates;
  if (d && !ctx.isTest) {
    await section(S.dedupe);
    await kv(S.groups, num(d.groups), 'count');
    await kv(S.accounts, num(d.accounts), 'count');
    await kv(S.removed, num(d.removed), 'count');
    await kv(S.bulk, num(d.flaggedBulk), 'count', S.bulkNote);
    if (Array.isArray(d.largest) && d.largest.length) await kv(S.largest, d.largest.join(T.sep.enum));
    await kv(S.noAddress, num(stats.noAddressRecipients), 'count');
  }

  // ---- amounts
  await section(S.amounts);
  await s.header(S.amountsHeader);
  if (ctx.isTest) {
    await trow([S.testTier(formatUsd(params.testGiftAmountCents ?? ctx.rows[0]?.r.amountCents)), null, null, num(stats.recipients ?? ctx.rows.length), money(stats.totalCents)], ['text', 'count', 'count', 'count', 'money']);
  } else {
    for (const t of objects(stats.tiers)) {
      await trow([txt(t.label ?? tierLabel(t.tier, ctx.tiersCents)), num(t.ordered), num(t.never), num(t.count), money(t.cents)], ['text', 'count', 'count', 'count', 'money']);
    }
    await trow([S.total, num(stats.orderedCount), num(stats.neverCount), num(stats.recipients), money(stats.totalCents)], ['bold', 'count', 'count', 'count', 'money']);
    const bf = stats.basisFromEarlier ?? {};
    const count = (v) => (typeof v === 'number' && Number.isFinite(v) ? v : 0);
    const why = [S.whyCancelled(fmtCount(count(bf.lastCancelled)))];
    if (count(bf.lastTest) > 0) why.push(S.whyTest(fmtCount(bf.lastTest))); // older lists have no lastTest
    why.push(S.whyZero(fmtCount(count(bf.lastZero))));
    await kv(S.orderedTotal, money(stats.orderedCents), 'money', S.personCount(fmtCount(stats.orderedCount)));
    await kv(S.neverTotal, money(stats.neverCents), 'money', S.personCount(fmtCount(stats.neverCount)));
    await kv(S.fromEarlier, S.personCount(fmtCount(count(bf.lastCancelled) + count(bf.lastTest) + count(bf.lastZero))), 'text', why.join(T.sep.comma));
    await kv(S.onlyCancelled, S.personCount(fmtCount(stats.neverReasons?.onlyCancelledOrZero)));
    await kv(S.median, money(stats.medianOrderedRawCents), 'money');
  }

  // ---- issuing progress
  const p = ctx.progress;
  await section(S.progress);
  await s.header(S.progressHeader);
  for (const [status, b] of p.byStatus) {
    await s.add([T.STATUS_LABELS[status] ?? status, b.count, money(b.cents)], [cellStyle('text', STATUS_FILLS[status] ?? null), cellStyle('count'), cellStyle('money')]);
  }
  await trow([S.total, p.total.count, money(p.total.cents)], ['bold', 'count', 'money']);
  // Only batches that created cards count (an issue --repair-only or date-blocked run creates none).
  await kv(S.lastBatch, ctx.journal.lastIssuedBatch ?? ctx.journal.lastBatch ?? null, 'int');
  await kv(S.nextSeq, p.nextPendingSeq ?? S.nothingPending, 'int');
  if (p.skipReasons.size) {
    await s.blank();
    await s.header(S.skipReasonsHeader);
    for (const [reason, n] of [...p.skipReasons].sort((a, b) => b[1] - a[1])) await trow([issueSkipText(reason, undefined, undefined, T) || S.reasonMissing, n], ['text', 'count']);
  }
  await s.add([S.progressNote], NOTE);

  // ---- usage
  if (ctx.usageSummary) {
    const u = ctx.usageSummary;
    await section(S.usage);
    await kv(S.asOf, ctx.time(ctx.usage.fetchedAt), 'datetime');
    await kv(S.issuedCards, u.issuedCards, 'count');
    await kv(S.issuedValue, money(u.issuedCents), 'money');
    await kv(S.usedCards, u.usedCards, 'count', u.usedRate, formatStyle(T.fmt.share));
    await kv(S.usedValue, money(u.usedCents), 'money', u.usedCentsRate, formatStyle(T.fmt.shareOfValue));
    await kv(S.orders, u.orders, 'count');
    await kv(S.ordersTotal, money(u.ordersTotalCents), 'money');
    await kv(S.perOrder, money(u.avgOrderCents), 'money');
    await kv(S.giftCardPaid, money(u.giftCardCents), 'money', T.USAGE_SPLIT_NOTE);
    await kv(S.customerPaid, money(u.customerPaidCents), 'money');
    if (u.topProducts.length) await kv(S.topProducts, txt(u.topProducts.slice(0, 3).map((x) => `${T.stored.productName(x.name)} × ${x.quantity}`).join(T.sep.list)));
    const ut = ctx.usedTagStats;
    await kv(S.usedTagged, ut.tagged, 'count', S.usedTagNote(fmtCount(ut.added), fmtCount(ut.failed)));
    await kv(S.dailyTrend, S.seeUsageReport);
  }

  // ---- run history
  await section(S.runs);
  await s.header(S.runsHeader);
  const runs = [...ctx.journal.runs].sort((a, b) => String(a.startedAt).localeCompare(String(b.startedAt)));
  const lastOpen = new Map(); // command → its newest run without run.end
  for (const run of runs) if (!run.endedAt) lastOpen.set(run.command, run);
  for (const run of runs) {
    let summary = summaryText(run.summary, { command: run.command, dryRun: run.dryRun }, T);
    if (!run.endedAt) summary = ctx.running === run.command && lastOpen.get(run.command) === run ? S.running : S.notEnded;
    await trow([ctx.time(run.startedAt), txt(run.command), run.dryRun ? S.dryRun : S.live, runBatchText(run), num(run.limit), num(run.exitCode), ctx.time(run.endedAt), txt(summary)], [
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
  if (!runs.length) await s.add([S.noRuns], NOTE);
  await s.finish();
}

/** "amazon" → "Amazon", "205641" → "Sellbrite（205641）" (in the pack's words). */
function sourceText(source, T) {
  const label = T.channelLabel(source);
  if (!label) return String(source ?? '');
  return label.toLowerCase() === String(source).toLowerCase() ? label : T.summary.sourceWithCode(label, source);
}

function runBatchText(run) {
  return Number.isInteger(run.batch) ? String(run.batch) : null;
}

// ---- 发放名单 (recipients) ---------------------------------------------------

/** Column definitions of the recipients sheet: header, width, cell kind, help text (help sheet) and value. */
function recipientColumns(ctx) {
  const { T } = ctx;
  const C = T.recipients.columns;
  /** Title and help of column `id`; a templated one gets `args` (the tag it names). */
  const text = (id, ...args) => {
    const pick = (v) => (typeof v === 'function' ? v(...args) : v);
    return { title: pick(C[id].title), help: pick(C[id].help) };
  };
  return [
    { ...text('seq'), width: 7, kind: 'int', value: (x) => num(x.r.seq) },
    { ...text('status'), width: 13, kind: 'text', value: (x) => T.STATUS_LABELS[x.status] ?? x.status },
    { ...text('batch'), width: 6, kind: 'int', value: (x) => num(x.state?.batch) },
    { ...text('customerId'), width: 16, kind: 'link', value: (x, c) => c.link('customers', x.r.customerId, x.r.numericId) },
    { ...text('name'), width: 18, kind: 'text', value: (x) => txt(x.r.name) },
    { ...text('email'), width: 28, kind: 'text', value: (x) => txt(x.r.email) },
    { ...text('marketing'), width: 10, kind: 'text', value: (x) => labelOf(T.MARKETING_LABELS, x.r.marketingState) || null },
    { ...text('kind'), width: 10, kind: 'text', value: (x) => labelOf(T.KIND_LABELS, x.r.kind) || null },
    { ...text('basisOrder'), width: 13, kind: 'link', value: (x, c) => (x.r.basis ? c.link('orders', x.r.basis.orderId, x.r.basis.orderName) : null) },
    { ...text('basisDate'), width: 12, kind: 'date', value: (x, c) => (x.r.basis ? c.day(x.r.basis.createdAt) : null) },
    { ...text('channel'), width: 13, kind: 'text', value: (x) => txt(T.channelLabel(x.r.lastOrderSource)) },
    { ...text('daysAgo'), width: 8, kind: 'int', value: (x, c) => (x.r.basis ? c.daysToExport(x.r.basis.createdAt) : null) },
    { ...text('basisTotal'), width: 12, kind: 'money', value: (x) => (x.r.basis ? money(x.r.basis.totalCents) : null) },
    { ...text('formula'), width: 42, kind: 'text', value: (x, c) => txt(T.stored.formula(x.r, c.selection)) },
    { ...text('tier'), width: 9, kind: 'text', value: (x, c) => c.tierText(x.r) },
    { ...text('amount'), width: 11, kind: 'money', value: (x) => money(x.cardCents) },
    { ...text('groupSize'), width: 8, kind: 'int', value: (x) => (x.r.addressKey ? num(x.r.groupSize ?? 1) : null) },
    { ...text('groupOthers'), width: 18, kind: 'text', value: (x) => (Array.isArray(x.r.groupOthers) ? txt(x.r.groupOthers.join(', ')) : txt(x.r.groupOthers)) },
    { ...text('hasTag', ctx.sentTag), width: 14, kind: 'text', value: (x, c) => (c.hasTag(x.state, x.r.customerId) ? T.yes : T.no) },
    { ...text('giftCardId'), width: 16, kind: 'link', value: (x, c) => (x.state?.giftCardId ? c.link('gift_cards', x.state.giftCardId) : null) },
    { ...text('last4'), width: 9, kind: 'text', value: (x) => txt(x.state?.last4) },
    { ...text('createdAt'), width: 16, kind: 'datetime', value: (x, c) => c.time(x.state?.createdAt) },
    { ...text('taggedAt'), width: 16, kind: 'datetime', value: (x, c) => c.time(x.state?.taggedAt) },
    { ...text('usedAmount'), width: 11, kind: 'money', value: (x) => (x.card ? money(x.card.usedCents) : null) },
    { ...text('balance'), width: 11, kind: 'money', value: (x) => (x.card ? money(x.card.balanceCents) : null) },
    { ...text('usedOrders'), width: 16, kind: 'text', value: (x, c) => c.ordersText(x.card) },
    { ...text('usedTag', ctx.usedTag), width: 10, kind: 'text', value: (x, c) => (c.hasUsedTag(x.state, x.r.customerId) ? T.yes : T.no) },
    { ...text('notes'), width: 44, kind: 'text', value: (x, c) => c.notesFor(x) },
    { ...text('city'), width: 14, kind: 'text', value: (x) => txt(x.r.city) },
    { ...text('province'), width: 6, kind: 'text', value: (x) => txt(x.r.provinceCode) },
    { ...text('zip'), width: 9, kind: 'text', value: (x) => txt(x.r.zip) },
    { ...text('orderCount'), width: 7, kind: 'int', value: (x) => num(x.r.numberOfOrders) },
    { ...text('amountSpent'), width: 12, kind: 'money', value: (x) => money(x.r.amountSpentCents) },
    { ...text('accountCreated'), width: 12, kind: 'date', value: (x, c) => c.day(x.r.accountCreatedAt) },
  ];
}

async function writeRecipientsSheet(book, ctx) {
  const columns = recipientColumns(ctx);
  const s = book.sheet(ctx.T.sheets.recipients, { widths: columns.map((c) => c.width), freezeRows: 1, freezeCols: 3 });
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

// ---- 未入选 (not selected) ---------------------------------------------------

async function writeNotSelectedSheet(book, ctx) {
  const { T } = ctx;
  const titles = T.notSelected.header;
  const s = book.sheet(T.sheets.notSelected, { widths: [16, 18, 28, 10, 38, 56, 12, 7, 16, 8, 30] });
  const header = await s.header(titles);
  const rowStyles = ['link', 'text', 'text', 'text', 'text', 'text', 'date', 'int', 'link', 'text', 'text'].map((k) => cellStyle(k));
  const info = { params: ctx.params, timeZone: ctx.timeZone };
  for (const n of ctx.notSelected) {
    await s.add(
      [
        ctx.link('customers', n.customerId, n.numericId),
        txt(n.name),
        txt(n.email),
        labelOf(T.MARKETING_LABELS, n.marketingState) || null,
        txt(T.stored.primaryReason(n, info)),
        txt(T.stored.allReasons(n, info)),
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

// ---- 同地址重复 (same address) -------------------------------------------------

async function writeDuplicatesSheet(book, ctx) {
  const { T } = ctx;
  const D = T.duplicates;
  const titles = D.header;
  const kinds = ['text', 'int', 'text', 'text', 'link', 'text', 'text', 'date', 'int', 'date', 'text'];
  const s = book.sheet(T.sheets.duplicates, { widths: [15, 12, 40, 16, 18, 18, 28, 14, 8, 12, 12] });
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
          m.kept ? D.kept : D.notSent,
          ctx.link('customers', m.customerId, m.numericId),
          txt(m.name),
          txt(m.email),
          ctx.day(m.lastPaidOrderAt),
          num(m.numberOfOrders),
          ctx.day(m.accountCreatedAt),
          g.flaggedBulk ? T.yes : null,
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
  await s.add([D.activeTitle(ctx.inactiveMonths, fmtCount(active.length))], SECTION, { height: 22 });
  await s.header(D.activeHeader);
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

// ---- 使用明细 (usage details) ------------------------------------------------

function itemsText(lineItems, T) {
  const items = objects(lineItems);
  if (!items.length) return null;
  return txt(items.map((l) => `${T.stored.productName(l.name ?? '')} × ${l.quantity ?? 1}`).join('; '));
}

async function writeUsageDetailSheet(book, ctx) {
  const { T } = ctx;
  const D = T.usageDetail;
  const u = ctx.usage;
  const s = book.sheet(T.sheets.usageDetail, { widths: [12, 17, 16, 17, 12, 14, 12, 8, 60] });
  const titles = D.header;
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
        order?.cancelled ? T.yes : null,
        itemsText(order?.lineItems, T),
      ],
      rowStyles,
    );
  }
  s.setFilter(header, titles.length);

  const unmatched = objects(u.unmatched);
  await s.blank();
  await s.blank();
  if (!unmatched.length) {
    await s.add([D.noUnmatched], NOTE);
  } else {
    await s.add([D.unmatchedTitle(fmtCount(unmatched.length))], SECTION, { height: 22 });
    await s.header(D.unmatchedHeader);
    const unmatchedStyles = ['link', 'datetime', 'link', 'datetime', 'money', 'text'].map((k) => cellStyle(k));
    for (const m of unmatched) {
      await s.add(
        [
          ctx.link('orders', m.orderId, m.orderName),
          ctx.time(m.orderCreatedAt),
          ctx.link('customers', m.customerId),
          ctx.time(m.processedAt),
          money(m.amountCents),
          txt(storedLabelOf(T.UNMATCHED_REASON_LABELS, m.reason, T)),
        ],
        unmatchedStyles,
      );
    }
  }
  await s.finish();
}

// ---- 核对 (verify) -------------------------------------------------------------

async function writeVerifySheet(book, ctx) {
  const { T } = ctx;
  const V = T.verify;
  const v = ctx.verify;
  const issues = objects(v.issues);
  const s = book.sheet(T.sheets.verify, { widths: [30, 16, 16, 36, 36, 46], freezeRows: 3 });
  await s.add([V.title(ctx.stamp(v.verifiedAt) || V.unknownTime, ctx.tzLabel, fmtCount(v.cardCount), ctx.sentTag, fmtCount(v.taggedCount))], SECTION, { height: 22 });
  const counts = Object.entries(v.counts ?? {})
    .filter(([, n]) => n > 0)
    .map(([type, n]) => V.countItem(labelOf(T.VERIFY_TYPE_LABELS, type), fmtCount(n)))
    .join(T.sep.list);
  await s.add([issues.length ? V.found(fmtCount(issues.length), counts) : V.none], NOTE);
  const titles = V.header;
  const header = await s.header(titles);
  const rowStyles = ['text', 'link', 'link', 'text', 'text', 'text'].map((k) => cellStyle(k));
  for (const i of issues) {
    await s.add(
      [
        txt(labelOf(T.VERIFY_TYPE_LABELS, i.type)),
        i.customerId ? ctx.link('customers', i.customerId) : null,
        i.giftCardId ? ctx.link('gift_cards', i.giftCardId) : null,
        txt(T.stored.verify(i, 'journal')),
        txt(T.stored.verify(i, 'shopify')),
        txt(T.stored.verify(i, 'action')),
      ],
      rowStyles,
    );
  }
  s.setFilter(header, titles.length);
  await s.finish();
}

// ---- 操作日志 (activity log) ---------------------------------------------------

/**
 * "结果/说明" of one journal entry. Skip details are made readable: times in
 * store time, customers by number, marketing states by label (ctx.stamp is
 * the workbook's store-local clock).
 */
function entryResult(e, ctx) {
  const { T } = ctx;
  const J = T.journal;
  const amount = Number.isInteger(e.amountCents) ? J.amount(formatUsd(e.amountCents)) : '';
  const last4 = e.last4 ? J.last4(e.last4) : '';
  const detail = { stamp: ctx.stamp };
  switch (e.op) {
    case 'create.start':
      return amount;
    case 'create.ok':
      return [amount, last4].filter(Boolean).join(T.sep.comma);
    case 'reconcile.found':
      return [labelOf(T.RECONCILE_SOURCE_LABELS, e.source), amount, last4].filter(Boolean).join(T.sep.comma);
    case 'skip':
      return issueSkipText(e.reason, e.detail, detail, T);
    case 'used.tag.ok':
      return J.usedTagAdded(ctx.usedTag);
    case 'used.tag.fail':
      return J.usedTagFailed(ctx.usedTag);
    default:
      return noteText(e.note, T);
  }
}

async function writeJournalSheet(book, ctx) {
  const { T } = ctx;
  const titles = T.journal.header;
  const s = book.sheet(T.sheets.journal, { widths: [16, 9, 12, 16, 18, 44, 16, 44] });
  const header = await s.header(titles);
  const rowStyles = ['datetime', 'text', 'text', 'text', 'text', 'text', 'text', 'warn'].map((k) => cellStyle(k));
  for (const e of ctx.entries) {
    if (!e || typeof e.op !== 'string' || e.op.startsWith('run.')) continue;
    const run = ctx.runsById.get(e.run);
    const batch = Number.isInteger(e.batch ?? run?.batch) ? String(e.batch ?? run.batch) : null;
    // Entries whose run.start is missing: the op itself tells usage and issue
    // apart; reconcile.* can also come from verify, so it stays blank.
    const command = run?.command ?? (/^used\.tag\./.test(e.op) ? 'usage' : /^(create|tag)\.|^skip$/.test(e.op) ? 'issue' : null);
    await s.add(
      [
        ctx.time(e.t),
        txt(command),
        batch,
        txt(numericId(e.cid) || e.cid),
        labelOf(T.OP_LABELS, e.op),
        txt(entryResult(e, ctx)),
        txt(numericId(e.giftCardId)),
        txt(errorText(e.error, T)),
      ],
      rowStyles,
    );
  }
  s.setFilter(header, titles.length);
  await s.finish();
}

// ---- 说明 (help) ---------------------------------------------------------------

async function writeHelpSheet(book, ctx) {
  const { T } = ctx;
  const H = T.help;
  const s = book.sheet(T.sheets.help, { widths: [30, 120] });
  const item = cellStyle('text');
  const section = async (title) => {
    await s.blank();
    await s.add([title], SECTION, { height: 22 });
  };
  const rows = async (pairs) => {
    for (const [k, v] of pairs) await s.add([txt(k), txt(v)], [item, item]);
  };
  await s.header(H.header);

  await section(H.sheetsSection);
  await rows(H.sheets);

  await section(H.columnsSection);
  await rows(recipientColumns(ctx).map((c) => [c.title, c.help]));

  await section(H.statusSection);
  for (const status of STATUS_ORDER) {
    await s.add([T.STATUS_LABELS[status], T.STATUS_HELP[status]], [cellStyle('text', STATUS_FILLS[status] ?? null), item]);
  }
  await rows([[H.issueSkipReasons, Object.values(T.ISSUE_SKIP_LABELS).join(T.sep.enum)]]);

  await section(H.reminderSection);
  await rows(reminderHelp(ctx.sentTag, T));

  await section(H.amountSection);
  const percent = ctx.params.giftPercent ?? ctx.config.giftPercent;
  if (ctx.isTest) {
    await rows([[H.testCampaign, H.testAmount(formatUsd(ctx.params.testGiftAmountCents ?? ctx.rows[0]?.r.amountCents))]]);
  } else {
    await rows([
      [H.base, H.baseText(percent)],
      [H.never, H.neverText(percent)],
      ...(ctx.tiersCents.length ? T.describeTiers(ctx.tiersCents).map((t, i) => H.tierRow(i, t)) : []),
      [H.roundFirst, H.roundFirstText],
    ]);
  }

  await section(H.usageSection);
  await rows([[H.usageAmounts, T.USAGE_SPLIT_NOTE]]);

  await section(H.otherSection);
  await rows([
    [H.time, H.timeText(ctx.timeZone, ctx.tzLabel)],
    [H.links, H.linksText],
    [H.readOnly, H.readOnlyText],
    [H.filter, H.filterText],
    ...T.EXIT_CODE_HELP.map(([code, text]) => [H.exitCode(code), text]),
  ]);
  await s.finish();
}
