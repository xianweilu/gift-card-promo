// Local preview of the "Gift card created" notification (subject + body).
//
// The two templates in notifications/ are pasted into the Shopify admin by
// hand and Shopify renders them when it sends the email. This module renders
// the same templates locally with liquidjs so the copy can be checked before
// pasting: the three promo stages (first email, reminder 1, reminder 2) and
// the original email for gift cards that are not part of the campaign.
//
// The templates cannot read .env, so the two reminder dates are also written
// into a fixed block at the top of each template:
//   {%- assign promo_remind_1 = 20261012 -%}{%- assign promo_remind_2 = 20261016 -%}
// syncTemplateDates() rewrites exactly those numbers from REMIND_1_DATE /
// REMIND_2_DATE and leaves every other byte of the templates alone.
//
// Nothing here contacts Shopify.

import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { Liquid } from 'liquidjs';
import { ROOT_DIR } from './config.js';
import { campaignPaths, readJson, writeFileAtomic } from './campaign.js';
import { localDate, localDateTime } from './time.js';
import { formatUsd } from './select/amount.js';

/** The emails that can be previewed, in display order. */
export const VARIANTS = Object.freeze(['first', 'remind1', 'remind2', 'original']);

export const DEFAULT_TEMPLATES_DIR = path.join(ROOT_DIR, 'notifications');

export const TEMPLATE_FILES = Object.freeze({
  subject: 'gift-card-created.subject.liquid',
  body: 'gift-card-created.body.liquid',
});

/** Sample data used when no recipient from the selection is given. */
export const SAMPLE_FIRST_NAME = 'Alex';
export const SAMPLE_AMOUNT_CENTS = 1533;

const DEFAULT_TEMPLATE_SUFFIX = 'gift-card-promo';
// The store's zone; campaign dates in .env are calendar days there.
const STORE_TIMEZONE = 'America/Los_Angeles';
const SHOPIFY_ASSET_BASE = 'https://cdn.shopify.com/shopifycloud/shopify/assets/';
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

/** Chinese names of the variants / template stages (they share the same keys). */
const LABELS = Object.freeze({ first: '首封', remind1: '第一次提醒', remind2: '第二次提醒', original: '原版' });

// `assign promo_remind_1 = 20261012` → [prefix, round ('1' | '2'), value].
// The value runs until whitespace or the closing `-%}` / `%}` of the tag, so a
// malformed value (e.g. '2026-10-12' or a quoted string) is captured and reported.
const REMIND_ASSIGNMENT = /(\bassign\s+promo_remind_([12])\s*=\s*)((?:(?!-?%\})\S)*)/g;

// The body template prints which copy it chose, e.g. <!-- promo-stage: remind1 -->.
const STAGE_MARKER = /<!--\s*promo-stage:\s*([\w-]+)\s*-->/;

// ---------------------------------------------------------------------------
// Template dates ⇄ .env
// ---------------------------------------------------------------------------

/**
 * Make the reminder dates written in both templates match REMIND_1_DATE /
 * REMIND_2_DATE. Every `assign promo_remind_1 = <YYYYMMDD>` and
 * `assign promo_remind_2 = <YYYYMMDD>` is set to the configured date; nothing
 * else in the files changes.
 *
 * Nothing is written when there is any problem (a date missing from .env, a
 * template missing, an assignment missing or not an 8-digit number).
 *
 * @param {object} a
 * @param {object} a.config loadConfig() result (remind1Date, remind2Date)
 * @param {string} [a.templatesDir] folder holding the two templates (default notifications/)
 * @param {boolean} [a.write] false = only report what differs
 * @param {object} [a.log]
 * @returns {{ changed: boolean, files: string[], problems: string[], changes: {file, name, from, to}[] }}
 *   changed/files: templates whose dates differ from .env (rewritten when `write`).
 */
export function syncTemplateDates({ config, templatesDir = DEFAULT_TEMPLATES_DIR, write = true, log = console } = {}) {
  const problems = [];

  const wanted = {}; // round → 'YYYYMMDD'
  for (const [round, key, value] of [['1', 'REMIND_1_DATE', config.remind1Date], ['2', 'REMIND_2_DATE', config.remind2Date]]) {
    if (value && DATE_RE.test(value)) wanted[round] = value.replaceAll('-', '');
    else problems.push(`.env 里没有设置 ${key}（格式 YYYY-MM-DD），无法更新邮件模板里的提醒日期`);
  }

  const templates = [];
  for (const name of Object.values(TEMPLATE_FILES)) {
    const file = path.join(templatesDir, name);
    let text;
    try {
      text = fs.readFileSync(file, 'utf8');
    } catch (err) {
      problems.push(err.code === 'ENOENT' ? `找不到邮件模板：${file}` : `读不了邮件模板 ${file}：${err.message}`);
      continue;
    }
    const found = { 1: 0, 2: 0 };
    for (const m of text.matchAll(REMIND_ASSIGNMENT)) {
      found[m[2]] += 1;
      if (!/^\d{8}$/.test(m[3])) {
        problems.push(`邮件模板 ${file} 里 promo_remind_${m[2]} 的值 "${m[3]}" 不是 8 位日期（应为 YYYYMMDD，例如 20261012）`);
      }
    }
    for (const round of ['1', '2']) {
      if (!found[round]) {
        problems.push(`邮件模板 ${file} 里找不到 {%- assign promo_remind_${round} = YYYYMMDD -%}：请恢复模板开头的活动日期区块`);
      }
    }
    templates.push({ file, text, next: text });
  }

  if (problems.length) return { changed: false, files: [], problems, changes: [] };

  const changes = [];
  for (const t of templates) {
    t.next = t.text.replace(REMIND_ASSIGNMENT, (all, prefix, round, value) => {
      if (value !== wanted[round]) changes.push({ file: t.file, name: `promo_remind_${round}`, from: value, to: wanted[round] });
      return `${prefix}${wanted[round]}`;
    });
  }
  const outdated = templates.filter((t) => t.next !== t.text);
  const changed = outdated.length > 0;

  if (changed && write) {
    for (const t of outdated) writeFileAtomic(t.file, t.next);
    log.warn('模板已按 .env 更新，请把主题和正文重新贴到 Shopify 后台');
    for (const c of changes) log.info(`  ${path.basename(c.file)}：${c.name} ${c.from} → ${c.to}`);
  }
  return { changed, files: outdated.map((t) => t.file), problems, changes };
}

// ---------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------

/** "YYYY-MM-DD" today in the store's time zone. */
function storeToday(config, now) {
  return localDate(now().getTime(), config.timezone || STORE_TIMEZONE);
}

/**
 * The day a variant is pretended to be sent: first / original on LAUNCH_DATE
 * (today when it is not set), the reminders on REMIND_1_DATE / REMIND_2_DATE.
 */
export function simulatedSendDate(config, variant, now = () => new Date()) {
  switch (variant) {
    case 'first':
    case 'original':
      return config.launchDate || storeToday(config, now);
    case 'remind1':
      if (!config.remind1Date) throw new Error('.env 里没有设置 REMIND_1_DATE，无法预览第一次提醒');
      return config.remind1Date;
    case 'remind2':
      if (!config.remind2Date) throw new Error('.env 里没有设置 REMIND_2_DATE，无法预览第二次提醒');
      return config.remind2Date;
    default:
      throw new Error(unknownVariantMessage(variant));
  }
}

/**
 * The stage the template should show for `variant` sent on `date` when its
 * dates agree with .env. Without both reminder dates in .env the variant's
 * own stage is assumed.
 */
export function expectedStage(config, variant, date) {
  if (variant === 'original') return 'original';
  if (!config.remind1Date || !config.remind2Date) return variant;
  if (date >= config.remind2Date) return 'remind2';
  if (date >= config.remind1Date) return 'remind1';
  return 'first';
}

/** Shopify-style money from integer cents: 1077 → "$10.77"; trimZeros: 1000 → "$10". Blank for non-numbers. */
function formatMoney(cents, { trimZeros = false } = {}) {
  if (cents === null || cents === undefined || cents === '') return '';
  const n = Number(cents);
  if (!Number.isFinite(n)) return '';
  const whole = Math.round(n);
  if (trimZeros && whole % 100 === 0) {
    return `${whole < 0 ? '-' : ''}$${(Math.abs(whole) / 100).toLocaleString('en-US')}`;
  }
  return formatUsd(whole);
}

/**
 * A liquidjs engine that renders like Shopify's notification renderer would
 * on `sendDate` (YYYY-MM-DD, a calendar day in the store's zone).
 *
 * Dates: JavaScript parses a date-only string such as gift_card.expires_on
 * ('2026-10-19') as UTC midnight. With liquidjs' defaults it is then shown in
 * this machine's zone, which in Los Angeles is October 18 — one day early.
 * Rendering every date in UTC (timezoneOffset 0) keeps calendar dates on their
 * day whatever zone the machine is in. 'now' / 'today' are pinned to 19:00 UTC
 * on sendDate, which is around noon in Los Angeles, so the template's
 * `'now' | date: '%Y%m%d'` gives sendDate both in UTC and in the store's zone.
 * locale en-US keeps month names English on a machine set to another language.
 */
export function createPreviewEngine({ sendDate, currency = 'USD' }) {
  if (!DATE_RE.test(String(sendDate ?? ''))) throw new Error(`模拟发送日期必须是 YYYY-MM-DD，收到 "${sendDate}"`);
  const simulatedNow = new Date(`${sendDate}T19:00:00Z`);
  const engine = new Liquid({
    timezoneOffset: 0,
    locale: 'en-US',
    strictVariables: false, // unknown variables render empty, as in Shopify
    strictFilters: false,
  });

  const builtinDate = engine.filters.date;
  engine.registerFilter('date', function date(input, ...args) {
    // Like Ruby Liquid (Shopify), 'now' / 'today' match in any letter case.
    const isNow = typeof input === 'string' && ['now', 'today'].includes(input.toLowerCase());
    return builtinDate.call(this, isNow ? simulatedNow : input, ...args);
  });

  // Money filters take integer cents, like Shopify's.
  engine.registerFilter('money', (cents) => formatMoney(cents));
  engine.registerFilter('money_with_currency', (cents) => {
    const money = formatMoney(cents);
    return money ? `${money} ${currency}` : '';
  });
  engine.registerFilter('money_without_trailing_zeros', (cents) => formatMoney(cents, { trimZeros: true }));
  engine.registerFilter('money_without_currency', (cents) => formatMoney(cents).replace('$', ''));
  engine.registerFilter('shopify_asset_url', (file) => `${SHOPIFY_ASSET_BASE}${String(file ?? '').replace(/^\/+/, '')}`);
  return engine;
}

/**
 * The expiry date a preview shows: `expiresOn` when the caller passes one (also
 * '' / null, meaning the cards have no expiry date), else GIFT_CARD_EXPIRES_ON.
 * Callers with a selection pass selection.params.giftCardExpiresOn: the cards
 * are created with that frozen date, even when .env was changed afterwards.
 * Returns 'YYYY-MM-DD' or null (null, not '': an empty string is truthy in Liquid).
 */
export function previewExpiresOn(config, expiresOn = undefined) {
  const value = expiresOn === undefined ? config.giftCardExpiresOn : expiresOn;
  if (value === null || value === undefined || value === '') return null;
  if (!DATE_RE.test(String(value))) throw new Error(`礼品卡到期日必须是 YYYY-MM-DD，收到 "${value}"`);
  return String(value);
}

/**
 * The Liquid variables of one preview. `recipient` is a selection.json
 * recipient (its first name and amount are used); without one, sample data.
 * `expiresOn`: see previewExpiresOn (omitted = GIFT_CARD_EXPIRES_ON from .env).
 */
export function previewData(config, variant, recipient = null, { expiresOn } = {}) {
  const cents = Number.isInteger(recipient?.amountCents) && recipient.amountCents > 0 ? recipient.amountCents : SAMPLE_AMOUNT_CENTS;
  return {
    shop: {
      name: 'LA Balloons',
      url: 'https://www.laballoons.com',
      // null, not '': an empty string is truthy in Liquid and would render an <img src="">
      // instead of the shop-name heading Shopify shows when no email logo is set.
      email_logo_url: null,
      email_logo_width: 150,
      email_accent_color: '#1990C6',
    },
    gift_card: {
      template_suffix: variant === 'original' ? '' : config.giftCardTemplateSuffix || DEFAULT_TEMPLATE_SUFFIX,
      balance: cents,
      initial_value: cents,
      expires_on: previewExpiresOn(config, expiresOn),
      url: '#',
      pass_url: null,
      // A real recipient keeps their real (possibly empty) first name, so the
      // preview shows exactly what Shopify would print.
      customer: { first_name: recipient ? recipient.firstName ?? '' : SAMPLE_FIRST_NAME },
    },
    order_name: '',
  };
}

/** The stage the body template chose (from its promo-stage marker), or null. */
export function parseStage(html) {
  return STAGE_MARKER.exec(String(html ?? ''))?.[1] ?? null;
}

function readTemplate(templatesDir, name) {
  const file = path.join(templatesDir, name);
  try {
    return fs.readFileSync(file, 'utf8');
  } catch (err) {
    if (err.code === 'ENOENT') throw new Error(`找不到邮件模板：${file}`);
    throw err;
  }
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
}

function unknownVariantMessage(variant) {
  return `未知的预览类型 "${variant}"。可选：first（首封）、remind1（第一次提醒）、remind2（第二次提醒）、original（原版），不写则四种全部生成`;
}

function describeRecipient(recipient, firstName) {
  if (!recipient) return `示例顾客 ${firstName}`;
  return `名单第 ${recipient.seq} 号顾客 ${firstName || '（没有名字）'}`;
}

/** The small notice shown at the top of a preview page (not part of the real email). */
function bannerHtml({ subject, sendDate, variant, recipient, data }) {
  const firstName = data.gift_card.customer.first_name;
  const style = [
    'font-family: -apple-system, BlinkMacSystemFont, \'PingFang SC\', \'Helvetica Neue\', Arial, sans-serif',
    'font-size: 13px',
    'line-height: 1.6',
    'text-align: left',
    'background: #fff8e1',
    'color: #333',
    'border-bottom: 1px solid #e0c97f',
    'padding: 8px 12px',
  ].join('; ');
  return `<div class="gift-card-promo-preview" style="${style}">`
    + `<strong>本地预览</strong>　主题：${escapeHtml(subject)}，模拟发送日期：${escapeHtml(sendDate)}，变体：${escapeHtml(variant)}（${LABELS[variant]}）<br>`
    + `${escapeHtml(describeRecipient(recipient, firstName))}，礼品卡金额 ${formatUsd(data.gift_card.balance)}。`
    + 'Shopify 后台的邮件样式表不在本地，排版以测试活动收到的真实邮件为准。'
    + '</div>';
}

/** The rendered body with the preview banner right after <body> (at the top when there is no <body>). */
function withBanner(html, banner) {
  const body = /<body\b[^>]*>/i.exec(html);
  if (!body) return `${banner}\n${html}`;
  const at = body.index + body[0].length;
  return `${html.slice(0, at)}\n${banner}${html.slice(at)}`;
}

/**
 * Render one email to `paths.previewDir/<variant>.html`.
 *
 * @param {object} a
 * @param {object} a.config
 * @param {object} [a.paths] campaignPaths(config)
 * @param {'first'|'remind1'|'remind2'|'original'} a.variant
 * @param {object|null} [a.recipient] a selection.json recipient (first name + amount), or sample data
 * @param {string} [a.templatesDir]
 * @param {object} [a.log]
 * @param {() => Date} [a.now] only used for "today" when LAUNCH_DATE is not set
 * @param {string|null} [a.sendDate] YYYY-MM-DD to render as if sent that day (default: the variant's day)
 * @param {string|null} [a.expiresOn] the cards' expiry date to show, overriding GIFT_CARD_EXPIRES_ON
 *   (callers with a selection pass selection.params.giftCardExpiresOn; '' / null = no expiry date)
 * @param {boolean} [a.checkDates] also warn when the templates' reminder dates differ from .env
 * @returns {Promise<{ file: string, subject: string, stage: string|null, expectedStage: string, sendDate: string, variant: string, expiresOn: string|null }>}
 *   stage: what the template actually showed; expectedStage: what .env's dates call for on sendDate;
 *   expiresOn: the expiry date the email showed (null = none).
 */
export async function renderPreviewFor({
  config,
  paths = campaignPaths(config),
  variant,
  recipient = null,
  templatesDir = DEFAULT_TEMPLATES_DIR,
  log = console,
  now = () => new Date(),
  sendDate = null,
  expiresOn = undefined,
  checkDates = true,
} = {}) {
  if (!VARIANTS.includes(variant)) throw new Error(unknownVariantMessage(variant));
  const date = sendDate ?? simulatedSendDate(config, variant, now);
  if (!DATE_RE.test(String(date))) throw new Error(`模拟发送日期必须是 YYYY-MM-DD，收到 "${date}"`);
  const expiry = previewExpiresOn(config, expiresOn); // validated before anything is read or written

  if (checkDates) {
    // Read-only: only `preview` itself rewrites the templates.
    const check = syncTemplateDates({ config, templatesDir, write: false, log });
    for (const problem of check.problems) log.warn(`邮件模板的提醒日期有问题：${problem}`);
    if (check.changed) {
      log.warn('模板里的提醒日期和 .env 不一致：请先运行 node index.js preview 更新模板，再把主题和正文重新贴到 Shopify 后台');
    }
  }

  const subjectTemplate = readTemplate(templatesDir, TEMPLATE_FILES.subject);
  const bodyTemplate = readTemplate(templatesDir, TEMPLATE_FILES.body);
  const engine = createPreviewEngine({ sendDate: date, currency: config.giftCardCurrency || 'USD' });
  const data = previewData(config, variant, recipient, { expiresOn: expiry });

  // A subject is a single line: collapse the template's whitespace and newlines.
  const subject = (await engine.parseAndRender(subjectTemplate, data)).replace(/\s+/g, ' ').trim();
  const body = await engine.parseAndRender(bodyTemplate, data);

  const stage = parseStage(body);
  const expected = expectedStage(config, variant, date);
  if (stage !== expected) {
    const shown = stage ? `${LABELS[stage] ?? stage}（${stage}）` : '没有 promo-stage 标记';
    const suffixHint = stage === 'original' && variant !== 'original'
      ? '；也可能是 .env 的 GIFT_CARD_TEMPLATE_SUFFIX 和模板里的 promo 后缀不一致'
      : '';
    log.warn(`预览的文案和预期不一致：模板里的日期可能和 .env 不一致（${LABELS[variant]}，模拟 ${date} 发送，预期${LABELS[expected]}（${expected}），模板显示${shown}${suffixHint}）`);
  } else if (expected !== variant) {
    log.info(`按 .env 的日期，${date} 发出的是${LABELS[expected]}文案（${expected}）`);
  }

  if (recipient && !data.gift_card.customer.first_name) {
    log.warn(`名单第 ${recipient.seq} 号顾客没有名字，邮件里会显示 "Hi ,"`);
  }

  const file = path.join(paths.previewDir, `${variant}.html`);
  writeFileAtomic(file, withBanner(body, bannerHtml({ subject, sendDate: date, variant, recipient, data })));
  return { file, subject, stage, expectedStage: expected, sendDate: date, variant, expiresOn: data.gift_card.expires_on };
}

// ---------------------------------------------------------------------------
// Overview page
// ---------------------------------------------------------------------------

/**
 * The overview page. `expiresOn` is the expiry date the previews showed (null = none);
 * `fromList` says it is the one frozen in the list rather than GIFT_CARD_EXPIRES_ON.
 */
function indexHtml({ config, results, recipient, sync, generatedAt, timezone, expiresOn, fromList }) {
  const sample = previewData(config, 'first', recipient, { expiresOn }); // the name and amount every page used
  const who = describeRecipient(recipient, sample.gift_card.customer.first_name);
  const envExpiry = config.giftCardExpiresOn || null;
  const expiryNote = fromList && expiresOn !== envExpiry
    ? `（按名单生成时的日期，卡按这个日期建；.env 现在是 ${envExpiry || '（未设置）'}）`
    : '';
  const rows = results.map((r) => {
    const mismatch = r.stage !== r.expectedStage
      ? `<br><span class="warn">⚠ 模板显示的是${escapeHtml(LABELS[r.stage] ?? r.stage ?? '（没有 promo-stage 标记）')}，按 .env 应为${escapeHtml(LABELS[r.expectedStage])}</span>`
      : '';
    const note = r.variant === 'original' ? '<br><span class="note">不是本活动的礼品卡收到的邮件</span>' : '';
    const name = path.basename(r.file);
    return `      <tr><td>${LABELS[r.variant]}（${r.variant}）${note}${mismatch}</td><td>${escapeHtml(r.sendDate)}</td>`
      + `<td>${escapeHtml(r.subject)}</td><td><a href="${escapeHtml(encodeURIComponent(name))}">${escapeHtml(name)}</a></td></tr>`;
  }).join('\n');
  const updated = sync.changed
    ? '\n  <p class="warn">模板已按 .env 更新，请把主题和正文重新贴到 Shopify 后台。</p>'
    : '';
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>邮件预览 · ${escapeHtml(config.campaignId)}</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Helvetica Neue", Arial, sans-serif; margin: 24px; color: #222; background: #fff; }
    h1 { font-size: 20px; margin: 0 0 12px; }
    p { margin: 6px 0; font-size: 14px; line-height: 1.6; }
    table { border-collapse: collapse; margin: 16px 0; font-size: 14px; }
    th, td { border: 1px solid #ddd; padding: 8px 10px; text-align: left; vertical-align: top; }
    th { background: #f5f5f5; }
    .warn { color: #b45309; }
    .note { color: #666; }
  </style>
</head>
<body>
  <h1>邮件预览（活动 ${escapeHtml(config.campaignId)}）</h1>
  <p>生成于 ${escapeHtml(generatedAt)}（${escapeHtml(timezone)}）。${escapeHtml(who)}，礼品卡金额 ${formatUsd(sample.gift_card.balance)}。</p>
  <p>模板里的提醒日期：第一次提醒 ${escapeHtml(config.remind1Date)}，第二次提醒 ${escapeHtml(config.remind2Date)}，与 .env 一致。礼品卡到期日：${escapeHtml(expiresOn || '（未设置）')}${escapeHtml(expiryNote)}。</p>${updated}
  <table>
    <thead>
      <tr><th>变体</th><th>模拟发送日期</th><th>主题</th><th>文件链接</th></tr>
    </thead>
    <tbody>
${rows}
    </tbody>
  </table>
  <p class="note">Shopify 后台的邮件样式表不在本地，排版和真实邮件会有小差别。预览用来核对文案和三种切换；真实样子以测试活动收到的邮件为准。</p>
</body>
</html>
`;
}

// ---------------------------------------------------------------------------
// The preview command
// ---------------------------------------------------------------------------

/** Open a file with the system's default application; failures are ignored (the paths are printed anyway). */
export function openInBrowser(file) {
  const [command, args] = process.platform === 'darwin'
    ? ['open', [file]]
    : process.platform === 'win32'
      ? ['explorer', [file]]
      : ['xdg-open', [file]];
  try {
    const child = spawn(command, args, { detached: true, stdio: 'ignore' });
    child.on('error', () => {});
    child.unref();
  } catch {
    /* opening the browser is only a convenience */
  }
}

/** --seq value → positive integer, null when absent, NaN when invalid. */
function parseSeq(seq) {
  if (seq === null || seq === undefined || seq === '') return null;
  const text = String(seq).trim();
  return /^\d+$/.test(text) && Number(text) >= 1 ? Number(text) : NaN;
}

/**
 * `node index.js preview [first|remind1|remind2|original] [--seq N] [--open]`
 *
 * Updates the templates' reminder dates from .env, renders the requested
 * emails to campaigns/<id>/preview/<variant>.html plus an overview page
 * index.html, and prints every subject and file path. Never contacts Shopify.
 * Sample data uses GIFT_CARD_EXPIRES_ON; with --seq the expiry frozen in the
 * list (selection.params.giftCardExpiresOn) is shown, as on the real cards.
 *
 * @returns {Promise<{ exitCode: number, files: string[], templatesChanged: boolean, index: string|null, results: object[] }>}
 *   files: the rendered email pages; index: the overview page.
 */
export async function runPreview({
  config,
  variant = 'all',
  seq = null,
  open = false,
  templatesDir = DEFAULT_TEMPLATES_DIR,
  log = console,
  now = () => new Date(),
  openFile = openInBrowser,
} = {}) {
  const stop = (exitCode, templatesChanged = false) => ({ exitCode, files: [], templatesChanged, index: null, results: [] });

  const requested = String(variant ?? 'all').trim().toLowerCase() || 'all';
  if (requested !== 'all' && !VARIANTS.includes(requested)) {
    log.error(unknownVariantMessage(variant));
    return stop(2);
  }
  const seqNo = parseSeq(seq);
  if (Number.isNaN(seqNo)) {
    log.error(`--seq 必须是名单里的序号（正整数），收到 "${seq}"`);
    return stop(2);
  }

  const paths = campaignPaths(config);

  const sync = syncTemplateDates({ config, templatesDir, write: true, log });
  if (sync.problems.length) {
    for (const p of sync.problems) log.error(p);
    log.error('邮件模板里的提醒日期没有更新，预览没有生成');
    return stop(1);
  }
  if (!sync.changed) {
    log.info(`模板里的提醒日期和 .env 一致（REMIND_1_DATE=${config.remind1Date}，REMIND_2_DATE=${config.remind2Date}）`);
  }

  let recipient = null;
  let expiresOn; // undefined = GIFT_CARD_EXPIRES_ON from .env (sample data)
  if (seqNo !== null) {
    let selection;
    try {
      selection = readJson(paths.selection, null);
    } catch (err) {
      log.error(err.message);
      return stop(1, sync.changed);
    }
    if (!selection) {
      log.error(`还没有名单（${paths.selection}）：--seq 需要先运行 select`);
      return stop(1, sync.changed);
    }
    const recipients = selection.recipients ?? [];
    recipient = recipients.find((r) => r.seq === seqNo) ?? null;
    if (!recipient) {
      log.error(`名单里没有序号 ${seqNo}（名单共 ${recipients.length} 人${recipients.length ? `，序号 1–${recipients.length}` : ''}）`);
      return stop(1, sync.changed);
    }
    log.info(`用名单第 ${seqNo} 号顾客的名字和金额：${recipient.firstName || '（没有名字）'}，${formatUsd(recipient.amountCents)}`);
    // issue creates the cards with the expiry frozen in the list (selection.params), not with
    // today's .env, so a real recipient's preview shows that date ('' / missing = no expiry).
    expiresOn = selection.params?.giftCardExpiresOn || null;
    if ((expiresOn ?? '') !== (config.giftCardExpiresOn || '')) {
      log.warn(`名单生成时的礼品卡到期日是 ${expiresOn || '（未设置）'}，.env 的 GIFT_CARD_EXPIRES_ON 现在是 ${config.giftCardExpiresOn || '（未设置）'}：`
        + '卡按名单里的日期建，预览也按名单里的日期显示');
    }
  }

  const variants = requested === 'all' ? [...VARIANTS] : [requested];
  const results = [];
  try {
    for (const name of variants) {
      // Heading first, so a warning about this email is printed under it.
      const sendDate = simulatedSendDate(config, name, now);
      log.info(`${LABELS[name]}（${name}），模拟 ${sendDate} 发送`);
      const r = await renderPreviewFor({ config, paths, variant: name, recipient, templatesDir, log, now, sendDate, expiresOn, checkDates: false });
      results.push(r);
      log.info(`  主题：${r.subject}`);
      log.info(`  文件：${r.file}`);
    }
  } catch (err) {
    log.error(`预览生成失败：${err.message}`);
    return { ...stop(1, sync.changed), files: results.map((r) => r.file), results };
  }

  const timezone = config.timezone || STORE_TIMEZONE;
  const index = path.join(paths.previewDir, 'index.html');
  writeFileAtomic(index, indexHtml({
    config, results, recipient, sync, generatedAt: localDateTime(now().getTime(), timezone), timezone,
    expiresOn: results[0]?.expiresOn ?? null, fromList: expiresOn !== undefined,
  }));
  log.info(`总览：${index}`);
  log.info('注意：Shopify 后台的邮件样式表不在本地，排版和真实邮件会有小差别；真实样子以测试活动收到的邮件为准。');

  if (open) {
    const target = requested === 'all' ? index : results[0].file;
    log.info(`在浏览器中打开：${target}`);
    try {
      await openFile(target);
    } catch {
      /* opening is a convenience; the paths are printed above */
    }
  }

  return { exitCode: 0, files: results.map((r) => r.file), templatesChanged: sync.changed, index, results };
}
