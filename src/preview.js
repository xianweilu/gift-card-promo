// Local preview of the "Gift card created" notification (subject + body).
//
// The two templates in notifications/ are pasted into the Shopify admin by
// hand and Shopify renders them when it sends the email. This module renders
// the same templates locally with liquidjs so the copy can be checked before
// pasting: the promo email (cards issued by this tool) and the original email
// for gift cards that are not part of the campaign. Reminders are sent with
// Shopify Email, so there is one promo copy only. Every date in the body comes
// from the card's expiry date (gift_card.expires_on); the subject hardcodes
// its date, because Shopify caps a subject template at 512 characters and the
// Liquid that computes "19th" does not fit. The preview warns when that
// hardcoded date is not the cards' expiry date, or the template is too long.
//
// Nothing here contacts Shopify and nothing here writes to the templates.

import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { Liquid } from 'liquidjs';
import { ROOT_DIR } from './config.js';
import { campaignPaths, readJson, writeFileAtomic } from './campaign.js';
import { localDate, localDateTime } from './time.js';
import { formatUsd } from './select/amount.js';

/** The emails that can be previewed, in display order. */
export const VARIANTS = Object.freeze(['first', 'original']);

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

/**
 * Shopify refuses a notification subject template longer than this ("Email
 * subject is too long (maximum is 512 characters)"). Liquid logic counts too,
 * which is why the promo subject hardcodes its date instead of computing it.
 */
export const SUBJECT_MAX_LENGTH = 512;

// The hardcoded date in the promo subject: "… by Oct. 19th!" (month name or
// abbreviation, optional period, day, optional ordinal suffix).
const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
const SUBJECT_DATE_RE = /\bby\s+((jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{1,2})(?:st|nd|rd|th)?)\b/i;

/** Chinese names of the variants. */
const LABELS = Object.freeze({ first: '首封', original: '原版' });
/** Chinese names of the template's copies (its promo-stage marker). */
const STAGE_LABELS = Object.freeze({ promo: '活动文案', original: '原版' });

// The body template prints which copy it chose: <!-- promo-stage: promo --> or original.
const STAGE_MARKER = /<!--\s*promo-stage:\s*([\w-]+)\s*-->/;

// ---------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------

/** "YYYY-MM-DD" today in the store's time zone. */
function storeToday(config, now) {
  return localDate(now().getTime(), config.timezone || STORE_TIMEZONE);
}

/**
 * The day a preview is pretended to be sent: LAUNCH_DATE, or today when it is
 * not set. The copy no longer depends on it; it only pins the template's 'now'.
 */
export function simulatedSendDate(config, variant, now = () => new Date()) {
  if (!VARIANTS.includes(variant)) throw new Error(unknownVariantMessage(variant));
  return config.launchDate || storeToday(config, now);
}

/** The copy the template must show for `variant`: 'promo' for the campaign email, 'original' otherwise. */
export function expectedStage(variant) {
  return variant === 'original' ? 'original' : 'promo';
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
 * on sendDate, which is around noon in Los Angeles, so a `'now' | date` gives
 * sendDate both in UTC and in the store's zone. locale en-US keeps month names
 * English on a machine set to another language.
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

/** The copy the body template chose (from its promo-stage marker), or null. */
export function parseStage(html) {
  return STAGE_MARKER.exec(String(html ?? ''))?.[1] ?? null;
}

// ---------------------------------------------------------------------------
// Subject template checks
// ---------------------------------------------------------------------------

/** The length Shopify measures: characters (code points, so 💰 is one), without a trailing newline. */
export function subjectLength(template) {
  return [...String(template ?? '').trimEnd()].length;
}

/**
 * The date written into the subject template by hand ("… by Oct. 19th!"), as
 * { text: 'Oct. 19th', month: 10, day: 19 }, or null when there is none. Only
 * literal text is searched; Liquid tags and outputs are ignored.
 */
export function subjectHardcodedDate(template) {
  const literal = String(template ?? '').replace(/\{\{[\s\S]*?\}\}|\{%[\s\S]*?%\}/g, ' ');
  const m = SUBJECT_DATE_RE.exec(literal);
  if (!m) return null;
  return { text: m[1], month: MONTHS.indexOf(m[2].toLowerCase()) + 1, day: Number(m[3]) };
}

/**
 * Warnings about the subject template: longer than Shopify accepts, or a
 * hardcoded date that is not the cards' expiry date (`expiresOn`: 'YYYY-MM-DD',
 * or null for cards without one). The year is not in the subject, so only the
 * month and day are compared. Returns [] when everything is fine.
 */
export function subjectTemplateWarnings(template, expiresOn) {
  const warnings = [];
  const length = subjectLength(template);
  if (length > SUBJECT_MAX_LENGTH) {
    warnings.push(`主题模板有 ${length} 个字符，超过 Shopify 的上限 ${SUBJECT_MAX_LENGTH} 个，后台会拒绝保存：请缩短 ${TEMPLATE_FILES.subject}`);
  }
  const hard = subjectHardcodedDate(template);
  if (hard) {
    if (!expiresOn) {
      warnings.push(`主题模板里写死了到期日 "${hard.text}"，但这些卡没有到期日：请改 ${TEMPLATE_FILES.subject} 后重新贴到 Shopify 后台`);
    } else {
      const [, month, day] = String(expiresOn).split('-').map(Number);
      if (hard.month !== month || hard.day !== day) {
        warnings.push(`主题模板里写死的到期日 "${hard.text}" 和礼品卡到期日 ${expiresOn} 不一致：请改 ${TEMPLATE_FILES.subject} 后重新贴到 Shopify 后台`);
      }
    }
  }
  return warnings;
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
  return `未知的预览类型 "${variant}"。可选：first（首封）、original（原版），不写则两种都生成`;
}

function describeRecipient(recipient, firstName) {
  if (!recipient) return `示例顾客 ${firstName}`;
  return `名单第 ${recipient.seq} 号顾客 ${firstName || '（没有名字）'}`;
}

/** The small notice shown at the top of a preview page (not part of the real email). */
function bannerHtml({ subject, variant, recipient, data }) {
  const firstName = data.gift_card.customer.first_name;
  const expiry = data.gift_card.expires_on;
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
    + `<strong>本地预览</strong>　主题：${escapeHtml(subject)}，变体：${escapeHtml(variant)}（${LABELS[variant]}），礼品卡到期日：${escapeHtml(expiry || '（未设置）')}<br>`
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
 * @param {'first'|'original'} a.variant
 * @param {object|null} [a.recipient] a selection.json recipient (first name + amount), or sample data
 * @param {string} [a.templatesDir]
 * @param {object} [a.log]
 * @param {() => Date} [a.now] only used for "today" when LAUNCH_DATE is not set
 * @param {string|null} [a.sendDate] YYYY-MM-DD the template's 'now' is pinned to (default: LAUNCH_DATE or today);
 *   the copy does not depend on it
 * @param {string|null} [a.expiresOn] the cards' expiry date to show, overriding GIFT_CARD_EXPIRES_ON
 *   (callers with a selection pass selection.params.giftCardExpiresOn; '' / null = no expiry date)
 * @returns {Promise<{ file: string, subject: string, stage: string|null, expectedStage: string, sendDate: string, variant: string, expiresOn: string|null }>}
 *   stage: what the template actually showed ('promo' | 'original'); expectedStage: what the variant calls for;
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
} = {}) {
  if (!VARIANTS.includes(variant)) throw new Error(unknownVariantMessage(variant));
  const date = sendDate ?? simulatedSendDate(config, variant, now);
  if (!DATE_RE.test(String(date))) throw new Error(`模拟发送日期必须是 YYYY-MM-DD，收到 "${date}"`);
  const expiry = previewExpiresOn(config, expiresOn); // validated before anything is read or written

  const subjectTemplate = readTemplate(templatesDir, TEMPLATE_FILES.subject);
  const bodyTemplate = readTemplate(templatesDir, TEMPLATE_FILES.body);
  const engine = createPreviewEngine({ sendDate: date, currency: config.giftCardCurrency || 'USD' });
  const data = previewData(config, variant, recipient, { expiresOn: expiry });

  // A subject is a single line: collapse the template's whitespace and newlines.
  const subject = (await engine.parseAndRender(subjectTemplate, data)).replace(/\s+/g, ' ').trim();
  const body = await engine.parseAndRender(bodyTemplate, data);

  const stage = parseStage(body);
  const expected = expectedStage(variant);
  if (stage !== expected) {
    const shown = stage ? `${STAGE_LABELS[stage] ?? stage}（${stage}）` : '没有 promo-stage 标记';
    const suffixHint = stage === 'original' && variant !== 'original'
      ? '；可能是 .env 的 GIFT_CARD_TEMPLATE_SUFFIX 和模板里的 promo 后缀不一致'
      : '';
    log.warn(`预览的文案和预期不一致：${LABELS[variant]}预期${STAGE_LABELS[expected]}（${expected}），模板显示${shown}${suffixHint}`);
  }

  if (recipient && !data.gift_card.customer.first_name) {
    log.warn(`名单第 ${recipient.seq} 号顾客没有名字，邮件里只会显示 "Hello,"`);
  }

  // The subject's length limit and hardcoded date concern the promo email; checked
  // once per run (the original variant shares the file, so it would repeat them).
  if (variant !== 'original') {
    for (const warning of subjectTemplateWarnings(subjectTemplate, data.gift_card.expires_on)) log.warn(warning);
  }

  const file = path.join(paths.previewDir, `${variant}.html`);
  writeFileAtomic(file, withBanner(body, bannerHtml({ subject, variant, recipient, data })));
  return { file, subject, stage, expectedStage: expected, sendDate: date, variant, expiresOn: data.gift_card.expires_on };
}

// ---------------------------------------------------------------------------
// Overview page
// ---------------------------------------------------------------------------

/**
 * The overview page. `expiresOn` is the expiry date the previews showed (null = none);
 * `fromList` says it is the one frozen in the list rather than GIFT_CARD_EXPIRES_ON.
 */
function indexHtml({ config, results, recipient, generatedAt, timezone, expiresOn, fromList }) {
  const sample = previewData(config, 'first', recipient, { expiresOn }); // the name and amount every page used
  const who = describeRecipient(recipient, sample.gift_card.customer.first_name);
  const envExpiry = config.giftCardExpiresOn || null;
  const expiryNote = fromList && expiresOn !== envExpiry
    ? `（按名单生成时的日期，卡按这个日期建；.env 现在是 ${envExpiry || '（未设置）'}）`
    : '';
  const rows = results.map((r) => {
    const mismatch = r.stage !== r.expectedStage
      ? `<br><span class="warn">⚠ 模板显示的是${escapeHtml(STAGE_LABELS[r.stage] ?? r.stage ?? '（没有 promo-stage 标记）')}，应为${escapeHtml(STAGE_LABELS[r.expectedStage])}</span>`
      : '';
    const note = r.variant === 'original' ? '<br><span class="note">不是本活动的礼品卡收到的邮件</span>' : '';
    const name = path.basename(r.file);
    return `      <tr><td>${LABELS[r.variant]}（${r.variant}）${note}${mismatch}</td>`
      + `<td>${escapeHtml(r.subject)}</td><td><a href="${escapeHtml(encodeURIComponent(name))}">${escapeHtml(name)}</a></td></tr>`;
  }).join('\n');
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
  <p>礼品卡到期日：${escapeHtml(expiresOn || '（未设置）')}${escapeHtml(expiryNote)}。邮件里的日期都由卡的到期日生成；提醒邮件用 Shopify Email 发送，不在这里预览。</p>
  <table>
    <thead>
      <tr><th>变体</th><th>主题</th><th>文件链接</th></tr>
    </thead>
    <tbody>
${rows}
    </tbody>
  </table>
  <p class="note">Shopify 后台的邮件样式表不在本地，排版和真实邮件会有小差别。预览用来核对文案；真实样子以测试活动收到的邮件为准。</p>
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
 * `node index.js preview [first|original] [--seq N] [--open]`
 *
 * Renders the requested emails to campaigns/<id>/preview/<variant>.html plus
 * an overview page index.html, and prints every subject and file path. Never
 * contacts Shopify and never changes the templates. Sample data uses
 * GIFT_CARD_EXPIRES_ON; with --seq the expiry frozen in the list
 * (selection.params.giftCardExpiresOn) is shown, as on the real cards.
 *
 * @returns {Promise<{ exitCode: number, files: string[], index: string|null, results: object[] }>}
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
  const stop = (exitCode) => ({ exitCode, files: [], index: null, results: [] });

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

  let recipient = null;
  let expiresOn; // undefined = GIFT_CARD_EXPIRES_ON from .env (sample data)
  if (seqNo !== null) {
    let selection;
    try {
      selection = readJson(paths.selection, null);
    } catch (err) {
      log.error(err.message);
      return stop(1);
    }
    if (!selection) {
      log.error(`还没有名单（${paths.selection}）：--seq 需要先运行 select`);
      return stop(1);
    }
    const recipients = selection.recipients ?? [];
    recipient = recipients.find((r) => r.seq === seqNo) ?? null;
    if (!recipient) {
      log.error(`名单里没有序号 ${seqNo}（名单共 ${recipients.length} 人${recipients.length ? `，序号 1–${recipients.length}` : ''}）`);
      return stop(1);
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
      log.info(`${LABELS[name]}（${name}）`);
      const r = await renderPreviewFor({ config, paths, variant: name, recipient, templatesDir, log, now, expiresOn });
      results.push(r);
      log.info(`  主题：${r.subject}`);
      log.info(`  文件：${r.file}`);
    }
  } catch (err) {
    log.error(`预览生成失败：${err.message}`);
    return { ...stop(1), files: results.map((r) => r.file), results };
  }

  const timezone = config.timezone || STORE_TIMEZONE;
  const index = path.join(paths.previewDir, 'index.html');
  writeFileAtomic(index, indexHtml({
    config, results, recipient, generatedAt: localDateTime(now().getTime(), timezone), timezone,
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

  return { exitCode: 0, files: results.map((r) => r.file), index, results };
}
