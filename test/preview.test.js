// Tests for src/preview.js: local rendering of the "Gift card created" email
// templates and the sync of their reminder dates with .env.
//
// Every test renders a private copy of the real templates (notifications/)
// in a temp folder, writes previews into a temp campaigns folder, and runs
// with globalThis.fetch replaced by a stub that fails the test if called.

import { test, beforeEach, afterEach, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { pathToFileURL } from 'node:url';
import { ROOT_DIR } from '../src/config.js';
import { campaignPaths, writeJsonAtomic } from '../src/campaign.js';
import { testConfig, memoryLog, makeCustomer, makeOrder, selectionFixture } from './helpers.js';
import {
  VARIANTS,
  TEMPLATE_FILES,
  syncTemplateDates,
  renderPreviewFor,
  runPreview,
  createPreviewEngine,
  previewData,
  previewExpiresOn,
  parseStage,
  expectedStage,
  simulatedSendDate,
} from '../src/preview.js';

const REAL_DIR = path.join(ROOT_DIR, 'notifications');
const REAL = {
  subject: fs.readFileSync(path.join(REAL_DIR, TEMPLATE_FILES.subject), 'utf8'),
  body: fs.readFileSync(path.join(REAL_DIR, TEMPLATE_FILES.body), 'utf8'),
};

// "Cha-Ching" in Unicode mathematical sans-serif bold, exactly as in the subject template.
const CHA_CHING = '𝗖𝗵𝗮-𝗖𝗵𝗶𝗻𝗴';
const LEGAL_LINE = 'PROMOTIONAL GIFT CARD · VALID THROUGH OCTOBER 19, 2026';
const MISMATCH = '预览的文案和预期不一致：模板里的日期可能和 .env 不一致';
const UPDATED = '模板已按 .env 更新，请把主题和正文重新贴到 Shopify 后台';
const NOW = () => new Date('2026-10-02T17:00:00Z'); // 10:00 in Los Angeles
const execFileAsync = promisify(execFile);

/** Set the reminder-date block of a template to the given YYYYMMDD values. */
function withDates(text, remind1, remind2) {
  return text
    .replace(/(assign\s+promo_remind_1\s*=\s*)\d{8}/g, `$1${remind1}`)
    .replace(/(assign\s+promo_remind_2\s*=\s*)\d{8}/g, `$1${remind2}`);
}

// The copies start with the dates testConfig() uses, so results do not depend
// on whichever dates were last synced into the real templates.
const BASE = {
  subject: withDates(REAL.subject, '20261012', '20261016'),
  body: withDates(REAL.body, '20261012', '20261016'),
};

let cleanups = [];
let fetchCalls = 0;
let realFetch;

beforeEach(() => {
  fetchCalls = 0;
  realFetch = globalThis.fetch;
  globalThis.fetch = async () => {
    fetchCalls += 1;
    throw new Error('preview must never use the network');
  };
});

afterEach(() => {
  globalThis.fetch = realFetch;
  for (const fn of cleanups.reverse()) fn();
  cleanups = [];
  assert.equal(fetchCalls, 0, 'preview called fetch');
});

after(() => {
  // No test may touch the real templates.
  assert.equal(fs.readFileSync(path.join(REAL_DIR, TEMPLATE_FILES.subject), 'utf8'), REAL.subject);
  assert.equal(fs.readFileSync(path.join(REAL_DIR, TEMPLATE_FILES.body), 'utf8'), REAL.body);
});

/** A fresh config, campaigns folder and template copies for one test. */
function setup(overrides = {}, { subject = BASE.subject, body = BASE.body } = {}) {
  const t = testConfig(overrides);
  const templatesDir = fs.mkdtempSync(path.join(os.tmpdir(), 'gcp-templates-'));
  fs.writeFileSync(path.join(templatesDir, TEMPLATE_FILES.subject), subject);
  fs.writeFileSync(path.join(templatesDir, TEMPLATE_FILES.body), body);
  cleanups.push(() => {
    t.cleanup();
    fs.rmSync(templatesDir, { recursive: true, force: true });
  });
  return {
    config: t.config,
    paths: campaignPaths(t.config),
    templatesDir,
    subjectFile: path.join(templatesDir, TEMPLATE_FILES.subject),
    bodyFile: path.join(templatesDir, TEMPLATE_FILES.body),
    log: memoryLog(),
  };
}

const read = (file) => fs.readFileSync(file, 'utf8');
const warnings = (log) => log.lines.filter((l) => l.startsWith('WARN '));
const errors = (log) => log.lines.filter((l) => l.startsWith('ERROR '));
// The email as Shopify would send it: the page minus the preview-only banner.
const emailOnly = (html) => html.replace(/<div class="gift-card-promo-preview"[^>]*>.*?<\/div>/s, '');
const buttonText = (html) => /class="button__text">([^<]*)</.exec(html)?.[1];
const emailTitle = (html) => /<title>([^<]*)<\/title>/.exec(html)?.[1].trim();

async function previewAll(s, options = {}) {
  const result = await runPreview({ config: s.config, templatesDir: s.templatesDir, log: s.log, now: NOW, openFile: () => {}, ...options });
  const byVariant = Object.fromEntries(result.results.map((r) => [r.variant, { ...r, html: read(r.file) }]));
  return { result, byVariant };
}

// ---------------------------------------------------------------------------
// Rendering the four emails
// ---------------------------------------------------------------------------

test('renders all four emails: each picks the copy for its send date', async () => {
  const s = setup();
  const { result, byVariant } = await previewAll(s);

  assert.equal(result.exitCode, 0);
  assert.equal(result.templatesChanged, false);
  assert.deepEqual(result.results.map((r) => r.variant), VARIANTS);
  assert.deepEqual(result.results.map((r) => r.sendDate), ['2026-10-05', '2026-10-12', '2026-10-16', '2026-10-05']);
  for (const variant of VARIANTS) {
    const r = byVariant[variant];
    assert.equal(r.stage, variant, `${variant}: stage`);
    assert.equal(r.expectedStage, variant);
    assert.ok(r.html.includes(`<!-- promo-stage: ${variant} -->`), `${variant}: marker`);
    assert.equal(parseStage(r.html), variant);
    assert.doesNotMatch(r.subject, /\n/, 'a subject is one line');
  }
  assert.deepEqual(warnings(s.log), [], 'templates agree with .env: no warnings');
  assert.deepEqual(errors(s.log), []);
});

test('subjects: first keeps the bold Cha-Ching, reminders name the expiry day, original shows the amount', async () => {
  assert.ok(REAL.subject.includes(CHA_CHING), 'the subject template still has the math-bold Cha-Ching');
  const s = setup();
  const { byVariant } = await previewAll(s);

  assert.ok(byVariant.first.subject.startsWith('💰 '));
  assert.ok(byVariant.first.subject.includes(CHA_CHING));
  assert.ok(byVariant.first.subject.includes('𝗟𝗔 𝗕𝗮𝗹𝗹𝗼𝗼𝗻𝘀'));
  assert.equal(byVariant.remind1.subject, "⏰ Don't forget: your LA Balloons credit expires October 19");
  assert.equal(byVariant.remind2.subject, '⌛ Last chance: your LA Balloons credit expires October 19');
  assert.equal(byVariant.original.subject, 'LA Balloons $15.33 Online Credit Code');
  for (const v of ['remind1', 'remind2', 'original']) assert.ok(!byVariant[v].subject.includes(CHA_CHING));
});

test('promo emails carry the legal line with the real expiry day (no off-by-one); the original does not', async () => {
  const s = setup();
  const { byVariant } = await previewAll(s);

  for (const v of ['first', 'remind1', 'remind2']) {
    const email = emailOnly(byVariant[v].html);
    assert.ok(email.includes(LEGAL_LINE), `${v}: legal line`);
    assert.ok(email.includes('valid through October 19, 2026') || email.includes('used through October 19, 2026'), `${v}: body date`);
    assert.ok(!email.includes('OCTOBER 18') && !email.includes('October 18'), `${v}: no day shift`);
    assert.ok(!email.includes('$15.33'), `${v}: the promo copy never shows the amount`);
  }
  assert.ok(!byVariant.original.html.includes('VALID THROUGH'));
  assert.ok(emailOnly(byVariant.original.html).includes('here is your $15.33 online credit code'));
});

test('dates stay on their calendar day whatever time zone the machine is in', async () => {
  const s = setup();
  const before = process.env.TZ;
  try {
    for (const tz of ['America/Los_Angeles', 'Asia/Shanghai', 'Pacific/Kiritimati', 'Etc/GMT+12', 'UTC']) {
      process.env.TZ = tz;
      for (const variant of ['first', 'remind1', 'remind2']) {
        const r = await renderPreviewFor({ config: s.config, paths: s.paths, variant, templatesDir: s.templatesDir, log: s.log, now: NOW });
        assert.equal(r.stage, variant, `${tz} ${variant}`);
        assert.ok(read(r.file).includes(LEGAL_LINE), `${tz} ${variant}: legal line`);
        if (variant !== 'first') assert.ok(r.subject.endsWith('expires October 19'), `${tz} ${variant}: ${r.subject}`);
      }
    }
  } finally {
    if (before === undefined) delete process.env.TZ;
    else process.env.TZ = before;
  }
  assert.deepEqual(warnings(s.log), []);
});

test('month names stay English on a machine set to Chinese', async (t) => {
  // The default Intl locale comes from the environment at start-up, so check in a child process.
  const script = `
    const { createPreviewEngine } = await import(${JSON.stringify(pathToFileURL(path.join(ROOT_DIR, 'src', 'preview.js')).href)});
    const out = await createPreviewEngine({ sendDate: '2026-10-12' }).parseAndRender("{{ '2026-10-19' | date: '%B %-d, %Y' }}");
    console.log(JSON.stringify({ locale: Intl.DateTimeFormat().resolvedOptions().locale, out }));
  `;
  const { stdout } = await execFileAsync(process.execPath, ['--input-type=module', '-e', script], {
    env: { ...process.env, LC_ALL: 'zh_CN.UTF-8', LANG: 'zh_CN.UTF-8' },
  });
  const { locale, out } = JSON.parse(stdout.trim());
  if (!locale.startsWith('zh')) {
    t.skip(`this Node build ignores LC_ALL (default locale ${locale})`);
    return;
  }
  assert.equal(out, 'October 19, 2026');
});

test('each email has its own title and button; the last reminder says "Use"', async () => {
  const s = setup();
  const { byVariant } = await previewAll(s);

  assert.equal(emailTitle(byVariant.first.html), 'A little surprise, just for you');
  assert.equal(emailTitle(byVariant.remind1.html), 'Your surprise is still waiting');
  assert.equal(emailTitle(byVariant.remind2.html), 'Last chance to use your credit');
  assert.equal(emailTitle(byVariant.original.html), 'Your Online Credit Code is now available!');

  assert.equal(buttonText(byVariant.first.html), 'Reveal your Online Credit Code');
  assert.equal(buttonText(byVariant.remind1.html), 'Reveal your Online Credit Code');
  assert.equal(buttonText(byVariant.remind2.html), 'Use your Online Credit Code');
  assert.equal(buttonText(byVariant.original.html), 'View your Online Credit Code');
  assert.ok(!byVariant.remind2.html.includes('Reveal your Online Credit Code'));

  assert.ok(byVariant.first.html.includes("Hi Alex, we've set aside an exclusive Online Credit Code"));
  assert.ok(byVariant.remind1.html.includes("hasn't been used yet"));
  assert.ok(byVariant.remind2.html.includes('this is your final reminder'));
  // Shop name heading (no email logo) and the store link around the button.
  assert.ok(byVariant.first.html.includes('<a href="https://www.laballoons.com">LA Balloons</a>'));
  assert.ok(!byVariant.first.html.includes('<img src=""'));
});

test('engine: now/today are the send date; money and asset filters behave like Shopify', async () => {
  const engine = createPreviewEngine({ sendDate: '2026-10-12' });
  const render = (tpl, data = {}) => engine.parseAndRender(tpl, data);

  assert.equal(await render("{{ 'now' | date: '%Y-%m-%d %H:%M' }}|{{ 'today' | date: '%Y%m%d' }}|{{ 'Now' | date: '%Y%m%d' | plus: 0 }}"), '2026-10-12 19:00|20261012|20261012');
  assert.equal(await render("{{ '2026-10-19' | date: '%B %-d, %Y' }}|{{ d | date: '%B %-d' }}", { d: '2026-11-01' }), 'October 19, 2026|November 1');

  assert.equal(await render('{{ 1077 | money_without_trailing_zeros }} {{ 1000 | money_without_trailing_zeros }} {{ 123400 | money_without_trailing_zeros }}'), '$10.77 $10 $1,234');
  assert.equal(await render('{{ 1077 | money }} {{ 1000 | money }} {{ 123456 | money }}'), '$10.77 $10.00 $1,234.56');
  assert.equal(await render('{{ 1077 | money_with_currency }}|{{ 1533 | money_without_currency }}'), '$10.77 USD|15.33');
  assert.equal(await render('[{{ nothing | money }}][{{ nothing | money_without_trailing_zeros }}]'), '[][]');
  assert.equal(await render("{{ 'gift-card/add-to-apple-wallet.png' | shopify_asset_url }}"), 'https://cdn.shopify.com/shopifycloud/shopify/assets/gift-card/add-to-apple-wallet.png');
  assert.equal(await render('[{{ no_such.thing }}]'), '[]', 'unknown variables render empty');

  const cad = createPreviewEngine({ sendDate: '2026-10-12', currency: 'CAD' });
  assert.equal(await cad.parseAndRender('{{ 1077 | money_with_currency }}'), '$10.77 CAD');
  assert.throws(() => createPreviewEngine({ sendDate: '10/12/2026' }), /YYYY-MM-DD/);
});

test('sample data, simulated send dates and the expected stage', () => {
  const { config } = setup();
  const data = previewData(config, 'first');
  assert.equal(data.gift_card.template_suffix, 'gift-card-promo');
  assert.equal(data.gift_card.balance, 1533);
  assert.equal(data.gift_card.initial_value, 1533);
  assert.equal(data.gift_card.expires_on, '2026-10-19');
  assert.equal(data.gift_card.customer.first_name, 'Alex');
  assert.equal(data.gift_card.pass_url, null);
  assert.equal(data.order_name, '');
  assert.equal(data.shop.name, 'LA Balloons');
  assert.equal(previewData(config, 'original').gift_card.template_suffix, '');
  const real = previewData(config, 'remind1', { seq: 3, firstName: 'Maria', amountCents: 1977 });
  assert.equal(real.gift_card.customer.first_name, 'Maria');
  assert.equal(real.gift_card.balance, 1977);

  assert.equal(simulatedSendDate(config, 'first'), '2026-10-05');
  assert.equal(simulatedSendDate(config, 'original'), '2026-10-05');
  assert.equal(simulatedSendDate(config, 'remind1'), '2026-10-12');
  assert.equal(simulatedSendDate(config, 'remind2'), '2026-10-16');
  assert.throws(() => simulatedSendDate(config, 'remind3'), /未知的预览类型/);

  assert.equal(expectedStage(config, 'first', '2026-10-11'), 'first');
  assert.equal(expectedStage(config, 'remind1', '2026-10-12'), 'remind1');
  assert.equal(expectedStage(config, 'remind1', '2026-10-15'), 'remind1');
  assert.equal(expectedStage(config, 'remind2', '2026-10-16'), 'remind2');
  assert.equal(expectedStage(config, 'first', '2026-10-20'), 'remind2');
  assert.equal(expectedStage(config, 'original', '2026-10-16'), 'original');
});

test('LAUNCH_DATE not set: first and original are rendered as sent today in the store time zone', async () => {
  const s = setup({ LAUNCH_DATE: '' });
  // 2026-10-03 05:00 UTC is still October 2 in Los Angeles.
  const { result } = await previewAll(s, { now: () => new Date('2026-10-03T05:00:00Z') });
  assert.equal(result.exitCode, 0);
  const first = result.results.find((r) => r.variant === 'first');
  const original = result.results.find((r) => r.variant === 'original');
  assert.equal(first.sendDate, '2026-10-02');
  assert.equal(first.stage, 'first');
  assert.equal(original.sendDate, '2026-10-02');
});

// ---------------------------------------------------------------------------
// Template dates ⇄ .env
// ---------------------------------------------------------------------------

test('syncTemplateDates rewrites both templates when REMIND_1_DATE changes, and nothing else', async () => {
  const s = setup({ REMIND_1_DATE: '2026-10-13' });

  const res = syncTemplateDates({ config: s.config, templatesDir: s.templatesDir, log: s.log });
  assert.equal(res.changed, true);
  assert.deepEqual(res.problems, []);
  assert.deepEqual([...res.files].sort(), [s.bodyFile, s.subjectFile].sort());
  assert.deepEqual(res.changes.map((c) => [path.basename(c.file), c.name, c.from, c.to]), [
    [TEMPLATE_FILES.subject, 'promo_remind_1', '20261012', '20261013'],
    [TEMPLATE_FILES.body, 'promo_remind_1', '20261012', '20261013'],
  ]);
  assert.equal(read(s.subjectFile), withDates(BASE.subject, '20261013', '20261016'), 'only the date changed in the subject');
  assert.equal(read(s.bodyFile), withDates(BASE.body, '20261013', '20261016'), 'only the date changed in the body');
  assert.ok(read(s.subjectFile).startsWith('{%- assign promo_remind_1 = 20261013 -%}{%- assign promo_remind_2 = 20261016 -%}'));
  assert.ok(read(s.bodyFile).includes('{%- assign promo_remind_1 = 20261013 -%}\n{%- assign promo_remind_2 = 20261016 -%}'));
  assert.deepEqual(warnings(s.log), [`WARN ${UPDATED}`]);

  // Running it again is a no-op.
  const again = syncTemplateDates({ config: s.config, templatesDir: s.templatesDir, log: s.log });
  assert.equal(again.changed, false);
  assert.deepEqual(again.files, []);
  assert.equal(warnings(s.log).length, 1);

  // An email sent on 10/12 now still shows the first-email copy; 10/13 shows reminder 1.
  const on1012 = await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'remind1', sendDate: '2026-10-12', templatesDir: s.templatesDir, log: s.log });
  assert.equal(on1012.stage, 'first');
  assert.equal(on1012.expectedStage, 'first');
  assert.equal(on1012.sendDate, '2026-10-12');
  const on1013 = await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'remind1', templatesDir: s.templatesDir, log: s.log });
  assert.equal(on1013.sendDate, '2026-10-13');
  assert.equal(on1013.stage, 'remind1');
  assert.ok(!s.log.lines.some((l) => l.includes(MISMATCH)), 'the copy agrees with .env');
});

test('write: false reports outdated templates without touching them', () => {
  const s = setup({ REMIND_2_DATE: '2026-10-17' });
  const res = syncTemplateDates({ config: s.config, templatesDir: s.templatesDir, write: false, log: s.log });
  assert.equal(res.changed, true);
  assert.equal(res.files.length, 2);
  assert.equal(read(s.subjectFile), BASE.subject);
  assert.equal(read(s.bodyFile), BASE.body);
  assert.deepEqual(s.log.lines, []);
});

test('runPreview syncs the templates first and says they must be pasted again', async () => {
  const s = setup({ REMIND_2_DATE: '2026-10-17' });
  const { result, byVariant } = await previewAll(s);

  assert.equal(result.exitCode, 0);
  assert.equal(result.templatesChanged, true);
  assert.ok(warnings(s.log).includes(`WARN ${UPDATED}`));
  assert.ok(read(s.subjectFile).includes('promo_remind_2 = 20261017'));
  assert.ok(read(s.bodyFile).includes('promo_remind_2 = 20261017'));
  assert.equal(byVariant.remind2.sendDate, '2026-10-17');
  assert.equal(byVariant.remind2.stage, 'remind2');
  assert.ok(read(result.index).includes(UPDATED));

  // 10/16 is now still inside the first-reminder window.
  const on1016 = await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'remind2', sendDate: '2026-10-16', templatesDir: s.templatesDir, log: s.log });
  assert.equal(on1016.stage, 'remind1');
  assert.ok(!s.log.lines.some((l) => l.includes(MISMATCH)));
});

test('a missing date assignment is a problem: nothing is written and preview stops', async () => {
  const subject = BASE.subject.replace(/\{%-\s*assign\s+promo_remind_2\s*=\s*\d{8}\s*-%\}/, '');
  assert.ok(!subject.includes('promo_remind_2 ='), 'test setup removed the assignment');
  const s = setup({ REMIND_1_DATE: '2026-10-13' }, { subject });

  const res = syncTemplateDates({ config: s.config, templatesDir: s.templatesDir, log: s.log });
  assert.equal(res.changed, false);
  assert.equal(res.problems.length, 1);
  assert.match(res.problems[0], /promo_remind_2/);
  assert.ok(res.problems[0].includes(s.subjectFile));
  assert.equal(read(s.bodyFile), BASE.body, 'the body is not rewritten either');
  assert.equal(read(s.subjectFile), subject);

  const result = await runPreview({ config: s.config, templatesDir: s.templatesDir, log: s.log, now: NOW, open: true, openFile: () => assert.fail('must not open') });
  assert.equal(result.exitCode, 1);
  assert.deepEqual(result.files, []);
  assert.ok(errors(s.log).some((l) => l.includes('promo_remind_2')));
  assert.ok(!fs.existsSync(s.paths.previewDir), 'no preview was written');

  // The issue/remind dry runs render directly: they still get a preview, plus a warning.
  const direct = await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'first', templatesDir: s.templatesDir, log: s.log });
  assert.ok(fs.existsSync(direct.file));
  assert.ok(warnings(s.log).some((l) => l.includes('邮件模板的提醒日期有问题') && l.includes('promo_remind_2')));
  assert.equal(read(s.bodyFile), BASE.body, 'still not rewritten');
});

test('missing .env dates, malformed values and missing templates are problems', () => {
  const noDate = setup({ REMIND_2_DATE: '' });
  const a = syncTemplateDates({ config: noDate.config, templatesDir: noDate.templatesDir, log: noDate.log });
  assert.equal(a.changed, false);
  assert.ok(a.problems.some((p) => p.includes('REMIND_2_DATE')));

  const body = BASE.body.replace('assign promo_remind_1 = 20261012', "assign promo_remind_1 = '2026-10-12'");
  const bad = setup({}, { body });
  const b = syncTemplateDates({ config: bad.config, templatesDir: bad.templatesDir, log: bad.log });
  assert.equal(b.problems.length, 1);
  assert.ok(b.problems[0].includes("'2026-10-12'"));
  assert.ok(b.problems[0].includes('不是 8 位日期'));
  assert.equal(read(bad.bodyFile), body);

  const gone = setup();
  fs.rmSync(gone.subjectFile);
  const c = syncTemplateDates({ config: gone.config, templatesDir: gone.templatesDir, log: gone.log });
  assert.ok(c.problems.some((p) => p.includes('找不到邮件模板') && p.includes(TEMPLATE_FILES.subject)));
  assert.deepEqual(gone.log.lines, []);
});

test('rendering directly (issue/remind dry runs) flags templates whose dates differ from .env', async () => {
  // The templates say reminder 1 starts 10/13, .env says 10/12.
  const s = setup({}, { subject: withDates(REAL.subject, '20261013', '20261016'), body: withDates(REAL.body, '20261013', '20261016') });
  const r = await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'remind1', templatesDir: s.templatesDir, log: s.log });
  assert.equal(r.sendDate, '2026-10-12');
  assert.equal(r.stage, 'first');
  assert.equal(r.expectedStage, 'remind1');
  assert.ok(warnings(s.log).some((l) => l.includes(MISMATCH)));
  assert.ok(warnings(s.log).some((l) => l.includes('模板里的提醒日期和 .env 不一致')));
  assert.ok(read(s.bodyFile).includes('promo_remind_1 = 20261013'), 'renderPreviewFor never rewrites the templates');
});

test('a template suffix that does not match the templates is flagged', async () => {
  const s = setup({ GIFT_CARD_TEMPLATE_SUFFIX: 'other-suffix' });
  const r = await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'first', templatesDir: s.templatesDir, log: s.log });
  assert.equal(r.stage, 'original');
  assert.ok(warnings(s.log).some((l) => l.includes(MISMATCH) && l.includes('GIFT_CARD_TEMPLATE_SUFFIX')));
});

// ---------------------------------------------------------------------------
// --seq, index page, --open, usage errors, output location
// ---------------------------------------------------------------------------

test('--seq uses the recipient\'s first name and amount', async () => {
  const s = setup();
  const selection = await selectionFixture(s.config, {
    customers: [
      makeCustomer({ n: 1, first: 'Maria', lastOrder: makeOrder({ n: 101, total: '200.00' }) }),
      makeCustomer({ n: 2, first: 'Ben', lastOrder: makeOrder({ n: 102, total: '50.00' }) }),
    ],
  });
  const maria = selection.recipients.find((r) => r.firstName === 'Maria');
  assert.ok(maria, 'fixture selected Maria');
  assert.equal(maria.amountCents, 1977);

  const { result, byVariant } = await previewAll(s, { seq: String(maria.seq) });
  assert.equal(result.exitCode, 0);
  assert.equal(byVariant.original.subject, 'LA Balloons $19.77 Online Credit Code');
  assert.ok(byVariant.original.html.includes('Hi Maria, here is your $19.77 online credit code'));
  for (const v of ['first', 'remind1', 'remind2']) {
    assert.ok(byVariant[v].html.includes('Hi Maria,'), `${v}: first name`);
    assert.ok(!byVariant[v].html.includes('Alex'));
    assert.ok(byVariant[v].html.includes(`名单第 ${maria.seq} 号顾客 Maria，礼品卡金额 $19.77`), `${v}: banner`);
  }
  assert.ok(read(result.index).includes(`名单第 ${maria.seq} 号顾客 Maria`));
  assert.ok(s.log.lines.some((l) => l.startsWith('INFO ') && l.includes(`名单第 ${maria.seq} 号`) && l.includes('$19.77')));
  // The list's expiry date is the one in .env: shown as is, nothing to point out.
  assert.equal(selection.params.giftCardExpiresOn, '2026-10-19');
  assert.ok(result.results.every((r) => r.expiresOn === '2026-10-19'));
  assert.deepEqual(warnings(s.log), []);
  assert.ok(read(result.index).includes('礼品卡到期日：2026-10-19。'));
});

test('--seq errors: no selection yet, unknown number, not a number', async () => {
  const s = setup();
  const noSelection = await runPreview({ config: s.config, seq: 1, templatesDir: s.templatesDir, log: s.log, now: NOW });
  assert.equal(noSelection.exitCode, 1);
  assert.ok(errors(s.log).some((l) => l.includes('还没有名单')));

  await selectionFixture(s.config, { customers: [makeCustomer({ n: 1, lastOrder: makeOrder({ n: 101, total: '80.00' }) })] });
  const unknown = await runPreview({ config: s.config, seq: 99, templatesDir: s.templatesDir, log: s.log, now: NOW });
  assert.equal(unknown.exitCode, 1);
  assert.ok(errors(s.log).some((l) => l.includes('名单里没有序号 99')));

  for (const seq of ['abc', '0', '-3', '1.5']) {
    const bad = await runPreview({ config: s.config, seq, templatesDir: s.templatesDir, log: s.log, now: NOW });
    assert.equal(bad.exitCode, 2, `seq ${seq}`);
  }
  assert.ok(!fs.existsSync(s.paths.previewDir), 'nothing was rendered');
});

test('index.html lists the four emails with their send dates, subjects and links', async () => {
  const s = setup();
  const { result, byVariant } = await previewAll(s);
  assert.equal(result.index, path.join(s.paths.previewDir, 'index.html'));
  const html = read(result.index);

  assert.ok(html.includes('<th>变体</th><th>模拟发送日期</th><th>主题</th><th>文件链接</th>'));
  const rows = html.match(/<tr><td>.*<\/tr>/g);
  assert.equal(rows.length, 4);
  VARIANTS.forEach((variant, i) => {
    assert.ok(rows[i].includes(`（${variant}）`), `row ${i}: ${variant}`);
    assert.ok(rows[i].includes(`<a href="${variant}.html">${variant}.html</a>`));
    assert.ok(rows[i].includes(`<td>${byVariant[variant].sendDate}</td>`));
  });
  assert.ok(rows[0].includes(CHA_CHING));
  assert.ok(rows[1].includes("Don't forget"));
  assert.ok(rows[2].includes('Last chance'));
  assert.ok(rows[3].includes('$15.33'));
  assert.ok(html.includes('首封') && html.includes('第一次提醒') && html.includes('第二次提醒') && html.includes('原版'));
  assert.ok(html.includes('生成于 2026-10-02 10:00'));
  assert.ok(s.log.lines.includes(`INFO 总览：${result.index}`));

  // A single variant: only that email is rendered and listed.
  const one = await runPreview({ config: s.config, variant: 'remind2', templatesDir: s.templatesDir, log: s.log, now: NOW });
  assert.equal(one.exitCode, 0);
  assert.deepEqual(one.files, [path.join(s.paths.previewDir, 'remind2.html')]);
  const oneRows = read(one.index).match(/<tr><td>.*<\/tr>/g);
  assert.equal(oneRows.length, 1);
  assert.ok(oneRows[0].includes('remind2.html'));
});

test('--open opens the overview (all) or the single email; errors from the opener are ignored', async () => {
  const s = setup();
  const opened = [];
  const openFile = (file) => opened.push(file);

  const all = await runPreview({ config: s.config, open: true, openFile, templatesDir: s.templatesDir, log: s.log, now: NOW });
  assert.equal(all.exitCode, 0);
  assert.deepEqual(opened, [all.index]);

  const one = await runPreview({ config: s.config, variant: 'remind1', open: true, openFile, templatesDir: s.templatesDir, log: s.log, now: NOW });
  assert.deepEqual(opened, [all.index, path.join(s.paths.previewDir, 'remind1.html')]);
  assert.equal(one.exitCode, 0);

  await runPreview({ config: s.config, openFile, templatesDir: s.templatesDir, log: s.log, now: NOW });
  assert.equal(opened.length, 2, 'not opened without --open');

  const failing = await runPreview({ config: s.config, open: true, openFile: async () => { throw new Error('no browser'); }, templatesDir: s.templatesDir, log: s.log, now: NOW });
  assert.equal(failing.exitCode, 0);
});

test('an unknown variant is a usage error (exit 2) and changes nothing', async () => {
  const s = setup({ REMIND_1_DATE: '2026-10-13' });
  const result = await runPreview({ config: s.config, variant: 'remind3', templatesDir: s.templatesDir, log: s.log, now: NOW });
  assert.equal(result.exitCode, 2);
  assert.deepEqual(result.files, []);
  assert.ok(errors(s.log).some((l) => l.includes('未知的预览类型 "remind3"')));
  assert.equal(read(s.bodyFile), BASE.body, 'validation happens before the templates are synced');
  assert.ok(!fs.existsSync(s.paths.previewDir));

  await assert.rejects(renderPreviewFor({ config: s.config, paths: s.paths, variant: 'all', templatesDir: s.templatesDir, log: s.log }), /未知的预览类型/);
  // Variant names are not case sensitive on the command line.
  const upper = await runPreview({ config: s.config, variant: 'FIRST', templatesDir: s.templatesDir, log: s.log, now: NOW });
  assert.equal(upper.exitCode, 0);
});

test('files are written under the campaign preview folder, with the banner right after <body>', async () => {
  const s = setup();
  const { result, byVariant } = await previewAll(s);

  assert.equal(s.paths.previewDir, path.join(s.config.campaignsDir, s.config.campaignId, 'preview'));
  for (const variant of VARIANTS) {
    const r = byVariant[variant];
    assert.equal(r.file, path.join(s.paths.previewDir, `${variant}.html`));
    assert.ok(fs.statSync(r.file).isFile());
    const banner = `主题：${r.subject}，模拟发送日期：${r.sendDate}，变体：${variant}`;
    assert.ok(r.html.includes(banner), `${variant}: banner text`);
    const bodyAt = r.html.indexOf('<body>');
    assert.ok(bodyAt > 0 && bodyAt < r.html.indexOf(banner) && r.html.indexOf(banner) < r.html.indexOf('<table class="body">'), `${variant}: banner position`);
    assert.ok(s.log.lines.includes(`INFO   主题：${r.subject}`));
    assert.ok(s.log.lines.includes(`INFO   文件：${r.file}`));
  }
  assert.deepEqual(fs.readdirSync(s.paths.previewDir).sort(), ['first.html', 'index.html', 'original.html', 'remind1.html', 'remind2.html']);
  assert.deepEqual(result.files, VARIANTS.map((v) => path.join(s.paths.previewDir, `${v}.html`)));
});

// ---------------------------------------------------------------------------
// The expiry date shown: .env, or the one frozen in the list
// ---------------------------------------------------------------------------

test('renderPreviewFor: expiresOn overrides GIFT_CARD_EXPIRES_ON; "" / null show no date; omitted uses .env; malformed is refused', async () => {
  const s = setup(); // .env: GIFT_CARD_EXPIRES_ON=2026-10-19
  const render = (variant, extra = {}) => renderPreviewFor({ config: s.config, paths: s.paths, variant, templatesDir: s.templatesDir, log: s.log, ...extra });

  // The list was made when the expiry was 10/26: the cards carry that date, so the email shows it.
  for (const variant of ['first', 'remind1', 'remind2']) {
    const r = await render(variant, { expiresOn: '2026-10-26' });
    assert.equal(r.expiresOn, '2026-10-26');
    const email = emailOnly(read(r.file));
    assert.ok(email.includes('PROMOTIONAL GIFT CARD · VALID THROUGH OCTOBER 26, 2026'), `${variant}: legal line`);
    assert.ok(!email.includes('OCTOBER 19') && !email.includes('October 19'), `${variant}: not the .env date`);
    if (variant !== 'first') assert.ok(r.subject.endsWith('expires October 26'), `${variant}: ${r.subject}`);
  }

  // Omitted (undefined): GIFT_CARD_EXPIRES_ON.
  const fromEnv = await render('remind1');
  assert.equal(fromEnv.expiresOn, '2026-10-19');
  assert.ok(read(fromEnv.file).includes(LEGAL_LINE));
  assert.equal((await render('remind1', { expiresOn: undefined })).expiresOn, '2026-10-19');

  // The list has no expiry date ('' or null): the cards never expire, so no date and no legal line.
  for (const none of ['', null]) {
    const r = await render('first', { expiresOn: none });
    assert.equal(r.expiresOn, null);
    const email = emailOnly(read(r.file));
    assert.ok(!email.includes('VALID THROUGH') && !email.includes('valid through'), `no expiry text for ${JSON.stringify(none)}`);
    assert.equal(r.stage, 'first');
  }

  // A malformed date is refused before anything is written (issue/remind turn this into a warning).
  fs.rmSync(s.paths.previewDir, { recursive: true, force: true });
  await assert.rejects(render('first', { expiresOn: '10/26/2026' }), /礼品卡到期日必须是 YYYY-MM-DD，收到 "10\/26\/2026"/);
  assert.ok(!fs.existsSync(path.join(s.paths.previewDir, 'first.html')));
  assert.deepEqual(warnings(s.log), []);

  // The pure helpers.
  assert.equal(previewExpiresOn(s.config), '2026-10-19');
  assert.equal(previewExpiresOn(s.config, '2026-10-26'), '2026-10-26');
  assert.equal(previewExpiresOn(s.config, ''), null);
  assert.equal(previewExpiresOn(s.config, null), null);
  assert.equal(previewExpiresOn(setup({ GIFT_CARD_EXPIRES_ON: '' }).config), null);
  assert.equal(previewData(s.config, 'first', null, { expiresOn: '2026-10-26' }).gift_card.expires_on, '2026-10-26');
  assert.equal(previewData(s.config, 'first').gift_card.expires_on, '2026-10-19');
});

test('preview --seq shows the expiry frozen in the list (the date the cards get) and says when .env differs', async () => {
  const s = setup(); // .env now: GIFT_CARD_EXPIRES_ON=2026-10-19
  const selection = await selectionFixture(s.config, {
    customers: [makeCustomer({ n: 1, first: 'Maria', lastOrder: makeOrder({ n: 101, total: '200.00' }) })],
  });
  // The list was made while .env still said 2026-10-26.
  writeJsonAtomic(s.paths.selection, { ...selection, params: { ...selection.params, giftCardExpiresOn: '2026-10-26' } });

  const { result, byVariant } = await previewAll(s, { seq: '1' });
  assert.equal(result.exitCode, 0);
  for (const v of ['first', 'remind1', 'remind2']) {
    assert.equal(byVariant[v].expiresOn, '2026-10-26');
    assert.ok(emailOnly(byVariant[v].html).includes('VALID THROUGH OCTOBER 26, 2026'), `${v}: the list's date`);
    assert.ok(!byVariant[v].html.includes('OCTOBER 19'), `${v}: not the .env date`);
  }
  assert.equal(byVariant.remind2.subject, '⌛ Last chance: your LA Balloons credit expires October 26');
  const drift = warnings(s.log).filter((l) => l.includes('名单生成时的礼品卡到期日'));
  assert.deepEqual(drift, ['WARN 名单生成时的礼品卡到期日是 2026-10-26，.env 的 GIFT_CARD_EXPIRES_ON 现在是 2026-10-19：卡按名单里的日期建，预览也按名单里的日期显示']);
  assert.ok(read(result.index).includes('礼品卡到期日：2026-10-26（按名单生成时的日期，卡按这个日期建；.env 现在是 2026-10-19）'), read(result.index));

  // Without --seq the sample data keeps using .env's date, with no drift warning.
  const sample = setup();
  const plain = await previewAll(sample);
  assert.ok(plain.result.results.every((r) => r.expiresOn === '2026-10-19'));
  assert.ok(read(plain.result.index).includes('礼品卡到期日：2026-10-19。'));
  assert.ok(!sample.log.lines.some((l) => l.includes('名单生成时的礼品卡到期日')));

  // A list made without an expiry date: the cards never expire, so the preview shows none.
  const none = setup();
  const noExpiry = await selectionFixture(none.config, { customers: [makeCustomer({ n: 1, lastOrder: makeOrder({ n: 101, total: '200.00' }) })] });
  writeJsonAtomic(none.paths.selection, { ...noExpiry, params: { ...noExpiry.params, giftCardExpiresOn: '' } });
  const shown = await previewAll(none, { seq: 1, variant: 'first' });
  assert.equal(shown.result.exitCode, 0);
  assert.equal(shown.byVariant.first.expiresOn, null);
  assert.ok(!emailOnly(shown.byVariant.first.html).includes('VALID THROUGH'));
  assert.ok(none.log.lines.includes('WARN 名单生成时的礼品卡到期日是 （未设置），.env 的 GIFT_CARD_EXPIRES_ON 现在是 2026-10-19：卡按名单里的日期建，预览也按名单里的日期显示'), none.log.lines.join('\n'));
});

test('renderPreviewFor works with only config and variant (as the issue/remind dry runs call it)', async () => {
  const s = setup();
  // Default templatesDir = the real notifications/ folder (read only); default paths = campaignPaths(config).
  const r = await renderPreviewFor({ config: s.config, variant: 'first', recipient: { seq: 1, firstName: 'Maria', amountCents: 1077 }, log: s.log });
  assert.equal(r.file, path.join(s.paths.previewDir, 'first.html'));
  assert.ok(r.subject.length > 0);
  assert.ok(read(r.file).includes('Maria'));

  const nameless = await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'first', recipient: { seq: 7, firstName: '', amountCents: 1077 }, templatesDir: s.templatesDir, log: s.log });
  assert.ok(read(nameless.file).includes('Hi , we'), 'shows exactly what Shopify would print');
  assert.ok(warnings(s.log).some((l) => l.includes('名单第 7 号顾客没有名字')));
});
