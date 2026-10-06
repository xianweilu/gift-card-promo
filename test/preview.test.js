// Tests for src/preview.js: local rendering of the "Gift card created" email
// templates (the promo email and the original one).
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
  renderPreviewFor,
  runPreview,
  createPreviewEngine,
  previewData,
  previewExpiresOn,
  parseStage,
  expectedStage,
  simulatedSendDate,
  SUBJECT_MAX_LENGTH,
  subjectLength,
  subjectHardcodedDate,
  subjectTemplateWarnings,
} from '../src/preview.js';

const REAL_DIR = path.join(ROOT_DIR, 'notifications');
const REAL = {
  subject: fs.readFileSync(path.join(REAL_DIR, TEMPLATE_FILES.subject), 'utf8'),
  body: fs.readFileSync(path.join(REAL_DIR, TEMPLATE_FILES.body), 'utf8'),
};

// The promo subject. Its date is hardcoded (Shopify caps the subject template at
// 512 characters), so it reads "Oct. 19th" whatever the cards' expiry date is;
// testConfig's GIFT_CARD_EXPIRES_ON is 2026-10-19, which matches.
const PROMO_SUBJECT = '💰 Cha-Ching! You’ve got LA Balloons CASH waiting for you to spend by Oct. 19th!';
const ORIGINAL_SUBJECT = 'LA Balloons $15.33 Online Credit Code';
const MISMATCH = '预览的文案和预期不一致';
const SUBJECT_DATE_MISMATCH = '主题模板里写死的到期日 "Oct. 19th" 和礼品卡到期日';
const SUBJECT_NO_EXPIRY = '主题模板里写死了到期日 "Oct. 19th"，但这些卡没有到期日';
const NOW = () => new Date('2026-10-02T17:00:00Z'); // 10:00 in Los Angeles
const execFileAsync = promisify(execFile);

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
function setup(overrides = {}, { subject = REAL.subject, body = REAL.body } = {}) {
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
// Rendering the two emails
// ---------------------------------------------------------------------------

test('renders the two emails: the promo copy for campaign cards, the original for all others', async () => {
  const s = setup();
  const { result, byVariant } = await previewAll(s);

  assert.equal(result.exitCode, 0);
  assert.deepEqual(VARIANTS, ['first', 'original']);
  assert.deepEqual(result.results.map((r) => r.variant), VARIANTS);
  assert.deepEqual(result.results.map((r) => [r.stage, r.expectedStage]), [['promo', 'promo'], ['original', 'original']]);
  assert.ok(byVariant.first.html.includes('<!-- promo-stage: promo -->'));
  assert.ok(byVariant.original.html.includes('<!-- promo-stage: original -->'));
  assert.equal(parseStage(byVariant.first.html), 'promo');
  assert.equal(parseStage(byVariant.original.html), 'original');
  for (const variant of VARIANTS) assert.doesNotMatch(byVariant[variant].subject, /\n/, 'a subject is one line');
  assert.equal(result.results.every((r) => r.expiresOn === '2026-10-19'), true);
  assert.deepEqual(warnings(s.log), []);
  assert.deepEqual(errors(s.log), []);
  // The templates no longer carry reminder dates or several promo stages.
  assert.doesNotMatch(REAL.subject, /promo_remind/);
  assert.doesNotMatch(REAL.body, /promo_remind|remind1|remind2/);
});

test('subjects: the promo email names the hardcoded expiry day, the original shows the amount', async () => {
  const s = setup();
  const { byVariant } = await previewAll(s);
  assert.equal(byVariant.first.subject, PROMO_SUBJECT);
  assert.equal(byVariant.original.subject, ORIGINAL_SUBJECT);
  assert.ok(!byVariant.first.subject.includes('$15.33'), 'the promo subject never shows the amount');
  assert.deepEqual(warnings(s.log), [], 'the hardcoded date matches GIFT_CARD_EXPIRES_ON');
});

test('the real subject template fits Shopify\'s 512-character limit and hardcodes Oct. 19', () => {
  assert.equal(SUBJECT_MAX_LENGTH, 512);
  assert.ok(subjectLength(REAL.subject) <= SUBJECT_MAX_LENGTH, `subject template is ${subjectLength(REAL.subject)} characters`);
  assert.ok(subjectLength(REAL.subject) < 300, 'the one-liner stays short');
  assert.doesNotMatch(REAL.subject, /\n(?!$)/, 'a single line');
  assert.doesNotMatch(REAL.subject, /expires_on|promo_day|promo_sfx/, 'no date logic in the subject');
  assert.deepEqual(subjectHardcodedDate(REAL.subject), { text: 'Oct. 19th', month: 10, day: 19 });
  assert.deepEqual(subjectTemplateWarnings(REAL.subject, '2026-10-19'), []);
});

test('subject checks: length counts characters like Shopify; a hardcoded date is compared with the expiry', () => {
  assert.equal(subjectLength('💰 ab\n'), 4, 'code points, trailing newline ignored');
  assert.equal(subjectLength(''), 0);
  assert.equal(subjectLength(null), 0);

  // Dates in literal text only; Liquid tags and outputs are ignored.
  assert.deepEqual(subjectHardcodedDate('spend by Oct. 19th!'), { text: 'Oct. 19th', month: 10, day: 19 });
  assert.deepEqual(subjectHardcodedDate('spend by November 1!'), { text: 'November 1', month: 11, day: 1 });
  assert.deepEqual(subjectHardcodedDate('by Dec 25th'), { text: 'Dec 25th', month: 12, day: 25 });
  assert.equal(subjectHardcodedDate("{% if x %}by {{ gift_card.expires_on | date: '%b. %-d' }}{% endif %}"), null);
  assert.equal(subjectHardcodedDate('Cha-Ching! CASH waiting for you!'), null);
  assert.equal(subjectHardcodedDate(''), null);

  assert.deepEqual(subjectTemplateWarnings('by Oct. 19th', '2026-10-19'), []);
  assert.deepEqual(subjectTemplateWarnings('by Oct. 19th', '2027-10-19'), [], 'the year is not in the subject');
  assert.deepEqual(subjectTemplateWarnings('no date here', null), []);
  assert.deepEqual(subjectTemplateWarnings('no date here', '2026-10-26'), []);
  assert.deepEqual(subjectTemplateWarnings('by Oct. 19th', '2026-10-26'), [
    `主题模板里写死的到期日 "Oct. 19th" 和礼品卡到期日 2026-10-26 不一致：请改 ${TEMPLATE_FILES.subject} 后重新贴到 Shopify 后台`,
  ]);
  assert.deepEqual(subjectTemplateWarnings('by Oct. 19th', '2026-11-19'), [
    `主题模板里写死的到期日 "Oct. 19th" 和礼品卡到期日 2026-11-19 不一致：请改 ${TEMPLATE_FILES.subject} 后重新贴到 Shopify 后台`,
  ]);
  assert.deepEqual(subjectTemplateWarnings('by Oct. 19th', null), [
    `主题模板里写死了到期日 "Oct. 19th"，但这些卡没有到期日：请改 ${TEMPLATE_FILES.subject} 后重新贴到 Shopify 后台`,
  ]);

  const long = `{% if a %}${'x'.repeat(600)}{% endif %}`;
  assert.deepEqual(subjectTemplateWarnings(long, '2026-10-19'), [
    `主题模板有 ${subjectLength(long)} 个字符，超过 Shopify 的上限 512 个，后台会拒绝保存：请缩短 ${TEMPLATE_FILES.subject}`,
  ]);
  assert.deepEqual(subjectTemplateWarnings('x'.repeat(512), '2026-10-19'), [], 'exactly 512 is accepted');
  assert.equal(subjectTemplateWarnings(`${'x'.repeat(600)} by Oct. 19th`, '2026-10-26').length, 2, 'both problems are reported');
});

test('a too-long subject template is flagged when the promo email is rendered, once per run', async () => {
  const subject = `{%- assign pad = '${'y'.repeat(560)}' -%}${REAL.subject}`;
  const s = setup({}, { subject });
  const { result, byVariant } = await previewAll(s);
  assert.equal(result.exitCode, 0, 'a warning, not a failure');
  assert.equal(byVariant.first.subject, PROMO_SUBJECT, 'the subject itself still renders');
  const tooLong = warnings(s.log).filter((l) => l.includes('超过 Shopify 的上限 512 个'));
  assert.equal(tooLong.length, 1, s.log.lines.join('\n'));
  assert.ok(tooLong[0].includes(`主题模板有 ${subjectLength(subject)} 个字符`));
});

test('the promo body: the image copy, every date from the card expiry, no amount; the original is unchanged', async () => {
  const s = setup();
  const { byVariant } = await previewAll(s);
  const promo = emailOnly(byVariant.first.html);

  assert.ok(promo.includes('<p>Hello Alex,</p>'));
  assert.ok(promo.includes('We did some "Balloon Math" and found some funds for you!'));
  assert.ok(promo.includes('You have an <strong>unused credit</strong> at <strong>LA Balloons</strong> waiting for you!'));
  assert.ok(promo.includes('It’s yours to use at <a href="https://www.laballoons.com">LABalloons.com</a>, but you must use it or lose it by <strong>October 19, 2026</strong>.'));
  assert.ok(promo.includes('Click below to claim your credit code before it\'s gone:'));
  assert.equal(buttonText(byVariant.first.html), 'CLAIM IT NOW');
  assert.ok(!promo.includes('Promotional reward only'), 'the Promotional reward line was dropped (10/6)');
  assert.ok(promo.includes('<span style="color: #FF2600; font-size: 12px">No cash value &amp; non-transferable.</span>'), 'red T&C sentence');
  assert.ok(promo.includes('No strings attached!'));
  assert.ok(promo.includes('We appreciate your business, and thank you for trusting us with your balloon and party supply needs!'));
  assert.ok(promo.includes('<p>- LA Balloons</p>'));
  assert.ok(promo.includes('** Terms &amp; Conditions:'));
  assert.ok(promo.includes('exempt from California’s mandatory gift certificate cash-out regulations'));
  assert.ok(promo.includes('Expires on 10/19/2026.'));
  assert.ok(!promo.includes('$15.33') && !promo.includes('15.33'), 'the promo copy never shows the amount');
  assert.ok(!promo.includes('October 18') && !promo.includes('10/18/2026'), 'no day shift');
  assert.ok(!promo.includes('VALID THROUGH'), 'the old legal line is gone');
  assert.equal(emailTitle(byVariant.first.html), '', 'the promo email has no title heading');
  // Shop name heading (no email logo), the store link, the footer.
  assert.ok(promo.includes('<a href="https://www.laballoons.com">LA Balloons</a>'));
  assert.ok(!promo.includes('<img src=""'));
  assert.ok(promo.includes('Visit our store'));
  assert.ok(promo.includes('1422 Gardena Ave., Glendale, CA 91204'));

  const original = emailOnly(byVariant.original.html);
  assert.equal(emailTitle(byVariant.original.html), 'Your Online Credit Code is now available!');
  assert.ok(original.includes('Hi Alex, here is your $15.33 online credit code.'));
  assert.equal(buttonText(byVariant.original.html), 'View your Online Credit Code');
  assert.ok(!original.includes('Balloon Math') && !original.includes('Promotional reward') && !original.includes('October 19, 2026'));
});

test('the subject date is hardcoded: the same "Oct. 19th" for every expiry, with a warning when the expiry differs', async () => {
  const s = setup();
  const render = async (expiresOn) => renderPreviewFor({ config: s.config, paths: s.paths, variant: 'first', templatesDir: s.templatesDir, log: s.log, expiresOn });

  // Matching expiry: no warning.
  assert.equal((await render('2026-10-19')).subject, PROMO_SUBJECT);
  assert.deepEqual(warnings(s.log), []);

  // Other expiry dates: the body follows the card, the subject does not, and the run says so.
  for (const date of ['2026-10-01', '2026-10-20', '2026-11-19', '2026-12-25']) {
    const r = await render(date);
    assert.equal(r.subject, PROMO_SUBJECT, date);
    assert.ok(emailOnly(read(r.file)).includes(`${date.slice(5, 7).replace(/^0/, '')}/${date.slice(8).replace(/^0/, '')}/2026`), `${date}: body date`);
    const w = warnings(s.log).filter((l) => l.includes(SUBJECT_DATE_MISMATCH));
    assert.equal(w.at(-1), `WARN ${SUBJECT_DATE_MISMATCH} ${date} 不一致：请改 ${TEMPLATE_FILES.subject} 后重新贴到 Shopify 后台`, date);
  }
  assert.equal(warnings(s.log).length, 4, 'one warning per mismatching render');

  // The original email shares the subject file but is not about the promo date: no warning.
  const before = warnings(s.log).length;
  await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'original', templatesDir: s.templatesDir, log: s.log, expiresOn: '2026-11-19' });
  assert.equal(warnings(s.log).length, before);
});

test('a card without an expiry date: the body sentences leave the date out; the hardcoded subject date is flagged', async () => {
  const s = setup();
  const r = await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'first', templatesDir: s.templatesDir, log: s.log, expiresOn: '' });
  assert.equal(r.expiresOn, null);
  assert.equal(r.subject, PROMO_SUBJECT, 'the subject cannot follow the card');
  assert.deepEqual(warnings(s.log), [`WARN ${SUBJECT_NO_EXPIRY}：请改 ${TEMPLATE_FILES.subject} 后重新贴到 Shopify 后台`]);
  const email = emailOnly(read(r.file));
  assert.ok(email.includes('It’s yours to use at <a href="https://www.laballoons.com">LABalloons.com</a>.</p>'));
  assert.ok(!email.includes('use it or lose it'));
  assert.ok(email.includes('will not be replaced.</p>'), 'the T&C ends without an expiry sentence');
  assert.ok(!email.includes('expires') && !email.includes('Expires on'));
  assert.equal(r.stage, 'promo');
});

test('dates stay on their calendar day whatever time zone the machine is in', async () => {
  const s = setup();
  const before = process.env.TZ;
  try {
    for (const tz of ['America/Los_Angeles', 'Asia/Shanghai', 'Pacific/Kiritimati', 'Etc/GMT+12', 'UTC']) {
      process.env.TZ = tz;
      const r = await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'first', templatesDir: s.templatesDir, log: s.log, now: NOW });
      assert.equal(r.stage, 'promo', tz);
      assert.equal(r.subject, PROMO_SUBJECT, tz);
      const email = emailOnly(read(r.file));
      assert.ok(email.includes('October 19, 2026') && email.includes('10/19/2026'), `${tz}: body dates`);
      assert.ok(!email.includes('October 18') && !email.includes('October 20'), `${tz}: no day shift`);
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
    const out = await createPreviewEngine({ sendDate: '2026-10-12' }).parseAndRender("{{ '2026-10-19' | date: '%B %-d, %Y' }}|{{ '2026-10-19' | date: '%b' }}");
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
  assert.equal(out, 'October 19, 2026|Oct');
});

test('engine: now/today are the send date; money and asset filters behave like Shopify', async () => {
  const engine = createPreviewEngine({ sendDate: '2026-10-12' });
  const render = (tpl, data = {}) => engine.parseAndRender(tpl, data);

  assert.equal(await render("{{ 'now' | date: '%Y-%m-%d %H:%M' }}|{{ 'today' | date: '%Y%m%d' }}|{{ 'Now' | date: '%Y%m%d' | plus: 0 }}"), '2026-10-12 19:00|20261012|20261012');
  assert.equal(await render("{{ '2026-10-19' | date: '%B %-d, %Y' }}|{{ d | date: '%B %-d' }}|{{ d | date: '%-m/%-d/%Y' }}", { d: '2026-11-01' }), 'October 19, 2026|November 1|11/1/2026');

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

test('sample data, the simulated send date and the expected stage', () => {
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
  const real = previewData(config, 'first', { seq: 3, firstName: 'Maria', amountCents: 1977 });
  assert.equal(real.gift_card.customer.first_name, 'Maria');
  assert.equal(real.gift_card.balance, 1977);

  assert.equal(simulatedSendDate(config, 'first'), '2026-10-05');
  assert.equal(simulatedSendDate(config, 'original'), '2026-10-05');
  assert.throws(() => simulatedSendDate(config, 'remind1'), /未知的预览类型/);

  assert.equal(expectedStage('first'), 'promo');
  assert.equal(expectedStage('original'), 'original');
});

test('LAUNCH_DATE not set: the send date is today in the store time zone (the copy does not depend on it)', async () => {
  const s = setup({ LAUNCH_DATE: '' });
  // 2026-10-03 05:00 UTC is still October 2 in Los Angeles.
  const { result } = await previewAll(s, { now: () => new Date('2026-10-03T05:00:00Z') });
  assert.equal(result.exitCode, 0);
  const first = result.results.find((r) => r.variant === 'first');
  assert.equal(first.sendDate, '2026-10-02');
  assert.equal(first.stage, 'promo');
  assert.equal(first.subject, PROMO_SUBJECT);
  // Any send date shows the same copy: there is no date-based switch any more.
  for (const sendDate of ['2026-10-05', '2026-10-12', '2026-10-16', '2026-10-19', '2027-01-01']) {
    const r = await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'first', templatesDir: s.templatesDir, log: s.log, sendDate });
    assert.equal(r.stage, 'promo', sendDate);
    assert.equal(r.subject, PROMO_SUBJECT, sendDate);
    assert.ok(emailOnly(read(r.file)).includes('Balloon Math'), sendDate);
  }
  assert.deepEqual(warnings(s.log), []);
});

test('a template suffix that does not match the templates is flagged (the original copy would go out)', async () => {
  const s = setup({ GIFT_CARD_TEMPLATE_SUFFIX: 'other-suffix' });
  const r = await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'first', templatesDir: s.templatesDir, log: s.log });
  assert.equal(r.stage, 'original');
  assert.equal(r.expectedStage, 'promo');
  assert.ok(warnings(s.log).some((l) => l.includes(MISMATCH) && l.includes('GIFT_CARD_TEMPLATE_SUFFIX')));
  // The overview page says so too.
  const { result } = await previewAll(s);
  assert.equal(result.exitCode, 0, 'a mismatch is a warning, not a failure');
  assert.ok(read(result.index).includes('模板显示的是原版，应为活动文案'));
});

test('a body without the promo-stage marker is flagged', async () => {
  const body = REAL.body.replace(/<!--\s*promo-stage:[^>]*-->/, '');
  assert.ok(!body.includes('promo-stage'), 'test setup removed the marker');
  const s = setup({}, { body });
  const r = await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'first', templatesDir: s.templatesDir, log: s.log });
  assert.equal(r.stage, null);
  assert.ok(warnings(s.log).some((l) => l.includes(MISMATCH) && l.includes('没有 promo-stage 标记')));
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
  assert.ok(byVariant.first.html.includes('<p>Hello Maria,</p>'));
  assert.ok(!byVariant.first.html.includes('Alex'));
  assert.ok(!emailOnly(byVariant.first.html).includes('19.77'), 'the promo copy never shows the amount');
  assert.ok(byVariant.first.html.includes(`名单第 ${maria.seq} 号顾客 Maria，礼品卡金额 $19.77`), 'banner');
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

test('index.html lists the two emails with their subjects and links', async () => {
  const s = setup();
  const { result, byVariant } = await previewAll(s);
  assert.equal(result.index, path.join(s.paths.previewDir, 'index.html'));
  const html = read(result.index);

  assert.ok(html.includes('<th>变体</th><th>主题</th><th>文件链接</th>'));
  const rows = html.match(/<tr><td>.*<\/tr>/g);
  assert.equal(rows.length, 2);
  VARIANTS.forEach((variant, i) => {
    assert.ok(rows[i].includes(`（${variant}）`), `row ${i}: ${variant}`);
    assert.ok(rows[i].includes(`<a href="${variant}.html">${variant}.html</a>`));
    assert.ok(rows[i].includes(`<td>${byVariant[variant].subject}</td>`));
  });
  assert.ok(rows[0].includes('Cha-Ching'));
  assert.ok(rows[1].includes('$15.33'));
  assert.ok(html.includes('首封') && html.includes('原版'));
  assert.ok(!html.includes('第一次提醒') && !html.includes('第二次提醒') && !html.includes('REMIND'));
  assert.ok(html.includes('提醒邮件用 Shopify Email 发送'));
  assert.ok(html.includes('生成于 2026-10-02 10:00'));
  assert.ok(s.log.lines.includes(`INFO 总览：${result.index}`));

  // A single variant: only that email is rendered and listed.
  const one = await runPreview({ config: s.config, variant: 'original', templatesDir: s.templatesDir, log: s.log, now: NOW });
  assert.equal(one.exitCode, 0);
  assert.deepEqual(one.files, [path.join(s.paths.previewDir, 'original.html')]);
  const oneRows = read(one.index).match(/<tr><td>.*<\/tr>/g);
  assert.equal(oneRows.length, 1);
  assert.ok(oneRows[0].includes('original.html'));
});

test('--open opens the overview (all) or the single email; errors from the opener are ignored', async () => {
  const s = setup();
  const opened = [];
  const openFile = (file) => opened.push(file);

  const all = await runPreview({ config: s.config, open: true, openFile, templatesDir: s.templatesDir, log: s.log, now: NOW });
  assert.equal(all.exitCode, 0);
  assert.deepEqual(opened, [all.index]);

  const one = await runPreview({ config: s.config, variant: 'first', open: true, openFile, templatesDir: s.templatesDir, log: s.log, now: NOW });
  assert.deepEqual(opened, [all.index, path.join(s.paths.previewDir, 'first.html')]);
  assert.equal(one.exitCode, 0);

  await runPreview({ config: s.config, openFile, templatesDir: s.templatesDir, log: s.log, now: NOW });
  assert.equal(opened.length, 2, 'not opened without --open');

  const failing = await runPreview({ config: s.config, open: true, openFile: async () => { throw new Error('no browser'); }, templatesDir: s.templatesDir, log: s.log, now: NOW });
  assert.equal(failing.exitCode, 0);
});

test('an unknown variant (the old reminder variants included) is a usage error (exit 2) and changes nothing', async () => {
  const s = setup();
  for (const variant of ['remind1', 'remind2', 'remind3', 'bogus']) {
    const result = await runPreview({ config: s.config, variant, templatesDir: s.templatesDir, log: s.log, now: NOW });
    assert.equal(result.exitCode, 2, variant);
    assert.deepEqual(result.files, []);
    assert.ok(errors(s.log).some((l) => l.includes(`未知的预览类型 "${variant}"`) && l.includes('first（首封）、original（原版）')));
  }
  assert.ok(!fs.existsSync(s.paths.previewDir));

  await assert.rejects(renderPreviewFor({ config: s.config, paths: s.paths, variant: 'all', templatesDir: s.templatesDir, log: s.log }), /未知的预览类型/);
  await assert.rejects(renderPreviewFor({ config: s.config, paths: s.paths, variant: 'remind1', templatesDir: s.templatesDir, log: s.log }), /未知的预览类型/);
  // Variant names are not case sensitive on the command line.
  const upper = await runPreview({ config: s.config, variant: 'FIRST', templatesDir: s.templatesDir, log: s.log, now: NOW });
  assert.equal(upper.exitCode, 0);
});

test('a missing template is reported (exit 1), nothing is written', async () => {
  const s = setup();
  fs.rmSync(s.subjectFile);
  const result = await runPreview({ config: s.config, templatesDir: s.templatesDir, log: s.log, now: NOW, open: true, openFile: () => assert.fail('must not open') });
  assert.equal(result.exitCode, 1);
  assert.deepEqual(result.files, []);
  assert.ok(errors(s.log).some((l) => l.includes('找不到邮件模板') && l.includes(TEMPLATE_FILES.subject)));
  assert.ok(!fs.existsSync(s.paths.previewDir), 'no preview was written');
  await assert.rejects(renderPreviewFor({ config: s.config, paths: s.paths, variant: 'first', templatesDir: s.templatesDir, log: s.log }), /找不到邮件模板/);
});

test('files are written under the campaign preview folder, with the banner right after <body>', async () => {
  const s = setup();
  const { result, byVariant } = await previewAll(s);

  assert.equal(s.paths.previewDir, path.join(s.config.campaignsDir, s.config.campaignId, 'preview'));
  for (const variant of VARIANTS) {
    const r = byVariant[variant];
    assert.equal(r.file, path.join(s.paths.previewDir, `${variant}.html`));
    assert.ok(fs.statSync(r.file).isFile());
    const banner = `主题：${r.subject}，变体：${variant}`;
    assert.ok(r.html.includes(banner), `${variant}: banner text`);
    assert.ok(r.html.includes('礼品卡到期日：2026-10-19'), `${variant}: banner expiry`);
    const bodyAt = r.html.indexOf('<body>');
    assert.ok(bodyAt > 0 && bodyAt < r.html.indexOf(banner) && r.html.indexOf(banner) < r.html.indexOf('<table class="body">'), `${variant}: banner position`);
    assert.ok(s.log.lines.includes(`INFO   主题：${r.subject}`));
    assert.ok(s.log.lines.includes(`INFO   文件：${r.file}`));
  }
  assert.deepEqual(fs.readdirSync(s.paths.previewDir).sort(), ['first.html', 'index.html', 'original.html']);
  assert.deepEqual(result.files, VARIANTS.map((v) => path.join(s.paths.previewDir, `${v}.html`)));
});

test('the real notifications/ templates are never modified by preview', async () => {
  const s = setup();
  await previewAll(s);
  await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'first', templatesDir: s.templatesDir, log: s.log, expiresOn: '2026-10-26' });
  assert.equal(read(s.subjectFile), REAL.subject);
  assert.equal(read(s.bodyFile), REAL.body);
});

// ---------------------------------------------------------------------------
// The expiry date shown: .env, or the one frozen in the list
// ---------------------------------------------------------------------------

test('renderPreviewFor: expiresOn overrides GIFT_CARD_EXPIRES_ON; "" / null show no date; omitted uses .env; malformed is refused', async () => {
  const s = setup(); // .env: GIFT_CARD_EXPIRES_ON=2026-10-19
  const render = (variant, extra = {}) => renderPreviewFor({ config: s.config, paths: s.paths, variant, templatesDir: s.templatesDir, log: s.log, ...extra });

  // The list was made when the expiry was 10/26: the cards carry that date, so the body shows it.
  // The subject's date is hardcoded, so it still says Oct. 19th and the run warns.
  const r = await render('first', { expiresOn: '2026-10-26' });
  assert.equal(r.expiresOn, '2026-10-26');
  assert.equal(r.subject, PROMO_SUBJECT);
  const email = emailOnly(read(r.file));
  assert.ok(email.includes('use it or lose it by <strong>October 26, 2026</strong>'));
  assert.ok(email.includes('Expires on 10/26/2026.'));
  assert.ok(!email.includes('October 19') && !email.includes('10/19/2026'), 'not the .env date');
  assert.deepEqual(warnings(s.log), [`WARN ${SUBJECT_DATE_MISMATCH} 2026-10-26 不一致：请改 ${TEMPLATE_FILES.subject} 后重新贴到 Shopify 后台`]);
  s.log.lines.length = 0;

  // Omitted (undefined): GIFT_CARD_EXPIRES_ON.
  const fromEnv = await render('first');
  assert.equal(fromEnv.expiresOn, '2026-10-19');
  assert.equal(fromEnv.subject, PROMO_SUBJECT);
  assert.equal((await render('first', { expiresOn: undefined })).expiresOn, '2026-10-19');

  // The list has no expiry date ('' or null): the cards never expire, so no date in the body;
  // the subject's hardcoded date is flagged.
  for (const none of ['', null]) {
    const r2 = await render('first', { expiresOn: none });
    assert.equal(r2.expiresOn, null);
    const e2 = emailOnly(read(r2.file));
    assert.ok(!e2.includes('use it or lose it') && !e2.includes('expires') && !e2.includes('Expires on'), `no expiry text for ${JSON.stringify(none)}`);
    assert.equal(r2.stage, 'promo');
  }
  assert.deepEqual(warnings(s.log), Array(2).fill(`WARN ${SUBJECT_NO_EXPIRY}：请改 ${TEMPLATE_FILES.subject} 后重新贴到 Shopify 后台`));
  s.log.lines.length = 0;

  // A malformed date is refused before anything is written (issue turns this into a warning).
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
  assert.equal(byVariant.first.expiresOn, '2026-10-26');
  assert.ok(emailOnly(byVariant.first.html).includes('October 26, 2026'), 'the list\'s date');
  assert.ok(!byVariant.first.html.includes('October 19'), 'not the .env date');
  assert.equal(byVariant.first.subject, PROMO_SUBJECT, 'the subject date is hardcoded');
  const drift = warnings(s.log).filter((l) => l.includes('名单生成时的礼品卡到期日'));
  assert.deepEqual(drift, ['WARN 名单生成时的礼品卡到期日是 2026-10-26，.env 的 GIFT_CARD_EXPIRES_ON 现在是 2026-10-19：卡按名单里的日期建，预览也按名单里的日期显示']);
  // …and the subject's hardcoded Oct. 19th no longer matches the cards: said once, for the promo email.
  assert.deepEqual(warnings(s.log).filter((l) => l.includes(SUBJECT_DATE_MISMATCH)), [`WARN ${SUBJECT_DATE_MISMATCH} 2026-10-26 不一致：请改 ${TEMPLATE_FILES.subject} 后重新贴到 Shopify 后台`]);
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
  assert.ok(!emailOnly(shown.byVariant.first.html).includes('use it or lose it'));
  assert.ok(none.log.lines.includes('WARN 名单生成时的礼品卡到期日是 （未设置），.env 的 GIFT_CARD_EXPIRES_ON 现在是 2026-10-19：卡按名单里的日期建，预览也按名单里的日期显示'), none.log.lines.join('\n'));
});

test('renderPreviewFor works as the issue dry run calls it: config, paths, variant first, recipient, log, expiresOn', async () => {
  const s = setup();
  // Default templatesDir = the real notifications/ folder (read only); default paths = campaignPaths(config).
  const r = await renderPreviewFor({ config: s.config, variant: 'first', recipient: { seq: 1, firstName: 'Maria', amountCents: 1077 }, log: s.log });
  assert.equal(r.file, path.join(s.paths.previewDir, 'first.html'));
  assert.equal(r.subject, PROMO_SUBJECT);
  assert.ok(read(r.file).includes('Hello Maria,'));

  // The call issue makes (selection.params.giftCardExpiresOn as expiresOn; a sendDate is accepted and harmless).
  const asIssue = await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'first', recipient: { seq: 2, firstName: 'Ben', amountCents: 1977 }, log: s.log, expiresOn: '2026-10-19', sendDate: '2026-10-05' });
  assert.equal(asIssue.subject, PROMO_SUBJECT);
  assert.equal(asIssue.stage, 'promo');
  assert.equal(asIssue.sendDate, '2026-10-05');

  const nameless = await renderPreviewFor({ config: s.config, paths: s.paths, variant: 'first', recipient: { seq: 7, firstName: '', amountCents: 1077 }, templatesDir: s.templatesDir, log: s.log });
  assert.ok(read(nameless.file).includes('<p>Hello,</p>'), 'shows exactly what Shopify would print');
  assert.ok(warnings(s.log).some((l) => l.includes('名单第 7 号顾客没有名字')));
});
