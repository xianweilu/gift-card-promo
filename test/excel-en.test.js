// The English workbook (gift-card-promo-<id>-en.xlsx): the TEXT.en pack and its stored-text hooks.
// The frozen scenarios of test/report-fixture.js (ASCII customer data, a journal with every op and
// every stored Chinese note / detail / error form, old-style notSelected texts, verify.json issues with
// and without *En fields, usage.json with unmatched reasons and '（无名称）') are written with the real
// writeReport; the English edition must contain no CJK or full-width character, mirror the Chinese
// edition's structure cell for cell, and show the English wording of the spec.
import { test, describe } from 'node:test';
import assert from 'node:assert/strict';

import { writeReport, englishWorkbookPath, statusLabel, remindLabel } from '../src/report/excel.js';
import { TEXT } from '../src/report/text.js';
import { errorText, summaryText, detailText, issueSkipText, remindSkipText, noteText, storedLabelOf, labelsFor } from '../src/report/labels.js';
import { ruleLabelEn, reasonTextEn, notSelectedReasonsEn, formulaEn, channelLabelEn, storedTextEn } from '../src/report/translate-en.js';
import { RULES } from '../src/select/rules.js';
import { loadScenarios, setupScenario, snapshotWorkbook, structureOf, cjkCells, CJK } from './report-fixture.js';
import { testConfig, selectionFixture, makeCustomer, makeOrder, gid } from './helpers.js';

const { zh, en } = TEXT;
const cust = (n) => gid('Customer', n);

// ---------------------------------------------------------------------------
// Helpers over snapshotWorkbook's shape: { name, rows: [[r, height, [[c, value, style], ...]], ...] }
// ---------------------------------------------------------------------------

const sheetNamed = (sheets, name) => {
  const s = sheets.find((x) => x.name === name);
  assert.ok(s, `sheet "${name}" exists (have ${sheets.map((x) => x.name).join(', ')})`);
  return s;
};
/** Row `r` (1-based) as an array indexed by column - 1 (a HYPERLINK cell → its label). */
const rowValues = (sheet, r) => {
  const row = sheet.rows.find(([n]) => n === r);
  const out = [];
  for (const [c, v] of row?.[2] ?? []) out[c - 1] = v && typeof v === 'object' && 'f' in v ? v.r : v;
  return out;
};
const allRows = (sheet) => sheet.rows.map(([r]) => rowValues(sheet, r));
/** Every string in a sheet (cell texts and HYPERLINK labels). */
const textsOf = (sheet) => allRows(sheet).flat().filter((v) => typeof v === 'string');
const findRow = (sheet, pred) => allRows(sheet).find(pred);

/** Every workbook of a scenario, both editions, as snapshots. */
async function writeScenario(scenario, patch = null) {
  const s = setupScenario(patch ? patch(structuredClone(scenario)) : scenario, writeReport);
  try {
    const result = await s.write();
    assert.equal(result.fileEn, englishWorkbookPath(s.paths.excel));
    return { result, zh: await snapshotWorkbook(result.file), en: await snapshotWorkbook(result.fileEn), paths: s.paths };
  } finally {
    s.cleanup();
  }
}

/** Keys of a pack value, recursively: "a.b.c" with the type of each leaf (array length, function, string, number). */
function shapeOf(value, prefix = '') {
  const out = [];
  if (Array.isArray(value)) {
    out.push(`${prefix} = array[${value.length}]`);
    value.forEach((v, i) => out.push(...shapeOf(v, `${prefix}[${i}]`)));
  } else if (typeof value === 'function') {
    out.push(`${prefix} = function/${value.length}`);
  } else if (value && typeof value === 'object') {
    for (const k of Object.keys(value).sort()) out.push(...shapeOf(value[k], prefix ? `${prefix}.${k}` : k));
  } else {
    out.push(`${prefix} = ${typeof value}`);
  }
  return out;
}

/** Every string leaf of a pack value with its path. */
function stringsOf(value, prefix = '') {
  const out = [];
  if (Array.isArray(value)) value.forEach((v, i) => out.push(...stringsOf(v, `${prefix}[${i}]`)));
  else if (typeof value === 'string') out.push([prefix, value]);
  else if (value && typeof value === 'object') for (const [k, v] of Object.entries(value)) out.push(...stringsOf(v, prefix ? `${prefix}.${k}` : k));
  return out;
}

const scenarios = loadScenarios();
const live = scenarios.find((s) => s.name === 'live');
const testCampaign = scenarios.find((s) => s.name === 'test');
const empty = scenarios.find((s) => s.name === 'empty');

// ---------------------------------------------------------------------------
// The pack
// ---------------------------------------------------------------------------

describe('TEXT.en pack', () => {
  test('has exactly the keys and shapes of TEXT.zh (arrays of the same length, functions where zh has functions)', () => {
    assert.equal(en.lang, 'en');
    assert.equal(zh.lang, 'zh');
    assert.deepEqual(shapeOf({ ...en, lang: 0 }), shapeOf({ ...zh, lang: 0 }));
  });

  test('every English text is free of CJK and full-width characters', () => {
    const bad = stringsOf({ ...en, lang: '' }).filter(([, s]) => CJK.test(s));
    assert.deepEqual(bad, []);
  });

  test('templated texts render English without CJK (sample arguments)', () => {
    const samples = [
      en.roundTagHelp('T-R1', 'T-R2').flat().join(' '),
      en.timeZone.other('Asia/Shanghai'),
      en.roundName(2),
      en.describeTiers([1077, 1533, 1977]).join(' '),
      en.describeTiers([1000]).join(' '),
      en.usageReport.title('2026-10-12 08:00', 'Los Angeles time'),
      en.usageReport.footer('x', 'y'),
      en.summary.subtitle('2026-10', true, 'now', 'tz'),
      en.summary.subtitle('2026-10', false, 'now', 'tz'),
      en.summary.exportedNote('Bulk', '42'),
      en.summary.cutoff(3),
      en.summary.days(7),
      en.summary.percentValue(10),
      en.summary.expiryChanged('a', 'b'),
      en.summary.dateChanged('a', 'b'),
      en.summary.sourceWithCode('Sellbrite', '205641'),
      en.summary.ruleRow(1, 'x'),
      en.summary.funnelNote('1', '2'),
      en.summary.testTier('$0.10'),
      en.summary.whyCancelled('1') + en.summary.whyTest('1') + en.summary.whyZero('1'),
      en.summary.personCount('1') + en.summary.personCount('2'),
      en.summary.remindDateChanged(1, 'note'),
      en.summary.remindNote(en.summary.audienceLive) + en.summary.remindNote(en.summary.audienceTest),
      en.recipients.columns.hasTag.title('TAG') + en.recipients.columns.hasTag.help('TAG'),
      en.recipients.columns.remind1.help('TAG-R1') + en.recipients.columns.remind2.help('TAG-R2'),
      en.recipients.remind.sentAt('10-12 09:05') + en.recipients.remind.skippedFor('x') + en.recipients.remind.failedWith('x'),
      en.recipients.notes.reconciledFrom('verify') + en.recipients.notes.amountDiffers('$1', '$2') + en.recipients.notes.bulk(10),
      en.duplicates.activeTitle(3, '1'),
      en.usageDetail.unmatchedTitle('3'),
      en.verify.title('a', 'b', '9', 'T', '7') + en.verify.countItem('x', '1') + en.verify.found('2', 'x') + en.verify.found('2', ''),
      en.journal.amount('$1') + en.journal.last4('x001') + en.journal.alreadyTagged('T') + en.journal.tagMissing('T') + en.journal.tagRepaired('T'),
      en.help.testAmount('$0.10') + en.help.baseText(10) + en.help.neverText(10) + en.help.tierRow(0, 'r').join(' ') + en.help.timeText('tz', 'l') + en.help.exitCode(2),
      en.storeTime('2026-10-04 11:30') + en.withDetail('a', 'b'),
      Object.values(en.shopifyErrors).map((f) => (typeof f === 'function' ? f('x', 'y') : f)).join(' '),
      Object.values(en.runSummary).map((f) => (typeof f === 'function' ? f('x', 'y') : f)).join(' '),
    ];
    for (const s of samples) assert.doesNotMatch(s, CJK, s);
    assert.equal(en.summary.personCount('1'), '1 person');
    assert.equal(en.summary.personCount('13'), '13 people');
    assert.deepEqual(en.describeTiers([1000]), ['any → $10.00']);
    assert.deepEqual(en.describeTiers([1077, 1533, 1977]), zh.describeTiers([1077, 1533, 1977]));
  });

  test('sheet names, statuses, kinds and marketing states follow the spec wording', () => {
    assert.deepEqual(Object.values(en.sheets), ['Usage report', 'Summary', 'Recipients', 'Not selected', 'Same address', 'Usage details', 'Verify', 'Activity log', 'Help']);
    assert.deepEqual(Object.values(en.STATUS_LABELS), ['Pending', 'In progress', 'Needs review', 'Card created, not tagged', 'Done', 'Failed', 'Skipped before issuing']);
    assert.deepEqual(Object.values(en.KIND_LABELS), ['Ordered before', 'Never ordered', 'Test']);
    assert.deepEqual(Object.values(en.MARKETING_LABELS), ['Subscribed', 'Not subscribed', 'Unsubscribed', 'Pending', 'Invalid', 'Redacted', 'No status']);
    assert.equal(en.timeZone.default, 'Los Angeles time');
    assert.equal(en.REMIND_SENT_BY_TAG, 'Sent (recorded from the round tag in Shopify)');
    assert.equal(labelsFor('en'), en);
    assert.equal(labelsFor(en), en);
  });

  test('number formats keep the money and date formats and translate the words', () => {
    assert.deepEqual(en.fmt, { cards: '#,##0" cards"', orders: '#,##0" orders"', share: '"share "0.0%', shareOfValue: '"of face value "0.0%' });
    for (const [k, v] of Object.entries(en.moneyPrefix)) {
      assert.doesNotMatch(v, /"/, `${k} has no quote (it goes inside a number format)`);
      assert.match(v, / $/, `${k} ends with a space like the Chinese prefix`);
    }
  });

  test('every run-summary key the commands write has an English label (keys as in src/*.js)', () => {
    const written = [
      'recipients', 'totalCents', 'source', 'snapshotReused',
      'attempted', 'created', 'tagged', 'tagFixed', 'reconciled', 'skipped', 'failed', 'rejected', 'unknown', 'amountCents', 'tagFailed', 'reconciledNone', 'stillUnknown', 'newCardsRefused', 'repairOnly', 'stoppedByDate',
      'eligible', 'planned', 'sent', 'retriedUnknown', 'retriedFailed', 'alreadySent', 'alreadySentByTag', 'previouslyFailed', 'waitingUnknown', 'notIssued', 'roundTagged', 'roundTagFailed', 'roundTagFixed', 'roundTagMissing',
      'issuedCards', 'usedCards', 'usedCents', 'orders', 'ordersTotalCents', 'giftCardCents', 'customerPaidCents', 'unmatched',
      'cardCount', 'taggedCount', 'issueCount', 'fixedCount', 'counts',
    ];
    for (const key of written) {
      assert.equal(typeof en.SUMMARY_KEY_LABELS[key], 'string', key);
      assert.ok(Object.hasOwn(zh.SUMMARY_KEY_LABELS, key), `${key} in zh too`);
    }
    assert.deepEqual(Object.keys(en.SUMMARY_KEY_LABELS), Object.keys(zh.SUMMARY_KEY_LABELS));
  });
});

// ---------------------------------------------------------------------------
// stored.* hooks and translate-en.js
// ---------------------------------------------------------------------------

describe('TEXT.en.stored', () => {
  const S = en.stored;

  test('ruleLabel: every rule from its code, with the configured numbers (params, else the stored label, else the defaults)', () => {
    const params = { minAccountAgeDays: 14, inactiveMonths: 6 };
    const labels = RULES.map((r) => S.ruleLabel({ ...r, label: r.label }, params));
    assert.deepEqual(labels, [
      'No email or invalid email format',
      'Email domain on the exclusion list',
      'Not subscribed to marketing emails',
      'Account younger than 14 days',
      'Carries an excluded tag',
      'Already sent (has the sent tag)',
      'Ordered in the last 6 months',
      'Same-address account ordered in the last 6 months',
      'Latest order from a marketplace channel',
      'Same address, another account kept',
      'Computed amount is $0',
      'Customer deleted at the order follow-up',
    ]);
    assert.equal(S.ruleLabel({ code: 'too-new', label: '注册不满 7 天' }, {}), 'Account younger than 7 days');
    assert.equal(S.ruleLabel({ code: 'recent-order', label: '近 3 个月有下单' }, {}), 'Ordered in the last 3 months');
    assert.equal(S.ruleLabel({ code: 'active-address' }, {}), 'Same-address account ordered in the last 3 months');
    assert.equal(S.ruleLabel({ code: 'future-rule', label: '新规则' }, {}), '新规则', 'an unknown rule keeps its stored label');
    assert.equal(ruleLabelEn('future-rule'), 'future-rule');
  });

  test('primaryReason / allReasons: from the new reasons field (codes, raw details, store dates)', () => {
    const info = { params: { minAccountAgeDays: 7, inactiveMonths: 3 }, timeZone: 'America/Los_Angeles' };
    const entry = {
      primaryText: '旧文本',
      allReasons: '旧文本；旧文本',
      reasons: [
        { n: 4, code: 'too-new', detail: '2026-09-29T03:00:00Z' }, // 2026-09-28 in Los Angeles
        { n: 9, code: 'marketplace-order', detail: '205641' },
        { n: 8, code: 'active-address', detail: cust(12) },
        { n: 10, code: 'duplicate-address', detail: cust(31) },
        { n: 11, code: 'zero-amount', detail: '没有可用来算平均值的有下单客户' },
        { n: 12, code: 'customer-deleted' },
        { n: 3, code: 'not-subscribed', detail: 'UNSUBSCRIBED' },
        { n: 5, code: 'excluded-tag', detail: 'WHS' },
      ],
    };
    assert.equal(S.primaryReason(entry, info), '4. Account younger than 7 days: 2026-09-28');
    assert.equal(S.allReasons(entry, info), [
      '4. Account younger than 7 days: 2026-09-28',
      '9. Latest order from a marketplace channel: Sellbrite',
      '8. Same-address account ordered in the last 3 months: ordering account 12',
      '10. Same address, another account kept: kept account 31',
      '11. Computed amount is $0: no ordering customers to compute the average from',
      '12. Customer deleted at the order follow-up',
      '3. Not subscribed to marketing emails: UNSUBSCRIBED',
      '5. Carries an excluded tag: WHS',
    ].join('; '));
    assert.equal(reasonTextEn({ n: 7, code: 'recent-order', detail: '2026-08-15' }, {}, 'America/Los_Angeles'), '7. Ordered in the last 3 months: 2026-08-15');
    assert.equal(reasonTextEn({ n: 99, code: 'future-rule', detail: 'x' }, {}), '99. future-rule: x');
  });

  test('primaryReason / allReasons: older lists without the field are parsed from the stored Chinese texts', () => {
    const info = { params: { minAccountAgeDays: 7, inactiveMonths: 3 } };
    const cases = [
      ['4. 注册不满 7 天：2026-09-28', '4. Account younger than 7 days: 2026-09-28'],
      ['5. 带排除 tag：WHS', '5. Carries an excluded tag: WHS'],
      ['6. 已经发过（已有发放 tag）：gift-card-sent-2026-10', '6. Already sent (has the sent tag): gift-card-sent-2026-10'],
      ['7. 近 3 个月有下单：2026-08-15', '7. Ordered in the last 3 months: 2026-08-15'],
      ['8. 同地址账户近 3 个月有下单：下单账户 12', '8. Same-address account ordered in the last 3 months: ordering account 12'],
      ['9. 最近订单来自平台渠道：Amazon', '9. Latest order from a marketplace channel: Amazon'],
      ['9. 最近订单来自平台渠道：网店', '9. Latest order from a marketplace channel: Online Store (web)'],
      ['9. 最近订单来自平台渠道：草稿订单', '9. Latest order from a marketplace channel: Draft order'],
      ['10. 同地址重复，保留了另一个账户：保留 31', '10. Same address, another account kept: kept account 31'],
      ['11. 算出的金额为 $0：没有可用来算平均值的有下单客户', '11. Computed amount is $0: no ordering customers to compute the average from'],
      ['11. 算出的金额为 $0', '11. Computed amount is $0'],
      ['12. 补查订单时客户已被删除', '12. Customer deleted at the order follow-up'],
      ['3. 未订阅邮件营销：NOT_SUBSCRIBED', '3. Not subscribed to marketing emails: NOT_SUBSCRIBED'],
      ['2. 邮箱域名在排除名单里：mail.codisto.com', '2. Email domain on the exclusion list: mail.codisto.com'],
      ['1. 没有邮箱或邮箱格式无效', '1. No email or invalid email format'],
    ];
    for (const [stored, expected] of cases) assert.equal(S.primaryReason({ primaryText: stored }, info), expected, stored);
    assert.equal(S.allReasons({ allReasons: '5. 带排除 tag：DISC；7. 近 3 个月有下单：2026-08-01' }, info), '5. Carries an excluded tag: DISC; 7. Ordered in the last 3 months: 2026-08-01');
    // numbers come from the stored label when params lack them
    assert.equal(S.primaryReason({ primaryText: '4. 注册不满 30 天：2026-09-28' }, { params: {} }), '4. Account younger than 30 days: 2026-09-28');
    assert.equal(S.primaryReason({ primaryText: '7. 近 6 个月有下单：2026-08-15' }, { params: {} }), '7. Ordered in the last 6 months: 2026-08-15');
    // unknown forms stay as they are, never throw
    assert.equal(S.primaryReason({ primaryText: '99. 新规则：x' }, info), '99. 新规则：x');
    assert.equal(S.primaryReason({ primaryText: 'free text' }, info), 'free text');
    assert.equal(S.primaryReason({}, info), '');
    assert.equal(S.allReasons({ allReasons: null }, info), '');
    assert.equal(notSelectedReasonsEn({ reasons: [] , primaryText: '12. 补查订单时客户已被删除' }, 'primary', info), '12. Customer deleted at the order follow-up', 'an empty reasons array falls back to the text');
  });

  test('the new reasons field and the stored texts of a real selection render the same English', async () => {
    const env = testConfig();
    try {
      const customers = [
        makeCustomer({ n: 1, lastOrder: makeOrder({ n: 11, createdAt: '2026-03-01T12:00:00Z', total: '150.00' }) }),
        makeCustomer({ n: 2, tags: ['WHS'], createdAt: '2026-09-30T12:00:00Z' }), // rules 4 + 5
        makeCustomer({ n: 3, lastOrder: makeOrder({ n: 31, createdAt: '2026-02-01T12:00:00Z', total: '90.00', source: 'amazon' }) }), // rule 9
        makeCustomer({ n: 4, lastOrder: makeOrder({ n: 41, createdAt: '2026-08-15T12:00:00Z', total: '40.00' }) }), // rule 7
      ];
      const selection = await selectionFixture(env.config, { customers, write: false });
      const info = { params: selection.params, timeZone: selection.params.timezone };
      assert.ok(selection.notSelected.length >= 3);
      for (const entry of selection.notSelected) {
        assert.ok(Array.isArray(entry.reasons) && entry.reasons.length, 'select writes reasons');
        const fromField = S.allReasons(entry, info);
        const fromText = S.allReasons({ allReasons: entry.allReasons }, info);
        assert.equal(fromField, fromText, entry.allReasons);
        assert.equal(S.primaryReason(entry, info), S.primaryReason({ primaryText: entry.primaryText }, info));
        assert.doesNotMatch(fromField, CJK);
      }
      const c2 = selection.notSelected.find((x) => x.customerId === cust(2));
      assert.equal(S.allReasons(c2, info), '4. Account younger than 7 days: 2026-09-30; 5. Carries an excluded tag: WHS');
    } finally {
      env.cleanup();
    }
  });

  test('formula: rebuilt from the stored numbers; translated word by word without them', () => {
    const selection = { params: { giftPercent: 10 }, averageCents: 10558 };
    assert.equal(S.formula({ kind: 'ordered', basis: { totalCents: 15000 }, rawCents: 1500, amountCents: 1533, formula: '10% × $150.00 = $15.00 → 档位 $15.33' }, selection), '10% × $150.00 = $15.00 → tier $15.33');
    assert.equal(S.formula({ kind: 'never', rawCents: 1056, amountCents: 1077, formula: '10% × 平均 $105.58 = $10.56 → 档位 $10.77' }, selection), '10% × average $105.58 = $10.56 → tier $10.77');
    assert.equal(S.formula({ kind: 'test', amountCents: 10, formula: '测试固定金额 $0.10' }, {}), 'Fixed test amount $0.10');
    assert.equal(S.formula({ kind: 'ordered', basis: { totalCents: 10775 }, rawCents: 1078, amountCents: 1533 }, { params: { giftPercent: 7.5 } }), '7.5% × $107.75 = $10.78 → tier $15.33');
    // without the numbers: the stored text, translated
    assert.equal(S.formula({ formula: '10% × 平均 $125.92 = $12.59 → 档位 $15.33' }, {}), '10% × average $125.92 = $12.59 → tier $15.33');
    assert.equal(S.formula({ formula: '测试固定金额 $0.10' }, {}), 'Fixed test amount $0.10');
    assert.equal(S.formula({}, {}), '');
    assert.equal(formulaEn(null, null), '');
    // the Chinese pack shows the stored text as it is
    assert.equal(zh.stored.formula({ formula: 'x' }, {}), 'x');
  });

  test('channel: codes and stored Chinese labels', () => {
    for (const [input, expected] of [['web', 'Online Store (web)'], ['WEB', 'Online Store (web)'], ['网店', 'Online Store (web)'], ['checkout_next', 'New checkout'], ['新版结账', 'New checkout'], ['shopify_draft_order', 'Draft order'], ['草稿订单', 'Draft order'], ['205641', 'Sellbrite'], ['amazon', 'Amazon'], ['pos', 'POS'], ['tiktok', 'tiktok'], ['Amazon', 'Amazon']]) {
      assert.equal(S.channel(input), expected, input);
      assert.equal(en.channelLabel(input), expected, input);
    }
    assert.equal(en.channelLabel(''), '');
    assert.equal(en.channelLabel(null), '');
    assert.equal(channelLabelEn(undefined), '');
    assert.equal(S.channel(''), '', 'an empty key stays empty (the sheet shows its own "(empty)")');
  });

  test('text: every form our own code writes into the journal / usage.json', () => {
    const cases = [
      // issue.js
      ['邮箱格式无效', 'invalid email format'],
      ['超过 10 分钟仍查不到这张卡，确认未建成，可以重试', 'Card still not found after 10 minutes: confirmed not created, can be retried'],
      ['超过 1 分钟仍查不到这张卡，确认未建成，可以重试', 'Card still not found after 1 minute: confirmed not created, can be retried'],
      ['超过 10 分钟仍查不到这张卡，确认未建成；正式活动从 REMIND_1_DATE（2026-10-12）起不再建新卡', 'Card still not found after 10 minutes: confirmed not created; from REMIND_1_DATE (2026-10-12) on the live campaign creates no new cards'],
      ['超过 90 秒仍查不到这张卡，确认未建成；本次是 --repair-only，不建新卡；之后正常运行 issue 时才会给他建卡', 'Card still not found after 90 seconds: confirmed not created; this run is --repair-only and creates no new cards; a later normal issue run will create it'],
      ['上次运行在建卡途中中断，Shopify 上暂时查不到这张卡', 'The previous run was interrupted while creating this card; Shopify does not show it yet'],
      ['未知错误', 'unknown error'],
      ['查卡也失败了', 'the card lookup failed too'],
      ['正式活动从 REMIND_1_DATE（2026-10-12）起不再建新卡', 'from REMIND_1_DATE (2026-10-12) on the live campaign creates no new cards'],
      ['本次是 --repair-only，不建新卡；之后正常运行 issue 时才会给他建卡', 'this run is --repair-only and creates no new cards; a later normal issue run will create it'],
      // verify.js
      ['verify：开始建卡 10 分钟后仍没在 Shopify 找到这张卡，确认没有建成', 'verify: card not found in Shopify 10 minutes after creation started; confirmed not created'],
      ['卡 9 张，发现 11 条，补记日志 2 条', '9 cards, 11 issues found, 2 journal entries added'],
      ['卡 9 张，没有发现问题', '9 cards, no issues found'],
      ['失败：Network error calling Shopify: fetch failed', 'Failed: Network error calling Shopify: fetch failed'],
      // remind.js
      ['上次运行在发送这封提醒时中断，不知道是否已发出', 'The previous run was interrupted while sending this reminder; unknown whether it went out'],
      ['余额 $5.33 / 面额 $15.33', 'balance $5.33 / face value $15.33'],
      ['到期日 2026-10-11', 'expiry date 2026-10-11'],
      ['2 张卡：x042、1043', '2 cards: x042, 1043'],
      ['3 张卡：a、b、c', '3 cards: a, b, c'],
      ['Shopify 上这位客户名下没有日志记录的卡 1950', 'Shopify has no card 1950 (the one in the journal) under this customer'],
      ['customer deleted', 'Customer deleted'],
      // usage.js
      ['回执里没有礼品卡 ID', 'No gift card ID in the receipt'],
      ['回执里没有礼品卡 ID（退款）', 'No gift card ID in the receipt (refund)'],
      ['（无名称）', '(no name)'],
      // durations
      ['10 分钟', '10 minutes'],
      ['90 秒', '90 seconds'],
    ];
    for (const [stored, expected] of cases) assert.equal(S.text(stored), expected, stored);
    // unknown texts come back unchanged (customer data, Shopify's words, newer commands), never throw
    for (const s of ['这是新的说明', 'fetch failed', '#7001', 'mail.codisto.com', 'already tagged in Shopify', ' spaced ', '']) assert.equal(S.text(s), s);
    assert.equal(S.text(null), '');
    assert.equal(storedTextEn(undefined), '');
    assert.equal(S.productName('（无名称）'), '(no name)');
    assert.equal(S.productName('Gold Balloon'), 'Gold Balloon');
    assert.equal(zh.stored.text('邮箱格式无效'), '邮箱格式无效');
  });

  test('verify: the *En field when present, else the Chinese text unchanged; summaryText: null in English', () => {
    const issue = { journal: '日志', shopify: 'Shopify 文本', action: '操作', journalEn: 'Journal', actionEn: '' };
    assert.equal(S.verify(issue, 'journal'), 'Journal');
    assert.equal(S.verify(issue, 'shopify'), 'Shopify 文本');
    assert.equal(S.verify(issue, 'action'), '操作', 'an empty *En field does not hide the Chinese text');
    assert.equal(S.verify({}, 'journal'), undefined);
    assert.equal(zh.stored.verify(issue, 'journal'), '日志');
    assert.equal(S.summaryText('卡 9 张，没有发现问题'), null);
    assert.equal(zh.stored.summaryText('x'), 'x');
  });
});

// ---------------------------------------------------------------------------
// labels.js formatters in English
// ---------------------------------------------------------------------------

describe('labels.js in English', () => {
  test('errorText: the Shopify error forms and the commands\' own texts', () => {
    const cases = [
      ['giftCardCreate rejected: input: Customer is invalid [INVALID]; input.expiresOn: must be in the future [GREATER_THAN]', 'Shopify refused: input: Customer is invalid (INVALID); input.expiresOn: must be in the future (GREATER_THAN)'],
      ['tagsAdd rejected: tags: Tag limit reached [INVALID]', 'Shopify refused: tags: Tag limit reached (INVALID)'],
      ['Network error calling Shopify: fetch failed', 'Network error: fetch failed'],
      ['Network error calling Shopify', 'Network error'],
      ['Shopify HTTP 502', 'Shopify returned HTTP 502'],
      ['Shopify HTTP 503: Service Unavailable', 'Shopify returned HTTP 503: Service Unavailable'],
      ['Throttled by Shopify 6 times in a row; giving up', 'Shopify rate limit: refused 6 times in a row, gave up'],
      ['HTTP 429 from Shopify 6 times in a row; giving up', 'Shopify rate limit: refused 6 times in a row, gave up'],
      ['GraphQL error: Throttled', 'Shopify query error: Throttled'],
      ['Shopify returned non-JSON (text/html)', 'Shopify returned an unreadable response'],
      ['Shopify response contained no data', 'Shopify returned no data'],
      ['tagsAdd returned no payload', 'Shopify returned no result'],
      ['Network error calling Shopify: fetch failed；查卡也失败了：GraphQL error: Internal error', 'Network error: fetch failed; the card lookup failed too: Shopify query error: Internal error'],
      ['未知错误；查卡也失败了：Shopify HTTP 500', 'unknown error; the card lookup failed too: Shopify returned HTTP 500'],
      ['上次运行在建卡途中中断，Shopify 上暂时查不到这张卡', 'The previous run was interrupted while creating this card; Shopify does not show it yet'],
      ['rejected', 'rejected'],
      ['outcome unknown', 'outcome unknown'],
      ['tag failed', 'tag failed'],
      ['some other error', 'some other error'],
    ];
    for (const [message, expected] of cases) assert.equal(errorText(message, 'en'), expected, message);
    assert.equal(errorText(null, 'en'), '');
  });

  test('summaryText: English labels, dry-run words, nested skips, seq ranges, stops and errors; verify\'s Chinese text ignored', () => {
    assert.equal(
      summaryText({ batch: 1, dryRun: false, attempted: 9, created: 3, tagged: 4, reconciled: 2, skipped: { 'ordered-since-snapshot': 3, 'no-email': 1 }, failed: 1, rejected: 1, unknown: 3, amountCents: 6000, tagFailed: 1, seqFrom: 1, seqTo: 13 }, { command: 'issue', dryRun: false }, 'en'),
      'attempted 9, cards created 3, tagged 4, recorded from Shopify 2, skipped (Ordered after the export 3, Email missing 1), failed 1, refused (can retry) 1, outcome unknown 3, amount $60.00, tag failed 1, seq 1-13',
    );
    assert.equal(summaryText({ dryRun: true, attempted: 1, tagFixed: 1, seqFrom: 9, seqTo: 9 }, { command: 'issue', dryRun: true }, 'en'), 'would create 1, tags added later 1, seq 9');
    assert.equal(summaryText({ repairOnly: true, reconciledNone: 3, newCardsRefused: 2, stoppedByDate: '2026-10-12' }, { command: 'issue' }, 'en'), 'records and tags only, confirmed not created 3, no new card (still pending) 2, stopped by date 2026-10-12');
    assert.equal(summaryText({ round: 1, dryRun: true, eligible: 2, planned: 2, alreadySent: 3, roundTagMissing: 1 }, { command: 'remind', dryRun: true }, 'en'), 'eligible 2, would send 2, sent earlier 3, round tag missing 1');
    assert.equal(summaryText({ round: 1, planned: 6, sent: 2, skipped: { used: 1 }, stopped: 'too-many-failures' }, { command: 'remind', dryRun: false }, 'en'), 'to send 6, sent 2, skipped (Card already used 1), Stopped: too many consecutive send failures');
    assert.equal(summaryText({ error: 'GraphQL error: Throttled' }, { command: 'select' }, 'en'), 'error: Shopify query error: Throttled');
    assert.equal(summaryText({ recipients: 13, totalCents: 18993, source: 'bulk', snapshotReused: true }, { command: 'select' }, 'en'), 'selected 13, total $189.93, export method Bulk export of the whole store, reused customer data exported within 24 hours');
    // verify: the Chinese line is ignored, the counters are shown
    const verifyLine = summaryText({ cardCount: 9, taggedCount: 7, issueCount: 11, fixedCount: 2, counts: { 'tag-missing': 1 }, text: '卡 9 张，发现 11 条，补记日志 2 条' }, { command: 'verify' }, 'en');
    assert.match(verifyLine, /^cards 9, tagged customers 7, issues 11, journal entries added 2, by type \(Card without the tag 1\)/);
    assert.doesNotMatch(verifyLine, CJK);
    const failedLine = summaryText({ error: 'Network error calling Shopify: fetch failed', text: '失败：Network error calling Shopify: fetch failed' }, { command: 'verify' }, 'en');
    assert.match(failedLine, /^error: Network error: fetch failed/);
    assert.doesNotMatch(failedLine, CJK);
    assert.equal(summaryText({ a: 0, b: false }, {}, 'en'), 'all counters are 0');
    assert.equal(summaryText({ cards: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12], flags: [true] }, {}, 'en'), 'cards 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 and more (12 in all), flags yes');
    // the Chinese pack is untouched
    assert.equal(summaryText({ cardCount: 9, text: '卡 9 张，没有发现问题' }, { command: 'verify' }), '卡 9 张，没有发现问题');
  });

  test('skip texts, details, notes and labels', () => {
    const stamp = () => '2026-10-04 11:30';
    assert.equal(issueSkipText('ordered-since-snapshot', '2026-10-04T18:30:00Z', { stamp }, 'en'), 'Ordered after the export (2026-10-04 11:30 store time)');
    assert.equal(issueSkipText('no-email', '邮箱格式无效', {}, 'en'), 'Email missing (invalid email format)');
    assert.equal(issueSkipText('address-ordered-since-snapshot', 'gid://shopify/Customer/12', {}, 'en'), 'Same-address account ordered after the export (customer 12)');
    assert.equal(issueSkipText('not-subscribed', 'NONE', {}, 'en'), 'Not subscribed to marketing emails (No status)');
    assert.equal(issueSkipText('future-reason', 'gid://shopify/Product/5', {}, 'en'), 'future-reason (5)');
    assert.equal(remindSkipText('used', '余额 $5.33 / 面额 $15.33', {}, 'en'), 'Card already used (balance $5.33 / face value $15.33)');
    assert.equal(remindSkipText('multiple-cards', '2 张卡：x042、1043', {}, 'en'), 'More than one card (2 cards: x042, 1043)');
    assert.equal(remindSkipText('card-expired', '到期日 2026-10-11', {}, 'en'), 'Card expired (expiry date 2026-10-11)');
    assert.equal(remindSkipText('no-card', 'Shopify 上这位客户名下没有日志记录的卡 1950', {}, 'en'), 'No card (Shopify has no card 1950 (the one in the journal) under this customer)');
    assert.equal(remindSkipText('no-card', undefined, {}, 'en'), 'No card');
    assert.equal(detailText('gid://shopify/Order/7', {}, 'en'), 'order 7');
    assert.equal(detailText('gid://shopify/GiftCard/9', {}, 'en'), 'gift card 9');
    assert.equal(noteText('already tagged in Shopify', 'en'), 'Already carries this tag in Shopify');
    assert.equal(noteText('customer deleted', 'en'), 'Customer deleted');
    assert.equal(noteText('something new', 'en'), 'something new');
    assert.equal(storedLabelOf(en.UNMATCHED_REASON_LABELS, '回执里没有礼品卡 ID（退款）', 'en'), 'No gift card ID in the receipt (refund)');
    assert.equal(storedLabelOf(en.UNMATCHED_REASON_LABELS, 'no-receipt', 'en'), 'No receipt');
    assert.equal(statusLabel('in_progress', null, 'en'), 'Needs review');
    assert.equal(statusLabel('in_progress', 'issue', 'en'), 'In progress');
    assert.equal(statusLabel('created', null, 'en'), 'Card created, not tagged');
  });

  test('remindLabel: every reminder cell form', () => {
    const tz = 'America/Los_Angeles';
    assert.equal(remindLabel({ status: 'sent', at: '2026-10-12T16:05:00Z' }, null, tz, 'en'), 'Sent 10-12 09:05');
    assert.equal(remindLabel({ status: 'sent', at: '2026-10-12T16:05:00Z', tagError: 'x' }, null, tz, 'en'), 'Sent 10-12 09:05; round tag missing, the next run adds it');
    assert.equal(remindLabel({ status: 'sent', source: 'tag' }, null, tz, 'en'), 'Sent (recorded from the round tag in Shopify)');
    assert.equal(remindLabel({ status: 'sent' }, null, tz, 'en'), 'Sent');
    assert.equal(remindLabel({ status: 'skipped', reason: 'used' }, null, tz, 'en'), 'Skipped: Card already used');
    assert.equal(remindLabel({ status: 'skipped' }, null, tz, 'en'), 'Skipped');
    assert.equal(remindLabel({ status: 'failed', error: 'giftCardSendNotificationToCustomer rejected: input: Customer has no email [INVALID]' }, null, tz, 'en'), 'Failed: Shopify refused: input: Customer has no email (INVALID)');
    assert.equal(remindLabel({ status: 'failed' }, null, tz, 'en'), 'Failed');
    assert.equal(remindLabel({ status: 'unknown' }, null, tz, 'en'), 'Outcome unknown');
    assert.equal(remindLabel({ status: 'in_progress' }, 'remind', tz, 'en'), 'In progress');
    assert.equal(remindLabel({ status: 'in_progress' }, null, tz, 'en'), 'Outcome unknown');
  });
});

// ---------------------------------------------------------------------------
// The workbooks
// ---------------------------------------------------------------------------

describe('English workbook', () => {
  for (const scenario of scenarios) {
    test(`${scenario.name}: no CJK or full-width character in any cell or sheet name, and the same structure as the Chinese edition`, async () => {
      const w = await writeScenario(scenario);
      assert.deepEqual(cjkCells(w.en), []);
      assert.deepEqual(w.en.map((s) => s.name), w.zh.map((s) => s.name).map((n) => en.sheets[Object.keys(zh.sheets).find((k) => zh.sheets[k] === n)]), 'the same sheets in the same order, in English');
      assert.deepEqual(structureOf(w.en), structureOf(w.zh), 'rows, columns, numbers, dates, links, styles and number-format shapes are identical');
      // the Chinese edition still shows Chinese (both files were written, each in its language)
      assert.ok(cjkCells(w.zh).length > 100);
    });
  }

  test('live: Recipients headers, statuses, reminder cells, formulas and notes', async () => {
    const { en: sheets } = await writeScenario(live);
    const R = sheetNamed(sheets, 'Recipients');
    assert.deepEqual(rowValues(R, 1), [
      'Seq', 'Status', 'Batch', 'Customer ID', 'Name', 'Email', 'Marketing', 'Customer type', 'Last valid order', 'Last order date',
      'Latest order channel', 'Days ago', 'Last order total', 'Amount formula', 'Tier', 'Gift card amount', 'Accounts at address', 'Other accounts at address',
      'Has gift-card-sent-2026-10', 'Gift card ID', 'Card last 4', 'Card created', 'Tagged', 'Reminder 1', 'Reminder 2',
      'Used amount', 'Balance', 'Orders using the card', 'Notes/errors', 'City', 'State', 'ZIP', 'Orders', 'Total spent', 'Signed up',
    ]);
    const COL = Object.fromEntries(rowValues(R, 1).map((h, i) => [h, i]));
    const bySeq = new Map(allRows(R).slice(1).map((r) => [r[COL.Seq], r]));
    const col = (seq, name) => bySeq.get(seq)[COL[name]];
    assert.deepEqual([...bySeq.values()].map((r) => r[COL.Status]), [
      'Done', 'Done', 'Done', 'Done', 'Card created, not tagged', 'Failed', 'Done', 'Skipped before issuing', 'Needs review', 'Pending', 'Needs review', 'Done', 'Skipped before issuing',
    ]);
    assert.equal(col(2, 'Reminder 1'), 'Sent 10-12 09:05');
    assert.equal(col(2, 'Reminder 2'), 'In progress');
    assert.equal(col(12, 'Reminder 1'), 'Sent (recorded from the round tag in Shopify)');
    assert.equal(col(12, 'Reminder 2'), 'Sent 10-16 09:00; round tag missing, the next run adds it');
    assert.equal(col(5, 'Reminder 1'), 'Skipped: Card already used');
    assert.equal(col(5, 'Reminder 2'), 'Failed');
    assert.equal(col(7, 'Reminder 1'), 'Failed: Shopify refused: input: Customer has no email (INVALID)');
    assert.equal(col(7, 'Reminder 2'), 'Skipped: Not subscribed to marketing emails');
    assert.equal(col(1, 'Reminder 1'), 'Outcome unknown');
    assert.equal(col(6, 'Reminder 2'), 'Skipped');
    assert.equal(col(2, 'Amount formula'), '10% × $150.00 = $15.00 → tier $15.33');
    assert.equal(col(11, 'Amount formula'), '10% × average $125.92 = $12.59 → tier $15.33');
    assert.equal(col(2, 'Latest order channel'), 'Online Store (web)');
    assert.equal(col(5, 'Latest order channel'), 'New checkout');
    assert.equal(col(7, 'Latest order channel'), 'Draft order');
    assert.equal(col(2, 'Customer type'), 'Ordered before');
    assert.equal(col(11, 'Customer type'), 'Never ordered');
    assert.equal(col(2, 'Marketing'), 'Subscribed');
    assert.equal(col(2, 'Has gift-card-sent-2026-10'), 'Yes');
    assert.equal(col(1, 'Has gift-card-sent-2026-10'), 'No');
    assert.equal(col(1, 'Notes/errors'), 'Card found in Shopify (check at issue start) and recorded');
    assert.equal(col(5, 'Notes/errors'), 'Shopify refused: tags: Tag limit reached (INVALID)');
    assert.equal(col(6, 'Notes/errors'), 'Shopify refused: input: Customer is invalid (INVALID); input.expiresOn: must be in the future (GREATER_THAN)');
    assert.equal(col(7, 'Notes/errors'), 'Card amount $15.33 differs from the list amount $10.77');
    assert.equal(col(8, 'Notes/errors'), 'Ordered after the export; Latest order is $0, based on an earlier paid order');
    assert.equal(col(9, 'Notes/errors'), 'The card request was sent but no outcome was recorded (the run was interrupted), the next issue run first looks for the card in Shopify; Latest order cancelled, based on an earlier paid order');
    assert.equal(col(10, 'Notes/errors'), 'verify: card not found in Shopify 10 minutes after creation started; confirmed not created; Latest order is a test order, based on an earlier paid order');
    assert.equal(col(11, 'Notes/errors'), 'Network error: fetch failed; the card lookup failed too: Shopify query error: Internal error');
    assert.equal(col(12, 'Notes/errors'), 'Card found in Shopify (verify) and recorded; 10 accounts at this address, suspected bulk sign-ups');
    assert.equal(col(13, 'Notes/errors'), 'All orders cancelled or $0, treated as never ordered');
  });

  test('live: Not selected reasons parsed from the stored Chinese texts (old-style list)', async () => {
    const { en: sheets } = await writeScenario(live);
    const N = sheetNamed(sheets, 'Not selected');
    assert.deepEqual(rowValues(N, 1), ['Customer ID', 'Name', 'Email', 'Marketing', 'Main reason', 'All reasons', 'Last order date', 'Orders', 'Related account', 'Group', 'Tags']);
    const reasons = allRows(N).slice(1).map((r) => [r[0], r[4], r[5]]);
    const by = (id) => reasons.find(([c]) => c === id);
    assert.deepEqual(by('25'), ['25', '4. Account younger than 7 days: 2026-09-28', '4. Account younger than 7 days: 2026-09-28']);
    assert.deepEqual(by('29'), ['29', '5. Carries an excluded tag: DISC', '5. Carries an excluded tag: DISC; 7. Ordered in the last 3 months: 2026-08-01']);
    assert.deepEqual(by('26'), ['26', '6. Already sent (has the sent tag): gift-card-sent-2026-10', '6. Already sent (has the sent tag): gift-card-sent-2026-10']);
    assert.deepEqual(by('9'), ['9', '7. Ordered in the last 3 months: 2026-08-15', '7. Ordered in the last 3 months: 2026-08-15']);
    assert.deepEqual(by('11'), ['11', '8. Same-address account ordered in the last 3 months: ordering account 12', '8. Same-address account ordered in the last 3 months: ordering account 12']);
    assert.deepEqual(by('27'), ['27', '9. Latest order from a marketplace channel: Amazon', '9. Latest order from a marketplace channel: Amazon']);
    assert.deepEqual(by('28'), ['28', '9. Latest order from a marketplace channel: Sellbrite', '9. Latest order from a marketplace channel: Sellbrite']);
    assert.deepEqual(by('4'), ['4', '10. Same address, another account kept: kept account 3', '10. Same address, another account kept: kept account 3']);
    assert.deepEqual(by('30'), ['30', '12. Customer deleted at the order follow-up', '12. Customer deleted at the order follow-up']);
  });

  test('live: Not selected reasons from the new reasons field (new-style list) give the same cells, whatever the stored texts say', async () => {
    const detailOf = (n) => ({
      'too-new': '2026-09-28T12:00:00Z',
      'excluded-tag': n.tags,
      'already-sent': 'gift-card-sent-2026-10',
      'recent-order': n.lastOrderAt,
      'active-address': n.relatedCustomerId,
      'marketplace-order': n.lastOrderSource,
      'duplicate-address': n.relatedCustomerId,
    });
    const patch = (scenario) => {
      for (const n of scenario.selection.notSelected) {
        const details = detailOf(n);
        const rule = RULES.find((r) => r.code === n.primaryCode);
        const reasons = [{ n: rule.n, code: n.primaryCode, ...(details[n.primaryCode] ? { detail: details[n.primaryCode] } : {}) }];
        // c29 carries a second reason (rule 7) in its stored text
        if (n.allReasons.includes('；')) reasons.push({ n: 7, code: 'recent-order', detail: n.lastOrderAt });
        n.reasons = reasons;
        n.primaryText = '旧的中文文本';
        n.allReasons = '旧的中文文本；旧的中文文本';
      }
      return scenario;
    };
    const [a, b] = await Promise.all([writeScenario(live), writeScenario(live, patch)]);
    const cells = (sheets, name) => allRows(sheetNamed(sheets, name)).map((r) => [r[0], r[4], r[5]]);
    assert.deepEqual(cells(b.en, 'Not selected'), cells(a.en, 'Not selected'));
    assert.deepEqual(cjkCells(b.en), []);
    // the Chinese edition shows whatever the list stores (the old texts), reasons or not
    assert.ok(cells(b.zh, '未入选').slice(1).every(([, p, all]) => p === '旧的中文文本' && all === '旧的中文文本；旧的中文文本'));
  });

  test('live: Summary funnel, channels, amounts, usage and run history', async () => {
    const { en: sheets } = await writeScenario(live);
    const S = sheetNamed(sheets, 'Summary');
    const texts = textsOf(S);
    for (const expected of [
      'Campaign parameters', 'Selection funnel', 'Same-address dedupe', 'Amounts', 'Issuing progress', 'Reminders', 'Usage', 'Run history',
      '1. No email or invalid email format', '4. Account younger than 7 days', '7. Ordered in the last 3 months', '8. Same-address account ordered in the last 3 months', '12. Customer deleted at the order follow-up',
      'counted only', 'listed row by row', 'Rule 9: by channel', 'Online Store (web)', 'New checkout', 'Amazon', 'Sellbrite', 'Rule 3: by marketing status', 'Unsubscribed', 'No status', '(empty)',
      'Reminder 1', 'Reminder 2', 'gift-card-sent-2026-10-R1', 'Card already used', 'More than one card', 'reason not recorded',
      'Pending', 'In progress', 'Needs review', 'Card created, not tagged', 'Done', 'Failed', 'Skipped before issuing', 'Ordered after the export',
      'Bulk export of the whole store, 42 customers', 'Live campaign: selected by the rules', '≤ $10.77 → $10.77; $10.78–$15.33 → $15.33; ≥ $15.34 → $19.77',
      '10 people', '3 people', '1 person', 'Gold Balloon × 20; Helium Tank × 19; (no name) × 18', 'See the "Usage report" sheet', 'live', 'dry run', 'running', 'did not end normally (possibly interrupted)',
      'Interrupted (Ctrl+C)', 'all counters are 0',
    ]) assert.ok(texts.includes(expected), `Summary has "${expected}"`);
    assert.ok(texts.some((t) => t.startsWith('Campaign 2026-10 · live campaign · generated 2026-10-12 10:00 (Los Angeles time)')));
    assert.ok(texts.includes('attempted 9, cards created 3, tagged 4, recorded from Shopify 2, skipped (Ordered after the export 3, Email missing 1), failed 1, refused (can retry) 1, outcome unknown 3, amount $60.00, tag failed 1, seq 1-13'));
    assert.ok(texts.includes('records and tags only, tags added later 2, recorded from Shopify 2, confirmed not created 3, no new card (still pending) 2, stopped by date 2026-10-12'));
    assert.ok(texts.some((t) => t.startsWith('cards 9, tagged customers 7, issues 11, journal entries added 2, by type (Card without the tag 1, Still needs review 1)')));
    assert.ok(texts.some((t) => t.startsWith('error: Network error: fetch failed')));
    assert.ok(texts.includes('2026-10-12 when the list was generated, now 2026-10-13 per .env'));
  });

  test('live: Activity log shows every journal note, detail and error in English', async () => {
    const { en: sheets } = await writeScenario(live);
    const J = sheetNamed(sheets, 'Activity log');
    assert.deepEqual(rowValues(J, 1), ['Time', 'Command', 'Batch/round', 'Customer ID', 'Action', 'Result/notes', 'Gift card ID', 'Error']);
    const rows = allRows(J).slice(1);
    const of = (cid, action) => rows.filter((r) => r[3] === cid && (!action || r[4] === action));
    const results = rows.map((r) => r[5]).filter(Boolean);
    const errors = rows.map((r) => r[7]).filter(Boolean);
    const actions = new Set(rows.map((r) => r[4]));
    for (const a of Object.values(en.OP_LABELS)) assert.ok(actions.has(a), `action "${a}" appears`);
    assert.ok(actions.has('future.op'), 'an unknown op shows its code');
    for (const expected of [
      'Card still not found after 10 minutes: confirmed not created, can be retried',
      'Card still not found after 10 minutes: confirmed not created; from REMIND_1_DATE (2026-10-12) on the live campaign creates no new cards',
      'Card still not found after 90 seconds: confirmed not created; this run is --repair-only and creates no new cards; a later normal issue run will create it',
      'verify: card not found in Shopify 10 minutes after creation started; confirmed not created',
      'Email missing (invalid email format)',
      'Ordered after the export (2026-10-04 11:30 store time)',
      'Ordered after the export (#7001)',
      'Same-address account ordered after the export (customer 12)',
      'Same-address account ordered after the export (order 7)',
      'Not subscribed to marketing emails (Unsubscribed)',
      'Not subscribed to marketing emails (No status)',
      'future-reason (5)',
      'Card already used (balance $5.33 / face value $15.33)',
      'More than one card (2 cards: x042, 1043)',
      'No card (Shopify has no card 1950 (the one in the journal) under this customer)',
      'Card expired (expiry date 2026-10-11)',
      'No email (c953@example.org)',
      'Already carries this tag in Shopify',
      'Already carries gift-card-sent-2026-10-R1 in Shopify',
      'Already carries the round tag in Shopify',
      'gift-card-sent-2026-10-R1 not added; the next run adds it',
      'gift-card-sent-2026-10-R1 added',
      'Customer deleted',
      'Resent (had failed or outcome unknown)',
      'amount $15.33, card last 4 x001',
      'check right after the unknown outcome, amount $15.33, card last 4 x042',
      'pre-flight check, amount $15.33, card last 4 x041',
      'something new',
    ]) assert.ok(results.includes(expected), `result "${expected}"`);
    for (const expected of [
      'Network error: fetch failed; the card lookup failed too: Shopify query error: Internal error',
      'The previous run was interrupted while creating this card; Shopify does not show it yet',
      'The previous run was interrupted while sending this reminder; unknown whether it went out',
      'Shopify refused: input: Customer is invalid (INVALID); input.expiresOn: must be in the future (GREATER_THAN)',
      'Shopify rate limit: refused 6 times in a row, gave up',
      'Shopify returned HTTP 502',
      'Shopify returned HTTP 503: Service Unavailable',
      'Shopify returned no result',
      'Shopify returned an unreadable response',
      'Shopify returned no data',
    ]) assert.ok(errors.includes(expected), `error "${expected}"`);
    assert.equal(of('964', 'Round tag added')[0][5], 'Customer deleted');
    assert.deepEqual(rows.filter((r) => r[2]).map((r) => r[2]).filter((b) => /^Reminder/.test(b)).every((b) => b === 'Reminder 1' || b === 'Reminder 2'), true);
  });

  test('live: Usage report, Usage details and Verify', async () => {
    const { en: sheets } = await writeScenario(live);
    const U = sheetNamed(sheets, 'Usage report');
    assert.equal(rowValues(U, 1)[0], 'Usage report (as of 2026-10-12 08:00, Los Angeles time)');
    assert.deepEqual(rowValues(U, 9), ['By tier', 'Issued', 'Used', 'Usage rate', 'Used amount']);
    assert.ok(textsOf(U).includes('(no name)'));
    assert.deepEqual(textsOf(U).filter((t) => ['Ordered before', 'Never ordered', 'Test', 'mystery'].includes(t)), ['Ordered before', 'Never ordered', 'Test', 'mystery']);
    const D = sheetNamed(sheets, 'Usage details');
    const dTexts = textsOf(D);
    assert.ok(dTexts.includes('Gift card payments needing review: no campaign gift card ID in the receipt (3)'));
    assert.ok(dTexts.includes('No gift card ID in the receipt'));
    assert.ok(dTexts.includes('No gift card ID in the receipt (refund)'));
    assert.ok(dTexts.includes('No receipt'));
    assert.ok(dTexts.includes('(no name) × 1;  × 1'));
    assert.ok(dTexts.includes('Yes'), 'a cancelled order');
    const V = sheetNamed(sheets, 'Verify');
    assert.equal(rowValues(V, 1)[0], 'Verified 2026-10-14 10:00 (Los Angeles time); 9 campaign cards; 7 customers carrying gift-card-sent-2026-10');
    assert.match(rowValues(V, 2)[0], /^11 issues found: Card in the journal, not in Shopify 1; .*; future-type 1$/);
    assert.deepEqual(rowValues(V, 3), ['Issue type', 'Customer ID', 'Gift card ID', 'Journal', 'Shopify', 'Suggested action']);
    const issueRows = allRows(V).slice(3);
    assert.deepEqual(issueRows.map((r) => r[0]), [...Object.values(en.VERIFY_TYPE_LABELS), 'future-type']);
    assert.deepEqual(issueRows[0].slice(3), ['Done: card x014, tagged 2026-10-05 09:06', 'Not among this campaign\'s cards', 'Check the customer in the admin'], 'the *En fields are shown');
  });

  test('live: Help explains the sheets, columns, statuses, reminder cells and round tags in English', async () => {
    const { en: sheets } = await writeScenario(live);
    const H = sheetNamed(sheets, 'Help');
    const texts = textsOf(H);
    assert.deepEqual(rowValues(H, 1), ['Item', 'Explanation']);
    for (const name of Object.values(en.sheets)) assert.ok(texts.includes(name), `Help lists the "${name}" sheet`);
    for (const status of Object.values(en.STATUS_LABELS)) assert.ok(texts.includes(status));
    for (const [item] of en.REMIND_HELP) assert.ok(texts.includes(item));
    for (const expected of ['Recipients columns', 'Has gift-card-sent-2026-10', 'Notes/errors', 'Round tag (at most one reminder per person per round)', 'Adding a missing round tag', 'Amount rules', 'Tier 1', 'base ≤ $10.77 → $10.77', 'Exit code 130', 'Read-only', 'Filtering']) {
      assert.ok(texts.includes(expected), `Help has "${expected}"`);
    }
    assert.ok(texts.some((t) => t.includes('gift-card-sent-2026-10-R1') && t.includes('gift-card-sent-2026-10-R2')));
    assert.ok(texts.includes('All times are in the store\'s time zone (America/Los_Angeles, Los Angeles time).'));
  });

  test('test campaign and single-tier list: "Fixed test amount", "Test" tier, another time zone, "any" tier', async () => {
    const t = await writeScenario(testCampaign);
    const R = sheetNamed(t.en, 'Recipients');
    const first = rowValues(R, 2);
    assert.equal(first[13], 'Fixed test amount $0.10');
    assert.equal(first[14], 'Test');
    assert.equal(first[7], 'Test');
    const S = textsOf(sheetNamed(t.en, 'Summary'));
    assert.ok(S.includes('Test campaign: only the customers in TEST_CUSTOMER_IDS, no selection rules, fixed amount'));
    assert.ok(S.includes('Fixed test amount $0.10'));
    assert.ok(S.some((x) => x.includes('(Asia/Shanghai time)')));
    assert.ok(S.includes('not set'), 'an unset date');
    assert.ok(textsOf(sheetNamed(t.en, 'Help')).includes('Fixed $0.10 per person, no tiers.'));
    const e = await writeScenario(empty);
    assert.ok(textsOf(sheetNamed(e.en, 'Summary')).includes('any → $10.00'));
    assert.ok(textsOf(sheetNamed(e.en, 'Help')).includes('base any → $10.00'));
    assert.ok(textsOf(sheetNamed(e.en, 'Summary')).includes('no runs yet'));
    assert.ok(textsOf(sheetNamed(e.en, 'Verify')).includes('No issues found'));
    assert.ok(textsOf(sheetNamed(e.en, 'Usage details')).includes('No gift card payments need review'));
  });

  test('a stored Chinese text the tables do not know is shown unchanged in the English edition (no crash), and verify issues without *En fields keep their Chinese', async () => {
    const patch = (scenario) => {
      scenario.journal.push(
        { t: '2026-10-17T10:00:00.000Z', op: 'reconcile.none', cid: cust(999), note: '这是一条新的说明文字', run: 'gone' },
        { t: '2026-10-17T10:00:01.000Z', op: 'skip', cid: cust(998), reason: 'ordered-since-snapshot', detail: '新的细节', run: 'gone' },
        { t: '2026-10-17T10:00:02.000Z', op: 'create.fail', cid: cust(997), error: '新的错误；查卡也失败了：Shopify HTTP 500', run: 'gone' },
      );
      scenario.verify.issues.push({ type: 'tag-missing', customerId: cust(996), giftCardId: null, journal: '没有英文的日志文本', shopify: '客户没有 tag', action: '运行 issue' });
      scenario.usage.unmatched.push({ orderId: gid('Order', 5009), orderName: '#5009', orderCreatedAt: '2026-10-11T20:00:00.000Z', customerId: null, amountCents: 100, processedAt: '2026-10-11T20:00:05.000Z', reason: '新的原因' });
      return scenario;
    };
    const w = await writeScenario(live, patch);
    const J = sheetNamed(w.en, 'Activity log');
    const rows = allRows(J);
    assert.equal(rows.find((r) => r[3] === '999')[5], '这是一条新的说明文字');
    assert.equal(rows.find((r) => r[3] === '998')[5], 'Ordered after the export (新的细节)');
    assert.equal(rows.find((r) => r[3] === '997')[7], '新的错误; the card lookup failed too: Shopify returned HTTP 500');
    const V = allRows(sheetNamed(w.en, 'Verify'));
    assert.deepEqual(V.find((r) => r[1] === '996').slice(3), ['没有英文的日志文本', '客户没有 tag', '运行 issue']);
    assert.ok(textsOf(sheetNamed(w.en, 'Usage details')).includes('新的原因'));
    const unknown = cjkCells(w.en);
    assert.equal(unknown.length, 7, `only the unknown stored texts stay Chinese: ${unknown.join('\n')}`);
    assert.deepEqual(structureOf(w.en), structureOf(w.zh));
  });
});
