import { test } from 'node:test';
import assert from 'node:assert/strict';
import { toCents, percentOf, isGiftableTotal, normalizeTiers, tierIndex, giftFor, formatUsd, formulaText, tierLabel, describeTiers } from '../src/select/amount.js';

const TIERS = [1077, 1533, 1977];
const gift = (orderCents) => giftFor(orderCents, { percent: 10, tiersCents: TIERS });

test('amount: decimal strings → integer cents', () => {
  assert.equal(toCents('28.59'), 2859);
  assert.equal(toCents('0.1'), 10);
  assert.equal(toCents('1077.5'), 107750);
  assert.equal(toCents(0), 0);
  assert.throws(() => toCents('abc'), /Not a money amount/);
  assert.throws(() => toCents(undefined), /Not a money amount/);
});

test('amount: percentages round half up to the cent', () => {
  assert.equal(percentOf(10775, 10), 1078); // $107.75 → $10.775 → $10.78
  assert.equal(percentOf(10774, 10), 1077);
  assert.equal(percentOf(15340, 10), 1534);
  assert.equal(percentOf(5, 10), 1);
  assert.equal(percentOf(4, 10), 0);
});

test('amount: orders whose 10% is below one cent are not a basis ($0 and under $0.05)', () => {
  assert.equal(isGiftableTotal(0, 10), false);
  assert.equal(isGiftableTotal(4, 10), false);
  assert.equal(isGiftableTotal(5, 10), true);
  assert.equal(isGiftableTotal(2300, 10), true);
  assert.equal(isGiftableTotal(12.5, 10), false, 'cents must be an integer');
});

test('amount: the plan tier table, including the exact boundaries', () => {
  const cases = [
    [8340, 834, 1077], // $83.40
    [10774, 1077, 1077], // $107.74 → $10.77 → $10.77
    [10775, 1078, 1533], // $107.75 → $10.78 → $15.33
    [15000, 1500, 1533],
    [15330, 1533, 1533], // exactly the upper bound
    [15334, 1533, 1533], // $153.34 → $15.334 → $15.33
    [15335, 1534, 1977], // $153.35 → $15.335 → $15.34 → top tier
    [15340, 1534, 1977],
    [35000, 3500, 1977],
    [2300, 230, 1077], // earlier $23.00 order
    [10558, 1056, 1077], // average $105.58 for never-ordered customers
    [5, 1, 1077], // smallest giftable order
  ];
  for (const [order, raw, amount] of cases) {
    const g = gift(order);
    assert.equal(g.rawCents, raw, `raw for ${order}`);
    assert.equal(g.amountCents, amount, `amount for ${order}`);
    assert.equal(TIERS[g.tier], amount);
  }
  assert.equal(tierIndex(1077, TIERS), 0);
  assert.equal(tierIndex(1078, TIERS), 1);
  assert.equal(tierIndex(1534, TIERS), 2);
  assert.equal(tierIndex(999999, TIERS), 2);
  assert.equal(tierIndex(500, [2000]), 0, 'a single tier pays everyone the same');
});

test('amount: tiers are validated', () => {
  assert.deepEqual(normalizeTiers([1077, 1533, 1977]), [1077, 1533, 1977]);
  assert.throws(() => normalizeTiers([]), /at least one/);
  assert.throws(() => normalizeTiers([1077, 1077]), /strictly increasing/);
  assert.throws(() => normalizeTiers([0]), /positive/);
  assert.throws(() => normalizeTiers([10.5]), /positive/);
});

test('amount: display helpers', () => {
  assert.equal(formatUsd(123456), '$1,234.56');
  assert.equal(formatUsd(10), '$0.10');
  assert.equal(formatUsd(-250), '-$2.50');
  assert.equal(formatUsd(0), '$0.00');
  assert.equal(formatUsd(null), '');
  assert.equal(formatUsd(undefined), '');
  assert.equal(tierLabel(1, TIERS), '$15.33');
  assert.deepEqual(describeTiers(TIERS), ['≤ $10.77 → $10.77', '$10.78–$15.33 → $15.33', '≥ $15.34 → $19.77']);
  assert.deepEqual(describeTiers([2000]), ['任意 → $20.00']);
});

test('amount: the formula text shown in the Excel', () => {
  assert.equal(formulaText({ kind: 'ordered', baseCents: 15000, percent: 10, ...gift(15000) }), '10% × $150.00 = $15.00 → 档位 $15.33');
  assert.equal(formulaText({ kind: 'never', baseCents: 10557.6, percent: 10, ...gift(10557.6) }), '10% × 平均 $105.58 = $10.56 → 档位 $10.77');
  assert.equal(formulaText({ kind: 'test', amountCents: 10 }), '测试固定金额 $0.10');
  assert.equal(formulaText({ kind: 'ordered', baseCents: 15000, percent: 12.5, rawCents: 1875, amountCents: 1977 }), '12.5% × $150.00 = $18.75 → 档位 $19.77');
});
