// Money is handled in integer cents end to end; only display converts to dollars.
//
// Amount rule (2026-10 campaign): base = 10% of the basis (last paid order
// total, or the average for never-ordered customers), rounded to the cent;
// the base is then mapped to a tier:
//   base ≤ $10.77 → $10.77,  base ≤ $15.33 → $15.33,  otherwise → $19.77.
// Tiers come from GIFT_TIERS ("10.77,15.33,19.77"): every value but the last
// is both an upper bound and the amount paid; the last is paid above them all.

/** Shopify decimal string ("28.59") → integer cents (2859). */
export function toCents(amount) {
  const n = Number(amount);
  if (!Number.isFinite(n)) throw new Error(`Not a money amount: ${amount}`);
  return Math.round(n * 100);
}

/** percent of cents, rounded half-up to a whole cent. */
export function percentOf(cents, percent) {
  return Math.round((cents * percent) / 100);
}

/** An order can be the basis of a gift only if its percentage is at least 1 cent (skips $0 and tiny orders). */
export function isGiftableTotal(totalCents, percent) {
  return Number.isInteger(totalCents) && percentOf(totalCents, percent) >= 1;
}

/** Validates and sorts tier amounts (cents). At least one tier, all positive, strictly increasing. */
export function normalizeTiers(tiersCents) {
  const tiers = [...tiersCents];
  if (!tiers.length) throw new Error('at least one gift tier is required');
  for (let i = 0; i < tiers.length; i += 1) {
    if (!Number.isInteger(tiers[i]) || tiers[i] <= 0) throw new Error(`gift tier must be a positive amount, got ${tiers[i]}`);
    if (i && tiers[i] <= tiers[i - 1]) throw new Error('gift tiers must be strictly increasing');
  }
  return tiers;
}

/** Index of the tier a base amount falls into (0-based). */
export function tierIndex(baseCents, tiersCents) {
  for (let i = 0; i < tiersCents.length - 1; i += 1) {
    if (baseCents <= tiersCents[i]) return i;
  }
  return tiersCents.length - 1;
}

/**
 * { rawCents, amountCents, tier } for a basis amount (an order total, or the
 * average for never-ordered customers). rawCents = percent of the basis,
 * rounded to the cent; amountCents = the tier it maps to; tier = 0-based index.
 */
export function giftFor(basisCents, { percent, tiersCents }) {
  const rawCents = percentOf(basisCents, percent);
  const tier = tierIndex(rawCents, tiersCents);
  return { rawCents, amountCents: tiersCents[tier], tier };
}

export function formatUsd(cents) {
  if (cents === null || cents === undefined || Number.isNaN(cents)) return '';
  const sign = cents < 0 ? '-' : '';
  return `${sign}$${(Math.abs(cents) / 100).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function pct(percent) {
  return `${Number.isInteger(percent) ? percent : percent.toFixed(2).replace(/0+$/, '').replace(/\.$/, '')}%`;
}

/** The formula shown in the Excel, e.g. "10% × $150.00 = $15.00 → 档位 $15.33". */
export function formulaText({ kind, baseCents, percent, rawCents, amountCents }) {
  if (kind === 'test') return `测试固定金额 ${formatUsd(amountCents)}`;
  const base = kind === 'never' ? `平均 ${formatUsd(Math.round(baseCents))}` : formatUsd(baseCents);
  return `${pct(percent)} × ${base} = ${formatUsd(rawCents)} → 档位 ${formatUsd(amountCents)}`;
}

/** Display label of a tier, e.g. "$15.33". */
export function tierLabel(tier, tiersCents) {
  return formatUsd(tiersCents[tier]);
}

/** Human description of the tier ranges, e.g. ["≤ $10.77 → $10.77", "$10.78–$15.33 → $15.33", "≥ $15.34 → $19.77"]. */
export function describeTiers(tiersCents) {
  return tiersCents.map((t, i) => {
    if (tiersCents.length === 1) return `任意 → ${formatUsd(t)}`;
    if (i === 0) return `≤ ${formatUsd(t)} → ${formatUsd(t)}`;
    if (i === tiersCents.length - 1) return `≥ ${formatUsd(tiersCents[i - 1] + 1)} → ${formatUsd(t)}`;
    return `${formatUsd(tiersCents[i - 1] + 1)}–${formatUsd(t)} → ${formatUsd(t)}`;
  });
}
