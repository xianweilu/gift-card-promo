// Builders for Shopify-shaped test data and an isolated config per test.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { loadConfig } from '../src/config.js';

export const NOW_ISO = '2026-10-01T20:00:00.000Z'; // cutoff in America/Los_Angeles = 2026-07-01T00:00:00-07:00
export const NOW_MS = Date.parse(NOW_ISO);
export const CUTOFF_MS = Date.parse('2026-07-01T00:00:00-07:00');

export function gid(type, n) {
  return `gid://shopify/${type}/${n}`;
}

/** An Order node as returned inside lastOrder / CustomerOrders. */
export function makeOrder({ n = 1, createdAt = '2026-03-01T12:00:00Z', total = '50.00', cancelled = false, test = false, source = 'web' } = {}) {
  return {
    id: gid('Order', n),
    name: `#${n}`,
    createdAt,
    cancelledAt: cancelled ? createdAt : null,
    test,
    sourceName: source,
    totalPriceSet: { shopMoney: { amount: total, currencyCode: 'USD' } },
  };
}

/** A Customer node with the fields of src/queries.js CUSTOMER. */
export function makeCustomer({
  n,
  first = `First${n}`,
  last = `Last${n}`,
  email = `c${n}@example.org`,
  validFormat = true,
  marketingState = 'SUBSCRIBED',
  createdAt = '2024-01-01T00:00:00Z',
  tags = [],
  address = { address1: `${n} Main St`, address2: '', city: 'Austin', provinceCode: 'TX', zip: `7${String(n).padStart(4, '0')}`, countryCodeV2: 'US' },
  lastOrder = null,
  numberOfOrders,
  amountSpent = '0.00',
} = {}) {
  return {
    id: gid('Customer', n),
    firstName: first,
    lastName: last,
    displayName: `${first} ${last}`.trim(),
    createdAt,
    numberOfOrders: String(numberOfOrders ?? (lastOrder ? 1 : 0)),
    tags: [...tags],
    amountSpent: { amount: amountSpent, currencyCode: 'USD' },
    defaultEmailAddress: email ? { emailAddress: email, marketingState, validFormat } : null,
    defaultAddress: address,
    lastOrder,
  };
}

/**
 * A GIFT_CARD_ORDERS node: an order paid (partly) with gift cards.
 * payments: [{ giftCardNumericId, amount: '10.77', kind: 'SALE'|'REFUND', status: 'SUCCESS', processedAt }]
 * lineItems: [{ name, sku, quantity, amount }]
 */
export function makeGiftCardOrder({ n = 1, customer = null, createdAt = '2026-10-06T18:00:00Z', total = '60.00', payments = [], lineItems = [], cancelled = false } = {}) {
  return {
    id: gid('Order', n),
    name: `#${n}`,
    createdAt,
    cancelledAt: cancelled ? createdAt : null,
    displayFinancialStatus: 'PAID',
    customer: customer ? { id: customer.id ?? customer } : null,
    totalPriceSet: { shopMoney: { amount: total, currencyCode: 'USD' } },
    transactions: [
      ...payments.map((p) => ({
        gateway: 'gift_card',
        kind: p.kind ?? 'SALE',
        status: p.status ?? 'SUCCESS',
        processedAt: p.processedAt ?? createdAt,
        amountSet: { shopMoney: { amount: p.amount } },
        receiptJson: p.noReceiptId ? JSON.stringify({}) : JSON.stringify({ gift_card_id: Number(p.giftCardNumericId), gift_card_last_characters: 'x000' }),
      })),
      { gateway: 'shopify_payments', kind: 'SALE', status: 'SUCCESS', processedAt: createdAt, amountSet: { shopMoney: { amount: '1.00' } }, receiptJson: '{}' },
    ],
    lineItems: { nodes: lineItems.map((l) => ({ name: l.name, sku: l.sku ?? null, quantity: l.quantity ?? 1, discountedTotalSet: { shopMoney: { amount: l.amount ?? '1.00' } } })) },
  };
}

/** An ORDERS_SINCE node (activity check). */
export function makeActiveOrder({ n = 1, customer, createdAt = '2026-08-15T12:00:00Z', cancelled = false, test = false } = {}) {
  return {
    id: gid('Order', n),
    createdAt,
    cancelledAt: cancelled ? createdAt : null,
    test,
    customer: customer ? { id: customer.id, defaultAddress: customer.defaultAddress } : null,
  };
}

/** A fresh campaigns dir and a validated config pointing at it. */
export function testConfig(overrides = {}) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'gcp-test-'));
  const env = {
    SHOP: 'teststore',
    CLIENT_ID: 'id',
    CLIENT_SECRET: 'secret',
    CAMPAIGN_ID: '2026-10',
    SENT_TAG: 'gift-card-sent-2026-10',
    GIFT_CARD_NOTE: 'gift-card-promo',
    GIFT_CARD_TEMPLATE_SUFFIX: 'gift-card-promo',
    TIMEZONE: 'America/Los_Angeles',
    LAUNCH_DATE: '2026-10-05',
    REMIND_1_DATE: '2026-10-12',
    REMIND_2_DATE: '2026-10-16',
    GIFT_CARD_EXPIRES_ON: '2026-10-19',
    DRY_RUN: 'false',
    CAMPAIGNS_DIR: dir,
    ...overrides,
  };
  const config = loadConfig(env, { envFile: false });
  return {
    config,
    dir,
    cleanup() {
      fs.rmSync(dir, { recursive: true, force: true });
    },
  };
}

/**
 * A selection object exactly as `select` would write it, built from fake
 * customer nodes with the real rules (src/select/selection.js). Also writes it
 * to selection.json when `write` is true. `history` maps customer gid → raw
 * Order nodes (newest first) for customers whose last order is cancelled/test/$0.
 */
export async function selectionFixture(config, { customers, activeOrders = [], history = {}, nowMs = NOW_MS, createdAt = NOW_ISO, write = true } = {}) {
  const { buildSelection } = await import('../src/select/selection.js');
  const { parseCustomer } = await import('../src/select/rules.js');
  const { campaignPaths, writeJsonAtomic } = await import('../src/campaign.js');
  const selection = buildSelection({
    config,
    customers: customers.map(parseCustomer),
    activeOrders,
    history,
    nowMs,
    cutoffMs: CUTOFF_MS,
    cutoffIso: '2026-07-01T00:00:00-07:00',
    timezone: 'America/Los_Angeles',
    createdAt,
    snapshot: { exportedAt: createdAt, source: 'bulk', count: customers.length },
  });
  if (write) writeJsonAtomic(campaignPaths(config).selection, selection);
  return selection;
}

/** A logger that records lines instead of printing them. */
export function memoryLog() {
  const lines = [];
  return {
    lines,
    info: (...a) => lines.push(`INFO ${a.join(' ')}`),
    warn: (...a) => lines.push(`WARN ${a.join(' ')}`),
    error: (...a) => lines.push(`ERROR ${a.join(' ')}`),
  };
}
