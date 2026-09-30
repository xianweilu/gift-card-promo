import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

import { loadConfig, ConfigError } from '../src/config.js';
import { missingScopes, parseScopes, REQUIRED_SCOPES } from '../src/auth.js';
import { loadProgress, saveProgress } from '../src/progress.js';
import * as giftcards from '../src/giftcards.js';
import { buildCreateInput, formatAmount, GIFT_CARD_CREATE } from '../src/giftcards.js';
import { TAGS_ADD, TAGGED_CUSTOMERS_QUERY, tagSearchQuery } from '../src/customers.js';
import { SEGMENT_MEMBERS_QUERY, toCustomerGid } from '../src/segment.js';

const baseEnv = {
  SHOP: 'balloonstore1',
  CLIENT_ID: 'id',
  CLIENT_SECRET: 'secret',
  SEGMENT_ID: 'gid://shopify/Segment/488535490623',
  SENT_TAG: 'gift-card-sent-2026-09',
};
const cfg = (extra = {}) => loadConfig({ ...baseEnv, ...extra }, { envFile: false });

test('config: defaults', () => {
  const c = cfg();
  assert.equal(c.dryRun, true);
  assert.equal(c.apiVersion, '2026-07');
  assert.equal(c.segmentPageSize, 250);
  assert.equal(c.giftCardValue, 0.1);
  assert.equal(c.giftCardCurrency, 'USD');
  assert.equal(c.giftCardExpiresOn, '');
  assert.equal(c.giftCardTemplateSuffix, '');
  assert.ok(c.progressFile.endsWith(`${path.sep}progress.json`));
});

test('config: DRY_RUN is only off for the literal string false', () => {
  assert.equal(cfg({ DRY_RUN: 'false' }).dryRun, false);
  assert.equal(cfg({ DRY_RUN: 'FALSE' }).dryRun, false);
  for (const v of ['0', 'no', 'off', '', 'true', 'yes']) {
    assert.equal(cfg({ DRY_RUN: v }).dryRun, true, `DRY_RUN=${JSON.stringify(v)} must stay in dry-run`);
  }
});

test('config: lists every missing required value', () => {
  assert.throws(
    () => loadConfig({ SHOP: 'x' }, { envFile: false }),
    (err) => err instanceof ConfigError && /CLIENT_ID, CLIENT_SECRET, SEGMENT_ID, SENT_TAG/.test(err.message),
  );
});

test('config: validation', () => {
  assert.throws(() => cfg({ GIFT_CARD_VALUE: '0' }), /GIFT_CARD_VALUE/);
  assert.throws(() => cfg({ GIFT_CARD_VALUE: '2000.01' }), /GIFT_CARD_VALUE/);
  assert.throws(() => cfg({ GIFT_CARD_VALUE: 'abc' }), /GIFT_CARD_VALUE/);
  assert.throws(() => cfg({ GIFT_CARD_EXPIRES_ON: '2027/01/01' }), /GIFT_CARD_EXPIRES_ON/);
  assert.throws(() => cfg({ SEGMENT_ID: '488535490623' }), /SEGMENT_ID/);
  assert.throws(() => cfg({ SEGMENT_PAGE_SIZE: '1001' }), /SEGMENT_PAGE_SIZE/);
  assert.throws(() => cfg({ SEGMENT_PAGE_SIZE: '2.5' }), /SEGMENT_PAGE_SIZE/);
  assert.throws(() => cfg({ SENT_TAG: 'a,b' }), /SENT_TAG/);
  assert.throws(() => cfg({ GIFT_CARD_CURRENCY: 'US' }), /GIFT_CARD_CURRENCY/);
  assert.throws(() => cfg({ API_VERSION: '2026' }), /API_VERSION/);
  assert.equal(cfg({ SHOP: 'https://balloonstore1.myshopify.com/' }).shop, 'balloonstore1');
  assert.equal(cfg({ GIFT_CARD_EXPIRES_ON: '2027-01-31' }).giftCardExpiresOn, '2027-01-31');
  assert.equal(cfg({ GIFT_CARD_CURRENCY: 'usd' }).giftCardCurrency, 'USD');
});

test('scopes: write_ implies read_, read_all_orders must be explicit', () => {
  assert.deepEqual(missingScopes(parseScopes('write_gift_cards, write_customers,write_orders,read_all_orders')), []);
  assert.deepEqual(missingScopes(parseScopes('write_gift_cards,write_customers,read_orders')), ['read_all_orders']);
  assert.deepEqual(missingScopes(parseScopes('read_gift_cards,read_customers')), ['write_gift_cards', 'write_customers', 'read_orders', 'read_all_orders']);
  assert.deepEqual(missingScopes(parseScopes(undefined)), REQUIRED_SCOPES);
  assert.deepEqual(REQUIRED_SCOPES, ['write_gift_cards', 'write_customers', 'read_orders', 'read_all_orders']);
});

test('progress: missing file -> {}, atomic save, corrupt file refuses to load', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'gift-card-promo-'));
  const file = path.join(dir, 'progress.json');
  try {
    assert.deepEqual(loadProgress(file), {});
    saveProgress(file, { a: { giftCardId: 'x' } });
    assert.deepEqual(loadProgress(file), { a: { giftCardId: 'x' } });
    assert.equal(fs.existsSync(`${file}.tmp`), false);
    assert.ok(fs.readFileSync(file, 'utf8').endsWith('\n'));
    fs.writeFileSync(file, '');
    assert.deepEqual(loadProgress(file), {});
    fs.writeFileSync(file, '{ not json');
    assert.throws(() => loadProgress(file), /not valid JSON/);
    fs.writeFileSync(file, '[]');
    assert.throws(() => loadProgress(file), /JSON object/);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('gift cards: amount formatting and input shape; the code is never requested', () => {
  assert.equal(formatAmount(0.1), '0.10');
  assert.equal(formatAmount('5'), '5.00');
  assert.equal(formatAmount(12.345), '12.35');
  assert.throws(() => formatAmount(0), /Invalid gift card amount/);
  assert.throws(() => formatAmount('abc'), /Invalid gift card amount/);
  assert.deepEqual(
    buildCreateInput('gid://shopify/Customer/1', { amount: 0.1, currencyCode: 'USD', note: '', expiresOn: '', templateSuffix: '' }),
    { customerId: 'gid://shopify/Customer/1', initialAmount: { amount: '0.10', currencyCode: 'USD' } },
  );
  assert.equal(GIFT_CARD_CREATE.includes('giftCardCode'), false);
  assert.equal(GIFT_CARD_CREATE.includes('initialValue'), false);
  assert.match(GIFT_CARD_CREATE, /\$input: GiftCardCreateInput!/);
});

test('gift cards: no explicit notification send exists (Shopify emails the card on creation)', () => {
  // giftCardCreate with customerId already emails the customer; a second call to
  // giftCardSendNotificationToCustomer produced duplicate emails (2026-09-30).
  assert.equal(giftcards.sendNotification, undefined);
  assert.equal(giftcards.GIFT_CARD_SEND, undefined);
  for (const doc of Object.values(giftcards).filter((v) => typeof v === 'string')) {
    assert.equal(doc.includes('giftCardSendNotificationToCustomer'), false);
  }
});

test('segment: member ids are normalised to Customer GIDs before any write', () => {
  assert.equal(toCustomerGid('gid://shopify/CustomerSegmentMember/9585672781887'), 'gid://shopify/Customer/9585672781887');
  assert.equal(toCustomerGid('gid://shopify/Customer/42'), 'gid://shopify/Customer/42');
  assert.throws(() => toCustomerGid('gid://shopify/Segment/1'), /Unexpected segment member id/);
  assert.throws(() => toCustomerGid('9585672781887'), /Unexpected segment member id/);
  assert.throws(() => toCustomerGid('gid://shopify/CustomerSegmentMember/abc'), /Unexpected segment member id/);
});

test('customers: tag search query is quoted and escaped; the lookup paginates', () => {
  assert.equal(tagSearchQuery('gift-card-sent-2026-09'), 'tag:"gift-card-sent-2026-09"');
  assert.equal(tagSearchQuery('a"b\\c'), 'tag:"a\\"b\\\\c"');
  assert.match(TAGGED_CUSTOMERS_QUERY, /customers\(first: \$first, after: \$after, query: \$query\)/);
  assert.match(TAGGED_CUSTOMERS_QUERY, /pageInfo \{ hasNextPage endCursor \}/);
});

test('GraphQL documents request only fields that exist on their payload types', () => {
  assert.match(TAGS_ADD, /userErrors \{ field message \}/);
  assert.equal(TAGS_ADD.includes('code'), false);
  assert.match(SEGMENT_MEMBERS_QUERY, /customerSegmentMembers\(segmentId: \$id, first: \$first, after: \$after\)/);
  assert.match(SEGMENT_MEMBERS_QUERY, /totalCount/);
  assert.match(SEGMENT_MEMBERS_QUERY, /pageInfo \{ hasNextPage endCursor \}/);
});
