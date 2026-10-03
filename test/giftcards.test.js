import { test, beforeEach, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import * as queries from '../src/queries.js';
import {
  GIFT_CARD_CREATE, centsToDecimal, campaignMarker, campaignNote, hasCampaignMarker, buildCreateInput,
  createGiftCard, campaignCardsQuery, findCampaignCards, sendGiftCardNotification,
} from '../src/giftcards.js';
import { addTag, fetchTaggedCustomerIds, tagSearchQuery, TAGS_ADD } from '../src/customers.js';
import { initClient, resetClient } from '../src/shopify.js';
import { installFakeShopify } from './fake-shopify.js';
import { makeCustomer, memoryLog } from './helpers.js';

let fake;
const cid = (n) => `gid://shopify/Customer/${n}`;
const card = (n, { customer = 1, note = 'gift-card-promo [campaign:2026-10]', createdAt = '2026-10-05T16:00:00Z', amount = '10.77', balance = amount, enabled = true, expiresOn = '2026-10-19' } = {}) => ({
  id: `gid://shopify/GiftCard/${n}`,
  createdAt,
  note,
  templateSuffix: 'gift-card-promo',
  enabled,
  expiresOn,
  lastCharacters: `a${String(n).padStart(3, '0')}`,
  initialValue: { amount, currencyCode: 'USD' },
  balance: { amount: balance, currencyCode: 'USD' },
  customer: customer ? { id: cid(customer) } : null,
});

function install(options) {
  fake?.restore(); // a second install in one test must not stack fakes
  fake = installFakeShopify(options);
  initClient({ shop: 'teststore', apiVersion: '2026-07', token: 'fake-token', log: memoryLog(), sleep: async () => {} });
  return fake;
}
beforeEach(() => resetClient());
afterEach(() => {
  fake?.restore();
  fake = null;
  resetClient();
});

test('gift cards: amounts are sent as exact decimals from integer cents', () => {
  assert.equal(centsToDecimal(1077), '10.77');
  assert.equal(centsToDecimal(10), '0.10');
  assert.equal(centsToDecimal(1977), '19.77');
  assert.equal(centsToDecimal(200000), '2000.00');
  for (const bad of [0, -5, 10.5, '1077', null]) assert.throws(() => centsToDecimal(bad), /Invalid gift card amount/);
});

test('gift cards: the campaign marker in the internal note', () => {
  assert.equal(campaignMarker('2026-10'), '[campaign:2026-10]');
  assert.equal(campaignNote('gift-card-promo', '2026-10'), 'gift-card-promo [campaign:2026-10]');
  assert.equal(campaignNote('', '2026-10'), '[campaign:2026-10]');
  assert.equal(hasCampaignMarker('IBC2026 test [campaign:2026-10]', '2026-10'), true);
  assert.equal(hasCampaignMarker('[campaign:2026-10-test]', '2026-10'), false, 'the test campaign is a different campaign');
  assert.equal(hasCampaignMarker(null, '2026-10'), false);
});

test('gift cards: create input; the card code is never requested', () => {
  assert.deepEqual(
    buildCreateInput(cid(1), { amountCents: 1533, currencyCode: 'USD', note: 'gift-card-promo [campaign:2026-10]', expiresOn: '2026-10-19', templateSuffix: 'gift-card-promo' }),
    { customerId: cid(1), initialAmount: { amount: '15.33', currencyCode: 'USD' }, note: 'gift-card-promo [campaign:2026-10]', expiresOn: '2026-10-19', templateSuffix: 'gift-card-promo' },
  );
  assert.deepEqual(buildCreateInput(cid(1), { amountCents: 10, currencyCode: 'USD', note: '', expiresOn: '', templateSuffix: '' }), { customerId: cid(1), initialAmount: { amount: '0.10', currencyCode: 'USD' } });
  const documents = [GIFT_CARD_CREATE, ...Object.values(queries).filter((v) => typeof v === 'string')];
  for (const doc of documents) {
    assert.equal(/giftCardCode|maskedCode/.test(doc), false, 'no document may read a gift card code');
  }
  assert.match(GIFT_CARD_CREATE, /giftCard \{ id lastCharacters \}/);
});

test('gift cards: create through the API returns the id and last 4 characters only', async () => {
  const f = install({ customers: [makeCustomer({ n: 1 })] });
  const created = await createGiftCard(cid(1), { amountCents: 1077, currencyCode: 'USD', note: 'n [campaign:2026-10]', expiresOn: '2026-10-19', templateSuffix: 'gift-card-promo' });
  assert.deepEqual(Object.keys(created).sort(), ['id', 'last4']);
  assert.match(created.id, /^gid:\/\/shopify\/GiftCard\/\d+$/);
  assert.equal(f.state.calls.create.length, 1);
  assert.equal(f.state.calls.create[0].initialAmount.amount, '10.77');
});

test('gift cards: a userError is a permanent, known rejection', async () => {
  install({ failures: { GiftCardCreate: [{ kind: 'userError', message: 'Customer is invalid' }] } });
  await assert.rejects(createGiftCard(cid(1), { amountCents: 1077, currencyCode: 'USD' }), (err) => err.code === 'USER_ERROR' && err.outcomeKnown === true && /Customer is invalid/.test(err.message));
});

test('gift cards: a lost response is an unknown outcome (the card may exist)', async () => {
  const f = install({ failures: { GiftCardCreate: [{ kind: 'appliedThenLost' }] } });
  await assert.rejects(createGiftCard(cid(1), { amountCents: 1077, currencyCode: 'USD' }), (err) => err.outcomeKnown === false);
  assert.equal(f.state.giftCards.length, 1, 'Shopify created it even though we never heard back');
});

test('gift cards: campaign search query', () => {
  assert.equal(campaignCardsQuery('2026-10-01T19:00:00.000Z'), "created_at:>='2026-10-01T19:00:00.000Z' source:api_client");
  assert.equal(campaignCardsQuery('2026-10-01T19:00:00.000Z', cid(42)), "created_at:>='2026-10-01T19:00:00.000Z' source:api_client customer_id:42");
  assert.equal(campaignCardsQuery('2026-10-01T19:00:00.000Z', '42'), "created_at:>='2026-10-01T19:00:00.000Z' source:api_client customer_id:42");
});

test('gift cards: findCampaignCards keeps only this campaign, maps money to cents and pages through', async () => {
  const many = Array.from({ length: 260 }, (_, i) => card(2000 + i, { customer: 100 + i }));
  install({
    giftCards: [
      card(1, { customer: 1, amount: '15.33', balance: '5.33', enabled: false }),
      card(2, { customer: 2, note: 'manual card' }),
      card(3, { customer: 3, note: 'x [campaign:2026-10-test]' }),
      card(4, { customer: 4, createdAt: '2026-09-01T00:00:00Z' }),
      card(5, { customer: null }),
      ...many,
    ],
  });
  await assert.rejects(findCampaignCards({ campaignId: '2026-10' }), /needs campaignId and sinceIso/);
  const cards = await findCampaignCards({ campaignId: '2026-10', sinceIso: '2026-10-01T00:00:00Z' });
  assert.equal(cards.length, 2 + 260, 'cards 1 and 5 plus 260 over two pages');
  const first = cards.find((c) => c.id === 'gid://shopify/GiftCard/1');
  assert.deepEqual(first, {
    id: 'gid://shopify/GiftCard/1', createdAt: '2026-10-05T16:00:00Z', note: 'gift-card-promo [campaign:2026-10]', templateSuffix: 'gift-card-promo',
    enabled: false, expiresOn: '2026-10-19', last4: 'a001', amountCents: 1533, balanceCents: 533, currencyCode: 'USD', customerId: cid(1),
  });
  assert.equal(cards.find((c) => c.id === 'gid://shopify/GiftCard/5').customerId, null);
  const one = await findCampaignCards({ campaignId: '2026-10', sinceIso: '2026-10-01T00:00:00Z', customerId: cid(1) });
  assert.deepEqual(one.map((c) => c.id), ['gid://shopify/GiftCard/1']);
  fake.state.hiddenCardIds.add('gid://shopify/GiftCard/1');
  assert.deepEqual(await findCampaignCards({ campaignId: '2026-10', sinceIso: '2026-10-01T00:00:00Z', customerId: cid(1) }), [], 'search-index lag hides a fresh card');
});

test('gift cards: re-sending the notification', async () => {
  const f = install({ giftCards: [card(1), card(2, { customer: null })] });
  assert.deepEqual(await sendGiftCardNotification('gid://shopify/GiftCard/1'), { id: 'gid://shopify/GiftCard/1' });
  assert.deepEqual(f.state.calls.notify, ['gid://shopify/GiftCard/1']);
  await assert.rejects(sendGiftCardNotification('gid://shopify/GiftCard/2'), (err) => err.code === 'USER_ERROR' && /no customer/.test(err.message));
  await assert.rejects(sendGiftCardNotification('gid://shopify/GiftCard/404'), (err) => err.code === 'USER_ERROR' && /not found/.test(err.message));
  assert.match(queries.SEND_GIFT_CARD_NOTIFICATION, /giftCardSendNotificationToCustomer\(id: \$id\)/);
  // The fake's userError failure kind uses this mutation's real payload field.
  const g = install({ giftCards: [card(1)], failures: { SendGiftCardNotification: [{ kind: 'userError', message: 'Not eligible' }] } });
  await assert.rejects(sendGiftCardNotification('gid://shopify/GiftCard/1'), (err) => err.code === 'USER_ERROR' && /Not eligible/.test(err.message));
  assert.deepEqual(g.state.calls.notify, ['gid://shopify/GiftCard/1']);
});

test('customers: tags are added one at a time and looked up with an exact, escaped search', async () => {
  assert.equal(tagSearchQuery('OCT26RTPROMO'), 'tag:"OCT26RTPROMO"');
  assert.equal(tagSearchQuery('a"b\\c'), 'tag:"a\\"b\\\\c"');
  assert.match(TAGS_ADD, /userErrors \{ field message \}/);
  const customers = Array.from({ length: 260 }, (_, i) => makeCustomer({ n: i + 1, tags: i % 2 ? ['OCT26RTPROMO'] : [] }));
  const f = install({ customers });
  const ids = await fetchTaggedCustomerIds('OCT26RTPROMO', { pageSize: 100 });
  assert.equal(ids.size, 130);
  assert.ok(ids.has(cid(2)));
  await addTag(cid(1), 'OCT26RTPROMO');
  assert.deepEqual(f.state.calls.tag, [{ id: cid(1), tags: ['OCT26RTPROMO'] }]);
  assert.equal((await fetchTaggedCustomerIds('oct26rtpromo')).size, 131, 'Shopify tag search is case-insensitive');
  await assert.rejects(addTag(cid(999), 'OCT26RTPROMO'), (err) => err.code === 'USER_ERROR');
});
