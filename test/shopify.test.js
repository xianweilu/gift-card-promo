import { test, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import { gql, initClient, resetClient, throwIfUserErrors, ShopifyError } from '../src/shopify.js';
import { missingScopes, parseScopes, assertScopes, REQUIRED_SCOPES } from '../src/auth.js';
import { connect } from '../src/connect.js';
import { SHOP_INFO, ORDERS_SINCE } from '../src/queries.js';
import { installFakeShopify } from './fake-shopify.js';
import { memoryLog, testConfig } from './helpers.js';

let fake;
let sleeps;
function install(options) {
  fake?.restore(); // a second install in one test must not stack fakes
  fake = installFakeShopify(options);
  sleeps = [];
  initClient({ shop: 'teststore', apiVersion: '2026-07', token: 'fake-token', log: memoryLog(), sleep: async (ms) => { sleeps.push(ms); } });
  return fake;
}
afterEach(() => {
  fake?.restore();
  fake = null;
  resetClient();
});

test('gql: refuses to run before the client is initialised', async () => {
  resetClient();
  await assert.rejects(gql(SHOP_INFO), /not initialised/);
});

test('gql: THROTTLED is retried with a wait, then succeeds', async () => {
  const f = install({ failures: { ShopInfo: [{ kind: 'throttle' }, { kind: 'throttle' }] } });
  const data = await gql(SHOP_INFO);
  assert.equal(data.shop.ianaTimezone, 'America/Los_Angeles');
  assert.equal(f.opsNamed('ShopInfo').length, 3);
  assert.equal(sleeps.length, 2);
  assert.ok(sleeps.every((ms) => ms >= 2000));
});

test('gql: throttled six times → a known rejection (safe to retry later)', async () => {
  install({ failures: { ShopInfo: Array.from({ length: 6 }, () => ({ kind: 'throttle' })) } });
  await assert.rejects(gql(SHOP_INFO), (err) => err instanceof ShopifyError && err.outcomeKnown === true && err.code === 'THROTTLED');
});

test('gql: HTTP 429 is retried; persistent 429 is a known rejection', async () => {
  const f = install({ failures: { ShopInfo: [{ kind: 'http', status: 429 }] } });
  await gql(SHOP_INFO);
  assert.equal(f.opsNamed('ShopInfo').length, 2);
  assert.deepEqual(sleeps, [2000]);
  install({ failures: { ShopInfo: Array.from({ length: 6 }, () => ({ kind: 'http', status: 429 })) } });
  await assert.rejects(gql(SHOP_INFO), (err) => err.outcomeKnown === true && err.status === 429);
});

test('gql: outcome of failures — unknown for network/5xx/internal errors, known for 4xx and other GraphQL errors', async () => {
  install({ failures: { ShopInfo: [{ kind: 'network' }] } });
  await assert.rejects(gql(SHOP_INFO), (err) => err.outcomeKnown === false && /Network error/.test(err.message));
  install({ failures: { ShopInfo: [{ kind: 'http', status: 502 }] } });
  await assert.rejects(gql(SHOP_INFO), (err) => err.outcomeKnown === false && err.status === 502);
  install({ failures: { ShopInfo: [{ kind: 'http', status: 400 }] } });
  await assert.rejects(gql(SHOP_INFO), (err) => err.outcomeKnown === true && err.status === 400);
  install({ failures: { ShopInfo: [{ kind: 'graphqlError', message: 'Internal error', code: 'INTERNAL_SERVER_ERROR' }] } });
  await assert.rejects(gql(SHOP_INFO), (err) => err.outcomeKnown === false && err.code === 'INTERNAL_SERVER_ERROR');
  install({ failures: { ShopInfo: [{ kind: 'graphqlError', message: 'Field x does not exist' }] } });
  await assert.rejects(gql(SHOP_INFO), (err) => err.outcomeKnown === true && /Field x does not exist/.test(err.message));
});

test('gql: a search filter Shopify ignored stops the command instead of returning everything', async () => {
  install({ searchWarnings: { OrdersSince: true } });
  await assert.rejects(gql(ORDERS_SINCE, { query: "created_at:>='2026-07-01T00:00:00-07:00'" }), (err) => err.code === 'SEARCH_WARNING' && err.outcomeKnown === true && /ignored part of a search filter/.test(err.message));
});

test('gql: the access token goes in the header, never in the URL or body', async () => {
  install();
  const seen = [];
  const inner = globalThis.fetch;
  globalThis.fetch = async (url, init) => {
    seen.push({ url: String(url), headers: init.headers, body: String(init.body) });
    return inner(url, init);
  };
  try {
    await gql(SHOP_INFO);
  } finally {
    globalThis.fetch = inner;
  }
  assert.equal(seen[0].url, 'https://teststore.myshopify.com/admin/api/2026-07/graphql.json');
  assert.equal(seen[0].headers['X-Shopify-Access-Token'], 'fake-token');
  assert.equal(seen[0].body.includes('fake-token'), false);
});

test('userErrors: none → ok; any → permanent rejection with the field path', () => {
  throwIfUserErrors('op', []);
  throwIfUserErrors('op', undefined);
  assert.throws(() => throwIfUserErrors('giftCardCreate', [{ field: ['input', 'customerId'], message: 'is invalid', code: 'INVALID' }]), (err) => err.code === 'USER_ERROR' && err.outcomeKnown === true && err.message === 'giftCardCreate rejected: input.customerId: is invalid [INVALID]');
});

test('scopes: write_ implies read_; read_all_orders must be granted explicitly', () => {
  assert.deepEqual(REQUIRED_SCOPES, ['write_gift_cards', 'write_customers', 'read_orders', 'read_all_orders']);
  assert.deepEqual(missingScopes(parseScopes('write_gift_cards, write_customers,write_orders,read_all_orders')), []);
  assert.deepEqual(missingScopes(parseScopes('write_gift_cards,write_customers,read_orders')), ['read_all_orders']);
  assert.deepEqual(missingScopes(parseScopes(undefined)), REQUIRED_SCOPES);
  assert.throws(() => assertScopes(['read_customers']), /missing required scopes: write_gift_cards, write_customers, read_orders, read_all_orders/);
});

test('connect: client-credentials token, scope check, then GraphQL works', async () => {
  const { config, cleanup } = testConfig();
  try {
    const f = installFakeShopify();
    fake = f;
    const { scopes } = await connect(config, { log: memoryLog(), sleep: async () => {} });
    assert.ok(scopes.includes('write_gift_cards'));
    assert.deepEqual(f.state.calls.token[0], { grant_type: 'client_credentials', client_id: 'id', client_secret: 'secret' });
    assert.equal((await gql(SHOP_INFO)).shop.currencyCode, 'USD');
    f.restore();

    fake = installFakeShopify({ scope: 'write_gift_cards,write_customers,read_orders' });
    await assert.rejects(connect(config, { log: memoryLog() }), /read_all_orders/);
  } finally {
    cleanup();
  }
});
