import { test, beforeEach, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

import { run } from '../src/run.js';
import { resetClient } from '../src/shopify.js';
import { installFakeShopify } from './fake-shopify.js';

const ANN = { id: 'gid://shopify/Customer/1', displayName: 'Ann Apple', numberOfOrders: '2', lastOrderId: 'gid://shopify/Order/11', defaultEmailAddress: { emailAddress: 'ann@example.com' } };
const BOB = { id: 'gid://shopify/Customer/2', displayName: 'Bob Berry', numberOfOrders: '1', lastOrderId: 'gid://shopify/Order/12', defaultEmailAddress: null };
const CY = { id: 'gid://shopify/Customer/3', displayName: 'Cy Cherry', numberOfOrders: '3', lastOrderId: 'gid://shopify/Order/13', defaultEmailAddress: { emailAddress: 'cy@example.com' } };
const PAGES = [[ANN, BOB], [CY]];
const NOW = '2026-09-30T12:00:00.000Z';
const TAG = 'gift-card-sent-2026-09';

let dir;
let fake;
let logs;
let log;

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'gift-card-promo-'));
  logs = [];
  log = {
    info: (m) => logs.push(`INFO ${m}`),
    warn: (m) => logs.push(`WARN ${m}`),
    error: (m) => logs.push(`ERROR ${m}`),
  };
});

afterEach(() => {
  if (fake) fake.restore();
  fake = null;
  resetClient();
  fs.rmSync(dir, { recursive: true, force: true });
});

const progressFile = () => path.join(dir, 'progress.json');
const makeConfig = (overrides = {}) => ({
  shop: 'teststore',
  clientId: 'id',
  clientSecret: 'secret',
  apiVersion: '2026-07',
  segmentId: 'gid://shopify/Segment/1',
  segmentPageSize: 2,
  giftCardValue: 0.1,
  giftCardCurrency: 'USD',
  giftCardNote: 'IBC2026 test',
  giftCardExpiresOn: '',
  giftCardTemplateSuffix: '',
  sentTag: TAG,
  dryRun: false,
  progressFile: progressFile(),
  ...overrides,
});
const go = ({ mode = 'run', limit = 0, config = {} } = {}) =>
  run({ config: makeConfig(config), mode, limit, log, now: () => NOW, sleep: async () => {} });
const readProgress = () => JSON.parse(fs.readFileSync(progressFile(), 'utf8'));
const writeProgress = (obj) => fs.writeFileSync(progressFile(), JSON.stringify(obj));
const mutationCount = (s) => s.calls.create.length + s.calls.tag.length;

test('--check: form-encoded token request, scope check, no GraphQL calls', async () => {
  fake = installFakeShopify({ pages: PAGES });
  const result = await go({ mode: 'check' });
  assert.equal(result.exitCode, 0);
  assert.equal(fake.state.calls.token.length, 1);
  assert.equal(fake.state.calls.token[0].contentType, 'application/x-www-form-urlencoded');
  assert.deepEqual(fake.state.calls.token[0].body, { grant_type: 'client_credentials', client_id: 'id', client_secret: 'secret' });
  assert.equal(fake.state.calls.queries.length, 0);
  assert.equal(logs.some((l) => l.includes('secret') || l.includes('fake-token')), false, 'must never log the secret or token');
});

test('--check: fails when a required scope is missing', async () => {
  fake = installFakeShopify({ scope: 'write_gift_cards,write_customers,read_orders' });
  await assert.rejects(go({ mode: 'check' }), /missing required scopes: read_all_orders/);
});

test('--list: pages through the whole segment with the configured page size, no writes', async () => {
  fake = installFakeShopify({ pages: PAGES, tagged: [CY.id] });
  const result = await go({ mode: 'list' });
  assert.equal(result.exitCode, 0);
  assert.equal(result.totalCount, 3);
  assert.deepEqual(result.members.map((m) => m.id), [ANN.id, BOB.id, CY.id]);
  const pageQueries = fake.state.calls.queries.filter((q) => q.op === 'SegmentMembers');
  assert.equal(pageQueries.length, 2);
  assert.deepEqual(pageQueries.map((q) => q.variables.first), [2, 2]);
  assert.equal(pageQueries[0].variables.after, null);
  assert.equal(pageQueries[1].variables.after, 'cursor-1');
  assert.equal(pageQueries[0].token, 'fake-token');
  assert.equal(mutationCount(fake.state), 0);
  assert.equal(fs.existsSync(progressFile()), false);
  assert.ok(logs.some((l) => l.includes('Bob Berry') && l.includes('(no email)') && l.includes('[no email]')));
  // the API hands back CustomerSegmentMember ids; the listing must show the Customer id
  assert.ok(logs.some((l) => l.startsWith('INFO gid://shopify/Customer/1  Ann Apple') && l.endsWith('[pending]')));
  assert.ok(logs.some((l) => l.startsWith('INFO gid://shopify/Customer/3  Cy Cherry') && l.endsWith('[done (tagged in Shopify)]')));
  assert.ok(logs.some((l) => l.includes(`1 customer(s) already carry tag "${TAG}" in Shopify`)));
  assert.equal(logs.some((l) => l.includes('CustomerSegmentMember')), false);
});

test('live run: creates and tags every member with an email, in order', async () => {
  fake = installFakeShopify({ pages: PAGES });
  const result = await go();
  assert.equal(result.exitCode, 0);
  assert.deepEqual(result.summary, { processed: 2, created: 2, tagged: 2, skippedDone: 0, skippedTagged: 0, skippedNoEmail: 1, needsReview: 0, failed: 0 });
  assert.deepEqual(fake.state.calls.create, [
    { customerId: ANN.id, initialAmount: { amount: '0.10', currencyCode: 'USD' }, note: 'IBC2026 test' },
    { customerId: CY.id, initialAmount: { amount: '0.10', currencyCode: 'USD' }, note: 'IBC2026 test' },
  ]);
  assert.deepEqual(fake.state.calls.tag, [{ id: ANN.id, tags: [TAG] }, { id: CY.id, tags: [TAG] }]);
  // all pages and the already-tagged set are fetched before the first write; per customer: create -> tag
  assert.deepEqual(
    fake.state.calls.queries.map((q) => q.op),
    ['SegmentMembers', 'SegmentMembers', 'TaggedCustomers', 'GiftCardCreate', 'TagsAdd', 'GiftCardCreate', 'TagsAdd'],
  );
  assert.equal(fake.state.calls.queries.find((q) => q.op === 'TaggedCustomers').variables.query, `tag:"${TAG}"`);
  // Shopify emails the card as part of giftCardCreate; an explicit send would be a second email
  assert.equal(fake.state.calls.queries.some((q) => q.op === 'GiftCardSend'), false);
  const p = readProgress();
  assert.deepEqual(p[ANN.id], { displayName: 'Ann Apple', giftCardId: 'gid://shopify/GiftCard/1', createdAt: NOW, taggedAt: NOW });
  assert.deepEqual(p[CY.id], { displayName: 'Cy Cherry', giftCardId: 'gid://shopify/GiftCard/2', createdAt: NOW, taggedAt: NOW });
  assert.equal(p[BOB.id], undefined);
  assert.equal(fs.existsSync(`${progressFile()}.tmp`), false);
});

test('re-run after success: everything skipped, zero mutations, progress untouched', async () => {
  fake = installFakeShopify({ pages: PAGES });
  await go();
  const before = fs.readFileSync(progressFile(), 'utf8');
  const result = await go();
  assert.equal(result.exitCode, 0);
  assert.deepEqual(result.summary, { processed: 0, created: 0, tagged: 0, skippedDone: 2, skippedTagged: 0, skippedNoEmail: 1, needsReview: 0, failed: 0 });
  assert.equal(mutationCount(fake.state), 4);
  assert.equal(fs.readFileSync(progressFile(), 'utf8'), before);
});

test('resume: card already created -> only tag, never a second card', async () => {
  fake = installFakeShopify({ pages: PAGES });
  writeProgress({ [ANN.id]: { displayName: 'Ann Apple', giftCardId: 'gid://shopify/GiftCard/99', createdAt: NOW } });
  const result = await go();
  assert.equal(result.exitCode, 0);
  assert.deepEqual(fake.state.calls.create.map((c) => c.customerId), [CY.id]);
  assert.deepEqual(fake.state.calls.tag, [{ id: ANN.id, tags: [TAG] }, { id: CY.id, tags: [TAG] }]);
  const p = readProgress();
  assert.equal(p[ANN.id].giftCardId, 'gid://shopify/GiftCard/99');
  assert.equal(p[ANN.id].taggedAt, NOW);
  assert.deepEqual(result.summary, { processed: 2, created: 1, tagged: 2, skippedDone: 0, skippedTagged: 0, skippedNoEmail: 1, needsReview: 0, failed: 0 });
});

test('dry run: reads everything, writes nothing', async () => {
  fake = installFakeShopify({ pages: PAGES });
  const result = await go({ config: { dryRun: true } });
  assert.equal(result.exitCode, 0);
  assert.equal(result.summary.processed, 2);
  assert.equal(result.summary.skippedNoEmail, 1);
  assert.equal(mutationCount(fake.state), 0);
  assert.equal(fs.existsSync(progressFile()), false);
  assert.equal(logs.filter((l) => l.includes('[dry-run] would create a 0.10 USD gift card')).length, 2);
});

test('--limit 1 stops after the first customer actually processed', async () => {
  fake = installFakeShopify({ pages: PAGES });
  const result = await go({ limit: 1 });
  assert.equal(result.exitCode, 0);
  assert.equal(result.summary.processed, 1);
  assert.deepEqual(fake.state.calls.create.map((c) => c.customerId), [ANN.id]);
  assert.equal(fake.state.calls.tag.length, 1);
  assert.equal(readProgress()[CY.id], undefined);
});

test('--limit does not count skipped customers', async () => {
  fake = installFakeShopify({ pages: PAGES });
  writeProgress({ [ANN.id]: { displayName: 'Ann Apple', giftCardId: 'gid://shopify/GiftCard/99', createdAt: NOW, taggedAt: NOW } });
  const result = await go({ limit: 1 });
  assert.equal(result.exitCode, 0);
  assert.deepEqual(fake.state.calls.create.map((c) => c.customerId), [CY.id]);
});

test('--limit in dry run counts previewed customers', async () => {
  fake = installFakeShopify({ pages: PAGES });
  const result = await go({ limit: 1, config: { dryRun: true } });
  assert.equal(result.summary.processed, 1);
  assert.equal(logs.filter((l) => l.includes('[dry-run] would create')).length, 1);
});

test('already tagged in Shopify but missing from progress.json -> skipped, no card, no progress write', async () => {
  fake = installFakeShopify({ pages: PAGES, tagged: [ANN.id] });
  const result = await go();
  assert.equal(result.exitCode, 0);
  assert.deepEqual(result.summary, { processed: 1, created: 1, tagged: 1, skippedDone: 0, skippedTagged: 1, skippedNoEmail: 1, needsReview: 0, failed: 0 });
  assert.deepEqual(fake.state.calls.create.map((c) => c.customerId), [CY.id]);
  assert.deepEqual(fake.state.calls.tag, [{ id: CY.id, tags: [TAG] }]);
  assert.equal(readProgress()[ANN.id], undefined, 'nothing is recorded for a customer the script did not touch');
  assert.ok(logs.some((l) => l.startsWith('WARN skip Ann Apple') && l.includes(`already has tag "${TAG}" in Shopify but no record in progress.json`)));
});

test('already tagged in Shopify with a card in progress.json -> skipped, progress untouched', async () => {
  fake = installFakeShopify({ pages: [[ANN]], tagged: [ANN.id] });
  writeProgress({ [ANN.id]: { displayName: 'Ann Apple', giftCardId: 'gid://shopify/GiftCard/99', createdAt: NOW } });
  const before = fs.readFileSync(progressFile(), 'utf8');
  const result = await go();
  assert.equal(result.exitCode, 0);
  assert.equal(result.summary.skippedTagged, 1);
  assert.equal(mutationCount(fake.state), 0);
  assert.equal(fs.readFileSync(progressFile(), 'utf8'), before);
  const line = logs.find((l) => l.startsWith('WARN skip Ann Apple'));
  assert.ok(line && line.includes('already has tag') && !line.includes('no record'));
});

test('dry run: customers already tagged in Shopify are skipped, nothing written', async () => {
  fake = installFakeShopify({ pages: PAGES, tagged: [ANN.id] });
  const result = await go({ config: { dryRun: true } });
  assert.equal(result.exitCode, 0);
  assert.equal(result.summary.skippedTagged, 1);
  assert.equal(result.summary.processed, 1);
  assert.equal(mutationCount(fake.state), 0);
  assert.equal(fs.existsSync(progressFile()), false);
});

test('--limit 1 does not count customers skipped for the Shopify tag', async () => {
  fake = installFakeShopify({ pages: PAGES, tagged: [ANN.id] });
  const result = await go({ limit: 1 });
  assert.equal(result.summary.processed, 1);
  assert.deepEqual(fake.state.calls.create.map((c) => c.customerId), [CY.id]);
});

test('optional gift card fields are sent only when configured', async () => {
  fake = installFakeShopify({ pages: [[ANN]] });
  await go({ config: { giftCardExpiresOn: '2027-01-31', giftCardTemplateSuffix: 'winback', giftCardNote: '' } });
  assert.deepEqual(fake.state.calls.create, [
    { customerId: ANN.id, initialAmount: { amount: '0.10', currencyCode: 'USD' }, expiresOn: '2027-01-31', templateSuffix: 'winback' },
  ]);
});

test('throttled response is retried; the card is created exactly once', async () => {
  fake = installFakeShopify({ pages: [[ANN]], failures: { GiftCardCreate: [{ kind: 'throttle' }] } });
  const result = await go();
  assert.equal(result.exitCode, 0);
  assert.equal(fake.state.calls.queries.filter((q) => q.op === 'GiftCardCreate').length, 2);
  assert.equal(fake.state.calls.create.length, 1);
  assert.ok(logs.some((l) => /throttled; waiting 2s/.test(l)));
});

test('HTTP 429 is retried using Retry-After', async () => {
  fake = installFakeShopify({ pages: [[ANN]], failures: { TagsAdd: [{ kind: 'http', status: 429, headers: { 'retry-after': '1' } }] } });
  const result = await go();
  assert.equal(result.exitCode, 0);
  assert.equal(fake.state.calls.tag.length, 1);
  assert.ok(logs.some((l) => /rate limited \(HTTP 429\); waiting 1s/.test(l)));
});

test('gives up after repeated throttling and records the error', async () => {
  fake = installFakeShopify({ pages: [[ANN]], failures: { TagsAdd: Array(6).fill({ kind: 'throttle' }) } });
  const result = await go();
  assert.equal(result.exitCode, 1);
  assert.match(result.error.message, /Throttled by Shopify 6 times in a row/);
  assert.equal(fake.state.calls.create.length, 1);
  const p = readProgress();
  assert.equal(p[ANN.id].giftCardId, 'gid://shopify/GiftCard/1');
  assert.equal(p[ANN.id].taggedAt, undefined);
  assert.match(p[ANN.id].error.message, /Throttled/);
});

test('network error during create: marker kept, run stops, re-run refuses until reviewed', async () => {
  fake = installFakeShopify({ pages: PAGES, failures: { GiftCardCreate: [{ kind: 'network' }] } });
  const first = await go();
  assert.equal(first.exitCode, 1);
  assert.equal(first.summary.failed, 1);
  assert.equal(fake.state.calls.create.length, 0);
  let p = readProgress();
  assert.equal(p[ANN.id].createStartedAt, NOW);
  assert.equal(p[ANN.id].giftCardId, undefined);
  assert.match(p[ANN.id].error.message, /Network error/);
  assert.equal(p[CY.id], undefined, 'stopped at the first failure');
  assert.ok(logs.some((l) => l.startsWith('ERROR Stopped at customer Ann Apple')));

  const second = await go();
  assert.equal(second.exitCode, 1);
  assert.match(second.error.message, /result is unknown/);
  assert.equal(fake.state.calls.create.length, 0);
  p = readProgress();
  assert.equal(p[ANN.id].createStartedAt, NOW, 'marker must survive');
});

test('stale creation marker: dry run reports it instead of stopping', async () => {
  fake = installFakeShopify({ pages: PAGES });
  writeProgress({ [ANN.id]: { displayName: 'Ann Apple', createStartedAt: NOW } });
  const before = fs.readFileSync(progressFile(), 'utf8');
  const result = await go({ config: { dryRun: true } });
  assert.equal(result.exitCode, 0);
  assert.equal(result.summary.needsReview, 1);
  assert.equal(result.summary.processed, 1);
  assert.equal(fs.readFileSync(progressFile(), 'utf8'), before);
});

test('userErrors during create: marker cleared, next run creates the card', async () => {
  const userErrors = [{ field: ['input', 'initialAmount'], message: 'must be greater than 0', code: 'GREATER_THAN' }];
  fake = installFakeShopify({
    pages: [[ANN]],
    failures: { GiftCardCreate: [{ kind: 'data', data: { giftCardCreate: { giftCard: null, userErrors } } }] },
  });
  const first = await go();
  assert.equal(first.exitCode, 1);
  assert.match(first.error.message, /giftCardCreate rejected: input\.initialAmount: must be greater than 0 \[GREATER_THAN\]/);
  let p = readProgress();
  assert.equal(p[ANN.id].createStartedAt, undefined, 'definitely-rejected create must not leave a marker');
  assert.equal(p[ANN.id].giftCardId, undefined);
  assert.ok(p[ANN.id].error);

  const second = await go();
  assert.equal(second.exitCode, 0);
  assert.equal(fake.state.calls.create.length, 1);
  p = readProgress();
  assert.equal(p[ANN.id].giftCardId, 'gid://shopify/GiftCard/1');
  assert.equal(p[ANN.id].error, undefined);
});

test('5xx during tag: card kept, tag retried next run, no second card', async () => {
  fake = installFakeShopify({ pages: [[ANN]], failures: { TagsAdd: [{ kind: 'http', status: 502, body: 'Bad Gateway' }] } });
  const first = await go();
  assert.equal(first.exitCode, 1);
  let p = readProgress();
  assert.equal(p[ANN.id].giftCardId, 'gid://shopify/GiftCard/1');
  assert.equal(p[ANN.id].taggedAt, undefined);
  assert.match(p[ANN.id].error.message, /HTTP 502/);

  const second = await go();
  assert.equal(second.exitCode, 0);
  assert.equal(fake.state.calls.create.length, 1);
  assert.equal(fake.state.calls.queries.filter((q) => q.op === 'TagsAdd').length, 2);
  assert.deepEqual(fake.state.calls.tag, [{ id: ANN.id, tags: [TAG] }]);
  p = readProgress();
  assert.equal(p[ANN.id].taggedAt, NOW);
  assert.equal(p[ANN.id].error, undefined);
});

test('MAX_COST_EXCEEDED while paging surfaces a page-size hint and writes nothing', async () => {
  fake = installFakeShopify({
    pages: PAGES,
    failures: { SegmentMembers: [{ kind: 'graphqlError', message: 'Query cost is 2002, which exceeds the single query max cost limit (1000).', code: 'MAX_COST_EXCEEDED' }] },
  });
  await assert.rejects(go(), /MAX_COST_EXCEEDED.*lower SEGMENT_PAGE_SIZE/);
  assert.equal(mutationCount(fake.state), 0);
  assert.equal(fs.existsSync(progressFile()), false);
});

test('corrupt progress.json refuses to run', async () => {
  fake = installFakeShopify({ pages: PAGES });
  fs.writeFileSync(progressFile(), '{ not json');
  await assert.rejects(go(), /not valid JSON/);
  assert.equal(mutationCount(fake.state), 0);
});
