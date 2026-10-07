import { test } from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { loadConfig, ConfigError, ROOT_DIR } from '../src/config.js';

// Never reads the project's .env: every call passes envFile: false.
const BASE = { SHOP: 'balloonstore1', CLIENT_ID: 'id', CLIENT_SECRET: 'secret', CAMPAIGN_ID: '2026-10', SENT_TAG: 'OCT26RTPROMO' };
const cfg = (extra = {}) => loadConfig({ ...BASE, ...extra }, { envFile: false });

test('config: lists every missing required value', () => {
  assert.throws(
    () => loadConfig({ SHOP: 'x' }, { envFile: false }),
    (err) => err instanceof ConfigError && /CLIENT_ID, CLIENT_SECRET, CAMPAIGN_ID, SENT_TAG/.test(err.message),
  );
});

test('config: the agreed campaign defaults', () => {
  const c = cfg();
  assert.equal(c.dryRun, true);
  assert.equal(c.apiVersion, '2026-07');
  assert.equal(c.inactiveMonths, 3);
  assert.equal(c.minAccountAgeDays, 7);
  assert.equal(c.requireEmailSubscribed, true);
  assert.equal(c.giftPercent, 10);
  assert.deepEqual(c.giftTiersCents, [1077, 1533, 1977]);
  assert.equal(c.giftCardCurrency, 'USD');
  assert.equal(c.issueMaxPerRun, 20000);
  assert.equal(c.testGiftAmountCents, 10);
  assert.deepEqual(c.testCustomerIds, []);
  assert.equal(c.timezone, '');
  assert.equal(c.campaignsDir, path.join(ROOT_DIR, 'campaigns'));
  assert.deepEqual(c.excludeTags, ['WHS', 'PotentiallyWHS', 'DISC', 'level1A', 'level2A', 'DO NOT SELL', 'likely-fake-account', 'walmart.com', 'Walmart', 'Amazon', 'eBay']);
  assert.equal(c.excludeTags.includes('former-fake-identified'), false, 'the owner asked not to exclude former-fake-identified');
  assert.deepEqual(c.excludeEmailDomains, [
    'example.com', 'mail.codisto.com', 'connectebay.com', 'marketplace.amazon.com', 'relay.walmart.com', 'members.ebay.com',
    'jokerpartysupply.com', 'loftus.com', 'burtonandburton.com', 'mayflower.com', 'toyworld.com', 'rainbowballoons.com',
  ]);
  assert.deepEqual(c.excludeOrderSources, ['amazon', 'walmart', 'ebay', 'etsy', '205641', '1456995', '1775805']);
  assert.equal(c.launchDate, '');
  assert.equal(c.giftCardExpiresOn, '');
});

test('config: DRY_RUN is only off for the literal string false', () => {
  assert.equal(cfg({ DRY_RUN: 'false' }).dryRun, false);
  assert.equal(cfg({ DRY_RUN: 'FALSE' }).dryRun, false);
  assert.equal(cfg({ DRY_RUN: ' false ' }).dryRun, false);
  for (const v of ['0', 'no', 'off', '', 'true', 'yes', 'flase']) {
    assert.equal(cfg({ DRY_RUN: v }).dryRun, true, `DRY_RUN=${JSON.stringify(v)} must stay a dry run`);
  }
});

test('config: lists — unset uses the default, set-but-empty means none, domains/sources are lower-cased', () => {
  assert.deepEqual(cfg({ EXCLUDE_TAGS: '' }).excludeTags, []);
  assert.deepEqual(cfg({ EXCLUDE_TAGS: ' WHS , DISC ,' }).excludeTags, ['WHS', 'DISC']);
  assert.deepEqual(cfg({ EXCLUDE_EMAIL_DOMAINS: 'Example.COM' }).excludeEmailDomains, ['example.com']);
  assert.deepEqual(cfg({ EXCLUDE_ORDER_SOURCES: 'Amazon,205641' }).excludeOrderSources, ['amazon', '205641']);
});

test('config: gift tiers and percent', () => {
  assert.deepEqual(cfg({ GIFT_TIERS: '5, 10.5,20' }).giftTiersCents, [500, 1050, 2000]);
  assert.deepEqual(cfg({ GIFT_TIERS: '12' }).giftTiersCents, [1200]);
  assert.throws(() => cfg({ GIFT_TIERS: '15.33,10.77' }), /strictly increasing/);
  assert.throws(() => cfg({ GIFT_TIERS: '10.77,10.77' }), /strictly increasing/);
  assert.throws(() => cfg({ GIFT_TIERS: '10.777' }), /dollar amounts/);
  assert.throws(() => cfg({ GIFT_TIERS: '0' }), /between/);
  assert.throws(() => cfg({ GIFT_TIERS: '$10' }), /dollar amounts/);
  assert.equal(cfg({ GIFT_PERCENT: '12.5' }).giftPercent, 12.5);
  assert.throws(() => cfg({ GIFT_PERCENT: '0' }), /GIFT_PERCENT/);
  assert.throws(() => cfg({ GIFT_PERCENT: '101' }), /GIFT_PERCENT/);
  assert.throws(() => cfg({ GIFT_PERCENT: 'ten' }), /GIFT_PERCENT/);
});

test('config: campaign dates are real YYYY-MM-DD dates, launch not after expiry', () => {
  const c = cfg({ LAUNCH_DATE: '2026-10-05', GIFT_CARD_EXPIRES_ON: '2026-10-19' });
  assert.deepEqual([c.launchDate, c.giftCardExpiresOn], ['2026-10-05', '2026-10-19']);
  // Launch may fall on the last usable day.
  assert.equal(cfg({ LAUNCH_DATE: '2026-10-19', GIFT_CARD_EXPIRES_ON: '2026-10-19' }).launchDate, '2026-10-19');
  assert.throws(() => cfg({ LAUNCH_DATE: '2026/10/05' }), /LAUNCH_DATE must be a real date in the format YYYY-MM-DD/);
  assert.throws(() => cfg({ LAUNCH_DATE: '10-05-2026' }), /LAUNCH_DATE/);
  assert.throws(() => cfg({ LAUNCH_DATE: '2026-02-30' }), /LAUNCH_DATE/);
  assert.throws(() => cfg({ GIFT_CARD_EXPIRES_ON: '2026-13-01' }), /GIFT_CARD_EXPIRES_ON/);
  assert.throws(() => cfg({ LAUNCH_DATE: '2026-10-20', GIFT_CARD_EXPIRES_ON: '2026-10-19' }), /LAUNCH_DATE \(2026-10-20\) must not be later than GIFT_CARD_EXPIRES_ON \(2026-10-19\)/);
  // Either date alone is fine.
  assert.equal(cfg({ GIFT_CARD_EXPIRES_ON: '2026-10-19' }).launchDate, '');
  assert.equal(cfg({ LAUNCH_DATE: '2026-10-05' }).giftCardExpiresOn, '');
});

test('config: LAUNCH_DATE_NEVER (the never-ordered group) is a date between LAUNCH_DATE and the expiry', () => {
  const c = cfg({ LAUNCH_DATE: '2026-10-07', LAUNCH_DATE_NEVER: '2026-10-12', GIFT_CARD_EXPIRES_ON: '2026-10-19' });
  assert.equal(c.launchDateNever, '2026-10-12');
  assert.equal(cfg({ LAUNCH_DATE: '2026-10-07' }).launchDateNever, '', 'optional');
  // The same day as LAUNCH_DATE or as the expiry is allowed.
  assert.equal(cfg({ LAUNCH_DATE: '2026-10-07', LAUNCH_DATE_NEVER: '2026-10-07' }).launchDateNever, '2026-10-07');
  assert.equal(cfg({ LAUNCH_DATE_NEVER: '2026-10-19', GIFT_CARD_EXPIRES_ON: '2026-10-19' }).launchDateNever, '2026-10-19');
  assert.throws(() => cfg({ LAUNCH_DATE_NEVER: '2026-10-1' }), /LAUNCH_DATE_NEVER must be a real date in the format YYYY-MM-DD/);
  assert.throws(() => cfg({ LAUNCH_DATE_NEVER: '2026-10-20', GIFT_CARD_EXPIRES_ON: '2026-10-19' }), /LAUNCH_DATE_NEVER \(2026-10-20\) must not be later than GIFT_CARD_EXPIRES_ON \(2026-10-19\)/);
  assert.throws(() => cfg({ LAUNCH_DATE: '2026-10-07', LAUNCH_DATE_NEVER: '2026-10-06' }), /LAUNCH_DATE_NEVER \(2026-10-06\) must not be earlier than LAUNCH_DATE \(2026-10-07\)/);
});

test('config: the reminder dates are gone (reminders are sent with Shopify Email); old .env keys are ignored', () => {
  const c = cfg({ LAUNCH_DATE: '2026-10-05', REMIND_1_DATE: '2026-10-12', REMIND_2_DATE: 'not-a-date', GIFT_CARD_EXPIRES_ON: '2026-10-19' });
  assert.equal('remind1Date' in c, false);
  assert.equal('remind2Date' in c, false);
  assert.equal(Object.keys(c).some((k) => /remind/i.test(k)), false);
});

test('config: test campaigns must say so in CAMPAIGN_ID', () => {
  assert.throws(() => cfg({ TEST_CUSTOMER_IDS: '123' }), /CAMPAIGN_ID must contain "test"/);
  const c = cfg({ CAMPAIGN_ID: '2026-10-test', SENT_TAG: 'OCT26RTPROMO-TEST', TEST_CUSTOMER_IDS: '123, gid://shopify/Customer/456', TEST_GIFT_AMOUNT: '10.77' });
  assert.deepEqual(c.testCustomerIds, ['gid://shopify/Customer/123', 'gid://shopify/Customer/456']);
  assert.equal(c.testGiftAmountCents, 1077);
  assert.throws(() => cfg({ CAMPAIGN_ID: 'x-test', TEST_CUSTOMER_IDS: 'gid://shopify/Order/1' }), /TEST_CUSTOMER_IDS/);
  assert.throws(() => cfg({ TEST_GIFT_AMOUNT: '0' }), /TEST_GIFT_AMOUNT/);
  assert.throws(() => cfg({ TEST_GIFT_AMOUNT: '25' }), /TEST_GIFT_AMOUNT/);
});

test('config: other validation', () => {
  assert.equal(cfg({ SHOP: 'https://balloonstore1.myshopify.com/admin' }).shop, 'balloonstore1');
  assert.throws(() => cfg({ SHOP: 'bad shop' }), /SHOP looks wrong/);
  assert.throws(() => cfg({ API_VERSION: 'unstable' }), /API_VERSION/);
  assert.throws(() => cfg({ API_VERSION: '2026' }), /API_VERSION/);
  assert.throws(() => cfg({ SENT_TAG: 'a,b' }), /SENT_TAG/);
  assert.throws(() => cfg({ CAMPAIGN_ID: '2026/10' }), /CAMPAIGN_ID/);
  assert.throws(() => cfg({ CAMPAIGN_ID: '../x' }), /CAMPAIGN_ID/);
  assert.throws(() => cfg({ GIFT_CARD_CURRENCY: 'US' }), /GIFT_CARD_CURRENCY/);
  assert.equal(cfg({ GIFT_CARD_CURRENCY: 'usd' }).giftCardCurrency, 'USD');
  assert.throws(() => cfg({ TIMEZONE: 'Mars/Base' }), /TIMEZONE/);
  assert.equal(cfg({ TIMEZONE: 'America/Los_Angeles' }).timezone, 'America/Los_Angeles');
  assert.throws(() => cfg({ INACTIVE_MONTHS: '0' }), /INACTIVE_MONTHS/);
  assert.throws(() => cfg({ MIN_ACCOUNT_AGE_DAYS: '1.5' }), /MIN_ACCOUNT_AGE_DAYS/);
  assert.throws(() => cfg({ ISSUE_MAX_PER_RUN: '0' }), /ISSUE_MAX_PER_RUN/);
  assert.throws(() => cfg({ REQUIRE_EMAIL_SUBSCRIBED: 'maybe' }), /REQUIRE_EMAIL_SUBSCRIBED/);
  assert.equal(cfg({ REQUIRE_EMAIL_SUBSCRIBED: 'false' }).requireEmailSubscribed, false);
});

test('config: values already in the environment win over the .env file', () => {
  // dotenv never overrides a key that is already set; envFile false skips the file entirely.
  const env = { ...BASE, DRY_RUN: 'false' };
  assert.equal(loadConfig(env, { envFile: false }).dryRun, false);
  assert.equal(loadConfig(env, { envFile: path.join(ROOT_DIR, 'does-not-exist.env') }).dryRun, false);
});
