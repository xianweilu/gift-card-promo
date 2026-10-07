import path from 'node:path';
import { fileURLToPath } from 'node:url';
import dotenv from 'dotenv';

export const ROOT_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

export class ConfigError extends Error {
  constructor(message) {
    super(message);
    this.name = 'ConfigError';
  }
}

const REQUIRED = ['SHOP', 'CLIENT_ID', 'CLIENT_SECRET', 'CAMPAIGN_ID', 'SENT_TAG'];

// Defaults agreed for the 2026-10 campaign. An explicitly empty value in .env
// (e.g. `EXCLUDE_TAGS=`) means "none"; leaving the key out uses the default.
export const DEFAULT_EXCLUDE_TAGS = 'WHS,PotentiallyWHS,DISC,level1A,level2A,DO NOT SELL,likely-fake-account,walmart.com,Walmart,Amazon,eBay';
// Marketplace relay/placeholder domains, plus company domains the owner excluded (2026-10).
export const DEFAULT_EXCLUDE_EMAIL_DOMAINS = 'example.com,mail.codisto.com,connectebay.com,marketplace.amazon.com,relay.walmart.com,members.ebay.com,jokerpartysupply.com,loftus.com,burtonandburton.com,mayflower.com,toyworld.com,rainbowballoons.com';
export const DEFAULT_GIFT_TIERS = '10.77,15.33,19.77';
// Order.sourceName values. The numeric ones are Shopify app ids:
// 205641 = Sellbrite, 1456995 = CedCommerce Walmart Connector, 1775805 = eBay.
export const DEFAULT_EXCLUDE_ORDER_SOURCES = 'amazon,walmart,ebay,etsy,205641,1456995,1775805';

/** Trimmed string value of env[key]; `fallback` when unset or blank. */
function str(env, key, fallback = '') {
  const raw = env[key];
  const value = raw === undefined || raw === null ? '' : String(raw).trim();
  return value || fallback;
}

/** Comma-separated list. Unset key → fallback list; set-but-empty → []. */
function list(env, key, fallback) {
  const raw = env[key] === undefined || env[key] === null ? fallback : String(env[key]);
  return raw.split(',').map((s) => s.trim()).filter(Boolean);
}

function int(env, key, fallback, { min, max }) {
  const raw = str(env, key, String(fallback));
  const n = Number(raw);
  if (!Number.isInteger(n) || n < min || n > max) {
    throw new ConfigError(`${key} must be an integer from ${min} to ${max}, got "${raw}"`);
  }
  return n;
}

function bool(env, key, fallback) {
  const raw = str(env, key, fallback ? 'true' : 'false').toLowerCase();
  if (['true', '1', 'yes'].includes(raw)) return true;
  if (['false', '0', 'no'].includes(raw)) return false;
  throw new ConfigError(`${key} must be true or false, got "${raw}"`);
}

/** Dollar amount like "20" or "0.10" → integer cents. */
function dollars(env, key, fallback, { minCents, maxCents }) {
  const raw = str(env, key, fallback);
  if (!/^\d+(\.\d{1,2})?$/.test(raw)) throw new ConfigError(`${key} must be a dollar amount like 20 or 0.10, got "${raw}"`);
  const cents = Math.round(Number(raw) * 100);
  if (cents < minCents || cents > maxCents) {
    throw new ConfigError(`${key} must be between $${(minCents / 100).toFixed(2)} and $${(maxCents / 100).toFixed(2)}, got "${raw}"`);
  }
  return cents;
}

/** "YYYY-MM-DD" (a calendar date in the store's time zone) or '' when optional. */
function dateValue(env, key, { required = false } = {}) {
  const v = str(env, key);
  if (!v) {
    if (required) throw new ConfigError(`${key} is required (format YYYY-MM-DD, e.g. 2026-10-05)`);
    return '';
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v) || Number.isNaN(Date.parse(`${v}T00:00:00Z`)) || new Date(`${v}T00:00:00Z`).toISOString().slice(0, 10) !== v) {
    throw new ConfigError(`${key} must be a real date in the format YYYY-MM-DD (e.g. 2026-10-05), got "${v}"`);
  }
  return v;
}

/** GIFT_TIERS="10.77,15.33,19.77" → [1077, 1533, 1977] (strictly increasing, positive). */
function tiers(env, key, fallback) {
  const raw = str(env, key, fallback);
  const parts = raw.split(',').map((x) => x.trim()).filter(Boolean);
  if (!parts.length) throw new ConfigError(`${key} must list at least one amount, e.g. 10.77,15.33,19.77`);
  const cents = parts.map((x) => {
    if (!/^\d+(\.\d{1,2})?$/.test(x)) throw new ConfigError(`${key} entries must be dollar amounts like 10.77, got "${x}"`);
    return Math.round(Number(x) * 100);
  });
  for (let i = 0; i < cents.length; i += 1) {
    if (cents[i] <= 0 || cents[i] > 200000) throw new ConfigError(`${key} amounts must be between $0.01 and $2,000.00, got "${parts[i]}"`);
    if (i && cents[i] <= cents[i - 1]) throw new ConfigError(`${key} must be strictly increasing, got "${raw}"`);
  }
  return cents;
}

function customerGid(raw) {
  const v = String(raw).trim();
  if (/^\d+$/.test(v)) return `gid://shopify/Customer/${v}`;
  if (/^gid:\/\/shopify\/Customer\/\d+$/.test(v)) return v;
  throw new ConfigError(`TEST_CUSTOMER_IDS entries must be customer ids, got "${v}"`);
}

/**
 * Load and validate configuration.
 *
 * Reads `.env` from the project root unless `envFile` is false. dotenv never
 * overrides variables that are already set in the environment, so
 * `DRY_RUN=false node index.js issue` takes precedence over the value in `.env`.
 */
export function loadConfig(env = process.env, { envFile = path.join(ROOT_DIR, '.env') } = {}) {
  if (envFile) dotenv.config({ path: envFile, processEnv: env, quiet: true });

  const missing = REQUIRED.filter((key) => !str(env, key));
  if (missing.length) {
    throw new ConfigError(`Missing required .env values: ${missing.join(', ')}`);
  }

  const shop = str(env, 'SHOP')
    .replace(/^https?:\/\//i, '')
    .replace(/\.myshopify\.com.*$/i, '')
    .replace(/\/.*$/, '');
  if (!/^[a-z0-9][a-z0-9-]*$/i.test(shop)) {
    throw new ConfigError(`SHOP looks wrong: "${shop}" (use the store handle, e.g. balloonstore1)`);
  }

  const apiVersion = str(env, 'API_VERSION', '2026-07');
  if (!/^\d{4}-\d{2}$/.test(apiVersion)) {
    throw new ConfigError(`API_VERSION must look like 2026-07, got "${apiVersion}" (unstable is not allowed for live campaigns)`);
  }

  const campaignId = str(env, 'CAMPAIGN_ID');
  if (!/^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$/.test(campaignId)) {
    throw new ConfigError(`CAMPAIGN_ID may only contain letters, digits, ".", "_" and "-", got "${campaignId}"`);
  }

  const sentTag = str(env, 'SENT_TAG');
  if (sentTag.includes(',')) {
    throw new ConfigError('SENT_TAG must not contain a comma (tagsAdd would split it into several tags)');
  }

  const giftPercentRaw = str(env, 'GIFT_PERCENT', '10');
  const giftPercent = Number(giftPercentRaw);
  if (!Number.isFinite(giftPercent) || giftPercent <= 0 || giftPercent > 100) {
    throw new ConfigError(`GIFT_PERCENT must be a number greater than 0 and at most 100, got "${giftPercentRaw}"`);
  }
  const giftTiersCents = tiers(env, 'GIFT_TIERS', DEFAULT_GIFT_TIERS);

  const giftCardCurrency = str(env, 'GIFT_CARD_CURRENCY', 'USD').toUpperCase();
  if (!/^[A-Z]{3}$/.test(giftCardCurrency)) {
    throw new ConfigError(`GIFT_CARD_CURRENCY must be a 3-letter ISO code, got "${giftCardCurrency}"`);
  }

  // Campaign dates, calendar days in the store's time zone. The recipients are
  // issued in two groups on two days: LAUNCH_DATE gates live issue runs for the
  // customers who have ordered (--group ordered), LAUNCH_DATE_NEVER for those
  // who never have (--group never). The expiry is the last usable day and the
  // day from which issue creates no new card. (Reminders are sent with Shopify
  // Email, so there are no reminder dates any more.)
  const giftCardExpiresOn = dateValue(env, 'GIFT_CARD_EXPIRES_ON');
  const launchDate = dateValue(env, 'LAUNCH_DATE');
  const launchDateNever = dateValue(env, 'LAUNCH_DATE_NEVER');
  if (giftCardExpiresOn && launchDate && launchDate > giftCardExpiresOn) {
    throw new ConfigError(`LAUNCH_DATE (${launchDate}) must not be later than GIFT_CARD_EXPIRES_ON (${giftCardExpiresOn})`);
  }
  if (giftCardExpiresOn && launchDateNever && launchDateNever > giftCardExpiresOn) {
    throw new ConfigError(`LAUNCH_DATE_NEVER (${launchDateNever}) must not be later than GIFT_CARD_EXPIRES_ON (${giftCardExpiresOn})`);
  }
  if (launchDate && launchDateNever && launchDateNever < launchDate) {
    throw new ConfigError(`LAUNCH_DATE_NEVER (${launchDateNever}) must not be earlier than LAUNCH_DATE (${launchDate})`);
  }

  const testCustomerIds = list(env, 'TEST_CUSTOMER_IDS', '').map(customerGid);
  if (testCustomerIds.length && !/test/i.test(campaignId)) {
    throw new ConfigError('TEST_CUSTOMER_IDS is set, so CAMPAIGN_ID must contain "test" (keeps test cards out of the real campaign)');
  }

  const timezone = str(env, 'TIMEZONE');
  if (timezone) {
    try {
      new Intl.DateTimeFormat('en-US', { timeZone: timezone });
    } catch {
      throw new ConfigError(`TIMEZONE is not a valid IANA time zone: "${timezone}"`);
    }
  }

  // Anything other than the literal string "false" keeps dry-run ON.
  const dryRun = str(env, 'DRY_RUN', 'true').toLowerCase() !== 'false';

  return {
    shop,
    clientId: str(env, 'CLIENT_ID'),
    clientSecret: str(env, 'CLIENT_SECRET'),
    apiVersion,
    campaignId,
    sentTag,
    inactiveMonths: int(env, 'INACTIVE_MONTHS', 3, { min: 1, max: 36 }),
    minAccountAgeDays: int(env, 'MIN_ACCOUNT_AGE_DAYS', 7, { min: 0, max: 3650 }),
    requireEmailSubscribed: bool(env, 'REQUIRE_EMAIL_SUBSCRIBED', true),
    excludeTags: list(env, 'EXCLUDE_TAGS', DEFAULT_EXCLUDE_TAGS),
    excludeEmailDomains: list(env, 'EXCLUDE_EMAIL_DOMAINS', DEFAULT_EXCLUDE_EMAIL_DOMAINS).map((d) => d.toLowerCase()),
    excludeOrderSources: list(env, 'EXCLUDE_ORDER_SOURCES', DEFAULT_EXCLUDE_ORDER_SOURCES).map((s) => s.toLowerCase()),
    giftPercent,
    giftTiersCents,
    giftCardCurrency,
    giftCardNote: str(env, 'GIFT_CARD_NOTE'),
    giftCardExpiresOn,
    giftCardTemplateSuffix: str(env, 'GIFT_CARD_TEMPLATE_SUFFIX'),
    launchDate,
    launchDateNever,
    issueMaxPerRun: int(env, 'ISSUE_MAX_PER_RUN', 20000, { min: 1, max: 100000 }),
    testCustomerIds,
    testGiftAmountCents: dollars(env, 'TEST_GIFT_AMOUNT', '0.10', { minCents: 1, maxCents: 2000 }),
    timezone, // empty → use the shop's ianaTimezone at select time
    dryRun,
    campaignsDir: str(env, 'CAMPAIGNS_DIR', path.join(ROOT_DIR, 'campaigns')),
  };
}
