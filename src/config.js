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

const REQUIRED = ['SHOP', 'CLIENT_ID', 'CLIENT_SECRET', 'SEGMENT_ID', 'SENT_TAG'];

/** Trimmed string value of env[key]; `fallback` when unset or blank. */
function str(env, key, fallback = '') {
  const raw = env[key];
  const value = raw === undefined || raw === null ? '' : String(raw).trim();
  return value || fallback;
}

/**
 * Load and validate configuration.
 *
 * Reads `.env` from the project root unless `envFile` is false. dotenv never
 * overrides variables that are already set in the environment, so
 * `DRY_RUN=false node index.js` takes precedence over the value in `.env`.
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
  if (!/^(\d{4}-\d{2}|unstable)$/.test(apiVersion)) {
    throw new ConfigError(`API_VERSION must look like 2026-07, got "${apiVersion}"`);
  }

  const segmentId = str(env, 'SEGMENT_ID');
  if (!/^gid:\/\/shopify\/Segment\/\d+$/.test(segmentId)) {
    throw new ConfigError(`SEGMENT_ID must look like gid://shopify/Segment/<number>, got "${segmentId}"`);
  }

  const pageSizeRaw = str(env, 'SEGMENT_PAGE_SIZE', '250');
  const segmentPageSize = Number(pageSizeRaw);
  if (!Number.isInteger(segmentPageSize) || segmentPageSize < 1 || segmentPageSize > 1000) {
    throw new ConfigError(`SEGMENT_PAGE_SIZE must be an integer from 1 to 1000, got "${pageSizeRaw}"`);
  }

  const valueRaw = str(env, 'GIFT_CARD_VALUE', '0.10');
  const giftCardValue = Number(valueRaw);
  if (!Number.isFinite(giftCardValue) || giftCardValue <= 0 || giftCardValue > 2000) {
    throw new ConfigError(`GIFT_CARD_VALUE must be a number greater than 0 and at most 2000, got "${valueRaw}"`);
  }

  const giftCardCurrency = str(env, 'GIFT_CARD_CURRENCY', 'USD').toUpperCase();
  if (!/^[A-Z]{3}$/.test(giftCardCurrency)) {
    throw new ConfigError(`GIFT_CARD_CURRENCY must be a 3-letter ISO code, got "${giftCardCurrency}"`);
  }

  const giftCardExpiresOn = str(env, 'GIFT_CARD_EXPIRES_ON');
  if (giftCardExpiresOn && !/^\d{4}-\d{2}-\d{2}$/.test(giftCardExpiresOn)) {
    throw new ConfigError(`GIFT_CARD_EXPIRES_ON must be YYYY-MM-DD or empty, got "${giftCardExpiresOn}"`);
  }

  const sentTag = str(env, 'SENT_TAG');
  if (sentTag.includes(',')) {
    throw new ConfigError('SENT_TAG must not contain a comma (tagsAdd would split it into several tags)');
  }

  // Anything other than the literal string "false" keeps dry-run ON.
  const dryRun = str(env, 'DRY_RUN', 'true').toLowerCase() !== 'false';

  return {
    shop,
    clientId: str(env, 'CLIENT_ID'),
    clientSecret: str(env, 'CLIENT_SECRET'),
    apiVersion,
    segmentId,
    segmentPageSize,
    giftCardValue,
    giftCardCurrency,
    giftCardNote: str(env, 'GIFT_CARD_NOTE'),
    giftCardExpiresOn,
    giftCardTemplateSuffix: str(env, 'GIFT_CARD_TEMPLATE_SUFFIX'),
    sentTag,
    dryRun,
    progressFile: path.join(ROOT_DIR, 'progress.json'),
  };
}
