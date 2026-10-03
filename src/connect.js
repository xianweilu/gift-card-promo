import { getAccessToken, assertScopes, REQUIRED_SCOPES } from './auth.js';
import { initClient } from './shopify.js';

/**
 * Exchange the app credentials for a token, check the scopes and point the
 * GraphQL client at the store. Every command calls this first.
 * Never logs the secret or the token.
 */
export async function connect(config, { log = console, sleep, timeoutMs } = {}) {
  const { token, scopes } = await getAccessToken(config);
  assertScopes(scopes, REQUIRED_SCOPES);
  initClient({ shop: config.shop, apiVersion: config.apiVersion, token, log, sleep, timeoutMs });
  return { scopes };
}
