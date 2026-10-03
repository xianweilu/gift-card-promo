const MAX_RETRIES = 5; // retries after the first attempt, for 429 / THROTTLED only
const LOW_WATER_MARK = 100; // points; pause when the cost bucket drops below this

export class ShopifyError extends Error {
  /**
   * @param {string} message
   * @param {{ outcomeKnown?: boolean, code?: string, status?: number }} [info]
   *   outcomeKnown=true  → Shopify definitely did NOT apply the request (a mutation may be retried)
   *   outcomeKnown=false → we cannot tell whether it was applied (a mutation must NOT be blindly retried)
   */
  constructor(message, { outcomeKnown = false, code, status } = {}) {
    super(message);
    this.name = 'ShopifyError';
    this.outcomeKnown = outcomeKnown;
    this.code = code;
    this.status = status;
  }
}

export const defaultSleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

let client = null;

export function initClient({ shop, apiVersion, token, log = console, sleep, timeoutMs = 30_000 }) {
  client = {
    url: `https://${shop}.myshopify.com/admin/api/${apiVersion}/graphql.json`,
    token,
    log,
    sleep: sleep ?? defaultSleep,
    timeoutMs,
  };
}

export function resetClient() {
  client = null;
}

/**
 * Run one GraphQL operation. Retries only when Shopify explicitly rejected the
 * request for rate-limiting (HTTP 429 or a THROTTLED error). Every other
 * failure is thrown as a ShopifyError whose `outcomeKnown` flag tells the
 * caller whether the request can have taken effect.
 */
export async function gql(query, variables = {}) {
  if (!client) throw new Error('Shopify client not initialised: call initClient() first');
  const { url, token, log, sleep, timeoutMs } = client;

  for (let attempt = 0; ; attempt += 1) {
    let res;
    try {
      res = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Accept: 'application/json',
          'X-Shopify-Access-Token': token,
        },
        body: JSON.stringify({ query, variables }),
        signal: AbortSignal.timeout(timeoutMs),
      });
    } catch (err) {
      throw new ShopifyError(`Network error calling Shopify: ${err.message}`, { outcomeKnown: false });
    }

    if (res.status === 429) {
      if (attempt >= MAX_RETRIES) {
        throw new ShopifyError(`HTTP 429 from Shopify ${MAX_RETRIES + 1} times in a row; giving up`, { outcomeKnown: true, status: 429 });
      }
      const retryAfter = Number(res.headers.get('retry-after')) || 2;
      log.warn(`  rate limited (HTTP 429); waiting ${retryAfter}s, retry ${attempt + 1}/${MAX_RETRIES}`);
      await sleep(retryAfter * 1000);
      continue;
    }

    const text = await res.text();
    if (!res.ok) {
      throw new ShopifyError(`Shopify HTTP ${res.status}: ${text.slice(0, 500)}`, {
        outcomeKnown: res.status < 500,
        status: res.status,
      });
    }

    let body;
    try {
      body = JSON.parse(text);
    } catch {
      throw new ShopifyError(`Shopify returned non-JSON (HTTP ${res.status}): ${text.slice(0, 200)}`, {
        outcomeKnown: false,
        status: res.status,
      });
    }

    if (body.errors && !Array.isArray(body.errors)) {
      throw new ShopifyError(`Shopify error: ${JSON.stringify(body.errors).slice(0, 500)}`, { outcomeKnown: true });
    }
    const errors = body.errors ?? [];
    const cost = body.extensions?.cost;
    const throttle = cost?.throttleStatus;

    if (errors.some((e) => e.extensions?.code === 'THROTTLED')) {
      if (attempt >= MAX_RETRIES) {
        throw new ShopifyError(`Throttled by Shopify ${MAX_RETRIES + 1} times in a row; giving up`, { outcomeKnown: true, code: 'THROTTLED' });
      }
      const shortfall = (cost?.requestedQueryCost ?? 0) - (throttle?.currentlyAvailable ?? 0);
      const waitMs = Math.max(2000, Math.ceil(shortfall / (throttle?.restoreRate || 50)) * 1000);
      log.warn(`  throttled; waiting ${waitMs / 1000}s, retry ${attempt + 1}/${MAX_RETRIES}`);
      await sleep(waitMs);
      continue;
    }

    if (errors.length) {
      const code = errors[0]?.extensions?.code;
      const messages = errors.map((e) => (e.extensions?.code ? `${e.message} [${e.extensions.code}]` : e.message));
      throw new ShopifyError(`GraphQL error: ${messages.join('; ')}`, {
        outcomeKnown: code !== 'INTERNAL_SERVER_ERROR',
        code,
      });
    }

    if (body.data === undefined || body.data === null) {
      throw new ShopifyError('Shopify response contained no data', { outcomeKnown: false });
    }

    // Shopify silently IGNORES a search filter it does not understand and
    // returns everything; the only signal is a warning here. Treat it as fatal.
    const searchWarnings = (body.extensions?.search ?? []).flatMap((s) => (s.warnings ?? []).map((w) => `${s.query ?? ''}: ${w.field ?? ''} ${w.message ?? ''}`.trim()));
    if (searchWarnings.length) {
      throw new ShopifyError(`Shopify ignored part of a search filter: ${searchWarnings.join('; ')}`, { outcomeKnown: true, code: 'SEARCH_WARNING' });
    }

    if (throttle && throttle.currentlyAvailable < LOW_WATER_MARK) {
      const waitMs = Math.ceil((LOW_WATER_MARK - throttle.currentlyAvailable) / (throttle.restoreRate || 50)) * 1000;
      log.info(`  rate-limit budget low (${throttle.currentlyAvailable} points); pausing ${waitMs / 1000}s`);
      await sleep(waitMs);
    }

    return body.data;
  }
}

/** Throw a ShopifyError (outcomeKnown: the mutation was rejected) when a payload carries userErrors. */
export function throwIfUserErrors(operation, userErrors) {
  if (!Array.isArray(userErrors) || userErrors.length === 0) return;
  const details = userErrors.map((e) => {
    const field = Array.isArray(e.field) ? e.field.join('.') : e.field;
    const code = e.code ? ` [${e.code}]` : '';
    return field ? `${field}: ${e.message}${code}` : `${e.message}${code}`;
  });
  throw new ShopifyError(`${operation} rejected: ${details.join('; ')}`, { outcomeKnown: true, code: 'USER_ERROR' });
}
