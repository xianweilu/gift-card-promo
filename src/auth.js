export const REQUIRED_SCOPES = ['write_gift_cards', 'write_customers', 'read_orders', 'read_all_orders'];

/**
 * Exchange the app's client credentials for a 24-hour Admin API access token
 * (Dev Dashboard app, client credentials grant; the body is form-encoded).
 * Never log the client secret or the returned token.
 */
export async function getAccessToken({ shop, clientId, clientSecret }) {
  const url = `https://${shop}.myshopify.com/admin/oauth/access_token`;
  let res;
  try {
    res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', Accept: 'application/json' },
      body: new URLSearchParams({ grant_type: 'client_credentials', client_id: clientId, client_secret: clientSecret }),
      signal: AbortSignal.timeout(30_000),
    });
  } catch (err) {
    throw new Error(`Token request to ${url} failed: ${err.message}`);
  }
  const text = await res.text();
  if (!res.ok) {
    throw new Error(`Token request failed: HTTP ${res.status} ${text.slice(0, 300)}`);
  }
  let json;
  try {
    json = JSON.parse(text);
  } catch {
    throw new Error('Token response is not JSON');
  }
  if (!json.access_token) throw new Error('Token response has no access_token');
  return { token: json.access_token, scopes: parseScopes(json.scope), expiresIn: json.expires_in };
}

export function parseScopes(scope) {
  return String(scope ?? '')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}

/** Scopes from `required` that `granted` does not cover. A write_X scope covers read_X. */
export function missingScopes(granted, required = REQUIRED_SCOPES) {
  const has = new Set(granted);
  return required.filter((scope) => {
    if (has.has(scope)) return false;
    if (scope.startsWith('read_') && has.has(`write_${scope.slice('read_'.length)}`)) return false;
    return true;
  });
}

export function assertScopes(granted, required = REQUIRED_SCOPES) {
  const missing = missingScopes(granted, required);
  if (missing.length) {
    throw new Error(
      `Access token is missing required scopes: ${missing.join(', ')}. Granted: ${granted.join(', ') || '(none)'}. `
        + 'Add them to the app in the Dev Dashboard, release a new version and reinstall it on the store.',
    );
  }
}
