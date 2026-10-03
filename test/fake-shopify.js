/**
 * In-process stand-in for the Shopify endpoints this tool uses, installed as
 * globalThis.fetch. It dispatches on the GraphQL operation name (see
 * src/queries.js), keeps mutable state (customers, gift cards, tags) and
 * records every call so tests can assert exactly what was sent.
 *
 * failures: { [operationName]: [ failure, ... ] } — each entry is consumed once, in order.
 *   { kind: 'throttle' }                 THROTTLED error (request not applied)
 *   { kind: 'network' }                  fetch throws before Shopify sees anything
 *   { kind: 'http', status: 502 }        HTTP error response
 *   { kind: 'userError', message, code } mutation returns userErrors
 *   { kind: 'appliedThenLost' }          the mutation IS applied, then the response is lost (network error)
 *   { kind: 'graphqlError', message, code }
 *
 * Gift cards carry `balance` (starts equal to `initialValue`); use spendCard(id, cents)
 * to simulate a purchase. giftCardOrders feeds the usage report (GiftCardOrders).
 */

export const BULK_URL = 'https://fake-bulk.test/export.jsonl';
const ISO = (ms) => new Date(ms).toISOString();

/** Payload field of each mutation, used by the `userError` failure kind. */
const USER_ERROR_PAYLOAD = {
  GiftCardCreate: 'giftCardCreate',
  TagsAdd: 'tagsAdd',
  SendGiftCardNotification: 'giftCardSendNotificationToCustomer',
  RunCustomerExport: 'bulkOperationRunQuery',
  CancelBulk: 'bulkOperationCancel',
};

function json(status, body) {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
}

function opName(query) {
  return /^\s*(?:query|mutation)\s+(\w+)/.exec(query)?.[1] ?? '?';
}

function page(items, after, size) {
  const start = after ? Number(String(after).replace('cur-', '')) : 0;
  const slice = items.slice(start, start + size);
  const end = start + slice.length;
  const hasNextPage = end < items.length;
  return { slice, pageInfo: { hasNextPage, endCursor: hasNextPage ? `cur-${end}` : null } };
}

export function installFakeShopify({
  customers = [],
  orders = [],
  orderHistory = {},
  giftCards = [],
  giftCardOrders = [], // GiftCardOrders nodes (orders with gift-card transactions + receipts + line items)
  scope = 'write_gift_cards,write_customers,read_orders,read_all_orders',
  shop = { ianaTimezone: 'America/Los_Angeles', currencyCode: 'USD' },
  bulk = { mode: 'ok', pollsBeforeComplete: 1 },
  failures = {},
  now = () => Date.now(),
  searchWarnings = {},
} = {}) {
  const state = {
    customers: customers.map((c) => structuredClone(c)),
    orders: orders.map((o) => structuredClone(o)),
    orderHistory: structuredClone(orderHistory),
    giftCards: giftCards.map((g) => structuredClone(g)),
    giftCardOrders: giftCardOrders.map((o) => structuredClone(o)),
    hiddenCardIds: new Set(), // simulate search-index lag: hidden cards exist but are not found by FindCampaignCards
    nextGiftCardId: 1000,
    shop,
    scope,
    bulk: { mode: 'ok', pollsBeforeComplete: 1, ...bulk, polls: 0, submitted: 0, cancelled: 0 },
    failures,
    searchWarnings,
    calls: { token: [], ops: [], create: [], tag: [], notify: [], bulkDownloads: 0 },
  };
  const realFetch = globalThis.fetch;

  const customerById = (id) => state.customers.find((c) => c.id === id) ?? null;

  function graphql(data, errors, extensions) {
    const body = { data, extensions: { cost: { requestedQueryCost: 10, actualQueryCost: 10, throttleStatus: { maximumAvailable: 20000, currentlyAvailable: 19990, restoreRate: 1000 } }, ...extensions } };
    if (errors) body.errors = errors;
    return json(200, body);
  }

  function failureResponse(f) {
    switch (f.kind) {
      case 'throttle':
        return graphql(null, [{ message: 'Throttled', extensions: { code: 'THROTTLED' } }], { cost: { requestedQueryCost: 10, actualQueryCost: null, throttleStatus: { maximumAvailable: 20000, currentlyAvailable: 5, restoreRate: 1000 } } });
      case 'network':
        throw new TypeError('fetch failed');
      case 'http':
        return new Response(f.body ?? 'error', { status: f.status ?? 502 });
      case 'graphqlError':
        return graphql(null, [{ message: f.message ?? 'boom', extensions: f.code ? { code: f.code } : undefined }]);
      default:
        return null;
    }
  }

  function createCard(input) {
    state.nextGiftCardId += 1;
    const n = state.nextGiftCardId;
    const card = {
      id: `gid://shopify/GiftCard/${n}`,
      createdAt: ISO(now()),
      note: input.note ?? '',
      templateSuffix: input.templateSuffix ?? null,
      enabled: true,
      lastCharacters: `x${String(n).slice(-3)}`,
      initialValue: { amount: input.initialAmount.amount, currencyCode: input.initialAmount.currencyCode },
      balance: { amount: input.initialAmount.amount, currencyCode: input.initialAmount.currencyCode },
      customer: input.customerId ? { id: input.customerId } : null,
      expiresOn: input.expiresOn ?? null,
    };
    state.giftCards.push(card);
    return card;
  }

  function handle(op, variables) {
    switch (op) {
      case 'ShopInfo':
        return graphql({ shop: state.shop });

      case 'RunCustomerExport': {
        state.bulk.submitted += 1;
        state.bulk.polls = 0;
        state.bulk.query = variables.query;
        return graphql({ bulkOperationRunQuery: { bulkOperation: { id: `gid://shopify/BulkOperation/${state.bulk.submitted}`, status: 'CREATED' }, userErrors: [] } });
      }
      case 'BulkStatus': {
        state.bulk.polls += 1;
        const base = { id: variables.id, objectCount: '0', rootObjectCount: '0', fileSize: null, url: null, partialDataUrl: null, completedAt: null, errorCode: null };
        if (state.bulk.mode === 'fail') return graphql({ bulkOperation: { ...base, status: 'FAILED', errorCode: 'INTERNAL_SERVER_ERROR' } });
        if (state.bulk.mode === 'stall' || state.bulk.polls < state.bulk.pollsBeforeComplete) return graphql({ bulkOperation: { ...base, status: 'RUNNING', objectCount: '100' } });
        const count = String(state.customers.length);
        return graphql({ bulkOperation: { ...base, status: 'COMPLETED', objectCount: count, rootObjectCount: count, fileSize: '1000', url: state.customers.length ? BULK_URL : null, completedAt: ISO(now()) } });
      }
      case 'CancelBulk':
        state.bulk.cancelled += 1;
        return graphql({ bulkOperationCancel: { bulkOperation: { id: variables.id, status: 'CANCELING' }, userErrors: [] } });

      case 'CustomersPage': {
        const { slice, pageInfo } = page(state.customers, variables.after, 250);
        return graphql({ customers: { nodes: slice, pageInfo } });
      }
      case 'OrdersSince': {
        const since = /created_at:>='([^']+)'/.exec(variables.query)?.[1];
        const sinceMs = since ? Date.parse(since) : 0;
        const matching = state.orders.filter((o) => Date.parse(o.createdAt) >= sinceMs);
        const { slice, pageInfo } = page(matching, variables.after, 250);
        const ext = state.searchWarnings.OrdersSince ? { search: [{ path: ['orders'], query: variables.query, warnings: [{ field: 'x', message: 'Invalid search field for this query.', code: 'invalid_field' }] }] } : undefined;
        return graphql({ orders: { nodes: slice, pageInfo } }, null, ext);
      }
      case 'CustomerOrders': {
        const list = state.orderHistory[variables.id] ?? [];
        const { slice, pageInfo } = page(list, variables.after, 50);
        return graphql({ customer: customerById(variables.id) ? { id: variables.id, orders: { nodes: slice, pageInfo } } : null });
      }
      case 'RefreshCustomers':
        return graphql({ nodes: variables.ids.map((id) => customerById(id)) });

      case 'GiftCardCreate': {
        state.calls.create.push(structuredClone(variables.input));
        const card = createCard(variables.input);
        return graphql({ giftCardCreate: { giftCard: { id: card.id, lastCharacters: card.lastCharacters }, userErrors: [] } });
      }
      case 'TagsAdd': {
        state.calls.tag.push({ id: variables.id, tags: variables.tags });
        const c = customerById(variables.id);
        if (c) for (const t of variables.tags) if (!c.tags.some((x) => x.toLowerCase() === t.toLowerCase())) c.tags.push(t);
        return graphql({ tagsAdd: { node: c ? { id: c.id } : null, userErrors: c ? [] : [{ field: ['id'], message: 'Customer not found' }] } });
      }
      case 'TaggedCustomers': {
        const tag = /tag:"((?:[^"\\]|\\.)*)"/.exec(variables.query)?.[1]?.replace(/\\(.)/g, '$1') ?? '';
        const matching = state.customers.filter((c) => c.tags.some((t) => t.toLowerCase() === tag.toLowerCase()));
        const { slice, pageInfo } = page(matching, variables.after, variables.first ?? 250);
        return graphql({ customers: { edges: slice.map((c) => ({ node: { id: c.id } })), pageInfo } });
      }
      case 'FindCampaignCards': {
        const since = /created_at:>='([^']+)'/.exec(variables.query)?.[1];
        const customerNumeric = /customer_id:(\d+)/.exec(variables.query)?.[1];
        const sinceMs = since ? Date.parse(since) : 0;
        const matching = state.giftCards.filter((g) => !state.hiddenCardIds.has(g.id)
          && Date.parse(g.createdAt) >= sinceMs
          && (!customerNumeric || g.customer?.id === `gid://shopify/Customer/${customerNumeric}`));
        const { slice, pageInfo } = page(matching, variables.after, 250);
        return graphql({ giftCards: { nodes: slice, pageInfo } });
      }
      case 'SendGiftCardNotification': {
        state.calls.notify.push(variables.id);
        const card = state.giftCards.find((g) => g.id === variables.id);
        if (!card) return graphql({ giftCardSendNotificationToCustomer: { giftCard: null, userErrors: [{ field: ['id'], message: 'Gift card not found', code: 'GIFT_CARD_NOT_FOUND' }] } });
        if (!card.customer) return graphql({ giftCardSendNotificationToCustomer: { giftCard: null, userErrors: [{ field: ['id'], message: 'The gift card has no customer', code: 'INVALID' }] } });
        return graphql({ giftCardSendNotificationToCustomer: { giftCard: { id: card.id }, userErrors: [] } });
      }
      case 'GiftCardOrders': {
        const since = /created_at:>='([^']+)'/.exec(variables.query)?.[1];
        const sinceMs = since ? Date.parse(since) : 0;
        const matching = state.giftCardOrders.filter((o) => Date.parse(o.createdAt) >= sinceMs);
        const { slice, pageInfo } = page(matching, variables.after, 25);
        return graphql({ orders: { nodes: slice, pageInfo } });
      }
      default:
        return graphql(null, [{ message: `fake: unknown operation ${op}` }]);
    }
  }

  /** Test helper: spend `cents` from a card (simulates a customer using it at checkout). */
  function spendCard(cardId, cents) {
    const card = state.giftCards.find((g) => g.id === cardId);
    if (!card) throw new Error(`fake: no card ${cardId}`);
    const bal = Math.round(Number(card.balance.amount) * 100) - cents;
    card.balance = { ...card.balance, amount: (Math.max(0, bal) / 100).toFixed(2) };
    return card;
  }

  globalThis.fetch = async (url, init = {}) => {
    const u = String(url);
    if (u.endsWith('/admin/oauth/access_token')) {
      state.calls.token.push(Object.fromEntries(new URLSearchParams(String(init.body))));
      return json(200, { access_token: 'fake-token', scope: state.scope, expires_in: 86399 });
    }
    if (u === BULK_URL) {
      state.calls.bulkDownloads += 1;
      const body = state.customers.map((c) => JSON.stringify(c)).join('\n') + (state.customers.length ? '\n' : '');
      return new Response(body, { status: 200, headers: { 'content-type': 'application/jsonl' } });
    }
    if (u.endsWith('/graphql.json')) {
      const { query, variables = {} } = JSON.parse(init.body);
      const op = opName(query);
      state.calls.ops.push({ op, variables: structuredClone(variables) });
      const f = state.failures[op]?.shift();
      if (f) {
        if (f.kind === 'userError') {
          if (op === 'GiftCardCreate') state.calls.create.push(structuredClone(variables.input));
          if (op === 'SendGiftCardNotification') state.calls.notify.push(variables.id);
          const payloadKey = USER_ERROR_PAYLOAD[op] ?? 'payload';
          return graphql({ [payloadKey]: { giftCard: null, node: null, bulkOperation: null, userErrors: [{ field: ['input'], message: f.message ?? 'rejected', code: f.code ?? 'INVALID' }] } });
        }
        if (f.kind === 'appliedThenLost') {
          handle(op, variables); // Shopify did the work...
          throw new TypeError('fetch failed: socket hang up'); // ...but the answer never arrived
        }
        const r = failureResponse(f);
        if (r) return r;
      }
      return handle(op, variables);
    }
    throw new Error(`fake fetch: unexpected URL ${u}`);
  };

  return {
    state,
    spendCard,
    opsNamed: (name) => state.calls.ops.filter((o) => o.op === name),
    restore() {
      globalThis.fetch = realFetch;
    },
  };
}
