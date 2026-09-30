/**
 * A tiny in-process stand-in for the two Shopify endpoints the script uses,
 * installed as globalThis.fetch. It records every call so tests can assert on
 * exactly which mutations were sent, and can inject failures per operation.
 *
 * failures: { [operationName]: [ { kind: 'throttle' | 'network' | 'http' | 'data' | 'graphqlError', ... }, ... ] }
 * Each entry is consumed once, in order, before the normal handler runs.
 */
export function installFakeShopify({
  pages = [[]],
  tagged = [], // Customer GIDs that already carry the sent tag in "Shopify"
  scope = 'write_gift_cards,write_customers,read_orders,read_all_orders',
  failures = {},
  throttleStatus = { maximumAvailable: 2000, currentlyAvailable: 1990, restoreRate: 100 },
} = {}) {
  const state = {
    scope,
    pages,
    tagged: [...tagged],
    failures,
    throttleStatus,
    nextGiftCardId: 1,
    calls: { token: [], queries: [], create: [], tag: [] },
  };
  const realFetch = globalThis.fetch;

  globalThis.fetch = async (url, init = {}) => {
    const u = String(url);
    if (u.endsWith('/admin/oauth/access_token')) {
      state.calls.token.push({
        contentType: init.headers?.['Content-Type'],
        body: Object.fromEntries(new URLSearchParams(String(init.body))),
      });
      return json(200, { access_token: 'fake-token', scope: state.scope, expires_in: 86399 });
    }
    if (u.endsWith('/graphql.json')) {
      const { query, variables } = JSON.parse(init.body);
      const op = /^\s*(?:query|mutation)\s+(\w+)/.exec(query)?.[1] ?? '?';
      state.calls.queries.push({ op, variables, token: init.headers?.['X-Shopify-Access-Token'] });
      const failure = state.failures[op]?.shift();
      if (failure) return fail(failure, state);
      return handle(op, variables, state);
    }
    throw new Error(`fake fetch: unexpected URL ${u}`);
  };

  return {
    state,
    restore() {
      globalThis.fetch = realFetch;
    },
  };
}

function handle(op, variables, state) {
  switch (op) {
    case 'SegmentMembers': {
      const index = variables.after ? Number(variables.after.replace('cursor-', '')) : 0;
      const page = state.pages[index] ?? [];
      const hasNextPage = index + 1 < state.pages.length;
      // The real API returns CustomerSegmentMember ids, never Customer ids.
      return graphql(state, {
        customerSegmentMembers: {
          totalCount: state.pages.flat().length,
          edges: page.map((node) => ({ node: { ...node, id: node.id.replace('gid://shopify/Customer/', 'gid://shopify/CustomerSegmentMember/') } })),
          pageInfo: { hasNextPage, endCursor: hasNextPage ? `cursor-${index + 1}` : null },
        },
      });
    }
    case 'TaggedCustomers':
      // Single page: the real API paginates, but tests never need more than one page.
      return graphql(state, {
        customers: {
          edges: state.tagged.map((id) => ({ node: { id } })),
          pageInfo: { hasNextPage: false, endCursor: null },
        },
      });
    case 'GiftCardCreate': {
      state.calls.create.push(variables.input);
      const id = `gid://shopify/GiftCard/${state.nextGiftCardId}`;
      state.nextGiftCardId += 1;
      return graphql(state, { giftCardCreate: { giftCard: { id }, userErrors: [] } });
    }
    case 'TagsAdd':
      state.calls.tag.push({ id: variables.id, tags: variables.tags });
      if (!state.tagged.includes(variables.id)) state.tagged.push(variables.id);
      return graphql(state, { tagsAdd: { node: { id: variables.id }, userErrors: [] } });
    default:
      // Deliberately no GiftCardSend handler: Shopify emails the card on creation,
      // so any explicit send call is a bug and must fail the test.
      return graphql(state, null, [{ message: `fake: unknown operation ${op}` }]);
  }
}

function fail(f, state) {
  switch (f.kind) {
    case 'throttle':
      return graphql(
        state,
        null,
        [{ message: 'Throttled', extensions: { code: 'THROTTLED', documentation: 'https://shopify.dev/api/usage/rate-limits' } }],
        { requestedQueryCost: 10, actualQueryCost: null, throttleStatus: { ...state.throttleStatus, currentlyAvailable: 5 } },
      );
    case 'network':
      throw new TypeError('fetch failed');
    case 'http':
      return new Response(f.body ?? 'error', { status: f.status ?? 500, headers: f.headers ?? {} });
    case 'data':
      return graphql(state, f.data);
    case 'graphqlError':
      return graphql(state, null, [{ message: f.message, extensions: f.code ? { code: f.code } : undefined }]);
    default:
      throw new Error(`fake: unknown failure kind ${f.kind}`);
  }
}

function graphql(state, data, errors, cost) {
  const body = {
    data,
    extensions: { cost: cost ?? { requestedQueryCost: 10, actualQueryCost: 10, throttleStatus: state.throttleStatus } },
  };
  if (errors) body.errors = errors;
  return json(200, body);
}

function json(status, body) {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
}
