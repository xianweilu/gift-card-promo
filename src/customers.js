import { gql, throwIfUserErrors } from './shopify.js';

// tagsAdd returns plain UserError objects: only `field` and `message` exist (no `code`).
export const TAGS_ADD = `
  mutation TagsAdd($id: ID!, $tags: [String!]!) {
    tagsAdd(id: $id, tags: $tags) {
      node { id }
      userErrors { field message }
    }
  }
`;

export const TAGGED_CUSTOMERS_QUERY = `
  query TaggedCustomers($query: String!, $first: Int!, $after: String) {
    customers(first: $first, after: $after, query: $query) {
      edges { node { id tags } }
      pageInfo { hasNextPage endCursor }
    }
  }
`;

/**
 * Shopify search syntax for "has this tag": the value is double-quoted with `\` and `"` escaped.
 * Shopify matches it as a prefix, not exactly: tag:"OCT26RTPROMO" also finds customers who only
 * carry OCT26RTPROMO-TEST7 or OCT26RTPROMO-USED (seen on the real store, 2026-10-06). Callers
 * must therefore check the tags that come back; fetchTaggedCustomerIds does.
 */
export function tagSearchQuery(tag) {
  return `tag:"${String(tag).replace(/[\\"]/g, (c) => `\\${c}`)}"`;
}

/**
 * Every customer that carries exactly `tag` (tags are case-insensitive), as a
 * Set of Customer GIDs (gid://shopify/Customer/<n>, directly comparable with
 * member ids). The search only narrows the candidates (it matches tag
 * prefixes, see tagSearchQuery); each customer's own tags decide.
 * Called once before any write so that the tag protects against duplicate
 * gift cards even when the journal is missing or was started from scratch.
 * Known limit: customers(query:) reads a search index, so a tag added moments
 * ago may not be visible yet; the journal stays the first line of defence.
 */
export async function fetchTaggedCustomerIds(tag, { pageSize = 250, log } = {}) {
  const ids = new Set();
  const query = tagSearchQuery(tag);
  const wanted = String(tag).toLowerCase();
  const hasTag = (node) => (node.tags ?? []).some((t) => String(t).toLowerCase() === wanted);
  let after = null;
  do {
    const data = await gql(TAGGED_CUSTOMERS_QUERY, { query, first: pageSize, after });
    const connection = data.customers;
    if (!connection) throw new Error('customers query returned nothing while looking up already-tagged customers');
    for (const edge of connection.edges ?? []) {
      if (edge?.node?.id && hasTag(edge.node)) ids.add(edge.node.id);
    }
    const next = connection.pageInfo?.hasNextPage ? connection.pageInfo.endCursor : null;
    if (next && next === after) throw new Error('customers query returned the same cursor twice; aborting to avoid an endless loop');
    after = next;
  } while (after);
  log?.info(`${ids.size} customer(s) already carry tag "${tag}" in Shopify; they will be skipped`);
  return ids;
}

/**
 * Add one tag (a string) or several (an array) to a customer in one call.
 * Idempotent on Shopify's side. This script never removes tags.
 */
export async function addTag(customerId, tag) {
  const tags = Array.isArray(tag) ? tag : [tag];
  const data = await gql(TAGS_ADD, { id: customerId, tags });
  const payload = data.tagsAdd;
  if (!payload) throw new Error('tagsAdd returned no payload');
  throwIfUserErrors('tagsAdd', payload.userErrors);
  return { id: payload.node?.id ?? customerId };
}
