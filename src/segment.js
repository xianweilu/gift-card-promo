import { gql } from './shopify.js';

const SEGMENT_MEMBER_GID = /^gid:\/\/shopify\/CustomerSegmentMember\/(\d+)$/;
const CUSTOMER_GID = /^gid:\/\/shopify\/Customer\/\d+$/;

/**
 * customerSegmentMembers returns CustomerSegmentMember nodes whose id is
 * gid://shopify/CustomerSegmentMember/<n>; the number is the customer's id.
 * giftCardCreate.customerId and tagsAdd.id only accept gid://shopify/Customer/<n>,
 * so every member id is normalised to that form before anything else sees it.
 * Customer GIDs pass through unchanged; any other shape throws so that no
 * write is ever attempted against an unexpected object.
 */
export function toCustomerGid(id) {
  const m = SEGMENT_MEMBER_GID.exec(id);
  if (m) return `gid://shopify/Customer/${m[1]}`;
  if (CUSTOMER_GID.test(id)) return id;
  throw new Error(`Unexpected segment member id "${id}" (expected gid://shopify/CustomerSegmentMember/<n>)`);
}

export const SEGMENT_MEMBERS_QUERY = `
  query SegmentMembers($id: ID!, $first: Int!, $after: String) {
    customerSegmentMembers(segmentId: $id, first: $first, after: $after) {
      totalCount
      edges {
        node {
          id
          displayName
          numberOfOrders
          lastOrderId
          defaultEmailAddress { emailAddress }
        }
      }
      pageInfo { hasNextPage endCursor }
    }
  }
`;

/**
 * Fetch EVERY member of the segment before returning. Nothing may be written
 * to Shopify while paging: tagging a member removes them from the segment in
 * real time, which would shift the cursor and skip people.
 */
export async function fetchAllMembers(segmentId, { pageSize = 250, log } = {}) {
  const byId = new Map();
  let after = null;
  let totalCount = 0;
  let pages = 0;
  do {
    const data = await gql(SEGMENT_MEMBERS_QUERY, { id: segmentId, first: pageSize, after });
    const connection = data.customerSegmentMembers;
    if (!connection) throw new Error('customerSegmentMembers returned nothing; check SEGMENT_ID');
    pages += 1;
    totalCount = connection.totalCount ?? totalCount;
    const edges = connection.edges ?? [];
    for (const edge of edges) {
      if (!edge?.node?.id) continue;
      const id = toCustomerGid(edge.node.id);
      byId.set(id, { ...edge.node, id });
    }
    log?.info(`  page ${pages}: ${edges.length} member(s), ${byId.size} unique so far`);
    const next = connection.pageInfo?.hasNextPage ? connection.pageInfo.endCursor : null;
    if (next && next === after) throw new Error('customerSegmentMembers returned the same cursor twice; aborting to avoid an endless loop');
    after = next;
  } while (after);
  return { members: [...byId.values()], totalCount, pages };
}
