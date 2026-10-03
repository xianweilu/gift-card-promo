// Every GraphQL document the tool sends, in one place. Operation names are
// stable: the test double (test/fake-shopify.js) dispatches on them.
// Validated against the Admin GraphQL schema; run on the store's 2026-07 API.

const ADDRESS = 'address1 address2 city provinceCode zip countryCodeV2';
const ORDER = 'id name createdAt cancelledAt test sourceName totalPriceSet { shopMoney { amount currencyCode } }';
const CUSTOMER = `
  id firstName lastName displayName createdAt numberOfOrders tags
  amountSpent { amount currencyCode }
  defaultEmailAddress { emailAddress marketingState validFormat }
  defaultAddress { ${ADDRESS} }
  lastOrder { ${ORDER} }
`;

export const SHOP_INFO = `
  query ShopInfo {
    shop { ianaTimezone currencyCode }
  }
`;

/**
 * Submitted as the `query` argument of bulkOperationRunQuery. One JSONL line
 * per customer: lastOrder, defaultAddress and tags stay inline (no child rows).
 */
export const CUSTOMER_EXPORT_BULK = `
  {
    customers {
      edges {
        node {
          ${CUSTOMER}
        }
      }
    }
  }
`;

export const RUN_BULK_QUERY = `
  mutation RunCustomerExport($query: String!) {
    bulkOperationRunQuery(query: $query) {
      bulkOperation { id status }
      userErrors { field message code }
    }
  }
`;

export const BULK_STATUS = `
  query BulkStatus($id: ID!) {
    bulkOperation(id: $id) {
      id status errorCode objectCount rootObjectCount fileSize url partialDataUrl completedAt
    }
  }
`;

export const BULK_CANCEL = `
  mutation CancelBulk($id: ID!) {
    bulkOperationCancel(id: $id) {
      bulkOperation { id status }
      userErrors { field message }
    }
  }
`;

/** Fallback when the bulk job fails or stalls: same fields, cursor pagination. */
export const CUSTOMERS_PAGE = `
  query CustomersPage($after: String) {
    customers(first: 250, after: $after, sortKey: ID) {
      nodes {
        ${CUSTOMER}
      }
      pageInfo { hasNextPage endCursor }
    }
  }
`;

/** Orders placed since a date, with the buyer's default address (activity check). */
export const ORDERS_SINCE = `
  query OrdersSince($query: String!, $after: String) {
    orders(first: 250, after: $after, query: $query, sortKey: CREATED_AT) {
      nodes {
        id createdAt cancelledAt test
        customer { id defaultAddress { ${ADDRESS} } }
      }
      pageInfo { hasNextPage endCursor }
    }
  }
`;

/** Newest orders of one customer, used when the last order is cancelled, a test or $0. */
export const CUSTOMER_ORDERS = `
  query CustomerOrders($id: ID!, $after: String) {
    customer(id: $id) {
      id
      orders(first: 50, after: $after, reverse: true, sortKey: CREATED_AT) {
        nodes { ${ORDER} }
        pageInfo { hasNextPage endCursor }
      }
    }
  }
`;

/** issue pre-flight and test campaigns: re-read up to 250 customers by id. */
export const REFRESH_CUSTOMERS = `
  query RefreshCustomers($ids: [ID!]!) {
    nodes(ids: $ids) {
      ... on Customer {
        ${CUSTOMER}
      }
    }
  }
`;

/**
 * Gift cards created by this tool, with balance and expiry. The campaign marker
 * lives in `note` (not searchable, matched client-side). `balance` equal to
 * `initialValue` means the card has not been used.
 */
export const FIND_CAMPAIGN_CARDS = `
  query FindCampaignCards($query: String!, $after: String) {
    giftCards(first: 250, after: $after, query: $query, sortKey: CREATED_AT) {
      nodes {
        id createdAt note templateSuffix enabled expiresOn lastCharacters
        initialValue { amount currencyCode }
        balance { amount currencyCode }
        customer { id }
      }
      pageInfo { hasNextPage endCursor }
    }
  }
`;

/** remind: Shopify re-sends the card's "Gift card created" email. Every call sends one email. */
export const SEND_GIFT_CARD_NOTIFICATION = `
  mutation SendGiftCardNotification($id: ID!) {
    giftCardSendNotificationToCustomer(id: $id) {
      giftCard { id }
      userErrors { field message code }
    }
  }
`;

/**
 * usage: orders paid (partly) with a gift card. Each gift-card transaction's
 * receiptJson carries gift_card_id (numeric) — verified on the store 2026-10-01.
 */
export const GIFT_CARD_ORDERS = `
  query GiftCardOrders($query: String!, $after: String) {
    orders(first: 25, after: $after, query: $query, sortKey: CREATED_AT) {
      nodes {
        id name createdAt cancelledAt displayFinancialStatus
        customer { id }
        totalPriceSet { shopMoney { amount currencyCode } }
        transactions(first: 20) { gateway kind status processedAt amountSet { shopMoney { amount } } receiptJson }
        lineItems(first: 30) { nodes { name sku quantity discountedTotalSet { shopMoney { amount } } } }
      }
      pageInfo { hasNextPage endCursor }
    }
  }
`;
