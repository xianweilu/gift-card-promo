import { gql, throwIfUserErrors } from './shopify.js';

// giftCardCode is deliberately NOT requested: Shopify returns it only on
// creation and this script must never see, print or store it.
//
// Because `customerId` is set, Shopify itself emails the customer the
// "Gift card created" notification as part of this mutation. Do NOT follow it
// with giftCardSendNotificationToCustomer: that uses the same template and the
// customer receives the same email twice (observed on balloonstore1, 2026-09-30).
export const GIFT_CARD_CREATE = `
  mutation GiftCardCreate($input: GiftCardCreateInput!) {
    giftCardCreate(input: $input) {
      giftCard { id }
      userErrors { field message code }
    }
  }
`;

/** Decimal string with two places, as Shopify's Decimal scalar expects. */
export function formatAmount(value) {
  const n = Number(value);
  if (!Number.isFinite(n) || n <= 0) throw new Error(`Invalid gift card amount: ${value}`);
  return n.toFixed(2);
}

export function buildCreateInput(customerId, { amount, currencyCode, note, expiresOn, templateSuffix }) {
  // API 2026-07: `initialValue` is deprecated in favour of `initialAmount` (MoneyInput).
  const input = { customerId, initialAmount: { amount: formatAmount(amount), currencyCode } };
  if (note) input.note = note;
  if (expiresOn) input.expiresOn = expiresOn;
  if (templateSuffix) input.templateSuffix = templateSuffix;
  return input;
}

/**
 * Create a gift card assigned to the customer; Shopify emails it to them. Returns { id }.
 * NOT idempotent: the caller must make sure it is called at most once per customer.
 */
export async function createGiftCard(customerId, options) {
  const data = await gql(GIFT_CARD_CREATE, { input: buildCreateInput(customerId, options) });
  const payload = data.giftCardCreate;
  if (!payload) throw new Error('giftCardCreate returned no payload');
  throwIfUserErrors('giftCardCreate', payload.userErrors);
  if (!payload.giftCard?.id) throw new Error('giftCardCreate returned no gift card id');
  return { id: payload.giftCard.id };
}
