import { gql, throwIfUserErrors } from './shopify.js';
import { FIND_CAMPAIGN_CARDS, SEND_GIFT_CARD_NOTIFICATION } from './queries.js';

// giftCardCode is deliberately NOT requested: Shopify returns it only on
// creation and this tool must never see, print or store it. lastCharacters
// (the last 4 characters, also shown in the admin) is enough to identify a card.
//
// Because `customerId` is set, Shopify itself emails the customer the
// "Gift card created" notification as part of this mutation. Do NOT follow it
// with giftCardSendNotificationToCustomer: that uses the same template and the
// customer receives the same email twice (observed on balloonstore1, 2026-09-30).
export const GIFT_CARD_CREATE = `
  mutation GiftCardCreate($input: GiftCardCreateInput!) {
    giftCardCreate(input: $input) {
      giftCard { id lastCharacters }
      userErrors { field message code }
    }
  }
`;

/** Integer cents → decimal string with two places, as Shopify's Decimal scalar expects. */
export function centsToDecimal(cents) {
  if (!Number.isInteger(cents) || cents <= 0) throw new Error(`Invalid gift card amount in cents: ${cents}`);
  return (cents / 100).toFixed(2);
}

/** The marker that ties a card to a campaign. `note` is internal (customers never see it). */
export function campaignMarker(campaignId) {
  return `[campaign:${campaignId}]`;
}

export function campaignNote(note, campaignId) {
  return [note, campaignMarker(campaignId)].filter(Boolean).join(' ');
}

export function hasCampaignMarker(note, campaignId) {
  return String(note ?? '').includes(campaignMarker(campaignId));
}

export function buildCreateInput(customerId, { amountCents, currencyCode, note, expiresOn, templateSuffix }) {
  // API 2026-07: `initialValue` is deprecated in favour of `initialAmount` (MoneyInput).
  const input = { customerId, initialAmount: { amount: centsToDecimal(amountCents), currencyCode } };
  if (note) input.note = note;
  if (expiresOn) input.expiresOn = expiresOn;
  if (templateSuffix) input.templateSuffix = templateSuffix;
  return input;
}

/**
 * Create a gift card assigned to the customer; Shopify emails it to them.
 * Returns { id, last4 }. NOT idempotent: the caller must make sure it is
 * called at most once per customer (see src/issue.js).
 */
export async function createGiftCard(customerId, options) {
  const data = await gql(GIFT_CARD_CREATE, { input: buildCreateInput(customerId, options) });
  const payload = data.giftCardCreate;
  if (!payload) throw new Error('giftCardCreate returned no payload');
  throwIfUserErrors('giftCardCreate', payload.userErrors);
  if (!payload.giftCard?.id) throw new Error('giftCardCreate returned no gift card id');
  return { id: payload.giftCard.id, last4: payload.giftCard.lastCharacters ?? '' };
}

/** Shopify search syntax for the campaign-card lookup. */
export function campaignCardsQuery(sinceIso, customerId) {
  const parts = [`created_at:>='${sinceIso}'`, 'source:api_client'];
  if (customerId) parts.push(`customer_id:${String(customerId).split('/').pop()}`);
  return parts.join(' ');
}

/**
 * Every gift card this campaign created since `sinceIso` (optionally for one
 * customer). The campaign marker is matched client-side because `note` is not
 * a search field. Returns
 * [{ id, createdAt, note, templateSuffix, enabled, expiresOn, last4, amountCents, balanceCents, currencyCode, customerId }].
 * amountCents is the initial value; balanceCents === amountCents means unused.
 */
export async function findCampaignCards({ campaignId, sinceIso, customerId } = {}) {
  if (!campaignId || !sinceIso) throw new Error('findCampaignCards needs campaignId and sinceIso');
  const query = campaignCardsQuery(sinceIso, customerId);
  const cards = [];
  let after = null;
  do {
    const data = await gql(FIND_CAMPAIGN_CARDS, { query, after });
    const connection = data.giftCards;
    if (!connection) throw new Error('giftCards query returned nothing');
    for (const n of connection.nodes ?? []) {
      if (!hasCampaignMarker(n.note, campaignId)) continue;
      cards.push({
        id: n.id,
        createdAt: n.createdAt,
        note: n.note ?? '',
        templateSuffix: n.templateSuffix ?? '',
        enabled: n.enabled !== false,
        expiresOn: n.expiresOn ?? null,
        last4: n.lastCharacters ?? '',
        amountCents: Math.round(Number(n.initialValue?.amount ?? 0) * 100),
        balanceCents: Math.round(Number(n.balance?.amount ?? 0) * 100),
        currencyCode: n.initialValue?.currencyCode ?? '',
        customerId: n.customer?.id ?? null,
      });
    }
    const next = connection.pageInfo?.hasNextPage ? connection.pageInfo.endCursor : null;
    if (next && next === after) throw new Error('giftCards query returned the same cursor twice');
    after = next;
  } while (after);
  return cards;
}

/**
 * Ask Shopify to e-mail the card's "Gift card created" notification again (used
 * for reminders). NOT idempotent: every successful call sends one more email.
 * Throws ShopifyError (code USER_ERROR) on userErrors.
 */
export async function sendGiftCardNotification(giftCardId) {
  const data = await gql(SEND_GIFT_CARD_NOTIFICATION, { id: giftCardId });
  const payload = data.giftCardSendNotificationToCustomer;
  if (!payload) throw new Error('giftCardSendNotificationToCustomer returned no payload');
  throwIfUserErrors('giftCardSendNotificationToCustomer', payload.userErrors);
  return { id: payload.giftCard?.id ?? giftCardId };
}
