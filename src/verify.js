// `verify`: read-only reconciliation of the local journal against Shopify.
//
// Lists every gift card this campaign created (its note carries
// [campaign:<CAMPAIGN_ID>]) and every customer carrying SENT_TAG, compares both
// with selection.json and the journal, and writes
//   verify.json  the issues found (the workbook's "核对" sheet); every issue carries its
//                texts in Chinese (journal / shopify / action) and in English
//                (journalEn / shopifyEn / actionEn, for the English workbook)
//   tags.json    the tag snapshot (the workbook's "是否有 <SENT_TAG>" column)
//
// Shopify is never written to and no card is ever deactivated. The only local
// changes are journal entries recording what Shopify proves:
//   reconcile.found  a recipient holds a campaign card the journal does not
//                    know about, or an in_progress/unknown row's card exists
//   reconcile.none   an in_progress/unknown row's card is still absent at least
//                    `unknownSettleMs` after create.start (search-index lag ruled
//                    out), so the row goes back to pending and may be retried
//                    (for the real campaign only before REMIND_1_DATE: from that
//                    store date on, issue creates no new card)
// verify holds the run lock, so no issue/remind can append to the journal or
// leave a live in_progress row while it runs.

import fs from 'node:fs';
import {
  STATUS,
  acquireRunLock,
  appendJournal,
  campaignPaths,
  foldJournal,
  newRunId,
  readJournal,
  readJson,
  writeJsonAtomic,
} from './campaign.js';
import { connect } from './connect.js';
import { campaignMarker, findCampaignCards } from './giftcards.js';
import { formatUsd } from './select/amount.js';
import { localDate, localDateTime } from './time.js';
import { compareGids, fetchTagSnapshot, missingSelectionMessage, refreshReport } from './export-command.js';

export const VERIFY_VERSION = 1;
export const DEFAULT_UNKNOWN_SETTLE_MS = 600_000;
/** Campaign cards are searched from one hour before the list was made (no card can be older). */
const SEARCH_MARGIN_MS = 3_600_000;
/** Issues printed one by one in the console; the workbook lists them all. */
const MAX_LISTED = 20;

/** Issue types in report order (needs attention first; already-fixed last) with their console labels. */
export const ISSUE_TYPES = Object.freeze([
  { type: 'still-unknown', label: '建卡结果仍不明' },
  { type: 'duplicate-cards', label: '同一客户多张卡' },
  { type: 'amount-mismatch', label: '金额不一致' },
  { type: 'missing-in-shopify', label: '日志有但 Shopify 没有' },
  { type: 'not-in-selection', label: '卡的客户不在名单里' },
  { type: 'tag-missing', label: '日志记录已打 tag，Shopify 上没有' },
  { type: 'tag-without-card', label: '带 tag 但没有卡' },
  { type: 'card-disabled', label: '卡已停用' },
  { type: 'not-in-journal', label: 'Shopify 有但日志没有（已补记）' },
  { type: 'resolved-unknown', label: '结果不明已查清（已补记）' },
]);
const TYPE_RANK = new Map(ISSUE_TYPES.map(({ type }, i) => [type, i]));
const TYPE_LABEL = new Map(ISSUE_TYPES.map(({ type, label }) => [type, label]));
/** Types whose journal side verify fixed itself (they need no action). */
const FIXED_TYPES = new Set(['not-in-journal', 'resolved-unknown']);

const STATUS_LABEL = {
  [STATUS.PENDING]: '待发放',
  [STATUS.IN_PROGRESS]: '进行中',
  [STATUS.UNKNOWN]: '需人工核对',
  [STATUS.CREATED]: '已建卡未打tag',
  [STATUS.DONE]: '已完成',
  [STATUS.FAILED]: '失败',
  [STATUS.SKIPPED]: '发放前跳过',
};

const STATUS_LABEL_EN = {
  [STATUS.PENDING]: 'Pending',
  [STATUS.IN_PROGRESS]: 'In progress',
  [STATUS.UNKNOWN]: 'Needs review',
  [STATUS.CREATED]: 'Card created, not tagged',
  [STATUS.DONE]: 'Done',
  [STATUS.FAILED]: 'Failed',
  [STATUS.SKIPPED]: 'Skipped before issuing',
};

const SKIP_LABEL = {
  'customer-deleted': '客户已删除',
  'no-email': '没有邮箱',
  'relay-email': '邮箱域名在排除名单里',
  'not-subscribed': '已退订营销邮件',
  'already-tagged': '已经带发放 tag',
  'ordered-since-snapshot': '导出名单后下过单',
  'address-ordered-since-snapshot': '同地址有账户在导出名单后下过单',
};

const SKIP_LABEL_EN = {
  'customer-deleted': 'customer deleted',
  'no-email': 'no email address',
  'relay-email': 'email domain is on the exclusion list',
  'not-subscribed': 'unsubscribed from marketing emails',
  'already-tagged': 'already carries the campaign tag',
  'ordered-since-snapshot': 'ordered after the list was exported',
  'address-ordered-since-snapshot': 'an account at the same address ordered after the list was exported',
};

// ---------------------------------------------------------------------------
// Text helpers: short sentences for verify.json / the "核对" sheet, each in both
// languages. Every helper returns a pair { zh, en }: `zh` is the Chinese text
// (stored as journal / shopify / action, unchanged wording) and `en` the same
// facts in operational English with ASCII punctuation (journalEn / shopifyEn /
// actionEn, shown by the English workbook).
// ---------------------------------------------------------------------------

const tail = (gid) => (gid ? String(gid).split('/').pop() : '-');

/** A bilingual text. */
const bi = (zh, en) => ({ zh, en });

/** English name of the store time zone used in the texts ("Los Angeles time"). */
function zoneNameEn(tz) {
  return tz === 'America/Los_Angeles' ? 'Los Angeles time' : `${tz} time`;
}

/** Store-local "YYYY-MM-DD HH:MM" of `iso` (English: with the zone's name), or "unknown time". */
function when(iso, tz) {
  const ms = Date.parse(iso ?? '');
  if (!Number.isFinite(ms)) return bi('时间不明', 'unknown time');
  const local = localDateTime(ms, tz);
  return bi(local, `${local} ${zoneNameEn(tz)}`);
}

/** "10 分钟" / "30 秒" and "10 minutes" / "30 seconds", rounded up. */
function durationText(ms) {
  if (ms >= 60_000) {
    const n = Math.ceil(ms / 60_000);
    return bi(`${n} 分钟`, `${n} minute${n === 1 ? '' : 's'}`);
  }
  const n = Math.max(1, Math.ceil(ms / 1000));
  return bi(`${n} 秒`, `${n} second${n === 1 ? '' : 's'}`);
}

function money(cents, currency) {
  const text = formatUsd(cents);
  return currency && currency !== 'USD' ? `${text} ${currency}` : text;
}

function clip(text, max = 120) {
  const s = String(text ?? '').replace(/\s+/g, ' ').trim();
  return s.length > max ? `${s.slice(0, max)}…` : s;
}

/** One line about a customer's journal state. */
function describeJournal(js, tz) {
  if (!js) return bi('日志里没有这位客户的记录', 'No journal record for this customer');
  const label = STATUS_LABEL[js.status] ?? js.status;
  const labelEn = STATUS_LABEL_EN[js.status] ?? js.status;
  switch (js.status) {
    case STATUS.PENDING:
      return js.attempts
        ? bi(`${label}：之前尝试过 ${js.attempts} 次，确认没有建成`, `${labelEn}: tried ${js.attempts} time${js.attempts === 1 ? '' : 's'} before, confirmed not created`)
        : bi(`${label}：还没有建卡记录`, `${labelEn}: no card creation recorded yet`);
    case STATUS.IN_PROGRESS: {
      const at = when(js.startedAt, tz);
      return bi(`${label}：${at.zh} 开始建卡，没有记下结果`, `${labelEn}: card creation started ${at.en}, no result recorded`);
    }
    case STATUS.UNKNOWN: {
      const at = when(js.startedAt, tz);
      return bi(`${label}：${at.zh} 开始建卡，结果不明`, `${labelEn}: card creation started ${at.en}, outcome unknown`);
    }
    case STATUS.CREATED: {
      const at = when(js.createdAt, tz);
      return bi(`${label}：卡尾号 ${js.last4 || '?'}，${at.zh} 建卡`, `${labelEn}: last 4 ${js.last4 || '?'}, created ${at.en}`);
    }
    case STATUS.DONE: {
      const at = when(js.taggedAt, tz);
      return bi(`${label}：卡尾号 ${js.last4 || '?'}，${at.zh} 打 tag`, `${labelEn}: last 4 ${js.last4 || '?'}, tagged ${at.en}`);
    }
    case STATUS.FAILED: {
      const error = clip(js.error);
      return bi(`${label}：${error || 'Shopify 拒绝建卡'}`, `${labelEn}: ${error || 'Shopify refused to create the card'}`);
    }
    case STATUS.SKIPPED:
      return bi(
        `${label}：${SKIP_LABEL[js.skipReason] ?? js.skipReason ?? '原因不明'}`,
        `${labelEn}: ${SKIP_LABEL_EN[js.skipReason] ?? js.skipReason ?? 'reason unknown'}`,
      );
    default:
      return bi(label, labelEn);
  }
}

/** One line about a Shopify card. Never includes the code: only the last characters. */
function describeCard(card, tz, { mentionDisabled = true } = {}) {
  const last4 = card.last4 || '?';
  const amount = money(card.amountCents, card.currencyCode);
  const created = when(card.createdAt, tz);
  const parts = [`尾号 ${last4}`, `原金额 ${amount}`];
  const partsEn = [`last 4 ${last4}`, `initial ${amount}`];
  if (card.balanceCents !== card.amountCents) {
    const balance = money(card.balanceCents, card.currencyCode);
    parts.push(`余额 ${balance}`);
    partsEn.push(`balance ${balance}`);
  }
  parts.push(`${created.zh} 建卡`);
  partsEn.push(`created ${created.en}`);
  if (!card.enabled && mentionDisabled) {
    parts.push('已停用');
    partsEn.push('disabled');
  }
  return bi(parts.join('，'), partsEn.join(', '));
}

/** Action of every row whose journal side verify fixed from Shopify's data. */
const RECORDED_ACTION = bi('已按 Shopify 补记到本地日志', 'Recorded in the local journal from Shopify');

/**
 * Action of a row verify puts back to pending (reconcile.none). `noNewCardsSince` is
 * REMIND_1_DATE once the store date has reached it for the real campaign: issue then
 * only repairs and never creates a card (its creation email would show reminder copy).
 */
function settledNoneAction(noNewCardsSince) {
  if (!noNewCardsSince) return bi('已在本地日志改回待发放，可以用 issue 重试', 'Set back to Pending in the local journal; issue can retry it');
  return bi(
    `已在本地日志改回待发放；正式活动从 REMIND_1_DATE（${noNewCardsSince}）起 issue 不再建新卡，确需补发请先改提醒日期并重贴模板`,
    `Set back to Pending in the local journal; for the live campaign issue creates no new cards from REMIND_1_DATE (${noNewCardsSince}) on. To issue it anyway, move the reminder date first and re-paste the templates`,
  );
}

/** Appended to the Shopify text of a recorded card whose customer does not carry the tag yet. */
function untaggedNote(tagged, tag) {
  return tagged ? bi('', '') : bi(`；客户还没有 ${tag}，下次运行 issue 会补打`, `; the customer does not carry ${tag} yet, the next issue run will add it`);
}

function duplicateAction(card, keeperLast4) {
  if (!card.enabled) return bi('已停用，不用再处理', 'Already disabled, nothing to do');
  if (card.balanceCents < card.amountCents) return bi('这张多出来的卡已经被用过，请人工决定怎么处理', 'This extra card has already been used; decide manually how to handle it');
  return bi(`核实后在后台停用这张多出来的卡（保留尾号 ${keeperLast4} 那张）`, `Disable this extra card in the admin after checking (keep the one ending ${keeperLast4})`);
}

function compareCards(a, b) {
  return (Date.parse(a.createdAt) || 0) - (Date.parse(b.createdAt) || 0) || compareGids(a.id, b.id);
}

function foundFix(cid, card) {
  return { op: 'reconcile.found', cid, giftCardId: card.id, last4: card.last4, amountCents: card.amountCents, createdAt: card.createdAt, source: 'verify' };
}

// ---------------------------------------------------------------------------
// The comparison (pure)
// ---------------------------------------------------------------------------

/**
 * Compare the list + journal with what Shopify holds. Pure: no I/O.
 *
 * @param {object} a
 * @param {object} a.selection selection.json
 * @param {{ customers: Map<string, object> }} a.state foldJournal() result
 * @param {object[]} a.cards findCampaignCards() result (campaign marker only)
 * @param {Set<string>} a.taggedIds customers carrying SENT_TAG in Shopify
 * @param {number} a.nowMs when the card search started (settle-time reference)
 * @param {number} [a.unknownSettleMs] an unknown row is settled this long after create.start
 * @param {string} a.tag SENT_TAG (for the texts)
 * @param {string} [a.timezone] store time zone for the texts
 * @param {string} [a.remind1Date] REMIND_1_DATE of the current config ('' = none): from that store
 *   date on, issue creates no new card for the real campaign (same rule as src/issue.js)
 * @returns {{ issues: object[], fixes: object[], counts: Record<string, number> }}
 *   issues: [{ type, customerId, giftCardId, journal, shopify, action, journalEn, shopifyEn, actionEn }]
 *   in report order (the *En fields say the same in English, for the English workbook);
 *   fixes: journal entries to append (without `run`), in customer order.
 */
export function analyzeVerify({ selection, state, cards, taggedIds, nowMs, unknownSettleMs = DEFAULT_UNKNOWN_SETTLE_MS, tag, timezone = 'UTC', remind1Date = '' }) {
  const tz = timezone;
  const currency = selection.params?.currency || '';
  // Test campaigns are not bound to the campaign dates (issue creates their cards any day).
  const noNewCardsSince = selection.mode !== 'test' && remind1Date && localDate(nowMs, tz) >= remind1Date ? remind1Date : '';
  const recipients = new Map(selection.recipients.map((r) => [r.customerId, r]));
  const journal = state.customers;
  const sortedCards = [...cards].sort(compareCards);
  const cardById = new Map(sortedCards.map((c) => [c.id, c]));
  const cardsOf = new Map(); // customer gid (or null) → that customer's campaign cards, oldest first
  for (const card of sortedCards) {
    const key = card.customerId ?? null;
    if (!cardsOf.has(key)) cardsOf.set(key, []);
    cardsOf.get(key).push(card);
  }
  const isTagged = (cid) => taggedIds.has(cid);
  const settle = durationText(unknownSettleMs);

  const issues = [];
  const fixes = [];
  /** `journalText`, `shopifyText` and `action` are bilingual pairs ({ zh, en }, see bi). */
  const add = (type, customerId, giftCardId, journalText, shopifyText, action) => {
    issues.push({
      type,
      customerId: customerId ?? null,
      giftCardId: giftCardId ?? null,
      journal: journalText.zh,
      shopify: shopifyText.zh,
      action: action.zh,
      journalEn: journalText.en,
      shopifyEn: shopifyText.en,
      actionEn: action.en,
    });
  };
  // Customer → the card the journal holds once this run's fix-ups are applied.
  const keeperOf = new Map();

  // Every recipient in issue order, then customers known only to the journal.
  const customerIds = [...recipients.keys(), ...[...journal.keys()].filter((cid) => !recipients.has(cid))];
  for (const cid of customerIds) {
    const js = journal.get(cid) ?? null;
    const own = cardsOf.get(cid) ?? [];
    const isRecipient = recipients.has(cid);
    const journalCardId = js?.giftCardId ?? null;
    let recorded = null; // the card this run records in the journal, if any

    if (js && (js.status === STATUS.IN_PROGRESS || js.status === STATUS.UNKNOWN)) {
      // A create whose answer never arrived. No create can be running now (we hold the run lock).
      const before = describeJournal(js, tz);
      if (own.length) {
        recorded = own.find((c) => c.id === journalCardId) ?? own[0];
        fixes.push(foundFix(cid, recorded));
        const card = describeCard(recorded, tz);
        const untagged = untaggedNote(isTagged(cid), tag);
        add('resolved-unknown', cid, recorded.id, before,
          bi(`找到了这张卡：${card.zh}${untagged.zh}`, `Found the card: ${card.en}${untagged.en}`),
          RECORDED_ACTION);
      } else {
        const startedMs = Date.parse(js.startedAt ?? '');
        const waitedMs = nowMs - startedMs;
        if (Number.isFinite(startedMs) && waitedMs >= unknownSettleMs) {
          fixes.push({ op: 'reconcile.none', cid, note: `verify：开始建卡 ${settle.zh}后仍没在 Shopify 找到这张卡，确认没有建成` });
          add('resolved-unknown', cid, null, before,
            bi(`开始建卡超过 ${settle.zh}仍没找到这张卡，确认没有建成`, `Card still not found more than ${settle.en} after creation started; confirmed not created`),
            settledNoneAction(noNewCardsSince));
        } else {
          // Too early to rule out search-index lag (or no start time at all): leave the row alone.
          let action;
          if (Number.isFinite(startedMs)) {
            const wait = durationText(unknownSettleMs - waitedMs);
            action = bi(`约 ${wait.zh}后再运行 verify 或 issue 复查`, `Run verify or issue again in about ${wait.en} to check`);
          } else {
            action = bi('在后台按客户查看有没有这张卡', 'Look up the customer in the admin to see whether the card exists');
          }
          add('still-unknown', cid, null, before,
            bi('还没找到这张卡（刚建的卡可能还没进搜索索引）', 'Card not found yet (a new card may not be in the search index yet)'),
            action);
        }
      }
    } else if (isRecipient && !journalCardId && own.length) {
      // Shopify has a card the journal never heard of (lost journal lines, or a card that only
      // appeared in the search index after a reconcile.none): record the oldest one.
      recorded = own[0];
      fixes.push(foundFix(cid, recorded));
      const card = describeCard(recorded, tz);
      const untagged = untaggedNote(isTagged(cid), tag);
      add('not-in-journal', cid, recorded.id, describeJournal(js, tz),
        bi(`有本活动的卡：${card.zh}${untagged.zh}`, `Has a card from this campaign: ${card.en}${untagged.en}`),
        RECORDED_ACTION);
    }

    const keeperId = recorded?.id ?? journalCardId;
    if (keeperId) keeperOf.set(cid, keeperId);

    if (journalCardId && !cardById.has(journalCardId)) {
      const created = when(js.createdAt, tz);
      add('missing-in-shopify', cid, journalCardId,
        bi(`日志记录了这张卡：尾号 ${js.last4 || '?'}，${created.zh} 建卡`, `Journal has this card: last 4 ${js.last4 || '?'}, created ${created.en}`),
        bi('没找到这张卡（刚建的卡可能还没进搜索索引，或者 note 被改过）', 'Card not found (a new card may not be in the search index yet, or its note was changed)'),
        bi('过几分钟再运行 verify；仍然没有就在后台按礼品卡 ID 查看', 'Run verify again in a few minutes; if still missing, look up the gift card by id in the admin'));
    }

    // More than one card once the journal's record and Shopify's list are combined.
    const allIds = new Set(own.map((c) => c.id));
    if (journalCardId) allIds.add(journalCardId);
    if (allIds.size > 1) {
      const keeper = cardById.get(keeperId) ?? null;
      const keeperLast4 = keeper?.last4 || (keeperId === journalCardId ? js?.last4 : '') || '?';
      const journalText = recorded && recorded.id !== journalCardId
        ? bi(`日志里原来没有卡，本次按 Shopify 补记了尾号 ${keeperLast4} 那张`, `The journal had no card; this run recorded the one ending ${keeperLast4} from Shopify`)
        : bi(
          `日志记录的是尾号 ${keeperLast4} 那张${keeper ? '' : '（这次在 Shopify 上没找到）'}`,
          `The journal holds the one ending ${keeperLast4}${keeper ? '' : ' (not found in Shopify this time)'}`,
        );
      for (const card of own) {
        if (card.id === keeperId) continue;
        const text = describeCard(card, tz);
        add('duplicate-cards', cid, card.id, journalText,
          bi(`这位客户一共有 ${allIds.size} 张本活动的卡，这张是多出来的：${text.zh}`, `This customer has ${allIds.size} cards from this campaign; this one is extra: ${text.en}`),
          duplicateAction(card, keeperLast4));
      }
    }

    if (js?.taggedAt && !isTagged(cid)) {
      const tagged = when(js.taggedAt, tz);
      add('tag-missing', cid, keeperId,
        bi(`日志记录 ${tagged.zh} 已打 tag`, `Journal says tagged ${tagged.en}`),
        bi(`客户没有 ${tag}（刚打的 tag 可能还没进搜索索引）`, `Customer does not carry ${tag} (a new tag may not be in the search index yet)`),
        bi(`过几分钟再运行 verify；仍然没有就在后台给客户补上 ${tag}`, `Run verify again in a few minutes; if still missing, add ${tag} to the customer in the admin`));
    }
    if (isRecipient && isTagged(cid) && !own.length && !journalCardId) {
      add('tag-without-card', cid, null, describeJournal(js, tz),
        bi(`客户带 ${tag}，但没有本活动的卡`, `Customer carries ${tag} but has no card from this campaign`),
        bi('核实 tag 是谁加的：带着这个 tag，issue 会跳过这位客户', 'Check who added the tag: with this tag, issue skips this customer'));
    }
  }

  // Facts about each card on its own.
  for (const card of sortedCards) {
    const cid = card.customerId ?? null;
    const recipient = cid ? recipients.get(cid) : undefined;
    const journalText = cid ? describeJournal(journal.get(cid) ?? null, tz) : bi('日志里没有这张卡', 'No journal record for this card');
    if (!recipient) {
      const text = describeCard(card, tz);
      add('not-in-selection', cid, card.id, journalText,
        cid
          ? bi(`有本活动的卡（${text.zh}），但这位客户不在发放名单里`, `Has a card from this campaign (${text.en}), but this customer is not on the recipient list`)
          : bi(`这张卡带本活动的标记，但没有关联客户：${text.zh}`, `This card carries the campaign marker but has no customer: ${text.en}`),
        bi('核实这张卡是怎么建的；不该发的话在后台停用', 'Check how this card was created; disable it in the admin if it should not have been issued'));
    } else {
      const currencyDiffers = Boolean(currency && card.currencyCode && card.currencyCode !== currency);
      if (card.amountCents !== recipient.amountCents || currencyDiffers) {
        // With different currencies both sides name theirs; otherwise money() adds a code only when not USD.
        const listText = currencyDiffers ? `${formatUsd(recipient.amountCents)} ${currency}` : money(recipient.amountCents, currency);
        const cardText = currencyDiffers ? `${formatUsd(card.amountCents)} ${card.currencyCode}` : money(card.amountCents, card.currencyCode);
        add('amount-mismatch', cid, card.id,
          bi(`名单金额 ${listText}`, `List amount ${listText}`),
          bi(
            `卡的原金额 ${cardText}${currencyDiffers ? '，币种不一致' : ''}，尾号 ${card.last4 || '?'}`,
            `Card initial amount ${cardText}${currencyDiffers ? ', currency differs' : ''}, last 4 ${card.last4 || '?'}`,
          ),
          bi('核实金额；需要的话在后台调整这张卡的余额', 'Check the amount; adjust the card balance in the admin if needed'));
      }
    }
    if (!card.enabled) {
      const isExtra = keeperOf.has(cid) && keeperOf.get(cid) !== card.id;
      const text = describeCard(card, tz, { mentionDisabled: false });
      add('card-disabled', cid, card.id, journalText,
        bi(`这张卡已停用：${text.zh}`, `This card is disabled: ${text.en}`),
        isExtra
          ? bi('这是多出来的卡，已停用，不用再处理', 'This is an extra card and already disabled, nothing to do')
          : bi('如果不是有意停用，请在后台重新启用这张卡', 'If it was not disabled on purpose, re-enable the card in the admin'));
    }
  }

  const seqOf = (cid) => recipients.get(cid)?.seq ?? Number.MAX_SAFE_INTEGER;
  issues.sort((a, b) => TYPE_RANK.get(a.type) - TYPE_RANK.get(b.type)
    || seqOf(a.customerId) - seqOf(b.customerId)
    || compareGids(a.customerId, b.customerId)
    || compareGids(a.giftCardId, b.giftCardId));

  const counts = Object.fromEntries(ISSUE_TYPES.map(({ type }) => [type, 0]));
  for (const issue of issues) counts[issue.type] += 1;
  return { issues, fixes, counts };
}

// ---------------------------------------------------------------------------
// The command
// ---------------------------------------------------------------------------

function loadSelection(paths, config) {
  let selection;
  try {
    selection = readJson(paths.selection, null);
  } catch (err) {
    throw new Error(`名单文件读不出来：${err.message}`);
  }
  if (!selection) throw new Error(missingSelectionMessage(paths, config.campaignId));
  if (selection.campaignId !== config.campaignId) {
    throw new Error(`名单属于活动 ${selection.campaignId}，和现在的 CAMPAIGN_ID=${config.campaignId} 不一致，不能核对`);
  }
  if (!Array.isArray(selection.recipients) || !Number.isFinite(Date.parse(selection.createdAt))) {
    throw new Error(`${paths.selection} 内容不完整（缺少 recipients 或 createdAt）`);
  }
  return selection;
}

function printSummary({ report, fixes, selection, tag, paths, log }) {
  log.info(`核对完成：Shopify 上本活动的卡 ${report.cardCount} 张，带 ${tag} 的客户 ${report.taggedCount} 个，名单 ${selection.recipients.length} 人`);
  if (!report.issues.length) {
    log.info('没有发现问题');
  } else {
    const open = report.issues.filter((i) => !FIXED_TYPES.has(i.type));
    log.info(`发现 ${report.issues.length} 条：${open.length} 条需要人工看，${report.issues.length - open.length} 条已按 Shopify 补记`);
    for (const { type, label } of ISSUE_TYPES) {
      if (report.counts[type]) log.info(`  ${label}：${report.counts[type]}`);
    }
    for (const i of open.slice(0, MAX_LISTED)) {
      log.info(`  · ${TYPE_LABEL.get(i.type)} | 客户 ${tail(i.customerId)} | 卡 ${tail(i.giftCardId)} | ${i.action}`);
    }
    if (open.length > MAX_LISTED) log.info(`  …还有 ${open.length - MAX_LISTED} 条，见 Excel 的"核对"表`);
  }
  if (fixes.length) {
    const found = fixes.filter((f) => f.op === 'reconcile.found').length;
    log.info(`本地日志补记了 ${fixes.length} 条：${found} 条按 Shopify 记下已建的卡，${fixes.length - found} 条确认没有建卡`);
  }
  log.info(`核对结果已写入 ${paths.verify}（Excel 的"核对"表）`);
}

/** The verification itself, run while holding the run lock. */
async function verifyCampaign({ config, paths, selection, log, now, sleep, unknownSettleMs, runId }) {
  const stamp = () => now().toISOString();
  const tz = selection.params?.timezone || config.timezone || 'UTC';
  const tag = config.sentTag;
  const sinceIso = new Date(Date.parse(selection.createdAt) - SEARCH_MARGIN_MS).toISOString();

  log.info(`开始核对活动 ${config.campaignId}：名单 ${selection.recipients.length} 人（只读，不会改 Shopify 上的任何东西）`);
  if (selection.params?.sentTag && selection.params.sentTag !== tag) {
    log.warn(`注意：生成名单时的 SENT_TAG 是 ${selection.params.sentTag}，现在是 ${tag}；按现在的 ${tag} 核对`);
  }

  await connect(config, { log, sleep });
  const state = foldJournal(readJournal(paths.journal));

  const searchStarted = now();
  log.info(`查询 ${localDateTime(Date.parse(sinceIso), tz)} 之后建的、note 带 ${campaignMarker(config.campaignId)} 的礼品卡…`);
  const cards = await findCampaignCards({ campaignId: config.campaignId, sinceIso });
  log.info(`Shopify 上本活动的卡：${cards.length} 张`);

  log.info(`查询带 ${tag} 的客户…`);
  const snapshot = await fetchTagSnapshot({ config, paths, now });
  log.info(`带 ${tag} 的客户：${snapshot.ids.length} 个（已写入 tags.json）`);

  const { issues, fixes, counts } = analyzeVerify({
    selection,
    state,
    cards,
    taggedIds: snapshot.idSet,
    nowMs: searchStarted.getTime(),
    unknownSettleMs,
    tag,
    timezone: tz,
    remind1Date: config.remind1Date,
  });

  // Local fix-ups only; each line is fsynced on its own, so a crash here loses nothing already written.
  for (const fix of fixes) appendJournal(paths.journal, { ...fix, run: runId }, { now: stamp });

  const report = {
    version: VERIFY_VERSION,
    verifiedAt: searchStarted.toISOString(),
    cardCount: cards.length,
    taggedCount: snapshot.ids.length,
    counts,
    issues,
  };
  writeJsonAtomic(paths.verify, report);
  printSummary({ report, fixes, selection, tag, paths, log });

  const summary = {
    cardCount: report.cardCount,
    taggedCount: report.taggedCount,
    issueCount: issues.length,
    fixedCount: fixes.length,
    counts: Object.fromEntries(Object.entries(counts).filter(([, n]) => n)),
    text: issues.length
      ? `卡 ${report.cardCount} 张，发现 ${issues.length} 条，补记日志 ${fixes.length} 条`
      : `卡 ${report.cardCount} 张，没有发现问题`,
  };
  return { report, summary };
}

/**
 * `node index.js verify`
 *
 * @param {object} options
 * @param {object} options.config loadConfig() result
 * @param {{info: Function, warn: Function, error: Function}} [options.log]
 * @param {() => Date} [options.now]
 * @param {(ms: number) => Promise<void>} [options.sleep] passed to the Shopify client (rate-limit waits)
 * @param {Function} [options.writeReport] defaults to src/report/excel.js
 * @param {number} [options.unknownSettleMs] how long after create.start a card that still cannot
 *   be found is taken as never created (same rule as issue: 10 minutes)
 * @returns {Promise<{ exitCode: number, result: object|null }>} result = the verify.json content.
 *   exitCode 0 = verification completed (issues found are still a success); 1 = not run or failed.
 */
export async function runVerify({
  config,
  log = console,
  now = () => new Date(),
  sleep,
  writeReport,
  unknownSettleMs = DEFAULT_UNKNOWN_SETTLE_MS,
} = {}) {
  const paths = campaignPaths(config);
  // Checked before taking the lock so a wrong CAMPAIGN_ID does not leave an empty campaign folder.
  if (!fs.existsSync(paths.selection)) {
    log.error(missingSelectionMessage(paths, config.campaignId));
    return { exitCode: 1, result: null };
  }

  let release;
  try {
    release = acquireRunLock(paths, 'verify');
  } catch (err) {
    log.error(err.message);
    return { exitCode: 1, result: null };
  }

  // Read under the lock: select cannot be rewriting the list now.
  let selection;
  try {
    selection = loadSelection(paths, config);
  } catch (err) {
    release();
    log.error(err.message);
    return { exitCode: 1, result: null };
  }

  const runId = newRunId(now());
  const stamp = () => now().toISOString();
  let exitCode = 1;
  let result = null;
  let summary = null;
  try {
    appendJournal(paths.journal, { op: 'run.start', run: runId, command: 'verify', dryRun: false, batch: null, limit: null, options: {} }, { now: stamp });
    const outcome = await verifyCampaign({ config, paths, selection, log, now, sleep, unknownSettleMs, runId });
    result = outcome.report;
    summary = outcome.summary;
    exitCode = 0;
  } catch (err) {
    log.error(`verify 没有完成：${err.message}`);
    summary = { error: clip(err.message, 500), text: `失败：${clip(err.message)}` };
  } finally {
    try {
      appendJournal(paths.journal, { op: 'run.end', run: runId, summary, exitCode }, { now: stamp });
    } catch (err) {
      log.warn(`本地日志没写上这次 verify 的结束记录：${err.message}`);
    }
    release();
    // After the lock is released; never throws, so it cannot hide the outcome above.
    await refreshReport({ config, paths, log, now, writeReport });
  }
  return { exitCode, result };
}
