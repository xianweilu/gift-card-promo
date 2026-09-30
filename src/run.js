import { getAccessToken, assertScopes, REQUIRED_SCOPES } from './auth.js';
import { initClient } from './shopify.js';
import { fetchAllMembers } from './segment.js';
import { createGiftCard, formatAmount } from './giftcards.js';
import { addTag, fetchTaggedCustomerIds } from './customers.js';
import { loadProgress, saveProgress } from './progress.js';

export const consoleLog = {
  info: (...args) => console.log(...args),
  warn: (...args) => console.warn(...args),
  error: (...args) => console.error(...args),
};

/**
 * Phase 2 hook. Phase 1 gives everyone the fixed GIFT_CARD_VALUE; phase 2 will
 * derive the amount from the member's last order (orders.js) here.
 */
export function resolveAmount(member, config) {
  return config.giftCardValue;
}

export function memberLabel(member) {
  return `${member.displayName || '(no name)'} <${member.id}>`;
}

function emailOf(member) {
  return member.defaultEmailAddress?.emailAddress || '';
}

function newSummary() {
  return { processed: 0, created: 0, tagged: 0, skippedDone: 0, skippedTagged: 0, skippedNoEmail: 0, needsReview: 0, failed: 0 };
}

function printSummary(summary, log) {
  const skipped = summary.skippedDone + summary.skippedTagged + summary.skippedNoEmail;
  let line = `Summary: processed=${summary.processed} created=${summary.created} tagged=${summary.tagged} `
    + `skipped=${skipped} (done=${summary.skippedDone}, already-tagged=${summary.skippedTagged}, no-email=${summary.skippedNoEmail}) failed=${summary.failed}`;
  if (summary.needsReview) line += ` needs-review=${summary.needsReview}`;
  log.info(line);
}

function statusOf(member, entry, tagged) {
  if (entry?.taggedAt) return 'done';
  if (tagged.has(member.id)) return 'done (tagged in Shopify)';
  if (entry?.giftCardId) return 'card created, not tagged';
  if (entry?.createStartedAt) return 'NEEDS REVIEW: creation result unknown';
  return emailOf(member) ? 'pending' : 'no email';
}

function printMembers(members, progress, tagged, log) {
  for (const m of members) {
    log.info([
      m.id,
      m.displayName || '(no name)',
      emailOf(m) || '(no email)',
      `orders=${m.numberOfOrders ?? '?'}`,
      `lastOrderId=${m.lastOrderId ?? '-'}`,
      `[${statusOf(m, progress[m.id], tagged)}]`,
    ].join('  '));
  }
}

/**
 * The whole flow. mode: 'check' | 'list' | 'run'. Returns { exitCode, ... }
 * and never throws once the per-customer loop has started; earlier failures
 * (auth, scopes, reading progress, fetching the segment or the already-tagged set) propagate.
 */
export async function run({
  config,
  mode = 'run',
  limit = 0,
  log = consoleLog,
  now = () => new Date().toISOString(),
  sleep,
} = {}) {
  log.info(`Shop: ${config.shop}.myshopify.com   API version: ${config.apiVersion}`);
  const { token, scopes } = await getAccessToken(config);
  log.info(`Access token obtained. Granted scopes: ${scopes.join(', ') || '(none)'}`);
  assertScopes(scopes, REQUIRED_SCOPES);
  log.info(`All required scopes present: ${REQUIRED_SCOPES.join(', ')}`);
  if (mode === 'check') return { exitCode: 0, scopes };

  initClient({ shop: config.shop, apiVersion: config.apiVersion, token, log, sleep });

  const progress = loadProgress(config.progressFile);
  const save = () => saveProgress(config.progressFile, progress);

  // Fetch ALL pages before any write: tagging removes members from the segment in real time.
  log.info(`Fetching all members of ${config.segmentId} (page size ${config.segmentPageSize})...`);
  const { members, totalCount, pages } = await fetchAllMembers(config.segmentId, { pageSize: config.segmentPageSize, log });
  log.info(`Segment totalCount=${totalCount}; fetched ${members.length} unique member(s) in ${pages} page(s)`);
  if (totalCount !== members.length) {
    log.warn(`  totalCount (${totalCount}) differs from the fetched count (${members.length}): the segment changed while paging, or duplicates were removed`);
  }

  // Second line of defence, independent of progress.json: anyone who already
  // carries the sent tag in Shopify is skipped, whatever progress.json says.
  const tagged = await fetchTaggedCustomerIds(config.sentTag, { pageSize: config.segmentPageSize, log });

  if (mode === 'list') {
    printMembers(members, progress, tagged, log);
    return { exitCode: 0, members, totalCount };
  }

  const summary = newSummary();
  log.info(config.dryRun
    ? 'Mode: DRY RUN - nothing will be created, sent or tagged'
    : 'Mode: LIVE - gift cards WILL be created (Shopify emails each one to the customer automatically) and customers tagged');
  if (limit) log.info(`Limit: at most ${limit} customer(s) this run`);

  let current = null;
  try {
    for (const member of members) {
      current = member;
      const entry = progress[member.id] ?? {};

      if (entry.taggedAt) {
        summary.skippedDone += 1;
        log.info(`skip ${memberLabel(member)}: already done (tagged ${entry.taggedAt})`);
        continue;
      }
      if (tagged.has(member.id)) {
        summary.skippedTagged += 1;
        log.warn(`skip ${memberLabel(member)}: already has tag "${config.sentTag}" in Shopify`
          + (entry.giftCardId ? '' : ' but no record in progress.json (card issued by an earlier run or by hand)'));
        continue;
      }
      if (!emailOf(member)) {
        summary.skippedNoEmail += 1;
        log.warn(`skip ${memberLabel(member)}: no email address`);
        continue;
      }
      if (entry.createStartedAt && !entry.giftCardId) {
        const msg = `gift card creation for ${memberLabel(member)} started at ${entry.createStartedAt} but its result is unknown. `
          + 'Check Shopify admin > Products > Gift cards for this customer. If a card exists, set "giftCardId" on this entry in '
          + 'progress.json and remove "createStartedAt"; if none exists, delete the entry. Then re-run.';
        if (config.dryRun) {
          summary.needsReview += 1;
          log.warn(`[dry-run] would stop here: ${msg}`);
          continue;
        }
        throw new Error(msg);
      }

      const amount = formatAmount(resolveAmount(member, config));

      if (config.dryRun) {
        log.info(`[dry-run] would create a ${amount} ${config.giftCardCurrency} gift card for ${memberLabel(member)} (Shopify emails it automatically) and add tag "${config.sentTag}"`);
        summary.processed += 1;
        if (limit && summary.processed >= limit) {
          log.info(`Reached --limit ${limit}; stopping`);
          break;
        }
        continue;
      }

      if (!entry.giftCardId) {
        // Shopify emails the customer the "Gift card created" notification as part of
        // giftCardCreate (customerId is set), so there is no separate send step.
        // Write-ahead marker: if the process dies between here and the save below,
        // the next run stops on this customer instead of creating a second card.
        progress[member.id] = { displayName: member.displayName, createStartedAt: now() };
        save();
        const card = await createGiftCard(member.id, {
          amount,
          currencyCode: config.giftCardCurrency,
          note: config.giftCardNote,
          expiresOn: config.giftCardExpiresOn,
          templateSuffix: config.giftCardTemplateSuffix,
        });
        progress[member.id] = { displayName: member.displayName, giftCardId: card.id, createdAt: now() };
        save();
        summary.created += 1;
        log.info(`created ${amount} ${config.giftCardCurrency} gift card ${card.id} for ${memberLabel(member)} (notification email sent by Shopify on creation)`);
      }

      const record = progress[member.id];
      if (!record.taggedAt) {
        await addTag(member.id, config.sentTag);
        record.taggedAt = now();
        delete record.error;
        save();
        summary.tagged += 1;
        log.info(`tagged ${memberLabel(member)} with "${config.sentTag}"`);
      }

      summary.processed += 1;
      if (limit && summary.processed >= limit) {
        log.info(`Reached --limit ${limit}; stopping`);
        break;
      }
    }
  } catch (err) {
    summary.failed += 1;
    if (current && !config.dryRun) {
      const record = progress[current.id] ?? (progress[current.id] = { displayName: current.displayName });
      record.error = { at: now(), message: err.message };
      // The create was definitely rejected → drop the write-ahead marker so the next run retries.
      if (record.createStartedAt && !record.giftCardId && err.outcomeKnown === true) delete record.createStartedAt;
      save();
    }
    log.error(`Stopped at customer ${current ? memberLabel(current) : '(none)'}: ${err.message}`);
    printSummary(summary, log);
    return { exitCode: 1, summary, error: err };
  }

  printSummary(summary, log);
  return { exitCode: 0, summary };
}
