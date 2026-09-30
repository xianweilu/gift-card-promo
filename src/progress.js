import fs from 'node:fs';

/**
 * progress.json shape:
 *   { [customerId]: { displayName, createStartedAt?, giftCardId?, createdAt?, taggedAt?, error? } }
 *
 * Entries written before 2026-09-30 may also carry `sentAt` from the removed
 * explicit-send step (Shopify emails the card on creation); it is ignored.
 *
 * createStartedAt without giftCardId means a giftCardCreate call was sent but
 * its result never came back; the run loop refuses to create again until a
 * human has checked the admin. Only saveProgress() may write this file.
 */
export function loadProgress(file) {
  let text;
  try {
    text = fs.readFileSync(file, 'utf8');
  } catch (err) {
    if (err.code === 'ENOENT') return {};
    throw new Error(`Cannot read ${file}: ${err.message}`);
  }
  if (!text.trim()) return {};
  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (err) {
    throw new Error(
      `${file} is not valid JSON (${err.message}). Refusing to start from scratch because that could create duplicate gift cards; fix or move the file, then re-run.`,
    );
  }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw new Error(`${file} must contain a JSON object keyed by customer id`);
  }
  return parsed;
}

/** Atomic write: temp file → fsync → rename over the real file. */
export function saveProgress(file, progress) {
  const tmp = `${file}.tmp`;
  const fd = fs.openSync(tmp, 'w');
  try {
    fs.writeFileSync(fd, `${JSON.stringify(progress, null, 2)}\n`, 'utf8');
    fs.fsyncSync(fd);
  } finally {
    fs.closeSync(fd);
  }
  fs.renameSync(tmp, file);
}
