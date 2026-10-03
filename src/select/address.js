// Same-address matching. Two customers are the same person when the keys of
// their default addresses are identical. No fuzzy matching: a missed duplicate
// is better than merging two different people.

const SUFFIX = {
  street: 'st', str: 'st', st: 'st',
  avenue: 'ave', av: 'ave', ave: 'ave', avn: 'ave', avenu: 'ave',
  road: 'rd', rd: 'rd',
  drive: 'dr', drv: 'dr', dr: 'dr',
  boulevard: 'blvd', boul: 'blvd', blvd: 'blvd',
  lane: 'ln', ln: 'ln',
  court: 'ct', ct: 'ct',
  circle: 'cir', cir: 'cir',
  place: 'pl', pl: 'pl',
  parkway: 'pkwy', pkwy: 'pkwy', pky: 'pkwy',
  highway: 'hwy', hwy: 'hwy',
  terrace: 'ter', ter: 'ter',
  trail: 'trl', trl: 'trl',
  square: 'sq', sq: 'sq',
  expressway: 'expy', expy: 'expy',
  freeway: 'fwy', fwy: 'fwy',
  alley: 'aly', aly: 'aly',
  crossing: 'xing', xing: 'xing',
  heights: 'hts', hts: 'hts',
  center: 'ctr', centre: 'ctr', ctr: 'ctr',
  route: 'rte', rte: 'rte',
  plaza: 'plz', plz: 'plz',
  point: 'pt', pt: 'pt',
  crescent: 'cres', cres: 'cres',
  mount: 'mt', mt: 'mt',
  mountain: 'mtn', mtn: 'mtn',
  creek: 'crk', crk: 'crk',
  estates: 'est', est: 'est',
};
const DIRECTION = { north: 'n', south: 's', east: 'e', west: 'w', northeast: 'ne', northwest: 'nw', southeast: 'se', southwest: 'sw' };
// Unit designators are dropped so that "Apt 4B", "#4B", "Unit 4B" and "4B" all become "4b".
const UNIT_WORDS = new Set(['apartment', 'apt', 'unit', 'suite', 'ste', 'no', 'number', 'num', 'rm', 'room', 'spc', 'space', '#']);

function clean(s) {
  return String(s ?? '').normalize('NFKD').replace(/[̀-ͯ]/g, '').toLowerCase();
}

/** One normalised street line from address1 + address2. */
export function normalizeStreet(s) {
  const t = clean(s)
    .replace(/#/g, ' # ')
    .replace(/\bp\s*\.?\s*o\s*\.?\s*box\b/g, ' pobox ')
    .replace(/\bpost\s+office\s+box\b/g, ' pobox ')
    .replace(/[^a-z0-9# ]+/g, ' ');
  const out = [];
  for (const w of t.split(/\s+/).filter(Boolean)) {
    if (UNIT_WORDS.has(w)) continue;
    out.push(DIRECTION[w] ?? SUFFIX[w] ?? (w === 'pobox' ? 'po box' : w));
  }
  return out.join(' ');
}

function normalizeCity(s) {
  return normalizeStreet(s).replace(/\bsaint\b/g, 'st').replace(/\bfort\b/g, 'ft');
}

/**
 * The comparison key of a MailingAddress, or '' when the address cannot be
 * compared (missing, no street, or neither ZIP nor city).
 *   with ZIP:    "<street>|<zip5>|<country>"            (city/state spelling ignored)
 *   without ZIP: "<street>|<city>|<state>|<country>"
 */
export function addressKey(a) {
  if (!a) return '';
  if (!normalizeStreet(a.address1)) return '';
  const street = normalizeStreet(`${a.address1 ?? ''} ${a.address2 ?? ''}`);
  const country = clean(a.countryCodeV2 || 'US').replace(/[^a-z]/g, '') || 'us';
  const zipRaw = clean(a.zip).replace(/[^a-z0-9]/g, '');
  const zip = country === 'us' ? zipRaw.replace(/[^0-9]/g, '').slice(0, 5) : zipRaw;
  if (zip.length >= 3) return `${street}|${zip}|${country}`;
  const city = normalizeCity(a.city);
  if (!city) return '';
  const province = clean(a.provinceCode).replace(/[^a-z]/g, '');
  return `${street}|${city}|${province}|${country}`;
}

/** Human-readable form of a key for the Excel ("1422 gardena ave · 91204 · US"). */
export function describeKey(key) {
  return key ? key.split('|').filter(Boolean).map((p, i, arr) => (i === arr.length - 1 ? p.toUpperCase() : p)).join(' · ') : '';
}
