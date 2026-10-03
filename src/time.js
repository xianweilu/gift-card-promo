// Calendar maths in the store's time zone, without dependencies.

/** Calendar date/time parts of `date` as seen in `timeZone`. */
export function zonedParts(date, timeZone) {
  const fmt = new Intl.DateTimeFormat('en-CA', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hourCycle: 'h23',
  });
  const p = Object.fromEntries(fmt.formatToParts(date).map((x) => [x.type, x.value]));
  return { year: +p.year, month: +p.month, day: +p.day, hour: +p.hour, minute: +p.minute, second: +p.second };
}

/** Offset of `timeZone` from UTC at instant `ms`, in minutes (local minus UTC). */
export function zoneOffsetMinutes(ms, timeZone) {
  const p = zonedParts(new Date(ms), timeZone);
  const asUtc = Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute, p.second);
  return Math.round((asUtc - Math.floor(ms / 1000) * 1000) / 60000);
}

/** Instant (ms) of local midnight on year-month-day in `timeZone`. */
export function zonedMidnight(year, month, day, timeZone) {
  const guess = Date.UTC(year, month - 1, day);
  let ms = guess - zoneOffsetMinutes(guess, timeZone) * 60000;
  ms = guess - zoneOffsetMinutes(ms, timeZone) * 60000; // re-check across a DST change
  return ms;
}

/** Local midnight `months` calendar months before `now` (day clamped to the month's length). */
export function monthsAgoMidnight(now, months, timeZone) {
  const p = zonedParts(now, timeZone);
  let year = p.year;
  let month = p.month - months;
  while (month <= 0) {
    month += 12;
    year -= 1;
  }
  const daysInMonth = new Date(Date.UTC(year, month, 0)).getUTCDate();
  return zonedMidnight(year, month, Math.min(p.day, daysInMonth), timeZone);
}

const pad = (n) => String(n).padStart(2, '0');

/** ISO-8601 with the zone's offset, e.g. 2026-07-01T00:00:00-07:00 (used in Shopify search). */
export function isoWithOffset(ms, timeZone) {
  const off = zoneOffsetMinutes(ms, timeZone);
  const local = new Date(ms + off * 60000);
  const sign = off >= 0 ? '+' : '-';
  const a = Math.abs(off);
  return `${local.getUTCFullYear()}-${pad(local.getUTCMonth() + 1)}-${pad(local.getUTCDate())}`
    + `T${pad(local.getUTCHours())}:${pad(local.getUTCMinutes())}:${pad(local.getUTCSeconds())}`
    + `${sign}${pad(Math.floor(a / 60))}:${pad(a % 60)}`;
}

/** "YYYY-MM-DD" of instant `ms` in `timeZone`. */
export function localDate(ms, timeZone) {
  const p = zonedParts(new Date(ms), timeZone);
  return `${p.year}-${pad(p.month)}-${pad(p.day)}`;
}

/** "YYYY-MM-DD HH:MM" of instant `ms` in `timeZone` (for people reading the Excel). */
export function localDateTime(ms, timeZone) {
  const p = zonedParts(new Date(ms), timeZone);
  return `${p.year}-${pad(p.month)}-${pad(p.day)} ${pad(p.hour)}:${pad(p.minute)}`;
}
