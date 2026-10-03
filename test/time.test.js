import { test } from 'node:test';
import assert from 'node:assert/strict';
import { zonedParts, zoneOffsetMinutes, zonedMidnight, monthsAgoMidnight, isoWithOffset, localDate, localDateTime } from '../src/time.js';

const LA = 'America/Los_Angeles';

test('time: three months before 2026-10-01 13:00 PDT is 2026-07-01 00:00 PDT', () => {
  const ms = monthsAgoMidnight(new Date('2026-10-01T20:00:00Z'), 3, LA);
  assert.equal(ms, Date.parse('2026-07-01T00:00:00-07:00'));
  assert.equal(isoWithOffset(ms, LA), '2026-07-01T00:00:00-07:00');
});

test('time: the cutoff uses the store calendar day, not the UTC day', () => {
  // 2026-10-02 05:00 UTC is still 2026-10-01 in Los Angeles.
  assert.equal(monthsAgoMidnight(new Date('2026-10-02T05:00:00Z'), 3, LA), Date.parse('2026-07-01T00:00:00-07:00'));
  assert.equal(monthsAgoMidnight(new Date('2026-10-02T08:00:00Z'), 3, LA), Date.parse('2026-07-02T00:00:00-07:00'));
});

test('time: month ends are clamped and DST is respected', () => {
  // May 31 → Feb 28 (2026 is not a leap year), in PST.
  assert.equal(monthsAgoMidnight(new Date('2026-05-31T19:00:00Z'), 3, LA), Date.parse('2026-02-28T00:00:00-08:00'));
  // Across a year boundary into standard time.
  assert.equal(monthsAgoMidnight(new Date('2026-03-15T19:00:00Z'), 3, LA), Date.parse('2025-12-15T00:00:00-08:00'));
  // 2028 is a leap year.
  assert.equal(monthsAgoMidnight(new Date('2028-05-31T19:00:00Z'), 3, LA), Date.parse('2028-02-29T00:00:00-08:00'));
  assert.equal(monthsAgoMidnight(new Date('2026-10-01T20:00:00Z'), 12, LA), Date.parse('2025-10-01T00:00:00-07:00'));
});

test('time: local midnight on DST change days', () => {
  assert.equal(zonedMidnight(2026, 3, 8, LA), Date.parse('2026-03-08T00:00:00-08:00')); // spring forward at 2am
  assert.equal(zonedMidnight(2026, 3, 9, LA), Date.parse('2026-03-09T00:00:00-07:00'));
  assert.equal(zonedMidnight(2026, 11, 1, LA), Date.parse('2026-11-01T00:00:00-07:00')); // fall back at 2am
  assert.equal(zonedMidnight(2026, 11, 2, LA), Date.parse('2026-11-02T00:00:00-08:00'));
  assert.equal(zonedMidnight(2026, 7, 1, 'UTC'), Date.parse('2026-07-01T00:00:00Z'));
});

test('time: offsets and ISO strings with offsets', () => {
  assert.equal(zoneOffsetMinutes(Date.parse('2026-07-01T12:00:00Z'), LA), -420);
  assert.equal(zoneOffsetMinutes(Date.parse('2026-12-01T12:00:00Z'), LA), -480);
  assert.equal(zoneOffsetMinutes(Date.parse('2026-07-01T12:00:00Z'), 'Asia/Shanghai'), 480);
  assert.equal(isoWithOffset(Date.parse('2026-12-01T08:00:00Z'), LA), '2026-12-01T00:00:00-08:00');
  assert.equal(isoWithOffset(Date.parse('2026-07-01T00:00:00Z'), 'Asia/Kolkata'), '2026-07-01T05:30:00+05:30');
  assert.equal(isoWithOffset(Date.parse('2026-07-01T00:00:00Z'), 'UTC'), '2026-07-01T00:00:00+00:00');
});

test('time: calendar date and time in the store time zone', () => {
  assert.equal(localDate(Date.parse('2026-10-02T06:59:59Z'), LA), '2026-10-01');
  assert.equal(localDate(Date.parse('2026-10-02T07:00:00Z'), LA), '2026-10-02');
  assert.equal(localDate(new Date('2026-10-05T16:00:00Z'), LA), '2026-10-05', 'accepts a Date too');
  assert.equal(localDateTime(Date.parse('2026-10-05T16:04:00Z'), LA), '2026-10-05 09:04');
  assert.equal(localDateTime(Date.parse('2026-10-05T07:00:00Z'), LA), '2026-10-05 00:00', 'midnight is 00, not 24');
  assert.deepEqual(zonedParts(new Date('2026-10-05T16:04:05Z'), LA), { year: 2026, month: 10, day: 5, hour: 9, minute: 4, second: 5 });
});
