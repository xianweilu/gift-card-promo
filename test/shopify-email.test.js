// Tests for the two reminder emails in notifications/shopify-email.remind-*.liquid.
//
// They are Shopify Email "code" templates, pasted into the admin by hand, so
// nothing in src/ reads them. These checks keep them pasteable: the variables
// Shopify Email requires are there, only variables Shopify Email provides are
// used, the copy carries the hardcoded expiry date, and the templates render.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { Liquid } from 'liquidjs';
import { ROOT_DIR } from '../src/config.js';

const DIR = path.join(ROOT_DIR, 'notifications');
const FILES = {
  remind1: 'shopify-email.remind-1.liquid',
  remind2: 'shopify-email.remind-2.liquid',
};
const TEMPLATES = Object.fromEntries(Object.entries(FILES).map(([k, f]) => [k, fs.readFileSync(path.join(DIR, f), 'utf8')]));

// What Shopify Email exposes to a code template (the subset these emails use).
const ALLOWED_VARIABLES = new Set(['shop.name', 'shop.url', 'customer.first_name', 'unsubscribe_link', 'open_tracking_block', 'promo_name']);
const SHOP = { name: 'LA Balloons', url: 'https://www.laballoons.com' };
const UNSUBSCRIBE = 'https://example.test/unsubscribe/abc';
const TRACKING = '<img src="https://example.test/open.gif" width="1" height="1">';

// strict: every variable must exist (catches typos); Shopify itself renders a missing one as nil.
function render(template, customer, { strict = true } = {}) {
  const engine = new Liquid({ strictVariables: strict, strictFilters: true });
  return engine.parseAndRender(template, { shop: SHOP, customer, unsubscribe_link: UNSUBSCRIBE, open_tracking_block: TRACKING });
}

const withoutComments = (t) => t.replace(/\{%-?\s*comment\s*-?%\}[\s\S]*?\{%-?\s*endcomment\s*-?%\}/g, '');

for (const [name, template] of Object.entries(TEMPLATES)) {
  test(`${FILES[name]}: contains what Shopify Email requires and nothing it lacks`, () => {
    const body = withoutComments(template);
    assert.ok(body.includes('{{ unsubscribe_link }}'), 'unsubscribe_link');
    assert.ok(body.includes('{{ open_tracking_block }}'), 'open_tracking_block');
    assert.ok(!/href="\{\{ unsubscribe_link \}\}"/.test(body), 'unsubscribe_link renders its own link; it is not wrapped in an <a>');
    // The editor takes a fragment: no document wrapper.
    assert.doesNotMatch(body, /<!DOCTYPE|<html|<head|<body|<\/html>/i);
    assert.ok(body.trim().startsWith('{%- assign promo_name'), 'starts with the name assigns, then the outer table');
    assert.ok(body.trim().endsWith('</table>'));
    assert.ok(/<!-- Header -->\s*<tr>\s*<td[^>]*>\s*<\/td>\s*<\/tr>/.test(body), 'the header cell is empty (no shop-name heading)');

    // Every {{ output }} uses a variable Shopify Email provides (or one assigned in the template).
    for (const m of body.matchAll(/\{\{-?\s*([\w.]+)[^}]*\}\}/g)) {
      assert.ok(ALLOWED_VARIABLES.has(m[1]), `unknown variable in output: ${m[0]}`);
    }
    assert.doesNotMatch(body, /gift_card/, 'Shopify Email has no gift_card');
    assert.doesNotMatch(body, /shopify_asset_url|email_logo_url|email_accent_color|notifications\/styles\.css/, 'notification-only things');
    assert.doesNotMatch(body, /\$\s?\d/, 'no amount anywhere');

    // The hardcoded expiry date, long form in the copy and short form in the T&C.
    assert.ok(body.includes('October 19, 2026'));
    assert.ok(body.includes('Expires on 10/19/2026.'));
    // Where the code is, and no gift-card page button.
    assert.match(body, /Just reply to this email and we&rsquo;ll send (your code|it) again/);
    // Dropped on 10/6: the Cha-Ching tip, the Promotional reward line, 'No strings attached'.
    for (const gone of ['Cha-Ching', 'Promotional reward only', 'No strings attached', 'Balloon Math']) {
      assert.ok(!body.includes(gone), `should be gone: ${gone}`);
    }
    // 'We appreciate…' and 'Your support…' are separate paragraphs.
    assert.ok(/needs!<\/p>\s*<p[^>]*>Your support means the world to our team\.<\/p>/.test(body), 'two paragraphs');
    // The T&C: one size smaller than the body, and the whole paragraph in the text colour (no red sentence).
    assert.ok(body.includes('<p style="margin: 0; font-size: 10px; line-height: 1.5;"><strong>** Terms &amp; Conditions:</strong><br>'));
    assert.ok(body.includes('paid in exchange for it. No cash value &amp; non-transferable. Valid only for merchandise online at'));
    assert.doesNotMatch(body, /FF2600|color: *red/i);
    assert.ok(body.includes('Shop Now at LABalloons.com'));
    assert.doesNotMatch(body, /Claim My Credit|View your Online Credit Code/);

    // The paste-in instructions at the top.
    // The first reminder goes to the ordered group only (its group tag); the second to everyone.
    const segment = name === 'remind1'
      ? "customer_tags CONTAINS 'OCT26RTPROMO-ORDERED' AND NOT customer_tags CONTAINS 'OCT26RTPROMO-USED'"
      : "customer_tags CONTAINS 'OCT26RTPROMO' AND NOT customer_tags CONTAINS 'OCT26RTPROMO-USED'";
    assert.ok(template.includes(`segment: ${segment}`), segment);
    assert.ok(template.includes('Subject'));
    assert.ok(template.includes('Preview text'));
  });

  test(`${FILES[name]}: renders with and without a first name`, async () => {
    const named = await render(template, { first_name: 'Maria' });
    assert.ok(named.includes('Hello Maria,'));
    assert.ok(named.includes(`marketing emails. ${UNSUBSCRIBE}</p>`), 'unsubscribe_link printed as is');
    assert.ok(named.includes(TRACKING));
    assert.ok(named.includes('href="https://www.laballoons.com"'));
    assert.ok(named.startsWith('<table width="100%"'), 'the comment and assigns leave nothing before the table');
    assert.ok(!named.includes('{{') && !named.includes('{%'), 'everything rendered');

    // A note after the name, as some accounts have, is dropped: "Bob (wholesale)" → "Bob".
    assert.ok((await render(template, { first_name: 'Bob (wholesale)' })).includes('Hello Bob,'));
    assert.ok((await render(template, { first_name: '  Ana  ' })).includes('Hello Ana,'));

    for (const customer of [{ first_name: '' }, { first_name: null }, {}]) {
      const nameless = await render(template, customer, { strict: false });
      assert.ok(nameless.includes('<p style="margin: 0 0 16px 0;">Hello,</p>'), `no name: ${JSON.stringify(customer)}`);
    }
  });
}

test('the first reminder follows the approved copy (10/6 image)', () => {
  const body = withoutComments(TEMPLATES.remind1);
  const order = [
    'This is a friendly reminder: you still have an <strong>unused credit</strong> at <strong>LA Balloons</strong>!',
    'Time is running out, and if you don&rsquo;t use it at <a href="{{ shop.url }}" style="color: #1990C6;">LABalloons.com</a> by <strong>October 19, 2026,</strong> you lose it.',
    'Shop Now at LABalloons.com',
    // The sent-on line and the 'can't find it' line share one paragraph (no gap between them).
    'Your credit code was sent to your email on October 7, 2026.<br>\n            Can&rsquo;t find it? Just reply to this email and we&rsquo;ll send your code again.</p>',
    'Shop now before it&rsquo;s too late!',
    'We appreciate your business, and thank you for trusting us with your balloon and party supply needs!',
    'Your support means the world to our team.',
    '- LA Balloons',
    '<strong>** Terms &amp; Conditions:</strong>',
    'paid in exchange for it. No cash value &amp; non-transferable. Valid only for merchandise online at',
    'Expires on 10/19/2026.',
  ];
  let at = 0;
  for (const text of order) {
    const i = body.indexOf(text, at);
    assert.ok(i >= 0, `missing or out of order: ${text}`);
    at = i + text.length;
  }
  for (const gone of ['Just a friendly reminder', 'October 5, 2026']) {
    assert.ok(!body.includes(gone), `should be gone: ${gone}`);
  }
  assert.ok(TEMPLATES.remind1.includes('Subject       ⏰ TIME IS RUNNING OUT: Your LA Balloons CASH expires Oct. 19th!'));
});

test('the second reminder follows the approved copy (10/6 image)', () => {
  const body = withoutComments(TEMPLATES.remind2);
  const order = [
    '<strong>This is your final reminder:</strong> your <strong>unused credit</strong> at <strong>LA Balloons</strong> expires on <strong>October 19, 2026!</strong></p>',
    'After that, the credit disappears for good.<br>\n            Don&rsquo;t let your money go to waste &mdash; spend it at <a href="{{ shop.url }}" style="color: #1990C6;">LABalloons.com</a> today!</p>',
    'Shop Now at LABalloons.com',
    'Can&rsquo;t find your credit code?<br>\n            Just reply to this email and we&rsquo;ll send it again!</p>',
    'We appreciate your business, and thank you for trusting us with your balloon and party supply needs!</p>',
    'Your support means the world to our team.</p>',
    '- LA Balloons</p>',
    '<strong>** Terms &amp; Conditions:</strong>',
    'paid in exchange for it. No cash value &amp; non-transferable. Valid only for merchandise online at',
    'Expires on 10/19/2026.',
  ];
  let at = 0;
  for (const text of order) {
    const i = body.indexOf(text, at);
    assert.ok(i >= 0, `missing or out of order: ${text}`);
    at = i + text.length;
  }
  for (const gone of ['only a few days away', 'October 19th', 'Your credit code was sent', 'Time is running out']) {
    assert.ok(!body.includes(gone), `should be gone: ${gone}`);
  }
});

test('the two reminders differ in urgency but share the T&C body', () => {
  assert.ok(withoutComments(TEMPLATES.remind1).includes('This is a friendly reminder'));
  assert.ok(withoutComments(TEMPLATES.remind2).includes('This is your final reminder'));
  assert.ok(TEMPLATES.remind1.includes('send on 2026-10-12') && TEMPLATES.remind2.includes('send on 2026-10-16'));
  const legal = (t) => /Valid only for merchandise online at[\s\S]*?Expires on 10\/19\/2026\.<\/p>/.exec(t)?.[0];
  assert.ok(legal(TEMPLATES.remind1));
  assert.equal(legal(TEMPLATES.remind1), legal(TEMPLATES.remind2));
});
