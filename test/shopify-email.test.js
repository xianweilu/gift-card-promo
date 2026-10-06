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
    assert.ok(body.includes('Can&rsquo;t find your credit code? Just reply to this email and we&rsquo;ll send it again.'));
    // Dropped on 10/6: the 'code was sent on October 5' line, the Cha-Ching tip, the Promotional reward line, 'No strings attached'.
    for (const gone of ['Your credit code was sent', 'Cha-Ching', 'Promotional reward only', 'No strings attached', 'Balloon Math']) {
      assert.ok(!body.includes(gone), `should be gone: ${gone}`);
    }
    // 'We appreciate…' and 'Your support…' are separate paragraphs.
    assert.ok(/needs!<\/p>\s*<p[^>]*>Your support means the world to our team\.<\/p>/.test(body), 'two paragraphs');
    assert.ok(body.includes('Shop Now at LABalloons.com'));
    assert.doesNotMatch(body, /Claim My Credit|View your Online Credit Code/);

    // The paste-in instructions at the top.
    assert.ok(template.includes("customer_tags CONTAINS 'OCT26RTPROMO' AND NOT customer_tags CONTAINS 'OCT26RTPROMO-USED'"));
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
    'Can&rsquo;t find your credit code? Just reply to this email and we&rsquo;ll send it again.',
    'Shop now before it&rsquo;s too late!',
    'We appreciate your business, and thank you for trusting us with your balloon and party supply needs!',
    'Your support means the world to our team.',
    '- LA Balloons',
    '<strong>** Terms &amp; Conditions:</strong>',
    'paid in exchange for it. <span style="color: #FF2600;">No cash value &amp; non-transferable.</span> Valid only for merchandise online at',
    'Expires on 10/19/2026.',
  ];
  let at = 0;
  for (const text of order) {
    const i = body.indexOf(text, at);
    assert.ok(i >= 0, `missing or out of order: ${text}`);
    at = i + text.length;
  }
  for (const gone of ['Just a friendly reminder', 'Your credit code was sent']) {
    assert.ok(!body.includes(gone), `should be gone: ${gone}`);
  }
  assert.ok(TEMPLATES.remind1.includes('Subject       ⏰ TIME IS RUNNING OUT: Your LA Balloons CASH expires Oct. 19th!'));
});

test('the two reminders differ in urgency but share the T&C body', () => {
  assert.ok(withoutComments(TEMPLATES.remind1).includes('This is a friendly reminder'));
  assert.ok(withoutComments(TEMPLATES.remind2).includes('This is your final reminder'));
  assert.ok(TEMPLATES.remind1.includes('send on 2026-10-12') && TEMPLATES.remind2.includes('send on 2026-10-16'));
  const legal = (t) => /Valid only for merchandise online at[\s\S]*?Expires on 10\/19\/2026\.<\/p>/.exec(t)?.[0];
  assert.ok(legal(TEMPLATES.remind1));
  assert.equal(legal(TEMPLATES.remind1), legal(TEMPLATES.remind2));
});
