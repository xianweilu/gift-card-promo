import { test } from 'node:test';
import assert from 'node:assert/strict';
import { addressKey, describeKey, normalizeStreet } from '../src/select/address.js';

const addr = (address1, { address2 = '', city = '', provinceCode = '', zip = '', countryCodeV2 = 'US' } = {}) => ({ address1, address2, city, provinceCode, zip, countryCodeV2 });
const same = (a, b) => {
  const ka = addressKey(a);
  assert.ok(ka, `expected a comparable key for ${JSON.stringify(a)}`);
  assert.equal(ka, addressKey(b));
};
const different = (a, b) => assert.notEqual(addressKey(a), addressKey(b));

test('address: the plan examples that are the same person', () => {
  same(addr('1422 Gardena Avenue', { city: 'Glendale', provinceCode: 'CA', zip: '91204-1234' }), addr('1422 gardena ave.', { city: 'glendale', provinceCode: 'ca', zip: '91204' }));
  same(addr('123 Main St Apt 4B', { city: 'Austin', provinceCode: 'TX', zip: '78701' }), addr('123 Main Street #4b', { city: 'Austin', provinceCode: 'TX', zip: '78701' }));
  same(addr('123 Main St 4B', { zip: '78701' }), addr('123 Main St Unit 4B', { zip: '78701' }));
  // ZIP decides; the city spelling does not matter.
  same(addr('500 N Broadway', { city: 'Los Angeles', provinceCode: 'CA', zip: '90012' }), addr('500 North Broadway', { city: 'LA', provinceCode: 'CA', zip: '90012' }));
});

test('address: the plan examples that are different people', () => {
  different(addr('123 Main St Apt 4B', { zip: '78701' }), addr('123 Main St Apt 5C', { zip: '78701' }));
  different(addr('123 Main St', { city: 'Austin', provinceCode: 'TX', zip: '78701' }), addr('123 Main St', { city: 'Dallas', provinceCode: 'TX', zip: '75201' }));
});

test('address: not comparable → empty key (never treated as a duplicate)', () => {
  assert.equal(addressKey(null), '');
  assert.equal(addressKey(undefined), '');
  assert.equal(addressKey(addr('', { address2: 'Apt 4', zip: '78701' })), '');
  assert.equal(addressKey(addr('   ', { zip: '78701' })), '');
  assert.equal(addressKey(addr('123 Main St')), '', 'no ZIP and no city');
});

test('address: unit words, address2, P.O. boxes, accents and directions', () => {
  same(addr('123 Main St', { address2: 'Suite 200', zip: '78701' }), addr('123 Main St Ste 200', { zip: '78701' }));
  same(addr('123 Main St', { address2: 'Apartment 7', zip: '78701' }), addr('123 Main St No. 7', { zip: '78701' }));
  same(addr('P.O. Box 12', { zip: '78701' }), addr('PO Box 12', { zip: '78701' }));
  same(addr('Post Office Box 12', { zip: '78701' }), addr('p o box 12', { zip: '78701' }));
  same(addr('5 Calle Niño', { zip: '78701' }), addr('5 calle nino', { zip: '78701' }));
  same(addr('10 Northeast Elm Boulevard', { zip: '78701' }), addr('10 NE Elm Blvd', { zip: '78701' }));
  assert.equal(normalizeStreet('1422 Gardena Avenue, Apt. #4B'), '1422 gardena ave 4b');
  assert.equal(normalizeStreet('P.O. Box 77'), 'po box 77');
});

test('address: without a ZIP the city, state and country are compared', () => {
  same(addr('9 Elm St', { city: 'Saint Louis', provinceCode: 'MO' }), addr('9 Elm Street', { city: 'St Louis', provinceCode: 'mo' }));
  different(addr('9 Elm St', { city: 'Springfield', provinceCode: 'IL' }), addr('9 Elm St', { city: 'Springfield', provinceCode: 'MO' }));
});

test('address: country defaults to US; other countries keep their postal code letters', () => {
  same(addr('1 Main St', { zip: '78701', countryCodeV2: '' }), addr('1 Main St', { zip: '78701', countryCodeV2: 'US' }));
  same(addr('1 Wellington St', { zip: 'K1A 0B1', countryCodeV2: 'CA' }), addr('1 Wellington Street', { zip: 'k1a0b1', countryCodeV2: 'CA' }));
  different(addr('1 Main St', { zip: '78701', countryCodeV2: 'US' }), addr('1 Main St', { zip: '78701', countryCodeV2: 'MX' }));
  assert.equal(addressKey(addr('1 Main St', { zip: '78701-0001' })), '1 main st|78701|us');
});

test('address: describeKey is readable', () => {
  assert.equal(describeKey('1422 gardena ave|91204|us'), '1422 gardena ave · 91204 · US');
  assert.equal(describeKey('9 elm st|st louis|mo|us'), '9 elm st · st louis · mo · US');
  assert.equal(describeKey(''), '');
});
