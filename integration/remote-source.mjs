import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import test from 'node:test';

import { fetchPresentation } from '../src/source.js';

const fixtureUrl =
  'https://gaia.cs.umass.edu/kurose_ross/ppt-9e/Chapter_1_v9.0.pptx';
const proxyUrl = 'https://cors-proxy.costineest.workers.dev/';

test('downloads the canonical deck through the CORS proxy fallback', async () => {
  const result = await fetchPresentation(fixtureUrl, {
    proxyUrl,
    fetchImpl: async (url, options) => {
      if (url.hostname === 'gaia.cs.umass.edu') throw new TypeError('Simulated browser CORS failure');
      return fetch(url, {
        ...options,
        headers: { Origin: 'https://costineest.github.io' },
      });
    },
  });
  const digest = createHash('sha256').update(new Uint8Array(result.buffer)).digest('hex');

  assert.equal(result.buffer.byteLength, 19_054_287);
  assert.equal(digest, '82b5c78c1845c85f714ca8f7973f7252edae00265f75af5a12075d1b87591419');
  assert.equal(result.source, 'url');
});
