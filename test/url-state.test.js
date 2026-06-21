import assert from 'node:assert/strict';
import test from 'node:test';

import { createViewerUrl } from '../src/url-state.js';

test('creates a readable share URL for a remote presentation', () => {
  const presentationUrl =
    'https://gaia.cs.umass.edu/kurose_ross/ppt-9e/Chapter_1_v9.0.pptx';
  const result = createViewerUrl('https://costineest.github.io/pptx-web/', presentationUrl);

  assert.equal(
    result.href,
    'https://costineest.github.io/pptx-web/?url=https://gaia.cs.umass.edu/kurose_ross/ppt-9e/Chapter_1_v9.0.pptx',
  );
  assert.equal(result.searchParams.get('url'), presentationUrl);
});

test('safely encodes reserved characters inside the presentation URL', () => {
  const presentationUrl = 'https://example.com/deck.pptx?download=1&token=a#slide=2';
  const result = createViewerUrl('https://costineest.github.io/pptx-web/?theme=dark', presentationUrl);

  assert.equal(result.searchParams.get('theme'), 'dark');
  assert.equal(result.searchParams.get('url'), presentationUrl);
});

test('removes a stale presentation URL for local files', () => {
  const result = createViewerUrl(
    'https://costineest.github.io/pptx-web/?theme=dark&url=https://example.com/deck.pptx#viewer',
    null,
  );

  assert.equal(result.href, 'https://costineest.github.io/pptx-web/?theme=dark#viewer');
});
