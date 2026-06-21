import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

import {
  assertPptxBytes,
  fetchPresentation,
  formatBytes,
  isPptxName,
  normalizeHttpUrl,
  readLocalPresentation,
} from '../src/source.js';

const zipBuffer = () => Uint8Array.from([0x50, 0x4b, 0x03, 0x04, 1, 2, 3]).buffer;
const fixtureUrl = new URL('./fixtures/chapter-1-v9.0.pptx', import.meta.url);
const fixtureSha256 = '82b5c78c1845c85f714ca8f7973f7252edae00265f75af5a12075d1b87591419';

test('recognizes PPTX names without accepting legacy PPT files', () => {
  assert.equal(isPptxName('Quarterly.PPTX'), true);
  assert.equal(isPptxName('https://example.com/deck.pptx?download=1'), true);
  assert.equal(isPptxName('deck.ppt'), false);
});

test('only accepts HTTP(S) URLs', () => {
  assert.equal(normalizeHttpUrl('/deck.pptx', 'https://example.com/app').href, 'https://example.com/deck.pptx');
  assert.throws(() => normalizeHttpUrl('file:///tmp/deck.pptx'), /Only HTTP and HTTPS/);
  assert.throws(() => normalizeHttpUrl(''), /Enter a URL/);
});

test('checks the ZIP signature and size before parsing', () => {
  assert.doesNotThrow(() => assertPptxBytes(zipBuffer()));
  assert.throws(() => assertPptxBytes(new ArrayBuffer(4)), /valid \.pptx/);
  assert.throws(() => assertPptxBytes(zipBuffer(), 2), /larger than 2 B/);
});

test('reads a local PPTX file through the same buffer contract', async () => {
  const file = new File([zipBuffer()], 'Demo.PPTX', { type: 'application/octet-stream' });
  const progress = [];
  const result = await readLocalPresentation(file, {
    onProgress: (value) => progress.push(value),
  });

  assert.equal(result.name, 'Demo.PPTX');
  assert.equal(result.source, 'file');
  assert.equal(result.buffer.byteLength, 7);
  assert.deepEqual(progress.at(-1), { loaded: 7, total: 7 });
});

test('accepts the canonical 89-slide Kurose and Ross presentation fixture', async () => {
  const bytes = await readFile(fixtureUrl);
  const file = new File([bytes], 'chapter-1-v9.0.pptx', { type: 'application/octet-stream' });
  const result = await readLocalPresentation(file);

  assert.equal(result.buffer.byteLength, 19_054_287);
  assert.equal(createHash('sha256').update(new Uint8Array(result.buffer)).digest('hex'), fixtureSha256);
});

test('rejects legacy local PowerPoint files', async () => {
  const file = new File([zipBuffer()], 'legacy.ppt');
  await assert.rejects(readLocalPresentation(file), /Legacy \.ppt files/);
});

test('fetches a streamed presentation and reports progress', async () => {
  const progress = [];
  const response = new Response(zipBuffer(), {
    headers: {
      'content-length': '7',
      'content-disposition': "attachment; filename*=UTF-8''Roadmap%202026.pptx",
    },
  });
  const result = await fetchPresentation('https://example.com/download', {
    fetchImpl: async () => response,
    onProgress: (value) => progress.push(value),
  });

  assert.equal(result.name, 'Roadmap 2026.pptx');
  assert.equal(result.buffer.byteLength, 7);
  assert.deepEqual(progress.at(-1), { loaded: 7, total: 7 });
});

test('falls back to the CORS proxy after a direct network failure', async () => {
  const requests = [];
  let fallbackCount = 0;
  const result = await fetchPresentation('https://example.com/deck.pptx', {
    proxyUrl: 'https://cors-proxy.costineest.workers.dev/',
    fetchImpl: async (url) => {
      requests.push(url.href);
      if (requests.length === 1) throw new TypeError('Failed to fetch');
      return new Response(zipBuffer(), { headers: { 'content-length': '7' } });
    },
    onProxyFallback: () => {
      fallbackCount += 1;
    },
  });

  assert.deepEqual(requests, [
    'https://example.com/deck.pptx',
    'https://cors-proxy.costineest.workers.dev/?url=https%3A%2F%2Fexample.com%2Fdeck.pptx',
  ]);
  assert.equal(fallbackCount, 1);
  assert.equal(result.url, 'https://example.com/deck.pptx');
  assert.equal(result.name, 'deck.pptx');
});

test('explains likely CORS failures', async () => {
  await assert.rejects(
    fetchPresentation('https://example.com/deck.pptx', {
      fetchImpl: async () => {
        throw new TypeError('Failed to fetch');
      },
    }),
    /CORS/,
  );
});

test('formats byte counts for UI messages', () => {
  assert.equal(formatBytes(1536), '1.5 KB');
  assert.equal(formatBytes(200 * 1024 * 1024), '200 MB');
});
