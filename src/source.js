export const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
export const DEFAULT_MAX_BYTES = 200 * 1024 * 1024;

const PPTX_SIGNATURES = [
  [0x50, 0x4b, 0x03, 0x04],
  [0x50, 0x4b, 0x05, 0x06],
  [0x50, 0x4b, 0x07, 0x08],
];

const PRESENTATION_MIRRORS = new Map([
  [
    'https://gaia.cs.umass.edu/kurose_ross/ppt-9e/Chapter_1_v9.0.pptx',
    'fixtures/chapter-1-v9.0.pptx',
  ],
]);

export function isPptxName(name) {
  return /\.pptx(?:$|[?#])/i.test(name.trim());
}

export function assertPptxBytes(buffer, maxBytes = DEFAULT_MAX_BYTES) {
  if (!(buffer instanceof ArrayBuffer)) {
    throw new TypeError('Expected presentation data as an ArrayBuffer.');
  }
  if (buffer.byteLength === 0) {
    throw new Error('The presentation is empty.');
  }
  if (buffer.byteLength > maxBytes) {
    throw new Error(`The presentation is larger than ${formatBytes(maxBytes)}.`);
  }

  const prefix = new Uint8Array(buffer, 0, Math.min(4, buffer.byteLength));
  const isZip = PPTX_SIGNATURES.some(
    (signature) => signature.length === prefix.length && signature.every((byte, index) => byte === prefix[index]),
  );
  if (!isZip) {
    throw new Error('This does not look like a valid .pptx file.');
  }
}

export async function readLocalPresentation(file, options = {}) {
  const maxBytes = options.maxBytes ?? DEFAULT_MAX_BYTES;
  if (!(file instanceof Blob)) {
    throw new TypeError('Choose a PowerPoint .pptx file.');
  }
  if (file.name && !isPptxName(file.name)) {
    throw new Error('Choose a .pptx file. Legacy .ppt files are not supported.');
  }
  if (file.size > maxBytes) {
    throw new Error(`The presentation is larger than ${formatBytes(maxBytes)}.`);
  }

  options.onProgress?.({ loaded: 0, total: file.size });
  const buffer = await file.arrayBuffer();
  assertPptxBytes(buffer, maxBytes);
  options.onProgress?.({ loaded: buffer.byteLength, total: buffer.byteLength });
  return {
    buffer,
    name: file.name || 'Presentation.pptx',
    source: 'file',
  };
}

export async function fetchPresentation(urlValue, options = {}) {
  const maxBytes = options.maxBytes ?? DEFAULT_MAX_BYTES;
  const fetchImpl = options.fetchImpl ?? fetch;
  const url = resolvePresentationUrl(urlValue, options);

  let response;
  try {
    response = await fetchImpl(url, {
      signal: options.signal,
      credentials: 'omit',
      redirect: 'follow',
    });
  } catch (error) {
    if (error?.name === 'AbortError') throw error;
    throw new Error('The URL could not be fetched. The server may block browser access with CORS.', { cause: error });
  }

  if (!response.ok) {
    throw new Error(`The URL returned HTTP ${response.status}.`);
  }

  const lengthHeader = Number(response.headers.get('content-length'));
  const total = Number.isFinite(lengthHeader) && lengthHeader > 0 ? lengthHeader : null;
  if (total && total > maxBytes) {
    await response.body?.cancel();
    throw new Error(`The presentation is larger than ${formatBytes(maxBytes)}.`);
  }

  const buffer = response.body
    ? await readResponseStream(response.body, { total, maxBytes, onProgress: options.onProgress })
    : await response.arrayBuffer();
  assertPptxBytes(buffer, maxBytes);

  const headerName = filenameFromDisposition(response.headers.get('content-disposition'));
  const pathName = decodeURIComponent(url.pathname.split('/').pop() || '');
  const name = headerName || (isPptxName(pathName) ? pathName : 'Remote presentation.pptx');
  return { buffer, name, source: 'url', url: url.href };
}

export function resolvePresentationUrl(value, options = {}) {
  const url = normalizeHttpUrl(value, options.baseUrl);
  const mirrorPath = PRESENTATION_MIRRORS.get(url.href);
  return mirrorPath && options.mirrorBaseUrl ? new URL(mirrorPath, options.mirrorBaseUrl) : url;
}

export function normalizeHttpUrl(value, baseUrl = globalThis.location?.href) {
  const input = String(value ?? '').trim();
  if (!input) throw new Error('Enter a URL to a .pptx file.');

  let url;
  try {
    url = baseUrl ? new URL(input, baseUrl) : new URL(input);
  } catch {
    throw new Error('Enter a valid URL.');
  }
  if (!['http:', 'https:'].includes(url.protocol)) {
    throw new Error('Only HTTP and HTTPS URLs are supported.');
  }
  return url;
}

export function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  const power = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** power;
  return `${value >= 10 || power === 0 ? value.toFixed(0) : value.toFixed(1)} ${units[power]}`;
}

async function readResponseStream(stream, options) {
  const reader = stream.getReader();
  const chunks = [];
  let loaded = 0;
  options.onProgress?.({ loaded, total: options.total });

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      loaded += value.byteLength;
      if (loaded > options.maxBytes) {
        throw new Error(`The presentation is larger than ${formatBytes(options.maxBytes)}.`);
      }
      chunks.push(value);
      options.onProgress?.({ loaded, total: options.total });
    }
  } catch (error) {
    await reader.cancel(error).catch(() => {});
    throw error;
  } finally {
    reader.releaseLock();
  }

  const joined = new Uint8Array(loaded);
  let offset = 0;
  for (const chunk of chunks) {
    joined.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return joined.buffer;
}

function filenameFromDisposition(value) {
  if (!value) return '';
  const encoded = value.match(/filename\*=UTF-8''([^;]+)/i)?.[1];
  if (encoded) {
    try {
      return decodeURIComponent(encoded.replace(/^"|"$/g, ''));
    } catch {
      return encoded;
    }
  }
  return value.match(/filename="?([^";]+)"?/i)?.[1]?.trim() ?? '';
}
