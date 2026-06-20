# pptx-web

A fast, browser-only PowerPoint `.pptx` viewer built with HTML, CSS, vanilla
JavaScript, Vite, and
[`@aiden0z/pptx-renderer`](https://github.com/aiden0z/pptx-renderer).

The GitHub Pages deployment target is:

**https://costineest.github.io/pptx-web/**

Files selected from the device stay in the browser. There is no application
backend and no upload step.

## Features

- Open a local `.pptx` with the file picker.
- Drop a presentation anywhere in the application.
- Open an HTTP(S) direct file URL when its server permits browser CORS access.
- Navigate with thumbnails, previous/next controls, a slide number, the
  keyboard, or a touch swipe.
- Search presentation text and highlight matching slide elements.
- Zoom, fit the slide to the stage, and enter fullscreen mode.
- Cancel active downloads and rendering work.
- Reject invalid input, legacy `.ppt` files, oversized files, and unsafe ZIP
  expansion through explicit limits.

## Architecture

The renderer parses PresentationML and produces browser-native HTML/SVG DOM.
Its supported model includes text and style inheritance, shapes, images,
tables, charts, SmartArt fallback data, groups, themes, backgrounds, and
gradients.

The application has two main modules:

1. [`src/source.js`](./src/source.js) normalizes local and remote presentations
   into the same bounded `ArrayBuffer` contract. It validates the ZIP signature,
   streams URL responses with progress, enforces a 200 MB compressed-file
   limit, and supports cancellation.
2. [`src/main.js`](./src/main.js) owns the viewer lifecycle and UI. It loads the
   renderer dynamically, renders slides as a windowed vertical document,
   mounts thumbnails as they approach the sidebar viewport, and disposes
   renderer resources when another file is opened.

### Why no Web Worker or WASM?

The chosen renderer builds DOM/SVG and relies on browser DOM parsing, so the
rendering phase belongs on the main thread. A separate outline worker would
unzip the presentation a second time, increasing CPU and peak memory. Lazy
slide/media parsing and single-slide rendering avoid most initial work without
that duplication.

A LibreOffice-class WASM engine would add a much larger download, startup cost,
and memory footprint. It is a potential fallback only when fidelity for
unsupported PowerPoint effects matters more than startup performance.

## Performance

- The application entry is about 17 kB minified and 6 kB gzip.
- The renderer is isolated in a lazy chunk and prefetched during idle time.
- `lazySlides` and `lazyMedia` defer unvisited content.
- Windowed document mode mounts only slides near the right-hand scrollport.
- `IntersectionObserver` progressively mounts thumbnail previews.
- New opens abort stale fetch and rendering work.
- `RECOMMENDED_ZIP_LIMITS` guards against oversized ZIP entries and
  decompression bombs.

## Local development

Node.js 22 or newer is recommended.

```bash
npm install
npm run dev
```

Because the production base path is `/pptx-web/`, Vite serves the development
app at:

```text
http://localhost:5173/pptx-web/
```

Create and inspect a production build with:

```bash
npm run build
npm run preview
```

## Tests

The canonical fixture is the 89-slide, 19.1 MB Chapter 1 presentation from
Kurose and Ross's *Computer Networking: A Top-Down Approach*, 9th edition:

[Chapter_1_v9.0.pptx](https://gaia.cs.umass.edu/kurose_ross/ppt-9e/Chapter_1_v9.0.pptx)

A pinned copy lives at
[`public/fixtures/chapter-1-v9.0.pptx`](./public/fixtures/chapter-1-v9.0.pptx)
so unit tests remain deterministic. The tests verify its exact size and SHA-256
digest in addition to synthetic validation and streaming cases.

```bash
# Deterministic unit tests, including the pinned real presentation
npm test

# Network integration test against the original UMass URL
npm run test:remote
```

The UMass host currently does not return an `Access-Control-Allow-Origin`
header. For this exact fixture URL, the app transparently uses the pinned
same-origin copy instead. Vite serves that mirror during local development and
copies it into the GitHub Pages artifact. Other remote URLs still require their
servers to enable CORS.

## URL opening

The URL dialog accepts HTTP and HTTPS direct file responses. A presentation can
also be opened at startup with a URL-encoded query parameter:

```text
https://costineest.github.io/pptx-web/?url=https%3A%2F%2Ffiles.example%2Fdeck.pptx
```

Sharing pages from Google Drive, SharePoint, and similar services are usually
not direct file responses. They need a provider-specific direct link or a
server-side proxy that explicitly allows the viewer origin.

## GitHub Pages deployment

[`deploy-pages.yml`](./.github/workflows/deploy-pages.yml) tests and builds the
application on every push to `main`, then publishes `dist/` through GitHub
Pages.

For the first deployment, open the repository on GitHub and select:

```text
Settings > Pages > Build and deployment > Source > GitHub Actions
```

The workflow needs only the standard Pages permissions declared in the file.
No repository secret is required.

## Project structure

```text
.
|-- .github/workflows/deploy-pages.yml
|-- index.html
|-- public/
|   `-- fixtures/chapter-1-v9.0.pptx
|-- src/
|   |-- main.js
|   |-- source.js
|   `-- styles.css
|-- test/
|   `-- source.test.js
|-- integration/remote-source.mjs
|-- vite.config.js
`-- package.json
```

## Known limitations

- Legacy binary `.ppt` files are not supported.
- PowerPoint animations, transitions, macros, and some proprietary effects do
  not have complete browser equivalents.
- Rendering can differ when a presentation uses fonts unavailable on the
  viewer's device.
- EMF content with embedded PDF previews needs the renderer's optional
  `pdfjs-dist` integration, which is intentionally excluded from the default
  bundle.

## License

[MIT](./LICENSE)
