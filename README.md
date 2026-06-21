# pptx-web

A fast, browser-only PowerPoint `.pptx` viewer built with HTML, CSS, vanilla
JavaScript, Vite, and
[`@aiden0z/pptx-renderer`](https://github.com/aiden0z/pptx-renderer).

The GitHub Pages deployment target is
[costineest.github.io/pptx-web](https://costineest.github.io/pptx-web/).

Files selected from the device stay in the browser and are never uploaded.
Remote URLs are fetched directly first, with a restricted Cloudflare Worker as
a fallback when browser CORS prevents access.

## Table of contents

- [Features](#features)
- [Using the viewer](#using-the-viewer)
- [Keyboard shortcuts](#keyboard-shortcuts)
- [Architecture](#architecture)
- [Why no Web Worker or WASM?](#why-no-web-worker-or-wasm)
- [Performance](#performance)
- [Local development](#local-development)
- [Tests](#tests)
- [Remote URLs and sharing](#remote-urls-and-sharing)
- [GitHub Pages deployment](#github-pages-deployment)
- [Project structure](#project-structure)
- [Known limitations](#known-limitations)
- [License](#license)

## Features

- Open local `.pptx` files with the file picker or drag and drop; local files
  remain in the browser and are never uploaded.
- Open direct HTTP or HTTPS file URLs with streamed download progress and
  automatic fallback to a restricted CORS proxy when direct access fails.
- Preserve successfully opened remote URLs in the address bar as reloadable,
  shareable `?url=` links, including remote URLs with their own query strings;
  opening a local file removes stale remote URL state.
- Load a shared presentation automatically when the page opens with a `url`
  query parameter.
- Scroll slides as a continuous, PDF-like vertical document with windowed
  rendering for large presentations.
- Browse progressively rendered thumbnails, track the active slide, and
  collapse the thumbnail sidebar when more viewing space is needed.
- Navigate with thumbnails, previous/next buttons, a typed slide number,
  keyboard shortcuts, or horizontal touch swipes.
- Search presentation text case-insensitively, move forward or backward
  through wrapping results, and highlight the matching slide element.
- Zoom from 25% to 250% in 10% steps, reset to a fitted 100% view, and enter
  fullscreen presentation mode.
- Show determinate progress when the file size is known, report proxy fallback
  and rendering status, and allow active work to be cancelled.
- Abort stale work when another presentation is opened and dispose previous
  viewer, thumbnail, and search-highlight resources.
- Reject unsupported URL protocols, invalid or empty files, legacy `.ppt`
  files, files larger than 200 MB, and unsafe ZIP expansion.
- Adapt the interface for desktop, mobile, touch, reduced-motion, and
  forced-colors environments, with native dialogs and labelled controls.

## Using the viewer

To open a local presentation, select **Open file** and choose a `.pptx`, or drag
the file anywhere onto the page. The file is read locally and does not leave
the device.

To open a remote presentation, select **Open URL**, paste a direct HTTP or
HTTPS `.pptx` response, and select **Open presentation**. After it renders, the
address bar contains a shareable link that will reopen the same presentation.

Once a presentation is open, scroll the document normally or use the
thumbnails, slide counter, navigation buttons, search box, zoom controls, fit
button, and fullscreen button. On a touch device, swipe left or right across
the presentation to change slides.

## Keyboard shortcuts

| Shortcut | Action |
| --- | --- |
| `Ctrl+O` or `Command+O` | Open the local file picker |
| `/` | Focus presentation search |
| `Enter` in search | Open the next search result |
| `Shift+Enter` in search | Open the previous search result |
| `Right Arrow`, `Page Down`, or `Space` | Go to the next slide |
| `Left Arrow` or `Page Up` | Go to the previous slide |
| `Home` | Go to the first slide |
| `End` | Go to the last slide |
| `F` | Enter or exit fullscreen mode |

## Architecture

The renderer parses PresentationML and produces browser-native HTML/SVG DOM.
Its supported model includes text and style inheritance, shapes, images,
tables, charts, SmartArt fallback data, groups, themes, backgrounds, and
gradients.

The application has three main JavaScript modules:

1. [`src/source.js`](./src/source.js) normalizes local and remote presentations
   into the same bounded `ArrayBuffer` contract. It validates the ZIP signature,
   streams URL responses with progress, enforces a 200 MB compressed-file
   limit, supports cancellation, and retries eligible HTTPS URLs through the
   configured CORS proxy after an eligible direct request fails.
2. [`src/url-state.js`](./src/url-state.js) creates safe, readable viewer URLs
   for remote presentations and removes stale remote state when a local file is
   opened.
3. [`src/main.js`](./src/main.js) owns the viewer lifecycle and UI. It loads the
   renderer dynamically, renders slides as a windowed vertical document,
   mounts thumbnails as they approach the sidebar viewport, and disposes
   renderer resources when another file is opened.

The separately deployed Cloudflare Worker is only a fallback transport for
eligible remote URLs. Local files and CORS-enabled remote hosts do not use it.

### Why no Web Worker or WASM?

The chosen renderer builds DOM/SVG and relies on browser DOM parsing, so the
rendering phase belongs on the main thread. A separate outline worker would
unzip the presentation a second time, increasing CPU and peak memory. Lazy
slide/media parsing and windowed document rendering avoid most initial work
without that duplication.

A LibreOffice-class WASM engine would add a much larger download, startup cost,
and memory footprint. It is a potential fallback only when fidelity for
unsupported PowerPoint effects matters more than startup performance.

## Performance

- The application entry is about 18.2 kB minified and 6.5 kB gzip.
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
Kurose and Ross's _Computer Networking: A Top-Down Approach_, 9th edition:

[Chapter_1_v9.0.pptx](https://gaia.cs.umass.edu/kurose_ross/ppt-9e/Chapter_1_v9.0.pptx)

A pinned copy lives at
[`test/fixtures/chapter-1-v9.0.pptx`](./test/fixtures/chapter-1-v9.0.pptx)
so unit tests remain deterministic. The tests verify its exact size and SHA-256
digest in addition to synthetic validation, streaming, CORS fallback, and
shareable URL cases.

```bash
# Deterministic unit tests, including the pinned real presentation
npm test

# Network integration test for the deployed CORS proxy fallback
npm run test:remote
```

The UMass host currently does not return an `Access-Control-Allow-Origin`
header. The app first tries that URL directly, then retries through the
restricted Cloudflare Worker. The pinned fixture is used only by automated
tests and is not included in the GitHub Pages artifact.

## Remote URLs and sharing

The URL dialog accepts HTTP and HTTPS direct file responses. Direct access is
attempted first; eligible HTTPS URLs fall back to the configured CORS proxy. A
presentation can also be opened at startup with the `url` query parameter:

```text
https://costineest.github.io/pptx-web/?url=https://gaia.cs.umass.edu/kurose_ross/ppt-9e/Chapter_1_v9.0.pptx
```

To open your own presentation this way:

1. Start with the viewer URL: `https://costineest.github.io/pptx-web/`.
2. Add `?url=` followed by the direct HTTPS URL of the `.pptx` file.
3. URL-encode the nested file URL when it contains its own query parameters or
   other reserved characters. The viewer does this automatically for links
   opened through the URL dialog.
4. Open or share the completed viewer URL. The presentation loads
   automatically when the page opens.

The example above opens the Chapter 1 presentation directly in the deployed
viewer. You can also open the viewer normally and paste a direct file URL into
the URL dialog. After the presentation opens successfully, the address bar is
updated with the same `?url=` link so it can be copied, shared, or reloaded.

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
No repository secret is required. The Cloudflare Worker is deployed and managed
separately from this Pages workflow.

## Project structure

```text
.
|-- .github/workflows/deploy-pages.yml
|-- .gitignore
|-- LICENSE
|-- README.md
|-- index.html
|-- integration/remote-source.mjs
|-- package-lock.json
|-- package.json
|-- src/
|   |-- main.js
|   |-- source.js
|   |-- styles.css
|   `-- url-state.js
|-- test/
|   |-- fixtures/chapter-1-v9.0.pptx
|   |-- source.test.js
|   `-- url-state.test.js
`-- vite.config.js
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
