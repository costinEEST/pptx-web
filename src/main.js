import './styles.css';

import { fetchPresentation, formatBytes, readLocalPresentation } from './source.js';

const DOCUMENT_LIST_OPTIONS = {
  windowed: true,
  batchSize: 4,
  initialSlides: 4,
  overscanViewport: 1.5,
};
const MIRROR_BASE_URL = new URL(import.meta.env.BASE_URL, window.location.origin).href;

const elements = {
  app: byId('app'),
  themeColor: byId('theme-color'),
  sidebar: byId('sidebar'),
  sidebarToggle: byId('sidebar-toggle'),
  sidebarToggleTooltip: byId('sidebar-toggle-tooltip'),
  documentTitle: byId('document-title'),
  fileInput: byId('file-input'),
  urlButton: byId('url-button'),
  welcomeUrlButton: byId('welcome-url-button'),
  urlDialog: byId('url-dialog'),
  urlForm: byId('url-form'),
  urlInput: byId('url-input'),
  urlCancel: byId('url-cancel'),
  closeUrlDialog: byId('close-url-dialog'),
  stage: byId('stage'),
  viewerContainer: byId('viewer-container'),
  thumbnailList: byId('thumbnail-list'),
  searchInput: byId('search-input'),
  searchCount: byId('search-count'),
  searchPrev: byId('search-prev'),
  searchNext: byId('search-next'),
  loading: byId('loading'),
  loadingTitle: byId('loading-title'),
  loadingDetail: byId('loading-detail'),
  loadingProgress: byId('loading-progress'),
  cancelButton: byId('cancel-button'),
  dropLayer: byId('drop-layer'),
  errorToast: byId('error-toast'),
  errorMessage: byId('error-message'),
  dismissError: byId('dismiss-error'),
  status: byId('status'),
  previousButton: byId('previous-button'),
  nextButton: byId('next-button'),
  slideNumber: byId('slide-number'),
  slideCount: byId('slide-count'),
  zoomOut: byId('zoom-out'),
  zoomIn: byId('zoom-in'),
  zoomValue: byId('zoom-value'),
  fitButton: byId('fit-button'),
  fullscreenButton: byId('fullscreen-button'),
};

let viewer = null;
let operationController = null;
let rendererPromise = null;
let thumbnailObserver = null;
let thumbnailHandles = new Map();
let activeHighlight = null;
let searchResults = [];
let searchIndex = -1;
let dragDepth = 0;
let pointerStart = null;

wireEvents();
setViewerEnabled(false);
prefetchRenderer();
openInitialUrl();

function wireEvents() {
  elements.fileInput.addEventListener('change', async (event) => {
    const [file] = event.target.files ?? [];
    event.target.value = '';
    if (file) await openLocalFile(file);
  });

  elements.urlButton.addEventListener('click', showUrlDialog);
  elements.welcomeUrlButton.addEventListener('click', showUrlDialog);
  elements.urlCancel.addEventListener('click', closeUrlDialog);
  elements.closeUrlDialog.addEventListener('click', closeUrlDialog);
  elements.urlDialog.addEventListener('click', (event) => {
    if (event.target === elements.urlDialog) closeUrlDialog();
  });
  elements.urlForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    const url = elements.urlInput.value;
    closeUrlDialog();
    await openRemoteUrl(url);
  });

  elements.cancelButton.addEventListener('click', cancelOperation);
  elements.dismissError.addEventListener('click', hideError);
  elements.previousButton.addEventListener('click', () => navigateRelative(-1));
  elements.nextButton.addEventListener('click', () => navigateRelative(1));
  elements.slideNumber.addEventListener('change', goToTypedSlide);
  elements.slideNumber.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') goToTypedSlide();
  });
  elements.zoomOut.addEventListener('click', () => setZoom(viewer?.zoomPercent - 10));
  elements.zoomIn.addEventListener('click', () => setZoom(viewer?.zoomPercent + 10));
  elements.fitButton.addEventListener('click', fitSlide);
  elements.fullscreenButton.addEventListener('click', toggleFullscreen);
  elements.sidebarToggle.addEventListener('click', toggleSidebar);

  elements.searchInput.addEventListener('search', clearSearchIfEmpty);
  elements.searchInput.addEventListener('input', clearSearchIfEmpty);
  elements.searchInput.addEventListener('keydown', (event) => {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    runSearch(event.shiftKey ? -1 : 1);
  });
  elements.searchPrev.addEventListener('click', () => runSearch(-1));
  elements.searchNext.addEventListener('click', () => runSearch(1));

  window.addEventListener('dragenter', handleDragEnter);
  window.addEventListener('dragover', handleDragOver);
  window.addEventListener('dragleave', handleDragLeave);
  window.addEventListener('drop', handleDrop);
  window.addEventListener('keydown', handleKeyboard);
  document.addEventListener('fullscreenchange', updateFullscreenButton);

  elements.stage.addEventListener('pointerdown', handlePointerDown);
  elements.stage.addEventListener('pointerup', handlePointerUp);
  elements.stage.addEventListener('pointercancel', () => {
    pointerStart = null;
  });
}

async function openLocalFile(file) {
  hideError();
  const controller = beginOperation('Reading presentation', file.name || 'Local file');
  try {
    const source = await readLocalPresentation(file, {
      onProgress: updateLoadingProgress,
    });
    ensureCurrentOperation(controller);
    await renderPresentation(source, controller);
  } catch (error) {
    handleOpenError(error, controller);
  }
}

async function openRemoteUrl(url) {
  hideError();
  const controller = beginOperation('Downloading presentation', 'Connecting to the file server...');
  try {
    const source = await fetchPresentation(url, {
      signal: controller.signal,
      mirrorBaseUrl: MIRROR_BASE_URL,
      onProgress: ({ loaded, total }) => {
        updateLoadingProgress({ loaded, total });
        elements.loadingDetail.textContent = total
          ? `${formatBytes(loaded)} of ${formatBytes(total)}`
          : `${formatBytes(loaded)} downloaded`;
      },
    });
    ensureCurrentOperation(controller);
    await renderPresentation(source, controller);
  } catch (error) {
    handleOpenError(error, controller);
  }
}

async function renderPresentation(source, controller) {
  elements.loadingTitle.textContent = 'Preparing the viewer';
  elements.loadingDetail.textContent = 'Loading the rendering engine...';
  setIndeterminateProgress();
  const { PptxViewer, RECOMMENDED_ZIP_LIMITS } = await loadRenderer();
  ensureCurrentOperation(controller);

  destroyViewer();
  elements.viewerContainer.replaceChildren();
  elements.viewerContainer.setAttribute('aria-busy', 'true');
  elements.app.classList.remove('is-empty');
  elements.themeColor.setAttribute('content', '#10131a');
  elements.loadingTitle.textContent = 'Rendering presentation';
  elements.loadingDetail.textContent = 'Building the document view...';

  const nextViewer = await PptxViewer.open(source.buffer, elements.viewerContainer, {
    signal: controller.signal,
    renderMode: 'list',
    listOptions: DOCUMENT_LIST_OPTIONS,
    fitMode: 'contain',
    zoomPercent: 100,
    zipLimits: RECOMMENDED_ZIP_LIMITS,
    lazySlides: true,
    lazyMedia: true,
    scrollContainer: elements.stage,
    onSlideChange: updateCurrentSlide,
    onSlideError: (index, error) => {
      console.warn(`Slide ${index + 1} failed to render`, error);
      setStatus(`Slide ${index + 1} contains unsupported content.`);
    },
    onNodeError: (nodeId, error) => {
      console.warn(`Presentation node ${nodeId} failed to render`, error);
    },
  });

  ensureCurrentOperation(controller);
  viewer = nextViewer;
  elements.app.style.setProperty('--deck-aspect', `${viewer.slideWidth} / ${viewer.slideHeight}`);
  elements.viewerContainer.removeAttribute('aria-busy');
  elements.documentTitle.textContent = source.name;
  elements.documentTitle.title = source.name;
  elements.slideCount.textContent = String(viewer.slideCount);
  elements.zoomValue.value = `${viewer.zoomPercent}%`;
  setViewerEnabled(true);
  updateCurrentSlide(viewer.currentSlideIndex);
  createThumbnails();
  updateCurrentSlide(viewer.currentSlideIndex);
  finishOperation(controller);
  setStatus(`${source.name} - ${viewer.slideCount} slide${viewer.slideCount === 1 ? '' : 's'}`);
}

function beginOperation(title, detail) {
  cancelOperation();
  const controller = new AbortController();
  operationController = controller;
  elements.loadingTitle.textContent = title;
  elements.loadingDetail.textContent = detail;
  elements.loading.hidden = false;
  elements.stage.setAttribute('aria-busy', 'true');
  elements.loadingProgress.hidden = false;
  elements.loadingProgress.value = 0;
  return controller;
}

function finishOperation(controller) {
  if (operationController !== controller) return;
  operationController = null;
  elements.loading.hidden = true;
  elements.stage.removeAttribute('aria-busy');
}

function cancelOperation() {
  operationController?.abort();
  operationController = null;
  elements.loading.hidden = true;
  elements.stage.removeAttribute('aria-busy');
}

function ensureCurrentOperation(controller) {
  if (operationController !== controller || controller.signal.aborted) {
    throw new DOMException('Opening was cancelled.', 'AbortError');
  }
}

function handleOpenError(error, controller) {
  if (operationController !== controller && error?.name === 'AbortError') return;
  finishOperation(controller);
  if (error?.name === 'AbortError') {
    setStatus('Opening cancelled');
    if (!viewer) elements.app.classList.add('is-empty');
    return;
  }
  console.error(error);
  showError(error instanceof Error ? error.message : 'The file could not be opened.');
  setStatus('Could not open presentation');
  if (!viewer) elements.app.classList.add('is-empty');
}

function destroyViewer() {
  disposeHighlight();
  searchResults = [];
  searchIndex = -1;
  thumbnailObserver?.disconnect();
  thumbnailObserver = null;
  for (const handle of thumbnailHandles.values()) handle?.dispose?.();
  thumbnailHandles.clear();
  viewer?.destroy?.();
  viewer = null;
  elements.app.style.removeProperty('--deck-aspect');
  elements.thumbnailList.replaceChildren();
  elements.slideCount.textContent = '0';
  elements.slideNumber.value = '1';
  setViewerEnabled(false);
  resetSearchUi();
}

function createThumbnails() {
  const fragment = document.createDocumentFragment();
  thumbnailObserver = new IntersectionObserver(handleThumbnailVisibility, {
    root: elements.thumbnailList,
    rootMargin: '360px 0px',
  });

  for (let index = 0; index < viewer.slideCount; index += 1) {
    const button = document.createElement('button');
    button.className = 'thumbnail-button';
    button.type = 'button';
    button.dataset.index = String(index);
    button.setAttribute('role', 'option');
    button.setAttribute('aria-label', `Go to slide ${index + 1}`);
    button.setAttribute('aria-selected', String(index === viewer.currentSlideIndex));

    const preview = document.createElement('span');
    preview.className = 'thumbnail-preview';
    const number = document.createElement('span');
    number.className = 'thumbnail-number';
    number.textContent = String(index + 1).padStart(2, '0');
    button.append(preview, number);
    button.addEventListener('click', () => goToSlide(index));
    fragment.append(button);
    thumbnailObserver.observe(preview);
  }
  elements.thumbnailList.append(fragment);
}

function handleThumbnailVisibility(entries) {
  for (const entry of entries) {
    if (!entry.isIntersecting || !viewer) continue;
    const button = entry.target.closest('.thumbnail-button');
    const index = Number(button?.dataset.index);
    if (!Number.isInteger(index) || thumbnailHandles.has(index)) continue;
    try {
      const width = Math.max(96, Math.floor(entry.target.clientWidth));
      const handle = viewer.renderThumbnailToContainer(index, entry.target, { width });
      thumbnailHandles.set(index, handle);
      handle?.ready?.catch((error) => {
        console.warn(`Thumbnail ${index + 1} failed to render`, error);
        button?.classList.add('thumbnail-failed');
      });
    } catch (error) {
      console.warn(`Thumbnail ${index + 1} failed to render`, error);
      button?.classList.add('thumbnail-failed');
    }
    thumbnailObserver.unobserve(entry.target);
  }
}

async function goToSlide(index) {
  if (!viewer) return;
  const target = Math.max(0, Math.min(viewer.slideCount - 1, index));
  try {
    await viewer.goToSlide(target);
    updateCurrentSlide(target);
  } catch (error) {
    showError(error instanceof Error ? error.message : 'That slide could not be rendered.');
  }
}

function navigateRelative(offset) {
  if (viewer) goToSlide(viewer.currentSlideIndex + offset);
}

function goToTypedSlide() {
  if (!viewer) return;
  const requested = Number(elements.slideNumber.value);
  if (!Number.isInteger(requested) || requested < 1 || requested > viewer.slideCount) {
    elements.slideNumber.value = String(viewer.currentSlideIndex + 1);
    setStatus(`Enter a slide number between 1 and ${viewer.slideCount}.`);
    return;
  }
  goToSlide(requested - 1);
}

function updateCurrentSlide(index) {
  if (!viewer && !Number.isInteger(index)) return;
  const current = Number.isInteger(index) ? index : viewer.currentSlideIndex;
  const count = viewer?.slideCount ?? 0;
  elements.slideNumber.value = String(current + 1);
  elements.previousButton.disabled = current <= 0;
  elements.nextButton.disabled = current >= count - 1;

  for (const button of elements.thumbnailList.querySelectorAll('.thumbnail-button')) {
    const active = Number(button.dataset.index) === current;
    button.classList.toggle('is-active', active);
    button.setAttribute('aria-selected', String(active));
    if (active) button.scrollIntoView({ block: 'nearest' });
  }
}

async function setZoom(value) {
  if (!viewer || !Number.isFinite(value)) return;
  const zoom = Math.max(25, Math.min(250, Math.round(value / 10) * 10));
  try {
    await viewer.setZoom(zoom);
    elements.zoomValue.value = `${viewer.zoomPercent}%`;
  } catch (error) {
    showError(error instanceof Error ? error.message : 'Zoom could not be changed.');
  }
}

async function fitSlide() {
  if (!viewer) return;
  try {
    await viewer.setFitMode('contain');
    await viewer.setZoom(100);
    elements.zoomValue.value = '100%';
  } catch (error) {
    showError(error instanceof Error ? error.message : 'The slide could not be fitted.');
  }
}

async function runSearch(direction) {
  if (!viewer) return;
  const query = elements.searchInput.value.trim();
  if (!query) {
    elements.searchInput.focus();
    return;
  }

  const sameQuery = elements.searchInput.dataset.query === query;
  if (!sameQuery) {
    disposeHighlight();
    searchResults = viewer.searchText(query, { matchCase: false, wholeWord: false, snippetRadius: 48 });
    searchIndex = direction < 0 ? 0 : -1;
    elements.searchInput.dataset.query = query;
  }

  if (searchResults.length === 0) {
    elements.searchCount.value = 'No results';
    elements.searchPrev.disabled = true;
    elements.searchNext.disabled = true;
    setStatus(`No results for "${query}".`);
    return;
  }

  searchIndex = (searchIndex + direction + searchResults.length) % searchResults.length;
  const result = searchResults[searchIndex];
  elements.searchCount.value = `${searchIndex + 1} / ${searchResults.length}`;
  elements.searchPrev.disabled = false;
  elements.searchNext.disabled = false;
  await goToSlide(result.slideIndex);
  disposeHighlight();
  activeHighlight = await viewer.highlightSearchResult(result);
  setStatus(`Result ${searchIndex + 1} of ${searchResults.length} on slide ${result.slideIndex + 1}.`);
}

function clearSearchIfEmpty() {
  if (elements.searchInput.value) {
    if (elements.searchInput.dataset.query !== elements.searchInput.value.trim()) {
      elements.searchCount.value = '';
    }
    return;
  }
  resetSearchUi();
  disposeHighlight();
}

function resetSearchUi() {
  elements.searchInput.value = '';
  delete elements.searchInput.dataset.query;
  elements.searchCount.value = '';
  elements.searchPrev.disabled = true;
  elements.searchNext.disabled = true;
  searchResults = [];
  searchIndex = -1;
}

function disposeHighlight() {
  activeHighlight?.dispose?.();
  activeHighlight = null;
  viewer?.clearSearchHighlights?.();
}

function setViewerEnabled(enabled) {
  for (const control of [
    elements.sidebarToggle,
    elements.searchInput,
    elements.slideNumber,
    elements.zoomOut,
    elements.zoomIn,
    elements.fitButton,
    elements.fullscreenButton,
  ]) {
    control.disabled = !enabled;
  }
  if (!enabled) {
    elements.previousButton.disabled = true;
    elements.nextButton.disabled = true;
  }
}

function toggleSidebar() {
  const isClosed = elements.app.classList.toggle('sidebar-closed');
  const isExpanded = !isClosed;
  const label = isExpanded ? 'Hide slide thumbnails' : 'Show slide thumbnails';
  elements.sidebarToggle.setAttribute('aria-expanded', String(isExpanded));
  elements.sidebarToggle.setAttribute('aria-label', label);
  elements.sidebarToggleTooltip.textContent = label;
}

async function toggleFullscreen() {
  try {
    if (document.fullscreenElement) await document.exitFullscreen();
    else await elements.stage.requestFullscreen();
  } catch (error) {
    showError(error instanceof Error ? error.message : 'Fullscreen is not available.');
  }
}

function updateFullscreenButton() {
  const active = Boolean(document.fullscreenElement);
  elements.fullscreenButton.setAttribute('aria-label', active ? 'Exit fullscreen' : 'Enter fullscreen');
}

function handleKeyboard(event) {
  const target = event.target;
  const typing = target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement || target instanceof HTMLSelectElement;
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'o') {
    event.preventDefault();
    elements.fileInput.click();
    return;
  }
  if (event.key === '/' && viewer && !typing) {
    event.preventDefault();
    elements.searchInput.focus();
    return;
  }
  if (!viewer || typing || elements.urlDialog.open) return;

  if (['ArrowRight', 'PageDown', ' '].includes(event.key)) {
    event.preventDefault();
    navigateRelative(1);
  } else if (['ArrowLeft', 'PageUp'].includes(event.key)) {
    event.preventDefault();
    navigateRelative(-1);
  } else if (event.key === 'Home') {
    event.preventDefault();
    goToSlide(0);
  } else if (event.key === 'End') {
    event.preventDefault();
    goToSlide(viewer.slideCount - 1);
  } else if (event.key.toLowerCase() === 'f') {
    event.preventDefault();
    toggleFullscreen();
  }
}

function handlePointerDown(event) {
  if (event.pointerType === 'mouse' || !viewer) return;
  pointerStart = { x: event.clientX, y: event.clientY, id: event.pointerId };
}

function handlePointerUp(event) {
  if (!pointerStart || event.pointerId !== pointerStart.id) return;
  const dx = event.clientX - pointerStart.x;
  const dy = event.clientY - pointerStart.y;
  pointerStart = null;
  if (Math.abs(dx) > 60 && Math.abs(dx) > Math.abs(dy) * 1.3) {
    navigateRelative(dx < 0 ? 1 : -1);
  }
}

function handleDragEnter(event) {
  if (!hasFileDrag(event)) return;
  event.preventDefault();
  dragDepth += 1;
  elements.dropLayer.hidden = false;
}

function handleDragOver(event) {
  if (!hasFileDrag(event)) return;
  event.preventDefault();
  event.dataTransfer.dropEffect = 'copy';
}

function handleDragLeave(event) {
  if (!hasFileDrag(event)) return;
  event.preventDefault();
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0) elements.dropLayer.hidden = true;
}

async function handleDrop(event) {
  if (!hasFileDrag(event)) return;
  event.preventDefault();
  dragDepth = 0;
  elements.dropLayer.hidden = true;
  const [file] = event.dataTransfer.files ?? [];
  if (file) await openLocalFile(file);
}

function hasFileDrag(event) {
  return Array.from(event.dataTransfer?.types ?? []).includes('Files');
}

function showUrlDialog() {
  elements.urlDialog.showModal();
  requestAnimationFrame(() => elements.urlInput.focus());
}

function closeUrlDialog() {
  elements.urlDialog.close();
}

function showError(message) {
  elements.errorMessage.textContent = message;
  elements.errorToast.hidden = false;
}

function hideError() {
  elements.errorToast.hidden = true;
}

function updateLoadingProgress({ loaded, total }) {
  if (total) {
    elements.loadingProgress.hidden = false;
    elements.loadingProgress.max = total;
    elements.loadingProgress.value = loaded;
  } else {
    setIndeterminateProgress();
  }
}

function setIndeterminateProgress() {
  elements.loadingProgress.hidden = false;
  elements.loadingProgress.removeAttribute('value');
}

function setStatus(message) {
  elements.status.textContent = message;
}

function prefetchRenderer() {
  const warm = () => loadRenderer().catch(() => {});
  if ('requestIdleCallback' in window) window.requestIdleCallback(warm, { timeout: 2500 });
  else window.setTimeout(warm, 1500);
}

function loadRenderer() {
  rendererPromise ??= import('@aiden0z/pptx-renderer');
  return rendererPromise;
}

function openInitialUrl() {
  const value = new URLSearchParams(location.search).get('url');
  if (value) openRemoteUrl(value);
}

function byId(id) {
  const element = document.getElementById(id);
  if (!element) throw new Error(`Missing #${id}`);
  return element;
}
