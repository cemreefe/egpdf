// Printing: render the document exactly as saving would produce it (form
// values baked, redactions rasterized, all edits applied) into high-res page
// images, show them in an M365-style print screen (settings panel + live
// preview), then print silently with the chosen options.
import { loadPdf, pdfjsLib } from './viewer.js';
import { buildSavedPdf } from './save.js';

const PRINT_DPI = 300;
const MAX_RENDER_DIM = 4500; // cap oversized pages (A0 plans etc.)
const $ = (id) => document.getElementById(id);

export async function preparePrint(tab, onProgress) {
  const bytes = await buildSavedPdf(tab);
  const pdf = await loadPdf(new Uint8Array(bytes));
  const container = $('print-container');
  container.replaceChildren();
  const total = pdf.numPages;
  const pages = []; // {src, landscape, rotatedSrc}
  try {
    for (let n = 1; n <= total; n++) {
      onProgress?.(n, total);
      const page = await pdf.getPage(n);
      const base = page.getViewport({ scale: 1 });
      const scale = Math.min(PRINT_DPI / 72, MAX_RENDER_DIM / Math.max(base.width, base.height));
      const viewport = page.getViewport({ scale });
      const canvas = document.createElement('canvas');
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      const ctx = canvas.getContext('2d', { alpha: false });
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({
        canvasContext: ctx,
        viewport,
        intent: 'print',
        annotationMode: pdfjsLib.AnnotationMode.ENABLE,
      }).promise;
      const src = canvas.toDataURL('image/png');
      pages.push({ src, landscape: base.width > base.height, rotatedSrc: null });
      const div = document.createElement('div');
      div.className = 'print-page';
      const img = new Image();
      img.src = src;
      div.appendChild(img);
      container.appendChild(div);
      canvas.width = 0; canvas.height = 0; // release backing store early
    }
  } finally {
    pdf.destroy();
  }
  return pages;
}

// Rotate a page image 90° counterclockwise — the printer convention for
// fitting landscape content onto a portrait sheet (content top lands on the
// sheet's left edge; turn the sheet clockwise to read).
function rotate90(src) {
  return new Promise((resolve, reject) => {
    const im = new Image();
    im.onload = () => {
      const c = document.createElement('canvas');
      c.width = im.naturalHeight;
      c.height = im.naturalWidth;
      const ctx = c.getContext('2d', { alpha: false });
      ctx.translate(0, c.height);
      ctx.rotate(-Math.PI / 2);
      ctx.drawImage(im, 0, 0);
      resolve(c.toDataURL('image/png'));
    };
    im.onerror = reject;
    im.src = src;
  });
}

export function clearPrint() {
  $('print-container').replaceChildren();
}

// ---- preview state -----------------------------------------------------------

let state = null;      // { pages: [{src, landscape, rotatedSrc}], page, total }
let statusCb = () => {};
let wired = false;

// One print job = one sheet orientation. "Auto" uses the document's dominant
// orientation; pages that don't match the sheet get rotated 90° to fill it.
function sheetIsLandscape() {
  const v = $('pp-orient').value;
  if (v === 'landscape') return true;
  if (v === 'portrait') return false;
  const landCount = state ? state.pages.filter((p) => p.landscape).length : 0;
  return state ? landCount > state.pages.length / 2 : false;
}

async function pageSrcForSheet(i) {
  const p = state.pages[i];
  if (p.landscape === sheetIsLandscape()) return p.src;
  if (!p.rotatedSrc) p.rotatedSrc = await rotate90(p.src);
  return p.rotatedSrc;
}

async function updatePreview() {
  if (!state) return;
  const cur = state.page;
  const src = await pageSrcForSheet(cur - 1);
  if (!state || state.page !== cur) return; // navigated away meanwhile
  $('pp-page-img').src = src;
  $('pp-pageinfo').textContent = `${state.page} of ${state.total}`;
  $('pp-prev').disabled = state.page <= 1;
  $('pp-next').disabled = state.page >= state.total;
}

// Exposed for the e2e suite.
export async function getPrintState() {
  if (!state) return null;
  return {
    orientations: state.pages.map((p) => p.landscape),
    sheetLandscape: sheetIsLandscape(),
    page1Rotated: (await pageSrcForSheet(0)) !== state.pages[0].src,
  };
}

export function isPrintPreviewOpen() {
  return !$('print-overlay').classList.contains('hidden');
}

export function closePrintPreview() {
  $('print-overlay').classList.add('hidden');
  clearPrint();
  $('pp-page-img').removeAttribute('src');
  state = null;
}

function parseRange(text, total) {
  const out = new Set();
  for (const part of String(text).split(',')) {
    const m = part.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/);
    if (!m) continue;
    let a = +m[1], b = m[2] ? +m[2] : +m[1];
    if (a > b) [a, b] = [b, a];
    for (let i = Math.max(1, a); i <= Math.min(total, b); i++) out.add(i);
  }
  return out;
}

async function doPrint() {
  if (!state) return;
  let range = null;
  if ($('pp-range-mode').value === 'custom') {
    range = parseRange($('pp-range').value, state.total);
    if (!range.size) {
      statusCb('Enter a valid page range, e.g. 1-3, 5', true);
      return;
    }
  }
  const divs = document.querySelectorAll('#print-container .print-page');
  divs.forEach((d, i) => { d.style.display = !range || range.has(i + 1) ? '' : 'none'; });

  // Match every page image to the chosen sheet orientation before printing.
  statusCb('Preparing pages…');
  const imgs = document.querySelectorAll('#print-container .print-page img');
  for (let i = 0; i < state.pages.length; i++) {
    if (!range || range.has(i + 1)) imgs[i].src = await pageSrcForSheet(i);
  }

  const opts = {
    deviceName: $('pp-printer').value,
    copies: Math.max(1, Math.min(99, parseInt($('pp-copies').value, 10) || 1)),
    collate: $('pp-collate').value === '1',
    landscape: sheetIsLandscape(),
    color: $('pp-color').value === '1',
    duplexMode: $('pp-duplex').value,
    margins: { marginType: $('pp-scale').value === 'actual' ? 'none' : 'default' },
  };
  if ($('pp-paper').value) opts.pageSize = $('pp-paper').value;

  statusCb('Printing…');
  const res = await window.native.printNow(opts);
  if (res?.ok) {
    statusCb('Sent to printer');
    closePrintPreview();
  } else {
    statusCb(`Print failed: ${res?.reason || 'unknown error'}`, true);
    divs.forEach((d) => { d.style.display = ''; });
  }
}

function wireOnce() {
  if (wired) return;
  wired = true;
  $('pp-prev').addEventListener('click', () => { if (state) { state.page = Math.max(1, state.page - 1); updatePreview(); } });
  $('pp-next').addEventListener('click', () => { if (state) { state.page = Math.min(state.total, state.page + 1); updatePreview(); } });
  $('pp-zoom').addEventListener('input', (e) => { $('pp-page-img').style.width = e.target.value + '%'; });
  $('pp-orient').addEventListener('change', () => updatePreview());
  $('pp-range-mode').addEventListener('change', (e) => {
    $('pp-range').classList.toggle('hidden', e.target.value !== 'custom');
    if (e.target.value === 'custom') $('pp-range').focus();
  });
  $('print-close').addEventListener('click', closePrintPreview);
  $('pp-print').addEventListener('click', () => doPrint());
  window.addEventListener('keydown', (e) => {
    if (!isPrintPreviewOpen()) return;
    if (e.key === 'Escape') { e.stopImmediatePropagation(); closePrintPreview(); }
    else if (e.key === 'ArrowLeft') $('pp-prev').click();
    else if (e.key === 'ArrowRight') $('pp-next').click();
  }, true);
}

export async function openPrintPreview(tab, setStatus = () => {}) {
  statusCb = setStatus;
  const pages = await preparePrint(tab, (n, t) => setStatus(`Preparing preview… page ${n} / ${t}`));
  setStatus('');
  state = { pages, page: 1, total: pages.length };

  const sel = $('pp-printer');
  sel.replaceChildren();
  let printers = [];
  try { printers = await window.native.listPrinters(); } catch {}
  for (const p of printers) {
    const o = document.createElement('option');
    o.value = p.name;
    o.textContent = p.displayName || p.name;
    if (p.isDefault) o.selected = true;
    sel.appendChild(o);
  }
  if (!printers.length) {
    const o = document.createElement('option');
    o.textContent = 'No printers found';
    sel.appendChild(o);
  }
  $('pp-print').disabled = printers.length === 0;

  wireOnce();
  $('print-overlay').classList.remove('hidden');
  await updatePreview();
  return { pages: state.total, printers: printers.length };
}
