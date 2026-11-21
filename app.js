/* MyPDF - all processing happens client-side */
(() => {
'use strict';
const { PDFDocument, rgb, StandardFonts } = PDFLib;
pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const CFG = window.MYPDF_CONFIG || {};

/* Utilities */
const loader = $('#loader'), loaderText = $('#loader-text');
function busy(on, text = 'Processing…') { loaderText.textContent = text; loader.hidden = !on; }
let toastTimer;
function toast(msg, isError = false) {
  const t = $('#toast'); t.textContent = msg; t.className = 'toast' + (isError ? ' error' : ''); t.hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => (t.hidden = true), 3500);
  t.style.animation = 'none'; void t.offsetWidth; t.style.animation = '';
  if (!isError && /ready|merged|Done/i.test(msg)) celebrate();
}
function celebrate() {
  if (matchMedia('(prefers-reduced-motion: reduce)').matches) return;
  const colors = ['#22d3ee', '#8b5cf6', '#ec4899', '#34d399', '#f59e0b', '#ffffff'];
  const ox = innerWidth / 2, oy = innerHeight - 70;
  for (let i = 0; i < 40; i++) {
    const s = document.createElement('span'); s.className = 'confetti';
    const a = -Math.PI / 2 + (Math.random() - .5) * 2.2, d = 180 + Math.random() * 320;
    Object.assign(s.style, { left: ox + 'px', top: oy + 'px', background: colors[i % colors.length], color: colors[i % colors.length] });
    s.style.setProperty('--dx', Math.cos(a) * d + 'px'); s.style.setProperty('--dy', Math.sin(a) * d + 120 + 'px');
    document.body.appendChild(s); setTimeout(() => s.remove(), 1400);
  }
}
function download(data, filename, type = 'application/pdf') {
  const blob = data instanceof Blob ? data : new Blob([data], { type });
  const url = URL.createObjectURL(blob);
  const a = Object.assign(document.createElement('a'), { href: url, download: filename });
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}
const baseName = n => n.replace(/\.[^.]+$/, '');
const fmtSize = b => b < 1024 ? b + ' B' : b < 1048576 ? (b / 1024).toFixed(1) + ' KB' : (b / 1048576).toFixed(1) + ' MB';
const isPdf = f => f.type === 'application/pdf' || /\.pdf$/i.test(f.name);
const isImg = f => f.type.startsWith('image/');
const nextFrame = () => new Promise(r => requestAnimationFrame(() => r()));

async function loadPdfLib(bytes) {
  try { return await PDFDocument.load(bytes, { ignoreEncryption: true }); }
  catch (e) { throw new Error('This PDF could not be read. It may be damaged or password-protected.'); }
}
async function renderPdfPage(pdf, num, targetWidth) {
  const page = await pdf.getPage(num);
  const v1 = page.getViewport({ scale: 1 });
  const vp = page.getViewport({ scale: targetWidth / v1.width });
  const c = document.createElement('canvas');
  c.width = Math.round(vp.width); c.height = Math.round(vp.height);
  await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
  return c;
}

/* Drop zone wiring */
function setupDrop(zoneSel, inputSel, accept, onFiles) {
  const zone = $(zoneSel), input = $(inputSel);
  input.addEventListener('change', () => { if (input.files.length) onFiles([...input.files]); input.value = ''; });
  ['dragenter', 'dragover'].forEach(ev => zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach(ev => zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.remove('drag'); }));
  zone.addEventListener('drop', e => {
    const files = [...e.dataTransfer.files].filter(accept);
    if (!files.length) return toast('Unsupported file type', true);
    onFiles(files);
  });
}

/* Reorderable card list (drag & drop) */
function makeSortable(container, items, rerender) {
  let from = null;
  $$('.file-card', container).forEach((card, i) => {
    card.draggable = true;
    card.addEventListener('dragstart', e => { from = i; card.classList.add('dragging'); e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', ''); });
    card.addEventListener('dragend', () => card.classList.remove('dragging'));
    card.addEventListener('dragover', e => { e.preventDefault(); e.stopPropagation(); card.classList.add('over'); });
    card.addEventListener('dragleave', () => card.classList.remove('over'));
    card.addEventListener('drop', e => {
      e.preventDefault(); e.stopPropagation(); card.classList.remove('over');
      if (from === null || from === i) return;
      const [m] = items.splice(from, 1); items.splice(i, 0, m); from = null; rerender();
    });
  });
}

/* Router */
const views = $$('.view');
function route() {
  let id = (location.hash || '#home').slice(1);
  if (!document.getElementById(id) || !document.getElementById(id).classList.contains('view')) id = 'home';
  views.forEach(v => { const show = v.id === id; v.hidden = !show; v.classList.remove('enter'); if (show) { void v.offsetWidth; v.classList.add('enter'); } });
  $$('.nav a').forEach(a => a.classList.toggle('active', a.getAttribute('href') === '#' + id));
  document.body.classList.remove('nav-open');
  const titles = { home: 'MyPDF – Free Online PDF Tools', 'jpg-to-pdf': 'JPG to PDF – MyPDF', 'pdf-to-jpg': 'PDF to JPG – MyPDF', merge: 'Merge PDF – MyPDF', split: 'Split PDF – MyPDF', sign: 'Sign PDF – MyPDF' };
  const h1 = document.getElementById(id).querySelector('h1');
  document.title = titles[id] || ((h1 ? h1.textContent.trim() : 'PDF tools') + ' – MyPDF');
  window.scrollTo(0, 0);
}
window.addEventListener('hashchange', route);
document.addEventListener('click', e => {
  const a = e.target.closest('a[data-scroll]'); if (!a) return;
  e.preventDefault(); const id = a.getAttribute('href').slice(1);
  const go = () => { const el = document.getElementById(id); if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' }); };
  if ($('#home').hidden) { location.hash = '#home'; setTimeout(go, 80); } else go();
});
{ const tc = document.getElementById('tool-count'); if (tc) tc.dataset.count = document.querySelectorAll('#tools-grid .tool-card').length; }
route();
$('.menu-btn').addEventListener('click', () => document.body.classList.toggle('nav-open'));
$('#year').textContent = new Date().getFullYear();
if (CFG.VERSION) $('#app-version').textContent = 'v' + CFG.VERSION;

/* Theme switch */
const root = document.documentElement;
function setTheme(t) {
  root.dataset.theme = t;
  try { localStorage.setItem('mypdf-theme', t); } catch (e) {}
  const m = document.querySelector('meta[name=theme-color]'); if (m) m.content = t === 'dark' ? '#05060f' : '#f3f5fc';
  window.dispatchEvent(new CustomEvent('mypdf-theme', { detail: t }));
}
$('#theme-toggle').addEventListener('click', e => {
  const next = root.dataset.theme === 'dark' ? 'light' : 'dark';
  const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
  if (!document.startViewTransition || reduce) return setTheme(next);
  const r = e.currentTarget.getBoundingClientRect(), x = r.left + r.width / 2, y = r.top + r.height / 2;
  const end = Math.hypot(Math.max(x, innerWidth - x), Math.max(y, innerHeight - y));
  const vt = document.startViewTransition(() => setTheme(next));
  vt.ready.then(() => root.animate({ clipPath: [`circle(0px at ${x}px ${y}px)`, `circle(${end}px at ${x}px ${y}px)`] },
    { duration: 650, easing: 'cubic-bezier(.6,0,.3,1)', pseudoElement: '::view-transition-new(root)' })).catch(() => {});
});

/* Scroll reveal */
const io = 'IntersectionObserver' in window ? new IntersectionObserver(entries => entries.forEach(en => {
  if (en.isIntersecting) { const t = en.target; t.classList.add('in'); io.unobserve(t); setTimeout(() => (t.style.transitionDelay = ''), 1000); }
}), { threshold: .12 }) : null;
$$('.reveal').forEach((el, i) => { if (io) { el.style.transitionDelay = (i % 4) * 70 + 'ms'; io.observe(el); } else el.classList.add('in'); });

/* Session stats (this device only) */
const SES_KEY = 'mypdf-stats';
let ses = { files: 0, pages: 0, bytes: 0 };
try { ses = Object.assign(ses, JSON.parse(localStorage.getItem(SES_KEY) || '{}')); } catch (e) {}
function renderSes(bump) {
  const set = (id, v) => { const el = $(id); if (!el) return; el.textContent = v; if (bump) { el.classList.remove('bump'); void el.offsetWidth; el.classList.add('bump'); } };
  set('#ses-files', ses.files.toLocaleString()); set('#ses-pages', ses.pages.toLocaleString()); set('#ses-bytes', fmtSize(ses.bytes));
}
function track(files, pages, bytes) {
  ses.files += files; ses.pages += pages; ses.bytes += bytes;
  try { localStorage.setItem(SES_KEY, JSON.stringify(ses)); } catch (e) {}
  renderSes(true);
}
renderSes(false);
$('#ses-reset').onclick = () => { ses = { files: 0, pages: 0, bytes: 0 }; try { localStorage.removeItem(SES_KEY); } catch (e) {} renderSes(true); };
const sumSize = arr => arr.reduce((s, f) => s + f.size, 0);

/* JPG → PDF */
const images = []; // {file, url}
setupDrop('#img-drop', '#img-input', isImg, files => {
  files.filter(isImg).forEach(file => images.push({ file, url: URL.createObjectURL(file) }));
  renderImages();
});
function renderImages() {
  const list = $('#img-list');
  list.innerHTML = images.map((im, i) => `
    <div class="file-card" style="--i:${i}"><span class="badge">${i + 1}</span>
      <button class="remove" data-i="${i}" title="Remove">✕</button>
      <div class="thumb"><img src="${im.url}" alt=""></div>
      <div class="name" title="${im.file.name}">${im.file.name}</div>
      <div class="meta">${fmtSize(im.file.size)}</div>
    </div>`).join('');
  $$('.remove', list).forEach(b => b.onclick = () => { URL.revokeObjectURL(images[b.dataset.i].url); images.splice(b.dataset.i, 1); renderImages(); });
  makeSortable(list, images, renderImages);
  $('#img-options').hidden = !images.length;
  $('#img-drop').classList.toggle('compact', images.length > 0);
}
async function imageToJpegBytes(file) {
  // Re-encode through canvas: handles EXIF rotation, WEBP/GIF/BMP, and transparency (white bg)
  let bmp;
  try { bmp = await createImageBitmap(file, { imageOrientation: 'from-image' }); }
  catch { bmp = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = URL.createObjectURL(file); }); }
  const c = document.createElement('canvas'); c.width = bmp.width; c.height = bmp.height;
  const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height); ctx.drawImage(bmp, 0, 0);
  const blob = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.92));
  return { bytes: await blob.arrayBuffer(), w: c.width, h: c.height };
}
$('#img-go').onclick = async () => {
  if (!images.length) return;
  busy(true, 'Creating PDF…');
  try {
    const size = $('#img-size').value, orient = $('#img-orient').value, margin = +$('#img-margin').value;
    const doc = await PDFDocument.create();
    for (let i = 0; i < images.length; i++) {
      loaderText.textContent = `Adding image ${i + 1} of ${images.length}…`; await nextFrame();
      const { bytes, w, h } = await imageToJpegBytes(images[i].file);
      const img = await doc.embedJpg(bytes);
      let pw, ph;
      if (size === 'fit') { pw = w * 0.75 + margin * 2; ph = h * 0.75 + margin * 2; }
      else {
        [pw, ph] = size === 'a4' ? [595.28, 841.89] : [612, 792];
        const land = orient === 'landscape' || (orient === 'auto' && w > h);
        if (land) [pw, ph] = [ph, pw];
      }
      const page = doc.addPage([pw, ph]);
      const aw = pw - margin * 2, ah = ph - margin * 2;
      const s = Math.min(aw / w, ah / h);
      const dw = w * s, dh = h * s;
      page.drawImage(img, { x: (pw - dw) / 2, y: (ph - dh) / 2, width: dw, height: dh });
    }
    const out = await doc.save();
    download(out, images.length === 1 ? baseName(images[0].file.name) + '.pdf' : 'mypdf-images.pdf');
    track(images.length, images.length, sumSize(images.map(x => x.file)));
    toast('Your PDF is ready!');
  } catch (e) { console.error(e); toast(e.message || 'Something went wrong', true); }
  busy(false);
};

/* PDF → JPG */
let p2jFile = null, p2jOutputs = [];
setupDrop('#p2j-drop', '#p2j-input', isPdf, files => {
  p2jFile = files[0]; p2jOutputs = [];
  $('#p2j-name').textContent = '📄 ' + p2jFile.name;
  $('#p2j-options').hidden = false; $('#p2j-result').hidden = true; $('#p2j-pages').innerHTML = '';
  $('#p2j-drop').classList.add('compact');
});
$('#p2j-go').onclick = async () => {
  if (!p2jFile) return;
  busy(true, 'Reading PDF…');
  try {
    const fmt = $('#p2j-format').value, scale = +$('#p2j-scale').value, ext = fmt === 'jpeg' ? 'jpg' : 'png';
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await p2jFile.arrayBuffer()) }).promise;
    p2jOutputs.forEach(o => URL.revokeObjectURL(o.url)); p2jOutputs = [];
    const grid = $('#p2j-pages'); grid.innerHTML = '';
    for (let n = 1; n <= pdf.numPages; n++) {
      loaderText.textContent = `Converting page ${n} of ${pdf.numPages}…`; await nextFrame();
      const page = await pdf.getPage(n);
      const vp = page.getViewport({ scale });
      const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
      const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
      await page.render({ canvasContext: ctx, viewport: vp }).promise;
      const blob = await new Promise(r => c.toBlob(r, 'image/' + fmt, 0.92));
      const name = `${baseName(p2jFile.name)}-page-${n}.${ext}`;
      const url = URL.createObjectURL(blob);
      p2jOutputs.push({ name, blob, url });
      grid.insertAdjacentHTML('beforeend', `<div class="file-card" style="--i:0"><span class="badge">${n}</span>
        <div class="thumb"><img src="${url}" alt="Page ${n}"></div>
        <div class="name">${name}</div><div class="meta">${c.width}×${c.height} · ${fmtSize(blob.size)}</div>
        <a class="dl" href="#">⬇ Download</a></div>`);
      grid.lastElementChild.querySelector('.dl').onclick = ev => { ev.preventDefault(); const o = p2jOutputs[n - 1]; download(o.blob, o.name); };
    }
    $('#p2j-count').textContent = `✅ ${pdf.numPages} image${pdf.numPages > 1 ? 's' : ''} ready`;
    $('#p2j-result').hidden = false;
    track(1, pdf.numPages, p2jFile.size);
    if (pdf.numPages === 1) download(p2jOutputs[0].blob, p2jOutputs[0].name);
  } catch (e) { console.error(e); toast('This PDF could not be read. It may be damaged or password-protected.', true); }
  busy(false);
};
$('#p2j-zip').onclick = async () => {
  busy(true, 'Creating ZIP…');
  const zip = new JSZip(); p2jOutputs.forEach(o => zip.file(o.name, o.blob));
  download(await zip.generateAsync({ type: 'blob' }), baseName(p2jFile.name) + '-images.zip', 'application/zip');
  busy(false);
};

/* MERGE */
const mergeFiles = []; // {file, thumb (canvas|null), pages}
setupDrop('#merge-drop', '#merge-input', isPdf, async files => {
  busy(true, 'Loading PDFs…');
  for (const file of files.filter(isPdf)) {
    const item = { file, thumb: null, pages: '?' };
    try {
      const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
      item.pages = pdf.numPages; item.thumb = (await renderPdfPage(pdf, 1, 220)).toDataURL();
    } catch { item.pages = '⚠'; }
    mergeFiles.push(item);
  }
  busy(false); renderMerge();
});
function renderMerge() {
  const list = $('#merge-list');
  list.innerHTML = mergeFiles.map((m, i) => `
    <div class="file-card" style="--i:${i}"><span class="badge">${i + 1}</span>
      <button class="remove" data-i="${i}" title="Remove">✕</button>
      <div class="thumb">${m.thumb ? `<img src="${m.thumb}" alt="">` : '📄'}</div>
      <div class="name" title="${m.file.name}">${m.file.name}</div>
      <div class="meta">${m.pages} page${m.pages === 1 ? '' : 's'} · ${fmtSize(m.file.size)}</div>
    </div>`).join('');
  $$('.remove', list).forEach(b => b.onclick = () => { mergeFiles.splice(b.dataset.i, 1); renderMerge(); });
  makeSortable(list, mergeFiles, renderMerge);
  $('#merge-options').hidden = !mergeFiles.length;
  $('#merge-go').disabled = mergeFiles.length < 2;
  $('#merge-go').textContent = mergeFiles.length < 2 ? 'Add at least 2 PDFs' : `Merge ${mergeFiles.length} PDFs`;
  $('#merge-drop').classList.toggle('compact', mergeFiles.length > 0);
}
$('#merge-go').onclick = async () => {
  if (mergeFiles.length < 2) return;
  busy(true, 'Merging…');
  try {
    const out = await PDFDocument.create();
    for (let i = 0; i < mergeFiles.length; i++) {
      loaderText.textContent = `Merging file ${i + 1} of ${mergeFiles.length}…`; await nextFrame();
      const src = await loadPdfLib(await mergeFiles[i].file.arrayBuffer());
      const pages = await out.copyPages(src, src.getPageIndices());
      pages.forEach(p => out.addPage(p));
    }
    download(await out.save(), 'mypdf-merged.pdf');
    track(mergeFiles.length, out.getPageCount(), sumSize(mergeFiles.map(m => m.file)));
    toast('Your PDFs were merged!');
  } catch (e) { console.error(e); toast(e.message, true); }
  busy(false);
};

/* SPLIT */
let splitFile = null, splitCount = 0;
setupDrop('#split-drop', '#split-input', isPdf, async files => {
  splitFile = files[0];
  busy(true, 'Loading PDF…');
  try {
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await splitFile.arrayBuffer()) }).promise;
    splitCount = pdf.numPages;
    $('#split-name').textContent = `📄 ${splitFile.name} · ${splitCount} pages`;
    $('#split-options').hidden = false; $('#split-drop').classList.add('compact');
    const grid = $('#split-pages'); grid.innerHTML = '';
    const max = Math.min(splitCount, 60);
    for (let n = 1; n <= max; n++) {
      const c = await renderPdfPage(pdf, n, 180);
      const card = document.createElement('div'); card.className = 'file-card'; card.style.cursor = 'default'; card.style.setProperty('--i', Math.min(n, 20));
      card.innerHTML = `<span class="badge">${n}</span><div class="thumb"></div><div class="meta">Page ${n}</div>`;
      $('.thumb', card).appendChild(c); grid.appendChild(card);
    }
    if (splitCount > max) grid.insertAdjacentHTML('beforeend', `<div class="file-card" style="display:grid;place-items:center">+${splitCount - max} more pages</div>`);
  } catch (e) { console.error(e); toast('This PDF could not be read.', true); }
  busy(false);
});
$$('input[name=split-mode]').forEach(r => r.onchange = () => { $('#split-ranges-wrap').hidden = r.value === 'all' && r.checked; });
function parseRanges(str, max) {
  const out = [];
  for (const part of str.split(',').map(s => s.trim()).filter(Boolean)) {
    const m = part.match(/^(\d+)\s*(?:-\s*(\d+)?)?$/);
    if (!m) throw new Error(`"${part}" is not a valid page or range`);
    const a = +m[1], b = m[2] ? +m[2] : (part.includes('-') ? max : a);
    if (a < 1 || b > max || a > b) throw new Error(`Range "${part}" is outside pages 1–${max}`);
    out.push([a, b]);
  }
  if (!out.length) throw new Error('Please enter at least one page or range');
  return out;
}
$('#split-go').onclick = async () => {
  if (!splitFile) return;
  const mode = $('input[name=split-mode]:checked').value;
  let ranges;
  try {
    ranges = mode === 'all' ? Array.from({ length: splitCount }, (_, i) => [i + 1, i + 1]) : parseRanges($('#split-ranges').value, splitCount);
  } catch (e) { return toast(e.message, true); }
  busy(true, 'Splitting…');
  try {
    const src = await loadPdfLib(await splitFile.arrayBuffer());
    const base = baseName(splitFile.name);
    const idx = ([a, b]) => Array.from({ length: b - a + 1 }, (_, i) => a - 1 + i);
    if (mode === 'extract') {
      const out = await PDFDocument.create();
      (await out.copyPages(src, ranges.flatMap(idx))).forEach(p => out.addPage(p));
      download(await out.save(), `${base}-extracted.pdf`);
    } else if (ranges.length === 1) {
      const out = await PDFDocument.create();
      (await out.copyPages(src, idx(ranges[0]))).forEach(p => out.addPage(p));
      download(await out.save(), `${base}-${ranges[0][0]}-${ranges[0][1]}.pdf`);
    } else {
      const zip = new JSZip();
      for (let i = 0; i < ranges.length; i++) {
        loaderText.textContent = `Creating file ${i + 1} of ${ranges.length}…`; await nextFrame();
        const [a, b] = ranges[i];
        const out = await PDFDocument.create();
        (await out.copyPages(src, idx(ranges[i]))).forEach(p => out.addPage(p));
        zip.file(a === b ? `${base}-page-${a}.pdf` : `${base}-pages-${a}-${b}.pdf`, await out.save());
      }
      loaderText.textContent = 'Creating ZIP…';
      download(await zip.generateAsync({ type: 'blob' }), `${base}-split.zip`, 'application/zip');
    }
    track(1, splitCount, splitFile.size);
    toast('Done! Your download has started.');
  } catch (e) { console.error(e); toast(e.message, true); }
  busy(false);
};

/* SIGN */
let signFile = null, signPdf = null, signPage = 1, signature = null; // signature: {url, ratio}
const placements = []; // {page, x, y, w, h (fractions of page), url, ratio, date}

// signature tabs
let activeTab = 'draw';
$$('.tab').forEach(t => t.onclick = () => {
  activeTab = t.dataset.tab;
  $$('.tab').forEach(x => x.classList.toggle('active', x === t));
  $$('.tab-pane').forEach(p => p.hidden = p.dataset.pane !== activeTab);
});
// draw pad
const pad = $('#sig-pad'), pctx = pad.getContext('2d');
let drawing = false, padDirty = false, last = null;
function padPos(e) { const r = pad.getBoundingClientRect(); return { x: (e.clientX - r.left) * pad.width / r.width, y: (e.clientY - r.top) * pad.height / r.height }; }
pad.addEventListener('pointerdown', e => { drawing = true; padDirty = true; last = padPos(e); pad.setPointerCapture(e.pointerId); });
pad.addEventListener('pointermove', e => {
  if (!drawing) return; const p = padPos(e);
  pctx.strokeStyle = $('#sig-color').value; pctx.lineWidth = 4; pctx.lineCap = 'round'; pctx.lineJoin = 'round';
  pctx.beginPath(); pctx.moveTo(last.x, last.y); pctx.lineTo(p.x, p.y); pctx.stroke(); last = p;
});
['pointerup', 'pointercancel'].forEach(ev => pad.addEventListener(ev, () => drawing = false));
$('#sig-clear').onclick = () => { pctx.clearRect(0, 0, pad.width, pad.height); padDirty = false; };
// type preview
function updateTypePreview() { const p = $('#sig-type-preview'); p.textContent = $('#sig-text').value || 'Your signature'; p.style.fontFamily = $('#sig-font').value; }
$('#sig-text').oninput = updateTypePreview; $('#sig-font').onchange = updateTypePreview; updateTypePreview();

function trimCanvas(c) {
  const ctx = c.getContext('2d'), { width: w, height: h } = c, d = ctx.getImageData(0, 0, w, h).data;
  let x0 = w, y0 = h, x1 = -1, y1 = -1;
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) if (d[(y * w + x) * 4 + 3] > 10) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
  if (x1 < 0) return null;
  const pad = 6; x0 = Math.max(0, x0 - pad); y0 = Math.max(0, y0 - pad); x1 = Math.min(w - 1, x1 + pad); y1 = Math.min(h - 1, y1 + pad);
  const o = document.createElement('canvas'); o.width = x1 - x0 + 1; o.height = y1 - y0 + 1;
  o.getContext('2d').drawImage(c, x0, y0, o.width, o.height, 0, 0, o.width, o.height);
  return o;
}
async function buildSignature() {
  if (activeTab === 'draw') {
    if (!padDirty) throw new Error('Please draw your signature first');
    return trimCanvas(pad);
  }
  if (activeTab === 'type') {
    const text = $('#sig-text').value.trim(); if (!text) throw new Error('Please type your name');
    const font = $('#sig-font').value; await document.fonts.load(`120px ${font}`);
    const c = document.createElement('canvas'); c.width = 1400; c.height = 300;
    const ctx = c.getContext('2d'); ctx.font = `120px ${font}`; ctx.fillStyle = $('#sig-color').value; ctx.textBaseline = 'middle';
    ctx.fillText(text, 30, 150);
    return trimCanvas(c);
  }
  const f = $('#sig-upload').files[0]; if (!f) throw new Error('Please choose an image');
  const bmp = await createImageBitmap(f);
  const c = document.createElement('canvas'); c.width = bmp.width; c.height = bmp.height; c.getContext('2d').drawImage(bmp, 0, 0);
  return c;
}
$('#sig-use').onclick = async () => {
  try {
    const c = await buildSignature(); if (!c) throw new Error('Signature is empty');
    signature = { url: c.toDataURL('image/png'), ratio: c.height / c.width };
    $('#sig-current-img').src = signature.url; $('#sig-current').hidden = false;
    toast('Signature ready. Click on the page to place it.');
  } catch (e) { toast(e.message, true); }
};

// load PDF
setupDrop('#sign-drop', '#sign-input', isPdf, async files => {
  signFile = files[0]; placements.length = 0;
  busy(true, 'Loading PDF…');
  try {
    signPdf = await pdfjsLib.getDocument({ data: new Uint8Array(await signFile.arrayBuffer()) }).promise;
    signPage = 1; $('#sign-workspace').hidden = false; $('#sign-drop').classList.add('compact');
    await renderSignPage();
  } catch (e) { console.error(e); toast('This PDF could not be read.', true); }
  busy(false);
});
async function renderSignPage() {
  const page = await signPdf.getPage(signPage);
  const avail = Math.min(900, $('.sign-viewer').clientWidth - 28);
  const v1 = page.getViewport({ scale: 1 });
  const vp = page.getViewport({ scale: (avail / v1.width) * (window.devicePixelRatio || 1) });
  const c = $('#sign-canvas'); c.width = vp.width; c.height = vp.height;
  c.style.width = avail + 'px';
  await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
  $('#sign-pageinfo').textContent = `Page ${signPage} / ${signPdf.numPages}`;
  $('#sign-prev').disabled = signPage <= 1; $('#sign-next').disabled = signPage >= signPdf.numPages;
  renderPlacements();
}
$('#sign-prev').onclick = () => { if (signPage > 1) { signPage--; renderSignPage(); } };
$('#sign-next').onclick = () => { if (signPage < signPdf.numPages) { signPage++; renderSignPage(); } };
let resizeT; window.addEventListener('resize', () => { if (signPdf && !$('#sign').hidden) { clearTimeout(resizeT); resizeT = setTimeout(renderSignPage, 200); } });

const fmtDate = () => new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
$('#sign-overlay').addEventListener('click', e => {
  if (e.target !== e.currentTarget) return;
  if (!signature) return toast('Create a signature first (step 1)', true);
  const r = e.currentTarget.getBoundingClientRect();
  const w = 0.25, h = w * signature.ratio * (r.width / r.height);
  placements.push({ page: signPage, x: Math.max(0, Math.min(1 - w, (e.clientX - r.left) / r.width - w / 2)), y: Math.max(0, Math.min(1 - h, (e.clientY - r.top) / r.height - h / 2)), w, h, url: signature.url, ratio: signature.ratio, date: $('#sig-date').checked ? fmtDate() : null });
  renderPlacements();
});
function renderPlacements() {
  const ov = $('#sign-overlay'); ov.innerHTML = '';
  placements.forEach((p, i) => {
    if (p.page !== signPage) return;
    const el = document.createElement('div'); el.className = 'placed';
    const place = () => Object.assign(el.style, { left: p.x * 100 + '%', top: p.y * 100 + '%', width: p.w * 100 + '%', height: p.h * 100 + '%' });
    place();
    el.innerHTML = `<img src="${p.url}" alt=""><button class="del" title="Remove">✕</button><span class="rs"></span>${p.date ? `<span class="date">${p.date}</span>` : ''}`;
    $('.del', el).onclick = ev => { ev.stopPropagation(); placements.splice(i, 1); renderPlacements(); };
    el.addEventListener('pointerdown', ev => {
      if (ev.target.classList.contains('del')) return;
      ev.preventDefault(); el.setPointerCapture(ev.pointerId);
      const r = ov.getBoundingClientRect(), resizing = ev.target.classList.contains('rs');
      const sx = ev.clientX, sy = ev.clientY, o = { ...p };
      const move = mv => {
        const dx = (mv.clientX - sx) / r.width, dy = (mv.clientY - sy) / r.height;
        if (resizing) { p.w = Math.max(0.04, Math.min(1 - p.x, o.w + dx)); p.h = p.w * p.ratio * (r.width / r.height); }
        else { p.x = Math.max(0, Math.min(1 - p.w, o.x + dx)); p.y = Math.max(0, Math.min(1 - p.h, o.y + dy)); }
        place();
      };
      const up = () => { el.removeEventListener('pointermove', move); el.removeEventListener('pointerup', up); };
      el.addEventListener('pointermove', move); el.addEventListener('pointerup', up);
    });
    ov.appendChild(el);
  });
}
$('#sign-reset').onclick = () => { placements.length = 0; renderPlacements(); };

$('#sign-go').onclick = async () => {
  if (!signFile) return;
  if (!placements.length) return toast('Click on the page to place your signature first', true);
  busy(true, 'Signing PDF…');
  try {
    const doc = await loadPdfLib(await signFile.arrayBuffer());
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const cache = {};
    for (const p of placements) {
      const page = doc.getPage(p.page - 1);
      cache[p.url] = cache[p.url] || await doc.embedPng(p.url);
      const cb = page.getCropBox(), rot = (page.getRotation().angle % 360 + 360) % 360;
      // map displayed (rotated) fractions back to unrotated page coordinates
      const { x, y, w, h } = p, W = cb.width, H = cb.height;
      let opts;
      if (rot === 0) opts = { x: cb.x + x * W, y: cb.y + H - (y + h) * H, width: w * W, height: h * H };
      else if (rot === 180) opts = { x: cb.x + (1 - x) * W, y: cb.y + (y + h) * H, width: w * W, height: h * H, rotate: PDFLib.degrees(180) };
      else if (rot === 90) opts = { x: cb.x + (y + h) * W, y: cb.y + x * H, width: w * H, height: h * W, rotate: PDFLib.degrees(90) };
      else if (rot === 270) opts = { x: cb.x + (1 - y - h) * W, y: cb.y + (1 - x) * H, width: w * H, height: h * W, rotate: PDFLib.degrees(270) };
      page.drawImage(cache[p.url], opts);
      if (p.date && rot === 0) {
        const size = Math.max(7, Math.min(12, h * H * 0.25));
        page.drawText(p.date, { x: opts.x, y: opts.y - size - 2, size, font, color: rgb(0.13, 0.13, 0.13) });
      }
    }
    download(await doc.save(), baseName(signFile.name) + '-signed.pdf');
    track(1, doc.getPageCount(), signFile.size);
    toast('Your signed PDF is ready!');
  } catch (e) { console.error(e); toast(e.message || 'Could not sign this PDF', true); }
  busy(false);
};
/* shared helpers for tools.js */
window.MYPDF = { $, $$, busy, toast, download, baseName, fmtSize, isPdf, isImg, nextFrame, loadPdfLib, renderPdfPage, setupDrop, makeSortable, track, sumSize, loaderText, parseRanges };
})();
