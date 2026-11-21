/* MyPDF - more tools: Office conversions, PDF to Word/text, organize, compress, page numbers & watermark.
   Everything runs in the browser. Big libraries load only when a tool needs them. */
(() => {
'use strict';
const M = window.MYPDF;
const { $, $$, busy, toast, download, baseName, fmtSize, isPdf, nextFrame, loadPdfLib, renderPdfPage, setupDrop, makeSortable, track, loaderText } = M;
const { PDFDocument, rgb, StandardFonts, degrees } = PDFLib;
const say = t => { loaderText.textContent = t; };

/* ---------- lazy script loader ---------- */
const loaded = {};
function loadScript(src) {
  if (!loaded[src]) loaded[src] = new Promise((res, rej) => {
    const s = document.createElement('script'); s.src = src; s.onload = res;
    s.onerror = () => { delete loaded[src]; rej(new Error('Could not load a component. Check your connection and try again.')); };
    document.head.appendChild(s);
  });
  return loaded[src];
}
// docx-preview and the docx library both use the global name "docx", so each is kept in its own variable
let DOCX_PREVIEW = null, DOCX_LIB = null;
async function loadKeep(src) { const prev = window.docx; await loadScript(src); const lib = window.docx; window.docx = prev; return lib; }
const need = {
  html2canvas: () => loadScript('lib/html2canvas.min.js'),
  docxPreview: async () => DOCX_PREVIEW || (DOCX_PREVIEW = await loadKeep('lib/docx-preview.min.js')),
  xlsx: () => loadScript('lib/xlsx.full.min.js'),
  docx: async () => DOCX_LIB || (DOCX_LIB = await loadKeep('lib/docx.min.js')),
};
const ext = n => (n.match(/\.([^.]+)$/) || [, ''])[1].toLowerCase();
const byExt = list => f => list.includes(ext(f.name));
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const openPdfJs = async file => pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
const canvasToJpeg = (c, q = 0.9) => new Promise(r => c.toBlob(b => b.arrayBuffer().then(r), 'image/jpeg', q));

/* ---------- paper sizes (points) ---------- */
const PAPER = { a4: [595.28, 841.89], letter: [612, 792] };

/* ---------- render a DOM block to PDF pages ----------
   The element is captured with html2canvas, then cut into pages.
   Cuts are placed between paragraphs / table rows where possible, never through a line of text. */
function breakCandidates(root, scale) {
  const top = root.getBoundingClientRect().top, out = [];
  root.querySelectorAll('p, li, tr, h1, h2, h3, h4, h5, h6, img, table, br, div.slide-break, section > *').forEach(el => {
    const r = el.getBoundingClientRect();
    if (r.height > 0) out.push(Math.round((r.bottom - top) * scale));
  });
  return [...new Set(out)].sort((a, b) => a - b);
}
function whiteRowNear(canvas, from, to) {
  // look upward from `to` for a fully blank row (fallback when no DOM break fits)
  const ctx = canvas.getContext('2d'), w = canvas.width;
  for (let y = to; y > from; y -= 2) {
    const d = ctx.getImageData(0, y, w, 1).data; let blank = true;
    for (let i = 0; i < d.length; i += 16) { if (d[i] < 245 || d[i + 1] < 245 || d[i + 2] < 245) { blank = false; break; } }
    if (blank) return y;
  }
  return -1;
}
function planCuts(canvas, cands, start, end, capacity) {
  const cuts = [];
  let y = start;
  while (end - y > capacity) {
    const limit = y + capacity;
    let best = -1;
    for (const c of cands) { if (c > y + capacity * 0.35 && c <= limit) best = c; }
    if (best < 0) best = whiteRowNear(canvas, y + capacity * 0.5, limit);
    if (best < 0) best = limit;
    cuts.push([y, best]); y = best;
  }
  cuts.push([y, end]);
  return cuts;
}
// Adds pages to `doc` from a captured canvas. Page size in points; margins in points.
async function canvasToPages(doc, canvas, cands, { pageW, pageH, margin, startPx = 0, endPx = canvas.height, fitWidth = true }) {
  const contentWpt = pageW - margin * 2, contentHpt = pageH - margin * 2;
  const pxPerPt = fitWidth ? canvas.width / contentWpt : 2;
  const capacity = Math.floor(contentHpt * pxPerPt);
  const cuts = planCuts(canvas, cands, startPx, endPx, capacity);
  for (const [a, b] of cuts) {
    const h = Math.max(1, b - a);
    const slice = document.createElement('canvas'); slice.width = canvas.width; slice.height = h;
    const sx = slice.getContext('2d'); sx.fillStyle = '#fff'; sx.fillRect(0, 0, slice.width, h);
    sx.drawImage(canvas, 0, a, canvas.width, h, 0, 0, canvas.width, h);
    const img = await doc.embedJpg(await canvasToJpeg(slice, 0.92));
    const page = doc.addPage([pageW, pageH]);
    page.drawImage(img, { x: margin, y: pageH - margin - h / pxPerPt, width: canvas.width / pxPerPt, height: h / pxPerPt });
    await nextFrame();
  }
  return cuts.length;
}
// Off-screen stage the user never sees (but the browser lays out normally)
function stage(widthPx) {
  const el = document.createElement('div');
  Object.assign(el.style, { position: 'fixed', left: '-10000px', top: '0', width: widthPx + 'px', background: '#fff', color: '#000', zIndex: '-1' });
  document.body.appendChild(el); return el;
}
async function capture(el, scale = 2) {
  await need.html2canvas();
  if (document.fonts && document.fonts.ready) await document.fonts.ready;
  return html2canvas(el, { scale, backgroundColor: '#ffffff', logging: false, useCORS: true, windowWidth: el.scrollWidth });
}
// Print through a hidden iframe: the browser does real pagination and keeps text selectable
function printHtml(html, css) {
  const f = document.createElement('iframe');
  Object.assign(f.style, { position: 'fixed', right: '0', bottom: '0', width: '0', height: '0', border: '0' });
  document.body.appendChild(f);
  const d = f.contentDocument;
  d.open(); d.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>MyPDF</title><style>${css}</style></head><body>${html}</body></html>`); d.close();
  const go = () => { f.contentWindow.focus(); f.contentWindow.print(); setTimeout(() => f.remove(), 60000); };
  const imgs = [...d.images];
  Promise.all(imgs.map(i => i.complete ? 1 : new Promise(r => { i.onload = i.onerror = r; }))).then(() => setTimeout(go, 150));
}

/* =====================================================================
   WORD → PDF  (docx-preview renders the document with its styles)
   ===================================================================== */
let w2pFile = null, w2pStyle = null;
setupDrop('#w2p-drop', '#w2p-input', byExt(['docx']), async files => {
  const f = files[0];
  if (ext(f.name) !== 'docx') return toast(ext(f.name) === 'doc' ? 'Old .doc files aren\'t supported. Save it as .docx first.' : 'Please choose a .docx file', true);
  w2pFile = f; $('#w2p-name').textContent = '📄 ' + f.name;
  busy(true, 'Reading document…');
  try {
    const DP = await need.docxPreview();
    const box = $('#w2p-preview'); box.innerHTML = '';
    w2pStyle = document.createElement('div');
    await DP.renderAsync(await f.arrayBuffer(), box, w2pStyle, { inWrapper: false, ignoreWidth: false, ignoreHeight: false, breakPages: true, renderHeaders: true, renderFooters: true, renderFootnotes: true, experimental: true, useBase64URL: true });
    box.prepend(w2pStyle);
    $('#w2p-options').hidden = false; $('#w2p-preview-wrap').hidden = false; $('#w2p-drop').classList.add('compact');
  } catch (e) { console.error(e); toast('This Word file could not be read. It may be damaged, password-protected or not a real .docx.', true); }
  busy(false);
});
function pxFromCss(v) { const n = parseFloat(v) || 0; return /pt$/.test(v) ? n * 96 / 72 : n; }
$('#w2p-go').onclick = async () => {
  if (!w2pFile) return;
  busy(true, 'Creating PDF…');
  const st = stage(1200);
  try {
    st.innerHTML = $('#w2p-preview').innerHTML;
    const sections = $$('section.docx', st);
    if (!sections.length) throw new Error('Nothing to convert in this document');
    const doc = await PDFDocument.create();
    for (let i = 0; i < sections.length; i++) {
      say(`Rendering part ${i + 1} of ${sections.length}…`); await nextFrame();
      const sec = sections[i], cs = getComputedStyle(sec);
      const wPx = sec.offsetWidth, hPx = pxFromCss(sec.style.minHeight) || wPx * 1.4142;
      const padT = parseFloat(cs.paddingTop) || 0, padB = parseFloat(cs.paddingBottom) || 0;
      sec.style.minHeight = '0'; sec.style.margin = '0'; sec.style.boxShadow = 'none';
      // capture this section alone
      const holder = stage(wPx); holder.appendChild(sec.cloneNode(true));
      const s = 2, canvas = await capture(holder.firstChild, s);
      const cands = breakCandidates(holder.firstChild, s);
      holder.remove();
      const pageW = wPx * 0.75, pageH = hPx * 0.75;
      const top = Math.round(padT * s), bottom = canvas.height - Math.round(padB * s);
      // content area of each page = page height minus top/bottom margins
      const capPx = Math.round((hPx - padT - padB) * s);
      const cuts = planCuts(canvas, cands, top, Math.max(top + 1, bottom), capPx);
      for (const [a, b] of cuts) {
        const pc = document.createElement('canvas'); pc.width = canvas.width; pc.height = Math.round(hPx * s);
        const x = pc.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, pc.width, pc.height);
        x.drawImage(canvas, 0, a, canvas.width, b - a, 0, top, canvas.width, b - a);
        const img = await doc.embedJpg(await canvasToJpeg(pc, 0.92));
        doc.addPage([pageW, pageH]).drawImage(img, { x: 0, y: 0, width: pageW, height: pageH });
        await nextFrame();
      }
    }
    download(await doc.save(), baseName(w2pFile.name) + '.pdf');
    track(1, doc.getPageCount(), w2pFile.size);
    toast('Your PDF is ready!');
  } catch (e) { console.error(e); toast(e.message || 'Could not convert this document', true); }
  st.remove(); busy(false);
};
$('#w2p-print').onclick = () => {
  if (!w2pFile) return;
  const css = (w2pStyle ? [...w2pStyle.querySelectorAll('style')].map(s => s.textContent).join('\n') : '') +
    '\n@page{margin:0}html,body{margin:0;background:#fff}section.docx{box-shadow:none!important;margin:0 auto!important;page-break-after:always;break-after:page}';
  const html = $('#w2p-preview').cloneNode(true); html.querySelectorAll('style').forEach(s => s.remove());
  printHtml(html.innerHTML, css);
};

/* =====================================================================
   EXCEL → PDF  (SheetJS reads the workbook, we lay out clean tables)
   ===================================================================== */
let x2pFile = null, x2pBook = null;
setupDrop('#x2p-drop', '#x2p-input', byExt(['xlsx', 'xls', 'xlsm', 'csv', 'ods']), async files => {
  const f = files[0];
  if (!['xlsx', 'xls', 'xlsm', 'csv', 'ods'].includes(ext(f.name))) return toast('Please choose an .xlsx, .xls, .csv or .ods file', true);
  busy(true, 'Reading spreadsheet…');
  try {
    await need.xlsx();
    x2pBook = XLSX.read(await f.arrayBuffer(), { type: 'array', cellDates: true });
    x2pFile = f; $('#x2p-name').textContent = '📊 ' + f.name;
    const used = x2pBook.SheetNames.filter(n => x2pBook.Sheets[n]['!ref']);
    if (!used.length) throw new Error('empty');
    $('#x2p-sheet').innerHTML = (used.length > 1 ? `<option value="__all">All sheets (${used.length})</option>` : '') + used.map(n => `<option value="${esc(n)}">${esc(n)}</option>`).join('');
    $('#x2p-options').hidden = false; $('#x2p-preview-wrap').hidden = false; $('#x2p-drop').classList.add('compact');
    renderSheets($('#x2p-preview'));
  } catch (e) { console.error(e); toast(e.message === 'empty' ? 'This spreadsheet is empty' : 'This spreadsheet could not be read. It may be damaged or password-protected.', true); }
  busy(false);
});
function sheetNames() {
  const v = $('#x2p-sheet').value;
  return v === '__all' ? x2pBook.SheetNames.filter(n => x2pBook.Sheets[n]['!ref']) : [v];
}
function sheetHtml(name) {
  const ws = x2pBook.Sheets[name];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '', blankrows: false });
  let maxC = 0; rows.forEach(r => { for (let c = r.length - 1; c >= 0; c--) if (String(r[c]).trim() !== '') { maxC = Math.max(maxC, c + 1); break; } });
  const grid = $('#x2p-grid').checked;
  const body = rows.map((r, i) => '<tr>' + Array.from({ length: maxC }, (_, c) => {
    const v = r[c] == null ? '' : String(r[c]);
    const num = /^[-+]?[\d,.\s]*\d[\d,.\s]*%?$/.test(v.trim()) || /^[-+]?[₱$€£¥]\s?[\d,.]+$/.test(v.trim());
    return `<td${num ? ' class="n"' : ''}>${esc(v)}</td>`;
  }).join('') + '</tr>').join('');
  return `<div class="xs${grid ? ' grid' : ''}"><div class="xs-name">${esc(name)}</div><table>${body}</table></div>`;
}
const SHEET_CSS = `.xs{font:12px/1.35 Arial,Helvetica,sans-serif;color:#111;padding:0 0 18px}
.xs-name{font-weight:700;font-size:13px;margin:0 0 8px;color:#0f5132}
.xs table{border-collapse:collapse}
.xs td{padding:4px 7px;vertical-align:top;white-space:pre-wrap;word-break:normal;max-width:360px}
.xs.grid td{border:1px solid #c9ced6}
.xs:not(.grid) tr:nth-child(odd) td{background:#f6f8fa}
.xs tr:first-child td{font-weight:700;background:#eef2f6}
.xs td.n{text-align:right;white-space:nowrap}`;
function renderSheets(box) {
  box.innerHTML = `<style>${SHEET_CSS}</style>` + sheetNames().map(sheetHtml).join('');
}
['#x2p-sheet', '#x2p-grid'].forEach(id => $(id).addEventListener('change', () => x2pBook && renderSheets($('#x2p-preview'))));
$('#x2p-go').onclick = async () => {
  if (!x2pBook) return;
  busy(true, 'Creating PDF…');
  try {
    const doc = await PDFDocument.create(), margin = 32;
    const names = sheetNames();
    for (let i = 0; i < names.length; i++) {
      say(`Rendering sheet ${i + 1} of ${names.length}…`); await nextFrame();
      const st = stage(4000);
      st.innerHTML = `<style>${SHEET_CSS}</style>` + sheetHtml(names[i]);
      const block = st.querySelector('.xs');
      block.style.display = 'inline-block';
      const natW = Math.ceil(block.getBoundingClientRect().width);
      let [pw, ph] = PAPER[$('#x2p-paper').value];
      const o = $('#x2p-orient').value;
      const land = o === 'landscape' || (o === 'auto' && natW > (pw - margin * 2) * 96 / 72 * 1.05);
      if (land) [pw, ph] = [ph, pw];
      const contentPx = (pw - margin * 2) * 96 / 72;
      const w = Math.max(natW, contentPx);
      st.style.width = w + 'px'; block.style.width = w + 'px';
      const scale = Math.min(3, Math.max(1.4, 2 * w / contentPx));
      const canvas = await capture(block, scale);
      const cands = breakCandidates(block, scale);
      st.remove();
      await canvasToPages(doc, canvas, cands, { pageW: pw, pageH: ph, margin });
    }
    download(await doc.save(), baseName(x2pFile.name) + '.pdf');
    track(1, doc.getPageCount(), x2pFile.size);
    toast('Your PDF is ready!');
  } catch (e) { console.error(e); toast(e.message || 'Could not convert this spreadsheet', true); }
  busy(false);
};
$('#x2p-print').onclick = () => {
  if (!x2pBook) return;
  const o = $('#x2p-orient').value, paper = $('#x2p-paper').value === 'letter' ? 'letter' : 'A4';
  const css = SHEET_CSS + `\n@page{size:${paper}${o === 'landscape' ? ' landscape' : o === 'portrait' ? ' portrait' : ''};margin:12mm}body{margin:0}.xs{page-break-after:always;break-after:page}.xs table{width:100%}tr{break-inside:avoid}`;
  printHtml(sheetNames().map(sheetHtml).join(''), css);
};

/* =====================================================================
   POWERPOINT → PDF  (our own slide renderer: text, pictures, backgrounds, basic shapes)
   ===================================================================== */
const EMU = 12700; // EMU per point
const NS = { a: 'http://schemas.openxmlformats.org/drawingml/2006/main', p: 'http://schemas.openxmlformats.org/presentationml/2006/main', r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' };
const parseXml = t => new DOMParser().parseFromString(t, 'application/xml');
const kids = (el, ns, name) => el ? [...el.children].filter(c => c.localName === name && (!ns || c.namespaceURI === NS[ns])) : [];
const kid = (el, ns, name) => kids(el, ns, name)[0] || null;
const desc = (el, name) => el ? el.getElementsByTagNameNS('*', name) : [];
const first = (el, name) => (el && el.getElementsByTagNameNS('*', name)[0]) || null;
function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/'); parts.pop();
  target.split('/').forEach(p => { if (p === '..') parts.pop(); else if (p !== '.') parts.push(p); });
  return parts.join('/');
}
async function readRels(zip, partPath) {
  const dir = partPath.split('/'); const file = dir.pop();
  const relPath = dir.join('/') + '/_rels/' + file + '.rels';
  const f = zip.file(relPath); const out = {};
  if (!f) return out;
  const x = parseXml(await f.async('text'));
  [...desc(x, 'Relationship')].forEach(r => { out[r.getAttribute('Id')] = { target: resolvePath(partPath, r.getAttribute('Target')), type: r.getAttribute('Type') || '', external: r.getAttribute('TargetMode') === 'External' }; });
  return out;
}
function hexToRgb(h) { const n = parseInt(h, 16); return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; }
function applyMods(rgbA, el) {
  let [r, g, b] = rgbA, a = 1;
  [...el.children].forEach(m => {
    const v = +m.getAttribute('val') / 100000;
    if (m.localName === 'lumMod') { r *= v; g *= v; b *= v; }
    else if (m.localName === 'lumOff') { r += 255 * v; g += 255 * v; b += 255 * v; }
    else if (m.localName === 'tint') { r = r + (255 - r) * (1 - v); g = g + (255 - g) * (1 - v); b = b + (255 - b) * (1 - v); }
    else if (m.localName === 'shade') { r *= v; g *= v; b *= v; }
    else if (m.localName === 'alpha') a = v;
  });
  const c = x => Math.max(0, Math.min(255, Math.round(x)));
  return `rgba(${c(r)},${c(g)},${c(b)},${a})`;
}
function colorOf(el, theme) {
  if (!el) return null;
  const s = kid(el, 'a', 'srgbClr'), sc = kid(el, 'a', 'schemeClr'), sys = kid(el, 'a', 'sysClr'), pr = kid(el, 'a', 'prstClr');
  if (s) return applyMods(hexToRgb(s.getAttribute('val')), s);
  if (sys) return applyMods(hexToRgb(sys.getAttribute('lastClr') || '000000'), sys);
  if (sc) {
    let key = sc.getAttribute('val');
    key = { tx1: 'dk1', tx2: 'dk2', bg1: 'lt1', bg2: 'lt2' }[key] || key;
    return applyMods(hexToRgb(theme[key] || '000000'), sc);
  }
  if (pr) { const m = { black: '000000', white: 'FFFFFF', red: 'FF0000', blue: '0000FF', green: '008000', yellow: 'FFFF00', gray: '808080' }; return applyMods(hexToRgb(m[pr.getAttribute('val')] || '000000'), pr); }
  return null;
}
async function loadImage(zip, path, cache) {
  if (cache[path]) return cache[path];
  const f = zip.file(path); if (!f) return null;
  const blob = await f.async('blob');
  const mime = /\.png$/i.test(path) ? 'image/png' : /\.gif$/i.test(path) ? 'image/gif' : /\.svg$/i.test(path) ? 'image/svg+xml' : 'image/jpeg';
  try { cache[path] = await createImageBitmap(new Blob([blob], { type: mime })); }
  catch { cache[path] = await new Promise(r => { const i = new Image(); i.onload = () => r(i); i.onerror = () => r(null); i.src = URL.createObjectURL(new Blob([blob], { type: mime })); }); }
  return cache[path];
}
function xfrmOf(sp) {
  const x = first(kid(sp, 'p', 'spPr') || kid(sp, 'p', 'grpSpPr') || sp, 'xfrm');
  if (!x) return null;
  const off = kid(x, 'a', 'off'), e = kid(x, 'a', 'ext');
  if (!off || !e) return null;
  const o = { x: +off.getAttribute('x'), y: +off.getAttribute('y'), w: +e.getAttribute('cx'), h: +e.getAttribute('cy'), rot: +(x.getAttribute('rot') || 0) / 60000, flipH: x.getAttribute('flipH') === '1', flipV: x.getAttribute('flipV') === '1' };
  const co = kid(x, 'a', 'chOff'), ce = kid(x, 'a', 'chExt');
  if (co && ce) Object.assign(o, { chX: +co.getAttribute('x'), chY: +co.getAttribute('y'), chW: +ce.getAttribute('cx') || 1, chH: +ce.getAttribute('cy') || 1 });
  return o;
}
function phKey(sp) {
  const ph = first(sp, 'ph'); if (!ph) return null;
  return { type: ph.getAttribute('type') || 'body', idx: ph.getAttribute('idx') };
}
function findPh(tree, key) {
  if (!tree || !key) return null;
  const sps = [...desc(tree, 'sp')];
  return sps.find(s => { const k = phKey(s); return k && k.idx != null && k.idx === key.idx && key.idx != null; }) ||
    sps.find(s => { const k = phKey(s); return k && (k.type === key.type || (key.type === 'ctrTitle' && k.type === 'title') || (key.type === 'subTitle' && k.type === 'body')); }) || null;
}
async function bgOf(cSld, theme, zip, rels, cache) {
  const bg = kid(cSld, 'p', 'bg'); if (!bg) return null;
  const pr = kid(bg, 'p', 'bgPr');
  if (pr) {
    const sf = kid(pr, 'a', 'solidFill'); if (sf) return { color: colorOf(sf, theme) };
    const bf = kid(pr, 'a', 'blipFill');
    if (bf) { const blip = first(bf, 'blip'); const id = blip && blip.getAttributeNS(NS.r, 'embed'); if (id && rels[id]) return { image: await loadImage(zip, rels[id].target, cache) }; }
    const gf = kid(pr, 'a', 'gradFill');
    if (gf) { const stops = [...desc(gf, 'gs')].map(g => ({ pos: +g.getAttribute('pos') / 100000, color: colorOf(g, theme) })); if (stops.length) return { grad: stops }; }
  }
  const ref = kid(bg, 'p', 'bgRef'); if (ref) { const c = colorOf(ref, theme); if (c) return { color: c }; }
  return null;
}
function wrapLines(ctx, text, maxW) {
  const out = [];
  text.split('\n').forEach(par => {
    const words = par.split(/(\s+)/); let line = '';
    words.forEach(w => {
      const t = line + w;
      if (ctx.measureText(t).width > maxW && line.trim()) { out.push(line.trimEnd()); line = w.trimStart(); }
      else line = t;
    });
    out.push(line);
  });
  return out;
}
async function renderSlide(ctx, S, slidePath, zip, ctxInfo) {
  const { theme, cache, scale, W, H, fonts } = ctxInfo;
  const sx = new XMLSerializer();
  const slideXml = parseXml(await zip.file(slidePath).async('text'));
  const rels = await readRels(zip, slidePath);
  const layoutRel = Object.values(rels).find(r => /slideLayout/.test(r.type));
  let layoutXml = null, layoutRels = {}, masterXml = null, masterRels = {};
  if (layoutRel && zip.file(layoutRel.target)) {
    layoutXml = parseXml(await zip.file(layoutRel.target).async('text')); layoutRels = await readRels(zip, layoutRel.target);
    const mRel = Object.values(layoutRels).find(r => /slideMaster/.test(r.type));
    if (mRel && zip.file(mRel.target)) { masterXml = parseXml(await zip.file(mRel.target).async('text')); masterRels = await readRels(zip, mRel.target); }
  }
  const cSld = first(slideXml, 'cSld'), lcSld = first(layoutXml, 'cSld'), mcSld = first(masterXml, 'cSld');
  // background
  let bg = await bgOf(cSld, theme, zip, rels, cache) || await bgOf(lcSld, theme, zip, layoutRels, cache) || await bgOf(mcSld, theme, zip, masterRels, cache);
  ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, W, H);
  if (bg && bg.color) { ctx.fillStyle = bg.color; ctx.fillRect(0, 0, W, H); }
  else if (bg && bg.grad) { const g = ctx.createLinearGradient(0, 0, 0, H); bg.grad.forEach(s => g.addColorStop(Math.min(1, Math.max(0, s.pos)), s.color)); ctx.fillStyle = g; ctx.fillRect(0, 0, W, H); }
  else if (bg && bg.image) ctx.drawImage(bg.image, 0, 0, W, H);
  // master + layout decorations (non-placeholder shapes), then the slide itself
  const showMaster = !(layoutXml && first(layoutXml, 'sldLayout') && first(layoutXml, 'sldLayout').getAttribute('showMasterSp') === '0');
  const env = (tree, relsMap, level) => ({ tree, rels: relsMap, level });
  const layers = [];
  if (mcSld && showMaster) layers.push(env(first(mcSld, 'spTree'), masterRels, 'master'));
  if (lcSld) layers.push(env(first(lcSld, 'spTree'), layoutRels, 'layout'));
  if (cSld) layers.push(env(first(cSld, 'spTree'), rels, 'slide'));
  const lTree = lcSld && first(lcSld, 'spTree'), mTree = mcSld && first(mcSld, 'spTree');
  const masterTx = masterXml && first(masterXml, 'txStyles');
  const T = { x: 0, y: 0, sx: scale, sy: scale };
  for (const L of layers) await drawTree(ctx, L.tree, L, T);

  async function drawTree(ctx, tree, L, tf) {
    if (!tree) return;
    for (const el of [...tree.children]) {
      const n = el.localName;
      if (n === 'sp' || n === 'cxnSp') await drawShape(ctx, el, L, tf);
      else if (n === 'pic') await drawPic(ctx, el, L, tf);
      else if (n === 'grpSp') {
        const g = xfrmOf(el); if (!g) { await drawTree(ctx, el, L, tf); continue; }
        const kx = g.w / (g.chW || g.w || 1), ky = g.h / (g.chH || g.h || 1);
        const t2 = { x: tf.x + (g.x - (g.chX || 0) * kx) * tf.sx, y: tf.y + (g.y - (g.chY || 0) * ky) * tf.sy, sx: tf.sx * kx, sy: tf.sy * ky };
        // child coordinates are relative to chOff; convert: X = g.x + (cx - chX) * kx
        await drawTree(ctx, el, L, { x: tf.x + g.x * tf.sx - (g.chX || 0) * kx * tf.sx, y: tf.y + g.y * tf.sy - (g.chY || 0) * ky * tf.sy, sx: t2.sx, sy: t2.sy });
      } else if (n === 'graphicFrame') await drawTable(ctx, el, L, tf);
      else if (n === 'AlternateContent') { const fb = kid(el, null, 'Fallback') || kid(el, null, 'Choice'); if (fb) await drawTree(ctx, fb, L, tf); }
    }
  }
  function box(el, L, tf) {
    let x = xfrmOf(el);
    const key = phKey(el);
    if (L.level !== 'slide' && key) return null; // placeholders on layout/master are only templates
    if (!x && key) {
      const lp = findPh(lTree, key), mp = findPh(mTree, key);
      x = (lp && xfrmOf(lp)) || (mp && xfrmOf(mp));
    }
    if (!x) return null;
    return { x: tf.x + x.x * tf.sx, y: tf.y + x.y * tf.sy, w: x.w * tf.sx, h: x.h * tf.sy, rot: x.rot, flipH: x.flipH, flipV: x.flipV };
  }
  function withRot(ctx, b, fn) {
    ctx.save();
    if (b.rot || b.flipH || b.flipV) { ctx.translate(b.x + b.w / 2, b.y + b.h / 2); ctx.rotate((b.rot || 0) * Math.PI / 180); ctx.scale(b.flipH ? -1 : 1, b.flipV ? -1 : 1); ctx.translate(-(b.x + b.w / 2), -(b.y + b.h / 2)); }
    fn(); ctx.restore();
  }
  function shapePath(ctx, geom, b) {
    ctx.beginPath();
    if (geom === 'ellipse') ctx.ellipse(b.x + b.w / 2, b.y + b.h / 2, Math.abs(b.w / 2), Math.abs(b.h / 2), 0, 0, Math.PI * 2);
    else if (geom === 'roundRect' && ctx.roundRect) ctx.roundRect(b.x, b.y, b.w, b.h, Math.min(b.w, b.h) * 0.16);
    else if (geom === 'line' || geom === 'straightConnector1') { ctx.moveTo(b.x, b.y); ctx.lineTo(b.x + b.w, b.y + b.h); }
    else if (geom === 'triangle') { ctx.moveTo(b.x + b.w / 2, b.y); ctx.lineTo(b.x + b.w, b.y + b.h); ctx.lineTo(b.x, b.y + b.h); ctx.closePath(); }
    else if (geom === 'rightArrow') { const t = b.h * 0.25, hx = b.x + b.w - b.h * 0.5; ctx.moveTo(b.x, b.y + t); ctx.lineTo(hx, b.y + t); ctx.lineTo(hx, b.y); ctx.lineTo(b.x + b.w, b.y + b.h / 2); ctx.lineTo(hx, b.y + b.h); ctx.lineTo(hx, b.y + b.h - t); ctx.lineTo(b.x, b.y + b.h - t); ctx.closePath(); }
    else ctx.rect(b.x, b.y, b.w, b.h);
  }
  function styleColor(el, path) { const st = first(el, 'style'); const r = st && first(st, path); return r ? colorOf(r, theme) : null; }
  async function drawShape(ctx, sp, L, tf) {
    const b = box(sp, L, tf); if (!b) return;
    const spPr = kid(sp, 'p', 'spPr');
    const geom = (first(spPr, 'prstGeom') || { getAttribute: () => 'rect' }).getAttribute('prst') || 'rect';
    let fill = null, line = null, lw = 1;
    if (spPr) {
      const sf = kid(spPr, 'a', 'solidFill'), nf = kid(spPr, 'a', 'noFill'), gf = kid(spPr, 'a', 'gradFill'), bf = kid(spPr, 'a', 'blipFill');
      if (sf) fill = colorOf(sf, theme);
      else if (gf) { const gs = first(gf, 'gs'); fill = gs ? colorOf(gs, theme) : null; }
      else if (!nf && !bf) fill = styleColor(sp, 'fillRef');
      if (bf) { const blip = first(bf, 'blip'); const id = blip && blip.getAttributeNS(NS.r, 'embed'); if (id && L.rels[id]) { const im = await loadImage(zip, L.rels[id].target, cache); if (im) withRot(ctx, b, () => ctx.drawImage(im, b.x, b.y, b.w, b.h)); } }
      const ln = kid(spPr, 'a', 'ln');
      if (ln && !kid(ln, 'a', 'noFill')) { const lf = kid(ln, 'a', 'solidFill'); line = lf ? colorOf(lf, theme) : styleColor(sp, 'lnRef'); lw = Math.max(1, (+ln.getAttribute('w') || 12700) * tf.sx); }
      else if (!ln) line = (geom === 'line' || sp.localName === 'cxnSp') ? styleColor(sp, 'lnRef') : null;
    }
    if (fill || line) withRot(ctx, b, () => {
      shapePath(ctx, geom, b);
      if (fill && geom !== 'line') { ctx.fillStyle = fill; ctx.fill(); }
      if (line) { ctx.strokeStyle = line; ctx.lineWidth = lw; ctx.stroke(); }
    });
    const txBody = kid(sp, 'p', 'txBody');
    if (txBody) withRot(ctx, b, () => drawText(ctx, txBody, b, sp, L));
  }
  // paragraph defaults for a list level, inherited: slide shape → layout placeholder → master placeholder → master text styles
  function levelProps(sp, lvl) {
    const key = phKey(sp); const type = key ? key.type : null;
    const isTitle = type === 'title' || type === 'ctrTitle';
    const styleName = !key ? 'otherStyle' : isTitle ? 'titleStyle' : (type === 'body' || type === 'obj' || type === 'subTitle' || /^(ph)?$/.test(type)) ? 'bodyStyle' : 'otherStyle';
    const srcs = [first(kid(sp, 'p', 'txBody'), 'lstStyle')];
    if (key) { const lp = findPh(lTree, key), mp = findPh(mTree, key); srcs.push(lp && first(lp, 'lstStyle'), mp && first(mp, 'lstStyle')); }
    srcs.push(masterTx && first(masterTx, styleName));
    const out = { size: null, bold: null, color: key ? null : styleColor(sp, 'fontRef'), algn: null, bullet: null };
    const tag = 'lvl' + (lvl + 1) + 'pPr';
    for (const src of srcs) {
      const pp = src && kid(src, 'a', tag); if (!pp) continue;
      if (out.algn == null && pp.getAttribute('algn')) out.algn = pp.getAttribute('algn');
      if (out.bullet == null) { if (kid(pp, 'a', 'buNone')) out.bullet = ''; else if (kid(pp, 'a', 'buChar')) out.bullet = kid(pp, 'a', 'buChar').getAttribute('char'); else if (kid(pp, 'a', 'buAutoNum')) out.bullet = '•'; }
      const d = kid(pp, 'a', 'defRPr');
      if (d) {
        if (out.size == null && d.getAttribute('sz')) out.size = +d.getAttribute('sz') / 100;
        if (out.bold == null && d.getAttribute('b') != null) out.bold = d.getAttribute('b') === '1';
        if (out.color == null) { const c = colorOf(kid(d, 'a', 'solidFill'), theme); if (c) out.color = c; }
      }
    }
    if (out.size == null) out.size = isTitle ? 44 : type === 'subTitle' ? 24 : key ? 28 - lvl * 4 : 18;
    if (out.color == null) out.color = styleColor(sp, 'fontRef') || applyMods(hexToRgb(theme.dk1 || '000000'), document.createElementNS(NS.a, 'x'));
    if (out.bullet == null) out.bullet = '';
    if (!key && out.bullet) out.bullet = '';
    return out;
  }
  function drawText(ctx, txBody, b, sp, L) {
    const bodyPr = kid(txBody, 'a', 'bodyPr');
    const ins = k => (bodyPr && bodyPr.getAttribute(k) != null ? +bodyPr.getAttribute(k) : (k === 'lIns' || k === 'rIns' ? 91440 : 45720)) * scale;
    const anchor = (bodyPr && bodyPr.getAttribute('anchor')) || (phKey(sp) && /title/i.test(phKey(sp).type) ? 'ctr' : 't');
    const fontScale = (() => { const na = bodyPr && first(bodyPr, 'normAutofit'); return na && na.getAttribute('fontScale') ? +na.getAttribute('fontScale') / 100000 : 1; })();
    const defFace = phKey(sp) && /title/i.test(phKey(sp).type) ? fonts.head : fonts.body;
    const left = b.x + ins('lIns'), top = b.y + ins('tIns'), maxW = Math.max(10, b.w - ins('lIns') - ins('rIns'));
    const lines = []; // {text runs per wrapped line, align, height}
    for (const p of kids(txBody, 'a', 'p')) {
      const pPr = kid(p, 'a', 'pPr');
      const lvl = +(pPr && pPr.getAttribute('lvl') || 0);
      const D = levelProps(sp, lvl);
      const algn = (pPr && pPr.getAttribute('algn')) || D.algn || 'l';
      const runs = [];
      [...p.children].forEach(r => {
        if (r.localName === 'r' || r.localName === 'fld') {
          const t = kid(r, 'a', 't'); const rPr = kid(r, 'a', 'rPr');
          const size = (rPr && rPr.getAttribute('sz') ? +rPr.getAttribute('sz') / 100 : D.size) * fontScale;
          const bold = rPr && rPr.getAttribute('b') != null ? rPr.getAttribute('b') === '1' : !!D.bold;
          const ital = rPr && rPr.getAttribute('i') === '1';
          const col = (rPr && colorOf(kid(rPr, 'a', 'solidFill'), theme)) || D.color;
          const latin = rPr && kid(rPr, 'a', 'latin');
          const face = latin && !/^\+/.test(latin.getAttribute('typeface')) ? latin.getAttribute('typeface') : defFace;
          runs.push({ text: t ? t.textContent : '', size, bold, ital, col, face });
        } else if (r.localName === 'br') runs.push({ br: true });
      });
      const endSize = ((kid(p, 'a', 'endParaRPr') && kid(p, 'a', 'endParaRPr').getAttribute('sz')) ? +kid(p, 'a', 'endParaRPr').getAttribute('sz') / 100 : D.size) * fontScale;
      const noBullet = pPr && kid(pPr, 'a', 'buNone');
      const buChar = pPr && kid(pPr, 'a', 'buChar');
      const bullet = noBullet ? '' : buChar ? buChar.getAttribute('char') : D.bullet;
      const text = runs.map(r => r.br ? '\n' : r.text).join('');
      const fr = runs.find(r => !r.br) || { size: endSize, bold: !!D.bold, ital: false, col: D.color, face: defFace };
      const indent = lvl * 28 * EMU * scale;
      const px = fr.size * scale * EMU;
      ctx.font = `${fr.ital ? 'italic ' : ''}${fr.bold ? 'bold ' : ''}${px}px ${cssFont(fr.face)}`;
      const prefix = bullet && text.trim() ? bullet + ' ' : '';
      const wrapped = text.trim() || runs.length ? wrapLines(ctx, prefix + text, maxW - indent) : [''];
      wrapped.forEach(w => lines.push({ text: w, font: ctx.font, color: fr.col, px, align: algn, indent }));
      if (!runs.length) lines[lines.length - 1].px = endSize * scale * EMU;
    }
    const lh = l => l.px * 1.18;
    const total = lines.reduce((a, l) => a + lh(l), 0);
    let y = anchor === 'ctr' ? b.y + (b.h - total) / 2 : anchor === 'b' ? b.y + b.h - ins('bIns') - total : top;
    ctx.save(); ctx.textBaseline = 'top';
    lines.forEach(l => {
      ctx.font = l.font; ctx.fillStyle = l.color;
      const w = ctx.measureText(l.text).width;
      const x = l.align === 'ctr' ? b.x + (b.w - w) / 2 : l.align === 'r' ? b.x + b.w - ins('rIns') - w : left + l.indent;
      ctx.fillText(l.text, x, y + l.px * 0.09); y += lh(l);
    });
    ctx.restore();
  }
  async function drawPic(ctx, pic, L, tf) {
    const b = box(pic, L, tf); if (!b) return;
    const blip = first(pic, 'blip'); const id = blip && blip.getAttributeNS(NS.r, 'embed');
    if (!id || !L.rels[id] || L.rels[id].external) return;
    const im = await loadImage(zip, L.rels[id].target, cache); if (!im) return;
    const src = first(pic, 'srcRect');
    let sx0 = 0, sy0 = 0, sw = im.width, sh = im.height;
    if (src) { const g = k => +(src.getAttribute(k) || 0) / 100000; sx0 = im.width * g('l'); sy0 = im.height * g('t'); sw = im.width * (1 - g('l') - g('r')); sh = im.height * (1 - g('t') - g('b')); }
    withRot(ctx, b, () => ctx.drawImage(im, sx0, sy0, Math.max(1, sw), Math.max(1, sh), b.x, b.y, b.w, b.h));
  }
  async function drawTable(ctx, gf, L, tf) {
    const tbl = first(gf, 'tbl'); if (!tbl) return;
    const x = (() => { const xf = kid(gf, 'p', 'xfrm'); if (!xf) return null; const o = kid(xf, 'a', 'off'), e = kid(xf, 'a', 'ext'); return { x: tf.x + +o.getAttribute('x') * tf.sx, y: tf.y + +o.getAttribute('y') * tf.sy }; })();
    if (!x) return;
    const cols = [...desc(tbl, 'gridCol')].map(c => +c.getAttribute('w') * tf.sx);
    let y = x.y;
    [...desc(tbl, 'tr')].forEach((tr, ri) => {
      const h = +tr.getAttribute('h') * tf.sy; let cx = x.x;
      [...kids(tr, 'a', 'tc')].forEach((tc, ci) => {
        const w = cols[ci] || 0;
        const tcPr = kid(tc, 'a', 'tcPr'); const f = tcPr && colorOf(kid(tcPr, 'a', 'solidFill'), theme);
        ctx.fillStyle = f || (ri === 0 ? applyMods(hexToRgb(theme.accent1 || '4472C4'), document.createElementNS(NS.a, 'x')) : (ri % 2 ? '#e9edf4' : '#f6f8fb'));
        ctx.fillRect(cx, y, w, h); ctx.strokeStyle = '#ffffff'; ctx.lineWidth = 1.5; ctx.strokeRect(cx, y, w, h);
        const txt = [...desc(tc, 't')].map(t => t.textContent).join(' ');
        const rp = first(tc, 'rPr'); const size = (rp && rp.getAttribute('sz') ? +rp.getAttribute('sz') / 100 : 14) * scale * EMU;
        ctx.font = `${ri === 0 ? 'bold ' : ''}${size}px ${cssFont(fonts.body)}`; ctx.fillStyle = ri === 0 && !f ? '#fff' : '#111'; ctx.textBaseline = 'top';
        wrapLines(ctx, txt, w - 12).slice(0, Math.max(1, Math.floor(h / (size * 1.2)))).forEach((l, i) => ctx.fillText(l, cx + 6, y + 5 + i * size * 1.2));
        cx += w;
      });
      y += h;
    });
  }
}
function cssFont(face) { return `"${(face || 'Calibri').replace(/"/g, '')}", Calibri, Carlito, Arial, Helvetica, sans-serif`; }
async function readTheme(zip) {
  const pres = parseXml(await zip.file('ppt/presentation.xml').async('text'));
  const prels = await readRels(zip, 'ppt/presentation.xml');
  const themeRel = Object.values(prels).find(r => /\/theme$/.test(r.type));
  let theme = {}, fonts = { body: 'Calibri', head: 'Calibri Light' };
  // theme usually belongs to the master; fall back to the first theme part
  let themePath = themeRel && themeRel.target;
  if (!themePath) { const t = Object.keys(zip.files).find(p => /^ppt\/theme\/theme\d+\.xml$/.test(p)); themePath = t; }
  if (themePath && zip.file(themePath)) {
    const tx = parseXml(await zip.file(themePath).async('text'));
    const cs = first(tx, 'clrScheme');
    if (cs) [...cs.children].forEach(c => { const s = first(c, 'srgbClr'), y = first(c, 'sysClr'); theme[c.localName] = s ? s.getAttribute('val') : y ? (y.getAttribute('lastClr') || '000000') : '000000'; });
    const mi = first(tx, 'minorFont'), ma = first(tx, 'majorFont');
    if (mi && first(mi, 'latin')) fonts.body = first(mi, 'latin').getAttribute('typeface');
    if (ma && first(ma, 'latin')) fonts.head = first(ma, 'latin').getAttribute('typeface');
  }
  const sz = first(pres, 'sldSz');
  const W = sz ? +sz.getAttribute('cx') : 12192000, H = sz ? +sz.getAttribute('cy') : 6858000;
  const order = [...desc(pres, 'sldId')].map(s => prels[s.getAttributeNS(NS.r, 'id')]).filter(Boolean).map(r => r.target);
  return { theme, fonts, W, H, order };
}
let pp2pFile = null, pp2pSlides = [], pp2pInfo = null;
setupDrop('#pp2p-drop', '#pp2p-input', byExt(['pptx']), files => {
  const f = files[0];
  if (ext(f.name) !== 'pptx') return toast(ext(f.name) === 'ppt' ? 'Old .ppt files aren\'t supported. Save it as .pptx first.' : 'Please choose a .pptx file', true);
  pp2pFile = f; pp2pSlides = [];
  $('#pp2p-name').textContent = '📽 ' + f.name;
  $('#pp2p-options').hidden = false; $('#pp2p-result').hidden = true; $('#pp2p-slides').innerHTML = '';
  $('#pp2p-drop').classList.add('compact');
});
$('#pp2p-go').onclick = async () => {
  if (!pp2pFile) return;
  busy(true, 'Reading presentation…');
  try {
    const zip = await JSZip.loadAsync(await pp2pFile.arrayBuffer());
    if (!zip.file('ppt/presentation.xml')) throw new Error('This doesn\'t look like a PowerPoint (.pptx) file');
    const info = await readTheme(zip); pp2pInfo = info;
    const targetW = 1600, scale = targetW / info.W, Wpx = targetW, Hpx = Math.round(info.H * scale);
    const cache = {}; pp2pSlides = [];
    const grid = $('#pp2p-slides'); grid.innerHTML = '';
    if (document.fonts && document.fonts.ready) await document.fonts.ready;
    for (let i = 0; i < info.order.length; i++) {
      say(`Building slide ${i + 1} of ${info.order.length}…`); await nextFrame();
      const c = document.createElement('canvas'); c.width = Wpx; c.height = Hpx;
      const ctx = c.getContext('2d');
      try { await renderSlide(ctx, null, info.order[i], zip, { theme: info.theme, fonts: info.fonts, cache, scale, W: Wpx, H: Hpx }); }
      catch (e) { console.warn('slide', i + 1, e); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, Wpx, Hpx); ctx.fillStyle = '#999'; ctx.font = '40px sans-serif'; ctx.fillText('This slide could not be rebuilt', 60, 100); }
      pp2pSlides.push(c);
      const card = document.createElement('div'); card.className = 'file-card'; card.style.cursor = 'default'; card.style.setProperty('--i', Math.min(i, 20));
      card.innerHTML = `<span class="badge">${i + 1}</span><div class="thumb"></div><div class="meta">Slide ${i + 1}</div>`;
      const t = document.createElement('img'); t.src = c.toDataURL('image/jpeg', 0.7); t.alt = 'Slide ' + (i + 1); $('.thumb', card).appendChild(t); grid.appendChild(card);
    }
    if (!pp2pSlides.length) throw new Error('No slides found in this presentation');
    $('#pp2p-count').textContent = `✅ ${pp2pSlides.length} slide${pp2pSlides.length > 1 ? 's' : ''} ready. Check the preview below.`;
    $('#pp2p-result').hidden = false;
  } catch (e) { console.error(e); toast(e.message && e.message.length < 120 ? e.message : 'This presentation could not be read. It may be damaged or password-protected.', true); }
  busy(false);
};
$('#pp2p-dl').onclick = async () => {
  if (!pp2pSlides.length) return;
  busy(true, 'Creating PDF…');
  try {
    const doc = await PDFDocument.create();
    const pw = pp2pInfo.W / EMU, ph = pp2pInfo.H / EMU;
    for (let i = 0; i < pp2pSlides.length; i++) {
      say(`Adding slide ${i + 1} of ${pp2pSlides.length}…`); await nextFrame();
      const img = await doc.embedJpg(await canvasToJpeg(pp2pSlides[i], 0.9));
      doc.addPage([pw, ph]).drawImage(img, { x: 0, y: 0, width: pw, height: ph });
    }
    download(await doc.save(), baseName(pp2pFile.name) + '.pdf');
    track(1, pp2pSlides.length, pp2pFile.size);
    toast('Your PDF is ready!');
  } catch (e) { console.error(e); toast('Could not create the PDF', true); }
  busy(false);
};

/* =====================================================================
   PDF text helpers (used by PDF → Word and PDF → text)
   ===================================================================== */
async function pageLines(page) {
  const tc = await page.getTextContent();
  let fontNames = {};
  try {
    await page.getOperatorList();
    tc.items.forEach(it => { if (it.fontName && !(it.fontName in fontNames)) { try { const f = page.commonObjs.get(it.fontName); fontNames[it.fontName] = (f && (f.name || f.loadedName)) || ''; } catch { fontNames[it.fontName] = ''; } } });
  } catch {}
  const items = tc.items.filter(it => it.str !== undefined).map(it => {
    const t = it.transform, size = Math.hypot(t[2], t[3]) || Math.hypot(t[0], t[1]) || 10;
    const fn = fontNames[it.fontName] || '';
    return { str: it.str, x: t[4], y: t[5], w: it.width, size, bold: /bold|black|heavy|semibold/i.test(fn), italic: /italic|oblique/i.test(fn), eol: it.hasEOL };
  }).filter(it => it.str.length);
  items.sort((a, b) => (Math.abs(b.y - a.y) > Math.min(a.size, b.size) * 0.5 ? b.y - a.y : a.x - b.x));
  const lines = [];
  for (const it of items) {
    const L = lines[lines.length - 1];
    if (L && Math.abs(L.y - it.y) <= Math.min(L.size, it.size) * 0.5) {
      const gap = it.x - L.xEnd;
      const needSpace = gap > it.size * 0.18 && !/\s$/.test(L.text) && !/^\s/.test(it.str);
      L.runs.push({ text: (needSpace ? ' ' : '') + it.str, bold: it.bold, italic: it.italic, size: it.size });
      L.text += (needSpace ? ' ' : '') + it.str; L.xEnd = Math.max(L.xEnd, it.x + it.w); L.size = Math.max(L.size, it.size);
    } else lines.push({ y: it.y, x: it.x, xEnd: it.x + it.w, size: it.size, text: it.str, runs: [{ text: it.str, bold: it.bold, italic: it.italic, size: it.size }] });
  }
  return lines;
}
function toParagraphs(lines, pageW) {
  if (!lines.length) return [];
  const sizes = {}; lines.forEach(l => { const k = Math.round(l.size); sizes[k] = (sizes[k] || 0) + l.text.length; });
  const body = +Object.entries(sizes).sort((a, b) => b[1] - a[1])[0][0] || 11;
  const mL = Math.min(...lines.map(l => l.x)), mR = Math.max(...lines.map(l => l.xEnd)), textW = Math.max(1, mR - mL), mid = (mL + mR) / 2;
  const paras = []; let cur = null, prev = null;
  for (const l of lines) {
    const gap = prev ? prev.y - l.y : 0;
    const sizeJump = prev && Math.abs(Math.round(l.size) - Math.round(prev.size)) >= 2;
    const bullet = /^\s*([•●▪◦‣\-–]|\d{1,3}[.)])\s+/.test(l.text);
    const newPara = !cur || gap > prev.size * 1.75 || sizeJump || bullet || (prev && /[.:!?]\s*$/.test(prev.text) && prev.xEnd < mL + textW * 0.75);
    if (newPara) {
      const width = l.xEnd - l.x, center = (l.x + l.xEnd) / 2;
      cur = { runs: [], size: l.size, body, center: l.x > mL + textW * 0.12 && Math.abs(center - mid) < textW * 0.05 && width < textW * 0.75, bullet };
      paras.push(cur);
    } else if (cur.runs.length && !/[\s-]$/.test(cur.runs[cur.runs.length - 1].text)) cur.runs.push({ text: ' ', bold: false, italic: false, size: l.size });
    l.runs.forEach(r => cur.runs.push(r));
    prev = l;
  }
  return paras;
}

/* =====================================================================
   PDF → WORD
   ===================================================================== */
let p2wFile = null;
setupDrop('#p2w-drop', '#p2w-input', isPdf, files => {
  p2wFile = files[0]; $('#p2w-name').textContent = '📄 ' + p2wFile.name;
  $('#p2w-options').hidden = false; $('#p2w-drop').classList.add('compact');
});
$('#p2w-go').onclick = async () => {
  if (!p2wFile) return;
  busy(true, 'Reading PDF…');
  try {
    const D = await need.docx();
    const pdf = await openPdfJs(p2wFile);
    let mode = $('#p2w-mode').value, switched = false;
    const sections = [];
    if (mode === 'text') {
      const children = []; let chars = 0; let firstW = 612, firstH = 792;
      for (let n = 1; n <= pdf.numPages; n++) {
        say(`Reading page ${n} of ${pdf.numPages}…`); await nextFrame();
        const page = await pdf.getPage(n); const vp = page.getViewport({ scale: 1 });
        if (n === 1) { firstW = vp.width; firstH = vp.height; }
        const paras = toParagraphs(await pageLines(page), vp.width);
        paras.forEach((p, i) => {
          const text = p.runs.map(r => r.text).join(''); chars += text.trim().length;
          const ratio = p.size / p.body;
          const heading = ratio >= 1.7 ? D.HeadingLevel.HEADING_1 : ratio >= 1.25 ? D.HeadingLevel.HEADING_2 : undefined;
          children.push(new D.Paragraph({
            heading, alignment: p.center ? D.AlignmentType.CENTER : undefined,
            pageBreakBefore: n > 1 && i === 0,
            spacing: { after: 120 },
            children: p.runs.map(r => new D.TextRun({ text: r.text, bold: r.bold || undefined, italics: r.italic || undefined, size: heading ? undefined : Math.max(16, Math.round(r.size * 2)) })),
          }));
        });
      }
      if (chars < 10 * pdf.numPages) { mode = 'image'; switched = true; }
      else sections.push({ properties: { page: { size: { width: Math.round(firstW * 20), height: Math.round(firstH * 20) }, margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } } }, children });
    }
    if (mode === 'image') {
      for (let n = 1; n <= pdf.numPages; n++) {
        say(`Copying page ${n} of ${pdf.numPages}…`); await nextFrame();
        const page = await pdf.getPage(n); const vp1 = page.getViewport({ scale: 1 });
        const vp = page.getViewport({ scale: 2 });
        const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
        const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
        await page.render({ canvasContext: ctx, viewport: vp }).promise;
        const data = await canvasToJpeg(c, 0.88);
        sections.push({
          properties: { page: { size: { width: Math.round(vp1.width * 20), height: Math.round(vp1.height * 20) }, margin: { top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0 } } },
          children: [new D.Paragraph({ spacing: { before: 0, after: 0, line: 240 }, children: [new D.ImageRun({ type: 'jpg', data, transformation: { width: Math.floor(vp1.width * 96 / 72) - 2, height: Math.floor(vp1.height * 96 / 72) - 4 } })] })],
        });
      }
    }
    say('Creating Word file…'); await nextFrame();
    const blob = await D.Packer.toBlob(new D.Document({ creator: 'MyPDF', sections }));
    download(blob, baseName(p2wFile.name) + '.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    track(1, pdf.numPages, p2wFile.size);
    toast(switched ? 'Done! This PDF has no real text (likely a scan), so pages were kept as pictures.' : 'Your Word file is ready!');
  } catch (e) { console.error(e); toast(e.message && e.message.length < 120 ? e.message : 'This PDF could not be converted. It may be damaged or password-protected.', true); }
  busy(false);
};

/* =====================================================================
   PDF → TEXT
   ===================================================================== */
let p2tFile = null;
setupDrop('#p2t-drop', '#p2t-input', isPdf, files => {
  p2tFile = files[0]; $('#p2t-name').textContent = '📄 ' + p2tFile.name;
  $('#p2t-options').hidden = false; $('#p2t-out').hidden = true; $('#p2t-drop').classList.add('compact');
});
$('#p2t-go').onclick = async () => {
  if (!p2tFile) return;
  busy(true, 'Reading PDF…');
  try {
    const pdf = await openPdfJs(p2tFile); const marks = $('#p2t-marks').checked; const parts = [];
    for (let n = 1; n <= pdf.numPages; n++) {
      say(`Reading page ${n} of ${pdf.numPages}…`); await nextFrame();
      const page = await pdf.getPage(n); const vp = page.getViewport({ scale: 1 });
      const paras = toParagraphs(await pageLines(page), vp.width).map(p => p.runs.map(r => r.text).join('').replace(/\s+/g, ' ').trim()).filter(Boolean);
      parts.push((marks ? `----- Page ${n} -----\n\n` : '') + paras.join('\n\n'));
    }
    const text = parts.join('\n\n').trim();
    $('#p2t-text').value = text;
    const words = text ? text.split(/\s+/).length : 0;
    $('#p2t-stats').textContent = words ? `${pdf.numPages} page${pdf.numPages > 1 ? 's' : ''} · ${words.toLocaleString()} words` : 'No text found. This PDF is probably a scan (pictures of pages).';
    $('#p2t-out').hidden = false;
    track(1, pdf.numPages, p2tFile.size);
    if (words) toast('Done! Your text is ready.');
  } catch (e) { console.error(e); toast('This PDF could not be read. It may be damaged or password-protected.', true); }
  busy(false);
};
$('#p2t-copy').onclick = async () => {
  const t = $('#p2t-text');
  try { await navigator.clipboard.writeText(t.value); } catch { t.select(); document.execCommand('copy'); }
  toast('Copied to clipboard');
};
$('#p2t-dl').onclick = () => download(new Blob([$('#p2t-text').value], { type: 'text/plain;charset=utf-8' }), baseName(p2tFile.name) + '.txt', 'text/plain');

/* =====================================================================
   ORGANIZE PAGES
   ===================================================================== */
let orgFile = null, orgPdf = null, orgPages = [], orgThumbs = {};
setupDrop('#org-drop', '#org-input', isPdf, async files => {
  orgFile = files[0];
  busy(true, 'Loading pages…');
  try {
    orgPdf = await openPdfJs(orgFile); orgThumbs = {};
    orgPages = Array.from({ length: orgPdf.numPages }, (_, i) => ({ src: i, rot: 0 }));
    $('#org-name').textContent = `📄 ${orgFile.name} · ${orgPdf.numPages} pages`;
    $('#org-options').hidden = false; $('#org-drop').classList.add('compact');
    renderOrg();
    for (let i = 0; i < orgPdf.numPages; i++) {
      if (i % 4 === 0) { say(`Loading page ${i + 1} of ${orgPdf.numPages}…`); await nextFrame(); }
      orgThumbs[i] = (await renderPdfPage(orgPdf, i + 1, 200)).toDataURL('image/jpeg', 0.75);
      const img = document.querySelector(`#org-pages img[data-src="${i}"]`); if (img) { img.src = orgThumbs[i]; img.hidden = false; }
      if (i === 11) busy(false); // show the first pages quickly, keep loading the rest
    }
  } catch (e) { console.error(e); toast('This PDF could not be read.', true); }
  busy(false);
});
function renderOrg() {
  const grid = $('#org-pages');
  grid.innerHTML = orgPages.map((p, i) => `
    <div class="file-card org-card" style="--i:${Math.min(i, 20)}"><span class="badge">${i + 1}</span>
      <button class="remove" data-a="del" data-i="${i}" title="Delete page">✕</button>
      <div class="thumb"><img data-src="${p.src}" ${orgThumbs[p.src] ? `src="${orgThumbs[p.src]}"` : 'hidden'} alt="Page ${p.src + 1}" style="transform:rotate(${p.rot}deg)${p.rot % 180 ? ' scale(.72)' : ''}"></div>
      <div class="org-tools">
        <button data-a="left" data-i="${i}" title="Move earlier" ${i === 0 ? 'disabled' : ''}>◀</button>
        <button data-a="rl" data-i="${i}" title="Rotate left">↺</button>
        <button data-a="rr" data-i="${i}" title="Rotate right">↻</button>
        <button data-a="right" data-i="${i}" title="Move later" ${i === orgPages.length - 1 ? 'disabled' : ''}>▶</button>
      </div>
      <div class="meta">Original page ${p.src + 1}</div>
    </div>`).join('');
  $$('button[data-a]', grid).forEach(b => b.onclick = ev => {
    ev.stopPropagation();
    const i = +b.dataset.i, a = b.dataset.a;
    if (a === 'del') { if (orgPages.length === 1) return toast('A PDF needs at least one page', true); orgPages.splice(i, 1); }
    else if (a === 'rl') orgPages[i].rot = (orgPages[i].rot + 270) % 360;
    else if (a === 'rr') orgPages[i].rot = (orgPages[i].rot + 90) % 360;
    else if (a === 'left' && i > 0) [orgPages[i - 1], orgPages[i]] = [orgPages[i], orgPages[i - 1]];
    else if (a === 'right' && i < orgPages.length - 1) [orgPages[i + 1], orgPages[i]] = [orgPages[i], orgPages[i + 1]];
    renderOrg();
  });
  makeSortable(grid, orgPages, renderOrg);
  $('#org-go').textContent = `Save PDF (${orgPages.length} page${orgPages.length > 1 ? 's' : ''})`;
}
$('#org-rotall').onclick = () => { orgPages.forEach(p => p.rot = (p.rot + 90) % 360); renderOrg(); };
$('#org-reset').onclick = () => { if (!orgPdf) return; orgPages = Array.from({ length: orgPdf.numPages }, (_, i) => ({ src: i, rot: 0 })); renderOrg(); };
$('#org-go').onclick = async () => {
  if (!orgFile || !orgPages.length) return;
  busy(true, 'Saving PDF…');
  try {
    const src = await loadPdfLib(await orgFile.arrayBuffer());
    const out = await PDFDocument.create();
    const copied = await out.copyPages(src, orgPages.map(p => p.src));
    copied.forEach((pg, i) => { const base = pg.getRotation().angle || 0; pg.setRotation(degrees(((base + orgPages[i].rot) % 360 + 360) % 360)); out.addPage(pg); });
    download(await out.save(), baseName(orgFile.name) + '-organized.pdf');
    track(1, orgPages.length, orgFile.size);
    toast('Done! Your organized PDF is ready.');
  } catch (e) { console.error(e); toast(e.message || 'Could not save this PDF', true); }
  busy(false);
};

/* =====================================================================
   COMPRESS
   ===================================================================== */
let cmpFile = null, cmpOut = null;
const LEVELS = { strong: { dpi: 72, q: 0.5 }, balanced: { dpi: 110, q: 0.62 }, light: { dpi: 150, q: 0.75 } };
setupDrop('#cmp-drop', '#cmp-input', isPdf, files => {
  cmpFile = files[0]; cmpOut = null;
  $('#cmp-name').textContent = `📄 ${cmpFile.name} · ${fmtSize(cmpFile.size)}`;
  $('#cmp-options').hidden = false; $('#cmp-result').hidden = true; $('#cmp-drop').classList.add('compact');
});
$('#cmp-go').onclick = async () => {
  if (!cmpFile) return;
  busy(true, 'Compressing…');
  try {
    const { dpi, q } = LEVELS[$('#cmp-level').value];
    const pdf = await openPdfJs(cmpFile); const out = await PDFDocument.create();
    for (let n = 1; n <= pdf.numPages; n++) {
      say(`Compressing page ${n} of ${pdf.numPages}…`); await nextFrame();
      const page = await pdf.getPage(n); const vp1 = page.getViewport({ scale: 1 }); const vp = page.getViewport({ scale: dpi / 72 });
      const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(vp.width)); c.height = Math.max(1, Math.round(vp.height));
      const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
      await page.render({ canvasContext: ctx, viewport: vp }).promise;
      const img = await out.embedJpg(await canvasToJpeg(c, q));
      out.addPage([vp1.width, vp1.height]).drawImage(img, { x: 0, y: 0, width: vp1.width, height: vp1.height });
      c.width = c.height = 0;
    }
    cmpOut = await out.save();
    const before = cmpFile.size, after = cmpOut.byteLength, pct = Math.round((1 - after / before) * 100);
    $('#cmp-sizes').innerHTML = after < before
      ? `✅ ${fmtSize(before)} → <b>${fmtSize(after)}</b> (${pct}% smaller)`
      : `⚠ The result (${fmtSize(after)}) isn't smaller than your file (${fmtSize(before)}). This PDF is already well compressed. Keep your original.`;
    $('#cmp-dl').hidden = after >= before;
    $('#cmp-result').hidden = false;
    track(1, pdf.numPages, before);
    if (after < before) toast('Done! Your PDF is smaller.');
  } catch (e) { console.error(e); toast('This PDF could not be compressed. It may be damaged or password-protected.', true); }
  busy(false);
};
$('#cmp-dl').onclick = () => cmpOut && download(cmpOut, baseName(cmpFile.name) + '-compressed.pdf');

/* =====================================================================
   PAGE NUMBERS & WATERMARK
   ===================================================================== */
let nwFile = null, nwPdf = null, nwBase = null;
setupDrop('#nw-drop', '#nw-input', isPdf, async files => {
  nwFile = files[0];
  busy(true, 'Loading PDF…');
  try {
    nwPdf = await openPdfJs(nwFile);
    $('#nw-name').textContent = `📄 ${nwFile.name} · ${nwPdf.numPages} pages`;
    $('#nw-workspace').hidden = false; $('#nw-drop').classList.add('compact');
    nwBase = await renderPdfPage(nwPdf, 1, 900);
    drawNwPreview();
  } catch (e) { console.error(e); toast('This PDF could not be read.', true); }
  busy(false);
});
function nwOpts() {
  return {
    num: $('#nw-num').checked, pos: $('#nw-pos').value, fmt: $('#nw-fmt').value, start: Math.max(0, parseInt($('#nw-start').value, 10) || 0),
    size: +$('#nw-size').value, skip: $('#nw-skip').checked,
    wm: $('#nw-wm').checked, text: $('#nw-text').value.trim(), diag: $('#nw-angle').value === 'diag', wsize: +$('#nw-wsize').value,
    color: $('#nw-color').value.split(',').map(Number), op: +$('#nw-op').value,
  };
}
function numLabel(o, i, total) {
  const n = o.start + i - (o.skip ? 1 : 0), N = o.start + total - 1 - (o.skip ? 1 : 0);
  return o.fmt === 'page' ? `Page ${n}` : o.fmt === 'of' ? `Page ${n} of ${N}` : o.fmt === 'slash' ? `${n} / ${N}` : String(n);
}
function drawNwPreview() {
  if (!nwBase) return;
  const o = nwOpts(), c = $('#nw-canvas'); c.width = nwBase.width; c.height = nwBase.height;
  const ctx = c.getContext('2d'); ctx.drawImage(nwBase, 0, 0);
  const k = c.width / (nwPdf ? 612 : 612); // approx px per pt (preview only)
  if (o.wm && o.text) {
    ctx.font = 'bold 100px Helvetica, Arial, sans-serif';
    const span = o.diag ? Math.hypot(c.width, c.height) : c.width;
    const size = Math.min(c.height * 0.3, 100 * span * o.wsize / ctx.measureText(o.text).width);
    ctx.save(); ctx.translate(c.width / 2, c.height / 2); if (o.diag) ctx.rotate(-Math.atan2(c.height, c.width));
    ctx.font = `bold ${size}px Helvetica, Arial, sans-serif`; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillStyle = `rgba(${o.color.map(v => Math.round(v * 255)).join(',')},${o.op})`; ctx.fillText(o.text, 0, 0); ctx.restore();
  }
  if (o.num && !(o.skip)) {
    const t = numLabel(o, 0, nwPdf.numPages), size = o.size * k, m = 28 * k;
    ctx.font = `${size}px Helvetica, Arial, sans-serif`; ctx.fillStyle = '#222'; ctx.textBaseline = o.pos[0] === 't' ? 'top' : 'bottom';
    ctx.textAlign = o.pos[1] === 'c' ? 'center' : o.pos[1] === 'r' ? 'right' : 'left';
    ctx.fillText(t, o.pos[1] === 'c' ? c.width / 2 : o.pos[1] === 'r' ? c.width - m : m, o.pos[0] === 't' ? m : c.height - m);
  }
}
['#nw-num', '#nw-pos', '#nw-fmt', '#nw-start', '#nw-size', '#nw-skip', '#nw-wm', '#nw-text', '#nw-angle', '#nw-wsize', '#nw-color', '#nw-op'].forEach(id => {
  $(id).addEventListener('input', drawNwPreview); $(id).addEventListener('change', drawNwPreview);
});
const toggleFields = () => { $('#nw-num-fields').classList.toggle('off', !$('#nw-num').checked); $('#nw-wm-fields').classList.toggle('off', !$('#nw-wm').checked); };
$('#nw-num').addEventListener('change', toggleFields); $('#nw-wm').addEventListener('change', toggleFields); toggleFields();
// draw text at "visual" coordinates (as the page is displayed), whatever the page's /Rotate value
function visualText(page, font, text, vx, vy, size, angle, color, opacity) {
  const rot = ((page.getRotation().angle || 0) % 360 + 360) % 360;
  const { width: W, height: H } = page.getSize(); const cb = page.getCropBox ? page.getCropBox() : { x: 0, y: 0 };
  let x, y;
  if (rot === 90) { x = W - vy; y = vx; }
  else if (rot === 180) { x = W - vx; y = H - vy; }
  else if (rot === 270) { x = vy; y = H - vx; }
  else { x = vx; y = vy; }
  page.drawText(text, { x: x + (cb.x || 0), y: y + (cb.y || 0), size, font, color, opacity, rotate: degrees(angle + rot) });
}
$('#nw-go').onclick = async () => {
  if (!nwFile) return;
  const o = nwOpts();
  if (!o.num && !(o.wm && o.text)) return toast('Turn on page numbers or a watermark first', true);
  busy(true, 'Stamping pages…');
  try {
    const doc = await loadPdfLib(await nwFile.arrayBuffer());
    const font = await doc.embedFont(StandardFonts.Helvetica), bold = await doc.embedFont(StandardFonts.HelveticaBold);
    try { if (o.wm) bold.encodeText(o.text); } catch { throw new Error('The watermark has letters this font can\'t show. Use basic Latin letters only.'); }
    const pages = doc.getPages(), total = pages.length;
    pages.forEach((page, i) => {
      const rot = ((page.getRotation().angle || 0) % 360 + 360) % 360;
      const { width, height } = page.getSize();
      const VW = rot % 180 ? height : width, VH = rot % 180 ? width : height;
      if (o.wm && o.text) {
        const span = o.diag ? Math.hypot(VW, VH) : VW;
        const size = Math.min(VH * 0.3, span * o.wsize / Math.max(1, bold.widthOfTextAtSize(o.text, 1)));
        const tw = bold.widthOfTextAtSize(o.text, size), th = size * 0.7;
        const a = o.diag ? Math.atan2(VH, VW) : 0;
        const cx = VW / 2, cy = VH / 2;
        const sx = cx - (tw / 2) * Math.cos(a) + (th / 2) * Math.sin(a), sy = cy - (tw / 2) * Math.sin(a) - (th / 2) * Math.cos(a);
        visualText(page, bold, o.text, sx, sy, size, a * 180 / Math.PI, rgb(...o.color), o.op);
      }
      if (o.num && !(o.skip && i === 0)) {
        const t = numLabel(o, i, total), size = o.size, m = 28, tw = font.widthOfTextAtSize(t, size);
        const vx = o.pos[1] === 'c' ? (VW - tw) / 2 : o.pos[1] === 'r' ? VW - m - tw : m;
        const vy = o.pos[0] === 't' ? VH - m - size * 0.75 : m;
        visualText(page, font, t, vx, vy, size, 0, rgb(0.13, 0.13, 0.13), 1);
      }
    });
    download(await doc.save(), baseName(nwFile.name) + '-stamped.pdf');
    track(1, total, nwFile.size);
    toast('Done! Your PDF is ready.');
  } catch (e) { console.error(e); toast(e.message || 'Could not stamp this PDF', true); }
  busy(false);
};
})();
