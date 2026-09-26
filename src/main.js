import {
  app, on, emit, change, undo, redo, setDoc, select, clearSelection, scheduleSave, loadSaved,
  selectedNodes, selectedEdges, nodeById, clone,
} from './state.js';
import {
  initCanvas, renderCanvas, renderOverlay, fit, zoomAt, setZoom, centerOn, addShape, editNodeText,
  editEdgeLabel, commitInline, isInlineEditing, isTyping, closeMenus, applyView, quickAdd, isDragging, cancelDrag,
} from './canvas.js';
import { initInspector, renderInspector, addLane, addPhase, deleteLane, h } from './inspector.js';
import { SHAPES, SHAPE_ORDER } from './shapes.js';
import { themeCSS, fontFor, FONTS, clearMeasureCache, renderSVG } from './render.js';
import { buildStandaloneHTML, exportSVGString, exportPNGBlob, download, fileBase, parseDocumentText, canEmbedFonts } from './io.js';
import { TEMPLATES } from './templates.js';
import { uid, makeNode } from './model.js';
import { parseOutline, OUTLINE_EXAMPLE, autoLayout, numberTags } from './layout.js';
import { parseDelimited, readXlsx, guessRoles, buildFromRows, ROLES, EXAMPLE_CSV } from './spreadsheet.js';
import { unionBox, center } from './geometry.js';

const $ = (s) => document.querySelector(s);
const THEME_KEY = 'lagos.theme';

// ---------------------------------------------------------------------------
// theme & fonts

const themeStyle = document.createElement('style');
themeStyle.textContent = themeCSS();
document.head.prepend(themeStyle);
const fontVars = document.createElement('style');
document.head.append(fontVars);
let currentFont = null;

function applyFont() {
  const key = app.doc.settings.font;
  if (key === currentFont) return;
  currentFont = key;
  const F = fontFor(app.doc);
  fontVars.textContent = `:root{--ui:${F.ui};--cond:${F.text};--mono:${F.mono}}`;
  clearMeasureCache();
  if (document.fonts?.ready) document.fonts.ready.then(() => { clearMeasureCache(); renderCanvas(); });
}

function applyEditorTheme() {
  let t = null;
  try { t = localStorage.getItem(THEME_KEY); } catch (_) { /* ignore */ }
  if (t === 'light' || t === 'dark') document.documentElement.setAttribute('data-theme', t);
  else document.documentElement.removeAttribute('data-theme');
}

function toggleTheme() {
  const cur = document.documentElement.getAttribute('data-theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  const next = cur === 'dark' ? 'light' : 'dark';
  try { localStorage.setItem(THEME_KEY, next); } catch (_) { /* ignore */ }
  applyEditorTheme();
  if (app.mode === 'preview') refreshPreview();
}

// ---------------------------------------------------------------------------
// toasts

let toastTimer = null;
export function toast(msg, ms = 2200) {
  let t = $('.toast');
  if (!t) { t = document.createElement('div'); t.className = 'toast'; t.setAttribute('role', 'status'); document.body.append(t); }
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.hidden = true; }, ms);
}

// ---------------------------------------------------------------------------
// palette

function shapeIcon(key) {
  const def = SHAPES[key];
  const w = def.w, hh = def.h;
  const parts = def.draw(w, hh);
  const inner = parts.map(([t, a], i) => {
    const attrs = { fill: i === 0 || def.layered ? (def.fill === 'transparent' ? 'none' : def.fill || 'var(--node)') : 'none', stroke: 'var(--node-stroke)', 'stroke-width': 1.2, 'vector-effect': 'non-scaling-stroke', 'stroke-dasharray': def.dash && i === 0 ? '3 2' : undefined, ...a };
    if (a.stroke === 'none' && key === 'text') return '';
    return `<${t} ${Object.entries(attrs).filter(([, v]) => v !== undefined).map(([k, v]) => `${k}="${v}"`).join(' ')}/>`;
  }).join('');
  const extra = key === 'text' ? `<text x="${w / 2}" y="${hh / 2 + 9}" text-anchor="middle" font-size="26" font-weight="600" fill="var(--muted)">T</text>` : '';
  return `<svg width="38" height="24" viewBox="-3 -3 ${w + 6} ${hh + 6}" preserveAspectRatio="xMidYMid meet" aria-hidden="true">${inner}${extra}</svg>`;
}

const SHORT_NAMES = {
  terminator: 'Start / end', document: 'Document', multidoc: 'Documents', data: 'Data / I/O', manualInput: 'Manual input',
  database: 'Data store', predefined: 'Subprocess', preparation: 'Prepare', manualOp: 'Manual op', storedData: 'Stored data',
  rounded: 'Rounded', offpage: 'Off-page', cloud: 'Cloud', actor: 'Person', annotation: 'Comment', group: 'Group',
};

function buildPalette() {
  const list = $('#shape-list');
  list.innerHTML = '';
  for (const key of SHAPE_ORDER) {
    const b = document.createElement('button');
    b.className = 'shape-btn';
    b.type = 'button';
    b.draggable = true;
    b.dataset.shape = key;
    b.title = SHAPES[key].name + ' — click to add, or drag onto the canvas';
    b.innerHTML = shapeIcon(key) + `<span>${SHORT_NAMES[key] || SHAPES[key].name}</span>`;
    b.addEventListener('click', () => { if (app.mode !== 'edit') return; addShape(key); $('#app').classList.remove('show-palette'); });
    b.addEventListener('dragstart', (ev) => {
      ev.dataTransfer.setData('application/x-flow-shape', key);
      ev.dataTransfer.effectAllowed = 'copy';
    });
    list.append(b);
  }
}

function buildConnectorDefaults() {
  const box = $('#connector-defaults');
  const kinds = app.doc.edgeKinds;
  const kindSel = h('select', { id: 'default-kind', 'aria-label': 'New connector type' }, kinds.map((k) => h('option', { value: k.id, selected: k.id === (app.connectDefaults.kind || kinds[0].id) }, k.name)));
  kindSel.addEventListener('change', () => { app.connectDefaults.kind = kindSel.value; });
  const rSel = h('select', { id: 'default-routing', 'aria-label': 'New connector routing' },
    [['orthogonal', 'Elbow'], ['straight', 'Straight'], ['curved', 'Curved']].map(([v, l]) => h('option', { value: v, selected: v === app.connectDefaults.routing }, l)));
  rSel.addEventListener('change', () => { app.connectDefaults.routing = rSel.value; });
  box.replaceChildren(kindSel, rSel, h('p', { class: 'hint-text' }, 'Hover a shape and drag from a blue dot to connect. Drop on empty space to create a new shape.'));
}

// ---------------------------------------------------------------------------
// commands

let clipboard = null;
let pasteCount = 0;

function copySelection() {
  const nodes = selectedNodes();
  const ids = new Set(nodes.map((n) => n.id));
  const inSel = (e, ep) => (ep.node ? ids.has(ep.node) : app.sel.edges.has(e.id) || ids.size > 0);
  const edges = app.doc.edges.filter((e) => (app.sel.edges.has(e.id) || ids.size) && inSel(e, e.from) && inSel(e, e.to));
  if (!nodes.length && !edges.length) return false;
  clipboard = { fcClip: 1, nodes: clone(nodes), edges: clone(edges) };
  pasteCount = 0;
  try { navigator.clipboard?.writeText(JSON.stringify(clipboard)).catch(() => {}); } catch (_) { /* ignore */ }
  return true;
}

function pasteClip(clip, offset) {
  if (!clip) return;
  const map = new Map();
  const nodes = clip.nodes.map((n) => { const c = clone(n); c.id = uid('n'); map.set(n.id, c.id); c.x += offset; c.y += offset; return c; });
  const edges = clip.edges.map((e) => {
    const c = clone(e); c.id = uid('e');
    for (const k of ['from', 'to']) {
      if (c[k].node) c[k].node = map.get(c[k].node);
      else c[k] = { x: c[k].x + offset, y: c[k].y + offset };
    }
    c.waypoints = c.waypoints.map(([x, y]) => [x + offset, y + offset]);
    return c;
  }).filter((e) => (e.from.node !== undefined || 'x' in e.from) && (e.to.node !== undefined || 'x' in e.to));
  change((doc) => { doc.nodes.push(...nodes); doc.edges.push(...edges); });
  select({ nodes: nodes.map((n) => n.id), edges: edges.map((e) => e.id) });
}

function deleteSelection() {
  if (app.sel.phase && !app.sel.nodes.size && !app.sel.edges.size) {
    const id = app.sel.phase;
    change((doc) => { doc.phases.items = doc.phases.items.filter((l) => l.id !== id); });
    return;
  }
  if (app.sel.lane && !app.sel.nodes.size && !app.sel.edges.size) {
    deleteLane(app.sel.lane);
    return;
  }
  const nodeIds = new Set(app.sel.nodes);
  const edgeIds = new Set(app.sel.edges);
  if (!nodeIds.size && !edgeIds.size) return;
  change((doc) => {
    doc.nodes = doc.nodes.filter((n) => !nodeIds.has(n.id));
    doc.edges = doc.edges.filter((e) => !edgeIds.has(e.id) && !nodeIds.has(e.from.node) && !nodeIds.has(e.to.node));
  });
  clearSelection();
}

function frameSelection() {
  const nodes = selectedNodes();
  if (!nodes.length) return;
  const u = unionBox(nodes);
  const g = makeNode('group', u.x - 20, u.y - 40, { w: u.w + 40, h: u.h + 60, text: 'Group' });
  change((doc) => { doc.nodes.unshift(g); });
  select({ nodes: [g.id] });
}

// Swap lane orientation. Lane coordinates are kept; the flow axis is transposed and
// stretched (or squeezed) just enough that no two shapes overlap.
function transposeLanes(doc) {
  const toH = doc.lanes.orientation !== 'horizontal';
  const H = doc.lanes.items.length ? doc.lanes.header || 40 : 0;
  const along = (v, f) => H + (v - H) * f;
  const place = (f) => doc.nodes.map((n) => {
    const cx = n.x + n.w / 2, cy = n.y + n.h / 2;
    return toH ? { x: along(cy, f) - n.w / 2, y: cx - n.h / 2, w: n.w, h: n.h } : { x: cy - n.w / 2, y: along(cx, f) - n.h / 2, w: n.w, h: n.h };
  });
  const solid = doc.nodes.map((n) => !SHAPES[n.shape]?.container);
  const overlaps = (boxes) => boxes.some((a, i) => solid[i] && boxes.some((b, j) => j > i && solid[j] && a.x < b.x + b.w + 24 && a.x + a.w + 24 > b.x && a.y < b.y + b.h + 24 && a.y + a.h + 24 > b.y));
  let f = 4;
  for (let t = 0.4; t <= 4; t += 0.05) { if (!overlaps(place(t))) { f = t; break; } }
  f = Math.round(f * 100) / 100;
  const boxes = place(f);
  doc.nodes.forEach((n, i) => { n.x = Math.round(boxes[i].x); n.y = Math.round(boxes[i].y); });
  const tp = ([x, y]) => (toH ? [along(y, f), x] : [y, along(x, f)]);
  for (const e of doc.edges) {
    e.waypoints = e.waypoints.map(tp);
    for (const k of ['from', 'to']) {
      if (!e[k].node) { const [x, y] = tp([e[k].x, e[k].y]); e[k] = { x, y }; }
      else if (e[k].side && e[k].side !== 'auto') e[k].side = { t: 'l', l: 't', b: 'r', r: 'b' }[e[k].side];
    }
  }
  for (const p of doc.phases.items) p.size = Math.round(p.size * f);
  doc.lanes.orientation = toH ? 'horizontal' : 'vertical';
}

// ---------------------------------------------------------------------------
// format painter

const NODE_STYLE_KEYS = ['fill', 'stroke', 'strokeWidth', 'dash', 'textColor', 'fontSize', 'bold', 'italic', 'align', 'valign', 'opacity', 'shadow', 'tagColor'];

function styleOf(sel) {
  const nodes = selectedNodes(), edges = selectedEdges();
  if (nodes.length === 1 && !edges.length) return { kind: 'node', style: clone(nodes[0].style || {}) };
  if (edges.length === 1 && !nodes.length) return { kind: 'edge', edgeKind: edges[0].kind, style: clone(edges[0].style || {}) };
  return null;
}

function applyStyle(src, nodes, edges) {
  for (const n of nodes) {
    const st = { ...(n.style || {}) };
    for (const k of NODE_STYLE_KEYS) { if (k in src.style) st[k] = clone(src.style[k]); else delete st[k]; }
    n.style = st;
  }
  for (const e of edges) { e.kind = src.edgeKind; e.style = clone(src.style); }
}

function startPainter(sticky) {
  const src = styleOf();
  if (!src) { toast('Select one shape or connector to copy its formatting'); return; }
  app.painter = { ...src, sticky };
  updatePainterUI();
}
function stopPainter() { app.painter = null; updatePainterUI(); }
function updatePainterUI() {
  const b = $('#painter-btn');
  b.classList.toggle('on', !!app.painter);
  b.setAttribute('aria-pressed', app.painter ? 'true' : 'false');
  $('#stage').classList.toggle('painting', !!app.painter);
  updateStatus();
}

function onPaint(target) {
  const P = app.painter;
  if (!P) return;
  if (target.none) { stopPainter(); return; }
  if (target.node && P.kind === 'node') { const n = nodeById(target.node); change(() => applyStyle(P, [n], [])); }
  else if (target.edge && P.kind === 'edge') { const e = app.doc.edges.find((x) => x.id === target.edge); change(() => applyStyle(P, [], [e])); }
  else { toast(P.kind === 'node' ? 'Shape formatting can only be pasted onto shapes' : 'Connector formatting can only be pasted onto connectors'); return; }
  if (!P.sticky) stopPainter();
}

let styleClip = null;
function copyStyle() { styleClip = styleOf(); toast(styleClip ? 'Formatting copied' : 'Select one shape or connector to copy its formatting'); }
function pasteStyle() {
  if (!styleClip) { toast('Copy formatting first (Ctrl+Alt+C)'); return; }
  const nodes = styleClip.kind === 'node' ? selectedNodes() : [];
  const edges = styleClip.kind === 'edge' ? selectedEdges() : [];
  if (!nodes.length && !edges.length) { toast(styleClip.kind === 'node' ? 'Select shapes to paste the formatting onto' : 'Select connectors to paste the formatting onto'); return; }
  change(() => applyStyle(styleClip, nodes, edges));
}

const commands = {
  new: () => openTemplates(),
  open: () => $('#file-input').click(),
  save: () => { download(fileBase(app.doc) + '.json', JSON.stringify(app.doc, null, 2), 'application/json'); toast('Saved ' + fileBase(app.doc) + '.json'); },
  'export-menu': (btn) => {
    const m = $('#export-menu');
    const open = m.hidden;
    closeMenus();
    if (open) {
      const small = document.querySelector('[data-cmd="export-html-small"]');
      if (small) small.hidden = !canEmbedFonts(app.doc); // only differs when fonts would be built in
      const r = btn.getBoundingClientRect();
      m.style.left = r.left + 'px'; m.style.top = r.bottom + 4 + 'px';
      m.hidden = false; btn.setAttribute('aria-expanded', 'true');
    }
  },
  'export-html': () => { download(fileBase(app.doc) + '.html', buildStandaloneHTML(app.doc), 'text/html'); toast('Exported interactive HTML'); },
  'export-html-small': () => {
    download(fileBase(app.doc) + '.html', buildStandaloneHTML(app.doc, { embedFonts: false }), 'text/html');
    toast('Exported interactive HTML (system fonts)');
  },
  'export-svg': () => { download(fileBase(app.doc) + '.svg', exportSVGString(app.doc), 'image/svg+xml'); toast('Exported SVG'); },
  'export-png': async () => {
    try { const blob = await exportPNGBlob(app.doc, { scale: 2 }); download(fileBase(app.doc) + '.png', blob); toast('Exported PNG'); }
    catch (err) { toast('PNG export failed: ' + err.message); }
  },
  'export-json': () => commands.save(),
  'copy-svg': async () => {
    try { await navigator.clipboard.writeText(exportSVGString(app.doc)); toast('SVG copied to clipboard'); }
    catch (_) { toast('Clipboard not available in this browser'); }
  },
  'import-text': () => openOutline(),
  'import-sheet': () => openSheetImport(),
  painter: () => { if (app.painter) stopPainter(); else startPainter(false); },
  'copy-style': () => copyStyle(),
  'paste-style': () => pasteStyle(),
  undo: () => { commitInline(); undo(); },
  redo: () => { commitInline(); redo(); },
  'zoom-in': () => zoomAt(1.2),
  'zoom-out': () => zoomAt(1 / 1.2),
  'zoom-reset': () => setZoom(1),
  fit: () => fit(),
  'mode-edit': () => setMode('edit'),
  'mode-preview': () => setMode('preview'),
  theme: () => toggleTheme(),
  help: () => openHelp(),
  'add-lane': () => addLane(),
  'add-phase': () => addPhase(),
  'toggle-lane-orient': () => change((doc) => transposeLanes(doc)),
  'toggle-palette': () => { const a = $('#app'); a.classList.remove('show-inspector'); a.classList.toggle('show-palette'); },
  'toggle-inspector': () => { const a = $('#app'); a.classList.remove('show-palette'); a.classList.toggle('show-inspector'); },
  delete: () => deleteSelection(),
  duplicate: () => { if (copySelection()) pasteClip(clipboard, 24); },
  copy: () => { if (copySelection()) toast('Copied'); },
  cut: () => { if (copySelection()) deleteSelection(); },
  paste: () => { pasteCount++; pasteClip(clipboard, 24 * pasteCount); },
  'select-all': () => select({ nodes: app.doc.nodes.map((n) => n.id), edges: app.doc.edges.map((e) => e.id) }),
  frame: () => frameSelection(),
};

$('#painter-btn').addEventListener('dblclick', () => startPainter(true));

document.addEventListener('click', (ev) => {
  const b = ev.target.closest('[data-cmd]');
  if (b) {
    const cmd = b.dataset.cmd;
    if (cmd !== 'export-menu') closeMenus();
    commands[cmd]?.(b);
    return;
  }
  if (!ev.target.closest('.menu')) closeMenus();
});

// ---------------------------------------------------------------------------
// mode / preview

// ---------------------------------------------------------------------------
// quick-add shape picker (hover an arrow next to the selected shape)

const PICKER_SHAPES = ['process', 'decision', 'terminator', 'document', 'data', 'manualInput', 'database', 'predefined', 'delay', 'offpage', 'connector', 'note'];
let pickerHide = null;
function openQuickPicker({ id, side, x, y }) {
  const m = $('#quick-picker');
  clearTimeout(pickerHide);
  m.replaceChildren(
    h('div', { class: 'qp-title' }, 'Add connected'),
    h('div', { class: 'qp-grid' }, PICKER_SHAPES.map((k) => h('button', {
      type: 'button', role: 'menuitem', title: SHAPES[k].name, 'aria-label': SHAPES[k].name, 'data-pick': k,
      onclick: () => { closeMenus(); quickAdd(id, side, k); },
    }))),
  );
  m.querySelectorAll('[data-pick]').forEach((b) => { b.innerHTML = shapeIcon(b.dataset.pick); });
  m.hidden = false;
  const r = m.getBoundingClientRect();
  const gap = 14;
  let left = side === 'l' ? x - r.width - gap : side === 'r' ? x + gap : x - r.width / 2;
  let top = side === 't' ? y - r.height - gap : side === 'b' ? y + gap : y - r.height / 2;
  m.style.left = Math.max(4, Math.min(innerWidth - r.width - 4, left)) + 'px';
  m.style.top = Math.max(50, Math.min(innerHeight - r.height - 4, top)) + 'px';
}
document.getElementById('quick-picker').addEventListener('mouseleave', () => scheduleHidePicker());
function scheduleHidePicker() {
  clearTimeout(pickerHide);
  pickerHide = setTimeout(() => { const m = $('#quick-picker'); if (!m.matches(':hover')) m.hidden = true; }, 400);
}

// ---------------------------------------------------------------------------
// context menu

function openContextMenu({ x, y, world }) {
  const m = $('#context-menu');
  const nodes = selectedNodes(), edges = selectedEdges();
  const item = (label, fn, key = '') => h('button', { type: 'button', role: 'menuitem', onclick: () => { closeMenus(); fn(); } }, label, key ? h('span', { class: 'muted', style: 'float:right;margin-left:16px' }, key) : null);
  const sep = () => h('hr');
  const items = [];
  if (nodes.length === 1 && !edges.length) items.push(item('Edit text', () => editNodeText(nodes[0].id), 'Enter'));
  if (edges.length === 1 && !nodes.length) items.push(item('Edit label', () => editEdgeLabel(edges[0].id), 'Enter'));
  if (nodes.length || edges.length) {
    items.push(item('Copy', () => commands.copy(), 'Ctrl+C'), item('Cut', () => commands.cut(), 'Ctrl+X'), item('Duplicate', () => commands.duplicate(), 'Ctrl+D'));
    if (nodes.length) {
      items.push(sep(),
        item('Bring to front', () => { const s = new Set(nodes.map((n) => n.id)); change((doc) => { doc.nodes = [...doc.nodes.filter((n) => !s.has(n.id)), ...nodes]; }); }, 'Ctrl+]'),
        item('Send to back', () => { const s = new Set(nodes.map((n) => n.id)); change((doc) => { doc.nodes = [...nodes, ...doc.nodes.filter((n) => !s.has(n.id))]; }); }, 'Ctrl+['),
        item('Frame in group', () => commands.frame(), 'Ctrl+G'));
    }
    if (edges.length) items.push(sep(), item('Reset route', () => change(() => edges.forEach((e) => { e.waypoints = []; }))));
    if ((nodes.length === 1 && !edges.length) || (edges.length === 1 && !nodes.length)) items.push(sep(), item('Format painter', () => startPainter(false)), item('Copy formatting', () => copyStyle(), 'Ctrl+Alt+C'));
    if (styleClip) items.push(item('Paste formatting', () => pasteStyle(), 'Ctrl+Alt+V'));
    items.push(sep(), item('Delete', () => commands.delete(), 'Del'));
  } else {
    items.push(
      item('Add shape here', () => { const n = addShape(app.lastShape || 'process', world); editNodeText(n.id); }),
      item('Add text here', () => { const n = addShape('text', world); editNodeText(n.id); }));
    if (clipboard) items.push(item('Paste', () => commands.paste(), 'Ctrl+V'));
    items.push(sep(), item('Select all', () => commands['select-all'](), 'Ctrl+A'), item('Fit diagram', () => fit(), 'Shift+1'));
  }
  m.replaceChildren(...items);
  m.hidden = false;
  const r = m.getBoundingClientRect();
  m.style.left = Math.min(x, innerWidth - r.width - 8) + 'px';
  m.style.top = Math.min(y, innerHeight - r.height - 8) + 'px';
}

function setMode(mode) {
  commitInline();
  app.mode = mode;
  $('#app').classList.toggle('preview', mode === 'preview');
  document.querySelectorAll('.seg button').forEach((b) => b.classList.toggle('on', b.dataset.cmd === 'mode-' + mode));
  const frame = $('#preview');
  frame.hidden = mode !== 'preview';
  $('#canvas').style.visibility = mode === 'preview' ? 'hidden' : '';
  if (mode === 'preview') refreshPreview();
  ['zoom-in', 'zoom-out', 'zoom-reset', 'fit'].forEach((c) => { const b = document.querySelector(`[data-cmd="${c}"]`); if (b) b.disabled = mode === 'preview'; });
  ['undo', 'redo'].forEach((c) => { const b = document.querySelector(`[data-cmd="${c}"]`); if (b) b.disabled = mode === 'preview' || (c === 'undo' ? !app.undoStack.length : !app.redoStack.length); });
  applyView();
  updateStatus();
}

function refreshPreview() {
  const frame = $('#preview');
  let html = buildStandaloneHTML(app.doc, { preview: true });
  const t = document.documentElement.getAttribute('data-theme');
  if (t && (app.doc.settings.theme || 'auto') === 'auto') html = html.replace('<html lang="en"', `<html lang="en" data-theme="${t}"`);
  frame.srcdoc = html;
}

// ---------------------------------------------------------------------------
// modals

function openModal(content) {
  const dlg = $('#modal');
  $('#modal-body').replaceChildren(...content);
  if (!dlg.open) dlg.showModal();
}
function closeModal() { const d = $('#modal'); if (d.open) d.close(); }

function openTemplates() {
  openModal([
    h('h2', {}, 'New diagram'),
    h('p', { class: 'hint-text' }, 'Your current diagram can be restored with Undo.'),
    h('div', { class: 'tpl-grid' }, TEMPLATES.map((t) => h('button', {
      class: 'tpl', type: 'button', 'data-template': t.id,
      onclick: () => { setDoc(t.make(), { keepHistory: true }); closeModal(); },
    }, h('b', {}, t.name), h('span', {}, t.desc)))),
    h('div', { class: 'modal-actions' }, h('button', { type: 'button', onclick: closeModal }, 'Cancel')),
  ]);
}

// ---------------------------------------------------------------------------
// spreadsheet import

function openSheetImport(file) {
  const st = { rows: null, roles: [], header: true, orientation: 'vertical', sequential: true, badges: true, title: 'Imported process' };
  const mapBox = h('div', { id: 'sheet-map' });
  const summary = h('p', { class: 'hint-text', id: 'sheet-summary' });
  const warn = h('div', { class: 'hint-text', id: 'sheet-warnings', style: 'color:var(--risk);max-height:90px;overflow:auto' });
  const paste = h('textarea', { class: 'code', id: 'sheet-paste', spellcheck: false, placeholder: 'Or paste rows copied from Excel / Google Sheets here (include the header row)…', style: 'min-height:90px' });
  const fileIn = h('input', { type: 'file', accept: '.xlsx,.csv,.tsv,.txt', id: 'sheet-file', hidden: true });
  const create = h('button', { type: 'button', class: 'primary', id: 'sheet-create', disabled: true }, 'Create diagram');

  const load = (rows, title) => {
    if (!rows || rows.length < (st.header ? 2 : 1)) { summary.textContent = 'No rows found.'; st.rows = null; render(); return; }
    st.rows = rows;
    if (title) st.title = title;
    const headers = st.header ? rows[0] : rows[0].map((_, i) => 'Column ' + (i + 1));
    st.roles = guessRoles(headers, st.header ? rows.slice(1) : rows);
    render();
  };
  const build = () => buildFromRows(st.rows, st.roles, { headerRow: st.header, orientation: st.orientation, sequential: st.sequential, badges: st.badges, title: st.title });
  function render() {
    create.disabled = !st.rows;
    if (!st.rows) { mapBox.replaceChildren(); warn.textContent = ''; return; }
    const headers = st.header ? st.rows[0] : st.rows[0].map((_, i) => 'Column ' + (i + 1));
    const sample = (i) => (st.rows.slice(st.header ? 1 : 0).find((r) => String(r[i] ?? '').trim()) || [])[i] || '';
    mapBox.replaceChildren(
      h('div', { class: 'sec-title', style: 'margin-top:10px;font-size:11px;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);font-weight:600' }, 'Columns'),
      h('table', { class: 'map-table' },
        h('tr', {}, h('th', {}, 'Column'), h('th', {}, 'Example'), h('th', {}, 'Use as')),
        headers.map((hd, i) => {
          const sel = h('select', { 'data-col': i, 'aria-label': 'Use column ' + hd + ' as' }, ROLES.map(([v, l]) => h('option', { value: v, selected: v === st.roles[i] }, l)));
          sel.addEventListener('change', () => { st.roles[i] = sel.value; render(); });
          return h('tr', {}, h('td', {}, String(hd)), h('td', { class: 'muted' }, String(sample(i)).slice(0, 48)), h('td', {}, sel));
        })),
    );
    try {
      const r = build();
      summary.textContent = `${r.stats.steps} steps · ${r.stats.connectors} connectors · ${r.stats.lanes} lanes · ${r.stats.phases} phases`;
      warn.replaceChildren(...r.warnings.slice(0, 8).map((w) => h('div', {}, '⚠ ' + w)), ...(r.warnings.length > 8 ? [h('div', {}, `…and ${r.warnings.length - 8} more`)] : []));
    } catch (e) { summary.textContent = 'Could not build: ' + e.message; }
  }
  const readFile = async (f) => {
    try {
      const rows = /\.xlsx$/i.test(f.name) ? await readXlsx(await f.arrayBuffer()) : parseDelimited(await f.text());
      paste.value = '';
      load(rows, f.name.replace(/\.[^.]+$/, ''));
      summary.prepend(h('b', {}, f.name + ': '));
    } catch (e) { summary.textContent = 'Could not read ' + f.name + ': ' + e.message; }
  };
  fileIn.addEventListener('change', () => { if (fileIn.files[0]) readFile(fileIn.files[0]); });
  paste.addEventListener('input', () => load(parseDelimited(paste.value)));
  const opt = (label, el) => h('label', { class: 'chk', style: 'margin-right:14px;display:inline-flex' }, el, label);
  const hdr = h('input', { type: 'checkbox', checked: true }); hdr.addEventListener('change', () => { st.header = hdr.checked; if (st.rows) load(st.rows); });
  const seq = h('input', { type: 'checkbox', checked: true }); seq.addEventListener('change', () => { st.sequential = seq.checked; render(); });
  const bdg = h('input', { type: 'checkbox', checked: true }); bdg.addEventListener('change', () => { st.badges = bdg.checked; render(); });
  const ori = h('select', { 'aria-label': 'Lanes as' }, h('option', { value: 'vertical' }, 'Lanes as columns'), h('option', { value: 'horizontal' }, 'Lanes as rows'));
  ori.addEventListener('change', () => { st.orientation = ori.value; render(); });
  create.addEventListener('click', () => {
    const r = build();
    setDoc(r.doc, { keepHistory: true });
    closeModal();
    toast(`Imported ${r.stats.steps} steps` + (r.warnings.length ? ` — ${r.warnings.length} warning(s)` : ''), 3500);
  });
  openModal([
    h('h2', {}, 'Import from a spreadsheet'),
    h('p', { class: 'hint-text' }, 'One row per step. Columns such as ', h('b', {}, 'Step, Activity, Owner, Phase, Type, Next'),
      ' are recognised automatically; other columns become detail fields. In ', h('b', {}, 'Next'), ' list the next step IDs separated by ',
      h('code', {}, ';'), ', with optional labels (', h('code', {}, 'Yes: CD-03; No: CD-01'), '). Prefix ', h('code', {}, '~'), ' for a dashed information feed.'),
    h('div', { class: 'btns', style: 'margin-bottom:8px' },
      h('button', { type: 'button', onclick: () => fileIn.click(), id: 'sheet-choose' }, 'Choose .xlsx / .csv file…'),
      h('button', { type: 'button', id: 'sheet-example', onclick: () => { paste.value = EXAMPLE_CSV; load(parseDelimited(EXAMPLE_CSV), 'Cash disbursements walkthrough'); } }, 'Use example'),
      h('button', { type: 'button', onclick: () => download('flowchart-import-example.csv', EXAMPLE_CSV, 'text/csv') }, 'Download example CSV')),
    fileIn, paste,
    h('div', { style: 'margin-top:6px' }, opt('First row is headers', hdr), opt('Connect rows in order if there is no Next column', seq), opt('R / C badges on rows with a risk or control', bdg), ori),
    mapBox, summary, warn,
    h('div', { class: 'modal-actions' }, h('button', { type: 'button', onclick: closeModal }, 'Cancel'), create),
  ]);
  if (file) readFile(file);
}

function openOutline() {
  const ta = h('textarea', { class: 'code', spellcheck: false, id: 'outline-text' }, OUTLINE_EXAMPLE);
  const err = h('p', { class: 'hint-text', style: 'color:var(--risk)' });
  openModal([
    h('h2', {}, 'Import from text outline'),
    h('p', { class: 'hint-text' },
      'One shape per line as ', h('code', {}, 'id: Text {shape} #TAG !R !C1'), ', indented ', h('code', {}, 'field: value'),
      ' lines add details, ', h('code', {}, '[Lane]'), ' switches lane, ', h('code', {}, 'a -> b : label'), ' connects (', h('code', {}, '-.->'), ' for the second connector type). Shapes are laid out automatically.'),
    ta, err,
    h('div', { class: 'modal-actions' },
      h('button', { type: 'button', onclick: closeModal }, 'Cancel'),
      h('button', {
        type: 'button', class: 'primary', id: 'outline-create',
        onclick: () => {
          try {
            const { doc, errors } = parseOutline(ta.value);
            if (!doc.nodes.length) { err.textContent = errors.join(' · ') || 'Nothing to import.'; return; }
            setDoc(doc, { keepHistory: true });
            closeModal();
            if (errors.length) toast(errors.length + ' line(s) skipped: ' + errors[0], 4000);
          } catch (e) { err.textContent = e.message; }
        },
      }, 'Create diagram')),
  ]);
  ta.focus();
}

function openHelp() {
  const K = [
    ['Double-click canvas', 'New shape (last used type)'], ['Double-click shape / line', 'Edit text or label'],
    ['Drag from blue dot', 'Connect (drop on empty space = new shape; hold Shift for a loose end)'],
    ['Click ▸ arrows', 'Add a connected shape in that direction'], ['Drag segment handle', 'Reroute an elbow connector'],
    ['Type while a shape is selected', 'Replace its text'], ['Enter / F2', 'Edit selected text'],
    ['Delete / Backspace', 'Delete selection'], ['Ctrl+C / X / V / D', 'Copy, cut, paste, duplicate'],
    ['Ctrl+Z / Ctrl+Shift+Z', 'Undo / redo'], ['Ctrl+A', 'Select all'], ['Ctrl+G', 'Frame selection in a group'],
    ['Ctrl+B / Ctrl+I', 'Bold / italic'], ['Ctrl+Alt+C / Ctrl+Alt+V', 'Copy / paste formatting (brush button = format painter)'], ['Ctrl+] / Ctrl+[', 'Bring to front / send to back'],
    ['Arrows (Shift = ×4)', 'Nudge by one grid step (Alt = 1px)'], ['Alt while dragging', 'Disable snapping'],
    ['Shift while resizing', 'Keep proportions'], ['Wheel / Shift+wheel', 'Pan'], ['Ctrl+wheel, + / −', 'Zoom'],
    ['Space+drag, middle drag', 'Pan'], ['Shift+1 / 0', 'Fit / 100%'], ['Ctrl+S / Ctrl+O', 'Save / open .json'],
    ['Ctrl+Shift+P', 'Toggle preview'], ['Esc', 'Clear selection'],
  ];
  openModal([
    h('h2', {}, 'Shortcuts & tips'),
    h('div', { class: 'keys' }, K.flatMap(([k, d]) => [h('div', {}, ...k.split(/( \/ |, | = )/).map((p) => (/^( \/ |, | = )$/.test(p) ? p : h('kbd', {}, p)))), h('div', {}, d)])),
    h('p', { class: 'hint-text' }, 'Your work is saved in this browser automatically. Use Save to keep a .json file, and Export → Interactive HTML to share a click-through version like the preview.'),
    h('div', { class: 'modal-actions' }, h('button', { type: 'button', class: 'primary', onclick: closeModal }, 'Close')),
  ]);
}

// ---------------------------------------------------------------------------
// file open

async function openFile(file) {
  if (/\.(xlsx|csv|tsv|txt)$/i.test(file.name)) { openSheetImport(file); return; }
  try {
    const text = await file.text();
    const doc = parseDocumentText(text);
    setDoc(doc, { keepHistory: true });
    toast('Opened ' + file.name);
  } catch (e) {
    toast('Could not open file: ' + e.message, 4000);
  }
}
$('#file-input').addEventListener('change', (ev) => {
  const f = ev.target.files[0];
  if (f) openFile(f);
  ev.target.value = '';
});

// ---------------------------------------------------------------------------
// keyboard

function nudge(dx, dy) {
  const nodes = selectedNodes();
  if (!nodes.length) return;
  change(() => { nodes.forEach((n) => { n.x += dx; n.y += dy; }); }, { inspector: true });
}

window.addEventListener('keydown', (ev) => {
  const dlg = $('#modal');
  if (dlg.open) return;
  const mod = ev.ctrlKey || ev.metaKey;
  const key = ev.key;
  if (mod && ev.shiftKey && key.toLowerCase() === 'p') { ev.preventDefault(); setMode(app.mode === 'edit' ? 'preview' : 'edit'); return; }
  if (mod && ev.altKey && (ev.code === 'KeyC' || ev.code === 'KeyV') && !isTyping(ev.target)) { ev.preventDefault(); ev.code === 'KeyC' ? copyStyle() : pasteStyle(); return; }
  if (mod && key.toLowerCase() === 's') { ev.preventDefault(); commands.save(); return; }
  if (mod && key.toLowerCase() === 'o') { ev.preventDefault(); commands.open(); return; }
  if (isTyping(ev.target) || isInlineEditing()) {
    if (key === 'Escape' && ev.target.closest?.('#inspector')) ev.target.blur();
    return;
  }
  if (app.mode !== 'edit') { if (key === 'Escape') setMode('edit'); return; }
  if (isDragging()) {
    // no commands mid-gesture; Escape cancels the drag
    if (key === 'Escape') cancelDrag();
    if (key !== 'Shift' && key !== 'Alt' && key !== 'Control' && key !== 'Meta') ev.preventDefault();
    return;
  }
  const k = key.toLowerCase();
  if (mod && k === 'z') { ev.preventDefault(); ev.shiftKey ? commands.redo() : commands.undo(); return; }
  if (mod && k === 'y') { ev.preventDefault(); commands.redo(); return; }
  if (mod && k === 'c') { ev.preventDefault(); commands.copy(); return; }
  if (mod && k === 'x') { ev.preventDefault(); commands.cut(); return; }
  if (mod && k === 'v') { return; /* handled by paste event */ }
  if (mod && k === 'd') { ev.preventDefault(); commands.duplicate(); return; }
  if (mod && k === 'a') { ev.preventDefault(); commands['select-all'](); return; }
  if (mod && k === 'g') { ev.preventDefault(); commands.frame(); return; }
  if (mod && (k === 'b' || k === 'i')) {
    ev.preventDefault();
    const nodes = selectedNodes();
    const prop = k === 'b' ? 'bold' : 'italic';
    const on = !nodes.every((n) => n.style?.[prop]);
    change(() => nodes.forEach((n) => { n.style = { ...n.style, [prop]: on }; }));
    return;
  }
  if (mod && (key === ']' || key === '[')) {
    ev.preventDefault();
    const nodes = selectedNodes();
    const s = new Set(nodes.map((n) => n.id));
    change((doc) => { const rest = doc.nodes.filter((n) => !s.has(n.id)); doc.nodes = key === ']' ? [...rest, ...nodes] : [...nodes, ...rest]; });
    return;
  }
  if (key === 'Delete' || key === 'Backspace') { ev.preventDefault(); commands.delete(); return; }
  if (key === 'Escape') { if (app.painter) stopPainter(); else if ([...document.querySelectorAll('.menu')].some((m) => !m.hidden)) closeMenus(); else clearSelection(); return; }
  if (key === '?' ) { openHelp(); return; }
  if (key === '+' || key === '=') { zoomAt(1.2); return; }
  if (key === '-' || key === '_') { zoomAt(1 / 1.2); return; }
  if (key === '0' && !mod) { setZoom(1); return; }
  if (key === '!' || (ev.shiftKey && ev.code === 'Digit1')) { fit(); return; }
  if (key.startsWith('Arrow')) {
    ev.preventDefault();
    const g = ev.altKey ? 1 : (app.doc.settings.snap ? app.doc.settings.grid : 1) * (ev.shiftKey ? 4 : 1);
    const d = { ArrowLeft: [-g, 0], ArrowRight: [g, 0], ArrowUp: [0, -g], ArrowDown: [0, g] }[key];
    nudge(...d);
    return;
  }
  const nodes = selectedNodes();
  const edges = selectedEdges();
  if (key === 'Enter' || key === 'F2') {
    ev.preventDefault();
    if (nodes.length === 1) editNodeText(nodes[0].id);
    else if (edges.length === 1 && !nodes.length) editEdgeLabel(edges[0].id);
    return;
  }
  if (key === 'Tab' && nodes.length === 1 && !mod) {
    // Tab adds a connected shape to the right (Shift+Tab below)
    ev.preventDefault();
    quickAdd(nodes[0].id, ev.shiftKey ? 'b' : 'r');
    return;
  }
  if (key.length === 1 && key !== ' ' && !mod && !ev.altKey && (nodes.length === 1 && !edges.length || edges.length === 1 && !nodes.length)) {
    ev.preventDefault();
    if (nodes.length) editNodeText(nodes[0].id); else editEdgeLabel(edges[0].id);
    const ta = $('#inline-editor');
    ta.value = key;
    ta.setSelectionRange(1, 1);
    ta.dispatchEvent(new Event('input'));
  }
});

window.addEventListener('paste', (ev) => {
  if (isTyping(ev.target) || isInlineEditing() || app.mode !== 'edit') return;
  const t = ev.clipboardData?.getData('text/plain') || '';
  ev.preventDefault();
  if (t.startsWith('{')) {
    try {
      const clip = JSON.parse(t);
      if (clip.fcClip) {
        const same = clipboard && JSON.stringify(clip) === JSON.stringify(clipboard);
        if (!same) { clipboard = clip; pasteCount = 0; }
        commands.paste();
        return;
      }
      if (clip.nodes && clip.version) { setDoc(parseDocumentText(t), { keepHistory: true }); return; }
    } catch (_) { /* fall through */ }
  }
  if (clipboard && !t) { commands.paste(); return; }
  if (t.trim()) {
    // paste plain text: one shape per line, stacked
    const lines = t.split(/\r?\n/).map((s) => s.trim()).filter(Boolean).slice(0, 50);
    const first = addShape(app.lastShape || 'process');
    change(() => { first.text = lines[0]; });
    let prev = first;
    for (const line of lines.slice(1)) {
      const nn = quickAdd(prev.id, 'b');
      commitInline(false);
      change(() => { nn.text = line; });
      prev = nn;
    }
  } else if (clipboard) commands.paste();
});

// ---------------------------------------------------------------------------
// search

let searchIdx = 0;
$('#search').addEventListener('keydown', (ev) => {
  if (ev.key === 'Escape') { ev.target.value = ''; ev.target.blur(); return; }
  if (ev.key !== 'Enter') return;
  const q = ev.target.value.trim().toLowerCase();
  if (!q) return;
  const hits = app.doc.nodes.filter((n) => (n.text + ' ' + n.tag + ' ' + Object.values(n.fields || {}).join(' ')).toLowerCase().includes(q));
  if (!hits.length) { toast('No match for "' + q + '"'); return; }
  const n = hits[searchIdx++ % hits.length];
  if (app.mode !== 'edit') setMode('edit');
  select({ nodes: [n.id] });
  centerOn(n);
  toast(`${(searchIdx - 1) % hits.length + 1} of ${hits.length}`);
});
$('#search').addEventListener('input', () => { searchIdx = 0; });

// ---------------------------------------------------------------------------
// title

const titleInput = $('#doc-title');
let titlePending = true;
titleInput.addEventListener('focus', () => { titlePending = true; });
titleInput.addEventListener('input', () => {
  if (titlePending) { titlePending = false; app.undoStack.push(JSON.stringify(app.doc)); app.redoStack.length = 0; }
  app.doc.title = titleInput.value;
  emit({ doc: true, title: true });
});
titleInput.addEventListener('keydown', (ev) => { if (ev.key === 'Enter' || ev.key === 'Escape') titleInput.blur(); });

// ---------------------------------------------------------------------------
// status

function updateStatus() {
  const n = app.sel.nodes.size, e = app.sel.edges.size;
  const parts = [];
  if (n) parts.push(`${n} shape${n > 1 ? 's' : ''}`);
  if (e) parts.push(`${e} connector${e > 1 ? 's' : ''}`);
  if (app.sel.lane) parts.push('lane');
  if (app.sel.phase) parts.push('phase');
  $('#status-sel').textContent = parts.length ? parts.join(', ') + ' selected' : `${app.doc.nodes.length} shapes · ${app.doc.edges.length} connectors`;
  const pb = document.getElementById('painter-btn');
  if (pb) pb.disabled = app.mode === 'preview' || (!app.painter && !((n === 1 && !e) || (e === 1 && !n)));
  $('#status-hint').textContent = app.painter
    ? `Format painter: click ${app.painter.kind === 'node' ? 'shapes' : 'connectors'} to apply the formatting${app.painter.sticky ? ' · Esc to stop' : ''}`
    : app.mode === 'preview'
    ? 'Preview: click shapes to see details — this is what the exported HTML looks like.'
    : n === 1 ? 'Double-click or type to edit text · drag blue dots to connect · arrows add linked shapes'
      : e === 1 ? 'Drag handles to reroute · double-click to edit the label'
        : 'Double-click to add a shape · drag to select · Space+drag or wheel to pan · Ctrl+wheel to zoom';
  const u = document.querySelector('[data-cmd="undo"]'), r = document.querySelector('[data-cmd="redo"]');
  if (u) u.disabled = app.mode === 'preview' || !app.undoStack.length;
  if (r) r.disabled = app.mode === 'preview' || !app.redoStack.length;
}

// ---------------------------------------------------------------------------
// wiring

let rafPending = false;
on((what) => {
  if (what.openFile) { openFile(what.openFile); return; }
  if (what.contextMenu) { openContextMenu(what.contextMenu); return; }
  if (what.paint) { onPaint(what.paint); return; }
  if (what.quickPicker) { openQuickPicker(what.quickPicker); return; }
  if (what.quickPickerLeave) { scheduleHidePicker(); return; }
  if (what.toast) { toast(what.toast, 3500); return; }
  if (what.command) { commands[what.command]?.(); return; }
  if (what.saved) { $('#status-save').textContent = 'Saved in this browser ' + app.lastSaved.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }); return; }
  if (what.doc) {
    applyFont();
    if (!what.title && document.activeElement !== titleInput) titleInput.value = app.doc.title;
    document.title = (app.doc.title || 'Untitled') + ' — Flowchart Editor';
    scheduleSave();
    buildConnectorDefaults();
    if (app.mode === 'preview') refreshPreview();
  }
  if (what.doc || what.sel) renderCanvas();
  if (what.inspector) renderInspector();
  if (what.fit) requestAnimationFrame(() => fit());
  if (!rafPending) { rafPending = true; requestAnimationFrame(() => { rafPending = false; updateStatus(); }); }
});

function start() {
  applyEditorTheme();
  initCanvas();
  initInspector();
  buildPalette();
  const params = new URLSearchParams(location.search);
  const tpl = params.get('template');
  let doc = null;
  if (tpl) doc = TEMPLATES.find((t) => t.id === tpl)?.make() || null;
  if (!doc) doc = loadSaved();
  if (!doc) doc = TEMPLATES.find((t) => t.id === 'basic').make();
  setDoc(doc);
  titleInput.value = app.doc.title;
  requestAnimationFrame(() => fit());
}

// test / automation hooks
window.fc = {
  app, commands, setDoc, select, clearSelection, fit, renderCanvas, renderOverlay, setMode, parseOutline, autoLayout, numberTags,
  getDoc: () => app.doc,
  exportHTML: () => buildStandaloneHTML(app.doc),
  exportSVG: () => exportSVGString(app.doc),
  exportPNG: async () => { const b = await exportPNGBlob(app.doc); return b ? b.size : 0; },
  templates: TEMPLATES.map((t) => t.id),
  loadTemplate: (id) => setDoc(TEMPLATES.find((t) => t.id === id).make()),
  nodeScreenBox: (id) => {
    const n = nodeById(id);
    const r = document.getElementById('canvas').getBoundingClientRect();
    const { x, y, k } = app.view;
    return { x: r.left + x + n.x * k, y: r.top + y + n.y * k, w: n.w * k, h: n.h * k, cx: r.left + x + (n.x + n.w / 2) * k, cy: r.top + y + (n.y + n.h / 2) * k };
  },
  worldToScreen: (px, py) => {
    const r = document.getElementById('canvas').getBoundingClientRect();
    const { x, y, k } = app.view;
    return [r.left + x + px * k, r.top + y + py * k];
  },
  renderSVG: (o) => renderSVG(app.doc, o),
  FONTS, center,
};

start();
