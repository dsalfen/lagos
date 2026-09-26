// Canvas: renders the diagram, draws selection overlays and handles all pointer input.
import {
  app, emit, change, commitFrom, snapshot, select, toggleSelect, clearSelection,
  nodeById, edgeById, selectedNodes,
} from './state.js';
import { renderDiagram, computeRoutes, renderNode, renderEdge, renderLabel, nodeStyle, labelGeometry, fontFor, esc, fitNodeText, jumpSize } from './render.js';
import {
  portPoint, projectOnPolyline, simplify, center, unionBox, pathD,
} from './geometry.js';
import { makeNode, makeEdge, phaseAxis, phaseStarts, growPhasesFor } from './model.js';
import { shapeDef, shapeTextBox } from './shapes.js';

const $ = (s) => document.querySelector(s);
let svg, world, diagramG, overlayG, stage, inline;
let last = null;
let hoverNode = null;
let drag = null;
let guides = [];
let marquee = null;
let connectPreview = null;
let dropTargetId = null;
let spaceDown = false;
let inlineState = null;

const OPP = { t: 'b', b: 't', l: 'r', r: 'l' };

export function initCanvas() {
  svg = $('#canvas'); world = $('#world'); diagramG = $('#diagram'); overlayG = $('#overlay');
  stage = $('#stage'); inline = $('#inline-editor');
  svg.addEventListener('pointerdown', onPointerDown);
  svg.addEventListener('pointermove', onPointerMove);
  svg.addEventListener('pointerup', onPointerUp);
  svg.addEventListener('pointercancel', onPointerUp);
  svg.addEventListener('contextmenu', (ev) => { ev.preventDefault(); if (app.mode === 'edit') emit({ contextMenu: { x: ev.clientX, y: ev.clientY, world: toWorld(ev) } }); });
  stage.addEventListener('wheel', onWheel, { passive: false });
  stage.addEventListener('dragover', (ev) => { if (app.mode === 'edit') { ev.preventDefault(); ev.dataTransfer.dropEffect = 'copy'; } });
  stage.addEventListener('drop', onDrop);
  svg.addEventListener('pointerleave', () => { if (!drag && hoverNode) { hoverNode = null; renderOverlay(); } });
  window.addEventListener('keydown', (ev) => {
    if (ev.code === 'Space' && !isTyping(ev.target) && app.mode === 'edit') { spaceDown = true; stage.classList.add('space-down'); if (ev.target === stage || ev.target === document.body) ev.preventDefault(); }
  });
  window.addEventListener('keyup', (ev) => { if (ev.code === 'Space') { spaceDown = false; stage.classList.remove('space-down'); } });
  inline.addEventListener('keydown', onInlineKey);
  inline.addEventListener('blur', () => commitInline());
  inline.addEventListener('input', sizeInline);
  new ResizeObserver(() => applyView()).observe(stage);
}

export const isDragging = () => !!drag;

export function cancelDrag() {
  if (!drag) return;
  const before = drag.before;
  drag = null;
  pendingMove = null;
  guides = []; marquee = null; connectPreview = null; dropTargetId = null;
  stage.classList.remove('panning');
  if (before !== snapshot()) { app.doc = JSON.parse(before); emit({ doc: true, inspector: true }); }
  else renderCanvas();
}

export const isTyping = (t) => t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable);

// ---------------------------------------------------------------------------
// rendering

const SVGNS = 'http://www.w3.org/2000/svg';
function replaceMarkup(el, markup) {
  const tmp = document.createElementNS(SVGNS, 'g');
  tmp.innerHTML = markup;
  const fresh = tmp.firstElementChild;
  if (el && fresh) el.replaceWith(fresh);
  return fresh;
}

// While dragging, patch only what changed instead of rebuilding the whole SVG.
function patchCanvas() {
  const doc = app.doc;
  const opts = { editor: true, selNodes: app.sel.nodes, selEdges: app.sel.edges };
  const routes = computeRoutes(doc, { prevRoutes: last.routes, fastNodes: drag.fastNodes, fastEdges: drag.fastEdges });
  const q = (sel) => diagramG.querySelector(sel);
  for (const id of drag.fastNodes || []) {
    const n = nodeById(id);
    const el = q(`[data-node="${CSS.escape(id)}"]`);
    if (!n || !el) continue;
    if (drag.type === 'resize') replaceMarkup(el, renderNode(doc, n, opts));
    else el.setAttribute('transform', `translate(${Math.round(n.x * 100) / 100} ${Math.round(n.y * 100) / 100})`);
  }
  const hopKey = (m, id) => JSON.stringify(m?.get(id) || []);
  for (const e of doc.edges) {
    const r = routes.get(e.id);
    const hopsChanged = hopKey(routes.jumps, e.id) !== hopKey(last.routes.jumps, e.id);
    if (r === last.routes.get(e.id) && !hopsChanged) continue;
    replaceMarkup(q(`[data-edge="${CSS.escape(e.id)}"]`), renderEdge(doc, e, r, opts, routes.jumps?.get(e.id)));
    const lab = q(`[data-label="${CSS.escape(e.id)}"]`);
    const markup = renderLabel(doc, e, r, opts);
    if (lab) { if (markup) replaceMarkup(lab, markup); else lab.remove(); }
    else if (markup) { const g = q('.edge-labels'); const tmp = document.createElementNS(SVGNS, 'g'); tmp.innerHTML = markup; g?.append(tmp.firstElementChild); }
  }
  last = { ...last, routes };
  renderOverlay();
}

export function renderCanvas() {
  if (drag && drag.moved && (drag.fastNodes || drag.fastEdges) && last && !dropTargetId && !drag.fullRender) { patchCanvas(); return; }
  const r = renderDiagram(app.doc, { editor: true, selNodes: app.sel.nodes, selEdges: app.sel.edges, extentPad: 240 });
  diagramG.innerHTML = r.body;
  last = r;
  if (app.sel.lane) diagramG.querySelector(`[data-lane="${CSS.escape(app.sel.lane)}"]`)?.classList.add('sel');
  if (app.sel.phase) diagramG.querySelector(`[data-phase="${CSS.escape(app.sel.phase)}"]`)?.classList.add('sel');
  if (dropTargetId) diagramG.querySelector(`[data-node="${CSS.escape(dropTargetId)}"]`)?.classList.add('drop-target');
  renderOverlay();
  applyView();
  $('#empty-hint').hidden = app.doc.nodes.length > 0 || app.doc.lanes.items.length > 0;
}

export const lastRender = () => last;

export function applyView() {
  const { x, y, k } = app.view;
  world.setAttribute('transform', `translate(${x} ${y}) scale(${k})`);
  const g = (app.doc?.settings.grid || 8) * k;
  const showGrid = app.doc?.settings.showGrid && app.mode === 'edit';
  stage.classList.toggle('grid', !!showGrid && g >= 4);
  let gs = g;
  while (gs < 8) gs *= 2;
  stage.style.backgroundSize = `${gs}px ${gs}px`;
  stage.style.backgroundPosition = `${x - gs / 2}px ${y - gs / 2}px`;
  const zl = $('#zoom-label');
  if (zl) zl.textContent = Math.round(k * 100) + '%';
  if (inlineState) positionInline();
}

function circle(x, y, r, a = {}) {
  return `<circle cx="${x}" cy="${y}" r="${r}"${Object.entries(a).map(([k, v]) => ` ${k}="${esc(v)}"`).join('')}/>`;
}
function rect(x, y, w, h, a = {}) {
  return `<rect x="${x}" y="${y}" width="${w}" height="${h}"${Object.entries(a).map(([k, v]) => ` ${k}="${esc(v)}"`).join('')}/>`;
}

export function renderOverlay() {
  if (!last || !overlayG) return;
  const k = app.view.k, s = 1 / k;
  const doc = app.doc;
  let o = '';

  // lane dividers
  if (app.mode === 'edit' && doc.lanes.items.length) {
    const horiz = doc.lanes.orientation === 'horizontal';
    let pos = 0;
    doc.lanes.items.forEach((l, i) => {
      pos += l.size;
      const cls = 'ov-divider ' + (horiz ? 'h' : 'v') + (drag?.divider === i ? ' active' : '');
      o += horiz
        ? rect(0, pos - 3 * s, last.extent, 6 * s, { class: cls, 'data-divider': i })
        : rect(pos - 3 * s, 0, 6 * s, last.extent, { class: cls, 'data-divider': i });
    });
  }

  if (app.mode === 'edit' && doc.phases.items.length) {
    const { along } = phaseAxis(doc);
    const H = doc.phases.header || 34;
    phaseStarts(doc).forEach(([ph, pos], i) => {
      const end = pos + ph.size;
      const cls = 'ov-divider ' + (along === 'y' ? 'h' : 'v') + (drag?.pdivider === i ? ' active' : '');
      o += along === 'y'
        ? rect(-H, end - 3 * s, H + last.crossLen, 6 * s, { class: cls, 'data-pdivider': i })
        : rect(end - 3 * s, -H, 6 * s, H + last.crossLen, { class: cls, 'data-pdivider': i });
    });
  }

  const sel = selectedNodes();
  // selected edges: halo + handles
  const selEdges = doc.edges.filter((e) => app.sel.edges.has(e.id));
  for (const e of selEdges) {
    const r = last.routes.get(e.id);
    if (!r) continue;
    o += `<path class="ov-halo" d="${pathD(r, doc.settings.cornerRadius, last.routes.jumps?.get(e.id), jumpSize(doc))}" stroke-width="${7 * s}"/>`;
  }
  if (selEdges.length === 1 && !sel.length && !(drag && drag.hideHandles)) {
    const e = selEdges[0];
    const r = last.routes.get(e.id);
    if (r) {
      const pts = r.pts;
      if (e.routing === 'orthogonal') {
        for (let i = 0; i < pts.length - 1; i++) {
          const a = pts[i], b = pts[i + 1];
          const L = Math.abs(b[0] - a[0]) + Math.abs(b[1] - a[1]);
          if (L < 16 * s) continue;
          const vert = Math.abs(b[0] - a[0]) < 0.5;
          let f = 0.5;
          const lg = e.label && labelGeometry(doc, e, r);
          if (lg && Math.hypot(lg.cx - (a[0] + b[0]) / 2, lg.cy - (a[1] + b[1]) / 2) < 30 * s + Math.max(lg.w, lg.h) / 2) f = 0.2;
          const mx = a[0] + (b[0] - a[0]) * f, my = a[1] + (b[1] - a[1]) * f;
          const w = (vert ? 6 : 14) * s, h = (vert ? 14 : 6) * s;
          o += rect(mx - w / 2, my - h / 2, w, h, { class: 'ov-seg', 'data-seg': i, 'data-id': e.id, rx: 2 * s, style: `cursor:${vert ? 'ew-resize' : 'ns-resize'}` });
        }
      } else {
        const ctrl = [pts[0], ...e.waypoints, pts[pts.length - 1]];
        for (let i = 0; i < ctrl.length - 1; i++) {
          const a = ctrl[i], b = ctrl[i + 1];
          o += circle((a[0] + b[0]) / 2, (a[1] + b[1]) / 2, 4 * s, { class: 'ov-addwp', 'data-addwp': i, 'data-id': e.id });
        }
        e.waypoints.forEach((p, i) => { o += circle(p[0], p[1], 5 * s, { class: 'ov-wp', 'data-wp': i, 'data-id': e.id }); });
      }
      const a = pts[0], b = pts[pts.length - 1];
      o += circle(a[0], a[1], 5.5 * s, { class: 'ov-ep' + (e.from.node ? '' : ' free'), 'data-ep': 'from', 'data-id': e.id });
      o += circle(b[0], b[1], 5.5 * s, { class: 'ov-ep' + (e.to.node ? '' : ' free'), 'data-ep': 'to', 'data-id': e.id });
    }
  }

  // selected nodes
  if (sel.length) {
    const hide = drag && drag.hideHandles;
    for (const n of sel) o += rect(n.x - 3 * s, n.y - 3 * s, n.w + 6 * s, n.h + 6 * s, { class: 'ov-box', 'stroke-width': s });
    if (sel.length > 1) {
      const u = unionBox(sel);
      o += rect(u.x - 8 * s, u.y - 8 * s, u.w + 16 * s, u.h + 16 * s, { class: 'ov-box', 'stroke-width': s, opacity: 0.6 });
    }
    if (sel.length === 1 && !hide) {
      const n = sel[0];
      const hs = 8 * s;
      const H = { nw: [0, 0], n: [0.5, 0], ne: [1, 0], e: [1, 0.5], se: [1, 1], s: [0.5, 1], sw: [0, 1], w: [0, 0.5] };
      const cursors = { nw: 'nwse-resize', se: 'nwse-resize', ne: 'nesw-resize', sw: 'nesw-resize', n: 'ns-resize', s: 'ns-resize', e: 'ew-resize', w: 'ew-resize' };
      for (const [h, [fx, fy]] of Object.entries(H)) {
        o += rect(n.x + n.w * fx - hs / 2, n.y + n.h * fy - hs / 2, hs, hs, { class: 'ov-handle', 'data-handle': h, 'data-id': n.id, 'stroke-width': s, rx: 1.5 * s, style: `cursor:${cursors[h]}` });
      }
      if (!shapeDef(n.shape).container) {
        // quick-add arrows
        const d = 22 * s, a = 6 * s;
        const [cx, cy] = center(n);
        const Q = {
          t: [cx, n.y - d, `M${cx - a} ${n.y - d + a / 2}L${cx} ${n.y - d - a / 1.2}L${cx + a} ${n.y - d + a / 2}Z`],
          b: [cx, n.y + n.h + d, `M${cx - a} ${n.y + n.h + d - a / 2}L${cx} ${n.y + n.h + d + a / 1.2}L${cx + a} ${n.y + n.h + d - a / 2}Z`],
          l: [n.x - d, cy, `M${n.x - d + a / 2} ${cy - a}L${n.x - d - a / 1.2} ${cy}L${n.x - d + a / 2} ${cy + a}Z`],
          r: [n.x + n.w + d, cy, `M${n.x + n.w + d - a / 2} ${cy - a}L${n.x + n.w + d + a / 1.2} ${cy}L${n.x + n.w + d - a / 2} ${cy + a}Z`],
        };
        for (const [side, [, , dd]] of Object.entries(Q)) {
          o += `<path class="ov-quick" d="${dd}" data-quick="${side}" data-id="${esc(n.id)}"><title>Click: add a connected shape · Hover: choose its type · Drag: draw a connector</title></path>`;
        }
      }
    }
  }

  // ports on hovered node (and on the node under a connection drag)
  const portNodes = new Set();
  if (!drag && hoverNode && !(app.sel.nodes.has(hoverNode) && app.sel.nodes.size === 1)) portNodes.add(hoverNode);
  if (drag && (drag.type === 'connect' || drag.type === 'reconnect') && dropTargetId) portNodes.add(dropTargetId);
  for (const id of portNodes) {
    const n = nodeById(id);
    if (!n || (shapeDef(n.shape).container && !(drag && dropTargetId === id))) continue;
    // keep ports from covering small (or zoomed-out) shapes, so they can still be dragged
    const screenMin = Math.min(n.w, n.h) * k;
    const isTarget = drag && (drag.type === 'connect' || drag.type === 'reconnect');
    if (screenMin < 22 && !isTarget) continue;
    const hit = Math.max(4, Math.min(10, screenMin * 0.2));
    for (const side of ['t', 'r', 'b', 'l']) {
      const [px, py] = portPoint(n, side, 0);
      o += circle(px, py, hit * s, { class: 'ov-port-hit', 'data-port': side, 'data-id': n.id });
      o += circle(px, py, Math.min(4.5, hit) * s, { class: 'ov-port', 'data-port': side, 'data-id': n.id, 'stroke-width': 1.2 * s });
    }
  }

  if (connectPreview) {
    o += `<path class="ov-preview" d="${connectPreview.d}" stroke-width="${1.6 * s}"/>`;
    if (connectPreview.snap) o += circle(connectPreview.snap[0], connectPreview.snap[1], 5 * s, { class: 'ov-snap' });
  }
  for (const g of guides) o += `<line class="ov-guide" x1="${g[0]}" y1="${g[1]}" x2="${g[2]}" y2="${g[3]}" stroke-width="${s}"/>`;
  if (marquee) {
    const x = Math.min(marquee[0], marquee[2]), y = Math.min(marquee[1], marquee[3]);
    o += rect(x, y, Math.abs(marquee[2] - marquee[0]), Math.abs(marquee[3] - marquee[1]), { class: 'ov-marquee', 'stroke-width': s });
  }
  overlayG.innerHTML = o;
}

// ---------------------------------------------------------------------------
// view

export function toWorld(ev) {
  const r = svg.getBoundingClientRect();
  return [(ev.clientX - r.left - app.view.x) / app.view.k, (ev.clientY - r.top - app.view.y) / app.view.k];
}

export function stageSize() {
  const r = stage.getBoundingClientRect();
  return { w: r.width, h: r.height };
}

export function zoomAt(factor, sx, sy) {
  const { w, h } = stageSize();
  if (sx === undefined) { sx = w / 2; sy = h / 2; }
  const k0 = app.view.k;
  const k = Math.max(0.1, Math.min(4, k0 * factor));
  const wx = (sx - app.view.x) / k0, wy = (sy - app.view.y) / k0;
  app.view.k = k;
  app.view.x = sx - wx * k;
  app.view.y = sy - wy * k;
  applyView();
  renderOverlay();
}

export function setZoom(k) { zoomAt(k / app.view.k); }

export function fit() {
  if (!last) renderCanvas();
  const b = last.bounds;
  const { w, h } = stageSize();
  if (!w || !h) return;
  const pad = 40;
  const k = Math.max(0.1, Math.min(1.25, (w - pad * 2) / Math.max(1, b.w), (h - pad * 2) / Math.max(1, b.h)));
  app.view.k = k;
  app.view.x = (w - b.w * k) / 2 - b.x * k;
  app.view.y = Math.max(pad - b.y * k, (h - b.h * k) / 2 - b.y * k);
  if (b.h * k > h - pad * 2) app.view.y = pad - b.y * k;
  applyView();
  renderOverlay();
}

export function centerOn(n) {
  const { w, h } = stageSize();
  const [cx, cy] = center(n);
  app.view.x = w / 2 - cx * app.view.k;
  app.view.y = h / 2 - cy * app.view.k;
  applyView();
  renderOverlay();
}

function onWheel(ev) {
  if (app.mode !== 'edit') return;
  ev.preventDefault();
  if (inlineState) commitInline();
  const r = svg.getBoundingClientRect();
  const unit = ev.deltaMode === 1 ? 16 : ev.deltaMode === 2 ? 400 : 1;
  if (ev.ctrlKey || ev.metaKey) {
    zoomAt(Math.max(0.8, Math.min(1.25, Math.exp(-ev.deltaY * unit * 0.0022))), ev.clientX - r.left, ev.clientY - r.top);
  } else {
    const dx = ev.shiftKey && !ev.deltaX ? ev.deltaY : ev.deltaX;
    const dy = ev.shiftKey && !ev.deltaX ? 0 : ev.deltaY;
    app.view.x -= dx * unit;
    app.view.y -= dy * unit;
    applyView();
  }
}

// ---------------------------------------------------------------------------
// hit testing helpers

function nodeAt(p, exclude = new Set()) {
  const nodes = app.doc.nodes;
  let container = null;
  for (let i = nodes.length - 1; i >= 0; i--) {
    const n = nodes[i];
    if (exclude.has(n.id)) continue;
    const m = 2 / app.view.k;
    if (p[0] >= n.x - m && p[0] <= n.x + n.w + m && p[1] >= n.y - m && p[1] <= n.y + n.h + m) {
      if (shapeDef(n.shape).container) { container = container || n; continue; }
      return n;
    }
  }
  return container;
}

// Group frames accept connections dropped close to their border (inside is for new shapes).
function nearBorder(n, p) {
  const m = 14 / app.view.k;
  const inX = p[0] > n.x + m && p[0] < n.x + n.w - m, inY = p[1] > n.y + m && p[1] < n.y + n.h - m;
  return !(inX && inY);
}

function nearPort(n, p) {
  let best = null, bd = 14 / app.view.k;
  for (const side of ['t', 'r', 'b', 'l']) {
    const q = portPoint(n, side, 0);
    const d = Math.hypot(q[0] - p[0], q[1] - p[1]);
    if (d < bd) { bd = d; best = { side, pt: q }; }
  }
  return best;
}

// ---------------------------------------------------------------------------
// snapping

function gridSnap(v) {
  const g = app.doc.settings.grid || 8;
  return Math.round(v / g) * g;
}

function snapBox(box, exclude, disable) {
  guides = [];
  if (disable) return [0, 0];
  const thr = 6 / app.view.k;
  const bx = [box.x, box.x + box.w / 2, box.x + box.w];
  const by = [box.y, box.y + box.h / 2, box.y + box.h];
  let bestX = null, bestY = null;
  const consider = (xs, ys, ref) => {
    for (let i = 0; i < 3; i++) for (let j = 0; j < xs.length; j++) {
      const d = xs[j] - bx[i];
      if (Math.abs(d) < thr && (!bestX || Math.abs(d) < Math.abs(bestX.d))) bestX = { d, x: xs[j], ref };
    }
    for (let i = 0; i < 3; i++) for (let j = 0; j < ys.length; j++) {
      const d = ys[j] - by[i];
      if (Math.abs(d) < thr && (!bestY || Math.abs(d) < Math.abs(bestY.d))) bestY = { d, y: ys[j], ref };
    }
  };
  for (const n of app.doc.nodes) {
    if (exclude.has(n.id)) continue;
    consider([n.x, n.x + n.w / 2, n.x + n.w], [n.y, n.y + n.h / 2, n.y + n.h], n);
  }
  // lane centres
  let pos = 0;
  for (const l of app.doc.lanes.items) {
    const c = pos + l.size / 2;
    if (app.doc.lanes.orientation === 'horizontal') {
      const d = c - by[1];
      if (Math.abs(d) < thr && (!bestY || Math.abs(d) < Math.abs(bestY.d))) bestY = { d, y: c, ref: null };
    } else {
      const d = c - bx[1];
      if (Math.abs(d) < thr && (!bestX || Math.abs(d) < Math.abs(bestX.d))) bestX = { d, x: c, ref: null };
    }
    pos += l.size;
  }
  const snap = app.doc.settings.snap;
  const dx = bestX ? bestX.d : snap ? gridSnap(bx[1]) - bx[1] : 0;
  const dy = bestY ? bestY.d : snap ? gridSnap(by[1]) - by[1] : 0;
  if (bestX) {
    const r = bestX.ref;
    const y0 = Math.min(box.y + dy, r ? r.y : box.y + dy) - 10, y1 = Math.max(box.y + box.h + dy, r ? r.y + r.h : 0) + 10;
    guides.push([bestX.x, y0, bestX.x, y1]);
  }
  if (bestY) {
    const r = bestY.ref;
    const x0 = Math.min(box.x + dx, r ? r.x : box.x + dx) - 10, x1 = Math.max(box.x + box.w + dx, r ? r.x + r.w : 0) + 10;
    guides.push([x0, bestY.y, x1, bestY.y]);
  }
  return [dx, dy];
}

// ---------------------------------------------------------------------------
// pointer handling

function begin(ev, d) {
  drag = { startClient: [ev.clientX, ev.clientY], moved: false, before: snapshot(), ...d };
  lastPointer = { clientX: ev.clientX, clientY: ev.clientY };
  try { svg.setPointerCapture(ev.pointerId); } catch (_) { /* ignore */ }
  if (drag.type !== 'pan' && !autoScrolling) { autoScrolling = true; requestAnimationFrame(autoScroll); }
}

// Scroll the view while a drag is held near the canvas edge.
let lastPointer = null, autoScrolling = false;
function autoScroll() {
  if (!drag || drag.type === 'pan') { autoScrolling = false; return; }
  requestAnimationFrame(autoScroll);
  if (!drag.moved || !lastPointer || !drag.move || pinch) return;
  const r = stage.getBoundingClientRect();
  const E = 36;
  const push = (v, lo, hi) => (v < lo + E ? -(lo + E - v) : v > hi - E ? v - (hi - E) : 0);
  const clamp = (v) => Math.max(-20, Math.min(20, v * 0.4));
  const vx = clamp(push(lastPointer.clientX, r.left, r.right)), vy = clamp(push(lastPointer.clientY, r.top, r.bottom));
  if (!vx && !vy) return;
  app.view.x -= vx; app.view.y -= vy;
  applyView();
  drag.move(toWorld(lastPointer), lastPointer);
}

function passedThreshold(ev) {
  if (drag.moved) return true;
  const d = Math.hypot(ev.clientX - drag.startClient[0], ev.clientY - drag.startClient[1]);
  if (d > 3) drag.moved = true;
  return drag.moved;
}

// Our own double-click detection: the overlay is re-rendered between clicks, which
// stops browsers from firing a native dblclick on the replaced element.
let lastDown = null;

// What a pointer event is aimed at, so two clicks only count as a double-click on the same thing.
function targetKey(t) {
  const d = t.dataset || {};
  if (d.quick || d.port || d.addwp !== undefined || d.divider !== undefined || d.pdivider !== undefined) return null;
  if (d.wp !== undefined) return 'wp:' + d.id + ':' + d.wp;
  if (d.seg !== undefined || d.ep) return 'edge:' + d.id;
  if (d.handle) return 'node:' + d.id;
  const l = t.closest('[data-label]'); if (l) return 'edge:' + l.dataset.label;
  const n = t.closest('[data-node]'); if (n) return 'node:' + n.dataset.node;
  const e = t.closest('[data-edge]'); if (e) return 'edge:' + e.dataset.edge;
  const h = t.closest('[data-lane-head]'); if (h) return 'lane:' + h.dataset.laneHead;
  const ph = t.closest('[data-phase-head]'); if (ph) return 'phase:' + ph.dataset.phaseHead;
  return 'empty';
}

// Two-finger pinch / pan on touch screens.
const touches = new Map();
let pinch = null;
function trackPointer(ev, down) {
  if (ev.pointerType !== 'touch') return false;
  if (down) touches.set(ev.pointerId, [ev.clientX, ev.clientY]);
  else if (touches.has(ev.pointerId)) touches.set(ev.pointerId, [ev.clientX, ev.clientY]);
  if (touches.size === 2) {
    const [a, b] = [...touches.values()];
    const mid = [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2];
    const dist = Math.hypot(a[0] - b[0], a[1] - b[1]) || 1;
    if (!pinch) {
      if (drag) cancelDrag();
      pinch = { mid, dist };
      return true;
    }
    const r = svg.getBoundingClientRect();
    app.view.x += mid[0] - pinch.mid[0];
    app.view.y += mid[1] - pinch.mid[1];
    zoomAt(dist / pinch.dist, mid[0] - r.left, mid[1] - r.top);
    pinch = { mid, dist };
    return true;
  }
  return !!pinch;
}
function endPointer(ev) {
  touches.delete(ev.pointerId);
  if (touches.size < 2 && pinch) { pinch = null; return true; }
  return false;
}

function onPointerDown(ev) {
  if (app.mode !== 'edit') return;
  clearTimeout(quickTimer); quickHover = null;
  if (trackPointer(ev, true)) return;
  if (ev.button === 0 && !spaceDown) {
    const now = performance.now();
    const key = targetKey(ev.target);
    if (key && lastDown && lastDown.key === key && now - lastDown.t < 450 && Math.hypot(ev.clientX - lastDown.x, ev.clientY - lastDown.y) < 6) {
      lastDown = null;
      ev.preventDefault();
      if (drag) drag = null;
      onDblClick(ev);
      return;
    }
    lastDown = { t: now, x: ev.clientX, y: ev.clientY, key };
  }
  closeMenus();
  if (inlineState) commitInline();
  if (document.activeElement && isTyping(document.activeElement)) document.activeElement.blur();
  stage.focus({ preventScroll: true });
  if (ev.button === 1 || (ev.button === 0 && spaceDown)) { ev.preventDefault(); return startPan(ev); }
  if (ev.button === 2) {
    // right click selects what is under the pointer
    const nodeEl = ev.target.closest('[data-node]');
    const edgeEl = ev.target.closest('[data-edge], [data-label]');
    if (nodeEl && !app.sel.nodes.has(nodeEl.dataset.node)) select({ nodes: [nodeEl.dataset.node] });
    else if (!nodeEl && edgeEl) { const id = edgeEl.dataset.edge || edgeEl.dataset.label; if (!app.sel.edges.has(id)) select({ edges: [id] }); }
    else if (!nodeEl && !edgeEl && !ev.target.dataset?.id) clearSelection();
    return;
  }
  if (ev.button !== 0) return;
  const t = ev.target;
  const d = t.dataset || {};
  const p = toWorld(ev);
  if (app.painter) {
    // format painter: apply to whatever was clicked
    const nodeEl = t.closest('[data-node]');
    const edgeEl = t.closest('[data-edge], [data-label]');
    const id = d.handle ? { node: d.id } : nodeEl ? { node: nodeEl.dataset.node } : edgeEl ? { edge: edgeEl.dataset.edge || edgeEl.dataset.label } : d.ep || d.seg !== undefined ? { edge: d.id } : null;
    emit({ paint: id || { none: true } });
    return;
  }
  if (d.handle) return startResize(ev, d.id, d.handle, p);
  if (d.port) return startConnect(ev, d.id, d.port, p);
  if (d.quick) { ev.preventDefault(); return startQuick(ev, d.id, d.quick, p); }
  if (d.ep) return startReconnect(ev, d.id, d.ep);
  if (d.seg !== undefined) return startSegment(ev, d.id, +d.seg, p);
  if (d.wp !== undefined) return startWaypoint(ev, d.id, +d.wp, false);
  if (d.addwp !== undefined) return startWaypoint(ev, d.id, +d.addwp, true);
  if (d.divider !== undefined) return startDivider(ev, +d.divider, p);
  if (d.pdivider !== undefined) return startDivider(ev, +d.pdivider, p, true);
  const label = t.closest('[data-label]');
  if (label) return startLabel(ev, label.dataset.label, p);
  const nodeEl = t.closest('[data-node]');
  if (nodeEl) return startMove(ev, nodeEl.dataset.node, p);
  const edgeEl = t.closest('[data-edge]');
  if (edgeEl) return startEdgeBody(ev, edgeEl.dataset.edge, p);
  const head = t.closest('[data-lane-head]');
  if (head) { select({ lane: head.dataset.laneHead }); return; }
  const phead = t.closest('[data-phase-head]');
  if (phead) { select({ phase: phead.dataset.phaseHead }); return; }
  startMarquee(ev, p);
}

function onPointerMove(ev) {
  if (trackPointer(ev, false)) return;
  if (!drag) {
    if (app.mode !== 'edit') return;
    const t = ev.target;
    // hovering a quick-add arrow opens a shape picker after a short pause
    const qk = t.dataset?.quick ? t.dataset.id + '|' + t.dataset.quick : null;
    if (qk !== quickHover) {
      clearTimeout(quickTimer);
      if (quickHover && !qk) emit({ quickPickerLeave: true });
      quickHover = qk;
      if (qk) {
        const r = t.getBoundingClientRect();
        const [qid, side] = qk.split('|');
        quickTimer = setTimeout(() => emit({ quickPicker: { id: qid, side, x: r.left + r.width / 2, y: r.top + r.height / 2 } }), 350);
      }
    }
    const nodeEl = t.closest?.('[data-node]');
    const id = nodeEl ? nodeEl.dataset.node : (t.dataset?.port ? t.dataset.id : (t.dataset?.handle || t.dataset?.quick ? hoverNode : null));
    const n = id && nodeById(id);
    const newHover = n && !shapeDef(n.shape).container ? id : null;
    if (newHover !== hoverNode) { hoverNode = newHover; renderOverlay(); }
    return;
  }
  if (!passedThreshold(ev)) return;
  // coalesce moves into animation frames; the last one is flushed on release
  const wasPending = !!pendingMove;
  pendingMove = { p: toWorld(ev), ev: { clientX: ev.clientX, clientY: ev.clientY, altKey: ev.altKey, shiftKey: ev.shiftKey, ctrlKey: ev.ctrlKey, metaKey: ev.metaKey } };
  lastPointer = pendingMove.ev;
  if (!wasPending) requestAnimationFrame(flushMove);
}

let pendingMove = null;
let quickHover = null, quickTimer = null;
function flushMove() {
  const m = pendingMove;
  pendingMove = null;
  if (m && drag && drag.move) drag.move(m.p, m.ev);
}

function onPointerUp(ev) {
  if (endPointer(ev)) { lastDown = null; return; }
  if (!drag) return;
  flushMove();
  if (!drag) return;
  const d = drag;
  drag = null;
  if (d.moved) lastDown = null; // a drag can't be the first half of a double-click
  try { svg.releasePointerCapture(ev.pointerId); } catch (_) { /* ignore */ }
  guides = [];
  connectPreview = null;
  const hadDrop = dropTargetId;
  dropTargetId = null;
  if (d.end) d.end(toWorld(ev), ev, d.moved);
  if (hadDrop || d.moved) renderCanvas(); else renderOverlay();
}

function startPan(ev) {
  const v0 = { ...app.view };
  begin(ev, { type: 'pan' });
  stage.classList.add('panning');
  drag.move = (p, e) => {
    app.view.x = v0.x + (e.clientX - drag.startClient[0]);
    app.view.y = v0.y + (e.clientY - drag.startClient[1]);
    applyView();
  };
  drag.end = () => stage.classList.remove('panning');
}

function containedIn(g) {
  return app.doc.nodes.filter((n) => n !== g && n.x >= g.x && n.y >= g.y && n.x + n.w <= g.x + g.w && n.y + n.h <= g.y + g.h);
}

function startMove(ev, id, p) {
  const additive = ev.shiftKey || ev.metaKey || ev.ctrlKey;
  let wasSelected = app.sel.nodes.has(id);
  if (additive) {
    toggleSelect('node', id);
    if (!app.sel.nodes.has(id)) return;
    wasSelected = false;
  } else if (!wasSelected) {
    select({ nodes: [id] });
  }
  const moving = new Map();
  for (const n of selectedNodes()) {
    moving.set(n.id, n);
    if (shapeDef(n.shape).container) for (const c of containedIn(n)) moving.set(c.id, c);
  }
  const orig = new Map([...moving.values()].map((n) => [n.id, [n.x, n.y]]));
  const endMoves = (e, ep) => (ep.node ? moving.has(ep.node) : app.sel.edges.has(e.id));
  const edges = app.doc.edges.filter((e) => endMoves(e, e.from) && endMoves(e, e.to))
    .map((e) => ({ e, wps: e.waypoints.map((q) => [...q]), from: { ...e.from }, to: { ...e.to } }));
  const box0 = unionBox([...moving.values()]);
  begin(ev, { type: 'move', hideHandles: true, fastNodes: new Set(moving.keys()), fastEdges: new Set(edges.map((it) => it.e.id)) });
  drag.move = (q, e) => {
    const dx0 = q[0] - p[0], dy0 = q[1] - p[1];
    const [sx, sy] = snapBox({ x: box0.x + dx0, y: box0.y + dy0, w: box0.w, h: box0.h }, new Set(moving.keys()), e.altKey);
    const dx = dx0 + sx, dy = dy0 + sy;
    for (const [nid, [x, y]] of orig) { const n = moving.get(nid); n.x = Math.round((x + dx) * 100) / 100; n.y = Math.round((y + dy) * 100) / 100; }
    for (const it of edges) {
      it.e.waypoints = it.wps.map(([x, y]) => [x + dx, y + dy]);
      if (!it.from.node) it.e.from = { x: it.from.x + dx, y: it.from.y + dy };
      if (!it.to.node) it.e.to = { x: it.to.x + dx, y: it.to.y + dy };
    }
    renderCanvas();
  };
  drag.end = (q, e, moved) => {
    if (moved) { growPhasesFor(app.doc, new Set(moving.keys())); commitFrom(drag0.before); }
    else if (wasSelected && !additive && app.sel.nodes.size > 1) select({ nodes: [id] });
  };
  const drag0 = drag;
}

function startResize(ev, id, handle, p) {
  const n = nodeById(id);
  const o = { x: n.x, y: n.y, w: n.w, h: n.h };
  begin(ev, { type: 'resize', fastNodes: new Set([id]) });
  const d0 = drag;
  drag.move = (q, e) => {
    let dx = q[0] - p[0], dy = q[1] - p[1];
    let { x, y, w, h } = o;
    if (handle.includes('e')) w = o.w + dx;
    if (handle.includes('w')) { w = o.w - dx; x = o.x + dx; }
    if (handle.includes('s')) h = o.h + dy;
    if (handle.includes('n')) { h = o.h - dy; y = o.y + dy; }
    if (e.shiftKey && handle.length === 2) {
      const ratio = o.w / o.h;
      if (w / h > ratio) w = h * ratio; else h = w / ratio;
      if (handle.includes('w')) x = o.x + o.w - w;
      if (handle.includes('n')) y = o.y + o.h - h;
    }
    if (app.doc.settings.snap && !e.altKey && !e.shiftKey) {
      const g = app.doc.settings.grid || 8;
      if (handle.includes('e')) w = Math.round((x + w) / g) * g - x;
      if (handle.includes('w')) { const nx = Math.round(x / g) * g; w += x - nx; x = nx; }
      if (handle.includes('s')) h = Math.round((y + h) / g) * g - y;
      if (handle.includes('n')) { const ny = Math.round(y / g) * g; h += y - ny; y = ny; }
    }
    const min = 16;
    if (w < min) { if (handle.includes('w')) x -= min - w; w = min; }
    if (h < min) { if (handle.includes('n')) y -= min - h; h = min; }
    Object.assign(n, { x, y, w, h });
    renderCanvas();
  };
  drag.end = (q, e, moved) => { if (moved) { growPhasesFor(app.doc, new Set([id])); commitFrom(d0.before); } };
}

function defaultEdgeProps() {
  const doc = app.doc;
  const kind = app.connectDefaults.kind && doc.edgeKinds.some((k) => k.id === app.connectDefaults.kind) ? app.connectDefaults.kind : doc.edgeKinds[0].id;
  return { kind, routing: app.connectDefaults.routing || 'orthogonal' };
}

function startConnect(ev, id, side, p) {
  const src = nodeById(id);
  const start = portPoint(src, side, 0);
  begin(ev, { type: 'connect' });
  const d0 = drag;
  let target = null, tside = 'auto';
  drag.move = (q) => {
    const n = nodeAt(q, new Set([id]));
    target = n && (!shapeDef(n.shape).container || nearBorder(n, q)) ? n : null;
    dropTargetId = target ? target.id : null;
    let endPt = q, snap = null;
    tside = 'auto';
    if (target) {
      const np = nearPort(target, q);
      if (np) { tside = np.side; endPt = np.pt; snap = np.pt; } else endPt = center(target);
    }
    connectPreview = { d: `M${start[0]} ${start[1]}L${endPt[0]} ${endPt[1]}`, snap };
    renderCanvas();
  };
  drag.end = (q, e, moved) => {
    if (!moved) return;
    const props = defaultEdgeProps();
    if (target) {
      change((doc) => { doc.edges.push(makeEdge({ node: id, side }, { node: target.id, side: tside }, props)); });
      const ne = app.doc.edges[app.doc.edges.length - 1];
      select({ edges: [ne.id] });
    } else if (e.altKey || e.shiftKey) {
      change((doc) => { doc.edges.push(makeEdge({ node: id, side }, { x: Math.round(q[0]), y: Math.round(q[1]) }, props)); });
      select({ edges: [app.doc.edges[app.doc.edges.length - 1].id] });
    } else {
      // drop on empty canvas: create a new shape there, connected
      const shape = ['decision', 'terminator', 'connector', 'offpage', 'text', 'annotation', 'note', 'actor'].includes(src.shape) ? app.lastShape : src.shape;
      const def = shapeDef(shape);
      const nn = makeNode(shape, gridSnap(q[0] - def.w / 2), gridSnap(q[1] - def.h / 2), { text: '' });
      change((doc) => {
        doc.nodes.push(nn);
        doc.edges.push(makeEdge({ node: id, side }, { node: nn.id, side: 'auto' }, props));
        growPhasesFor(doc, new Set([nn.id]));
      });
      select({ nodes: [nn.id] });
      renderCanvas();
      editNodeText(nn.id);
    }
    void d0;
  };
}

// Quick-add arrows: click adds a connected shape, drag draws a connector from that side.
function startQuick(ev, id, side, p) {
  begin(ev, { type: 'pending' });
  drag.move = (q, e2) => {
    drag = null;
    startConnect({ pointerId: ev.pointerId, clientX: ev.clientX, clientY: ev.clientY }, id, side, p);
    drag.moved = true;
    drag.move(q, e2);
  };
  drag.end = (q, e2, moved) => { if (!moved) quickAdd(id, side); };
}

export function quickAdd(id, side, chosen) {
  const n = nodeById(id);
  if (!n) return;
  // steps continue as the same kind of step; anything else (start, decision, data…) is followed by a process
  const shape = chosen || (['process', 'rounded', 'predefined', 'document', 'multidoc'].includes(n.shape) ? n.shape : 'process');
  const def = shapeDef(shape);
  const gap = 48;
  const [cx, cy] = center(n);
  let x, y;
  const place = (mult) => {
    if (side === 'r') { x = n.x + n.w + gap + mult * (def.w + gap); y = cy - def.h / 2; }
    if (side === 'l') { x = n.x - gap - def.w - mult * (def.w + gap); y = cy - def.h / 2; }
    if (side === 'b') { x = cx - def.w / 2; y = n.y + n.h + gap + mult * (def.h + gap); }
    if (side === 't') { x = cx - def.w / 2; y = n.y - gap - def.h - mult * (def.h + gap); }
  };
  let m = 0;
  place(0);
  while (m < 12 && app.doc.nodes.some((o) => o.x < x + def.w + 8 && o.x + o.w > x - 8 && o.y < y + def.h + 8 && o.y + o.h > y - 8 && !shapeDef(o.shape).container)) place(++m);
  // moving across lanes: centre the new shape in the lane it lands in
  const L = app.doc.lanes;
  const across = L.orientation === 'horizontal' ? side === 't' || side === 'b' : side === 'l' || side === 'r';
  if (L.items.length && across) {
    let pos = 0;
    const c = L.orientation === 'horizontal' ? y + def.h / 2 : x + def.w / 2;
    for (const l of L.items) {
      if (c >= pos && c < pos + l.size) {
        if (L.orientation === 'horizontal') y = pos + l.size / 2 - def.h / 2; else x = pos + l.size / 2 - def.w / 2;
        break;
      }
      pos += l.size;
    }
  }
  const nn = makeNode(shape, Math.round(x), Math.round(y), { text: '' });
  change((doc) => {
    doc.nodes.push(nn);
    doc.edges.push(makeEdge({ node: id, side }, { node: nn.id, side: OPP[side] }, defaultEdgeProps()));
    growPhasesFor(doc, new Set([nn.id]));
  });
  select({ nodes: [nn.id] });
  renderCanvas();
  editNodeText(nn.id);
  return nn;
}

function startReconnect(ev, id, which) {
  const e = edgeById(id);
  const other = which === 'from' ? e.to : e.from;
  begin(ev, { type: 'reconnect', hideHandles: false, fastEdges: new Set([id]) });
  const d0 = drag;
  let target = null, tside = 'auto';
  drag.move = (q) => {
    const n = nodeAt(q);
    target = n && (!shapeDef(n.shape).container || nearBorder(n, q)) ? n : null;
    dropTargetId = target ? target.id : null;
    tside = 'auto';
    let snap = null;
    if (target) { const np = nearPort(target, q); if (np) { tside = np.side; snap = np.pt; } }
    e[which] = target ? { node: target.id, side: tside, offset: 0 } : { x: Math.round(q[0]), y: Math.round(q[1]) };
    connectPreview = snap ? { d: '', snap } : null;
    renderCanvas();
  };
  drag.end = (q, ev2, moved) => { if (moved) commitFrom(d0.before); };
}

function startSegment(ev, id, seg, p) {
  const e = edgeById(id);
  const r = last.routes.get(id);
  let pts = r.pts.map((q) => [...q]);
  let i = seg;
  if (i === 0) { pts.splice(1, 0, [...pts[0]]); i = 1; }
  if (i === pts.length - 2) { pts.splice(pts.length - 1, 0, [...pts[pts.length - 1]]); }
  const a0 = [...pts[i]], b0 = [...pts[i + 1]];
  const vert = Math.abs(a0[0] - b0[0]) < 0.5;
  select({ edges: [id] });
  if (e.label && (e.labelPos === null || e.labelPos === undefined)) {
    // pin an auto-placed label so rerouting doesn't send it somewhere else
    const g = labelGeometry(app.doc, e, r);
    const pr = projectOnPolyline(r.poly, [g.cx, g.cy]);
    e.labelPos = Math.round(pr.frac * 1000) / 1000;
    e.labelOffset = [Math.round(g.cx - pr.pt[0]), Math.round(g.cy - pr.pt[1])];
  }
  begin(ev, { type: 'segment', fastEdges: new Set([id]) });
  const d0 = drag;
  drag.move = (q, e2) => {
    let delta = vert ? q[0] - p[0] : q[1] - p[1];
    if (app.doc.settings.snap && !e2.altKey) {
      const base = vert ? a0[0] : a0[1];
      delta = gridSnap(base + delta) - base;
    }
    if (vert) { pts[i] = [a0[0] + delta, a0[1]]; pts[i + 1] = [b0[0] + delta, b0[1]]; }
    else { pts[i] = [a0[0], a0[1] + delta]; pts[i + 1] = [b0[0], b0[1] + delta]; }
    e.waypoints = pts.slice(1, -1).map((q2) => [Math.round(q2[0] * 100) / 100, Math.round(q2[1] * 100) / 100]);
    renderCanvas();
  };
  drag.end = (q, e2, moved) => {
    if (!moved) return;
    // drop redundant waypoints
    const full = simplify([pts[0], ...e.waypoints, pts[pts.length - 1]]);
    e.waypoints = full.slice(1, -1);
    commitFrom(d0.before);
  };
}

function startWaypoint(ev, id, idx, insert) {
  const e = edgeById(id);
  begin(ev, { type: 'waypoint', fastEdges: new Set([id]) });
  const d0 = drag;
  let index = idx;
  let inserted = false;
  drag.move = (q, e2) => {
    if (insert && !inserted) { e.waypoints.splice(index, 0, [q[0], q[1]]); inserted = true; }
    const snapOn = app.doc.settings.snap && !e2.altKey;
    e.waypoints[index] = snapOn ? [gridSnap(q[0]), gridSnap(q[1])] : [q[0], q[1]];
    renderCanvas();
  };
  drag.end = (q, e2, moved) => { if (moved) commitFrom(d0.before); };
}

function startEdgeBody(ev, id, p) {
  const additive = ev.shiftKey || ev.metaKey || ev.ctrlKey;
  if (additive) { toggleSelect('edge', id); return; }
  if (!app.sel.edges.has(id) || app.sel.nodes.size) select({ edges: [id] });
  const e = edgeById(id);
  const r = last.routes.get(id);
  const proj = projectOnPolyline(r.pts, p);
  if (e.routing === 'orthogonal') {
    // behaves like dragging the nearest segment
    const d = { startClient: [ev.clientX, ev.clientY] };
    begin(ev, { type: 'pending' });
    drag.move = (q, e2) => {
      const moved = drag.moved;
      drag = null;
      startSegment({ ...ev, pointerId: ev.pointerId, clientX: d.startClient[0], clientY: d.startClient[1] }, id, proj.seg, p);
      drag.moved = moved;
      drag.move(q, e2);
    };
    drag.end = () => {};
  } else {
    // insert a waypoint on the segment of the control polygon nearest the pointer
    const ctrl = [r.pts[0], ...e.waypoints, r.pts[r.pts.length - 1]];
    const cp = projectOnPolyline(ctrl, p);
    begin(ev, { type: 'pending' });
    drag.move = (q, e2) => {
      const moved = drag.moved;
      drag = null;
      startWaypoint({ ...ev, clientX: ev.clientX, clientY: ev.clientY, pointerId: ev.pointerId }, id, cp.seg, true);
      drag.moved = moved;
      drag.move(q, e2);
    };
    drag.end = () => {};
  }
}

function startLabel(ev, id, p) {
  if (!app.sel.edges.has(id) || app.sel.nodes.size) select({ edges: [id] });
  const e = edgeById(id);
  const r = last.routes.get(id);
  const g = labelGeometry(app.doc, e, r);
  const grab = [p[0] - g.cx, p[1] - g.cy];
  begin(ev, { type: 'label', fastEdges: new Set([id]) });
  const d0 = drag;
  drag.move = (q) => {
    const c = [q[0] - grab[0], q[1] - grab[1]];
    const pr = projectOnPolyline(r.poly, c);
    e.labelPos = Math.round(pr.frac * 1000) / 1000;
    let off = [c[0] - pr.pt[0], c[1] - pr.pt[1]];
    if (Math.hypot(off[0], off[1]) < 6 / app.view.k) off = [0, 0];
    e.labelOffset = [Math.round(off[0]), Math.round(off[1])];
    e.labelAlign = undefined;
    renderCanvas();
  };
  drag.end = (q, e2, moved) => { if (moved) commitFrom(d0.before); };
}

function startDivider(ev, i, p, phase = false) {
  const doc = app.doc;
  let horiz, lane, boundary;
  if (phase) {
    horiz = phaseAxis(doc).along === 'y';
    const [ph, pos] = phaseStarts(doc)[i];
    lane = ph; boundary = pos + ph.size;
  } else {
    horiz = doc.lanes.orientation === 'horizontal';
    lane = doc.lanes.items[i];
    boundary = 0;
    for (let j = 0; j <= i; j++) boundary += doc.lanes.items[j].size;
  }
  const size0 = lane.size;
  const axis = horiz ? 1 : 0;
  const nodes = doc.nodes.filter((n) => (phase ? n[horiz ? 'y' : 'x'] >= boundary - 1 : center(n)[axis] > boundary)).map((n) => [n, n[horiz ? 'y' : 'x']]);
  const edges = doc.edges.map((e) => ({ e, wps: e.waypoints.map((q) => [...q]), from: { ...e.from }, to: { ...e.to } }));
  select(phase ? { phase: lane.id } : { lane: lane.id });
  begin(ev, phase ? { type: 'divider', pdivider: i } : { type: 'divider', divider: i });
  const d0 = drag;
  drag.move = (q, e2) => {
    const min = phase ? 24 : 40;
    let size = Math.max(min, size0 + (q[axis] - p[axis]));
    if (doc.settings.snap && !e2.altKey) size = Math.max(min, gridSnap(size));
    const delta = size - size0;
    lane.size = size;
    for (const [n, v] of nodes) n[horiz ? 'y' : 'x'] = v + delta;
    for (const it of edges) {
      it.e.waypoints = it.wps.map((w) => { const c = [...w]; if (c[axis] > boundary) c[axis] += delta; return c; });
      for (const k of ['from', 'to']) {
        if (!it[k].node) { const c = { ...it[k] }; const key = horiz ? 'y' : 'x'; if (c[key] > boundary) c[key] += delta; it.e[k] = c; }
      }
    }
    renderCanvas();
  };
  drag.end = (q, e2, moved) => { if (moved) commitFrom(d0.before); };
}

function startMarquee(ev, p) {
  const additive = ev.shiftKey || ev.metaKey || ev.ctrlKey;
  const prevNodes = additive ? new Set(app.sel.nodes) : new Set();
  const prevEdges = additive ? new Set(app.sel.edges) : new Set();
  begin(ev, { type: 'marquee' });
  drag.move = (q) => {
    marquee = [p[0], p[1], q[0], q[1]];
    renderOverlay();
  };
  drag.end = (q, e, moved) => {
    if (!moved) { marquee = null; if (!additive) clearSelection(); return; }
    const x0 = Math.min(p[0], q[0]), x1 = Math.max(p[0], q[0]);
    const y0 = Math.min(p[1], q[1]), y1 = Math.max(p[1], q[1]);
    const inside = (b) => b.x >= x0 && b.y >= y0 && b.x + b.w <= x1 && b.y + b.h <= y1;
    const nodes = app.doc.nodes.filter(inside).map((n) => n.id);
    const nodeSet = new Set(nodes);
    const edges = app.doc.edges.filter((ed) => {
      const r = last.routes.get(ed.id);
      const allIn = r.poly.every(([x, y]) => x >= x0 && x <= x1 && y >= y0 && y <= y1);
      return allIn || ((ed.from.node ? nodeSet.has(ed.from.node) : false) && (ed.to.node ? nodeSet.has(ed.to.node) : false));
    }).map((ed) => ed.id);
    marquee = null;
    select({ nodes: [...prevNodes, ...nodes], edges: [...prevEdges, ...edges] });
  };
}

// ---------------------------------------------------------------------------
// double click & inline editing

function onDblClick(ev) {
  if (app.mode !== 'edit') return;
  const t = document.elementFromPoint(ev.clientX, ev.clientY) || ev.target;
  const d = t.dataset || {};
  const p = toWorld(ev);
  if (d.wp !== undefined) {
    const e = edgeById(d.id);
    change(() => { e.waypoints.splice(+d.wp, 1); });
    return;
  }
  if (d.seg !== undefined || d.ep) { editEdgeLabel(d.id); return; }
  const label = t.closest('[data-label]');
  if (label) { editEdgeLabel(label.dataset.label); return; }
  const nodeEl = t.closest('[data-node]') || (d.handle && document.querySelector(`[data-node="${CSS.escape(d.id)}"]`));
  if (nodeEl) { editNodeText(nodeEl.dataset.node); return; }
  const edgeEl = t.closest('[data-edge]');
  if (edgeEl) { editEdgeLabel(edgeEl.dataset.edge, p); return; }
  const head = t.closest('[data-lane-head]');
  if (head) { editLaneTitle(head.dataset.laneHead); return; }
  const phead = t.closest('[data-phase-head]');
  if (phead) { editPhaseTitle(phead.dataset.phaseHead); return; }
  // empty canvas: create a shape
  const shape = app.lastShape || 'process';
  const def = shapeDef(shape);
  const nn = makeNode(shape, gridSnap(p[0] - def.w / 2), gridSnap(p[1] - def.h / 2), { text: '' });
  change((doc) => { doc.nodes.push(nn); });
  select({ nodes: [nn.id] });
  renderCanvas();
  editNodeText(nn.id);
}

function startInline(state) {
  if (inlineState) commitInline();
  inlineState = state;
  inline.hidden = false;
  inline.value = state.value;
  inline.style.textAlign = state.align || 'center';
  inline.style.fontFamily = state.fontFamily || '';
  inline.style.fontWeight = state.fontWeight || 500;
  positionInline();
  inline.focus();
  inline.select();
}

function positionInline() {
  const s = inlineState;
  if (!s) return;
  const k = app.view.k;
  const b = s.box();
  const fs = s.fontSize * k;
  inline.style.fontSize = fs + 'px';
  inline.style.left = (b.x * k + app.view.x) + 'px';
  inline.style.top = (b.y * k + app.view.y) + 'px';
  inline.style.width = Math.max(60, b.w * k) + 'px';
  inline.style.minHeight = Math.max(fs * 1.6, b.h * k) + 'px';
  sizeInline();
}

function sizeInline() {
  // grow with the text, and centre it vertically like the shape does
  const minH = parseFloat(inline.style.minHeight) || 0;
  inline.style.paddingTop = inline.style.paddingBottom = '2px';
  inline.style.height = 'auto';
  const content = inline.scrollHeight;
  const pad = Math.max(2, (minH - content) / 2 + 2);
  inline.style.paddingTop = inline.style.paddingBottom = pad + 'px';
  inline.style.height = Math.max(content + 2, minH) + 'px';
}

function onInlineKey(ev) {
  ev.stopPropagation();
  if (ev.key === 'Escape') { ev.preventDefault(); commitInline(false); stage.focus(); }
  else if (ev.key === 'Enter' && !ev.shiftKey && !ev.altKey) { ev.preventDefault(); commitInline(); stage.focus(); }
  else if (ev.key === 'Tab') { ev.preventDefault(); commitInline(); stage.focus(); }
}

export function commitInline(save = true) {
  const s = inlineState;
  if (!s) return;
  inlineState = null;
  const v = inline.value;
  inline.hidden = true;
  if (save && v !== s.value) s.commit(v);
  else if (save && s.commitAlways) s.commit(v);
}

export const isInlineEditing = () => !!inlineState;

export function editNodeText(id) {
  const n = nodeById(id);
  if (!n) return;
  const st = nodeStyle(app.doc, n);
  startInline({
    value: n.text || '',
    align: st.align,
    fontSize: st.fontSize,
    fontFamily: fontFor(app.doc).text,
    fontWeight: st.bold ? 600 : 500,
    box: () => {
      const tb = shapeTextBox(n.shape, n.w, n.h);
      const w = Math.max(tb.w, 80), h = Math.max(tb.h, 24);
      return { x: n.x + tb.x + tb.w / 2 - w / 2, y: n.y + tb.y + tb.h / 2 - h / 2, w, h };
    },
    commit: (v) => change(() => { n.text = v; fitNodeText(app.doc, n); growPhasesFor(app.doc, new Set([n.id])); }),
  });
}

export function editEdgeLabel(id) {
  const e = edgeById(id);
  if (!e) return;
  select({ edges: [id] });
  const r = last.routes.get(id);
  startInline({
    value: e.label || '',
    align: 'center',
    fontSize: 11,
    fontFamily: fontFor(app.doc).text,
    box: () => {
      const g = labelGeometry(app.doc, { ...e, label: e.label || 'M' }, last.routes.get(id) || r);
      const w = Math.max(100, g.w + 20);
      return { x: g.cx - w / 2, y: g.cy - 10, w, h: 20 };
    },
    commit: (v) => change(() => { e.label = v; }),
  });
}

export function editLaneTitle(id) {
  const lane = app.doc.lanes.items.find((l) => l.id === id);
  if (!lane) return;
  select({ lane: id });
  startInline({
    value: lane.title || '',
    align: 'center',
    fontSize: 13,
    fontFamily: fontFor(app.doc).text,
    fontWeight: 600,
    box: () => {
      let pos = 0;
      for (const l of app.doc.lanes.items) { if (l === lane) break; pos += l.size; }
      const H = app.doc.lanes.header || 40;
      return app.doc.lanes.orientation === 'horizontal' ? { x: 0, y: pos + lane.size / 2 - 12, w: Math.max(H, 120), h: 24 } : { x: pos + 4, y: 8, w: lane.size - 8, h: 24 };
    },
    commit: (v) => change(() => { lane.title = v; }),
  });
}

export function editPhaseTitle(id) {
  const ph = app.doc.phases.items.find((l) => l.id === id);
  if (!ph) return;
  select({ phase: id });
  startInline({
    value: ph.title || '',
    align: 'center',
    fontSize: 12,
    fontFamily: fontFor(app.doc).text,
    fontWeight: 600,
    box: () => {
      const pos = phaseStarts(app.doc).find(([p]) => p === ph)[1];
      const H = app.doc.phases.header || 34;
      return phaseAxis(app.doc).along === 'y' ? { x: -H, y: pos + ph.size / 2 - 12, w: 180, h: 24 } : { x: pos + 4, y: -H + 4, w: ph.size - 8, h: 24 };
    },
    commit: (v) => change(() => { ph.title = v; }),
  });
}

// ---------------------------------------------------------------------------
// adding shapes

// Nearest position (spiralling out from x,y) where a w×h shape doesn't touch anything.
export function findFreeSpot(x, y, w, h, ignoreAll = false) {
  if (ignoreAll) return [x, y];
  const solid = app.doc.nodes.filter((n) => !shapeDef(n.shape).container);
  const free = (px, py) => !solid.some((n) => px < n.x + n.w + 16 && px + w + 16 > n.x && py < n.y + n.h + 16 && py + h + 16 > n.y);
  if (free(x, y)) return [x, y];
  for (let r = 24; r < 1600; r += 24) {
    for (let a = 0; a < 16; a++) {
      const px = x + Math.cos((a / 16) * Math.PI * 2) * r, py = y + Math.sin((a / 16) * Math.PI * 2) * r;
      if (free(px, py)) return [px, py];
    }
  }
  return [x, y];
}

export function addShape(shape, at) {
  const def = shapeDef(shape);
  let x, y;
  if (at) { x = at[0] - def.w / 2; y = at[1] - def.h / 2; }
  else {
    const { w, h } = stageSize();
    [x, y] = findFreeSpot((w / 2 - app.view.x) / app.view.k - def.w / 2, (h / 2 - app.view.y) / app.view.k - def.h / 2, def.w, def.h, def.container);
  }
  const nn = makeNode(shape, Math.round(x + def.w / 2 - def.w / 2), Math.round(y));
  if (app.doc.settings.snap) {
    nn.x = gridSnap(nn.x + def.w / 2) - def.w / 2;
    nn.y = gridSnap(nn.y + def.h / 2) - def.h / 2;
  }
  if (shape !== 'text' && shape !== 'group' && shape !== 'annotation' && shape !== 'note') app.lastShape = shape;
  change((doc) => {
    if (def.container) doc.nodes.unshift(nn); else doc.nodes.push(nn);
    growPhasesFor(doc, new Set([nn.id]));
  });
  select({ nodes: [nn.id] });
  return nn;
}

function onDrop(ev) {
  if (app.mode !== 'edit') return;
  const shape = ev.dataTransfer.getData('application/x-flow-shape');
  if (shape) {
    ev.preventDefault();
    addShape(shape, toWorld(ev));
    return;
  }
  const f = ev.dataTransfer.files && ev.dataTransfer.files[0];
  if (f) { ev.preventDefault(); emit({ openFile: f }); }
}

export function closeMenus() {
  document.querySelectorAll('.menu').forEach((m) => { m.hidden = true; });
  document.querySelectorAll('[aria-expanded="true"]').forEach((b) => b.setAttribute('aria-expanded', 'false'));
}

export { gridSnap };
