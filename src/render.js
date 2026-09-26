// Turns a document into SVG markup. Used by the editor canvas and by every export,
// so what you see while editing is exactly what gets exported.
import { SHAPES, SHAPE_ORDER, shapeDef, shapeTextBox } from './shapes.js';
import { routeEdge, pathD, routePolyline, obstaclesFor, longestSegmentMid, pointAt, unionBox, findJumps } from './geometry.js';
import { laneRects, shapeLabel, phaseAxis, phaseStarts, nodeBadges } from './model.js';

// All fonts are local: IBM Plex is embedded (src/fonts.css); the others use fonts installed on the computer.
export const FONTS = {
  plex: {
    name: 'IBM Plex (built in)',
    text: '"IBM Plex Sans Condensed", "Arial Narrow", system-ui, sans-serif',
    ui: '"IBM Plex Sans", system-ui, -apple-system, "Segoe UI", sans-serif',
    mono: '"IBM Plex Mono", ui-monospace, "SFMono-Regular", Menlo, monospace',
    embedded: true,
  },
  system: {
    name: 'System sans',
    text: 'system-ui, -apple-system, "Segoe UI", Roboto, sans-serif',
    ui: 'system-ui, -apple-system, "Segoe UI", Roboto, sans-serif',
    mono: 'ui-monospace, "SFMono-Regular", Menlo, Consolas, monospace',
  },
  inter: {
    name: 'Sans (Inter if installed)',
    text: 'Inter, "Helvetica Neue", Arial, system-ui, sans-serif',
    ui: 'Inter, "Helvetica Neue", Arial, system-ui, sans-serif',
    mono: '"JetBrains Mono", ui-monospace, Consolas, monospace',
  },
  serif: {
    name: 'Serif',
    text: '"Source Serif 4", Georgia, "Times New Roman", serif',
    ui: '"Source Sans 3", system-ui, sans-serif',
    mono: '"Source Code Pro", ui-monospace, Consolas, monospace',
  },
  hand: {
    name: 'Handwritten',
    text: '"Patrick Hand", "Segoe Print", "Comic Sans MS", cursive',
    ui: '"Patrick Hand", "Segoe Print", "Comic Sans MS", cursive',
    mono: '"Patrick Hand", ui-monospace, monospace',
  },
};

export const fontFor = (doc) => FONTS[doc.settings.font] || FONTS.plex;

// Theme variables shared by the editor and exported pages.
export const THEME_LIGHT = {
  '--bg': '#F3F5F4', '--panel': '#FFFFFF', '--ink': '#18201D', '--muted': '#58645E', '--faint': '#8A958F',
  '--line': '#C6CEC9', '--lane': '#FFFFFF', '--lane-alt': '#EAEEEC', '--lane-head': '#DCE3DF',
  '--flow': '#2B5A84', '--feed': '#7E8A84', '--node': '#FFFFFF', '--node-stroke': '#33413B',
  '--dec': '#E9F0F7', '--term': '#E4E9E6', '--note': '#FFF6D6', '--group': 'rgba(43,90,132,0.05)',
  '--risk': '#B8432B', '--risk-ink': '#FFFFFF', '--tag': '#2B5A84', '--control': '#2E7D32', '--control-ink': '#FFFFFF',
  '--sel': '#2B5A84', '--sel-fill': '#DCE9F5', '--label-bg': '#F3F5F4',
};
export const THEME_DARK = {
  '--bg': '#121715', '--panel': '#1A211E', '--ink': '#E4EAE7', '--muted': '#A3AFA9', '--faint': '#76827C',
  '--line': '#34403A', '--lane': '#161C19', '--lane-alt': '#1B2320', '--lane-head': '#232D29',
  '--flow': '#7FB0DD', '--feed': '#8C9892', '--node': '#1F2724', '--node-stroke': '#9AA8A1',
  '--dec': '#1D2A36', '--term': '#26302C', '--note': '#3A3522', '--group': 'rgba(127,176,221,0.06)',
  '--risk': '#E0775C', '--risk-ink': '#121715', '--tag': '#7FB0DD', '--control': '#7CC47F', '--control-ink': '#121715',
  '--sel': '#7FB0DD', '--sel-fill': '#22364A', '--label-bg': '#161C19',
};

const varsBlock = (t) => Object.entries(t).map(([k, v]) => `${k}:${v};`).join('');
export function themeCSS(scope = ':root') {
  return `${scope}{color-scheme:light;${varsBlock(THEME_LIGHT)}}
@media (prefers-color-scheme: dark){${scope}:not([data-theme="light"]){color-scheme:dark;${varsBlock(THEME_DARK)}}}
${scope}[data-theme="dark"]{color-scheme:dark;${varsBlock(THEME_DARK)}}`;
}

// Replace var(--x) references with concrete colours (for SVG/PNG export).
export function resolveVars(markup, dark = false) {
  const t = dark ? THEME_DARK : THEME_LIGHT;
  return markup.replace(/var\((--[a-z-]+)\)/g, (m, k) => t[k] || m);
}

// ---- text measurement ------------------------------------------------------

let ctx = null;
const widthCache = new Map();
export function measure(text, font) {
  const key = font + '|' + text;
  let w = widthCache.get(key);
  if (w !== undefined) return w;
  if (!ctx && typeof document !== 'undefined') ctx = document.createElement('canvas').getContext('2d');
  if (ctx) { ctx.font = font; w = ctx.measureText(text).width; }
  else w = text.length * 6.2;
  if (widthCache.size > 5000) widthCache.clear();
  widthCache.set(key, w);
  return w;
}
export function clearMeasureCache() { widthCache.clear(); }

export function wrapText(text, maxW, font) {
  const out = [];
  for (const para of String(text ?? '').split('\n')) {
    const words = para.split(/\s+/).filter(Boolean);
    if (!words.length) { out.push(''); continue; }
    let cur = '';
    for (let w of words) {
      const trial = cur ? cur + ' ' + w : w;
      if (measure(trial, font) <= maxW || !cur) {
        if (!cur && measure(w, font) > maxW && w.length > 1) {
          // break a very long word
          let piece = '';
          for (const ch of w) {
            if (measure(piece + ch, font) > maxW && piece) { out.push(piece); piece = ch; }
            else piece += ch;
          }
          cur = piece;
        } else cur = trial;
      } else { out.push(cur); cur = w; if (measure(w, font) > maxW) { out.push(w); cur = ''; } }
    }
    if (cur) out.push(cur);
  }
  return out;
}

// ---- markup helpers ----------------------------------------------------------

export const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const attrs = (o) => Object.entries(o).filter(([, v]) => v !== undefined && v !== null && v !== '').map(([k, v]) => ` ${k}="${esc(v)}"`).join('');
const tag = (t, a, inner) => (inner === undefined ? `<${t}${attrs(a)}/>` : `<${t}${attrs(a)}>${inner}</${t}>`);

function hash(s) {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36);
}

export const ARROWS = {
  none: null,
  arrow: { d: 'M0 0L10 5L0 10z', fill: true },
  open: { d: 'M0 0L10 5L0 10', fill: false },
  diamond: { d: 'M0 5L5 0L10 5L5 10z', fill: true },
  diamondOpen: { d: 'M0 5L5 0L10 5L5 10z', fill: false },
  circle: { d: 'M5 1A4 4 0 1 1 5 9A4 4 0 1 1 5 1z', fill: true },
  circleOpen: { d: 'M5 1A4 4 0 1 1 5 9A4 4 0 1 1 5 1z', fill: false },
  bar: { d: 'M9 0V10', fill: false },
};

export function edgeStyle(doc, e) {
  const kind = doc.edgeKinds.find((k) => k.id === e.kind) || doc.edgeKinds[0] || {};
  const s = e.style || {};
  return {
    color: s.color || kind.color || 'var(--flow)',
    width: s.width ?? kind.width ?? 1.6,
    dash: s.dash ?? kind.dash ?? '',
    start: s.startArrow ?? kind.startArrow ?? 'none',
    end: s.endArrow ?? kind.endArrow ?? 'arrow',
    labelColor: s.labelColor || kind.labelColor || s.color || kind.color || 'var(--flow)',
  };
}

export function nodeStyle(doc, n) {
  const def = shapeDef(n.shape);
  const s = n.style || {};
  return {
    fill: s.fill || def.fill || 'var(--node)',
    stroke: s.stroke || (def.stroke === 'none' ? 'none' : 'var(--node-stroke)'),
    strokeWidth: s.strokeWidth ?? 1.4,
    dash: s.dash ?? def.dash ?? '',
    textColor: s.textColor || 'var(--ink)',
    fontSize: s.fontSize || doc.settings.fontSize || 12.5,
    bold: s.bold ?? false,
    italic: s.italic ?? false,
    align: s.align || def.align || 'center',
    valign: s.valign || def.valign || 'middle',
    opacity: s.opacity ?? 1,
    shadow: s.shadow ?? false,
  };
}

// Grow a shape's height (never shrink) so its text fits inside the text box.
export function fitNodeText(doc, n) {
  const st = nodeStyle(doc, n);
  const F = fontFor(doc);
  const font = `${st.italic ? 'italic ' : ''}${st.bold ? 600 : 500} ${st.fontSize}px ${F.text}`;
  const lh = st.fontSize * 1.08;
  const need = () => {
    const box = shapeTextBox(n.shape, n.w, n.h);
    const lines = wrapText(n.text, Math.max(10, box.w), font).length + (doc.settings.showTags && n.tag ? 1 : 0);
    return { box, h: lines * lh + st.fontSize * 0.3 };
  };
  let guard = 0;
  while (guard++ < 200) {
    const { box, h } = need();
    if (box.h >= h) break;
    n.h = Math.ceil(n.h + Math.max(2, (h - box.h) / 2));
  }
  return n;
}

// Arrowheads are sized in user units from the connector's own width, so they keep their size
// when a hovered/selected connector is drawn thicker.
function markerId(type, color, width) {
  return `mk-${type}-${hash(color)}-${String(width).replace('.', '_')}`;
}

// ---- rendering ---------------------------------------------------------------

// Fast path used while dragging: reuse previous routes except for connectors that touch
// the moving shapes (or run near them). A full-quality pass runs when the drag ends.
function computeRoutesFast(doc, { fastNodes, fastEdges, prevRoutes }) {
  const nodes = new Map(doc.nodes.map((n) => [n.id, n]));
  const obstacles = obstaclesFor(doc);
  const moved = [...(fastNodes || [])].map((id) => nodes.get(id)).filter(Boolean).map((n) => ({ x: n.x - 40, y: n.y - 40, w: n.w + 80, h: n.h + 80 }));
  const routes = new Map();
  for (const e of doc.edges) {
    const prev = prevRoutes.get(e.id);
    let dirty = !prev || fastEdges?.has(e.id) || fastNodes?.has(e.from.node) || fastNodes?.has(e.to.node);
    if (!dirty && moved.length) {
      // dirty only if one of its segments actually passes through a moving shape's area
      const P = prev.poly || prev.pts;
      for (let i = 0; i < P.length - 1 && !dirty; i++) {
        const a = P[i], b = P[i + 1];
        const x0 = Math.min(a[0], b[0]), x1 = Math.max(a[0], b[0]), y0 = Math.min(a[1], b[1]), y1 = Math.max(a[1], b[1]);
        dirty = moved.some((m) => x1 >= m.x && x0 <= m.x + m.w && y1 >= m.y && y0 <= m.y + m.h);
      }
    }
    if (!dirty) { routes.set(e.id, prev); continue; }
    const r = routeEdge(e, nodes, obstacles, doc.settings, {});
    r.poly = routePolyline(r);
    routes.set(e.id, r);
  }
  return routes;
}

// Routes depend only on geometry, so cache them across renders that don't move anything
// (selection changes, text and style edits).
let routeCache = { key: null, routes: null };
function geometryKey(doc) {
  let k = (doc.settings.stub ?? 18) + ';';
  for (const n of doc.nodes) k += `${n.id},${n.shape},${n.x},${n.y},${n.w},${n.h};`;
  for (const e of doc.edges) k += `${e.id},${JSON.stringify(e.from)},${JSON.stringify(e.to)},${e.routing},${JSON.stringify(e.waypoints)};`;
  return k;
}

// Line hops are kept on the Map itself (routes.jumps), never on the cached route objects.
function withJumps(doc, routes) {
  routes.jumps = doc.settings.lineJumps === 'none' ? new Map() : findJumps(routes, jumpSize(doc));
  return routes;
}
export const jumpSize = (doc) => Math.max(2, Math.min(12, Number(doc.settings.jumpSize) || 5));

export function computeRoutes(doc, opts = {}) {
  if (opts.prevRoutes && (opts.fastNodes || opts.fastEdges)) return withJumps(doc, computeRoutesFast(doc, opts));
  const key = geometryKey(doc) + '|' + (doc.settings.lineJumps || 'arc') + jumpSize(doc);
  if (routeCache.key === key) return routeCache.routes;
  const routes = withJumps(doc, computeRoutesFull(doc));
  routeCache = { key, routes };
  return routes;
}

function computeRoutesFull(doc) {
  const nodes = new Map(doc.nodes.map((n) => [n.id, n]));
  const obstacles = obstaclesFor(doc);
  const routes = new Map();
  const used = new Map(); // "node|side" -> {in, out, members}
  const note = (node, side, dir, e, end) => {
    if (!node || !side) return;
    const k = node + '|' + side;
    if (!used.has(k)) used.set(k, { in: 0, out: 0, members: [] });
    const u = used.get(k);
    u[dir]++;
    u.members.push({ e, end });
  };
  const auto = (e) => (e.routing || 'orthogonal') === 'orthogonal' && !(e.waypoints || []).length &&
    ((e.from.node && (e.from.side || 'auto') === 'auto') || (e.to.node && (e.to.side || 'auto') === 'auto'));
  for (const e of doc.edges) {
    const r = routeEdge(e, nodes, obstacles, doc.settings, { used });
    routes.set(e.id, r);
    note(e.from.node, r.s1, 'out', e, 'from');
    note(e.to.node, r.s2, 'in', e, 'to');
  }
  // Second pass: re-route auto connectors knowing where every other connector runs.
  const autoEdges = doc.edges.filter(auto);
  if (autoEdges.length && doc.edges.length <= 400) {
    const segsOf = (r) => { const out = []; for (let i = 0; i < r.pts.length - 1; i++) out.push([r.pts[i], r.pts[i + 1]]); return out; };
    for (const e of autoEdges) {
      const others = [];
      for (const [id, r] of routes) if (id !== e.id) others.push(...segsOf(r));
      const prev = routes.get(e.id);
      // release this edge's own port usage while it is re-routed
      const ka = e.from.node + '|' + prev.s1, kb = e.to.node + '|' + prev.s2;
      if (used.get(ka)) used.get(ka).out--;
      if (used.get(kb)) used.get(kb).in--;
      const r = routeEdge(e, nodes, obstacles, doc.settings, { used, others });
      for (const [k, dir, s] of [[ka, 'out', 'from'], [kb, 'in', 'to']]) {
        const u = used.get(k);
        if (u) u.members = u.members.filter((m) => !(m.e === e && m.end === s));
      }
      note(e.from.node, r.s1, 'out', e, 'from');
      note(e.to.node, r.s2, 'in', e, 'to');
      routes.set(e.id, r);
    }
  }
  // Where one side of a shape has both incoming and outgoing connectors on the same
  // point, spread them out along that side so they don't overlap.
  const forced = new Map();
  for (const [k, u] of used) {
    if (!u.in || !u.out) continue;
    if (u.members.some(({ e, end }) => e[end].offset)) continue;
    const [nid, side] = k.split('|');
    const n = nodes.get(nid);
    const len = side === 't' || side === 'b' ? n.w : n.h;
    const step = Math.min(24, len / (u.members.length + 1));
    const axis = side === 't' || side === 'b' ? 0 : 1;
    const other = ({ e, end }) => {
      const r = routes.get(e.id);
      const p = end === 'from' ? r.pts[r.pts.length - 1] : r.pts[0];
      return p[axis];
    };
    const sorted = [...u.members].sort((a, b) => other(a) - other(b));
    sorted.forEach((m, i) => {
      const f = forced.get(m.e.id) || {};
      f[m.end] = { ...m.e[m.end], side, offset: (i - (sorted.length - 1) / 2) * step };
      forced.set(m.e.id, f);
    });
  }
  // Parallel connectors between the same two shapes: offset them so each is visible.
  const groups = new Map();
  for (const e of doc.edges) {
    const r = routes.get(e.id);
    if (!e.from.node || !e.to.node || e.from.node === e.to.node || (e.waypoints || []).length || (e.routing || 'orthogonal') === 'curved') continue;
    if (e.from.offset || e.to.offset || forced.has(e.id)) continue;
    const fwd = e.from.node < e.to.node;
    const key = fwd ? `${e.from.node}|${r.s1}|${e.to.node}|${r.s2}` : `${e.to.node}|${r.s2}|${e.from.node}|${r.s1}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(e);
  }
  for (const list of groups.values()) {
    if (list.length < 2) continue;
    const r0 = routes.get(list[0].id);
    list.forEach((e, i) => {
      const off = (i - (list.length - 1) / 2) * 14;
      const r = routes.get(e.id);
      forced.set(e.id, { from: { ...e.from, side: r.s1 || r0.s1, offset: off }, to: { ...e.to, side: r.s2 || r0.s2, offset: off } });
    });
  }
  for (const [id, f] of forced) {
    const e = doc.edges.find((x) => x.id === id);
    routes.set(id, routeEdge({ ...e, ...f }, nodes, obstacles, doc.settings, {}));
  }
  for (const r of routes.values()) r.poly = routePolyline(r);
  return routes;
}

function contentBounds(doc, routes) {
  const boxes = doc.nodes.map((n) => ({ x: n.x, y: n.y, w: n.w, h: n.h }));
  for (const r of routes.values()) for (const p of r.poly) boxes.push({ x: p[0] - 4, y: p[1] - 4, w: 8, h: 8 });
  return unionBox(boxes);
}

export function labelGeometry(doc, e, route) {
  if (!e.label) return null;
  const F = fontFor(doc);
  const size = e.style?.labelSize || 10.5;
  const font = `500 ${size}px ${F.text}`;
  const lines = String(e.label).split('\n').flatMap((l) => wrapText(l, e.style?.labelWidth || 180, font));
  const w = Math.max(...lines.map((l) => measure(l, font)));
  const lh = size * 1.25;
  const h = lines.length * lh;
  let cx, cy, anchor = 'middle', rotate = false;
  if (e.labelPos !== null && e.labelPos !== undefined) {
    const p = pointAt(route.poly, e.labelPos).pt;
    const off = e.labelOffset || [0, 0];
    cx = p[0] + off[0]; cy = p[1] + off[1];
  } else {
    const m = longestSegmentMid(route.poly);
    if (route.routing === 'curved') { const p = pointAt(route.poly, 0.5).pt; cx = p[0]; cy = p[1]; }
    else if (!m.vertical) { cx = m.pt[0]; cy = m.pt[1] - h / 2 - 3; }
    else if (m.length < 70) { cx = m.pt[0] + 6 + w / 2; cy = m.pt[1]; anchor = 'start'; }
    else { cx = m.pt[0]; cy = m.pt[1]; rotate = doc.settings.verticalLabels === 'rotate' && lines.length === 1; }
  }
  if (e.labelRotate !== undefined && e.labelRotate !== null) rotate = !!e.labelRotate;
  if (rotate) {
    return { x: cx - h / 2, y: cy - w / 2, w: h, h: w, cx, cy, lines, lh, size, font, anchor: 'middle', rotate: true, tw: w, th: h };
  }
  if (e.labelAlign === 'start' && anchor === 'middle') { cx += w / 2; anchor = 'start'; }
  if (e.labelAlign === 'end' && anchor === 'middle') { cx -= w / 2; anchor = 'end'; }
  const x = cx - w / 2, y = cy - h / 2;
  return { x, y, w, h, cx, cy, lines, lh, size, font, anchor };
}

export function renderEdge(doc, e, route, opts, jumps = null) {
  const st = edgeStyle(doc, e);
  const d = pathD(route, doc.settings.cornerRadius, jumps, jumpSize(doc));
  const selected = opts.selEdges?.has(e.id);
  let inner = '';
  if (opts.editor || opts.interactive) inner += tag('path', { class: 'edge-hit', d, fill: 'none', stroke: 'transparent', 'stroke-width': 14 });
  inner += tag('path', {
    class: 'edge-line', d, fill: 'none', stroke: st.color, 'stroke-width': st.width,
    'stroke-dasharray': st.dash || undefined, 'stroke-linejoin': 'round', 'stroke-linecap': st.dash ? 'butt' : 'round',
    'marker-end': ARROWS[st.end] ? `url(#${markerId(st.end, st.color, st.width)})` : undefined,
    'marker-start': ARROWS[st.start] ? `url(#${markerId(st.start, st.color, st.width)})` : undefined,
  });
  const a = { class: 'edge' + (selected ? ' sel' : ''), 'data-edge': e.id };
  if (opts.interactive) { a.tabindex = 0; a.role = 'button'; a['aria-label'] = 'Connector' + (e.label ? ': ' + e.label : ''); }
  return tag('g', a, inner);
}

export function renderLabel(doc, e, route, opts) {
  const g = labelGeometry(doc, e, route);
  if (!g) return '';
  const st = edgeStyle(doc, e);
  const F = fontFor(doc);
  if (g.rotate) {
    let rin = tag('rect', { class: 'label-bg', x: round(-g.tw / 2 - 3), y: round(-g.th / 2 - 1), width: round(g.tw + 6), height: round(g.th + 2), rx: 2, fill: e.style?.labelBg || 'var(--label-bg)' });
    g.lines.forEach((l, i) => {
      rin += tag('text', { x: 0, y: round(-g.th / 2 + g.lh * i + g.size * 0.95), 'text-anchor': 'middle', 'font-family': F.text, 'font-size': g.size, 'font-weight': 500, fill: st.labelColor }, esc(l));
    });
    return tag('g', { class: 'edge-label', 'data-label': e.id, transform: `translate(${round(g.cx)} ${round(g.cy)}) rotate(-90)`, ...(opts.editor ? { style: 'cursor:move' } : {}) }, rin);
  }
  let inner = tag('rect', { class: 'label-bg', x: round(g.x - 3), y: round(g.y - 1), width: round(g.w + 6), height: round(g.h + 2), rx: 2, fill: e.style?.labelBg || 'var(--label-bg)' });
  const tx = g.anchor === 'start' ? g.x : g.anchor === 'end' ? g.x + g.w : g.cx;
  g.lines.forEach((l, i) => {
    inner += tag('text', {
      x: round(tx), y: round(g.y + g.lh * i + g.size * 0.95), 'text-anchor': g.anchor, 'font-family': F.text,
      'font-size': g.size, 'font-weight': 500, fill: st.labelColor,
    }, esc(l));
  });
  return tag('g', { class: 'edge-label', 'data-label': e.id, ...(opts.editor ? { style: 'cursor:move' } : {}) }, inner);
}

const round = (v) => Math.round(v * 100) / 100;

export function renderNode(doc, n, opts = {}) {
  const def = shapeDef(n.shape);
  const st = nodeStyle(doc, n);
  const F = fontFor(doc);
  const parts = def.draw(n.w, n.h);
  let inner = '';
  parts.forEach(([t, a], i) => {
    const main = i === 0 || def.layered;
    const base = main
      ? { class: 'shape', fill: st.fill, stroke: st.stroke, 'stroke-width': st.strokeWidth, 'stroke-dasharray': st.dash || undefined }
      : { class: 'deco', fill: 'none', stroke: st.stroke === 'none' ? 'var(--node-stroke)' : st.stroke, 'stroke-width': Math.max(1, st.strokeWidth * 0.85) };
    inner += tag(t, { ...base, ...a });
  });

  // text
  const box = shapeTextBox(n.shape, n.w, n.h);
  const weight = st.bold ? 600 : 500;
  const font = `${st.italic ? 'italic ' : ''}${weight} ${st.fontSize}px ${F.text}`;
  const lines = wrapText(n.text, Math.max(10, box.w), font);
  const showTag = doc.settings.showTags && n.tag;
  const lh = st.fontSize * 1.08;
  const total = lines.length + (showTag ? 1 : 0);
  let ty;
  if (st.valign === 'top') ty = box.y + st.fontSize * 0.95;
  else if (st.valign === 'bottom') ty = box.y + box.h - (total - 1) * lh - st.fontSize * 0.25;
  else ty = box.y + box.h / 2 - ((total - 1) * lh) / 2 + st.fontSize * 0.35;
  const tx = st.align === 'left' ? box.x : st.align === 'right' ? box.x + box.w : box.x + box.w / 2;
  const anchor = st.align === 'left' ? 'start' : st.align === 'right' ? 'end' : 'middle';
  let text = '';
  if (showTag) {
    text += tag('text', { class: 'tag', x: round(tx), y: round(ty), 'text-anchor': anchor, 'font-family': F.mono, 'font-size': Math.max(8, st.fontSize - 2), 'font-weight': 600, fill: n.style?.tagColor || 'var(--tag)', 'letter-spacing': '0.03em' }, esc(n.tag));
    ty += lh;
  }
  for (const l of lines) {
    text += tag('text', {
      x: round(tx), y: round(ty), 'text-anchor': anchor, 'font-family': F.text, 'font-size': st.fontSize,
      'font-weight': weight, 'font-style': st.italic ? 'italic' : undefined, fill: st.textColor,
    }, esc(l));
    ty += lh;
  }
  inner += tag('g', { class: 'label', 'pointer-events': 'none' }, text);

  // badges (a circle, or a pill for numbered ones like C1), stepping left from the corner
  const badges = nodeBadges(doc, n);
  if (badges.length) {
    let [bx] = def.badge ? def.badge(n.w, n.h) : [n.w - 2, 2];
    const by = (def.badge ? def.badge(n.w, n.h) : [0, 2])[1];
    const bfont = `700 10px ${F.mono}`;
    badges.forEach(({ kind: b, label }) => {
      const w = Math.max(16, measure(label, bfont) + 8);
      const cx = bx - (w - 16) / 2;
      inner += tag('g', { class: 'badge', 'data-badge': b.id },
        (w === 16 ? tag('circle', { cx: round(cx), cy: round(by), r: 8, fill: b.color || 'var(--risk)' })
          : tag('rect', { x: round(cx - w / 2), y: round(by - 8), width: round(w), height: 16, rx: 8, fill: b.color || 'var(--risk)' })) +
        tag('text', { x: round(cx), y: round(by + 3.6), 'text-anchor': 'middle', 'font-size': 10, 'font-weight': 700, 'font-family': F.mono, fill: b.textColor || 'var(--risk-ink)' }, esc(label)));
      bx -= w + 3;
    });
  }

  const cls = ['node', `shape-${n.shape}`];
  if (opts.selNodes?.has(n.id)) cls.push('sel');
  if (def.container) cls.push('container');
  const a = {
    class: cls.join(' '), 'data-node': n.id, transform: `translate(${round(n.x)} ${round(n.y)})`,
    opacity: st.opacity !== 1 ? st.opacity : undefined,
  };
  if (st.shadow) a.filter = 'url(#fc-shadow)';
  if (opts.interactive) {
    a.tabindex = 0; a.role = 'button';
    a['aria-label'] = (n.tag ? n.tag + ' ' : '') + (n.text || def.name);
  }
  return tag('g', a, inner);
}

function renderLanes(doc, extent, crossStart = 0) {
  const L = doc.lanes;
  if (!L.items.length) return '';
  const F = fontFor(doc);
  const H = L.header || 40;
  let out = '';
  const rects = laneRects(doc, extent);
  rects.forEach((r, i) => {
    const fill = r.lane.fill || (i % 2 ? 'var(--lane-alt)' : 'var(--lane)');
    const horiz = L.orientation === 'horizontal';
    let inner = tag('rect', { class: 'lane-body', x: r.x, y: r.y, width: r.w, height: r.h, fill });
    const hx = r.x, hy = r.y, hw = horiz ? H : r.w, hh = horiz ? r.h : H;
    inner += tag('rect', { class: 'lane-head', 'data-lane-head': r.lane.id, x: hx, y: hy, width: hw, height: hh, fill: r.lane.headFill || 'var(--lane-head)' });
    const title = esc(String(r.lane.title || '').toUpperCase());
    const tcol = r.lane.textColor || 'var(--muted)';
    if (horiz) {
      const cx = hx + hw / 2, cy = hy + hh / 2;
      inner += tag('text', { x: 0, y: 0, 'text-anchor': 'middle', transform: `translate(${cx + 4} ${cy}) rotate(-90)`, 'font-family': F.text, 'font-size': 13, 'font-weight': 600, 'letter-spacing': '0.08em', fill: tcol, 'pointer-events': 'none' }, title);
      if (r.lane.subtitle) inner += tag('text', { x: 0, y: 0, 'text-anchor': 'middle', transform: `translate(${hx + hw + 12} ${cy}) rotate(-90)`, 'font-size': 10, fill: 'var(--faint)', 'font-family': F.ui, 'pointer-events': 'none' }, esc(r.lane.subtitle));
    } else {
      inner += tag('text', { x: hx + hw / 2, y: hy + H / 2 + 5, 'text-anchor': 'middle', 'font-family': F.text, 'font-size': 13, 'font-weight': 600, 'letter-spacing': '0.08em', fill: tcol, 'pointer-events': 'none' }, title);
      if (r.lane.subtitle) inner += tag('text', { x: hx + hw / 2, y: hy + H + 12, 'text-anchor': 'middle', 'font-size': 10, fill: 'var(--faint)', 'font-family': F.ui, 'pointer-events': 'none' }, esc(r.lane.subtitle));
    }
    if (i > 0) {
      inner += horiz
        ? tag('line', { x1: r.x, y1: r.y, x2: r.x + r.w, y2: r.y, stroke: 'var(--line)' })
        : tag('line', { x1: r.x, y1: r.y, x2: r.x, y2: r.y + r.h, stroke: 'var(--line)' });
    }
    out += tag('g', { class: 'lane', 'data-lane': r.lane.id }, inner);
  });
  return out;
}

function renderPhases(doc, crossLen, stretchTo = 0) {
  if (!doc.phases.items.length) return '';
  const F = fontFor(doc);
  const H = doc.phases.header || 34;
  const { along, start } = phaseAxis(doc);
  const PH = doc.phases;
  let out = '';
  const list = phaseStarts(doc);
  // optional heading for the phase column (e.g. "BD"), in the corner beside the lane headers
  if (PH.label && start > 0) {
    const corner = along === 'y' ? { x: -H, y: 0, width: H, height: start } : { x: 0, y: -H, width: start, height: H };
    out += tag('g', { class: 'phase-label' },
      tag('rect', { ...corner, fill: PH.headFill || 'var(--lane-head)' }) +
      tag('text', { x: corner.x + corner.width / 2, y: corner.y + corner.height / 2 + 4.5, 'text-anchor': 'middle', 'font-family': F.text, 'font-size': 12, 'font-weight': 600, 'letter-spacing': '0.07em', fill: 'var(--muted)' }, esc(String(PH.label).toUpperCase())));
  }
  list.forEach(([p0, pos], i) => {
    // the last phase stretches to the end of the lanes so there is no unlabelled strip
    const p = i === list.length - 1 && stretchTo > pos + p0.size ? { ...p0, size: stretchTo - pos } : p0;
    let inner = '';
    const fill = p.headFill || PH.headFill || (i % 2 ? 'var(--lane-alt)' : 'var(--lane-head)');
    const title = esc(p.title || '');
    // title placement: centred in the band (default) or at its start; rotated (default) or horizontal
    const mid = PH.titleAt === 'start' ? pos + Math.min(p.size / 2, 46) : pos + p.size / 2;
    const tFont = PH.titleFont === 'mono' ? F.mono : F.text;
    const tColor = p.textColor || (PH.titleFont === 'mono' ? 'var(--ink)' : 'var(--muted)');
    if (along === 'y') {
      if (p.fill) inner += tag('rect', { class: 'phase-body', x: 0, y: pos, width: crossLen, height: p.size, fill: p.fill, 'pointer-events': 'none' });
      inner += tag('rect', { class: 'phase-head', 'data-phase-head': p.id, x: -H, y: pos, width: H, height: p.size, fill });
      inner += PH.titleStyle === 'horizontal'
        ? tag('text', { x: -H / 2, y: mid + 4, 'text-anchor': 'middle', 'font-family': tFont, 'font-size': 11, 'font-weight': 600, fill: tColor, 'pointer-events': 'none' }, title)
        : tag('text', { x: 0, y: 0, 'text-anchor': 'middle', transform: `translate(${-H / 2 + 4} ${mid}) rotate(-90)`, 'font-family': tFont, 'font-size': 11.5, 'font-weight': 600, 'letter-spacing': '0.04em', fill: tColor, 'pointer-events': 'none' }, title);
      inner += tag('line', { x1: -H, y1: pos + p.size, x2: crossLen, y2: pos + p.size, stroke: 'var(--line)', 'stroke-dasharray': PH.lineDash || '6 4', 'pointer-events': 'none' });
      if (i === 0) inner += tag('line', { x1: -H, y1: pos, x2: crossLen, y2: pos, stroke: 'var(--line)', 'pointer-events': 'none' });
    } else {
      if (p.fill) inner += tag('rect', { class: 'phase-body', x: pos, y: 0, width: p.size, height: crossLen, fill: p.fill, 'pointer-events': 'none' });
      inner += tag('rect', { class: 'phase-head', 'data-phase-head': p.id, x: pos, y: -H, width: p.size, height: H, fill });
      inner += tag('text', { x: mid, y: -H / 2 + 4, 'text-anchor': 'middle', 'font-family': tFont, 'font-size': 11.5, 'font-weight': 600, 'letter-spacing': '0.04em', fill: tColor, 'pointer-events': 'none' }, title);
      inner += tag('line', { x1: pos + p.size, y1: -H, x2: pos + p.size, y2: crossLen, stroke: 'var(--line)', 'stroke-dasharray': PH.lineDash || '6 4', 'pointer-events': 'none' });
      if (i === 0) inner += tag('line', { x1: pos, y1: -H, x2: pos, y2: crossLen, stroke: 'var(--line)', 'pointer-events': 'none' });
    }
    out += tag('g', { class: 'phase', 'data-phase': p.id }, inner);
  });
  return out;
}

export function phasesEnd(doc) {
  const st = phaseStarts(doc);
  if (!st.length) return 0;
  const [p, pos] = st[st.length - 1];
  return pos + p.size;
}

export function lanesSize(doc) {
  return doc.lanes.items.reduce((s, l) => s + l.size, 0);
}

// Main entry. opts: {editor, interactive, selNodes, selEdges, minExtent}
export function renderDiagram(doc, opts = {}) {
  const routes = computeRoutes(doc, opts);
  const cb = contentBounds(doc, routes);
  const horiz = doc.lanes.orientation === 'horizontal';
  const hasLanes = doc.lanes.items.length > 0;
  let extent = 0;
  if (hasLanes) {
    const far = cb ? (horiz ? cb.x + cb.w : cb.y + cb.h) : 0;
    extent = Math.max(far + (opts.extentPad ?? 40), (doc.lanes.header || 40) + 200, opts.minExtent || 0, doc.lanes.minLength || 0, phasesEnd(doc));
  }
  const hasPhases = doc.phases.items.length > 0;
  const crossLen = hasLanes ? lanesSize(doc) : Math.max(300, cb ? (horiz ? cb.y + cb.h : cb.x + cb.w) + 40 : 0);

  // markers
  const needed = new Map();
  for (const e of doc.edges) {
    const st = edgeStyle(doc, e);
    for (const t of [st.start, st.end]) if (ARROWS[t]) needed.set(markerId(t, st.color, st.width), [t, st.color, st.width]);
  }
  let defs = '';
  for (const [id, [t, color, width]] of needed) {
    const A = ARROWS[t];
    const size = round(7 * (Number(width) || 1.6));
    defs += tag('marker', { id, viewBox: '-1 -1 12 12', refX: 9, refY: 5, markerWidth: size, markerHeight: size, orient: 'auto-start-reverse', markerUnits: 'userSpaceOnUse' },
      tag('path', { d: A.d, fill: A.fill ? color : 'none', stroke: A.fill ? 'none' : color, 'stroke-width': 1.6, 'stroke-linejoin': 'round' }));
  }
  defs += '<filter id="fc-shadow" x="-10%" y="-10%" width="130%" height="140%"><feDropShadow dx="0" dy="2" stdDeviation="2.5" flood-opacity="0.18"/></filter>';

  const containers = doc.nodes.filter((n) => shapeDef(n.shape).container);
  const others = doc.nodes.filter((n) => !shapeDef(n.shape).container);
  let body = `<defs>${defs}</defs>`;
  body += tag('g', { class: 'lanes' }, renderLanes(doc, extent));
  const alongEnd = hasLanes ? extent : (cb ? (horiz ? cb.x + cb.w : cb.y + cb.h) + 40 : 0);
  body += tag('g', { class: 'phases' }, renderPhases(doc, crossLen, alongEnd));
  body += tag('g', { class: 'containers' }, containers.map((n) => renderNode(doc, n, opts)).join(''));
  body += tag('g', { class: 'edges' }, doc.edges.map((e) => renderEdge(doc, e, routes.get(e.id), opts, routes.jumps?.get(e.id))).join(''));
  body += tag('g', { class: 'edge-labels' }, doc.edges.map((e) => renderLabel(doc, e, routes.get(e.id), opts)).join(''));
  body += tag('g', { class: 'nodes' }, others.map((n) => renderNode(doc, n, opts)).join(''));

  // overall bounds
  const boxes = [];
  if (cb) boxes.push(cb);
  if (hasLanes) {
    const size = lanesSize(doc);
    boxes.push(horiz ? { x: 0, y: 0, w: extent, h: size } : { x: 0, y: 0, w: size, h: extent });
  }
  for (const e of doc.edges) {
    const g = labelGeometry(doc, e, routes.get(e.id));
    if (g) boxes.push({ x: g.x - 4, y: g.y - 2, w: g.w + 8, h: g.h + 4 });
  }
  if (hasPhases) {
    const H = doc.phases.header || 34;
    const start = phaseStarts(doc)[0][1], end = Math.max(phasesEnd(doc), alongEnd);
    boxes.push(horiz ? { x: start, y: -H, w: end - start, h: H + crossLen } : { x: -H, y: start, w: H + crossLen, h: end - start });
    if (doc.phases.label) boxes.push(horiz ? { x: 0, y: -H, w: start, h: H } : { x: -H, y: 0, w: H, h: start });
  }
  const bounds = unionBox(boxes) || { x: 0, y: 0, w: 400, h: 300 };
  return { body, bounds, routes, extent, crossLen };
}

// Full standalone <svg> for export.
export function renderSVG(doc, { padding = 16, resolve = false, dark = false, interactive = false, background = true, header = false } = {}) {
  const { body, bounds } = renderDiagram(doc, { interactive });
  const hasLanes = doc.lanes.items.length > 0;
  const pad = hasLanes && !header ? 0 : padding;
  const F = fontFor(doc);
  // optional title / subtitle / legend block above the diagram (image exports)
  let head = '', headH = 0;
  if (header) {
    const parts = [];
    let y0 = 0;
    const maxW = Math.max(bounds.w, 480);
    if (doc.title) { parts.push(tag('text', { x: 0, y: 22, 'font-family': F.text, 'font-size': 22, 'font-weight': 600, fill: 'var(--ink)' }, esc(doc.title))); y0 = 32; }
    if (doc.subtitle) { parts.push(tag('text', { x: 0, y: y0 + 13, 'font-family': F.ui, 'font-size': 12.5, fill: 'var(--muted)' }, esc(doc.subtitle))); y0 += 22; }
    const items = doc.settings.showLegend !== false ? legendItems(doc) : [];
    if (items.length) {
      let lx = 0, ly = y0 + 6;
      for (const it of items) {
        const iw = 34 + measure(it.label, `400 12px ${F.ui}`) + 14;
        if (lx + iw > maxW && lx > 0) { lx = 0; ly += 22; }
        parts.push(legendSwatch(doc, it).replace('<svg ', `<svg x="${lx}" y="${ly}" `));
        parts.push(tag('text', { x: lx + 32, y: ly + 12, 'font-family': F.ui, 'font-size': 12, fill: 'var(--muted)' }, esc(it.label)));
        lx += iw;
      }
      y0 = ly + 22;
    }
    headH = y0 ? y0 + 14 : 0;
    head = `<g transform="translate(${bounds.x} ${bounds.y - headH})">${parts.join('')}</g>`;
  }
  const x = Math.floor(bounds.x - pad), y = Math.floor(bounds.y - pad - headH);
  const w = Math.ceil(bounds.w + pad * 2), h = Math.ceil(bounds.h + pad * 2 + headH);
  const bg = background ? tag('rect', { x, y, width: w, height: h, fill: 'var(--panel)' }) : '';
  let svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${x} ${y} ${w} ${h}" width="${w}" height="${h}" role="img" aria-label="${esc(doc.title)}">${bg}${head}${body}</svg>`;
  if (resolve) svg = resolveVars(svg, dark);
  return { svg, w, h };
}

// ---- legend -------------------------------------------------------------------

export function legendItems(doc) {
  const items = [];
  const hidden = new Set(doc.legendHidden || []);
  // shapes listed in legendAlways appear even when unused (a fixed, house-style legend)
  const used = new Set([...doc.nodes.map((n) => n.shape), ...(doc.legendAlways || [])]);
  for (const k of SHAPE_ORDER) {
    if (!used.has(k) || SHAPES[k].noLegend || hidden.has('shape:' + k)) continue;
    items.push({ type: 'shape', key: k, label: shapeLabel(doc, k) });
  }
  const dashed = doc.nodes.some((n) => n.style?.dash && !shapeDef(n.shape).container);
  if (dashed && doc.dashedNodeLabel && !hidden.has('dashed')) items.push({ type: 'dashed', key: 'dashed', label: doc.dashedNodeLabel });
  const kinds = new Set(doc.edges.map((e) => e.kind));
  for (const k of doc.edgeKinds) if (kinds.has(k.id) && !hidden.has('edge:' + k.id)) items.push({ type: 'edge', key: k.id, label: k.name, kind: k });
  const badges = new Set(doc.nodes.flatMap((n) => nodeBadges(doc, n).map((b) => b.kind.id)));
  for (const b of doc.badgeKinds) if (badges.has(b.id) && !hidden.has('badge:' + b.id)) items.push({ type: 'badge', key: b.id, label: b.name, badge: b });
  // optional explicit order, e.g. ["shape:process", "shape:decision", ...]
  if (Array.isArray(doc.legendOrder) && doc.legendOrder.length) {
    const rank = (it) => { const i = doc.legendOrder.indexOf(it.type + ':' + it.key); return i < 0 ? 1e3 + items.indexOf(it) : i; };
    items.sort((a, b) => rank(a) - rank(b));
  }
  return items;
}

export function legendSwatch(doc, it) {
  const F = fontFor(doc);
  if (it.type === 'shape') {
    const def = SHAPES[it.key];
    const w = 26, h = 16;
    const parts = def.draw(w - 2, h - 2);
    const inner = parts.map(([t, a], i) => tag(t, { fill: i === 0 || def.layered ? def.fill || 'var(--node)' : 'none', stroke: 'var(--node-stroke)', 'stroke-width': 1, ...a })).join('');
    return `<svg width="${w}" height="${h}" aria-hidden="true"><g transform="translate(1 1)">${inner}</g></svg>`;
  }
  if (it.type === 'dashed') return `<svg width="26" height="16" aria-hidden="true"><rect x="1" y="2" width="24" height="12" rx="2" fill="var(--node)" stroke="var(--node-stroke)" stroke-dasharray="3 2"/></svg>`;
  if (it.type === 'edge') {
    const k = it.kind;
    return `<svg width="28" height="10" aria-hidden="true"><line x1="1" y1="5" x2="27" y2="5" stroke="${esc(k.color)}" stroke-width="${esc(k.width || 1.6)}"${k.dash ? ` stroke-dasharray="${esc(k.dash)}"` : ''}/></svg>`;
  }
  const b = it.badge;
  return `<svg width="16" height="16" aria-hidden="true"><circle cx="8" cy="8" r="7" fill="${esc(b.color)}"/><text x="8" y="11.5" text-anchor="middle" font-size="10" font-weight="700" fill="${esc(b.textColor || 'var(--risk-ink)')}" font-family="${esc(F.mono)}">${esc(b.text)}</text></svg>`;
}

export function legendHTML(doc) {
  return legendItems(doc).map((it) => `<span data-legend="${esc(it.type + ':' + it.key)}">${legendSwatch(doc, it)}${esc(it.label)}</span>`).join('');
}
