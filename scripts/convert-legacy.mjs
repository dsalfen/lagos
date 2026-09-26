// Converts hand-coded "interactive flowchart" HTML pages (one page per process, with the chart data
// in an inline script: N = shapes, E = connectors, LANES, optional BD column) into editor documents
// (.json) that open with File → Open.
//
//   node scripts/convert-legacy.mjs <input.html> [output.json]
//
// The original page is run in a headless browser with all network access blocked; its chart data
// (shapes, connectors, lanes, business-day rows) is captured just before it draws, then mapped onto
// the editor's model. Connector ends are matched to the exact side/offset they touch, and bends are
// kept, so the layout matches the original.
import { chromium } from '@playwright/test';
import { readFileSync, writeFileSync } from 'node:fs';
import { basename } from 'node:path';
import { portPoint } from '../src/geometry.js';
import { normalizeDoc } from '../src/model.js';

const [input, output] = process.argv.slice(2);
if (!input) { console.error('usage: node scripts/convert-legacy.mjs <input.html> [output.json]'); process.exit(1); }
const src = readFileSync(input, 'utf8');

// ---- capture the chart data from the original page --------------------------------------------
const HOOK = 'document.getElementById("chart").appendChild(svg);';
if (!src.includes(HOOK)) throw new Error('This does not look like one of the interactive flowchart pages.');
const capture = `window.__LEGACY = JSON.parse(JSON.stringify({
  N, E, LANES, LW, TOP, ROW, BW, BH, DH, W, H,
  LX: typeof LX === 'undefined' ? 0 : LX,
  BD: typeof BD === 'undefined' ? null : BD,
  NROWS: typeof NROWS === 'undefined' ? null : NROWS,
  sub: typeof sub === 'undefined' ? '' : sub.textContent,
}));`;
const html = src.replace(HOOK, capture + HOOK);
const initial = (src.match(/select\("([^"]+)"\);\s*\}\)\(\);/) || [])[1] || null;
const stepPrefix = (src.match(/n\.id\.startsWith\("([^"]+)"\)/) || [])[1] || '';
const hint = (src.match(/<p class="hint">([^<]+)<\/p>/) || [])[1] || '';

const browser = await chromium.launch();
const ctx = await browser.newContext();
await ctx.route('**/*', (r) => r.abort()); // never touch the network
const page = await ctx.newPage();
await page.setContent(html, { waitUntil: 'domcontentloaded' });
await page.waitForFunction(() => window.__LEGACY);
const L = await page.evaluate(() => window.__LEGACY);
const meta = await page.evaluate(() => ({
  title: document.querySelector('h1')?.textContent.trim() || document.title,
  subtitle: document.querySelector('.sub')?.textContent.trim() || '',
  notes: [...document.querySelectorAll('.notes p')].map((p) => p.textContent.trim()),
  legend: [...document.querySelectorAll('.legend span')].map((s) => s.textContent.trim()),
}));
await browser.close();

// ---- map onto the editor model -----------------------------------------------------------------
const SHAPE = { proc: 'process', dec: 'decision', doc: 'document', xfer: 'data', man: 'manualInput', data: 'database', term: 'terminator', tbc: 'process' };
const dx = -L.LX; // lanes start at x = 0; a business-day column (if any) becomes the phase strip at negative x
const isStep = (id) => (stepPrefix ? id.startsWith(stepPrefix) : /^[A-Z]+-\d+/.test(id));

const nodes = L.N.map((n) => {
  const w = n.s === 'term' ? L.BW - 28 : L.BW;
  const h = n.s === 'dec' ? L.DH : n.s === 'term' ? 40 : L.BH;
  const style = {};
  if (n.s === 'tbc') style.dash = '5 4';
  if (n.s === 'dec') style.fontSize = 11.5;
  const fields = { sp: n.sp, what: n.what, sys: n.sys, out: n.out };
  if (n.prp) fields.prp = n.prp;
  return {
    id: n.id, shape: SHAPE[n.s] || 'process', x: n.x + dx - w / 2, y: n.y - h / 2, w, h, text: n.t,
    tag: isStep(n.id) ? n.id : '', style, badges: n.prp && isStep(n.id) ? ['risk'] : [], laneLabel: n.lane, fields,
  };
});

// Which shape, side and offset does a point sit on?
function attach(pt) {
  let best = null;
  for (const n of nodes) {
    const cx = n.x + n.w / 2, cy = n.y + n.h / 2;
    for (const side of ['t', 'b', 'l', 'r']) {
      const offset = side === 't' || side === 'b' ? pt[0] - cx : pt[1] - cy;
      if (Math.abs(offset) > (side === 't' || side === 'b' ? n.w : n.h) / 2) continue;
      const q = portPoint(n, side, offset);
      // on the outline (exact), or on the bounding box edge (originals drew some shapes' ports there)
      const bx = side === 'l' ? n.x : side === 'r' ? n.x + n.w : pt[0];
      const by = side === 't' ? n.y : side === 'b' ? n.y + n.h : pt[1];
      const d = Math.min(Math.hypot(q[0] - pt[0], q[1] - pt[1]), Math.hypot(bx - pt[0], by - pt[1]) + 0.5);
      const score = d + Math.abs(offset) * 0.01; // at a corner, prefer the side the point is centred on
      if (d < 3 && (!best || score < best.score)) best = { score, node: n.id, side, offset: Math.round(offset * 10) / 10 };
    }
  }
  return best ? { node: best.node, side: best.side, offset: best.offset } : { x: pt[0], y: pt[1] };
}

const segLen = (a, b) => Math.hypot(b[0] - a[0], b[1] - a[1]);
const hasCollinear = (pts) => pts.some((p, i) => i > 0 && i < pts.length - 1 &&
  ((Math.abs(pts[i - 1][0] - p[0]) < 0.5 && Math.abs(p[0] - pts[i + 1][0]) < 0.5) || (Math.abs(pts[i - 1][1] - p[1]) < 0.5 && Math.abs(p[1] - pts[i + 1][1]) < 0.5)));

let k = 0;
const edges = L.E.map((ed) => {
  const pts = ed.pts.map(([x, y]) => [x + dx, y]);
  const e = {
    id: 'e' + ++k, from: attach(pts[0]), to: attach(pts[pts.length - 1]), waypoints: pts.slice(1, -1),
    kind: ed.dash ? 'feed' : 'flow', routing: 'orthogonal', label: '', labelPos: null, labelOffset: null, style: {},
  };
  if (ed.stub) {
    // short "stub" arrow with its text beyond the loose end
    const left = ed.dir === 'l';
    Object.assign(e, { label: ed.stub, labelPos: 1, labelOffset: [left ? -5 : 5, 0], labelAlign: left ? 'end' : 'start' });
    e.style.labelColor = 'var(--muted)';
    e.style.labelBg = 'none';
    return e;
  }
  if (!ed.label) return e;
  e.label = ed.label;
  const f = ed.lt ?? 0.5;
  if (f === 0.5 && !hasCollinear(pts)) return e; // the editor's automatic placement is the same rule
  // explicit position: same point on the longest segment as the original
  let best = 0, bi = 0;
  for (let i = 0; i < pts.length - 1; i++) { const l = segLen(pts[i], pts[i + 1]); if (l > best) { best = l; bi = i; } }
  const total = pts.slice(1).reduce((s, p, i) => s + segLen(pts[i], p), 0);
  const before = pts.slice(1, bi + 1).reduce((s, p, i) => s + segLen(pts[i], p), 0);
  e.labelPos = Math.round(((before + best * f) / total) * 1000) / 1000;
  const vert = Math.abs(pts[bi + 1][0] - pts[bi][0]) < 1;
  if (vert && best < 70) { e.labelOffset = [6, 0]; e.labelAlign = 'start'; }
  else if (vert) { e.labelOffset = [0, 0]; e.labelRotate = true; }
  else e.labelOffset = [0, -8];
  return e;
});

// lanes (the last one takes the remaining width)
const lanes = L.LANES.map((title, i) => ({
  id: 'L' + (i + 1), title, subtitle: i === 0 ? L.sub : '',
  size: i === L.LANES.length - 1 ? L.W - L.LX - i * L.LW : L.LW,
}));

// business-day rows become phases in a strip left of the lanes
let phases = { header: 34, items: [] };
if (L.BD) {
  const rows = Object.keys(L.BD).map(Number).sort((a, b) => a - b);
  const bnd = (r) => L.TOP + r * L.ROW;
  const end = L.TOP + (L.NROWS || rows[rows.length - 1] + 3) * L.ROW;
  phases = {
    header: L.LX, label: 'BD', titleStyle: 'horizontal', titleAt: 'start', titleFont: 'mono', headFill: 'var(--lane-head)', lineDash: '2 4',
    items: rows.map((r, i) => {
      const start = i === 0 ? 40 : bnd(r);
      const stop = i + 1 < rows.length ? bnd(rows[i + 1]) : end;
      return { id: 'P' + (i + 1), title: L.BD[r], size: stop - start };
    }),
  };
}

// the badge swatch's own "R" is part of the span's text
const riskName = (meta.legend.find((t) => /risk/i.test(t)) || 'Risk point').replace(/^R(?=[A-Z])/, '');
// the originals show a fixed legend: keep its entries and order
const LEGEND_KEYS = { Process: 'shape:process', Decision: 'shape:decision', Report: 'shape:document', 'Data transfer': 'shape:data',
  'Manual input': 'shape:manualInput', 'Start / end': 'shape:terminator', 'Data store': 'shape:database', Flow: 'edge:flow',
  'Information feed': 'edge:feed' };
const legendOrder = meta.legend.map((t) => (/risk/i.test(t) ? 'badge:risk' : LEGEND_KEYS[t])).filter(Boolean);
const legendAlways = legendOrder.filter((k) => k.startsWith('shape:')).map((k) => k.slice(6));
const doc = normalizeDoc({
  title: meta.title, subtitle: meta.subtitle, notes: meta.notes.join('\n'),
  settings: { font: 'plex', verticalLabels: 'rotate', detailHint: hint, initialSelection: initial, noTagLabel: 'Connector' },
  lanes: { orientation: 'vertical', header: 40, items: lanes },
  phases,
  fields: [
    { key: 'sp', label: 'Sub-process' },
    { key: 'what', label: 'What happens' },
    { key: 'sys', label: 'System / data' },
    { key: 'out', label: 'Output' },
    { key: 'prp', label: 'Risk point to probe', highlight: true },
  ],
  edgeKinds: [
    { id: 'flow', name: 'Flow', color: 'var(--flow)', dash: '', width: 1.6 },
    { id: 'feed', name: 'Information feed', color: 'var(--feed)', dash: '5 4', width: 1.6, labelColor: 'var(--muted)' },
  ],
  badgeKinds: [{ id: 'risk', text: 'R', color: 'var(--risk)', name: riskName }],
  shapeLabels: { document: 'Report', data: 'Data transfer', terminator: 'Start / end' },
  dashedNodeLabel: 'TBC',
  legendHidden: ['dashed', ...(legendOrder.includes('shape:database') ? [] : ['shape:database'])],
  legendOrder, legendAlways,
  nodes, edges,
});

const unattached = doc.edges.filter((e) => !e.from.node).length;
const out = output || basename(input).replace(/\.html?$/i, '') + '.json';
writeFileSync(out, JSON.stringify(doc, null, 2));
console.log(`${out}: ${doc.nodes.length} shapes, ${doc.edges.length} connectors, ${doc.lanes.items.length} lanes, ${doc.phases.items.length} phases` +
  (unattached ? ` — ${unattached} connector(s) start at a loose point` : ''));
