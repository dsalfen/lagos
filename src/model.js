import { shapeDef } from './shapes.js';

export const FORMAT_VERSION = 1;

let counter = 0;
export function uid(prefix = 'n') {
  counter = (counter + 1) % 1e6;
  return prefix + Date.now().toString(36).slice(-5) + Math.random().toString(36).slice(2, 6) + counter.toString(36);
}

export const clone = (o) => JSON.parse(JSON.stringify(o));

export function defaultSettings() {
  return {
    grid: 8,
    snap: true,
    showGrid: true,
    showTags: true,
    cornerRadius: 0,
    font: 'plex',
    fontSize: 12.5,
    stub: 18,
    theme: 'auto',
    showLegend: true,
    detailHint: 'Select any shape to see its details.',
    highlightColor: 'var(--risk)',
    verticalLabels: 'horizontal',
    lineJumps: 'arc',
    jumpSize: 5,
  };
}

export function defaultEdgeKinds() {
  return [
    { id: 'flow', name: 'Flow', color: 'var(--flow)', dash: '', width: 1.6 },
    { id: 'feed', name: 'Information feed', color: 'var(--feed)', dash: '5 4', width: 1.6 },
  ];
}

// A badge kind can be linked to a detail field: the badge then shows automatically on any shape
// whose field has text (e.g. C when the Control field is filled in).
export function controlBadgeKind(field) {
  return { id: 'control', text: 'C', color: 'var(--control)', textColor: 'var(--control-ink)', name: 'Control point', ...(field ? { field } : {}) };
}
export function defaultBadgeKinds() {
  return [{ id: 'risk', text: 'R', color: 'var(--risk)', name: 'Risk point', field: 'risk' }, controlBadgeKind('control')];
}

export function defaultFields() {
  return [
    { key: 'what', label: 'What happens' },
    { key: 'system', label: 'System / data' },
    { key: 'output', label: 'Output' },
    { key: 'control', label: 'Control', highlight: true, color: 'var(--control)' },
    { key: 'risk', label: 'Risk point to probe', highlight: true },
  ];
}

export function emptyDoc() {
  return {
    version: FORMAT_VERSION,
    title: 'Untitled flowchart',
    subtitle: '',
    notes: '',
    settings: defaultSettings(),
    lanes: { orientation: 'vertical', header: 40, items: [] },
    phases: { header: 34, items: [] },
    fields: defaultFields(),
    edgeKinds: defaultEdgeKinds(),
    badgeKinds: defaultBadgeKinds(),
    shapeLabels: {},
    nodes: [],
    edges: [],
  };
}

export function makeNode(shape, x, y, extra = {}) {
  const def = shapeDef(shape);
  return {
    id: uid('n'),
    shape,
    x, y,
    w: def.w, h: def.h,
    text: extra.text ?? (shape === 'text' ? 'Text' : shape === 'group' ? 'Group' : def.name),
    tag: '',
    style: {},
    badges: [],
    fields: {},
    ...extra,
  };
}

export function makeEdge(from, to, extra = {}) {
  return {
    id: uid('e'),
    from, to,
    kind: 'flow',
    routing: 'orthogonal',
    waypoints: [],
    label: '',
    labelPos: null,
    labelOffset: null,
    style: {},
    ...extra,
  };
}

// Fill in defaults on a loaded document so older/hand-written files keep working.
export function normalizeDoc(raw) {
  const d = Object.assign(emptyDoc(), clone(raw || {}));
  d.settings = Object.assign(defaultSettings(), d.settings || {});
  d.lanes = Object.assign({ orientation: 'vertical', header: 40, items: [] }, d.lanes || {});
  d.lanes.items = (d.lanes.items || []).map((l) => ({ id: l.id || uid('l'), title: '', subtitle: '', size: 240, ...l }));
  d.phases = Object.assign({ header: 34, items: [] }, d.phases || {});
  d.phases.items = (d.phases.items || []).map((l) => ({ id: l.id || uid('p'), title: '', size: 200, ...l }));
  if (!Array.isArray(d.fields)) d.fields = defaultFields();
  if (!Array.isArray(d.edgeKinds) || !d.edgeKinds.length) d.edgeKinds = defaultEdgeKinds();
  if (!Array.isArray(d.badgeKinds)) d.badgeKinds = defaultBadgeKinds();
  // charts saved before control points existed get the C badge (unlinked; link it in document settings)
  if (!d.badgeKinds.some((b) => b.id === 'control')) d.badgeKinds.push(controlBadgeKind());
  d.shapeLabels = d.shapeLabels || {};
  d.nodes = (d.nodes || []).map((n) => {
    const def = shapeDef(n.shape);
    return {
      id: n.id || uid('n'), shape: n.shape || 'process', x: 0, y: 0, w: def.w, h: def.h,
      text: '', tag: '', ...n,
      style: n.style || {}, badges: n.badges || [], fields: n.fields || {},
    };
  });
  const ids = new Set(d.nodes.map((n) => n.id));
  const okEnd = (p) => p && ((p.node && ids.has(p.node)) || (typeof p.x === 'number' && typeof p.y === 'number'));
  d.edges = (d.edges || [])
    .map((e) => ({
      id: e.id || uid('e'), kind: d.edgeKinds[0].id, routing: 'orthogonal', label: '', labelPos: null, labelOffset: null,
      ...e,
      waypoints: e.waypoints || [], style: e.style || {},
    }))
    .filter((e) => okEnd(e.from) && okEnd(e.to));
  d.version = FORMAT_VERSION;
  return d;
}

export function laneRects(doc, extent) {
  // extent: length of lanes along the flow direction
  const L = doc.lanes;
  const out = [];
  let pos = 0;
  for (const lane of L.items) {
    if (L.orientation === 'horizontal') out.push({ lane, x: 0, y: pos, w: extent, h: lane.size });
    else out.push({ lane, x: pos, y: 0, w: lane.size, h: extent });
    pos += lane.size;
  }
  return out;
}

export function laneAt(doc, x, y) {
  const L = doc.lanes;
  let pos = 0;
  const v = L.orientation === 'horizontal' ? y : x;
  for (const lane of L.items) {
    if (v >= pos && v < pos + lane.size) return lane;
    pos += lane.size;
  }
  return null;
}

// Phases are bands that run across the lanes (rows when lanes are columns, and vice versa).
// They start below the lane headers; their own header strip sits outside the lanes (negative coords).
export function phaseAxis(doc) {
  const horiz = doc.lanes.orientation === 'horizontal';
  return { along: horiz ? 'x' : 'y', start: doc.lanes.items.length ? (doc.lanes.header || 40) : 0 };
}

export function phaseStarts(doc) {
  const { start } = phaseAxis(doc);
  const out = [];
  let pos = start;
  for (const p of doc.phases.items) { out.push([p, pos]); pos += p.size; }
  return out;
}

export function phaseAt(doc, x, y) {
  const { along } = phaseAxis(doc);
  const v = along === 'y' ? y : x;
  const list = phaseStarts(doc);
  for (const [p, pos] of list) if (v >= pos && v < pos + p.size) return p;
  // anything past the end belongs to the last phase (it is drawn stretched)
  if (list.length && v >= list[list.length - 1][1]) return list[list.length - 1][0];
  return null;
}

export function shapeLabel(doc, shape) {
  return (doc.shapeLabels && doc.shapeLabels[shape]) || shapeDef(shape).name;
}

// Grow the phase a shape starts in so the shape fits inside it; content further along moves with it.
export function growPhasesFor(doc, ids) {
  if (!doc.phases.items.length || !ids || !ids.size) return false;
  const { along } = phaseAxis(doc);
  const size = along === 'y' ? 'h' : 'w';
  let changed = false;
  for (const id of ids) {
    const n = doc.nodes.find((x) => x.id === id);
    if (!n || n.shape === 'group') continue;
    const starts = phaseStarts(doc);
    const hit = starts.find(([p, pos]) => n[along] >= pos && n[along] < pos + p.size);
    if (!hit) continue;
    const [p, pos] = hit;
    const end = pos + p.size;
    const need = n[along] + n[size] + 16 - end;
    if (need <= 0) continue;
    const delta = Math.ceil(need / 8) * 8;
    for (const m of doc.nodes) if (m !== n && m[along] >= end - 1) m[along] += delta;
    for (const e of doc.edges) {
      const i = along === 'x' ? 0 : 1;
      e.waypoints = e.waypoints.map((w) => { const q = [...w]; if (q[i] >= end) q[i] += delta; return q; });
      for (const k of ['from', 'to']) if (!e[k].node && e[k][along] >= end) e[k] = { ...e[k], [along]: e[k][along] + delta };
    }
    p.size += delta;
    changed = true;
  }
  return changed;
}

// The badges a shape shows: ticked ones plus those whose linked field has text, each with its
// optional number (C → C1).
export function nodeBadges(doc, n) {
  const out = [];
  for (const kind of doc.badgeKinds) {
    const manual = (n.badges || []).includes(kind.id);
    const auto = !!(kind.field && String(n.fields?.[kind.field] ?? '').trim());
    if (!manual && !auto) continue;
    const num = String(n.badgeNums?.[kind.id] ?? '').trim();
    out.push({ kind, label: (kind.text || '!') + num, auto: auto && !manual });
  }
  return out;
}
