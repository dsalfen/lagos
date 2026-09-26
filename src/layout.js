// Layout helpers: automatic layered layout, tag numbering and the text-outline importer.
import { shapeDef, SHAPES } from './shapes.js';
import { center } from './geometry.js';
import { emptyDoc, normalizeDoc, laneAt, uid } from './model.js';

export function numberTags(nodes, { prefix = 'STEP-', start = 1, digits = 2, order = 'rows' } = {}) {
  const tol = 20;
  const sorted = [...nodes].sort((a, b) => {
    const [ax, ay] = center(a), [bx, by] = center(b);
    if (order === 'cols') return Math.abs(ax - bx) > tol ? ax - bx : ay - by;
    return Math.abs(ay - by) > tol ? ay - by : ax - bx;
  });
  sorted.forEach((n, i) => { n.tag = prefix + String(start + i).padStart(Math.max(1, digits || 1), '0'); });
}

const layoutable = (n) => !shapeDef(n.shape).container && n.shape !== 'text' && n.shape !== 'annotation';

// Layered layout following the connectors. Respects lanes when present.
export function autoLayout(doc, { gap = 44, laneGap = 60, phaseOf = null } = {}) {
  const nodes = doc.nodes.filter(layoutable);
  if (!nodes.length) return;
  const ids = new Set(nodes.map((n) => n.id));
  const out = new Map(nodes.map((n) => [n.id, []]));
  const inc = new Map(nodes.map((n) => [n.id, []]));
  for (const e of doc.edges) {
    const a = e.from.node, b = e.to.node;
    if (a && b && ids.has(a) && ids.has(b) && a !== b) { out.get(a).push(b); inc.get(b).push(a); }
  }
  // break cycles with a DFS in document order
  const state = new Map(), back = new Set();
  const dfs = (v) => {
    state.set(v, 1);
    for (const w of out.get(v)) {
      if (state.get(w) === 1) back.add(v + '>' + w);
      else if (!state.get(w)) dfs(w);
    }
    state.set(v, 2);
  };
  const roots = nodes.filter((n) => !inc.get(n.id).length);
  for (const n of [...roots, ...nodes]) if (!state.get(n.id)) dfs(n.id);
  const fwd = (v) => out.get(v).filter((w) => !back.has(v + '>' + w));

  // longest-path ranks
  const rank = new Map(nodes.map((n) => [n.id, 0]));
  let changed = true, guard = 0;
  while (changed && guard++ < nodes.length + 5) {
    changed = false;
    for (const n of nodes) for (const w of fwd(n.id)) {
      if (rank.get(w) < rank.get(n.id) + 1) { rank.set(w, rank.get(n.id) + 1); changed = true; }
    }
  }

  // Phases (optional): every step of phase p must come after every step of earlier phases.
  const enforcePhases = () => {
    if (!phaseOf) return false;
    let moved = false;
    const idx = [...new Set([...phaseOf.values()])].sort((a, b) => a - b);
    let prevMax = -1;
    for (const p of idx) {
      const members = nodes.filter((n) => phaseOf.get(n.id) === p);
      if (!members.length) continue;
      const minP = Math.min(...members.map((n) => rank.get(n.id)));
      if (minP <= prevMax) {
        const shift = prevMax + 1 - minP;
        for (const n of members) rank.set(n.id, rank.get(n.id) + shift);
        moved = true;
      }
      prevMax = Math.max(prevMax, ...members.map((n) => rank.get(n.id)));
    }
    return moved;
  };
  const propagate = () => {
    let moved = false;
    for (const n of nodes) for (const w of fwd(n.id)) if (rank.get(w) < rank.get(n.id) + 1) { rank.set(w, rank.get(n.id) + 1); moved = true; }
    return moved;
  };
  for (let i = 0; i < nodes.length + 5 && (enforcePhases() | propagate()); i++);

  const lanes = doc.lanes.items;
  const horiz = doc.lanes.orientation === 'horizontal';
  const H = doc.lanes.header || 40;
  if (lanes.length) {
    // keep each node in its lane; avoid two nodes sharing a lane+rank
    const laneOf = new Map(nodes.map((n) => {
      const l = n.laneId ? lanes.find((x) => x.id === n.laneId) : laneAt(doc, ...center(n));
      return [n.id, l ? l.id : lanes[0].id];
    }));
    for (let pass = 0; pass < nodes.length * 2; pass++) {
      let bumped = false;
      const seen = new Map();
      for (const n of [...nodes].sort((a, b) => rank.get(a.id) - rank.get(b.id))) {
        const key = laneOf.get(n.id) + '|' + rank.get(n.id);
        if (seen.has(key)) { rank.set(n.id, rank.get(n.id) + 1); bumped = true; } else seen.set(key, n.id);
      }
      // re-propagate
      if (propagate()) bumped = true;
      if (enforcePhases()) bumped = true;
      if (!bumped) break;
    }
    const starts = new Map(); let p = 0;
    for (const l of lanes) { starts.set(l.id, p); p += l.size; }
    const maxRank = Math.max(...rank.values());
    const along = [];
    let cur = H + (horiz ? 40 : 30);
    for (let r = 0; r <= maxRank; r++) {
      const inRank = nodes.filter((n) => rank.get(n.id) === r);
      const size = Math.max(40, ...inRank.map((n) => (horiz ? n.w : n.h)));
      along.push(cur + size / 2);
      cur += size + (horiz ? laneGap : gap);
    }
    for (const n of nodes) {
      const l = lanes.find((x) => x.id === laneOf.get(n.id));
      const c = starts.get(l.id) + l.size / 2;
      const a = along[rank.get(n.id)];
      if (horiz) { n.x = Math.round(a - n.w / 2); n.y = Math.round(c - n.h / 2); }
      else { n.x = Math.round(c - n.w / 2); n.y = Math.round(a - n.h / 2); }
    }
  } else {
    const byRank = [];
    for (const n of nodes) (byRank[rank.get(n.id)] ||= []).push(n);
    const pos = new Map();
    let y = 40;
    for (const row of byRank) {
      if (!row) continue;
      // order by average position of predecessors (barycentre)
      row.sort((a, b) => bary(a) - bary(b));
      const colW = Math.max(...row.map((n) => n.w)) + laneGap;
      const rowH = Math.max(...row.map((n) => n.h));
      row.forEach((n, i) => {
        const cx = (i - (row.length - 1) / 2) * colW + 400;
        n.x = Math.round(cx - n.w / 2);
        n.y = Math.round(y + (rowH - n.h) / 2);
        pos.set(n.id, cx);
      });
      y += rowH + gap;
    }
    function bary(n) {
      const ps = inc.get(n.id).filter((p) => pos.has(p));
      return ps.length ? ps.reduce((s, p) => s + pos.get(p), 0) / ps.length : 0;
    }
  }
  // size the phase bands around their steps
  if (phaseOf && doc.phases.items.length) {
    const key = horiz && lanes.length ? 'x' : 'y', size = key === 'x' ? 'w' : 'h';
    let b = lanes.length ? H : 0;
    doc.phases.items.forEach((ph, p) => {
      const mem = nodes.filter((n) => phaseOf.get(n.id) === p);
      const next = nodes.filter((n) => phaseOf.get(n.id) === p + 1);
      const hi = mem.length ? Math.max(...mem.map((n) => n[key] + n[size])) : b + 60;
      const end = next.length ? (hi + Math.min(...next.map((n) => n[key]))) / 2 : hi + 30;
      ph.size = Math.max(40, Math.round(end - b));
      b += ph.size;
    });
  }
  for (const e of doc.edges) {
    e.waypoints = [];
    if (e.from.node) e.from = { node: e.from.node, side: 'auto', offset: 0 };
    if (e.to.node) e.to = { node: e.to.node, side: 'auto', offset: 0 };
    e.labelPos = null; e.labelOffset = null;
  }
}

// ---------------------------------------------------------------------------
// Text outline importer

export const OUTLINE_EXAMPLE = `title: Purchase to pay
lanes: Requester | Procurement | Finance

[Requester]
start: Need identified {terminator}
req: Raise purchase request #P2P-01
  what: Requester fills in the purchase request form.

[Procurement]
ok: Within budget? {decision} #P2P-02
po: Issue purchase order #P2P-03 !R
  risk: POs could be raised without an approved request.

[Finance]
inv: Match invoice to PO {document} #P2P-04
pay: Pay supplier #P2P-05
end: Paid {terminator}

start -> req -> ok
ok -> po : Yes
ok -> req : No, revise
po -> inv -> pay -> end
po -.-> pay : PO data`;

export const shapeKey = (s) => {
  const t = s.trim().toLowerCase().replace(/[\s/_-]+/g, '');
  for (const k of Object.keys(SHAPES)) {
    if (k.toLowerCase() === t || SHAPES[k].name.toLowerCase().replace(/[\s/_-]+/g, '') === t) return k;
  }
  const aliases = { proc: 'process', dec: 'decision', diamond: 'decision', term: 'terminator', start: 'terminator', end: 'terminator', doc: 'document', report: 'document', io: 'data', input: 'data', transfer: 'data', manual: 'manualInput', db: 'database', store: 'database', cylinder: 'database', sub: 'predefined', subprocess: 'predefined', circle: 'connector', person: 'actor' };
  return aliases[t] || null;
};

export function parseOutline(src) {
  const doc = emptyDoc();
  doc.title = 'Imported flowchart';
  const nodes = new Map();
  const lanesByName = new Map();
  let lane = null;
  let lastNode = null;
  const errors = [];
  const getLane = (name) => {
    const key = name.trim().toLowerCase();
    if (!lanesByName.has(key)) {
      const l = { id: uid('l'), title: name.trim(), subtitle: '', size: 240 };
      lanesByName.set(key, l); doc.lanes.items.push(l);
    }
    return lanesByName.get(key);
  };
  const ensureNode = (id) => {
    if (!nodes.has(id)) {
      const text = id;
      const lower = id.toLowerCase();
      const shape = /\?$/.test(text) ? 'decision' : ['start', 'end', 'stop', 'begin', 'finish', 'done'].includes(lower) ? 'terminator' : 'process';
      const d = shapeDef(shape);
      const n = { id: uid('n'), key: id, shape, x: 0, y: 0, w: d.w, h: d.h, text, tag: '', style: {}, badges: [], fields: {}, laneId: lane?.id };
      nodes.set(id, n);
    }
    return nodes.get(id);
  };
  const fieldKey = (name) => {
    const t = name.trim().toLowerCase();
    let f = doc.fields.find((x) => x.key.toLowerCase() === t || x.label.toLowerCase() === t);
    if (!f) { f = { key: t.replace(/\W+/g, '_'), label: name.trim().replace(/^\w/, (c) => c.toUpperCase()) }; doc.fields.push(f); }
    return f.key;
  };
  src.split(/\r?\n/).forEach((raw, i) => {
    const line = raw.replace(/\s+$/, '');
    if (!line.trim() || line.trim().startsWith('//') || /^\s*#\s/.test(line)) return;
    const indented = /^\s+/.test(line);
    const s = line.trim();
    let m;
    if (!indented && (m = s.match(/^(title|subtitle|notes)\s*:\s*(.*)$/i))) { doc[m[1].toLowerCase()] = m[2]; return; }
    if (!indented && (m = s.match(/^lanes\s*:\s*(.*)$/i))) { m[1].split('|').map((x) => x.trim()).filter(Boolean).forEach(getLane); return; }
    if (!indented && (m = s.match(/^rows\s*:\s*(.*)$/i))) { doc.lanes.orientation = 'horizontal'; m[1].split('|').map((x) => x.trim()).filter(Boolean).forEach((n) => { getLane(n).size = 160; }); return; }
    if ((m = s.match(/^\[(.+)\]$/))) { lane = getLane(m[1]); return; }
    if (s.includes('->') || s.includes('-.->') || s.includes('==>')) {
      // chain: a -> b -.-> c : label
      let label = '';
      let body = s;
      const lm = s.match(/^(.*?)\s:\s(.*)$/);
      if (lm) { body = lm[1]; label = lm[2].trim(); }
      const parts = body.split(/\s*(-\.->|->|==>)\s*/);
      for (let j = 0; j + 2 < parts.length; j += 2) {
        const a = ensureNode(parts[j].trim()), b = ensureNode(parts[j + 2].trim());
        const kind = parts[j + 1] === '-.->' ? (doc.edgeKinds[1]?.id || doc.edgeKinds[0].id) : doc.edgeKinds[0].id;
        const isLast = j + 2 === parts.length - 1;
        doc.edges.push({ id: uid('e'), from: { node: a.id, side: 'auto' }, to: { node: b.id, side: 'auto' }, kind, routing: 'orthogonal', waypoints: [], label: isLast ? label : '', labelPos: null, labelOffset: null, style: parts[j + 1] === '==>' ? { width: 2.6 } : {} });
      }
      return;
    }
    if (indented && lastNode && (m = s.match(/^([\w \-/]+?)\s*:\s*(.*)$/))) {
      const k = fieldKey(m[1]);
      lastNode.fields[k] = (lastNode.fields[k] ? lastNode.fields[k] + '\n' : '') + m[2];
      return;
    }
    if ((m = s.match(/^([\w\-.]+)\s*:\s*(.*)$/))) {
      const n = ensureNode(m[1]);
      let rest = m[2];
      const sm = rest.match(/\{([^}]+)\}/);
      if (sm) {
        const k = shapeKey(sm[1]);
        if (k) { n.shape = k; const d = shapeDef(k); n.w = d.w; n.h = d.h; } else errors.push(`Line ${i + 1}: unknown shape "${sm[1]}"`);
        rest = rest.replace(sm[0], '');
      }
      rest = rest.replace(/(^|\s)#([\w\-.]+)/g, (_, sp, t) => { n.tag = t; return sp; });
      rest = rest.replace(/(^|\s)!(\w+)/g, (_, sp, b) => {
        const bk = doc.badgeKinds.find((x) => x.text.toLowerCase() === b.toLowerCase() || x.id === b.toLowerCase());
        if (bk) n.badges.push(bk.id); else errors.push(`Line ${i + 1}: unknown badge "!${b}"`);
        return sp;
      });
      n.text = rest.trim() || n.text;
      if (lane) n.laneId = lane.id;
      lastNode = n;
      return;
    }
    errors.push(`Line ${i + 1}: could not understand "${s}"`);
  });
  doc.nodes = [...nodes.values()].map((n) => { const { key, ...rest } = n; void key; return rest; });
  const out = normalizeDoc(doc);
  out.nodes.forEach((n, i) => { n.laneId = doc.nodes[i].laneId; });
  autoLayout(out);
  out.nodes.forEach((n) => { delete n.laneId; });
  return { doc: out, errors };
}
