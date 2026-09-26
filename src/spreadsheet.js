// Spreadsheet import: CSV / TSV text, rows pasted from Excel or Google Sheets, and .xlsx files.
// One row per step; columns are mapped to roles (step id, text, lane, phase, next step, ...).
import { emptyDoc, normalizeDoc, uid } from './model.js';
import { shapeDef } from './shapes.js';
import { autoLayout, shapeKey } from './layout.js';

// ---------------------------------------------------------------------------
// parsing

export function parseDelimited(text) {
  const src = String(text).replace(/^﻿/, '');
  const first = src.split(/\r?\n/, 1)[0] || '';
  const delim = first.includes('\t') ? '\t' : (first.split(';').length > first.split(',').length ? ';' : ',');
  const rows = [];
  let row = [], cell = '', q = false;
  for (let i = 0; i < src.length; i++) {
    const c = src[i];
    if (q) {
      if (c === '"') { if (src[i + 1] === '"') { cell += '"'; i++; } else q = false; }
      else cell += c;
    } else if (c === '"' && cell === '') q = true;
    else if (c === delim) { row.push(cell); cell = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && src[i + 1] === '\n') i++;
      row.push(cell); rows.push(row); row = []; cell = '';
    } else cell += c;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return rows.filter((r) => r.some((c) => String(c).trim() !== ''));
}

async function inflateRaw(bytes) {
  const ds = new DecompressionStream('deflate-raw');
  const stream = new Blob([bytes]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

// Minimal .xlsx reader: unzip with the browser's DecompressionStream, read the first worksheet.
export async function readXlsx(buffer) {
  const bytes = new Uint8Array(buffer);
  const dv = new DataView(buffer);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('Not a valid .xlsx file.');
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const files = new Map();
  const dec = new TextDecoder();
  for (let i = 0; i < count; i++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
    const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
    const local = dv.getUint32(p + 42, true);
    const name = dec.decode(bytes.subarray(p + 46, p + 46 + nlen));
    files.set(name, { method, csize, local });
    p += 46 + nlen + xlen + clen;
  }
  const read = async (name) => {
    const f = files.get(name);
    if (!f) return null;
    const start = f.local + 30 + dv.getUint16(f.local + 26, true) + dv.getUint16(f.local + 28, true);
    const data = bytes.subarray(start, start + f.csize);
    return dec.decode(f.method === 8 ? await inflateRaw(data) : data);
  };
  const xml = (s) => new DOMParser().parseFromString(s, 'application/xml');
  // first sheet via workbook relationships, falling back to sheet1
  let sheetPath = 'xl/worksheets/sheet1.xml';
  const wb = await read('xl/workbook.xml'), rels = await read('xl/_rels/workbook.xml.rels');
  if (wb && rels) {
    const sheet = xml(wb).getElementsByTagName('sheet')[0];
    const rid = sheet && (sheet.getAttribute('r:id') || sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'));
    const rel = [...xml(rels).getElementsByTagName('Relationship')].find((r) => r.getAttribute('Id') === rid);
    if (rel) { const t = rel.getAttribute('Target'); sheetPath = t.startsWith('/') ? t.slice(1) : 'xl/' + t.replace(/^\.\//, ''); }
  }
  const ss = await read('xl/sharedStrings.xml');
  const shared = ss ? [...xml(ss).getElementsByTagName('si')].map((si) => [...si.getElementsByTagName('t')].map((t) => t.textContent).join('')) : [];
  const sheetXml = await read(sheetPath);
  if (!sheetXml) throw new Error('Could not find a worksheet in this file.');
  const colIndex = (ref) => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };
  const rows = [];
  for (const r of xml(sheetXml).getElementsByTagName('row')) {
    const out = [];
    let auto = 0;
    for (const c of r.getElementsByTagName('c')) {
      const ref = c.getAttribute('r');
      const ci = ref ? colIndex(ref) : auto;
      auto = ci + 1;
      const t = c.getAttribute('t');
      const v = c.getElementsByTagName('v')[0]?.textContent ?? '';
      let val;
      if (t === 's') val = shared[+v] ?? '';
      else if (t === 'inlineStr') val = [...c.getElementsByTagName('t')].map((x) => x.textContent).join('');
      else if (t === 'b') val = v === '1' ? 'TRUE' : 'FALSE';
      else val = v;
      out[ci] = val;
    }
    rows.push(Array.from(out, (x) => x ?? ''));
  }
  return rows.filter((r) => r.some((c) => String(c).trim() !== ''));
}

// ---------------------------------------------------------------------------
// column roles

export const ROLES = [
  ['ignore', 'Ignore'], ['id', 'Step ID / tag'], ['text', 'Step text'], ['lane', 'Lane (who)'], ['phase', 'Phase / stage'],
  ['shape', 'Shape type'], ['next', 'Next step(s)'], ['prev', 'Previous step(s)'], ['label', 'Connector label'], ['field', 'Detail field'],
  ['riskNum', 'Risk number (R1, R2…)'], ['controlNum', 'Control number (C1, C2…)'],
];

const SYN = {
  id: /^(id|step ?id|key|ref(erence)?|#|no\.?|num(ber)?|step ?(#|no\.?|num(ber)?|ref)|tag|code)$/,
  text: /^(name|step ?name|activity|task|action|title|text|process ?step|step ?(description|text)|what|shape ?text)$/,
  lane: /^(lane|swim ?lane|role|owner|actor|dep(ar)?t(ment)?|team|performed ?by|responsible|function|party|who|organi[sz]ation|system owner)$/,
  phase: /^(phase|stage|sub-?process|section|process ?area|area|milestone|group|cycle)$/,
  shape: /^(shape|type|symbol|kind|step ?type|shape ?type)$/,
  next: /^(next|next ?steps?|goes ?to|leads ?to|to|then|successors?|outputs? ?to|connects? ?to|flows? ?to|next ?id)$/,
  prev: /^(prev(ious)?|previous ?steps?|follows|from|predecessors?|after|comes ?from|inputs? ?from)$/,
  riskNum: /^(key )?risk ?(id|no\.?|#|num(ber)?|ref)$/,
  controlNum: /^(key )?control ?(id|no\.?|#|num(ber)?|ref)$/,
  label: /^(condition|label|connector ?label|edge ?label|branch|outcome|if|decision ?outcome|flow ?label)$/,
};

const norm = (h) => String(h || '').trim().toLowerCase().replace(/[_:]+/g, ' ').replace(/\s+/g, ' ');
const looksLikeCodes = (vals) => {
  const v = vals.filter((x) => String(x).trim());
  if (!v.length) return false;
  const codes = v.filter((x) => /^[A-Za-z]{0,6}[-_. ]?\d{1,4}[a-z]?$/.test(String(x).trim()));
  return codes.length / v.length >= 0.7 && new Set(v).size === v.length;
};

export function guessRoles(headers, rows) {
  const roles = headers.map(() => 'field');
  const used = new Set();
  const col = (i) => rows.map((r) => r[i] ?? '');
  headers.forEach((h, i) => {
    const n = norm(h);
    if (!n) { roles[i] = 'ignore'; return; }
    if (n === 'step' || n === 'step #') { roles[i] = looksLikeCodes(col(i)) ? 'id' : 'text'; return; }
    for (const r of ['riskNum', 'controlNum', 'id', 'text', 'lane', 'phase', 'shape', 'next', 'prev', 'label']) {
      if (SYN[r].test(n) && !used.has(r)) { roles[i] = r; used.add(r); return; }
    }
    if (n === 'description' && !used.has('text') && !headers.some((x) => SYN.text.test(norm(x)))) { roles[i] = 'text'; used.add('text'); }
  });
  if (!roles.includes('id')) {
    const i = roles.findIndex((r, k) => r === 'field' && looksLikeCodes(col(k)));
    if (i >= 0 && i < 2) roles[i] = 'id';
  }
  if (!roles.includes('text')) {
    const i = roles.findIndex((r) => r === 'field');
    if (i >= 0) roles[i] = 'text';
  }
  return roles;
}

// ---------------------------------------------------------------------------
// building the diagram

const RISKY = /risk|issue|gap|deficien|exception|concern|red ?flag/i;
const CONTROL = /\bcontrols?\b/i;
const HIGHLIGHT = new RegExp(RISKY.source + '|' + CONTROL.source, 'i');
// "C1", "C-01" in a Control number column → "1", "01"; anything else is kept as typed
const badgeNum = (v, letter) => { const m = String(v).trim().match(new RegExp('^' + letter + '\\s*[-.]?\\s*(\\d.*)$', 'i')); return m ? m[1] : String(v).trim(); };

// "CD-03", "Yes: CD-03", "CD-03 (Yes)", "Yes -> CD-03", "~CD-05" (second connector type)
function parseTargets(cell) {
  const s = String(cell || '').trim();
  if (!s) return [];
  let parts = s.split(/\s*(?:;|\n|\|)\s*/);
  if (parts.length === 1 && s.includes(',')) parts = s.split(/\s*,\s*/);
  return parts.filter(Boolean).map((t) => {
    let label = '', target = t.trim(), alt = false;
    let m;
    if ((m = target.match(/^(.*?)\s*(?:->|→|=>|:)\s*(.+)$/))) { label = m[1]; target = m[2]; }
    if ((m = target.match(/^(.+?)\s*[([]\s*(.+?)\s*[)\]]$/))) { target = m[1]; label = label || m[2]; }
    if (target.startsWith('~')) { alt = true; target = target.slice(1).trim(); }
    if (label.startsWith('~')) { alt = true; label = label.slice(1).trim(); }
    return { target: target.trim(), label: label.trim(), alt };
  });
}

export function buildFromRows(rows, roles, { headerRow = true, orientation = 'vertical', sequential = true, badges = true, title = 'Imported process' } = {}) {
  const headers = headerRow ? rows[0].map((h, i) => String(h || '').trim() || `Column ${i + 1}`) : rows[0].map((_, i) => `Column ${i + 1}`);
  const data = (headerRow ? rows.slice(1) : rows).filter((r) => r.some((c) => String(c).trim()));
  const warnings = [];
  const doc = emptyDoc();
  doc.title = title;
  doc.lanes.orientation = orientation;
  const idx = (role) => roles.indexOf(role);
  const fieldCols = roles.map((r, i) => (r === 'field' ? i : -1)).filter((i) => i >= 0);
  if (fieldCols.length) {
    const seen = new Set();
    doc.fields = fieldCols.map((i) => {
      let key = norm(headers[i]).replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '') || 'field';
      while (seen.has(key)) key += '_';
      seen.add(key);
      const control = CONTROL.test(headers[i]) && !RISKY.test(headers[i]);
      return { key, label: headers[i], highlight: HIGHLIGHT.test(headers[i]), ...(control ? { color: 'var(--control)' } : {}), col: i };
    });
  }
  // R and C badges appear automatically where their field has text
  for (const kind of doc.badgeKinds) {
    const test = kind.id === 'control' ? (f) => f.color === 'var(--control)' : kind.id === 'risk' ? (f) => f.highlight && !f.color : null;
    const f = badges && test ? doc.fields.find(test) : null;
    if (f) kind.field = f.key; else delete kind.field;
  }
  const lanes = new Map(), phases = new Map();
  const get = (r, role) => { const i = idx(role); return i >= 0 ? String(r[i] ?? '').trim() : ''; };
  const nodes = [];
  const byKey = new Map();
  data.forEach((r, n) => {
    const idv = get(r, 'id');
    const text = get(r, 'text') || idv || `Step ${n + 1}`;
    let shape = null;
    const sv = get(r, 'shape');
    if (sv) { shape = shapeKey(sv); if (!shape) warnings.push(`Row ${n + 1}: unknown shape "${sv}" — used a process box`); }
    if (!shape) shape = /\?\s*$/.test(text) ? 'decision' : /^(start|end|begin|finish|stop|done)\b/i.test(text) ? 'terminator' : 'process';
    const def = shapeDef(shape);
    const laneName = get(r, 'lane'), phaseName = get(r, 'phase');
    if (laneName && !lanes.has(laneName.toLowerCase())) lanes.set(laneName.toLowerCase(), { id: uid('l'), title: laneName, subtitle: '', size: orientation === 'horizontal' ? 150 : 240 });
    if (phaseName && !phases.has(phaseName.toLowerCase())) phases.set(phaseName.toLowerCase(), { id: uid('p'), title: phaseName, size: 200, index: phases.size });
    const node = {
      id: uid('n'), shape, x: 0, y: 0, w: def.w, h: def.h, text, style: {}, badges: [], fields: {},
      tag: idv && /\d/.test(idv) && idv.toLowerCase() !== text.toLowerCase() ? idv : '', // step numbers, not labels like START
      laneId: laneName ? lanes.get(laneName.toLowerCase()).id : undefined,
      phaseIndex: phaseName ? phases.get(phaseName.toLowerCase()).index : undefined,
    };
    for (const f of doc.fields) { const v = String(r[f.col] ?? '').trim(); if (v) node.fields[f.key] = v; }
    for (const [role, kind, letter] of [['riskNum', 'risk', 'R'], ['controlNum', 'control', 'C']]) {
      const v = get(r, role);
      if (!v) continue;
      (node.badgeNums ||= {})[kind] = badgeNum(v, letter);
      if (badges && !doc.badgeKinds.find((k) => k.id === kind)?.field) node.badges.push(kind);
    }
    nodes.push({ node, row: r, n });
    if (idv) byKey.set(idv.toLowerCase(), node);
    if (!byKey.has(text.toLowerCase())) byKey.set(text.toLowerCase(), node);
  });
  const find = (t) => byKey.get(String(t).toLowerCase());
  const edges = [];
  const edgeKey = new Set();
  const addEdge = (a, b, label, alt) => {
    const k = a.id + '>' + b.id;
    if (edgeKey.has(k)) return;
    edgeKey.add(k);
    edges.push({ id: uid('e'), from: { node: a.id, side: 'auto' }, to: { node: b.id, side: 'auto' }, kind: alt ? 'feed' : 'flow', routing: 'orthogonal', waypoints: [], label, labelPos: null, labelOffset: null, style: {} });
  };
  const hasLinks = idx('next') >= 0 || idx('prev') >= 0;
  for (const { node, row, n } of nodes) {
    for (const t of parseTargets(get(row, 'next'))) {
      const b = find(t.target);
      if (b) addEdge(node, b, t.label, t.alt); else warnings.push(`Row ${n + 1}: next step "${t.target}" not found`);
    }
    const prevs = parseTargets(get(row, 'prev'));
    prevs.forEach((t) => {
      const a = find(t.target);
      if (a) addEdge(a, node, t.label || (prevs.length === 1 ? get(row, 'label') : ''), t.alt); else warnings.push(`Row ${n + 1}: previous step "${t.target}" not found`);
    });
    // a connector label column with only a Next column labels that row's single outgoing connector
    if (idx('prev') < 0 && idx('label') >= 0 && get(row, 'label')) {
      const outs = edges.filter((e) => e.from.node === node.id);
      if (outs.length === 1 && !outs[0].label) outs[0].label = get(row, 'label');
    }
  }
  if (!hasLinks && sequential) {
    for (let i = 0; i + 1 < nodes.length; i++) {
      if (nodes[i].node.shape === 'terminator' && i > 0) continue;
      addEdge(nodes[i].node, nodes[i + 1].node, idx('label') >= 0 ? get(nodes[i + 1].row, 'label') : '', false);
    }
  }
  doc.lanes.items = [...lanes.values()];
  doc.phases.items = [...phases.values()].map(({ index, ...p }) => { void index; return p; });
  doc.fields = doc.fields.map(({ col, ...f }) => { void col; return f; });
  // widen lanes to fit their widest shape
  if (orientation !== 'horizontal') for (const l of doc.lanes.items) l.size = Math.max(240, ...nodes.filter((x) => x.node.laneId === l.id).map((x) => x.node.w + 56));
  doc.nodes = nodes.map((x) => x.node);
  doc.edges = edges;
  const out = normalizeDoc(doc);
  const phaseOf = new Map();
  out.nodes.forEach((n) => { if (n.phaseIndex !== undefined) phaseOf.set(n.id, n.phaseIndex); });
  autoLayout(out, { phaseOf: phaseOf.size ? phaseOf : null });
  out.nodes.forEach((n) => { delete n.laneId; delete n.phaseIndex; });
  return { doc: out, warnings, stats: { steps: out.nodes.length, lanes: out.lanes.items.length, phases: out.phases.items.length, connectors: out.edges.length } };
}

export const EXAMPLE_CSV = `Step,Activity,Owner,Phase,Type,Next,What happens,Control,Control #,Risk
CD-00,Start,Requester,Approve,start,CD-01,,,,
CD-01,Submit invoice,Requester,Approve,manual input,CD-02,Requester uploads the supplier invoice to the AP inbox.,,,
CD-02,Match PO?,AP,Approve,decision,Yes: CD-03; No: CD-01,AP matches the invoice to the purchase order and receipt.,"Three-way match of invoice, PO and receipt before approval.",C1,Invoices without a PO could be paid.
CD-03,Schedule payment,AP,Pay,process,CD-04; ~posts: ERP,AP schedules the invoice in the next payment run.,,,
ERP,ERP,AP,Pay,data store,,Invoices and payments are recorded in the ERP.,,,
CD-04,Release payment,Treasury,Pay,process,CD-05,Treasury releases the payment run.,Payment run approved by a second signatory in the bank portal.,C2,Payments could be released without approval.
CD-05,Paid,Treasury,Pay,end,,,,,`;
