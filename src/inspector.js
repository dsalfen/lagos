// Right-hand properties panel. Rebuilt when the selection changes; inputs edit the model live.
import {
  app, emit, change, checkpoint, select, clearSelection, nodeById, selectedNodes, selectedEdges,
} from './state.js';
import { SHAPES, SHAPE_ORDER, shapeDef } from './shapes.js';
import { FONTS, ARROWS, edgeStyle, nodeStyle, fitNodeText } from './render.js';
import { uid, makeNode, laneAt, shapeLabel, phaseAxis, phaseStarts } from './model.js';
import { center, unionBox } from './geometry.js';
import { numberTags, autoLayout } from './layout.js';
import { editLaneTitle, editPhaseTitle } from './canvas.js';

let root;

export function h(tag, props = {}, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(props || {})) {
    if (v === undefined || v === null || v === false) continue;
    if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
    else if (k === 'class') el.className = v;
    else if (k === 'style') el.style.cssText = v;
    else if (['value', 'checked', 'selected', 'disabled', 'hidden', 'title'].includes(k)) el[k] = v;
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const c of kids.flat(Infinity)) {
    if (c === undefined || c === null || c === false) continue;
    el.append(c.nodeType ? c : String(c));
  }
  return el;
}

// Live text input: one undo step per focus session.
function live(el, apply, { rebuild = false } = {}) {
  let pending = true;
  el.addEventListener('focus', () => { pending = true; });
  el.addEventListener('input', () => {
    if (pending) { checkpoint(); pending = false; }
    apply(el.type === 'number' ? (el.value === '' ? null : Number(el.value)) : el.value);
    emit({ doc: true, inspector: rebuild });
  });
  el.addEventListener('change', () => { pending = true; });
  return el;
}

const field = (label, control, cls = '') => h('div', { class: 'f ' + cls }, h('label', {}, label), control);
const section = (title, ...kids) => h('section', {}, title ? h('div', { class: 'sec-title' }, title) : null, ...kids);
const text = (value, apply, attrs = {}) => live(h('input', { type: 'text', value: value ?? '', ...attrs }), apply);
const num = (value, apply, attrs = {}) => live(h('input', { type: 'number', value: value ?? '', step: 'any', ...attrs }), (v) => { if (v !== null && !Number.isNaN(v)) apply(v); });
const area = (value, apply, attrs = {}) => live(h('textarea', { rows: 2, ...attrs }, value ?? ''), apply);
function selectEl(options, value, onChange, attrs = {}) {
  const el = h('select', attrs, options.map(([v, l]) => h('option', { value: v, selected: v === value }, l)));
  el.addEventListener('change', () => change(() => onChange(el.value)));
  return el;
}
function check(label, value, onChange, attrs = {}) {
  const inp = h('input', { type: 'checkbox', checked: !!value, ...attrs });
  inp.addEventListener('change', () => change(() => onChange(inp.checked)));
  return h('label', { class: 'chk' }, inp, label);
}
function btn(label, onClick, attrs = {}) {
  return h('button', { type: 'button', ...attrs, onclick: onClick }, label);
}
function segButtons(options, value, onPick, attrs = {}) {
  return h('div', { class: 'btns', ...attrs }, options.map(([v, l, t]) => btn(l, () => change(() => onPick(v)), { class: v === value ? 'on' : '', title: t || '', 'data-value': v })));
}

const FILLS = ['var(--node)', 'var(--dec)', 'var(--term)', 'var(--note)', 'var(--sel-fill)', '#FDE8E4', '#FFF1CC', '#E3F4E6', '#E1EEFB', '#EFE6F8', '#E9ECEF', 'transparent'];
const LINES = ['var(--node-stroke)', 'var(--flow)', 'var(--feed)', 'var(--risk)', '#2E7D32', '#B7791F', '#6A1B9A', '#0B7285', '#C2185B', '#000000', '#FFFFFF'];

function colorField(value, swatches, onPick, { allowDefault = true, name = 'color' } = {}) {
  const wrap = h('div', { class: 'swatches', 'data-color-field': name });
  if (allowDefault) wrap.append(h('button', { type: 'button', class: 'sw auto' + (!value ? ' on' : ''), title: 'Default', 'data-swatch': '', onclick: () => change(() => onPick('')) }));
  for (const c of swatches) {
    wrap.append(h('button', {
      type: 'button', class: 'sw' + (value === c ? ' on' : ''), title: SWATCH_NAMES[c] || c, 'data-swatch': c,
      style: `background:${c === 'transparent' ? 'repeating-conic-gradient(var(--line) 0 25%, transparent 0 50%) 0 0/8px 8px' : c}`,
      onclick: () => change(() => onPick(c)),
    }));
  }
  const custom = h('input', { type: 'color', value: /^#[0-9a-f]{6}$/i.test(value || '') ? value : '#888888', title: 'Custom colour' });
  live(custom, (v) => onPick(v));
  const isCustom = value && !swatches.includes(value);
  wrap.append(h('span', { class: 'sw custom' + (isCustom ? ' on' : ''), style: isCustom ? `background:${value}` : '' }, custom));
  return wrap;
}

const SWATCH_NAMES = {
  'var(--node)': 'Theme: shape', 'var(--dec)': 'Theme: decision tint', 'var(--term)': 'Theme: grey tint', 'var(--note)': 'Theme: note yellow',
  'var(--sel-fill)': 'Theme: highlight', 'var(--node-stroke)': 'Theme: outline', 'var(--flow)': 'Theme: flow blue', 'var(--feed)': 'Theme: feed grey',
  'var(--risk)': 'Theme: risk red', 'var(--control)': 'Theme: control green', 'var(--ink)': 'Theme: text', 'var(--muted)': 'Theme: muted text', 'var(--lane)': 'Theme: lane',
  'var(--lane-alt)': 'Theme: alternate lane', 'var(--lane-head)': 'Theme: lane header', transparent: 'Transparent', none: 'No border',
};
const HIGHLIGHTS = ['var(--risk)', 'var(--control)', 'var(--flow)', '#B7791F', '#6A1B9A'];
const DASHES = [['', 'Solid'], ['6 4', 'Dashed'], ['2 3', 'Dotted'], ['10 4 2 4', 'Dash-dot']];

export function initInspector() {
  root = document.getElementById('inspector');
}

let shownFor = '';

export function renderInspector() {
  if (!root) return;
  const key = [...app.sel.nodes].join(',') + '|' + [...app.sel.edges].join(',') + '|' + app.sel.lane + '|' + app.sel.phase;
  const active = document.activeElement;
  if (key === shownFor && active && root.contains(active) && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA') && active.type !== 'checkbox' && active.type !== 'color') {
    // don't yank the field out from under the user; the model is already live-updated
    return;
  }
  shownFor = key;
  const scroll = root.scrollTop;
  root.innerHTML = '';
  const nodes = selectedNodes();
  const edges = selectedEdges();
  if (nodes.length === 1 && !edges.length) root.append(...nodeInspector(nodes[0]));
  else if (nodes.length > 1 || (nodes.length && edges.length)) root.append(...multiInspector(nodes, edges));
  else if (edges.length === 1) root.append(...edgeInspector(edges[0]));
  else if (edges.length > 1) root.append(...multiEdgeInspector(edges));
  else if (app.sel.lane) root.append(...laneInspector(app.sel.lane));
  else if (app.sel.phase) root.append(...phaseInspector(app.sel.phase));
  else root.append(...docInspector());
  root.scrollTop = scroll;
}

// ---------------------------------------------------------------------------
function styleSection(nodes) {
  const n0 = nodes[0];
  const st = n0.style || {};
  const set = (k, v) => nodes.forEach((n) => { n.style = { ...n.style }; if (v === '' || v === null) delete n.style[k]; else n.style[k] = v; });
  const ns = nodeStyle(app.doc, n0);
  return section('Style',
    field('Fill', colorField(st.fill, FILLS, (v) => set('fill', v), { name: 'fill' })),
    field('Border', colorField(st.stroke, [...LINES, 'none'], (v) => set('stroke', v), { name: 'stroke' })),
    field('Border', h('div', { class: 'row' },
      num(st.strokeWidth ?? '', (v) => set('strokeWidth', v), { min: 0, max: 12, placeholder: '1.4', title: 'Border width', 'aria-label': 'Border width' }),
      selectEl(DASHES, st.dash ?? (shapeDef(n0.shape).dash || ''), (v) => set('dash', v), { 'aria-label': 'Border style' }))),
    field('Text', colorField(st.textColor, ['var(--ink)', 'var(--muted)', 'var(--flow)', 'var(--risk)', '#FFFFFF', '#2E7D32', '#6A1B9A'], (v) => set('textColor', v), { name: 'text' })),
    field('Font', h('div', { class: 'row' },
      num(st.fontSize ?? '', (v) => set('fontSize', v), { min: 6, max: 72, placeholder: String(app.doc.settings.fontSize), 'aria-label': 'Font size' }),
      h('div', { class: 'btns fixed' },
        btn('B', () => change(() => set('bold', !ns.bold)), { class: ns.bold ? 'on' : '', title: 'Bold (Ctrl+B)', style: 'font-weight:700' }),
        btn('I', () => change(() => set('italic', !ns.italic)), { class: ns.italic ? 'on' : '', title: 'Italic (Ctrl+I)', style: 'font-style:italic' })))),
    field('Align', h('div', { class: 'row' },
      segButtons([['left', '⇤', 'Left'], ['center', '↔', 'Centre'], ['right', '⇥', 'Right']], ns.align, (v) => set('align', v)),
      segButtons([['top', '⤒', 'Top'], ['middle', '↕', 'Middle'], ['bottom', '⤓', 'Bottom']], ns.valign, (v) => set('valign', v)))),
    field('Opacity', live(h('input', { type: 'range', min: 0.1, max: 1, step: 0.05, value: st.opacity ?? 1 }), (v) => set('opacity', Number(v) === 1 ? '' : Number(v)))),
    check('Drop shadow', st.shadow, (v) => set('shadow', v || '')),
  );
}

// Updates the badge ticks while a linked detail field is typed into (the panel is not rebuilt while typing).
let refreshBadges = () => {};

function badgesSection(nodes) {
  refreshBadges = () => {};
  if (!app.doc.badgeKinds.length) return null;
  const fieldLabel = (key) => app.doc.fields.find((f) => f.key === key)?.label || key;
  const rows = app.doc.badgeKinds.map((b) => {
    const box = check(`${b.text} · ${b.name}`, false, (on) => {
      nodes.forEach((n) => { n.badges = n.badges.filter((x) => x !== b.id); if (on) n.badges.push(b.id); });
    }, { 'data-badge-toggle': b.id });
    const nums = new Set(nodes.map((n) => String(n.badgeNums?.[b.id] ?? '')));
    const no = text(nums.size === 1 ? [...nums][0] : '', (v) => {
      for (const n of nodes) {
        n.badgeNums = { ...(n.badgeNums || {}) };
        if (String(v).trim()) n.badgeNums[b.id] = String(v).trim(); else delete n.badgeNums[b.id];
        if (!Object.keys(n.badgeNums).length) delete n.badgeNums;
      }
    }, { class: 'fixed', style: 'width:84px', maxlength: 16, placeholder: 'No. / ID', title: `A number (1 shows as ${b.text}1) or an ID such as K${b.text}-04, shown on the badge and in the details panel`, 'aria-label': `${b.name} number or ID`, 'data-badge-num': b.id });
    const note = h('p', { class: 'hint-text', style: 'margin:0 0 4px 22px' }, `automatic: ${fieldLabel(b.field)} has text`);
    const update = () => {
      // shown automatically because its linked field has text on every selected shape
      const auto = !!b.field && nodes.every((n) => String(n.fields?.[b.field] ?? '').trim());
      const manual = nodes.every((n) => n.badges.includes(b.id));
      const inp = box.querySelector('input');
      inp.checked = manual || auto;
      inp.disabled = auto && !manual;
      note.hidden = !(auto && !manual);
    };
    update();
    return { el: h('div', {}, h('div', { class: 'row' }, box, no), note), update };
  });
  refreshBadges = () => rows.forEach((r) => r.update());
  return section('Badges', rows.map((r) => r.el));
}

function shapeSelect(value, onPick) {
  return selectEl(SHAPE_ORDER.map((k) => [k, SHAPES[k].name]), value, onPick, { 'aria-label': 'Shape type', id: 'insp-shape' });
}

function arrangeButtons(nodes) {
  return h('div', { class: 'btns' },
    btn('Duplicate', () => emit({ command: 'duplicate' }), { title: 'Ctrl+D' }),
    btn('To front', () => change((doc) => { const s = new Set(nodes.map((n) => n.id)); doc.nodes = [...doc.nodes.filter((n) => !s.has(n.id)), ...nodes]; }), { title: 'Ctrl+]' }),
    btn('To back', () => change((doc) => { const s = new Set(nodes.map((n) => n.id)); doc.nodes = [...nodes, ...doc.nodes.filter((n) => !s.has(n.id))]; }), { title: 'Ctrl+[' }),
    btn('Delete', () => emit({ command: 'delete' }), { class: 'danger', title: 'Delete' }),
  );
}

function nodeInspector(n) {
  const def = shapeDef(n.shape);
  const lane = laneAt(app.doc, ...center(n));
  const out = [];
  out.push(h('h3', {}, h('span', { class: 'kind' }, 'Shape'), shapeLabel(app.doc, n.shape)));
  out.push(section(null,
    field('Text', area(n.text, (v) => { n.text = v; fitNodeText(app.doc, n); }, { id: 'insp-text', rows: 3 }), 'stack'),
    field('Tag / ID', text(n.tag, (v) => { n.tag = v; }, { id: 'insp-tag', placeholder: 'e.g. STEP-01' })),
    field('Type', shapeSelect(n.shape, (v) => {
      // keep the shape's size (so alignment and connectors stay put); only tiny symbols snap to their own size
      const d = shapeDef(v);
      if (v === 'connector' || v === 'offpage' || def.w < 80) { const cx = n.x + n.w / 2, cy = n.y + n.h / 2; n.w = d.w; n.h = d.h; n.x = cx - d.w / 2; n.y = cy - d.h / 2; }
      n.shape = v;
      fitNodeText(app.doc, n);
    })),
    lane ? h('p', { class: 'hint-text' }, 'In lane: ', h('b', {}, lane.title || '(untitled)')) : null,
  ));
  const xywh = ['x', 'y', 'w', 'h'].map((k) => h('div', {}, h('label', {}, k.toUpperCase()), num(Math.round(n[k] * 10) / 10, (v) => { if ((k === 'w' || k === 'h') && v < 8) return; n[k] = v; }, { 'data-geom': k })));
  out.push(section('Position & size', h('div', { class: 'grid4' }, xywh)));
  out.push(styleSection([n]));
  out.push(badgesSection([n]));
  const fields = app.doc.fields;
  out.push(section(h('span', {}, 'Details ', h('span', { class: 'muted' }, '(shown in preview/export)')),
    ...fields.map((f) => field(f.label, area(n.fields[f.key], (v) => { n.fields = { ...n.fields, [f.key]: v }; if (!v) delete n.fields[f.key]; refreshBadges(); }, { 'data-field': f.key }), 'stack')),
    field('Lane label', text(n.laneLabel, (v) => { if (v) n.laneLabel = v; else delete n.laneLabel; }, { placeholder: lane?.title || 'override lane name in details' })),
    h('div', { class: 'btns' }, btn('Edit detail fields…', () => { clearSelection(); setTimeout(() => document.querySelector('[data-sec="fields"]')?.scrollIntoView(), 0); })),
  ));
  out.push(section('Arrange', arrangeButtons([n])));
  return out;
}

function multiInspector(nodes, edges) {
  const out = [h('h3', {}, h('span', { class: 'kind' }, 'Selection'), `${nodes.length} shape${nodes.length === 1 ? '' : 's'}${edges.length ? `, ${edges.length} connector${edges.length === 1 ? '' : 's'}` : ''}`)];
  if (nodes.length > 1) {
    const al = (fn) => () => change(() => { const u = unionBox(nodes); nodes.forEach((n) => fn(n, u)); });
    const dist = (axis) => () => change(() => {
      const k = axis === 'x' ? 'x' : 'y', s = axis === 'x' ? 'w' : 'h';
      const sorted = [...nodes].sort((a, b) => a[k] + a[s] / 2 - (b[k] + b[s] / 2));
      const first = sorted[0], lastN = sorted[sorted.length - 1];
      const c0 = first[k] + first[s] / 2, c1 = lastN[k] + lastN[s] / 2;
      sorted.forEach((n, i) => { n[k] = c0 + ((c1 - c0) * i) / (sorted.length - 1) - n[s] / 2; });
    });
    out.push(section('Align',
      h('div', { class: 'btns' },
        btn('Left', al((n, u) => { n.x = u.x; }), { 'data-align': 'left' }),
        btn('Centre', al((n, u) => { n.x = u.x + u.w / 2 - n.w / 2; }), { 'data-align': 'hcenter' }),
        btn('Right', al((n, u) => { n.x = u.x + u.w - n.w; }), { 'data-align': 'right' }),
        btn('Top', al((n, u) => { n.y = u.y; }), { 'data-align': 'top' }),
        btn('Middle', al((n, u) => { n.y = u.y + u.h / 2 - n.h / 2; }), { 'data-align': 'vcenter' }),
        btn('Bottom', al((n, u) => { n.y = u.y + u.h - n.h; }), { 'data-align': 'bottom' })),
      h('div', { class: 'btns', style: 'margin-top:4px' },
        btn('Distribute ↔', dist('x'), { 'data-align': 'dist-x' }),
        btn('Distribute ↕', dist('y'), { 'data-align': 'dist-y' }),
        btn('Same width', () => change(() => nodes.forEach((n) => { n.w = nodes[0].w; })), {}),
        btn('Same height', () => change(() => nodes.forEach((n) => { n.h = nodes[0].h; })), {})),
    ));
  }
  if (nodes.length) {
    out.push(section(null, field('Type', shapeSelect(nodes.every((n) => n.shape === nodes[0].shape) ? nodes[0].shape : '', (v) => nodes.forEach((n) => { n.shape = v; })))));
    out.push(styleSection(nodes));
    out.push(badgesSection(nodes));
    const pre = h('input', { type: 'text', value: 'STEP-', 'aria-label': 'Tag prefix' });
    const start = h('input', { type: 'number', value: 1, 'aria-label': 'First number' });
    const digits = h('input', { type: 'number', value: 2, min: 1, max: 5, 'aria-label': 'Digits' });
    const order = h('select', { 'aria-label': 'Order' }, h('option', { value: 'rows' }, 'Rows (top→bottom)'), h('option', { value: 'cols' }, 'Columns (left→right)'));
    out.push(section('Number tags',
      h('div', { class: 'row' }, pre, start, digits),
      h('div', { class: 'row' }, order, btn('Apply', () => change(() => numberTags(nodes, { prefix: pre.value, start: +start.value, digits: +digits.value, order: order.value })), { class: 'ibtn fixed', id: 'number-apply' })),
    ));
    out.push(section('Arrange', arrangeButtons(nodes)));
  }
  if (edges.length && !nodes.length) out.push(...multiEdgeInspector(edges).slice(1));
  return out;
}

function kindOptions() { return app.doc.edgeKinds.map((k) => [k.id, k.name]); }
const ARROW_OPTS = Object.keys(ARROWS).map((k) => [k, { none: 'None', arrow: 'Arrow', open: 'Open arrow', diamond: 'Diamond', diamondOpen: 'Open diamond', circle: 'Dot', circleOpen: 'Open dot', bar: 'Bar' }[k]]);

function edgeStyleSection(edges) {
  const e0 = edges[0];
  const s = e0.style || {};
  const set = (k, v) => edges.forEach((e) => { e.style = { ...e.style }; if (v === '' || v === null) delete e.style[k]; else e.style[k] = v; });
  const st = edgeStyle(app.doc, e0);
  return section('Line',
    field('Type', selectEl(kindOptions(), e0.kind, (v) => edges.forEach((e) => { e.kind = v; }), { id: 'insp-kind' })),
    field('Routing', segButtons([['orthogonal', 'Elbow'], ['straight', 'Straight'], ['curved', 'Curved']], e0.routing, (v) => edges.forEach((e) => { e.routing = v; if (v === 'orthogonal') e.waypoints = []; }), { id: 'insp-routing' })),
    field('Start', selectEl(ARROW_OPTS, st.start, (v) => set('startArrow', v), { id: 'insp-start' })),
    field('End', selectEl(ARROW_OPTS, st.end, (v) => set('endArrow', v), { id: 'insp-end' })),
    field('Colour', colorField(s.color, LINES, (v) => set('color', v), { name: 'edge' })),
    field('Width', h('div', { class: 'row' },
      num(s.width ?? '', (v) => set('width', v), { min: 0.5, max: 12, placeholder: String(st.width), 'aria-label': 'Line width' }),
      selectEl(DASHES, s.dash ?? st.dash ?? '', (v) => set('dash', v), { 'aria-label': 'Line style' }))),
  );
}

const SIDE_OPTS = [['auto', 'Auto'], ['t', 'Top'], ['r', 'Right'], ['b', 'Bottom'], ['l', 'Left']];
function endControls(e, which) {
  const ep = e[which];
  if (!ep.node) {
    return h('div', {},
      h('div', { class: 'row' }, h('span', { class: 'hint-text' }, 'Free end at'), num(Math.round(ep.x), (v) => { ep.x = v; }), num(Math.round(ep.y), (v) => { ep.y = v; })),
    );
  }
  const n = nodeById(ep.node);
  return h('div', {},
    h('div', { class: 'hint-text' }, (n?.tag ? n.tag + ' · ' : '') + (n?.text || n?.id || '')),
    h('div', { class: 'row' },
      selectEl(SIDE_OPTS, ep.side || 'auto', (v) => { ep.side = v; e.waypoints = []; }, { 'aria-label': which + ' side', 'data-end-side': which }),
      num(ep.offset || 0, (v) => { ep.offset = v; }, { 'aria-label': which + ' offset', title: 'Offset along the side (px)' })),
    btn('Detach end', () => change(() => { const r = document.querySelector(`[data-edge="${CSS.escape(e.id)}"] .edge-line`); const d = r?.getAttribute('d') || ''; const nums = d.match(/-?\d+(\.\d+)?/g)?.map(Number) || [0, 0]; const pt = which === 'from' ? [nums[0], nums[1]] : [nums[nums.length - 2], nums[nums.length - 1]]; const [cx, cy] = center(n); const dx = pt[0] - cx, dy = pt[1] - cy; const L = Math.hypot(dx, dy) || 1; e[which] = { x: Math.round(pt[0] + (dx / L) * 30), y: Math.round(pt[1] + (dy / L) * 30) }; }), { class: 'ibtn sm' }),
  );
}

function edgeInspector(e) {
  const out = [h('h3', {}, h('span', { class: 'kind' }, 'Connector'), e.label ? e.label.split('\n')[0] : 'Connector')];
  out.push(section(null,
    field('Label', area(e.label, (v) => { e.label = v; }, { id: 'insp-label', rows: 2 }), 'stack'),
    h('div', { class: 'btns' },
      btn('Reset label position', () => change(() => { e.labelPos = null; e.labelOffset = null; e.labelAlign = undefined; })),
      btn('Label size', () => change(() => { e.style = { ...e.style, labelSize: (e.style.labelSize || 10.5) >= 14 ? 10.5 : (e.style.labelSize || 10.5) + 1.5 }; }), { title: 'Cycle label size' })),
  ));
  out.push(edgeStyleSection([e]));
  if (app.doc.fields.length) {
    out.push(section(h('span', {}, 'Details ', h('span', { class: 'muted' }, '(shown when the connector is clicked in preview/export)')),
      ...app.doc.fields.map((f) => field(f.label, area(e.fields?.[f.key], (v) => { e.fields = { ...(e.fields || {}), [f.key]: v }; if (!v) delete e.fields[f.key]; }, { 'data-edge-field': f.key }), 'stack'))));
  }
  out.push(section('Ends',
    field('From', endControls(e, 'from'), 'stack'),
    field('To', endControls(e, 'to'), 'stack'),
    h('div', { class: 'btns' },
      btn('Reverse', () => change(() => { [e.from, e.to] = [e.to, e.from]; e.waypoints = [...e.waypoints].reverse(); if (e.labelPos != null) e.labelPos = 1 - e.labelPos; }), { id: 'edge-reverse' }),
      btn('Reset route', () => change(() => { e.waypoints = []; }), { id: 'edge-reset', title: 'Remove manual bends' }),
      btn('Delete', () => emit({ command: 'delete' }), { class: 'danger' })),
    h('p', { class: 'hint-text' }, 'Drag a segment handle to reroute. Drag the end dots to reconnect. Double-click a bend (straight/curved) to remove it.'),
  ));
  return out;
}

function multiEdgeInspector(edges) {
  return [
    h('h3', {}, h('span', { class: 'kind' }, 'Selection'), `${edges.length} connectors`),
    edgeStyleSection(edges),
    section(null, h('div', { class: 'btns' },
      btn('Reset routes', () => change(() => edges.forEach((e) => { e.waypoints = []; }))),
      btn('Delete', () => emit({ command: 'delete' }), { class: 'danger' }))),
  ];
}

function laneInspector(id) {
  const L = app.doc.lanes;
  const i = L.items.findIndex((l) => l.id === id);
  const lane = L.items[i];
  if (!lane) return docInspector();
  const move = (d) => () => change(() => {
    const j = i + d;
    if (j < 0 || j >= L.items.length) return;
    // move the nodes with their lanes
    const starts = []; let p = 0; for (const l of L.items) { starts.push(p); p += l.size; }
    const a = L.items[Math.min(i, j)], b = L.items[Math.max(i, j)];
    const sa = starts[Math.min(i, j)], sb = starts[Math.max(i, j)];
    const key = L.orientation === 'horizontal' ? 'y' : 'x';
    const axis = key === 'x' ? 0 : 1;
    for (const n of app.doc.nodes) {
      const c = center(n)[axis];
      if (c >= sa && c < sa + a.size) n[key] += b.size;
      else if (c >= sb && c < sb + b.size) n[key] -= a.size;
    }
    for (const e of app.doc.edges) {
      e.waypoints = e.waypoints.map((w) => { const q = [...w]; if (q[axis] >= sa && q[axis] < sa + a.size) q[axis] += b.size; else if (q[axis] >= sb && q[axis] < sb + b.size) q[axis] -= a.size; return q; });
    }
    [L.items[i], L.items[j]] = [L.items[j], L.items[i]];
  });
  return [
    h('h3', {}, h('span', { class: 'kind' }, 'Lane'), lane.title || '(untitled lane)'),
    section(null,
      field('Title', text(lane.title, (v) => { lane.title = v; }, { id: 'lane-title' })),
      field('Subtitle', text(lane.subtitle, (v) => { lane.subtitle = v; }, { id: 'lane-subtitle' })),
      field(L.orientation === 'horizontal' ? 'Height' : 'Width', num(lane.size, (v) => setLaneSize(lane, v), { id: 'lane-size', min: 40 })),
      field('Fill', colorField(lane.fill, ['var(--lane)', 'var(--lane-alt)', '#F6FAFF', '#FFFBF0', '#F4FBF5', '#FBF5FF'], (v) => { lane.fill = v || undefined; }, { name: 'lane-fill' })),
      field('Header', colorField(lane.headFill, ['var(--lane-head)', 'var(--sel-fill)', '#FDE8E4', '#FFF1CC', '#E3F4E6', '#EFE6F8'], (v) => { lane.headFill = v || undefined; }, { name: 'lane-head' })),
    ),
    section('Arrange', h('div', { class: 'btns' },
      btn(L.orientation === 'horizontal' ? 'Move up' : 'Move left', move(-1), { id: 'lane-prev' }),
      btn(L.orientation === 'horizontal' ? 'Move down' : 'Move right', move(1), { id: 'lane-next' }),
      btn('Insert after', () => { const nl = { id: uid('l'), title: 'New lane', subtitle: '', size: lane.size }; change(() => { shiftContent(L.orientation === 'horizontal' ? 'y' : 'x', laneStart(i) + lane.size, nl.size); L.items.splice(i + 1, 0, nl); }); select({ lane: nl.id }); }),
      btn('Rename on canvas', () => editLaneTitle(id)),
      btn('Delete lane', () => deleteLane(id), { class: 'danger', id: 'lane-delete', title: 'Deletes the lane and the shapes in it' }),
    ), h('p', { class: 'hint-text' }, 'Drag a lane border on the canvas to resize; shapes to the right/below move with it.')),
  ];
}

function phaseInspector(id) {
  const P = app.doc.phases;
  const i = P.items.findIndex((l) => l.id === id);
  const ph = P.items[i];
  if (!ph) return docInspector();
  const { along } = phaseAxis(app.doc);
  const key = along, axis = along === 'x' ? 0 : 1;
  const shift = (from, delta) => {
    for (const n of app.doc.nodes) if (center(n)[axis] >= from) n[key] += delta;
    for (const e of app.doc.edges) e.waypoints = e.waypoints.map((w) => { const q = [...w]; if (q[axis] >= from) q[axis] += delta; return q; });
  };
  const move = (d) => () => change(() => {
    const j = i + d;
    if (j < 0 || j >= P.items.length) return;
    const starts = phaseStarts(app.doc).map(([, pos]) => pos);
    const lo = Math.min(i, j), hi = Math.max(i, j);
    const a = P.items[lo], b = P.items[hi], sa = starts[lo], sb = starts[hi];
    for (const n of app.doc.nodes) {
      const c = center(n)[axis];
      if (c >= sa && c < sa + a.size) n[key] += b.size;
      else if (c >= sb && c < sb + b.size) n[key] -= a.size;
    }
    [P.items[i], P.items[j]] = [P.items[j], P.items[i]];
  });
  const start = phaseStarts(app.doc)[i][1];
  return [
    h('h3', {}, h('span', { class: 'kind' }, 'Phase'), ph.title || '(untitled phase)'),
    section(null,
      field('Title', text(ph.title, (v) => { ph.title = v; }, { id: 'phase-title' })),
      field(along === 'y' ? 'Height' : 'Width', num(ph.size, (v) => setPhaseSize(ph, v), { id: 'phase-size', min: 24 })),
      field('Band', colorField(ph.fill, ['rgba(43,90,132,0.05)', 'rgba(184,67,43,0.05)', 'rgba(46,125,50,0.06)', 'rgba(183,121,31,0.07)', 'rgba(106,27,154,0.05)'], (v) => { ph.fill = v || undefined; }, { name: 'phase-fill' })),
      field('Header', colorField(ph.headFill, ['var(--lane-head)', 'var(--lane-alt)', 'var(--sel-fill)', '#FDE8E4', '#FFF1CC', '#E3F4E6'], (v) => { ph.headFill = v || undefined; }, { name: 'phase-head' })),
    ),
    section('Arrange', h('div', { class: 'btns' },
      btn(along === 'y' ? 'Move up' : 'Move left', move(-1), { id: 'phase-prev' }),
      btn(along === 'y' ? 'Move down' : 'Move right', move(1), { id: 'phase-next' }),
      btn('Insert after', () => { const np = { id: uid('p'), title: 'New phase', size: ph.size }; change(() => { shift(start + ph.size, np.size); P.items.splice(i + 1, 0, np); }); select({ phase: np.id }); }),
      btn('Rename on canvas', () => editPhaseTitle(id)),
      btn('Delete phase', () => change(() => { P.items.splice(i, 1); }), { class: 'danger', id: 'phase-delete' }),
    ), h('p', { class: 'hint-text' }, 'Phases are bands across the lanes, e.g. stages or sub-processes. Drag a phase border to resize; shapes below/right of it move too.')),
  ];
}

export function addPhase() {
  const P = app.doc.phases;
  const np = { id: uid('p'), title: `Phase ${P.items.length + 1}`, size: 200 };
  change(() => { P.items.push(np); });
  select({ phase: np.id });
}

function laneStart(i) {
  let p = 0;
  for (let j = 0; j < i; j++) p += app.doc.lanes.items[j].size;
  return p;
}

// shift shapes, bends and loose ends that lie at/after `from` along an axis
function shiftContent(axisKey, from, delta) {
  const axis = axisKey === 'x' ? 0 : 1;
  for (const n of app.doc.nodes) if (center(n)[axis] >= from) n[axisKey] += delta;
  for (const e of app.doc.edges) {
    e.waypoints = e.waypoints.map((w) => { const q = [...w]; if (q[axis] >= from) q[axis] += delta; return q; });
    for (const k of ['from', 'to']) if (!e[k].node && e[k][axisKey] >= from) e[k] = { ...e[k], [axisKey]: e[k][axisKey] + delta };
  }
}

function setLaneSize(lane, v) {
  if (!(v >= 40)) return;
  const i = app.doc.lanes.items.indexOf(lane);
  const key = app.doc.lanes.orientation === 'horizontal' ? 'y' : 'x';
  shiftContent(key, laneStart(i) + lane.size, v - lane.size);
  lane.size = v;
}

function setPhaseSize(ph, v) {
  if (!(v >= 24)) return;
  const start = phaseStarts(app.doc).find(([p]) => p === ph)[1];
  shiftContent(phaseAxis(app.doc).along, start + ph.size, v - ph.size);
  ph.size = v;
}

// Deleting a lane removes the shapes inside it and closes the gap.
export function deleteLane(id) {
  const L = app.doc.lanes;
  const i = L.items.findIndex((l) => l.id === id);
  if (i < 0) return;
  const lane = L.items[i];
  const key = L.orientation === 'horizontal' ? 'y' : 'x';
  const axis = key === 'x' ? 0 : 1;
  const start = laneStart(i), end = start + lane.size;
  const gone = new Set(app.doc.nodes.filter((n) => { const c = center(n)[axis]; return c >= start && c < end; }).map((n) => n.id));
  change((doc) => {
    doc.nodes = doc.nodes.filter((n) => !gone.has(n.id));
    doc.edges = doc.edges.filter((e) => !gone.has(e.from.node) && !gone.has(e.to.node));
    shiftContent(key, end, -lane.size);
    L.items.splice(i, 1);
  });
  emit({ toast: `Deleted lane “${lane.title || 'untitled'}”${gone.size ? ` and ${gone.size} shape${gone.size > 1 ? 's' : ''}` : ''} — Ctrl+Z to undo` });
}

export function addLane() {
  const L = app.doc.lanes;
  const nl = { id: uid('l'), title: `Lane ${L.items.length + 1}`, subtitle: '', size: L.orientation === 'horizontal' ? 160 : 240 };
  change(() => { L.items.push(nl); });
  select({ lane: nl.id });
}

function listEditor(items, render, onAdd, addLabel, attrs = {}) {
  return h('div', attrs,
    items.map((it, i) => h('div', { class: 'list-item' }, render(it, i),
      h('div', { class: 'btns', style: 'margin-top:4px' },
        btn('↑', () => change(() => { if (i > 0) [items[i - 1], items[i]] = [items[i], items[i - 1]]; }), { class: 'ibtn sm', title: 'Move up', disabled: i === 0 }),
        btn('↓', () => change(() => { if (i < items.length - 1) [items[i + 1], items[i]] = [items[i], items[i + 1]]; }), { class: 'ibtn sm', title: 'Move down', disabled: i === items.length - 1 }),
        btn('Remove', () => change(() => { items.splice(i, 1); }), { class: 'ibtn sm danger' })))),
    h('div', { class: 'btns' }, btn(addLabel, () => change(onAdd))),
  );
}

function docInspector() {
  const doc = app.doc;
  const S = doc.settings;
  const out = [h('h3', {}, h('span', { class: 'kind' }, 'Document'), 'Diagram settings')];
  out.push(section(null,
    field('Title', text(doc.title, (v) => { doc.title = v; document.getElementById('doc-title').value = v; }, { id: 'doc-title-insp' })),
    field('Subtitle', text(doc.subtitle, (v) => { doc.subtitle = v; }, { id: 'doc-subtitle' })),
    field('Notes', area(doc.notes, (v) => { doc.notes = v; }, { rows: 3, placeholder: 'Shown under the chart in preview/export. One paragraph per line.' }), 'stack'),
  ));
  out.push(section('Canvas',
    field('Theme', selectEl([['auto', 'Follow system'], ['light', 'Light'], ['dark', 'Dark']], S.theme, (v) => { S.theme = v; }, { id: 'set-theme' })),
    field('Font', selectEl(Object.entries(FONTS).map(([k, f]) => [k, f.name]), S.font, (v) => { S.font = v; }, { id: 'set-font' })),
    field('Text size', num(S.fontSize, (v) => { if (v >= 6) S.fontSize = v; }, { min: 6, max: 40 })),
    field('Grid', h('div', { class: 'row' }, num(S.grid, (v) => { if (v >= 2) S.grid = v; }, { min: 2, max: 100, 'aria-label': 'Grid size' }), h('span', { class: 'fixed hint-text' }, 'px'))),
    check('Snap to grid & guides', S.snap, (v) => { S.snap = v; }, { id: 'set-snap' }),
    check('Show grid', S.showGrid, (v) => { S.showGrid = v; }),
    check('Show tags / IDs on shapes', S.showTags, (v) => { S.showTags = v; }, { id: 'set-tags' }),
    field('Corners', num(S.cornerRadius, (v) => { S.cornerRadius = Math.max(0, v); }, { min: 0, max: 30, title: 'Rounded connector corners', id: 'set-corner' })),
    field('V. labels', selectEl([['horizontal', 'Keep horizontal'], ['rotate', 'Rotate along line']], S.verticalLabels || 'horizontal', (v) => { S.verticalLabels = v; }, { title: 'Labels on long vertical connector runs', id: 'set-vlabels' })),
    field('Line hops', h('div', { class: 'row' },
      selectEl([['arc', 'Arc over crossings'], ['none', 'None']], S.lineJumps || 'arc', (v) => { S.lineJumps = v; }, { id: 'set-hops', 'aria-label': 'Line hops' }),
      num(S.jumpSize ?? 5, (v) => { S.jumpSize = Math.max(2, Math.min(12, v)); }, { class: 'fixed', style: 'width:56px', min: 2, max: 12, 'aria-label': 'Hop size', title: 'Hop radius (px)' }))),
    field('Stub', num(S.stub, (v) => { S.stub = Math.max(4, v); }, { min: 4, max: 80, title: 'How far connectors leave a shape before turning' })),
  ));
  out.push(section('Lanes',
    field('Direction', (() => {
      const el = h('select', { id: 'lane-orient' }, [['vertical', 'Columns'], ['horizontal', 'Rows']].map(([v, l]) => h('option', { value: v, selected: v === doc.lanes.orientation }, l)));
      el.addEventListener('change', () => { if (el.value !== doc.lanes.orientation) emit({ command: 'toggle-lane-orient' }); });
      return el;
    })()),
    field('Header', num(doc.lanes.header, (v) => { if (v >= 16) doc.lanes.header = v; }, { min: 16, max: 200 })),
    ...doc.lanes.items.map((l) => h('div', { class: 'row' },
      text(l.title, (v) => { l.title = v; }, { 'aria-label': 'Lane title' }),
      num(l.size, (v) => setLaneSize(l, v), { class: 'fixed', style: 'width:64px', 'aria-label': 'Lane size' }),
      btn('✎', () => select({ lane: l.id }), { class: 'ibtn sm fixed', title: 'More lane options' }))),
    h('div', { class: 'btns' }, btn('+ Add lane', addLane, { id: 'doc-add-lane' })),
  ));
  out.push(section('Phases',
    h('p', { class: 'hint-text' }, 'Bands across the lanes for stages or sub-processes.'),
    ...doc.phases.items.map((l) => h('div', { class: 'row' },
      text(l.title, (v) => { l.title = v; }, { 'aria-label': 'Phase title' }),
      num(l.size, (v) => setPhaseSize(l, v), { class: 'fixed', style: 'width:64px', 'aria-label': 'Phase size' }),
      btn('✎', () => select({ phase: l.id }), { class: 'ibtn sm fixed', title: 'More phase options' }))),
    h('div', { class: 'btns' }, btn('+ Add phase', addPhase, { id: 'doc-add-phase' })),
    field('Heading', text(doc.phases.label, (v) => { doc.phases.label = v; }, { id: 'phase-label', placeholder: 'e.g. BD, Stage' })),
    field('Titles', selectEl([['rotated', 'Rotated, centred'], ['horizontal', 'Horizontal, centred'], ['horizontal-start', 'Horizontal, at top']],
      doc.phases.titleStyle === 'horizontal' ? (doc.phases.titleAt === 'start' ? 'horizontal-start' : 'horizontal') : 'rotated',
      (v) => { doc.phases.titleStyle = v === 'rotated' ? undefined : 'horizontal'; doc.phases.titleAt = v === 'horizontal-start' ? 'start' : undefined; }, { id: 'phase-titles' })),
    field('Strip', num(doc.phases.header, (v) => { if (v >= 20) doc.phases.header = v; }, { min: 20, max: 200, title: 'Width of the phase title strip' })),
  ));
  out.push(section(h('span', { 'data-sec': 'fields' }, 'Detail fields'),
    h('p', { class: 'hint-text' }, 'Each shape can hold these fields. They appear in the details panel of the preview and exported page.'),
    listEditor(doc.fields, (f) => h('div', {},
      h('div', { class: 'row' }, text(f.label, (v) => { f.label = v; }, { 'aria-label': 'Field label' })),
      check('Highlight (e.g. risks)', f.highlight, (v) => { f.highlight = v; }),
      f.highlight ? colorField(f.color || '', HIGHLIGHTS, (v) => { f.color = v || undefined; }, { name: 'field-' + f.key }) : null),
    () => { doc.fields.push({ key: uid('f'), label: 'New field' }); }, '+ Add field', { id: 'fields-list' }),
    field('Hint', text(S.detailHint, (v) => { S.detailHint = v; })),
    field('No tag', text(S.noTagLabel, (v) => { S.noTagLabel = v; }, { placeholder: 'heading for untagged shapes', id: 'set-notag' })),
    check('Show details panel', S.showDetail !== false, (v) => { S.showDetail = v; }),
    check('Make web addresses in details clickable', S.linkify !== false, (v) => { S.linkify = v; }, { id: 'set-linkify' }),
  ));
  out.push(section('Connector types',
    listEditor(doc.edgeKinds, (k) => h('div', {},
      h('div', { class: 'row' }, text(k.name, (v) => { k.name = v; }, { 'aria-label': 'Connector type name' }), num(k.width, (v) => { k.width = v; }, { class: 'fixed', style: 'width:56px', 'aria-label': 'Width' })),
      colorField(k.color, LINES, (v) => { k.color = v || 'var(--flow)'; }, { allowDefault: false, name: 'kind-' + k.id }),
      h('div', { class: 'row', style: 'margin-top:4px' },
        selectEl(DASHES, k.dash || '', (v) => { k.dash = v; }, { 'aria-label': 'Dash' }),
        selectEl(ARROW_OPTS, k.endArrow || 'arrow', (v) => { k.endArrow = v; }, { 'aria-label': 'End arrow' }))),
    () => { doc.edgeKinds.push({ id: uid('k'), name: 'New type', color: '#2E7D32', dash: '', width: 1.6 }); }, '+ Add type'),
  ));
  out.push(section('Badges',
    listEditor(doc.badgeKinds, (b) => h('div', {},
      h('div', { class: 'row' }, text(b.text, (v) => { b.text = v; }, { class: 'fixed', style: 'width:44px', maxlength: 3, 'aria-label': 'Badge text' }), text(b.name, (v) => { b.name = v; }, { 'aria-label': 'Badge meaning' })),
      field('Auto', selectEl([['', 'Only when ticked'], ...doc.fields.map((f) => [f.key, `When “${f.label}” has text`])], b.field || '',
        (v) => { b.field = v; }, { 'aria-label': 'Show automatically', title: 'Show this badge automatically on shapes where this detail field has text', 'data-badge-field': b.id })),
      colorField(b.color, ['var(--risk)', 'var(--control)', 'var(--flow)', '#2E7D32', '#B7791F', '#6A1B9A', '#0B7285', '#C2185B'], (v) => { b.color = v || 'var(--risk)'; }, { allowDefault: false, name: 'badge-' + b.id })),
    () => { doc.badgeKinds.push({ id: uid('b'), text: 'B', name: 'New badge', color: '#B7791F' }); }, '+ Add badge'),
  ));
  const used = [...new Set(doc.nodes.map((n) => n.shape))].filter((k) => !SHAPES[k]?.noLegend);
  out.push(section('Legend',
    check('Show legend', S.showLegend !== false, (v) => { S.showLegend = v; }),
    ...used.map((k) => field(SHAPES[k].name, text(doc.shapeLabels[k] || '', (v) => { if (v) doc.shapeLabels[k] = v; else delete doc.shapeLabels[k]; }, { placeholder: SHAPES[k].name }))),
    field('Dashed', text(doc.dashedNodeLabel || '', (v) => { doc.dashedNodeLabel = v; }, { placeholder: 'meaning of dashed outline' })),
  ));
  out.push(section('Tools',
    h('div', { class: 'btns' },
      btn('Auto-layout', () => change(() => autoLayout(doc)), { id: 'auto-layout', title: 'Arrange shapes in ranks following the connectors' }),
      btn('Select all', () => select({ nodes: doc.nodes.map((n) => n.id), edges: doc.edges.map((e) => e.id) })),
    ),
    h('p', { class: 'hint-text' }, `${doc.nodes.length} shapes · ${doc.edges.length} connectors · ${doc.lanes.items.length} lanes`),
  ));
  return out;
}

export { makeNode };
