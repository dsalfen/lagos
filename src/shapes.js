// Shape library. Every shape is drawn in local coordinates (0,0)-(w,h).
// draw() returns [tag, attrs] pairs: the first is the outline (class "shape"),
// the rest are decorations drawn on top with no fill.
// outline() returns a polygon used for connection points and hit maths.

const rectPoly = (w, h) => [[0, 0], [w, 0], [w, h], [0, h]];

function ellipsePoly(cx, cy, rx, ry, a0 = 0, a1 = Math.PI * 2, n = 32) {
  const pts = [];
  for (let i = 0; i <= n; i++) {
    const a = a0 + ((a1 - a0) * i) / n;
    pts.push([cx + rx * Math.cos(a), cy + ry * Math.sin(a)]);
  }
  return pts;
}

function docPath(x, y, w, h) {
  const b = y + h - Math.min(8, h * 0.14);
  const amp = Math.min(10, h * 0.16);
  return `M${x} ${y}H${x + w}V${b}Q${x + (w * 3) / 4} ${b - amp} ${x + w / 2} ${b}T${x} ${b}Z`;
}

const skew = (w, h) => Math.min(16, w * 0.15, h * 0.5);

export const SHAPES = {
  process: {
    name: 'Process', w: 184, h: 60,
    draw: (w, h) => [['rect', { x: 0, y: 0, width: w, height: h, rx: 3 }]],
  },
  rounded: {
    name: 'Alternate process', w: 184, h: 60,
    draw: (w, h) => [['rect', { x: 0, y: 0, width: w, height: h, rx: Math.min(14, h / 4) }]],
  },
  decision: {
    name: 'Decision', w: 184, h: 80, fill: 'var(--dec)',
    draw: (w, h) => [['path', { d: `M${w / 2} 0L${w} ${h / 2}L${w / 2} ${h}L0 ${h / 2}Z` }]],
    outline: (w, h) => [[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]],
    text: (w, h) => ({ x: w * 0.2, y: h * 0.18, w: w * 0.6, h: h * 0.64 }),
    badge: (w, h) => [w / 2 + 26, 14],
  },
  terminator: {
    name: 'Start / end', w: 156, h: 40, fill: 'var(--term)',
    draw: (w, h) => [['rect', { x: 0, y: 0, width: w, height: h, rx: Math.min(h, w) / 2 }]],
    outline: (w, h) => {
      const r = Math.min(h, w) / 2;
      return [
        ...ellipsePoly(w - r, h / 2, r, h / 2, -Math.PI / 2, Math.PI / 2, 12),
        ...ellipsePoly(r, h / 2, r, h / 2, Math.PI / 2, (Math.PI * 3) / 2, 12),
      ];
    },
    text: (w, h) => ({ x: h / 3, y: 2, w: w - (h * 2) / 3, h: h - 4 }),
  },
  document: {
    name: 'Document / report', w: 184, h: 60,
    draw: (w, h) => [['path', { d: docPath(0, 0, w, h) }]],
    text: (w, h) => ({ x: 8, y: 4, w: w - 16, h: h - 16 }),
  },
  multidoc: {
    name: 'Multiple documents', w: 184, h: 66,
    draw: (w, h) => [
      ['path', { d: docPath(8, 0, w - 8, h - 8) }],
      ['path', { d: docPath(4, 4, w - 8, h - 8) }],
      ['path', { d: docPath(0, 8, w - 8, h - 8) }],
    ],
    layered: true,
    text: (w, h) => ({ x: 8, y: 12, w: w - 24, h: h - 26 }),
  },
  data: {
    name: 'Data / transfer', w: 184, h: 60,
    draw: (w, h) => { const s = skew(w, h); return [['path', { d: `M${s} 0H${w}L${w - s} ${h}H0Z` }]]; },
    outline: (w, h) => { const s = skew(w, h); return [[s, 0], [w, 0], [w - s, h], [0, h]]; },
    text: (w, h) => { const s = skew(w, h); return { x: s + 4, y: 4, w: w - 2 * s - 8, h: h - 8 }; },
  },
  manualInput: {
    name: 'Manual input', w: 184, h: 60,
    draw: (w, h) => { const s = Math.min(12, h * 0.3); return [['path', { d: `M0 ${s}L${w} 0V${h}H0Z` }]]; },
    outline: (w, h) => { const s = Math.min(12, h * 0.3); return [[0, s], [w, 0], [w, h], [0, h]]; },
    text: (w, h) => { const s = Math.min(12, h * 0.3); return { x: 8, y: s + 2, w: w - 16, h: h - s - 6 }; },
  },
  database: {
    name: 'Data store', w: 184, h: 64,
    draw: (w, h) => {
      const ry = Math.min(8, h * 0.15);
      return [
        ['path', { d: `M0 ${ry}A${w / 2} ${ry} 0 0 1 ${w} ${ry}V${h - ry}A${w / 2} ${ry} 0 0 1 0 ${h - ry}Z` }],
        ['path', { d: `M0 ${ry}A${w / 2} ${ry} 0 0 0 ${w} ${ry}` }],
      ];
    },
    text: (w, h) => { const ry = Math.min(8, h * 0.15); return { x: 8, y: ry * 2 + 1, w: w - 16, h: h - ry * 3 - 2 }; },
  },
  predefined: {
    name: 'Predefined process', w: 184, h: 60,
    draw: (w, h) => {
      const i = Math.min(10, w * 0.08);
      return [
        ['rect', { x: 0, y: 0, width: w, height: h }],
        ['path', { d: `M${i} 0V${h}M${w - i} 0V${h}` }],
      ];
    },
    text: (w, h) => { const i = Math.min(10, w * 0.08); return { x: i + 4, y: 4, w: w - 2 * i - 8, h: h - 8 }; },
  },
  preparation: {
    name: 'Preparation', w: 184, h: 60,
    draw: (w, h) => { const s = Math.min(20, w * 0.15); return [['path', { d: `M${s} 0H${w - s}L${w} ${h / 2}L${w - s} ${h}H${s}L0 ${h / 2}Z` }]]; },
    outline: (w, h) => { const s = Math.min(20, w * 0.15); return [[s, 0], [w - s, 0], [w, h / 2], [w - s, h], [s, h], [0, h / 2]]; },
    text: (w, h) => { const s = Math.min(20, w * 0.15); return { x: s, y: 4, w: w - 2 * s, h: h - 8 }; },
  },
  manualOp: {
    name: 'Manual operation', w: 184, h: 60,
    draw: (w, h) => { const s = Math.min(18, w * 0.12); return [['path', { d: `M0 0H${w}L${w - s} ${h}H${s}Z` }]]; },
    outline: (w, h) => { const s = Math.min(18, w * 0.12); return [[0, 0], [w, 0], [w - s, h], [s, h]]; },
    text: (w, h) => { const s = Math.min(18, w * 0.12); return { x: s, y: 4, w: w - 2 * s, h: h - 8 }; },
  },
  delay: {
    name: 'Delay', w: 150, h: 60,
    draw: (w, h) => { const r = Math.min(h / 2, w / 2); return [['path', { d: `M0 0H${w - r}A${r} ${h / 2} 0 0 1 ${w - r} ${h}H0Z` }]]; },
    outline: (w, h) => { const r = Math.min(h / 2, w / 2); return [[0, 0], ...ellipsePoly(w - r, h / 2, r, h / 2, -Math.PI / 2, Math.PI / 2, 12), [0, h]]; },
    text: (w, h) => ({ x: 8, y: 4, w: w - Math.min(h / 2, w / 2) - 8, h: h - 8 }),
  },
  display: {
    name: 'Display', w: 184, h: 60,
    draw: (w, h) => {
      const s = Math.min(20, w * 0.14);
      return [['path', { d: `M0 ${h / 2}L${s} 0H${w - s}A${s} ${h / 2} 0 0 1 ${w - s} ${h}H${s}Z` }]];
    },
    outline: (w, h) => {
      const s = Math.min(20, w * 0.14);
      return [[0, h / 2], [s, 0], ...ellipsePoly(w - s, h / 2, s, h / 2, -Math.PI / 2, Math.PI / 2, 12), [s, h]];
    },
    text: (w, h) => { const s = Math.min(20, w * 0.14); return { x: s, y: 4, w: w - 2 * s, h: h - 8 }; },
  },
  storedData: {
    name: 'Stored data', w: 184, h: 60,
    draw: (w, h) => {
      const r = Math.min(14, w * 0.1);
      return [['path', { d: `M${r} 0H${w}A${r} ${h / 2} 0 0 0 ${w} ${h}H${r}A${r} ${h / 2} 0 0 1 ${r} 0Z` }]];
    },
    text: (w, h) => { const r = Math.min(14, w * 0.1); return { x: r + 2, y: 4, w: w - 2 * r - 4, h: h - 8 }; },
  },
  ellipse: {
    name: 'Ellipse', w: 150, h: 70,
    draw: (w, h) => [['ellipse', { cx: w / 2, cy: h / 2, rx: w / 2, ry: h / 2 }]],
    outline: (w, h) => ellipsePoly(w / 2, h / 2, w / 2, h / 2),
    text: (w, h) => ({ x: w * 0.15, y: h * 0.15, w: w * 0.7, h: h * 0.7 }),
  },
  connector: {
    name: 'Connector', w: 40, h: 40, fill: 'var(--term)',
    draw: (w, h) => [['ellipse', { cx: w / 2, cy: h / 2, rx: w / 2, ry: h / 2 }]],
    outline: (w, h) => ellipsePoly(w / 2, h / 2, w / 2, h / 2),
    text: (w, h) => ({ x: 2, y: 2, w: w - 4, h: h - 4 }),
    badge: (w) => [w - 4, 4],
  },
  offpage: {
    name: 'Off-page connector', w: 60, h: 56,
    draw: (w, h) => [['path', { d: `M0 0H${w}V${h * 0.62}L${w / 2} ${h}L0 ${h * 0.62}Z` }]],
    outline: (w, h) => [[0, 0], [w, 0], [w, h * 0.62], [w / 2, h], [0, h * 0.62]],
    text: (w, h) => ({ x: 3, y: 3, w: w - 6, h: h * 0.62 - 3 }),
  },
  merge: {
    name: 'Merge', w: 80, h: 60,
    draw: (w, h) => [['path', { d: `M0 0H${w}L${w / 2} ${h}Z` }]],
    outline: (w, h) => [[0, 0], [w, 0], [w / 2, h]],
    text: (w, h) => ({ x: w * 0.2, y: 3, w: w * 0.6, h: h * 0.5 }),
  },
  extract: {
    name: 'Extract', w: 80, h: 60,
    draw: (w, h) => [['path', { d: `M${w / 2} 0L${w} ${h}H0Z` }]],
    outline: (w, h) => [[w / 2, 0], [w, h], [0, h]],
    text: (w, h) => ({ x: w * 0.2, y: h * 0.45, w: w * 0.6, h: h * 0.52 }),
  },
  hexagon: {
    name: 'Hexagon', w: 120, h: 70,
    draw: (w, h) => { const s = w * 0.25; return [['path', { d: `M${s} 0H${w - s}L${w} ${h / 2}L${w - s} ${h}H${s}L0 ${h / 2}Z` }]]; },
    outline: (w, h) => { const s = w * 0.25; return [[s, 0], [w - s, 0], [w, h / 2], [w - s, h], [s, h], [0, h / 2]]; },
    text: (w, h) => ({ x: w * 0.2, y: 4, w: w * 0.6, h: h - 8 }),
  },
  cloud: {
    name: 'Cloud / external', w: 170, h: 80,
    draw: (w, h) => [['path', {
      d: `M${w * 0.25} ${h * 0.85}C${w * 0.05} ${h * 0.85} ${w * 0.02} ${h * 0.55} ${w * 0.18} ${h * 0.5}` +
        `C${w * 0.12} ${h * 0.2} ${w * 0.4} ${h * 0.08} ${w * 0.5} ${h * 0.25}` +
        `C${w * 0.6} ${h * 0.02} ${w * 0.88} ${h * 0.12} ${w * 0.83} ${h * 0.4}` +
        `C${w * 1.02} ${h * 0.45} ${w * 0.98} ${h * 0.88} ${w * 0.75} ${h * 0.85}Z`,
    }]],
    outline: (w, h) => ellipsePoly(w / 2, h * 0.5, w * 0.46, h * 0.38),
    text: (w, h) => ({ x: w * 0.2, y: h * 0.3, w: w * 0.6, h: h * 0.5 }),
  },
  actor: {
    name: 'Person / actor', w: 60, h: 80, textBelow: true,
    draw: (w, h) => {
      const hh = h - 18; const r = Math.min(w, hh) * 0.16; const cx = w / 2;
      return [
        ['rect', { x: 0, y: 0, width: w, height: h, fill: 'transparent', stroke: 'none' }],
        ['path', {
          d: `M${cx} ${r * 2}m${-r} 0a${r} ${r} 0 1 0 ${r * 2} 0a${r} ${r} 0 1 0 ${-r * 2} 0` +
            `M${cx} ${r * 3}V${hh * 0.7}M${cx - w * 0.3} ${hh * 0.45}H${cx + w * 0.3}` +
            `M${cx} ${hh * 0.7}L${cx - w * 0.25} ${hh}M${cx} ${hh * 0.7}L${cx + w * 0.25} ${hh}`,
        }],
      ];
    },
    text: (w, h) => ({ x: -30, y: h - 18, w: w + 60, h: 18 }),
  },
  note: {
    name: 'Note', w: 170, h: 70, fill: 'var(--note)', align: 'left',
    draw: (w, h) => [
      ['path', { d: `M0 0H${w - 14}L${w} 14V${h}H0Z` }],
      ['path', { d: `M${w - 14} 0V14H${w}` }],
    ],
    text: (w, h) => ({ x: 8, y: 6, w: w - 24, h: h - 12 }),
  },
  annotation: {
    name: 'Annotation', w: 170, h: 50, fill: 'transparent', align: 'left', noLegend: true,
    draw: (w, h) => [
      ['rect', { x: 0, y: 0, width: w, height: h, fill: 'transparent', stroke: 'none' }],
      ['path', { d: `M12 0H0V${h}H12` }],
    ],
    text: (w, h) => ({ x: 8, y: 2, w: w - 10, h: h - 4 }),
  },
  text: {
    name: 'Text', w: 160, h: 30, fill: 'transparent', stroke: 'none', noLegend: true,
    draw: (w, h) => [['rect', { x: 0, y: 0, width: w, height: h, fill: 'transparent', stroke: 'none' }]],
    text: (w, h) => ({ x: 2, y: 2, w: w - 4, h: h - 4 }),
  },
  group: {
    name: 'Group / frame', w: 320, h: 220, fill: 'var(--group)', dash: '6 4', container: true,
    align: 'left', valign: 'top', noLegend: true,
    draw: (w, h) => [['rect', { x: 0, y: 0, width: w, height: h, rx: 6 }]],
    text: (w, h) => ({ x: 10, y: 6, w: w - 20, h: Math.min(h - 12, 40) }),
  },
};

export const SHAPE_ORDER = [
  'process', 'decision', 'terminator', 'document', 'multidoc', 'data', 'manualInput',
  'database', 'predefined', 'preparation', 'manualOp', 'delay', 'display', 'storedData',
  'rounded', 'ellipse', 'connector', 'offpage', 'merge', 'extract', 'hexagon', 'cloud',
  'actor', 'note', 'annotation', 'text', 'group',
];

export function shapeDef(key) {
  return SHAPES[key] || SHAPES.process;
}

export function shapeOutline(key, w, h) {
  const s = shapeDef(key);
  return s.outline ? s.outline(w, h) : rectPoly(w, h);
}

export function shapeTextBox(key, w, h) {
  const s = shapeDef(key);
  return s.text ? s.text(w, h) : { x: 8, y: 5, w: w - 16, h: h - 10 };
}
