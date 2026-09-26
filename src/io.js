import { renderSVG, themeCSS, legendHTML, fontFor, esc } from './render.js';
import { laneAt, phaseAt, shapeLabel, normalizeDoc } from './model.js';
import { center } from './geometry.js';

const safeJSON = (o) => JSON.stringify(o).replace(/</g, '\\u003c');

// The IBM Plex fonts as @font-face rules with the font files inlined: from the <style> in the
// single-file build, or from the local src/fonts.css stylesheet when running from the repository.
let fontCSSCache = null;
export function embeddedFontCSS() {
  if (typeof document === 'undefined') return '';
  const inline = document.getElementById('embedded-fonts')?.textContent;
  if (inline) return inline;
  if (fontCSSCache) return fontCSSCache;
  try {
    const sheet = document.getElementById('font-css')?.sheet;
    const css = sheet ? [...sheet.cssRules].map((r) => r.cssText).join('\n') : '';
    if (css) fontCSSCache = css;
    return css;
  } catch (_) { return ''; }
}
const usesEmbeddable = (doc) => (doc.settings.font || 'plex') === 'plex';
export const canEmbedFonts = (doc) => usesEmbeddable(doc) && !!embeddedFontCSS();

// Exported pages get the same rule as the editor: the browser may not fetch anything.
export const EXPORT_CSP = "default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; font-src data:; img-src data:; connect-src 'none'; object-src 'none'; base-uri 'none'; form-action 'none'";

export function detailData(doc) {
  const out = {};
  for (const n of doc.nodes) {
    const lane = laneAt(doc, ...center(n));
    const fields = doc.fields
      .map((f) => ({ label: f.label, value: n.fields?.[f.key] ?? '', highlight: !!f.highlight }))
      .filter((f) => String(f.value).trim());
    const phase = phaseAt(doc, ...center(n));
    out[n.id] = {
      tag: n.tag || '', text: n.text || '', lane: n.laneLabel || lane?.title || '',
      phase: phase?.title ? (doc.phases.label ? doc.phases.label + ' ' : '') + phase.title : '',
      shape: shapeLabel(doc, n.shape) + (n.style?.dash && !n.shape.startsWith('group') && doc.dashedNodeLabel ? ` (${doc.dashedNodeLabel})` : ''),
      fields,
    };
  }
  // connectors: from → to plus any detail fields
  const nameOf = (ep) => {
    if (!ep.node) return 'loose end';
    const n = doc.nodes.find((x) => x.id === ep.node);
    return n ? (n.tag ? n.tag + ' · ' : '') + (n.text || shapeLabel(doc, n.shape)).replace(/\n/g, ' ') : '';
  };
  for (const e of doc.edges) {
    const kind = doc.edgeKinds.find((k) => k.id === e.kind) || doc.edgeKinds[0];
    const fields = doc.fields
      .map((f) => ({ label: f.label, value: e.fields?.[f.key] ?? '', highlight: !!f.highlight }))
      .filter((f) => String(f.value).trim());
    out['edge:' + e.id] = {
      edge: true, tag: '', text: e.label ? e.label.replace(/\n/g, ' ') : `${nameOf(e.from)} → ${nameOf(e.to)}`,
      lane: '', phase: '', shape: kind?.name || 'Connector',
      fields: [{ label: 'From', value: nameOf(e.from) }, { label: 'To', value: nameOf(e.to) }, ...fields],
      own: fields.length,
    };
  }
  return out;
}

// Runs inside the exported page. Must be self-contained.
function viewerScript() {
  var data = JSON.parse(document.getElementById('fc-detail').textContent);
  var cfg = JSON.parse(document.getElementById('fc-config').textContent);
  var det = document.getElementById('detail');
  var esc = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var fmt = function (s) {
    var t = esc(s);
    // web addresses become links only when the author allowed it; clicking one is the reader's choice
    if (cfg.linkify) t = t.replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" target="_blank" rel="noopener noreferrer">$1</a>');
    return t.replace(/\n/g, '<br>');
  };
  var nodes = document.querySelectorAll('#chart .node[data-node], #chart .edge[data-edge]');
  var keyOf = function (g) { return g.hasAttribute('data-edge') ? 'edge:' + g.getAttribute('data-edge') : g.getAttribute('data-node'); };
  function select(id) {
    var n = data[id];
    if (!n || !det) return;
    nodes.forEach(function (g) { g.classList.toggle('sel', keyOf(g) === id); });
    var chips = '';
    if (n.lane) chips += '<span class="chip">' + esc(n.lane) + '</span>';
    if (n.phase) chips += '<span class="chip">' + esc(n.phase) + '</span>';
    if (n.shape) chips += '<span class="chip">' + esc(n.shape) + '</span>';
    var dl = n.fields.map(function (f) {
      return '<div class="' + (f.highlight ? 'prp' : '') + '"><dt>' + esc(f.label) + '</dt><dd>' + fmt(f.value) + '</dd></div>';
    }).join('');
    det.innerHTML = '<div class="id">' + esc(n.edge ? 'Connector' : (n.tag || cfg.noTag || '')) + '</div><h2>' + esc(n.text) + '</h2>' +
      '<div class="meta">' + chips + '</div>' + (dl ? '<dl>' + dl + '</dl>' : '<p class="empty">No details recorded for this step.</p>') +
      (cfg.hint ? '<p class="hint">' + esc(cfg.hint) + '</p>' : '');
    try { history.replaceState(null, '', '#' + encodeURIComponent(n.tag || id)); } catch (e) { /* sandboxed / srcdoc */ }
  }
  document.querySelectorAll('#chart [data-label]').forEach(function (g) {
    g.style.cursor = 'pointer';
    g.addEventListener('click', function () { select('edge:' + g.getAttribute('data-label')); });
  });
  nodes.forEach(function (g) {
    var id = keyOf(g);
    g.addEventListener('click', function () { select(id); });
    g.addEventListener('keydown', function (ev) { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); select(id); } });
  });
  var hash = decodeURIComponent(location.hash.slice(1));
  var initial = null;
  if (hash) for (var k in data) if ((data[k].tag && data[k].tag === hash) || k === hash) initial = k;
  initial = initial || cfg.initial;
  if (initial && data[initial]) select(initial);
  else if (det) det.innerHTML = '<p class="hint" style="margin:0">' + esc(cfg.hint || 'Select a shape to see its details.') + '</p>';
  // zoom controls
  var svg = document.querySelector('#chart svg');
  var chart = document.getElementById('chart');
  if (svg && chart) {
    var W = parseFloat(svg.getAttribute('width')), H = parseFloat(svg.getAttribute('height')), z = 1;
    var setZ = function (v) {
      z = Math.max(0.2, Math.min(3, v));
      svg.setAttribute('width', W * z); svg.setAttribute('height', H * z);
      var lab = document.getElementById('zoom-level'); if (lab) lab.textContent = Math.round(z * 100) + '%';
    };
    var fitZ = function () { return Math.min(1, (chart.clientWidth - 2) / W); };
    document.querySelectorAll('[data-zoom]').forEach(function (b) {
      b.addEventListener('click', function () {
        var k = b.getAttribute('data-zoom');
        setZ(k === 'fit' ? fitZ() : k === 'in' ? z * 1.25 : z / 1.25);
      });
    });
    if (fitZ() < 0.6) setZ(fitZ()); // small screens start fitted
  }
  var tb = document.getElementById('theme-toggle');
  if (tb) tb.addEventListener('click', function () {
    var root = document.documentElement;
    var cur = root.getAttribute('data-theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
    root.setAttribute('data-theme', cur === 'dark' ? 'light' : 'dark');
  });
}

export const VIEWER_CSS = `
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font-family:var(--ui);font-size:14px;line-height:1.5}
.wrap{max-width:1400px;margin:0 auto;padding-inline:16px;padding-block:24px 40px}
header.top{display:flex;flex-wrap:wrap;align-items:flex-end;justify-content:space-between;gap:12px 24px;margin-bottom:16px}
h1{font-family:var(--cond);font-weight:600;font-size:26px;letter-spacing:-0.01em;margin:0;text-wrap:balance}
.sub{color:var(--muted);margin:2px 0 0;font-size:13px}
.legend{display:flex;flex-wrap:wrap;gap:6px 14px;font-size:12px;color:var(--muted);align-items:center}
.legend span{display:inline-flex;align-items:center;gap:6px}
.legend svg{flex:none}
.layout{display:grid;grid-template-columns:minmax(0,1fr) 320px;gap:16px;align-items:start}
.layout.nodetail{grid-template-columns:minmax(0,1fr)}
@media (max-width:900px){.layout{grid-template-columns:minmax(0,1fr)}}
.chart{background:var(--panel);border:1px solid var(--line);border-radius:6px;overflow:auto}
.chart svg{display:block;max-width:none}
.chart-wrap{position:relative;min-width:0;margin-top:34px}
.zoom{position:absolute;right:0;top:-34px;z-index:2;display:flex;align-items:center;gap:2px;background:var(--panel);border:1px solid var(--line);border-radius:6px;padding:2px;font-size:12px;color:var(--muted);box-shadow:0 1px 4px rgba(0,0,0,.08)}
.zoom button{font:inherit;background:none;border:0;color:var(--ink);padding:2px 8px;border-radius:4px;cursor:pointer}
.zoom button:hover{background:var(--lane-alt)}
.zoom span{min-width:40px;text-align:center;font-variant-numeric:tabular-nums}
@media print{.zoom,.theme-btn{display:none}}
.detail{background:var(--panel);border:1px solid var(--line);border-radius:6px;padding:16px 18px;position:sticky;top:12px}
@media (max-width:900px){.detail{position:static}}
.detail .id{font-family:var(--mono);font-size:12px;color:var(--tag);font-weight:600;letter-spacing:.02em;min-height:1em}
.detail h2{font-family:var(--cond);font-size:20px;font-weight:600;margin:2px 0 10px;text-wrap:balance}
.meta{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:12px}
.chip{font-size:11px;text-transform:uppercase;letter-spacing:.06em;padding:2px 8px;border-radius:99px;background:var(--lane-alt);color:var(--muted);font-weight:500}
.detail dl{margin:0;display:grid;gap:10px}
.detail dt{font-size:11px;text-transform:uppercase;letter-spacing:.07em;color:var(--faint);font-weight:600}
.detail dd{margin:2px 0 0;overflow-wrap:anywhere}
.detail a{color:var(--flow)}
.prp{border-left:3px solid var(--hl);padding-left:10px}
.prp dt{color:var(--hl)}
.hint{margin-top:14px;font-size:12px;color:var(--faint)}
.empty{color:var(--faint);font-size:13px}
.node{cursor:pointer}
.node:focus{outline:none}
.node:focus-visible .shape{stroke:var(--sel);stroke-width:3}
.node.sel .shape{fill:var(--sel-fill);stroke:var(--sel);stroke-width:2.5}
.node:hover .shape{stroke:var(--sel)}
.edge{cursor:pointer}
.edge:focus{outline:none}
.edge:hover .edge-line,.edge:focus-visible .edge-line{stroke-width:2.6px}
.edge.sel .edge-line{stroke:var(--sel);stroke-width:3px}
@media (prefers-reduced-motion:no-preference){.node .shape{transition:fill .15s,stroke .15s}}
.notes{margin-top:16px;color:var(--muted);font-size:13px;max-width:75ch}
.notes p{margin:0 0 6px}
.theme-btn{font:inherit;font-size:12px;background:none;border:1px solid var(--line);color:var(--muted);border-radius:4px;padding:2px 8px;cursor:pointer}
`;

// Only allow plain colour values into CSS (they are user/document supplied).
export function safeColor(c, fallback) {
  const v = String(c ?? '').trim();
  return /^(#[0-9a-f]{3,8}|var\(--[a-z-]+\)|(rgb|hsl)a?\([\d\s.,%]+\)|[a-z]{3,20})$/i.test(v) ? v : fallback;
}

export function buildStandaloneHTML(doc, { preview = false, embedFonts = true } = {}) {
  const embed = embedFonts && canEmbedFonts(doc);
  const F = fontFor(doc);
  const { svg } = renderSVG(doc, { interactive: true, padding: 16 });
  const detail = detailData(doc);
  const hasDetail = doc.settings.showDetail !== false && Object.values(detail).some((d) => (d.edge ? d.own : d.fields.length));
  const notes = String(doc.notes || '').split(/\n+/).filter((s) => s.trim()).map((p) => `<p>${esc(p)}</p>`).join('');
  const theme = doc.settings.theme && doc.settings.theme !== 'auto' ? ` data-theme="${esc(doc.settings.theme)}"` : '';
  const legend = doc.settings.showLegend !== false ? legendHTML(doc) : '';
  const cfg = { hint: doc.settings.detailHint || '', initial: doc.settings.initialSelection || null, noTag: doc.settings.noTagLabel || '', linkify: doc.settings.linkify !== false };
  const vars = `:root{--ui:${F.ui};--cond:${F.text};--mono:${F.mono};--hl:${safeColor(doc.settings.highlightColor, 'var(--risk)')}}`;
  return `<!doctype html>
<html lang="en"${theme}><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="${EXPORT_CSP}">
<title>${esc(doc.title)}</title>
${embed ? `<style id="embedded-fonts">${embeddedFontCSS()}</style>` : ''}
<style>${themeCSS()}
${vars}
${VIEWER_CSS}</style></head><body>
<div class="wrap">
  <header class="top">
    <div><h1>${esc(doc.title)}</h1>${doc.subtitle ? `<p class="sub">${esc(doc.subtitle)}</p>` : ''}</div>
    <div class="legend" aria-label="Legend">${legend}${preview ? '' : '<button class="theme-btn" id="theme-toggle" type="button" aria-label="Toggle dark mode">◐</button>'}</div>
  </header>
  <div class="layout${hasDetail ? '' : ' nodetail'}">
    <div class="chart-wrap">
      <div class="zoom" role="group" aria-label="Zoom"><button type="button" data-zoom="out" aria-label="Zoom out">−</button><span id="zoom-level">100%</span><button type="button" data-zoom="in" aria-label="Zoom in">+</button><button type="button" data-zoom="fit">Fit</button></div>
      <div class="chart" id="chart" role="region" aria-label="Flowchart" tabindex="0">${svg}</div>
    </div>
    ${hasDetail ? '<aside class="detail" id="detail" aria-live="polite"></aside>' : ''}
  </div>
  ${notes ? `<div class="notes">${notes}</div>` : ''}
</div>
<script type="application/json" id="fc-detail">${safeJSON(detail)}</script>
<script type="application/json" id="fc-config">${safeJSON(cfg)}</script>
<script type="application/json" id="fc-doc">${safeJSON(doc)}</script>
<script>(${viewerScript.toString()})();</script>
</body></html>`;
}

// Standalone SVG: concrete colours, no external fonts (the font stacks fall back to local fonts),
// with the title, subtitle and legend above the diagram.
export function exportSVGString(doc, dark = false) {
  const { svg } = renderSVG(doc, { resolve: true, dark, padding: 24, header: true });
  return '<?xml version="1.0" encoding="UTF-8"?>\n' + svg;
}

export async function exportPNGBlob(doc, { scale = 2, dark = false } = {}) {
  let { svg, w, h } = renderSVG(doc, { resolve: true, dark, padding: 24, header: true });
  // images can't see the page's fonts, so carry the embedded ones inside the SVG
  if (canEmbedFonts(doc)) svg = svg.replace(/^<svg([^>]*)>/, (m) => `${m}<style>${embeddedFontCSS()}</style>`);
  const url = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
  const img = new Image();
  await new Promise((res, rej) => { img.onload = res; img.onerror = rej; img.src = url; });
  const canvas = document.createElement('canvas');
  canvas.width = Math.ceil(w * scale); canvas.height = Math.ceil(h * scale);
  const c = canvas.getContext('2d');
  c.scale(scale, scale);
  c.drawImage(img, 0, 0, w, h);
  return new Promise((res) => canvas.toBlob(res, 'image/png'));
}

export function download(name, data, type) {
  const blob = data instanceof Blob ? data : new Blob([data], { type });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

export function fileBase(doc) {
  return (doc.title || 'flowchart').replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '_') || 'flowchart';
}

// Accepts editor JSON or an exported HTML page (which embeds the source document).
export function parseDocumentText(text) {
  const t = text.trim();
  if (t.startsWith('{')) return normalizeDoc(JSON.parse(t));
  const m = t.match(/<script type="application\/json" id="fc-doc">([\s\S]*?)<\/script>/);
  if (m) return normalizeDoc(JSON.parse(m[1]));
  throw new Error('This file is not a flowchart document (expected .json or an exported .html).');
}
