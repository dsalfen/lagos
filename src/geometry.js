import { shapeDef, shapeOutline } from './shapes.js';

export const DIRS = { t: [0, -1], r: [1, 0], b: [0, 1], l: [-1, 0] };
const SIDE_KEYS = ['t', 'r', 'b', 'l'];

const round = (v) => Math.round(v * 100) / 100;
const same = (a, b) => Math.abs(a - b) < 0.5;

// ---- basic helpers --------------------------------------------------------

export function center(n) {
  return [n.x + n.w / 2, n.y + n.h / 2];
}

export function boxesIntersect(a, b) {
  return a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
}

export function unionBox(boxes) {
  if (!boxes.length) return null;
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const b of boxes) {
    x0 = Math.min(x0, b.x); y0 = Math.min(y0, b.y);
    x1 = Math.max(x1, b.x + b.w); y1 = Math.max(y1, b.y + b.h);
  }
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

function raySegment(o, d, a, b) {
  // returns t along ray o + t*d where it crosses segment a-b, or null
  const ex = b[0] - a[0], ey = b[1] - a[1];
  const den = d[0] * ey - d[1] * ex;
  if (Math.abs(den) < 1e-9) return null;
  const t = ((a[0] - o[0]) * ey - (a[1] - o[1]) * ex) / den;
  const u = ((a[0] - o[0]) * d[1] - (a[1] - o[1]) * d[0]) / den;
  if (t < 0 || u < -1e-6 || u > 1 + 1e-6) return null;
  return t;
}

function rayPolygon(o, d, poly, far = false) {
  let best = null;
  for (let i = 0; i < poly.length; i++) {
    const t = raySegment(o, d, poly[i], poly[(i + 1) % poly.length]);
    if (t !== null && (best === null || (far ? t > best : t < best))) best = t;
  }
  return best;
}

// Point on a node outline for a given side (t/r/b/l) and offset along that side.
export function portPoint(n, side, offset = 0) {
  const poly = shapeOutline(n.shape, n.w, n.h).map((p) => [p[0] + n.x, p[1] + n.y]);
  const [cx, cy] = center(n);
  const far = 1e4;
  let o, d;
  const off = Math.max(-n.w / 2, Math.min(n.w / 2, offset));
  const offV = Math.max(-n.h / 2, Math.min(n.h / 2, offset));
  if (side === 't') { o = [cx + off, n.y - far]; d = [0, 1]; }
  else if (side === 'b') { o = [cx + off, n.y + n.h + far]; d = [0, -1]; }
  else if (side === 'l') { o = [n.x - far, cy + offV]; d = [1, 0]; }
  else { o = [n.x + n.w + far, cy + offV]; d = [-1, 0]; }
  const t = rayPolygon(o, d, poly);
  if (t === null) {
    if (side === 't') return [cx + off, n.y];
    if (side === 'b') return [cx + off, n.y + n.h];
    if (side === 'l') return [n.x, cy + offV];
    return [n.x + n.w, cy + offV];
  }
  return [round(o[0] + d[0] * t), round(o[1] + d[1] * t)];
}

// Where a line from the node centre toward `target` leaves the outline.
export function boundaryToward(n, target) {
  const [cx, cy] = center(n);
  let dx = target[0] - cx, dy = target[1] - cy;
  const len = Math.hypot(dx, dy);
  if (len < 1e-6) return [cx, cy];
  dx /= len; dy /= len;
  const poly = shapeOutline(n.shape, n.w, n.h).map((p) => [p[0] + n.x, p[1] + n.y]);
  const t = rayPolygon([cx, cy], [dx, dy], poly, true);
  if (t === null) return [cx, cy];
  return [round(cx + dx * t), round(cy + dy * t)];
}

export function sideFacing(n, pt) {
  const [cx, cy] = center(n);
  const dx = (pt[0] - cx) / Math.max(1, n.w), dy = (pt[1] - cy) / Math.max(1, n.h);
  if (Math.abs(dx) > Math.abs(dy)) return dx > 0 ? 'r' : 'l';
  return dy > 0 ? 'b' : 't';
}

export function nearestSide(n, pt) {
  let best = 't', bd = Infinity;
  for (const s of SIDE_KEYS) {
    const p = portPoint(n, s, 0);
    const d = Math.hypot(p[0] - pt[0], p[1] - pt[1]);
    if (d < bd) { bd = d; best = s; }
  }
  return best;
}

// ---- polyline helpers -----------------------------------------------------

export function simplify(pts) {
  const out = [];
  for (const p of pts) {
    const q = out[out.length - 1];
    if (q && same(q[0], p[0]) && same(q[1], p[1])) continue;
    out.push([round(p[0]), round(p[1])]);
  }
  for (let i = out.length - 2; i > 0; i--) {
    const a = out[i - 1], b = out[i], c = out[i + 1];
    const col = (same(a[0], b[0]) && same(b[0], c[0])) || (same(a[1], b[1]) && same(b[1], c[1]));
    if (!col) continue;
    // only drop when b lies between a and c (don't erase a doubling-back)
    const between = (Math.min(a[0], c[0]) - 0.5 <= b[0] && b[0] <= Math.max(a[0], c[0]) + 0.5 &&
      Math.min(a[1], c[1]) - 0.5 <= b[1] && b[1] <= Math.max(a[1], c[1]) + 0.5);
    if (between) out.splice(i, 1);
  }
  return out;
}

export function polyLength(pts) {
  let L = 0;
  for (let i = 0; i < pts.length - 1; i++) L += Math.hypot(pts[i + 1][0] - pts[i][0], pts[i + 1][1] - pts[i][1]);
  return L;
}

export function pointAt(pts, frac) {
  const total = polyLength(pts);
  let target = Math.max(0, Math.min(1, frac)) * total;
  for (let i = 0; i < pts.length - 1; i++) {
    const a = pts[i], b = pts[i + 1];
    const L = Math.hypot(b[0] - a[0], b[1] - a[1]);
    if (target <= L || i === pts.length - 2) {
      const f = L ? Math.min(1, target / L) : 0;
      return { pt: [a[0] + (b[0] - a[0]) * f, a[1] + (b[1] - a[1]) * f], seg: i, vertical: Math.abs(b[0] - a[0]) < Math.abs(b[1] - a[1]) };
    }
    target -= L;
  }
  return { pt: pts[0] || [0, 0], seg: 0, vertical: false };
}

// nearest point on polyline -> {frac, pt, dist, seg}
export function projectOnPolyline(pts, p) {
  const total = polyLength(pts) || 1;
  let acc = 0, best = { dist: Infinity, frac: 0, pt: pts[0], seg: 0 };
  for (let i = 0; i < pts.length - 1; i++) {
    const a = pts[i], b = pts[i + 1];
    const dx = b[0] - a[0], dy = b[1] - a[1];
    const L2 = dx * dx + dy * dy;
    let t = L2 ? ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / L2 : 0;
    t = Math.max(0, Math.min(1, t));
    const q = [a[0] + dx * t, a[1] + dy * t];
    const d = Math.hypot(p[0] - q[0], p[1] - q[1]);
    const L = Math.sqrt(L2);
    if (d < best.dist) best = { dist: d, frac: (acc + t * L) / total, pt: q, seg: i };
    acc += L;
  }
  return best;
}

export function longestSegmentMid(pts) {
  let best = -1, bi = 0;
  for (let i = 0; i < pts.length - 1; i++) {
    const L = Math.hypot(pts[i + 1][0] - pts[i][0], pts[i + 1][1] - pts[i][1]);
    if (L > best + 0.5) { best = L; bi = i; }
  }
  const a = pts[bi], b = pts[bi + 1] || a;
  return { pt: [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2], vertical: Math.abs(b[0] - a[0]) < Math.abs(b[1] - a[1]), length: best };
}

// ---- orthogonal router ----------------------------------------------------

function segHitsBox(a, b, r) {
  const x0 = Math.min(a[0], b[0]), x1 = Math.max(a[0], b[0]);
  const y0 = Math.min(a[1], b[1]), y1 = Math.max(a[1], b[1]);
  return x1 > r.x && x0 < r.x + r.w && y1 > r.y && y0 < r.y + r.h;
}

function scoreRoute(pts, obstacles, d1, d2) {
  // pts: [p1, s, ..., e, p2]; stubs p1->s and e->p2 are not tested against obstacles
  let len = 0, bends = 0, hits = 0, reverse = 0;
  let prev = null;
  for (let i = 0; i < pts.length - 1; i++) {
    const a = pts[i], b = pts[i + 1];
    const dx = b[0] - a[0], dy = b[1] - a[1];
    const L = Math.abs(dx) + Math.abs(dy);
    if (L < 0.5) continue;
    len += L;
    const dir = [Math.sign(dx), Math.sign(dy)];
    if (prev) {
      if (prev[0] !== dir[0] || prev[1] !== dir[1]) bends++;
      if (prev[0] === -dir[0] && prev[1] === -dir[1]) reverse++;
    }
    prev = dir;
    const stub = i === 0 || i === pts.length - 2;
    for (const o of obstacles) {
      if (stub && o.own) continue;
      if (segHitsBox(a, b, o)) hits += o.own ? 2 : 1;
    }
  }
  if (d2 && prev && (prev[0] !== -d2[0] || prev[1] !== -d2[1])) reverse++;
  // prefer jogs in the middle of a run rather than tight against a shape
  let tight = 0;
  const sp = simplify(pts);
  if (sp.length >= 4) {
    const first = Math.abs(sp[1][0] - sp[0][0]) + Math.abs(sp[1][1] - sp[0][1]);
    const lastL = Math.abs(sp[sp.length - 1][0] - sp[sp.length - 2][0]) + Math.abs(sp[sp.length - 1][1] - sp[sp.length - 2][1]);
    tight = Math.max(0, 36 - first) + Math.max(0, 36 - lastL);
  }
  return len + bends * 24 + hits * 3000 + reverse * 8000 + tight * 0.6;
}

function candidates(s, e, obstacles, margin, reach) {
  const xs = new Set([round((s[0] + e[0]) / 2), s[0], e[0]]);
  const ys = new Set([round((s[1] + e[1]) / 2), s[1], e[1]]);
  const bx0 = Math.min(s[0], e[0]) - reach, bx1 = Math.max(s[0], e[0]) + reach;
  const by0 = Math.min(s[1], e[1]) - reach, by1 = Math.max(s[1], e[1]) + reach;
  for (const o of obstacles) {
    if (o.x + o.w < bx0 || o.x > bx1 || o.y + o.h < by0 || o.y > by1) continue;
    xs.add(round(o.x - margin)); xs.add(round(o.x + o.w + margin));
    ys.add(round(o.y - margin)); ys.add(round(o.y + o.h + margin));
  }
  const out = [
    [s, [e[0], s[1]], e],
    [s, [s[0], e[1]], e],
  ];
  for (const x of xs) out.push([s, [x, s[1]], [x, e[1]], e]);
  for (const y of ys) out.push([s, [s[0], y], [e[0], y], e]);
  return { out, xs: [...xs], ys: [...ys] };
}

function orthoBetween(p1, d1, p2, d2, obstacles, stub, fallback = false) {
  // Facing ports that are close together: shorten the stubs so they don't cross.
  if (d1 && d2 && d1[0] === -d2[0] && d1[1] === -d2[1]) {
    const gap = (p2[0] - p1[0]) * d1[0] + (p2[1] - p1[1]) * d1[1];
    if (gap > 0 && gap < stub * 2) stub = gap / 2;
  }
  const s = d1 ? [p1[0] + d1[0] * stub, p1[1] + d1[1] * stub] : p1;
  const e = d2 ? [p2[0] + d2[0] * stub, p2[1] + d2[1] * stub] : p2;
  let best = null, bestScore = Infinity;
  const test = (mid) => {
    const pts = [p1, ...mid, p2];
    const sc = scoreRoute(pts, obstacles, d1, d2);
    if (sc < bestScore) { bestScore = sc; best = pts; }
  };
  if (!fallback) {
    // 1) simple L / Z shapes; 2) only if blocked, detours around the shapes in the way
    const mx = round((s[0] + e[0]) / 2), my = round((s[1] + e[1]) / 2);
    const simple = [[s, [e[0], s[1]], e], [s, [s[0], e[1]], e], [s, [mx, s[1]], [mx, e[1]], e], [s, [s[0], my], [e[0], my], e]];
    simple.forEach(test);
    if (bestScore >= 3000) {
      const blocking = obstacles.filter((o) => !o.own && simple.some((m) => {
        const pts = [p1, ...m, p2];
        for (let i = 0; i < pts.length - 1; i++) if (segHitsBox(pts[i], pts[i + 1], o)) return true;
        return false;
      }));
      candidates(s, e, blocking.length ? blocking : obstacles, stub, 40).out.forEach(test);
      if (bestScore >= 3000 && blocking.length) candidates(s, e, obstacles, stub, 40).out.forEach(test);
    }
  } else {
    const { out, xs: allX, ys: allY } = candidates(s, e, obstacles, stub, 260);
    out.forEach(test);
    if (bestScore >= 3000) {
      // limit the 5-segment search to the channels closest to either end
      const near = (vals, a, b) => vals.sort((u, v) => Math.min(Math.abs(u - a), Math.abs(u - b)) - Math.min(Math.abs(v - a), Math.abs(v - b))).slice(0, 10);
      const xs = near(allX, s[0], e[0]), ys = near(allY, s[1], e[1]);
      for (const x of xs) for (const y of ys) {
        test([s, [s[0], y], [x, y], [x, e[1]], e]);
        test([s, [x, s[1]], [x, y], [e[0], y], e]);
      }
    }
  }
  return { pts: simplify(best), score: bestScore };
}

// Route through user waypoints, inserting elbows so every segment is axis aligned.
function orthoThrough(p1, d1, wps, p2, d2) {
  const pts = [p1];
  let horiz = d1 ? d1[1] === 0 : null;
  const all = [...wps, p2];
  for (let i = 0; i < all.length; i++) {
    const cur = pts[pts.length - 1];
    const q = all[i];
    const last = i === all.length - 1;
    if (!same(cur[0], q[0]) && !same(cur[1], q[1])) {
      let h;
      if (last && d2) h = d2[1] !== 0; // arrive vertically if target side is t/b -> go horizontal first
      else h = horiz === null ? Math.abs(q[0] - cur[0]) > Math.abs(q[1] - cur[1]) : horiz;
      pts.push(h ? [q[0], cur[1]] : [cur[0], q[1]]);
      horiz = !h;
    } else {
      horiz = same(cur[1], q[1]);
    }
    pts.push(q);
  }
  return simplify(pts);
}

// ---- endpoint resolution ---------------------------------------------------

function resolveEnd(end, nodes) {
  if (end && end.node) {
    const n = nodes.get(end.node);
    if (n) return { node: n, side: end.side || 'auto', offset: end.offset || 0 };
  }
  return { pt: [end?.x ?? 0, end?.y ?? 0] };
}

function refPoint(r) {
  return r.node ? center(r.node) : r.pt;
}

// Penalty for running along or across other connectors' segments.
function edgePenalty(pts, others) {
  if (!others || !others.length) return 0;
  let p = 0;
  for (let i = 0; i < pts.length - 1; i++) {
    const a = pts[i], b = pts[i + 1];
    const h = Math.abs(a[1] - b[1]) < 0.5;
    const lo = h ? Math.min(a[0], b[0]) : Math.min(a[1], b[1]);
    const hi = h ? Math.max(a[0], b[0]) : Math.max(a[1], b[1]);
    for (const [c, d] of others) {
      const h2 = Math.abs(c[1] - d[1]) < 0.5;
      if (h === h2) {
        // collinear overlap
        if (h ? Math.abs(a[1] - c[1]) < 3 : Math.abs(a[0] - c[0]) < 3) {
          const lo2 = h ? Math.min(c[0], d[0]) : Math.min(c[1], d[1]);
          const hi2 = h ? Math.max(c[0], d[0]) : Math.max(c[1], d[1]);
          if (Math.min(hi, hi2) - Math.max(lo, lo2) > 2) p += 45;
        }
      } else {
        // perpendicular crossing (strictly inside both segments)
        const [hx0, hx1, hy] = h ? [lo, hi, a[1]] : [Math.min(c[0], d[0]), Math.max(c[0], d[0]), c[1]];
        const [vy0, vy1, vx] = h ? [Math.min(c[1], d[1]), Math.max(c[1], d[1]), c[0]] : [lo, hi, a[0]];
        if (vx > hx0 + 1 && vx < hx1 - 1 && hy > vy0 + 1 && hy < vy1 - 1) p += 18;
      }
    }
  }
  return p;
}

export function routeEdge(edge, nodes, obstacleList, settings = {}, ctx = {}) {
  const used = ctx.used, others = ctx.others;
  let stub = settings.stub ?? 18;
  if (edge.from?.node && edge.from.node === edge.to?.node) stub = Math.max(stub, 32);
  const A = resolveEnd(edge.from, nodes);
  const B = resolveEnd(edge.to, nodes);
  const wps = (edge.waypoints || []).map((p) => [p[0], p[1]]);
  const routing = edge.routing || 'orthogonal';
  const firstRef = wps.length ? wps[0] : refPoint(B);
  const lastRef = wps.length ? wps[wps.length - 1] : refPoint(A);

  const sidesFor = (R, ref) => {
    if (!R.node) return [null];
    if (R.side !== 'auto') return [R.side];
    if (routing !== 'orthogonal' || wps.length) return [sideFacing(R.node, ref)];
    return SIDE_KEYS;
  };
  const endPoint = (R, side, ref) => {
    if (!R.node) return R.pt;
    if (routing === 'straight' && R.side === 'auto') return boundaryToward(R.node, ref);
    return portPoint(R.node, side, R.offset);
  };

  let pts, s1 = null, s2 = null;
  if (routing === 'orthogonal') {
    // Only shapes and connectors near this connector can affect its route.
    const pad = 320;
    const pts0 = [refPoint(A), refPoint(B), ...wps];
    const rx0 = Math.min(...pts0.map((p) => p[0])) - pad, rx1 = Math.max(...pts0.map((p) => p[0])) + pad;
    const ry0 = Math.min(...pts0.map((p) => p[1])) - pad, ry1 = Math.max(...pts0.map((p) => p[1])) + pad;
    const near = (x0, y0, x1, y1) => x1 >= rx0 && x0 <= rx1 && y1 >= ry0 && y0 <= ry1;
    const obstacles = [];
    for (const o of obstacleList) {
      if (near(o.x, o.y, o.x + o.w, o.y + o.h)) obstacles.push({ ...o, own: o.id === A.node?.id || o.id === B.node?.id });
    }
    const nearOthers = others && others.filter(([c, d]) => near(Math.min(c[0], d[0]), Math.min(c[1], d[1]), Math.max(c[0], d[0]), Math.max(c[1], d[1])));
    if (wps.length) {
      s1 = sidesFor(A, firstRef)[0];
      s2 = sidesFor(B, lastRef)[0];
      pts = orthoThrough(endPoint(A, s1, firstRef), s1 && DIRS[s1], wps, endPoint(B, s2, lastRef), s2 && DIRS[s2]);
    } else {
      let best = null;
      const sa = sidesFor(A, firstRef), sb = sidesFor(B, lastRef);
      // Rank side pairs by a cheap estimate and only route the most promising ones.
      let pairs = [];
      for (const a of sa) for (const b of sb) {
        const p1 = endPoint(A, a, firstRef), p2 = endPoint(B, b, lastRef);
        const d1 = a && DIRS[a], d2 = b && DIRS[b];
        const q1 = d1 ? [p1[0] + d1[0] * stub, p1[1] + d1[1] * stub] : p1;
        const q2 = d2 ? [p2[0] + d2[0] * stub, p2[1] + d2[1] * stub] : p2;
        let est = Math.abs(q1[0] - q2[0]) + Math.abs(q1[1] - q2[1]);
        if (d1 && (q2[0] - q1[0]) * d1[0] + (q2[1] - q1[1]) * d1[1] < 0) est += 120; // leaving away from the target
        if (d2 && (q1[0] - q2[0]) * d2[0] + (q1[1] - q2[1]) * d2[1] < 0) est += 120;
        pairs.push({ a, b, p1, p2, d1, d2, est });
      }
      pairs.sort((x, y) => x.est - y.est);
      if (pairs.length > 6) pairs = pairs.slice(0, 6);
      let results = pairs.map((q) => ({ q, r: orthoBetween(q.p1, q.d1, q.p2, q.d2, obstacles, stub) }));
      if (results.every(({ r }) => r.score >= 3000)) {
        results = pairs.map((q) => ({ q, r: orthoBetween(q.p1, q.d1, q.p2, q.d2, obstacles, stub, true) }));
      }
      for (const { q, r } of results) {
        const { a, b } = q;
        // mild preference for the "natural" facing sides
        let sc = r.score;
        if (A.node && a !== sideFacing(A.node, refPoint(B))) sc += 14;
        if (B.node && b !== sideFacing(B.node, refPoint(A))) sc += 30; // enter the side that faces the source
        // avoid mixing incoming and outgoing connectors on one side of a shape
        if (used) {
          const ua = A.node && used.get(A.node.id + '|' + a);
          const ub = B.node && used.get(B.node.id + '|' + b);
          if (ua?.in) sc += 60;
          if (ub?.out) sc += 60;
          if (ua?.out) sc += A.node.shape === 'decision' ? 70 : 20; // branches should leave by different sides
          if (ub?.in) sc += 25; // prefer a free side over merging into another incoming line
        }
        sc += edgePenalty(r.pts, nearOthers);
        if (!best || sc < best.sc) best = { sc, pts: r.pts, a, b };
      }
      pts = best.pts; s1 = best.a; s2 = best.b;
    }
  } else {
    s1 = sidesFor(A, firstRef)[0];
    s2 = sidesFor(B, lastRef)[0];
    pts = [endPoint(A, s1, firstRef), ...wps, endPoint(B, s2, lastRef)];
  }
  return { pts, routing, s1, s2, curve: routing === 'curved' };
}

// ---- path strings ------------------------------------------------------------

export function pathD(route, cornerRadius = 0) {
  const pts = route.pts;
  if (pts.length < 2) return '';
  if (route.routing === 'curved') {
    if (pts.length === 2) {
      const [a, b] = pts;
      const L = Math.max(40, Math.hypot(b[0] - a[0], b[1] - a[1]) / 2.5);
      const d1 = route.s1 ? DIRS[route.s1] : [(b[0] - a[0]) / (L * 2.5), (b[1] - a[1]) / (L * 2.5)];
      const d2 = route.s2 ? DIRS[route.s2] : [(a[0] - b[0]) / (L * 2.5), (a[1] - b[1]) / (L * 2.5)];
      return `M${a[0]} ${a[1]}C${round(a[0] + d1[0] * L)} ${round(a[1] + d1[1] * L)} ${round(b[0] + d2[0] * L)} ${round(b[1] + d2[1] * L)} ${b[0]} ${b[1]}`;
    }
    // Catmull-Rom through all points
    let d = `M${pts[0][0]} ${pts[0][1]}`;
    for (let i = 0; i < pts.length - 1; i++) {
      const p0 = pts[i - 1] || pts[i], p1 = pts[i], p2 = pts[i + 1], p3 = pts[i + 2] || p2;
      const c1 = [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6];
      const c2 = [p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6];
      d += `C${round(c1[0])} ${round(c1[1])} ${round(c2[0])} ${round(c2[1])} ${p2[0]} ${p2[1]}`;
    }
    return d;
  }
  if (!cornerRadius || pts.length < 3) return 'M' + pts.map((p) => p.join(' ')).join('L');
  let d = `M${pts[0][0]} ${pts[0][1]}`;
  for (let i = 1; i < pts.length - 1; i++) {
    const a = pts[i - 1], b = pts[i], c = pts[i + 1];
    const l1 = Math.hypot(b[0] - a[0], b[1] - a[1]), l2 = Math.hypot(c[0] - b[0], c[1] - b[1]);
    const r = Math.min(cornerRadius, l1 / 2, l2 / 2);
    const p = [b[0] - ((b[0] - a[0]) / l1) * r, b[1] - ((b[1] - a[1]) / l1) * r];
    const q = [b[0] + ((c[0] - b[0]) / l2) * r, b[1] + ((c[1] - b[1]) / l2) * r];
    d += `L${round(p[0])} ${round(p[1])}Q${b[0]} ${b[1]} ${round(q[0])} ${round(q[1])}`;
  }
  const z = pts[pts.length - 1];
  return d + `L${z[0]} ${z[1]}`;
}

// Sampled polyline for curves (used for hit tests and label placement)
export function routePolyline(route) {
  if (route.routing !== 'curved') return route.pts;
  const d = pathD(route);
  const nums = d.match(/-?\d+(\.\d+)?/g).map(Number);
  const out = [[nums[0], nums[1]]];
  for (let i = 2; i + 5 < nums.length; i += 6) {
    const p0 = out[out.length - 1];
    const c1 = [nums[i], nums[i + 1]], c2 = [nums[i + 2], nums[i + 3]], p3 = [nums[i + 4], nums[i + 5]];
    for (let k = 1; k <= 16; k++) {
      const t = k / 16, u = 1 - t;
      out.push([
        u * u * u * p0[0] + 3 * u * u * t * c1[0] + 3 * u * t * t * c2[0] + t * t * t * p3[0],
        u * u * u * p0[1] + 3 * u * u * t * c1[1] + 3 * u * t * t * c2[1] + t * t * t * p3[1],
      ]);
    }
  }
  return out;
}

export function obstaclesFor(doc) {
  return doc.nodes
    .filter((n) => !shapeDef(n.shape).container && n.shape !== 'text')
    .map((n) => ({ id: n.id, x: n.x - 2, y: n.y - 2, w: n.w + 4, h: n.h + 4 }));
}
