// Property-based checks of the connector router on random layouts.
import { test, expect } from '@playwright/test';
import { load } from './helpers.js';

test('random layouts: elbow routes are orthogonal, attached, and avoid their own shapes', async ({ page }) => {
  await load(page, 'blank');
  const failures = await page.evaluate(async () => {
    const G = await import('/src/geometry.js');
    const R = await import('/src/render.js');
    let seed = 42;
    const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
    const shapes = ['process', 'decision', 'terminator', 'document', 'data', 'database', 'ellipse', 'hexagon'];
    const out = [];
    for (let trial = 0; trial < 60; trial++) {
      const nodes = [];
      for (let i = 0; i < 6; i++) {
        const w = 80 + Math.round(rnd() * 120), h = 40 + Math.round(rnd() * 40);
        nodes.push({ id: 'n' + i, shape: shapes[Math.floor(rnd() * shapes.length)], x: Math.round(rnd() * 900), y: Math.round(rnd() * 700), w, h, text: '', style: {}, badges: [], fields: {} });
      }
      // drop overlapping shapes: routing between overlapping shapes is ill-defined
      const clean = nodes.filter((n, i) => !nodes.some((m, j) => j < i && n.x < m.x + m.w + 30 && n.x + n.w + 30 > m.x && n.y < m.y + m.h + 30 && n.y + n.h + 30 > m.y));
      const edges = [];
      for (let i = 0; i + 1 < clean.length; i++) edges.push({ id: 'e' + i, from: { node: clean[i].id, side: 'auto' }, to: { node: clean[i + 1].id, side: 'auto' }, routing: 'orthogonal', waypoints: [], style: {} });
      const doc = { nodes: clean, edges, settings: { stub: 18 }, lanes: { items: [] }, phases: { items: [] } };
      const routes = R.computeRoutes(doc);
      for (const e of edges) {
        const pts = routes.get(e.id).pts;
        const a = clean.find((n) => n.id === e.from.node), b = clean.find((n) => n.id === e.to.node);
        for (let i = 0; i < pts.length - 1; i++) {
          const dx = Math.abs(pts[i + 1][0] - pts[i][0]), dy = Math.abs(pts[i + 1][1] - pts[i][1]);
          if (dx > 0.5 && dy > 0.5) out.push(`trial ${trial} ${e.id}: diagonal segment`);
        }
        const onEdge = (n, p) => p[0] >= n.x - 1 && p[0] <= n.x + n.w + 1 && p[1] >= n.y - 1 && p[1] <= n.y + n.h + 1;
        if (!onEdge(a, pts[0])) out.push(`trial ${trial} ${e.id}: start not on source`);
        if (!onEdge(b, pts[pts.length - 1])) out.push(`trial ${trial} ${e.id}: end not on target`);
        // interior segments must not cut through source or target
        for (let i = 1; i < pts.length - 2; i++) {
          for (const n of [a, b]) {
            const p = pts[i], q = pts[i + 1];
            const x0 = Math.min(p[0], q[0]), x1 = Math.max(p[0], q[0]), y0 = Math.min(p[1], q[1]), y1 = Math.max(p[1], q[1]);
            if (x1 > n.x + 2 && x0 < n.x + n.w - 2 && y1 > n.y + 2 && y0 < n.y + n.h - 2) out.push(`trial ${trial} ${e.id}: cuts through ${n.id}`);
          }
        }
      }
    }
    return out;
  });
  expect(failures).toEqual([]);
});
