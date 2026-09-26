import { test, expect } from '@playwright/test';
import { load, doc, box, toScreen, drag, clickNode, noErrors, setup, N } from './helpers.js';

async function hoverPort(page, id, side) {
  const b = await box(page, id);
  await page.mouse.move(b.cx, b.cy);
  const port = page.locator(`circle.ov-port[data-id="${id}"][data-port="${side}"]`);
  await expect(port).toHaveCount(1);
  const pb = await port.boundingBox();
  return [pb.x + pb.width / 2, pb.y + pb.height / 2];
}

test('drag from a port to another shape creates a connector', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 400, 300)]);
  const from = await hoverPort(page, 'a', 'r');
  const b = await box(page, 'b');
  await drag(page, from, [b.cx, b.cy]);
  const d = await doc(page);
  expect(d.edges).toHaveLength(1);
  expect(d.edges[0].from).toMatchObject({ node: 'a', side: 'r' });
  expect(d.edges[0].to).toMatchObject({ node: 'b', side: 'auto' });
  await expect(page.locator('#diagram .edge')).toHaveCount(1);
  await noErrors(page);
});

test('dropping onto a target port pins that side', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 400, 300)]);
  const from = await hoverPort(page, 'a', 'b');
  const tb = await box(page, 'b');
  await drag(page, from, [tb.x + 1, tb.cy]); // left port of b
  expect((await doc(page)).edges[0].to).toMatchObject({ node: 'b', side: 'l' });
});

test('dropping a connection on empty space creates a new connected shape', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  const from = await hoverPort(page, 'a', 'b');
  const target = await toScreen(page, 180, 320);
  await drag(page, from, target);
  await expect(page.locator('#inline-editor')).toBeVisible();
  await page.keyboard.type('Next step');
  await page.keyboard.press('Enter');
  const d = await doc(page);
  expect(d.nodes).toHaveLength(2);
  expect(d.nodes[1].text).toBe('Next step');
  expect(d.edges[0]).toMatchObject({ from: { node: 'a', side: 'b' }, to: { node: d.nodes[1].id } });
});

test('shift-drop leaves a loose connector end', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  const from = await hoverPort(page, 'a', 'r');
  const target = await toScreen(page, 400, 130);
  await drag(page, from, target, { modifiers: ['Shift'] });
  const d = await doc(page);
  expect(d.nodes).toHaveLength(1);
  expect(d.edges[0].to).toEqual({ x: 400, y: 130 });
});

test('quick-add arrows create a connected shape in that direction', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  await clickNode(page, 'a');
  await page.locator('[data-quick="r"]').click();
  await page.keyboard.type('Right');
  await page.keyboard.press('Enter');
  let d = await doc(page);
  expect(d.nodes).toHaveLength(2);
  const n = d.nodes[1];
  expect(n.x).toBeGreaterThan(260);
  expect(n.y + n.h / 2).toBe(130);
  expect(d.edges[0]).toMatchObject({ from: { node: 'a', side: 'r' }, to: { node: n.id, side: 'l' } });
  // Tab from the new shape adds another to its right
  await page.keyboard.press('Tab');
  await page.keyboard.press('Escape');
  d = await doc(page);
  expect(d.nodes).toHaveLength(3);
});

test('clicking a connector selects it; double-click edits the label', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 100, 300)], [{ id: 'e1', from: { node: 'a' }, to: { node: 'b' } }]);
  const p = await toScreen(page, 180, 200);
  await page.mouse.click(p[0], p[1]);
  expect(await page.evaluate(() => [...window.fc.app.sel.edges])).toEqual(['e1']);
  await expect(page.locator('#inspector h3')).toContainText('Connector');
  await page.mouse.dblclick(p[0], p[1]);
  await page.keyboard.type('Yes');
  await page.keyboard.press('Enter');
  expect((await doc(page)).edges[0].label).toBe('Yes');
  await expect(page.locator('#diagram .edge-label text')).toHaveText('Yes');
});

test('inspector edits connector label, type, arrows and routing', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 400, 300)], [{ id: 'e1', from: { node: 'a' }, to: { node: 'b' } }]);
  await page.evaluate(() => window.fc.select({ edges: ['e1'] }));
  await page.fill('#insp-label', 'feeds');
  await page.selectOption('#insp-kind', 'feed');
  await page.selectOption('#insp-start', 'diamond');
  await page.click('#insp-routing [data-value="curved"]');
  const e = (await doc(page)).edges[0];
  expect(e).toMatchObject({ label: 'feeds', kind: 'feed', routing: 'curved', style: { startArrow: 'diamond' } });
  const d = await page.locator('#diagram .edge-line').getAttribute('d');
  expect(d).toContain('C');
  expect(await page.locator('#diagram .edge-line').getAttribute('stroke-dasharray')).toBe('5 4');
  expect(await page.locator('#diagram .edge-line').getAttribute('marker-start')).toMatch(/mk-diamond/);
  // one undo reverts the routing change only
  await page.locator('#stage').focus();
  await page.keyboard.press('Control+z');
  expect((await doc(page)).edges[0].routing).toBe('orthogonal');
});

test('dragging an elbow segment reroutes it; reset route clears bends', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 400, 300)], [{ id: 'e1', from: { node: 'a', side: 'r' }, to: { node: 'b', side: 'l' } }]);
  await page.evaluate(() => window.fc.select({ edges: ['e1'] }));
  const seg = page.locator('.ov-seg[data-id="e1"]').nth(1);
  const sb = await seg.boundingBox();
  await drag(page, [sb.x + sb.width / 2, sb.y + sb.height / 2], [sb.x + sb.width / 2 + 60, sb.y + sb.height / 2]);
  let e = (await doc(page)).edges[0];
  expect(e.waypoints.length).toBeGreaterThanOrEqual(2);
  const xs = e.waypoints.map((p) => p[0]);
  expect(Math.max(...xs)).toBeGreaterThan(330);
  // path is still orthogonal
  const pts = await page.locator('#diagram .edge-line').getAttribute('d');
  const nums = pts.match(/-?\d+(\.\d+)?/g).map(Number);
  for (let i = 2; i < nums.length; i += 2) {
    const dx = nums[i] - nums[i - 2], dy = nums[i + 1] - nums[i - 1];
    expect(Math.abs(dx) < 0.01 || Math.abs(dy) < 0.01).toBe(true);
  }
  await page.click('#edge-reset');
  e = (await doc(page)).edges[0];
  expect(e.waypoints).toEqual([]);
});

test('dragging the first segment keeps the connector attached', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 400, 100)], [{ id: 'e1', from: { node: 'a', side: 'r' }, to: { node: 'b', side: 'l' } }]);
  await page.evaluate(() => window.fc.select({ edges: ['e1'] }));
  const seg = page.locator('.ov-seg[data-id="e1"]').first();
  const sb = await seg.boundingBox();
  await drag(page, [sb.x + sb.width / 2, sb.y + sb.height / 2], [sb.x + sb.width / 2, sb.y + sb.height / 2 + 80]);
  const d = await page.locator('#diagram .edge-line').getAttribute('d');
  const nums = d.match(/-?\d+(\.\d+)?/g).map(Number);
  expect([nums[0], nums[1]]).toEqual([260, 130]); // still leaves a's right port
  expect([nums[nums.length - 2], nums[nums.length - 1]]).toEqual([400, 130]);
  expect(nums.some((v, i) => i % 2 === 1 && Math.abs(v - 210) < 9)).toBe(true);
});

test('straight connectors: drag body adds a bend, double-click removes it', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 400, 300)], [{ id: 'e1', from: { node: 'a' }, to: { node: 'b' }, routing: 'straight' }]);
  await page.evaluate(() => window.fc.select({ edges: ['e1'] }));
  const add = page.locator('.ov-addwp[data-id="e1"]').first();
  const ab = await add.boundingBox();
  await drag(page, [ab.x + ab.width / 2, ab.y + ab.height / 2], [ab.x + 100, ab.y - 60]);
  let e = (await doc(page)).edges[0];
  expect(e.waypoints).toHaveLength(1);
  const wp = page.locator('.ov-wp[data-id="e1"]');
  const wb = await wp.boundingBox();
  await page.mouse.dblclick(wb.x + wb.width / 2, wb.y + wb.height / 2);
  e = (await doc(page)).edges[0];
  expect(e.waypoints).toHaveLength(0);
});

test('dragging an end handle reconnects the connector', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 400, 100), N('c', 400, 300)], [{ id: 'e1', from: { node: 'a' }, to: { node: 'b' } }]);
  await page.evaluate(() => window.fc.select({ edges: ['e1'] }));
  const ep = page.locator('.ov-ep[data-ep="to"]');
  const eb = await ep.boundingBox();
  const c = await box(page, 'c');
  await drag(page, [eb.x + eb.width / 2, eb.y + eb.height / 2], [c.cx, c.cy]);
  expect((await doc(page)).edges[0].to.node).toBe('c');
  // drag the end onto empty space -> loose end
  const eb2 = await page.locator('.ov-ep[data-ep="to"]').boundingBox();
  const empty = await toScreen(page, 700, 500);
  await drag(page, [eb2.x + eb2.width / 2, eb2.y + eb2.height / 2], empty);
  expect((await doc(page)).edges[0].to).toEqual({ x: 700, y: 500 });
});

test('dragging a label moves it along the connector', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 100, 400)], [{ id: 'e1', from: { node: 'a' }, to: { node: 'b' }, label: 'go' }]);
  const lb = await page.locator('#diagram .edge-label').boundingBox();
  await drag(page, [lb.x + lb.width / 2, lb.y + lb.height / 2], [lb.x + lb.width / 2, lb.y + lb.height / 2 + 60]);
  const e = (await doc(page)).edges[0];
  expect(e.labelPos).toBeGreaterThan(0.5);
  await page.evaluate(() => window.fc.select({ edges: ['e1'] }));
  await page.getByRole('button', { name: 'Reset label position' }).click();
  expect((await doc(page)).edges[0].labelPos).toBe(null);
});

test('reverse swaps the ends', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 100, 400)], [{ id: 'e1', from: { node: 'a' }, to: { node: 'b' } }]);
  await page.evaluate(() => window.fc.select({ edges: ['e1'] }));
  await page.click('#edge-reverse');
  const e = (await doc(page)).edges[0];
  expect([e.from.node, e.to.node]).toEqual(['b', 'a']);
});

test('connectors follow shapes when they move', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 100, 400)], [{ id: 'e1', from: { node: 'a' }, to: { node: 'b' } }]);
  const before = await page.locator('#diagram .edge-line').getAttribute('d');
  const b = await box(page, 'b');
  await drag(page, [b.cx, b.cy], [b.cx + 300, b.cy]);
  const after = await page.locator('#diagram .edge-line').getAttribute('d');
  expect(after).not.toBe(before);
  const nums = after.match(/-?\d+(\.\d+)?/g).map(Number);
  expect(nums[nums.length - 2]).toBeGreaterThan(350);
});

// Routing quality: auto-routed connectors should never pass through unrelated shapes.
test('auto routing avoids shapes in every template', async ({ page }) => {
  await load(page, 'blank');
  const ids = await page.evaluate(() => window.fc.templates);
  for (const id of ids) {
    await page.evaluate((i) => { window.fc.loadTemplate(i); window.fc.getDoc().settings.lineJumps = 'none'; window.fc.renderCanvas(); }, id);
    const hits = await page.evaluate(() => {
      const d = window.fc.getDoc();
      const out = [];
      const routes = window.fc.renderSVG ? null : null;
      document.querySelectorAll('#diagram .edge').forEach((g) => {
        const e = d.edges.find((x) => x.id === g.dataset.edge);
        const nums = g.querySelector('.edge-line').getAttribute('d').match(/-?\d+(\.\d+)?/g).map(Number);
        const pts = []; for (let i = 0; i < nums.length; i += 2) pts.push([nums[i], nums[i + 1]]);
        for (const n of d.nodes) {
          if (n.id === e.from.node || n.id === e.to.node || n.shape === 'group' || n.shape === 'text') continue;
          for (let i = 0; i < pts.length - 1; i++) {
            const [a, b] = [pts[i], pts[i + 1]];
            const x0 = Math.min(a[0], b[0]), x1 = Math.max(a[0], b[0]), y0 = Math.min(a[1], b[1]), y1 = Math.max(a[1], b[1]);
            if (x1 > n.x + 1 && x0 < n.x + n.w - 1 && y1 > n.y + 1 && y0 < n.y + n.h - 1) out.push(`${e.id} crosses ${n.id}`);
          }
        }
      });
      return out;
    });
    expect(hits, `template ${id}`).toEqual([]);
  }
});

test('auto routing detours around an obstacle', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('mid', 100, 250), N('b', 100, 400)], [{ id: 'e1', from: { node: 'a' }, to: { node: 'b' } }]);
  const d = await page.locator('#diagram .edge-line').getAttribute('d');
  const nums = d.match(/-?\d+(\.\d+)?/g).map(Number);
  const xs = nums.filter((_, i) => i % 2 === 0);
  expect(xs.some((x) => x < 100 || x > 260)).toBe(true);
});

test('side resize handles are usable on a selected shape (ports hidden)', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  await clickNode(page, 'a');
  const h = page.locator('[data-handle="e"]');
  const hb = await h.boundingBox();
  await page.mouse.move(hb.x + hb.width / 2, hb.y + hb.height / 2);
  await expect(page.locator('.ov-port')).toHaveCount(0);
  await drag(page, [hb.x + hb.width / 2, hb.y + hb.height / 2], [hb.x + 60, hb.y + hb.height / 2]);
  const n = (await doc(page)).nodes[0];
  expect(n.w).toBeGreaterThanOrEqual(208);
  expect(n.h).toBe(60);
});

test('dragging a quick-add arrow draws a connector to another shape', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 100, 400)]);
  await clickNode(page, 'a');
  const q = await page.locator('[data-quick="b"]').boundingBox();
  const b = await box(page, 'b');
  await drag(page, [q.x + q.width / 2, q.y + q.height / 2], [b.cx, b.cy]);
  const d = await doc(page);
  expect(d.nodes).toHaveLength(2);
  expect(d.edges[0]).toMatchObject({ from: { node: 'a', side: 'b' }, to: { node: 'b' } });
});

test('incoming and outgoing connectors do not share one port', async ({ page }) => {
  await load(page, 'swim');
  const pts = await page.evaluate(() => {
    const d = window.fc.getDoc();
    const ends = [];
    document.querySelectorAll('#diagram .edge').forEach((g) => {
      const nums = g.querySelector('.edge-line').getAttribute('d').match(/-?\d+(\.\d+)?/g).map(Number);
      ends.push({ id: g.dataset.edge, start: nums.slice(0, 2).join(','), end: nums.slice(-2).join(',') });
    });
    return ends;
  });
  const starts = new Set(pts.map((p) => p.start));
  for (const p of pts) expect(starts.has(p.end), `edge ${p.id} ends where another starts`).toBe(false);
});
