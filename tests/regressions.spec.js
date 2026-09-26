// Regression tests for bugs found by exploratory QA.
import { test, expect } from '@playwright/test';
import { load, doc, box, toScreen, drag, clickNode, noErrors, setup, N } from './helpers.js';

test('holding Space to pan does not replace the selected shape text', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  await clickNode(page, 'a');
  await page.keyboard.down('Space');
  await drag(page, [900, 600], [960, 640]);
  await page.keyboard.up('Space');
  await expect(page.locator('#inline-editor')).toBeHidden();
  expect((await doc(page)).nodes[0].text).toBe('A');
});

test('small shapes at low zoom can still be dragged (ports do not cover them)', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  await page.evaluate(() => { window.fc.app.view = { x: 60, y: 60, k: 0.3 }; window.fc.renderCanvas(); });
  const b = await box(page, 'a');
  await page.mouse.move(b.cx, b.cy);
  await drag(page, [b.cx, b.cy], [b.cx + 30, b.cy + 30]);
  const d = await doc(page);
  expect(d.nodes).toHaveLength(1);
  expect(d.edges).toHaveLength(0);
  expect(d.nodes[0].x).toBeGreaterThan(150);
});

test('keys pressed during a drag are ignored; Escape cancels the drag', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  const b = await box(page, 'a');
  await page.mouse.move(b.cx, b.cy);
  await page.mouse.down();
  await page.mouse.move(b.cx + 50, b.cy + 50, { steps: 5 });
  await page.keyboard.press('Delete');
  await page.keyboard.press('Control+z');
  expect((await doc(page)).nodes).toHaveLength(1);
  await page.keyboard.press('Escape');
  await page.mouse.move(b.cx + 120, b.cy + 50, { steps: 3 });
  await page.mouse.up();
  const d = await doc(page);
  expect([d.nodes[0].x, d.nodes[0].y]).toEqual([100, 100]);
  expect(await page.evaluate(() => window.fc.app.undoStack.length)).toBe(0);
});

test('highlight colour cannot inject script into preview / export', async ({ page }) => {
  await load(page, 'blank');
  const html = await page.evaluate(() => {
    window.fc.setDoc({ nodes: [{ id: 'a', shape: 'process', x: 0, y: 0, w: 100, h: 50, text: '<script>window.__x=1</script>', fields: { what: '<img src=x onerror=window.__y=1>' } }],
      settings: { highlightColor: 'red}</style><script>window.__pwn=1</script><style>' } });
    return window.fc.exportHTML();
  });
  expect(html).not.toContain('<script>window.__pwn');
  expect(html).not.toContain('<script>window.__x');
  await page.click('[data-cmd="mode-preview"]');
  const frame = page.frameLocator('#preview');
  await frame.locator('[data-node="a"]').click();
  await page.waitForTimeout(200);
  const pwned = await page.evaluate(() => {
    const w = document.getElementById('preview').contentWindow;
    return !!(w.__pwn || w.__x || w.__y);
  });
  expect(pwned).toBe(false);
});

test('deleting a lane (key or button) removes its shapes and closes the gap', async ({ page }) => {
  const lanes = { items: [{ id: 'l1', title: 'A', size: 240 }, { id: 'l2', title: 'B', size: 240 }, { id: 'l3', title: 'C', size: 240 }] };
  for (const how of ['key', 'button']) {
    await setup(page, [N('a', 40, 100), N('b', 280, 100), N('c', 520, 100)], [{ from: { node: 'a' }, to: { node: 'b' } }, { from: { node: 'a' }, to: { node: 'c' } }], { lanes });
    await page.evaluate(() => window.fc.select({ lane: 'l2' }));
    if (how === 'key') { await page.locator('#stage').focus(); await page.keyboard.press('Delete'); } else await page.click('#lane-delete');
    const d = await doc(page);
    expect(d.lanes.items.map((l) => l.id), how).toEqual(['l1', 'l3']);
    expect(d.nodes.map((n) => n.id), how).toEqual(['a', 'c']);
    expect(d.nodes[1].x, how).toBe(280);
    expect(d.edges, how).toHaveLength(1);
    await expect(page.locator('.toast')).toContainText('Ctrl+Z');
  }
});

test('lane Direction dropdown transposes shapes like Rotate lanes', async ({ page }) => {
  await setup(page, [N('a', 280, 100)], [], { lanes: { items: [{ id: 'l1', title: 'A', size: 240 }, { id: 'l2', title: 'B', size: 240 }] } });
  await page.locator('#stage').click({ position: { x: 900, y: 700 } });
  await page.selectOption('#lane-orient', 'horizontal');
  const d = await doc(page);
  expect(d.lanes.orientation).toBe('horizontal');
  const n = d.nodes[0];
  expect(n.y + n.h / 2).toBeGreaterThanOrEqual(240); // still in lane B
});

test('typing a lane width moves shapes in later lanes', async ({ page }) => {
  await setup(page, [N('a', 40, 100), N('b', 280, 100)], [], { lanes: { items: [{ id: 'l1', title: 'A', size: 240 }, { id: 'l2', title: 'B', size: 240 }] } });
  await page.evaluate(() => window.fc.select({ lane: 'l1' }));
  await page.fill('#lane-size', '300');
  const d = await doc(page);
  expect(d.nodes.find((n) => n.id === 'b').x).toBe(340);
  expect(d.nodes.find((n) => n.id === 'a').x).toBe(40);
});

test('parallel connectors between the same shapes are offset', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 500, 300)], [1, 2, 3].map((i) => ({ id: 'e' + i, from: { node: 'a' }, to: { node: 'b' }, label: 'L' + i })));
  const ds = await page.locator('#diagram .edge-line').evaluateAll((els) => els.map((e) => e.getAttribute('d')));
  expect(new Set(ds).size).toBe(3);
  const labels = await page.locator('#diagram .edge-label').evaluateAll((els) => els.map((e) => { const r = e.getBoundingClientRect(); return `${Math.round(r.x)},${Math.round(r.y)}`; }));
  expect(new Set(labels).size).toBe(3);
});

test('shapes grow to fit long text; long connector labels wrap', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 100, 400)], [{ id: 'e1', from: { node: 'a' }, to: { node: 'b' } }]);
  await clickNode(page, 'a');
  await page.fill('#insp-text', 'This is a much longer description of the step that will certainly need several lines to fit inside the box');
  const n = (await doc(page)).nodes[0];
  expect(n.h).toBeGreaterThan(60);
  const shape = await page.locator('#diagram [data-node="a"] .shape').boundingBox();
  const text = await page.locator('#diagram [data-node="a"] .label').boundingBox();
  expect(text.y).toBeGreaterThanOrEqual(shape.y - 1);
  expect(text.y + text.height).toBeLessThanOrEqual(shape.y + shape.height + 1);
  await page.evaluate(() => window.fc.select({ edges: ['e1'] }));
  await page.fill('#insp-label', 'a very long connector label that would otherwise run far across the whole diagram');
  const lb = await page.locator('#diagram .edge-label').boundingBox();
  expect(lb.width).toBeLessThan(200);
});

test('a connector can be dropped on a group frame border', async ({ page }) => {
  await setup(page, [N('a', 100, 100), { id: 'g', shape: 'group', x: 400, y: 60, w: 300, h: 200, text: 'Group' }]);
  const b = await box(page, 'a');
  await page.mouse.move(b.cx, b.cy);
  const port = await page.locator('circle.ov-port[data-id="a"][data-port="r"]').boundingBox();
  const border = await toScreen(page, 404, 160);
  await drag(page, [port.x + port.width / 2, port.y + port.height / 2], border);
  const d = await doc(page);
  expect(d.edges).toHaveLength(1);
  expect(d.edges[0].to.node).toBe('g');
  expect(d.nodes).toHaveLength(2);
});

test('a drag followed quickly by a click is not a double-click', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  const b = await box(page, 'a');
  await drag(page, [b.cx, b.cy], [b.cx + 2, b.cy + 5], { steps: 3 });
  const b2 = await box(page, 'a');
  await page.mouse.click(b2.cx, b2.cy);
  await expect(page.locator('#inline-editor')).toBeHidden();
});

test('elbow jogs sit mid-way, not tight against a shape', async ({ page }) => {
  await setup(page, [{ id: 'd', shape: 'decision', x: 300, y: 100, w: 120, h: 80, text: 'Q?' }, N('z', 300, 300)], [{ id: 'e1', from: { node: 'd' }, to: { node: 'z' } }]);
  const dd = await page.locator('#diagram .edge-line').getAttribute('d');
  const nums = dd.match(/-?\d+(\.\d+)?/g).map(Number);
  const pts = []; for (let i = 0; i < nums.length; i += 2) pts.push([nums[i], nums[i + 1]]);
  expect(pts.length).toBe(4);
  const jogY = pts[1][1];
  expect(jogY).toBeGreaterThan(200); // between 180 (diamond bottom) and 300 (target top), not at the stub
  expect(jogY).toBeLessThan(290);
});

test('self-loop connectors have room', async ({ page }) => {
  await setup(page, [N('a', 200, 200)], [{ id: 'e1', from: { node: 'a', side: 'r' }, to: { node: 'a', side: 't' } }]);
  const dd = await page.locator('#diagram .edge-line').getAttribute('d');
  const nums = dd.match(/-?\d+(\.\d+)?/g).map(Number);
  const xs = nums.filter((_, i) => i % 2 === 0), ys = nums.filter((_, i) => i % 2 === 1);
  expect(Math.max(...xs)).toBeGreaterThanOrEqual(360 + 30);
  expect(Math.min(...ys)).toBeLessThanOrEqual(200 - 30);
  await noErrors(page);
});

// --- from the UI-only usability test ---

const overlapping = (nodes) => nodes.some((a, i) => nodes.some((b, j) => j > i && a.shape !== 'group' && b.shape !== 'group' &&
  a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y));

test('rotating lanes keeps lane sizes and never overlaps shapes (both ways)', async ({ page }) => {
  await load(page, 'swim');
  const before = await doc(page);
  await page.click('[data-cmd="toggle-lane-orient"]');
  let d = await doc(page);
  expect(d.lanes.items.map((l) => l.size)).toEqual(before.lanes.items.map((l) => l.size));
  expect(overlapping(d.nodes)).toBe(false);
  await page.click('[data-cmd="toggle-lane-orient"]');
  d = await doc(page);
  expect(d.lanes.orientation).toBe('vertical');
  expect(overlapping(d.nodes)).toBe(false);
});

test('clicking a palette shape places it in free space', async ({ page }) => {
  await load(page, 'basic');
  for (let i = 0; i < 3; i++) await page.click('.shape-btn[data-shape="database"]');
  expect(overlapping((await doc(page)).nodes)).toBe(false);
});

test('quick-add after start/decision/manual input creates a process', async ({ page }) => {
  await setup(page, [{ id: 's', shape: 'terminator', x: 100, y: 100, w: 156, h: 40, text: 'Start' }]);
  await page.click('.shape-btn[data-shape="terminator"]'); // make "last used" a start/end
  await page.evaluate(() => window.fc.select({ nodes: ['s'] }));
  await page.locator('[data-quick="b"]').click();
  await page.keyboard.press('Escape');
  const d = await doc(page);
  expect(d.nodes.at(-1).shape).toBe('process');
});

test('changing a shape type keeps its size', async ({ page }) => {
  await setup(page, [{ id: 's', shape: 'terminator', x: 100, y: 100, w: 156, h: 40, text: 'Go' }]);
  await clickNode(page, 's');
  await page.selectOption('#insp-shape', 'manualInput');
  const n = (await doc(page)).nodes[0];
  expect([n.shape, n.w, n.h]).toEqual(['manualInput', 156, 40]);
});

test('image exports include the title and legend', async ({ page }) => {
  await load(page, 'basic');
  const svg = await page.evaluate(() => window.fc.exportSVG());
  expect(svg).toContain('>Basic flowchart<');
  expect(svg).toContain('>Decision<');
  expect(svg).not.toContain('@import');
});

test('dragging a segment keeps its label where it was', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 400, 300)], [{ id: 'e1', from: { node: 'a', side: 'r' }, to: { node: 'b', side: 't' }, label: 'go' }]);
  const l0 = await page.locator('#diagram .edge-label').boundingBox();
  await page.evaluate(() => window.fc.select({ edges: ['e1'] }));
  const seg = page.locator('.ov-seg[data-id="e1"]').last();
  const sb = await seg.boundingBox();
  await drag(page, [sb.x + sb.width / 2, sb.y + sb.height / 2], [sb.x + sb.width / 2 + 40, sb.y + sb.height / 2 + 40]);
  const l1 = await page.locator('#diagram .edge-label').boundingBox();
  expect(Math.hypot(l1.x - l0.x, l1.y - l0.y)).toBeLessThan(60);
});

// --- remaining usability items ---

test('hovering a quick-add arrow offers a shape picker', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  await clickNode(page, 'a');
  const q = await page.locator('[data-quick="r"]').boundingBox();
  await page.mouse.move(q.x + q.width / 2, q.y + q.height / 2);
  const picker = page.locator('#quick-picker');
  await expect(picker).toBeVisible();
  await picker.locator('[data-pick="decision"]').click();
  await expect(picker).toBeHidden();
  await page.keyboard.press('Escape');
  const d = await doc(page);
  expect(d.nodes).toHaveLength(2);
  expect(d.nodes[1].shape).toBe('decision');
  expect(d.edges[0]).toMatchObject({ from: { node: 'a', side: 'r' }, to: { node: d.nodes[1].id } });
});

test('the picker closes when the pointer moves away without choosing', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  await clickNode(page, 'a');
  const q = await page.locator('[data-quick="b"]').boundingBox();
  await page.mouse.move(q.x + q.width / 2, q.y + q.height / 2);
  await expect(page.locator('#quick-picker')).toBeVisible();
  const far = await toScreen(page, 700, 500);
  await page.mouse.move(far[0], far[1], { steps: 4 });
  await expect(page.locator('#quick-picker')).toBeHidden();
  expect((await doc(page)).nodes).toHaveLength(1);
});

test('dragging a shape to the canvas edge scrolls the view', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  const b = await box(page, 'a');
  const stage = await page.locator('#stage').boundingBox();
  const v0 = await page.evaluate(() => ({ ...window.fc.app.view }));
  await page.mouse.move(b.cx, b.cy);
  await page.mouse.down();
  await page.mouse.move(stage.x + stage.width - 8, b.cy, { steps: 10 });
  await page.waitForTimeout(500); // hold at the edge
  await page.mouse.up();
  const v1 = await page.evaluate(() => ({ ...window.fc.app.view }));
  expect(v1.x).toBeLessThan(v0.x - 40);
  const n = (await doc(page)).nodes[0];
  // the shape travelled further than the pointer did across the screen
  expect(n.x - 100).toBeGreaterThan(stage.x + stage.width - 8 - b.cx + 30);
});

test('connectors are clickable in preview and can carry details', async ({ page }) => {
  await setup(page, [N('a', 100, 100, { text: 'Approve', fields: { what: 'x' } }), N('b', 100, 300, { text: 'Pay' })],
    [{ id: 'e1', from: { node: 'a' }, to: { node: 'b' }, label: 'approved', fields: { what: 'Sent overnight by batch job' } }]);
  await page.evaluate(() => window.fc.select({ edges: ['e1'] }));
  await expect(page.locator('[data-edge-field="what"]')).toHaveValue('Sent overnight by batch job');
  await page.click('[data-cmd="mode-preview"]');
  const frame = page.frameLocator('#preview');
  await frame.locator('.edge[data-edge="e1"] .edge-hit').click({ force: true });
  await expect(frame.locator('#detail .id')).toHaveText('Connector');
  await expect(frame.locator('#detail h2')).toHaveText('approved');
  await expect(frame.locator('#detail')).toContainText('Approve');
  await expect(frame.locator('#detail')).toContainText('Sent overnight by batch job');
  // clicking the label works too
  await frame.locator('[data-node="a"]').click();
  await frame.locator('[data-label="e1"]').click();
  await expect(frame.locator('#detail h2')).toHaveText('approved');
});

test('phases grow when a shape is dropped across their end', async ({ page }) => {
  await setup(page, [N('a', 40, 100), N('b', 40, 400)], [], { lanes: { items: [{ id: 'l1', title: 'T', size: 240 }] }, phases: { items: [{ id: 'p1', title: 'One', size: 200 }, { id: 'p2', title: 'Two', size: 200 }] } });
  // phase One spans y 40..240; drag a down so it straddles 240
  const b = await box(page, 'a');
  await drag(page, [b.cx, b.cy], [b.cx, b.cy + 100], { modifiers: ['Alt'] });
  const d = await doc(page);
  const a = d.nodes.find((n) => n.id === 'a');
  expect(a.y).toBe(200);
  expect(40 + d.phases.items[0].size).toBeGreaterThanOrEqual(a.y + a.h);
  expect(d.nodes.find((n) => n.id === 'b').y).toBeGreaterThan(400); // pushed down with phase Two
  // undo restores both in one step
  await page.keyboard.press('Control+z');
  const u = await doc(page);
  expect([u.phases.items[0].size, u.nodes[0].y]).toEqual([200, 100]);
});

test('arrowheads keep their size when a connector is selected in preview', async ({ page }) => {
  await setup(page, [N('a', 100, 100, { fields: { what: 'x' } }), N('b', 400, 100)], [{ id: 'e1', from: { node: 'a', side: 'r' }, to: { node: 'b', side: 'l' } }]);
  await page.click('[data-cmd="mode-preview"]');
  const frame = page.frameLocator('#preview');
  await frame.locator('.edge[data-edge="e1"] .edge-hit').click({ force: true });
  await expect(frame.locator('.edge[data-edge="e1"]')).toHaveClass(/sel/);
  const after = await frame.locator('.edge[data-edge="e1"]').evaluate((g) => {
    const m = document.querySelector(g.querySelector('.edge-line').getAttribute('marker-end').slice(4, -1));
    return { units: m.getAttribute('markerUnits'), w: +m.getAttribute('markerWidth') };
  });
  expect(after.units).toBe('userSpaceOnUse');
  expect(after.w).toBeCloseTo(11.2, 1); // 7 × the 1.6px line width, independent of the drawn stroke
});
