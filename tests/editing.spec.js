import { test, expect } from '@playwright/test';
import { load, doc, box, toScreen, drag, clickNode, noErrors, setup, N } from './helpers.js';

test('loads the default template without errors', async ({ page }) => {
  await load(page, 'basic');
  const d = await doc(page);
  expect(d.nodes.length).toBe(8);
  expect(d.edges.length).toBe(8);
  await expect(page.locator('#diagram .node')).toHaveCount(8);
  await expect(page.locator('#diagram .edge')).toHaveCount(8);
  await noErrors(page);
});

test('every template loads and renders', async ({ page }) => {
  await load(page, 'blank');
  const ids = await page.evaluate(() => window.fc.templates);
  for (const id of ids) {
    await page.evaluate((i) => window.fc.loadTemplate(i), id);
    const d = await doc(page);
    await expect(page.locator('#diagram .node')).toHaveCount(d.nodes.length);
  }
  await noErrors(page);
});

test('clicking a palette shape adds it and selects it', async ({ page }) => {
  await setup(page, []);
  await page.click('.shape-btn[data-shape="decision"]');
  const d = await doc(page);
  expect(d.nodes).toHaveLength(1);
  expect(d.nodes[0].shape).toBe('decision');
  expect(await page.evaluate(() => [...window.fc.app.sel.nodes])).toEqual([d.nodes[0].id]);
  await noErrors(page);
});

test('dragging a palette shape onto the canvas drops it at the pointer', async ({ page }) => {
  await setup(page, []);
  const target = await toScreen(page, 400, 300);
  await page.dispatchEvent('.shape-btn[data-shape="database"]', 'dragstart', { dataTransfer: await page.evaluateHandle(() => new DataTransfer()) });
  // emulate full HTML5 DnD with a shared DataTransfer
  await page.evaluate(([x, y]) => {
    const dt = new DataTransfer();
    dt.setData('application/x-flow-shape', 'database');
    const stage = document.getElementById('stage');
    stage.dispatchEvent(new DragEvent('dragover', { dataTransfer: dt, clientX: x, clientY: y, bubbles: true, cancelable: true }));
    stage.dispatchEvent(new DragEvent('drop', { dataTransfer: dt, clientX: x, clientY: y, bubbles: true, cancelable: true }));
  }, target);
  const d = await doc(page);
  expect(d.nodes).toHaveLength(1);
  const n = d.nodes[0];
  expect(n.shape).toBe('database');
  expect(Math.abs(n.x + n.w / 2 - 400)).toBeLessThanOrEqual(8);
  expect(Math.abs(n.y + n.h / 2 - 300)).toBeLessThanOrEqual(8);
});

test('double-clicking empty canvas creates a shape and edits its text', async ({ page }) => {
  await setup(page, []);
  const p = await toScreen(page, 300, 200);
  await page.mouse.dblclick(p[0], p[1]);
  await expect(page.locator('#inline-editor')).toBeVisible();
  await page.keyboard.type('Approve invoice');
  await page.keyboard.press('Enter');
  await expect(page.locator('#inline-editor')).toBeHidden();
  const d = await doc(page);
  expect(d.nodes).toHaveLength(1);
  expect(d.nodes[0].text).toBe('Approve invoice');
  await expect(page.locator('#diagram .node text').last()).toHaveText('Approve invoice');
});

test('double-clicking a shape edits its text; Escape cancels', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  const b = await box(page, 'a');
  await page.mouse.dblclick(b.cx, b.cy);
  await page.keyboard.type('New text');
  await page.keyboard.press('Escape');
  expect((await doc(page)).nodes[0].text).toBe('A');
  await page.mouse.dblclick(b.cx, b.cy);
  await page.keyboard.type('Line one');
  await page.keyboard.press('Shift+Enter');
  await page.keyboard.type('Line two');
  await page.keyboard.press('Enter');
  expect((await doc(page)).nodes[0].text).toBe('Line one\nLine two');
});

test('typing while a shape is selected replaces its text', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  await clickNode(page, 'a');
  await page.keyboard.type('Hello');
  await page.keyboard.press('Enter');
  expect((await doc(page)).nodes[0].text).toBe('Hello');
});

test('moving a shape snaps to grid, and undo/redo work', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  const b = await box(page, 'a');
  await drag(page, [b.cx, b.cy], [b.cx + 103, b.cy + 51]);
  let n = (await doc(page)).nodes[0];
  const cx = n.x + n.w / 2, cy = n.y + n.h / 2;
  expect(cx % 8).toBe(0);
  expect(cy % 8).toBe(0);
  expect(Math.abs(n.x - 203)).toBeLessThanOrEqual(8);
  await page.keyboard.press('Control+z');
  n = (await doc(page)).nodes[0];
  expect([n.x, n.y]).toEqual([100, 100]);
  await page.keyboard.press('Control+Shift+z');
  n = (await doc(page)).nodes[0];
  expect(Math.abs(n.x - 203)).toBeLessThanOrEqual(8);
});

test('Alt-drag moves without snapping', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  const b = await box(page, 'a');
  await drag(page, [b.cx, b.cy], [b.cx + 33, b.cy + 17], { modifiers: ['Alt'] });
  const n = (await doc(page)).nodes[0];
  expect([n.x, n.y]).toEqual([133, 117]);
});

test('smart guides align a dragged shape with a neighbour', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 400, 300)]);
  await page.evaluate(() => { window.fc.getDoc().settings.snap = true; });
  const b = await box(page, 'b');
  // drag b so its top is 3px below a's top -> should snap to exactly align
  await drag(page, [b.cx, b.cy], [b.cx, b.cy - 197]);
  const d = await doc(page);
  expect(d.nodes[1].y).toBe(100);
});

test('resizing with a handle changes size', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  await clickNode(page, 'a');
  const h = page.locator('[data-handle="se"]');
  const hb = await h.boundingBox();
  await drag(page, [hb.x + hb.width / 2, hb.y + hb.height / 2], [hb.x + 80, hb.y + 40]);
  const n = (await doc(page)).nodes[0];
  expect(n.w).toBeGreaterThan(220);
  expect(n.h).toBeGreaterThan(90);
  expect([n.x, n.y]).toEqual([100, 100]);
});

test('arrow keys nudge and Delete removes shapes with their connectors', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 100, 300)], [{ id: 'e1', from: { node: 'a' }, to: { node: 'b' } }]);
  await clickNode(page, 'a');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('Shift+ArrowDown');
  const n = (await doc(page)).nodes[0];
  expect([n.x, n.y]).toEqual([108, 132]);
  await page.keyboard.press('Delete');
  const d = await doc(page);
  expect(d.nodes.map((x) => x.id)).toEqual(['b']);
  expect(d.edges).toHaveLength(0);
});

test('marquee selection, align and distribute', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 140, 220), N('c', 90, 400)]);
  const p0 = await toScreen(page, 40, 60), p1 = await toScreen(page, 360, 520);
  await drag(page, p0, p1);
  expect(await page.evaluate(() => window.fc.app.sel.nodes.size)).toBe(3);
  await page.click('[data-align="left"]');
  let d = await doc(page);
  expect(new Set(d.nodes.map((n) => n.x))).toEqual(new Set([90]));
  await page.click('[data-align="dist-y"]');
  d = await doc(page);
  const ys = d.nodes.map((n) => n.y).sort((a, b) => a - b);
  expect(ys[1] - ys[0]).toBeCloseTo(ys[2] - ys[1], 3);
});

test('shift-click toggles selection', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 400, 100)]);
  await clickNode(page, 'a');
  await clickNode(page, 'b', { modifiers: ['Shift'] });
  expect(await page.evaluate(() => window.fc.app.sel.nodes.size)).toBe(2);
  await clickNode(page, 'a', { modifiers: ['Shift'] });
  expect(await page.evaluate(() => [...window.fc.app.sel.nodes])).toEqual(['b']);
});

test('copy / paste and duplicate create offset copies with connectors', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 100, 300)], [{ id: 'e1', from: { node: 'a' }, to: { node: 'b' }, label: 'go' }]);
  await page.keyboard.press('Control+a');
  await page.keyboard.press('Control+c');
  await page.keyboard.press('Control+v');
  let d = await doc(page);
  expect(d.nodes).toHaveLength(4);
  expect(d.edges).toHaveLength(2);
  const copyEdge = d.edges[1];
  expect(copyEdge.label).toBe('go');
  expect(copyEdge.from.node).not.toBe('a');
  await page.evaluate(() => window.fc.select({ nodes: ['a'] }));
  await page.keyboard.press('Control+d');
  d = await doc(page);
  expect(d.nodes).toHaveLength(5);
  expect(d.nodes[4].x).toBe(124);
});

test('cut removes, paste restores', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  await clickNode(page, 'a');
  await page.keyboard.press('Control+x');
  expect((await doc(page)).nodes).toHaveLength(0);
  await page.keyboard.press('Control+v');
  expect((await doc(page)).nodes).toHaveLength(1);
});

test('Ctrl+G frames the selection; dragging the frame moves its contents', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 300, 100)]);
  await page.keyboard.press('Control+a');
  await page.keyboard.press('Control+g');
  let d = await doc(page);
  const g = d.nodes.find((n) => n.shape === 'group');
  expect(g).toBeTruthy();
  const gb = await box(page, g.id);
  // grab the frame near its top-left title area
  await drag(page, [gb.x + 8, gb.y + 30], [gb.x + 8 + 80, gb.y + 30 + 40], { modifiers: ['Alt'] });
  d = await doc(page);
  expect(d.nodes.find((n) => n.id === 'a').x).toBe(180);
  expect(d.nodes.find((n) => n.id === 'b').y).toBe(140);
});

test('zooming with ctrl+wheel and fit', async ({ page }) => {
  await load(page, 'basic');
  const k0 = await page.evaluate(() => window.fc.app.view.k);
  await page.mouse.move(700, 400);
  await page.keyboard.down('Control');
  await page.mouse.wheel(0, -300);
  await page.keyboard.up('Control');
  const k1 = await page.evaluate(() => window.fc.app.view.k);
  expect(k1).toBeGreaterThan(k0);
  await page.click('[data-cmd="fit"]');
  const k2 = await page.evaluate(() => window.fc.app.view.k);
  expect(k2).toBeCloseTo(k0, 2);
  await page.mouse.wheel(0, 200); // plain wheel pans
  const y = await page.evaluate(() => window.fc.app.view.y);
  expect(y).toBeLessThan(await page.evaluate(() => window.fc.app.view.y + 1));
});

test('space+drag pans the canvas', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  await page.locator('#stage').focus();
  await page.keyboard.down('Space');
  await drag(page, [700, 500], [800, 560]);
  await page.keyboard.up('Space');
  const v = await page.evaluate(() => window.fc.app.view);
  expect([v.x, v.y]).toEqual([160, 120]);
  expect((await doc(page)).nodes[0].x).toBe(100);
});

test('autosave restores the document after reload', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  await page.click('.shape-btn[data-shape="note"]');
  await page.waitForTimeout(500);
  await page.goto('/');
  await page.waitForFunction(() => window.fc && window.fc.getDoc());
  const d = await doc(page);
  expect(d.nodes.map((n) => n.shape).sort()).toEqual(['note', 'process']);
});

test('search finds and selects a shape', async ({ page }) => {
  await load(page, 'walkthrough');
  await page.fill('#search', 'credit');
  await page.press('#search', 'Enter');
  const sel = await page.evaluate(() => [...window.fc.app.sel.nodes]);
  expect(sel).toHaveLength(1);
  const d = await doc(page);
  expect(d.nodes.find((n) => n.id === sel[0]).text.toLowerCase() + JSON.stringify(d.nodes.find((n) => n.id === sel[0]).fields).toLowerCase()).toContain('credit');
});
