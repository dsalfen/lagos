import { test, expect } from '@playwright/test';
import { load, doc, box, toScreen, drag, clickNode, noErrors, setup, N } from './helpers.js';

test('inspector edits shape text, tag, type, geometry and details', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  await clickNode(page, 'a');
  await page.fill('#insp-text', 'Post journal');
  await page.fill('#insp-tag', 'JE-01');
  await page.selectOption('#insp-shape', 'manualInput');
  await page.fill('[data-geom="w"]', '220');
  await page.fill('[data-field="what"]', 'Keyed in manually.');
  const n = (await doc(page)).nodes[0];
  expect(n).toMatchObject({ text: 'Post journal', tag: 'JE-01', shape: 'manualInput', w: 220, fields: { what: 'Keyed in manually.' } });
  await expect(page.locator('#diagram .node .tag')).toHaveText('JE-01');
  await expect(page.locator('#diagram .node.shape-manualInput')).toHaveCount(1);
  await noErrors(page);
});

test('typing in the inspector is a single undo step', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  await clickNode(page, 'a');
  await page.fill('#insp-text', '');
  await page.locator('#insp-text').pressSequentially('abc');
  await page.locator('#stage').focus();
  expect((await doc(page)).nodes[0].text).toBe('abc');
  await page.keyboard.press('Control+z');
  expect((await doc(page)).nodes[0].text).toBe('A');
  await page.keyboard.press('Control+Shift+z');
  expect((await doc(page)).nodes[0].text).toBe('abc');
});

test('style controls: fill, border, dash, bold, alignment, badges', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  await clickNode(page, 'a');
  await page.click('[data-color-field="fill"] [data-swatch="#FFF1CC"]');
  await page.click('[data-color-field="stroke"] [data-swatch="var(--risk)"]');
  await page.locator('select[aria-label="Border style"]').selectOption('6 4');
  await page.getByRole('button', { name: 'B', exact: true }).click();
  await page.click('[data-value="left"]');
  await page.click('[data-badge-toggle="risk"]');
  const n = (await doc(page)).nodes[0];
  expect(n.style).toMatchObject({ fill: '#FFF1CC', stroke: 'var(--risk)', dash: '6 4', bold: true, align: 'left' });
  expect(n.badges).toEqual(['risk']);
  const shape = page.locator('#diagram .node .shape');
  await expect(shape).toHaveAttribute('fill', '#FFF1CC');
  await expect(shape).toHaveAttribute('stroke-dasharray', '6 4');
  await expect(page.locator('#diagram .node .badge')).toHaveCount(1);
  await expect(page.locator('#diagram .node .label text').last()).toHaveAttribute('text-anchor', 'start');
});

test('multi-select style applies to all selected shapes', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 400, 100)]);
  await page.keyboard.press('Control+a');
  await page.click('[data-color-field="fill"] [data-swatch="#E3F4E6"]');
  const d = await doc(page);
  expect(d.nodes.every((n) => n.style.fill === '#E3F4E6')).toBe(true);
});

test('number tags in reading order', async ({ page }) => {
  await setup(page, [N('a', 100, 300), N('b', 100, 100), N('c', 400, 100)]);
  await page.keyboard.press('Control+a');
  await page.locator('input[aria-label="Tag prefix"]').fill('P-');
  await page.click('#number-apply');
  const d = await doc(page);
  const tag = (id) => d.nodes.find((n) => n.id === id).tag;
  expect([tag('b'), tag('c'), tag('a')]).toEqual(['P-01', 'P-02', 'P-03']);
});

test('document settings: title, fields, connector types and badges', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  await page.locator('#stage').click({ position: { x: 700, y: 600 } });
  await page.fill('#doc-title-insp', 'My process');
  await expect(page.locator('#doc-title')).toHaveValue('My process');
  await page.getByRole('button', { name: '+ Add field' }).click();
  await page.getByRole('button', { name: '+ Add type' }).click();
  await page.getByRole('button', { name: '+ Add badge' }).click();
  const d = await doc(page);
  expect(d.title).toBe('My process');
  expect(d.fields.at(-1).label).toBe('New field');
  expect(d.edgeKinds).toHaveLength(3);
  expect(d.badgeKinds).toHaveLength(3); // risk, control + the new one
  // title box in the toolbar also edits
  await page.fill('#doc-title', 'Renamed');
  expect((await doc(page)).title).toBe('Renamed');
});

test('toggling tags hides tag lines', async ({ page }) => {
  await setup(page, [N('a', 100, 100, { tag: 'X-1' })]);
  await expect(page.locator('#diagram .tag')).toHaveCount(1);
  await page.locator('#stage').click({ position: { x: 700, y: 600 } });
  await page.click('#set-tags');
  await expect(page.locator('#diagram .tag')).toHaveCount(0);
});

test('lanes: add, rename, resize by dragging the border, reorder, delete', async ({ page }) => {
  await setup(page, [N('a', 40, 100), N('b', 280, 100)]);
  await page.click('[data-cmd="add-lane"]');
  await page.click('[data-cmd="add-lane"]');
  let d = await doc(page);
  expect(d.lanes.items).toHaveLength(2);
  await page.fill('#lane-title', 'Finance');
  await expect(page.locator('#diagram .lane').nth(1)).toContainText('FINANCE');
  // drag the border between lane 1 and 2 (x=240) to the right by 80
  const p = await toScreen(page, 240, 300);
  await drag(page, p, [p[0] + 80, p[1]]);
  d = await doc(page);
  expect(d.lanes.items[0].size).toBe(320);
  expect(d.nodes.find((n) => n.id === 'b').x).toBe(360); // moved with its lane
  expect(d.nodes.find((n) => n.id === 'a').x).toBe(40);
  // select first lane header and move it right
  const hp = await toScreen(page, 100, 20);
  await page.mouse.click(hp[0], hp[1]);
  await expect(page.locator('#inspector h3')).toContainText('Lane');
  await page.click('#lane-next');
  d = await doc(page);
  expect(d.lanes.items[1].title).toBe('Lane 1');
  expect(d.nodes.find((n) => n.id === 'a').x).toBeGreaterThan(240);
  await page.click('#lane-delete');
  expect((await doc(page)).lanes.items).toHaveLength(1);
});

test('double-clicking a lane header renames it', async ({ page }) => {
  await setup(page, [], [], { lanes: { items: [{ id: 'l1', title: 'Old', size: 240 }] } });
  const hp = await toScreen(page, 120, 20);
  await page.mouse.dblclick(hp[0], hp[1]);
  await page.keyboard.type('Sales');
  await page.keyboard.press('Enter');
  expect((await doc(page)).lanes.items[0].title).toBe('Sales');
});

test('rotating lanes switches orientation and keeps shapes in their lane', async ({ page }) => {
  await setup(page, [N('a', 280, 100)], [], { lanes: { items: [{ id: 'l1', title: 'A', size: 240 }, { id: 'l2', title: 'B', size: 240 }] } });
  await page.click('[data-cmd="toggle-lane-orient"]');
  const d = await doc(page);
  expect(d.lanes.orientation).toBe('horizontal');
  const n = d.nodes[0];
  const cy = n.y + n.h / 2;
  let pos = 0; let lane = null;
  for (const l of d.lanes.items) { if (cy >= pos && cy < pos + l.size) lane = l.id; pos += l.size; }
  expect(lane).toBe('l2');
});

test('preview mode shows the interactive viewer with details', async ({ page }) => {
  await load(page, 'walkthrough');
  await page.click('[data-cmd="mode-preview"]');
  const frame = page.frameLocator('#preview');
  await expect(frame.locator('h1')).toHaveText('Order-to-cash walkthrough (example)');
  await expect(frame.locator('#detail h2')).toHaveText('Within credit limit?');
  await frame.locator('[data-node="OTC-10"]').click();
  await expect(frame.locator('#detail .id')).toHaveText('OTC-10');
  await expect(frame.locator('#detail .prp dd').last()).toContainText('wrong customer');
  await expect(frame.locator('#detail .prp dt').first()).toHaveText('Control (C3)');
  await expect(frame.locator('.legend')).toContainText('Risk point');
  await expect(page.locator('#palette')).toBeHidden();
  await page.click('[data-cmd="mode-edit"]');
  await expect(page.locator('#palette')).toBeVisible();
  await noErrors(page);
});

test('theme toggle switches the editor to dark', async ({ page }) => {
  await load(page, 'basic');
  const before = await page.evaluate(() => getComputedStyle(document.body).backgroundColor);
  await page.click('[data-cmd="theme"]');
  const after = await page.evaluate(() => getComputedStyle(document.body).backgroundColor);
  expect(after).not.toBe(before);
});

test('templates dialog replaces the diagram, undo brings it back', async ({ page }) => {
  await load(page, 'basic');
  await page.click('[data-cmd="new"]');
  await page.click('[data-template="swim"]');
  let d = await doc(page);
  expect(d.lanes.items).toHaveLength(3);
  await page.locator('#stage').focus();
  await page.keyboard.press('Control+z');
  d = await doc(page);
  expect(d.title).toBe('Basic flowchart');
});

test('help dialog opens with ?', async ({ page }) => {
  await load(page, 'basic');
  await page.locator('#stage').focus();
  await page.keyboard.press('?');
  await expect(page.locator('#modal')).toBeVisible();
  await expect(page.locator('#modal')).toContainText('Shortcuts');
  await page.keyboard.press('Escape');
  await expect(page.locator('#modal')).toBeHidden();
});

test('right-click menu on a shape and on empty canvas', async ({ page }) => {
  await setup(page, [N('a', 100, 100)]);
  const b = await box(page, 'a');
  await page.mouse.click(b.cx, b.cy, { button: 'right' });
  const menu = page.locator('#context-menu');
  await expect(menu).toBeVisible();
  await menu.getByRole('menuitem', { name: /Duplicate/ }).click();
  await expect(menu).toBeHidden();
  expect((await doc(page)).nodes).toHaveLength(2);
  const p = await toScreen(page, 600, 400);
  await page.mouse.click(p[0], p[1], { button: 'right' });
  await menu.getByRole('menuitem', { name: 'Add shape here' }).click();
  await page.keyboard.type('Here');
  await page.keyboard.press('Enter');
  const d = await doc(page);
  const n = d.nodes.find((x) => x.text === 'Here');
  expect(Math.abs(n.x + n.w / 2 - 600)).toBeLessThanOrEqual(8);
});

test('phases: add, rename, resize, and they appear as detail chips', async ({ page }) => {
  await setup(page, [N('a', 40, 80, { fields: { what: 'x' } }), N('b', 40, 300)], [], { lanes: { items: [{ id: 'l1', title: 'Team', size: 240 }] } });
  await page.click('[data-cmd="add-phase"]');
  await page.click('[data-cmd="add-phase"]');
  let d = await doc(page);
  expect(d.phases.items).toHaveLength(2);
  await page.fill('#phase-title', 'Close');
  await expect(page.locator('#diagram .phase').nth(1)).toContainText('Close');
  // first phase spans y 40..240; drag its bottom border down by 40 -> b (centre 330) moves too
  const p = await toScreen(page, 120, 240);
  await drag(page, p, [p[0], p[1] + 40]);
  d = await doc(page);
  expect(d.phases.items[0].size).toBe(240);
  expect(d.nodes.find((n) => n.id === 'b').y).toBe(340);
  expect(d.nodes.find((n) => n.id === 'a').y).toBe(80);
  // double-click a phase header to rename
  const hp = await toScreen(page, -17, 100);
  await page.mouse.dblclick(hp[0], hp[1]);
  await page.keyboard.type('Intake');
  await page.keyboard.press('Enter');
  expect((await doc(page)).phases.items[0].title).toBe('Intake');
  // preview shows phase chip
  await page.click('[data-cmd="mode-preview"]');
  await page.evaluate(() => {});
  const frame = page.frameLocator('#preview');
  await frame.locator('[data-node="a"]').click({ force: true });
  await expect(frame.locator('#detail .meta')).toContainText('Intake');
  await noErrors(page);
});
