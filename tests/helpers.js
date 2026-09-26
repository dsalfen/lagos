import { expect } from '@playwright/test';

// Loads the editor with a template (or blank) and fails the test on any page error.
export async function load(page, template = 'basic') {
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => {
    if (m.type() === 'error' && !/Failed to load resource|ERR_CERT|fonts\.g/.test(m.text())) errors.push(m.text());
  });
  await page.route(/fonts\.(googleapis|gstatic)\.com/, (r) => r.abort());
  await page.goto('/?template=' + template);
  await page.waitForFunction(() => window.fc && window.fc.getDoc(), null, { timeout: 5000 })
    .catch(() => { throw new Error('Editor failed to boot: ' + errors.join(' | ')); });
  await page.waitForTimeout(50);
  page.__errors = errors;
  return errors;
}

export const doc = (page) => page.evaluate(() => JSON.parse(JSON.stringify(window.fc.getDoc())));
export const box = (page, id) => page.evaluate((i) => window.fc.nodeScreenBox(i), id);
export const toScreen = (page, x, y) => page.evaluate(([a, b]) => window.fc.worldToScreen(a, b), [x, y]);

export async function drag(page, from, to, { steps = 12, modifiers = [] } = {}) {
  for (const m of modifiers) await page.keyboard.down(m);
  await page.mouse.move(from[0], from[1]);
  await page.mouse.down();
  for (let i = 1; i <= steps; i++) {
    await page.mouse.move(from[0] + ((to[0] - from[0]) * i) / steps, from[1] + ((to[1] - from[1]) * i) / steps);
  }
  await page.mouse.up();
  for (const m of modifiers) await page.keyboard.up(m);
}

// Select a node by clicking its centre.
export async function clickNode(page, id, { modifiers = [] } = {}) {
  const b = await box(page, id);
  for (const m of modifiers) await page.keyboard.down(m);
  await page.mouse.click(b.cx, b.cy);
  for (const m of modifiers) await page.keyboard.up(m);
}

export async function noErrors(page) {
  expect(page.__errors || [], 'page errors').toEqual([]);
}

// Blank canvas with a known set of nodes, for deterministic interaction tests.
export async function setup(page, nodes, edges = [], extra = {}) {
  await load(page, 'blank');
  await page.evaluate(([nodes, edges, extra]) => window.fc.setDoc({ title: 'Test', nodes, edges, ...extra }), [nodes, edges, extra]);
  await page.waitForTimeout(60); // let the post-load "fit" run first
  await page.evaluate(() => { window.fc.app.view = { x: 60, y: 60, k: 1 }; window.fc.renderCanvas(); });
}

export const N = (id, x, y, extra = {}) => ({ id, shape: 'process', x, y, w: 160, h: 60, text: id.toUpperCase(), ...extra });
