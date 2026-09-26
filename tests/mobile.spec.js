import { test, expect } from '@playwright/test';
import { load, doc } from './helpers.js';

test.use({ viewport: { width: 390, height: 800 }, hasTouch: true });

test('narrow screens: shape and properties drawers', async ({ page }) => {
  await load(page, 'basic');
  await expect(page.locator('#palette')).toBeHidden();
  await page.click('[data-cmd="toggle-palette"]');
  await expect(page.locator('#palette')).toBeVisible();
  await page.click('.shape-btn[data-shape="decision"]');
  await expect(page.locator('#palette')).toBeHidden();
  const d = await doc(page);
  expect(d.nodes.at(-1).shape).toBe('decision');
  await page.click('[data-cmd="toggle-inspector"]');
  await expect(page.locator('#inspector')).toBeVisible();
  await expect(page.locator('#insp-text')).toBeVisible();
  const scrollW = await page.evaluate(() => document.documentElement.scrollWidth);
  expect(scrollW).toBeLessThanOrEqual(390);
});

test('two-finger pinch zooms the canvas', async ({ page }) => {
  await load(page, 'basic');
  const k0 = await page.evaluate(() => window.fc.app.view.k);
  await page.evaluate(() => {
    const svg = document.getElementById('canvas');
    const fire = (type, id, x, y) => svg.dispatchEvent(new PointerEvent(type, { pointerId: id, pointerType: 'touch', clientX: x, clientY: y, bubbles: true, isPrimary: id === 1, button: 0, buttons: 1 }));
    fire('pointerdown', 1, 150, 400); fire('pointerdown', 2, 250, 400);
    for (let i = 1; i <= 5; i++) { fire('pointermove', 1, 150 - i * 10, 400); fire('pointermove', 2, 250 + i * 10, 400); }
    fire('pointerup', 1, 100, 400); fire('pointerup', 2, 300, 400);
  });
  const k1 = await page.evaluate(() => window.fc.app.view.k);
  expect(k1).toBeGreaterThan(k0 * 1.5);
  expect((await doc(page)).nodes).toHaveLength(8); // pinch didn't create or move anything
});
