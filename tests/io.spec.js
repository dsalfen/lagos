import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { load, doc, noErrors, setup, N, badgesOf } from './helpers.js';

async function exportVia(page, cmd) {
  await page.click('[data-cmd="export-menu"]');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click(`[data-cmd="${cmd}"]`)]);
  return dl;
}

test('interactive HTML export is standalone and round-trips through Open', async ({ page, browser }) => {
  await load(page, 'walkthrough');
  const dl = await exportVia(page, 'export-html');
  expect(dl.suggestedFilename()).toBe('Order-to-cash_walkthrough_example.html');
  const path = test.info().outputPath('export.html');
  await dl.saveAs(path);
  const html = readFileSync(path, 'utf8');
  expect(html).toContain('id="fc-doc"');
  expect(html).not.toContain('src/'); // no references to editor modules
  // open it on its own
  const p2 = await browser.newPage();
  const errs = [];
  p2.on('pageerror', (e) => errs.push(e.message));
  await p2.route(/fonts\./, (r) => r.abort());
  await p2.goto('file://' + path);
  await expect(p2.locator('#detail h2')).toHaveText('Within credit limit?');
  await p2.locator('[data-node="OTC-07"]').click();
  await expect(p2.locator('#detail .id')).toHaveText('OTC-07');
  // keyboard access
  await p2.locator('[data-node="OTC-01"]').focus();
  await p2.keyboard.press('Enter');
  await expect(p2.locator('#detail .id')).toHaveText('OTC-01');
  expect(errs).toEqual([]);
  await p2.close();
  // re-open in the editor
  await page.click('[data-cmd="new"]');
  await page.click('[data-template="blank"]');
  await page.setInputFiles('#file-input', path);
  await expect.poll(async () => (await doc(page)).nodes.length).toBe(15);
  expect((await doc(page)).title).toBe('Order-to-cash walkthrough (example)');
});

test('SVG export resolves theme colours and is valid XML', async ({ page }) => {
  await load(page, 'basic');
  const dl = await exportVia(page, 'export-svg');
  const svg = readFileSync(await dl.path(), 'utf8');
  expect(svg.startsWith('<?xml')).toBe(true);
  expect(svg).not.toMatch(/var\(--/);
  const ok = await page.evaluate((s) => !new DOMParser().parseFromString(s, 'image/svg+xml').querySelector('parsererror'), svg);
  expect(ok).toBe(true);
});

test('PNG export produces an image', async ({ page }) => {
  await load(page, 'basic');
  const dl = await exportVia(page, 'export-png');
  const buf = readFileSync(await dl.path());
  expect(buf.slice(1, 4).toString()).toBe('PNG');
  expect(buf.length).toBeGreaterThan(5000);
});

test('JSON save and open round-trip', async ({ page }) => {
  await setup(page, [N('a', 100, 100, { tag: 'T1', fields: { what: 'x' } }), N('b', 100, 300)], [{ id: 'e1', from: { node: 'a' }, to: { node: 'b' }, label: 'L' }]);
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('[data-cmd="save"]')]);
  const path = await dl.path();
  const saved = JSON.parse(readFileSync(path, 'utf8'));
  expect(saved.nodes).toHaveLength(2);
  await page.evaluate(() => window.fc.loadTemplate('blank'));
  await page.setInputFiles('#file-input', path);
  await expect.poll(async () => (await doc(page)).nodes.length).toBe(2);
  const d = await doc(page);
  expect(d.edges[0].label).toBe('L');
  expect(d.nodes[0].fields.what).toBe('x');
});

test('opening a bad file shows an error and keeps the diagram', async ({ page }) => {
  await load(page, 'basic');
  await page.setInputFiles('#file-input', { name: 'x.json', mimeType: 'application/json', buffer: Buffer.from('not json') });
  await expect(page.locator('.toast')).toContainText('Could not open');
  expect((await doc(page)).nodes).toHaveLength(8);
});

test('text outline import builds a laned diagram', async ({ page }) => {
  await load(page, 'blank');
  await page.click('[data-cmd="export-menu"]');
  await page.click('[data-cmd="import-text"]');
  await page.click('#outline-create');
  const d = await doc(page);
  expect(d.title).toBe('Purchase to pay');
  expect(d.lanes.items.map((l) => l.title)).toEqual(['Requester', 'Procurement', 'Finance']);
  expect(d.nodes).toHaveLength(7);
  expect(d.edges).toHaveLength(8);
  expect(d.edges.filter((e) => e.kind === 'feed')).toHaveLength(1);
  const ok = d.nodes.find((n) => n.text === 'Within budget?');
  expect(ok.shape).toBe('decision');
  const po = d.nodes.find((n) => n.tag === 'P2P-03');
  expect(await badgesOf(page, po.id)).toEqual(['risk', 'control']);
  expect(po.badgeNums).toEqual({ control: '1' });
  expect(po.fields.risk).toContain('without an approved request');
  // every node sits inside its declared lane
  const laneOf = (n) => { let p = 0; for (const l of d.lanes.items) { const c = n.x + n.w / 2; if (c >= p && c < p + l.size) return l.title; p += l.size; } };
  expect(laneOf(d.nodes.find((n) => n.tag === 'P2P-04'))).toBe('Finance');
  expect(laneOf(d.nodes.find((n) => n.text === 'Need identified'))).toBe('Requester');
  // no two shapes overlap
  for (const a of d.nodes) for (const b of d.nodes) {
    if (a === b) continue;
    const overlap = a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
    expect(overlap, `${a.text} overlaps ${b.text}`).toBe(false);
  }
  await noErrors(page);
});

test('outline parser reports unknown lines', async ({ page }) => {
  await load(page, 'blank');
  const r = await page.evaluate(() => { const x = window.fc.parseOutline('a: One {nonsense}\n?? what\na -> b'); return { errors: x.errors, n: x.doc.nodes.length }; });
  expect(r.n).toBe(2);
  expect(r.errors.length).toBe(2);
});

test('auto-layout arranges an untidy diagram without overlaps', async ({ page }) => {
  await setup(page, [N('a', 0, 0), N('b', 10, 10), N('c', 20, 20), N('d', 30, 30)], [
    { from: { node: 'a' }, to: { node: 'b' } }, { from: { node: 'a' }, to: { node: 'c' } }, { from: { node: 'b' }, to: { node: 'd' } }, { from: { node: 'c' }, to: { node: 'd' } }]);
  await page.locator('#stage').click({ position: { x: 900, y: 700 } });
  await page.click('#auto-layout');
  const d = await doc(page);
  const y = (id) => d.nodes.find((n) => n.id === id).y;
  expect(y('a')).toBeLessThan(y('b'));
  expect(y('b')).toBe(y('c'));
  expect(y('d')).toBeGreaterThan(y('b'));
  const b = d.nodes.find((n) => n.id === 'b'), c = d.nodes.find((n) => n.id === 'c');
  expect(Math.abs(b.x - c.x)).toBeGreaterThanOrEqual(160);
});

test('pasting plain text creates shapes', async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await setup(page, []);
  await page.evaluate(() => navigator.clipboard.writeText('First step\nSecond step'));
  await page.locator('#stage').focus();
  await page.keyboard.press('Control+v');
  const d = await doc(page);
  expect(d.nodes.map((n) => n.text)).toEqual(['First step', 'Second step']);
  expect(d.edges).toHaveLength(1);
});

test('standalone build works from file://', async ({ page }) => {
  const { execSync } = await import('node:child_process');
  execSync('node scripts/build.js', { stdio: 'ignore' });
  const errs = [];
  page.on('pageerror', (e) => errs.push(e.message));
  await page.route(/fonts\./, (r) => r.abort());
  await page.goto('file://' + process.cwd() + '/dist/flowchart-editor.html?template=walkthrough');
  await page.waitForFunction(() => window.fc && window.fc.getDoc());
  await expect(page.locator('#diagram .node')).toHaveCount(15);
  await page.click('[data-cmd="mode-preview"]');
  await expect(page.frameLocator('#preview').locator('h1')).toHaveText('Order-to-cash walkthrough (example)');
  expect(errs).toEqual([]);
});

test('exported page has zoom controls and fits on a phone', async ({ browser, page }) => {
  await load(page, 'walkthrough');
  const html = await page.evaluate(() => window.fc.exportHTML());
  const p2 = await browser.newPage({ viewport: { width: 390, height: 800 } });
  await p2.route(/fonts\./, (r) => r.abort());
  await p2.setContent(html);
  const w0 = await p2.evaluate(() => document.querySelector('#chart svg').getBoundingClientRect().width);
  const cw = await p2.evaluate(() => document.getElementById('chart').clientWidth);
  expect(w0).toBeLessThanOrEqual(cw + 1); // started fitted
  await p2.click('[data-zoom="in"]');
  const w1 = await p2.evaluate(() => document.querySelector('#chart svg').getBoundingClientRect().width);
  expect(w1).toBeGreaterThan(w0);
  const scrollW = await p2.evaluate(() => document.documentElement.scrollWidth);
  expect(scrollW).toBeLessThanOrEqual(390);
  await p2.close();
});

test('single-file build is fully offline, with embedded fonts in the editor, preview and export', async ({ browser }) => {
  const { execSync } = await import('node:child_process');
  execSync('node scripts/build.js', { stdio: 'ignore' });
  const ctx = await browser.newContext({ offline: true, acceptDownloads: true });
  const page = await ctx.newPage();
  const external = [];
  page.on('request', (r) => { if (!/^(file|data|blob|about):/.test(r.url())) external.push(r.url()); });
  await page.goto('file://' + process.cwd() + '/dist/flowchart-editor.html?template=walkthrough');
  await page.waitForFunction(() => window.fc && window.fc.getDoc());
  await page.evaluate(() => document.fonts.ready);
  const loaded = await page.evaluate(() => ['IBM Plex Sans', 'IBM Plex Sans Condensed', 'IBM Plex Mono'].map((f) => document.fonts.check(`600 12px "${f}"`)));
  expect(loaded).toEqual([true, true, true]);
  expect(external).toEqual([]);
  // preview iframe uses the embedded fonts too
  await page.click('[data-cmd="mode-preview"]');
  const frame = page.frameLocator('#preview');
  await expect(frame.locator('h1')).toHaveText('Order-to-cash walkthrough (example)');
  expect(await frame.locator('#embedded-fonts').count()).toBe(1);
  await page.click('[data-cmd="mode-edit"]');
  // the standard export has the fonts built in
  await page.click('[data-cmd="export-menu"]');
  await expect(page.locator('[data-cmd="export-html-small"]')).toBeVisible();
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('[data-cmd="export-html"]')]);
  const path = test.info().outputPath('offline.html');
  await dl.saveAs(path);
  const { readFileSync } = await import('node:fs');
  const html = readFileSync(path, 'utf8');
  expect(html).toContain('@font-face');
  expect(html).not.toContain('fonts.googleapis.com');
  // the exported page renders with the fonts, offline
  const p2 = await ctx.newPage();
  await p2.goto('file://' + path);
  await p2.evaluate(() => document.fonts.ready);
  expect(await p2.evaluate(() => document.fonts.check('500 12px "IBM Plex Sans Condensed"'))).toBe(true);
  await p2.locator('[data-node="OTC-10"]').click();
  await expect(p2.locator('#detail .id')).toHaveText('OTC-10');
  // a PNG still exports offline
  await page.click('[data-cmd="export-menu"]');
  const [png] = await Promise.all([page.waitForEvent('download'), page.click('[data-cmd="export-png"]')]);
  expect(png.suggestedFilename()).toMatch(/\.png$/);
  await ctx.close();
});

test('the development version also embeds fonts in exports (local font file)', async ({ page }) => {
  await load(page, 'basic');
  const html = await page.evaluate(() => window.fc.exportHTML());
  expect(html).toContain('@font-face');
  expect(html).toContain('IBM Plex Sans Condensed');
});
