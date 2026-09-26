import { test, expect } from '@playwright/test';
import { load, doc, noErrors, setup, N, clickNode } from './helpers.js';

// A horizontal connector (a → b) crossed by a vertical one (c → d).
const CROSS = [
  [N('a', 0, 100), N('b', 400, 100), N('c', 180, -100), N('d', 180, 300)],
  [
    { id: 'h', from: { node: 'a', side: 'r' }, to: { node: 'b', side: 'l' } },
    { id: 'v', from: { node: 'c', side: 'b' }, to: { node: 'd', side: 't' } },
  ],
];
const pathOf = (page, id) => page.locator(`#canvas [data-edge="${id}"] path.edge-line`).getAttribute('d');

test('line hops: the horizontal connector arcs over a crossing, the vertical one does not', async ({ page }) => {
  await setup(page, ...CROSS);
  expect(await pathOf(page, 'h')).toMatch(/A5 5 0 0 1/);
  expect(await pathOf(page, 'v')).not.toContain('A');
  // exports carry the same hop
  expect(await page.evaluate(() => window.fc.exportSVG())).toMatch(/A5 5 0 0 1/);
  expect(await page.evaluate(() => window.fc.exportHTML())).toMatch(/A5 5 0 0 1/);
  // size setting
  await page.evaluate(() => { window.fc.getDoc().settings.jumpSize = 8; window.fc.renderCanvas(); });
  expect(await pathOf(page, 'h')).toMatch(/A8 8 0 0 1/);
  await noErrors(page);
});

test('line hops: "None" turns them off, from the document settings', async ({ page }) => {
  await setup(page, ...CROSS);
  await page.evaluate(() => window.fc.clearSelection());
  await page.selectOption('#set-hops', 'none');
  await expect.poll(() => pathOf(page, 'h')).not.toContain('A');
  expect((await doc(page)).settings.lineJumps).toBe('none');
  await page.selectOption('#set-hops', 'arc');
  await expect.poll(() => pathOf(page, 'h')).toContain('A');
});

test('line hops: connectors that meet at a shared shape do not hop', async ({ page }) => {
  // two connectors leaving the same port, and one ending next to a corner of the other
  await setup(page, [N('a', 0, 0), N('b', 300, 0), N('c', 0, 200)], [
    { id: 'e1', from: { node: 'a', side: 'r' }, to: { node: 'b', side: 'l' } },
    { id: 'e2', from: { node: 'a', side: 'b' }, to: { node: 'c', side: 't' } },
    { id: 'e3', from: { node: 'a', side: 'r' }, to: { node: 'c', side: 'r' } },
  ]);
  for (const id of ['e1', 'e2', 'e3']) expect(await pathOf(page, id)).not.toContain('A');
});

test('line hops follow a dragged shape', async ({ page }) => {
  await setup(page, ...CROSS);
  expect(await pathOf(page, 'h')).toContain('A');
  // move the vertical pair clear of the horizontal connector
  await page.evaluate(() => { const d = window.fc.getDoc(); for (const n of d.nodes) if (n.id === 'c' || n.id === 'd') n.x = 700; window.fc.renderCanvas(); });
  await expect.poll(() => pathOf(page, 'h')).not.toContain('A');
});

test('control point: typing in the Control field shows a green C, a number makes it C1', async ({ page }) => {
  await setup(page, [N('a', 0, 0)], []);
  const badge = page.locator('#canvas [data-node="a"] [data-badge="control"]');
  await expect(badge).toHaveCount(0);
  await clickNode(page, 'a');
  await page.fill('[data-field="control"]', 'Manager approves the request');
  await expect(badge).toHaveCount(1);
  await expect(badge.locator('text')).toHaveText('C');
  await expect(badge.locator('circle')).toHaveAttribute('fill', 'var(--control)');
  // the inspector shows it as automatic, while typing
  await expect(page.locator('[data-badge-toggle="control"]')).toBeChecked();
  await expect(page.locator('[data-badge-toggle="control"]')).toBeDisabled();
  await page.fill('[data-badge-num="control"]', '1');
  await expect(badge.locator('text')).toHaveText('C1');
  await expect(badge.locator('rect')).toHaveCount(1); // wider pill
  expect((await doc(page)).nodes[0].badgeNums).toEqual({ control: '1' });
  // legend and details panel
  const html = await page.evaluate(() => window.fc.exportHTML());
  expect(html).toContain('Control point');
  await page.click('[data-cmd="mode-preview"]');
  const frame = page.frameLocator('#preview');
  await frame.locator('[data-node="a"]').click();
  await expect(frame.locator('#detail .prp dt')).toHaveText('Control (C1)');
  expect(await frame.locator('#detail .prp').getAttribute('style')).toContain('--hl:var(--control)');
  await expect(frame.locator('.legend')).toContainText('Control point');
  await page.click('[data-cmd="mode-edit"]');
  // clearing the field removes it
  await page.evaluate(() => window.fc.select({ nodes: ['a'] }));
  await page.fill('[data-field="control"]', '');
  await expect(badge).toHaveCount(0);
  await noErrors(page);
});

test('risk badges still work, with numbers too', async ({ page }) => {
  await setup(page, [N('a', 0, 0, { badges: ['risk'], badgeNums: { risk: '2' } }), N('b', 300, 0, { fields: { risk: 'x' } })], []);
  await expect(page.locator('#canvas [data-node="a"] [data-badge="risk"] text')).toHaveText('R2');
  await expect(page.locator('#canvas [data-node="b"] [data-badge="risk"] text')).toHaveText('R');
});

const OLD = (fields, nodeFields) => ({
  title: 'Old', fields, badgeKinds: [{ id: 'risk', text: 'R', color: 'var(--risk)', name: 'Risk point' }],
  nodes: [{ id: 'a', shape: 'process', x: 0, y: 0, w: 160, h: 60, text: 'A', badges: ['risk'], fields: nodeFields }], edges: [],
});

test('older charts gain hops and an empty Control field; a ticked C gets a description and ID', async ({ page }) => {
  await load(page, 'blank');
  await page.evaluate((d) => window.fc.setDoc(d), OLD([{ key: 'what', label: 'What happens' }, { key: 'prp', label: 'Risk point to probe', highlight: true }], { prp: 'p' }));
  const d = await doc(page);
  expect(d.settings.lineJumps).toBe('arc');
  expect(d.badgeKinds.map((b) => b.id)).toEqual(['risk', 'control']);
  expect(d.fields.map((f) => f.key)).toEqual(['what', 'control', 'prp']);
  expect(d.badgeKinds[1].field).toBe('control');
  // nothing changes on the chart until C is used
  await expect(page.locator('#canvas [data-badge="control"]')).toHaveCount(0);
  expect(await page.evaluate(() => window.fc.exportSVG())).not.toContain('Control point');
  // tick C on a shape: there is a Control field for its description, and an ID box
  await page.evaluate(() => window.fc.select({ nodes: ['a'] }));
  await page.click('[data-badge-toggle="control"]');
  await expect(page.locator('#canvas [data-node="a"] [data-badge="control"] text')).toHaveText('C');
  await page.fill('[data-field="control"]', 'Second signatory approves payments');
  await page.fill('[data-badge-num="control"]', 'KC-04');
  await expect(page.locator('#canvas [data-node="a"] [data-badge="control"] text')).toHaveText('KC-04');
  await page.click('[data-cmd="mode-preview"]');
  const frame = page.frameLocator('#preview');
  await frame.locator('[data-node="a"]').click();
  await expect(frame.locator('#detail .prp dt').first()).toHaveText('Control (KC-04)');
  await noErrors(page);
});

test('older charts with their own control field link C to it; "Only when ticked" sticks', async ({ page }) => {
  await load(page, 'blank');
  await page.evaluate((d) => window.fc.setDoc(d), OLD([{ key: 'prp', label: 'Risk point to probe', highlight: true }, { key: 'kc', label: 'Key control' }], { prp: 'p', kc: 'k' }));
  expect((await doc(page)).fields.map((f) => f.key)).toEqual(['prp', 'kc']);
  await expect(page.locator('#canvas [data-badge="control"]')).toHaveCount(1);
  await page.evaluate(() => window.fc.clearSelection());
  await page.selectOption('[data-badge-field="control"]', '');
  await expect(page.locator('#canvas [data-badge="control"]')).toHaveCount(0);
  // survives save and re-open
  await page.evaluate(() => window.fc.setDoc(JSON.parse(JSON.stringify(window.fc.getDoc()))));
  await expect(page.locator('#canvas [data-badge="control"]')).toHaveCount(0);
  expect((await doc(page)).badgeKinds[1].field).toBe('');
  await noErrors(page);
});

test('spreadsheet import: Control and Control # columns give C1/C2 and green fields', async ({ page }) => {
  await load(page, 'blank');
  const r = await page.evaluate(async () => {
    const S = await import('/src/spreadsheet.js');
    const { nodeBadges } = await import('/src/model.js');
    const rows = S.parseDelimited(S.EXAMPLE_CSV);
    const roles = S.guessRoles(rows[0], rows.slice(1));
    const { doc } = S.buildFromRows(rows, roles);
    return {
      roles,
      control: doc.fields.find((f) => f.key === 'control'),
      badges: Object.fromEntries(doc.nodes.map((n) => [n.tag || n.text, nodeBadges(doc, n).map((b) => b.label).join(' ')])),
      off: S.buildFromRows(rows, roles, { badges: false }).doc.nodes.some((n) => nodeBadges(S.buildFromRows(rows, roles, { badges: false }).doc, n).length),
    };
  });
  expect(r.roles).toContain('controlNum');
  expect(r.control).toMatchObject({ highlight: true, color: 'var(--control)' });
  expect(r.badges['CD-02']).toBe('R C1');
  expect(r.badges['CD-04']).toBe('R C2');
  expect(r.badges['CD-01']).toBe('');
  expect(r.off).toBe(false);
});

test('outline: !C1 adds a numbered control badge', async ({ page }) => {
  await load(page, 'blank');
  const b = await page.evaluate(() => {
    const x = window.fc.parseOutline('a: Approve {process} !C1 !R\nb: Next');
    return { errors: x.errors, badges: x.doc.nodes[0].badges, nums: x.doc.nodes[0].badgeNums };
  });
  expect(b.errors).toEqual([]);
  expect(b.badges).toEqual(['control', 'risk']);
  expect(b.nums).toEqual({ control: '1' });
});

test('walkthrough template shows numbered controls', async ({ page }) => {
  await load(page, 'walkthrough');
  await expect(page.locator('#canvas [data-node="OTC-02"] [data-badge="control"] text')).toHaveText('C1');
  await expect(page.locator('#canvas [data-badge="control"]')).toHaveCount(3);
  await expect(page.locator('#canvas [data-node="OTC-02"] [data-badge="risk"] text')).toHaveText('R');
  await noErrors(page);
});
