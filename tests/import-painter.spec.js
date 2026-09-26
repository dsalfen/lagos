import { test, expect } from '@playwright/test';
import { deflateRawSync } from 'node:zlib';
import { writeFileSync } from 'node:fs';
import { load, doc, box, toScreen, clickNode, noErrors, setup, N } from './helpers.js';

// --- a tiny .xlsx writer (deflated zip entries, shared + inline strings) ----------
function zip(files) {
  const parts = [], central = [];
  let offset = 0;
  for (const [name, text] of Object.entries(files)) {
    const nameB = Buffer.from(name), raw = Buffer.from(text), data = deflateRawSync(raw);
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(8, 8);
    lh.writeUInt32LE(data.length, 18); lh.writeUInt32LE(raw.length, 22); lh.writeUInt16LE(nameB.length, 26);
    parts.push(lh, nameB, data);
    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6); ch.writeUInt16LE(8, 10);
    ch.writeUInt32LE(data.length, 20); ch.writeUInt32LE(raw.length, 24); ch.writeUInt16LE(nameB.length, 28); ch.writeUInt32LE(offset, 42);
    central.push(ch, nameB);
    offset += 30 + nameB.length + data.length;
  }
  const cd = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(Object.keys(files).length, 8); end.writeUInt16LE(Object.keys(files).length, 10);
  end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...parts, cd, end]);
}
function xlsx(rows) {
  const shared = [];
  const si = (s) => { let i = shared.indexOf(s); if (i < 0) { shared.push(s); i = shared.length - 1; } return i; };
  const col = (i) => String.fromCharCode(65 + i);
  const sheetRows = rows.map((r, ri) => `<row r="${ri + 1}">${r.map((v, ci) => (v === '' ? '' : ci === 1
    ? `<c r="${col(ci)}${ri + 1}" t="inlineStr"><is><t>${v}</t></is></c>`
    : `<c r="${col(ci)}${ri + 1}" t="s"><v>${si(v)}</v></c>`)).join('')}</row>`).join('');
  return zip({
    'xl/workbook.xml': '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Steps" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="worksheet" Target="worksheets/steps.xml"/></Relationships>',
    'xl/sharedStrings.xml': `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${shared.map((s) => `<si><t>${s}</t></si>`).join('')}</sst>`,
    'xl/worksheets/steps.xml': `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetRows}</sheetData></worksheet>`,
  });
}

test('spreadsheet import: example data builds lanes, phases, branches and details', async ({ page }) => {
  await load(page, 'blank');
  await page.click('[data-cmd="export-menu"]');
  await page.click('[data-cmd="import-sheet"]');
  await page.click('#sheet-example');
  await expect(page.locator('#sheet-summary')).toContainText('7 steps · 7 connectors · 3 lanes · 2 phases');
  await expect(page.locator('select[data-col="2"]')).toHaveValue('lane');
  await page.click('#sheet-create');
  const d = await doc(page);
  expect(d.title).toBe('Cash disbursements walkthrough');
  expect(d.lanes.items.map((l) => l.title)).toEqual(['Requester', 'AP', 'Treasury']);
  expect(d.phases.items.map((p) => p.title)).toEqual(['Approve', 'Pay']);
  const by = (tag) => d.nodes.find((n) => n.tag === tag);
  expect(by('CD-02').shape).toBe('decision');
  expect(by('CD-01').shape).toBe('manualInput');
  expect(d.nodes.find((n) => n.text === 'ERP').shape).toBe('database');
  expect(by('CD-04').badges).toEqual(['risk']);
  expect(by('CD-04').fields.risk).toContain('without approval');
  const e = (a, b) => d.edges.find((x) => x.from.node === a.id && x.to.node === b.id);
  expect(e(by('CD-02'), by('CD-03')).label).toBe('Yes');
  expect(e(by('CD-02'), by('CD-01')).label).toBe('No');
  const feed = e(by('CD-03'), d.nodes.find((n) => n.text === 'ERP'));
  expect([feed.kind, feed.label]).toEqual(['feed', 'posts']);
  // lanes: each shape sits in its lane; phases: every Pay step is below every Approve step
  const laneOf = (n) => { let p = 0; for (const l of d.lanes.items) { const c = n.x + n.w / 2; if (c >= p && c < p + l.size) return l.title; p += l.size; } };
  expect(laneOf(by('CD-04'))).toBe('Treasury');
  const approveEnd = 40 + d.phases.items[0].size;
  for (const t of ['CD-00', 'CD-01', 'CD-02']) expect(by(t).y + by(t).h).toBeLessThanOrEqual(approveEnd);
  for (const t of ['CD-03', 'CD-04', 'CD-05']) expect(by(t).y).toBeGreaterThanOrEqual(approveEnd);
  await noErrors(page);
});

test('spreadsheet import: pasted Excel rows (tabs) and column remapping', async ({ page }) => {
  await load(page, 'blank');
  await page.click('[data-cmd="export-menu"]');
  await page.click('[data-cmd="import-sheet"]');
  const tsv = 'Ref\tTask\tWho\tNotes\n1\tReceive order\tSales\tBy email\n2\tCheck stock\tWarehouse\t\n3\tShip\tWarehouse\tCourier';
  await page.fill('#sheet-paste', tsv);
  await expect(page.locator('#sheet-summary')).toContainText('3 steps · 2 connectors · 2 lanes'); // no Next column: rows connect in order
  await expect(page.locator('select[data-col="0"]')).toHaveValue('id');
  // treat "Notes" as ignored instead of a detail field
  await page.selectOption('select[data-col="3"]', 'ignore');
  await page.click('#sheet-create');
  const d = await doc(page);
  expect(d.nodes.map((n) => n.text)).toEqual(['Receive order', 'Check stock', 'Ship']);
  expect(d.fields.some((f) => f.label === 'Notes')).toBe(false);
  expect(d.edges).toHaveLength(2);
});

test('spreadsheet import: reads an .xlsx file (via Open) and reports bad references', async ({ page }) => {
  await load(page, 'blank');
  const path = test.info().outputPath('steps.xlsx');
  writeFileSync(path, xlsx([
    ['Step', 'Activity', 'Lane', 'Next', 'Risk'],
    ['P-1', 'Raise PO', 'Buyer', 'P-2', ''],
    ['P-2', 'Approved?', 'Manager', 'Yes: P-3; No: P-1; Maybe: P-9', 'Approval could be skipped'],
    ['P-3', 'Send PO', 'Buyer', '', ''],
  ]));
  await page.setInputFiles('#file-input', path);
  await expect(page.locator('#sheet-summary')).toContainText('3 steps · 3 connectors · 2 lanes');
  await expect(page.locator('#sheet-warnings')).toContainText('"P-9" not found');
  await page.click('#sheet-create');
  const d = await doc(page);
  expect(d.title).toBe('steps');
  const p2 = d.nodes.find((n) => n.tag === 'P-2');
  expect(p2.shape).toBe('decision');
  expect(p2.badges).toEqual(['risk']);
  expect(d.fields.find((f) => f.label === 'Risk').highlight).toBe(true);
});

test('spreadsheet parser handles quotes, commas and blank rows', async ({ page }) => {
  await load(page, 'blank');
  const rows = await page.evaluate(async () => (await import('/src/spreadsheet.js')).parseDelimited('A,B\n"x, y","say ""hi""\nthere"\n\n1,2\n'));
  expect(rows).toEqual([['A', 'B'], ['x, y', 'say "hi"\nthere'], ['1', '2']]);
});

// --- format painter -------------------------------------------------------------

const styled = (id, x, y) => N(id, x, y, { style: { fill: '#FFF1CC', stroke: 'var(--risk)', dash: '6 4', bold: true, fontSize: 15 } });

test('format painter copies a shape style onto another shape once', async ({ page }) => {
  await setup(page, [styled('a', 100, 100), N('b', 100, 300), N('c', 400, 300)]);
  await clickNode(page, 'a');
  await page.click('#painter-btn');
  await expect(page.locator('#painter-btn')).toHaveClass(/on/);
  await expect(page.locator('#status-hint')).toContainText('Format painter');
  await clickNode(page, 'b');
  await expect(page.locator('#painter-btn')).not.toHaveClass(/on/);
  await clickNode(page, 'c'); // painter is off now: this just selects
  const d = await doc(page);
  expect(d.nodes[1].style).toEqual(d.nodes[0].style);
  expect(d.nodes[2].style).toEqual({});
  expect(d.nodes[1].text).toBe('B'); // only formatting changes
  await page.keyboard.press('Control+z');
  expect((await doc(page)).nodes[1].style).toEqual({});
});

test('double-clicking the painter keeps it on until Escape', async ({ page }) => {
  await setup(page, [styled('a', 100, 100), N('b', 100, 300), N('c', 400, 300), N('d', 400, 100)]);
  await clickNode(page, 'a');
  await page.dblclick('#painter-btn');
  await clickNode(page, 'b');
  await clickNode(page, 'c');
  await expect(page.locator('#painter-btn')).toHaveClass(/on/);
  await page.keyboard.press('Escape');
  await expect(page.locator('#painter-btn')).not.toHaveClass(/on/);
  await clickNode(page, 'd');
  const d = await doc(page);
  expect(d.nodes[1].style.fill).toBe('#FFF1CC');
  expect(d.nodes[2].style.fill).toBe('#FFF1CC');
  expect(d.nodes[3].style).toEqual({});
});

test('format painter works on connectors and refuses mismatched targets', async ({ page }) => {
  await setup(page, [N('a', 100, 100), N('b', 100, 300), N('c', 400, 100), N('d', 400, 300)], [
    { id: 'e1', from: { node: 'a' }, to: { node: 'b' }, kind: 'feed', style: { color: '#2E7D32', width: 3, endArrow: 'open' } },
    { id: 'e2', from: { node: 'c' }, to: { node: 'd' } }]);
  await page.evaluate(() => window.fc.select({ edges: ['e1'] }));
  await page.click('#painter-btn');
  await clickNode(page, 'c'); // a shape: not applicable
  await expect(page.locator('.toast')).toContainText('only be pasted onto connectors');
  const p = await toScreen(page, 480, 200);
  await page.mouse.click(p[0], p[1]);
  const e2 = (await doc(page)).edges[1];
  expect(e2).toMatchObject({ kind: 'feed', style: { color: '#2E7D32', width: 3, endArrow: 'open' } });
});

test('copy / paste formatting onto a multi-selection with Ctrl+Alt+C / V', async ({ page }) => {
  await setup(page, [styled('a', 100, 100), N('b', 100, 300), N('c', 400, 300)]);
  await clickNode(page, 'a');
  await page.keyboard.press('Control+Alt+c');
  await page.evaluate(() => window.fc.select({ nodes: ['b', 'c'] }));
  await page.locator('#stage').focus();
  await page.keyboard.press('Control+Alt+v');
  const d = await doc(page);
  expect(d.nodes[1].style).toEqual(d.nodes[0].style);
  expect(d.nodes[2].style).toEqual(d.nodes[0].style);
  // and one undo reverts both
  await page.keyboard.press('Control+z');
  const u = await doc(page);
  expect([u.nodes[1].style, u.nodes[2].style]).toEqual([{}, {}]);
});
