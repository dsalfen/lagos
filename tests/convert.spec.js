// The converter for the earlier hand-coded walkthrough pages (scripts/convert-legacy.mjs).
import { test, expect } from '@playwright/test';
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';

test('converts a legacy walkthrough page into an editor document that opens and matches', async ({ page }, info) => {
  const out = info.outputPath('legacy.json');
  execFileSync('node', ['scripts/convert-legacy.mjs', 'tests/fixtures/legacy-walkthrough.html', out]);
  const d = JSON.parse(readFileSync(out, 'utf8'));
  expect(d.title).toBe('Test Cycle Flowchart');
  expect(d.lanes.items.map((l) => [l.title, l.size])).toEqual([['External', 200], ['Team', 200], ['Finance', 240]]);
  expect(d.lanes.items[0].subtitle).toBe('Suppliers');
  expect(d.nodes).toHaveLength(6);
  const n = (id) => d.nodes.find((x) => x.id === id);
  expect(n('TST-02')).toMatchObject({ shape: 'decision', tag: 'TST-02', badges: ['risk'], laneLabel: 'Team' });
  expect(n('EXT')).toMatchObject({ shape: 'database', tag: '', badges: [] });
  expect(n('TST-03').shape).toBe('data');
  expect(n('TST-04').shape).toBe('manualInput');
  // every connector is attached at both ends, except the deliberate stub
  const loose = d.edges.filter((e) => !e.from.node || (!e.to.node && !e.labelAlign));
  expect(loose).toEqual([]);
  const feed = d.edges.find((e) => e.label === 'remittance');
  expect(feed).toMatchObject({ kind: 'feed', from: { node: 'TST-04', side: 'l' }, to: { node: 'EXT', side: 'r', offset: 10 } });
  expect(feed.labelPos).toBeGreaterThan(0.2);
  const stub = d.edges.find((e) => e.label === 'No: reject');
  expect(stub).toMatchObject({ labelAlign: 'end', from: { node: 'TST-02', side: 'l' } });
  const yes = d.edges.find((e) => e.label === 'Yes');
  expect(yes.waypoints).toHaveLength(2);
  expect(d.settings).toMatchObject({ initialSelection: 'TST-02', noTagLabel: 'Connector', verticalLabels: 'rotate' });
  expect(d.legendOrder.slice(0, 3)).toEqual(['shape:process', 'shape:decision', 'shape:document']);
  expect(d.badgeKinds[0].name).toBe('Risk point');

  // it opens in the editor and the exported page behaves like the original
  await page.goto('/?template=blank');
  await page.waitForFunction(() => window.fc);
  await page.setInputFiles('#file-input', out);
  await expect.poll(async () => page.evaluate(() => window.fc.getDoc().nodes.length)).toBe(6);
  const html = await page.evaluate(() => window.fc.exportHTML());
  await page.setContent(html);
  await expect(page.locator('#detail .id')).toHaveText('TST-02');
  await expect(page.locator('.legend')).toContainText('Report'); // fixed legend: listed although unused
  await page.locator('[data-node="EXT"]').click();
  await expect(page.locator('#detail .id')).toHaveText('Connector');
  await expect(page.locator('#detail .meta')).toContainText('External');
});
