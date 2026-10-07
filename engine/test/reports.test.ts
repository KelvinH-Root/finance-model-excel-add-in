// Summary and report modules: the full model recipe's dashboards and reports, the saved versions,
// the data table of scenario results, and a report module inserted into a model by change plan.
// (LibreOffice checks every chart table against the Python reference in tests/test_full_model.py.)

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { assemble, buildWorkbook, Model, planChange, rchartRefs, rchartXml, type Recipe } from '../src/index.ts';
import { loadLibrary } from '../src/node/library.ts';

const ROOT = join(import.meta.dirname, '..', '..');
const lib = loadLibrary(join(ROOT, 'library', 'hfg'));
const recipe = JSON.parse(readFileSync(join(ROOT, 'library', 'hfg', 'recipes', 'full_model.json'), 'utf8')) as Recipe;
const full = () => Model.fromRecipe(lib, recipe);

test('the full model brings every chart in the register, each named by its id', () => {
  const layout = assemble(full());
  const ids = layout.rcharts.map(c => c.id);
  const expected = [...Array.from({ length: 95 }, (_, k) => `C${String(k + 1).padStart(2, '0')}`), 'S01', 'S02', 'S03', 'S04', 'S05', 'S06'];
  assert.deepEqual([...ids].sort(), [...expected].sort());
  for (const sheet of ['Income summary', 'Income report', 'Scenario report']) assert.equal(layout.kindOf(sheet), 'report');
  const pos = layout.positions();
  for (const c of layout.rcharts) {
    for (const s of c.series) assert.ok(pos.has(s.row), `${c.id}: series row ${s.row}`);
    const xml = rchartXml(rchartRefs(layout, c));
    assert.match(xml, c.type === 'pie' ? /<c:pieChart>/ : /<c:(barChart|lineChart)>/);
    assert.match(xml, /<c:title><c:tx><c:strRef><c:f>'[^']+'!\$I\$\d+<\/c:f>/, `${c.id} title reads its cell`);
  }
});

test('stacked bars never ask for labels outside the end of a bar (Excel refuses it)', () => {
  const layout = assemble(full());
  for (const c of layout.rcharts) {
    if (c.grouping !== 'stacked') continue;
    for (const s of c.series) assert.notEqual(s.labels?.pos, 'outEnd', `${c.id}`);
  }
});

test('scenario charts read a data table on the Scenarios sheet, keyed to the active scenario', () => {
  const model = full();
  const layout = assemble(model);
  assert.equal(layout.dataTables.length, 1);
  const t = layout.dataTables[0];
  assert.equal(t.sheet, 'Scenarios');
  assert.equal(t.input, 'Sel_Scenario');
  const parts = unzipSync(buildWorkbook(layout, model, { created: new Date('2026-10-07T00:00:00Z') }));
  const sheets = Object.keys(parts).filter(p => /^xl\/worksheets\/sheet\d+\.xml$/.test(p)).map(p => strFromU8(parts[p]));
  const dt = sheets.flatMap(x => [...x.matchAll(/<f t="dataTable" ref="(J\d+:L\d+)" dt2D="0" dtr="1" r1="(I\d+)"\/>/g)]);
  assert.equal(dt.length, 1);
});

test('saved versions: a register row and a store block for each, and the Compared with list names them', () => {
  const layout = assemble(full());
  const versions = recipe.instances.find(i => i.module === 'fm.versions')!.data!.versions!;
  const reg = layout.sheetRows().get('Versions')!.filter(r => r.role === 'r.reg');
  assert.equal(reg.length, versions.length);
  const store = layout.sheetRows().get('Version store')!.filter(r => r.role === 'r.store');
  assert.equal(store.length, versions.reduce((n, v) => n + Object.keys(v.values).length, 0));
  for (const nm of ['VS_Keys', 'VS_Values', 'VR_Label', 'VR_BudgetKey', 'VR_RefKey']) assert.ok(layout.ranges.has(nm), nm);
  const items = layout.sheetRows().get('Lookups')!.filter(r => /^lookups\/List_Compare\/\d+$/.test(r.id));
  assert.equal(items.length, 4 + versions.length);
});

test('inserting a summary module plans its sheet, its rows and its charts', () => {
  const before = full();
  before.remove(before.instances.find(i => i.module === 'fm.income_summary')!.uid);
  const after = before.copy();
  after.insert('fm.income_summary', { year: 2 });
  const plan = planChange(assemble(before), assemble(after));
  const added = plan.ops.filter(o => o.op === 'add_rchart');
  assert.deepEqual(added.map(o => (o as { id: string }).id), ['C01', 'C02', 'C03', 'C04', 'C05']);
  assert.ok(plan.ops.some(o => o.op === 'add_sheet' && o.sheet === 'Income summary'));
  assert.ok(plan.preview.some(p => p.includes('chart C01')));
});

test('a report module needs the statements, and its charts need the selections they read', () => {
  const m = new Model(lib, 12);
  m.info = full().info;
  m.insert('fm.income_summary', { year: 1 });
  assert.throws(() => assemble(m), /reads the financial statements/);
});
