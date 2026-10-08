// The Scenarios sheet (Kelvin, 8 October 2026): a band over the scenario columns with the active one
// picked out, the names across their columns, a description of each, the adjustments grouped by what
// they adjust (one heading per group, a line per module beneath it), the scenario in use shaded by
// conditional formats, and the results the reports show grouped by measure.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { assemble, buildWorkbook, Model, planChange, type Recipe } from '../src/index.ts';
import { loadLibrary } from '../src/node/library.ts';

const ROOT = join(import.meta.dirname, '..', '..');
const lib = loadLibrary(join(ROOT, 'library', 'hfg'));
const recipe = JSON.parse(readFileSync(join(ROOT, 'library', 'hfg', 'recipes', 'full_model.json'), 'utf8')) as Recipe;
const full = () => Model.fromRecipe(lib, recipe);

function scenarioSheetXml(): string {
  const model = full();
  const layout = assemble(model);
  const parts = unzipSync(buildWorkbook(layout, model, { created: new Date('2026-10-08T00:00:00Z') }));
  const n = layout.sheets.findIndex(([s]) => s === 'Scenarios') + 1;
  return strFromU8(parts[`xl/worksheets/sheet${n}.xml`]);
}

test('adjustments sit under one heading per group, with a line for each module and no heading of their own', () => {
  const rows = assemble(full()).sheetRows().get('Scenarios')!;
  const groups = rows.filter(r => r.role === 'group').map(r => r.label);
  assert.deepEqual(groups, ['Revenue', 'Cost of sales', 'Staff costs', 'Operating expenses', 'Interest rates']);
  assert.equal(rows.filter(r => r.kind === 'section').length, 0);
  const revenue = rows.findIndex(r => r.role === 'group' && r.label === 'Revenue');
  const next = rows.findIndex((r, k) => k > revenue && r.role === 'group');
  const lines = rows.slice(revenue, next).filter(r => r.role === 'line');
  assert.equal(lines.length, 3);
  for (const l of lines) assert.match(String(l.cells[4]), /^=«B\|.+\/heading»$/, 'a line is labelled by its module title');
  assert.equal(rows.filter(r => r.kind === 'heading').length, 2, 'one bar for the scenarios, one for the results');
});

test('each group uses the active scenario unless set to another, and the line reads the scenario it uses', () => {
  const layout = assemble(full());
  const rows = layout.sheetRows().get('Scenarios')!;
  const use = rows.find(r => r.name === 'Sel_Scn_Revenue')!;
  assert.deepEqual(use.control, { kind: 'drop', list: 'List_Scenario_Use' });
  const line = rows.find(r => r.role === 'line')!;
  assert.match(String(line.cells[9]), /INDEX\(.+,IF\(Sel_Scn_Revenue=1,Sel_Scenario,Sel_Scn_Revenue-1\)\)/);
  const items = layout.sheetRows().get('Lookups')!.filter(r => /^lookups\/List_Scenario_Use\/\d+$/.test(r.id));
  assert.equal(items.length, 4);
});

test('the band, the shaded scenario column and the counts are conditional formats', () => {
  const xml = scenarioSheetXml();
  assert.match(xml, /<c r="J5"[^>]*t="(?:s|inlineStr|str)"/, 'the band names each scenario');
  assert.match(xml, /<c r="I6"[^>]*>/, 'the active tab sits over column I');
  assert.match(xml, /<conditionalFormatting sqref="J6:L6"><cfRule type="expression"[^>]*><formula>COLUMNS\(\$J6:J6\)=Sel_Scenario<\/formula>/);
  assert.match(xml, /<conditionalFormatting sqref="J7:L7"><cfRule type="expression"[^>]*><formula>COLUMNS\(\$J7:J7\)=Sel_Scenario<\/formula>/);
  const shaded = [...xml.matchAll(/<formula>IF\(Sel_Scn_\w+=1,Sel_Scenario,Sel_Scn_\w+-1\)=COLUMNS\(\$J\d+:J\d+\)<\/formula>/g)];
  assert.equal(shaded.length, 18, 'every line shades the column of the scenario it uses');
  assert.match(xml, /<formula>Sel_Scenario=COLUMNS\(\$J\d+:J\d+\)<\/formula>/, 'the names row shades the active scenario');
  assert.match(xml, /<mergeCells count="3">(<mergeCell ref="I\d+:L\d+"\/>){3}<\/mergeCells>/, 'descriptions run across the scenario columns');
  assert.match(xml, /<pane ySplit="8" topLeftCell="A9"/, 'the band stays in view');
});

test('results sit under a heading per measure inside the data table, and repeated figures share a row', () => {
  const layout = assemble(full());
  const rows = layout.sheetRows().get('Scenarios')!;
  const t = layout.dataTables[0];
  const ids = rows.map(r => r.id);
  assert.equal(ids.indexOf(t.first), ids.indexOf('scenarios/results/input') + 1, 'the input row sits directly above the table');
  const inTable = rows.slice(ids.indexOf(t.first), ids.indexOf(t.last) + 1);
  for (const r of inTable) assert.ok(9 in r.cells, `${r.id}: every row in the table has a formula in column I`);
  const formulas = inTable.filter(r => r.role === 'r.result').map(r => r.cells[9]);
  assert.equal(new Set(formulas).size, formulas.length);
  assert.ok(inTable.some(r => r.role === 'r.group' && r.label === 'Revenue, each financial year'));
});

test('a structural change keeps the names, descriptions and adjustments someone typed', () => {
  const before = full();
  const after = before.copy();
  after.insert('fm.revenue', {}, 'Advisory fees');
  const plan = planChange(assemble(before), assemble(after));
  const writes = plan.ops.filter(o => o.op === 'write' && o.sheet === 'Scenarios') as { row: number; cells: Record<number, unknown>; why: string }[];
  assert.ok(writes.length > 0);
  const layout = assemble(after);
  const pos = layout.positions();
  const at = (id: string) => pos.get(id)![1];
  for (const w of writes) {
    if (w.why !== 'rewire') continue;
    if (w.row === at('scenarios/names')) assert.ok(!(10 in w.cells), 'names kept');
    for (let k = 1; k <= 3; k++) if (w.row === at(`scenarios/desc/${k}`)) assert.ok(!(9 in w.cells), 'description kept');
  }
  const lines = (l: typeof layout) => l.sheetRows().get('Scenarios')!.filter(r => r.role === 'line');
  assert.equal(lines(layout).length, lines(assemble(before)).length + 1);
  const revenue = lines(layout).filter(r => /fm\.revenue#/.test(r.id));
  assert.equal(revenue.length, 4, 'the new line joins the Revenue group');
});
