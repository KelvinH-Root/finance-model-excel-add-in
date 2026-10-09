// Calculation sheets closer to the reference models (Kelvin, 9 October 2026): a category's lines sit
// under one bar for the category, each line a subsection with its sections as bold group labels, and
// a summary of the lines with their total after the last; the contents list categories rather than
// every line; checks sit under a label for their module; a driver the chosen method does not use is
// greyed, label and all; the statements carry group labels and gaps before their key results.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { assemble, buildWorkbook, Model, planChange, type Recipe } from '../src/index.ts';
import { loadLibrary } from '../src/node/library.ts';

const ROOT = join(import.meta.dirname, '..', '..');
const lib = loadLibrary(join(ROOT, 'library', 'hfg'));
const recipe = (name: string) => JSON.parse(readFileSync(join(ROOT, 'library', 'hfg', 'recipes', `${name}.json`), 'utf8')) as Recipe;
const budget = () => Model.fromRecipe(lib, recipe('budget_model'));
const full = () => Model.fromRecipe(lib, recipe('full_model'));

test('a category has one bar, each line a subsection, and a summary of its lines with their total', () => {
  const rows = assemble(budget()).sheetRows().get('Revenue and expenses')!;
  const bars = rows.filter(r => r.kind === 'heading').map(r => r.label);
  assert.deepEqual(bars, ['Revenue', 'Cost of sales', 'Operating expenses', 'Staff costs', 'Other income', 'Other expenses']);
  const lines = rows.filter(r => r.kind === 'subheading' && r.id.endsWith('/heading'));
  assert.equal(lines.length, 25, 'every GL line is a subsection');
  assert.ok(lines.every(r => r.level === 0), 'a line stays in view when its body is collapsed');
  const groups = rows.filter(r => r.kind === 'section');
  assert.ok(groups.length && groups.every(r => r.role === 'group' || r.role === 'working'), 'a line\'s sections are group labels');
  const ids = rows.map(r => r.id);
  const summary = ids.indexOf('cat/fm.revenue/summary');
  assert.ok(summary > ids.indexOf('fm.revenue#5/end') && summary < ids.indexOf('cat/fm.cogs'), 'the summary follows the last line');
  const items = rows.filter(r => r.id.startsWith('cat/fm.revenue/summary/fm.revenue#'));
  assert.equal(items.length, 5);
  assert.equal(items[0].formula, '=«R|fm.revenue#1/revenue»');
  assert.equal(items[0].cells[3], '=«B|fm.revenue#1/heading»', 'labelled by the line\'s title');
  const total = rows.find(r => r.id === 'cat/fm.revenue/summary/total')!;
  assert.equal(total.label, 'Total revenue');
  assert.equal(total.formula, '=SUM(«R|cat/fm.revenue/summary/fm.revenue#1»:«R|cat/fm.revenue/summary/fm.revenue#5»)');
  assert.ok(!rows.some(r => r.id === 'cat/fm.other_income/summary'), 'one line needs no summary');
});

test('the contents list categories, and every line keeps its Go_ name', () => {
  const layout = assemble(budget());
  const contents = layout.sheetRows().get('Contents')!;
  const entries = contents.filter(r => r.role === 'toc3').map(r => r.label);
  assert.ok(entries.includes('- Revenue') && entries.includes('- Debtors') && entries.includes('- Capital expenditure'));
  assert.ok(!entries.some(e => /4100 Reactive/.test(e)), 'not each line');
  assert.ok(!entries.includes('- Income summary'), 'a sheet holding one module of its own name lists the sheet only');
  assert.equal(layout.names.get('Go_cat_fm_revenue'), 'cat/fm.revenue');
  assert.equal(layout.names.get('Go_fm_revenue_1'), 'fm.revenue#1/heading');
});

test('a line inserted later joins its category, so the plan only inserts rows', () => {
  const before = full();
  const after = before.copy();
  after.insert('fm.revenue', {}, 'Advisory fees');
  const rows = assemble(after).sheetRows().get('Revenue and expenses')!;
  const ids = rows.map(r => r.id);
  const added = rows.find(r => r.label === 'Advisory fees' && r.kind === 'subheading')!;
  const at = ids.indexOf(added.id);
  assert.ok(at > ids.indexOf('cat/fm.revenue') && at < ids.indexOf('cat/fm.revenue/summary'), 'under the Revenue bar, before its summary');
  assert.ok(ids.some(id => id.startsWith('cat/fm.revenue/summary/') && id.endsWith(added.id.replace('/heading', ''))));
  const plan = planChange(assemble(before), assemble(after));
  assert.ok(plan.ops.some(o => o.op === 'insert_rows' && o.sheet === 'Revenue and expenses'));
});

test('checks sit under a label for their module, which links to it', () => {
  const rows = assemble(budget()).sheetRows().get('Checks')!;
  const labels = rows.filter(r => r.kind === 'section' && r.role === 'group');
  assert.ok(labels.length > 20);
  const plant = labels.find(r => r.label === 'Plant and fleet')!;
  assert.match(String(plant.cells[3]), /^=«B\|fm\.fixed_asset#1\/heading»$/);
  assert.equal(plant.links?.[3].to, 'Go_fm_fixed_asset_1');
  const item = rows[rows.indexOf(plant) + 1];
  assert.equal(item.label, 'Book values do not roll forward in actual months', 'no title before the check');
});

test('a driver the chosen method does not use greys out, label and all', () => {
  const model = budget();
  const layout = assemble(model);
  const rows = layout.sheetRows().get('Revenue and expenses')!;
  const amount = rows.find(r => r.id === 'fm.revenue#1/amount')!;
  assert.ok(amount.dim && /Sel_Rev1_Method/.test(amount.dim));
  const parts = unzipSync(buildWorkbook(layout, model, { created: new Date('2026-10-09T00:00:00Z') }));
  const n = layout.sheets.findIndex(([s]) => s === 'Revenue and expenses') + 1;
  const xml = strFromU8(parts[`xl/worksheets/sheet${n}.xml`]);
  const r = layout.positions().get('fm.revenue#1/amount')![1];
  assert.match(xml, new RegExp(`<conditionalFormatting sqref="C${r}:I${r}"><cfRule type="expression"[^>]*><formula>NOT\\(Sel_Rev1_Method=1\\)</formula>`));
});

test('the statements carry group labels and gaps before their key results', () => {
  const rows = assemble(budget()).sheetRows().get('Financials')!;
  const groups = rows.filter(r => r.role === 'group').map(r => r.label);
  assert.deepEqual(groups.slice(0, 7), ['Revenue', 'Cost of sales', 'Other income', 'Staff costs', 'Operating expenses', 'Other expenses',
    'Depreciation and amortisation']);
  const ids = rows.map(r => r.id);
  const gross = ids.indexOf('fm.statements#1/gross');
  assert.equal(rows[gross - 1].kind, 'blank', 'a gap before gross margin');
  const cash = rows.find(r => r.id === 'fm.statements#1/cash')!;
  assert.equal(cash.indent, 2, 'a line under a group label sits in from it');
});
