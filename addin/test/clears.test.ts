// A rewired row's old cells: removing a revenue line from the full model shrinks the report tables,
// and the live writer must clear the columns a rewired row no longer holds (in Excel they kept a
// formula that read #REF!, found by the Excel self-test on 11 November 2026). Linked cells keep
// their formula after their hyperlink is set.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { assemble, Model, type Recipe } from '../../engine/src/index.ts';
import { loadLibrary } from '../../engine/src/node/library.ts';
import { image } from '../../engine/test/image.ts';
import { planRemove } from '../src/insert/core.ts';
import { applyPlan } from '../src/live/apply.ts';
import { withClears, type WriteOp } from '../src/live/clears.ts';
import { fakeExcel } from './fake-excel.ts';

const ROOT = join(import.meta.dirname, '..', '..');
const hfg = loadLibrary(join(ROOT, 'library', 'hfg'));
const recipe = JSON.parse(readFileSync(join(ROOT, 'library', 'hfg', 'recipes', 'full_model.json'), 'utf8')) as Recipe;

test('removing a line clears the columns rewired report rows no longer hold', () => {
  const model = Model.fromRecipe(hfg, recipe);
  const open = { model, layout: assemble(model) };
  const lines = model.instances.filter(i => i.module === 'fm.revenue');
  const p = planRemove(open, lines[lines.length - 1].uid);
  withClears(open.layout, p.layout, p.plan);
  const cleared = p.plan.ops.filter((o): o is WriteOp => o.op === 'write' && !!(o as WriteOp).clear?.length);
  assert.ok(cleared.some(o => o.sheet === 'Income report'), cleared.map(o => `${o.sheet} ${o.row}`).join(', '));
  for (const o of cleared) for (const c of o.clear!) assert.ok(!Object.hasOwn(o.cells, c), 'a cleared column is not also written');
});

test('the live writer clears those columns, and a linked cell keeps its formula after its link', async () => {
  const model = Model.fromRecipe(hfg, recipe);
  const open = { model, layout: assemble(model) };
  const lines = model.instances.filter(i => i.module === 'fm.revenue');
  const p = planRemove(open, lines[lines.length - 1].uid);
  withClears(open.layout, p.layout, p.plan);
  const img = image(open.layout);
  const { ctx } = fakeExcel(img);
  // the stand-in keeps the last value written; a hyperlink set without text would put its target there in Excel
  await applyPlan(ctx, p.plan.ops.filter(o => !/chart$|^data_table$/.test(o.op)), '<?xml version="1.0"?><hfgModel xmlns="urn:hfg:model-metadata:v0">{}</hfgModel>');
  const want = image(p.layout);
  for (const o of p.plan.ops.filter((x): x is WriteOp => x.op === 'write' && !!(x as WriteOp).clear?.length)) {
    for (const c of o.clear!) {
      assert.equal(img.sheets.get(o.sheet)!.cells.get(o.row)?.get(c)?.v, want.sheets.get(o.sheet)!.cells.get(o.row)?.get(c)?.v, `${o.sheet} row ${o.row} column ${c}`);
    }
  }
});
