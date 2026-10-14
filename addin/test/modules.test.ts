// The module commands (Kelvin, 15 October 2026: build the module commands so testing can start):
// Add above, Duplicate, Rename, Delete and Properties on the budget model, the drop-downs that
// follow their list's items when a line moves, changes its title or goes, and the same commands
// through the model service outside Excel.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { assemble, Model, TOTAL_COL, type PlanOp, type Recipe } from '../../engine/src/index.ts';
import { loadLibrary } from '../../engine/src/node/library.ts';
import { LIBRARY } from '../../engine/test/demo.ts';
import { cell, image } from '../../engine/test/image.ts';
import { planInsert, planRemove, type OpenModel } from '../src/insert/core.ts';
import { applyPlan } from '../src/live/apply.ts';
import {
  changedLists, dropRows, listItems, planAbove, planDelete, planDuplicate, planProperties, planRename, PROPERTY_ROWS, relist, settingFromCell,
} from '../src/modules/core.ts';
import { ModelService } from '../src/service.ts';
import { fakeExcel } from './fake-excel.ts';

const ROOT = join(import.meta.dirname, '..', '..');
const hfg = loadLibrary(join(ROOT, 'library', 'hfg'));
const demo = loadLibrary(LIBRARY);
const recipe = JSON.parse(readFileSync(join(ROOT, 'library', 'hfg', 'recipes', 'budget_model.json'), 'utf8')) as Recipe;
const fresh = (): OpenModel => { const model = Model.fromRecipe(hfg, recipe); return { model, layout: assemble(model) }; };
const BUDGET = fresh();
const budget = (): OpenModel => ({ model: BUDGET.model.copy(), layout: BUDGET.layout });

const of = (open: OpenModel, module: string) => open.model.instances.filter(i => i.module === module);
/** What each drop-down in the lists a plan changes holds, as positions (a file built from the model) or as text (written live). */
const held = (open: OpenModel, p: Parameters<typeof relist>[1], text = false) => new Map(dropRows(open.layout, changedLists(open, p)).map(r => {
  const list = r.control!.kind === 'drop' ? r.control!.list : '';
  return [r.id, text ? listItems(open.layout, open.model, list)[(r.value as number) - 1]?.label : r.value];
}));
const writes = (ops: PlanOp[]) => ops.filter((o): o is Extract<PlanOp, { op: 'write' }> => o.op === 'write' && o.why === 'rewire' && !o.format);

/** The budget model with its first cost of sales lines choosing revenue lines 2, 3 and 5. */
function choosing(): OpenModel {
  const m = BUDGET.model.copy();
  const cogs = m.instances.filter(i => i.module === 'fm.cogs');
  cogs[0].settings.source = 2;
  cogs[1].settings.source = 3;
  cogs[2].settings.source = 5;
  return { model: m, layout: assemble(m) };
}

test('Add above puts the new line above the selected one, and drop-downs keep the lines they chose', () => {
  const open = choosing();
  const revs = of(open, 'fm.revenue');
  const p = planAbove(open, revs[2].uid, { bu: 2 });
  const order = p.next.instances.filter(i => i.module === 'fm.revenue').map(i => i.uid);
  assert.equal(order.indexOf(p.uid!), 2, 'the new line is third, above the line selected');
  assert.equal(p.next.instance(p.uid!).settings.bu, 2);
  assert.deepEqual(changedLists(open, p), ['List_Revenue_Lines']);
  const moved = relist(open, p, held(open, p));
  assert.equal(moved, 2, 'the lines that chose revenue lines 3 and 5 now choose 4 and 6; line 2 stays');
  const cogs = p.next.instances.filter(i => i.module === 'fm.cogs');
  assert.deepEqual(cogs.slice(0, 3).map(c => c.settings.source), [2, 4, 6]);
  const w = writes(p.plan.ops);
  const pos = p.layout.positions();
  assert.deepEqual(w.map(o => [o.sheet, o.row, o.cells[TOTAL_COL]]),
    [cogs[1], cogs[2]].map((c, k) => [...pos.get(`${c.uid}/set.source`)!, [4, 6][k]]));
  assert.match(p.plan.preview.at(-1)!, /2 drop-downs keep their choice/);
});

test('Add above goes in after a list\'s fixed first items and a line added at the end moves nothing', () => {
  const open = budget();
  const debtors = of(open, 'fm.collection');
  const p = planAbove(open, debtors[0].uid, {});
  const items = listItems(p.layout, p.next, 'List_Collection');
  assert.equal(items[0].label, 'Received on sale', 'the fixed first item stays first');
  assert.equal(items[1].id, p.uid);
  assert.deepEqual(changedLists(open, planInsert(open, 'fm.collection', {})), [], 'a line added at the end moves no item');
});

test('removing a line: drop-downs that chose it take the first item and say so, later ones close up', () => {
  const open = choosing();
  const revs = of(open, 'fm.revenue');
  const p = planRemove(open, revs[2].uid);
  relist(open, p, held(open, p));
  const cogs = p.next.instances.filter(i => i.module === 'fm.cogs');
  assert.deepEqual(cogs.slice(0, 3).map(c => c.settings.source), [2, 1, 4]);
  assert.ok(p.plan.preview.some(l => l.includes(`chose ${open.model.title(revs[2])}, which goes; it will choose ${open.model.title(revs[0])}`)));
});

test('rename: drop-downs written live hold their item\'s text, so they take the new title; positions need nothing', () => {
  const open = choosing();
  const revs = of(open, 'fm.revenue');
  const p = planRename(open, revs[2].uid, 'Insurance repairs');
  assert.equal(p.next.title(p.next.instance(revs[2].uid)), 'Insurance repairs');
  assert.equal(relist(open, p, held(open, p)), 0, 'a position still points at the same line');
  const q = planRename(open, revs[2].uid, 'Insurance repairs');
  assert.equal(relist(open, q, held(open, q, true)), 1);
  assert.deepEqual(writes(q.plan.ops).map(o => o.cells[TOTAL_COL]), ['Insurance repairs']);
  assert.throws(() => planRename(open, revs[2].uid, open.model.title(revs[3])), /already called/);
  assert.throws(() => planRename(open, revs[2].uid, '  '), /Type the new title/);
  assert.throws(() => planRename(open, of(open, 'fm.statements')[0].uid, 'Accounts'), /keeps its library title/);
});

test('duplicate: a copy straight after the line, with its settings and month amounts, titled (copy), (copy 2)', () => {
  const open = budget();
  const src = of(open, 'fm.opex')[0];
  const series = { amount: Array.from({ length: open.model.periods }, (_, t) => 100 + t) };
  const p = planDuplicate(open, src.uid, { ...src.settings }, series);
  const opex = p.next.instances.filter(i => i.module === 'fm.opex');
  assert.equal(opex[1].uid, p.uid, 'placed after the line it copies');
  assert.equal(p.title, `${open.model.title(src)} (copy)`);
  assert.deepEqual(p.next.instance(p.uid!).data.series, series);
  assert.equal(p.next.instance(p.uid!).data.history, undefined, 'historical actuals stay with the original');
  const again = planDuplicate({ model: p.next, layout: p.layout }, src.uid, { ...src.settings }, series);
  assert.equal(again.title, `${open.model.title(src)} (copy 2)`);
  assert.throws(() => planDuplicate(open, of(open, 'fm.gst')[0].uid, {}, {}), /goes in once/);
});

test('delete: any module goes with what it sends; one the model cannot do without is refused with the reason', () => {
  const open = budget();
  const p = planDelete(open, of(open, 'fm.interest_cash')[0].uid);
  assert.ok(!p.next.instances.some(i => i.module === 'fm.interest_cash'));
  assert.equal(p.next.modules['fm.interest_cash'], undefined);
  assert.throws(() => planDelete(open, of(open, 'fm.statements')[0].uid), /Financial statements cannot go: Income summary reads the financial statements/);
});

test('properties: title, Prepared by and the status line are written on the Settings sheet; notes are contents rows', () => {
  const open = budget();
  const info = open.model.info!;
  const p = planProperties(open, { title: 'HCL budget FY2028', preparedBy: info.preparedBy, notes: info.notes, display: { errors: true, alerts: false } });
  assert.equal(p.next.info!.title, 'HCL budget FY2028');
  const pos = p.layout.positions();
  assert.deepEqual(writes(p.plan.ops).map(o => [o.sheet, o.row, o.cells[TOTAL_COL]]), [
    [...pos.get(PROPERTY_ROWS.title)!, 'HCL budget FY2028'], [...pos.get(PROPERTY_ROWS.alerts)!, false]]);
  assert.match(p.plan.preview[0], /^Model title: HCL budget FY2028 \(was /);
  // what the workbook holds wins over the metadata: a title already typed on the Settings sheet is not written again
  const q = planProperties(open, { title: 'HCL budget FY2028', preparedBy: info.preparedBy, notes: ['One note'], display: info.display },
    { [PROPERTY_ROWS.title]: 'HCL budget FY2028' });
  assert.equal(writes(q.plan.ops).length, 0);
  assert.ok(q.plan.preview.some(l => l.startsWith('Contents:')));
  assert.throws(() => planProperties(open, { title: info.title, preparedBy: info.preparedBy, notes: info.notes, display: info.display }), /Nothing has changed/);
  assert.throws(() => planProperties(open, { title: ' ', preparedBy: '', notes: [], display: info.display }), /needs a title/);
});

test('settings read from cells: text from a live drop-down becomes its position', () => {
  const open = budget();
  assert.equal(settingFromCell(open.model, open.layout, 'fm.cogs', 'method', 'Amount'), 2);
  assert.equal(settingFromCell(open.model, open.layout, 'fm.cogs', 'source', open.model.title(of(open, 'fm.revenue')[1])), 2);
  assert.equal(settingFromCell(open.model, open.layout, 'fm.cogs', 'source', 3), 3);
  assert.equal(settingFromCell(open.model, open.layout, 'fm.cogs', 'method', ''), 1, 'an empty cell takes the default');
});

test('live: Add above applied to the workbook moves a classic drop-down\'s position with its line', async () => {
  const open = choosing();
  const img = image(open.layout);
  const { ctx } = fakeExcel(img);
  const p = planAbove(open, of(open, 'fm.revenue')[2].uid, {});
  relist(open, p, held(open, p));
  await applyPlan(ctx, p.plan.ops.filter(o => !/chart$|^data_table$/.test(o.op)), p.metadata);
  const cogs = p.next.instances.filter(i => i.module === 'fm.cogs');
  const pos = p.layout.positions();
  const at = (uid: string) => { const [s, r] = pos.get(`${uid}/set.source`)!; return cell(img.sheets.get(s)!, r, TOTAL_COL).v; };
  assert.deepEqual(cogs.slice(0, 3).map(c => at(c.uid)), [2, 4, 6]);
});

test('the service runs the module commands outside Excel on the selection (the first category line stands in)', async () => {
  const svc = new ModelService({ libs: { hfg, demo }, logos: {}, sample: budget, sampleFull: budget });
  const first = budget().model.instances.find(i => hfg.kind(i.module) === 'category')!;
  const above = await svc.addAboveAtSelection();
  assert.ok(above.ok, above.message);
  const dup = await svc.duplicateAtSelection();
  assert.ok(dup.ok, dup.message);
  assert.match(dup.message, /\(copy\) added/);
  const r = await svc.renameState();
  assert.equal(r.uid, first.uid);
  const rp = await svc.renamePreview('Renamed line');
  assert.equal(rp.title, 'Renamed line');
  assert.match((await svc.renameApply()).message, /^Renamed Renamed line/);
  const d = await svc.deletePreview();
  assert.equal(d.category, true);
  assert.match((await svc.deleteApply()).message, /deleted/);
  const props = await svc.propertiesState();
  assert.equal(props.title, 'Maintenance and construction budget');
  assert.ok(props.modules.length > 20 && props.modules.every(m => m.held === m.library));
  assert.deepEqual(props.facts.map(f => f[0]), ['Entity', 'Timeline', 'Financial year ends', 'Actuals to', 'Denomination', 'Library']);
  const pp = await svc.propertiesPreview({ ...props, title: 'New title' });
  assert.match(pp.lines[0], /Model title: New title/);
  assert.match((await svc.propertiesApply()).message, /^Properties saved/);
});
