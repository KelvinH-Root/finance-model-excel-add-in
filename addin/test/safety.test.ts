// Safety of change in the model service, against the stand-in for Excel: every change reads the key
// outputs before and after and logs itself in the model's metadata; a change that stops part way
// leaves a mark, keeps a copy of the workbook to open, and the add-in refuses further changes until
// it is put right; unlocking an approved budget needs a reason, which the log keeps.

import assert from 'node:assert/strict';
import { afterEach, beforeEach, test } from 'node:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { assemble, Model, type Recipe } from '../../engine/src/index.ts';
import { loadLibrary } from '../../engine/src/node/library.ts';
import { image } from '../../engine/test/image.ts';
import { metadataXml } from '../../engine/src/xlsx/package.ts';
import { changeLog, entryFor, expectedOutputs, logChange, LOG_LIMIT, resultLines, type LogEntry } from '../src/guard/core.ts';
import { planInsert, readModel } from '../src/insert/core.ts';
import { applyPlan } from '../src/live/apply.ts';
import { PENDING_NS, ModelService } from '../src/service.ts';
import { planLock, versionsOf } from '../src/versions/core.ts';
import { fakeExcel } from './fake-excel.ts';

const ROOT = join(import.meta.dirname, '..', '..');
const hfg = loadLibrary(join(ROOT, 'library', 'hfg'));
const recipe = JSON.parse(readFileSync(join(ROOT, 'library', 'hfg', 'recipes', 'budget_model.json'), 'utf8')) as Recipe;
const budget = () => { const model = Model.fromRecipe(hfg, recipe); return { model, layout: assemble(model) }; };

const g = globalThis as unknown as Record<string, unknown>;
let created: string[] = [];

/** The stand-in for Excel draws no charts (the live tests leave them out too). */
const noCharts = (svc: ModelService) => { svc.pending!.plan.ops = svc.pending!.plan.ops.filter(o => !/chart$|^data_table$/.test(o.op)); };

/** The budget model open in the stand-in for Excel, with the globals the service looks for. */
function inFakeExcel() {
  const open = budget();
  const { ctx, log } = fakeExcel(image(open.layout));
  log.parts.push(metadataXml(open.model, open.layout));
  created = [];
  g.Excel = {
    run: (fn: (c: unknown) => unknown) => fn(ctx),
    CalculationType: { recalculate: 'Recalculate', full: 'Full' },
    createWorkbook: async (b64: string) => { created.push(b64); },
  };
  g.Office = {
    FileType: { Compressed: 'compressed' },
    context: {
      requirements: { isSetSupported: () => true },
      document: {
        getFileAsync(_t: string, _o: unknown, cb: (r: unknown) => void) {
          cb({ status: 'succeeded', value: {
            sliceCount: 1,
            getSliceAsync: (_i: number, cb2: (r: unknown) => void) => cb2({ status: 'succeeded', value: { data: [80, 75, 3, 4], index: 0 } }),
            closeAsync: (cb3: () => void) => cb3(),
          } });
        },
      },
    },
  };
  const svc = new ModelService({ libs: { hfg }, logos: {}, sample: budget, sampleFull: budget });
  const meta = () => readModel(log.parts.find(x => x.includes('<hfgModel'))!, { hfg });
  return { svc, log, meta };
}

beforeEach(() => { delete g.Excel; delete g.Office; });
afterEach(() => { delete g.Excel; delete g.Office; });

test('a change logs itself: the command, what it wrote and the key outputs before and after', async () => {
  const { svc, meta, log } = inFakeExcel();
  await svc.load();
  await svc.insertPreview('fm.opex', {}, 'Self-test costs');
  noCharts(svc);
  const done = await svc.insertApply();
  assert.match(done.message, /No key output moved\./);
  const entries = changeLog(meta().model);
  assert.equal(entries.length, 1);
  assert.equal(entries[0].command, 'Insert');
  assert.equal(entries[0].title, 'Self-test costs');
  assert.ok(Object.keys(entries[0].outputs).includes('Closing cash'), Object.keys(entries[0].outputs).join(', '));
  assert.ok(entries[0].expected.includes('Net profit after tax for the period'), 'a new cost line can move profit');
  assert.ok(!log.parts.some(x => x.includes(PENDING_NS)), 'the in-progress mark is taken off');
});

test('a change that stops part way leaves a mark, keeps a copy, and the add-in refuses until it is put right', async () => {
  const { svc, log } = inFakeExcel();
  await svc.load();
  await svc.insertPreview('fm.opex', {}, 'Broken costs');
  noCharts(svc);
  const ops = svc.pending!.plan.ops;
  ops.splice(Math.floor(ops.length / 2), 0, { op: 'delete_sheet', sheet: 'No such sheet' });
  await assert.rejects(svc.insertApply(), /stopped part way/);
  assert.ok(log.parts.some(x => x.includes(PENDING_NS)), 'the mark stays in the workbook');
  assert.ok(svc.broken?.copy);
  await assert.rejects(svc.insertPreview('fm.opex', {}, 'Another'), /no longer matches its model/);

  // the mark is read back from the workbook by a fresh runtime too
  const again = new ModelService({ libs: { hfg }, logos: {}, sample: budget, sampleFull: budget });
  await again.load();
  assert.equal(again.broken?.title, 'Insert Broken costs');
  assert.equal(again.broken?.copy, false);
  const props = await again.propertiesState();
  assert.equal(props.broken?.title, 'Insert Broken costs');

  await svc.restoreCopy();
  assert.equal(created.length, 1, 'the copy opens in a new window');
  assert.equal(created[0], btoa(String.fromCharCode(80, 75, 3, 4)));
  await svc.markRepaired();
  assert.ok(!log.parts.some(x => x.includes(PENDING_NS)));
  await svc.load();
  await svc.insertPreview('fm.opex', {}, 'After the repair');
});

test('unlocking an approved budget needs a reason, and the change log keeps it', async () => {
  const open = budget();
  const approved = versionsOf(open.model).find(v => v.status === 'Approved' && v.locked)!;
  assert.ok(approved, 'the demo has an approved budget');
  assert.throws(() => planLock(open, approved.id, false), /Give a reason/);
  assert.throws(() => planLock(open, approved.id, false, 'x'.repeat(201)), /200 characters/);
  const draft = versionsOf(open.model).find(v => v.status !== 'Approved');
  if (draft) planLock(open, draft.id, !draft.locked);   // other versions lock and unlock without one

  const { svc, meta } = inFakeExcel();
  await svc.versionsState();
  await assert.rejects(svc.versionsPreview({ kind: 'lock', id: approved.id, locked: false }), /Give a reason/);
  const shown = await svc.versionsPreview({ kind: 'lock', id: approved.id, locked: false, reason: 'Board asked for a revised FY2027 budget' });
  assert.match(shown.lines[0], /Board asked/);
  await svc.versionsApply();
  const e = changeLog(meta().model).at(-1)!;
  assert.equal(e.command, 'Versions');
  assert.equal(e.effect, 'none');
  assert.equal(e.reason, 'Board asked for a revised FY2027 budget');
});

test('the live writer calculates by hand while it applies, puts calculation back, and names an operation Excel refuses', async () => {
  const open = budget();
  const p = planInsert(open, 'fm.opex', {}, 'More costs');
  const ops = p.plan.ops.filter(o => !/chart$|^data_table$/.test(o.op));
  const { ctx } = fakeExcel(image(open.layout));
  const app = (ctx.workbook as unknown as { application: { calculationMode: string } }).application;
  const seen: string[] = [];
  const sync = ctx.sync.bind(ctx);
  ctx.sync = async () => { seen.push(app.calculationMode); await sync(); };
  await applyPlan(ctx, ops, metadataXml(p.next, p.layout));
  assert.ok(seen.length > 3 && seen.slice(1, -1).every(m => m === 'Manual'), 'manual while the plan applies');
  assert.equal(app.calculationMode, 'Automatic', 'put back afterwards');

  const bad = [...ops];
  bad.splice(5, 0, { op: 'delete_sheet', sheet: 'No such sheet' });
  const again = fakeExcel(image(open.layout));
  await assert.rejects(applyPlan(again.ctx, bad, metadataXml(p.next, p.layout), { stepwise: true }), /at operation 6 of \d+: delete_sheet No such sheet/);
  assert.equal((again.ctx.workbook as unknown as { application: { calculationMode: string } }).application.calculationMode, 'Automatic',
    'put back when a plan stops too');
});

test('what a change can reach, what moved without cause, and how the result reads', () => {
  const open = budget();
  const p = planInsert(open, 'fm.opex', {}, 'More costs');
  const expected = expectedOutputs(p.layout, p.plan, 'reach');
  assert.ok(expected.includes('EBITDA for the period'));
  assert.deepEqual(expectedOutputs(p.layout, p.plan, 'none'), []);
  const e = entryFor({ command: 'Rename', title: 'x', effect: 'none', plan: p.plan, layout: p.layout,
    before: { 'Closing cash': 100, 'Peak debt': 50 }, after: { 'Closing cash': 90, 'Peak debt': 50.001 } });
  assert.deepEqual(e.moved, ['Closing cash']);
  assert.deepEqual(e.unexplained, ['Closing cash']);
  const lines = resultLines(e);
  assert.match(lines[0], /^Check this change: Closing cash moved, and this command should not move any key output\./);
  assert.match(lines[1], /^Closing cash: 100\.00 to 90\.00 \(-10\.00\)$/);
  const still = entryFor({ command: 'Add', title: 'x', effect: 'reach', plan: p.plan, layout: p.layout, before: { a: 1 }, after: { a: 1 } });
  assert.deepEqual(resultLines(still), ['No key output moved.']);

  const m = open.model.copy();
  for (let k = 0; k < LOG_LIMIT + 5; k++) logChange(m, { ...still, title: String(k) } as LogEntry);
  assert.equal(changeLog(m).length, LOG_LIMIT);
  assert.equal(changeLog(m).at(-1)!.title, String(LOG_LIMIT + 4));
});
