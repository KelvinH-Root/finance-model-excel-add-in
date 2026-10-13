// The add-in's dialogs (Kelvin, 13 October 2026: commands open where they belong: a dialog, the
// selection, the task pane or a sheet). The model service the dialogs call, the Explorer on the
// budget model, removing and adding a category line, and the message round trip between a
// dialog and the runtime that opened it.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { assemble, Model, type Recipe } from '../../engine/src/index.ts';
import { loadLibrary } from '../../engine/src/node/library.ts';
import { LIBRARY, demoModel } from '../../engine/test/demo.ts';
import { explorerTree, moduleInfo, ownerOfRow } from '../src/explorer/core.ts';
import { planRemove } from '../src/insert/core.ts';
import { openDialog, remote } from '../src/rpc.ts';
import { ModelService } from '../src/service.ts';

const ROOT = join(import.meta.dirname, '..', '..');
const hfg = loadLibrary(join(ROOT, 'library', 'hfg'));
const demo = loadLibrary(LIBRARY);
const recipe = JSON.parse(readFileSync(join(ROOT, 'library', 'hfg', 'recipes', 'budget_model.json'), 'utf8')) as Recipe;
const budget = () => { const model = Model.fromRecipe(hfg, recipe); return { model, layout: assemble(model) }; };
const small = () => { const model = demoModel(demo, 'HF'); return { model, layout: assemble(model) }; };
const service = () => new ModelService({ libs: { hfg, demo }, logos: {}, sample: small, sampleFull: budget });

test('the Explorer tree is the model: sections, sheets, categories with their lines, single modules', () => {
  const t = explorerTree(budget());
  assert.equal(t.title, 'Maintenance and construction budget');
  const model = t.sections.find(s => s.title === 'Financial Model')!;
  const rev = model.sheets.find(s => s.sheet === 'Revenue and expenses')!;
  assert.deepEqual(rev.entries.map(e => e.title).slice(0, 2), ['Revenue', 'Cost of sales']);
  const revenue = rev.entries[0];
  assert.equal(revenue.kind, 'category');
  assert.equal(revenue.modules.length, 5);
  assert.equal(revenue.modules[0].title, '4100 Reactive maintenance, Auckland');
  const wc = model.sheets.find(s => s.sheet === 'Working capital')!;
  assert.deepEqual(wc.entries.find(e => e.title === 'Debtors')!.modules.map(m => m.title), ['Trade debtors', 'Maintenance clients', 'Capital works clients']);
  assert.ok(t.sections.some(s => s.title === 'Dashboards'));
});

test('a module in the Explorer: its rows, what it takes from and sends to, settings in words and checks', () => {
  const o = budget();
  const uid = o.model.instances.find(i => i.name === 'Maintenance clients')!.uid;
  const m = moduleInfo(o, uid);
  assert.equal(m.title, 'Maintenance clients');
  assert.match(m.where, /^Working capital rows \d+ to \d+$/);
  assert.ok(m.components[0].rows.some(r => r.label === 'Days to collect' && r.what === 'Setting'));
  assert.ok(m.components[0].rows.some(r => r.label === 'Receipts' && r.what === 'Calculation'));
  assert.ok(m.takes.some(r => r.title === '4110 Planned maintenance, Auckland'));
  assert.ok(m.sends.length > 0, 'it sends receipts and closing debtors');
  assert.deepEqual(m.settings, [{ label: 'Days to collect', value: '30 days' }]);
  assert.deepEqual(m.checks.map(c => c.kind), ['alert', 'error']);
  const line = o.model.instances.find(i => i.name === '4100 Reactive maintenance, Auckland')!.uid;
  const s = moduleInfo(o, line).settings;
  assert.equal(s.find(x => x.label === 'Driver')!.value, 'Annual budget spread by a seasonality profile');
  assert.equal(s.find(x => x.label === 'Collected through')!.value, 'Trade debtors');   // after Received on sale
  assert.equal(ownerOfRow(o, `${line}/revenue`).uid, line);
  assert.equal(ownerOfRow(o, 'cat/fm.revenue').module, 'fm.revenue');
});

test('removing a category line plans its rows away and the totals close up', () => {
  const o = budget();
  const uid = o.model.instances.find(i => i.name === 'Workshop and office')!.uid;
  const p = planRemove(o, uid);
  assert.equal(p.title, 'Workshop and office');
  assert.ok(!p.next.instances.some(i => i.uid === uid));
  assert.ok(p.plan.ops.some(op => op.op === 'delete_rows'));
  assert.ok(p.plan.preview.some(l => /rows removed/.test(l)), p.plan.preview.join(' | '));
});

test('the service outside Excel: insert choices, a preview, then the insert on the demo model', async () => {
  const svc = service();
  const c = await svc.insertChoices();
  assert.match(c.source, /demo model/);
  const debt = c.choices.find(x => x.id === 'demo.facility')!;
  assert.equal(debt.category, true);
  assert.ok(debt.settings.some(s => s.unit === '%'));
  assert.equal(c.choices.find(x => x.id === 'demo.statements')!.blocked, 'Already in the model; it goes in once.');
  const p = await svc.insertPreview('demo.facility', Object.fromEntries(debt.settings.map(s => [s.key, s.value])), 'Equipment loan');
  assert.equal(p.title, 'Equipment loan');
  assert.ok(p.lines.length > 2);
  const r = await svc.insertApply();
  assert.match(r.message, /^Equipment loan inserted: \d+ operations/);
  await assert.rejects(() => svc.insertApply(), /Nothing to insert/);
});

test('the service outside Excel: Add category copies the line under the selection; Remove asks first', async () => {
  const svc = service();
  const r = await svc.addCategoryAtSelection();
  assert.equal(r.ok, true, r.message);
  assert.match(r.message, /Revenue line 3 added/);
  const svc2 = service();
  const rp = await svc2.removePreview();
  assert.equal(rp.title, 'Revenue line 1');
  assert.ok(rp.lines[0].startsWith('What it sends stops:'));
  assert.match((await svc2.removeApply()).message, /^Revenue line 1 removed/);
});

test('the service outside Excel: versions on the budget demo, and the Explorer tree with no check status', async () => {
  const svc = service();
  const v = await svc.versionsState();
  assert.deepEqual(v.options.map(o => o.label), ['Reforecast Sep 2026 (6+6)', 'Budget FY2028']);
  assert.ok(v.register.some(x => x.status === 'Approved'));
  const p = await svc.versionsPreview({ kind: 'save', option: 1, savedBy: 'Group Finance' });
  assert.equal(p.title, 'Budget FY2028');
  const t = await svc.explorerTree();
  assert.deepEqual(t.status, {});
  const m = await svc.explorerModule(t.sections[1].sheets[0].entries[0].modules[0].uid);
  assert.ok(m.raised.every(x => x === null));
  const c = await svc.checksState();
  assert.equal(c.read, false);
  assert.ok(c.groups.length >= 2 && c.groups.every(g => g.checks.length && g.goName.startsWith('Go_')), JSON.stringify(c.groups.map(g => g.title)));
  assert.equal(await svc.selectRow(c.groups[0].checks[0].row), false);   // nothing to select outside Excel
});

test('a dialog calls the runtime by message and gets the answer back; close and open go to the runtime', async () => {
  // One stand-in for Office plays both sides: the runtime's dialog object and the dialog's own Office.context.ui.
  const toParent: ((arg: { message: string }) => void)[] = [];
  const toChild: ((arg: { message: string }) => void)[] = [];
  const closed: string[] = [];
  const dialog = {
    addEventHandler: (type: string, fn: (arg: { message: string }) => void) => { if (type === 'msg') toParent.push(fn); },
    messageChild: (m: string) => toChild.forEach(f => f({ message: m })),
    close: () => closed.push('closed'),
  };
  (globalThis as Record<string, unknown>).Office = {
    AsyncResultStatus: { Succeeded: 'ok' },
    EventType: { DialogMessageReceived: 'msg', DialogEventReceived: 'event', DialogParentMessageReceived: 'parent' },
    context: {
      requirements: { isSetSupported: () => true },
      ui: {
        displayDialogAsync: (_url: string, _o: unknown, cb: (r: unknown) => void) => cb({ status: 'ok', value: dialog }),
        addHandlerAsync: (_t: string, fn: (arg: { message: string }) => void) => toChild.push(fn),
        messageParent: (m: string) => setTimeout(() => toParent.forEach(f => f({ message: m })), 0),
      },
    },
  };
  const opened: string[] = [];
  await openDialog('https://localhost/hfg-dialog.html?view=mod-insert', [50, 50], {
    call: async (m, a) => { if (m === 'boom') throw new Error('no such thing'); return { m, a }; },
    open: k => opened.push(k),
  });
  const svc = remote<{ hello(x: number): Promise<{ m: string; a: unknown[] }>; boom(): Promise<void> }>();
  assert.deepEqual(await svc.hello(3), { m: 'hello', a: [3] });
  await assert.rejects(() => svc.boom(), /no such thing/);
  const ui = (globalThis as unknown as { Office: { context: { ui: { messageParent(m: string): void } } } }).Office.context.ui;
  ui.messageParent(JSON.stringify({ t: 'open', key: 'grp-add' }));
  await new Promise(r => setTimeout(r, 500));
  assert.deepEqual(closed, ['closed']);
  assert.deepEqual(opened, ['grp-add']);
  delete (globalThis as Record<string, unknown>).Office;
});
