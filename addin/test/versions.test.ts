// Saved versions in the add-in: Save version, Approve budget and Lock, as change plans on the full
// model demo.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { assemble, Model, type Recipe } from '../../engine/src/index.ts';
import { loadLibrary } from '../../engine/src/node/library.ts';
import { checksum, planApprove, planLock, planSave, REFORECAST_LINES, suggestions, versionsOf } from '../src/versions/core.ts';

const ROOT = join(import.meta.dirname, '..', '..');
const lib = loadLibrary(join(ROOT, 'library', 'hfg'));
const recipe = JSON.parse(readFileSync(join(ROOT, 'library', 'hfg', 'recipes', 'full_model.json'), 'utf8')) as Recipe;

function open() {
  const model = Model.fromRecipe(lib, recipe);
  return { model, layout: assemble(model) };
}

const flat = (n: number, periods: number) => Object.fromEntries(REFORECAST_LINES.map(k => [k, Array(periods).fill(n)]));

test('Save version offers this month\'s reforecast (replacing an unlocked one) and the budget being built', () => {
  const o = open();
  const s = suggestions(o.model);
  assert.equal(s[0].label, 'Reforecast Sep 2026 (6+6)');
  assert.equal(s[0].replaces, 'V21');
  assert.equal(s[1].label, 'Budget FY2028');
  assert.deepEqual(s[1].months, [25, 36]);
});

test('saving a budget adds a register row, its store rows and a Compared with choice', () => {
  const o = open();
  const budget = suggestions(o.model)[1];
  const p = planSave(o, budget, flat(1000, 36), 'Group Finance', '2026-10-07');
  const v = versionsOf(p.next).at(-1)!;
  assert.equal(v.id, 'V22');
  assert.equal(v.values.rev[23], null);
  assert.equal(v.values.rev[24], 1000);
  assert.equal(v.checksum, checksum(v.values));
  const rows = p.layout.sheetRows();
  assert.equal(rows.get('Versions')!.filter(r => r.role === 'r.reg').length, versionsOf(o.model).length + 1);
  assert.ok(p.plan.preview.some(x => x.startsWith('Version store:')), p.plan.preview.join(' | '));
  assert.ok(p.plan.ops.some(op => op.op === 'add_name' && op.name === 'VR_Label'));   // the store's names grow with inserted rows
});

test('saving this month\'s reforecast again replaces it; approving a budget supersedes the one before', () => {
  const o = open();
  const p = planSave(o, suggestions(o.model)[0], flat(5, 36), 'Group Finance', '2026-10-07');
  const vs = versionsOf(p.next);
  assert.equal(vs.length, versionsOf(o.model).length);
  assert.equal(vs.find(v => v.id === 'V21')!.values.cash[0], 5);
  const draft = versionsOf(o.model).find(v => v.status === 'Superseded')!;
  const a = planApprove(o, draft.id);
  const after = versionsOf(a.next);
  assert.equal(after.find(v => v.id === draft.id)!.status, 'Approved');
  assert.equal(after.find(v => v.id === 'V14')!.status, 'Superseded');
  assert.equal(after.filter(v => v.status === 'Approved' && v.year === 'FY2027').length, 1);
  const l = planLock(o, 'V21', true);
  assert.equal(versionsOf(l.next).find(v => v.id === 'V21')!.locked, true);
  assert.equal(suggestions(l.next)[0].replaces, undefined);
});
