// Module versions (13 October 2026): a model records the library version of each module it holds,
// so a later library can find the models built from an older version and bring them forward with a
// change plan. Every instance of a module is built from the same version.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { Model, type Recipe } from '../src/index.ts';
import { loadLibrary } from '../src/node/library.ts';
import { LIBRARY, demoModel } from './demo.ts';

const ROOT = join(import.meta.dirname, '..', '..');
const hfg = () => loadLibrary(join(ROOT, 'library', 'hfg'));
const recipe = JSON.parse(readFileSync(join(ROOT, 'library', 'hfg', 'recipes', 'budget_model.json'), 'utf8')) as Recipe;

test('a model records the version of every module it holds, in its metadata and back', () => {
  const lib = hfg();
  const m = Model.fromRecipe(lib, recipe);
  const d = m.toDict();
  const held = [...new Set(m.instances.map(i => i.module))];
  assert.deepEqual(Object.keys(d.modules!), held, 'one entry per module, in the order first inserted');
  assert.ok(Object.values(d.modules!).every(v => v === 1), 'version 1 when a module does not say');
  const back = Model.fromDict(lib, JSON.parse(JSON.stringify(d)));
  assert.deepEqual(back.toDict().modules, d.modules);
  assert.deepEqual(back.outdated(), []);
});

test('a library update shows which modules a model holds at an older version', () => {
  const old = Model.fromRecipe(hfg(), recipe).toDict();
  const lib = hfg();
  lib.module('fm.revenue').version = 2;
  const m = Model.fromDict(lib, old);
  assert.deepEqual(m.outdated(), [{ module: 'fm.revenue', title: 'Revenue line', held: 1, library: 2 }]);
  assert.throws(() => m.insert('fm.revenue'), /holds version 1 of this module and the library has version 2/);
  m.insert('fm.cogs');   // a module the update leaves alone still goes in
  assert.equal(m.toDict().modules!['fm.cogs'], 1);
});

test('metadata from before the record reads as version 1, and a module removed drops its entry', () => {
  const d = Model.fromRecipe(hfg(), recipe).toDict();
  delete d.modules;
  const lib = hfg();
  lib.module('fm.debt').version = 3;
  const m = Model.fromDict(lib, d);
  assert.equal(m.versionHeld('fm.debt'), 1);
  assert.equal(m.versionHeld('fm.intangible_asset_none'), null);
  assert.deepEqual(m.outdated().map(x => x.module), ['fm.debt']);
  for (const i of m.instances.filter(x => x.module === 'fm.debt')) m.remove(i.uid);
  assert.equal(m.versionHeld('fm.debt'), null);
  assert.ok(!('fm.debt' in (m.toDict().modules ?? {})));
  m.insert('fm.debt');   // with none left, the library's version goes in
  assert.equal(m.toDict().modules!['fm.debt'], 3);
});

test('the Phase 0 proof keeps its metadata as it was', () => {
  const d = demoModel(loadLibrary(LIBRARY), null).toDict();
  assert.equal(d.modules, undefined);
  assert.equal(d.library, undefined);
});
