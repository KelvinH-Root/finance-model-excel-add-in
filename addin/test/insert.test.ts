// Insert module: reading the model from a workbook's metadata, what can go in, and the plan.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { strFromU8, unzipSync } from 'fflate';
import { assemble, buildWorkbook } from '../../engine/src/index.ts';
import { loadLibrary } from '../../engine/src/node/library.ts';
import { demoModel, LIBRARY } from '../../engine/test/demo.ts';
import { choices, planInsert, readModel } from '../src/insert/core.ts';

const lib = loadLibrary(LIBRARY);

function open() {
  const m = demoModel(lib, 'HF');
  const bytes = buildWorkbook(assemble(m), m, { created: new Date('2026-10-07T00:00:00Z') });
  return readModel(strFromU8(unzipSync(bytes)['customXml/item1.xml']), lib);
}

test('the model comes back from the workbook it was built into', () => {
  const o = open();
  assert.equal(o.model.instances.length, 8);
  assert.equal(o.model.info?.entity.brand, 'HF');
  assert.equal(o.layout.frame.id, 'standard');
  assert.throws(() => readModel('<other/>', lib), /no HFG model metadata/);
});

test('modules that go in once are offered only while absent', () => {
  const c = choices(lib, open().model);
  assert.match(c.find(x => x.id === 'demo.statements')!.blocked ?? '', /Already in the model/);
  assert.equal(c.find(x => x.id === 'demo.revenue_line')!.blocked, null);
  assert.equal(c.find(x => x.id === 'demo.assurance')!.blocked, null);
  assert.deepEqual(c.map(x => x.area).filter((a, i, all) => all.indexOf(a) === i).slice(0, 3), ['Revenue', 'Costs', 'Working capital']);
});

test('inserting a revenue line plans rows on four sheets and new metadata', () => {
  const p = planInsert(open(), 'demo.revenue_line', { base: 80, growth: 0.04 });
  assert.equal(p.title, 'Revenue line 3');
  assert.ok(p.plan.preview.some(l => l.startsWith('Revenue: 4 new rows')));
  assert.ok(p.plan.preview.some(l => l.startsWith('Working capital:')));
  assert.ok(p.plan.ops.every(o => o.op !== 'write' || o.format), 'every write carries its formats');
  assert.match(p.metadata, /demo.revenue_line/);
  assert.equal(readModel(p.metadata, lib).model.instances.length, 9);
});
