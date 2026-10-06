// The change plan carries everything the live writer needs in the standard frame. A workbook
// image (cells, formats, row heights, outline, links, checks formatting, validations, names) is
// built from one layout, the plan to the next layout is applied to it the way the live writer
// applies it, and the result must equal the image of a fresh build of the next layout. Formulas
// are left to LibreOffice (tests/test_engine_writer.py and the assembly proof): Excel moves the
// references in rows the plan does not touch, which an image cannot.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { assemble, planChange, type Model } from '../src/index.ts';
import { loadLibrary } from '../src/node/library.ts';
import { demoModel, LIBRARY } from './demo.ts';
import { apply, flat, image } from './image.ts';

const lib = loadLibrary(LIBRARY);

function check(before: Model, change: (m: Model) => void): void {
  const after = before.copy();
  change(after);
  const lo = assemble(before);
  const ln = assemble(after);
  const img = image(lo);
  apply(img, planChange(lo, ln).ops);
  const got = flat(img);
  const want = flat(image(ln));
  assert.deepEqual(got.order, want.order);
  for (const s of want.order) assert.deepEqual(got.sheets[s], want.sheets[s], s);
  assert.deepEqual(got.names, want.names);
}

const base = () => demoModel(lib, 'HF');

test('live: inserting a revenue line and a facility formats itself as a fresh build', () => {
  check(base(), m => { m.insert('demo.revenue_line', { base: 150, growth: 0.02 }); m.insert('demo.facility', { amount: 500, rate: 0.08, instalment: 25 }); });
});

test('live: removing a revenue line moves the dashed rule to the new last item', () => {
  check(base(), m => m.remove('demo.revenue_line#2'));
});

test('live: the first module of a section adds its cover, contents entries and links', () => {
  const m = demoModel(lib, 'HCP');
  m.remove('demo.dashboard#1');
  check(m, x => x.insert('demo.dashboard', { first: 1 }));
});

test('live: a blank model grows into the demo', () => {
  const blank = demoModel(lib, 'TWK');
  for (const i of [...blank.instances]) blank.remove(i.uid);
  check(blank, m => {
    m.insert('demo.statements'); m.insert('demo.checks'); m.insert('demo.revenue_line', { base: 10, growth: 0 });
  });
});

test('live: changing a setting rewrites nothing structural', () => {
  const m = base();
  const lo = assemble(m);
  const after = m.copy();
  after.instance('demo.revenue_line#1').settings.base = 120;
  const plan = planChange(lo, assemble(after));
  assert.ok(!plan.ops.some(o => o.op === 'outline' || o.op === 'insert_rows' || o.op === 'add_sheet'));
});
