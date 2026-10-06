// The Office.js live writer, against a stand-in for Excel working on the engine's workbook image:
// a plan applied by applyPlan must leave the workbook as a fresh build of the new model writes it
// (named styles, row heights, outline, links, check formatting, validations, names, metadata).

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { assemble, metadataXml, planChange, type Model } from '../../engine/src/index.ts';
import { loadLibrary } from '../../engine/src/node/library.ts';
import { demoModel, LIBRARY } from '../../engine/test/demo.ts';
import { flat, image } from '../../engine/test/image.ts';
import { applyPlan, asTyped, splitRef, validationRule, widthToPoints } from '../src/live/apply.ts';
import { fakeExcel } from './fake-excel.ts';

const lib = loadLibrary(LIBRARY);

/** Validations in the form the stand-in records them (Office.js rule and message); rows the plan left alone keep the image's form. */
function normalise(f: ReturnType<typeof flat>) {
  for (const s of Object.values(f.sheets) as { valid: Record<string, string[]> }[]) {
    for (const [r, list] of Object.entries(s.valid)) {
      s.valid[r] = list.map(entry => {
        const json = entry.slice(entry.indexOf('|') + 1);
        if (!json.startsWith('{"kind"')) return entry;
        const v = JSON.parse(json);
        return `${entry.slice(0, entry.indexOf('|'))}|${JSON.stringify(validationRule(v))}|${v.message}`;
      });
    }
  }
  return f;
}

async function check(before: Model, change: (m: Model) => void) {
  const after = before.copy();
  change(after);
  const lo = assemble(before);
  const ln = assemble(after);
  const img = image(lo);
  const { ctx, log } = fakeExcel(img);
  const plan = planChange(lo, ln, 'excel').ops.filter(o => !o.op.endsWith('_chart'));
  const report = await applyPlan(ctx, plan, metadataXml(after, ln));
  const got = normalise(flat(img));
  const want = normalise(flat(image(ln)));
  assert.deepEqual(got.order, want.order);
  for (const s of want.order) assert.deepEqual(got.sheets[s], want.sheets[s], s);
  assert.deepEqual(got.names, want.names);
  assert.equal(log.parts.length, 1);
  assert.match(log.parts[0], /^<\?xml[\s\S]*<hfgModel xmlns="urn:hfg:model-metadata:v0"/);
  return { report, log };
}

test('live writer: inserting a revenue line and a facility', async () => {
  const { report, log } = await check(demoModel(lib, 'HF'), m => {
    m.insert('demo.revenue_line', { base: 150, growth: 0.02 });
    m.insert('demo.facility', { amount: 500, rate: 0.08, instalment: 25 });
  });
  assert.ok(log.calls.some(c => c.startsWith('Revenue insert ')));
  assert.ok(report.syncs >= 2);
});

test('live writer: removing a module', async () => {
  await check(demoModel(lib, 'KM'), m => m.remove('demo.revenue_line#2'));
});

test('live writer: the first module of a section brings its cover', async () => {
  const m = demoModel(lib, 'HCL');
  m.remove('demo.dashboard#1');
  await check(m, x => x.insert('demo.dashboard', { first: 1 }));
});

test('live writer: a blank model grows into the demo', async () => {
  const blank = demoModel(lib, 'TWK');
  for (const i of [...blank.instances]) blank.remove(i.uid);
  await check(blank, m => { m.insert('demo.statements'); m.insert('demo.checks'); m.insert('demo.revenue_line', { base: 10, growth: 0 }); });
});

test('helpers', () => {
  assert.equal(asTyped('1.'), "'1.");
  assert.equal(asTyped('a.'), 'a.');
  assert.equal(asTyped('TRUE'), "'TRUE");
  assert.equal(asTyped('12/3'), "'12/3");
  assert.deepEqual(splitRef("'Working capital'!$J$20:$O$20"), ['Working capital', '$J$20:$O$20']);
  assert.deepEqual(splitRef('Dashboard!$C$13'), ['Dashboard', '$C$13']);
  assert.deepEqual(splitRef("'Bob''s'!A1"), ["Bob's", 'A1']);
  assert.equal(widthToPoints(11.75), 56.63);
});
