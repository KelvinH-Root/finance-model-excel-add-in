// Behaviour the golden fixtures do not reach: insert rules, binding, errors and the helpers.

import assert from 'node:assert/strict';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import {
  AssemblyError, assemble, code, codeWords, colLetter, Library, Model, namePart, planChange, renderFormula,
  unoSeparators,
} from '../src/index.ts';
import { loadLibrary } from '../src/node/library.ts';

const LIBRARY = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'prototypes', 'assembly', 'library');
const lib = loadLibrary(LIBRARY);

test('helpers', () => {
  assert.equal(colLetter(1), 'A');
  assert.equal(colLetter(26), 'Z');
  assert.equal(colLetter(27), 'AA');
  assert.equal(colLetter(703), 'AAA');
  assert.equal(namePart('lending_rate'), 'LendingRate');
  assert.equal(namePart('Opening Balances'), 'Opening balances');
  assert.equal(code('demo.revenue_line#1'), 'demo_revenue_line_1');
  assert.equal(code('Kāinga Maha'), 'Kāinga_Maha');
  assert.equal(codeWords('Closing cash (peak)'), 'ClosingCashPeak');
  assert.equal(unoSeparators('=IF(A1,"a, b",2)'), '=IF(A1;"a, b";2)');
});

test('insert: categories number on, singles once, unknown settings refused', () => {
  const m = new Model(lib);
  assert.equal(m.insert('demo.revenue_line').uid, 'demo.revenue_line#1');
  assert.equal(m.insert('demo.revenue_line', { base: 50 }).settings.base, 50);
  m.remove('demo.revenue_line#2');
  assert.equal(m.insert('demo.revenue_line').uid, 'demo.revenue_line#3', 'numbers are never reused');
  m.insert('demo.statements');
  assert.throws(() => m.insert('demo.statements'), AssemblyError);
  assert.throws(() => m.insert('demo.revenue_line', { colour: 'red' }), /has no setting colour/);
  assert.throws(() => m.insert('demo.nothing'), /no module/);
  assert.throws(() => m.remove('demo.debtors#1'), /no instance/);
});

test('a bound setting writes the group formula; unbinding brings back a typed value', () => {
  const facility = lib.module('demo.facility');
  const bindable = (facility.settings || []).find(s => s.group);
  assert.ok(bindable, 'the demo facility has a setting with a group assumption');
  const m = new Model(lib);
  m.insert('demo.statements');
  const f = m.insert('demo.facility', { amount: 1000, rate: 0.06, instalment: 50 });
  const before = assemble(m);
  m.bind(f.uid, bindable.key);
  const after = assemble(m);
  const row = after.sheets.flatMap(([, rows]) => rows).find(r => r.id === `${f.uid}/set.${bindable.key}`)!;
  assert.match(row.link || '', /GA_/);
  assert.equal(row.value, null);
  const write = planChange(before, after).ops.find(o => o.op === 'write' && o.why === 'rewire' && o.kind === 'setting');
  assert.ok(write && write.op === 'write' && typeof write.cells[9] === 'string', 'the rewire writes the formula over the value');
  m.unbind(f.uid, bindable.key, 0.065);
  const back = planChange(after, assemble(m)).ops.find(o => o.op === 'write' && o.kind === 'setting');
  assert.ok(back && back.op === 'write' && back.cells[9] === 0.065, 'unbinding writes the local value');
  assert.throws(() => m.bind(f.uid, 'nothing'), /has no setting/);
});

test('a typed input survives a structural change', () => {
  const m = new Model(lib);
  m.insert('demo.statements');
  m.insert('demo.revenue_line', { base: 100, growth: 0.01 });
  const before = assemble(m);
  m.insert('demo.revenue_line', { base: 60, growth: 0.03 });
  const plan = planChange(before, assemble(m));
  const touched = plan.ops.filter(o => o.op === 'write' && o.why === 'rewire' && o.kind === 'setting');
  for (const o of touched) assert.ok(o.op === 'write' && !(9 in o.cells));
  assert.ok(plan.preview.some(p => p.startsWith('Revenue: ')), plan.preview.join('\n'));
});

test('render: markers in both dialects', () => {
  const pos = new Map<string, [string, number]>([['a', ['Revenue', 7]], ['b', ['Working capital', 9]]]);
  assert.equal(renderFormula('=«R|a»+«P|a»', 'Revenue', 10, pos), '=J7+0');
  assert.equal(renderFormula('=«R|a»+«P|a»', 'Revenue', 11, pos), '=K7+J7');
  assert.equal(renderFormula('=SUM(«A|b»)', 'Revenue', 10, pos, 'excel', 3), "=SUM('Working capital'!$J$9:$L$9)");
  assert.equal(renderFormula('=INDEX(«A|a»,1,«T»)', 'Costs', 10, pos, 'uno', 3), '=INDEX($Revenue.$J$7:$L$7;1;3)');
  assert.equal(renderFormula('=«S|Revenue»&" "&«N»', 'Costs', 12, pos), '=Revenue!$B$1&" "&3');
  assert.throws(() => renderFormula('=«R|zz»', 'Revenue', 10, pos), /no row zz/);
});
