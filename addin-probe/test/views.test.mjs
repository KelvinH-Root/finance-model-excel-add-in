import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const read = rel => readFileSync(new URL(rel, import.meta.url), 'utf8');

function load() {
  const context = vm.createContext({ console });
  for (const f of ['model-sample.js', 'model-calc.js', 'view-kit.js', 'explorer.js', 'impacts.js', 'commands.js']) {
    vm.runInContext(read(`../src/${f}`), context);
  }
  return context;
}
const arr = (xs, f) => (f ? Array.from(xs, f) : Array.from(xs));                   // plain arrays: the modules run in their own realm
const close = (a, b, tol = 1e-6) => Math.abs(a - b) < tol;
const KEYS = ['revenue', 'opex', 'interest', 'receipts', 'cash', 'assets', 'debt', 'equity'];

test('the sample model\'s calculation matches the assembly reference, with and without a change', () => {
  const { HfgModelSample: M, HfgModelCalc: C } = load();
  const check = (got, want, what) => KEYS.forEach(k => got.lines[`demo.statements#1/${k}`].forEach((v, t) =>
    assert.ok(close(v, want[k][t]), `${what} ${k} month ${t + 1}: ${v} against ${want[k][t]}`)));
  check(C.calc(M), M.reference.base, 'base');
  M.reference.cases.forEach(cs => check(C.calc(M, { [cs.name]: cs.value }), cs.reference, cs.name));
});

test('the Explorer tree follows the sections and sheets, and Links names what each module takes and sends', () => {
  const { HfgModelSample: M, HfgExplorer: X } = load();
  const tree = X.tree(M);
  assert.deepEqual(arr(tree, s => s.title), ['Dashboards', 'Financial Model', 'Appendices']);
  assert.deepEqual(arr(tree[1].sheets[0].modules), ['demo.revenue_line#1', 'demo.revenue_line#2']);
  const debtors = X.relations(M, 'demo.debtors#1');
  assert.deepEqual(arr(debtors.takes, r => `${r.title}: ${arr(r.links).join()}`), ['Revenue line 1: is.revenue', 'Revenue line 2: is.revenue']);
  assert.deepEqual(arr(debtors.sends, r => `${r.title}: ${arr(r.links).join()}`), ['Financial statements: cf.receipts,bs.debtors']);
  const facility = X.relations(M, 'demo.facility#1');
  assert.equal(facility.takes.length, 0);
  assert.deepEqual(arr(facility.sends, r => r.title), ['Financial statements', 'Checks']);
  const statuses = X.checkStatus(M, 'demo.statements#1', load().HfgModelCalc.calc(M));
  assert.deepEqual(arr(statuses, s => `${s.label}:${s.raised}`), ['Balance sheet does not balance:0', 'Cash below zero:0']);
});

test('Impact of a change: every tie holds for every input, and the movement matches the reference', () => {
  const { HfgModelSample: M, HfgImpacts: I, HfgModelCalc: C } = load();
  const names = arr(M.modules.flatMap(m => arr(m.settings, s => [s.name, s.value])));
  for (const [name, value] of names) {
    for (const month of [1, 6, 12]) {
      const r = I.liveImpact(M, name, value * 1.25 + 0.01, month);
      assert.ok(close(r.ties.surplus.carried, r.ties.surplus.equity), `${name} surplus tie`);
      assert.ok(close(r.ties.cash.flows, r.ties.cash.cash), `${name} cash tie`);
      assert.ok(close(r.ties.balance, 0), `${name} balance`);
    }
  }
  M.reference.cases.forEach(cs => {
    const r = I.liveImpact(M, cs.name, cs.value, 12);
    KEYS.forEach(k => {
      const want = cs.reference[k][11] - M.reference.base[k][11];
      assert.ok(close(r.rows.find(x => x.id === `demo.statements#1/${k}`).month, want), `${cs.name} ${k}`);
    });
  });
  assert.equal(C.settingsOf(M).Rev1_Base, 100);                       // the sample is left as it was
  assert.deepEqual(arr(I.chain(M, 'demo.revenue_line#1'), e => arr(e).join(' > ')).slice(0, 2),
    ['Revenue line 1 > is.revenue > Debtors: Revenue line 1', 'Revenue line 1 > is.revenue > Financial statements']);
});

test('Impacts sheets: the pane works every item out as the Python reference does, in both contexts', () => {
  const { HfgImpactsSample: S, HfgImpacts: I } = load();
  assert.deepEqual(arr(S.contexts, c => c.items.length), [4, 11]);
  for (const ctx of S.contexts) {
    for (const item of ctx.items) {
      const { panels } = I.panels(item, ctx, {}, S.cashClasses);
      for (const [col, want] of Object.entries(item.expected)) {
        for (const [k, v] of Object.entries(want)) {
          assert.ok(close(panels[col][k] || 0, v), `${ctx.model} ${item.key} ${col} ${k}: ${panels[col][k]} against ${v}`);
        }
      }
    }
  }
  const claim = S.contexts[1].items.find(i => i.key === 'claim');
  const unpaid = I.panels(claim, S.contexts[1], { settled: 0 }, S.cashClasses).panels;
  assert.ok(close(unpaid.Group.wip, 850) && close(unpaid.Group.payables, 0) && close(unpaid.Group.receivables, 0));
});

test('Analysis > Impacts and the right-click item open the Impacts views; Explorer and Links open the Explorer', () => {
  const c = load();
  const C = c.HfgCommands;
  assert.deepEqual(arr(C.all().filter(x => x.parent === 'an-impacts'), x => x.key), ['imp-live', 'imp-sheets']);
  assert.ok(C.CONTEXT.includes('ctx-impacts'));
  // render needs a DOM, so check the keys each view answers to
  const src = read('../src/impacts.js') + read('../src/explorer.js');
  ['an-impacts', 'imp-live', 'imp-sheets', 'ctx-impacts', 'model-explorer', 'b-explorer', 'mod-links'].forEach(k => assert.ok(src.includes(`'${k}'`), k));
});
