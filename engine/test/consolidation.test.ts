// The group consolidation model type (framework: consolidation), built from the fictional group in
// examples/consolidation (recipe consolidation_model.json). The figures are compared with the Python
// reference month by month in tests/test_consolidation_model.py (LibreOffice); these tests cover the
// layout: its sheets, names and lists, the model's own timeline, and the errors a bad group raises.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { assemble, buildWorkbook, Model, type Recipe } from '../src/index.ts';
import { colLetter, FIRST_PERIOD_COL } from '../src/frame.ts';
import { loadLibrary } from '../src/node/library.ts';
import type { GroupData } from '../src/consolidation.ts';

const ROOT = join(import.meta.dirname, '..', '..');
const lib = loadLibrary(join(ROOT, 'library', 'hfg'));
const recipe = () => JSON.parse(readFileSync(join(ROOT, 'library', 'hfg', 'recipes', 'consolidation_model.json'), 'utf8')) as Recipe;
const groupOf = (r: Recipe) => r.instances[0].data!.group as GroupData;
const B = 'fm.consolidation#1';

test('the consolidation lays out its statements, registers and workings in the Group section', () => {
  const layout = assemble(Model.fromRecipe(lib, recipe()));
  assert.deepEqual(layout.sheets.map(([s]) => s), ['Contents', 'Group', 'Group statements', 'By group', 'Entities', 'Entity data',
    'Intercompany', 'Intergroup balances', 'Sites', 'Unrealised margin', 'Eliminations', 'NCI', 'Appendices', 'Settings', 'Lookups', 'Checks']);
  assert.equal(layout.kindOf('Entities'), 'table');
  assert.equal(layout.kindOf('Intercompany'), 'table');
  assert.equal(layout.kindOf('By group'), 'timeline');
  assert.equal(layout.totalHeads['By group'], 'Opening');
  assert.deepEqual(layout.warnings, []);
  for (const nm of ['Cn_Chk_TB', 'Cn_Chk_Tree', 'Cn_Chk_Join', 'Cn_Chk_Pairs', 'Cn_Chk_Inv', 'Cn_Chk_Elim', 'Cn_Chk_Cons', 'Cn_Chk_NCI',
    'Cn_Chk_Top', 'Cn_Chk_CF', 'Cn_Alert_Planned', 'Cn_Alert_Netting', 'Sel_Cn1_Group', 'Chk_Errors', 'Chk_Alerts']) {
    assert.ok(layout.names.has(nm), nm);
  }
  for (const nm of ['List_Groups', 'List_IC_Types', 'ED_Mov', 'EL_Val', 'MG_Val', 'NCI_9000_Val', 'Ent_In_9008']) assert.ok(layout.ranges.has(nm), nm);
  // five groups: the four that produce accounts and Partner LP, an NCI node
  const lookups = new Map(layout.sheets)!.get('Lookups')!;
  assert.deepEqual(lookups.filter(r => /^lookups\/List_Groups\/\d+$/.test(r.id)).map(r => r.cells[4]),
    ['Foundation group', 'Holdings group', 'Devco group', 'Fund group', 'Partner LP']);
  // the statements choose the group shown; every group's consolidation is on By group
  const gs = new Map(layout.sheets).get('Group statements')!;
  const cash = gs.find(r => r.id === `${B}/gs/bs/ca/1000`)!;
  assert.match(cash.formula!, /^=CHOOSE\(Sel_Cn1_Group,«R\|fm\.consolidation#1\/bg\/9000\/cons\/1000»,/);
  const by = new Map(layout.sheets).get('By group')!;
  assert.equal(by.filter(r => r.kind === 'heading' && / \(head \d+\)$/.test(r.label)).length, 5);
  const bytes = buildWorkbook(layout, Model.fromRecipe(lib, recipe()), { created: new Date('2026-10-26T00:00:00Z') });
  assert.ok(bytes.length > 100_000);
});

test('the timeline is the model\'s own: a shorter one moves the keys past it', () => {
  const r = recipe();
  r.timeline.start = '2026-07';
  r.periods = 24;
  const layout = assemble(Model.fromRecipe(lib, r));
  const ed = layout.ranges.get('ED_Mov')!;
  assert.equal(ed.toCol, 'timeline');
  assert.equal(layout.periods, 24);
  const keys = FIRST_PERIOD_COL + 24 + 1;
  assert.deepEqual(layout.ranges.get('ED_Ent'), { ...layout.ranges.get('ED_Ent')!, col: keys });
  assert.deepEqual(layout.heads['Entity data'].slice(0, 2), [[5, keys, 'Entity'], [5, keys + 1, 'Account']]);
  assert.equal(colLetter(keys), 'AI');
});

test('a group the consolidation cannot work with is refused with the reason', () => {
  const bad = (change: (g: GroupData) => void) => {
    const r = recipe();
    change(groupOf(r));
    return () => assemble(Model.fromRecipe(lib, r));
  };
  assert.throws(bad(g => { g.entities = []; }), /has no entities/);
  assert.throws(bad(g => { g.entities[3].parent = 9999; }), /entity 9003: no entity 9999 to be its parent/);
  assert.throws(bad(g => { g.entities[2].code = 9001; }), /two entities share a code/);
  assert.throws(bad(g => {
    for (const e of g.entities) { delete e.group; e.held = 1; }
  }), /no entity produces group accounts/);
  assert.throws(bad(g => { g.transactions![1].site = 'SZ'; }), /no site SZ/);
});
