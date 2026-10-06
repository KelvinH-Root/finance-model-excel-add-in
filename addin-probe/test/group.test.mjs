import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const read = rel => readFileSync(new URL(rel, import.meta.url), 'utf8');

function load() {
  const context = vm.createContext({ console });
  vm.runInContext(read('../src/group-sample.js'), context);
  vm.runInContext(read('../src/group.js'), context);
  return { G: context.HfgGroup, sample: context.HfgGroupSample };
}

const names = gs => Array.from(gs, g => g.name);      // plain arrays: the module runs in its own realm
const arr = xs => Array.from(xs);
const close = (a, b) => Math.abs(a - b) < 1e-6;

test('the tree, ownership and the groups each entity rolls into', () => {
  const { G, sample } = load();
  const st = G.makeState(sample);
  assert.deepEqual(names(G.groupsOf(st, 9009)), ['Fund group', 'Holdings group', 'Foundation group']);
  assert.deepEqual(names(G.groupsOf(st, 9001)), ['Foundation group']);
  assert.ok(close(G.ownedByTop(st, 9009), 0.6));
  assert.equal(G.parentGroup(st, 9005).name, 'Holdings group');
  assert.equal(G.parentGroup(st, 9000), null);
  assert.deepEqual(names(G.childGroups(st, 9004)), ['Devco group', 'Fund group', 'Partner LP']);
  assert.deepEqual(arr(G.directMembers(st, 9004)), [9004]);
  assert.deepEqual(arr(G.directMembers(st, 9000)), [9000, 9001, 9002, 9003]);
  assert.equal(G.depth(st, 9006), 3);
});

test('every roll-up adds up to its group\'s consolidated figures, every group and year', () => {
  const { G, sample } = load();
  const st = G.makeState(sample);
  for (const g of st.groups) {
    for (let p = 0; p < st.years.length; p++) {
      const r = G.rollup(st, g.head, p);
      assert.ok(close(r.total.surplus, g.surplus[p]), `${g.name} ${st.years[p]} surplus ${r.total.surplus} ${g.surplus[p]}`);
      assert.ok(close(r.total.na, g.na[p]), `${g.name} ${st.years[p]} net assets`);
      assert.ok(close(r.owners.na + r.nci.na, g.na[p]));
    }
  }
  // Holdings FY2028: Holdings' own figures, three sub-groups consolidated, and what is eliminated in Holdings itself
  const r = G.rollup(st, 9004, 2);
  assert.deepEqual(Array.from(r.lines, l => l.kind), ['entity', 'group', 'group', 'group', 'elim']);
  assert.ok(close(r.nci.na, 1207.6));
});

test('Add entity puts a planned entity under its parent, in tree order, and a partly held one becomes a group', () => {
  const { G, sample } = load();
  const st = G.makeState(sample);
  const base = { from: '1 April 2027', gst: false, role: 'Limited partnership' };
  assert.deepEqual(arr(G.validate(st, Object.assign({ code: '9006', name: 'X', parent: 9005, held: 100 }, base))), ['Code 9006 is already in the group.']);
  assert.equal(G.validate(st, Object.assign({ code: '9011', name: 'Dev LP C', parent: 9005, held: 0 }, base)).length, 1);
  const v = G.preview(st, Object.assign({ code: '9012', name: 'Fund Two LP', parent: 9004, held: 50 }, base));
  assert.ok(v.node && close(v.outside, 0.5));
  assert.deepEqual(arr(v.into), ['Fund Two LP', 'Holdings group', 'Foundation group']);
  assert.ok(G.addEntity(st, Object.assign({ code: '9011', name: 'Dev LP C', parent: 9005, held: 100 }, base)).ok);
  assert.ok(G.addEntity(st, Object.assign({ code: '9012', name: 'Fund Two LP', parent: 9004, held: 50 }, base)).ok);
  assert.deepEqual(Array.from(st.entities, e => e.code), [9000, 9001, 9002, 9003, 9004, 9005, 9006, 9007, 9011, 9008, 9009, 9010, 9012]);
  assert.equal(st.entities.find(e => e.code === 9011).status, 'Planned');
  assert.deepEqual(names(G.groupsOf(st, 9011)), ['Devco group', 'Holdings group', 'Foundation group']);
  assert.deepEqual(names(G.childGroups(st, 9004)), ['Devco group', 'Fund group', 'Partner LP', 'Fund Two LP']);
  // Planned entities have no figures yet, so every group still adds up
  for (const g of st.groups) {
    const r = G.rollup(st, g.head, 2);
    assert.ok(close(r.total.na, G.figures(st, g.head, 2).na), g.name);
  }
});

test('every Group menu item opens a designed view or a described placeholder', () => {
  const { G } = load();
  const context = vm.createContext({});
  vm.runInContext(read('../src/commands.js'), context);
  const menu = context.HfgCommands.find('model-group');
  assert.equal(menu.type, 'menu');
  const items = Array.from(context.HfgCommands.all().filter(c => c.parent === 'model-group'), c => c.key);
  assert.deepEqual(items, ['grp-structure', 'grp-add', 'grp-ownership', 'grp-remove', 'grp-refresh']);
  ['model-group', 'grp-structure', 'grp-add', 'grp-ownership'].forEach(k => assert.ok(G.VIEWS.includes(k), k));
});
