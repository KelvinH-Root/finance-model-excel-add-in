import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import vm from 'node:vm';

const read = rel => readFileSync(new URL(rel, import.meta.url), 'utf8');

// A stand-in for Office: requirement sets come from `sets`, settings live in memory.
function loadProbe(sets = {}) {
  const store = new Map();
  const associated = [];
  const Office = {
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    context: {
      diagnostics: { host: 'Excel', platform: 'PC', version: '16.0.19231.20156' },
      officeTheme: { isDarkTheme: false, bodyBackgroundColor: '#ffffff' },
      requirements: { isSetSupported: (name, version) => (sets[name] || []).includes(version) },
      document: {
        settings: {
          set: (k, v) => store.set(k, v), get: k => store.get(k), remove: k => store.delete(k),
          saveAsync: cb => cb({ status: 'succeeded' })
        }
      }
    },
    actions: { associate: (name, fn) => associated.push(name) }
  };
  const context = vm.createContext({ Office, console, setTimeout, performance, navigator: { userAgent: 'test' } });
  vm.runInContext(read('../src/commands.js'), context);
  vm.runInContext(read('../src/probe.js'), context);
  return { probe: context.HfgProbe, associated, commands: context.HfgCommands };
}

const ALL = {
  ExcelApi: ['1.1', '1.4', '1.5', '1.7', '1.8', '1.9', '1.10', '1.12', '1.13', '1.14', '1.15', '1.16', '1.17', '1.18', '1.19', '1.20', '1.21'],
  SharedRuntime: ['1.1'], RibbonApi: ['1.1'], ContextMenuApi: ['1.1'], KeyboardShortcuts: ['1.1'],
  NestedAppAuth: ['1.1'], DialogApi: ['1.1']
};

test('every probe has a unique id, and every automatic probe can run', () => {
  const { probe } = loadProbe();
  const ids = probe.PROBES.map(p => p.id);
  assert.equal(new Set(ids).size, ids.length);
  probe.PROBES.filter(p => p.auto).forEach(p => assert.equal(typeof p.run, 'function', p.id));
});

test('the environment probe passes only when every required set is present', async () => {
  let { probe } = loadProbe(ALL);
  let r = await probe.runProbe('env');
  assert.equal(r.status, 'pass', r.detail);
  ({ probe } = loadProbe({ ExcelApi: ['1.1', '1.17'] }));
  r = await probe.runProbe('env');
  assert.equal(r.status, 'warn');
  assert.match(r.detail, /ExcelApi 1\.18/);
});

test('a probe whose requirement set is missing is skipped, not failed', async () => {
  const { probe } = loadProbe({ ExcelApi: ['1.1'] });
  const r = await probe.runProbe('trace');
  assert.equal(r.status, 'skip');
  assert.match(r.detail, /ExcelApi 1\.12/);
});

test('document settings probe round-trips three sizes', async () => {
  const { probe } = loadProbe(ALL);
  const r = await probe.runProbe('settings');
  assert.equal(r.status, 'pass', r.detail);
  assert.match(r.detail, /3,000,000 characters saved/);
});

test('a confirmation keeps the probe detail and records the answer', async () => {
  const { probe } = loadProbe(ALL);
  probe.record('pdf', { status: 'manual', detail: 'PDF of 10 bytes.' });
  probe.confirm('pdf', false, 'the PDF ignored print areas');
  const r = probe.state.results.get('pdf');
  assert.equal(r.status, 'fail');
  assert.match(r.detail, /PDF of 10 bytes\. User confirmed: the PDF ignored print areas\./);
});

test('every manifest function and shortcut action has a handler', () => {
  const { probe, associated } = loadProbe(ALL);
  probe.registerCommands();
  const manifest = read('../manifest.xml');
  const functions = [...manifest.matchAll(/<FunctionName>([^<]+)<\/FunctionName>/g)].map(m => m[1]);
  const shortcuts = JSON.parse(read('../shortcuts.json')).actions.map(a => a.id);
  [...functions, ...shortcuts].forEach(name => assert.ok(associated.includes(name), `no handler for ${name}`));
});

test('the task pane has every element the UI script uses, and every confirm names a probe', () => {
  const { probe } = loadProbe(ALL);
  const html = read('../src/taskpane.html');
  const ui = read('../src/ui.js');
  const ids = new Set([...html.matchAll(/id="([^"]+)"/g)].map(m => m[1]));
  const used = [...ui.matchAll(/(?:\$|bind)\('([^']+)'/g)].map(m => m[1]);
  used.forEach(id => assert.ok(ids.has(id), `taskpane.html has no #${id}`));
  const probeIds = new Set(probe.PROBES.map(p => p.id));
  [...html.matchAll(/data-probe="([^"]+)"/g)].forEach(m => assert.ok(probeIds.has(m[1]), m[1]));
  [...ui.matchAll(/runProbe\('([^']+)'/g)].forEach(m => assert.ok(probeIds.has(m[1]), m[1]));
});

test('manifest ids used by ribbon and menu updates exist', () => {
  const manifest = read('../manifest.xml');
  const probeSource = read('../src/probe.js');
  ['HFG.Probe.Tab', 'HFG.Probe.Group', 'HFG.Probe.Toggle', 'HFG.Probe.Ctx.Mark'].forEach(id => {
    assert.ok(manifest.includes(`id="${id}"`), `manifest has no ${id}`);
    assert.ok(probeSource.includes(`'${id}'`), `probe.js does not use ${id}`);
  });
});

test('every designed command has a label, a supertip within Office limits, icons and a view', () => {
  const { commands: C } = loadProbe(ALL);
  assert.ok(C, 'commands.js did not load');
  const all = C.all();
  const keys = all.map(c => c.key);
  assert.equal(new Set(keys).size, keys.length, 'duplicate command keys');
  all.forEach(c => {
    assert.ok(c.label && c.label.length <= 32, `${c.key}: label`);
    assert.ok(c.tip && c.tip.length <= 250, `${c.key}: supertip ${c.tip && c.tip.length}`);
    assert.ok(c.view, `${c.key}: view text`);
    if (!c.parent) [16, 32, 80].forEach(s => assert.ok(existsSync(new URL(`../assets/cmd/${c.key}-${s}.png`, import.meta.url)), `${c.key}: icon ${s}`));
  });
  C.CONTEXT.forEach(k => assert.ok(C.find(k), `right-click ${k}`));
  C.SHORTCUTS.forEach(sc => assert.ok(C.find(sc.key), `shortcut ${sc.key}`));
});

test('the manifest is generated from the registry and the Build tab definition is complete', () => {
  const { probe, commands: C } = loadProbe(ALL);
  const manifest = read('../manifest.xml');
  C.MAIN.forEach(g => g.controls.forEach(c => {
    assert.ok(manifest.includes(`id="${C.controlId(c.key)}"`), `manifest lacks ${c.key}; run node tools/build-ribbon.mjs`);
    (c.items || []).forEach(i => assert.ok(manifest.includes(`id="${C.controlId(i.key)}"`), `manifest lacks ${i.key}`));
  }));
  C.CONTEXT.forEach(k => assert.ok(manifest.includes(`id="HFG.ctx.${k}"`), `manifest lacks right-click ${k}`));
  const ids = [...manifest.matchAll(/ id="([^"]+)"/g)].map(m => m[1]).filter(id => id.startsWith('HFG.'));
  const def = probe.buildTabDefinition('https://localhost:3000');
  const tab = def.tabs[0];
  assert.equal(tab.id, C.BUILD_TAB.id);
  assert.ok(def.actions.length && def.version);
  const buildIds = [];
  tab.groups.forEach(g => {
    assert.ok(g.icon.length === 3 && g.controls.length, g.id);
    buildIds.push(g.id);
    g.controls.forEach(c => {
      buildIds.push(c.id);
      assert.ok(c.superTip.description.length <= 250 && c.icon.length === 3, c.id);
      if (c.type === 'Menu') c.items.forEach(i => { buildIds.push(i.id); assert.ok(i.actionId && i.icon.length === 3, i.id); });
      else assert.ok(c.actionId, c.id);
    });
  });
  const clash = buildIds.filter(id => ids.includes(id));
  assert.deepEqual(clash, [], 'Build tab ids clash with manifest ids');
  assert.equal(new Set(buildIds).size, buildIds.length, 'duplicate Build tab ids');
});

test('a click on any designed control opens its view, whichever surface it came from', async () => {
  const { probe } = loadProbe(ALL);
  const seen = [];
  probe.onView = key => seen.push(key);
  let done = 0;
  for (const id of ['HFG.mod-insert', 'HFG.ctx.cat-add', 'HFG.b-style-total', 'HFG.ch-z']) {
    await probe.openView({ source: { id }, completed: () => { done++; } });
  }
  assert.deepEqual(seen, ['mod-insert', 'cat-add', 'b-style-total', 'ch-z']);
  assert.equal(done, 4);
  assert.equal(probe.state.results.get('ribbon-full').status, 'manual');
});

test('the speed check finds volatile functions, whole-column references and links to other workbooks', () => {
  const { probe } = loadProbe();
  const s = probe.scanFormulas([
    ['=SUM(A:A)', '=INDIRECT("Sheet1!A1")+OFFSET(B2,1,0)', 'text', 3],
    ['=SUM(J9:U9)', '=TODAY()', '=[Budget.xlsx]Data!B4*2', '=IF(B5="12:30","A:B",1)'],
    ['=SUM($3:$3)', '=Rev1_Base*1.05', '=CELL("address",A1)', null]
  ]);
  assert.equal(s.formulas, 9);
  assert.equal(s.volatile, 4);
  assert.deepEqual({ ...s.fns }, { INDIRECT: 1, OFFSET: 1, TODAY: 1, CELL: 1 });
  assert.equal(s.whole, 2);         // A:A and $3:$3; text in quotes is ignored
  assert.equal(s.external, 1);
});
