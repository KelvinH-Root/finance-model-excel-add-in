import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
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
  vm.runInContext(read('../src/probe.js'), context);
  return { probe: context.HfgProbe, associated };
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
