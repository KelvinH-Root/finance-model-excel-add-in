// The Excel self-test's own pieces outside Excel: the test pane goes into a workbook and comes out
// again, a fresh build reads back cell for cell, and the comparison sees only real differences.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { strFromU8, unzipSync } from 'fflate';
import { assemble, buildWorkbook } from '../../engine/src/index.ts';
import { loadLibrary } from '../../engine/src/node/library.ts';
import { demoModel, LIBRARY } from '../../engine/test/demo.ts';
import { compare, expectedOf, normFormula, withoutTaskPane, withTaskPane } from '../src/selftest/compare.ts';
import type { Snapshot } from '../src/selftest/run.ts';

const lib = loadLibrary(LIBRARY);
const model = demoModel(lib, 'HF');
const bytes = buildWorkbook(assemble(model), model);
const ID = 'ede549b5-a63a-4693-8d17-665a7b251b8b';

/** A snapshot as Excel would send it for a workbook that holds exactly what was built. */
function snapOf(want: ReturnType<typeof expectedOf>): Snapshot {
  const sheets: Snapshot['sheets'] = {};
  for (const [name, cells] of want.sheets) {
    let maxR = 1;
    let maxC = 1;
    const at = [...cells.keys()].map(a => {
      const m = /^([A-Z]+)(\d+)$/.exec(a)!;
      const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
      maxR = Math.max(maxR, Number(m[2]));
      maxC = Math.max(maxC, c);
      return [Number(m[2]), c, cells.get(a)!] as const;
    });
    const grid: unknown[][] = Array.from({ length: maxR }, () => Array(maxC).fill(''));
    for (const [r, c, cell] of at) grid[r - 1][c - 1] = 'f' in cell ? `=${cell.f.replace(/_xlfn\./g, '')}` : cell.v;
    sheets[name] = { address: `'${name}'!A1:Z${maxR}`, formulas: grid };
  }
  const names = Object.fromEntries([...want.names].map(([k, v]) => [k, `=${v}`]));
  return { label: 'test', order: [...want.order], sheets, names, metadata: '', errorCells: [], checks: { errors: 0, alerts: 0 } };
}

test('the test pane goes into a workbook and comes out again', () => {
  const withPane = unzipSync(withTaskPane(bytes, ID));
  assert.match(strFromU8(withPane['xl/webextensions/webextension1.xml']), new RegExp(`<we:reference id="${ID}"`));
  assert.match(strFromU8(withPane['_rels/.rels']), /webextensiontaskpanes/);
  assert.match(strFromU8(withPane['[Content_Types].xml']), /webextensiontaskpanes\+xml/);
  const back = unzipSync(withoutTaskPane(withTaskPane(bytes, ID)));
  assert.deepEqual(Object.keys(back).sort(), Object.keys(unzipSync(bytes)).sort());
  assert.doesNotMatch(strFromU8(back['_rels/.rels']), /webextension/);
  assert.doesNotMatch(strFromU8(back['[Content_Types].xml']), /webextension/);
});

test('a fresh build reads back with its sheets, cells and names', () => {
  const want = expectedOf(bytes);
  assert.deepEqual(want.order, assemble(model).sheets.map(([s]) => s));
  const cells = [...want.sheets.values()].reduce((n, s) => n + s.size, 0);
  assert.ok(cells > 200, `${cells} cells`);
  assert.ok([...want.sheets.values()].some(s => [...s.values()].some(c => 'f' in c)), 'formulas are read');
  assert.ok(want.names.size > 10);
});

test('a workbook that holds what was built compares clean; a changed cell, a lost name and a lost sheet are caught', () => {
  const want = expectedOf(bytes);
  const snap = snapOf(want);
  const clean = compare(snap, want);
  assert.equal(clean.mismatches, 0, JSON.stringify(clean.examples.slice(0, 3)));
  assert.ok(clean.orderMatches);
  assert.deepEqual(clean.names, { missing: [], extra: [], different: [] });

  const sheet = want.order.find(s => [...want.sheets.get(s)!.values()].some(c => 'f' in c))!;
  const grid = snap.sheets[sheet].formulas;
  const r = grid.findIndex(row => row.some(x => typeof x === 'string' && x.startsWith('=')));
  const c = grid[r].findIndex(x => typeof x === 'string' && x.startsWith('='));
  grid[r][c] = 42;
  const name = [...want.names.keys()][0];
  delete snap.names[name];
  const gone = snap.order.pop()!;
  const bad = compare(snap, want);
  assert.equal(bad.mismatches, 1);
  assert.equal(bad.examples[0].sheet, sheet);
  assert.equal(bad.examples[0].got, '42');
  assert.deepEqual(bad.names.missing, [name]);
  assert.deepEqual(bad.missingSheets, [gone]);
  assert.equal(bad.orderMatches, false);
});

test('a drop-down written live (text in the cell, a MATCH in its name) is counted apart, not as a mismatch', () => {
  const want = expectedOf(bytes);
  const snap = snapOf(want);
  const [name, ref] = [...want.names].find(([, f]) => /!\$[A-Z]+\$\d+$/.test(f))!;
  const at = /^(?:'((?:[^']|'')+)'|([^!]+))!\$([A-Z]+)\$(\d+)$/.exec(ref)!;
  const sheet = (at[1] ?? at[2]).replace(/''/g, "'");
  const col = [...at[3]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const grid = snap.sheets[sheet].formulas;
  while (grid.length < Number(at[4])) grid.push([]);
  const want1 = want.sheets.get(sheet)!.get(`${at[3]}${at[4]}`);
  want.sheets.get(sheet)!.set(`${at[3]}${at[4]}`, { v: 2 });
  grid[Number(at[4]) - 1][col - 1] = 'Second item';
  snap.names[name] = `=MATCH(${ref},List_Test,0)`;
  const c = compare(snap, want);
  assert.equal(c.mismatches, 0, JSON.stringify(c.examples));
  assert.equal(c.liveDropdowns, 1);
  assert.deepEqual(c.names.different, []);
  if (want1) want.sheets.get(sheet)!.set(`${at[3]}${at[4]}`, want1);
});

test('formulas compare without Excel\'s prefixes and implicit-intersection marks', () => {
  assert.equal(normFormula('=_xlfn.BETA.DIST(A1,2,3,TRUE)'), 'BETA.DIST(A1,2,3,TRUE)');
  assert.equal(normFormula('=@INDEX(List_Months,1)+@Tl_Start'), 'INDEX(List_Months,1)+Tl_Start');
  assert.equal(normFormula('="a@b"&A1'), '"a@b"&A1');
});
