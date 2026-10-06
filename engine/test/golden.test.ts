// The engine against the Python proof it replaces: every fixture in test/fixtures was written by
// tools/golden.py from prototypes/assembly. Layout, rendered cells and the change plan must match
// exactly, in both dialects.

import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import {
  assemble, chartRefs, FIRST_ROW, frameCells, Library, Model, planChange, rowCells,
  type Dialect, type Layout, type ModelDict,
} from '../src/index.ts';

const HERE = dirname(fileURLToPath(import.meta.url));
const FIXTURES = join(HERE, 'fixtures');
const LIBRARY = join(HERE, '..', '..', 'prototypes', 'assembly', 'library');

interface Fixture {
  name: string;
  old: ModelDict;
  new: ModelDict;
  layout: unknown;
  cells: Record<Dialect, unknown>;
  plan: Record<Dialect, unknown>;
}

/** JSON in and out, so numeric keys become strings and class instances plain objects, as in the fixture. */
const plain = (x: unknown): unknown => JSON.parse(JSON.stringify(x));

function layoutDict(lay: Layout) {
  return plain({
    periods: lay.periods,
    sheets: lay.sheets,
    names: [...lay.names],
    records: lay.records,
    warnings: lay.warnings,
    blocks: [...lay.blocks],
    charts: lay.charts,
    kinds: lay.kinds,
    titles: lay.titles,
    name_cols: lay.nameCols,
    headlines: lay.headlines,
  });
}

function cellsDict(lay: Layout, dialect: Dialect) {
  const pos = lay.positions();
  const out: Record<string, unknown> = {};
  for (const [s, rows] of lay.sheets) {
    const sheet: Record<string, Record<string, unknown>> = {};
    const put = (r: number, c: number, v: unknown) => ((sheet[r] ??= {})[c] = v);
    for (const [r, c, v] of frameCells(lay, s, dialect)) put(r, c, v);
    rows.forEach((row, k) => {
      const rn = FIRST_ROW + k;
      for (const [c, v] of Object.entries(rowCells(lay, s, row, rn, pos, dialect))) put(rn, Number(c), v);
    });
    out[s] = sheet;
  }
  out['@charts'] = lay.charts.map(c => chartRefs(lay, c, dialect));
  return plain(out);
}

const lib = Library.load(LIBRARY);
const files = readdirSync(FIXTURES).filter(f => f.endsWith('.json')).sort();

test('fixtures are present', () => {
  assert.ok(files.length >= 11, `expected the golden fixtures, found ${files.length}; run npm run golden`);
});

for (const f of files) {
  const fx = JSON.parse(readFileSync(join(FIXTURES, f), 'utf8')) as Fixture;
  test(`${fx.name}: model round trip`, () => {
    assert.deepEqual(plain(Model.fromDict(lib, fx.new).toDict()), fx.new);
  });
  const lo = assemble(Model.fromDict(lib, fx.old));
  const ln = assemble(Model.fromDict(lib, fx.new));
  test(`${fx.name}: layout`, () => {
    assert.deepEqual(layoutDict(ln), fx.layout);
  });
  for (const dialect of ['excel', 'uno'] as const) {
    test(`${fx.name}: cells (${dialect})`, () => {
      assert.deepEqual(cellsDict(ln, dialect), fx.cells[dialect]);
    });
    test(`${fx.name}: change plan (${dialect})`, () => {
      assert.deepEqual(plain(planChange(lo, ln, dialect)), fx.plan[dialect]);
    });
  }
}
