// The change plan carries everything the live writer needs in the standard frame. A workbook
// image (cells, formats, row heights, outline, links, checks formatting, validations, names) is
// built from one layout, the plan to the next layout is applied to it the way the live writer
// applies it, and the result must equal the image of a fresh build of the next layout. Formulas
// are left to LibreOffice (tests/test_engine_writer.py and the assembly proof): Excel moves the
// references in rows the plan does not touch, which an image cannot.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import {
  assemble, frameCells, headerFormat, planChange, rowCells, rowFormat, sheetOutline, type Layout, type Model, type PlanOp,
  type RowFormat,
} from '../src/index.ts';
import { loadLibrary } from '../src/node/library.ts';
import { demoModel, LIBRARY } from './demo.ts';

const lib = loadLibrary(LIBRARY);

interface Cell { v?: unknown; style?: string }
interface Sheet {
  cells: Map<number, Map<number, Cell>>;
  heights: Map<number, number>;
  levels: Map<number, number>;
  hidden: Set<number>;
  links: Map<string, string>;
  conds: Map<number, string[]>;
  valid: Map<number, string[]>;
}
interface Image { sheets: Map<string, Sheet>; order: string[]; names: Map<string, string> }

const empty = (): Sheet => ({ cells: new Map(), heights: new Map(), levels: new Map(), hidden: new Set(), links: new Map(), conds: new Map(), valid: new Map() });
const isFormula = (v: unknown) => typeof v === 'string' && v.startsWith('=');

function cell(sh: Sheet, r: number, c: number): Cell {
  let row = sh.cells.get(r);
  if (!row) sh.cells.set(r, row = new Map());
  let x = row.get(c);
  if (!x) row.set(c, x = {});
  return x;
}

/** A row's formats onto the image, replacing what the row had (Excel: clear the row's formats, then apply). */
function applyFormat(sh: Sheet, r: number, f: RowFormat, replace: boolean): void {
  if (replace) {
    for (const x of sh.cells.get(r)?.values() ?? []) delete x.style;
    for (const k of [...sh.links.keys()]) if (k.startsWith(`${r},`)) sh.links.delete(k);
    sh.conds.delete(r);
    sh.valid.delete(r);
  }
  for (const [c, fmt] of Object.entries(f.cells)) cell(sh, r, Number(c)).style = fmt.style;
  if (f.height !== undefined) sh.heights.set(r, f.height);
  for (const l of f.links ?? []) sh.links.set(`${r},${l.col}`, `${l.to}|${l.tip}`);
  if (f.conds?.length) sh.conds.set(r, f.conds.map(c => `${c.sqref.replace(/\d+/g, '#')}|${JSON.stringify(c.rule)}`));
  if (f.valid?.length) sh.valid.set(r, f.valid.map(v => `${v.col}|${JSON.stringify(v.rule)}`));
}

function applyOutline(sh: Sheet, runs: { from: number; to: number; level: number; hidden?: boolean }[]): void {
  sh.levels.clear();
  sh.hidden.clear();
  for (const run of runs) {
    for (let r = run.from; r <= run.to; r++) {
      if (run.level) sh.levels.set(r, run.level);
      if (run.hidden) sh.hidden.add(r);
    }
  }
}

function writeCells(sh: Sheet, r: number, cells: Record<number, unknown>): void {
  for (const [c, v] of Object.entries(cells)) cell(sh, r, Number(c)).v = v;
}

/** The image of a fresh build. */
function image(layout: Layout): Image {
  const pos = layout.positions();
  const img: Image = { sheets: new Map(), order: [], names: new Map() };
  for (const [s, rows] of layout.sheets) {
    const sh = empty();
    img.sheets.set(s, sh);
    img.order.push(s);
    for (const [r, c, v] of frameCells(layout, s, 'excel')) cell(sh, r, c).v = v;
    for (const [r, f] of Object.entries(headerFormat(layout, s))) applyFormat(sh, Number(r), f, false);
    const first = layout.firstRowOf(s);
    rows.forEach((row, k) => {
      writeCells(sh, first + k, rowCells(layout, s, row, first + k, pos, 'excel'));
      applyFormat(sh, first + k, rowFormat(layout, s, k, pos), false);
    });
    applyOutline(sh, sheetOutline(layout, s));
  }
  for (const [nm, rid] of layout.names) {
    const [s, r] = pos.get(rid)!;
    img.names.set(nm, `${s}!${r},${layout.nameCol(nm)}`);
  }
  return img;
}

/** Move every row-keyed thing on a sheet from row `at` by `by` (rows removed first when by < 0). */
function shift(img: Image, s: string, at: number, by: number): void {
  const sh = img.sheets.get(s)!;
  const move = <T>(m: Map<number, T>) => {
    const out = new Map<number, T>();
    for (const [r, v] of m) {
      if (by < 0 && r >= at && r < at - by) continue;
      out.set(r >= at ? r + by : r, v);
    }
    m.clear();
    for (const [r, v] of out) m.set(r, v);
  };
  move(sh.cells); move(sh.heights); move(sh.levels); move(sh.conds); move(sh.valid);
  const hidden = new Map([...sh.hidden].map(r => [r, true]));
  move(hidden);
  sh.hidden = new Set(hidden.keys());
  const links = new Map<number, [string, string][]>();
  for (const [k, v] of sh.links) {
    const r = Number(k.split(',')[0]);
    links.set(r, [...(links.get(r) ?? []), [k.split(',')[1], v]]);
  }
  move(links);
  sh.links = new Map([...links].flatMap(([r, list]) => list.map(([c, v]) => [`${r},${c}`, v] as [string, string])));
  for (const [nm, at2] of img.names) {
    const [sheet, rc] = at2.split('!');
    if (sheet !== s) continue;
    const [r, c] = rc.split(',').map(Number);
    if (by < 0 && r >= at && r < at - by) img.names.delete(nm);
    else if (r >= at) img.names.set(nm, `${sheet}!${r + by},${c}`);
  }
}

function apply(img: Image, ops: PlanOp[]): void {
  for (const op of ops) {
    switch (op.op) {
      case 'delete_name': img.names.delete(op.name); break;
      case 'delete_sheet': img.sheets.delete(op.sheet); img.order = img.order.filter(s => s !== op.sheet); break;
      case 'add_sheet': {
        const sh = empty();
        img.sheets.set(op.sheet, sh);
        img.order.splice(op.index, 0, op.sheet);
        for (const [r, c, v] of op.frame) cell(sh, r, c).v = v;
        for (const [r, f] of Object.entries(op.format!.header)) applyFormat(sh, Number(r), f, false);
        break;
      }
      case 'insert_rows': shift(img, op.sheet, op.row, op.count); break;
      case 'delete_rows': shift(img, op.sheet, op.row, -op.count); break;
      case 'write': {
        const sh = img.sheets.get(op.sheet)!;
        writeCells(sh, op.row, op.cells);
        applyFormat(sh, op.row, op.format!, true);
        break;
      }
      case 'outline': applyOutline(img.sheets.get(op.sheet)!, op.runs); break;
      case 'add_name': img.names.set(op.name, `${op.sheet}!${op.row},${op.col}`); break;
      default: break;   // charts are compared by the chart tests
    }
  }
}

/** Everything but formulas, in a form assert can compare. */
function flat(img: Image) {
  const sheets: Record<string, unknown> = {};
  for (const s of img.order) {
    const sh = img.sheets.get(s)!;
    const cells: Record<string, unknown> = {};
    for (const [r, row] of sh.cells) {
      for (const [c, x] of row) {
        if (x.v === undefined && !x.style) continue;
        cells[`${r},${c}`] = { v: isFormula(x.v) ? 'formula' : x.v ?? null, style: x.style ?? null };
      }
    }
    sheets[s] = {
      cells, heights: Object.fromEntries(sh.heights), levels: Object.fromEntries(sh.levels), hidden: [...sh.hidden].sort((a, b) => a - b),
      links: Object.fromEntries(sh.links), conds: Object.fromEntries(sh.conds), valid: Object.fromEntries(sh.valid),
    };
  }
  return { order: img.order, sheets, names: Object.fromEntries([...img.names].sort()) };
}

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
