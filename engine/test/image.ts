// A workbook image for the live tests: cells (values and named styles), row heights, outline,
// hyperlinks, check formatting, validations and names, built from a layout as a fresh build
// writes them. engine/test/live.test.ts applies plans to it; addin/test/live.test.ts compares the
// Office.js live writer's result with it.

import {
  frameCells, headerFormat, rowCells, rowFormat, sheetOutline, type Layout, type PlanOp, type RowFormat,
} from '../src/index.ts';

export interface Cell { v?: unknown; style?: string }
export interface Sheet {
  cells: Map<number, Map<number, Cell>>;
  heights: Map<number, number>;
  levels: Map<number, number>;
  hidden: Set<number>;
  links: Map<string, string>;
  conds: Map<number, string[]>;
  valid: Map<number, string[]>;
}
export interface Image { sheets: Map<string, Sheet>; order: string[]; names: Map<string, string> }

export const empty = (): Sheet => ({ cells: new Map(), heights: new Map(), levels: new Map(), hidden: new Set(), links: new Map(), conds: new Map(), valid: new Map() });
const isFormula = (v: unknown) => typeof v === 'string' && v.startsWith('=');

export function cell(sh: Sheet, r: number, c: number): Cell {
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
export function image(layout: Layout): Image {
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
export function shift(img: Image, s: string, at: number, by: number): void {
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

export function apply(img: Image, ops: PlanOp[]): void {
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
export function flat(img: Image) {
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

