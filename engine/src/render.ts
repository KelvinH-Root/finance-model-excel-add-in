// Rendering markers into cell formulas, in Excel's dialect (A1 with !) or LibreOffice's ($Sheet.A1, semicolons).

import {
  AssemblyError, colLetter, FIRST_PERIOD_COL, LABEL_COLS, PERIOD_ROW, TOTAL_COL, UNIT_COL,
} from './frame.ts';
import type { ChartSpec, Layout, LRow, RChart } from './layout.ts';
import type { RChartRefs } from './xlsx/charts.ts';
import { standardFrameCells } from './standard.ts';

export type Dialect = 'excel' | 'uno';
export type Positions = Map<string, [string, number]>;

const MARK = /«([RPQABSVO]|P\d+|C\d+)\|([^»|]+)»/g;
/** «G|first|last»: a block of rows across the timeline; «K|first|last|col»: one column down a block of rows (both absolute). */
const BLOCK = /«([GK])\|([^»|]+)\|([^»|]+)(?:\|(\d+))?»/g;
/** «F14»: this period's cell in timeline block row 14 (row absolute). */
const FRAME_MARK = /«F(\d+)»/g;
const PLAIN_SHEET = /^[A-Za-z_][A-Za-z0-9_]*$/;

export function sheetPrefix(sheet: string, dialect: Dialect): string {
  const plain = PLAIN_SHEET.test(sheet);
  if (dialect === 'uno') return plain ? `$${sheet}.` : `$'${sheet}'.`;
  return plain ? `${sheet}!` : `'${sheet}'!`;
}

/**
 * A marker formula as the formula for one cell. «R|id» is this period of a row, «P|id» the
 * period before (0 before the first) and «Pn|id» n periods before, «A|id» the whole timeline,
 * «B|id» the label in column B, «Cn|id» column n, «V|id» column I with the row fixed (a scalar),
 * «S|sheet» a sheet's title cell, «Fn» this period of timeline block row n, «N» the period number
 * and «T» the count.
 */
export function renderFormula(f: string, sheet: string, col: number, pos: Positions, dialect: Dialect = 'excel',
  periods = 12): string {
  let out = f.replace(BLOCK, (_m, kind: string, a: string, b: string, c?: string) => {
    const pa = pos.get(a);
    const pb = pos.get(b);
    if (!pa || !pb) throw new AssemblyError(`no row ${pa ? b : a} in the layout`);
    if (pa[0] !== pb[0]) throw new AssemblyError(`rows ${a} and ${b} are on different sheets`);
    const [c0, c1] = kind === 'G' ? [FIRST_PERIOD_COL, FIRST_PERIOD_COL + periods - 1] : [Number(c), Number(c)];
    const a1 = `$${colLetter(c0)}$${pa[1]}:$${colLetter(c1)}$${pb[1]}`;
    return pa[0] === sheet ? a1 : sheetPrefix(pa[0], dialect) + a1;
  });
  out = out.replace(MARK, (_m, kind: string, rid: string) => {
    if (kind === 'S') return (rid === sheet ? '' : sheetPrefix(rid, dialect)) + '$B$1';
    const at = pos.get(rid);
    if (!at) throw new AssemblyError(`no row ${rid} in the layout`);
    const [s, r] = at;
    let a1: string;
    if (kind === 'B') {
      a1 = `$B$${r}`;
    } else if (kind.startsWith('C')) {
      a1 = `${colLetter(Number(kind.slice(1)))}${r}`;
    } else if (kind === 'A') {
      a1 = `$${colLetter(FIRST_PERIOD_COL)}$${r}:$${colLetter(FIRST_PERIOD_COL + periods - 1)}$${r}`;
    } else if (kind === 'V') {
      a1 = `$${colLetter(TOTAL_COL)}$${r}`;
    } else if ((kind === 'Q' || kind === 'O') && col - 1 < FIRST_PERIOD_COL) {
      // the month before the first: the opening balance on the historical balance sheet
      // («O|id» is 0 for a row that has none)
      const h = pos.get(`hist/${rid}`);
      if (!h && kind === 'O') return '0';
      if (!h) throw new AssemblyError(`no historical line for ${rid}`);
      return (h[0] === sheet ? '' : sheetPrefix(h[0], dialect)) + `$${colLetter(TOTAL_COL)}$${h[1]}`;
    } else {
      const back = kind === 'R' ? 0 : kind === 'P' || kind === 'Q' || kind === 'O' ? 1 : Number(kind.slice(1));
      const c = col - back;
      if (c < FIRST_PERIOD_COL) return '0';   // a period before the first one
      a1 = `${colLetter(c)}${r}`;
    }
    return s === sheet ? a1 : sheetPrefix(s, dialect) + a1;
  });
  out = out.replace(FRAME_MARK, (_m, row: string) => `${colLetter(Math.max(col, FIRST_PERIOD_COL))}$${row}`);
  out = out.replaceAll('«N»', String(col - FIRST_PERIOD_COL + 1)).replaceAll('«T»', String(periods));
  return dialect === 'uno' ? unoSeparators(out) : out;
}

/** Excel's argument commas as LibreOffice semicolons, leaving commas inside quoted text alone. */
export function unoSeparators(f: string): string {
  return f.split('"').map((p, k) => (k % 2 === 0 ? p.replaceAll(',', ';') : p)).join('"');
}

export interface ChartRefs {
  sheet: string;
  title: string;
  anchor: { row: number; col: number };
  categories: string;
  series: { label: string; values: string; kind: 'column' | 'line' }[];
}

/** A chart's categories and series as cell ranges in one dialect. */
export function chartRefs(layout: Layout, chart: ChartSpec, dialect: Dialect = 'excel'): ChartRefs {
  const pos = layout.positions();
  const rows = new Map<string, LRow>();
  for (const [, rs] of layout.sheets) for (const r of rs) rows.set(r.id, r);
  const at = (rid: string) => {
    const p = pos.get(rid);
    if (!p) throw new AssemblyError(`chart ${chart.title}: no row ${rid} in the layout`);
    return p;
  };
  const area = (rid: string, c0: number, c1: number) => {
    const [s, r] = at(rid);
    return sheetPrefix(s, dialect) + `$${colLetter(c0)}$${r}` + (c1 !== c0 ? `:$${colLetter(c1)}$${r}` : '');
  };
  const last = FIRST_PERIOD_COL + chart.span - 1;
  const [, r] = at(chart.anchor);
  return {
    sheet: chart.sheet, title: chart.title, anchor: { row: r, col: FIRST_PERIOD_COL + layout.periods + 1 },
    categories: area(chart.categories, FIRST_PERIOD_COL, last),
    series: chart.series.map(([rid, kind]) => {
      const lab = LABEL_COLS[Math.min(rows.get(rid)!.indent, 2)];
      return { label: area(rid, lab, lab), values: area(rid, FIRST_PERIOD_COL, last), kind };
    }),
  };
}

const isFormula = (v: unknown): v is string => typeof v === 'string' && v.startsWith('=');

/** Every cell the engine owns in one row, rendered for one dialect. */
export function rowCells(layout: Layout, sheet: string, row: LRow, rownum: number, pos: Positions,
  dialect: Dialect = 'excel'): Record<number, unknown> {
  const cells: Record<number, unknown> = {};
  if (row.kind === 'blank') return cells;
  if (row.kind !== 'toc' && row.kind !== 'item') cells[LABEL_COLS[Math.min(row.indent, 2)]] = row.label;
  if (row.unit && row.unit !== 'text') cells[UNIT_COL] = row.unit;
  if (row.kind === 'setting') {
    cells[TOTAL_COL] = row.link ? renderFormula(row.link, sheet, TOTAL_COL, pos, dialect, layout.periods) : row.value;
  }
  if (row.kind === 'scenario') {
    (row.values ?? []).forEach((v, k) => { cells[FIRST_PERIOD_COL + k] = v; });
  }
  if (row.kind === 'series' && row.role === 'opening') cells[TOTAL_COL] = row.value ?? null;
  if (row.kind === 'series') {
    const last = FIRST_PERIOD_COL + (row.span || layout.periods) - 1;
    for (let c = FIRST_PERIOD_COL; c <= last; c++) {
      if (row.input) {   // a time series input: its starting values, typed
        cells[c] = row.values?.[c - FIRST_PERIOD_COL] ?? null;
        continue;
      }
      const tpl = c === FIRST_PERIOD_COL && row.first ? row.first : row.formula;
      if (tpl === null) throw new AssemblyError(`${sheet}: ${row.label} has no formula`);
      cells[c] = renderFormula(tpl, sheet, c, pos, dialect, layout.periods);
    }
    if (row.total === 'sum') {
      cells[TOTAL_COL] = `=SUM(${colLetter(FIRST_PERIOD_COL)}${rownum}:${colLetter(last)}${rownum})`;
    } else if (row.total === 'last') {
      cells[TOTAL_COL] = `=${colLetter(last)}${rownum}`;
    }
  }
  for (const [c, v] of Object.entries(row.cells)) {
    const col = Number(c);
    cells[col] = isFormula(v) ? renderFormula(v, sheet, col, pos, dialect, layout.periods) : v;
  }
  return cells;
}

/** [row, column, value] for one header cell. */
export type FrameCell = [number, number, unknown];

/**
 * Header cells every sheet carries, written by both writers: the links to the contents (A1)
 * and the checks (A2), the title, and the period row on timeline sheets. In the order written.
 */
export function frameCells(layout: Layout, sheet: string, dialect: Dialect = 'excel'): FrameCell[] {
  if (layout.frame.id === 'standard') return standardFrameCells(layout, sheet, dialect);
  const kind = layout.kindOf(sheet);
  const cells: FrameCell[] = [
    [1, 2, Object.hasOwn(layout.titles, sheet) ? layout.titles[sheet] : sheet],
    [2, 2, 'Assembly proof (demo data)'],
  ];
  if (kind !== 'contents') {
    cells.push([1, 1, '=HYPERLINK("#HL_Home","<")']);
    if (layout.hasChecks()) cells.push([2, 1, '=HYPERLINK("#HL_Err_Chk",IF(Chk_Errors=0,"✓","!"))']);
  }
  if (kind === 'timeline') {
    cells.push([PERIOD_ROW, 2, 'Month'], [PERIOD_ROW, TOTAL_COL, 'Total']);
    for (let p = 0; p < layout.periods; p++) cells.push([PERIOD_ROW, FIRST_PERIOD_COL + p, p + 1]);
  }
  return dialect === 'uno' ? cells.map(([r, c, v]) => [r, c, isFormula(v) ? unoSeparators(v) : v]) : cells;
}

/** A report chart's title, categories and series as cell references in one dialect. */
export function rchartRefs(layout: Layout, c: RChart, dialect: Dialect = 'excel'): RChartRefs {
  const pos = layout.positions();
  const rows = new Map<string, LRow>();
  for (const [, rs] of layout.sheets) for (const r of rs) rows.set(r.id, r);
  const at = (rid: string) => {
    const p = pos.get(rid);
    if (!p) throw new AssemblyError(`chart ${c.id}: no row ${rid} in the layout`);
    return p;
  };
  const area = (rid: string, c0: number, c1: number) => {
    const [s, r] = at(rid);
    return sheetPrefix(s, dialect) + `$${colLetter(c0)}$${r}` + (c1 !== c0 ? `:$${colLetter(c1)}$${r}` : '');
  };
  const last = FIRST_PERIOD_COL + c.n - 1;
  return {
    title: { ref: area(c.titleRow, TOTAL_COL, TOTAL_COL), cache: c.title },
    cats: area(c.cats, FIRST_PERIOD_COL, last),
    series: c.series.map(s => {
      const lab = LABEL_COLS[Math.min(rows.get(s.row)!.indent, 2)];
      const { row, ...look } = s;
      return { ...look, tx: area(row, lab, lab), values: area(row, FIRST_PERIOD_COL, last) };
    }),
    type: c.type, dir: c.dir ?? 'col', grouping: c.grouping ?? 'clustered', gap: c.gap ?? 60, overlap: c.overlap ?? 0,
    valueAxis: c.valueAxis ?? true, reverse: c.reverse ?? false, yFmt: c.yFmt ?? '#,##0', legend: c.legend === undefined ? 'b' : c.legend,
  };
}
