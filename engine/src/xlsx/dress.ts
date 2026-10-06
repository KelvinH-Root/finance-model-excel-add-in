// Dressing: a layout's rows become cells with cell formats, row heights and outline levels,
// column widths, frozen panes, hyperlinks, conditional formats and validations. Formats belong to
// the row definitions (kind, style, role, unit), so a row written by a change plan formats itself
// the same way as one written in a fresh build.

import { RC } from '../assurance.ts';
import { FIRST_PERIOD_COL, LABEL_COLS, STD, TOTAL_COL, UNIT_COL } from '../frame.ts';
import type { Layout, LRow, SheetKind, Validation } from '../layout.ts';
import { frameCells, rowCells, type Positions } from '../render.ts';
import { standardFrameLinks } from '../standard.ts';
import { CHECK_RED, formatForUnit, type Modifier, type StyleBook } from '../styles.ts';
import { esc, ref, SheetOut, type CellOut } from './sheet.ts';

/** Row heights in points (Look and wiring, decided 7 October 2026). */
export const HEIGHTS = { body: 11.4, heading: 12, check: 12, title: 15, name: 13.5, entity: 12, spacer: 6 } as const;

const isFormula = (v: unknown): v is string => typeof v === 'string' && v.startsWith('=');

function toCell(v: unknown, s: number): CellOut {
  if (isFormula(v)) return { f: v.slice(1), s };
  if (v === null || v === undefined) return { s };
  if (typeof v === 'number' || typeof v === 'boolean' || typeof v === 'string') return { v, s };
  throw new Error(`a cell cannot hold ${JSON.stringify(v)}`);
}

function validationXml(v: Validation, sqref: string): string {
  const head = (type: string, extra = '') => `<dataValidation type="${type}"${extra} allowBlank="1" showInputMessage="1" `
    + `showErrorMessage="1" errorTitle="Not accepted" error="${esc(v.message)}" sqref="${sqref}">`;
  const f = (n: number | string) => esc(String(n));
  switch (v.kind) {
    case 'list':
      return `${head('list')}<formula1>${esc(`"${v.items.join(',')}"`)}</formula1></dataValidation>`;
    case 'whole':
      return `${head('whole', ' operator="between"')}<formula1>${f(v.min)}</formula1><formula2>${f(v.max)}</formula2></dataValidation>`;
    case 'decimal':
      return `${head('decimal', ' operator="between"')}<formula1>-1E+15</formula1><formula2>1E+15</formula2></dataValidation>`;
    case 'date':
      return `${head('date', ' operator="between"')}<formula1>32874</formula1><formula2>109939</formula2></dataValidation>`;
    case 'text':
      return `${head('textLength', ' operator="lessThanOrEqual"')}<formula1>${v.max}</formula1></dataValidation>`;
  }
}

/** The style a value of this unit takes, as a calculation or as an input. */
const calcStyle = (unit: string) => formatForUnit(unit);
const inputStyle = (unit: string) => `in.${formatForUnit(unit)}`;

interface Ctx {
  layout: Layout;
  book: StyleBook;
  pos: Positions;
  sheet: SheetOut;
  name: string;
  kind: SheetKind;
  lastCol: number;
  std: boolean;
  redDxf: number;
}

function lastColumn(layout: Layout, kind: SheetKind, rows: LRow[]): number {
  if (kind === 'timeline' || kind === 'settings') return FIRST_PERIOD_COL + layout.periods - 1;
  let max = TOTAL_COL;
  for (const r of rows) for (const c of Object.keys(r.cells)) max = Math.max(max, Number(c));
  return max;
}

function band(ctx: Ctx, r: number, style: string, from = 2): void {
  const s = ctx.book.xf(style);
  for (let c = from; c <= ctx.lastCol; c++) ctx.sheet.style(r, c, s);
}

function checkFormat(ctx: Ctx, sqref: string): void {
  ctx.sheet.conds.push({ sqref, rules: [`<cfRule type="cellIs" dxfId="${ctx.redDxf}" priority="{p}" operator="notEqual"><formula>0</formula></cfRule>`] });
}

function dressRow(ctx: Ctx, row: LRow, r: number, prev: LRow | undefined): void {
  const { book, sheet } = ctx;
  const cells = rowCells(ctx.layout, ctx.name, row, r, ctx.pos, 'excel');
  const put = (c: number, style: string, mods: Modifier[] = []) => {
    if (c in cells) sheet.set(r, c, toCell(cells[c], book.xf(style, mods)));
  };
  const labelCol = LABEL_COLS[Math.min(row.indent, 2)];
  const level = row.level;
  switch (row.kind) {
    case 'blank':
      if (ctx.std) sheet.row(r, { ht: row.space ?? HEIGHTS.spacer, level: level ?? 1 });
      return;
    case 'heading':
      band(ctx, r, 'h1');
      put(labelCol, 'h1');
      sheet.row(r, { ht: HEIGHTS.heading, level: level ?? 0 });
      return;
    case 'subheading':
    case 'section':
      band(ctx, r, 'h2');
      put(labelCol, 'h2');
      sheet.row(r, { ht: HEIGHTS.heading, level: level ?? 1 });
      return;
    case 'setting': {
      put(labelCol, 'label');
      put(UNIT_COL, 'unit');
      const style = row.link ? calcStyle(row.unit) : row.role ?? inputStyle(row.unit);
      put(TOTAL_COL, style);
      if (row.valid && !row.link) sheet.validations.push(validationXml(row.valid, ref(r, TOTAL_COL)));
      sheet.row(r, { level: level ?? 1 });
      return;
    }
    case 'fixed': {
      put(labelCol, 'label');
      put(UNIT_COL, 'unit');
      put(TOTAL_COL, row.role ?? (row.style === 'rate' ? 'pct' : calcStyle(row.unit)));
      for (const c of Object.keys(cells).map(Number)) if (c > TOTAL_COL) put(c, 'text');
      sheet.row(r, { level: level ?? 1 });
      return;
    }
    case 'series': {
      const check = row.style === 'check';
      // A total under a list (whose last item has the dashed rule) is plain; a total that stands on
      // its own is a major result: bold, with a rule above.
      const major = row.style === 'total' && !(ctx.std && prev?.role === 'last');
      const mods: Modifier[] = [];
      if (major) mods.push('total');
      if (row.role === 'last') mods.push('last');
      put(labelCol, major ? 'h3' : 'label');
      put(UNIT_COL, 'unit');
      for (const c of Object.keys(cells).map(Number)) {
        if (c >= TOTAL_COL) put(c, check ? 'check' : calcStyle(row.unit), mods);
      }
      if (check) checkFormat(ctx, `${ref(r, TOTAL_COL)}:${ref(r, Math.max(TOTAL_COL, ...Object.keys(cells).map(Number)))}`);
      sheet.row(r, { ht: check ? HEIGHTS.check : undefined, level: level ?? (row.role === 'working' ? 2 : 1) });
      return;
    }
    case 'toc':
    case 'text': {
      dressToc(ctx, row, r, cells, put);
      return;
    }
  }
}

function dressToc(ctx: Ctx, row: LRow, r: number, cells: Record<number, unknown>,
  put: (c: number, style: string, mods?: Modifier[]) => void): void {
  const { sheet } = ctx;
  const linked = (c: number) => Boolean(row.links && c in row.links);
  const role = row.role;
  if (row.style === 'reg') {
    // An input register line: the link to the input, its value, then the record's columns.
    const bound = String(cells[RC.source] ?? '').startsWith('Group assumptions');
    const local = String(cells[RC.group] ?? '').startsWith('Local');
    put(2, 'link');
    put(UNIT_COL, 'unit');
    put(TOTAL_COL, row.unit === '%' ? 'pct' : calcStyle(row.unit));
    for (const k of ['source', 'owner', 'updated', 'evidence', 'reason'] as const) {
      const typed = !bound && (k !== 'reason' || local);
      const style = k === 'updated' ? (typed ? 'in.date' : 'date') : typed ? 'in.text' : 'text';
      put(RC[k], style);
      if (!(RC[k] in cells) && typed) sheet.set(r, RC[k], { s: ctx.book.xf(style) });
    }
    put(RC.age, 'int');
    put(RC.group, 'text');
    put(RC.status, 'text');
  } else if (row.style === 'bold') {
    for (const c of Object.keys(cells).map(Number)) put(c, c === TOTAL_COL && typeof cells[c] !== 'string' ? 'int' : 'h3');
    if (TOTAL_COL in cells && isFormula(cells[TOTAL_COL])) put(TOTAL_COL, 'int', ['bold']);
  } else {
    const style = role === 'toc1' || role === 'toc2' || role === 'toc3' ? role
      : role === 'sectionNo' ? 'sectionNo' : role === 'note' ? 'note' : role === 'h3' ? 'h3' : role === 'linkU' ? 'linkU' : 'label';
    for (const c of Object.keys(cells).map(Number)) {
      if (c === TOTAL_COL) put(c, row.style === 'check' ? 'check' : calcStyle(row.unit));
      else if (c === UNIT_COL) put(c, 'unit');
      else put(c, linked(c) && style === 'label' ? 'link' : style);
    }
    if (row.style === 'check' && TOTAL_COL in cells) checkFormat(ctx, ref(r, TOTAL_COL));
  }
  for (const [c, l] of Object.entries(row.links || {})) sheet.links.push({ row: r, col: Number(c), to: l.to, tip: l.tip });
  sheet.row(r, { ht: role === 'toc1' || role === 'sectionNo' ? HEIGHTS.heading : undefined, level: row.level ?? (ctx.std ? 0 : undefined) });
}

const BLOCK_STYLE: Record<number, string> = { 7: 'date', 8: 'date', 9: 'int', 10: 'year', 11: 'int', 12: 'int', 13: 'int', 14: 'int', 15: 'int' };

/** Header: rows 1 to 3, the navigation links in A1 and A2, and the timeline block on timeline sheets. */
function dressHeader(ctx: Ctx): void {
  const { book, sheet, layout, name, kind } = ctx;
  const timeline = kind === 'timeline' || kind === 'settings';
  for (const [r, c, v] of frameCells(layout, name, 'excel')) {
    let style = 'label';
    if (!ctx.std) style = r === 1 ? 'title' : r === 2 ? 'entity' : 'h3';
    else if (c === 1) style = 'nav';
    else if (r === STD.titleRow) style = 'title';
    else if (r === STD.nameRow) style = 'modelName';
    else if (r === STD.entityRow) style = 'entity';
    else if (r === 5) style = c === 2 ? 'periodLabel' : 'period';
    else if (r === 6) style = c === 2 ? 'period2Label' : 'period2';
    else if (c > 2) style = BLOCK_STYLE[r] ?? 'int';
    sheet.set(r, c, toCell(v, book.xf(style)));
  }
  if (!ctx.std) return;
  sheet.row(STD.titleRow, { ht: HEIGHTS.title });
  sheet.row(STD.nameRow, { ht: HEIGHTS.name });
  sheet.row(STD.entityRow, { ht: HEIGHTS.entity });
  sheet.row(4, { ht: HEIGHTS.spacer });
  for (const l of standardFrameLinks(layout, name)) sheet.links.push({ row: l.row, col: l.col, to: l.link.to, tip: l.link.tip });
  if (layout.hasChecks() && kind !== 'contents') {
    sheet.conds.push({ sqref: 'A2', rules: [`<cfRule type="expression" dxfId="${ctx.redDxf}" priority="{p}"><formula>Chk_Errors&lt;&gt;0</formula></cfRule>`] });
  }
  if (!timeline) return;
  for (let c = 3; c <= TOTAL_COL; c++) {   // the bars run from B to I; cells already written keep their format
    if (!sheet.get(5, c)) sheet.style(5, c, book.xf('periodLabel'));
    if (!sheet.get(6, c)) sheet.style(6, c, book.xf('period2Label'));
  }
  for (let r = 7; r <= STD.blockLast; r++) sheet.row(r, { level: 2, hidden: true });
  sheet.row(STD.freezeRow, { ht: HEIGHTS.spacer, collapsed: true });
}

function columns(ctx: Ctx): void {
  const { sheet, kind, layout } = ctx;
  const col = (min: number, max: number, width: number) => sheet.cols.push({ min, max, width });
  if (kind === 'contents') {
    col(1, 1, 2.5); col(2, 4, 3.75); col(5, 5, 48); col(6, 8, 2.5); col(9, 9, 12); col(10, 10, 30);
    return;
  }
  if (kind === 'cover') {
    col(1, 1, 3.75); col(2, 2, 70);
    return;
  }
  col(1, 1, 3.75); col(2, 6, 2.5); col(7, 7, 34); col(8, 8, 7);
  col(9, 9, kind === 'settings' ? 28 : kind === 'timeline' ? 11.75 : 14);
  if (kind === 'timeline' || kind === 'settings') {
    col(FIRST_PERIOD_COL, FIRST_PERIOD_COL + layout.periods - 1, 11.75);
  } else if (kind === 'register' || kind === 'list') {
    const widths: Record<keyof typeof RC, number> = { source: 44, owner: 16, updated: 14, evidence: 32, age: 12, group: 46, reason: 34, status: 18 };
    for (const [k, w] of Object.entries(widths)) col(RC[k as keyof typeof RC], RC[k as keyof typeof RC], w);
  }
}

/** Every sheet of a layout, dressed for the package writer. */
export function dress(layout: Layout, book: StyleBook): SheetOut[] {
  const pos = layout.positions();
  const std = layout.frame.id === 'standard';
  const redDxf = book.dxf(`<dxf><font><b/><color rgb="${CHECK_RED}"/></font></dxf>`);
  return layout.sheets.map(([name, rows]) => {
    const kind = layout.kindOf(name);
    const sheet = new SheetOut(name);
    const ctx: Ctx = { layout, book, pos, sheet, name, kind, lastCol: lastColumn(layout, kind, rows), std, redDxf };
    dressHeader(ctx);
    const first = layout.firstRowOf(name);
    rows.forEach((row, k) => dressRow(ctx, row, first + k, rows[k - 1]));
    if (std) {
      columns(ctx);
      sheet.summaryBelow = kind !== 'contents';
      sheet.freeze = kind === 'timeline' || kind === 'settings' ? { row: STD.freezeRow, col: FIRST_PERIOD_COL } : { row: 4, col: 1 };
    } else {
      sheet.defaultHeight = 15;
      sheet.cols.push({ min: 1, max: 6, width: 2.5 }, { min: 7, max: 7, width: 34 }, { min: 8, max: 8, width: 6 },
        { min: 9, max: 9, width: 12 }, { min: FIRST_PERIOD_COL, max: FIRST_PERIOD_COL + layout.periods - 1, width: 10 });
    }
    return sheet;
  });
}
