// Dressing: a layout's rows become cells with cell formats, row heights and outline levels,
// column widths, frozen panes, hyperlinks, conditional formats and validations. Formats belong to
// the row definitions (kind, style, role, unit), so a row written by a change plan formats itself
// the same way as one written in a fresh build: the same code dresses a row into a worksheet for
// the package writer (SheetSink) and into a change plan's write operation for the live writer
// (RecordSink, rowFormat).

import { RC } from '../assurance.ts';
import { FIRST_PERIOD_COL, LABEL_COLS, STD, TOTAL_COL, UNIT_COL } from '../frame.ts';
import type { Layout, LRow, SheetKind, Validation } from '../layout.ts';
import { frameCells, rowCells, type Positions } from '../render.ts';
import { standardFrameLinks } from '../standard.ts';
import { catalogue, CHECK_RED, formatForUnit, styleName, type Modifier, type StyleBook } from '../styles.ts';
import { esc, ref, SheetOut, type CellOut, type ColOut, type RowOut } from './sheet.ts';

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

/** A cell's format: a catalogue style (by its key) and modifiers. */
export interface CellFormat {
  style: string;
  mods?: Modifier[];
}

export type CondRule = { kind: 'notZero' } | { kind: 'expression'; formula: string };

/** Where dressing goes: a worksheet for the package writer, or a record for a change plan. */
export interface Sink {
  /** Write a value with its format. */
  put(r: number, c: number, value: unknown, fmt: CellFormat): void;
  /** Format a cell and keep its value; with onlyIfEmpty, leave a cell that already has something. */
  format(r: number, c: number, fmt: CellFormat, onlyIfEmpty?: boolean): void;
  row(r: number, props: RowOut): void;
  link(r: number, c: number, to: string, tip: string): void;
  cond(sqref: string, rule: CondRule): void;
  valid(r: number, c: number, v: Validation): void;
}

export class SheetSink implements Sink {
  readonly sheet: SheetOut;
  private book: StyleBook;
  private redDxf: number;

  constructor(sheet: SheetOut, book: StyleBook) {
    this.sheet = sheet;
    this.book = book;
    this.redDxf = book.dxf(`<dxf><font><b/><color rgb="${CHECK_RED}"/></font></dxf>`);
  }

  put(r: number, c: number, value: unknown, fmt: CellFormat): void {
    this.sheet.set(r, c, toCell(value, this.book.xf(fmt.style, fmt.mods)));
  }

  format(r: number, c: number, fmt: CellFormat, onlyIfEmpty = false): void {
    if (onlyIfEmpty && this.sheet.get(r, c)) return;
    this.sheet.style(r, c, this.book.xf(fmt.style, fmt.mods));
  }

  row(r: number, props: RowOut): void {
    this.sheet.row(r, props);
  }

  link(r: number, c: number, to: string, tip: string): void {
    this.sheet.links.push({ row: r, col: c, to, tip });
  }

  cond(sqref: string, rule: CondRule): void {
    const xml = rule.kind === 'notZero'
      ? `<cfRule type="cellIs" dxfId="${this.redDxf}" priority="{p}" operator="notEqual"><formula>0</formula></cfRule>`
      : `<cfRule type="expression" dxfId="${this.redDxf}" priority="{p}"><formula>${esc(rule.formula)}</formula></cfRule>`;
    this.sheet.conds.push({ sqref, rules: [xml] });
  }

  valid(r: number, c: number, v: Validation): void {
    this.sheet.validations.push(validationXml(v, ref(r, c)));
  }
}

/** One row's formats as a change plan carries them for the live writer: every cell's named style. */
export interface RowFormat {
  cells: Record<number, { style: string }>;
  height?: number;
  level?: number;
  hidden?: boolean;
  collapsed?: boolean;
  links?: { col: number; to: string; tip: string }[];
  /** Conditional formats on the row: checks that turn bold red. Replace the row's existing rules. */
  conds?: { sqref: string; rule: CondRule }[];
  valid?: { col: number; rule: Validation }[];
}

const STYLE_NAMES = Object.fromEntries(Object.entries(catalogue('HF')).map(([k, s]) => [k, s.name]));

export class RecordSink implements Sink {
  readonly rows = new Map<number, RowFormat>();

  at(r: number): RowFormat {
    let f = this.rows.get(r);
    if (!f) this.rows.set(r, f = { cells: {} });
    return f;
  }

  private cell(r: number, c: number, fmt: CellFormat): void {
    const name = STYLE_NAMES[fmt.style];
    if (!name) throw new Error(`no style ${fmt.style} in the catalogue`);
    this.at(r).cells[c] = { style: styleName(name, fmt.mods) };
  }

  put(r: number, c: number, _value: unknown, fmt: CellFormat): void {
    this.cell(r, c, fmt);
  }

  format(r: number, c: number, fmt: CellFormat, onlyIfEmpty = false): void {
    if (onlyIfEmpty && c in this.at(r).cells) return;
    this.cell(r, c, fmt);
  }

  row(r: number, props: RowOut): void {
    const f = this.at(r);
    if (props.ht !== undefined) f.height = props.ht;
    if (props.level !== undefined) f.level = props.level;
    if (props.hidden) f.hidden = true;
    if (props.collapsed) f.collapsed = true;
  }

  link(r: number, c: number, to: string, tip: string): void {
    (this.at(r).links ??= []).push({ col: c, to, tip });
  }

  cond(sqref: string, rule: CondRule): void {
    const r = Number(/\d+/.exec(sqref)![0]);
    (this.at(r).conds ??= []).push({ sqref, rule });
  }

  valid(r: number, c: number, v: Validation): void {
    (this.at(r).valid ??= []).push({ col: c, rule: v });
  }
}

interface Ctx {
  layout: Layout;
  pos: Positions;
  sink: Sink;
  name: string;
  kind: SheetKind;
  lastCol: number;
  std: boolean;
}

function lastColumn(layout: Layout, kind: SheetKind, rows: LRow[]): number {
  if (kind === 'timeline' || kind === 'settings') return FIRST_PERIOD_COL + layout.periods - 1;
  let max = TOTAL_COL;
  for (const r of rows) for (const c of Object.keys(r.cells)) max = Math.max(max, Number(c));
  return max;
}

function band(ctx: Ctx, r: number, style: string, from = 2): void {
  for (let c = from; c <= ctx.lastCol; c++) ctx.sink.format(r, c, { style });
}

function dressRow(ctx: Ctx, row: LRow, r: number, prev: LRow | undefined): void {
  const { sink } = ctx;
  const cells = rowCells(ctx.layout, ctx.name, row, r, ctx.pos, 'excel');
  const put = (c: number, style: string, mods: Modifier[] = []) => {
    if (c in cells) sink.put(r, c, cells[c], { style, mods });
  };
  const labelCol = LABEL_COLS[Math.min(row.indent, 2)];
  const level = row.level;
  switch (row.kind) {
    case 'blank':
      if (ctx.std) sink.row(r, { ht: row.space ?? HEIGHTS.spacer, level: level ?? 1 });
      return;
    case 'heading':
      band(ctx, r, 'h1');
      put(labelCol, 'h1');
      sink.row(r, { ht: HEIGHTS.heading, level: level ?? 0 });
      return;
    case 'subheading':
    case 'section':
      band(ctx, r, 'h2');
      put(labelCol, 'h2');
      sink.row(r, { ht: HEIGHTS.heading, level: level ?? 1 });
      return;
    case 'setting': {
      put(labelCol, 'label');
      put(UNIT_COL, 'unit');
      const style = row.link ? calcStyle(row.unit) : row.role ?? inputStyle(row.unit);
      put(TOTAL_COL, style);
      if (row.valid && !row.link) sink.valid(r, TOTAL_COL, row.valid);
      sink.row(r, { level: level ?? 1 });
      return;
    }
    case 'fixed': {
      put(labelCol, 'label');
      put(UNIT_COL, 'unit');
      put(TOTAL_COL, row.role ?? (row.style === 'rate' ? 'pct' : calcStyle(row.unit)));
      for (const c of Object.keys(cells).map(Number)) if (c > TOTAL_COL) put(c, 'text');
      sink.row(r, { level: level ?? 1 });
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
      if (check) sink.cond(`${ref(r, TOTAL_COL)}:${ref(r, Math.max(TOTAL_COL, ...Object.keys(cells).map(Number)))}`, { kind: 'notZero' });
      sink.row(r, { ht: check ? HEIGHTS.check : undefined, level: level ?? (row.role === 'working' ? 2 : 1) });
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
  const { sink } = ctx;
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
      if (!(RC[k] in cells) && typed) sink.format(r, RC[k], { style });
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
    if (row.style === 'check' && TOTAL_COL in cells) sink.cond(ref(r, TOTAL_COL), { kind: 'notZero' });
  }
  for (const [c, l] of Object.entries(row.links || {})) sink.link(r, Number(c), l.to, l.tip);
  sink.row(r, { ht: role === 'toc1' || role === 'sectionNo' ? HEIGHTS.heading : undefined, level: row.level ?? (ctx.std ? 0 : undefined) });
}

const BLOCK_STYLE: Record<number, string> = { 7: 'date', 8: 'date', 9: 'int', 10: 'year', 11: 'int', 12: 'int', 13: 'int', 14: 'int', 15: 'int' };

/** Header: rows 1 to 3, the navigation links in A1 and A2, and the timeline block on timeline sheets. */
function dressHeader(ctx: Ctx): void {
  const { sink, layout, name, kind } = ctx;
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
    sink.put(r, c, v, { style });
  }
  if (!ctx.std) return;
  sink.row(STD.titleRow, { ht: HEIGHTS.title });
  sink.row(STD.nameRow, { ht: HEIGHTS.name });
  sink.row(STD.entityRow, { ht: HEIGHTS.entity });
  sink.row(4, { ht: HEIGHTS.spacer });
  for (const l of standardFrameLinks(layout, name)) sink.link(l.row, l.col, l.link.to, l.link.tip);
  if (layout.hasChecks() && kind !== 'contents') sink.cond('A2', { kind: 'expression', formula: 'Chk_Errors<>0' });
  if (!timeline) return;
  for (let c = 3; c <= TOTAL_COL; c++) {   // the bars run from B to I; cells already written keep their format
    sink.format(5, c, { style: 'periodLabel' }, true);
    sink.format(6, c, { style: 'period2Label' }, true);
  }
  for (let r = 7; r <= STD.blockLast; r++) sink.row(r, { level: 2, hidden: true });
  sink.row(STD.freezeRow, { ht: HEIGHTS.spacer, collapsed: true });
}

/** A sheet's own settings: columns, frozen panes, outline direction and default row height. */
export interface SheetFormat {
  cols: ColOut[];
  freeze: { row: number; col: number } | null;
  summaryBelow: boolean;
  defaultHeight: number;
}

export function sheetFormat(layout: Layout, sheet: string): SheetFormat {
  const kind = layout.kindOf(sheet);
  if (layout.frame.id !== 'standard') {
    return { defaultHeight: 15, freeze: null, summaryBelow: true, cols: [{ min: 1, max: 6, width: 2.5 }, { min: 7, max: 7, width: 34 },
      { min: 8, max: 8, width: 6 }, { min: 9, max: 9, width: 12 }, { min: FIRST_PERIOD_COL, max: FIRST_PERIOD_COL + layout.periods - 1, width: 10 }] };
  }
  const cols: ColOut[] = [];
  const col = (min: number, max: number, width: number) => cols.push({ min, max, width });
  if (kind === 'contents') {
    col(1, 1, 2.5); col(2, 4, 3.75); col(5, 5, 48); col(6, 8, 2.5); col(9, 9, 12); col(10, 10, 30);
  } else if (kind === 'cover') {
    col(1, 1, 3.75); col(2, 2, 70);
  } else {
    col(1, 1, 3.75); col(2, 6, 2.5); col(7, 7, 34); col(8, 8, 7);
    col(9, 9, kind === 'settings' ? 28 : kind === 'timeline' ? 11.75 : 14);
    if (kind === 'timeline' || kind === 'settings') {
      col(FIRST_PERIOD_COL, FIRST_PERIOD_COL + layout.periods - 1, 11.75);
    } else if (kind === 'register' || kind === 'list') {
      const widths: Record<keyof typeof RC, number> = { source: 44, owner: 16, updated: 14, evidence: 32, age: 12, group: 46, reason: 34, status: 18 };
      for (const [k, w] of Object.entries(widths)) col(RC[k as keyof typeof RC], RC[k as keyof typeof RC], w);
    }
  }
  return {
    cols, defaultHeight: HEIGHTS.body, summaryBelow: kind !== 'contents',
    freeze: kind === 'timeline' || kind === 'settings' ? { row: STD.freezeRow, col: FIRST_PERIOD_COL } : { row: 4, col: 1 },
  };
}

function context(layout: Layout, name: string, sink: Sink, pos = layout.positions()): Ctx {
  const kind = layout.kindOf(name);
  const rows = layout.sheets.find(([s]) => s === name)?.[1] ?? [];
  return { layout, pos, sink, name, kind, lastCol: lastColumn(layout, kind, rows), std: layout.frame.id === 'standard' };
}

/** Every sheet of a layout, dressed for the package writer. */
export function dress(layout: Layout, book: StyleBook): SheetOut[] {
  const pos = layout.positions();
  return layout.sheets.map(([name, rows]) => {
    const sheet = new SheetOut(name);
    const ctx = context(layout, name, new SheetSink(sheet, book), pos);
    dressHeader(ctx);
    const first = layout.firstRowOf(name);
    rows.forEach((row, k) => dressRow(ctx, row, first + k, rows[k - 1]));
    Object.assign(sheet, sheetFormat(layout, name));
    return sheet;
  });
}

/** The formats of one layout row (by its index on the sheet), as a change plan carries them. */
export function rowFormat(layout: Layout, sheet: string, k: number, pos = layout.positions()): RowFormat {
  const rec = new RecordSink();
  const rows = layout.sheets.find(([s]) => s === sheet)![1];
  const r = layout.firstRowOf(sheet) + k;
  dressRow(context(layout, sheet, rec, pos), rows[k], r, rows[k - 1]);
  const f = rec.rows.get(r) ?? { cells: {} };
  // A row inserted live takes the height of the row above it, so the plan always says the height.
  if (layout.frame.id === 'standard' && f.height === undefined) f.height = HEIGHTS.body;
  return f;
}

/** The formats of a sheet's header rows (1 to 3, the timeline block), by row. */
export function headerFormat(layout: Layout, sheet: string): Record<number, RowFormat> {
  const rec = new RecordSink();
  dressHeader(context(layout, sheet, rec));
  return Object.fromEntries(rec.rows);
}

/** One run of rows at an outline level; hidden runs are collapsed. */
export interface OutlineRun {
  from: number;
  to: number;
  level: number;
  hidden?: boolean;
}

/**
 * A sheet's whole row outline as runs. Office.js can group and ungroup rows but cannot read a
 * row's level, so the live writer clears a sheet's outline and applies these runs whenever its
 * rows change.
 */
export function sheetOutline(layout: Layout, sheet: string): OutlineRun[] {
  const rec = new RecordSink();
  const ctx = context(layout, sheet, rec);
  dressHeader(ctx);
  const rows = layout.sheets.find(([s]) => s === sheet)![1];
  const first = layout.firstRowOf(sheet);
  rows.forEach((row, k) => dressRow(ctx, row, first + k, rows[k - 1]));
  const last = first + rows.length - 1;
  const runs: OutlineRun[] = [];
  for (let r = 1; r <= last; r++) {
    const f = rec.rows.get(r);
    const level = f?.level ?? 0;
    const hidden = Boolean(f?.hidden);
    const prev = runs[runs.length - 1];
    if (prev && prev.to === r - 1 && prev.level === level && Boolean(prev.hidden) === hidden) prev.to = r;
    else runs.push(hidden ? { from: r, to: r, level, hidden } : { from: r, to: r, level });
  }
  return runs.filter(x => x.level > 0 || x.hidden);
}
