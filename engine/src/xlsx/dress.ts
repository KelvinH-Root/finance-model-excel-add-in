// Dressing: a layout's rows become cells with cell formats, row heights and outline levels,
// column widths, frozen panes, hyperlinks, conditional formats and validations. Formats belong to
// the row definitions (kind, style, role, unit), so a row written by a change plan formats itself
// the same way as one written in a fresh build: the same code dresses a row into a worksheet for
// the package writer (SheetSink) and into a change plan's write operation for the live writer
// (RecordSink, rowFormat).

import { RC } from '../assurance.ts';
import { REG, REG_STYLES, REG_WIDTHS } from '../versions.ts';
import { colLetter, FIRST_PERIOD_COL, LABEL_COLS, STD, TOTAL_COL, UNIT_COL } from '../frame.ts';
import type { Control, Layout, LRow, SheetKind, Validation } from '../layout.ts';
import { frameCells, renderFormula, rowCells, type Positions } from '../render.ts';
import { standardFrameLinks } from '../standard.ts';
import { catalogue, CHECK_RED, formatForUnit, styleName, type Modifier, type StyleBook } from '../styles.ts';
import { SLOT, TINT } from '../theme.ts';
import { esc, ref, SheetOut, type CellOut, type ColOut, type RowOut } from './sheet.ts';

/** Report sheets: value columns from J and their width. */
export const REPORT_COLS = 24;
export const REPORT_COL_WIDTH = 10;

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
    case 'logical':
      return `${head('custom')}<formula1>${esc(`ISLOGICAL(${sqref.split(':')[0]})`)}</formula1></dataValidation>`;
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

export type CondRule = { kind: 'notZero' } | { kind: 'expression'; formula: string }
  /** The Scenarios sheet: the column of the scenario in use, shaded a step darker than an input. */
  | { kind: 'selected'; formula: string }
  /** The Scenarios band: the active scenario's name upright and bold (the others are italic). */
  | { kind: 'upright'; formula: string }
  /** The Scenarios band: the marker under the active scenario, dark (the others are faint). */
  | { kind: 'marker'; formula: string }
  /** A count that is not zero, in bold: information, not an error. */
  | { kind: 'on' }
  /** Dashboard tables: the last actual month's column, on the theme's light 2 (a subtle alternate fill). */
  | { kind: 'current'; formula: string }
  /** Dashboard tables: an unfavourable variance (below zero), in red. */
  | { kind: 'adverse' }
  /** Report table cells blank outside the timeline (#N/A, a gap in the chart) in a light grey. */
  | { kind: 'na'; formula: string }
  /** Input cells greyed out while the formula is true (an input not used in that month or for that method). */
  | { kind: 'inactive'; formula: string };

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
  /** A classic control over a cell, linked to the defined name on it. */
  control(r: number, c: number, control: Control, link: string): void;
  /** Merge cells of one row (a description that runs across the scenario columns). */
  merge(r: number, from: number, to: number): void;
}

export class SheetSink implements Sink {
  readonly sheet: SheetOut;
  private book: StyleBook;
  private redDxf: number;

  private greyDxf: number;
  private naDxf: number;
  private dxfs: Record<'selected' | 'upright' | 'marker' | 'on' | 'current' | 'adverse', number>;

  constructor(sheet: SheetOut, book: StyleBook) {
    this.sheet = sheet;
    this.book = book;
    this.redDxf = book.dxf(`<dxf><font><b/><color rgb="${CHECK_RED}"/></font></dxf>`);
    this.naDxf = book.dxf('<dxf><font><color theme="0" tint="-0.249977111117893"/></font></dxf>');
    // Inactive inputs: background-grey text, no fill, no border (white rules over the input's border).
    const white = '<color theme="0"/>';
    this.greyDxf = book.dxf(`<dxf><font><color theme="0" tint="-0.249977111117893"/></font><fill><patternFill><bgColor theme="0"/></patternFill></fill>`
      + `<border><left style="thin">${white}</left><right style="thin">${white}</right><top style="thin">${white}</top><bottom style="thin">${white}</bottom></border></dxf>`);
    this.dxfs = {
      selected: book.dxf(`<dxf><fill><patternFill><bgColor theme="${SLOT.accent1}" tint="${TINT.lighter60}"/></patternFill></fill></dxf>`),
      upright: book.dxf('<dxf><font><b/><i val="0"/></font></dxf>'),
      marker: book.dxf(`<dxf><font><color theme="${SLOT.dk2}"/></font></dxf>`),
      on: book.dxf(`<dxf><font><b/><color theme="${SLOT.dk2}"/></font></dxf>`),
      current: book.dxf(`<dxf><fill><patternFill><bgColor theme="${SLOT.lt2}"/></patternFill></fill></dxf>`),
      adverse: book.dxf(`<dxf><font><color rgb="${CHECK_RED}"/></font></dxf>`),
    };
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
    let xml: string;
    if (rule.kind === 'notZero' || rule.kind === 'on') {
      xml = `<cfRule type="cellIs" dxfId="${rule.kind === 'on' ? this.dxfs.on : this.redDxf}" priority="{p}" operator="notEqual"><formula>0</formula></cfRule>`;
    } else if (rule.kind === 'adverse') {
      xml = `<cfRule type="cellIs" dxfId="${this.dxfs.adverse}" priority="{p}" operator="lessThan"><formula>0</formula></cfRule>`;
    } else {
      const dxf = rule.kind === 'inactive' ? this.greyDxf : rule.kind === 'na' ? this.naDxf
        : rule.kind === 'selected' || rule.kind === 'upright' || rule.kind === 'marker' || rule.kind === 'current' ? this.dxfs[rule.kind] : this.redDxf;
      xml = `<cfRule type="expression" dxfId="${dxf}" priority="{p}"><formula>${esc(rule.formula)}</formula></cfRule>`;
    }
    this.sheet.conds.push({ sqref, rules: [xml] });
  }

  merge(r: number, from: number, to: number): void {
    this.sheet.merges.push(`${ref(r, from)}:${ref(r, to)}`);
  }

  valid(r: number, c: number, v: Validation): void {
    this.sheet.validations.push(validationXml(v, ref(r, c)));
  }

  control(r: number, c: number, control: Control, link: string): void {
    this.sheet.controls.push({ row: r, col: c, control, link, value: this.sheet.get(r, c)?.v ?? null });
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
  /** A classic control the package writer draws over the cell; the live writer uses an in-cell control instead. */
  control?: { col: number; control: Control; link: string };
  /** Cells merged across the row (first and last column). */
  merges?: { from: number; to: number }[];
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

  control(r: number, c: number, control: Control, link: string): void {
    this.at(r).control = { col: c, control, link };
  }

  merge(r: number, from: number, to: number): void {
    (this.at(r).merges ??= []).push({ from, to });
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
  if (kind === 'timeline') return FIRST_PERIOD_COL + layout.periods - 1;
  if (kind === 'lookups') return 5;
  let max = TOTAL_COL;
  for (const r of rows) {
    for (const c of Object.keys(r.cells)) max = Math.max(max, Number(c));
    if (r.kind === 'series' && r.span) max = Math.max(max, FIRST_PERIOD_COL + r.span - 1);
  }
  return max;
}

/** A row's inactive condition as a conditional format formula, relative to the first cell it covers. */
function condition(ctx: Ctx, marker: string, col: number): string {
  return renderFormula(`=${marker}`, ctx.name, col, ctx.pos, 'excel', ctx.layout.periods).slice(1);
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
      for (const c of Object.keys(cells).map(Number)) put(c, 'h2');
      sink.row(r, { ht: HEIGHTS.heading, level: level ?? 1 });
      return;
    case 'item':
      // A lookup list's heading or item: its position in C and its value in I, in a thin grid.
      for (const c of Object.keys(cells).map(Number)) {
        put(c, row.role === 'luHead' ? 'luHead' : c === 4 ? (row.role ?? 'lu.text') : 'lu.int');
      }
      sink.row(r, { level: level ?? 1 });
      return;
    case 'setting': {
      put(labelCol, 'label');
      put(UNIT_COL, 'unit');
      const style = row.link ? calcStyle(row.unit) : row.role ?? inputStyle(row.unit);
      put(TOTAL_COL, style);
      if (row.valid && !row.link) sink.valid(r, TOTAL_COL, row.valid);
      if (row.control && row.name && !row.link) sink.control(r, TOTAL_COL, row.control, row.name);
      if (row.inactive && !row.control) sink.cond(ref(r, TOTAL_COL), { kind: 'inactive', formula: condition(ctx, row.inactive, TOTAL_COL) });
      if (row.style === 'wide' && ctx.lastCol > TOTAL_COL) {
        // a text input that runs across the columns to the right (a scenario's description)
        for (let c = TOTAL_COL + 1; c <= ctx.lastCol; c++) sink.format(r, c, { style });
        sink.merge(r, TOTAL_COL, ctx.lastCol);
      }
      sink.row(r, { level: level ?? 1 });
      return;
    }
    case 'scenario': {
      dressScenario(ctx, row, r, cells, put);
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
      if (row.style === 'italic') mods.push('italic');
      put(labelCol, major ? 'h3' : 'label', row.style === 'italic' ? ['italic'] : []);
      put(UNIT_COL, 'unit');
      for (const c of Object.keys(cells).map(Number)) {
        if (c === TOTAL_COL && row.role === 'opening') {
          put(c, inputStyle(row.unit));
        } else if (c === TOTAL_COL || !row.input) {
          if (c >= TOTAL_COL) put(c, check ? 'check' : calcStyle(row.unit), mods);
        } else if (c > TOTAL_COL) {
          put(c, inputStyle(row.unit), mods.filter(m => m !== 'total'));
        }
      }
      if (row.inactive) {
        const lastCol = Math.max(...Object.keys(cells).map(Number));
        sink.cond(`${ref(r, FIRST_PERIOD_COL)}:${ref(r, lastCol)}`, { kind: 'inactive', formula: condition(ctx, row.inactive, FIRST_PERIOD_COL) });
      }
      if (check) sink.cond(`${ref(r, TOTAL_COL)}:${ref(r, Math.max(TOTAL_COL, ...Object.keys(cells).map(Number)))}`, { kind: 'notZero' });
      // workings (rows that only feed other rows) sit at level 2, hidden from the reading view
      sink.row(r, { ht: check ? HEIGHTS.check : undefined, level: level ?? (row.role === 'working' ? 2 : 1), hidden: row.role === 'working' && ctx.std });
      return;
    }
    case 'toc':
    case 'text': {
      dressToc(ctx, row, r, cells, put);
      return;
    }
    case 'table': {
      dressTable(ctx, row, r, cells, put);
      return;
    }
  }
}

/**
 * A row of the Scenarios sheet: a group's heading (with a link to its first module), its count of
 * lines adjusted, a line (the value used, boxed, then the scenarios' inputs with the one in use
 * shaded), the scenarios' names, or the heading over their descriptions.
 */
function dressScenario(ctx: Ctx, row: LRow, r: number, cells: Record<number, unknown>,
  put: (c: number, style: string, mods?: Modifier[]) => void): void {
  const { sink } = ctx;
  const labelCol = LABEL_COLS[Math.min(row.indent, 2)];
  const scn = Object.keys(cells).map(Number).filter(c => c > TOTAL_COL);
  const shade = () => {
    if (!row.selected || !scn.length) return;
    const a = ref(r, FIRST_PERIOD_COL);
    const fixed = `$${colLetter(FIRST_PERIOD_COL)}${r}`;
    sink.cond(`${a}:${ref(r, Math.max(...scn))}`, { kind: 'selected', formula: `${row.selected}=COLUMNS(${fixed}:${a})` });
  };
  switch (row.role) {
    case 'group':
      put(1, 'nav');
      put(labelCol, 'h3');
      for (const [c, l] of Object.entries(row.links || {})) sink.link(r, Number(c), l.to, l.tip);
      sink.row(r, { ht: HEIGHTS.heading, level: row.level ?? 1 });
      return;
    case 'count':
      put(labelCol, 'label');
      put(UNIT_COL, 'unit');
      put(TOTAL_COL, 'int');
      sink.cond(ref(r, TOTAL_COL), { kind: 'on' });
      break;
    case 'names':
      put(labelCol, 'label');
      put(TOTAL_COL, 'selText');
      for (const c of scn) {
        put(c, 'in.switch');
        if (row.valid) sink.valid(r, c, row.valid);
      }
      shade();
      break;
    case 'deschead':
      put(labelCol, 'h3');
      put(TOTAL_COL, 'h3');
      break;
    default:   // a line
      put(labelCol, 'label');
      put(UNIT_COL, 'unit');
      put(TOTAL_COL, 'selPct');
      for (const c of scn) put(c, 'in.pct');
      shade();
  }
  sink.row(r, { level: row.level ?? 1 });
}

/** A row of a report's chart table (or of the Scenarios sheet's results): label, unit, then cells by column. */
function dressTable(ctx: Ctx, row: LRow, r: number, cells: Record<number, unknown>,
  put: (c: number, style: string, mods?: Modifier[]) => void): void {
  const { sink } = ctx;
  const role = row.role ?? '';
  const labelCol = LABEL_COLS[Math.min(row.indent, 2)];
  if (role === 'r.blocks') {
    // a dashboard table's block headings, each merged across its columns
    put(labelCol, 'h3');
    for (const [a, b] of row.merges ?? []) {
      for (let c = a; c <= b; c++) {
        if (c in cells) put(c, 'blockHead');
        else sink.format(r, c, { style: 'blockHead' });
      }
      if (b > a) sink.merge(r, a, b);
    }
    sink.row(r, { ht: HEIGHTS.heading, level: row.level ?? 1 });
    return;
  }
  if (role === 'r.sub' || role === 'r.sum' || role === 'r.var') {
    // a dashboard table: the actual or forecast line under the months, or a statement line with its comparisons
    const styles = new Set(row.style.split('+'));
    const major = styles.has('bold');
    const mods: Modifier[] = [];
    if (major) mods.push('total');
    if (styles.has('italic')) mods.push('italic');
    if (styles.has('last')) mods.push('last');
    put(labelCol, role === 'r.sub' ? 'muted' : major ? 'h3' : 'label', styles.has('italic') ? ['italic'] : []);
    put(UNIT_COL, 'unit');
    const style = role === 'r.sub' ? 'colSub' : calcStyle(row.unit || '$');
    const cols = Object.keys(cells).map(Number).filter(c => c > TOTAL_COL);
    for (const c of cols) {
      const own = row.units?.[c];
      put(c, typeof cells[c] === 'string' && !String(cells[c]).startsWith('=') ? 'colSub' : own ? calcStyle(own) : style, role === 'r.sub' ? [] : mods);
    }
    if (row.shade && cols.length) {
      sink.cond(`${ref(r, FIRST_PERIOD_COL)}:${ref(r, FIRST_PERIOD_COL + 11)}`, { kind: 'current', formula: condition(ctx, row.shade, FIRST_PERIOD_COL) });
    }
    if (role === 'r.var' && cols.length) sink.cond(`${ref(r, FIRST_PERIOD_COL)}:${ref(r, Math.max(...cols))}`, { kind: 'adverse' });
    sink.row(r, { level: row.level ?? 1 });
    return;
  }
  if (role === 'r.gap' || role === 'r.group') {
    // inside the Scenarios sheet's data table: a gap, or a heading over a measure's results (each left blank by the table)
    if (role === 'r.group') put(labelCol, 'h3');
    for (const c of Object.keys(cells).map(Number).filter(c => c >= TOTAL_COL)) put(c, 'text');
    sink.row(r, { ht: role === 'r.gap' ? HEIGHTS.spacer : HEIGHTS.heading, level: row.level ?? 1 });
    return;
  }
  const bold = row.style === 'bold';
  const muted = role === 'r.index' || role === 'r.rank' || role === 'r.muted';
  put(labelCol, role === 'r.head' || bold ? 'h3' : muted ? 'muted' : 'label');
  put(UNIT_COL, 'unit');
  const valueStyle = row.style === 'signed' ? 'signed' : row.unit === 'text' ? 'text' : calcStyle(row.unit || '$');
  const cols = Object.keys(cells).map(Number).filter(c => c >= TOTAL_COL);
  for (const c of cols) {
    if (role === 'r.reg') put(c, REG_STYLES[c] ?? 'text');
    else if (role === 'r.store') put(c, c === TOTAL_COL ? 'key' : 'int');
    else if (role === 'r.title') put(c, 'h3');
    else if (role === 'r.head') put(c, 'colHead');
    else if (role === 'r.check') put(c, 'check');
    else if (muted) put(c, row.unit === 'text' ? 'muted' : 'mutedNum');
    else put(c, valueStyle, bold ? ['bold'] : []);
  }
  if (role === 'r.check') sink.cond(ref(r, TOTAL_COL), { kind: 'notZero' });
  if (role === 'r.reg') sink.cond(ref(r, REG.changed), { kind: 'notZero' });
  const values = cols.filter(c => c > TOTAL_COL);
  if (!role && values.length) {
    const last = Math.max(...values);
    sink.cond(`${ref(r, FIRST_PERIOD_COL)}:${ref(r, last)}`, { kind: 'na', formula: `ISNA(${ref(r, FIRST_PERIOD_COL)})` });
  }
  sink.row(r, { level: row.level ?? 1 });
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
  const timeline = kind === 'timeline';
  const band = kind === 'scenarios' && ctx.std;
  for (const [r, c, v] of frameCells(layout, name, 'excel')) {
    let style = 'label';
    if (!ctx.std) style = r === 1 ? 'title' : r === 2 ? 'entity' : 'h3';
    else if (c === 1) style = 'nav';
    else if (r === STD.titleRow) style = 'title';
    else if (r === STD.nameRow) style = 'modelName';
    else if (r === STD.entityRow) style = 'entity';
    else if (band) style = r === STD.scnHead ? (c === 2 ? 'scnBand' : 'scnBandHead') : r === STD.scnNames ? (c === TOTAL_COL ? 'scnActive' : 'scnBandName') : 'scnMarker';
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
  if (band && layout.scenarioCount) {
    // The band runs from B to the last scenario; the active scenario's name stands upright and its marker goes dark.
    const J = FIRST_PERIOD_COL;
    const L = J + layout.scenarioCount - 1;
    for (let c = 3; c <= TOTAL_COL; c++) sink.format(STD.scnHead, c, { style: 'scnBand' }, true);
    const at = (r: number) => `COLUMNS($${colLetter(J)}${r}:${colLetter(J)}${r})=Sel_Scenario`;
    sink.cond(`${ref(STD.scnNames, J)}:${ref(STD.scnNames, L)}`, { kind: 'upright', formula: at(STD.scnNames) });
    sink.cond(`${ref(STD.scnMarker, J)}:${ref(STD.scnMarker, L)}`, { kind: 'marker', formula: at(STD.scnMarker) });
    sink.row(STD.scnHead, { ht: HEIGHTS.heading });
    sink.row(STD.scnNames, { ht: HEIGHTS.heading });
    sink.row(STD.scnMarker, { ht: HEIGHTS.heading });
    sink.row(STD.scnMarker + 1, { ht: HEIGHTS.spacer });
    return;
  }
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
  } else if (kind === 'scenarios') {
    col(1, 1, 3.75); col(2, 6, 2.5); col(7, 7, 34); col(8, 8, 7); col(9, 9, 14); col(FIRST_PERIOD_COL, FIRST_PERIOD_COL + 5, 14);
  } else if (kind === 'lookups') {
    col(1, 1, 3.75); col(2, 2, 2.5); col(3, 3, 5); col(4, 4, 30); col(5, 5, 30);
  } else if (kind === 'report') {
    col(1, 1, 3.75); col(2, 6, 2.5); col(7, 7, 34); col(8, 8, 7);
    const rows = layout.sheets.find(([s]) => s === sheet)?.[1] ?? [];
    if (rows.some(r => r.role === 'r.reg')) {
      // the register of saved versions: columns sized to what they hold
      col(9, 9, 6);
      for (const [c, w] of Object.entries(REG_WIDTHS)) col(Number(c), Number(c), w);
    } else {
      col(9, 9, 14);
      col(FIRST_PERIOD_COL, FIRST_PERIOD_COL + REPORT_COLS - 1, REPORT_COL_WIDTH);
    }
  } else {
    col(1, 1, 3.75); col(2, 6, 2.5); col(7, 7, 34); col(8, 8, 7);
    col(9, 9, kind === 'settings' ? 28 : kind === 'timeline' ? 11.75 : 14);
    if (kind === 'timeline') {
      col(FIRST_PERIOD_COL, FIRST_PERIOD_COL + layout.periods - 1, 11.75);
    } else if (kind === 'register' || kind === 'list') {
      const widths: Record<keyof typeof RC, number> = { source: 44, owner: 16, updated: 14, evidence: 32, age: 12, group: 46, reason: 34, status: 18 };
      for (const [k, w] of Object.entries(widths)) col(RC[k as keyof typeof RC], RC[k as keyof typeof RC], w);
    }
  }
  return {
    cols, defaultHeight: HEIGHTS.body, summaryBelow: kind !== 'contents',
    freeze: kind === 'timeline' ? { row: STD.freezeRow, col: FIRST_PERIOD_COL } : kind === 'scenarios' ? { row: STD.scnFreezeRow, col: 1 } : { row: 4, col: 1 },
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
    for (const t of layout.dataTables.filter(x => x.sheet === name)) {
      const [, r0] = pos.get(t.first)!;
      const [, r1] = pos.get(t.last)!;
      const inputRow = layout.names.get(t.input);
      const at = inputRow ? pos.get(inputRow) : undefined;
      if (!at || at[0] !== name) throw new Error(`data table on ${name}: its input ${t.input} must be on the same sheet`);
      const cell = sheet.get(r0, FIRST_PERIOD_COL) ?? { s: 0 };
      sheet.set(r0, FIRST_PERIOD_COL, { s: cell.s, dt: { ref: `${ref(r0, FIRST_PERIOD_COL)}:${ref(r1, FIRST_PERIOD_COL + t.cols - 1)}`, r1: ref(at[1], layout.nameCol(t.input)) } });
    }
    Object.assign(sheet, sheetFormat(layout, name));
    if (ctx.std) {
      // Every row carries its height, so a reader that sizes unmarked rows to their font (LibreOffice)
      // keeps charts and controls over the rows they are anchored to.
      const used = new Set([...sheet.rows.keys(), ...sheet.cells.keys()]);
      for (const r of used) {
        const p = sheet.rows.get(r) ?? {};
        if (p.ht === undefined) sheet.row(r, { ...p, ht: sheet.defaultHeight });
      }
    }
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
