// What the engine lays out: rows on sheets, the names on them, the charts, and the link records.

import { PROOF_FRAME, TOTAL_COL, type Frame } from './frame.ts';

export type CellValue = string | number | boolean | null;
/** Column number -> a value, or a formula holding markers until it is rendered. */
export type Cells = Record<number, unknown>;

export type RowKind =
  | 'heading' | 'subheading' | 'section' | 'setting' | 'series' | 'text' | 'blank' | 'toc' | 'fixed' | 'item' | 'scenario';

export interface LRowFields {
  indent: number;
  unit: string;
  /** total, check, bold, reg, rate */
  style: string;
  /** A setting's value. */
  value: unknown;
  /** The defined name on the value or total cell. */
  name: string | null;
  /** Marker formula for the first period. */
  first: string | null;
  /** Marker formula for later periods. */
  formula: string | null;
  /** sum, last or none. */
  total: string;
  /** Column -> marker formula or text (contents and assurance rows). */
  cells: Cells;
  /** Periods written, from the first; null means every period. */
  span: number | null;
  /** A setting bound to a group assumption: the marker formula written in place of a value. */
  link: string | null;
  /** Standard frame: a spacer row's height in points (3, 6 or 9). */
  space?: number;
  /** Standard frame: the row's outline level (0 to 2), when it is not the frame's default for its kind. */
  level?: number;
  /** Standard frame: cell hyperlinks by column, each to a defined name with a screen tip. */
  links?: Record<number, CellLink>;
  /** Standard frame: what the row is for, when its kind alone does not say (last item of a list, a section number, a note). */
  role?: string;
  /** Standard frame: the validation on a setting's value cell. */
  valid?: Validation;
  /** Standard frame: a classic control over a setting's value cell, which holds its link. */
  control?: Control;
  /** A time series input: which months take typed values (the others are greyed). */
  input?: 'all' | 'forecast' | 'actual';
  /** A time series input's starting values by period (null for none). */
  values?: (number | null)[];
  /** Standard frame: a marker condition under which the row's input cells are greyed out (a conditional format). */
  inactive?: string;
}

/**
 * A form control over a setting's value cell. A drop-down's cell holds the position chosen in a
 * List_ range on the Lookups sheet (Sel_ names); a check box's cell holds TRUE or FALSE (Opt_ names).
 */
export type Control = { kind: 'drop'; list: string; lines?: number } | { kind: 'check' };

/** A defined name over a run of rows in one column (the List_ names on the Lookups sheet). */
export interface RangeName {
  /** First and last row ids. */
  from: string;
  to: string;
  col: number;
}

/** What a setting's value cell accepts; every rule carries its own error message. */
export type Validation =
  | { kind: 'list'; items: string[]; message: string }
  | { kind: 'whole'; min: number | string; max: number | string; message: string }
  | { kind: 'decimal'; message: string }
  | { kind: 'date'; message: string }
  | { kind: 'text'; max: number; message: string }
  /** TRUE or FALSE only: a check box's cell. */
  | { kind: 'logical'; message: string };

export interface CellLink {
  /** The defined name the link goes to. */
  to: string;
  /** Screen tip. */
  tip: string;
}

/** One row of a sheet. Field names and order follow the Python proof so the two compare directly. */
export class LRow implements LRowFields {
  id: string;
  kind: RowKind;
  label: string;
  indent = 0;
  unit = '';
  style = '';
  value: unknown = null;
  name: string | null = null;
  first: string | null = null;
  formula: string | null = null;
  total = 'sum';
  cells: Cells = {};
  span: number | null = null;
  link: string | null = null;
  declare space?: number;
  declare level?: number;
  declare links?: Record<number, CellLink>;
  declare role?: string;
  declare valid?: Validation;
  declare control?: Control;
  declare input?: 'all' | 'forecast' | 'actual';
  declare values?: (number | null)[];
  declare inactive?: string;

  constructor(id: string, kind: RowKind, label: string, init: Partial<LRowFields> = {}) {
    this.id = id;
    this.kind = kind;
    this.label = label;
    for (const [k, v] of Object.entries(init)) {
      if (v !== undefined) (this as Record<string, unknown>)[k] = v;
    }
  }

  /** What the engine owns in this row. Input values are left out: a structural change never overwrites what someone typed. */
  signature(): string {
    const cells = Object.keys(this.cells).map(Number).sort((a, b) => a - b).map(c => [c, this.cells[c]]);
    const sig: unknown[] = [this.kind, this.label, this.indent, this.unit, this.style, this.name, this.first,
      this.formula, this.total, cells, this.span, this.link];
    if (this.space !== undefined || this.level !== undefined || this.links !== undefined || this.role !== undefined
      || this.valid !== undefined || this.control !== undefined) {
      sig.push(this.space ?? null, this.level ?? null, this.links ?? null, this.role ?? null, this.valid ?? null);
      if (this.control !== undefined) sig.push(this.control);
    }
    if (this.input !== undefined || this.inactive !== undefined) sig.push(this.input ?? null, this.inactive ?? null);
    return JSON.stringify(sig);
  }
}

/** A chart a module carries. Series point at rows by id, so the engine can re-point them. */
export class ChartSpec {
  id: string;
  sheet: string;
  title: string;
  /** Row id the chart's top-left corner sits on. */
  anchor: string;
  /** Row id whose cells label the x axis. */
  categories: string;
  /** Periods shown. */
  span: number;
  /** [row id, "column" or "line"]. */
  series: [string, 'column' | 'line'][];

  constructor(id: string, sheet: string, title: string, anchor: string, categories: string, span: number,
    series: [string, 'column' | 'line'][]) {
    this.id = id;
    this.sheet = sheet;
    this.title = title;
    this.anchor = anchor;
    this.categories = categories;
    this.span = span;
    this.series = series;
  }

  signature(): string {
    return JSON.stringify([this.sheet, this.title, this.anchor, this.categories, this.span, this.series]);
  }
}

export interface LinkRecord {
  link: string;
  mode: 'each' | 'total' | 'mirror';
  from: string;
  to: string;
}

/** A key output: what every structural change is checked against. */
export interface Headline {
  id: string;
  label: string;
  measure: string;
  unit: string;
  name: string;
}

export type SheetKind = 'contents' | 'cover' | 'timeline' | 'settings' | 'lookups' | 'scenarios' | 'list' | 'register';

export class Layout {
  periods: number;
  sheets: [string, LRow[]][];
  /** Name -> row id, in the order the names were made. */
  names: Map<string, string>;
  records: LinkRecord[];
  warnings: string[];
  /** Block id -> title. */
  blocks: Map<string, string>;
  charts: ChartSpec[];
  kinds: Record<string, SheetKind> = {};
  /** Sheet -> the title shown in B1. */
  titles: Record<string, string> = {};
  /** Names that point somewhere other than column I. */
  nameCols: Record<string, number> = {};
  /** Names over a run of rows (List_ ranges), in the order made. */
  ranges = new Map<string, RangeName>();
  /** Sheet -> the heading of column I on its timeline block ("Total" unless set: "Opening" on the historical balance sheet). */
  totalHeads: Record<string, string> = {};
  headlines: Headline[];
  frame: Frame = PROOF_FRAME;

  constructor(periods: number, sheets: [string, LRow[]][], names: Map<string, string>, records: LinkRecord[],
    warnings: string[], blocks: Map<string, string>, charts: ChartSpec[] = [], headlines: Headline[] = []) {
    this.periods = periods;
    this.sheets = sheets;
    this.names = names;
    this.records = records;
    this.warnings = warnings;
    this.blocks = blocks;
    this.charts = charts;
    this.headlines = headlines;
  }

  /** The kind of a sheet; sheets the frame does not list are timeline sheets. */
  kindOf(sheet: string): SheetKind {
    return Object.hasOwn(this.kinds, sheet) ? this.kinds[sheet] : 'timeline';
  }

  /** The row a sheet's first layout row sits on. */
  firstRowOf(sheet: string): number {
    return this.frame.firstRow(this.kindOf(sheet));
  }

  /**
   * Row id -> [sheet, row number]. "@Sheet" is the top of a sheet, for navigation names; in the
   * standard frame "@Sheet/2" and "@Sheet/3" are its model name and entity rows.
   */
  positions(): Map<string, [string, number]> {
    const pos = new Map<string, [string, number]>();
    for (const [sheet, rows] of this.sheets) {
      pos.set(`@${sheet}`, [sheet, 1]);
      if (this.frame.id === 'standard') {
        pos.set(`@${sheet}/2`, [sheet, 2]);
        pos.set(`@${sheet}/3`, [sheet, 3]);
      }
      const first = this.firstRowOf(sheet);
      rows.forEach((r, k) => pos.set(r.id, [sheet, first + k]));
    }
    return pos;
  }

  nameCol(name: string): number {
    return Object.hasOwn(this.nameCols, name) ? this.nameCols[name] : TOTAL_COL;
  }

  hasChecks(): boolean {
    return this.names.has('Chk_Errors');
  }

  sheetRows(): Map<string, LRow[]> {
    return new Map(this.sheets);
  }
}
