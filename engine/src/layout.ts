// What the engine lays out: rows on sheets, the names on them, the charts, and the link records.

import { FIRST_ROW, TOTAL_COL } from './frame.ts';

export type CellValue = string | number | boolean | null;
/** Column number -> a value, or a formula holding markers until it is rendered. */
export type Cells = Record<number, unknown>;

export type RowKind =
  | 'heading' | 'subheading' | 'section' | 'setting' | 'series' | 'text' | 'blank' | 'toc' | 'fixed';

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
    return JSON.stringify([this.kind, this.label, this.indent, this.unit, this.style, this.name, this.first,
      this.formula, this.total, cells, this.span, this.link]);
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

export type SheetKind = 'contents' | 'cover' | 'timeline' | 'list' | 'register';

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
  headlines: Headline[];

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

  /** Row id -> [sheet, row number]; "@Sheet" is the top of a sheet, for navigation names. */
  positions(): Map<string, [string, number]> {
    const pos = new Map<string, [string, number]>();
    for (const [sheet, rows] of this.sheets) {
      pos.set(`@${sheet}`, [sheet, 1]);
      rows.forEach((r, k) => pos.set(r.id, [sheet, FIRST_ROW + k]));
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
