// The module library: area order, sections and module definitions. Kept in YAML in the repo and
// bundled as JSON for the add-in (LibraryBundle); src/node/library.ts reads the YAML folder.

import { AssemblyError } from './frame.ts';

export interface SettingDef {
  key: string;
  label: string;
  unit?: string;
  default?: unknown;
  display?: boolean;
  group?: { item: string; formula?: string };
  /** A drop-down of these items: one list for the module on the Lookups sheet; the value is the position chosen. */
  choice?: string[];
  /** A drop-down of an existing list (a category list another module keeps, or a frame list). */
  list?: string;
  /** A check box: TRUE or FALSE. */
  check?: boolean;
  /** A model-wide name in place of the instance name (a single module's rate other modules read, such as GST_Rate). */
  name?: string;
  /** Inactive (greyed) unless this marker condition holds, such as "$method=2". */
  when?: string;
}

export interface RowDef {
  key?: string;
  label?: string;
  unit?: string;
  style?: string;
  name?: string;
  first?: string;
  formula?: string;
  total?: string;
  span?: number;
  check?: string;
  section?: string;
  collect?: string;
  headline?: HeadlineDef | HeadlineDef[];
  /** A time series input: typed in every month, in forecast months only, or in actual months only (greyed elsewhere). */
  input?: 'all' | 'forecast' | 'actual';
  /** A time series input's value when the instance brings none. */
  default?: number;
  /** Inactive (greyed) unless this marker condition holds, such as "$method=2". */
  when?: string;
  /** The row has a line on the historical income statement or balance sheet, which [hist:key] reads. */
  history?: { in: 'is' | 'bs'; group: string; label?: string };
  /**
   * The row has a scenario adjustment on the Scenarios sheet, which {scenario} (or [scn:key]) reads.
   * The adjustment sits under its group there (Revenue, Cost of sales) as a line named after the
   * module's title; a string, or a label, adds the row's own name where a module has more than one.
   */
  scenario?: boolean | string | { group?: string; label?: string };
  /** A working row: grouped at level 2, out of the reading view. */
  working?: boolean;
  /** Italic detail or ratio line. */
  italic?: boolean;
  /** A collect row: a range name over the rows it collects, from column J across its span (the seasonality profiles). */
  range?: string;
}

export interface HeadlineDef {
  label: string;
  measure?: string;
}

export interface ChartDef {
  key: string;
  title: string;
  categories: string;
  series: { each?: string; row?: string; line?: boolean }[];
}

/**
 * A chart a summary or report module brings (framework: report): a recipe the engine expands into
 * formula rows on the module's sheet (every number reads the statements) and a native chart over them.
 */
export interface ReportChartDef {
  /** The chart's id in the chart register (C01). */
  id: string;
  title: string;
  recipe: 'compare' | 'mix' | 'depth' | 'pie' | 'combo' | 'budget' | 'scenario' | 'bridge' | 'movement'
    | 'variance' | 'trend' | 'versions' | 'accuracy' | 'walk' | 'statement'
    /** Dashboard tables (no chart): a statement month by month with comparisons, quarters against a year earlier, or the budget against the comparison. */
    | 'summary' | 'quarters' | 'budget_table';
  /** A statement line (a key of the statements module's report lines). */
  line?: string;
  /** A group (category lines, such as rev) or a fixed set (such as current_assets). */
  group?: string;
  /** Lines or columns, for compare. */
  kind?: 'line' | 'column';
  /** The months shown: the year shown, the 12 months to the month shown, the month shown, or the year to it. */
  frame?: 'year' | 'rolling' | 'at' | 'ytd' | 'budget';
  cumulative?: boolean;
  periods?: ('prior' | 'shown' | 'next')[];
  /** Rank a group by the period's total and show the top N, the rest as Other. */
  top?: number;
  /** A line stacked in front of the group (mix). */
  lead?: string;
  /** The line the prior and next periods' totals read (mix); the group's total by default. */
  total?: string;
  /** Draw the prior and next periods as lines (mix); true by default. */
  compare?: boolean;
  bars?: string[];
  lines?: string[];
  grouping?: 'clustered' | 'stacked';
  /** Scenario charts: by financial year or by month of the year shown. */
  by?: 'year' | 'month';
  /** Dashboard tables: the rows, top to bottom. */
  rows?: SummaryRowDef[];
  /** Summary tables: the comparison blocks to the right of the months (the 12 months before, the year to date, a year earlier). */
  against?: ('prior' | 'ytd' | 'year_ago')[];
  /** Summary tables: whether the months add up to a total column (flows), or the table holds balances (no total). */
  totals?: boolean;
}

/**
 * A row of a dashboard table: a statement line, a group's lines (each in italic, the last with a
 * dashed rule), or a ratio of two lines in the same table (with a line taken off the numerator).
 */
export interface SummaryRowDef {
  line?: string;
  group?: string;
  ratio?: [string, string];
  less?: string;
  label?: string;
  /** bold: a major result (bold, rule above); italic: detail; last: the last item of a list (a dashed rule under it); plain by default. */
  style?: 'bold' | 'italic' | 'last';
  /** A ratio's unit: % (default) or x (times). */
  unit?: '%' | 'x';
  /** A balance read at the start of each month rather than the end (opening cash). */
  at?: 'start';
}

/** What reports can read from the statements module (framework: statements): lines, category groups and fixed sets. */
export interface ReportSpec {
  /** Key -> the statements row and the label charts show; flows add up over months, balances are read at a month. */
  lines: Record<string, { row: string; label: string; kind?: 'flow' | 'balance' | 'ratio' }>;
  /** Key -> the link whose senders are the group's lines (one row each on the statements). */
  groups: Record<string, { link: string; label: string; other: string }>;
  /** Fixed sets of lines: stacks (with a total and an optional current line), items, or a bridge (start, steps, end). A key with a leading minus is shown negative. */
  sets: Record<string, { stacks?: string[]; items?: string[]; total?: string; current?: string; start?: string; steps?: string[]; end?: string;
    startLabel?: string; endLabel?: string }>;
}

export interface ModuleDef {
  id: string;
  title: string;
  area: string;
  code?: string;
  as_category?: boolean;
  /** A category module whose instances make a drop-down list (collection profiles, asset categories). */
  list?: { name: string; title: string; first?: string[] };
  /** Help shown in the add-in. */
  description?: string;
  mirror?: string;
  framework?: string;
  settings?: SettingDef[];
  inputs?: { link: string; mode?: string; required?: boolean }[];
  rows?: RowDef[];
  outputs?: { link: string; row: string }[];
  charts?: ChartDef[];
  /** A summary or report module's charts (framework: report). */
  reports?: ReportChartDef[];
  /** The statements module's lines, groups and sets for reports (framework: statements). */
  report?: ReportSpec;
  /** The sheet the versions module keeps its values on (framework: versions). */
  store?: string;
}

export interface SectionDef {
  title: string;
  cover?: string | null;
  note?: string;
  areas: string[];
}

export type ModuleKind = 'single' | 'category' | 'mirror';

/** The historical statements: their sheets and the groups lines sit in, in order. */
export interface HistoryDef {
  is?: { sheet: string; title: string; groups: string[] };
  bs?: { sheet: string; title: string; groups: { name: string; side: 'asset' | 'liability' | 'equity' }[] };
}

/** The Scenarios sheet: its area, the scenarios' names and a description of each. */
export interface ScenariosDef {
  sheet: string;
  names: string[];
  descriptions?: string[];
}

/** The library as one JSON document: what the add-in ships. */
export interface LibraryBundle {
  /** Which library a model was built from, kept in its metadata ("demo" when absent: the Phase 0 proof's). */
  id?: string;
  areas: string[];
  sections?: SectionDef[];
  modules: ModuleDef[];
  history?: HistoryDef;
  /** The Scenarios sheet's area, when modules carry scenario adjustments: its scenarios and what each stands for. */
  scenarios?: ScenariosDef;
  /** Lists several modules' drop-downs share (GST treatment): put on the Lookups sheet when a setting reads one. */
  lists?: { name: string; title: string; items: string[] }[];
}

export class Library {
  id = 'demo';
  areas: string[];
  modules: Map<string, ModuleDef>;
  sections: SectionDef[];
  history: HistoryDef;
  scenarios: ScenariosDef | null;
  lists: { name: string; title: string; items: string[] }[];

  constructor(areas: string[], modules: Map<string, ModuleDef>, sections: SectionDef[] = [], history: HistoryDef = {},
    scenarios: ScenariosDef | null = null, lists: { name: string; title: string; items: string[] }[] = []) {
    this.areas = areas;
    this.modules = modules;
    this.sections = sections;
    this.history = history;
    this.scenarios = scenarios;
    this.lists = lists;
  }

  /** A library from its bundle: areas.yaml's content and every module definition in file name order. */
  static fromBundle(bundle: LibraryBundle): Library {
    const areas = bundle.areas;
    const sections = bundle.sections || [];
    const placed = sections.flatMap(s => s.areas);
    if (sections.length && [...placed].sort().join('\u0000') !== [...areas].sort().join('\u0000')) {
      throw new AssemblyError('areas.yaml: every area must sit in exactly one section');
    }
    const modules = new Map<string, ModuleDef>();
    for (const d of bundle.modules) {
      if (!areas.includes(d.area)) throw new AssemblyError(`${d.id}: area '${d.area}' is not in areas.yaml`);
      modules.set(d.id, d);
    }
    for (const h of [bundle.history?.is?.sheet, bundle.history?.bs?.sheet, bundle.scenarios?.sheet]) {
      if (h && !areas.includes(h)) throw new AssemblyError(`areas.yaml: the sheet '${h}' is not an area`);
    }
    const lib = new Library(areas, modules, sections, bundle.history ?? {}, bundle.scenarios ?? null, bundle.lists ?? []);
    if (bundle.id) lib.id = bundle.id;
    return lib;
  }

  /** The bundle this library came from, for the add-in. */
  toBundle(): LibraryBundle {
    const b: LibraryBundle = { areas: this.areas, sections: this.sections, modules: [...this.modules.values()] };
    if (this.id !== 'demo') b.id = this.id;
    if (Object.keys(this.history).length) b.history = this.history;
    if (this.scenarios) b.scenarios = this.scenarios;
    if (this.lists.length) b.lists = this.lists;
    return b;
  }

  module(id: string): ModuleDef {
    const m = this.modules.get(id);
    if (!m) throw new AssemblyError(`no module '${id}' in the library`);
    return m;
  }

  kind(id: string): ModuleKind {
    const m = this.module(id);
    if (m.mirror) return 'mirror';
    return m.as_category ? 'category' : 'single';
  }
}
