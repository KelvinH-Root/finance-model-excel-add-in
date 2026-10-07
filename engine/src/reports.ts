// Summary and report modules (framework: report). Each chart a module brings is a recipe; the
// engine expands it into a table on the module's sheet, every number a formula on the statements,
// and a native chart over the table. The selections at the top of the module (the year shown, the
// month shown) are drop-downs, so changing one, the active scenario or the last actual month
// redraws every chart without the add-in. Charts sit in a grid above the tables.
//
// Recipes (the chart register, prototypes/reports/register.yaml, in HFG's frame):
//   compare   one line over the period before, the period shown and the period after
//   mix       a group stacked by month, with the group total for the periods before and after as lines
//   depth     a group's totals for the period, one column per line
//   pie       a group's make-up for the period
//   combo     chosen lines as columns (clustered or stacked) with lines drawn over them
//   budget    actual (solid) and forecast (hatched) by month against the comparison as a line
//   scenario  the line in each scenario, by financial year or by month of the year shown
//   bridge    a waterfall: net assets built up from the balance sheet, or opening to closing cash
//   movement  the change in each balance sheet line against a year earlier

import { AssemblyError, colLetter, FIRST_PERIOD_COL, LABEL_COLS, TOTAL_COL, UNIT_COL } from './frame.ts';
import { LRow, type RChart, type RSeries } from './layout.ts';
import type { ReportChartDef, ReportSpec } from './library.ts';

const J = FIRST_PERIOD_COL;
/** Charts per row of the grid, and sheet rows per row of charts. */
export const GRID = { cols: 3, rows: 20 } as const;
/** Report charts: 12.2 cm by 7.6 cm, as in the chart register proof. */
export const RCHART_CM = { w: 12.2, h: 7.6 } as const;

type Frame = 'year' | 'rolling' | 'at' | 'ytd';

/** A scenario result: one output formula the Scenarios sheet's data table works out for every scenario. */
export interface ScenarioResult {
  id: string;
  label: string;
  /** Marker formula in column I. */
  formula: string;
}

export interface ReportContext {
  /** The report block's id (rows are made under it) and its instance's code for names (IncS1). */
  block: string;
  code: string;
  title: string;
  charts: ReportChartDef[];
  /** The statements block's id, and what reports can read from it. */
  fs: string;
  spec: ReportSpec;
  /** The statements rows of a group's lines, in order. */
  members: (link: string) => string[];
  /** The module's selections: Sel_ names of the year shown, the month shown and Compared with, when it has them. */
  year: string | null;
  month: string | null;
  compare: string | null;
  /** Labels as the default selections show them, for chart title caches. */
  fyLabel: (k: number) => string;
  monthLabel: (p: number) => string;
  yearShown: number;
  monthShown: number;
  /** Financial years in the timeline. */
  years: number;
  /** Scenario names, when the model has a Scenarios sheet. */
  scenarios: string[] | null;
  /** Data table rows the Scenarios sheet makes; a report reads its row's J to L. */
  results: ScenarioResult[];
}

export interface ReportOut {
  rows: LRow[];
  charts: RChart[];
  /** Marker references to the module's per-chart check cells. */
  errors: string[];
  alerts: string[];
  /** Names over the module's scalar rows (year shown starts at, and so on) -> row id. */
  names: Map<string, string>;
}

/** A value of a line or group at period i, blank (#N/A, a gap in a chart) outside the timeline. */
const pickAt = (rng: string, i: string) => `IF(OR(${i}<1,${i}>Tl_Term),NA(),INDEX(${rng},${i}))`;
const spanSum = (rng: string, a: string, b: string) => `SUM(INDEX(${rng},${a}):INDEX(${rng},${b}))`;

interface TableRow {
  id: string;
  label?: string;
  /** A label that is a formula (it reads a cell). */
  labelF?: string;
  unit?: string;
  total?: string;
  values: string[];
  role?: string;
  style?: string;
}

class Table {
  readonly rows: LRow[] = [];
  readonly head: string;
  readonly titleRow: string;
  readonly base: string;
  ncat = 0;
  private k = 0;

  readonly rep: Expander;
  readonly ch: ReportChartDef;

  constructor(rep: Expander, ch: ReportChartDef, titleFormula: string) {
    this.rep = rep;
    this.ch = ch;
    this.base = `${rep.ctx.block}/r/${ch.id}`;
    this.rows.push(new LRow(`${this.base}/sp`, 'blank', '', { space: 6 }));
    this.rows.push(new LRow(`${this.base}/section`, 'section', `${ch.id}  ${ch.title}`, { indent: 1 }));
    this.titleRow = `${this.base}/title`;
    this.rows.push(new LRow(this.titleRow, 'table', 'Chart title', { indent: 2, role: 'r.title', cells: { [TOTAL_COL]: titleFormula } }));
    this.head = `${this.base}/head`;
  }

  header(cats: string[], label: string, total = ''): void {
    const cells: Record<number, unknown> = { [TOTAL_COL]: total };
    cats.forEach((c, j) => { cells[J + j] = c; });
    this.rows.push(new LRow(this.head, 'table', label, { indent: 2, role: 'r.head', cells }));
    this.ncat = cats.length;
  }

  /** The id the next row will take. */
  nextId(): string {
    return `${this.base}/${this.k + 1}`;
  }

  add(r: Omit<TableRow, 'id'>): string {
    this.k += 1;
    const id = `${this.base}/${this.k}`;
    const cells: Record<number, unknown> = {};
    if (r.labelF) cells[LABEL_COLS[2]] = r.labelF;
    if (r.total) cells[TOTAL_COL] = r.total.replaceAll('{self}', id);
    r.values.forEach((v, j) => { cells[J + j] = v.replaceAll('{self}', id); });
    this.rows.push(new LRow(id, 'table', r.label ?? '', { indent: 2, unit: r.unit ?? '', role: r.role, style: r.style ?? '', cells }));
    return id;
  }

  /** A check on the table's figures: 1 when it fails, in column I. */
  check(label: string, formula: string, kind: 'error' | 'alert'): void {
    const id = `${this.base}/chk/${kind}/${this.rows.length}`;
    this.rows.push(new LRow(id, 'table', label, { indent: 2, unit: 'flag', role: 'r.check', cells: { [TOTAL_COL]: formula } }));
    (kind === 'error' ? this.rep.errors : this.rep.alerts).push(`«V|${id}»`);
  }
}

/** The month a period-of-year column is in, across a table: J{r} to U{r} of a row. */
const across = (id: string, n: number) => `«C${J}|${id}»:«C${J + n - 1}|${id}»`;
/** A flow's total for twelve months, blank unless all twelve are in the timeline. */
const fullTotal = (n = 12) => `=IF(COUNT(${across('{self}', n)})=${n},SUM(${across('{self}', n)}),"")`;

const PRIOR_STYLE: Partial<RSeries> = { colour: 'grey', width: 1.25 };
const SHOWN_STYLE: Partial<RSeries> = { colour: 'tx2', width: 2.25, marker: true };
const NEXT_STYLE: Partial<RSeries> = { colour: 'accent1', width: 1.5, dash: true };
const ACCENTS = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];

class Expander {
  readonly rows: LRow[] = [];
  readonly charts: RChart[] = [];
  readonly errors: string[] = [];
  readonly alerts: string[] = [];
  readonly names = new Map<string, string>();
  readonly ctx: ReportContext;
  /** Row ids of the selection scalars and index rows. */
  private y0 = '';
  private ys = '';
  private ye = '';
  private ytd0 = '';
  private cmpRow = '';
  private cmpLabel = '';
  private yidx = '';
  private ridx = '';
  private mcats = '';
  private rcats = '';
  readonly grid: string;

  constructor(ctx: ReportContext) {
    this.ctx = ctx;
    this.grid = `${ctx.block}/r/grid`;
  }

  n(what: string): string {
    return `${this.ctx.code}_${what}`;
  }

  where(ch: ReportChartDef): string {
    return `${this.ctx.title} chart ${ch.id}`;
  }

  line(ch: ReportChartDef, key: string): { rng: string; label: string; kind: 'flow' | 'balance' | 'ratio' } {
    const l = this.ctx.spec.lines[key];
    if (!l) throw new AssemblyError(`${this.where(ch)}: the statements have no report line '${key}'`);
    return { rng: `«A|${this.ctx.fs}/${l.row}»`, label: l.label, kind: l.kind ?? 'flow' };
  }

  group(ch: ReportChartDef, key: string): { ids: string[]; blk: string; labels: string; label: string; other: string } {
    const g = this.ctx.spec.groups[key];
    if (!g) throw new AssemblyError(`${this.where(ch)}: the statements have no report group '${key}'`);
    const ids = this.ctx.members(g.link);
    return {
      ids, label: g.label, other: g.other,
      blk: ids.length ? `«G|${ids[0]}|${ids[ids.length - 1]}»` : '',
      labels: ids.length ? `«K|${ids[0]}|${ids[ids.length - 1]}|${LABEL_COLS[2]}»` : '',
    };
  }

  /** A line or a set member, signed: "-cl" is current liabilities shown negative. */
  signed(ch: ReportChartDef, key: string): { rng: string; label: string; sign: string; kind: string } {
    const neg = key.startsWith('-');
    const l = this.line(ch, neg ? key.slice(1) : key);
    return { rng: l.rng, label: l.label, sign: neg ? '-' : '', kind: l.kind };
  }

  // --- the module's selections and the index rows every table reads -----------------------------
  scalar(key: string, label: string, formula: string, unit = '', name: string | null = null): string {
    const id = `${this.ctx.block}/r/${key}`;
    this.rows.push(new LRow(id, 'table', label, { indent: 2, unit, role: 'r.scalar', cells: { [TOTAL_COL]: formula } }));
    if (name) this.names.set(name, id);
    return id;
  }

  indexRow(key: string, label: string, cell: (j: number) => string, unit = ''): string {
    const id = `${this.ctx.block}/r/${key}`;
    const cells: Record<number, unknown> = {};
    for (let j = 0; j < 12; j++) cells[J + j] = cell(j);
    this.rows.push(new LRow(id, 'table', label, { indent: 2, role: 'r.index', unit, cells }));
    return id;
  }

  /** The window's label for chart titles. */
  windowLabel(frame: Frame): string {
    const M = this.ctx.month;
    const ml = `TEXT(INDEX(List_Months,${M}),"mmm yyyy")`;
    return {
      year: `INDEX(List_Years,${this.ctx.year})`, rolling: `"12 months to "&${ml}`, at: `"at "&${ml}`, ytd: `"year to "&${ml}`,
    }[frame];
  }

  defaultLabel(frame: Frame | 'movement' | 'each'): string {
    const c = this.ctx;
    const ml = c.monthLabel(c.monthShown);
    return {
      year: c.fyLabel(c.yearShown), rolling: `12 months to ${ml}`, at: `at ${ml}`, ytd: `year to ${ml}`,
      movement: `${ml} against a year earlier`, each: 'each year',
    }[frame];
  }

  periodLabel(frame: Frame, which: 'prior' | 'shown' | 'next'): string {
    if (which === 'shown') return `=${this.windowLabel(frame)}`;
    if (frame === 'year') return `="FY"&(Tl_First_FY+${this.ctx.year}${which === 'prior' ? '-2' : ''})`;
    return which === 'prior' ? '="Prior 12 months"' : '="Next 12 months"';
  }

  /** Period index of column j for a frame, shifted by months. */
  idx(frame: Frame, j: number, shift = 0): string {
    const row = frame === 'year' ? this.yidx : this.ridx;
    const base = `«C${J + j}|${row}»`;
    return shift ? `(${base}${shift > 0 ? '+' : ''}${shift})` : base;
  }

  cats(frame: Frame): string[] {
    const row = frame === 'year' ? this.mcats : this.rcats;
    return Array.from({ length: 12 }, (_, j) => `=«C${J + j}|${row}»`);
  }

  bounds(frame: Frame): [string, string] {
    const M = this.ctx.month!;
    if (frame === 'year') return [`«V|${this.ys}»`, `«V|${this.ye}»`];
    if (frame === 'rolling') return [`MAX(1,${M}-11)`, M];
    if (frame === 'ytd') return [`«V|${this.ytd0}»`, M];
    return [M, M];
  }

  needs(frame: Frame | undefined): void {
    if ((frame ?? 'year') === 'year' && !this.ctx.year) throw new AssemblyError(`${this.ctx.title}: a chart reads the year shown, but the module has no year setting`);
    if (frame && frame !== 'year' && !this.ctx.month) throw new AssemblyError(`${this.ctx.title}: a chart reads the month shown, but the module has no month setting`);
  }

  selections(): void {
    const c = this.ctx;
    this.rows.push(new LRow(`${c.block}/r/sel/sp`, 'blank', '', { space: 6 }));
    this.rows.push(new LRow(`${c.block}/r/sel/section`, 'section', 'What the charts show', { indent: 1 }));
    if (c.scenarios) {
      this.scalar('sel/scenario', 'Scenario shown (change it on the Scenarios sheet)', '=Scn_Active_Name', 'text');
    }
    if (c.year) {
      this.scalar('sel/year_label', 'Year shown', `=INDEX(List_Years,${c.year})`, 'text');
      // Month 1 of the year shown as a period number (below 1 when the first year is part of a year).
      this.y0 = this.scalar('sel/y0', 'Year shown: its first month as a period', `=(${c.year}-1)*12-Sel_Start_Month+2`, 'period', this.n('Y0'));
      this.ys = this.scalar('sel/ys', 'Year shown: first month in the timeline', `=MAX(1,«V|${this.y0}»)`, 'period', this.n('Y_Start'));
      this.ye = this.scalar('sel/ye', 'Year shown: last month in the timeline', `=MIN(Tl_Term,«V|${this.y0}»+11)`, 'period', this.n('Y_End'));
    }
    if (c.month) {
      this.scalar('sel/month_label', 'Month shown', `=TEXT(INDEX(List_Months,${c.month}),"mmm yyyy")`, 'text');
      this.ytd0 = this.scalar('sel/ytd0', 'Month shown: first month of its financial year in the timeline',
        `=MAX(1,INT((${c.month}+Sel_Start_Month-2)/12)*12-Sel_Start_Month+2)`, 'period', this.n('YTD_Start'));
    }
    if (c.compare) {
      if (!c.year) throw new AssemblyError(`${c.title}: Compared with needs a year shown`);
      const sel = c.compare;
      // The register row of the version compared with: 0 for the budget being built, -1 when nothing is saved for it.
      this.cmpRow = this.scalar('sel/cmp_row', "Compared with: the version's row in the register (0: the budget being built)",
        `=IFERROR(CHOOSE(MIN(${sel},5),MATCH("Budget|Approved|"&INDEX(List_Years,${c.year}),VR_BudgetKey,0),`
        + `MATCH("Reforecast|"&(Tl_Last_Actual-1),VR_RefKey,0),MATCH("Reforecast|"&MAX(VR_RefAt),VR_RefKey,0),0,${sel}-4),-1)`,
        '#', this.n('Cmp_Row'));
      this.scalar('sel/cmp_id', 'Compared with: version id', `=IF(«V|${this.cmpRow}»>0,INDEX(VR_Id,«V|${this.cmpRow}»),"")`, 'text', this.n('Cmp_Id'));
      this.cmpLabel = this.scalar('sel/cmp_label', 'Compared with',
        `=IF(«V|${this.cmpRow}»=0,"Budget being built",IF(«V|${this.cmpRow}»<0,"Nothing saved",INDEX(VR_Label,«V|${this.cmpRow}»)))`, 'text', this.n('Cmp_Label'));
      const id = `${c.block}/r/sel/cmp_check`;
      this.rows.push(new LRow(id, 'table', 'Nothing is saved for the comparison chosen, so it is blank', { indent: 2, unit: 'flag', role: 'r.check',
        cells: { [TOTAL_COL]: `=IF(«V|${this.cmpRow}»<0,1,0)` } }));
      this.alerts.push(`«V|${id}»`);
    }
    // The chart grid
    this.rows.push(new LRow(`${c.block}/r/grid/sp`, 'blank', '', { space: 6 }));
    this.rows.push(new LRow(`${c.block}/r/grid/section`, 'section', 'Charts', { indent: 1 }));
    const gridRows = Math.ceil(c.charts.length / GRID.cols) * GRID.rows;
    for (let k = 0; k < gridRows; k++) {
      this.rows.push(new LRow(k === 0 ? this.grid : `${this.grid}/${k}`, 'blank', '', { space: 11.4 }));
    }
    this.rows.push(new LRow(`${c.block}/r/data/sp`, 'blank', '', { space: 6 }));
    this.rows.push(new LRow(`${c.block}/r/data/section`, 'section', 'Chart data: every number is a formula on the statements', { indent: 1 }));
    if (c.year) {
      this.yidx = this.indexRow('idx/year', 'Year shown: period of each month', j => `=«V|${this.y0}»+${j}`);
      this.mcats = this.indexRow('idx/year_months', 'Year shown: months', j => `=LEFT(INDEX(List_Month_Names,MOD(Sel_FY_End_Month+${j},12)+1),3)`, 'text');
    }
    if (c.charts.some(ch => ch.frame === 'rolling')) {
      this.ridx = this.indexRow('idx/rolling', '12 months to the month shown: period of each month', j => `=${c.month}-11+${j}`);
      this.rcats = this.indexRow('idx/rolling_months', '12 months to the month shown: months',
        j => `=IF(OR(«C${J + j}|${this.ridx}»<1,«C${J + j}|${this.ridx}»>Tl_Term),"",TEXT(INDEX(List_Months,«C${J + j}|${this.ridx}»),"mmm yy"))`, 'text');
    }
  }

  // --- ranking for top N ----------------------------------------------------------------------
  /** Rows below a table that rank a group by its total for the period (column K holds the line at each rank). */
  ranking(t: Table, g: ReturnType<Expander['group']>, a: string, b: string): void {
    const n = g.ids.length;
    const base = `${t.base}/rank`;
    const first = `${base}/1`;
    const last = `${base}/${n}`;
    t.rows.push(new LRow(`${base}/head`, 'table', 'Ranking for the period', { indent: 2, role: 'r.head',
      cells: { [TOTAL_COL]: 'Total', [J]: 'Rank key', [J + 1]: 'Line at rank' } }));
    g.ids.forEach((_, k) => {
      const id = `${base}/${k + 1}`;
      t.rows.push(new LRow(id, 'table', '', { indent: 2, role: 'r.rank', cells: {
        [LABEL_COLS[2]]: `=INDEX(${g.labels},${k + 1})`,
        [TOTAL_COL]: `=SUM(INDEX(${g.blk},${k + 1},${a}):INDEX(${g.blk},${k + 1},${b}))`,
        [J]: `=«C${TOTAL_COL}|${id}»-${k + 1}/1000000000`,
        [J + 1]: `=MATCH(LARGE(«K|${first}|${last}|${J}»,${k + 1}),«K|${first}|${last}|${J}»,0)`,
      } }));
    });
  }

  /** The group's lines shown: (label formula, line number expression), ranked or in model order. */
  members(t: Table, ch: ReportChartDef, g: ReturnType<Expander['group']>, rank: { at: string[] } | null): [string, string][] {
    const n = g.ids.length;
    if (ch.top && rank) {
      return rank.at.slice(0, Math.min(ch.top, n)).map(cell => [`=INDEX(${g.labels},${cell})`, cell]);
    }
    return g.ids.map((_, k) => [`=INDEX(${g.labels},${k + 1})`, String(k + 1)]);
  }

  push(chart: Omit<RChart, 'sheet' | 'grid' | 'slot'>): void {
    this.charts.push({ ...chart, sheet: '', grid: this.grid, slot: this.charts.length });
  }

  // --- recipes ----------------------------------------------------------------------------------
  compare(ch: ReportChartDef): Table {
    const frame = (ch.frame ?? 'year') as Frame;
    this.needs(frame);
    const l = this.line(ch, ch.line!);
    const pct = l.kind === 'ratio';
    const t = new Table(this, ch, `="${ch.title}, "&${this.windowLabel(frame)}`);
    t.header(this.cats(frame), 'Series', ch.cumulative ? '' : 'Total');
    const series: RSeries[] = [];
    for (const which of ch.periods ?? ['prior', 'shown', 'next']) {
      const shift = { prior: -12, shown: 0, next: 12 }[which];
      const vals = Array.from({ length: 12 }, (_, j) => {
        const i = this.idx(frame, j, shift);
        if (!ch.cumulative) return `=${pickAt(l.rng, i)}`;
        const start = j === 0 ? i : `(${i}-${j})`;
        return `=IF(OR(${start}<1,${i}>Tl_Term),NA(),${spanSum(l.rng, start, i)})`;
      });
      const id = t.add({ labelF: this.periodLabel(frame, which), unit: pct ? '%' : '$', values: vals, style: which === 'shown' ? 'bold' : '',
        total: !ch.cumulative && l.kind === 'flow' ? fullTotal() : undefined });
      const look = { prior: PRIOR_STYLE, shown: SHOWN_STYLE, next: NEXT_STYLE }[which];
      series.push(ch.kind === 'column'
        ? { row: id, as: 'bar', colour: look.colour! }
        : { row: id, as: 'line', ...look } as RSeries);
    }
    this.push({ id: ch.id, title: `${ch.title}, ${this.defaultLabel(frame)}`, titleRow: t.titleRow, cats: t.head, n: 12, type: 'bar',
      dir: 'col', grouping: 'clustered', gap: 60, overlap: -10, series, yFmt: pct ? '0%' : '#,##0', legend: 'b' });
    return t;
  }

  mix(ch: ReportChartDef): Table {
    const frame = (ch.frame ?? 'year') as Frame;
    this.needs(frame);
    const compare = ch.compare ?? true;
    const t = new Table(this, ch, `="${ch.title}, "&${this.windowLabel(frame)}`);
    t.header(this.cats(frame), 'Series', 'Total');
    const series: RSeries[] = [];
    const total = (v: string) => (frame === 'year' ? fullTotal() : `=SUM(${across(v, 12)})`);
    let colours = [...ACCENTS];
    let totalAt: (i: string) => string;
    const stack: string[] = [];
    const set = this.ctx.spec.sets[ch.group!];
    if (!set) {
      const g = this.group(ch, ch.group!);
      const n = g.ids.length;
      const [a, b] = this.bounds(frame);
      let rank: { at: string[] } | null = null;
      if (ch.top && n) {
        // the ranking sits below the table
        rank = { at: g.ids.map((_, k) => `«C${J + 1}|${t.base}/rank/${k + 1}»`) };
      }
      if (ch.lead) {
        const l = this.line(ch, ch.lead);
        const id = t.add({ label: l.label, unit: '$', values: Array.from({ length: 12 }, (_, j) => `=${pickAt(l.rng, this.idx(frame, j))}`), total: total('{self}') });
        stack.push(id);
        series.push({ row: id, as: 'bar', colour: 'accent2' });
        colours = ['accent1', 'accent3', 'accent4', 'accent5', 'accent6', ...ACCENTS];
      }
      const firstMember = t.nextId();
      let lastMember = firstMember;
      const shown = n ? this.members(t, ch, g, rank) : [];
      shown.forEach(([label, k], m) => {
        const id = t.add({ labelF: label, unit: '$', total: total('{self}'),
          values: Array.from({ length: 12 }, (_, j) => `=IF(OR(${this.idx(frame, j)}<1,${this.idx(frame, j)}>Tl_Term),NA(),INDEX(${g.blk},${k},${this.idx(frame, j)}))`) });
        stack.push(id);
        lastMember = id;
        series.push({ row: id, as: 'bar', colour: colours[m % colours.length] });
      });
      const groupAt = (i: string) => (n ? `IF(OR(${i}<1,${i}>Tl_Term),NA(),SUM(INDEX(${g.blk},0,${i})))` : `IF(OR(${i}<1,${i}>Tl_Term),NA(),0)`);
      if (ch.top && n > ch.top) {
        const id = t.add({ label: g.other, unit: '$', total: total('{self}'),
          values: Array.from({ length: 12 }, (_, j) => `=${groupAt(this.idx(frame, j))}-SUM(«C${J + j}|${firstMember}»:«C${J + j}|${lastMember}»)`) });
        stack.push(id);
        series.push({ row: id, as: 'bar', colour: 'light' });
      }
      if (ch.total) {
        const tl = this.line(ch, ch.total);
        totalAt = i => pickAt(tl.rng, i);
      } else {
        totalAt = groupAt;
      }
      if (ch.top && n) {
        this.ranking(t, g, a, b);
        if (frame === 'year') {
          const totalSpan = ch.total ? spanSum(this.line(ch, ch.total).rng, a, b)
            : `SUM(INDEX(${g.blk},0,${a}):INDEX(${g.blk},0,${b}))`;
          const shownTotals = `SUM(«C${TOTAL_COL}|${stack[0]}»:«C${TOTAL_COL}|${stack[stack.length - 1]}»)`;
          t.check(`${ch.id}: the lines shown do not add to the total`,
            `=IF(COUNT(${across(stack[0], 12)})<12,0,IF(ABS(${shownTotals}-${totalSpan})>0.001,1,0))`, 'error');
        }
      }
    } else {
      for (const key of set.stacks ?? []) {
        const s = this.signed(ch, key);
        const id = t.add({ label: s.label + (s.sign ? ' (shown below zero)' : ''), unit: '$',
          values: Array.from({ length: 12 }, (_, j) => `=${s.sign}${pickAt(s.rng, this.idx(frame, j))}`) });
        stack.push(id);
        series.push({ row: id, as: 'bar', colour: colours[(stack.length - 1) % colours.length] });
      }
      const tl = this.line(ch, set.total!);
      totalAt = i => pickAt(tl.rng, i);
      if (set.current) {
        const cl = this.line(ch, set.current);
        const id = t.add({ label: cl.label, unit: '$', style: 'bold', values: Array.from({ length: 12 }, (_, j) => `=${pickAt(cl.rng, this.idx(frame, j))}`) });
        series.push({ row: id, as: 'line', ...SHOWN_STYLE } as RSeries);
      }
    }
    if (compare) {
      for (const which of ['prior', 'next'] as const) {
        const shift = which === 'prior' ? -12 : 12;
        const id = t.add({ labelF: this.periodLabel(frame, which), unit: '$',
          values: Array.from({ length: 12 }, (_, j) => `=${totalAt(this.idx(frame, j, shift))}`) });
        series.push({ row: id, as: 'line', ...(which === 'prior' ? { colour: 'grey', width: 1.5 } : { colour: 'tx2', width: 1.5, dash: true }) } as RSeries);
      }
    }
    this.push({ id: ch.id, title: `${ch.title}, ${this.defaultLabel(frame)}`, titleRow: t.titleRow, cats: t.head, n: 12, type: 'bar',
      dir: 'col', grouping: 'stacked', gap: 55, overlap: 100, series, yFmt: '#,##0', legend: 'b' });
    return t;
  }

  depth(ch: ReportChartDef): Table {
    const frame = (ch.frame ?? 'year') as Frame;
    this.needs(frame);
    const g = this.group(ch, ch.group!);
    const n = g.ids.length;
    const [a, b] = this.bounds(frame);
    const t = new Table(this, ch, `="${ch.title}, "&${this.windowLabel(frame)}`);
    t.header([`=${this.windowLabel(frame)}`], 'Line');
    const rank = ch.top && n ? { at: g.ids.map((_, k) => `«C${J + 1}|${t.base}/rank/${k + 1}»`) } : null;
    const series: RSeries[] = [];
    const first = t.nextId();
    let last = first;
    this.members(t, ch, g, rank).forEach(([label, k], m) => {
      const v = rank ? `=INDEX(«K|${t.base}/rank/1|${t.base}/rank/${n}|${TOTAL_COL}»,${k})` : `=SUM(INDEX(${g.blk},${k},${a}):INDEX(${g.blk},${k},${b}))`;
      last = t.add({ labelF: label, unit: '$', values: [v] });
      series.push({ row: last, as: 'bar', colour: ACCENTS[m % ACCENTS.length], labels: { fmt: '#,##0', pos: 'outEnd' } });
    });
    if (ch.top && n > ch.top) {
      const id = t.add({ label: g.other, unit: '$', values: [`=SUM(INDEX(${g.blk},0,${a}):INDEX(${g.blk},0,${b}))-SUM(«C${J}|${first}»:«C${J}|${last}»)`] });
      series.push({ row: id, as: 'bar', colour: 'light', labels: { fmt: '#,##0', pos: 'outEnd' } });
      t.check(`${ch.id}: Other is below zero, so the ranking is wrong`, `=IF(«C${J}|${id}»<-0.001,1,0)`, 'error');
    }
    if (rank) this.ranking(t, g, a, b);
    this.push({ id: ch.id, title: `${ch.title}, ${this.defaultLabel(frame)}`, titleRow: t.titleRow, cats: t.head, n: 1, type: 'bar',
      dir: 'col', grouping: 'clustered', gap: 40, overlap: -5, series, yFmt: '#,##0', legend: 'b' });
    return t;
  }

  pie(ch: ReportChartDef): Table {
    const frame = (ch.frame ?? 'year') as Frame;
    this.needs(frame);
    const t = new Table(this, ch, `="${ch.title}, "&${this.windowLabel(frame)}`);
    const set = this.ctx.spec.sets[ch.group!];
    let values: string;
    let colours: string[];
    let n: number;
    if (!set) {
      const g = this.group(ch, ch.group!);
      const total = g.ids.length;
      const [a, b] = this.bounds(frame);
      const rank = total ? { at: g.ids.map((_, k) => `«C${J + 1}|${t.base}/rank/${k + 1}»`) } : null;
      const mem = total ? this.members(t, { ...ch, top: ch.top ?? total }, g, rank) : [];
      const other = Boolean(ch.top && total > ch.top);
      const cats = [...mem.map(([lab]) => lab), ...(other ? [`="${g.other}"`] : [])];
      t.header(cats, 'Line');
      n = cats.length;
      const vals = mem.map(([, k]) => `=INDEX(«K|${t.base}/rank/1|${t.base}/rank/${total}|${TOTAL_COL}»,${k})`);
      const row = t.nextId();
      if (other) vals.push(`=SUM(INDEX(${g.blk},0,${a}):INDEX(${g.blk},0,${b}))-SUM(«C${J}|${row}»:«C${J + mem.length - 1}|${row}»)`);
      values = t.add({ label: 'Total for the period', unit: '$', values: vals, total: `=SUM(${across('{self}', vals.length)})` });
      colours = [...mem.map((_, m) => ACCENTS[m % ACCENTS.length]), ...(other ? ['light'] : [])];
      if (total) this.ranking(t, g, a, b);
    } else if (frame === 'at') {
      const items = set.items!.map(k => this.signed(ch, k));
      t.header(items.map(i => i.label), 'Line');
      n = items.length;
      values = t.add({ label: 'Balance at the month shown', unit: '$', values: items.map(i => `=${i.sign}${pickAt(i.rng, this.ctx.month!)}`) });
      colours = items.map((_, m) => (m < 5 ? ACCENTS[m] : m === 5 ? 'grey' : 'tx2'));
    } else {
      const items = set.items!.map(k => this.signed(ch, k));
      const [a, b] = this.bounds(frame);
      const net = t.nextId();
      n = items.length;
      t.header(items.map((i, j) => `="${i.label}"&IF(«C${J + j}|${net}»<0," (out)"," (in)")`), 'Flow');
      t.add({ label: 'Net flow for the period', unit: '$', values: items.map(i => `=${i.sign}${spanSum(i.rng, a, b)}`) });
      values = t.add({ label: 'Size (the pie shows sizes)', unit: '$', values: items.map((_, j) => `=ABS(«C${J + j}|${net}»)`) });
      colours = items.map((_, m) => ACCENTS[m % ACCENTS.length]);
    }
    this.push({ id: ch.id, title: `${ch.title}, ${this.defaultLabel(frame)}`, titleRow: t.titleRow, cats: t.head, n, type: 'pie',
      series: [{ row: values, as: 'bar', colour: 'accent1', points: colours, labels: { fmt: '0%', pos: 'outEnd', pct: true } }], legend: 'r' });
    return t;
  }

  combo(ch: ReportChartDef): Table {
    const frame = (ch.frame ?? 'year') as Frame;
    this.needs(frame);
    const t = new Table(this, ch, `="${ch.title}, "&${this.windowLabel(frame)}`);
    t.header(this.cats(frame), 'Series', 'Total');
    const series: RSeries[] = [];
    const row = (key: string, bold: boolean) => {
      const l = this.line(ch, key);
      return t.add({ label: l.label, unit: '$', style: bold ? 'bold' : '', total: l.kind === 'flow' ? fullTotal() : undefined,
        values: Array.from({ length: 12 }, (_, j) => `=${pickAt(l.rng, this.idx(frame, j))}`) });
    };
    (ch.bars ?? []).forEach((k, m) => series.push({ row: row(k, false), as: 'bar', colour: ACCENTS[m % ACCENTS.length] }));
    (ch.lines ?? []).forEach((k, m) => series.push({ row: row(k, true), as: 'line', colour: m ? 'grey' : 'tx2', width: 2.25, marker: true }));
    const stacked = ch.grouping === 'stacked';
    this.push({ id: ch.id, title: `${ch.title}, ${this.defaultLabel(frame)}`, titleRow: t.titleRow, cats: t.head, n: 12, type: 'bar',
      dir: 'col', grouping: stacked ? 'stacked' : 'clustered', gap: stacked ? 55 : 60, overlap: stacked ? 100 : -10, series, yFmt: '#,##0', legend: 'b' });
    return t;
  }

  movement(ch: ReportChartDef): Table {
    this.needs('at');
    const set = this.ctx.spec.sets[ch.group!];
    if (!set?.items) throw new AssemblyError(`${this.where(ch)}: no set of items '${ch.group}'`);
    const items = set.items.map(k => this.signed(ch, k));
    const M = this.ctx.month!;
    const t = new Table(this, ch, `="${ch.title}, "&TEXT(INDEX(List_Months,${M}),"mmm yyyy")&" against a year earlier"`);
    t.header(items.map(i => i.label), 'Line');
    const n = items.length;
    const r1 = t.add({ label: 'Month shown', unit: '$', values: items.map(i => `=${i.sign}${pickAt(i.rng, M)}`) });
    const r2 = t.add({ label: 'A year earlier', unit: '$', values: items.map(i => `=${i.sign}${pickAt(i.rng, `(${M}-12)`)}`) });
    const r3 = t.add({ label: 'Change', unit: '$', style: 'signed', values: items.map((_, j) => `=«C${J + j}|${r1}»-«C${J + j}|${r2}»`) });
    const ri = t.add({ label: 'Increase', unit: '$', values: items.map((_, j) => `=IF(«C${J + j}|${r3}»>=0,«C${J + j}|${r3}»,0)`) });
    const rd = t.add({ label: 'Decrease', unit: '$', values: items.map((_, j) => `=IF(«C${J + j}|${r3}»<0,«C${J + j}|${r3}»,0)`) });
    this.push({ id: ch.id, title: `${ch.title}, ${this.defaultLabel('movement')}`, titleRow: t.titleRow, cats: t.head, n, type: 'bar',
      dir: 'bar', grouping: 'clustered', gap: 40, overlap: 100, valueAxis: false, reverse: true, legend: null,
      series: [{ row: ri, as: 'bar', colour: 'accent1', labels: { fmt: '+#,##0;-#,##0;;', pos: 'outEnd' } },
        { row: rd, as: 'bar', colour: 'accent3', labels: { fmt: '+#,##0;-#,##0;;', pos: 'outEnd' } }] });
    return t;
  }

  bridge(ch: ReportChartDef): Table {
    const set = this.ctx.spec.sets[ch.group!];
    if (!set?.steps || !set.end) throw new AssemblyError(`${this.where(ch)}: no bridge '${ch.group}'`);
    const frame = (ch.frame ?? 'year') as Frame;
    this.needs(frame);
    const end = this.line(ch, set.end);
    const items: { label: string; value: string | null; kind: 'start' | 'step' | 'end' }[] = [];
    let endRef: string;
    if (frame === 'at') {
      const M = this.ctx.month!;
      for (const k of set.steps) {
        const s = this.signed(ch, k);
        items.push({ label: s.label, value: `=${s.sign}${pickAt(s.rng, M)}`, kind: 'step' });
      }
      endRef = pickAt(end.rng, M);
    } else {
      const [a, b] = this.bounds(frame);
      if (set.start) {
        const st = this.line(ch, set.start);
        items.push({ label: set.startLabel ?? st.label, value: `=INDEX(${st.rng},${a})`, kind: 'start' });
      }
      for (const k of set.steps) {
        const s = this.signed(ch, k);
        items.push({ label: s.label, value: `=${s.sign}${spanSum(s.rng, a, b)}`, kind: 'step' });
      }
      endRef = `INDEX(${end.rng},${b})`;
    }
    items.push({ label: set.endLabel ?? end.label, value: null, kind: 'end' });
    const t = new Table(this, ch, `="${ch.title}, "&${this.windowLabel(frame)}`);
    t.header(items.map(i => i.label), 'Bridge');
    const vr = `${t.base}/${1}`;
    const rr = `${t.base}/${2}`;
    const c = (row: string, j: number) => `«C${J + j}|${row}»`;
    const value: string[] = [];
    const run: string[] = [];
    const base: string[] = [];
    const tot: string[] = [];
    const inc: string[] = [];
    const dec: string[] = [];
    items.forEach((it, j) => {
      const prev = j ? c(rr, j - 1) : '0';
      value.push(it.kind === 'end' ? `=${c(rr, j - 1)}` : it.value!);
      if (it.kind === 'step') {
        run.push(`=${prev}+${c(vr, j)}`);
        base.push(`=MIN(${prev},${c(rr, j)})`);
        tot.push('=0');
        inc.push(`=MAX(${c(vr, j)},0)`);
        dec.push(`=MAX(-${c(vr, j)},0)`);
      } else {
        run.push(`=${c(vr, j)}`);
        base.push('=0');
        tot.push(`=${c(vr, j)}`);
        inc.push('=0');
        dec.push('=0');
      }
    });
    t.add({ label: 'Value', unit: '$', values: value, style: 'bold' });
    t.add({ label: 'Running total', unit: '$', values: run });
    const br = t.add({ label: 'Base (not drawn)', unit: '$', values: base, role: 'r.muted' });
    const tr = t.add({ label: 'Total', unit: '$', values: tot });
    const ir = t.add({ label: 'Increase', unit: '$', values: inc });
    const dr = t.add({ label: 'Decrease', unit: '$', values: dec });
    const n = items.length;
    t.check(`${ch.id}: the bridge does not reach the statement figure`, `=IF(ABS(${c(vr, n - 1)}-${endRef})>0.001,1,0)`, 'error');
    t.check(`${ch.id}: the running total crosses zero, so a step is drawn from the wrong base`, `=IF(MIN(${across(rr, n)})<0,1,0)`, 'alert');
    const dir = frame === 'at' ? 'col' : 'bar';
    this.push({ id: ch.id, title: `${ch.title}, ${this.defaultLabel(frame)}`, titleRow: t.titleRow, cats: t.head, n, type: 'bar',
      dir, grouping: 'stacked', gap: 35, overlap: 100, valueAxis: false, reverse: dir === 'bar', legend: null,
      series: [
        { row: br, as: 'bar', colour: 'none' },
        { row: tr, as: 'bar', colour: 'grey', outline: true, labels: { fmt: '#,##0;-#,##0;', pos: 'inEnd' } },
        { row: ir, as: 'bar', colour: 'accent1', outline: true, labels: { fmt: '"+"#,##0;;', pos: 'ctr' } },
        { row: dr, as: 'bar', colour: 'accent3', outline: true, labels: { fmt: '"-"#,##0;;', pos: 'ctr' } },
      ] });
    return t;
  }

  budget(ch: ReportChartDef): Table {
    this.needs('year');
    if (!this.cmpRow) throw new AssemblyError(`${this.where(ch)}: a budget chart needs a Compared with setting`);
    const l = this.line(ch, ch.line!);
    const lead = ch.title.replace(/ against budget$/, '');
    const yl = `INDEX(List_Years,${this.ctx.year})`;
    const cl = `«V|${this.cmpLabel}»`;
    const t = new Table(this, ch, `="${lead} against "&${cl}&IF(ISNUMBER(SEARCH(${yl},${cl})),"",", "&${yl})`);
    t.header(this.cats('year'), 'Series', 'Total');
    const m = `${t.base}/store_row`;
    t.rows.push(new LRow(m, 'table', 'Its row in the Version store', { indent: 2, unit: '#', role: 'r.scalar',
      cells: { [TOTAL_COL]: `=IFERROR(MATCH(${this.n('Cmp_Id')}&"|${ch.line}",VS_Keys,0),0)` } }));
    const out = (i: string) => `OR(${i}<1,${i}>Tl_Term)`;
    const vals = (f: (i: string) => string) => Array.from({ length: 12 }, (_, j) => {
      const i = this.idx('year', j);
      return `=IF(${out(i)},NA(),${f(i)})`;
    });
    const ra = t.add({ label: 'Actual', unit: '$', total: fullTotal(), values: vals(i => `IF(${i}<=Tl_Last_Actual,INDEX(${l.rng},${i}),0)`) });
    const rf = t.add({ label: 'Forecast', unit: '$', total: fullTotal(), values: vals(i => `IF(${i}<=Tl_Last_Actual,0,INDEX(${l.rng},${i}))`) });
    const stored = (i: string) => `INDEX(VS_Values,«V|${m}»,${i})`;
    const rc = t.add({ labelF: `=${cl}`, unit: '$', style: 'bold', total: fullTotal(),
      values: vals(i => `IF(«V|${this.cmpRow}»=0,INDEX(${l.rng},${i}),IF(«V|${m}»=0,NA(),IF(${stored(i)}="",NA(),${stored(i)})))`) });
    t.add({ label: 'Actual and forecast less the comparison', unit: '$', style: 'signed', total: fullTotal(),
      values: Array.from({ length: 12 }, (_, j) => `=«C${J + j}|${ra}»+«C${J + j}|${rf}»-«C${J + j}|${rc}»`) });
    this.push({ id: ch.id, title: `${lead} against the approved budget, ${this.defaultLabel('year')}`, titleRow: t.titleRow, cats: t.head,
      n: 12, type: 'bar', dir: 'col', grouping: 'stacked', gap: 55, overlap: 100, yFmt: '#,##0', legend: 'b',
      series: [{ row: ra, as: 'bar', colour: 'tx2' }, { row: rf, as: 'bar', colour: 'tx2', hatch: true },
        { row: rc, as: 'line', colour: 'accent1', width: 2.25, marker: true }] });
    return t;
  }

  scenario(ch: ReportChartDef): Table {
    const c = this.ctx;
    if (!c.scenarios) throw new AssemblyError(`${this.where(ch)}: the model has no scenarios`);
    const l = this.line(ch, ch.line!);
    const by = ch.by ?? 'year';
    if (by === 'month') this.needs('year');
    const t = new Table(this, ch, by === 'month' ? `="${ch.title}, "&${this.windowLabel('year')}` : `="${ch.title}, each year"`);
    const slots = by === 'month' ? 12 : c.years;
    // One data table row per slot: the line for that month of the year shown, or that financial year.
    const res: string[] = [];
    for (let j = 0; j < slots; j++) {
      const id = `scnres/${c.block}/${ch.id}/${j + 1}`;
      let f: string;
      let label: string;
      if (by === 'month') {
        f = `=IF(OR(${this.n('Y0')}+${j}<1,${this.n('Y0')}+${j}>Tl_Term),0,INDEX(${l.rng},${this.n('Y0')}+${j}))`;
        label = `${c.title} ${ch.id}: ${l.label}, month ${j + 1} of the year shown`;
      } else {
        const a = `MAX(1,${j * 12}-Sel_Start_Month+2)`;
        const b = `MIN(Tl_Term,${(j + 1) * 12}-Sel_Start_Month+1)`;
        f = l.kind === 'flow' ? `=${spanSum(l.rng, a, b)}` : `=INDEX(${l.rng},${b})`;
        label = `${c.title} ${ch.id}: ${l.label}, ${c.fyLabel(j + 1)}`;
      }
      c.results.push({ id, label, formula: f });
      res.push(id);
    }
    if (by === 'month') {
      t.header(this.cats('year'), 'Scenario', l.kind === 'flow' ? 'Total' : '');
    } else {
      t.header(Array.from({ length: slots }, (_, y) => `=INDEX(List_Years,${y + 1})`), 'Scenario');
    }
    const looks: Partial<RSeries>[] = [{ colour: 'tx2', width: 2.25 }, { colour: 'accent1', width: 1.75 }, { colour: 'accent3', width: 1.75, dash: true }];
    const series: RSeries[] = [];
    c.scenarios.forEach((_, s) => {
      const vals = res.map((rid, j) => {
        const v = `INDEX(«C${J}|${rid}»:«C${J + c.scenarios!.length - 1}|${rid}»,${s + 1})`;
        return by === 'month' ? `=IF(OR(${this.idx('year', j)}<1,${this.idx('year', j)}>Tl_Term),NA(),${v})` : `=${v}`;
      });
      const id = t.add({ labelF: `=INDEX(List_Scenarios,${s + 1})`, unit: '$', values: vals, style: s === 0 ? 'bold' : '',
        total: by === 'month' && l.kind === 'flow' ? fullTotal() : undefined });
      series.push({ row: id, as: 'line', ...looks[s % looks.length], marker: by === 'year' } as RSeries);
    });
    if (by === 'year') t.add({ label: 'Flows are totals for the year; balances are at the end of the year.', role: 'r.note', values: [] });
    this.push({ id: ch.id, title: `${ch.title}, ${this.defaultLabel(by === 'month' ? 'year' : 'each')}`, titleRow: t.titleRow, cats: t.head,
      n: slots, type: 'bar', dir: 'col', grouping: 'clustered', series, yFmt: '#,##0', legend: 'b' });
    return t;
  }
}

/** Expand a report module's charts into its rows and charts. */
export function expandReport(ctx: ReportContext): ReportOut {
  const e = new Expander(ctx);
  e.selections();
  for (const ch of ctx.charts) {
    const fn = (e as unknown as Record<string, (c: ReportChartDef) => Table>)[ch.recipe];
    if (typeof fn !== 'function' || !['compare', 'mix', 'depth', 'pie', 'combo', 'movement', 'bridge', 'scenario', 'budget'].includes(ch.recipe)) {
      throw new AssemblyError(`${ctx.title} chart ${ch.id}: no recipe '${ch.recipe}'`);
    }
    const t = fn.call(e, ch);
    e.rows.push(...t.rows);
  }
  return { rows: e.rows, charts: e.charts, errors: e.errors, alerts: e.alerts, names: e.names };
}

export const REPORT_UNIT_COL = UNIT_COL;
export { colLetter };
