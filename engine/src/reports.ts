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
import { COST_LINES, VERSION_LINES } from './versions.ts';

const J = FIRST_PERIOD_COL;
/** Charts per row of the grid, and sheet rows per row of charts. */
export const GRID = { cols: 3, rows: 20 } as const;
/** Report charts: 12.2 cm by 7.6 cm, as in the chart register proof. */
export const RCHART_CM = { w: 12.2, h: 7.6 } as const;

type Frame = 'year' | 'rolling' | 'at' | 'ytd' | 'budget';

/** A scenario result: one output formula the Scenarios sheet's data table works out for every scenario. */
export interface ScenarioResult {
  id: string;
  label: string;
  /** What it measures, the heading it sits under on the Scenarios sheet ("Revenue, each financial year"). */
  group: string;
  /** The period it covers, the row's label under that heading (a formula: the year, or the month's name). */
  item: string;
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
  /** The months a dashboard table shows (List_Table_Months): as the charts, the 12 months to or after the last actual month, or 12 months from any month. */
  table?: string | null;
  /** A second comparison and the line shown (the version comparison module). */
  compare2?: string | null;
  line?: string | null;
  /** Months in the timeline. */
  periods: number;
  /** The budget window New model set (for chart title caches). */
  budget?: { first: number; months: number };
  /** A target the statement recipe shows the gap to (the Budget module's target profit), as a setting name. */
  target?: string | null;
  /** Labels as the default selections show them, for chart title caches. */
  fyLabel: (k: number) => string;
  monthLabel: (p: number) => string;
  /** The financial year (1 for the first) a period falls in. */
  fyOf: (p: number) => number;
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

  constructor(rep: Expander, ch: ReportChartDef, titleFormula: string, dashboard = false) {
    this.rep = rep;
    this.ch = ch;
    this.base = `${rep.ctx.block}/r/${ch.id}`;
    this.rows.push(new LRow(`${this.base}/sp`, 'blank', '', { space: dashboard ? 9 : 6 }));
    // a dashboard table is headed by its title alone; a chart's table by its register id, and the title its chart shows
    this.rows.push(new LRow(`${this.base}/section`, 'section', dashboard ? ch.title : `${ch.id}  ${ch.title}`, { indent: 1 }));
    this.titleRow = `${this.base}/title`;
    if (!dashboard) this.rows.push(new LRow(this.titleRow, 'table', 'Chart title', { indent: 2, role: 'r.title', cells: { [TOTAL_COL]: titleFormula } }));
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
  /** Each comparison: its register row, version id and label scalars (names). */
  readonly cmps: { row: string; id: string; label: string }[] = [];
  /** The year position the year frame reads: the year shown, or the financial year of the month shown. */
  year: string | null;
  private lineRow = '';
  private bidx = '';
  private bcats = '';
  /** Index rows, made with the selections (so tables above the grid can read them) and placed under Chart data. */
  private readonly pendingIndex: LRow[] = [];
  private yidx = '';
  private ridx = '';
  private mcats = '';
  private rcats = '';
  /** A dashboard table's twelve months: the first as a period, the period and label of each, and its description. */
  private w0 = '';
  private widx = '';
  private wcats = '';
  private wlabel = '';
  /** The year to date: its first month (the financial year of the last actual month). */
  private ya = '';
  readonly grid: string;

  constructor(ctx: ReportContext) {
    this.ctx = ctx;
    this.grid = `${ctx.block}/r/grid`;
    this.year = ctx.year;
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
    this.pendingIndex.push(new LRow(id, 'table', label, { indent: 2, role: 'r.index', unit, cells }));
    return id;
  }

  /** The window's label for chart titles. */
  windowLabel(frame: Frame): string {
    const M = this.ctx.month;
    const ml = `TEXT(INDEX(List_Months,${M}),"mmm yyyy")`;
    const bl = (p: string) => `TEXT(INDEX(List_Months,${p}),"mmm yyyy")`;
    return {
      year: `INDEX(List_Years,${this.year})`, rolling: `"12 months to "&${ml}`, at: `"at "&${ml}`, ytd: `"year to "&${ml}`,
      budget: `"budget "&${bl('Sel_Budget_First')}&" to "&${bl('MIN(Tl_Term,Sel_Budget_First+Tl_Budget_Term-1)')}`,
    }[frame];
  }

  defaultLabel(frame: Frame | 'movement' | 'each'): string {
    const c = this.ctx;
    const ml = c.monthLabel(c.monthShown);
    return {
      year: c.fyLabel(c.year ? c.yearShown : c.fyOf(c.monthShown)), rolling: `12 months to ${ml}`, at: `at ${ml}`, ytd: `year to ${ml}`,
      movement: `${ml} against a year earlier`, each: 'each year',
      budget: `budget ${c.monthLabel(c.budget?.first ?? 1)} to ${c.monthLabel((c.budget?.first ?? 1) + (c.budget?.months ?? 12) - 1)}`,
    }[frame];
  }

  periodLabel(frame: Frame, which: 'prior' | 'shown' | 'next'): string {
    if (which === 'shown') return `=${this.windowLabel(frame)}`;
    if (frame === 'year') return `="FY"&(Tl_First_FY+${this.year}${which === 'prior' ? '-2' : ''})`;
    return which === 'prior' ? '="Prior 12 months"' : '="Next 12 months"';
  }

  /** Period index of column j for a frame, shifted by months. */
  idx(frame: Frame, j: number, shift = 0): string {
    const row = frame === 'year' ? this.yidx : frame === 'budget' ? this.bidx : this.ridx;
    const base = `«C${J + j}|${row}»`;
    return shift ? `(${base}${shift > 0 ? '+' : ''}${shift})` : base;
  }

  cats(frame: Frame): string[] {
    const row = frame === 'year' ? this.mcats : frame === 'budget' ? this.bcats : this.rcats;
    return Array.from({ length: 12 }, (_, j) => `=«C${J + j}|${row}»`);
  }

  bounds(frame: Frame): [string, string] {
    const M = this.ctx.month!;
    if (frame === 'year') return [`«V|${this.ys}»`, `«V|${this.ye}»`];
    if (frame === 'rolling') return [`MAX(1,${M}-11)`, M];
    if (frame === 'ytd') return [`«V|${this.ytd0}»`, M];
    if (frame === 'budget') return ['Sel_Budget_First', 'MIN(Tl_Term,Sel_Budget_First+Tl_Budget_Term-1)'];
    return [M, M];
  }

  needs(frame: Frame | undefined): void {
    if ((frame ?? 'year') === 'year' && !this.year) throw new AssemblyError(`${this.ctx.title}: a chart reads the year shown, but the module has no year setting`);
    if (frame && frame !== 'year' && frame !== 'budget' && !this.ctx.month) throw new AssemblyError(`${this.ctx.title}: a chart reads the month shown, but the module has no month setting`);
  }

  selections(): void {
    const c = this.ctx;
    this.rows.push(new LRow(`${c.block}/r/sel/sp`, 'blank', '', { space: 6 }));
    const tables = c.charts.some(ch => TABLES.includes(ch.recipe));
    this.rows.push(new LRow(`${c.block}/r/sel/section`, 'section', tables ? 'What the tables and charts show' : 'What the charts show', { indent: 1 }));
    if (c.scenarios) {
      this.scalar('sel/scenario', 'Scenario shown (change it on the Scenarios sheet)', '=Scn_Active_Name', 'text');
    }
    if (!c.year && c.month && (c.compare || c.line || c.charts.some(ch => (ch.frame ?? 'year') === 'year' && !['movement', 'pie'].includes(ch.recipe) || ch.recipe === 'budget'))) {
      // modules that show a month read the financial year it falls in
      this.year = this.n('FY');
      this.scalar('sel/fy', 'Financial year of the month shown', `=INT((${c.month}+Sel_Start_Month-2)/12)+1`, '#', this.year);
    }
    if (this.year) {
      const y = this.year;
      this.scalar('sel/year_label', c.year ? 'Year shown' : 'Year of the month shown', `=INDEX(List_Years,${y})`, 'text');
      // Month 1 of the year shown as a period number (below 1 when the first year is part of a year).
      this.y0 = this.scalar('sel/y0', 'Year shown: its first month as a period', `=(${y}-1)*12-Sel_Start_Month+2`, 'period', this.n('Y0'));
      this.ys = this.scalar('sel/ys', 'Year shown: first month in the timeline', `=MAX(1,«V|${this.y0}»)`, 'period', this.n('Y_Start'));
      this.ye = this.scalar('sel/ye', 'Year shown: last month in the timeline', `=MIN(Tl_Term,«V|${this.y0}»+11)`, 'period', this.n('Y_End'));
    }
    if (c.month) {
      this.scalar('sel/month_label', 'Month shown', `=TEXT(INDEX(List_Months,${c.month}),"mmm yyyy")`, 'text');
      this.ytd0 = this.scalar('sel/ytd0', 'Month shown: first month of its financial year in the timeline',
        `=MAX(1,INT((${c.month}+Sel_Start_Month-2)/12)*12-Sel_Start_Month+2)`, 'period', this.n('YTD_Start'));
    }
    if (c.line) {
      this.scalar('sel/line_label', 'Line shown', `=INDEX(List_Version_Lines,${c.line})`, 'text', this.n('Line_Label'));
      this.scalar('sel/line_key', 'Line shown: its key', `=INDEX(List_Version_Keys,${c.line})`, 'text', this.n('Line_Key'));
    }
    for (const [k, sel] of [[1, c.compare], [2, c.compare2]] as const) {
      if (!sel) continue;
      if (!this.year) throw new AssemblyError(`${c.title}: Compared with needs a year shown`);
      const sfx = k === 1 ? 'Cmp' : 'Cmp2';
      const what = k === 1 ? 'Compared with' : 'Second comparison';
      // The register row of the version compared with: 0 for the budget being built, -1 when nothing is saved for it.
      const row = this.scalar(`sel/${sfx}_row`.toLowerCase(), `${what}: the version's row in the register (0: the budget being built)`,
        `=IFERROR(CHOOSE(MIN(${sel},5),MATCH("Budget|Approved|"&INDEX(List_Years,${this.year}),VR_BudgetKey,0),`
        + `MATCH("Reforecast|"&(Tl_Last_Actual-1),VR_RefKey,0),MATCH("Reforecast|"&MAX(VR_RefAt),VR_RefKey,0),0,${sel}-4),-1)`,
        '#', this.n(`${sfx}_Row`));
      this.scalar(`sel/${sfx}_id`.toLowerCase(), `${what}: version id`, `=IF(«V|${row}»>0,INDEX(VR_Id,«V|${row}»),"")`, 'text', this.n(`${sfx}_Id`));
      const label = this.scalar(`sel/${sfx}_label`.toLowerCase(), what,
        `=IF(«V|${row}»=0,"Budget being built",IF(«V|${row}»<0,"Nothing saved",INDEX(VR_Label,«V|${row}»)))`, 'text', this.n(`${sfx}_Label`));
      this.cmps.push({ row, id: this.n(`${sfx}_Id`), label });
      const id = `${c.block}/r/sel/${sfx.toLowerCase()}_check`;
      this.rows.push(new LRow(id, 'table', `Nothing is saved for the ${k === 1 ? 'comparison' : 'second comparison'} chosen, so it is blank`,
        { indent: 2, unit: 'flag', role: 'r.check', cells: { [TOTAL_COL]: `=IF(«V|${row}»<0,1,0)` } }));
      this.alerts.push(`«V|${id}»`);
    }
    if (this.cmps.length) {
      this.cmpRow = this.cmps[0].row;
      this.cmpLabel = this.cmps[0].label;
    }
    if (c.line) {
      // the line shown, every month of the timeline, so charts can read it with INDEX
      const keys = VERSION_LINES.filter(k => c.spec.lines[k]);
      const cells: Record<number, unknown> = {};
      for (let t = 0; t < c.periods; t++) {
        cells[J + t] = `=CHOOSE(${c.line},${keys.map(k => `«R|${c.fs}/${c.spec.lines[k].row}»`).join(',')})`;
      }
      this.lineRow = `${c.block}/r/idx/line`;
      this.pendingIndex.push(new LRow(this.lineRow, 'table', 'Line shown, every month', { indent: 2, role: 'r.index', cells }));
    }
    if (c.charts.some(ch => ch.recipe === 'summary' || ch.recipe === 'quarters')) {
      // A dashboard table's twelve months: as the charts show them, or the table's own choice.
      const start = this.year ? `«V|${this.y0}»` : c.month ? `(${c.month}-11)` : '(Tl_Last_Actual-11)';
      const f = c.table ? `=CHOOSE(MIN(${c.table},4),${start},Tl_Last_Actual-11,Tl_Last_Actual+1,${c.table}-3)` : `=${start}`;
      this.w0 = this.scalar('sel/w0', 'Table: its first month as a period', f, 'period', this.n('W0'));
      const W = `«V|${this.w0}»`;
      const end = `EOMONTH(Tl_Start,${W}+10)`;
      this.wlabel = this.scalar('sel/w_label', 'Table shows', `=IF(MOD(${W}+Sel_Start_Month-2,12)=0,"FY"&(YEAR(${end})+IF(MONTH(${end})>Sel_FY_End_Month,1,0)),`
        + `"12 months to "&TEXT(${end},"mmmm yyyy"))`, 'text', this.n('W_Label'));
      if (c.charts.some(ch => ch.against?.includes('ytd'))) {
        this.ya = this.scalar('sel/ytd_a', 'Year to date: its first month (the financial year of the last actual month)',
          '=IF(Tl_Last_Actual<1,0,MAX(1,INT((Tl_Last_Actual+Sel_Start_Month-2)/12)*12-Sel_Start_Month+2))', 'period', this.n('YTD_A'));
      }
    }
    if (this.w0) {
      this.widx = this.indexRow('idx/window', 'Table: period of each month', j => `=«V|${this.w0}»+${j}`);
      const p = (j: number) => `«C${J + j}|${this.widx}»`;
      this.wcats = this.indexRow('idx/window_months', 'Table: months',
        j => `=IF(OR(${p(j)}<1,${p(j)}>Tl_Term),"",TEXT(INDEX(List_Months,${p(j)}),"mmm yy"))`, 'text');
    }
    if (this.year) {
      this.yidx = this.indexRow('idx/year', 'Year shown: period of each month', j => `=«V|${this.y0}»+${j}`);
      this.mcats = this.indexRow('idx/year_months', 'Year shown: months', j => `=LEFT(INDEX(List_Month_Names,MOD(Sel_FY_End_Month+${j},12)+1),3)`, 'text');
    }
    if (c.charts.some(ch => ch.frame === 'budget')) {
      // the budget window: its months, blank past its length
      this.bidx = this.indexRow('idx/budget', 'Budget: period of each month', j => `=IF(${j}<Tl_Budget_Term,Sel_Budget_First+${j},0)`);
      this.bcats = this.indexRow('idx/budget_months', 'Budget: months',
        j => `=IF(«C${J + j}|${this.bidx}»<1,"",TEXT(INDEX(List_Months,«C${J + j}|${this.bidx}»),"mmm yy"))`, 'text');
    }
    if (c.charts.some(ch => ch.frame === 'rolling')) {
      this.ridx = this.indexRow('idx/rolling', '12 months to the month shown: period of each month', j => `=${c.month}-11+${j}`);
      this.rcats = this.indexRow('idx/rolling_months', '12 months to the month shown: months',
        j => `=IF(OR(«C${J + j}|${this.ridx}»<1,«C${J + j}|${this.ridx}»>Tl_Term),"",TEXT(INDEX(List_Months,«C${J + j}|${this.ridx}»),"mmm yy"))`, 'text');
    }
  }

  /** The chart grid, then the chart data's index rows. */
  gridRows(): void {
    const c = this.ctx;
    this.rows.push(new LRow(`${c.block}/r/grid/sp`, 'blank', '', { space: 6 }));
    this.rows.push(new LRow(`${c.block}/r/grid/section`, 'section', 'Charts', { indent: 1 }));
    const gridRows = Math.ceil(c.charts.filter(ch => !TABLES.includes(ch.recipe)).length / GRID.cols) * GRID.rows;
    for (let k = 0; k < gridRows; k++) {
      this.rows.push(new LRow(k === 0 ? this.grid : `${this.grid}/${k}`, 'blank', '', { space: 11.4 }));
    }
    this.rows.push(new LRow(`${c.block}/r/data/sp`, 'blank', '', { space: 6 }));
    this.rows.push(new LRow(`${c.block}/r/data/section`, 'section', 'Chart data: every number is a formula on the statements', { indent: 1 }));
    this.rows.push(...this.pendingIndex);
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
    this.waterfall(ch, t, items, frame === 'at' ? 'col' : 'bar', endRef);
    return t;
  }

  budget(ch: ReportChartDef): Table {
    this.needs('year');
    if (!this.cmpRow) throw new AssemblyError(`${this.where(ch)}: a budget chart needs a Compared with setting`);
    const l = this.line(ch, ch.line!);
    const lead = ch.title.replace(/ against budget$/, '');
    const yl = `INDEX(List_Years,${this.year})`;
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

  /** A waterfall's rows and chart: items' values, then the running total, the base, totals, increases and decreases. */
  waterfall(ch: ReportChartDef, t: Table, items: { label: string; labelF?: string; value: string | null; kind: 'start' | 'step' | 'end' }[],
    dir: 'col' | 'bar', endRef: string, incLabel = 'Increase', decLabel = 'Decrease', tolerance = 0.01): void {
    t.header(items.map(i => i.labelF ?? i.label), 'Bridge');
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
    const ir = t.add({ label: incLabel, unit: '$', values: inc });
    const dr = t.add({ label: decLabel, unit: '$', values: dec });
    const n = items.length;
    t.check(`${ch.id}: the bridge does not reach the statement figure`, `=IF(ISNA(${c(vr, n - 1)}),0,IF(ABS(${c(vr, n - 1)}-${endRef})>${tolerance},1,0))`, 'error');
    t.check(`${ch.id}: the running total crosses zero, so a step is drawn from the wrong base`, `=IF(ISNA(${c(vr, n - 1)}),0,IF(MIN(${across(rr, n)})<0,1,0))`, 'alert');
    this.push({ id: ch.id, title: `${ch.title}, ${this.defaultLabel(ch.frame === 'at' ? 'at' : 'year')}`, titleRow: t.titleRow, cats: t.head, n, type: 'bar',
      dir, grouping: 'stacked', gap: 35, overlap: 100, valueAxis: false, reverse: dir === 'bar', legend: null,
      series: [
        { row: br, as: 'bar', colour: 'none' },
        { row: tr, as: 'bar', colour: 'grey', outline: true, labels: { fmt: '#,##0;-#,##0;', pos: 'inEnd' } },
        { row: ir, as: 'bar', colour: 'accent1', outline: true, labels: { fmt: '"+"#,##0;;', pos: 'ctr' } },
        { row: dr, as: 'bar', colour: 'accent3', outline: true, labels: { fmt: '"-"#,##0;;', pos: 'ctr' } },
      ] });
  }

  // --- the version comparison (HFG additions with no reference chart) ----------------------------
  /** A comparison's total for a line over periods a to b: the live line for the budget being built, else the stored version. */
  cmpSpan(k: number, key: string, a: string, b: string, live: string): string {
    const cmp = this.cmps[k];
    const m = `MATCH(${cmp.id}&"|"&${key},VS_Keys,0)`;
    return `IF(«V|${cmp.row}»=0,${live},IF(«V|${cmp.row}»<0,NA(),IFERROR(SUM(INDEX(VS_Values,${m},${a}):INDEX(VS_Values,${m},${b})),NA())))`;
  }

  /** A comparison's value for a line at period i; blank where the version has none. */
  cmpAt(k: number, key: string, i: string, live: string): string {
    const cmp = this.cmps[k];
    const v = `INDEX(VS_Values,MATCH(${cmp.id}&"|"&${key},VS_Keys,0),${i})`;
    return `IF(«V|${cmp.row}»=0,${live},IF(«V|${cmp.row}»<0,NA(),IFERROR(IF(${v}="",NA(),${v}),NA())))`;
  }

  needsVersions(ch: ReportChartDef, k = 1): void {
    if (this.cmps.length < k) throw new AssemblyError(`${this.where(ch)}: needs ${k === 1 ? 'a Compared with setting' : 'two comparisons'}`);
    if (!this.lineRow && ch.recipe !== 'variance' && ch.recipe !== 'walk') throw new AssemblyError(`${this.where(ch)}: needs a line shown`);
  }

  /** The month, year to date and full year against the comparisons, for every line a version keeps (no chart). */
  variance(ch: ReportChartDef): Table {
    this.needs('at');
    this.needsVersions(ch);
    const c = this.ctx;
    const M = c.month!;
    const [ys, ye] = [`«V|${this.ys}»`, `«V|${this.ye}»`];
    const ytd0 = `«V|${this.ytd0}»`;
    const t = new Table(this, ch, `="${ch.title}: "&TEXT(INDEX(List_Months,${M}),"mmm yyyy")&", year to date and full year"`);
    const groups = [
      { key: 'month', label: `="Month: "&TEXT(INDEX(List_Months,${M}),"mmm yyyy")`, a: M, b: M, head: `=IF(${M}<=Tl_Last_Actual,"Actual","Forecast")` },
      { key: 'ytd', label: `="Year to date to "&TEXT(INDEX(List_Months,${M}),"mmm yyyy")`, a: ytd0, b: M, head: `=IF(${M}<=Tl_Last_Actual,"Actual","To date")` },
      { key: 'fy', label: `="Full year: "&INDEX(List_Years,${this.year})`, a: ys, b: ye, head: 'Outturn' },
    ];
    const per = 1 + 2 * this.cmps.length;
    const groupCells: Record<number, unknown> = {};
    groups.forEach((g, gi) => { groupCells[J + per * gi] = g.label; });
    t.rows.push(new LRow(`${t.base}/groups`, 'table', '', { indent: 2, role: 'r.head', cells: groupCells }));
    const heads: string[] = [];
    for (const g of groups) {
      heads.push(g.head);
      this.cmps.forEach((_, k) => heads.push(k === 0 ? 'Compared' : 'Second', 'Variance'));
    }
    t.header(heads, 'Line');
    for (const key of VERSION_LINES) {
      const l = c.spec.lines[key];
      if (!l) continue;
      const rng = `«A|${c.fs}/${l.row}»`;
      const values: string[] = [];
      for (const g of groups) {
        const live = g.a === g.b ? `INDEX(${rng},${g.a})` : spanSum(rng, g.a, g.b);
        const actualCol = J + values.length;
        values.push(`=${live}`);
        this.cmps.forEach((_, k) => {
          const cmpCol = J + values.length;
          values.push(`=${g.a === g.b ? this.cmpAt(k, `"${key}"`, g.a, live) : this.cmpSpan(k, `"${key}"`, g.a, g.b, live)}`);
          const [x, y] = COST_LINES.has(key) ? [cmpCol, actualCol] : [actualCol, cmpCol];
          values.push(`=«C${x}|{self}»-«C${y}|{self}»`);
        });
      }
      t.add({ label: l.label, unit: '$', values, style: ['rev', 'gm', 'ebitda', 'npat'].includes(key) ? 'bold' : '' });
    }
    t.add({ labelF: `="Compared: "&«V|${this.cmps[0].label}»${this.cmps[1] ? `&". Second: "&«V|${this.cmps[1].label}»` : ''}&". Variances are favourable when positive: for costs, spending less than the comparison."`,
      role: 'r.note', values: [] });
    return t;
  }

  /** Full-year outturn by version: the approved budget, each reforecast saved in the year, and the current forecast. */
  trend(ch: ReportChartDef): Table {
    this.needs('at');
    this.needsVersions(ch);
    const c = this.ctx;
    const M = c.month!;
    const key = this.n('Line_Key');
    const [ys, ye] = [`«V|${this.ys}»`, `«V|${this.ye}»`];
    const t = new Table(this, ch, `="${ch.title}: "&${this.n('Line_Label')}&", "&INDEX(List_Years,${this.year})`);
    const asat = `${t.base}/asat`;
    const ids = `${t.base}/ids`;
    const n = 13;
    t.header(['Budget', ...Array.from({ length: 11 }, (_, j) => `=TEXT(INDEX(List_Months,MAX(1,«C${J + j + 1}|${asat}»)),"mmm")`), 'Current'], 'Version');
    const asatCells: Record<number, unknown> = {};
    for (let j = 1; j <= 11; j++) asatCells[J + j] = `=«V|${this.y0}»-1+${j}`;
    t.rows.push(new LRow(asat, 'table', 'Reforecast as at month', { indent: 2, role: 'r.index', cells: asatCells }));
    const idCells: Record<number, unknown> = { [J]: `=IFERROR(INDEX(VR_Id,MATCH("Budget|Approved|"&INDEX(List_Years,${this.year}),VR_BudgetKey,0)),"")` };
    for (let j = 1; j <= 11; j++) {
      const a = `«C${J + j}|${asat}»`;
      idCells[J + j] = `=IF(OR(${a}<1,${a}>${M}),"",IFERROR(INDEX(VR_Id,MATCH("Reforecast|"&${a},VR_RefKey,0)),""))`;
    }
    t.rows.push(new LRow(ids, 'table', 'Version id', { indent: 2, role: 'r.index', unit: 'text', cells: idCells }));
    const stored = (j: number) => {
      const id = `«C${J + j}|${ids}»`;
      const m = `MATCH(${id}&"|"&${key},VS_Keys,0)`;
      return `=IF(${id}="","",IFERROR(SUM(INDEX(VS_Values,${m},${ys}):INDEX(VS_Values,${m},${ye})),""))`;
    };
    const blank = (from: number, to: number) => Array.from({ length: to - from }, () => '=""');
    const rb = t.add({ label: 'Approved budget', unit: '$', values: [stored(0), ...blank(1, n)] });
    const rr = t.add({ label: 'Reforecasts', unit: '$', values: ['=""', ...Array.from({ length: 11 }, (_, j) => stored(j + 1)), '=""'] });
    const rc = t.add({ label: 'Current forecast', unit: '$', style: 'bold', values: [...blank(0, n - 1), `=${spanSum(`«A|${this.lineRow}»`, ys, ye)}`] });
    t.add({ label: "Each reforecast's total for the year: actuals to the month it was saved, then its forecast.", role: 'r.note', values: [] });
    this.push({ id: ch.id, title: `${ch.title}, ${this.defaultLabel('year')}`, titleRow: t.titleRow, cats: t.head, n, type: 'bar', dir: 'col',
      grouping: 'stacked', gap: 45, overlap: 100, yFmt: '#,##0', legend: 'b',
      series: [{ row: rb, as: 'bar', colour: 'accent1' }, { row: rr, as: 'bar', colour: 'grey' }, { row: rc, as: 'bar', colour: 'tx2' }] });
    return t;
  }

  /** Actual (solid) and forecast (hatched) by month of the year, with both comparisons as lines. */
  versions(ch: ReportChartDef): Table {
    this.needs('at');
    this.needsVersions(ch);
    const key = this.n('Line_Key');
    const live = `«A|${this.lineRow}»`;
    const t = new Table(this, ch, `="${ch.title}: "&${this.n('Line_Label')}&", "&INDEX(List_Years,${this.year})`);
    t.header(this.cats('year'), 'Series', 'Total');
    const out = (i: string) => `OR(${i}<1,${i}>Tl_Term)`;
    const vals = (f: (i: string) => string) => Array.from({ length: 12 }, (_, j) => {
      const i = this.idx('year', j);
      return `=IF(${out(i)},NA(),${f(i)})`;
    });
    const ra = t.add({ label: 'Actual', unit: '$', total: fullTotal(), values: vals(i => `IF(${i}<=Tl_Last_Actual,INDEX(${live},${i}),0)`) });
    const rf = t.add({ label: 'Forecast', unit: '$', total: fullTotal(), values: vals(i => `IF(${i}<=Tl_Last_Actual,0,INDEX(${live},${i}))`) });
    const series: RSeries[] = [{ row: ra, as: 'bar', colour: 'tx2' }, { row: rf, as: 'bar', colour: 'tx2', hatch: true }];
    this.cmps.forEach((cmp, k) => {
      const r = t.add({ labelF: `=«V|${cmp.label}»`, unit: '$', style: k ? '' : 'bold', total: fullTotal(),
        values: vals(i => this.cmpAt(k, key, i, `INDEX(${live},${i})`)) });
      series.push({ row: r, as: 'line', colour: k ? 'accent3' : 'accent1', width: k ? 1.75 : 2.25, dash: k > 0, marker: true });
    });
    this.push({ id: ch.id, title: `${ch.title}, ${this.defaultLabel('year')}`, titleRow: t.titleRow, cats: t.head, n: 12, type: 'bar', dir: 'col',
      grouping: 'stacked', gap: 55, overlap: 100, yFmt: '#,##0', legend: 'b', series });
    return t;
  }

  /** What the budget and each of the twelve reforecasts before the month shown expected for it, against the actual. */
  accuracy(ch: ReportChartDef): Table {
    this.needs('at');
    this.needsVersions(ch);
    const M = this.ctx.month!;
    const key = this.n('Line_Key');
    const t = new Table(this, ch, `="${ch.title}: "&${this.n('Line_Label')}&", "&TEXT(INDEX(List_Months,${M}),"mmm yyyy")`);
    const asat = `${t.base}/asat`;
    const ids = `${t.base}/ids`;
    const n = 13;
    t.header(['Budget', ...Array.from({ length: 12 }, (_, j) => `=IF(«C${J + j + 1}|${asat}»<1,"",TEXT(INDEX(List_Months,«C${J + j + 1}|${asat}»),"mmm yy"))`)], 'Saved as at');
    const asatCells: Record<number, unknown> = {};
    for (let j = 1; j <= 12; j++) asatCells[J + j] = `=${M}-${13 - j}`;
    t.rows.push(new LRow(asat, 'table', 'Reforecast as at month', { indent: 2, role: 'r.index', cells: asatCells }));
    const idCells: Record<number, unknown> = { [J]: `=IFERROR(INDEX(VR_Id,MATCH("Budget|Approved|"&INDEX(List_Years,${this.year}),VR_BudgetKey,0)),"")` };
    for (let j = 1; j <= 12; j++) {
      const a = `«C${J + j}|${asat}»`;
      idCells[J + j] = `=IF(${a}<1,"",IFERROR(INDEX(VR_Id,MATCH("Reforecast|"&${a},VR_RefKey,0)),""))`;
    }
    t.rows.push(new LRow(ids, 'table', 'Version id', { indent: 2, role: 'r.index', unit: 'text', cells: idCells }));
    const stored = (j: number) => {
      const id = `«C${J + j}|${ids}»`;
      const v = `INDEX(VS_Values,MATCH(${id}&"|"&${key},VS_Keys,0),${M})`;
      return `=IF(${id}="","",IFERROR(IF(${v}="","",${v}),""))`;
    };
    const rb = t.add({ label: 'Approved budget', unit: '$', values: [stored(0), ...Array.from({ length: 12 }, () => '=""')] });
    const rr = t.add({ label: 'Reforecasts', unit: '$', values: ['=""', ...Array.from({ length: 12 }, (_, j) => stored(j + 1))] });
    const rc = t.add({ labelF: `=IF(${M}<=Tl_Last_Actual,"Actual","Current forecast")`, unit: '$', style: 'bold',
      values: Array.from({ length: n }, () => `=INDEX(«A|${this.lineRow}»,${M})`) });
    t.add({ label: 'Each bar is what that version expected for the month shown; the line is what happened (or the current forecast).', role: 'r.note', values: [] });
    this.push({ id: ch.id, title: `${ch.title}, ${this.defaultLabel('at')}`, titleRow: t.titleRow, cats: t.head, n, type: 'bar', dir: 'col',
      grouping: 'stacked', gap: 45, overlap: 100, yFmt: '#,##0', legend: 'b',
      series: [{ row: rb, as: 'bar', colour: 'accent1' }, { row: rr, as: 'bar', colour: 'grey' }, { row: rc, as: 'line', colour: 'tx2', width: 2.25 }] });
    return t;
  }

  /** A waterfall from the comparison's profit after tax for the year to the outturn, one step per line's variance. */
  walk(ch: ReportChartDef): Table {
    this.needs('at');
    this.needsVersions(ch);
    const c = this.ctx;
    const [ys, ye] = [`«V|${this.ys}»`, `«V|${this.ye}»`];
    const out = (k: string) => spanSum(`«A|${c.fs}/${c.spec.lines[k].row}»`, ys, ye);
    const cmp = (k: string) => this.cmpSpan(0, `"${k}"`, ys, ye, out(k));
    const t = new Table(this, ch, `="${ch.title}, full year "&INDEX(List_Years,${this.year})`);
    const items: { label: string; labelF?: string; value: string | null; kind: 'start' | 'step' | 'end' }[] = [
      { label: 'Compared with', labelF: `=«V|${this.cmps[0].label}»`, value: `=${cmp('npat')}`, kind: 'start' }];
    for (const k of ['rev', 'other_income', 'cogs', 'staff', 'opex', 'other_expense', 'da', 'interest', 'tax']) {
      if (!c.spec.lines[k]) continue;
      items.push({ label: c.spec.lines[k].label, value: COST_LINES.has(k) ? `=${cmp(k)}-${out(k)}` : `=${out(k)}-${cmp(k)}`, kind: 'step' });
    }
    items.push({ label: 'Outturn', value: null, kind: 'end' });
    // saved lines are rounded to the cent, so their sum can differ from the saved profit by cents
    this.waterfall(ch, t, items, 'col', out('npat'), 'Favourable', 'Unfavourable', 1);
    return t;
  }

  // --- dashboard tables ----------------------------------------------------------------------------
  /** A dashboard table's rows, resolved: what each reads, its label and style. */
  private summaryRows(ch: ReportChartDef): { key: string; label: string; labelF?: string; rng?: string; kind: 'flow' | 'balance';
    style: string; unit: string; ratio?: [string, string]; less?: string; start?: boolean }[] {
    const out: ReturnType<Expander['summaryRows']> = [];
    for (const d of ch.rows ?? []) {
      if (d.group) {
        const g = this.group(ch, d.group);
        g.ids.forEach((id, k) => out.push({ key: `${d.group}/${k + 1}`, label: '', labelF: `=INDEX(${g.labels},${k + 1})`,
          rng: `«A|${id}»`, kind: 'flow', style: k === g.ids.length - 1 ? 'italic+last' : 'italic', unit: '$' }));
      } else if (d.line) {
        const l = this.line(ch, d.line);
        if (l.kind === 'ratio') throw new AssemblyError(`${this.where(ch)}: show ${d.line} as a ratio of two lines in the table`);
        out.push({ key: d.line, label: d.label ?? l.label, rng: l.rng, kind: l.kind === 'balance' ? 'balance' : 'flow',
          style: d.style ?? '', unit: '$', start: d.at === 'start' });

      } else if (d.ratio) {
        out.push({ key: `ratio/${d.ratio.join('/')}${d.less ? `-${d.less}` : ''}`, label: d.label ?? `${d.ratio[0]} / ${d.ratio[1]}`,
          kind: 'flow', style: 'italic', unit: d.unit ?? '%', ratio: d.ratio, less: d.less });
      } else {
        throw new AssemblyError(`${this.where(ch)}: a table row needs a line, a group or a ratio`);
      }
    }
    return out;
  }

  /**
   * A statement month by month for the table's twelve months (as the charts, the 12 months to or
   * after the last actual month, or 12 months from any month), a total, then the comparisons asked
   * for: the 12 months before, the year to date against a year earlier, or the last month against a
   * year earlier. Ratios are worked out from the table's own lines, column by column. The last actual
   * month's column is shaded.
   */
  summary(ch: ReportChartDef): Table {
    const W = `«V|${this.w0}»`;
    const t = new Table(this, ch, '', true);
    const rows = this.summaryRows(ch);
    const totals = ch.totals !== false;
    const last = J + 11;
    const p = (j: number) => `«C${J + j}|${this.widx}»`;
    const inT = (i: string) => `OR(${i}<1,${i}>Tl_Term)`;
    const at = (rng: string, i: string) => `IF(${inT(i)},"",INDEX(${rng},${i}))`;
    const span = (rng: string, a: string, b: string) => `IF(OR(${a}<1,${b}>Tl_Term,${b}<${a}),"",SUM(INDEX(${rng},${a}):INDEX(${rng},${b})))`;
    const L = 'Tl_Last_Actual';
    const YA = this.ya ? `«V|${this.ya}»` : '0';

    // columns: the months, the total, then each comparison block after a gap
    type Col = { col: number; head: string; sub?: string; role: 'month' | 'total' | 'cur' | 'cmp' | 'chg' | 'pct'; j?: number;
      block?: 'prior' | 'ytd' | 'year_ago'; cur?: number; cmp?: number };
    const cols: Col[] = [];
    for (let j = 0; j < 12; j++) cols.push({ col: J + j, head: `=«C${J + j}|${this.wcats}»`, role: 'month', j,
      sub: `=IF(${inT(p(j))},"",IF(${p(j)}<=${L},Tl_Actual_Label,Tl_Forecast_Label))` });
    let next = last + 1;
    const blocks: { from: number; to: number; label: string }[] = [];
    if (totals) cols.push({ col: next++, head: 'Total', role: 'total' });
    blocks.push({ from: J, to: next - 1, label: `=«V|${this.wlabel}»` });
    for (const b of ch.against ?? []) {
      next += 1;   // a gap column
      const from = next;
      if (b === 'prior') {
        const cur = totals ? last + 1 : last;
        cols.push({ col: next++, head: 'Year before', role: 'cmp', block: b });
        cols.push({ col: next++, head: 'Change', role: 'chg', cur, cmp: next - 2 });
        cols.push({ col: next++, head: 'Change %', role: 'pct', cur, cmp: next - 3 });
        blocks.push({ from, to: next - 1, label: 'Against the 12 months before' });
      } else {
        const lab = b === 'ytd' ? 'To date' : `=TEXT(EOMONTH(Tl_Start,${W}+10),"mmm yy")`;
        cols.push({ col: next++, head: lab, role: 'cur', block: b, sub: b === 'ytd' ? `=IF(${L}<1,"",Tl_Actual_Label)` : undefined });
        cols.push({ col: next++, head: 'Year before', role: 'cmp', block: b });
        cols.push({ col: next++, head: 'Change', role: 'chg', cur: from, cmp: from + 1 });
        cols.push({ col: next++, head: 'Change %', role: 'pct', cur: from, cmp: from + 1 });
        blocks.push({ from, to: next - 1, label: b === 'ytd'
          ? `=IF(${L}<1,"Year to date: no actuals yet","Year to date to "&TEXT(INDEX(List_Months,${L}),"mmm yyyy")&" against a year earlier")`
          : `="At "&TEXT(EOMONTH(Tl_Start,${W}+10),"mmm yyyy")&" against a year earlier"` });
      }
    }

    // headings: the blocks, the months, then actual or forecast under each month
    const blockCells: Record<number, unknown> = {};
    for (const b of blocks) blockCells[b.from] = b.label;
    t.rows.push(new LRow(`${t.base}/blocks`, 'table', '', { indent: 2, role: 'r.blocks', cells: blockCells,
      merges: blocks.map(b => [b.from, b.to] as [number, number]) }));
    t.rows.push(new LRow(t.head, 'table', 'Month', { indent: 2, role: 'r.head', cells: Object.fromEntries(cols.map(c => [c.col, c.head])) }));
    t.rows.push(new LRow(`${t.base}/type`, 'table', 'Actual or forecast', { indent: 2, role: 'r.sub',
      cells: Object.fromEntries(cols.filter(c => c.sub).map(c => [c.col, c.sub])) }));

    // the lines
    const ids = new Map<string, string>();
    rows.forEach((r, k) => ids.set(r.key, `${t.base}/${k + 1}`));
    rows.forEach((r, k) => {
      const id = `${t.base}/${k + 1}`;
      const self = (col: number) => `«C${col}|${id}»`;
      const ratioAt = (col: number) => {
        const [a, b] = r.ratio!;
        const ref = (key: string) => {
          const rid = ids.get(key);
          if (!rid) throw new AssemblyError(`${this.where(ch)}: the ratio reads ${key}, which is not a line of the table`);
          return `«C${col}|${rid}»`;
        };
        return `=IFERROR((${ref(a)}${r.less ? `-${ref(r.less)}` : ''})/${ref(b)},"")`;
      };
      const flow = r.kind === 'flow';
      const cells: Record<number, unknown> = {};
      for (const c of cols) {
        if (c.role === 'chg') {
          cells[c.col] = `=IF(OR(${self(c.cur!)}="",${self(c.cmp!)}=""),"",${self(c.cur!)}-${self(c.cmp!)})`;
          continue;
        }
        if (c.role === 'pct') {
          cells[c.col] = r.ratio ? '=""' : `=IF(OR(${self(c.cur!)}="",${self(c.cmp!)}="",${self(c.cmp!)}=0),"",(${self(c.cur!)}-${self(c.cmp!)})/ABS(${self(c.cmp!)}))`;
          continue;
        }
        if (r.ratio) {
          cells[c.col] = ratioAt(c.col);
          continue;
        }
        const rng = r.rng!;
        // a balance is read at a month's end (or its start, for opening cash); a flow adds up over months
        const bal = (i: string, startI?: string) => `=${at(rng, r.start && startI ? startI : i)}`;
        if (c.role === 'month') cells[c.col] = `=${at(rng, p(c.j!))}`;
        else if (c.role === 'total') {
          cells[c.col] = flow ? `=IF(COUNT(${self(J)}:${self(last)})=0,"",SUM(${self(J)}:${self(last)}))` : bal(`(${W}+11)`, W);
        } else if (c.block === 'prior') {
          cells[c.col] = flow ? `=${span(rng, `(${W}-12)`, `(${W}-1)`)}` : bal(`(${W}-1)`, `(${W}-12)`);
        } else if (c.block === 'ytd') {
          const shift = c.role === 'cmp' ? '-12' : '';
          cells[c.col] = flow ? `=IF(${YA}<1,"",${span(rng, `(${YA}${shift})`, `(${L}${shift})`)})` : bal(`(${L}${shift})`, `(${YA}${shift})`);
        } else {   // year_ago: the last month against the same month a year earlier
          cells[c.col] = c.role === 'cmp' ? bal(`(${W}-1)`, `(${W}-12)`) : bal(`(${W}+11)`, W);
        }
      }
      const shade = `${p(0)}=${L}`;
      const units: Record<number, string> = Object.fromEntries(cols.filter(c => c.role === 'pct').map(c => [c.col, '%']));
      if (r.ratio && r.unit === '%') for (const c of cols) if (c.role === 'chg') units[c.col] = '%';
      t.rows.push(new LRow(id, 'table', r.label, { indent: 2, unit: r.ratio ? r.unit : '$', role: 'r.sum', style: r.style, shade, units,
        cells: r.labelF ? { [LABEL_COLS[2]]: r.labelF, ...cells } : cells }));
    });
    return t;
  }

  /** The table's four quarters, each against the same quarter a year earlier. */
  quarters(ch: ReportChartDef): Table {
    const W = `«V|${this.w0}»`;
    const t = new Table(this, ch, '', true);
    const rows = this.summaryRows(ch);
    const blockCells: Record<number, unknown> = {};
    const merges: [number, number][] = [];
    const heads: Record<number, unknown> = {};
    for (let q = 0; q < 4; q++) {
      const c0 = J + 3 * q;
      const endP = `(${W}+${3 * q + 2})`;
      blockCells[c0] = `=IF(OR(${endP}<1,${endP}>Tl_Term),"","3 months to "&TEXT(INDEX(List_Months,${endP}),"mmm yy"))`;
      merges.push([c0, c0 + 2]);
      heads[c0] = 'This year';
      heads[c0 + 1] = 'Year before';
      heads[c0 + 2] = 'Change %';
    }
    t.rows.push(new LRow(`${t.base}/blocks`, 'table', '', { indent: 2, role: 'r.blocks', cells: blockCells, merges }));
    t.rows.push(new LRow(t.head, 'table', 'Quarter', { indent: 2, role: 'r.head', cells: heads }));
    rows.forEach((r, k) => {
      if (!r.rng || r.kind !== 'flow') throw new AssemblyError(`${this.where(ch)}: a quarter adds up a flow; ${r.key} is not one`);
      const id = `${t.base}/${k + 1}`;
      const cells: Record<number, unknown> = {};
      for (let q = 0; q < 4; q++) {
        const c0 = J + 3 * q;
        const sum = (shift: number) => {
          const a = `(${W}+${3 * q + shift})`;
          const b = `(${W}+${3 * q + 2 + shift})`;
          return `=IF(OR(${a}<1,${b}>Tl_Term),"",SUM(INDEX(${r.rng},${a}):INDEX(${r.rng},${b})))`;
        };
        const cur = `«C${c0}|${id}»`;
        const cmp = `«C${c0 + 1}|${id}»`;
        cells[c0] = sum(0);
        cells[c0 + 1] = sum(-12);
        cells[c0 + 2] = `=IF(OR(${cur}="",${cmp}="",${cmp}=0),"",(${cur}-${cmp})/ABS(${cmp}))`;
      }
      const units = Object.fromEntries([0, 1, 2, 3].map(q => [J + 3 * q + 2, '%']));
      // the quarter the last actual month falls in is shaded
      const q0 = `«V|${this.w0}»+3*INT((COLUMNS($J1:J1)-1)/3)`;
      const shade = `AND(Tl_Last_Actual>=${q0},Tl_Last_Actual<=${q0}+2)`;
      t.rows.push(new LRow(id, 'table', r.label, { indent: 2, unit: '$', role: 'r.sum', style: r.style, units, shade,
        cells: r.labelF ? { [LABEL_COLS[2]]: r.labelF, ...cells } : cells }));
    });
    return t;
  }

  /**
   * The year shown month by month three ways: the model (actual and forecast), the version it is
   * compared with, and the variance (favourable when positive, an unfavourable one in red).
   */
  budget_table(ch: ReportChartDef): Table {
    this.needs('year');
    if (!this.cmps.length) throw new AssemblyError(`${this.where(ch)}: needs a Compared with setting`);
    const c = this.ctx;
    const t = new Table(this, ch, '', true);
    const keys = VERSION_LINES.filter(k => k !== 'opcosts' && c.spec.lines[k]);
    const heads: Record<number, unknown> = { [J + 12]: 'Total' };
    const subs: Record<number, unknown> = {};
    for (let j = 0; j < 12; j++) {
      const i = this.idx('year', j);
      heads[J + j] = `=«C${J + j}|${this.mcats}»`;
      subs[J + j] = `=IF(OR(${i}<1,${i}>Tl_Term),"",IF(${i}<=Tl_Last_Actual,Tl_Actual_Label,Tl_Forecast_Label))`;
    }
    t.rows.push(new LRow(`${t.base}/blocks`, 'table', '', { indent: 2, role: 'r.blocks', cells: { [J]: `=INDEX(List_Years,${this.year})` },
      merges: [[J, J + 12]] }));
    t.rows.push(new LRow(t.head, 'table', 'Month', { indent: 2, role: 'r.head', cells: heads }));
    t.rows.push(new LRow(`${t.base}/type`, 'table', 'Actual or forecast', { indent: 2, role: 'r.sub', cells: subs }));
    const bold = new Set(['gm', 'ebitda', 'ebit', 'npbt', 'npat']);
    const parts: { key: string; label: string; labelF?: string; role: string }[] = [
      { key: 'model', label: 'Actual and forecast', role: 'r.sum' },
      { key: 'cmp', label: '', labelF: `=«V|${this.cmps[0].label}»`, role: 'r.sum' },
      { key: 'var', label: 'Variance (favourable when positive)', role: 'r.var' },
    ];
    for (const part of parts) {
      t.rows.push(new LRow(`${t.base}/${part.key}/sp`, 'blank', '', { space: 6 }));
      t.rows.push(new LRow(`${t.base}/${part.key}`, 'table', part.label, { indent: 2, role: 'r.head',
        cells: part.labelF ? { [LABEL_COLS[2]]: part.labelF } : {} }));
      for (const key of keys) {
        const l = c.spec.lines[key];
        const rng = `«A|${c.fs}/${l.row}»`;
        const id = `${t.base}/${part.key}/${key}`;
        const cells: Record<number, unknown> = {};
        for (let j = 0; j < 12; j++) {
          const i = this.idx('year', j);
          const live = `INDEX(${rng},${i})`;
          const model = `«C${J + j}|${t.base}/model/${key}»`;
          const cmp = `«C${J + j}|${t.base}/cmp/${key}»`;
          cells[J + j] = part.key === 'model' ? `=IF(OR(${i}<1,${i}>Tl_Term),"",${live})`
            : part.key === 'cmp' ? `=IF(OR(${i}<1,${i}>Tl_Term),"",IFERROR(${this.cmpAt(0, `"${key}"`, i, live)},""))`
              : `=IF(OR(${model}="",${cmp}=""),"",${COST_LINES.has(key) ? `${cmp}-${model}` : `${model}-${cmp}`})`;
        }
        cells[J + 12] = `=IF(COUNT(«C${J}|${id}»:«C${J + 11}|${id}»)=0,"",SUM(«C${J}|${id}»:«C${J + 11}|${id}»))`;
        t.rows.push(new LRow(id, 'table', l.label, { indent: 2, unit: '$', role: part.role, style: bold.has(key) ? 'bold' : '', cells }));
      }
    }
    return t;
  }

  /** Statement lines by month over the frame, with totals, and the gap to a target profit when the module has one (no chart). */
  statement(ch: ReportChartDef): Table {
    const frame = (ch.frame ?? 'year') as Frame;
    this.needs(frame);
    const c = this.ctx;
    const t = new Table(this, ch, `="${ch.title}, "&${this.windowLabel(frame)}`);
    t.header(this.cats(frame), 'Line', 'Total');
    let npat = '';
    for (const key of ch.lines ?? VERSION_LINES) {
      const l = this.line(ch, key);
      const id = t.add({ label: l.label, unit: '$', style: ['rev', 'gm', 'ebitda', 'npat'].includes(key) ? 'bold' : '',
        total: `=SUM(${across('{self}', 12)})`,
        values: Array.from({ length: 12 }, (_, j) => `=IF(${this.idx(frame, j)}<1,NA(),${pickAt(l.rng, this.idx(frame, j))})`) });
      if (key === 'npat') npat = id;
    }
    if (c.target && npat) {
      t.add({ label: 'Target profit after tax', unit: '$', values: [], total: `=${c.target}` });
      t.add({ label: 'Gap to the target (above it when positive)', unit: '$', style: 'signed', values: [], total: `=«C${TOTAL_COL}|${npat}»-${c.target}` });
    }
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
      let group: string;
      let item: string;
      if (by === 'month') {
        f = `=IF(OR(${this.n('Y0')}+${j}<1,${this.n('Y0')}+${j}>Tl_Term),0,INDEX(${l.rng},${this.n('Y0')}+${j}))`;
        label = `${c.title} ${ch.id}: ${l.label}, month ${j + 1} of the year shown`;
        group = `${l.label}, each month of the year shown on the ${c.title}`;
        item = `=TEXT(DATE(2000,Sel_FY_End_Month+${j + 1},1),"mmmm")`;
      } else {
        const a = `MAX(1,${j * 12}-Sel_Start_Month+2)`;
        const b = `MIN(Tl_Term,${(j + 1) * 12}-Sel_Start_Month+1)`;
        f = l.kind === 'flow' ? `=${spanSum(l.rng, a, b)}` : `=INDEX(${l.rng},${b})`;
        label = `${c.title} ${ch.id}: ${l.label}, ${c.fyLabel(j + 1)}`;
        group = `${l.label}, ${l.kind === 'flow' ? 'each financial year' : 'at each year end'}`;
        item = `=INDEX(List_Years,${j + 1})`;
      }
      // Two charts that show the same figures share one row of the data table.
      const same = c.results.find(x => x.formula === f);
      if (same) {
        res.push(same.id);
        continue;
      }
      c.results.push({ id, label, group, item, formula: f });
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
  const run = (ch: ReportChartDef) => {
    const fn = (e as unknown as Record<string, (c: ReportChartDef) => Table>)[ch.recipe];
    if (typeof fn !== 'function' || !RECIPES.includes(ch.recipe)) throw new AssemblyError(`${ctx.title} chart ${ch.id}: no recipe '${ch.recipe}'`);
    e.rows.push(...fn.call(e, ch).rows);
  };
  // tables without a chart (the variance table) sit above the chart grid
  for (const ch of ctx.charts) if (TABLES.includes(ch.recipe)) run(ch);
  e.gridRows();
  for (const ch of ctx.charts) if (!TABLES.includes(ch.recipe)) run(ch);
  return { rows: e.rows, charts: e.charts, errors: e.errors, alerts: e.alerts, names: e.names };
}

const RECIPES = ['compare', 'mix', 'depth', 'pie', 'combo', 'movement', 'bridge', 'scenario', 'budget', 'variance', 'trend', 'versions', 'accuracy', 'walk',
  'statement', 'summary', 'quarters', 'budget_table'];
/** Recipes that make a table and no chart; they sit above the chart grid. */
const TABLES = ['variance', 'statement', 'summary', 'quarters', 'budget_table'];

export const REPORT_UNIT_COL = UNIT_COL;
export { colLetter };
