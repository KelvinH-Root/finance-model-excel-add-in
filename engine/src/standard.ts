// The HFG standard frame (docs/frame-standard.md, Look and wiring): the Settings sheet the timeline
// comes from, the contents with the entity, model name line, notes and a grouped table of contents,
// section covers, cell hyperlinks to Go_ names with screen tips, and the header every sheet carries.
//
// HFG's own name prefixes (7 October 2026): Go_ navigation targets, Tl_ timeline settings, Opt_
// switches, Model_ the model's own lines, Chk_ check totals; GA_, Reg_ and KO_ as before.

import { FRAME_KINDS, keyOutputFormula } from './assurance.ts';
import { AssemblyError, code, colLetter, CONTENTS, FIRST_PERIOD_COL, pick, STANDARD_FRAME, STD, TOTAL_COL, UNIT_COL } from './frame.ts';
import { LRow, type CellLink, type DataTable, type Layout, type RangeName } from './layout.ts';
import type { SectionDef } from './library.ts';
import type { Model, ModelInfo } from './model.ts';
import { unoSeparators, type Dialect, type FrameCell } from './render.ts';
import type { Block } from './resolve.ts';
import type { FrameRequests } from './assemble.ts';

export const SETTINGS = 'Settings';
export const GO_CONTENTS = 'Go_Contents';
export const GO_CHECKS = 'Go_Checks';
export const goSheet = (sheet: string) => `Go_Sheet_${code(sheet)}`;
export const goBlock = (uid: string) => `Go_${code(uid)}`;

/** Navigation symbols, drawn in Segoe UI Symbol. */
export const SYMBOL = { home: '⌂', clear: '✓', failing: '✗', prev: '◀', next: '▶' } as const;

/** Days since 30 December 1899, as Excel counts dates. */
export function excelDate(y: number, m: number, d: number): number {
  return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
}

/** "2026-04" -> 1 April 2026 as an Excel date. */
export function monthStart(ym: string): number {
  const m = /^(\d{4})-(\d{2})$/.exec(ym);
  if (!m) throw new AssemblyError(`timeline start ${ym} is not a month (yyyy-mm)`);
  return excelDate(Number(m[1]), Number(m[2]), 1);
}

/** The timeline block on every timeline sheet: label and formula. Each sheet works its block out from the Settings choices (the Tl_ and Sel_ names); Settings itself has no months across it. */
export const BLOCK: { row: number; label: string; source: (c: number, col: string, prev: string | null) => string }[] = [
  { row: 5, label: 'Month ending', source: (_c, L) => `=${L}8` },
  { row: 6, label: 'Actual or forecast', source: (_c, L) => `=IF(${L}9<=Tl_Last_Actual,Tl_Actual_Label,Tl_Forecast_Label)` },
  { row: 7, label: 'Period start', source: (_c, L) => `=EDATE(Tl_Start,${L}9-1)` },
  { row: 8, label: 'Period end', source: (_c, L) => `=EOMONTH(${L}7,0)` },
  { row: 9, label: 'Period', source: (_c, _L, P) => (P ? `=${P}9+1` : '=1') },
  { row: 10, label: 'Financial year', source: (_c, L) => `=YEAR(${L}8)+IF(MONTH(${L}8)>Sel_FY_End_Month,1,0)` },
  { row: 11, label: 'Month of the year', source: (_c, L) => `=MOD(MONTH(${L}8)-Sel_FY_End_Month-1,12)+1` },
  { row: 12, label: 'Quarter of the year', source: (_c, L) => `=INT((${L}11-1)/3)+1` },
  { row: 13, label: 'Half of the year', source: (_c, L) => `=INT((${L}11-1)/6)+1` },
  { row: 14, label: 'Actual month', source: (_c, L) => `=IF(${L}9<=Tl_Last_Actual,1,0)` },
  { row: 15, label: 'Forecast month number', source: (_c, L) => `=MAX(0,${L}9-Tl_Last_Actual)` },
];

export const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September',
  'October', 'November', 'December'];
export const DENOMINATIONS = ['$', '$000', '$m'] as const;
export const LOOKUPS = 'Lookups';
/** The column a lookup list's items sit in (D), with each item's position in C and the List_ name in E. */
export const LIST_COL = 4;

/** The financial year a "yyyy-mm" month falls in, and its position in that year (1 is the month after the year end). */
export function fiscalPosition(ym: string, fyEndMonth: number): { year: number; month: number } {
  const m = /^(\d{4})-(\d{2})$/.exec(ym);
  if (!m) throw new AssemblyError(`timeline start ${ym} is not a month (yyyy-mm)`);
  const y = Number(m[1]);
  const mo = Number(m[2]);
  return { year: y + (mo > fyEndMonth ? 1 : 0), month: ((mo - fyEndMonth - 1) % 12 + 12) % 12 + 1 };
}

/** The budget a model starts with: the twelve months after its actuals, inside the timeline. */
export function defaultBudget(lastActual: number, periods: number): { first: number; months: number } {
  const first = Math.min(lastActual + 1, periods);
  return { first, months: Math.min(12, periods - first + 1) };
}

/** A lookup list: a List_ range on the Lookups sheet that a drop-down reads. */
export interface ListSpec {
  name: string;
  title: string;
  /** The heading the list sits under on the Lookups sheet. */
  group: string;
  /** Each item's value (or formula) and the style key its cell takes (lu.text, lu.monthYear, lu.int). */
  items: { value: unknown; style: string }[];
}

/** The timeline's lists: months, start months, the model's months, last actual month and denominations. */
export function timelineLists(periods: number, years = 0): ListSpec[] {
  const group = 'Timeline lists';
  const months = (first: unknown[] = []) => [...first.map(value => ({ value, style: 'lu.text' })),
    ...Array.from({ length: periods }, (_, k) => ({ value: `=EOMONTH(Tl_Start,${k})`, style: 'lu.monthYear' }))];
  return [
    { name: 'List_Month_Names', title: 'Month names', group, items: MONTH_NAMES.map(value => ({ value, style: 'lu.text' })) },
    { name: 'List_Start_Months', title: 'Months of the first financial year', group,
      items: Array.from({ length: 12 }, (_, k) => ({ value: `=DATE(Tl_First_FY-1,Sel_FY_End_Month+${k + 1},1)`, style: 'lu.monthYear' })) },
    { name: 'List_Months', title: 'Months in the model', group, items: months() },
    { name: 'List_Last_Actual', title: 'Last month of actuals', group, items: months(['No actuals']) },
    { name: 'List_Denominations', title: 'Denominations', group, items: DENOMINATIONS.map(value => ({ value, style: 'lu.text' })) },
    { name: 'List_Denom_Factors', title: 'Denomination factors', group, items: [1, 1000, 1000000].map(value => ({ value, style: 'lu.int' })) },
    ...(years ? [{ name: 'List_Years', title: 'Financial years in the model', group,
      items: Array.from({ length: years }, (_, k) => ({ value: `="FY"&(Tl_First_FY+${k})`, style: 'lu.text' })) }] : []),
  ];
}

/** Financial years a timeline touches. */
export function yearsIn(info: ModelInfo, periods: number): number {
  return Math.ceil((fiscalPosition(info.timeline.start, info.timeline.fyEndMonth).month - 1 + periods) / 12);
}

/** Lists the frame keeps, which module settings may also read. */
export const FRAME_LISTS = new Set(['List_Month_Names', 'List_Start_Months', 'List_Months', 'List_Last_Actual',
  'List_Denominations', 'List_Denom_Factors', 'List_Scenarios', 'List_Years']);

/** The Lookups sheet: a heading per group, then each list's title, its items and the List_ range over them. */
export function lookupRows(lists: ListSpec[], ranges: Map<string, RangeName>): LRow[] {
  const rows: LRow[] = [];
  const groups = [...new Set(lists.map(l => l.group))];
  for (const g of groups) {
    const gid = `lookups/${code(g)}`;
    rows.push(new LRow(`${gid}/heading`, 'heading', g));
    lists.filter(l => l.group === g).forEach((l, k) => {
      if (k) rows.push(new LRow(`lookups/${l.name}/gap`, 'blank', '', { space: 6 }));
      rows.push(new LRow(`lookups/${l.name}/title`, 'section', l.title, { indent: 1, cells: { [LIST_COL + 1]: l.name } }));
      rows.push(new LRow(`lookups/${l.name}/head`, 'item', `${l.title} heading`, { role: 'luHead',
        cells: { [LIST_COL - 1]: '#', [LIST_COL]: 'Item' } }));
      // a list nothing has joined yet holds one placeholder, so its drop-downs and range stay valid
      const items = l.items.length ? l.items : [{ value: '(none yet)', style: 'lu.text' }];
      items.forEach((it, i) => rows.push(new LRow(`lookups/${l.name}/${i + 1}`, 'item', `${l.title} ${i + 1}`, {
        role: it.style, cells: { [LIST_COL - 1]: i + 1, [LIST_COL]: it.value } })));
      ranges.set(l.name, { from: `lookups/${l.name}/1`, to: `lookups/${l.name}/${items.length}`, col: LIST_COL });
    });
    rows.push(new LRow(`${gid}/end`, 'blank', '', { space: 9, level: 0 }));
  }
  return rows;
}

function settingsRows(info: ModelInfo, periods: number): LRow[] {
  const t = info.timeline;
  const set = (key: string, label: string, init: ConstructorParameters<typeof LRow>[3]) =>
    new LRow(`settings/${key}`, 'setting', label, { indent: 1, ...init });
  const fixed = (key: string, label: string, init: ConstructorParameters<typeof LRow>[3]) =>
    new LRow(`settings/${key}`, 'fixed', label, { indent: 1, ...init });
  const textValid = { kind: 'text' as const, max: 120, message: 'Keep it under 120 characters.' };
  const drop = (list: string) => ({ role: 'cellLink', control: { kind: 'drop' as const, list },
    valid: { kind: 'whole' as const, min: 1, max: `ROWS(${list})`, message: 'Choose from the drop-down list.' } });
  const check = { role: 'cellLink', control: { kind: 'check' as const }, valid: { kind: 'logical' as const, message: 'Tick or clear the box.' } };
  const end = (key: string) => new LRow(`settings/${key}/end`, 'blank', '', { space: 9, level: 0 });
  const first = fiscalPosition(t.start, t.fyEndMonth);
  const budget = t.budget ?? defaultBudget(t.lastActual, periods);
  return [
    new LRow('settings/model/heading', 'heading', 'Model'),
    set('model/title', 'Model title', { value: info.title, name: 'Model_Title', role: 'in.text', valid: textValid }),
    fixed('model/entity', 'Entity', { name: 'Model_Entity', role: 'text', cells: { [TOTAL_COL]: info.entity.name } }),
    set('model/prepared', 'Prepared by line', { value: info.preparedBy, name: 'Model_Prepared_By', role: 'in.text', valid: textValid }),
    end('model'),
    new LRow('settings/time/heading', 'heading', 'Timeline'),
    fixed('time/periodicity', 'Periodicity', { role: 'text', cells: { [TOTAL_COL]: 'Monthly' } }),
    set('time/fy', 'Financial year ends in', { value: t.fyEndMonth, name: 'Sel_FY_End_Month', ...drop('List_Month_Names') }),
    set('time/first_fy', 'First financial year', { value: first.year, unit: 'year', name: 'Tl_First_FY', role: 'in.year',
      valid: { kind: 'whole', min: 1990, max: 2200, message: 'Type a year such as 2027: the year the first financial year ends in.' } }),
    set('time/start_month', 'First month of the model', { value: first.month, name: 'Sel_Start_Month', ...drop('List_Start_Months') }),
    fixed('time/term', 'Months in the model', { unit: 'months', name: 'Tl_Term', role: 'int', cells: { [TOTAL_COL]: periods } }),
    fixed('time/start', 'Model start date', { unit: 'date', name: 'Tl_Start', role: 'date',
      cells: { [TOTAL_COL]: '=DATE(Tl_First_FY-1,Sel_FY_End_Month+Sel_Start_Month,1)' } }),
    fixed('time/end_date', 'Model end date', { unit: 'date', name: 'Tl_End', role: 'date', cells: { [TOTAL_COL]: '=EOMONTH(Tl_Start,Tl_Term-1)' } }),
    fixed('time/years', 'Financial years in the model', { unit: '#', name: 'Tl_Years', role: 'int', cells: { [TOTAL_COL]:
      '=YEAR(Tl_End)+IF(MONTH(Tl_End)>Sel_FY_End_Month,1,0)-YEAR(Tl_Start)-IF(MONTH(Tl_Start)>Sel_FY_End_Month,1,0)+1' } }),
    set('time/denom', 'Denomination', { value: DENOMINATIONS.indexOf(t.denomination) + 1, name: 'Sel_Denom', ...drop('List_Denominations') }),
    fixed('time/denom_label', 'Denomination shown in titles', { name: 'Tl_Denom', role: 'text',
      cells: { [TOTAL_COL]: '=INDEX(List_Denominations,Sel_Denom)' } }),
    fixed('time/denom_factor', 'Denomination factor', { unit: '#', name: 'Tl_Denom_Factor', role: 'int',
      cells: { [TOTAL_COL]: '=INDEX(List_Denom_Factors,Sel_Denom)' } }),
    end('time'),
    new LRow('settings/actual/heading', 'heading', 'Actuals and forecast'),
    set('actual/last', 'Last month of actuals', { value: t.lastActual + 1, name: 'Sel_Last_Actual', ...drop('List_Last_Actual') }),
    fixed('actual/period', 'Last actual period', { unit: 'period', name: 'Tl_Last_Actual', role: 'int', cells: { [TOTAL_COL]: '=Sel_Last_Actual-1' } }),
    fixed('actual/date', 'Actuals to', { unit: 'date', name: 'Tl_Last_Actual_Date', role: 'date',
      cells: { [TOTAL_COL]: '=EOMONTH(Tl_Start,Tl_Last_Actual-1)' } }),
    fixed('actual/forecast_fy', 'Financial year of the first forecast month', { unit: 'year', name: 'Tl_First_Forecast_FY', role: 'year',
      cells: { [TOTAL_COL]: '=YEAR(EDATE(Tl_Start,Tl_Last_Actual))+IF(MONTH(EDATE(Tl_Start,Tl_Last_Actual))>Sel_FY_End_Month,1,0)' } }),
    set('actual/label', 'Label for actual months', { value: 'Actual', name: 'Tl_Actual_Label', role: 'in.text', valid: textValid }),
    set('actual/forecast_label', 'Label for forecast months', { value: 'Forecast', name: 'Tl_Forecast_Label', role: 'in.text', valid: textValid }),
    end('actual'),
    new LRow('settings/budget/heading', 'heading', 'Budget'),
    set('budget/first', 'First month of the budget', { value: budget.first, name: 'Sel_Budget_First', ...drop('List_Months') }),
    set('budget/months', 'Months in the budget', { value: budget.months, unit: 'months', name: 'Tl_Budget_Term', role: 'in.count',
      valid: { kind: 'whole', min: 1, max: 'Tl_Term-Sel_Budget_First+1', message: 'Type a number of months that ends inside the timeline.' } }),
    fixed('budget/start', 'Budget start date', { unit: 'date', name: 'Tl_Budget_Start', role: 'date',
      cells: { [TOTAL_COL]: '=EDATE(Tl_Start,Sel_Budget_First-1)' } }),
    fixed('budget/end_date', 'Budget end date', { unit: 'date', name: 'Tl_Budget_End', role: 'date',
      cells: { [TOTAL_COL]: '=EOMONTH(Tl_Budget_Start,Tl_Budget_Term-1)' } }),
    end('budget'),
    new LRow('settings/display/heading', 'heading', 'Display'),
    set('display/errors', 'Show the error count in the model name', { value: info.display.errors, name: 'Opt_Show_Errors', ...check }),
    set('display/alerts', 'Show the alert count in the model name', { value: info.display.alerts, name: 'Opt_Show_Alerts', ...check }),
    end('display'),
  ];
}

const HYPERLINK = /^=HYPERLINK\("#([^"]+)",(.*)\)$/;

/** `=HYPERLINK("#X",text)` written by shared code becomes the text with a cell hyperlink to X. */
function liftHyperlinks(row: LRow): void {
  for (const [c, v] of Object.entries(row.cells)) {
    if (typeof v !== 'string') continue;
    const m = HYPERLINK.exec(v);
    if (!m) continue;
    const literal = /^"([^"]*)"$/.exec(m[2]);
    row.cells[Number(c)] = literal ? literal[1] : `=${m[2]}`;
    row.links = { ...(row.links || {}), [Number(c)]: { to: m[1], tip: literal ? `Go to ${literal[1]}` : `Go to ${m[1]}` } };
  }
}

/**
 * The sheets the frame writes for module rows that ask for them: the historical income statement
 * and balance sheet (a typed line for each row that declares one, grouped as the library says,
 * with group totals) and the Scenarios sheet (the active scenario, the scenario names, and an
 * adjustment line for each row that declares one, by scenario).
 */
export function frameSheets(model: Model, sheets: Map<string, LRow[]>, names: Map<string, string>, req: FrameRequests): DataTable[] {
  const lib = model.lib;
  for (const side of ['is', 'bs'] as const) {
    const def = lib.history[side];
    const lines = req.history.filter(h => h.in === side);
    if (!lines.length) continue;
    if (!def) throw new AssemblyError(`modules declare historical ${side === 'is' ? 'income statement' : 'balance sheet'} lines, but the library has no sheet for them`);
    const groups: string[] = side === 'is' ? def.groups as string[] : (def.groups as { name: string }[]).map(g => g.name);
    for (const h of lines) if (!groups.includes(h.group)) throw new AssemblyError(`historical line ${h.label}: no group '${h.group}' in the library`);
    const rows: LRow[] = [new LRow(`hist/${side}/heading`, 'heading', def.title)];
    for (const g of groups) {
      const mine = lines.filter(h => h.group === g);
      if (!mine.length) continue;
      const gid = `hist/${side}/${code(g)}`;
      rows.push(new LRow(`${gid}/sp`, 'blank', '', { space: 6 }));
      rows.push(new LRow(`${gid}/section`, 'section', g, { indent: 1 }));
      mine.forEach(h => rows.push(new LRow(h.id, 'series', h.label, { indent: 2, unit: h.unit, input: 'actual', values: h.values,
        inactive: '«F14»=0', total: side === 'is' ? 'sum' : 'none', role: side === 'bs' ? 'opening' : undefined, value: h.opening })));
      if (side === 'is' && mine.length > 1) rows[rows.length - 1].role = 'last';
      if (mine.length > 1) {
        rows.push(new LRow(`${gid}/total`, 'series', `Total ${g.toLowerCase()}`, { indent: 1, unit: mine[0].unit, style: 'total',
          formula: `=SUM(«R|${mine[0].id}»:«R|${mine[mine.length - 1].id}»)`, total: side === 'is' ? 'sum' : 'none',
          cells: side === 'bs' ? { [TOTAL_COL]: `=SUM(«C${TOTAL_COL}|${mine[0].id}»:«C${TOTAL_COL}|${mine[mine.length - 1].id}»)` } : {} }));
      }
    }
    rows.push(new LRow(`hist/${side}/end`, 'blank', '', { space: 9, level: 0 }));
    sheets.set(def.sheet, rows);
  }

  if (!req.scenarios.length) {
    if (req.results.length) throw new AssemblyError('a report shows results by scenario, but no module has a scenario adjustment');
    return [];
  }
  const sc = lib.scenarios;
  if (!sc) throw new AssemblyError('modules declare scenario adjustments, but the library has no Scenarios sheet');
  const n = sc.names.length;
  const rows: LRow[] = [
    new LRow('scenarios/heading', 'heading', 'Scenarios'),
    new LRow('scenarios/active', 'setting', 'Active scenario', { indent: 1, value: 1, name: 'Sel_Scenario', role: 'cellLink',
      control: { kind: 'drop', list: 'List_Scenarios' },
      valid: { kind: 'whole', min: 1, max: 'ROWS(List_Scenarios)', message: 'Choose from the drop-down list.' } }),
    new LRow('scenarios/active_name', 'fixed', 'Active scenario name', { indent: 1, name: 'Scn_Active_Name', role: 'text',
      cells: { [TOTAL_COL]: '=INDEX(List_Scenarios,Sel_Scenario)' } }),
    ...sc.names.map((nm, k) => new LRow(`scenarios/name/${k + 1}`, 'setting', `Scenario ${k + 1} name`, { indent: 1, value: nm,
      role: 'in.text', valid: { kind: 'text', max: 40, message: 'Keep it under 40 characters.' } })),
    new LRow('scenarios/end', 'blank', '', { space: 9, level: 0 }),
    new LRow('scenarios/lines/heading', 'heading', 'Scenario adjustments'),
    new LRow('scenarios/lines/head', 'toc', 'columns', { style: 'bold', cells: { 3: 'Line', [UNIT_COL]: 'Unit', [TOTAL_COL]: 'Active',
      ...Object.fromEntries(sc.names.map((_, k) => [FIRST_PERIOD_COL + k, `=«V|scenarios/name/${k + 1}»`])) } }),
  ];
  names.set('Sel_Scenario', 'scenarios/active');
  names.set('Scn_Active_Name', 'scenarios/active_name');
  let block = '';
  for (const s of req.scenarios) {
    if (s.block !== block) {
      block = s.block;
      rows.push(new LRow(`scenarios/block/${code(s.block)}`, 'section', s.title, { indent: 1 }));
    }
    rows.push(new LRow(s.id, 'scenario', s.label, { indent: 2, unit: '%', input: 'all', values: s.values,
      cells: { [TOTAL_COL]: `=INDEX(«C${FIRST_PERIOD_COL}|${s.id}»:«C${FIRST_PERIOD_COL + n - 1}|${s.id}»,Sel_Scenario)` } }));
  }
  rows.push(new LRow('scenarios/lines/end', 'blank', '', { space: 9, level: 0 }));
  const tables: DataTable[] = [];
  if (req.results.length) {
    // Results for every scenario at once: a data table substitutes each scenario's number into
    // Sel_Scenario and works out the formulas in column I, so report charts can show all three.
    rows.push(new LRow('scenarios/results/heading', 'heading', 'Scenario results'));
    rows.push(new LRow('scenarios/results/note', 'text', 'A data table: each column works the model out under that scenario. '
      + 'Excel recalculates it with the model (unless calculation is set to automatic except tables).', { indent: 1, role: 'note' }));
    rows.push(new LRow('scenarios/results/input', 'table', 'Scenario number', { indent: 2, role: 'r.head',
      cells: Object.fromEntries(sc.names.map((_, k) => [FIRST_PERIOD_COL + k, k + 1])) }));
    rows.push(new LRow('scenarios/results/names', 'table', 'Scenario', { indent: 2, role: 'r.head',
      cells: Object.fromEntries(sc.names.map((_, k) => [FIRST_PERIOD_COL + k, `=«V|scenarios/name/${k + 1}»`])) }));
    // The input row must sit directly above the results, so the names go above it.
    const [inputRow, namesRow] = rows.splice(rows.length - 2, 2);
    rows.push(namesRow, inputRow);
    for (const r of req.results) {
      rows.push(new LRow(r.id, 'table', r.label, { indent: 2, unit: '$', role: 'r.result',
        cells: { [TOTAL_COL]: r.formula, ...Object.fromEntries(sc.names.map((_, k) => [FIRST_PERIOD_COL + k, null])) } }));
    }
    rows.push(new LRow('scenarios/results/end', 'blank', '', { space: 9, level: 0 }));
    tables.push({ sheet: sc.sheet, head: 'scenarios/results/input', first: req.results[0].id, last: req.results[req.results.length - 1].id,
      cols: n, input: 'Sel_Scenario' });
  }
  sheets.set(sc.sheet, rows);
  return tables;
}

export function navigateStandard(layout: Layout, model: Model, blocks: Block[], sheets: Map<string, LRow[]>, moduleLists: ListSpec[] = []): void {
  const info = model.info!;
  const lib = model.lib;
  layout.frame = STANDARD_FRAME;
  sheets.set(SETTINGS, settingsRows(info, model.periods));
  const reporting = blocks.some(b => b.mod.framework === 'report');
  const lists = [...timelineLists(model.periods, reporting ? yearsIn(info, model.periods) : 0), ...moduleLists];
  if (lib.history.bs && sheets.has(lib.history.bs.sheet)) layout.totalHeads[lib.history.bs.sheet] = 'Opening';
  if (lib.scenarios && sheets.has(lib.scenarios.sheet)) {
    lists.push({ name: 'List_Scenarios', title: 'Scenario names', group: 'Scenario lists',
      items: lib.scenarios.names.map((_, k) => ({ value: `=«V|scenarios/name/${k + 1}»`, style: 'lu.text' })) });
  }
  sheets.set(LOOKUPS, lookupRows(lists, layout.ranges));
  for (const [, rows] of sheets) rows.forEach(liftHyperlinks);

  const sections: SectionDef[] = lib.sections.length ? lib.sections
    : [{ title: 'Model', cover: null, note: '', areas: lib.areas }];
  const home = sections.find(s => s.areas.includes('Checks')) ?? sections[sections.length - 1];
  const present = sections
    .map(sec => [sec, [...(sec === home ? [SETTINGS, LOOKUPS] : []), ...sec.areas.filter(a => sheets.has(a))]] as [SectionDef, string[]])
    .filter(([, areas]) => areas.length);
  const order: string[] = [CONTENTS];
  for (const [sec, areas] of present) {
    if (sec.cover) order.push(sec.cover);
    order.push(...areas);
  }
  // A sheet that holds only summary and report modules has no timeline: selections, a chart grid and tables.
  const plain = (b: Block) => b.mod.framework === 'report' || b.mod.framework === 'versions' || b.mod.framework === 'table';
  const reportAreas = new Set(blocks.filter(plain).map(b => b.mod.area));
  for (const b of blocks) if (!plain(b)) reportAreas.delete(b.mod.area);
  layout.kinds = { [CONTENTS]: 'contents' };
  layout.titles = { [CONTENTS]: info.entity.name };
  for (const [sec, areas] of present) {
    if (sec.cover) {
      layout.kinds[sec.cover] = 'cover';
      layout.titles[sec.cover] = sec.title;
    }
    for (const a of areas) {
      layout.kinds[a] = a === SETTINGS ? 'settings' : a === LOOKUPS ? 'lookups' : a === lib.scenarios?.sheet ? 'scenarios'
        : Object.hasOwn(FRAME_KINDS, a) ? FRAME_KINDS[a] : reportAreas.has(a) ? 'report' : 'timeline';
      layout.titles[a] = a;
    }
  }

  // The blocks each sheet holds, for the contents.
  const headings = new Map<string, [string, string][]>();
  const add = (sheet: string, uid: string, title: string) => {
    const list = headings.get(sheet);
    if (list) list.push([uid, title]);
    else headings.set(sheet, [[uid, title]]);
  };
  const seen = new Set<string>();
  for (const b of blocks) {
    if (seen.has(b.inst.uid)) continue;
    seen.add(b.inst.uid);
    add(b.mod.area, b.inst.uid, model.title(b.inst));
  }
  for (const sh of [SETTINGS, LOOKUPS]) for (const r of sheets.get(sh)!) if (r.kind === 'heading') add(sh, r.id.replace(/\/heading$/, ''), r.label);

  const names = layout.names;
  const nameAt = (nm: string, target: string, col: number) => {
    names.set(nm, target);
    layout.nameCols[nm] = col;
  };
  nameAt(GO_CONTENTS, `@${CONTENTS}`, 2);
  if (order.includes('Checks')) nameAt(GO_CHECKS, '@Checks', 2);
  for (const sh of order) nameAt(goSheet(sh), `@${sh}`, 2);
  nameAt('Model_Name', `@${CONTENTS}/2`, 2);
  for (const r of sheets.get(SETTINGS)!) if (r.name) names.set(r.name, r.id);

  const link = (to: string, tip: string): CellLink => ({ to, tip });
  const contents: LRow[] = [new LRow('contents/notes/heading', 'heading', 'Notes')];
  const notes = info.notes.length ? info.notes : ['No notes yet.'];
  notes.forEach((n, i) => contents.push(new LRow(`contents/note/${i + 1}`, 'toc', `note ${i + 1}`, {
    role: 'note', cells: { 2: info.notes.length ? `${i + 1}.` : '', 3: n } })));
  contents.push(new LRow('contents/notes/end', 'blank', '', { space: 9, level: 0 }));
  contents.push(new LRow('contents/title', 'heading', 'Table of contents'));
  contents.push(new LRow('contents/gap', 'blank', '', { space: 6, level: 0 }));
  present.forEach(([sec, areas], k0) => {
    const n = k0 + 1;
    const cover = sec.cover;
    if (cover) {
      contents.push(new LRow(`contents/section/${cover}`, 'toc', `section ${n} ${sec.title}`, {
        role: 'toc1', level: 0, cells: { 2: `${n}.`, 3: `=«S|${cover}»` },
        links: { 2: link(goSheet(cover), `Go to ${sec.title}`), 3: link(goSheet(cover), `Go to ${sec.title}`) } }));
    }
    areas.forEach((a, k) => {
      const letter = String.fromCharCode(97 + k);
      contents.push(new LRow(`contents/sheet/${a}`, 'toc', `${letter}. ${a}`, {
        role: 'toc2', level: 1, cells: { 3: `${letter}.`, 4: `=«S|${a}»` },
        links: { 3: link(goSheet(a), `Go to ${a}`), 4: link(goSheet(a), `Go to ${a}`) } }));
      for (const [uid, title] of headings.get(a) || []) {
        const nm = goBlock(uid);
        nameAt(nm, `${uid}/heading`, 2);
        contents.push(new LRow(`contents/module/${uid}`, 'toc', `- ${title}`, {
          role: 'toc3', level: 2, cells: { 4: '-', 5: `=«B|${uid}/heading»` },
          links: { 4: link(nm, `Go to ${title}`), 5: link(nm, `Go to ${title}`) } }));
      }
    });
    contents.push(new LRow(`contents/section/${sec.cover ?? n}/end`, 'blank', '', { space: 6, level: 0 }));
  });
  if (layout.hasChecks()) {
    contents.push(new LRow('contents/checks', 'heading', 'Checks'));
    contents.push(new LRow('contents/errors', 'toc', 'error checks', { style: 'check', role: 'link', cells: {
      3: 'Error checks failing', [TOTAL_COL]: '=Chk_Errors' }, links: { 3: link(GO_CHECKS, 'Go to the checks') } }));
    if (names.has('Chk_Alerts')) {
      contents.push(new LRow('contents/alerts', 'toc', 'alerts', { style: 'check', role: 'link', cells: {
        3: 'Alerts raised', [TOTAL_COL]: '=Chk_Alerts' }, links: { 3: link(GO_CHECKS, 'Go to the checks') } }));
    }
    contents.push(new LRow('contents/checks/end', 'blank', '', { space: 9, level: 0 }));
  }
  if (model.assured() && layout.headlines.length) {
    contents.push(new LRow('contents/ko', 'heading', 'Key outputs'));
    for (const h of layout.headlines) {
      const rid = `contents/ko/${h.name}`;
      contents.push(new LRow(rid, 'toc', h.label, { unit: h.unit, role: 'text', cells: { 3: h.label, [TOTAL_COL]: keyOutputFormula(h) } }));
      names.set(h.name, rid);
    }
  }

  const covers = present.map(([sec]) => sec.cover);
  const out: [string, LRow[]][] = [[CONTENTS, contents]];
  for (let i = 1; i < order.length; i++) {
    const sh = order[i];
    if (layout.kinds[sh] !== 'cover') {
      out.push([sh, sheets.get(sh)!]);
      continue;
    }
    const n = covers.indexOf(sh) + 1;
    const sec = present[n - 1][0];
    const prev = order[i - 1];
    const nxt = order[i + 1];
    if (nxt === undefined) throw new AssemblyError(`${sh}: a section cover cannot be the last sheet`);
    out.push([sh, [
      new LRow(`cover/${sh}/number`, 'toc', 'section number', { role: 'sectionNo', cells: { 2: `Section ${n}.` } }),
      new LRow(`cover/${sh}/gap`, 'blank', '', { space: 6 }),
      new LRow(`cover/${sh}/home`, 'toc', 'link to the contents', { role: 'linkU', cells: { 2: 'Go to contents' },
        links: { 2: link(GO_CONTENTS, 'Go to the table of contents') } }),
      new LRow(`cover/${sh}/prev`, 'toc', 'link to the previous sheet', { role: 'link',
        cells: { 2: `="${SYMBOL.prev} "&«S|${prev}»` }, links: { 2: link(goSheet(prev), prev === CONTENTS ? 'Go to the table of contents' : `Go to ${prev}`) } }),
      new LRow(`cover/${sh}/next`, 'toc', 'link to the next sheet', { role: 'link',
        cells: { 2: `=«S|${nxt}»&" ${SYMBOL.next}"` }, links: { 2: link(goSheet(nxt), `Go to ${nxt}`) } }),
      new LRow(`cover/${sh}/gap2`, 'blank', '', { space: 9 }),
      new LRow(`cover/${sh}/notes_title`, 'toc', 'notes heading', { role: 'h3', cells: { 2: 'Section notes' } }),
      new LRow(`cover/${sh}/notes`, 'toc', 'notes', { role: 'note', cells: { 2: pick(sec, 'note', '') } }),
    ]]);
  }
  layout.sheets = out;
  const ids = new Set<string>();
  for (const [sh, rows] of out) {
    for (const r of rows) {
      if (ids.has(r.id)) throw new AssemblyError(`${sh}: row id ${r.id} is used twice`);
      ids.add(r.id);
    }
  }
}

/** The model name line: the title, then the error and alert counts when they are not clear and their switch is on. */
export function modelNameFormula(layout: Layout): string {
  let f = '=Model_Title';
  if (layout.names.has('Sel_Scenario')) f += '&IF(Sel_Scenario>1," ("&Scn_Active_Name&" scenario)","")';
  if (layout.names.has('Chk_Errors')) {
    f += '&IF(AND(Opt_Show_Errors,Chk_Errors>0)," ("&Chk_Errors&IF(Chk_Errors=1," error)"," errors)"),"")';
  }
  if (layout.names.has('Chk_Alerts')) {
    f += '&IF(AND(Opt_Show_Alerts,Chk_Alerts>0)," ("&Chk_Alerts&IF(Chk_Alerts=1," alert)"," alerts)"),"")';
  }
  return f;
}

/** Header cells of a sheet in the standard frame, in the order written. */
export function standardFrameCells(layout: Layout, sheet: string, dialect: Dialect = 'excel'): FrameCell[] {
  const kind = layout.kindOf(sheet);
  const cells: FrameCell[] = [];
  if (kind === 'contents') {
    cells.push([STD.titleRow, 2, '=Model_Entity'], [STD.nameRow, 2, modelNameFormula(layout)], [STD.entityRow, 2, '=Model_Prepared_By']);
  } else {
    cells.push([STD.titleRow, 1, SYMBOL.home]);
    if (layout.hasChecks()) cells.push([STD.nameRow, 1, `=IF(Chk_Errors=0,"${SYMBOL.clear}","${SYMBOL.failing}")`]);
    cells.push([STD.titleRow, 2, Object.hasOwn(layout.titles, sheet) ? layout.titles[sheet] : sheet],
      [STD.nameRow, 2, '=Model_Name'], [STD.entityRow, 2, '=Model_Entity']);
  }
  if (kind === 'timeline') {
    cells.push([5, TOTAL_COL, layout.totalHeads[sheet] ?? 'Total']);
    for (const b of BLOCK) {
      cells.push([b.row, 2, b.label]);
      for (let p = 0; p < layout.periods; p++) {
        const c = FIRST_PERIOD_COL + p;
        cells.push([b.row, c, b.source(c, colLetter(c), p ? colLetter(c - 1) : null)]);
      }
    }
  }
  return dialect === 'uno'
    ? cells.map(([r, c, v]) => [r, c, typeof v === 'string' && v.startsWith('=') ? unoSeparators(v) : v])
    : cells;
}

/** The hyperlinks in a sheet's header: A1 to the contents and A2 to the checks. */
export function standardFrameLinks(layout: Layout, sheet: string): { row: number; col: number; link: CellLink }[] {
  if (layout.kindOf(sheet) === 'contents') return [];
  const out: { row: number; col: number; link: CellLink }[] = [
    { row: STD.titleRow, col: 1, link: { to: GO_CONTENTS, tip: 'Go to the table of contents' } }];
  if (layout.hasChecks() && layout.names.has(GO_CHECKS)) {
    out.push({ row: STD.nameRow, col: 1, link: { to: GO_CHECKS, tip: 'Go to the checks' } });
  }
  return out;
}
