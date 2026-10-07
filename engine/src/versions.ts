// Saved versions (framework: versions): every budget and reforecast saved as values and kept. The
// Versions sheet holds the register, one row per version; the Version store holds the values, one
// row per version and line, keyed "id|line" so a version saved before a line was added still reads
// correctly. Save version writes both; reports read them through the Compared with drop-down.

import { AssemblyError, FIRST_PERIOD_COL, LABEL_COLS, TOTAL_COL } from './frame.ts';
import { LRow, type RangeName } from './layout.ts';
import type { ReportSpec } from './library.ts';
import type { ListSpec } from './standard.ts';

/** One saved version, as Save version writes it (and a recipe brings it). */
export interface VersionData {
  id: string;
  type: 'Budget' | 'Reforecast' | 'Other';
  label: string;
  /** A budget's financial year, as its label in List_Years (FY2027). */
  year?: string;
  /** A reforecast's last actual month, as a period number. */
  asAt?: number;
  status: string;
  locked?: boolean;
  source?: string;
  savedBy?: string;
  /** "yyyy-mm-dd". */
  savedOn?: string;
  /** Each value times its sheet column, added up when saved. */
  checksum: number;
  /** Report line key -> values by period (null where the version has none). */
  values: Record<string, (number | null)[]>;
}

/** Register columns. */
export const REG = {
  id: TOTAL_COL, type: 10, year: 11, asAt: 12, asAtLabel: 13, status: 14, locked: 15, source: 16, savedBy: 17, savedOn: 18,
  first: 19, last: 20, saved: 21, now: 22, budgetKey: 23, refKey: 24, refAt: 25, changed: 26,
} as const;

/** The style each register column takes. */
export const REG_STYLES: Record<number, string> = {
  [REG.id]: 'text', [REG.type]: 'text', [REG.year]: 'text', [REG.asAt]: 'count', [REG.asAtLabel]: 'text', [REG.status]: 'text',
  [REG.locked]: 'text', [REG.source]: 'text', [REG.savedBy]: 'text', [REG.savedOn]: 'date', [REG.first]: 'count', [REG.last]: 'count',
  [REG.saved]: 'num', [REG.now]: 'num', [REG.budgetKey]: 'muted', [REG.refKey]: 'muted', [REG.refAt]: 'mutedNum', [REG.changed]: 'check',
};

/** The lines a budget keeps (the income statement), in order; reforecasts keep these plus cash and net assets. */
export const VERSION_LINES = ['rev', 'cogs', 'gm', 'other_income', 'staff', 'opex', 'opcosts', 'other_expense', 'ebitda', 'da', 'ebit',
  'interest', 'npbt', 'tax', 'npat'];

/** Lines where spending less than the comparison is favourable. */
export const COST_LINES = new Set(['cogs', 'staff', 'opex', 'opcosts', 'other_expense', 'da', 'interest', 'tax']);

/** Register column widths (characters). */
export const REG_WIDTHS: Record<number, number> = {
  [REG.type]: 10, [REG.year]: 8, [REG.asAt]: 6, [REG.asAtLabel]: 10, [REG.status]: 11, [REG.locked]: 7, [REG.source]: 13,
  [REG.savedBy]: 15, [REG.savedOn]: 12, [REG.first]: 8, [REG.last]: 8, [REG.saved]: 17, [REG.now]: 17, [REG.budgetKey]: 25,
  [REG.refKey]: 15, [REG.refAt]: 11, [REG.changed]: 8,
};

/** The choices before the saved versions in a Compared with drop-down. */
export const COMPARE_CHOICES = ['Approved budget for the year shown', "Last month's reforecast", 'Latest reforecast', 'Budget being built'];

export interface VersionsOut {
  register: LRow[];
  store: LRow[];
  ranges: Map<string, RangeName>;
  lists: ListSpec[];
}

const excelDate = (iso: string) => {
  const [y, m, d] = iso.split('-').map(Number);
  return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
};

/** The register and store rows for a versions block, and the names reports read. */
export function versionRows(block: string, versions: VersionData[], spec: ReportSpec, periods: number): VersionsOut {
  const reg: LRow[] = [];
  const store: LRow[] = [];
  const ranges = new Map<string, RangeName>();
  const head: Record<number, unknown> = {
    [REG.id]: 'Id', [REG.type]: 'Type', [REG.year]: 'Year', [REG.asAt]: 'As at', [REG.asAtLabel]: 'Actuals to', [REG.status]: 'Status',
    [REG.locked]: 'Locked', [REG.source]: 'Source', [REG.savedBy]: 'Saved by', [REG.savedOn]: 'Saved on', [REG.first]: 'First row',
    [REG.last]: 'Last row', [REG.saved]: 'Checksum saved', [REG.now]: 'Checksum now', [REG.budgetKey]: 'Budget key',
    [REG.refKey]: 'Reforecast key', [REG.refAt]: 'Reforecast at', [REG.changed]: 'Changed',
  };
  reg.push(new LRow(`${block}/reg/sp`, 'blank', '', { space: 6 }));
  reg.push(new LRow(`${block}/reg/section`, 'section', 'Saved versions', { indent: 1 }));
  reg.push(new LRow(`${block}/reg/head`, 'table', 'Version', { indent: 2, role: 'r.head', cells: head }));
  store.push(new LRow('store/heading', 'heading', 'Version store'));
  store.push(new LRow('store/note', 'text', 'Values written by Save version, one row per version and line. Do not type here: '
    + 'a changed value fails the checksum check.', { indent: 1, role: 'note' }));
  let k = 0;   // rows in the store range so far
  const storeFirst = 'store/first';
  store.push(new LRow(storeFirst, 'blank', '', { space: 6 }));
  k += 1;
  const seen = new Set<string>();
  for (const v of versions) {
    if (seen.has(v.id)) throw new AssemblyError(`two saved versions have the id ${v.id}`);
    seen.add(v.id);
    const sec = `store/${v.id}`;
    store.push(new LRow(sec, 'section', v.label, { indent: 1 }));
    k += 1;
    const first = k + 1;
    for (const [key, vals] of Object.entries(v.values)) {
      const line = spec.lines[key];
      if (!line) throw new AssemblyError(`version ${v.id}: the statements have no report line '${key}'`);
      const cells: Record<number, unknown> = { [TOTAL_COL]: `${v.id}|${key}` };
      for (let t = 0; t < periods; t++) cells[FIRST_PERIOD_COL + t] = vals[t] ?? null;
      store.push(new LRow(`${sec}/${key}`, 'table', line.label, { indent: 2, unit: '$', role: 'r.store', cells }));
      k += 1;
    }
    const last = k;
    const rid = `${block}/reg/${v.id}`;
    const c = (col: number) => `«C${col}|${rid}»`;
    const blk = `INDEX(VS_Values,${c(REG.first)},0):INDEX(VS_Values,${c(REG.last)},0)`;
    reg.push(new LRow(rid, 'table', v.label, { indent: 2, role: 'r.reg', cells: {
      [REG.id]: v.id, [REG.type]: v.type, [REG.year]: v.year ?? '', [REG.asAt]: v.asAt ?? null,
      [REG.asAtLabel]: `=IF(${c(REG.asAt)}="","",TEXT(INDEX(List_Months,${c(REG.asAt)}),"mmm yyyy"))`,
      [REG.status]: v.status, [REG.locked]: v.locked ? 'Yes' : 'No', [REG.source]: v.source ?? '', [REG.savedBy]: v.savedBy ?? '',
      [REG.savedOn]: v.savedOn ? excelDate(v.savedOn) : null, [REG.first]: first, [REG.last]: last, [REG.saved]: v.checksum,
      [REG.now]: `=SUMPRODUCT(${blk}*COLUMN(${blk}))`,
      [REG.budgetKey]: `=IF(${c(REG.type)}="Budget",${c(REG.type)}&"|"&${c(REG.status)}&"|"&${c(REG.year)},"")`,
      [REG.refKey]: `=IF(${c(REG.type)}="Reforecast","Reforecast|"&${c(REG.asAt)},"")`,
      [REG.refAt]: `=IF(${c(REG.type)}="Reforecast",${c(REG.asAt)},0)`,
      [REG.changed]: `=IF(ABS(${c(REG.now)}-${c(REG.saved)})>0.01,1,0)`,
    } }));
  }
  const storeLast = `store/end`;
  store.push(new LRow(storeLast, 'blank', '', { space: 9, level: 0 }));
  ranges.set('VS_Keys', { from: storeFirst, to: storeLast, col: TOTAL_COL });
  ranges.set('VS_Values', { from: storeFirst, to: storeLast, col: FIRST_PERIOD_COL, toCol: 'timeline' });
  {
    // with nothing saved yet, the names cover the heading row, so lookups find nothing rather than fail
    const a = versions.length ? `${block}/reg/${versions[0].id}` : `${block}/reg/head`;
    const b = versions.length ? `${block}/reg/${versions[versions.length - 1].id}` : `${block}/reg/head`;
    const col = (name: string, c: number) => ranges.set(name, { from: a, to: b, col: c });
    col('VR_Label', LABEL_COLS[2]);
    col('VR_Id', REG.id);
    col('VR_BudgetKey', REG.budgetKey);
    col('VR_RefKey', REG.refKey);
    col('VR_RefAt', REG.refAt);
    col('VR_Changed', REG.changed);
    col('VR_BudgetStatus', REG.status);
  }
  reg.push(new LRow(`${block}/reg/end`, 'blank', '', { space: 6 }));
  const list: ListSpec = { name: 'List_Compare', title: 'Compared with', group: 'Saved versions', items: [
    ...COMPARE_CHOICES.map(value => ({ value, style: 'lu.text' })),
    ...versions.map(v => ({ value: `=«B|${block}/reg/${v.id}»`.replace('«B|', `«C${LABEL_COLS[2]}|`), style: 'lu.text' })),
  ] };
  const keys = VERSION_LINES.filter(k => spec.lines[k]);
  const lines: ListSpec = { name: 'List_Version_Lines', title: 'Lines a version keeps', group: 'Saved versions',
    items: keys.map(k => ({ value: spec.lines[k].label, style: 'lu.text' })) };
  const keyList: ListSpec = { name: 'List_Version_Keys', title: 'Their keys in the Version store', group: 'Saved versions',
    items: keys.map(k => ({ value: k, style: 'lu.text' })) };
  return { register: reg, store, ranges, lists: [list, lines, keyList] };
}
