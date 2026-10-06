// Model assurance sheets: Group assumptions (the set the model uses, by item) and the Input
// register (every named input with its source, owner, date, evidence and status). Both are
// rebuilt from the model and its metadata, so a structural change rewrites them through the
// ordinary plan; values typed into the register are read back into the metadata before each change.

import { groupName, pick, TOTAL_COL, UNIT_COL } from './frame.ts';
import { LRow, type Cells, type Headline, type SheetKind } from './layout.ts';
import type { Model } from './model.ts';

export const GROUP_SHEET = 'Group assumptions';
export const REGISTER_SHEET = 'Input register';
export const FRAME_KINDS: Record<string, SheetKind> = { [GROUP_SHEET]: 'list', [REGISTER_SHEET]: 'register' };
/** 1 October 2026 as an Excel date. */
export const REGISTER_DATE = 46296;
/** Months before an input counts as past its review age. */
export const REVIEW_AGE = 12;
/** Register columns: input (B), unit (H), value (I), then these. */
export const RC = { source: 10, owner: 11, updated: 12, evidence: 13, age: 14, group: 15, reason: 16, status: 17 } as const;
export const REGISTER_HEADINGS: Cells = {
  2: 'Input', [UNIT_COL]: 'Unit', [TOTAL_COL]: 'Value', [RC.source]: 'Source', [RC.owner]: 'Owner',
  [RC.updated]: 'Updated', [RC.evidence]: 'Evidence', [RC.age]: 'Age (months)', [RC.group]: 'Group assumption',
  [RC.reason]: 'Reason for a local value', [RC.status]: 'Status',
};

/** One named input, as assemble finds it, for the register. */
export interface InputEntry {
  name: string;
  id: string;
  label: string;
  title: string;
  unit: string;
  /** The group item it is bound to, or false. */
  bound: string | false;
  /** The group item its module names for it, bound or not. */
  item: string | null;
}

interface GroupItem { key: string; label: string; value: unknown; unit?: string; source?: string }
interface GroupSet { title?: string; version?: number; published?: number; owner?: string; items?: GroupItem[] }

export function assuranceSheets(model: Model, sheets: Map<string, LRow[]>, names: Map<string, string>, inputs: InputEntry[]): void {
  const a = model.assurance;
  const gset: GroupSet = (pick(a, 'set', null) as GroupSet | null) || {};
  const items = new Map((pick(gset, 'items', []) as GroupItem[]).map(i => [i.key, i]));
  const used = new Map<string, number>();
  for (const i of inputs) if (i.bound) used.set(i.bound, (used.get(i.bound) || 0) + 1);

  const g = [new LRow('ga/heading', 'heading', 'Group assumptions')];
  const head: [string, string, unknown, string, string | null][] = [
    ['set', 'Set', pick(gset, 'title', 'No group set in use'), '', null],
    ['version', 'Version', pick(gset, 'version', 0), '', 'GA_Version'],
    ['published', 'Published', pick(gset, 'published', 0), 'date', 'GA_Published'],
    ['owner', 'Owner', pick(gset, 'owner', ''), '', null],
    ['latest', 'Latest version published', pick(a, 'latest', pick(gset, 'version', 0)), '', 'GA_Latest'],
  ];
  for (const [key, label, value, unit, nm] of head) {
    g.push(new LRow(`ga/${key}`, 'fixed', label, { indent: 1, unit, name: nm, cells: { [TOTAL_COL]: value } }));
    if (nm) names.set(nm, `ga/${key}`);
  }
  g.push(new LRow('ga/gap', 'blank', ''));
  if (items.size) {
    g.push(new LRow('ga/items', 'toc', 'columns', { style: 'bold', cells: {
      3: 'Item', [UNIT_COL]: 'Unit', [TOTAL_COL]: 'Value', [RC.source]: 'Source', [RC.owner]: 'Used by' } }));
    for (const [key, it] of items) {
      const nm = groupName(key);
      const n = used.get(key) || 0;
      g.push(new LRow(`ga/item/${key}`, 'fixed', it.label, { indent: 1, unit: pick(it, 'unit', ''), name: nm, style: 'rate',
        cells: { [TOTAL_COL]: it.value, [RC.source]: pick(it, 'source', ''),
          [RC.owner]: n ? `used by ${n} input${n !== 1 ? 's' : ''}` : 'not used' } }));
      names.set(nm, `ga/item/${key}`);
    }
  }
  sheets.set(GROUP_SHEET, g);

  const recs = pick(a, 'inputs', {}) as Record<string, Record<string, unknown>>;
  const r = [
    new LRow('register/heading', 'heading', 'Input register'),
    new LRow('register/date', 'setting', 'Register date', { indent: 1, unit: 'date',
      value: pick(a, 'register_date', REGISTER_DATE), name: 'Reg_Date' }),
    new LRow('register/age', 'setting', 'Review age', { indent: 1, unit: 'months',
      value: pick(a, 'review_age', REVIEW_AGE), name: 'Reg_MaxAge' }),
    new LRow('register/gap', 'blank', ''),
    new LRow('register/columns', 'toc', 'columns', { style: 'bold', cells: { ...REGISTER_HEADINGS } }),
  ];
  names.set('Reg_Date', 'register/date');
  names.set('Reg_MaxAge', 'register/age');
  const rowsIn: LRow[] = [];
  for (const i of inputs) {
    const rid = `register/in/${i.name}`;
    const rec = pick(recs, i.name, {} as Record<string, unknown>);
    const me = (c: number) => `«C${c}|${rid}»`;
    const cells: Cells = { 2: `=HYPERLINK("#${i.name}","${i.title}: ${i.label}")`, [TOTAL_COL]: `=${i.name}` };
    if (i.bound) {
      cells[RC.source] = `Group assumptions, version ${pick(gset, 'version', '')}`;
      cells[RC.owner] = pick(gset, 'owner', '');
      cells[RC.updated] = '=GA_Published';
      const bound = items.get(i.bound);
      cells[RC.group] = bound ? pick(bound, 'label', i.bound) : i.bound;
      cells[RC.evidence] = cells[RC.reason] = null;
    } else {   // typed columns are always written (blank when empty), so clearing one in the pane clears the cell
      for (const k of ['source', 'owner', 'updated', 'evidence'] as const) {
        const v = pick(rec, k, null);
        cells[RC[k]] = v === '' ? null : v;
      }
      const local = Boolean(i.item && items.has(i.item));
      cells[RC.group] = local ? `Local, in place of ${items.get(i.item!)!.label}` : null;
      cells[RC.reason] = local ? pick(rec, 'reason', null) : null;
    }
    cells[RC.age] = `=IF(${me(RC.updated)}="","",(YEAR(Reg_Date)-YEAR(${me(RC.updated)}))*12`
      + `+MONTH(Reg_Date)-MONTH(${me(RC.updated)}))`;
    cells[RC.status] = `=IF(AND(LEFT(${me(RC.group)},5)="Local",${me(RC.reason)}=""),"Local, no reason",`
      + `IF(${me(RC.source)}="","No source",`
      + `IF(AND(${me(RC.age)}<>"",${me(RC.age)}>Reg_MaxAge),"Past review age",`
      + `IF(${me(RC.group)}="","OK",IF(LEFT(${me(RC.group)},5)="Local","Local","Group")))))`;
    rowsIn.push(new LRow(rid, 'toc', i.label, { unit: i.unit, style: 'reg', cells }));
  }
  r.push(...rowsIn);
  r.push(new LRow('register/gap2', 'blank', ''));
  const counts: [string, string, string, string][] = [
    ['nosource', 'Inputs with no source', 'No source', 'Reg_NoSource'],
    ['stale', 'Inputs past their review age', 'Past review age', 'Reg_Stale'],
    ['local', 'Local values in place of a group assumption, with no reason', 'Local, no reason', 'Reg_Local'],
    ['group', 'Inputs drawn from the group assumptions', 'Group', 'Reg_Group'],
  ];
  for (const [key, label, word, nm] of counts) {
    const f = rowsIn.length
      ? `=COUNTIF(«C${RC.status}|${rowsIn[0].id}»:«C${RC.status}|${rowsIn[rowsIn.length - 1].id}»,"${word}")`
      : '=0';
    r.push(new LRow(`register/count/${key}`, 'toc', label, { style: 'bold', name: nm, cells: { 2: label, [TOTAL_COL]: f } }));
    names.set(nm, `register/count/${key}`);
  }
  sheets.set(REGISTER_SHEET, r);
}

/** A key output's formula on the contents: the row's total, least, greatest or last period. */
export function keyOutputFormula(h: Headline): string {
  const rng = `«A|${h.id}»`;
  const by: Record<string, string> = { sum: `=SUM(${rng})`, min: `=MIN(${rng})`, max: `=MAX(${rng})` };
  return Object.hasOwn(by, h.measure) ? by[h.measure] : `=INDEX(${rng},1,«T»)`;
}
