// Saved versions in the add-in: Save version (this month's reforecast, or the budget being built),
// Approve budget and Lock or unlock. Each is a change to the versions module's data, planned and
// applied like an insert, so the register, the Version store, the Compared with list and the names
// move together. DOM free; view.ts draws it.

import {
  assemble, FIRST_PERIOD_COL, metadataXml, planChange, type Layout, type Model, type VersionData,
} from '../../../engine/src/index.ts';
import type { InsertPlan, OpenModel } from '../insert/core.ts';

/** The lines each kind of version keeps: budgets the income statement, reforecasts that plus cash and net assets. */
export const BUDGET_LINES = ['rev', 'cogs', 'gm', 'other_income', 'staff', 'opex', 'opcosts', 'other_expense', 'ebitda', 'da', 'ebit',
  'interest', 'npbt', 'tax', 'npat'];
export const REFORECAST_LINES = [...BUDGET_LINES, 'cash', 'na'];

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

export interface SaveChoice {
  type: 'Reforecast' | 'Budget';
  label: string;
  /** A budget's year label (FY2027) and its months (period numbers); a reforecast's last actual month. */
  year?: string;
  months?: [number, number];
  asAt?: number;
  /** The saved version this one replaces (this month's reforecast saved again), if any. */
  replaces?: string;
}

export function versionsOf(model: Model): VersionData[] {
  return model.instances.find(i => model.lib.module(i.module).framework === 'versions')?.data.versions ?? [];
}

/** "Sep 2026" for a period of the model. */
export function monthLabel(model: Model, p: number): string {
  const [y, m] = model.info!.timeline.start.split('-').map(Number);
  const k = m - 1 + p - 1;
  return `${MONTHS[k % 12]} ${y + Math.floor(k / 12)}`;
}

/** The financial year label (FY2027) of a period. */
export function fyLabel(model: Model, p: number): string {
  const [y, m] = model.info!.timeline.start.split('-').map(Number);
  const k = m - 1 + p - 1;
  const cy = y + Math.floor(k / 12);
  const cm = (k % 12) + 1;
  return `FY${cy + (cm > model.info!.timeline.fyEndMonth ? 1 : 0)}`;
}

/** What Save version offers: this month's reforecast and the budget being built. */
export function suggestions(model: Model): SaveChoice[] {
  const t = model.info!.timeline;
  const last = t.lastActual;
  const fyEnd = t.fyEndMonth;
  const out: SaveChoice[] = [];
  if (last > 0) {
    // months of the financial year done: the calendar month of the last actual against the year end
    const [y, m] = t.start.split('-').map(Number);
    const cm = ((m - 1 + last - 1) % 12) + 1;
    const done = ((cm - fyEnd - 1) % 12 + 12) % 12 + 1;
    const existing = versionsOf(model).find(v => v.type === 'Reforecast' && v.asAt === last);
    out.push({ type: 'Reforecast', label: `Reforecast ${monthLabel(model, last)} (${done}+${12 - done})`, asAt: last,
      replaces: existing && !existing.locked ? existing.id : undefined });
    void y;
  }
  const b = t.budget ?? { first: Math.min(last + 1, model.periods), months: 12 };
  out.push({ type: 'Budget', label: `Budget ${fyLabel(model, b.first)}`, year: fyLabel(model, b.first), months: [b.first, b.first + b.months - 1] });
  return out;
}

/** Where each line Save version keeps sits in the workbook: the statements row's address across the timeline. */
export function lineAddresses(layout: Layout, model: Model, keys: string[]): Record<string, string> {
  const fs = model.instances.find(i => model.lib.module(i.module).framework === 'statements');
  if (!fs) throw new Error('The model has no financial statements to save.');
  const spec = model.lib.module(fs.module).report!;
  const pos = layout.positions();
  const out: Record<string, string> = {};
  const col = (c: number) => { let s = ''; while (c) { const r = (c - 1) % 26; c = Math.floor((c - 1) / 26); s = String.fromCharCode(65 + r) + s; } return s; };
  for (const k of keys) {
    const at = pos.get(`${fs.uid}/${spec.lines[k].row}`);
    if (!at) throw new Error(`The statements have no row for ${k}.`);
    out[k] = `'${at[0]}'!${col(FIRST_PERIOD_COL)}${at[1]}:${col(FIRST_PERIOD_COL + model.periods - 1)}${at[1]}`;
  }
  return out;
}

const round2 = (x: number) => Math.round(x * 100) / 100;

/** Each value times its sheet column (J is 10), added up: the checksum the register keeps. */
export function checksum(values: Record<string, (number | null)[]>): number {
  let s = 0;
  for (const vals of Object.values(values)) vals.forEach((v, t) => { if (v !== null) s += v * (FIRST_PERIOD_COL + t); });
  return round2(s);
}

function nextId(versions: VersionData[]): string {
  const n = Math.max(0, ...versions.map(v => Number(/^V(\d+)$/.exec(v.id)?.[1] ?? 0)));
  return `V${String(n + 1).padStart(2, '0')}`;
}

function planWith(open: OpenModel, change: (versions: VersionData[]) => VersionData[], title: string): InsertPlan {
  const next = open.model.copy();
  const inst = next.instances.find(i => next.lib.module(i.module).framework === 'versions');
  if (!inst) throw new Error('The model has no Saved versions module; insert it first (Modules > Insert).');
  inst.data.versions = change(structuredClone(inst.data.versions ?? []));
  const layout = assemble(next);
  return { next, layout, plan: planChange(open.layout, layout, 'excel'), metadata: metadataXml(next, layout), title };
}

/**
 * Save version: the statements' values (read from the workbook, by line) as a new version, or in
 * place of this month's reforecast when it was saved before and is not locked.
 */
export function planSave(open: OpenModel, choice: SaveChoice, read: Record<string, number[]>, savedBy: string, savedOn: string): InsertPlan {
  const keys = choice.type === 'Budget' ? BUDGET_LINES : REFORECAST_LINES;
  const values: Record<string, (number | null)[]> = {};
  for (const k of keys) {
    const row = read[k];
    if (!row) throw new Error(`No values were read for ${k}.`);
    values[k] = row.map((v, t) => {
      const p = t + 1;
      if (choice.months && (p < choice.months[0] || p > choice.months[1])) return null;
      return round2(Number(v) || 0);
    });
  }
  return planWith(open, versions => {
    const v: VersionData = {
      id: choice.replaces ?? nextId(versions), type: choice.type, label: choice.label, status: 'Saved', locked: false,
      source: 'Save version', savedBy, savedOn, checksum: checksum(values), values,
      ...(choice.year ? { year: choice.year } : {}), ...(choice.asAt ? { asAt: choice.asAt } : {}),
    };
    return choice.replaces ? versions.map(x => (x.id === choice.replaces ? v : x)) : [...versions, v];
  }, choice.label);
}

/** Approve budget: the budget becomes the approved one for its year (and is locked); the one before is superseded. */
export function planApprove(open: OpenModel, id: string): InsertPlan {
  const v = versionsOf(open.model).find(x => x.id === id);
  if (!v || v.type !== 'Budget') throw new Error('Choose a saved budget to approve.');
  return planWith(open, versions => versions.map(x => {
    if (x.id === id) return { ...x, status: 'Approved', locked: true };
    if (x.type === 'Budget' && x.year === v.year && x.status === 'Approved') return { ...x, status: 'Superseded' };
    return x;
  }), `${v.label} approved`);
}

/** Lock or unlock a version: a locked one cannot be replaced. */
export function planLock(open: OpenModel, id: string, locked: boolean): InsertPlan {
  const v = versionsOf(open.model).find(x => x.id === id);
  if (!v) throw new Error('Choose a saved version.');
  return planWith(open, versions => versions.map(x => (x.id === id ? { ...x, locked } : x)), `${v.label} ${locked ? 'locked' : 'unlocked'}`);
}
