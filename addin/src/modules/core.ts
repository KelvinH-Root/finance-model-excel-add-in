// The everyday module commands as change plans: add a category line above another, duplicate a
// line (with its inputs), rename a line, delete a module, and change the model's own properties.
// Each is planned on a copy of the model and applied like an insert, so the plan only adds,
// removes or rewrites rows. A line that moves, changes its title or goes also moves, renames or
// drops an item of the drop-down list its category keeps, so relist() keeps every drop-down on the
// item it chose. DOM free.

import {
  assemble, metadataXml, planChange, TOTAL_COL, type InstanceData, type Layout, type LRow, type Model,
} from '../../../engine/src/index.ts';
import type { InsertPlan, OpenModel } from '../insert/core.ts';

function finish(open: OpenModel, next: Model, title: string, uid?: string): InsertPlan {
  const layout = assemble(next);
  return { next, layout, plan: planChange(open.layout, layout, 'excel'), metadata: metadataXml(next, layout), title, uid };
}

const categoryOnly = (open: OpenModel, uid: string, what: string) => {
  const inst = open.model.instance(uid);
  if (open.model.lib.kind(inst.module) !== 'category') {
    throw new Error(`${open.model.title(inst)} goes in once, so ${what}.`);
  }
  return inst;
};

/** A new line of the same category just above another. */
export function planAbove(open: OpenModel, beforeUid: string, settings: Record<string, unknown>, name: string | null = null): InsertPlan {
  const before = categoryOnly(open, beforeUid, 'it takes no line above it');
  const next = open.model.copy();
  const inst = next.insert(before.module, settings, name);
  next.instances.pop();
  next.instances.splice(next.instances.findIndex(i => i.uid === beforeUid), 0, inst);
  return finish(open, next, next.title(inst), inst.uid);
}

/** A copy of a line, straight after it: its settings, the amounts typed in its months, and a title of its own. */
export function planDuplicate(open: OpenModel, uid: string, settings: Record<string, unknown>, series: Record<string, (number | null)[]>): InsertPlan {
  const src = categoryOnly(open, uid, 'it cannot be duplicated');
  const next = open.model.copy();
  const data: InstanceData = { series };
  const titles = new Set(next.instances.map(i => next.title(i)));
  const base = open.model.title(src);
  let name = `${base} (copy)`;
  for (let k = 2; titles.has(name); k++) name = `${base} (copy ${k})`;
  const inst = next.insert(src.module, settings, name, data);
  next.instances.pop();
  next.instances.splice(next.instances.findIndex(i => i.uid === uid) + 1, 0, inst);
  return finish(open, next, name, inst.uid);
}

/** A line's new title: its labels, contents entry, drop-down item and every label that reads it follow. */
export function planRename(open: OpenModel, uid: string, name: string): InsertPlan {
  const inst = open.model.instance(uid);
  if (open.model.lib.kind(inst.module) !== 'category') {
    throw new Error(`${open.model.title(inst)} is a single module and keeps its library title; only lines of a category take their own.`);
  }
  const clean = name.trim();
  if (!clean) throw new Error('Type the new title.');
  if (clean.length > 80) throw new Error('Keep the title to 80 characters.');
  if (clean === open.model.title(inst)) throw new Error('That is its title now.');
  if (open.model.instances.some(i => i.uid !== uid && open.model.title(i) === clean)) throw new Error(`Another line is already called ${clean}.`);
  const next = open.model.copy();
  next.instance(uid).name = clean;
  return finish(open, next, clean, uid);
}

/**
 * Taking any module out (a category line or a single module): its rows, names and checks go and
 * the totals it fed close up. A module others cannot do without (the statements, which every
 * summary reads) is refused with the engine's reason.
 */
export function planDelete(open: OpenModel, uid: string): InsertPlan {
  const title = open.model.title(open.model.instance(uid));
  const next = open.model.copy();
  next.remove(uid);
  try {
    return finish(open, next, title, uid);
  } catch (e) {
    throw new Error(`${title} cannot go: ${(e as Error).message}.`);
  }
}

export interface PropertiesEdit {
  title: string;
  preparedBy: string;
  notes: string[];
  /** Which counts the model name line on every sheet shows. */
  display: { errors: boolean; alerts: boolean };
}

/** The Settings sheet rows that hold the properties (inputs there too), by the property they hold. */
export const PROPERTY_ROWS = {
  title: 'settings/model/title', preparedBy: 'settings/model/prepared', errors: 'settings/display/errors', alerts: 'settings/display/alerts',
} as const;

/**
 * The model's title, Prepared by line, notes and status line. The title, Prepared by and the
 * status line are inputs on the Settings sheet (every header reads them), so they are written
 * there when they differ from what the workbook holds now (current, by row id); the notes are
 * rows on the contents.
 */
export function planProperties(open: OpenModel, edit: PropertiesEdit, current: Record<string, unknown> = {}): InsertPlan {
  const info = open.model.info;
  if (!info) throw new Error('This model was not made with New model, so it has no properties to change.');
  const title = edit.title.trim();
  if (!title) throw new Error('The model needs a title.');
  if (title.length > 80) throw new Error('Keep the title to 80 characters.');
  const preparedBy = edit.preparedBy.trim();
  if (preparedBy.length > 120) throw new Error('Keep the Prepared by line under 120 characters.');
  const next = open.model.copy();
  const display = { errors: !!edit.display.errors, alerts: !!edit.display.alerts };
  next.info = { ...info, title, preparedBy, notes: edit.notes.map(n => n.trim()).filter(Boolean), display };
  const p = finish(open, next, title);
  const now = (rid: string, held: unknown) => (rid in current ? current[rid] : held);
  const lines: string[] = [];
  const set = (rid: string, label: string, was: unknown, to: unknown, words: (v: unknown) => string) => {
    if (now(rid, was) === to) return;
    const at = p.layout.positions().get(rid);
    if (!at) return;
    p.plan.ops.push({ op: 'write', sheet: at[0], row: at[1], why: 'rewire', kind: 'setting', style: '', unit: '', cells: { [TOTAL_COL]: to } });
    lines.push(`${label}: ${words(to)} (was ${words(now(rid, was))}).`);
  };
  const text = (v: unknown) => (v === '' || v === null || v === undefined ? 'blank' : String(v));
  const shown = (v: unknown) => (v ? 'shown' : 'not shown');
  set(PROPERTY_ROWS.title, 'Model title', info.title, title, text);
  set(PROPERTY_ROWS.preparedBy, 'Prepared by line', info.preparedBy, preparedBy, text);
  set(PROPERTY_ROWS.errors, 'Error count in the model name', info.display.errors, display.errors, shown);
  set(PROPERTY_ROWS.alerts, 'Alert count in the model name', info.display.alerts, display.alerts, shown);
  p.plan.preview.unshift(...lines);
  if (!p.plan.ops.length) throw new Error('Nothing has changed.');
  return p;
}

// ------------------------------------------------------------------ drop-downs follow their items

export interface ListItem {
  /** The line's uid, or text: and the item for a fixed first item. */
  id: string;
  label: string;
}

const rowsOf = (layout: Layout): LRow[] => layout.sheets.flatMap(([, rows]) => rows);

/** The items of a drop-down list, in order, as a layout has them. */
export function listItems(layout: Layout, model: Model, name: string): ListItem[] {
  const rg = layout.ranges.get(name);
  if (!rg) return [];
  const rows = rowsOf(layout);
  const a = rows.findIndex(r => r.id === rg.from);
  const b = rows.findIndex(r => r.id === rg.to);
  if (a < 0 || b < a) return [];
  return rows.slice(a, b + 1).map(r => {
    const v = String(r.cells[rg.col] ?? '');
    const m = /^=«B\|(.+)\/heading»$/.exec(v);
    if (!m) return { id: `text:${v}`, label: v };
    const inst = model.instances.find(i => i.uid === m[1]);
    return { id: m[1], label: inst ? model.title(inst) : m[1] };
  });
}

/** The list a drop-down row chooses from ('' for any other row). */
const listOf = (r: LRow): string => (r.control?.kind === 'drop' ? r.control.list : '');

/** Every drop-down row (settings and selections) choosing from one of these lists. */
export function dropRows(layout: Layout, lists: Iterable<string>): LRow[] {
  const want = new Set(lists);
  return rowsOf(layout).filter(r => listOf(r) && want.has(listOf(r)));
}

/** The lists whose items move, change title or go between the open model and a plan (an item added at the end moves nothing). */
export function changedLists(open: OpenModel, p: InsertPlan): string[] {
  const names = new Set(dropRows(open.layout, open.layout.ranges.keys()).map(listOf));
  return [...names].filter(n => {
    const after = listItems(p.layout, p.next, n);
    return listItems(open.layout, open.model, n).some((it, k) => after[k]?.id !== it.id || after[k]?.label !== it.label);
  });
}

/**
 * Keep every drop-down on the item it chose when a plan moves, renames or removes items of its
 * list. values holds what each drop-down holds in the workbook now (its position, or its item's
 * text where the add-in wrote it live). A drop-down whose item goes takes the list's first item,
 * and the preview says so. The plan gains a write for each drop-down that changes, and the model
 * (which keeps positions) follows.
 */
export function relist(open: OpenModel, p: InsertPlan, values: Map<string, unknown>): number {
  const lists = changedLists(open, p);
  if (!lists.length) return 0;
  const pos = p.layout.positions();
  const ownerTitle = (rid: string) => {
    const uid = rid.split('/')[0];
    const inst = p.next.instances.find(i => i.uid === uid);
    return inst ? p.next.title(inst) : uid;
  };
  const notes: string[] = [];
  let moved = 0;
  for (const list of lists) {
    const before = listItems(open.layout, open.model, list);
    const after = listItems(p.layout, p.next, list);
    if (!after.length) continue;
    for (const r of dropRows(p.layout, [list])) {
      if (!values.has(r.id)) continue;   // a new row: the plan writes it
      const v = values.get(r.id);
      const item = typeof v === 'number' ? before[v - 1] : typeof v === 'string' ? before.find(i => i.label === v) : undefined;
      if (!item) continue;   // not a choice that can be read: left as it is
      let k = after.findIndex(i => i.id === item.id);
      if (k < 0) {
        k = 0;
        notes.push(`${ownerTitle(r.id)}: ${r.label} chose ${item.label}, which goes; it will choose ${after[0].label}.`);
      }
      const to = typeof v === 'number' ? k + 1 : after[k].label;
      const m = /^(.+)\/set\.(\w+)$/.exec(r.id);
      if (m && p.next.instances.some(i => i.uid === m[1])) p.next.instance(m[1]).settings[m[2]] = k + 1;
      r.value = k + 1;
      if (to === v) continue;
      const at = pos.get(r.id);
      if (!at) continue;
      const col = r.name ? p.layout.nameCol(r.name) : TOTAL_COL;
      p.plan.ops.push({ op: 'write', sheet: at[0], row: at[1], why: 'rewire', kind: r.kind, style: r.style, unit: r.unit, cells: { [col]: to } });
      moved += 1;
    }
  }
  if (moved) {
    p.metadata = metadataXml(p.next, p.layout);
    p.plan.preview.push(...notes, `${moved} drop-down${moved > 1 ? 's keep their' : ' keeps its'} choice as the list changes.`);
  }
  return moved;
}

/** Choice and list settings as the model keeps them (positions) from what the cells hold. */
export function settingFromCell(model: Model, layout: Layout, moduleId: string, key: string, v: unknown): unknown {
  const s = (model.lib.module(moduleId).settings ?? []).find(x => x.key === key);
  if (!s) return v;
  if (v === '' || v === null || v === undefined) return s.default ?? null;
  if (typeof v === 'string' && s.choice) {
    const k = s.choice.indexOf(v);
    return k >= 0 ? k + 1 : s.default ?? 1;
  }
  if (typeof v === 'string' && s.list) {
    const k = listItems(layout, model, s.list).findIndex(i => i.label === v);
    if (k >= 0) return k + 1;
    const shared = model.lib.lists.find(l => l.name === s.list);
    const j = shared ? shared.items.indexOf(v) : -1;
    return j >= 0 ? j + 1 : s.default ?? 1;
  }
  return v;
}
