// Explorer: the open model as a tree (sections, sheets, categories and the modules on them) and,
// for one module, its rows, what it takes from and sends to, its settings and its checks. Worked
// out from the model's metadata and the engine's layout, so it is the model in the workbook, not
// a sample. DOM free; view.ts draws it.

import { categoryHeading, goBlock, type LRow, type Model } from '../../../engine/src/index.ts';
import type { OpenModel } from '../insert/core.ts';

export interface TreeModule {
  uid: string;
  title: string;
  area: string;
}

/** A category's instances under its heading, or a single module on its own. */
export interface TreeEntry {
  kind: 'category' | 'single';
  /** The category module (fm.revenue) or the instance uid. */
  id: string;
  title: string;
  area: string;
  modules: TreeModule[];
}

export interface ExplorerTree {
  title: string;
  summary: string;
  areas: string[];
  sections: { title: string; sheets: { sheet: string; entries: TreeEntry[] }[] }[];
}

export interface RowInfo {
  row: number;
  label: string;
  /** What the row is, in words: Setting, Input, Link in, Calculation, Working, Check, or a heading. */
  what: string;
  head: boolean;
}

export interface Relation {
  uid: string;
  title: string;
  /** What travels: the labels of the rows that carry it. */
  carries: string[];
}

export interface CheckInfo {
  label: string;
  kind: 'error' | 'alert';
  /** Row id, so the add-in can read whether it is raised. */
  row: string;
}

export interface ModuleInfo {
  uid: string;
  module: string;
  title: string;
  area: string;
  kind: string;
  number: number;
  version: number;
  description: string;
  where: string;
  goName: string;
  components: { sheet: string; rows: RowInfo[] }[];
  takes: Relation[];
  sends: Relation[];
  settings: { label: string; value: string }[];
  checks: CheckInfo[];
}

const instanceOf = (rowId: string) => rowId.split('/')[0];

function rowWhat(r: LRow): string {
  if (r.kind === 'heading' || r.kind === 'subheading') return 'Heading';
  if (r.kind === 'section') return 'Group';
  if (r.kind === 'setting') return 'Setting';
  if (r.style === 'check' || r.unit === 'flag') return 'Check';
  if (r.input) return 'Input';
  if (r.id.includes('/in/')) return 'Link in';
  if (r.level === 2 || r.role === 'working') return 'Working';
  return 'Calculation';
}

/** The items of a drop-down list a setting reads, when the model can say what they are. */
function listItems(model: Model, list: string): string[] | null {
  for (const m of model.lib.modules.values()) {
    if (m.list?.name !== list) continue;
    const titles = model.instances.filter(i => i.module === m.id).map(i => model.title(i));
    return [...(m.list.first ?? []), ...titles];
  }
  const own = model.lib.lists.find(l => l.name === list);
  return own ? own.items : null;
}

function settingText(model: Model, uid: string): { label: string; value: string }[] {
  const inst = model.instance(uid);
  const mod = model.lib.module(inst.module);
  return (mod.settings ?? []).map(s => {
    const v = inst.settings[s.key];
    let value: string;
    if (v && typeof v === 'object' && 'group' in (v as object)) value = `Group assumption ${(v as { group: string }).group}`;
    else if (s.choice && typeof v === 'number') value = s.choice[v - 1] ?? String(v);
    else if (s.list && typeof v === 'number') value = listItems(model, s.list)?.[v - 1] ?? String(v);
    else if (s.check) value = v ? 'Yes' : 'No';
    else if (s.unit === '%' && typeof v === 'number') value = `${Math.round(v * 10000) / 100}%`;
    else if (v === null || v === undefined || v === '') value = '';
    else value = String(v) + (s.unit && s.unit !== '$' && s.unit !== '%' ? ` ${s.unit}` : '');
    return { label: s.label, value };
  });
}

/** The model as a tree: sections, their sheets in workbook order, and on each sheet its categories and modules in row order. */
export function explorerTree(open: OpenModel): ExplorerTree {
  const { model, layout } = open;
  const uids = new Set(model.instances.map(i => i.uid));
  const lib = model.lib;
  const onSheet = new Map<string, string[]>();
  for (const [sheet, rows] of layout.sheets) {
    const seen: string[] = [];
    for (const r of rows) {
      const u = instanceOf(r.id);
      if (uids.has(u) && !seen.includes(u)) seen.push(u);
    }
    if (seen.length) onSheet.set(sheet, seen);
  }
  const entries = (sheet: string): TreeEntry[] => {
    const out: TreeEntry[] = [];
    for (const uid of onSheet.get(sheet) ?? []) {
      const inst = model.instance(uid);
      const mod = lib.module(inst.module);
      const tm = { uid, title: model.title(inst), area: mod.area };
      if (lib.kind(inst.module) === 'category') {
        const cat = out.find(e => e.kind === 'category' && e.id === inst.module);
        if (cat) cat.modules.push(tm);
        else out.push({ kind: 'category', id: inst.module, title: categoryHeading(mod), area: mod.area, modules: [tm] });
      } else {
        out.push({ kind: 'single', id: uid, title: tm.title, area: mod.area, modules: [tm] });
      }
    }
    return out;
  };
  const order = layout.sheets.map(([s]) => s).filter(s => onSheet.has(s));
  const placed = new Set<string>();
  const sections = lib.sections.map(sec => {
    const sheets = order.filter(s => sec.areas.includes(s));
    sheets.forEach(s => placed.add(s));
    return { title: sec.title, sheets: sheets.map(sheet => ({ sheet, entries: entries(sheet) })) };
  }).filter(s => s.sheets.length);
  const rest = order.filter(s => !placed.has(s));
  if (rest.length) sections.push({ title: 'Other sheets', sheets: rest.map(sheet => ({ sheet, entries: entries(sheet) })) });
  const sheetCount = sections.reduce((n, s) => n + s.sheets.length, 0);
  return {
    title: model.info?.title ?? 'Model',
    summary: `${model.instances.length} modules on ${sheetCount} sheets in ${sections.length} sections`,
    areas: lib.areas,
    sections,
  };
}

/** The rows of a module's checks, with their kind, for reading whether each is raised. */
export function checkRows(open: OpenModel, uid?: string): CheckInfo[] {
  const { model, layout } = open;
  const out: CheckInfo[] = [];
  for (const [, rows] of layout.sheets) {
    for (const r of rows) {
      if (r.style !== 'check') continue;
      const u = instanceOf(r.id);
      if (uid ? u !== uid : !model.instances.some(i => i.uid === u)) continue;
      const inst = model.instances.find(i => i.uid === u);
      if (!inst) continue;
      const key = r.id.slice(u.length + 1);
      const def = (model.lib.module(inst.module).rows ?? []).find(d => d.key === key);
      out.push({ label: r.label, kind: def?.check === 'alert' ? 'alert' : 'error', row: r.id });
    }
  }
  return out;
}

/** One module: where it sits, its rows, its links in and out (by module, named by what they carry), settings and checks. */
export function moduleInfo(open: OpenModel, uid: string): ModuleInfo {
  const { model, layout } = open;
  const inst = model.instance(uid);
  const mod = model.lib.module(inst.module);
  const pos = layout.positions();
  const labelOf = new Map<string, string>();
  const components: ModuleInfo['components'] = [];
  for (const [sheet, rows] of layout.sheets) {
    const mine = rows.filter(r => instanceOf(r.id) === uid && r.kind !== 'blank');
    for (const r of rows) labelOf.set(r.id, r.label);
    if (!mine.length) continue;
    components.push({
      sheet,
      rows: mine.map(r => ({ row: pos.get(r.id)![1], label: r.label, what: rowWhat(r), head: ['heading', 'subheading', 'section'].includes(r.kind) })),
    });
  }
  const titles = new Map(model.instances.map(i => [i.uid, model.title(i)]));
  const gather = (pick: (from: string, to: string) => string | null) => {
    const by = new Map<string, Set<string>>();
    for (const rec of layout.records) {
      const other = pick(instanceOf(rec.from), instanceOf(rec.to));
      if (!other || other === uid || !titles.has(other)) continue;
      const what = labelOf.get(rec.from) || rec.link;
      if (!by.has(other)) by.set(other, new Set());
      by.get(other)!.add(what);
    }
    return [...by].map(([u, carries]) => ({ uid: u, title: titles.get(u)!, carries: [...carries] }));
  };
  const takes = gather((from, to) => (to === uid ? from : null));
  const sends = gather((from, to) => (from === uid ? to : null));
  const where = components.map(c => `${c.sheet} rows ${Math.min(...c.rows.map(r => r.row))} to ${Math.max(...c.rows.map(r => r.row))}`).join('; ');
  const kind = model.lib.kind(inst.module);
  return {
    uid, module: inst.module, title: model.title(inst), area: mod.area,
    kind: kind === 'category' ? `${mod.title}, one of a category (${categoryHeading(mod)})` : kind === 'mirror' ? `${mod.title}, a mirror` : mod.title,
    number: inst.number, version: model.versionHeld(inst.module) ?? 1, description: mod.description ?? '',
    where, goName: goBlock(uid), components, takes, sends, settings: settingText(model, uid), checks: checkRows(open, uid),
  };
}

/** The instance a row belongs to, or the category module for a category's bar or summary row. */
export function ownerOfRow(open: OpenModel, rowId: string): { uid: string | null; module: string | null } {
  if (rowId.startsWith('cat/')) return { uid: null, module: rowId.split('/')[1] };
  const u = instanceOf(rowId);
  const inst = open.model.instances.find(i => i.uid === u);
  return inst ? { uid: inst.uid, module: inst.module } : { uid: null, module: null };
}
