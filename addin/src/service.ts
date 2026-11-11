// The model service: everything the add-in's commands do to the open workbook, behind one set of
// plain async methods whose arguments and results are JSON. It runs where Excel is (the shared
// runtime behind the ribbon and the task pane). A dialog cannot reach the workbook, so its views
// call these methods through rpc.ts; the task pane calls them directly. Outside Excel it works on
// the demo models so the views can be seen in a browser.

import { goBlock, metadataXml, type Brand, type Layout, type Logo } from '../../engine/src/index.ts';
import { colLetter, FIRST_PERIOD_COL, META_NS, TOTAL_COL } from '../../engine/src/frame.ts';
import { checkRows, explorerTree, moduleInfo, ownerOfRow, type ExplorerTree, type ModuleInfo } from './explorer/core.ts';
import { changeLog, entryFor, logChange, measure, resultLines, type Effect, type LogEntry } from './guard/core.ts';
import { choices, planInsert, planRemove, readModel, type InsertPlan, type OpenModel } from './insert/core.ts';
import { applyPlan, type XContext } from './live/apply.ts';
import { withClears } from './live/clears.ts';
import {
  changedLists, dropRows, planAbove, planDelete, planDuplicate, planProperties, planRename, PROPERTY_ROWS, relist, settingFromCell,
  type PropertiesEdit,
} from './modules/core.ts';
import {
  BUDGET_LINES, lineAddresses, planApprove, planLock, planSave, REFORECAST_LINES, suggestions, versionsOf,
} from './versions/core.ts';
import { buildFile, fileName, toBase64, type Libraries, type WizardState } from './wizard/core.ts';

/* eslint-disable @typescript-eslint/no-explicit-any */
declare const Excel: any;
declare const Office: any;

export interface SettingInfo {
  key: string;
  label: string;
  unit: string;
  value: unknown;
  /** The items of a drop-down, when the setting is one. */
  options: string[] | null;
  check: boolean;
  /** Greyed unless this holds: [setting key, value]. */
  when: [string, number] | null;
}

export interface InsertChoiceInfo {
  id: string;
  title: string;
  area: string;
  blocked: string | null;
  description: string;
  category: boolean;
  settings: SettingInfo[];
}

export interface Preview {
  title: string;
  lines: string[];
}

export interface VersionsInfo {
  source: string;
  options: { label: string; note: string }[];
  register: { id: string; label: string; type: string; status: string; locked: boolean; canApprove: boolean }[];
}

export type VersionRequest =
  | { kind: 'save'; option: number; savedBy: string }
  | { kind: 'approve'; id: string }
  /** Unlocking an approved budget needs a reason, which the change log keeps. */
  | { kind: 'lock'; id: string; locked: boolean; reason?: string };

/** A change that stopped part way: the workbook no longer matches its model until it is restored. */
export interface Broken {
  title: string;
  /** ISO date and time it stopped. */
  at: string;
  /** Whether a copy of the workbook from just before the change can be opened. */
  copy: boolean;
}

/** Where the add-in marks a change in progress; it stays only if the change stops part way. */
export const PENDING_NS = 'urn:hfg:change-pending:v0';

export type Status = 'error' | 'alert' | 'ok';

/** The methods a dialog may call. */
export interface ServiceApi {
  insertChoices(): Promise<{ source: string; summary: string; choices: InsertChoiceInfo[] }>;
  insertPreview(id: string, values: Record<string, unknown>, name: string | null): Promise<Preview>;
  insertApply(): Promise<{ message: string }>;
  versionsState(): Promise<VersionsInfo>;
  versionsPreview(req: VersionRequest): Promise<Preview>;
  versionsApply(): Promise<{ message: string }>;
  createModel(state: WizardState): Promise<{ opened: boolean; message: string; base64?: string; name?: string }>;
  explorerTree(): Promise<ExplorerTree & { source: string; status: Record<string, Status> }>;
  explorerModule(uid: string): Promise<ModuleInfo & { raised: (boolean | null)[] }>;
  select(name: string): Promise<boolean>;
  removePreview(): Promise<Preview & { uid: string }>;
  removeApply(): Promise<{ message: string }>;
  checksState(): Promise<ChecksInfo>;
  selectRow(rowId: string): Promise<boolean>;
  renameState(): Promise<RenameInfo>;
  renamePreview(name: string): Promise<Preview>;
  renameApply(): Promise<{ message: string }>;
  deletePreview(): Promise<Preview & { uid: string; category: boolean }>;
  deleteApply(): Promise<{ message: string }>;
  propertiesState(): Promise<PropertiesInfo>;
  propertiesPreview(edit: PropertiesEdit): Promise<Preview>;
  propertiesApply(): Promise<{ message: string }>;
  restoreCopy(): Promise<{ message: string }>;
  markRepaired(): Promise<{ message: string }>;
}

export interface RenameInfo {
  uid: string;
  title: string;
  /** The category it is a line of. */
  category: string;
  sheet: string;
}

export interface PropertiesInfo {
  source: string;
  title: string;
  preparedBy: string;
  notes: string[];
  display: { errors: boolean; alerts: boolean };
  /** What the model was built with, shown but changed elsewhere (Time settings, New model). */
  facts: [string, string][];
  modules: { title: string; area: string; count: number; held: number; library: number }[];
  /** The newest changes the add-in made, newest first, with what each did to the key outputs. */
  changes: { when: string; command: string; title: string; lines: string[]; reason?: string }[];
  /** A change that stopped part way, if there is one. */
  broken: Broken | null;
}

export interface ChecksInfo {
  source: string;
  read: boolean;
  errors: number;
  alerts: number;
  passing: number;
  groups: { uid: string; title: string; sheet: string; goName: string; checks: { label: string; kind: 'error' | 'alert'; raised: boolean | null; row: string }[] }[];
}

export interface ServiceAssets {
  libs: Libraries;
  logos: Partial<Record<Brand, Logo>>;
  /** The models to show outside Excel: the small demo for most views, the full model for versions. */
  sample: () => OpenModel;
  sampleFull: () => OpenModel;
}

export const inExcel = () => typeof Excel !== 'undefined' && typeof Office !== 'undefined'
  && !!Office?.context?.requirements?.isSetSupported('ExcelApi', '1.10');

function whenOf(expr: string | undefined): [string, number] | null {
  const m = expr ? /^\$(\w+)=(\d+)$/.exec(expr.trim()) : null;
  return m ? [m[1], Number(m[2])] : null;
}

/** What each kind of change is called in the change log, and whether it may move a number. */
const COMMANDS: Record<string, [string, Effect]> = {
  insert: ['Insert', 'reach'], add: ['Add', 'reach'], above: ['Add above', 'reach'], duplicate: ['Duplicate', 'reach'],
  remove: ['Remove', 'reach'], delete: ['Delete', 'reach'], rename: ['Rename', 'none'], properties: ['Properties', 'none'],
  versions: ['Versions', 'none'],
};

/** Base64 for Excel.createWorkbook, in chunks so a large file does not overflow the call stack. */
function base64Of(bytes: Uint8Array): string {
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}

export class ModelService implements ServiceApi {
  assets: ServiceAssets;
  open: OpenModel | null = null;
  source = '';
  pending: InsertPlan | null = null;
  pendingKind = '';
  /** A reason the pending change carries into the change log (unlocking an approved budget). */
  pendingReason = '';
  /** A change that stopped part way, read from the workbook (or set when it stopped here). */
  broken: Broken | null = null;
  /** The workbook as it was just before the last change, while a change that stopped part way waits to be put right. */
  private backup: string | null = null;
  /** Apply each operation on its own and time it, so a refusal names the operation (the Excel self-test sets it). */
  diagnose = false;

  constructor(assets: ServiceAssets) {
    this.assets = assets;
  }

  /** Refuse to change a workbook that no longer matches its model. */
  private refuseIfBroken(): void {
    if (!this.broken) return;
    const when = new Date(this.broken.at).toLocaleString('en-NZ', { dateStyle: 'medium', timeStyle: 'short' });
    throw new Error(`${this.broken.title} stopped part way (${when}), so this workbook no longer matches its model and the add-in will not change it. `
      + `${this.broken.copy ? 'Open the copy from just before that change (Model > Properties, Open the copy) and save it in place of this one, or undo' : 'Undo'} `
      + 'it in Excel (Ctrl+Z) and then choose I have undone it in Model > Properties.');
  }

  // ------------------------------------------------------------------ the workbook

  /** The model in the open workbook (read fresh), or a demo model outside Excel. */
  async load(full = false): Promise<OpenModel> {
    this.pending = null;
    if (inExcel()) {
      const [xml, stopped]: [string, string] = await Excel.run(async (ctx: any) => {
        const parts = ctx.workbook.customXmlParts.getByNamespace(META_NS);
        const marks = ctx.workbook.customXmlParts.getByNamespace(PENDING_NS);
        parts.load('items');
        marks.load('items');
        await ctx.sync();
        const x = parts.items.length ? parts.items[0].getXml() : null;
        const m = marks.items.length ? marks.items[0].getXml() : null;
        await ctx.sync();
        return [x ? x.value : '', m ? m.value : ''];
      });
      if (!xml) throw new Error('This workbook has no HFG model metadata. Create a model with New model first, or open one made with it.');
      this.open = readModel(xml, this.assets.libs);
      this.source = 'the open workbook';
      if (stopped) {
        const title = /\btitle="([^"]*)"/.exec(stopped)?.[1] ?? 'A change';
        const at = /\bat="([^"]*)"/.exec(stopped)?.[1] ?? new Date().toISOString();
        this.broken = { title: title.replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&amp;/g, '&'), at, copy: !!this.backup };
      } else {
        this.broken = null;
      }
    } else {
      this.open = full ? this.assets.sampleFull() : this.assets.sample();
      this.source = 'the demo model (Excel is not running this view)';
    }
    return this.open;
  }

  /** Each key output (the library's headline rows) as the open workbook works it out now. */
  private async keyOutputs(layout: Layout): Promise<Record<string, number>> {
    const heads = layout.headlines;
    if (!inExcel() || !heads.length || !this.open) return {};
    const pos = layout.positions();
    const periods = layout.periods;
    const first = colLetter(FIRST_PERIOD_COL);
    const last = colLetter(FIRST_PERIOD_COL + periods - 1);
    return Excel.run(async (ctx: any) => {
      ctx.workbook.application.calculate(Excel.CalculationType.recalculate);
      const ranges = heads.map(h => {
        const p = pos.get(h.id);
        if (!p) return null;
        const r = ctx.workbook.worksheets.getItem(p[0]).getRange(`${first}${p[1]}:${last}${p[1]}`);
        r.load('values');
        return r;
      });
      await ctx.sync();
      const out: Record<string, number> = {};
      heads.forEach((h, k) => { const r = ranges[k]; if (r) out[h.label] = measure(h, r.values[0] as unknown[], periods); });
      return out;
    });
  }

  /** The whole workbook as it stands, as base64 (Office's compressed file), so a change that stops part way can be taken back. */
  private async copyOfFile(): Promise<string | null> {
    const doc = typeof Office !== 'undefined' ? Office?.context?.document : null;
    if (!doc?.getFileAsync) return null;
    const call = <T>(fn: (cb: (r: any) => void) => void) => new Promise<T | null>(ok => fn(r => ok(r.status === 'succeeded' ? r.value : null)));
    const file = await call<any>(cb => doc.getFileAsync(Office.FileType.Compressed, { sliceSize: 4 * 1024 * 1024 }, cb));
    if (!file) return null;
    try {
      const parts: Uint8Array[] = [];
      for (let i = 0; i < file.sliceCount; i++) {
        const slice = await call<any>(cb => file.getSliceAsync(i, cb));
        if (!slice) return null;
        parts.push(Uint8Array.from(slice.data as number[]));
      }
      const bytes = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
      let at = 0;
      for (const p of parts) { bytes.set(p, at); at += p.length; }
      return base64Of(bytes);
    } finally {
      await new Promise(ok => file.closeAsync(ok));
    }
  }

  /** Mark a change in progress in the workbook, or take the mark off. */
  private async markPending(title: string | null): Promise<void> {
    await Excel.run(async (ctx: any) => {
      const marks = ctx.workbook.customXmlParts.getByNamespace(PENDING_NS);
      marks.load('items');
      await ctx.sync();
      for (const m of marks.items) m.delete();
      if (title !== null) {
        const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
        ctx.workbook.customXmlParts.add(`<?xml version="1.0" encoding="UTF-8"?><hfgPending xmlns="${PENDING_NS}" title="${esc(title)}" at="${new Date().toISOString()}"/>`);
      }
      await ctx.sync();
    });
  }

  /**
   * Apply the pending plan, guarded: read the key outputs, keep a copy of the workbook and mark the
   * change in progress; apply; read the key outputs again, log the change in the model's metadata
   * and take the mark off. If it stops part way the mark stays, the copy is kept to open, and the
   * add-in refuses further changes until the workbook is put right.
   */
  private async applyPending(kind = this.pendingKind): Promise<string> {
    const p = this.pending;
    if (!p) throw new Error('Nothing to apply: preview it first.');
    this.refuseIfBroken();
    const [command, effect] = COMMANDS[kind] ?? [kind, 'reach'];
    let message: string;
    if (inExcel()) {
      withClears(this.open!.layout, p.layout, p.plan);
      const t0 = Date.now();
      const before = await this.keyOutputs(this.open!.layout);
      const t1 = Date.now();
      const copy = await this.copyOfFile();
      const t2 = Date.now();
      await this.markPending(`${command} ${p.title}`);
      let report: Awaited<ReturnType<typeof applyPlan>>;
      try {
        report = await Excel.run((ctx: XContext) => applyPlan(ctx, p.plan.ops, p.metadata, { stepwise: this.diagnose }));
        if (this.diagnose) {
          const slow = Object.entries(report.timing ?? {}).sort((a, b) => b[1].ms - a[1].ms).slice(0, 5)
            .map(([k, v]) => `${k} ${(v.ms / 1000).toFixed(1)} s (${v.count})`);
          report.notes.push(`Timing: key outputs ${t1 - t0} ms, copy of the file ${t2 - t1} ms; ${slow.join(', ')}.`);
        }
      } catch (e) {
        this.backup = copy;
        this.broken = { title: `${command} ${p.title}`, at: new Date().toISOString(), copy: !!copy };
        this.pending = null;
        throw new Error(`${(e as Error).message || e}. ${this.brokenAdvice()}`);
      }
      const after = await this.keyOutputs(p.layout);
      const entry = entryFor({ command, title: p.title, effect, plan: p.plan, layout: p.layout, before, after, reason: this.pendingReason || undefined });
      logChange(p.next, entry);
      await Excel.run(async (ctx: XContext) => {
        const parts = ctx.workbook.customXmlParts.getByNamespace(META_NS);
        parts.load('items');
        await ctx.sync();
        for (const part of parts.items) part.delete();
        ctx.workbook.customXmlParts.add(metadataXml(p.next, p.layout));
        await ctx.sync();
      });
      await this.markPending(null);
      this.backup = null;
      const lines = resultLines(entry);
      message = `${report.operations} operations in ${report.syncs} round trips.${report.notes.length ? ' ' + report.notes.join(' ') : ''}${lines.length ? ' ' + lines.join(' ') : ''}`;
    } else {
      logChange(p.next, entryFor({ command, title: p.title, effect, plan: p.plan, layout: p.layout, before: {}, after: {}, reason: this.pendingReason || undefined }));
      message = `${p.plan.ops.length} operations; open it in Excel to apply them.`;
    }
    this.open = { model: p.next, layout: p.layout };
    this.pending = null;
    this.pendingReason = '';
    return message;
  }

  private brokenAdvice(): string {
    return `The change stopped part way, so the workbook no longer matches its model. ${this.backup
      ? 'Model > Properties opens a copy from just before the change (save it in place of this one); or undo'
      : 'Undo'} it in Excel (Ctrl+Z), then choose I have undone it in Model > Properties. Until then the add-in will not change this workbook.`;
  }

  /** Open the copy of the workbook from just before the change that stopped part way, in a new window. */
  async restoreCopy(): Promise<{ message: string }> {
    if (!this.backup) throw new Error('There is no copy from before the change in this session. Undo it in Excel (Ctrl+Z) instead.');
    await Excel.createWorkbook(this.backup);
    return { message: 'The copy from just before the change is open in a new window. Save it in place of this workbook, then close this one without saving.' };
  }

  /** The person has undone the change in Excel: take the mark off so the add-in works on the workbook again. */
  async markRepaired(): Promise<{ message: string }> {
    if (inExcel()) await this.markPending(null);
    this.broken = null;
    this.backup = null;
    return { message: 'Marked as put right. Check Model > Checks before going on.' };
  }

  private preview(p: InsertPlan, kind: string): Preview {
    this.refuseIfBroken();
    this.pending = p;
    this.pendingKind = kind;
    this.pendingReason = '';
    return { title: p.title, lines: p.plan.preview };
  }

  /** The row id under the selected cell, read from the active cell and the model's layout. */
  private async selectedRow(): Promise<{ sheet: string; row: number; id: string | null }> {
    const open = await this.load();
    if (!inExcel()) {
      // outside Excel: the first category line of the demo stands in for the selection
      const inst = open.model.instances.find(i => open.model.lib.kind(i.module) === 'category')!;
      const [sheet, row] = open.layout.positions().get(`${inst.uid}/heading`) ?? ['', 0];
      return { sheet, row, id: `${inst.uid}/heading` };
    }
    const at: { sheet: string; row: number } = await Excel.run(async (ctx: any) => {
      const cell = ctx.workbook.getActiveCell();
      cell.load('rowIndex');
      cell.worksheet.load('name');
      await ctx.sync();
      return { sheet: cell.worksheet.name, row: cell.rowIndex + 1 };
    });
    let id: string | null = null;
    for (const [rid, [s, r]] of open.layout.positions()) {
      if (s === at.sheet && r === at.row && !rid.startsWith('@')) { id = rid; break; }
    }
    return { ...at, id };
  }

  async select(name: string): Promise<boolean> {
    if (!inExcel()) return false;
    try {
      return await Excel.run(async (ctx: any) => {
        const nm = ctx.workbook.names.getItemOrNullObject(name);
        await ctx.sync();
        if (nm.isNullObject) return false;
        const r = nm.getRange();
        r.worksheet.activate();
        r.select();
        await ctx.sync();
        return true;
      });
    } catch {
      return false;
    }
  }

  /** Whether each check row is raised (any month flagged), read in one round trip; null outside Excel. */
  private async raised(rowIds: string[]): Promise<(boolean | null)[]> {
    if (!inExcel() || !rowIds.length || !this.open) return rowIds.map(() => null);
    const pos = this.open.layout.positions();
    const last = colLetter(TOTAL_COL + this.open.model.periods);
    const first = colLetter(TOTAL_COL);
    return Excel.run(async (ctx: any) => {
      const ranges = rowIds.map(id => {
        const p = pos.get(id);
        if (!p) return null;
        const r = ctx.workbook.worksheets.getItem(p[0]).getRange(`${first}${p[1]}:${last}${p[1]}`);
        r.load('values');
        return r;
      });
      await ctx.sync();
      return ranges.map(r => (r ? (r.values[0] as unknown[]).some(v => (typeof v === 'number' && v !== 0) || v === true) : null));
    });
  }

  // ------------------------------------------------------------------ insert

  private settingsOf(id: string, from?: Record<string, unknown>): SettingInfo[] {
    const open = this.open!;
    const mod = open.model.lib.module(id);
    const lists = (name: string): string[] | null => {
      for (const m of open.model.lib.modules.values()) {
        if (m.list?.name !== name) continue;
        return [...(m.list.first ?? []), ...open.model.instances.filter(i => i.module === m.id).map(i => open.model.title(i))];
      }
      return open.model.lib.lists.find(l => l.name === name)?.items ?? null;
    };
    return (mod.settings ?? []).map(s => ({
      key: s.key, label: s.label, unit: s.unit ?? '', value: from && s.key in from ? from[s.key] : s.default ?? null,
      options: s.choice ?? (s.list ? lists(s.list) : null), check: !!s.check, when: whenOf(s.when),
    }));
  }

  async insertChoices() {
    const open = await this.load();
    const lib = open.model.lib;
    const list = choices(lib, open.model).map(c => ({
      id: c.id, title: c.title, area: c.area, blocked: c.blocked, description: lib.module(c.id).description ?? '',
      category: lib.kind(c.id) === 'category', settings: this.settingsOf(c.id),
    }));
    const sheets = open.layout.sheets.length;
    return { source: this.source, summary: `${open.model.instances.length} modules on ${sheets} sheets`, choices: list };
  }

  async insertPreview(id: string, values: Record<string, unknown>, name: string | null) {
    if (!this.open) await this.load();
    return this.preview(planInsert(this.open!, id, values, name), 'insert');
  }

  async insertApply() {
    const p = this.pending;
    if (!p || this.pendingKind !== 'insert') throw new Error('Nothing to insert: choose a module first.');
    const done = await this.applyPending();
    if (p.uid) await this.select(goBlock(p.uid));
    return { message: `${p.title} inserted: ${done}` };
  }

  // ------------------------------------------------------------------ what the cells hold now

  /** Cells of the open workbook by row id: one cell at col, or span cells from it. Null outside Excel. */
  private async readRows(want: { id: string; col: number; span?: number }[]): Promise<Map<string, unknown[]> | null> {
    if (!inExcel() || !this.open) return null;
    const pos = this.open.layout.positions();
    return Excel.run(async (ctx: any) => {
      const got = want.map(w => {
        const p = pos.get(w.id);
        if (!p) return null;
        const a = `${colLetter(w.col)}${p[1]}`;
        const r = ctx.workbook.worksheets.getItem(p[0]).getRange(w.span && w.span > 1 ? `${a}:${colLetter(w.col + w.span - 1)}${p[1]}` : a);
        r.load('values');
        return r;
      });
      await ctx.sync();
      const out = new Map<string, unknown[]>();
      got.forEach((r, k) => { if (r) out.set(want[k].id, r.values[0] as unknown[]); });
      return out;
    });
  }

  /** Keep each drop-down on the item it chose when a plan moves, renames or removes items of its list. */
  private async relisted(p: InsertPlan): Promise<InsertPlan> {
    const open = this.open!;
    const lists = changedLists(open, p);
    if (!lists.length) return p;
    const kept = p.layout.positions();
    const rows = dropRows(open.layout, lists).filter(r => kept.has(r.id));
    const read = await this.readRows(rows.map(r => ({ id: r.id, col: r.name ? open.layout.nameCol(r.name) : TOTAL_COL })));
    relist(open, p, new Map(rows.map(r => [r.id, read ? read.get(r.id)?.[0] : r.value])));
    return p;
  }

  /** A line's settings as the workbook holds them now (drop-downs as positions), and the amounts typed in its months. */
  private async lineNow(uid: string): Promise<{ settings: Record<string, unknown>; series: Record<string, (number | null)[]> }> {
    const open = this.open!;
    const inst = open.model.instance(uid);
    const settings = JSON.parse(JSON.stringify(inst.settings)) as Record<string, unknown>;
    const bound = (v: unknown) => typeof v === 'object' && v !== null && 'group' in v;
    const setRows = (open.model.lib.module(inst.module).settings ?? []).filter(s => !bound(inst.settings[s.key]))
      .map(s => ({ id: `${uid}/set.${s.key}`, col: TOTAL_COL, key: s.key }));
    const inputs = open.layout.sheets.flatMap(([, rows]) => rows)
      .filter(r => r.input && r.id.startsWith(`${uid}/`) && !r.id.slice(uid.length + 1).includes('/'));
    const read = await this.readRows([...setRows, ...inputs.map(r => ({ id: r.id, col: FIRST_PERIOD_COL, span: open.model.periods }))]);
    if (read) {
      for (const s of setRows) if (read.has(s.id)) settings[s.key] = settingFromCell(open.model, open.layout, inst.module, s.key, read.get(s.id)![0]);
    }
    const series: Record<string, (number | null)[]> = {};
    for (const r of inputs) {
      const vals = read?.get(r.id) ?? r.values ?? [];
      series[r.id.slice(uid.length + 1)] = vals.map(v => (typeof v === 'number' ? v : null));
    }
    return { settings, series };
  }

  // ------------------------------------------------------------------ on the selection (no window)

  /** The module under the selected cell: a line of a category, a category's bar (module only) or a single module. */
  private async atSelection(): Promise<{ where: string; uid: string | null; module: string | null }> {
    const at = await this.selectedRow();
    const owner = at.id ? ownerOfRow(this.open!, at.id) : { uid: null, module: null };
    return { where: `${at.sheet} row ${at.row}`, ...owner };
  }

  /** The category under the selection and the choices its selected line carries to a new one; a message when there is none. */
  private async categoryAtSelection(needLine: boolean): Promise<{ module: string; uid: string | null; carried: Record<string, unknown> } | string> {
    const s = await this.atSelection();
    const lib = this.open!.model.lib;
    if (!s.module || (needLine && !s.uid)) {
      return `The selected cell (${s.where}) is not ${needLine ? 'on a line of' : 'in'} a category. Select a line of one first: a revenue or cost line, a debtor profile, an asset class or a facility.`;
    }
    if (lib.kind(s.module) !== 'category') return `${lib.module(s.module).title} is not a category: it goes in once, so it takes no other line.`;
    // the selected line's choices (driver, business unit, profile, GST and the like) carry across; amounts start at their defaults
    const now = s.uid ? (await this.lineNow(s.uid)).settings : {};
    const carried = Object.fromEntries((lib.module(s.module).settings ?? [])
      .filter(x => (x.choice || x.list || x.check) && x.key in now).map(x => [x.key, now[x.key]]));
    return { module: s.module, uid: s.uid, carried };
  }

  private async applyHere(p: InsertPlan, kind: string, words: string): Promise<{ ok: boolean; message: string }> {
    this.refuseIfBroken();
    this.pending = p;
    this.pendingKind = kind;
    const done = await this.applyPending();
    if (p.uid) await this.select(goBlock(p.uid));
    return { ok: true, message: `${p.title} ${words}: ${done}` };
  }

  /** Add a line at the end of the category under the selection, with the selected line's choices, and select it. */
  async addCategoryAtSelection(): Promise<{ ok: boolean; message: string }> {
    const c = await this.categoryAtSelection(false);
    if (typeof c === 'string') return { ok: false, message: c };
    let p: InsertPlan;
    try {
      p = planInsert(this.open!, c.module, c.carried, null);
    } catch (e) {
      return { ok: false, message: (e as Error).message };
    }
    return this.applyHere(p, 'add', 'added');
  }

  /** Add a line just above the selected one, with its choices; drop-downs that chose lines below keep them. */
  async addAboveAtSelection(): Promise<{ ok: boolean; message: string }> {
    const c = await this.categoryAtSelection(true);
    if (typeof c === 'string') return { ok: false, message: c };
    let p: InsertPlan;
    try {
      p = await this.relisted(planAbove(this.open!, c.uid!, c.carried, null));
    } catch (e) {
      return { ok: false, message: (e as Error).message };
    }
    return this.applyHere(p, 'above', 'added above the selected line');
  }

  /** Copy the selected line straight after it: its settings and the amounts typed in its months (historical actuals stay with the original). */
  async duplicateAtSelection(): Promise<{ ok: boolean; message: string }> {
    const s = await this.atSelection();
    if (!s.uid) return { ok: false, message: `The selected cell (${s.where}) is not in a module. Select a line of a category to copy it.` };
    let p: InsertPlan;
    try {
      const now = await this.lineNow(s.uid);
      p = await this.relisted(planDuplicate(this.open!, s.uid, now.settings, now.series));
    } catch (e) {
      return { ok: false, message: (e as Error).message };
    }
    return this.applyHere(p, 'duplicate', 'added with the settings and amounts of the line it copies');
  }

  // ------------------------------------------------------------------ remove a category line (asks first)

  async removePreview() {
    const s = await this.atSelection();
    const open = this.open!;
    if (!s.uid) throw new Error(`The selected cell (${s.where}) is not on a line of a category. Select the line to remove first.`);
    const mod = open.model.lib.module(s.module!);
    if (open.model.lib.kind(s.module!) !== 'category') {
      throw new Error(`${mod.title} is not a category line; removing a whole module is Modules > Delete.`);
    }
    const info = moduleInfo(open, s.uid);
    const p = await this.relisted(planRemove(open, s.uid));
    const lines = [...p.plan.preview];
    if (info.sends.length) lines.unshift(`What it sends stops: ${info.sends.map(x => x.title).join(', ')}.`);
    this.pending = p;
    this.pendingKind = 'remove';
    return { title: p.title, lines, uid: s.uid };
  }

  async removeApply() {
    const p = this.pending;
    if (!p || this.pendingKind !== 'remove') throw new Error('Nothing to remove: preview it first.');
    const done = await this.applyPending();
    return { message: `${p.title} removed: ${done}` };
  }

  // ------------------------------------------------------------------ modules: rename, delete

  private target: string | null = null;

  /** The line under the selection, to rename. */
  async renameState(): Promise<RenameInfo> {
    const s = await this.atSelection();
    const open = this.open!;
    if (!s.uid) throw new Error(`The selected cell (${s.where}) is not in a module. Select a line of a category to rename it.`);
    const inst = open.model.instance(s.uid);
    const lib = open.model.lib;
    if (lib.kind(inst.module) !== 'category') {
      throw new Error(`${open.model.title(inst)} is a single module and keeps its library title. Only lines of a category (a revenue line, a debtor profile, an asset class) take their own.`);
    }
    this.target = s.uid;
    return { uid: s.uid, title: open.model.title(inst), category: lib.module(inst.module).title, sheet: s.where.split(' row ')[0] };
  }

  async renamePreview(name: string) {
    if (!this.open || !this.target) throw new Error('Nothing to rename: select a line first.');
    return this.preview(await this.relisted(planRename(this.open, this.target, name)), 'rename');
  }

  async renameApply() {
    const p = this.pending;
    if (!p || this.pendingKind !== 'rename') throw new Error('Nothing to rename: type the new title first.');
    const done = await this.applyPending();
    if (p.uid) await this.select(goBlock(p.uid));
    return { message: `Renamed ${p.title}: ${done}` };
  }

  /** Any module under the selection (a category line or a single module), and what goes with it. */
  async deletePreview() {
    const s = await this.atSelection();
    const open = this.open!;
    if (!s.uid) throw new Error(`The selected cell (${s.where}) is not in a module. Select a cell in the module to delete first.`);
    const info = moduleInfo(open, s.uid);
    const p = await this.relisted(planDelete(open, s.uid));
    const lines = [...p.plan.preview];
    if (info.sends.length) lines.unshift(`What it sends stops: ${info.sends.map(x => x.title).join(', ')}.`);
    this.pending = p;
    this.pendingKind = 'delete';
    return { title: p.title, lines, uid: s.uid, category: open.model.lib.kind(s.module!) === 'category' };
  }

  async deleteApply() {
    const p = this.pending;
    if (!p || this.pendingKind !== 'delete') throw new Error('Nothing to delete: preview it first.');
    const done = await this.applyPending();
    return { message: `${p.title} deleted: ${done}` };
  }

  // ------------------------------------------------------------------ model properties

  private propsNow: Record<string, unknown> = {};

  async propertiesState(): Promise<PropertiesInfo> {
    const open = await this.load();
    const info = open.model.info;
    if (!info) throw new Error('This model was not made with New model, so it has no properties to show.');
    const read = await this.readRows(Object.values(PROPERTY_ROWS).map(id => ({ id, col: TOTAL_COL })));
    this.propsNow = read ? Object.fromEntries([...read].map(([k, v]) => [k, v[0]])) : {};
    const now = (rid: string, held: unknown) => (rid in this.propsNow ? this.propsNow[rid] : held);
    const t = info.timeline;
    const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    const [y, m] = t.start.split('-').map(Number);
    const monthOf = (period: number) => { const k = m - 1 + period - 1; return `${months[k % 12]} ${y + Math.floor(k / 12)}`; };
    const lib = open.model.lib;
    const ids = [...new Set(open.model.instances.map(i => i.module))];
    return {
      source: this.source,
      title: String(now(PROPERTY_ROWS.title, info.title) ?? ''),
      preparedBy: String(now(PROPERTY_ROWS.preparedBy, info.preparedBy) ?? ''),
      notes: [...info.notes],
      display: { errors: !!now(PROPERTY_ROWS.errors, info.display.errors), alerts: !!now(PROPERTY_ROWS.alerts, info.display.alerts) },
      facts: [
        ['Entity', info.entity.name],
        ['Timeline', `${monthOf(1)} to ${monthOf(open.model.periods)}, ${open.model.periods} months`],
        ['Financial year ends', months[t.fyEndMonth - 1]],
        ['Actuals to', t.lastActual ? monthOf(t.lastActual) : 'none yet'],
        ['Denomination', t.denomination],
        ['Library', lib.id],
      ],
      modules: ids.map(id => ({
        title: lib.module(id).title, area: lib.module(id).area, count: open.model.instances.filter(i => i.module === id).length,
        held: open.model.versionHeld(id) ?? 1, library: lib.versionOf(id),
      })),
      changes: [...changeLog(open.model)].reverse().slice(0, 20).map((e: LogEntry) => ({
        when: e.when, command: e.command, title: e.title, lines: resultLines(e), ...(e.reason ? { reason: e.reason } : {}),
      })),
      broken: this.broken,
    };
  }

  async propertiesPreview(edit: PropertiesEdit) {
    if (!this.open) await this.propertiesState();
    return this.preview(planProperties(this.open!, edit, this.propsNow), 'properties');
  }

  async propertiesApply() {
    const p = this.pending;
    if (!p || this.pendingKind !== 'properties') throw new Error('Nothing to save: change a property first.');
    const done = await this.applyPending();
    return { message: `Properties saved: ${done}` };
  }

  // ------------------------------------------------------------------ versions

  async versionsState(): Promise<VersionsInfo> {
    const open = await this.load(true);
    if (!open.model.instances.some(i => open.model.lib.module(i.module).framework === 'versions')) {
      throw new Error('This model has no Saved versions module. Insert it first (Modules > Insert, Versions).');
    }
    const opts = suggestions(open.model);
    return {
      source: this.source,
      options: opts.map(o => ({ label: o.label, note: o.replaces ? 'Replaces the one saved before' : o.type === 'Budget' ? `${o.year}, the budget being built` : 'This month\'s reforecast' })),
      register: [...versionsOf(open.model)].reverse().map(v => ({
        id: v.id, label: v.label, type: v.type, status: v.status, locked: !!v.locked, canApprove: v.type === 'Budget' && v.status !== 'Approved',
      })),
    };
  }

  async versionsPreview(req: VersionRequest) {
    const open = this.open ?? await this.load(true);
    if (req.kind === 'approve') return this.preview(planApprove(open, req.id), 'versions');
    if (req.kind === 'lock') {
      const reason = (req.reason ?? '').trim();
      const shown = this.preview(planLock(open, req.id, req.locked, reason), 'versions');
      if (!reason) return shown;
      this.pendingReason = reason;
      return { ...shown, lines: [`Reason, kept in the change log: ${reason}`, ...shown.lines] };
    }
    const choice = suggestions(open.model)[req.option];
    if (!choice) throw new Error('Choose what to save.');
    const keys = choice.type === 'Budget' ? BUDGET_LINES : REFORECAST_LINES;
    let values: Record<string, number[]>;
    if (inExcel()) {
      const addr = lineAddresses(open.layout, open.model, keys);
      const read: { values: Record<string, number[]>; errors: number } = await Excel.run(async (ctx: any) => {
        const ranges = Object.entries(addr).map(([k, a]) => {
          const sheet = a.slice(1, a.indexOf("'!"));
          const r = ctx.workbook.worksheets.getItem(sheet).getRange(a.slice(a.indexOf("'!") + 2));
          r.load('values');
          return [k, r] as const;
        });
        const chk = ctx.workbook.names.getItem('Chk_Errors').getRange();
        chk.load('values');
        await ctx.sync();
        return {
          values: Object.fromEntries(ranges.map(([k, r]) => [k, (r.values[0] as unknown[]).map(v => Number(v) || 0)])),
          errors: Number(chk.values[0][0]) || 0,
        };
      });
      if (read.errors > 0) throw new Error(`The model has ${read.errors} error check${read.errors > 1 ? 's' : ''} failing; clear them before saving a version.`);
      values = read.values;
    } else {
      // outside Excel: the latest saved reforecast stands in for the statements' values
      const last = [...versionsOf(open.model)].reverse().find(v => v.type === 'Reforecast');
      values = Object.fromEntries(keys.map(k => [k, (last?.values[k] ?? []).map(v => v ?? 0).concat(Array(open.model.periods).fill(0)).slice(0, open.model.periods)]));
    }
    const today = new Date().toISOString().slice(0, 10);
    return this.preview(planSave(open, choice, values, req.savedBy || 'Group Finance', today), 'versions');
  }

  async versionsApply() {
    const p = this.pending;
    if (!p || this.pendingKind !== 'versions') throw new Error('Nothing to apply: preview it first.');
    const done = await this.applyPending();
    return { message: `${p.title}: ${done}` };
  }

  // ------------------------------------------------------------------ new model

  async createModel(state: WizardState) {
    const bytes = buildFile(state, this.assets.libs, this.assets.logos);
    const name = fileName(state);
    const canOpen = typeof Excel !== 'undefined' && typeof Office !== 'undefined'
      && !!Office?.context?.requirements?.isSetSupported('ExcelApi', '1.8');
    if (canOpen) {
      await Excel.createWorkbook(toBase64(bytes));
      return { opened: true, message: `${name} is open in a new window. Save it where the team keeps models.` };
    }
    return { opened: false, message: `${name} is ready (${Math.round(bytes.length / 1024)} KB). Excel is not running this view, so download it instead.`, base64: toBase64(bytes), name };
  }

  // ------------------------------------------------------------------ checks

  /** Every check in the model by module, and whether each is raised now (read from the workbook). */
  async checksState(): Promise<ChecksInfo> {
    const open = await this.load();
    const rows = checkRows(open);
    const flags = await this.raised(rows.map(r => r.row));
    const pos = open.layout.positions();
    const groups: ChecksInfo['groups'] = [];
    rows.forEach((c, k) => {
      const uid = c.row.split('/')[0];
      let g = groups.find(x => x.uid === uid);
      if (!g) {
        const inst = open.model.instance(uid);
        g = { uid, title: open.model.title(inst), sheet: pos.get(c.row)?.[0] ?? '', goName: goBlock(uid), checks: [] };
        groups.push(g);
      }
      g.checks.push({ label: c.label, kind: c.kind, raised: flags[k], row: c.row });
    });
    const all = groups.flatMap(g => g.checks);
    return {
      source: this.source, read: all.some(c => c.raised !== null),
      errors: all.filter(c => c.raised && c.kind === 'error').length,
      alerts: all.filter(c => c.raised && c.kind === 'alert').length,
      passing: all.filter(c => c.raised === false).length,
      groups,
    };
  }

  /** Select a row of the model (a check's row, say) on its sheet, from its first month. */
  async selectRow(rowId: string): Promise<boolean> {
    if (!inExcel() || !this.open) return false;
    const p = this.open.layout.positions().get(rowId);
    if (!p) return false;
    await Excel.run(async (ctx: any) => {
      const ws = ctx.workbook.worksheets.getItem(p[0]);
      ws.activate();
      ws.getRange(`B${p[1]}`).select();
      await ctx.sync();
    });
    return true;
  }

  // ------------------------------------------------------------------ explorer

  async explorerTree() {
    const open = await this.load();
    const tree = explorerTree(open);
    const checks = checkRows(open);
    const flags = await this.raised(checks.map(c => c.row));
    const status: Record<string, Status> = {};
    checks.forEach((c, k) => {
      const uid = c.row.split('/')[0];
      if (flags[k] === null) return;
      const s: Status = flags[k] ? (c.kind === 'error' ? 'error' : 'alert') : 'ok';
      const was = status[uid];
      if (!was || (s === 'error') || (s === 'alert' && was === 'ok')) status[uid] = s;
    });
    return { ...tree, source: this.source, status };
  }

  async explorerModule(uid: string) {
    const open = this.open ?? await this.load();
    const info = moduleInfo(open, uid);
    return { ...info, raised: await this.raised(info.checks.map(c => c.row)) };
  }
}
