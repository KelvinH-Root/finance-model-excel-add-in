// The Excel self-test: runs inside Excel (a task pane the harness opens with a test workbook) and
// drives the model service exactly as the ribbon does, one step at a time. After each step it
// recalculates and sends what Excel now holds (every sheet's formulas, the names, the metadata,
// cells showing errors, the check totals) to the harness, which compares it with a fresh build of
// the model the metadata describes (tools/excel-selftest.ts). Fictional test workbooks only.

import { META_NS } from '../../../engine/src/frame.ts';
import type { ModelService } from '../service.ts';

/* eslint-disable @typescript-eslint/no-explicit-any */
declare const Excel: any;
declare const Office: any;

/** One step of a scenario. Lines are found by module: the first instance, the nth, or the last. */
export type Step =
  | { op: 'add'; module: string; nth?: number; last?: boolean }
  | { op: 'above'; module: string; nth?: number; last?: boolean }
  | { op: 'duplicate'; module: string; nth?: number; last?: boolean }
  | { op: 'rename'; module: string; nth?: number; last?: boolean; title: string }
  | { op: 'remove'; module: string; nth?: number; last?: boolean }
  | { op: 'delete'; module: string; nth?: number; last?: boolean }
  | { op: 'insert'; module: string; settings?: Record<string, unknown>; name?: string }
  | { op: 'save-version'; option?: number }
  | { op: 'properties'; title?: string; preparedBy?: string };

export interface Config {
  run: string;
  book: string;
  steps: Step[];
  /** Save the workbook when done (for a look afterwards) and close it. */
  close: boolean;
  /** Apply each operation on its own and time it, so a refusal names the operation. */
  diagnose?: boolean;
}

export interface StepResult {
  label: string;
  ok: boolean;
  message: string;
  ms: number;
  snapshotMs: number;
}

/** What Excel holds after a step. */
export interface Snapshot {
  label: string;
  order: string[];
  sheets: Record<string, { address: string; formulas: unknown[][] }>;
  names: Record<string, string>;
  metadata: string;
  errorCells: string[];
  checks: { errors: number | null; alerts: number | null };
}

const ERROR_TEXT = /^#(REF!|NAME\?|VALUE!|NUM!|DIV\/0!|NULL!|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|GETTING_DATA)/;

async function post(path: string, body: unknown): Promise<void> {
  const r = await fetch(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  if (!r.ok) throw new Error(`${path}: ${r.status}`);
}

/** Recalculate fully, then read every sheet, the names, the metadata and the check totals. */
export async function snapshot(label: string): Promise<Snapshot> {
  return Excel.run(async (ctx: any) => {
    ctx.workbook.application.calculate(Excel.CalculationType.full);
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name');
    const names = ctx.workbook.names;
    names.load('items/name,items/formula');
    const parts = ctx.workbook.customXmlParts.getByNamespace(META_NS);
    parts.load('items');
    await ctx.sync();
    const used = sheets.items.map((ws: any) => {
      const u = ws.getUsedRangeOrNullObject(true);
      u.load('address,formulas,valueTypes,values');
      return u;
    });
    const xml = parts.items.length ? parts.items[0].getXml() : null;
    const chk = ['Chk_Errors', 'Chk_Alerts'].map(n => {
      const nm = ctx.workbook.names.getItemOrNullObject(n);
      return nm;
    });
    await ctx.sync();
    const chkRanges = chk.map((nm: any) => {
      if (nm.isNullObject) return null;
      const r = nm.getRange();
      r.load('values');
      return r;
    });
    await ctx.sync();
    const out: Snapshot = {
      label, order: sheets.items.map((s: any) => s.name), sheets: {}, names: {}, metadata: xml ? xml.value : '', errorCells: [],
      checks: { errors: chkRanges[0] ? Number(chkRanges[0].values[0][0]) : null, alerts: chkRanges[1] ? Number(chkRanges[1].values[0][0]) : null },
    };
    sheets.items.forEach((ws: any, k: number) => {
      const u = used[k];
      if (u.isNullObject) return;
      out.sheets[ws.name] = { address: u.address, formulas: u.formulas };
      const top = /!\$?([A-Z]+)\$?(\d+)/.exec(u.address);
      const col0 = top ? [...top[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) : 1;
      const row0 = top ? Number(top[2]) : 1;
      (u.valueTypes as string[][]).forEach((row, r) => row.forEach((t, c) => {
        const v = u.values[r][c];
        if (t === 'Error' && typeof v === 'string' && ERROR_TEXT.test(v)) out.errorCells.push(`${ws.name}!R${row0 + r}C${col0 + c} ${v}`);
      }));
    });
    for (const n of names.items) out.names[n.name] = n.formula;
    return out;
  });
}

/** The instance a step means: by module, the nth (from 1) or the last. */
function instanceOf(svc: ModelService, s: { module: string; nth?: number; last?: boolean }): string {
  const list = svc.open!.model.instances.filter(i => i.module === s.module);
  if (!list.length) throw new Error(`The model has no ${s.module}.`);
  const inst = s.last ? list[list.length - 1] : list[(s.nth ?? 1) - 1];
  if (!inst) throw new Error(`The model has no ${s.module} number ${s.nth}.`);
  return inst.uid;
}

/** Select a line's first row, as a person would before a command that works on the selection. */
async function selectLine(svc: ModelService, s: { module: string; nth?: number; last?: boolean }): Promise<void> {
  await svc.load();
  const uid = instanceOf(svc, s);
  const pos = svc.open!.layout.positions();
  const row = pos.has(`${uid}/heading`) ? `${uid}/heading` : [...pos.keys()].find(k => k.startsWith(`${uid}/`));
  if (!row || !(await svc.selectRow(row))) throw new Error(`Could not select ${uid}.`);
}

function describe(s: Step): string {
  const which = 'module' in s && s.op !== 'insert' ? `${s.module}${'last' in s && s.last ? ' (last)' : s.nth ? ` #${s.nth}` : ''}` : '';
  switch (s.op) {
    case 'insert': return `Insert ${s.module}`;
    case 'save-version': return `Save version (choice ${s.option ?? 0})`;
    case 'properties': return 'Properties';
    case 'rename': return `Rename ${which} to "${s.title}"`;
    default: return `${s.op[0].toUpperCase()}${s.op.slice(1)} ${which}`;
  }
}

async function runStep(svc: ModelService, s: Step): Promise<string> {
  const sel = async (fn: () => Promise<{ ok: boolean; message: string }>) => {
    const r = await fn();
    if (!r.ok) throw new Error(r.message);
    return r.message;
  };
  switch (s.op) {
    case 'add': await selectLine(svc, s); return sel(() => svc.addCategoryAtSelection());
    case 'above': await selectLine(svc, s); return sel(() => svc.addAboveAtSelection());
    case 'duplicate': await selectLine(svc, s); return sel(() => svc.duplicateAtSelection());
    case 'rename':
      await selectLine(svc, s);
      await svc.renameState();
      await svc.renamePreview(s.title);
      return (await svc.renameApply()).message;
    case 'remove':
      await selectLine(svc, s);
      await svc.removePreview();
      return (await svc.removeApply()).message;
    case 'delete':
      await selectLine(svc, s);
      await svc.deletePreview();
      return (await svc.deleteApply()).message;
    case 'insert':
      await svc.load();
      await svc.insertPreview(s.module, s.settings ?? {}, s.name ?? null);
      return (await svc.insertApply()).message;
    case 'save-version':
      await svc.versionsState();
      await svc.versionsPreview({ kind: 'save', option: s.option ?? 0, savedBy: 'Excel self-test' });
      return (await svc.versionsApply()).message;
    case 'properties': {
      const now = await svc.propertiesState();
      await svc.propertiesPreview({
        title: s.title ?? now.title, preparedBy: s.preparedBy ?? now.preparedBy, notes: now.notes, display: now.display,
      });
      return (await svc.propertiesApply()).message;
    }
  }
}

/** Run a scenario's steps, sending a snapshot after each, then the results; save and close the workbook if asked. */
export async function runSelfTest(svc: ModelService, cfg: Config, say: (line: string) => void = () => {}): Promise<StepResult[]> {
  const base = `/selftest/${encodeURIComponent(cfg.run)}/${encodeURIComponent(cfg.book)}`;
  const results: StepResult[] = [];
  svc.diagnose = !!cfg.diagnose;
  const host = { ...(Office?.context?.diagnostics ?? {}), excelApi: ['1.11', '1.13', '1.17', '1.20'].filter(v => Office.context.requirements.isSetSupported('ExcelApi', v)) };
  const record = async (label: string, work: () => Promise<string>) => {
    say(label);
    const t0 = performance.now();
    let ok = true;
    let message = '';
    try {
      message = await work();
    } catch (e) {
      ok = false;
      message = (e as Error).message || String(e);
    }
    const t1 = performance.now();
    let snapMs = 0;
    try {
      const snap = await snapshot(label);
      snapMs = performance.now() - t1;
      await post(`${base}/snapshot`, snap);
    } catch (e) {
      ok = false;
      message += ` Snapshot failed: ${(e as Error).message}`;
    }
    results.push({ label, ok, message, ms: Math.round(t1 - t0), snapshotMs: Math.round(snapMs) });
    say(`${ok ? 'ok' : 'FAILED'}: ${label}${ok ? '' : ` (${message})`}`);
  };
  await record('Opened as built', async () => 'The package writer\'s workbook, recalculated by Excel.');
  for (const s of cfg.steps) await record(describe(s), () => runStep(svc, s));
  let closed = false;
  if (cfg.close) {
    try {
      await Excel.run(async (ctx: any) => { ctx.workbook.save(Excel.SaveBehavior.save); await ctx.sync(); });
      closed = true;
    } catch (e) {
      say(`Save failed: ${(e as Error).message}`);
    }
  }
  await post(`${base}/done`, { host, results, saved: closed });
  if (closed) {
    await Excel.run(async (ctx: any) => { ctx.workbook.close(Excel.CloseBehavior.skipSave); await ctx.sync(); });
  }
  return results;
}
