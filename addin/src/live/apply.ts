// The live writer: a change plan applied to the open workbook through Office.js. Each plan
// operation maps to a few Office.js calls, batched and synced once where Excel does not need to
// answer a question first. Formats arrive as named styles (the package writer put them in the
// workbook), so nothing here types a colour except the red of a failing check.
//
// Requirement sets: ExcelApi 1.7 (styles by name, hyperlinks, chart series), 1.8 (validation),
// 1.10 (row grouping). The probe checks each before the command is offered.

import type { PlanOp, RowFormat, SheetFormat } from '../../../engine/src/index.ts';
import { colLetter, META_NS } from '../../../engine/src/frame.ts';
import type { CondRule } from '../../../engine/src/xlsx/dress.ts';
import type { Validation } from '../../../engine/src/layout.ts';

/* The parts of the Office.js object model the writer uses, so tests can stand in for Excel. */
export interface XRange {
  formulas: unknown[][];
  values: unknown[][];
  style: string;
  rowHidden: boolean;
  hyperlink: { documentReference?: string; screenTip?: string } | null;
  format: { rowHeight: number; columnWidth: number };
  conditionalFormats: { clearAll(): void; add(type: string): XConditionalFormat };
  dataValidation: { clear(): void; rule: unknown; errorAlert: unknown };
  insert(shift: string): unknown;
  delete(shift: string): void;
  clear(applyTo?: string): void;
  group(by: string): void;
  ungroup(by: string): void;
  load(p: string): void;
}
export interface XConditionalFormat {
  cellValue: { format: { font: { color: string; bold: boolean } }; rule: unknown };
  custom: { format: { font: { color: string; bold: boolean } }; rule: { formula: string } };
}
export interface XChart {
  title: { text: string };
  series: { getItemAt(i: number): { delete(): void }; add(name?: string): XSeries; load(p: string): void; items?: unknown[] };
  setPosition(start: XRange | string): void;
  width: number;
  height: number;
  delete(): void;
  load(p: string): void;
}
export interface XSeries { setValues(r: XRange): void; setXAxisValues(r: XRange): void; chartType: string; name: string }
export interface XSheet {
  name: string;
  position: number;
  showGridlines: boolean;
  getRange(address?: string): XRange;
  getCell(row: number, col: number): XRange;
  delete(): void;
  freezePanes: { freezeAt(r: XRange | string): void; freezeRows(n: number): void; unfreeze(): void };
  charts: { add(type: string, source: XRange, seriesBy?: string): XChart; load(p: string): void; items: XChart[] };
}
export interface XContext {
  workbook: {
    worksheets: { getItem(name: string): XSheet; add(name: string): XSheet };
    names: { add(name: string, ref: XRange | string): unknown; load(p: string): void; items: { name: string; delete(): void }[] };
    customXmlParts: {
      getByNamespace(ns: string): { load(p: string): void; items: { delete(): void }[] };
      add(xml: string): unknown;
    };
  };
  sync(): Promise<void>;
}

export const CHECK_RED_HEX = '#CB2840';
const NUMERIC_TEXT = /^\s*([-+]?[\d.,]+%?|TRUE|FALSE|\d{1,2}[/-]\d{1,2}([/-]\d{2,4})?)\s*$/i;

/** A plain text value Excel would read as a number, a date or TRUE/FALSE keeps its text with a leading apostrophe. */
export function asTyped(v: string): string {
  return NUMERIC_TEXT.test(v) ? `'${v}` : v;
}

/** "'Working capital'!$J$20:$O$20" -> ["Working capital", "$J$20:$O$20"]. */
export function splitRef(ref: string): [string, string] {
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(ref);
  if (!m) throw new Error(`not a sheet reference: ${ref}`);
  return [(m[1] ?? m[2]).replace(/''/g, "'"), m[3]];
}

/** Column width in Excel's characters to Office.js points (Segoe UI 9 body: about 6 pixels a character). */
export const widthToPoints = (chars: number) => Math.round((chars * 6 + 5) * 0.75 * 100) / 100;

function writeValue(cell: XRange, v: unknown): void {
  if (v === null || v === undefined || v === '') cell.clear('Contents');
  else if (typeof v === 'string' && v.startsWith('=')) cell.formulas = [[v]];
  else if (typeof v === 'string') cell.values = [[asTyped(v)]];
  else cell.values = [[v]];
}

export function validationRule(v: Validation): unknown {
  switch (v.kind) {
    case 'list': return { list: { inCellDropDown: true, source: v.items.join(',') } };
    case 'whole': return { wholeNumber: { operator: 'Between', formula1: typeof v.min === 'number' ? v.min : `=${v.min}`, formula2: typeof v.max === 'number' ? v.max : `=${v.max}` } };
    case 'decimal': return { decimal: { operator: 'Between', formula1: -1e15, formula2: 1e15 } };
    case 'date': return { date: { operator: 'Between', formula1: '1990-01-01', formula2: '2200-12-31' } };
    case 'text': return { textLength: { operator: 'LessThanOrEqualTo', formula1: v.max } };
  }
}

function addCond(ws: XSheet, sqref: string, rule: CondRule): void {
  const range = ws.getRange(sqref);
  if (rule.kind === 'notZero') {
    const cf = range.conditionalFormats.add('CellValue');
    cf.cellValue.format.font.color = CHECK_RED_HEX;
    cf.cellValue.format.font.bold = true;
    cf.cellValue.rule = { formula1: '=0', operator: 'NotEqualTo' };
  } else {
    const cf = range.conditionalFormats.add('Custom');
    cf.custom.rule.formula = `=${rule.formula}`;
    cf.custom.format.font.color = CHECK_RED_HEX;
    cf.custom.format.font.bold = true;
  }
}

/** A row's formats: clear what the row had, then styles by name, height, links, check formatting and validation. */
export function applyRowFormat(ws: XSheet, row: number, f: RowFormat, lastCol: number, clearFirst: boolean): void {
  if (clearFirst) {
    const band = ws.getRange(`A${row}:${colLetter(lastCol)}${row}`);
    band.style = 'Normal';
    band.clear('Hyperlinks');
    band.conditionalFormats.clearAll();
    band.dataValidation.clear();
  }
  for (const [c, fmt] of Object.entries(f.cells)) ws.getCell(row - 1, Number(c) - 1).style = fmt.style;
  if (f.height !== undefined) ws.getRange(`${row}:${row}`).format.rowHeight = f.height;
  for (const l of f.links ?? []) ws.getCell(row - 1, l.col - 1).hyperlink = { documentReference: l.to, screenTip: l.tip };
  for (const c of f.conds ?? []) addCond(ws, c.sqref, c.rule);
  for (const v of f.valid ?? []) {
    const dv = ws.getCell(row - 1, v.col - 1).dataValidation;
    dv.rule = validationRule(v.rule);
    dv.errorAlert = { showAlert: true, style: 'Stop', title: 'Not accepted', message: v.rule.message };
  }
}

function applySheetFormat(ws: XSheet, f: SheetFormat): void {
  ws.showGridlines = false;
  for (const c of f.cols) ws.getRange(`${colLetter(c.min)}:${colLetter(c.max)}`).format.columnWidth = widthToPoints(c.width);
  if (f.freeze) {
    if (f.freeze.col > 1) ws.freezePanes.freezeAt(ws.getRange(`A1:${colLetter(f.freeze.col - 1)}${f.freeze.row - 1}`));
    else ws.freezePanes.freezeRows(f.freeze.row - 1);
  }
}

/** The widest column a sheet's rows reach, for clearing a row's formats before rewriting it. */
function lastColumnOf(ops: PlanOp[], sheet: string, fallback: number): number {
  let max = fallback;
  for (const op of ops) {
    if (op.op === 'write' && op.sheet === sheet && op.format) for (const c of Object.keys(op.format.cells)) max = Math.max(max, Number(c));
  }
  return max;
}

export interface LiveReport {
  operations: number;
  syncs: number;
  notes: string[];
}

/**
 * Apply a plan made in the Excel dialect, then replace the model metadata. Rows move with
 * Excel's own reference shifting; only rows the plan names are written.
 */
export async function applyPlan(ctx: XContext, ops: PlanOp[], metadataXml: string): Promise<LiveReport> {
  const wb = ctx.workbook;
  const report: LiveReport = { operations: ops.length, syncs: 0, notes: [] };
  const sync = async () => { report.syncs += 1; await ctx.sync(); };
  const sheet = (name: string) => wb.worksheets.getItem(name);

  wb.names.load('items/name');
  await sync();
  const existing = new Map(wb.names.items.map(n => [n.name, n]));
  const lastCols = new Map<string, number>();
  const lastCol = (s: string) => {
    if (!lastCols.has(s)) lastCols.set(s, lastColumnOf(ops, s, 9));
    return lastCols.get(s)!;
  };

  async function removeChart(sheetName: string, title: string) {
    const ws = sheet(sheetName);
    ws.charts.load('items/title/text');
    await sync();
    for (const ch of ws.charts.items) if (ch.title.text === title) ch.delete();
  }

  /** Stacked columns with any line series over them; series named from their label cells as they read now. */
  async function drawChart(op: Extract<PlanOp, { op: 'add_chart' | 'set_chart' }>) {
    const ws = sheet(op.sheet);
    const at = (ref: string) => { const [s, a] = splitRef(ref); return sheet(s).getRange(a); };
    const labels = op.series.map(s => { const r = at(s.label); r.load('values'); return r; });
    const chart = ws.charts.add(op.series.some(s => s.kind === 'column') ? 'ColumnStacked' : 'Line', at(op.series[0]?.values ?? op.categories), 'Rows');
    chart.series.load('items');
    await sync();
    for (let i = (chart.series.items ?? []).length - 1; i >= 0; i--) chart.series.getItemAt(i).delete();
    op.series.forEach((s, i) => {
      const ser = chart.series.add(String(labels[i].values?.[0]?.[0] ?? ''));
      ser.setValues(at(s.values));
      ser.setXAxisValues(at(op.categories));
      if (s.kind === 'line') ser.chartType = 'Line';
    });
    chart.title.text = op.title;
    chart.setPosition(ws.getCell(op.anchor.row - 1, op.anchor.col - 1));
    chart.width = 453.5;    // 16 cm
    chart.height = 212.6;   // 7.5 cm
    report.notes.push(`${op.sheet}: chart ${op.title} ${op.op === 'add_chart' ? 'added' : 're-pointed'}; its series names are text until the model is next rebuilt.`);
  }

  for (const op of ops) {
    switch (op.op) {
      case 'delete_name': {
        existing.get(op.name)?.delete();
        existing.delete(op.name);
        break;
      }
      case 'delete_sheet':
        sheet(op.sheet).delete();
        break;
      case 'add_sheet': {
        const ws = wb.worksheets.add(op.sheet);
        ws.position = op.index;
        for (const [r, c, v] of op.frame) writeValue(ws.getCell(r - 1, c - 1), v);
        if (op.format) {
          applySheetFormat(ws, op.format.sheet);
          for (const [r, f] of Object.entries(op.format.header)) applyRowFormat(ws, Number(r), f, lastCol(op.sheet), false);
        }
        break;
      }
      case 'insert_rows':
        sheet(op.sheet).getRange(`${op.row}:${op.row + op.count - 1}`).insert('Down');
        break;
      case 'delete_rows':
        sheet(op.sheet).getRange(`${op.row}:${op.row + op.count - 1}`).delete('Up');
        break;
      case 'write': {
        const ws = sheet(op.sheet);
        for (const [c, v] of Object.entries(op.cells)) writeValue(ws.getCell(op.row - 1, Number(c) - 1), v);
        if (op.format) applyRowFormat(ws, op.row, op.format, lastCol(op.sheet), op.why !== 'frame');
        break;
      }
      case 'outline': {
        await sync();
        const ws = sheet(op.sheet);
        const last = Math.max(...op.runs.map(r => r.to), 1);
        for (let i = 0; i < 8; i++) {   // Office.js cannot read a row's level: take every level off, then regroup
          try {
            ws.getRange(`1:${last + 50}`).ungroup('ByRows');
            await sync();
          } catch {
            break;
          }
        }
        const levels = new Map<number, number>();
        for (const run of op.runs) for (let r = run.from; r <= run.to; r++) levels.set(r, run.level);
        const top = Math.max(0, ...op.runs.map(r => r.level));
        for (let k = 1; k <= top; k++) {
          let start = 0;
          for (let r = 1; r <= last + 1; r++) {
            const inside = (levels.get(r) ?? 0) >= k;
            if (inside && !start) start = r;
            if (!inside && start) {
              ws.getRange(`${start}:${r - 1}`).group('ByRows');
              start = 0;
            }
          }
        }
        for (const run of op.runs) if (run.hidden) ws.getRange(`${run.from}:${run.to}`).rowHidden = true;
        break;
      }
      case 'add_name': {
        existing.get(op.name)?.delete();
        const ws = sheet(op.sheet);
        existing.set(op.name, wb.names.add(op.name, ws.getRange(`$${colLetter(op.col)}$${op.row}`)) as { name: string; delete(): void });
        break;
      }
      case 'delete_chart':
        await removeChart(op.sheet, op.title);
        break;
      case 'set_chart':
        await removeChart(op.sheet, op.title);
        await drawChart(op);
        break;
      case 'add_chart':
        await drawChart(op);
        break;
    }
  }

  const parts = wb.customXmlParts.getByNamespace(META_NS);
  parts.load('items');
  await sync();
  for (const p of parts.items) p.delete();
  wb.customXmlParts.add(metadataXml);
  await sync();
  return report;
}
