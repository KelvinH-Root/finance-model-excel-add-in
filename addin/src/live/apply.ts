// The live writer: a change plan applied to the open workbook through Office.js. Each plan
// operation maps to a few Office.js calls, batched and synced once where Excel does not need to
// answer a question first. Formats arrive as named styles (the package writer put them in the
// workbook), so nothing here types a colour except the red of a failing check.
//
// Requirement sets: ExcelApi 1.7 (styles by name, hyperlinks, chart series), 1.8 (validation),
// 1.10 (row grouping). The probe checks each before the command is offered.

import type { PlanOp, RowFormat, SheetFormat } from '../../../engine/src/index.ts';
import { colLetter, META_NS } from '../../../engine/src/frame.ts';
import type { Brand } from '../../../engine/src/model.ts';
import { THEMES } from '../../../engine/src/theme.ts';
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
  merge(across?: boolean): void;
  unmerge(): void;
  delete(shift: string): void;
  clear(applyTo?: string): void;
  group(by: string): void;
  ungroup(by: string): void;
  load(p: string): void;
  /** An in-cell control (ExcelApi 1.18): the check box a classic one becomes when written live. */
  control?: unknown;
}
export interface XConditionalFormat {
  cellValue: { format: { font: { color: string; bold: boolean } }; rule: unknown };
  custom: {
    format: { font: { color: string; bold: boolean; italic?: boolean }; fill: { color: string };
      borders: { getItem(edge: string): { style: string; color: string } } };
    rule: { formula: string };
  };
}
export interface XChart {
  title: { text: string; setFormula?(f: string): void };
  name?: string;
  top?: number;
  left?: number;
  legend?: { position: string; visible: boolean };
  axes?: { categoryAxis: { reversePlotOrder: boolean }; valueAxis: { visible: boolean; numberFormat: string } };
  series: { getItemAt(i: number): { delete(): void }; add(name?: string): XSeries; load(p: string): void; items?: unknown[] };
  setPosition(start: XRange | string): void;
  width: number;
  height: number;
  delete(): void;
  load(p: string): void;
}
export interface XSeries {
  setValues(r: XRange): void;
  setXAxisValues(r: XRange): void;
  chartType: string;
  name: string;
  /** Report charts (ExcelApi 1.7 and 1.8): look, gaps and labels. */
  format?: { fill: { setSolidColor(c: string): void; clear(): void }; line: { color: string; weight: number; lineStyle: string } };
  markerStyle?: string;
  gapWidth?: number;
  overlap?: number;
  hasDataLabels?: boolean;
  dataLabels?: { numberFormat: string; position: string; showPercentage: boolean; showValue: boolean };
  points?: { getItemAt(i: number): { format: { fill: { setSolidColor(c: string): void } } } };
}
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
    names: {
      add(name: string, ref: XRange | string): unknown; load(p: string): void; items: { name: string; delete(): void }[];
      getItem(name: string): { getRange(): XRange };
    };
    customXmlParts: {
      getByNamespace(ns: string): { load(p: string): void; items: { delete(): void }[] };
      add(xml: string): unknown;
    };
  };
  sync(): Promise<void>;
}

export const CHECK_RED_HEX = '#CB2840';
/** Background 1 darker 25%: the grey of an input that is not in use. */
export const INACTIVE_GREY_HEX = '#BFBFBF';
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

/** A sheet name as a formula writes it. */
const sheetRef = (name: string) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);

/** Column width in Excel's characters to Office.js points (Segoe UI 9 body: about 6 pixels a character). */
export const widthToPoints = (chars: number) => Math.round((chars * 6 + 5) * 0.75 * 100) / 100;

function writeValue(cell: XRange, v: unknown): void {
  if (v === null || v === undefined || v === '') cell.clear('Contents');
  else if (typeof v === 'string' && v.startsWith('=')) cell.formulas = [[v]];
  else if (typeof v === 'string') cell.values = [[asTyped(v)]];
  else cell.values = [[v]];
}

export function validationRule(v: Validation, address = 'A1'): unknown {
  switch (v.kind) {
    case 'logical': return { custom: { formula: `=ISLOGICAL(${address})` } };
    case 'list': return { list: { inCellDropDown: true, source: v.items.join(',') } };
    case 'whole': return { wholeNumber: { operator: 'Between', formula1: typeof v.min === 'number' ? v.min : `=${v.min}`, formula2: typeof v.max === 'number' ? v.max : `=${v.max}` } };
    case 'decimal': return { decimal: { operator: 'Between', formula1: -1e15, formula2: 1e15 } };
    case 'date': return { date: { operator: 'Between', formula1: '1990-01-01', formula2: '2200-12-31' } };
    case 'text': return { textLength: { operator: 'LessThanOrEqualTo', formula1: v.max } };
  }
}

/** Colours the live writer types where the package writer uses theme slots (Office.js cannot set theme colours). */
export interface LivePalette {
  /** The scenario column in use: accent 1 lighter 60%. */
  selected: string;
  /** Dark 2: the active marker and a count that is on. */
  dark: string;
}

/** A theme colour lightened as Excel tints it (0 to 1). */
export function tintHex(hex: string, tint: number): string {
  const ch = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16)).map(c => Math.round(c + (255 - c) * tint));
  return `#${ch.map(c => c.toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

export function livePalette(brand: Brand): LivePalette {
  const t = THEMES[brand];
  return { selected: tintHex(t.accents[0], 0.6), dark: `#${t.dk2}` };
}

function addCond(ws: XSheet, sqref: string, rule: CondRule, palette: LivePalette = livePalette('HF')): void {
  const range = ws.getRange(sqref);
  if (rule.kind === 'on') {
    const cf = range.conditionalFormats.add('CellValue');
    cf.cellValue.format.font.color = palette.dark;
    cf.cellValue.format.font.bold = true;
    cf.cellValue.rule = { formula1: '=0', operator: 'NotEqualTo' };
    return;
  }
  if (rule.kind === 'selected' || rule.kind === 'upright' || rule.kind === 'marker') {
    const cf = range.conditionalFormats.add('Custom');
    cf.custom.rule.formula = `=${rule.formula}`;
    if (rule.kind === 'selected') cf.custom.format.fill.color = palette.selected;
    if (rule.kind === 'marker') cf.custom.format.font.color = palette.dark;
    if (rule.kind === 'upright') {
      cf.custom.format.font.bold = true;
      cf.custom.format.font.italic = false;
    }
    return;
  }
  if (rule.kind === 'na') {
    // a report table's cell outside the timeline (#N/A, a gap in the chart): light grey text
    const cf = range.conditionalFormats.add('Custom');
    cf.custom.rule.formula = `=${rule.formula}`;
    cf.custom.format.font.color = INACTIVE_GREY_HEX;
    return;
  }
  if (rule.kind === 'notZero') {
    const cf = range.conditionalFormats.add('CellValue');
    cf.cellValue.format.font.color = CHECK_RED_HEX;
    cf.cellValue.format.font.bold = true;
    cf.cellValue.rule = { formula1: '=0', operator: 'NotEqualTo' };
  } else if (rule.kind === 'inactive') {
    // An input not used in this month or for this method: background-grey text, white fill and rules.
    const cf = range.conditionalFormats.add('Custom');
    cf.custom.rule.formula = `=${rule.formula}`;
    cf.custom.format.font.color = INACTIVE_GREY_HEX;
    cf.custom.format.fill.color = '#FFFFFF';
    for (const edge of ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight']) {
      const b = cf.custom.format.borders.getItem(edge);
      b.style = 'Continuous';
      b.color = '#FFFFFF';
    }
  } else {
    const cf = range.conditionalFormats.add('Custom');
    cf.custom.rule.formula = `=${rule.formula}`;
    cf.custom.format.font.color = CHECK_RED_HEX;
    cf.custom.format.font.bold = true;
  }
}

/** A row's formats: clear what the row had, then styles by name, height, links, check formatting and validation. */
export function applyRowFormat(ws: XSheet, row: number, f: RowFormat, lastCol: number, clearFirst: boolean,
  palette: LivePalette = livePalette('HF')): void {
  if (clearFirst) {
    const band = ws.getRange(`A${row}:${colLetter(lastCol)}${row}`);
    band.style = 'Normal';
    band.clear('Hyperlinks');
    band.conditionalFormats.clearAll();
    band.dataValidation.clear();
    band.unmerge();
  }
  for (const [c, fmt] of Object.entries(f.cells)) ws.getCell(row - 1, Number(c) - 1).style = fmt.style;
  if (f.height !== undefined) ws.getRange(`${row}:${row}`).format.rowHeight = f.height;
  for (const l of f.links ?? []) ws.getCell(row - 1, l.col - 1).hyperlink = { documentReference: l.to, screenTip: l.tip };
  for (const c of f.conds ?? []) addCond(ws, c.sqref, c.rule, palette);
  for (const m of f.merges ?? []) ws.getRange(`${colLetter(m.from)}${row}:${colLetter(m.to)}${row}`).merge(false);
  const ctl = f.control;
  for (const v of f.valid ?? []) {
    if (ctl?.control.kind === 'drop' && v.col === ctl.col) continue;   // an in-cell list takes the place of the position rule
    const dv = ws.getCell(row - 1, v.col - 1).dataValidation;
    dv.rule = validationRule(v.rule, `${colLetter(v.col)}${row}`);
    dv.errorAlert = { showAlert: true, style: 'Stop', title: 'Not accepted', message: v.rule.message };
  }
  if (ctl) {
    // Office.js cannot draw classic controls, so a control written live is an in-cell one on the same cell:
    // a check box keeps TRUE or FALSE; a drop-down holds the item's text and its Sel_ name reads the position.
    const cell = ws.getCell(row - 1, ctl.col - 1);
    if (ctl.control.kind === 'check') {
      cell.style = 'HFG Unit';
      cell.control = { type: 'Checkbox' };
    } else {
      cell.style = 'HFG Input Text';
      cell.dataValidation.rule = { list: { inCellDropDown: true, source: `=${ctl.control.list}` } };
      cell.dataValidation.errorAlert = { showAlert: true, style: 'Stop', title: 'Not accepted', message: 'Choose from the drop-down list.' };
    }
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

  // Report charts: Office.js types colours rather than theme slots, so they come from the entity's
  // palette, and a hatched forecast is drawn in a light tint. The package writer's charts keep the
  // theme; a Rebuild restores it.
  const meta = /<hfgModel[^>]*>([\s\S]*)<\/hfgModel>/.exec(metadataXml);
  let brand: Brand = 'HF';
  try {
    const d = JSON.parse((meta?.[1] ?? '{}').replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&'));
    brand = d?.model?.info?.entity?.brand ?? 'HF';
  } catch { /* the default palette */ }
  const theme = THEMES[brand];
  const palette = livePalette(brand);
  const hex = (slot: string, hatch = false): string | null => {
    if (slot === 'none') return null;
    if (hatch) return `#${theme.accents[5]}`;
    if (slot === 'tx2') return `#${theme.dk2}`;
    if (slot === 'grey') return '#A6A6A6';
    if (slot === 'light') return '#D9D9D9';
    const k = Number(/^accent(\d)$/.exec(slot)?.[1] ?? 1);
    return `#${theme.accents[k - 1]}`;
  };

  async function removeRChart(sheetName: string, id: string) {
    const ws = sheet(sheetName);
    ws.charts.load('items/name');
    await sync();
    for (const ch of ws.charts.items) if (ch.name === id) ch.delete();
  }

  async function drawRChart(op: Extract<PlanOp, { op: 'add_rchart' | 'set_rchart' }>) {
    const ws = sheet(op.sheet);
    const at = (ref: string) => { const [s, a] = splitRef(ref); return sheet(s).getRange(a); };
    const labels = op.series.map(s => { const r = at(s.tx); r.load('values'); return r; });
    const bars = op.series.filter(s => s.as === 'bar');
    const type = op.type === 'pie' ? 'Pie' : !bars.length ? 'Line'
      : `${op.dir === 'bar' ? 'Bar' : 'Column'}${op.grouping === 'stacked' ? 'Stacked' : 'Clustered'}`;
    const chart = ws.charts.add(type, at(op.series[0].values), 'Rows');
    chart.series.load('items');
    const cell = ws.getCell(op.at.row - 1, op.at.col - 1);
    cell.load('left,top');
    await sync();
    for (let i = (chart.series.items ?? []).length - 1; i >= 0; i--) chart.series.getItemAt(i).delete();
    op.series.forEach((s, i) => {
      const ser = chart.series.add(String(labels[i].values?.[0]?.[0] ?? ''));
      ser.setValues(at(s.values));
      ser.setXAxisValues(at(op.cats));
      if (s.as === 'line' && bars.length && op.type !== 'pie') ser.chartType = 'Line';
      const colour = hex(s.colour, s.hatch);
      if (s.as === 'line' && ser.format) {
        if (colour) ser.format.line.color = colour;
        ser.format.line.weight = s.width ?? 1.5;
        ser.format.line.lineStyle = s.dash ? 'Dash' : 'Continuous';
        ser.markerStyle = s.marker ? 'Circle' : 'None';
      } else if (ser.format) {
        if (colour) ser.format.fill.setSolidColor(colour); else ser.format.fill.clear();
      }
      if (s.as === 'bar' && op.type !== 'pie') {
        ser.gapWidth = op.gap;
        ser.overlap = op.grouping === 'stacked' ? 100 : op.overlap;
      }
      s.points?.forEach((p, j) => { const c = hex(p); if (c) ser.points?.getItemAt(j).format.fill.setSolidColor(c); });
      if (s.labels && ser.dataLabels) {
        ser.hasDataLabels = true;
        ser.dataLabels.numberFormat = s.labels.fmt;
        ser.dataLabels.showPercentage = Boolean(s.labels.pct);
        ser.dataLabels.showValue = !s.labels.pct;
        const position = ({ outEnd: 'OutsideEnd', inEnd: 'InsideEnd', ctr: 'Center', bestFit: 'BestFit' } as Record<string, string>)[s.labels.pos ?? ''];
        if (position) ser.dataLabels.position = position;
      }
    });
    chart.name = op.id;
    if (chart.title.setFormula) chart.title.setFormula(`=${op.title.ref}`); else chart.title.text = op.title.cache;
    if (chart.legend) {
      chart.legend.visible = op.legend !== null;
      if (op.legend) chart.legend.position = op.legend === 'r' ? 'Right' : 'Bottom';
    }
    if (chart.axes && op.type !== 'pie') {
      chart.axes.categoryAxis.reversePlotOrder = op.reverse;
      chart.axes.valueAxis.visible = op.valueAxis;
      chart.axes.valueAxis.numberFormat = op.yFmt;
    }
    chart.top = (cell as unknown as { top: number }).top + 3;
    chart.left = (cell as unknown as { left: number }).left + op.at.offPx * 0.75;
    chart.width = 345.8;    // 12.2 cm
    chart.height = 215.4;   // 7.6 cm
    if (op.series.some(s => s.hatch)) report.notes.push(`${op.sheet}: chart ${op.id} shows forecast months in a light tint; a Rebuild draws them hatched.`);
  }

  for (const op of ops) {
    switch (op.op) {
      case 'data_table':
        report.notes.push(`${op.sheet}: the scenario results (${op.ref}) need a data table, which Office.js cannot write; `
          + 'rebuild the workbook to show every scenario at once.');
        break;
      case 'delete_rchart':
        await removeRChart(op.sheet, op.id);
        break;
      case 'set_rchart':
        await removeRChart(op.sheet, op.id);
        await drawRChart(op);
        break;
      case 'add_rchart':
        await drawRChart(op);
        break;
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
          for (const [r, f] of Object.entries(op.format.header)) applyRowFormat(ws, Number(r), f, lastCol(op.sheet), false, palette);
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
        const ctl = op.format?.control;
        if (ctl?.control.kind === 'drop' && typeof op.cells[ctl.col] === 'number') {
          // A new drop-down written live holds its item's text (the position is the plan's value).
          const list = wb.names.getItem(ctl.control.list).getRange();
          list.load('values');
          await sync();
          const pick = Number(op.cells[ctl.col]);
          op.cells[ctl.col] = String((list.values as unknown[][])[pick - 1]?.[0] ?? '');
        }
        for (const [c, v] of Object.entries(op.cells)) writeValue(ws.getCell(op.row - 1, Number(c) - 1), v);
        if (op.format) applyRowFormat(ws, op.row, op.format, lastCol(op.sheet), op.why !== 'frame', palette);
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
        const L = colLetter(op.col);
        const R = op.toCol ? colLetter(op.toCol) : L;
        const cell = ws.getRange(`$${L}$${op.row}${op.toRow ? `:$${R}$${op.toRow}` : ''}`);
        let ref: XRange | string = cell;
        if (op.choice) {
          // A classic drop-down's cell holds a position; an in-cell one holds text, so the name reads its position.
          cell.load('values');
          await sync();
          const v = (cell.values as unknown[][])[0][0];
          if (typeof v === 'string') ref = `=MATCH(${sheetRef(op.sheet)}!$${L}$${op.row},${op.choice},0)`;
        }
        existing.set(op.name, wb.names.add(op.name, ref) as { name: string; delete(): void });
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
