// The change plan: the difference between two layouts as operations, and a preview in plain words.
// The same plan builds a new workbook (package writer) or is applied to an open one (live writer).

import { AssemblyError, TOTAL_COL } from './frame.ts';
import type { ChartSpec, Layout, LRow, RowKind } from './layout.ts';
import { chartRefs, frameCells, rowCells, type ChartRefs, type Dialect, type FrameCell } from './render.ts';

export type PlanOp =
  | { op: 'delete_name'; name: string }
  | { op: 'delete_sheet'; sheet: string }
  | { op: 'add_sheet'; sheet: string; index: number; periods: number; frame: FrameCell[] }
  | { op: 'delete_rows' | 'insert_rows'; sheet: string; row: number; count: number }
  | { op: 'write'; sheet: string; row: number; why: 'frame' | 'new' | 'rewire'; kind: RowKind | 'frame'; style: string;
      unit: string; cells: Record<number, unknown> }
  | { op: 'add_name'; name: string; sheet: string; row: number; col: number }
  | { op: 'delete_chart'; sheet: string; title: string }
  | ({ op: 'add_chart' | 'set_chart' } & ChartRefs);

export interface Plan {
  ops: PlanOp[];
  preview: string[];
}

/** Consecutive numbers as [start, count] runs. */
function runs(nums: number[]): [number, number][] {
  const out: [number, number][] = [];
  for (const n of nums) {
    const last = out[out.length - 1];
    if (last && n === last[0] + last[1]) last[1] += 1;
    else out.push([n, 1]);
  }
  return out;
}

function seriesWords(c: ChartSpec): string {
  const cols = c.series.filter(([, k]) => k === 'column').length;
  const lines = c.series.length - cols;
  const words: string[] = [];
  if (cols) words.push(cols > 1 ? `${cols} column series` : '1 column series');
  if (lines) words.push(`${lines} line` + (lines > 1 ? 's' : ''));
  return words.join(' and ') || 'no series';
}

const plural = (n: number, word: string) => `${n} ${word}${n > 1 ? 's' : ''}`;

function frameMap(layout: Layout, sheet: string, dialect: Dialect): Map<string, [number, number, unknown]> {
  return new Map(frameCells(layout, sheet, dialect).map(cell => [`${cell[0]},${cell[1]}`, cell]));
}

export function planChange(old: Layout, nw: Layout, dialect: Dialect = 'excel'): Plan {
  const ops: PlanOp[] = [];
  const oldSheets = old.sheetRows();
  const newSheets = nw.sheetRows();
  const pos = nw.positions();
  const preview: string[] = [];

  for (const nm of [...old.names.keys()].filter(n => !nw.names.has(n)).sort()) ops.push({ op: 'delete_name', name: nm });
  for (const s of oldSheets.keys()) {
    if (!newSheets.has(s)) {
      ops.push({ op: 'delete_sheet', sheet: s });
      preview.push(`${s}: sheet removed (no module left on it).`);
    }
  }
  nw.sheets.forEach(([s], idx) => {
    if (oldSheets.has(s)) return;
    ops.push({ op: 'add_sheet', sheet: s, index: idx, periods: nw.periods, frame: frameCells(nw, s, dialect) });
    const kind = nw.kinds[s];
    preview.push(`${s}: new ${kind === 'cover' ? 'section cover' : 'sheet'}` + (kind === 'cover' ? ` (${nw.titles[s]}).` : '.'));
  });
  for (const s of newSheets.keys()) {
    if (!oldSheets.has(s)) continue;
    const a = frameMap(old, s, dialect);
    const b = frameMap(nw, s, dialect);
    const val = (m: Map<string, [number, number, unknown]>, k: string) => (m.has(k) ? m.get(k)![2] : null);
    const differ = [...new Set([...a.keys(), ...b.keys()])].filter(k => val(a, k) !== val(b, k));
    const rowsOf = [...new Set(differ.map(k => Number(k.split(',')[0])))].sort((x, y) => x - y);
    for (const r of rowsOf) {
      const cells: Record<number, unknown> = {};
      for (const k of differ) {
        const [rr, c] = k.split(',').map(Number);
        if (rr === r) cells[c] = val(b, k);
      }
      ops.push({ op: 'write', sheet: s, row: r, why: 'frame', kind: 'frame', style: '', unit: '', cells });
    }
  }

  const changedRows = new Map<string, [number, LRow, 'new' | 'rewire'][]>();
  for (const [s, rows] of nw.sheets) {
    const newIds = rows.map(r => r.id);
    const oldRows = oldSheets.get(s);
    const first = nw.firstRowOf(s);
    if (!oldRows) {
      changedRows.set(s, rows.map((r, k) => [first + k, r, 'new']));
      continue;
    }
    const oldFirst = old.firstRowOf(s);
    const oldIds = oldRows.map(r => r.id);
    const oldSet = new Set(oldIds);
    const newSet = new Set(newIds);
    const kept = oldIds.filter(i => newSet.has(i));
    const keptNew = newIds.filter(i => oldSet.has(i));
    if (kept.length !== keptNew.length || kept.some((id, k) => id !== keptNew[k])) {
      throw new AssemblyError(`${s}: rows would change order, which needs a move the plan does not support`);
    }
    const gone = oldIds.flatMap((id, k) => (newSet.has(id) ? [] : [oldFirst + k]));
    for (const [start, count] of runs(gone).reverse()) ops.push({ op: 'delete_rows', sheet: s, row: start, count });
    const added = newIds.flatMap((id, k) => (oldSet.has(id) ? [] : [first + k]));
    for (const [start, count] of runs(added)) ops.push({ op: 'insert_rows', sheet: s, row: start, count });
    const oldById = new Map(oldRows.map(r => [r.id, r]));
    const list: [number, LRow, 'new' | 'rewire'][] = [];
    rows.forEach((r, k) => {
      if (!oldSet.has(r.id)) list.push([first + k, r, 'new']);
      else if (r.signature() !== oldById.get(r.id)!.signature()) list.push([first + k, r, 'rewire']);
    });
    changedRows.set(s, list);
  }

  const oldLinks = new Map<string, string | null>();
  for (const [, rows] of old.sheets) for (const r of rows) if (r.kind === 'setting') oldLinks.set(r.id, r.link);
  for (const [s, list] of changedRows) {
    for (const [rownum, r, why] of list) {
      const cells = rowCells(nw, s, r, rownum, pos, dialect);
      if (why === 'rewire' && r.kind === 'setting' && !r.link && !oldLinks.get(r.id)) {
        delete cells[TOTAL_COL];   // keep the input someone typed
      }
      ops.push({ op: 'write', sheet: s, row: rownum, why, kind: r.kind, style: r.style, unit: r.unit, cells });
    }
  }

  for (const [nm, rid] of nw.names) {
    if (old.names.get(nm) !== rid || old.nameCol(nm) !== nw.nameCol(nm)) {
      const at = pos.get(rid);
      if (!at) throw new AssemblyError(`name ${nm} points at ${rid}, which is not in the layout`);
      ops.push({ op: 'add_name', name: nm, sheet: at[0], row: at[1], col: nw.nameCol(nm) });
    }
  }

  // Charts last, once every row is where it ends up. Ranges shift with inserted and deleted
  // rows on their own; a chart is rewritten only when the rows it shows change.
  const oldCharts = new Map(old.charts.map(c => [c.id, c]));
  const newCharts = new Map(nw.charts.map(c => [c.id, c]));
  const chartPreview: string[] = [];
  for (const [cid, c] of oldCharts) {
    if (!newCharts.has(cid) && newSheets.has(c.sheet)) {
      ops.unshift({ op: 'delete_chart', sheet: c.sheet, title: c.title });
      chartPreview.push(`${c.sheet}: chart ${c.title} removed.`);
    }
  }
  for (const [cid, c] of newCharts) {
    const was = oldCharts.get(cid);
    if (!was) {
      ops.push({ op: 'add_chart', ...chartRefs(nw, c, dialect) });
      chartPreview.push(`${c.sheet}: chart ${c.title} added (${seriesWords(c)}).`);
    } else if (c.signature() !== was.signature()) {
      ops.push({ op: 'set_chart', ...chartRefs(nw, c, dialect) });
      chartPreview.push(`${c.sheet}: chart ${c.title} re-pointed (${seriesWords(c)}, was ${seriesWords(was)}).`);
    }
  }

  // Preview, in plain words, before anything is touched.
  for (const [s, list] of changedRows) {
    const newN = list.filter(([, r, w]) => w === 'new' && r.kind !== 'blank').length;
    const rew = list.filter(([, , w]) => w === 'rewire').map(([, r]) => r.label);
    if (newN) {
      const spans = runs(list.filter(([, , w]) => w === 'new').map(([n]) => n))
        .map(([a, n]) => (n === 1 ? `${a}` : `${a} to ${a + n - 1}`));
      const where = spans.length === 1 ? spans[0] : spans.slice(0, -1).join(', ') + ' and ' + spans[spans.length - 1];
      preview.push(`${s}: ${plural(newN, 'new row')} (row${newN > 1 ? 's' : ''} ${where}).`);
    }
    if (rew.length) preview.push(`${s}: ${plural(rew.length, 'row')} rewired (${rew.join(', ')}).`);
  }
  for (const [s, rows] of old.sheets) {
    const now = newSheets.get(s);
    if (!now) continue;
    const ids = new Set(now.map(x => x.id));
    const gone = rows.filter(r => !ids.has(r.id) && r.kind !== 'blank');
    if (gone.length) preview.push(`${s}: ${plural(gone.length, 'row')} removed.`);
  }
  preview.push(...chartPreview);
  const key = (rec: { link: string; from: string; to: string }) => `${rec.link}\u0000${rec.from}\u0000${rec.to}`;
  const oldKeys = new Set(old.records.map(key));
  const newKeys = new Set(nw.records.map(key));
  const addedRecs = nw.records.filter(r => !oldKeys.has(key(r)));
  const removedRecs = old.records.filter(r => !newKeys.has(key(r)));
  for (const [label, recs] of [['New links', addedRecs], ['Links removed', removedRecs]] as const) {
    if (!recs.length) continue;
    const byLink = new Map<string, number>();
    for (const rec of recs) byLink.set(rec.link, (byLink.get(rec.link) || 0) + 1);
    const parts = [...byLink.entries()].sort((x, y) => (x[0] < y[0] ? -1 : x[0] > y[0] ? 1 : 0)).map(([k, v]) => `${k} x${v}`);
    preview.push(`${label}: ${parts.join(', ')}.`);
  }
  for (const w of nw.warnings) if (!old.warnings.includes(w)) preview.push(`Warning: ${w}.`);
  return { ops, preview };
}
