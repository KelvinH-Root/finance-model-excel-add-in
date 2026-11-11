// The cells a rewired row held before a change and does not hold after it. A fresh build never
// has them, but the live writer only writes a row's new cells, so without this a formula left
// behind (a report table's ranking column, say) would read #REF! once the rows it read have gone.
// The plan stays the engine's; the live writer clears the columns listed here before it writes.

import { rowCells, type Layout, type LRow, type Plan, type PlanOp } from '../../../engine/src/index.ts';

export type WriteOp = Extract<PlanOp, { op: 'write' }> & { clear?: number[] };

/** Mark each rewired row's write with the columns its old version held and its new one does not. */
export function withClears(old: Layout, nw: Layout, plan: Plan): Plan {
  const oldPos = old.positions();
  const newPos = nw.positions();
  const before = new Map<string, [string, LRow, number]>();
  for (const [s, rows] of old.sheets) {
    const first = old.firstRowOf(s);
    rows.forEach((r, k) => before.set(r.id, [s, r, first + k]));
  }
  const at = new Map<string, LRow>();
  for (const [s, rows] of nw.sheets) {
    const first = nw.firstRowOf(s);
    rows.forEach((r, k) => at.set(`${s}|${first + k}`, r));
  }
  for (const op of plan.ops) {
    if (op.op !== 'write' || op.why !== 'rewire') continue;
    const now = at.get(`${op.sheet}|${op.row}`);
    const was = now && before.get(now.id);
    if (!now || !was) continue;
    const held = Object.keys(rowCells(old, was[0], was[1], was[2], oldPos, 'excel')).map(Number);
    const holds = new Set(Object.keys(rowCells(nw, op.sheet, now, op.row, newPos, 'excel')).map(Number));
    const gone = held.filter(c => !holds.has(c));
    if (gone.length) (op as WriteOp).clear = gone;
  }
  return plan;
}
