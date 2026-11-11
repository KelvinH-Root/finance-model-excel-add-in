// Key outputs, the before-and-after check and the change log, on every command that changes the
// model (the add-in's port of prototypes/assurance/guard.py). The service reads the key outputs
// (the statement rows marked `headline` in the library: closing cash, peak debt, profit and the
// like) before a plan is applied and again after Excel recalculates, works out which of them the
// plan could reach from the rows it writes, and logs one line in the model's metadata: when, the
// command, what it wrote, the key outputs before and after, what moved and what moved without
// cause. Nothing here knows about revenue or debt; reach is followed through the layout's formulas.
// DOM free.

import type { Headline, Layout, Model, Plan } from '../../../engine/src/index.ts';

/** Half a cent: a smaller movement is rounding. */
export const TOLERANCE = 0.005;

/** Commands that must not move a number ('none'), and those whose reach is worked out from the plan. */
export type Effect = 'none' | 'reach';

const MARK = /«([RPQABSVOW]|P\d+|C\d+)\|([^»|]+)»/g;
const BLOCK = /«([GK])\|([^»|]+)\|([^»|]+)(?:\|\d+)?»/g;
const WORD = /[A-Za-z_][A-Za-z0-9_]*/g;

/** Row id -> the rows that read it, by marker, block of rows or defined name. */
export function dependents(layout: Layout): Map<string, Set<string>> {
  const order = new Map(layout.sheets.map(([s, rows]) => [s, rows.map(r => r.id)]));
  const where = new Map<string, [string, number]>();
  for (const [s, ids] of order) ids.forEach((id, k) => where.set(id, [s, k]));
  const out = new Map<string, Set<string>>();
  const add = (src: string, reader: string) => { let s = out.get(src); if (!s) out.set(src, s = new Set()); s.add(reader); };
  for (const [, rows] of layout.sheets) {
    for (const r of rows) {
      const texts = [r.first, r.formula, r.link, ...Object.values(r.cells ?? {})]
        .filter((t): t is string => typeof t === 'string' && t.includes('«') || (typeof t === 'string' && t.startsWith('=')));
      for (const t of texts) {
        for (const m of t.matchAll(MARK)) add(m[2], r.id);
        for (const m of t.matchAll(BLOCK)) {
          const a = where.get(m[2]);
          const b = where.get(m[3]);
          if (!a || !b || a[0] !== b[0]) continue;
          for (const id of order.get(a[0])!.slice(Math.min(a[1], b[1]), Math.max(a[1], b[1]) + 1)) add(id, r.id);
        }
        for (const w of t.replace(MARK, '').replace(BLOCK, '').match(WORD) ?? []) {
          const src = layout.names.get(w);
          if (src) add(src, r.id);
        }
      }
    }
  }
  return out;
}

/** The rows a plan writes (new or rewired), as row ids in the new layout. */
export function writtenRows(plan: Plan, layout: Layout): Set<string> {
  const at = new Map<string, string>();
  for (const [id, [s, r]] of layout.positions()) if (!id.startsWith('@')) at.set(`${s}|${r}`, id);
  const out = new Set<string>();
  for (const op of plan.ops) {
    if (op.op !== 'write' || (op.why !== 'new' && op.why !== 'rewire')) continue;
    const id = at.get(`${op.sheet}|${op.row}`);
    if (id) out.add(id);
  }
  return out;
}

/** Every row a change to the start rows can reach through the formulas. */
export function reachable(layout: Layout, start: Set<string>): Set<string> {
  const deps = dependents(layout);
  const seen = new Set(start);
  const todo = [...start];
  while (todo.length) {
    for (const next of deps.get(todo.pop()!) ?? []) {
      if (!seen.has(next)) { seen.add(next); todo.push(next); }
    }
  }
  return seen;
}

/**
 * The key outputs a change may move. A removed row is not in the new layout, so a plan that deletes
 * rows may move anything that read them: every key output, unless the command moves none.
 */
export function expectedOutputs(layout: Layout, plan: Plan, effect: Effect): string[] {
  if (effect === 'none') return [];
  if (plan.ops.some(o => o.op === 'delete_rows' || o.op === 'delete_sheet')) return layout.headlines.map(h => h.label);
  const reach = reachable(layout, writtenRows(plan, layout));
  return layout.headlines.filter(h => reach.has(h.id)).map(h => h.label);
}

/** A key output from its row's values across the timeline, as the contents' KO_ cell works it out. */
export function measure(h: Headline, values: unknown[], last: number): number {
  const nums = values.map(v => (typeof v === 'number' ? v : 0));
  if (!nums.length) return 0;
  if (h.measure === 'sum') return nums.reduce((a, b) => a + b, 0);
  if (h.measure === 'min') return Math.min(...nums);
  if (h.measure === 'max') return Math.max(...nums);
  return nums[Math.min(Math.max(last, 1), nums.length) - 1];
}

export interface LogEntry {
  /** ISO date and time. */
  when: string;
  who: string;
  command: string;
  title: string;
  effect: Effect;
  /** What the preview said it would write. */
  preview: string[];
  /** Key output label -> [before, after]. */
  outputs: Record<string, [number, number]>;
  expected: string[];
  moved: string[];
  unexplained: string[];
  /** A reason given where the command asks for one (unlocking an approved budget). */
  reason?: string;
}

/** The log entry for a change, from the key outputs read before and after it. */
export function entryFor(a: {
  command: string; title: string; effect: Effect; plan: Plan; layout: Layout;
  before: Record<string, number>; after: Record<string, number>; who?: string; reason?: string; when?: Date;
}): LogEntry {
  const expected = expectedOutputs(a.layout, a.plan, a.effect);
  const labels = [...new Set([...Object.keys(a.before), ...Object.keys(a.after)])];
  const moved = labels.filter(k => Math.abs((a.after[k] ?? 0) - (a.before[k] ?? 0)) > TOLERANCE);
  const round = (v: number | undefined) => Math.round((v ?? 0) * 100) / 100;
  return {
    when: (a.when ?? new Date()).toISOString(), who: a.who ?? '', command: a.command, title: a.title, effect: a.effect,
    preview: a.plan.preview.slice(0, 12),
    outputs: Object.fromEntries(labels.map(k => [k, [round(a.before[k]), round(a.after[k])]])),
    expected, moved, unexplained: moved.filter(k => !expected.includes(k)),
    ...(a.reason ? { reason: a.reason } : {}),
  };
}

const money = (v: number) => v.toLocaleString('en-NZ', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

/** The lines a result shows: a movement without cause first, then each key output that moved. */
export function resultLines(e: LogEntry): string[] {
  const out: string[] = [];
  if (e.unexplained.length) {
    const n = e.unexplained;
    const listed = n.length === 1 ? n[0] : `${n.slice(0, -1).join(', ')} and ${n[n.length - 1]}`;
    out.push(`Check this change: ${listed} moved, ${e.effect === 'none' ? 'and this command should not move any key output.'
      : `but nothing this command changed feeds ${n.length === 1 ? 'it' : 'them'}.`}`);
  }
  for (const k of e.moved) {
    const [a, b] = e.outputs[k];
    out.push(`${k}: ${money(a)} to ${money(b)} (${b - a >= 0 ? '+' : ''}${money(b - a)})`);
  }
  if (!e.moved.length && Object.keys(e.outputs).length) out.push('No key output moved.');
  return out;
}

/** How many entries the log keeps in the metadata; older ones drop off. */
export const LOG_LIMIT = 500;

/** The model's change log (newest last). */
export function changeLog(model: Model): LogEntry[] {
  return (model.assurance.log as LogEntry[] | undefined) ?? [];
}

/** Add an entry to the model's change log, keeping the newest LOG_LIMIT. */
export function logChange(model: Model, e: LogEntry): void {
  const log = [...changeLog(model), e];
  model.assurance.log = log.slice(-LOG_LIMIT);
}
