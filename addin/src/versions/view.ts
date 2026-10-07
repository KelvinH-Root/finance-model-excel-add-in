// Saved versions in the task pane: Save version, Approve budget, Lock or unlock. Each previews the
// change plan first, then applies it to the open workbook with the live writer. Outside Excel it
// runs on the demo model so the steps can be seen.

import type { Libraries } from '../wizard/core.ts';
import { applyPlan, type XContext } from '../live/apply.ts';
import type { InsertPlan, OpenModel } from '../insert/core.ts';
import { el, inExcel, readOpenWorkbook } from '../insert/view.ts';
import {
  BUDGET_LINES, lineAddresses, planApprove, planLock, planSave, REFORECAST_LINES, suggestions, versionsOf, type SaveChoice,
} from './core.ts';

declare const Excel: { run<T>(fn: (ctx: XContext) => Promise<T>): Promise<T> } | undefined;

const CSS = `
.ver-table { width: 100%; border-collapse: collapse; margin: 6px 0; font-size: 12px; }
.ver-table th, .ver-table td { text-align: left; padding: 2px 4px; border-bottom: 1px solid var(--line); }
.ver-table button { font-size: 11px; padding: 1px 6px; }
.ver-choice { display: block; margin: 3px 0; }
`;

export interface VersionAssets {
  libs: Libraries;
  sample: () => OpenModel;
}

/** The statements' values for the lines a version keeps, read from the open workbook. */
async function readValues(open: OpenModel, keys: string[]): Promise<{ values: Record<string, number[]>; errors: number }> {
  const addr = lineAddresses(open.layout, open.model, keys);
  return Excel!.run(async ctx => {
    const ranges = Object.entries(addr).map(([k, a]) => {
      const [s, cells] = [a.slice(1, a.indexOf("'!")), a.slice(a.indexOf("'!") + 2)];
      const r = ctx.workbook.worksheets.getItem(s).getRange(cells);
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
}

export function mountVersions(host: HTMLElement, assets: VersionAssets): void {
  if (!document.getElementById('hfg-versions-css')) document.head.append(el('style', { id: 'hfg-versions-css' }, CSS));
  let open: OpenModel | null = null;
  let source = '';
  let message = '';
  let picked = 0;
  let savedBy = 'Group Finance';
  let pending: InsertPlan | null = null;

  async function load() {
    try {
      if (inExcel()) {
        open = await readOpenWorkbook(assets.libs);
        source = 'the open workbook';
      } else {
        open = assets.sample();
        source = 'the demo model (Excel is not running this pane)';
      }
    } catch (e) {
      message = (e as Error).message;
    }
    render();
  }

  async function preview(make: () => Promise<InsertPlan> | InsertPlan) {
    try {
      pending = await make();
      message = '';
    } catch (e) {
      message = (e as Error).message;
      pending = null;
    }
    render();
  }

  async function save(choice: SaveChoice) {
    const keys = choice.type === 'Budget' ? BUDGET_LINES : REFORECAST_LINES;
    let values: Record<string, number[]>;
    if (inExcel()) {
      const read = await readValues(open!, keys);
      if (read.errors > 0) throw new Error(`The model has ${read.errors} error check${read.errors > 1 ? 's' : ''} failing; clear them before saving a version.`);
      values = read.values;
    } else {
      // outside Excel: the latest saved reforecast stands in for the statements' values
      const last = [...versionsOf(open!.model)].reverse().find(v => v.type === 'Reforecast');
      values = Object.fromEntries(keys.map(k => [k, (last?.values[k] ?? []).map(v => v ?? 0).concat(Array(open!.model.periods).fill(0)).slice(0, open!.model.periods)]));
    }
    const today = new Date().toISOString().slice(0, 10);
    return planSave(open!, choice, values, savedBy, today);
  }

  async function apply(button: HTMLButtonElement) {
    if (!pending || !open) return;
    button.disabled = true;
    try {
      if (inExcel()) {
        const report = await Excel!.run(ctx => applyPlan(ctx, pending!.plan.ops, pending!.metadata));
        message = `${pending.title}: ${report.operations} operations in ${report.syncs} round trips. ${report.notes.join(' ')}`;
      } else {
        message = `${pending.title}: ${pending.plan.ops.length} operations; open the pane in Excel to apply them.`;
      }
      open = { model: pending.next, layout: pending.layout };
      pending = null;
    } catch (e) {
      message = `Stopped: ${(e as Error).message}. Undo in Excel (Ctrl+Z) takes back what was written.`;
    }
    render();
  }

  function render() {
    host.innerHTML = '';
    const status = el('p', { class: 'quiet', role: 'status' }, message);
    if (!open) {
      host.append(status.textContent ? status : el('p', { class: 'quiet' }, 'Reading the model...'));
      return;
    }
    const versions = versionsOf(open.model);
    const parts: (Node | string)[] = [el('p', { class: 'quiet' }, `${versions.length} saved versions, read from ${source}.`)];

    // Save version
    const opts = suggestions(open.model);
    const box = el('div', { class: 'card' }, el('strong', {}, 'Save version'));
    opts.forEach((o, k) => {
      const r = el('input', { type: 'radio', name: 'ver-choice', value: String(k) });
      r.checked = picked === k;
      r.addEventListener('change', () => { picked = k; pending = null; render(); });
      const note = o.replaces ? ` (replaces the one saved before)` : '';
      box.append(el('label', { class: 'ver-choice' }, r, ` ${o.label}${note}`));
    });
    const by = el('input', { type: 'text', value: savedBy, 'aria-label': 'Saved by' });
    by.addEventListener('input', () => { savedBy = by.value; });
    const go = el('button', { type: 'button' }, 'Preview');
    go.addEventListener('click', () => void preview(() => save(opts[picked])));
    box.append(el('label', {}, 'Saved by ', by), ' ', go);
    parts.push(box);

    // The register
    const table = el('table', { class: 'ver-table' }, el('tr', {}, el('th', {}, 'Version'), el('th', {}, 'Status'), el('th', {}, '')));
    for (const v of [...versions].reverse()) {
      const actions = el('td', {});
      if (v.type === 'Budget' && v.status !== 'Approved') {
        const b = el('button', { type: 'button' }, 'Approve');
        b.addEventListener('click', () => void preview(() => planApprove(open!, v.id)));
        actions.append(b, ' ');
      }
      const l = el('button', { type: 'button' }, v.locked ? 'Unlock' : 'Lock');
      l.addEventListener('click', () => void preview(() => planLock(open!, v.id, !v.locked)));
      actions.append(l);
      table.append(el('tr', {}, el('td', {}, v.label), el('td', {}, `${v.status}${v.locked ? ', locked' : ''}`), actions));
    }
    parts.push(el('div', { class: 'card' }, el('strong', {}, 'Saved versions'), table));

    if (pending) {
      const ul = el('ul', { class: 'ins-preview' });
      for (const line of pending.plan.preview) ul.append(el('li', {}, line));
      const ok = el('button', { type: 'button', class: 'primary' }, 'Apply');
      ok.addEventListener('click', () => void apply(ok));
      parts.push(el('div', { class: 'card' }, el('strong', {}, `What "${pending.title}" writes`), ul, ok));
    }
    parts.push(status);
    host.append(el('div', {}, ...parts));
  }

  render();
  void load();
}
