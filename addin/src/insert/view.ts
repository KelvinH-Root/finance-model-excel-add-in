// Insert module in the task pane: pick a module, set its inputs, preview what the insert writes,
// then apply it to the open workbook with the live writer. Outside Excel it runs on the demo model
// so the steps can be seen.

import type { Library } from '../../../engine/src/index.ts';
import { META_NS } from '../../../engine/src/frame.ts';
import { applyPlan, type XContext } from '../live/apply.ts';
import { choices, planInsert, readModel, type InsertPlan, type OpenModel } from './core.ts';

declare const Excel: { run<T>(fn: (ctx: XContext) => Promise<T>): Promise<T> } | undefined;
declare const Office: { context?: { requirements?: { isSetSupported(set: string, version: string): boolean } } } | undefined;

const CSS = `
.ins-list { max-height: 220px; overflow: auto; border: 1px solid var(--line); border-radius: 6px; padding: 4px 6px; margin: 6px 0; }
.ins-area { font-size: 11px; color: var(--quiet); margin: 6px 0 2px; }
.ins-item { display: block; padding: 2px 0; }
.ins-item.off { color: var(--quiet); }
.ins-set { display: grid; grid-template-columns: 1fr 90px 28px; gap: 4px 6px; align-items: center; margin: 6px 0; }
.ins-set input { width: 100%; box-sizing: border-box; font: inherit; padding: 3px; border: 1px solid var(--line); border-radius: 4px; color: var(--ink); background: transparent; text-align: right; }
.ins-preview { margin: 6px 0; padding-left: 16px; }
.ins-preview li { margin: 2px 0; }
`;

function el<K extends keyof HTMLElementTagNameMap>(tag: K, attrs: Record<string, string> = {}, ...kids: (Node | string)[]) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
  for (const k of kids) e.append(k);
  return e;
}

export interface InsertAssets {
  lib: Library;
  /** The model to show when the pane is not in Excel. */
  sample: () => OpenModel;
}

const inExcel = () => typeof Excel !== 'undefined' && typeof Office !== 'undefined'
  && !!Office?.context?.requirements?.isSetSupported('ExcelApi', '1.10');

async function readOpenWorkbook(lib: Library): Promise<OpenModel> {
  const xml = await Excel!.run(async ctx => {
    const parts = ctx.workbook.customXmlParts.getByNamespace(META_NS) as unknown as {
      load(p: string): void; items: { getXml(): { value: string } }[] };
    parts.load('items');
    await ctx.sync();
    if (!parts.items.length) return '';
    const x = parts.items[0].getXml();
    await ctx.sync();
    return x.value;
  });
  if (!xml) throw new Error('This workbook has no HFG model metadata. Create a model with New model first.');
  return readModel(xml, lib);
}

export function mountInsert(host: HTMLElement, assets: InsertAssets): void {
  if (!document.getElementById('hfg-insert-css')) document.head.append(el('style', { id: 'hfg-insert-css' }, CSS));
  let open: OpenModel | null = null;
  let source = '';
  let picked = '';
  let values: Record<string, unknown> = {};
  let pending: InsertPlan | null = null;
  let message = '';

  async function load() {
    try {
      if (inExcel()) {
        open = await readOpenWorkbook(assets.lib);
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

  function pick(id: string) {
    picked = id;
    pending = null;
    const c = choices(assets.lib, open!.model).find(x => x.id === id)!;
    values = Object.fromEntries(c.settings.map(s => [s.key, s.default ?? null]));
    render();
  }

  async function insert(button: HTMLButtonElement) {
    if (!pending || !open) return;
    button.disabled = true;
    try {
      if (inExcel()) {
        const report = await Excel!.run(ctx => applyPlan(ctx, pending!.plan.ops, pending!.metadata));
        message = `${pending.title} inserted: ${report.operations} operations in ${report.syncs} round trips. ${report.notes.join(' ')}`;
      } else {
        message = `${pending.title} would be inserted with ${pending.plan.ops.length} operations; open the pane in Excel to apply it.`;
      }
      open = { model: pending.next, layout: pending.layout };
      pending = null;
      picked = '';
    } catch (e) {
      message = `The insert stopped: ${(e as Error).message}. Undo in Excel (Ctrl+Z) takes back what was written.`;
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
    const counts = `${open.model.instances.length} modules on ${open.layout.sheets.length} sheets`;
    const list = el('div', { class: 'ins-list', role: 'radiogroup', 'aria-label': 'Module' });
    let area = '';
    for (const c of choices(assets.lib, open.model)) {
      if (c.area !== area) { area = c.area; list.append(el('div', { class: 'ins-area' }, area)); }
      const r = el('input', { type: 'radio', name: 'ins-module', value: c.id });
      r.checked = picked === c.id;
      r.disabled = !!c.blocked;
      r.addEventListener('change', () => pick(c.id));
      list.append(el('label', { class: `ins-item${c.blocked ? ' off' : ''}`, title: c.blocked ?? '' }, r, ` ${c.title}`));
    }
    const parts: (Node | string)[] = [el('p', { class: 'quiet' }, `Model: ${counts}, read from ${source}.`), list];
    if (picked) {
      const c = choices(assets.lib, open.model).find(x => x.id === picked)!;
      if (c.settings.length) {
        const grid = el('div', { class: 'ins-set' });
        for (const s of c.settings) {
          // Percentages are typed as percentages (6 for 6%) and held as fractions.
          const pct = s.unit === '%';
          const shown = values[s.key] === null || values[s.key] === undefined ? '' : String(pct ? Math.round(Number(values[s.key]) * 1e6) / 1e4 : values[s.key]);
          const i = el('input', { type: 'number', step: 'any', value: shown, 'aria-label': s.label });
          i.addEventListener('input', () => { values[s.key] = i.value === '' ? null : Number(i.value) / (pct ? 100 : 1); pending = null; });
          grid.append(el('span', {}, s.label), i, el('span', { class: 'quiet' }, s.unit ?? ''));
        }
        parts.push(grid);
      }
      const preview = el('button', { type: 'button' }, 'Preview');
      preview.addEventListener('click', () => {
        try { pending = planInsert(open!, picked, values); message = ''; } catch (e) { message = (e as Error).message; }
        render();
      });
      parts.push(preview);
    }
    if (pending) {
      const ul = el('ul', { class: 'ins-preview' });
      for (const line of pending.plan.preview) ul.append(el('li', {}, line));
      const go = el('button', { type: 'button', class: 'primary' }, `Insert ${pending.title}`);
      go.addEventListener('click', () => void insert(go));
      parts.push(el('div', { class: 'card' }, el('strong', {}, `What inserting ${pending.title} writes`), ul, go));
    }
    parts.push(status);
    host.append(el('div', {}, ...parts));
  }

  render();
  void load();
}
