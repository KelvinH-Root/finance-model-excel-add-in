// New model wizard in the task pane: one step at a time, checks before moving on, a preview of
// the contents before Create. Create builds the workbook with the engine's package writer and
// opens it with Excel.createWorkbook (ExcelApi 1.8); outside Excel it offers the file to download.

import { THEMES, type Brand, type Library, type Logo } from '../../../engine/src/index.ts';
import {
  BRANDS, buildFile, DENOMINATIONS, fileName, initialState, MONTHS, periodMonth, preview, problems, RECIPES, recipeOf, STEPS,
  toBase64, type Libraries, type WizardState,
} from './core.ts';

export interface WizardAssets {
  lib: Library | Libraries;
  logos: Partial<Record<Brand, Logo & { dataUrl: string }>>;
  today?: Date;
}

declare const Office: { context?: { requirements?: { isSetSupported(set: string, version: string): boolean } } } | undefined;
declare const Excel: { createWorkbook(base64: string): Promise<void> } | undefined;

const CSS = `
.wiz { margin: 8px 0; }
.wsteps { display: flex; gap: 2px; list-style: none; padding: 0; margin: 0 0 10px; }
.wsteps li { flex: 1 1 auto; text-align: center; white-space: nowrap; font-size: 10.5px; color: var(--quiet);
  border-bottom: 2px solid var(--line); padding: 2px 2px 3px; }
.wsteps li.on { color: var(--ink); font-weight: 600; border-color: var(--accent-line); }
.wsteps li.done { color: var(--accent-text); }
.wfield { margin: 8px 0; }
.wfield > span { display: block; font-weight: 600; margin-bottom: 2px; }
.wfield small { display: block; color: var(--quiet); margin-top: 2px; }
.wfield input[type=text], .wfield input[type=month], .wfield input[type=number], .wfield select, .wfield textarea {
  width: 100%; box-sizing: border-box; font: inherit; padding: 4px; border: 1px solid var(--line); border-radius: 4px;
  color: var(--ink); background: transparent; }
.wfield textarea { min-height: 54px; resize: vertical; }
.wgrid { display: grid; grid-template-columns: 1fr 1fr; gap: 0 10px; }
.wbrands { display: grid; grid-template-columns: 1fr; gap: 6px; }
.wbrand { display: grid; grid-template-columns: 14px 90px 1fr; align-items: center; gap: 8px; padding: 6px 8px; margin: 0;
  border: 1px solid var(--line); border-radius: 6px; cursor: pointer; }
.wbrand.on { border-color: var(--accent-line); background: var(--accent-tint); }
.wlogo { height: 22px; display: flex; align-items: center; justify-content: center; border-radius: 3px; }
.wlogo.dark { background: var(--wtile, #0a132e); }
.wlogo img { max-height: 16px; max-width: 82px; }
.wbar { height: 6px; border-radius: 3px; margin-top: 3px; }
.wrecipe { display: block; padding: 6px 8px; margin: 4px 0; border: 1px solid var(--line); border-radius: 6px; cursor: pointer; }
.wrecipe.on { border-color: var(--accent-line); background: var(--accent-tint); }
.wrecipe small { display: block; color: var(--quiet); }
.werr { color: var(--fail); margin: 6px 0; padding-left: 16px; }
.wactions { display: flex; gap: 6px; margin-top: 10px; }
.wreview dl { display: grid; grid-template-columns: max-content 1fr; gap: 2px 10px; margin: 4px 0 8px; }
.wreview dt { color: var(--quiet); }
.wreview dd { margin: 0; }
.wcontents { margin: 4px 0; padding-left: 18px; }
.wcontents ol { list-style: lower-alpha; padding-left: 18px; }
`;

function el<K extends keyof HTMLElementTagNameMap>(tag: K, attrs: Record<string, string> = {}, ...kids: (Node | string)[]) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
  for (const k of kids) e.append(k);
  return e;
}

function field(label: string, control: HTMLElement, hint?: string): HTMLElement {
  const f = el('label', { class: 'wfield' }, el('span', {}, label), control);
  if (hint) f.append(el('small', {}, hint));
  return f;
}

export function mountWizard(host: HTMLElement, assets: WizardAssets): { state: WizardState; rerender: () => void } {
  if (!document.getElementById('hfg-wizard-css')) document.head.append(el('style', { id: 'hfg-wizard-css' }, CSS));
  const state = initialState(assets.today ?? new Date());
  let message = '';
  let download: { url: string; name: string } | null = null;

  const set = <K extends keyof WizardState>(k: K, v: WizardState[K], redraw = false) => {
    state[k] = v;
    if (redraw) render();
  };

  function input(type: string, value: string, onChange: (v: string) => void, attrs: Record<string, string> = {}) {
    const i = el('input', { type, value, ...attrs });
    i.addEventListener('input', () => onChange(i.value));
    return i;
  }

  function select(options: [string, string][], value: string, onChange: (v: string) => void) {
    const s = el('select');
    for (const [v, label] of options) {
      const o = el('option', { value: v }, label);
      if (v === value) o.selected = true;
      s.append(o);
    }
    s.addEventListener('change', () => onChange(s.value));
    return s;
  }

  function entityStep(): HTMLElement {
    const list = el('div', { class: 'wbrands', role: 'radiogroup', 'aria-label': 'Entity' });
    for (const b of BRANDS) {
      const t = THEMES[b];
      const logo = assets.logos[b];
      const tile = el('span', { class: `wlogo${t.logo.onDark ? ' dark' : ''}` });
      if (logo) tile.append(el('img', { src: logo.dataUrl, alt: `${t.name} logo` }));
      const radio = el('input', { type: 'radio', name: 'wbrand', value: b, 'aria-label': t.name });
      radio.checked = state.brand === b;
      radio.addEventListener('change', () => {
        const wasDefault = BRANDS.some(x => THEMES[x].name === state.entityName) || !state.entityName.trim();
        state.brand = b;
        if (wasDefault) state.entityName = t.name;
        render();
      });
      const bar = el('div', { class: 'wbar', style: `background: linear-gradient(90deg, ${t.accents.map((c, i) => `#${c} ${i * 16.6}% ${(i + 1) * 16.6}%`).join(', ')})` });
      list.append(el('label', { class: `wbrand${state.brand === b ? ' on' : ''}` }, radio, tile, el('span', {}, el('strong', {}, b), ` ${t.name}`, bar)));
    }
    return el('div', {},
      el('p', { class: 'quiet' }, 'The entity sets the name on the contents, the logo and the theme every style and chart takes its colours from. A model keeps its entity.'),
      list,
      field('Entity name', input('text', state.entityName, v => set('entityName', v), { maxlength: '80' }),
        'As it should appear at the top of the contents, for example a development LP that uses its parent\'s brand.'));
  }

  function modelStep(): HTMLElement {
    const recipes = el('div', { role: 'radiogroup', 'aria-label': 'Starts from' });
    for (const r of RECIPES) {
      const radio = el('input', { type: 'radio', name: 'wrecipe', value: r.id });
      radio.checked = state.recipe === r.id;
      radio.addEventListener('change', () => {
        if (r.model && !state.title.trim()) state.title = r.model.title;   // a whole model brings its own title
        set('recipe', r.id, true);
      });
      recipes.append(el('label', { class: `wrecipe${state.recipe === r.id ? ' on' : ''}` }, radio, ` ${r.label}`, el('small', {}, r.note)));
    }
    const notes = el('textarea', { rows: '3' }, state.notes);
    notes.addEventListener('input', () => set('notes', notes.value));
    return el('div', {},
      field('Model title', input('text', state.title, v => set('title', v), { maxlength: '120', placeholder: 'For example, Operating budget FY2027' }),
        'Shown on every sheet with the error and alert counts.'),
      el('div', { class: 'wfield' }, el('span', {}, 'Starts from'), recipes,
        el('small', {}, 'Model types from the catalogue join as their modules reach the library.')),
      field('Prepared by line', input('text', state.preparedBy, v => set('preparedBy', v), { maxlength: '120' }), 'The third line of the contents header.'),
      field('Notes', notes, 'One note per line; they are numbered on the contents.'));
  }

  function timelineStep(): HTMLElement {
    const own = recipeOf(state)?.model;
    if (own) {
      return el('div', {},
        el('p', {}, 'This model carries its own history and drivers, so it keeps its timeline:'),
        el('p', { class: 'quiet' }, preview(state, assets.lib).timeline),
        el('p', { class: 'quiet' }, 'Change the last month of actuals or the budget window on the Settings sheet once it is open; the drop-downs move every sheet.'));
    }
    const hint = el('small', {});
    const lastHint = () => { hint.textContent = state.lastActual ? `Actuals to ${periodMonth(state.start, state.lastActual)}.` : 'No actual months yet.'; };
    lastHint();
    const last = input('number', String(state.lastActual), v => { set('lastActual', Number(v)); lastHint(); }, { min: '0', step: '1' });
    const lastField = el('label', { class: 'wfield' }, el('span', {}, 'Last month of actuals (period)'), last, hint);
    return el('div', {},
      el('div', { class: 'wgrid' },
        field('First month', input('month', state.start, v => { set('start', v); lastHint(); })),
        field('Months', input('number', String(state.months), v => set('months', Number(v)), { min: '1', max: '600', step: '1' }))),
      el('div', { class: 'wgrid' },
        field('Financial year ends in', select(MONTHS.map((m, i) => [String(i + 1), m]), String(state.fyEndMonth), v => set('fyEndMonth', Number(v)))),
        field('Denomination', select(DENOMINATIONS.map(d => [d, d]), state.denomination, v => set('denomination', v as WizardState['denomination'])))),
      lastField,
      el('div', { class: 'wgrid' },
        field('Budget starts (period, 0 for after the actuals)', input('number', String(state.budgetFirst), v => set('budgetFirst', Number(v)), { min: '0', step: '1' })),
        field('Budget months', input('number', String(state.budgetMonths), v => set('budgetMonths', Number(v)), { min: '1', step: '1' }))),
      el('p', { class: 'quiet' }, 'These go on the Settings sheet as drop-downs, where they can be changed later; every sheet\'s timeline reads them.'));
  }

  function displayStep(): HTMLElement {
    const box = (label: string, key: 'showErrors' | 'showAlerts') => {
      const c = el('input', { type: 'checkbox' });
      c.checked = state[key];
      c.addEventListener('change', () => set(key, c.checked));
      return el('label', { class: 'wfield' }, c, ` ${label}`);
    };
    return el('div', {},
      el('p', { class: 'quiet' }, 'The model name line on every sheet can say how many checks are failing, so a printed page shows the model\'s state.'),
      box('Show the error count in the model name', 'showErrors'),
      box('Show the alert count in the model name', 'showAlerts'));
  }

  function reviewStep(): HTMLElement {
    const errs = problems(state);
    if (errs.length) return el('div', {}, el('p', {}, 'Some answers need changing before the model can be created.'));
    const p = preview(state, assets.lib);
    const recipe = recipeOf(state)!;
    const contents = el('ol', { class: 'wcontents' });
    for (const s of p.sections) {
      const sheets = el('ol');
      for (const sh of s.sheets) sheets.append(el('li', {}, sh));
      contents.append(el('li', {}, s.title, sheets));
    }
    const dl = el('dl');
    const row = (k: string, v: string) => dl.append(el('dt', {}, k), el('dd', {}, v));
    row('Entity', `${state.entityName} (${THEMES[state.brand].name} theme)`);
    row('Title', state.title);
    row('Starts from', recipe.label);
    row('Timeline', p.timeline);
    row('Model name shows', [state.showErrors && 'errors', state.showAlerts && 'alerts'].filter(Boolean).join(' and ') || 'the title only');
    return el('div', { class: 'wreview' }, dl, el('strong', {}, 'Contents'), contents);
  }

  async function create(button: HTMLButtonElement) {
    button.disabled = true;
    message = 'Building the workbook...';
    render();
    try {
      const bytes = buildFile(state, assets.lib, assets.logos);
      const name = fileName(state);
      if (download) URL.revokeObjectURL(download.url);
      download = { url: URL.createObjectURL(new Blob([bytes as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })), name };
      const canOpen = typeof Excel !== 'undefined' && typeof Office !== 'undefined'
        && !!Office?.context?.requirements?.isSetSupported('ExcelApi', '1.8');
      if (canOpen) {
        await Excel!.createWorkbook(toBase64(bytes));
        message = `${name} is open in a new window. Save it where the team keeps models.`;
      } else {
        message = `${name} is ready (${Math.round(bytes.length / 1024)} KB). Excel is not running this pane, so download it instead.`;
      }
    } catch (e) {
      message = `The model could not be created: ${(e as Error).message}`;
    }
    render();
  }

  function render() {
    host.innerHTML = '';
    const steps = el('ol', { class: 'wsteps' });
    STEPS.forEach((s, i) => steps.append(el('li', { class: i === state.step ? 'on' : i < state.step ? 'done' : '' }, `${i + 1}. ${s}`)));
    const body = [entityStep, modelStep, timelineStep, displayStep, reviewStep][state.step]();
    const errs = el('ul', { class: 'werr', role: 'alert' });
    const back = el('button', { type: 'button' }, 'Back');
    back.disabled = state.step === 0;
    back.addEventListener('click', () => { state.step -= 1; message = ''; render(); });
    const last = state.step === STEPS.length - 1;
    const next = el('button', { type: 'button', class: 'primary' }, last ? 'Create model' : 'Next');
    if (last && problems(state).length) next.disabled = true;
    next.addEventListener('click', () => {
      const found = problems(state);
      errs.innerHTML = '';
      if (found.length) {
        for (const f of found) errs.append(el('li', {}, f));
        return;
      }
      if (last) void create(next);
      else { state.step += 1; render(); }
    });
    const status = el('p', { class: 'quiet', role: 'status' }, message);
    if (download) status.append(' ', el('a', { href: download.url, download: download.name, class: 'button' }, 'Download a copy'));
    host.append(el('div', { class: 'wiz' }, steps, body, errs, el('div', { class: 'wactions' }, back, next), status));
  }

  render();
  return { state, rerender: render };
}
