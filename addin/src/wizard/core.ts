// New model wizard: the choices, their checks, and the workbook they build. DOM free, so the
// Node tests run it; view.ts draws it in the task pane.
//
// Steps follow the spec (Look and wiring, New model wizard): entity, model, timeline, display,
// review. Scenarios join when the Scenarios module is in the library.

import {
  assemble, buildWorkbook, Library, Model, THEMES, type Brand, type Layout, type Logo, type ModelInfo, type Recipe as ModelRecipe,
} from '../../../engine/src/index.ts';

/** The libraries the add-in carries, by id ("hfg" the Phase 1 library, "demo" the Phase 0 proof's). */
export type Libraries = Record<string, Library>;

/** The library a recipe or model needs; a single library stands for every id (the tests pass one). */
export function libFor(libs: Library | Libraries, id: string): Library {
  if (libs instanceof Library) return libs;
  const lib = libs[id];
  if (!lib) throw new Error(`The add-in has no library '${id}'.`);
  return lib;
}

export const STEPS = ['Entity', 'Model', 'Timeline', 'Display', 'Review'] as const;
export const BRANDS: Brand[] = ['HF', 'HCP', 'HCL', 'KM', 'TWK'];
export const DENOMINATIONS = ['$', '$000', '$m'] as const;
export const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
  'November', 'December'];

/**
 * What a model starts from. A recipe lists modules to insert, or is a whole model with its data
 * (history and drivers), which brings its own timeline. Catalogue recipes join as the library grows.
 */
export interface Recipe {
  id: string;
  label: string;
  note: string;
  library: string;
  modules: [string, Record<string, unknown>][];
  model?: ModelRecipe;
}

export const RECIPES: Recipe[] = [
  { id: 'blank', label: 'Blank model', note: 'The frame only: contents, settings, lookups and the timeline. Insert modules afterwards.',
    library: 'hfg', modules: [] },
  {
    id: 'demo', label: 'Assembly demo (fictional data)', library: 'demo',
    note: 'The assembly demo: two revenue lines, a cost line, debtors, a debt facility, statements, checks and an income summary.',
    modules: [
      ['demo.statements', {}], ['demo.checks', {}],
      ['demo.revenue_line', { base: 100, growth: 0.01 }], ['demo.revenue_line', { base: 60, growth: 0.03 }],
      ['demo.cost_line', { amount: 90, inflation: 0.002 }], ['demo.debtors', {}],
      ['demo.facility', { amount: 1000, rate: 0.06, instalment: 50 }], ['demo.dashboard', { first: 1 }],
    ],
  },
];

/** Recipes the build bundles (whole models with their data) join the list. */
export function addRecipes(list: Recipe[]): void {
  for (const r of list) if (!RECIPES.some(x => x.id === r.id)) RECIPES.push(r);
}

export const recipeOf = (s: { recipe: string }) => RECIPES.find(r => r.id === s.recipe);

export interface WizardState {
  step: number;
  brand: Brand;
  entityName: string;
  title: string;
  recipe: string;
  preparedBy: string;
  notes: string;
  /** "yyyy-mm" */
  start: string;
  fyEndMonth: number;
  months: number;
  lastActual: number;
  denomination: (typeof DENOMINATIONS)[number];
  /** The budget's first month as a period (0 for the month after the actuals) and its length. */
  budgetFirst: number;
  budgetMonths: number;
  showErrors: boolean;
  showAlerts: boolean;
}

/** The financial year a date falls in starts the month after the year end month. */
export function yearStart(today: Date, fyEndMonth: number): string {
  let y = today.getFullYear();
  const startMonth = (fyEndMonth % 12) + 1;
  if (today.getMonth() + 1 < startMonth) y -= 1;
  return `${y}-${String(startMonth).padStart(2, '0')}`;
}

export function initialState(today: Date): WizardState {
  return {
    step: 0, brand: 'HF', entityName: THEMES.HF.name, title: '', recipe: 'blank', preparedBy: 'Prepared by Group Finance',
    notes: '', start: yearStart(today, 3), fyEndMonth: 3, months: 12, lastActual: 0, denomination: '$',
    budgetFirst: 0, budgetMonths: 12, showErrors: true, showAlerts: true,
  };
}

/** What is wrong with the answers so far, for one step (or every step before Review). */
export function problems(s: WizardState, step = s.step): string[] {
  const out: string[] = [];
  const all = step >= STEPS.length - 1;
  if (all || step === 0) {
    if (!BRANDS.includes(s.brand)) out.push('Choose the entity\'s brand.');
    if (!s.entityName.trim()) out.push('Type the entity\'s name.');
    else if (s.entityName.length > 80) out.push('Keep the entity\'s name under 80 characters.');
  }
  if (all || step === 1) {
    if (!s.title.trim()) out.push('Give the model a title.');
    else if (s.title.length > 120) out.push('Keep the title under 120 characters.');
    if (s.preparedBy.length > 120) out.push('Keep the Prepared by line under 120 characters.');
    if (!RECIPES.some(r => r.id === s.recipe)) out.push('Choose what the model starts from.');
  }
  if ((all || step === 2) && !recipeOf(s)?.model) {
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(s.start)) out.push('Choose the first month of the model.');
    if (!Number.isInteger(s.fyEndMonth) || s.fyEndMonth < 1 || s.fyEndMonth > 12) out.push('Choose the month the financial year ends.');
    if (!Number.isInteger(s.months) || s.months < 1 || s.months > 600) out.push('The model runs for 1 to 600 months.');
    else if (!Number.isInteger(s.lastActual) || s.lastActual < 0 || s.lastActual > s.months) {
      out.push(`The last month of actuals is a period from 0 (none) to ${s.months}.`);
    }
    if (!DENOMINATIONS.includes(s.denomination)) out.push('Choose the denomination.');
    const b = budgetOf(s);
    if (Number.isInteger(s.months) && (b.first < 1 || b.first > s.months)) out.push(`The budget starts in a period from 1 to ${s.months}.`);
    else if (!Number.isInteger(s.budgetMonths) || s.budgetMonths < 1 || (s.budgetFirst && b.first + s.budgetMonths - 1 > s.months)) {
      out.push('The budget has to end inside the timeline.');
    }
  }
  return out;
}

/** The budget's first period and length: by default the month after the actuals. */
export function budgetOf(s: WizardState): { first: number; months: number } {
  const first = s.budgetFirst || Math.min(s.lastActual + 1, s.months);
  return { first, months: Math.min(s.budgetMonths, Math.max(1, s.months - first + 1)) };
}

export function toInfo(s: WizardState): ModelInfo {
  return {
    title: s.title.trim(), entity: { name: s.entityName.trim(), brand: s.brand }, preparedBy: s.preparedBy.trim(),
    notes: s.notes.split('\n').map(n => n.trim()).filter(Boolean),
    timeline: { start: s.start, fyEndMonth: s.fyEndMonth, lastActual: s.lastActual, denomination: s.denomination, budget: budgetOf(s) },
    display: { errors: s.showErrors, alerts: s.showAlerts },
  };
}

export function toModel(s: WizardState, libs: Library | Libraries): Model {
  const recipe = recipeOf(s)!;
  const lib = libFor(libs, recipe.library);
  if (recipe.model) {   // a whole model with its data keeps its own timeline; the answers name and brand it
    const m = Model.fromRecipe(lib, recipe.model);
    const own = toInfo(s);
    m.info = { ...m.info!, title: own.title, entity: own.entity, preparedBy: own.preparedBy,
      notes: own.notes.length ? own.notes : m.info!.notes, display: own.display };
    return m;
  }
  const m = new Model(lib, s.months);
  m.info = toInfo(s);
  for (const [id, settings] of recipe.modules) if (lib.modules.has(id)) m.insert(id, settings);
  return m;
}

/** The month a period falls in, as "April 2026". */
export function periodMonth(start: string, period: number): string {
  const [y, m] = start.split('-').map(Number);
  const k = m - 1 + period - 1;
  return `${MONTHS[((k % 12) + 12) % 12]} ${y + Math.floor(k / 12)}`;
}

export interface Preview {
  sections: { title: string; sheets: string[] }[];
  /** Plain words about the timeline. */
  timeline: string;
  layout: Layout;
}

/** The contents the model will have, and its timeline in words. */
export function preview(s: WizardState, libs: Library | Libraries): Preview {
  const model = toModel(s, libs);
  const layout = assemble(model);
  const sections: Preview['sections'] = [];
  for (const [sheet] of layout.sheets.slice(1)) {
    if (layout.kindOf(sheet) === 'cover') sections.push({ title: layout.titles[sheet], sheets: [] });
    else if (sections.length) sections[sections.length - 1].sheets.push(sheet);
    else sections.push({ title: 'Model', sheets: [sheet] });
  }
  if (recipeOf(s)?.model) s = { ...s, ...timelineOf(model) };
  const end = periodMonth(s.start, s.months);
  const actual = s.lastActual ? `actuals to ${periodMonth(s.start, s.lastActual)}` : 'no actuals yet';
  const b = budgetOf(s);
  const budget = `budget ${periodMonth(s.start, b.first)} to ${periodMonth(s.start, b.first + b.months - 1)}`;
  const timeline = `${s.months} months, ${periodMonth(s.start, 1)} to ${end}; financial year ends in ${MONTHS[s.fyEndMonth - 1]}; ${actual}; ${budget}; in ${s.denomination}.`;
  return { sections, timeline, layout };
}

/** A model's timeline as the wizard's answers. */
export function timelineOf(m: Model): Pick<WizardState, 'start' | 'fyEndMonth' | 'months' | 'lastActual' | 'denomination' | 'budgetFirst' | 'budgetMonths'> {
  const t = m.info!.timeline;
  return { start: t.start, fyEndMonth: t.fyEndMonth, months: m.periods, lastActual: t.lastActual, denomination: t.denomination,
    budgetFirst: t.budget?.first ?? 0, budgetMonths: t.budget?.months ?? 12 };
}

/** A file name from the title: letters, digits, spaces and dashes. */
export function fileName(s: WizardState): string {
  const base = s.title.trim().replace(/[^A-Za-z0-9 \-]+/g, '').replace(/\s+/g, ' ').trim() || 'HFG model';
  return `${base}.xlsx`;
}

export function buildFile(s: WizardState, lib: Library | Libraries, logos: Partial<Record<Brand, Logo>>, created = new Date()): Uint8Array {
  const errors = problems(s, STEPS.length - 1);
  if (errors.length) throw new Error(errors.join(' '));
  const model = toModel(s, lib);
  return buildWorkbook(assemble(model), model, { logo: logos[s.brand], created });
}

/** Base64 for Excel.createWorkbook, in chunks so large files do not overflow the call stack. */
export function toBase64(bytes: Uint8Array): string {
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}
