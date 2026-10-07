// Insert module: the model in the open workbook (from its metadata part), the modules that can go
// in, and the change plan an insert makes. DOM free; view.ts draws it, ../live/apply.ts applies it.

import {
  assemble, Library, metadataXml, Model, planChange, type Layout, type ModelDict, type ModuleDef, type Plan,
} from '../../../engine/src/index.ts';
import { libFor, type Libraries } from '../wizard/core.ts';

export interface OpenModel {
  model: Model;
  layout: Layout;
}

const unescape = (s: string) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** The model a workbook's metadata part describes, with the library it was built from. */
export function readModel(xml: string, libs: Library | Libraries): OpenModel {
  const m = /<hfgModel\b[^>]*>([\s\S]*)<\/hfgModel>/.exec(xml);
  if (!m) throw new Error('This workbook has no HFG model metadata.');
  const payload = JSON.parse(unescape(m[1])) as { model: ModelDict };
  const model = Model.fromDict(libFor(libs, payload.model.library ?? 'demo'), payload.model);
  return { model, layout: assemble(model) };
}

export interface Choice {
  id: string;
  title: string;
  area: string;
  /** Why it cannot go in, when it cannot. */
  blocked: string | null;
  settings: NonNullable<ModuleDef['settings']>;
}

/** Every module in the library, by area order, and whether this model can take it. */
export function choices(lib: Library, model: Model): Choice[] {
  return [...lib.modules.values()]
    .sort((a, b) => lib.areas.indexOf(a.area) - lib.areas.indexOf(b.area) || a.title.localeCompare(b.title))
    .map(d => ({
      id: d.id, title: d.title, area: d.area, settings: d.settings ?? [],
      blocked: lib.kind(d.id) !== 'category' && model.instances.some(i => i.module === d.id)
        ? 'Already in the model; it goes in once.' : null,
    }));
}

export interface InsertPlan {
  next: Model;
  layout: Layout;
  plan: Plan;
  metadata: string;
  title: string;
}

export function planInsert(open: OpenModel, moduleId: string, settings: Record<string, unknown>): InsertPlan {
  const next = open.model.copy();
  const inst = next.insert(moduleId, settings);
  const layout = assemble(next);
  return { next, layout, plan: planChange(open.layout, layout, 'excel'), metadata: metadataXml(next, layout), title: next.title(inst) };
}
