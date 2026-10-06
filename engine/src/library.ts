// The module library: area order, sections and module definitions read from YAML files.

import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { parse } from 'yaml';
import { AssemblyError } from './frame.ts';

export interface SettingDef {
  key: string;
  label: string;
  unit?: string;
  default?: unknown;
  display?: boolean;
  group?: { item: string; formula?: string };
}

export interface RowDef {
  key?: string;
  label?: string;
  unit?: string;
  style?: string;
  name?: string;
  first?: string;
  formula?: string;
  total?: string;
  span?: number;
  check?: string;
  section?: string;
  collect?: string;
  headline?: HeadlineDef | HeadlineDef[];
}

export interface HeadlineDef {
  label: string;
  measure?: string;
}

export interface ChartDef {
  key: string;
  title: string;
  categories: string;
  series: { each?: string; row?: string; line?: boolean }[];
}

export interface ModuleDef {
  id: string;
  title: string;
  area: string;
  code?: string;
  as_category?: boolean;
  mirror?: string;
  framework?: string;
  settings?: SettingDef[];
  inputs?: { link: string; mode?: string; required?: boolean }[];
  rows?: RowDef[];
  outputs?: { link: string; row: string }[];
  charts?: ChartDef[];
}

export interface SectionDef {
  title: string;
  cover?: string | null;
  note?: string;
  areas: string[];
}

export type ModuleKind = 'single' | 'category' | 'mirror';

export class Library {
  areas: string[];
  modules: Map<string, ModuleDef>;
  sections: SectionDef[];

  constructor(areas: string[], modules: Map<string, ModuleDef>, sections: SectionDef[] = []) {
    this.areas = areas;
    this.modules = modules;
    this.sections = sections;
  }

  /** Read areas.yaml and every other YAML file in a folder, in file name order. */
  static load(path: string): Library {
    const spec = parse(readFileSync(join(path, 'areas.yaml'), 'utf8')) as { areas: string[]; sections?: SectionDef[] };
    const areas = spec.areas;
    const sections = spec.sections || [];
    const placed = sections.flatMap(s => s.areas);
    if (sections.length && [...placed].sort().join('\u0000') !== [...areas].sort().join('\u0000')) {
      throw new AssemblyError('areas.yaml: every area must sit in exactly one section');
    }
    const modules = new Map<string, ModuleDef>();
    for (const f of readdirSync(path).filter(n => n.endsWith('.yaml')).sort()) {
      if (f === 'areas.yaml') continue;
      const d = parse(readFileSync(join(path, f), 'utf8')) as ModuleDef;
      if (!areas.includes(d.area)) throw new AssemblyError(`${f}: area '${d.area}' is not in areas.yaml`);
      modules.set(d.id, d);
    }
    return new Library(areas, modules, sections);
  }

  module(id: string): ModuleDef {
    const m = this.modules.get(id);
    if (!m) throw new AssemblyError(`no module '${id}' in the library`);
    return m;
  }

  kind(id: string): ModuleKind {
    const m = this.module(id);
    if (m.mirror) return 'mirror';
    return m.as_category ? 'category' : 'single';
  }
}
