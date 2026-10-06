// The module library: area order, sections and module definitions. Kept in YAML in the repo and
// bundled as JSON for the add-in (LibraryBundle); src/node/library.ts reads the YAML folder.

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

/** The library as one JSON document: what the add-in ships. */
export interface LibraryBundle {
  areas: string[];
  sections?: SectionDef[];
  modules: ModuleDef[];
}

export class Library {
  areas: string[];
  modules: Map<string, ModuleDef>;
  sections: SectionDef[];

  constructor(areas: string[], modules: Map<string, ModuleDef>, sections: SectionDef[] = []) {
    this.areas = areas;
    this.modules = modules;
    this.sections = sections;
  }

  /** A library from its bundle: areas.yaml's content and every module definition in file name order. */
  static fromBundle(bundle: LibraryBundle): Library {
    const areas = bundle.areas;
    const sections = bundle.sections || [];
    const placed = sections.flatMap(s => s.areas);
    if (sections.length && [...placed].sort().join('\u0000') !== [...areas].sort().join('\u0000')) {
      throw new AssemblyError('areas.yaml: every area must sit in exactly one section');
    }
    const modules = new Map<string, ModuleDef>();
    for (const d of bundle.modules) {
      if (!areas.includes(d.area)) throw new AssemblyError(`${d.id}: area '${d.area}' is not in areas.yaml`);
      modules.set(d.id, d);
    }
    return new Library(areas, modules, sections);
  }

  /** The bundle this library came from, for the add-in. */
  toBundle(): LibraryBundle {
    return { areas: this.areas, sections: this.sections, modules: [...this.modules.values()] };
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
