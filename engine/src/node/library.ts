// Node only: read a library folder (areas.yaml and one YAML file per module, in file name order).

import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { parse } from 'yaml';
import { Library, type LibraryBundle, type ModuleDef, type SectionDef } from '../library.ts';

export function readLibraryBundle(path: string): LibraryBundle {
  const spec = parse(readFileSync(join(path, 'areas.yaml'), 'utf8')) as { areas: string[]; sections?: SectionDef[] };
  const modules = readdirSync(path).filter(n => n.endsWith('.yaml') && n !== 'areas.yaml').sort()
    .map(f => parse(readFileSync(join(path, f), 'utf8')) as ModuleDef);
  return { areas: spec.areas, sections: spec.sections || [], modules };
}

export function loadLibrary(path: string): Library {
  return Library.fromBundle(readLibraryBundle(path));
}
