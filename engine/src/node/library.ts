// Node only: read a library folder (areas.yaml and one YAML file per module, in file name order).

import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { parse } from 'yaml';
import { Library, type LibraryBundle, type ModuleDef } from '../library.ts';

export function readLibraryBundle(path: string): LibraryBundle {
  const spec = parse(readFileSync(join(path, 'areas.yaml'), 'utf8')) as Omit<LibraryBundle, 'modules'>;
  const modules = readdirSync(path).filter(n => n.endsWith('.yaml') && n !== 'areas.yaml').sort()
    .map(f => parse(readFileSync(join(path, f), 'utf8')) as ModuleDef);
  const b: LibraryBundle = { areas: spec.areas, sections: spec.sections || [], modules };
  if (spec.id) b.id = spec.id;
  if (spec.history) b.history = spec.history;
  if (spec.scenarios) b.scenarios = spec.scenarios;
  if (spec.lists) b.lists = spec.lists;
  return b;
}

export function loadLibrary(path: string): Library {
  return Library.fromBundle(readLibraryBundle(path));
}
