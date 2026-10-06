// Node-only helpers: the entity logos shipped with the engine. The task pane loads the same files
// from the add-in's assets.

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { Brand } from '../model.ts';
import { THEMES } from '../theme.ts';
import type { Logo } from '../xlsx/package.ts';

export const ASSETS = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'assets');

/** Width and height from a PNG's header. */
export function pngSize(png: Uint8Array): { width: number; height: number } {
  const sig = [0x89, 0x50, 0x4e, 0x47];
  if (!sig.every((b, i) => png[i] === b)) throw new Error('not a PNG file');
  const dv = new DataView(png.buffer, png.byteOffset, png.byteLength);
  return { width: dv.getUint32(16), height: dv.getUint32(20) };
}

export function loadLogo(brand: Brand): Logo {
  const png = new Uint8Array(readFileSync(join(ASSETS, 'logos', THEMES[brand].logo.file)));
  return { png, ...pngSize(png) };
}
