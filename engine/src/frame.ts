// Frame constants and small helpers shared by the engine.
//
// These follow the Phase 0 assembly proof (prototypes/assembly) so the engine can be checked
// against it cell for cell. The Look and wiring standard (spec, 7 October 2026) replaces the
// frame layer next; the resolution, rendering and planning above it stay.

/** Label columns by indent: B, C, D. */
export const LABEL_COLS: Record<number, number> = { 0: 2, 1: 3, 2: 4 };
export const UNIT_COL = 8;
export const TOTAL_COL = 9;
export const FIRST_PERIOD_COL = 10;
export const PERIOD_ROW = 5;
export const FIRST_ROW = 7;
export const CONTENTS = 'Contents';
export const META_NS = 'urn:hfg:model-metadata:v0';

/**
 * Where a sheet's content starts and which header the writers put above it. The proof frame is
 * the Phase 0 assembly proof's (content from row 7), kept so the golden tests compare like with
 * like. The standard frame is HFG's (docs/frame-standard.md and the Look and wiring standard):
 * rows 1 to 3 header, the timeline block in rows 5 to 15 on timeline sheets, content from row 17
 * there and from row 5 on every other sheet (Settings, the contents, covers and lists).
 */
export interface Frame {
  id: 'proof' | 'standard';
  firstRow(kind: string): number;
}

export const PROOF_FRAME: Frame = { id: 'proof', firstRow: () => FIRST_ROW };

/** Standard frame rows. */
export const STD = {
  titleRow: 1,
  nameRow: 2,
  entityRow: 3,
  blockFirst: 5,
  blockLast: 15,
  freezeRow: 16,
  contentRow: 17,
  plainRow: 5,
} as const;

export const STANDARD_FRAME: Frame = {
  id: 'standard',
  firstRow: kind => (kind === 'timeline' ? STD.contentRow : STD.plainRow),
};

/** A module cannot be inserted or the links cannot be resolved. */
export class AssemblyError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'AssemblyError';
  }
}

/** A key's value when the key is present (even when it is null), otherwise the default, as Python's dict.get. */
export function pick<T extends object, K extends keyof T, D>(o: T, k: K, d: D): Exclude<T[K], undefined> | D {
  return Object.prototype.hasOwnProperty.call(o, k) ? (o[k] as Exclude<T[K], undefined>) : d;
}

/** 1 -> A, 27 -> AA. */
export function colLetter(c: number): string {
  let s = '';
  while (c) {
    const r = (c - 1) % 26;
    c = Math.floor((c - 1) / 26);
    s = String.fromCharCode(65 + r) + s;
  }
  return s;
}

/** "lending_rate" -> "LendingRate": each part capitalised, the rest lower case. */
export function namePart(key: string): string {
  return key.split('_').map(p => (p ? p[0].toUpperCase() + p.slice(1).toLowerCase() : '')).join('');
}

/** Every character that is not a letter, digit or underscore becomes an underscore. */
export function code(text: string): string {
  return text.replace(/[^\p{L}\p{N}_]/gu, '_');
}

/** "Closing cash" -> "ClosingCash". */
export function codeWords(label: string): string {
  return (label.match(/[A-Za-z0-9]+/g) || []).map(w => w.slice(0, 1).toUpperCase() + w.slice(1)).join('');
}

export function sheetLinkName(sheet: string): string {
  return `HL_Sheet_${code(sheet)}`;
}

export function groupName(item: string): string {
  return 'GA_' + namePart(item);
}
