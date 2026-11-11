// The Excel self-test's side outside Excel: put the test task pane into a workbook (and take it
// out again for the copy Kelvin looks at), read what a package-written workbook holds cell by
// cell, and compare a snapshot of Excel with it. Node only (tools/excel-selftest.ts and the tests).

import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate';
import type { Snapshot } from './run.ts';

const TASKPANES_REL = 'http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes';

/** The workbook with a task pane that opens the add-in `id` (registered on this PC) when the file opens. */
export function withTaskPane(bytes: Uint8Array, id: string): Uint8Array {
  const files = unzipSync(bytes) as Zippable;
  const text = (p: string) => strFromU8(files[p] as Uint8Array);
  files['xl/webextensions/webextension1.xml'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{7A1B5E3C-2F4D-4E6A-9B8C-1D2E3F4A5B6C}">'
    + `<we:reference id="${id}" version="1.0.0.0" store="developer" storeType="Registry"/>`
    + '<we:alternateReferences/><we:properties/><we:bindings/><we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
    + '</we:webextension>');
  files['xl/webextensions/taskpanes.xml'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">'
    + '<wetp:taskpane dockstate="right" visibility="1" width="350" row="1">'
    + '<wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdWe1"/>'
    + '</wetp:taskpane></wetp:taskpanes>');
  files['xl/webextensions/_rels/taskpanes.xml.rels'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rIdWe1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/>'
    + '</Relationships>');
  files['_rels/.rels'] = strToU8(text('_rels/.rels').replace('</Relationships>',
    `<Relationship Id="rIdTp1" Type="${TASKPANES_REL}" Target="xl/webextensions/taskpanes.xml"/></Relationships>`));
  files['[Content_Types].xml'] = strToU8(text('[Content_Types].xml').replace('</Types>',
    '<Override PartName="/xl/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>'
    + '<Override PartName="/xl/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/></Types>'));
  return zipSync(files, { level: 6 });
}

/** The workbook without any task pane or web extension parts (as Excel saved it, or as written above). */
export function withoutTaskPane(bytes: Uint8Array): Uint8Array {
  const files = unzipSync(bytes) as Zippable;
  for (const p of Object.keys(files)) if (p.startsWith('xl/webextensions/')) delete files[p];
  files['_rels/.rels'] = strToU8(strFromU8(files['_rels/.rels'] as Uint8Array)
    .replace(/<Relationship\b[^>]*webextensiontaskpanes[^>]*\/>/g, ''));
  files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml'] as Uint8Array)
    .replace(/<Override\b[^>]*PartName="\/xl\/webextensions\/[^"]*"[^>]*\/>/g, ''));
  return zipSync(files, { level: 6 });
}

const unxml = (s: string) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const colNum = (s: string) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const colName = (n: number) => { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };

/** A cell as a fresh build writes it: a formula (without "="), or a value. */
export type Cell = { f: string } | { v: string | number | boolean };

export interface Expected {
  order: string[];
  sheets: Map<string, Map<string, Cell>>;
  /** Workbook names and what they refer to (without "="). */
  names: Map<string, string>;
  /** Cells inside one-variable data tables, by sheet: Office.js reads them as TABLE() and cannot write them. */
  tables: Map<string, Set<string>>;
}

/** Every cell, name and data table a package-written workbook holds. */
export function expectedOf(bytes: Uint8Array): Expected {
  const files = unzipSync(bytes);
  const text = (p: string) => strFromU8(files[p]);
  const wb = text('xl/workbook.xml');
  const rels = new Map([...text('xl/_rels/workbook.xml.rels').matchAll(/<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"/g)].map(m => [m[1], m[2]]));
  const out: Expected = { order: [], sheets: new Map(), names: new Map(), tables: new Map() };
  for (const m of wb.matchAll(/<sheet\b[^>]*\bname="([^"]*)"[^>]*\br:id="([^"]+)"/g)) {
    const name = unxml(m[1]);
    out.order.push(name);
    const target = rels.get(m[2])!.replace(/^\/?(xl\/)?/, '');
    const xml = text(`xl/${target}`);
    const cells = new Map<string, Cell>();
    const tables = new Set<string>();
    for (const c of xml.matchAll(/<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const [, addr, attrs, body = ''] = c;
      const dt = /<f t="dataTable" ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/.exec(body);
      if (dt) {
        for (let r = Number(dt[2]); r <= Number(dt[4]); r++) for (let k = colNum(dt[1]); k <= colNum(dt[3]); k++) tables.add(`${colName(k)}${r}`);
        continue;
      }
      const f = /<f>([\s\S]*?)<\/f>/.exec(body);
      if (f) { cells.set(addr, { f: unxml(f[1]) }); continue; }
      const t = /\bt="(\w+)"/.exec(attrs)?.[1];
      if (t === 'inlineStr') { cells.set(addr, { v: unxml(/<t\b[^>]*>([\s\S]*?)<\/t>/.exec(body)?.[1] ?? '') }); continue; }
      const v = /<v>([\s\S]*?)<\/v>/.exec(body)?.[1];
      if (v === undefined) continue;
      cells.set(addr, { v: t === 'b' ? v === '1' : t === 's' || t === 'str' ? unxml(v) : Number(v) });
    }
    out.sheets.set(name, cells);
    if (tables.size) out.tables.set(name, tables);
  }
  for (const m of wb.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    const name = /\bname="([^"]+)"/.exec(m[1])![1];
    if (/\blocalSheetId=/.test(m[1]) || name.startsWith('_xlnm.')) continue;
    out.names.set(name, unxml(m[2]));
  }
  return out;
}

/** A formula as both sides can be compared: no "=", no future-function prefixes, no implicit-intersection "@". */
export function normFormula(f: string): string {
  return f.replace(/^=/, '').replace(/_xl(fn|ws|pm)\./g, '').replace(/(^|[(,+\-*/&=<>^:\s])@(?=[A-Za-z_$'])/g, '$1');
}

const sameValue = (a: unknown, b: unknown) => (typeof a === 'number' && typeof b === 'number'
  ? Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b)) : a === b);

export interface Mismatch {
  sheet: string;
  cell: string;
  want: string;
  got: string;
}

export interface Comparison {
  label: string;
  orderMatches: boolean;
  missingSheets: string[];
  extraSheets: string[];
  compared: number;
  mismatches: number;
  /** Drop-downs written live: the cell holds the item's text and its Sel_ name reads its position with MATCH, where a build holds the position. */
  liveDropdowns: number;
  bySheet: Record<string, number>;
  examples: Mismatch[];
  names: { missing: string[]; extra: string[]; different: { name: string; want: string; got: string }[] };
  errorCells: string[];
  checks: Snapshot['checks'];
}

const show = (c: Cell | undefined) => (!c ? '(empty)' : 'f' in c ? `=${c.f}` : JSON.stringify(c.v));

/** The cells Excel holds (from a snapshot) against a fresh build of the model it says it is. */
export function compare(snap: Snapshot, want: Expected, limit = 60): Comparison {
  const out: Comparison = {
    label: snap.label, orderMatches: JSON.stringify(snap.order) === JSON.stringify(want.order),
    missingSheets: want.order.filter(s => !snap.order.includes(s)), extraSheets: snap.order.filter(s => !want.order.includes(s)),
    compared: 0, mismatches: 0, liveDropdowns: 0, bySheet: {}, examples: [],
    names: { missing: [], extra: [], different: [] }, errorCells: snap.errorCells, checks: snap.checks,
  };
  // Names a live drop-down gave a MATCH of its cell, and the cells they read, by sheet and address
  const gotNames = new Map(Object.entries(snap.names).map(([k, v]) => [k, normFormula(String(v))]));
  const dropNames = new Set<string>();
  const dropCells = new Set<string>();
  for (const [n, f] of want.names) {
    const m = /^MATCH\((.+),(List_\w+),0\)$/.exec(gotNames.get(n) ?? '');
    if (!m || m[1] !== normFormula(f)) continue;
    const at = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)$/.exec(m[1]);
    if (!at) continue;
    dropNames.add(n);
    dropCells.add(`${(at[1] ?? at[2]).replace(/''/g, "'")}!${at[3]}${at[4]}`);
  }
  for (const sheet of want.order) {
    if (!snap.order.includes(sheet)) continue;
    const got = new Map<string, Cell>();
    const s = snap.sheets[sheet];
    if (s) {
      const top = /!\$?([A-Z]+)\$?(\d+)/.exec(s.address)!;
      const c0 = colNum(top[1]);
      const r0 = Number(top[2]);
      s.formulas.forEach((row, r) => row.forEach((x, c) => {
        if (x === '' || x === null || x === undefined) return;
        got.set(`${colName(c0 + c)}${r0 + r}`, typeof x === 'string' && x.startsWith('=') ? { f: x.slice(1) } : { v: x as string | number | boolean });
      }));
    }
    const exp = want.sheets.get(sheet)!;
    const tables = want.tables.get(sheet) ?? new Set<string>();
    for (const addr of new Set([...exp.keys(), ...got.keys()])) {
      if (tables.has(addr)) continue;
      const a = exp.get(addr);
      const b = got.get(addr);
      // an empty text cell and a blank cell are the same thing
      if ((!a || ('v' in a && a.v === '')) && (!b || ('v' in b && b.v === ''))) continue;
      out.compared++;
      const same = a && b && ('f' in a ? 'f' in b && normFormula(a.f) === normFormula(b.f) : 'v' in b && sameValue(a.v, b.v));
      if (same) continue;
      if (dropCells.has(`${sheet}!${addr}`) && a && 'v' in a && typeof a.v === 'number' && b && 'v' in b && typeof b.v === 'string') {
        out.liveDropdowns++;
        continue;
      }
      out.mismatches++;
      out.bySheet[sheet] = (out.bySheet[sheet] ?? 0) + 1;
      if (out.examples.length < limit) out.examples.push({ sheet, cell: addr, want: show(a), got: show(b) });
    }
  }
  for (const [n, f] of want.names) {
    if (!gotNames.has(n)) out.names.missing.push(n);
    else if (normFormula(f) !== gotNames.get(n) && !dropNames.has(n)) out.names.different.push({ name: n, want: f, got: gotNames.get(n)! });
  }
  // Excel adds hidden _xlfn. names of its own for some functions when it opens a file
  for (const n of gotNames.keys()) if (!want.names.has(n) && !/^_xl(nm|fn|ws)\./.test(n)) out.names.extra.push(n);
  return out;
}
