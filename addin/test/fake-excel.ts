// A stand-in for the Office.js object model the live writer uses, working on the engine's
// workbook image (engine/test/image.ts): row inserts and deletes move rows, styles are set by
// name, rows group and ungroup one level at a time (ungroup fails where there is no group, as in
// Excel), and conditional formats and validations are read back in the image's form.

import { colLetter } from '../../engine/src/frame.ts';
import { cell, empty, shift, type Image } from '../../engine/test/image.ts';
import type { XContext } from '../src/live/apply.ts';

function parse(address: string): { r1: number; r2: number; c1: number; c2: number } {
  const a = address.replace(/\$/g, '');
  const col = (s: string) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  let m = /^(\d+):(\d+)$/.exec(a);
  if (m) return { r1: +m[1], r2: +m[2], c1: 1, c2: 16384 };
  m = /^([A-Z]+):([A-Z]+)$/.exec(a);
  if (m) return { r1: 1, r2: 1048576, c1: col(m[1]), c2: col(m[2]) };
  m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a);
  if (!m) throw new Error(`fake Excel cannot read the address ${address}`);
  return { r1: +m[2], r2: +(m[4] ?? m[2]), c1: col(m[1]), c2: col(m[3] ?? m[1]) };
}

export interface FakeLog {
  calls: string[];
  parts: string[];
}

export function fakeExcel(img: Image): { ctx: XContext; log: FakeLog } {
  const log: FakeLog = { calls: [], parts: [] };
  const pending: (() => void)[] = [];

  function range(sheetName: string, address: string): any {
    const { r1, r2, c1, c2 } = parse(address);
    const sh = () => img.sheets.get(sheetName)!;
    const rows = () => Array.from({ length: Math.min(r2, 2000) - r1 + 1 }, (_, i) => r1 + i);
    const cellsIn = () => {
      const out: [number, number][] = [];
      for (const r of rows()) {
        if (c2 - c1 > 200) for (const c of sh().cells.get(r)?.keys() ?? []) { if (c >= c1 && c <= c2) out.push([r, c]); }
        else for (let c = c1; c <= c2; c++) out.push([r, c]);
      }
      return out;
    };
    const one = () => cell(sh(), r1, c1);
    const r: any = {
      _sheet: sheetName, _row: r1, _col: c1, _toRow: r2,
      set formulas(v: unknown[][]) { one().v = v[0][0]; },
      set values(v: unknown[][]) {
        const x = v[0][0];
        one().v = typeof x === 'string' && x.startsWith("'") ? x.slice(1) : x;
      },
      get values() { return rows().map(rr => [sh().cells.get(rr)?.get(c1)?.v ?? null]); },
      set style(name: string) {
        for (const [rr, cc] of cellsIn()) {
          if (name === 'Normal') { const x = sh().cells.get(rr)?.get(cc); if (x) delete x.style; }
          else cell(sh(), rr, cc).style = name;
        }
      },
      set rowHidden(h: boolean) { for (const rr of rows()) { if (h) sh().hidden.add(rr); else sh().hidden.delete(rr); } },
      set hyperlink(h: { documentReference: string; screenTip: string }) { sh().links.set(`${r1},${c1}`, `${h.documentReference}|${h.screenTip}`); },
      format: {
        set rowHeight(h: number) { for (const rr of rows()) sh().heights.set(rr, h); },
        set columnWidth(_w: number) { /* not compared */ },
      },
      conditionalFormats: {
        clearAll() { for (const rr of rows()) sh().conds.delete(rr); },
        add(type: string) {
          const edges: Record<string, any> = {};
          const cf: any = { cellValue: { format: { font: {} }, rule: null },
            custom: { format: { font: {}, fill: {}, borders: { getItem: (e: string) => (edges[e] ??= {}) } }, rule: { formula: '' } } };
          pending.push(() => {
            const font = (type === 'CellValue' ? cf.cellValue : cf.custom).format.font;
            const fill = cf.custom.format.fill.color;
            const grey = type !== 'CellValue' && font.color === '#BFBFBF';
            const red = font.color === '#CB2840';
            // The Scenarios sheet's formats: a shaded column, an upright name, a dark marker, a count in bold.
            const kind = type === 'CellValue' ? (red ? 'notZero' : 'on')
              : grey ? (fill === undefined ? 'na' : 'inactive') : fill && fill !== '#FFFFFF' ? 'selected' : font.italic === false ? 'upright'
                : !red && font.color && !font.bold ? 'marker' : 'expression';
            const rule = type === 'CellValue' ? { kind } : { kind, formula: cf.custom.rule.formula.replace(/^=/, '') };
            if (kind === 'inactive' && (fill !== '#FFFFFF' || Object.keys(edges).length !== 4)) throw new Error('an inactive format is not grey on white');
            if (kind === 'expression' && (!red || !font.bold)) throw new Error('a check format is not bold red');
            if (kind === 'on' && (!font.bold || !/^#[0-9A-F]{6}$/.test(font.color))) throw new Error('a count format is not bold');
            if (kind === 'upright' && !font.bold) throw new Error('an upright name is not bold');
            if (type === 'CellValue' && (cf.cellValue.rule.operator !== 'NotEqualTo' || cf.cellValue.rule.formula1 !== '=0')) throw new Error('bad check rule');
            const list = sh().conds.get(r1) ?? [];
            list.push(`${address.replace(/\d+/g, '#')}|${JSON.stringify(rule)}`);
            sh().conds.set(r1, list);
          });
          return cf;
        },
      },
      dataValidation: {
        clear() { for (const rr of rows()) sh().valid.delete(rr); },
        set rule(rule: unknown) {
          pending.push(() => {
            const list = sh().valid.get(r1) ?? [];
            list.push(`${c1}|${JSON.stringify(rule)}|${r.dataValidation._message}`);
            sh().valid.set(r1, list);
          });
        },
        set errorAlert(a: { message: string }) { r.dataValidation._message = a.message; },
      },
      merge(across: boolean) {
        if (across || r1 !== r2) throw new Error('only one row is merged');
        const list = sh().merges.get(r1) ?? [];
        list.push(`${c1}-${c2}`);
        sh().merges.set(r1, list);
      },
      unmerge() { for (const rr of rows()) sh().merges.delete(rr); },
      insert(dir: string) { log.calls.push(`${sheetName} insert ${address} ${dir}`); shift(img, sheetName, r1, r2 - r1 + 1); },
      delete(dir: string) { log.calls.push(`${sheetName} delete ${address} ${dir}`); shift(img, sheetName, r1, -(r2 - r1 + 1)); },
      clear(what: string) {
        if (what === 'Formats' || what === 'All') for (const rr of rows()) sh().merges.delete(rr);
        for (const [rr, cc] of cellsIn()) {
          if (what === 'Contents') { const x = sh().cells.get(rr)?.get(cc); if (x) delete x.v; }
          if (what === 'Hyperlinks') sh().links.delete(`${rr},${cc}`);
        }
      },
      group(by: string) { if (by !== 'ByRows') throw new Error(by); for (const rr of rows()) sh().levels.set(rr, (sh().levels.get(rr) ?? 0) + 1); },
      ungroup(_by: string) {
        const grouped = rows().filter(rr => (sh().levels.get(rr) ?? 0) > 0);
        if (!grouped.length) throw new Error('InvalidOperation: no group to remove');
        for (const rr of grouped) { const l = sh().levels.get(rr)! - 1; if (l) sh().levels.set(rr, l); else sh().levels.delete(rr); }
      },
      load() { /* values are live */ },
    };
    return r;
  }

  function sheetApi(name: string): any {
    if (!img.sheets.has(name)) throw new Error(`ItemNotFound: ${name}`);
    return {
      name,
      set position(i: number) { img.order = img.order.filter(s => s !== name); img.order.splice(i, 0, name); },
      showGridlines: true,
      getRange: (a: string) => range(name, a),
      getCell: (r: number, c: number) => range(name, `${colLetter(c + 1)}${r + 1}`),
      delete() { img.sheets.delete(name); img.order = img.order.filter(s => s !== name); },
      freezePanes: { freezeAt() {}, freezeRows() {}, unfreeze() {} },
      charts: { add() { throw new Error('charts are not part of this test'); }, load() {}, items: [] },
    };
  }

  const ctx: XContext = {
    workbook: {
      worksheets: {
        getItem: sheetApi,
        add(name: string) { img.sheets.set(name, empty()); img.order.push(name); return sheetApi(name); },
      },
      names: {
        add(name: string, rng: any) {
          img.names.set(name, typeof rng === 'string' ? rng : `${rng._sheet}!${rng._row},${rng._col}${rng._toRow !== rng._row ? `:${rng._toRow}` : ''}`);
          return { name, delete() { img.names.delete(name); } };
        },
        getItem(name: string) {
          const at = img.names.get(name);
          if (!at) throw new Error(`ItemNotFound: ${name}`);
          const m = /^(.+)!(\d+),(\d+)(?::(\d+))?$/.exec(at)!;
          return { getRange: () => range(m[1], `${colLetter(Number(m[3]))}${m[2]}:${colLetter(Number(m[3]))}${m[4] ?? m[2]}`) };
        },
        load() {},
        get items() { return [...img.names.keys()].map(name => ({ name, delete() { img.names.delete(name); } })); },
      },
      customXmlParts: {
        getByNamespace() { return { load() {}, items: log.parts.map((_, i) => ({ delete() { log.parts.splice(i, 1); } })) }; },
        add(xml: string) { log.parts.push(xml); },
      },
    } as unknown as XContext['workbook'],
    async sync() { while (pending.length) pending.shift()!(); },
  };

  return { ctx, log };
}
