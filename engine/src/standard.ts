// The HFG standard frame (docs/frame-standard.md, Look and wiring): the Settings sheet the timeline
// comes from, the contents with the entity, model name line, notes and a grouped table of contents,
// section covers, cell hyperlinks to Go_ names with screen tips, and the header every sheet carries.
//
// HFG's own name prefixes (7 October 2026): Go_ navigation targets, Tl_ timeline settings, Opt_
// switches, Model_ the model's own lines, Chk_ check totals; GA_, Reg_ and KO_ as before.

import { FRAME_KINDS, keyOutputFormula } from './assurance.ts';
import { AssemblyError, code, colLetter, CONTENTS, FIRST_PERIOD_COL, pick, STANDARD_FRAME, STD, TOTAL_COL } from './frame.ts';
import { LRow, type CellLink, type Layout } from './layout.ts';
import type { SectionDef } from './library.ts';
import type { Model, ModelInfo } from './model.ts';
import { sheetPrefix, unoSeparators, type Dialect, type FrameCell } from './render.ts';
import type { Block } from './resolve.ts';

export const SETTINGS = 'Settings';
export const GO_CONTENTS = 'Go_Contents';
export const GO_CHECKS = 'Go_Checks';
export const goSheet = (sheet: string) => `Go_Sheet_${code(sheet)}`;
export const goBlock = (uid: string) => `Go_${code(uid)}`;

/** Navigation symbols, drawn in Segoe UI Symbol. */
export const SYMBOL = { home: '⌂', clear: '✓', failing: '✗', prev: '◀', next: '▶' } as const;

/** Days since 30 December 1899, as Excel counts dates. */
export function excelDate(y: number, m: number, d: number): number {
  return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
}

/** "2026-04" -> 1 April 2026 as an Excel date. */
export function monthStart(ym: string): number {
  const m = /^(\d{4})-(\d{2})$/.exec(ym);
  if (!m) throw new AssemblyError(`timeline start ${ym} is not a month (yyyy-mm)`);
  return excelDate(Number(m[1]), Number(m[2]), 1);
}

/** The timeline block on every timeline sheet: label and the formula on the Settings sheet, where it is worked out. */
export const BLOCK: { row: number; label: string; source: (c: number, col: string, prev: string | null) => string }[] = [
  { row: 5, label: 'Month ending', source: (_c, L) => `=${L}8` },
  { row: 6, label: 'Actual or forecast', source: (_c, L) => `=IF(${L}9<=Tl_Last_Actual,"Actual","Forecast")` },
  { row: 7, label: 'Period start', source: (_c, L) => `=EDATE(Tl_Start,${L}9-1)` },
  { row: 8, label: 'Period end', source: (_c, L) => `=EOMONTH(${L}7,0)` },
  { row: 9, label: 'Period', source: (_c, _L, P) => (P ? `=${P}9+1` : '=1') },
  { row: 10, label: 'Financial year', source: (_c, L) => `=YEAR(${L}8)+IF(MONTH(${L}8)>Tl_FY_End_Month,1,0)` },
  { row: 11, label: 'Month of the year', source: (_c, L) => `=MOD(MONTH(${L}8)-Tl_FY_End_Month-1,12)+1` },
  { row: 12, label: 'Quarter of the year', source: (_c, L) => `=INT((${L}11-1)/3)+1` },
  { row: 13, label: 'Half of the year', source: (_c, L) => `=INT((${L}11-1)/6)+1` },
  { row: 14, label: 'Actual month', source: (_c, L) => `=IF(${L}9<=Tl_Last_Actual,1,0)` },
  { row: 15, label: 'Forecast month number', source: (_c, L) => `=MAX(0,${L}9-Tl_Last_Actual)` },
];

function settingsRows(info: ModelInfo, periods: number): LRow[] {
  const set = (key: string, label: string, init: ConstructorParameters<typeof LRow>[3]) =>
    new LRow(`settings/${key}`, 'setting', label, { indent: 1, ...init });
  const textValid = { kind: 'text' as const, max: 120, message: 'Keep it under 120 characters.' };
  const onOff = { kind: 'list' as const, items: ['TRUE', 'FALSE'], message: 'Choose TRUE or FALSE.' };
  const end = (key: string) => new LRow(`settings/${key}/end`, 'blank', '', { space: 9, level: 0 });
  return [
    new LRow('settings/model/heading', 'heading', 'Model'),
    set('model/title', 'Model title', { value: info.title, name: 'Model_Title', role: 'in.text', valid: textValid }),
    new LRow('settings/model/entity', 'fixed', 'Entity', { indent: 1, name: 'Model_Entity', role: 'text',
      cells: { [TOTAL_COL]: info.entity.name } }),
    set('model/prepared', 'Prepared by line', { value: info.preparedBy, name: 'Model_Prepared_By', role: 'in.text', valid: textValid }),
    end('model'),
    new LRow('settings/time/heading', 'heading', 'Timeline'),
    set('time/start', 'First month of the model', { value: monthStart(info.timeline.start), name: 'Tl_Start', role: 'in.date',
      valid: { kind: 'date', message: 'Type a date, such as 1 April 2026.' } }),
    set('time/fy', 'Month the financial year ends (1 to 12)', { value: info.timeline.fyEndMonth, name: 'Tl_FY_End_Month',
      role: 'in.count', valid: { kind: 'whole', min: 1, max: 12, message: 'Type a month number from 1 to 12; 3 is March.' } }),
    set('time/last', 'Last month of actuals (period number, 0 for none)', { value: info.timeline.lastActual,
      name: 'Tl_Last_Actual', role: 'in.count',
      valid: { kind: 'whole', min: 0, max: 'Tl_Term', message: 'Type a period number from 0 to the months in the model.' } }),
    set('time/denom', 'Denomination', { value: info.timeline.denomination, name: 'Tl_Denom', role: 'in.text',
      valid: { kind: 'list', items: ['$', '$000', '$m'], message: 'Choose $, $000 or $m.' } }),
    new LRow('settings/time/term', 'fixed', 'Months in the model', { indent: 1, unit: 'months', name: 'Tl_Term', role: 'int',
      cells: { [TOTAL_COL]: periods } }),
    end('time'),
    new LRow('settings/display/heading', 'heading', 'Display'),
    set('display/errors', 'Show the error count in the model name', { value: info.display.errors, name: 'Opt_Show_Errors',
      role: 'in.switch', valid: onOff }),
    set('display/alerts', 'Show the alert count in the model name', { value: info.display.alerts, name: 'Opt_Show_Alerts',
      role: 'in.switch', valid: onOff }),
    end('display'),
  ];
}

const HYPERLINK = /^=HYPERLINK\("#([^"]+)",(.*)\)$/;

/** `=HYPERLINK("#X",text)` written by shared code becomes the text with a cell hyperlink to X. */
function liftHyperlinks(row: LRow): void {
  for (const [c, v] of Object.entries(row.cells)) {
    if (typeof v !== 'string') continue;
    const m = HYPERLINK.exec(v);
    if (!m) continue;
    const literal = /^"([^"]*)"$/.exec(m[2]);
    row.cells[Number(c)] = literal ? literal[1] : `=${m[2]}`;
    row.links = { ...(row.links || {}), [Number(c)]: { to: m[1], tip: literal ? `Go to ${literal[1]}` : `Go to ${m[1]}` } };
  }
}

export function navigateStandard(layout: Layout, model: Model, blocks: Block[], sheets: Map<string, LRow[]>): void {
  const info = model.info!;
  const lib = model.lib;
  layout.frame = STANDARD_FRAME;
  sheets.set(SETTINGS, settingsRows(info, model.periods));
  for (const [, rows] of sheets) rows.forEach(liftHyperlinks);

  const sections: SectionDef[] = lib.sections.length ? lib.sections
    : [{ title: 'Model', cover: null, note: '', areas: lib.areas }];
  const home = sections.find(s => s.areas.includes('Checks')) ?? sections[sections.length - 1];
  const present = sections
    .map(sec => [sec, [...(sec === home ? [SETTINGS] : []), ...sec.areas.filter(a => sheets.has(a))]] as [SectionDef, string[]])
    .filter(([, areas]) => areas.length);
  const order: string[] = [CONTENTS];
  for (const [sec, areas] of present) {
    if (sec.cover) order.push(sec.cover);
    order.push(...areas);
  }
  layout.kinds = { [CONTENTS]: 'contents' };
  layout.titles = { [CONTENTS]: info.entity.name };
  for (const [sec, areas] of present) {
    if (sec.cover) {
      layout.kinds[sec.cover] = 'cover';
      layout.titles[sec.cover] = sec.title;
    }
    for (const a of areas) {
      layout.kinds[a] = a === SETTINGS ? 'settings' : Object.hasOwn(FRAME_KINDS, a) ? FRAME_KINDS[a] : 'timeline';
      layout.titles[a] = a;
    }
  }

  // The blocks each sheet holds, for the contents.
  const headings = new Map<string, [string, string][]>();
  const add = (sheet: string, uid: string, title: string) => {
    const list = headings.get(sheet);
    if (list) list.push([uid, title]);
    else headings.set(sheet, [[uid, title]]);
  };
  const seen = new Set<string>();
  for (const b of blocks) {
    if (seen.has(b.inst.uid)) continue;
    seen.add(b.inst.uid);
    add(b.mod.area, b.inst.uid, model.title(b.inst));
  }
  for (const r of sheets.get(SETTINGS)!) if (r.kind === 'heading') add(SETTINGS, r.id.replace(/\/heading$/, ''), r.label);

  const names = layout.names;
  const nameAt = (nm: string, target: string, col: number) => {
    names.set(nm, target);
    layout.nameCols[nm] = col;
  };
  nameAt(GO_CONTENTS, `@${CONTENTS}`, 2);
  if (order.includes('Checks')) nameAt(GO_CHECKS, '@Checks', 2);
  for (const sh of order) nameAt(goSheet(sh), `@${sh}`, 2);
  nameAt('Model_Name', `@${CONTENTS}/2`, 2);
  for (const r of sheets.get(SETTINGS)!) if (r.name) names.set(r.name, r.id);

  const link = (to: string, tip: string): CellLink => ({ to, tip });
  const contents: LRow[] = [new LRow('contents/notes/heading', 'heading', 'Notes')];
  const notes = info.notes.length ? info.notes : ['No notes yet.'];
  notes.forEach((n, i) => contents.push(new LRow(`contents/note/${i + 1}`, 'toc', `note ${i + 1}`, {
    role: 'note', cells: { 2: info.notes.length ? `${i + 1}.` : '', 3: n } })));
  contents.push(new LRow('contents/notes/end', 'blank', '', { space: 9, level: 0 }));
  contents.push(new LRow('contents/title', 'heading', 'Table of contents'));
  contents.push(new LRow('contents/gap', 'blank', '', { space: 6, level: 0 }));
  present.forEach(([sec, areas], k0) => {
    const n = k0 + 1;
    const cover = sec.cover;
    if (cover) {
      contents.push(new LRow(`contents/section/${cover}`, 'toc', `section ${n} ${sec.title}`, {
        role: 'toc1', level: 0, cells: { 2: `${n}.`, 3: `=«S|${cover}»` },
        links: { 2: link(goSheet(cover), `Go to ${sec.title}`), 3: link(goSheet(cover), `Go to ${sec.title}`) } }));
    }
    areas.forEach((a, k) => {
      const letter = String.fromCharCode(97 + k);
      contents.push(new LRow(`contents/sheet/${a}`, 'toc', `${letter}. ${a}`, {
        role: 'toc2', level: 1, cells: { 3: `${letter}.`, 4: `=«S|${a}»` },
        links: { 3: link(goSheet(a), `Go to ${a}`), 4: link(goSheet(a), `Go to ${a}`) } }));
      for (const [uid, title] of headings.get(a) || []) {
        const nm = goBlock(uid);
        nameAt(nm, `${uid}/heading`, 2);
        contents.push(new LRow(`contents/module/${uid}`, 'toc', `- ${title}`, {
          role: 'toc3', level: 2, cells: { 4: '-', 5: `=«B|${uid}/heading»` },
          links: { 4: link(nm, `Go to ${title}`), 5: link(nm, `Go to ${title}`) } }));
      }
    });
    contents.push(new LRow(`contents/section/${sec.cover ?? n}/end`, 'blank', '', { space: 6, level: 0 }));
  });
  if (layout.hasChecks()) {
    contents.push(new LRow('contents/checks', 'heading', 'Checks'));
    contents.push(new LRow('contents/errors', 'toc', 'error checks', { style: 'check', role: 'link', cells: {
      3: 'Error checks failing', [TOTAL_COL]: '=Chk_Errors' }, links: { 3: link(GO_CHECKS, 'Go to the checks') } }));
    if (names.has('Chk_Alerts')) {
      contents.push(new LRow('contents/alerts', 'toc', 'alerts', { style: 'check', role: 'link', cells: {
        3: 'Alerts raised', [TOTAL_COL]: '=Chk_Alerts' }, links: { 3: link(GO_CHECKS, 'Go to the checks') } }));
    }
    contents.push(new LRow('contents/checks/end', 'blank', '', { space: 9, level: 0 }));
  }
  if (model.assured() && layout.headlines.length) {
    contents.push(new LRow('contents/ko', 'heading', 'Key outputs'));
    for (const h of layout.headlines) {
      const rid = `contents/ko/${h.name}`;
      contents.push(new LRow(rid, 'toc', h.label, { unit: h.unit, role: 'text', cells: { 3: h.label, [TOTAL_COL]: keyOutputFormula(h) } }));
      names.set(h.name, rid);
    }
  }

  const covers = present.map(([sec]) => sec.cover);
  const out: [string, LRow[]][] = [[CONTENTS, contents]];
  for (let i = 1; i < order.length; i++) {
    const sh = order[i];
    if (layout.kinds[sh] !== 'cover') {
      out.push([sh, sheets.get(sh)!]);
      continue;
    }
    const n = covers.indexOf(sh) + 1;
    const sec = present[n - 1][0];
    const prev = order[i - 1];
    const nxt = order[i + 1];
    if (nxt === undefined) throw new AssemblyError(`${sh}: a section cover cannot be the last sheet`);
    out.push([sh, [
      new LRow(`cover/${sh}/number`, 'toc', 'section number', { role: 'sectionNo', cells: { 2: `Section ${n}.` } }),
      new LRow(`cover/${sh}/gap`, 'blank', '', { space: 6 }),
      new LRow(`cover/${sh}/home`, 'toc', 'link to the contents', { role: 'linkU', cells: { 2: 'Go to contents' },
        links: { 2: link(GO_CONTENTS, 'Go to the table of contents') } }),
      new LRow(`cover/${sh}/prev`, 'toc', 'link to the previous sheet', { role: 'link',
        cells: { 2: `="${SYMBOL.prev} "&«S|${prev}»` }, links: { 2: link(goSheet(prev), prev === CONTENTS ? 'Go to the table of contents' : `Go to ${prev}`) } }),
      new LRow(`cover/${sh}/next`, 'toc', 'link to the next sheet', { role: 'link',
        cells: { 2: `=«S|${nxt}»&" ${SYMBOL.next}"` }, links: { 2: link(goSheet(nxt), `Go to ${nxt}`) } }),
      new LRow(`cover/${sh}/gap2`, 'blank', '', { space: 9 }),
      new LRow(`cover/${sh}/notes_title`, 'toc', 'notes heading', { role: 'h3', cells: { 2: 'Section notes' } }),
      new LRow(`cover/${sh}/notes`, 'toc', 'notes', { role: 'note', cells: { 2: pick(sec, 'note', '') } }),
    ]]);
  }
  layout.sheets = out;
}

/** The model name line: the title, then the error and alert counts when they are not clear and their switch is on. */
export function modelNameFormula(layout: Layout): string {
  let f = '=Model_Title';
  if (layout.names.has('Chk_Errors')) {
    f += '&IF(AND(Opt_Show_Errors,Chk_Errors>0)," ("&Chk_Errors&IF(Chk_Errors=1," error)"," errors)"),"")';
  }
  if (layout.names.has('Chk_Alerts')) {
    f += '&IF(AND(Opt_Show_Alerts,Chk_Alerts>0)," ("&Chk_Alerts&IF(Chk_Alerts=1," alert)"," alerts)"),"")';
  }
  return f;
}

/** Header cells of a sheet in the standard frame, in the order written. */
export function standardFrameCells(layout: Layout, sheet: string, dialect: Dialect = 'excel'): FrameCell[] {
  const kind = layout.kindOf(sheet);
  const cells: FrameCell[] = [];
  if (kind === 'contents') {
    cells.push([STD.titleRow, 2, '=Model_Entity'], [STD.nameRow, 2, modelNameFormula(layout)], [STD.entityRow, 2, '=Model_Prepared_By']);
  } else {
    cells.push([STD.titleRow, 1, SYMBOL.home]);
    if (layout.hasChecks()) cells.push([STD.nameRow, 1, `=IF(Chk_Errors=0,"${SYMBOL.clear}","${SYMBOL.failing}")`]);
    cells.push([STD.titleRow, 2, Object.hasOwn(layout.titles, sheet) ? layout.titles[sheet] : sheet],
      [STD.nameRow, 2, '=Model_Name'], [STD.entityRow, 2, '=Model_Entity']);
  }
  if (kind === 'timeline' || kind === 'settings') {
    const source = kind === 'settings';
    const prefix = sheetPrefix(SETTINGS, dialect);
    if (!source) cells.push([5, TOTAL_COL, 'Total']);
    for (const b of BLOCK) {
      cells.push([b.row, 2, b.label]);
      for (let p = 0; p < layout.periods; p++) {
        const c = FIRST_PERIOD_COL + p;
        const L = colLetter(c);
        cells.push([b.row, c, source ? b.source(c, L, p ? colLetter(c - 1) : null) : `=${prefix}${L}${b.row}`]);
      }
    }
  }
  return dialect === 'uno'
    ? cells.map(([r, c, v]) => [r, c, typeof v === 'string' && v.startsWith('=') ? unoSeparators(v) : v])
    : cells;
}

/** The hyperlinks in a sheet's header: A1 to the contents and A2 to the checks. */
export function standardFrameLinks(layout: Layout, sheet: string): { row: number; col: number; link: CellLink }[] {
  if (layout.kindOf(sheet) === 'contents') return [];
  const out: { row: number; col: number; link: CellLink }[] = [
    { row: STD.titleRow, col: 1, link: { to: GO_CONTENTS, tip: 'Go to the table of contents' } }];
  if (layout.hasChecks() && layout.names.has(GO_CHECKS)) {
    out.push({ row: STD.nameRow, col: 1, link: { to: GO_CHECKS, tip: 'Go to the checks' } });
  }
  return out;
}
