// Sections, cover sheets, the contents and the navigation names, from the sheets the modules fill.
//
// A numbered contents with lettered sheets and the modules on each, a cover sheet before each
// section, a link to the contents (A1) and to the checks (A2) on every sheet, and links that
// are names (HL_...) so they survive a rename. Everything here is rebuilt from the layout, so a
// structural change rewrites it through the ordinary plan.

import { FRAME_KINDS, keyOutputFormula } from './assurance.ts';
import { AssemblyError, code, CONTENTS, pick, sheetLinkName, TOTAL_COL } from './frame.ts';
import { LRow, type Layout } from './layout.ts';
import type { SectionDef } from './library.ts';
import type { Model } from './model.ts';
import type { Block } from './resolve.ts';

export function navigate(layout: Layout, model: Model, blocks: Block[], sheets: Map<string, LRow[]>): void {
  const lib = model.lib;
  const sections: SectionDef[] = lib.sections.length ? lib.sections
    : [{ title: 'Model', cover: null, note: '', areas: lib.areas }];
  const present = sections
    .map(sec => [sec, sec.areas.filter(a => sheets.has(a))] as [SectionDef, string[]])
    .filter(([, areas]) => areas.length);
  const order: string[] = [CONTENTS];
  for (const [sec, areas] of present) {
    if (sec.cover) order.push(sec.cover);
    order.push(...areas);
  }
  layout.kinds = { [CONTENTS]: 'contents' };
  layout.titles = { [CONTENTS]: 'Contents' };
  for (const [sec, areas] of present) {
    if (sec.cover) {
      layout.kinds[sec.cover] = 'cover';
      layout.titles[sec.cover] = sec.title;
    }
    for (const a of areas) {
      layout.kinds[a] = Object.hasOwn(FRAME_KINDS, a) ? FRAME_KINDS[a] : 'timeline';
      layout.titles[a] = a;
    }
  }
  const headings = new Map<string, [string, string][]>();
  const seen = new Set<string>();
  for (const b of blocks) {
    if (seen.has(b.inst.uid)) continue;
    seen.add(b.inst.uid);
    const list = headings.get(b.mod.area);
    const entry: [string, string] = [b.inst.uid, model.title(b.inst)];
    if (list) list.push(entry);
    else headings.set(b.mod.area, [entry]);
  }

  const names = layout.names;
  names.set('HL_Home', `@${CONTENTS}`);
  layout.nameCols['HL_Home'] = 2;
  if (sheets.has('Checks')) {
    names.set('HL_Err_Chk', '@Checks');
    layout.nameCols['HL_Err_Chk'] = 2;
  }
  for (const sh of order) {
    names.set(sheetLinkName(sh), `@${sh}`);
    layout.nameCols[sheetLinkName(sh)] = 1;
  }

  const contents = [new LRow('contents/title', 'heading', 'Table of contents'), new LRow('contents/gap', 'blank', '')];
  present.forEach(([sec, areas], k0) => {
    const n = k0 + 1;
    const cover = sec.cover;
    if (cover) {
      contents.push(new LRow(`contents/section/${cover}`, 'toc', `section ${n} ${sec.title}`, { cells: {
        2: n, 3: `=HYPERLINK("#${sheetLinkName(cover)}",«S|${cover}»)` } }));
    }
    areas.forEach((a, k) => {
      const letter = String.fromCharCode(97 + k);
      contents.push(new LRow(`contents/sheet/${a}`, 'toc', `${letter}. ${a}`, { cells: {
        3: `${letter}.`, 4: `=HYPERLINK("#${sheetLinkName(a)}",«S|${a}»)` } }));
      const onSheet = headings.get(a) || [];
      const titleOf = new Map(onSheet);
      for (const [uid] of onSheet) {
        const nm = `HL_Toc_${code(uid)}`;
        names.set(nm, `${uid}/heading`);
        layout.nameCols[nm] = 2;
        contents.push(new LRow(`contents/module/${uid}`, 'toc', `- ${titleOf.get(uid)}`, { cells: {
          4: '-', 5: `=HYPERLINK("#${nm}",«B|${uid}/heading»)` } }));
      }
    });
  });
  contents.push(new LRow('contents/end', 'blank', ''));
  if (layout.hasChecks()) {
    contents.push(new LRow('contents/checks', 'heading', 'Checks'));
    contents.push(new LRow('contents/errors', 'toc', 'error checks', { style: 'check', cells: {
      3: '=HYPERLINK("#HL_Err_Chk","Error checks failing")', [TOTAL_COL]: '=Chk_Errors' } }));
    contents.push(new LRow('contents/alerts', 'toc', 'alerts', { cells: {
      3: '=HYPERLINK("#HL_Err_Chk","Alerts raised")', [TOTAL_COL]: '=Chk_Alerts' } }));
  }
  if (model.assured() && layout.headlines.length) {
    // Key outputs: what every structural change is checked against.
    contents.push(new LRow('contents/ko/gap', 'blank', ''));
    contents.push(new LRow('contents/ko', 'heading', 'Key outputs'));
    for (const h of layout.headlines) {
      const rid = `contents/ko/${h.name}`;
      contents.push(new LRow(rid, 'toc', h.label, { unit: h.unit, cells: { 3: h.label, [TOTAL_COL]: keyOutputFormula(h) } }));
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
      new LRow(`cover/${sh}/number`, 'toc', 'section number', { cells: { 2: `Section ${n}.` } }),
      new LRow(`cover/${sh}/home`, 'toc', 'link to the contents', { cells: { 2: '=HYPERLINK("#HL_Home","Go to contents")' } }),
      new LRow(`cover/${sh}/prev`, 'toc', 'link to the previous sheet', { cells: {
        2: `=HYPERLINK("#${sheetLinkName(prev)}","< "&«S|${prev}»)` } }),
      new LRow(`cover/${sh}/next`, 'toc', 'link to the next sheet', { cells: {
        2: `=HYPERLINK("#${sheetLinkName(nxt)}",«S|${nxt}»&" >")` } }),
      new LRow(`cover/${sh}/gap`, 'blank', ''),
      new LRow(`cover/${sh}/notes_title`, 'toc', 'notes heading', { style: 'bold', cells: { 2: 'Section notes' } }),
      new LRow(`cover/${sh}/notes`, 'toc', 'notes', { cells: { 2: pick(sec, 'note', '') } }),
    ]]);
  }
  layout.sheets = out;
}
