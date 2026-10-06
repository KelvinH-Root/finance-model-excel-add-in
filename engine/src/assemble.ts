// Layout (steps 4 and 7 of an insert): module rows on their area's sheet, formulas held as
// structural markers naming row ids, names, link records, charts and key outputs.

import { assuranceSheets, type InputEntry } from './assurance.ts';
import { AssemblyError, codeWords, groupName, namePart, pick } from './frame.ts';
import { ChartSpec, Layout, LRow, type Headline, type LinkRecord } from './layout.ts';
import type { RowDef } from './library.ts';
import type { Model } from './model.ts';
import { navigate } from './navigate.ts';
import { resolve, type Block } from './resolve.ts';

const SUM = /\[sum:([\w.]+)\]/g;
const RANGE = /\[range:(\w+)\]/g;
const PREV = /\[(\w+)@prev\]/g;
const REF = /\[(\w+)\]/g;
const SET = /\$([a-z]\w*)/g;

/**
 * A module's formula template as a marker formula: [key] is this period of a row, [key@prev]
 * the period before, [range:key] the whole timeline, [sum:link] the rows collected for a link,
 * $setting a setting's name, {p} the period number, {periods} the count, {src} and {src_range}
 * the sending row of a mirror block or collect row.
 */
export function compile(tpl: string | null | undefined, b: Block, keys: Map<string, string>, names: Map<string, string>,
  collects: Map<string, string[]>, src: string | null): string | null {
  if (tpl === null || tpl === undefined) return null;
  const where = `${b.mod.id} (${b.title})`;
  const key = (k: string, kind: string) => {
    const id = keys.get(k);
    if (id === undefined) throw new AssemblyError(`${where}: no row '${k}'`);
    return `«${kind}|${id}»`;
  };
  let out = tpl.replace(SUM, (_m, link: string) => {
    const ids = collects.get(link);
    if (ids === undefined) throw new AssemblyError(`${where}: [sum:${link}] but the module does not take ${link}`);
    return ids.length ? `SUM(«R|${ids[0]}»:«R|${ids[ids.length - 1]}»)` : '0';
  });
  out = out.replace(RANGE, (_m, k: string) => key(k, 'A'));
  out = out.replace(PREV, (_m, k: string) => key(k, 'P'));
  out = out.replace(REF, (_m, k: string) => key(k, 'R'));
  out = out.replace(SET, (_m, k: string) => {
    const nm = names.get(k);
    if (nm === undefined) throw new AssemblyError(`${where}: no setting '${k}'`);
    return nm;
  });
  out = out.replaceAll('{p}', '«N»').replaceAll('{periods}', '«T»');
  for (const [token, kind] of [['{src_range}', 'A'], ['{src}', 'R']] as const) {
    if (out.includes(token)) {
      if (src === null) throw new AssemblyError(`${where}: ${token} is only allowed in a mirror module or a collect row`);
      out = out.replaceAll(token, `«${kind}|${src}»`);
    }
  }
  return out;
}

const isGroupBinding = (v: unknown): v is { group: string } =>
  v !== null && typeof v === 'object' && !Array.isArray(v) && 'group' in v;

type CollectRow = { rid: string; pb: Block; prow: string } | { rid: string; pb: null; senders: [Block, string][] };

export function assemble(model: Model): Layout {
  const { blocks, producers } = resolve(model);
  const byId = new Map(blocks.map(b => [b.id, b]));
  const rowId = (b: Block, key: string) => `${b.id}/${key}`;
  const records: LinkRecord[] = [];
  const warnings: string[] = [];
  const names = new Map<string, string>();
  const sheets = new Map<string, LRow[]>();
  const headed = new Set<string>();
  const charts: ChartSpec[] = [];
  const headlines: Headline[] = [];
  const inputs: InputEntry[] = [];

  for (const b of blocks) {
    const mod = b.mod;
    let rows = sheets.get(mod.area);
    if (!rows) sheets.set(mod.area, rows = []);
    const instTitle = model.title(b.inst);
    const mirror = b.src !== null;
    if (!headed.has(b.inst.uid)) {   // one section bar per instance; a mirror has one sub-block per sender
      headed.add(b.inst.uid);
      rows.push(new LRow(`${b.inst.uid}/heading`, 'heading', instTitle));
    }
    if (mirror) rows.push(new LRow(`${b.id}/subheading`, 'subheading', b.title.split(': ').slice(1).join(': '), { indent: 1 }));

    // Settings become named input cells (step 7 registers the names).
    const setNames = new Map<string, string>();
    const prefix = pick(mod, 'code', 'M');
    for (const s of mod.settings || []) {
      const nm = `${prefix}${b.inst.number}_${namePart(s.key)}`;
      setNames.set(s.key, nm);
      const rid = rowId(b, `set.${s.key}`);
      const stored = pick(b.inst.settings, s.key, null);
      let value: unknown = stored;
      let link: string | null = null;
      if (isGroupBinding(stored)) {   // bound to a group assumption
        if (!s.group) throw new AssemblyError(`${instTitle}: ${s.label} is bound but has no group assumption`);
        link = pick(s.group, 'formula', '={item}').replaceAll('{item}', groupName(stored.group));
        value = null;
      }
      rows.push(new LRow(rid, 'setting', s.label, { indent: 1, unit: pick(s, 'unit', ''), value, name: nm, link }));
      names.set(nm, rid);
      if (s.display) continue;   // a view setting (the month a chart starts on) is not an assumption, so not in the register
      inputs.push({
        name: nm, id: rid, label: s.label, title: instTitle, unit: pick(s, 'unit', ''),
        bound: value === null && link !== null && isGroupBinding(stored) ? stored.group : false,
        item: s.group ? pick(s.group, 'item', null) : null,
      });
    }

    // Row ids for keys first, so formulas can point forwards as well as back.
    const keys = new Map<string, string>();
    for (const r of mod.rows || []) if (r.key !== undefined) keys.set(r.key, rowId(b, r.key));
    const collects = new Map<string, string[]>();
    const planned: [RowDef, CollectRow[]][] = [];
    for (const r of mod.rows || []) {
      if (r.collect !== undefined) {
        const link = r.collect;
        const spec = (mod.inputs || []).find(i => i.link === link);
        if (!spec) throw new AssemblyError(`${mod.id}: collects ${link} but does not list it in inputs`);
        const found = producers.get(link) || [];
        if (pick(spec, 'mode', 'each') === 'each') {
          // a keyed collect is a second set of rows for the same link (a chart window, say)
          const collectRows: CollectRow[] = found.map(([pb, prow]) => ({
            rid: r.key !== undefined ? `${b.id}/${r.key}/${pb.id}` : `${b.id}/in/${link}/${pb.id}`, pb, prow,
          }));
          collects.set(r.key !== undefined ? r.key : link, collectRows.map(c => c.rid));
          planned.push([r, collectRows]);
        } else {   // total: one row adding every sender
          const rid = `${b.id}/in/${link}`;
          collects.set(link, [rid]);
          planned.push([r, [{ rid, pb: null, senders: found }]]);
        }
      } else {
        planned.push([r, []]);
      }
    }
    for (const spec of mod.inputs || []) {
      if (spec.required && !(producers.get(spec.link) || []).length) {
        warnings.push(`${b.title}: required link ${spec.link} has no sender`);
      }
    }

    const src = b.src ? rowId(byId.get(b.src[0])!, b.src[1]) : null;
    if (mirror) records.push({ link: mod.mirror!, mode: 'mirror', from: src!, to: b.id });
    for (const [r, collectRows] of planned) {
      if (r.section !== undefined) {
        rows.push(new LRow(`${b.id}/section/${namePart(r.section)}`, 'section', r.section, { indent: 1 }));
        continue;
      }
      if (r.collect !== undefined) {
        const link = r.collect;
        for (const c of collectRows) {
          if (c.pb === null) {   // total mode
            const f = '=' + (c.senders.length ? c.senders.map(([x, k]) => `«R|${rowId(x, k)}»`).join('+') : '0');
            rows.push(new LRow(c.rid, 'series', `${link} (all senders)`, {
              indent: 2, unit: pick(r, 'unit', ''), formula: f, total: pick(r, 'total', 'sum'),
            }));
            for (const [x, k] of c.senders) records.push({ link, mode: 'total', from: rowId(x, k), to: c.rid });
          } else {
            const prowDef = (c.pb.mod.rows || []).find(x => x.key === c.prow);
            if (!prowDef) throw new AssemblyError(`${c.pb.mod.id}: no row '${c.prow}' for ${link}`);
            const label = link.startsWith('check.') ? `${c.pb.title}: ${prowDef.label}` : c.pb.title;
            const sender = rowId(c.pb, c.prow);
            const f = r.formula !== undefined ? compile(r.formula, b, keys, setNames, collects, sender) : `=«R|${sender}»`;
            rows.push(new LRow(c.rid, 'series', label, {
              indent: 2, unit: pick(r, 'unit', pick(prowDef, 'unit', '')), formula: f,
              total: pick(r, 'total', pick(prowDef, 'total', 'sum')), span: pick(r, 'span', null),
            }));
            records.push({ link, mode: 'each', from: sender, to: c.rid });
          }
        }
        continue;
      }
      if (r.key === undefined) throw new AssemblyError(`${mod.id}: a row has no key, section or collect`);
      const rid = keys.get(r.key)!;
      const nm = pick(r, 'name', null);
      if (nm) names.set(nm, rid);
      const heads = r.headline || [];
      for (const h of Array.isArray(heads) ? heads : [heads]) {
        headlines.push({ id: rid, label: h.label, measure: pick(h, 'measure', 'last'), unit: pick(r, 'unit', ''),
          name: 'KO_' + codeWords(h.label) });
      }
      rows.push(new LRow(rid, 'series', r.label as string, {
        indent: mirror ? 2 : 1, unit: pick(r, 'unit', ''), style: r.check ? 'check' : pick(r, 'style', ''), name: nm,
        first: compile(r.first, b, keys, setNames, collects, src),
        formula: compile(r.formula, b, keys, setNames, collects, src),
        total: pick(r, 'total', 'sum'), span: pick(r, 'span', null),
      }));
    }
    rows.push(new LRow(`${b.id}/end`, 'blank', ''));
    charts.push(...moduleCharts(b, model, keys, collects, rows));
  }

  const linkedIn = new Set(records.map(rec => rec.from));
  for (const [link, list] of producers) {
    for (const [pb, prow] of list) {
      if (!linkedIn.has(rowId(pb, prow)) && !link.startsWith('check.')) {
        warnings.push(`${pb.title}: ${link} is not taken by any module`);
      }
    }
  }

  if (model.assured()) assuranceSheets(model, sheets, names, inputs);
  const layout = new Layout(model.periods, [], names, records, warnings, new Map(blocks.map(b => [b.id, b.title])),
    charts, headlines);
  navigate(layout, model, blocks, sheets);
  return layout;
}

/** The charts a module declares, with every series resolved to the rows it shows now. */
function moduleCharts(b: Block, model: Model, keys: Map<string, string>, collects: Map<string, string[]>,
  rows: LRow[]): ChartSpec[] {
  const out: ChartSpec[] = [];
  const spans = new Map(rows.map(r => [r.id, r.span]));
  for (const c of b.mod.charts || []) {
    const where = `${b.mod.id} chart ${c.key}`;
    const cat = keys.get(c.categories);
    if (cat === undefined) throw new AssemblyError(`${where}: no row '${c.categories}' for the categories`);
    const series: [string, 'column' | 'line'][] = [];
    for (const spec of c.series) {
      const kind = spec.line ? 'line' : 'column';
      if (spec.each !== undefined) {
        const ids = collects.get(spec.each);
        if (ids === undefined) throw new AssemblyError(`${where}: no collected rows '${spec.each}'`);
        for (const rid of ids) series.push([rid, kind]);
      } else if (spec.row !== undefined && keys.has(spec.row)) {
        series.push([keys.get(spec.row)!, kind]);
      } else {
        throw new AssemblyError(`${where}: no row '${spec.row}'`);
      }
    }
    out.push(new ChartSpec(`${b.id}/chart/${c.key}`, b.mod.area, c.title, `${b.inst.uid}/heading`, cat,
      spans.get(cat) || model.periods, series));
  }
  return out;
}
