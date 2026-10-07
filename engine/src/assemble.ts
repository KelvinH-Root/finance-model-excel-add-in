// Layout (steps 4 and 7 of an insert): module rows on their area's sheet, formulas held as
// structural markers naming row ids, names, link records, charts and key outputs.

import { assuranceSheets, type InputEntry } from './assurance.ts';
import { AssemblyError, codeWords, groupName, namePart, pick } from './frame.ts';
import { ChartSpec, Layout, LRow, type Headline, type LinkRecord, type LRowFields, type RangeName, type RChart } from './layout.ts';
import type { ModuleDef, RowDef, SettingDef } from './library.ts';
import type { Model } from './model.ts';
import { navigate } from './navigate.ts';
import { expandReport, type ScenarioResult } from './reports.ts';
import { versionRows } from './versions.ts';
import { fiscalPosition, FRAME_LISTS, frameSheets, MONTH_NAMES, navigateStandard, type ListSpec } from './standard.ts';
import { resolve, type Block } from './resolve.ts';

const SUM = /\[sum:([\w.]+)\]/g;
const COL = /\[col:([\w.]+)\]/g;
const RANGE = /\[range:(\w+)\]/g;
const PREV = /\[(\w+)@prev(\d*)\]/g;
const HIST = /\[hist:(\w+)\]/g;
const SCN = /\[scn:(\w+)\]/g;
const REF = /\[(\w+)\]/g;
const SET = /\$([a-z]\w*)/g;
const SRC_SET = /\{src\.set:(\w+)\}/g;
const PREVSUM = /\[prevsum:([\w.]+)\]/g;
/** Timeline block rows a formula can read in its own period. */
const FRAME_TOKENS: Record<string, string> = {
  '{actual}': '«F14»=1', '{forecast}': '«F14»=0', '{fy}': '«F10»', '{month}': '«F11»', '{quarter}': '«F12»',
  '{period}': '«F9»', '{start}': '«F7»', '{end}': '«F8»',
};

/** What a formula can reach beyond its own block's rows and settings (the standard frame's tokens). */
export interface CompileExtras {
  /** This row's id, for {scenario}. */
  self?: string;
  /** Row keys with a historical line, and with a scenario adjustment. */
  hist?: Set<string>;
  /** Row keys with a historical balance sheet line, whose opening balance [key@prev] reads in the first month. */
  opening?: Set<string>;
  scn?: Set<string>;
  /** The sending block of a collect row, for {src.set:key}. */
  srcSettings?: Map<string, string>;
  /** This block's position in its module's list. */
  pos?: number;
  /** Link -> the rows that send it, for [prevsum:link]. */
  senders?: Map<string, string[]>;
  /** A report module's per-chart check cells, for {report.errors} and {report.alerts}. */
  reportChecks?: { errors: string[]; alerts: string[] };
}

/**
 * A module's formula template as a marker formula: [key] is this period of a row, [key@prev] the
 * period before and [key@prevN] N periods before, [range:key] the whole timeline, [sum:link] the
 * rows collected for a link and [col:link] the same rows as one range in this period, [hist:key]
 * the row's line on the historical statements, [scn:key] and {scenario} its scenario adjustment,
 * $setting a setting's name, {src.set:key} the sending block's setting, {pos} this block's place in
 * its list, {actual}, {fy}, {month} and the other timeline tokens this period's flags, {p} the
 * period number, {periods} the count, {src} and {src_range} the sending row of a mirror block or
 * collect row.
 */
export function compile(tpl: string | null | undefined, b: Block, keys: Map<string, string>, names: Map<string, string>,
  collects: Map<string, string[]>, src: string | null, extra: CompileExtras = {}): string | null {
  if (tpl === null || tpl === undefined) return null;
  const where = `${b.mod.id} (${b.title})`;
  const key = (k: string, kind: string) => {
    const id = keys.get(k);
    if (id === undefined) throw new AssemblyError(`${where}: no row '${k}'`);
    return `«${kind}|${id}»`;
  };
  let out = tpl.replace(PREVSUM, (_m, link: string) => {
    const ids = extra.senders?.get(link);
    if (ids === undefined) throw new AssemblyError(`${where}: [prevsum:${link}] but the module does not take ${link}`);
    return ids.length ? `(${ids.map(i => `«O|${i}»`).join('+')})` : '0';
  });
  for (const k of ['errors', 'alerts'] as const) {
    const token = `{report.${k}}`;
    if (!out.includes(token)) continue;
    if (!extra.reportChecks) throw new AssemblyError(`${where}: ${token} is only allowed in a report module`);
    const cells = extra.reportChecks[k];
    out = out.replaceAll(token, cells.length ? `(${cells.join('+')})` : '0');
  }
  out = out.replace(SUM, (_m, link: string) => {
    const ids = collects.get(link);
    if (ids === undefined) throw new AssemblyError(`${where}: [sum:${link}] but the module does not take ${link}`);
    return ids.length ? `SUM(«R|${ids[0]}»:«R|${ids[ids.length - 1]}»)` : '0';
  });
  out = out.replace(COL, (_m, link: string) => {
    const ids = collects.get(link);
    if (ids === undefined) throw new AssemblyError(`${where}: [col:${link}] but the module does not take ${link}`);
    return ids.length ? `«R|${ids[0]}»:«R|${ids[ids.length - 1]}»` : '0';
  });
  out = out.replace(HIST, (_m, k: string) => {
    if (!extra.hist?.has(k)) throw new AssemblyError(`${where}: [hist:${k}] but row '${k}' has no historical line`);
    return `«R|hist/${keys.get(k)}»`;
  });
  out = out.replace(SCN, (_m, k: string) => {
    if (!extra.scn?.has(k)) throw new AssemblyError(`${where}: [scn:${k}] but row '${k}' has no scenario adjustment`);
    return `«V|scn/${keys.get(k)}»`;
  });
  if (out.includes('{scenario}')) {
    if (!extra.self || !extra.scn || ![...extra.scn].some(k => keys.get(k) === extra.self)) {
      throw new AssemblyError(`${where}: {scenario} in a row that has no scenario adjustment`);
    }
    out = out.replaceAll('{scenario}', `«V|scn/${extra.self}»`);
  }
  out = out.replace(RANGE, (_m, k: string) => key(k, 'A'));
  // The month before a balance with a historical line: in the first month, its opening balance on the historical balance sheet.
  out = out.replace(PREV, (_m, k: string, n: string) => key(k, n ? `P${n}` : extra.opening?.has(k) ? 'Q' : 'P'));
  out = out.replace(REF, (_m, k: string) => key(k, 'R'));
  out = out.replace(SRC_SET, (_m, k: string) => {
    const nm = extra.srcSettings?.get(k);
    if (nm === undefined) throw new AssemblyError(`${where}: {src.set:${k}} but the sending module has no setting '${k}'`);
    return nm;
  });
  out = out.replace(SET, (_m, k: string) => {
    const nm = names.get(k);
    if (nm === undefined) throw new AssemblyError(`${where}: no setting '${k}'`);
    return nm;
  });
  if (out.includes('{pos}')) {
    if (extra.pos === undefined) throw new AssemblyError(`${where}: {pos} but the module keeps no list`);
    out = out.replaceAll('{pos}', String(extra.pos));
  }
  for (const [t, m] of Object.entries(FRAME_TOKENS)) out = out.replaceAll(t, m);
  out = out.replaceAll('{p}', '«N»').replaceAll('{periods}', '«T»');
  for (const [token, kind] of [['{src_range}', 'A'], ['{src}', 'R']] as const) {
    if (out.includes(token)) {
      if (src === null) throw new AssemblyError(`${where}: ${token} is only allowed in a mirror module or a collect row`);
      out = out.replaceAll(token, `«${kind}|${src}»`);
    }
  }
  return out;
}

/** The defined name a module setting's cell carries: Sel_ for drop-downs, Opt_ for check boxes. */
export function settingName(mod: ModuleDef, number: number, s: SettingDef): string {
  if (s.name) return s.name;
  const base = `${pick(mod, 'code', 'M')}${number}_${namePart(s.key)}`;
  if (s.choice || s.list) return `Sel_${base}`;
  if (s.check) return `Opt_${base}`;
  return base;
}

/** The list a module's choice setting reads: one list per module and setting. */
export const choiceList = (mod: ModuleDef, s: SettingDef) => `List_${pick(mod, 'code', 'M')}_${namePart(s.key)}`;

const isGroupBinding = (v: unknown): v is { group: string } =>
  v !== null && typeof v === 'object' && !Array.isArray(v) && 'group' in v;

type CollectRow = { rid: string; pb: Block; prow: string } | { rid: string; pb: null; senders: [Block, string][] };

/** A line on the historical statements, made for a module row that declares one. */
export interface HistoryLine {
  in: 'is' | 'bs';
  group: string;
  id: string;
  label: string;
  unit: string;
  values: (number | null)[];
  /** A balance's opening (the month before the model starts), typed in column I of the historical balance sheet. */
  opening: number | null;
}

/** A scenario adjustment line, made for a module row that declares one. */
export interface ScenarioLine {
  id: string;
  block: string;
  title: string;
  label: string;
  values: number[];
}

/** What the module rows ask the frame for: lists, historical lines and scenario lines. */
export interface FrameRequests {
  lists: ListSpec[];
  history: HistoryLine[];
  scenarios: ScenarioLine[];
  /** Outputs the Scenarios sheet's data table works out for every scenario (report modules' scenario charts). */
  results: ScenarioResult[];
}

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
  const std = model.info !== null;   // the standard frame adds spacer rows, list rules and validation
  const req: FrameRequests = { lists: [], history: [], scenarios: [], results: [] };
  const rcharts: RChart[] = [];
  const statements = blocks.find(x => x.mod.framework === 'statements') ?? null;
  const ranges = new Map<string, RangeName>();
  let store: { sheet: string; rows: LRow[] } | null = null;
  const listOf = (name: string, title: string, group: string): ListSpec => {
    let l = req.lists.find(x => x.name === name);
    if (!l) req.lists.push(l = { name, title, group, items: [] });
    return l;
  };
  const positions = new Map<string, number>();   // instance uid -> its place in its module's list
  const referenced = new Set<string>();
  const settingNames = new Map<string, Map<string, string>>();   // instance uid -> setting key -> name
  for (const inst of model.instances) {
    const mod = model.lib.module(inst.module);
    settingNames.set(inst.uid, new Map((mod.settings || []).map(s => [s.key, settingName(mod, inst.number, s)])));
  }
  const periods = model.periods;
  const seriesOf = (v: number | (number | null)[] | undefined, dflt: number | undefined): (number | null)[] =>
    Array.from({ length: periods }, (_, t) => (Array.isArray(v) ? (v[t] ?? null) : v ?? dflt ?? null));

  for (const b of blocks) {
    const mod = b.mod;
    let rows = sheets.get(mod.area);
    if (!rows) sheets.set(mod.area, rows = []);
    const instTitle = model.title(b.inst);
    const mirror = b.src !== null;
    const category = model.lib.kind(b.inst.module) === 'category';
    if (!headed.has(b.inst.uid)) {   // one section bar per instance; a mirror has one sub-block per sender
      headed.add(b.inst.uid);
      rows.push(new LRow(`${b.inst.uid}/heading`, 'heading', instTitle));
      if (mod.list) {
        if (!std) throw new AssemblyError(`${mod.id} keeps a list, which needs the standard frame`);
        const l = listOf(mod.list.name, mod.list.title, mod.title);
        if (!l.items.length) for (const f of mod.list.first || []) l.items.push({ value: f, style: 'lu.text' });
        l.items.push({ value: `=«B|${b.inst.uid}/heading»`, style: 'lu.text' });
        positions.set(b.inst.uid, l.items.length);
      }
    }
    if (mirror) rows.push(new LRow(`${b.id}/subheading`, 'subheading', b.title.split(': ').slice(1).join(': '), { indent: 1 }));
    const bodyStart = rows.length;   // the standard frame spaces sections and checks from what comes before them in the block

    // Settings become named input cells (step 7 registers the names).
    const setNames = settingNames.get(b.inst.uid)!;
    const keys = new Map<string, string>();
    for (const r of mod.rows || []) if (r.key !== undefined) keys.set(r.key, rowId(b, r.key));
    const hist = new Set((mod.rows || []).filter(r => r.history && r.key).map(r => r.key!));
    const opening = new Set((mod.rows || []).filter(r => r.history?.in === 'bs' && r.key).map(r => r.key!));
    const scn = new Set((mod.rows || []).filter(r => r.scenario && r.key).map(r => r.key!));
    const pos = positions.get(b.inst.uid);
    const when = (cond: string | undefined) => (cond ? `NOT(${compile(`=${cond}`, b, keys, setNames, new Map(), null, { pos })!.slice(1)})` : null);
    for (const s of mod.settings || []) {
      const nm = setNames.get(s.key)!;
      const rid = rowId(b, `set.${s.key}`);
      const stored = pick(b.inst.settings, s.key, null);
      let value: unknown = stored;
      let link: string | null = null;
      if (isGroupBinding(stored)) {   // bound to a group assumption
        if (!s.group) throw new AssemblyError(`${instTitle}: ${s.label} is bound but has no group assumption`);
        link = pick(s.group, 'formula', '={item}').replaceAll('{item}', groupName(stored.group));
        value = null;
      }
      const init: Partial<LRowFields> = { indent: 1, unit: pick(s, 'unit', ''), value, name: nm, link };
      if (s.choice || s.list) {
        if (!std) throw new AssemblyError(`${mod.id}: ${s.label} is a drop-down, which needs the standard frame`);
        const list = s.list ?? choiceList(mod, s);
        referenced.add(list);
        if (s.choice) {
          const l = listOf(list, `${mod.title}: ${s.label}`, mod.title);
          if (!l.items.length) for (const c of s.choice) l.items.push({ value: c, style: 'lu.text' });
        } else {
          const shared = model.lib.lists.find(x => x.name === list);
          if (shared && !req.lists.some(x => x.name === list)) {
            req.lists.push({ name: list, title: shared.title, group: 'Shared lists', items: shared.items.map(value => ({ value, style: 'lu.text' })) });
          }
        }
        Object.assign(init, { role: 'cellLink', control: { kind: 'drop', list },
          valid: { kind: 'whole', min: 1, max: `ROWS(${list})`, message: 'Choose from the drop-down list.' } });
      } else if (s.check) {
        if (!std) throw new AssemblyError(`${mod.id}: ${s.label} is a check box, which needs the standard frame`);
        Object.assign(init, { role: 'cellLink', control: { kind: 'check' }, valid: { kind: 'logical', message: 'Tick or clear the box.' } });
      } else if (std && !link) {
        init.valid = { kind: 'decimal', message: 'Type a number.' };
      }
      const inactive = when(s.when);
      if (inactive) init.inactive = inactive;
      rows.push(new LRow(rid, 'setting', s.label, init));
      names.set(nm, rid);
      if (s.display) continue;   // a view setting (the month a chart starts on) is not an assumption, so not in the register
      inputs.push({
        name: nm, id: rid, label: s.label, title: instTitle, unit: pick(s, 'unit', ''),
        bound: value === null && link !== null && isGroupBinding(stored) ? stored.group : false,
        item: s.group ? pick(s.group, 'item', null) : null,
      });
    }

    // A report module's charts: its selections, the chart grid and a table per chart.
    let reportChecks: CompileExtras['reportChecks'];
    if (mod.framework === 'report') {
      if (!std || !model.info) throw new AssemblyError(`${mod.id} is a report module, which needs the standard frame`);
      if (!statements?.mod.report) throw new AssemblyError(`${instTitle} reads the financial statements, which are not in the model`);
      const fsb = statements;
      const t = model.info.timeline;
      const first = fiscalPosition(t.start, t.fyEndMonth);
      const [y0, m0] = t.start.split('-').map(Number);
      const settingOf = (k: string) => ((mod.settings || []).some(x => x.key === k) ? setNames.get(k)! : null);
      const out = expandReport({
        block: b.id, code: `${pick(mod, 'code', 'M')}${b.inst.number}`, title: instTitle, charts: mod.reports || [],
        fs: fsb.id, spec: fsb.mod.report!,
        members: (link: string) => {
          const spec = (fsb.mod.inputs || []).find(i => i.link === link);
          if (!spec || pick(spec, 'mode', 'each') !== 'each') throw new AssemblyError(`${instTitle}: the statements do not list ${link} line by line`);
          const found = producers.get(link) || [];
          return found.map(([pb, prow]) => `${fsb.id}/in/${link}/${pb.id}` + (found.filter(([x]) => x.id === pb.id).length > 1 ? `/${prow}` : ''));
        },
        year: settingOf('year'), month: settingOf('month'), compare: settingOf('compare'),
        yearShown: Number(b.inst.settings.year ?? 1), monthShown: Number(b.inst.settings.month ?? 1),
        fyLabel: (k: number) => `FY${first.year + k - 1}`,
        monthLabel: (p: number) => {
          const m = m0 - 1 + p - 1;
          return `${MONTH_NAMES[((m % 12) + 12) % 12].slice(0, 3)} ${y0 + Math.floor(m / 12)}`;
        },
        years: Math.ceil((first.month - 1 + periods) / 12),
        fyOf: (p: number) => Math.floor((first.month - 1 + p - 1) / 12) + 1,
        compare2: settingOf('compare2'), line: settingOf('line'), periods, target: settingOf('target'),
        budget: model.info.timeline.budget,
        scenarios: model.lib.scenarios?.names ?? null,
        results: req.results,
      });
      rows.push(...out.rows);
      for (const [nm, rid] of out.names) names.set(nm, rid);
      for (const c of out.charts) rcharts.push({ ...c, sheet: mod.area });
      reportChecks = { errors: out.errors, alerts: out.alerts };
    }

    // The versions module: the register of saved versions here, their values on the Version store sheet.
    if (mod.framework === 'versions') {
      if (!std) throw new AssemblyError(`${mod.id} keeps saved versions, which needs the standard frame`);
      if (!statements?.mod.report) throw new AssemblyError(`${instTitle} saves the financial statements, which are not in the model`);
      const out = versionRows(b.id, b.inst.data.versions ?? [], statements.mod.report, periods);
      rows.push(...out.register);
      store = { sheet: mod.store ?? 'Version store', rows: out.store };
      for (const [nm, r] of out.ranges) ranges.set(nm, r);
      req.lists.push(...out.lists);
    }

    // Row ids for keys first, so formulas can point forwards as well as back.
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
          // a block that sends the link from more than one row (two checks, say) gets a row id for each
          // (the Phase 0 proof frame keeps the proof's ids, which its golden fixtures compare)
          const twice = (id: string) => std && found.filter(([x]) => x.id === id).length > 1;
          const collectRows: CollectRow[] = found.map(([pb, prow]) => ({
            rid: (r.key !== undefined ? `${b.id}/${r.key}/${pb.id}` : `${b.id}/in/${link}/${pb.id}`) + (twice(pb.id) ? `/${prow}` : ''), pb, prow,
          }));
          collects.set(r.key !== undefined ? r.key : link, collectRows.map(c => c.rid));
          planned.push([r, collectRows]);
        } else {   // total: one row adding every sender (a keyed one can be read by its key)
          const rid = r.key !== undefined ? rowId(b, r.key) : `${b.id}/in/${link}`;
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

    const senders = new Map((mod.inputs || []).map(i => [i.link, (producers.get(i.link) || []).map(([pb, prow]) => rowId(pb, prow))]));
    const src = b.src ? rowId(byId.get(b.src[0])!, b.src[1]) : null;
    if (mirror) records.push({ link: mod.mirror!, mode: 'mirror', from: src!, to: b.id });
    let checksSpaced = false;
    const spacer = (key: string, space: number) => {
      if (std && rows!.length > bodyStart) rows!.push(new LRow(`${b.id}/sp/${key}`, 'blank', '', { space }));
    };
    for (const [r, collectRows] of planned) {
      if (r.section !== undefined) spacer(namePart(r.section), 6);
      if (r.check && !checksSpaced) {
        spacer('checks', 6);
        checksSpaced = true;
      }
      if (r.section !== undefined) {
        rows.push(new LRow(`${b.id}/section/${namePart(r.section)}`, 'section', r.section, { indent: 1 }));
        continue;
      }
      if (r.collect !== undefined) {
        const link = r.collect;
        for (const c of collectRows) {
          if (c.pb === null) {   // total mode
            const f = '=' + (c.senders.length ? c.senders.map(([x, k]) => `«R|${rowId(x, k)}»`).join('+') : '0');
            rows.push(new LRow(c.rid, 'series', r.label ?? `${link} (all senders)`, {
              indent: r.label !== undefined ? 1 : 2, unit: pick(r, 'unit', ''), formula: f, total: pick(r, 'total', 'sum'),
              role: r.working ? 'working' : undefined, style: r.style ?? '',
            }));
            for (const [x, k] of c.senders) records.push({ link, mode: 'total', from: rowId(x, k), to: c.rid });
          } else {
            const prowDef = (c.pb.mod.rows || []).find(x => x.key === c.prow);
            if (!prowDef) throw new AssemblyError(`${c.pb.mod.id}: no row '${c.prow}' for ${link}`);
            const label = link.startsWith('check.') ? `${c.pb.title}: ${prowDef.label}`
              : r.label !== undefined ? r.label.replaceAll('{title}', c.pb.title) : c.pb.title;
            const sender = rowId(c.pb, c.prow);
            const f = r.formula !== undefined
              ? compile(r.formula, b, keys, setNames, collects, sender, { pos, srcSettings: settingNames.get(c.pb.inst.uid) })
              : `=«R|${sender}»`;
            rows.push(new LRow(c.rid, 'series', label, {
              indent: 2, unit: pick(r, 'unit', pick(prowDef, 'unit', '')), formula: f,
              total: pick(r, 'total', pick(prowDef, 'total', 'sum')), span: pick(r, 'span', null),
              role: r.working ? 'working' : undefined, style: r.italic ? 'italic' : pick(r, 'style', ''),
            }));
            records.push({ link, mode: 'each', from: sender, to: c.rid });
          }
        }
        if (std && collectRows.length && collectRows[0].pb !== null && !r.working) rows[rows.length - 1].role = 'last';
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
      const extra: CompileExtras = { self: rid, hist, scn, pos, opening, senders, reportChecks };
      const init: Partial<LRowFields> = {
        indent: mirror ? 2 : 1, unit: pick(r, 'unit', ''), style: r.check ? 'check' : r.italic ? 'italic' : pick(r, 'style', ''), name: nm,
        first: compile(r.first, b, keys, setNames, collects, src, extra),
        formula: compile(r.formula, b, keys, setNames, collects, src, extra),
        // rates, days and month counts have no meaningful total
        total: pick(r, 'total', std && ['%', 'days', 'months'].includes(pick(r, 'unit', '')) ? 'none' : 'sum'),
        span: pick(r, 'span', null),
      };
      if (r.working) init.role = 'working';
      if (r.input) {
        if (!std) throw new AssemblyError(`${mod.id}: ${r.label} is a time series input, which needs the standard frame`);
        init.input = r.input;
        init.values = seriesOf(b.inst.data.series?.[r.key], r.default);
        init.first = init.formula = null;
      }
      const off = [r.input === 'forecast' ? '«F14»=1' : r.input === 'actual' ? '«F14»=0' : null, when(r.when)].filter(Boolean);
      if (off.length) init.inactive = off.length > 1 ? `OR(${off.join(',')})` : off[0]!;
      rows.push(new LRow(rid, 'series', r.label as string, init));
      if (r.history) {
        if (!std) throw new AssemblyError(`${mod.id}: ${r.label} has a historical line, which needs the standard frame`);
        const tpl = r.history.label ?? (category ? '{title}' : r.label ?? r.key);
        req.history.push({ in: r.history.in, group: r.history.group, id: `hist/${rid}`, unit: pick(r, 'unit', ''),
          label: tpl.replaceAll('{title}', instTitle), values: seriesOf(b.inst.data.history?.[r.key], undefined),
          opening: r.history.in === 'bs' ? b.inst.data.opening?.[r.key] ?? 0 : null });
      }
      if (r.scenario) {
        if (!std) throw new AssemblyError(`${mod.id}: ${r.label} has a scenario adjustment, which needs the standard frame`);
        const n = model.lib.scenarios?.names.length ?? 3;
        const v = b.inst.data.scenarios?.[r.key] ?? [];
        req.scenarios.push({ id: `scn/${rid}`, block: b.inst.uid, title: instTitle, label: typeof r.scenario === 'string' ? r.scenario : r.label ?? r.key,
          values: Array.from({ length: n }, (_, k) => v[k] ?? 0) });
      }
    }
    rows.push(new LRow(`${b.id}/end`, 'blank', '', std ? { space: 9, level: 0 } : {}));
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

  for (const list of referenced) {   // a category list no instance has joined yet holds its first items only
    if (req.lists.some(x => x.name === list) || list.startsWith('List_') && FRAME_LISTS.has(list)) continue;
    const keeper = [...model.lib.modules.values()].find(m => m.list?.name === list);
    if (!keeper) throw new AssemblyError(`a drop-down reads ${list}, which no module or library list keeps`);
    req.lists.push({ name: list, title: keeper.list!.title, group: keeper.title,
      items: (keeper.list!.first || []).map(value => ({ value, style: 'lu.text' })) });
  }

  if (model.assured()) assuranceSheets(model, sheets, names, inputs);
  if (store) sheets.set(store.sheet, store.rows);
  const tables = std ? frameSheets(model, sheets, names, req) : [];
  const layout = new Layout(model.periods, [], names, records, warnings, new Map(blocks.map(b => [b.id, b.title])),
    charts, headlines);
  layout.rcharts = rcharts;
  layout.dataTables = tables;
  for (const [nm, r] of ranges) layout.ranges.set(nm, r);
  if (std) navigateStandard(layout, model, blocks, sheets, req.lists);
  else navigate(layout, model, blocks, sheets);
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
