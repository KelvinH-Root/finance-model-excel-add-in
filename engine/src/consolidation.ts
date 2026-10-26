// Group consolidation (framework: consolidation): every entity's figures by group account and
// month in, every group's statements out, all as formulas, so the workbook consolidates without
// the add-in. The logic is the Phase 0 proof's (prototypes/consolidation), month by month:
//
// - Entities hold the tree (parent, share held), from which each entity's ancestors, its path
//   and whether it sits in each group are worked out. A group is an entity that produces
//   accounts, or an NCI node (an entity its parent holds less than 100% of).
// - Entity data holds each entity's opening balance and monthly movement by group account. Every
//   figure after that is cumulative from the model's start: a balance sheet account is its
//   balance at the month end, an income statement account everything since the start, so each
//   month's trial balance nets to nil. The statements take the month's movement from it.
// - The intercompany register (transactions) and the intergroup balances (pairs, month by month)
//   are what Home Hub's matching hands over. A pair is eliminated in every group holding both
//   sides; below that it is a related party.
// - Eliminations: the register's lines by type, balance pairs against each other, investments
//   against capital (the outside investors' share to NCI), and the margin still held by site
//   (out of WIP, then investment property, released to cost of sale as homes leave the group).
// - NCI is worked out once at each node from its own consolidation and carried into every group
//   above it, bearing its share of upstream margin.
// - By group: each group combined, its eliminations and the result, with its cash flow worked out
//   from the consolidated balance sheet (indirect). Group statements show the group chosen.

import { AssemblyError, FIRST_PERIOD_COL, TOTAL_COL } from './frame.ts';
import { LRow, type RangeName, type SheetKind } from './layout.ts';
import type { ListSpec } from './standard.ts';

/** One account of the group chart (the consolidation module's chart). Debit positive throughout. */
export interface GroupAccount {
  code: number;
  name: string;
  class: 'asset' | 'liability' | 'equity' | 'income' | 'expense' | 'attribution';
  /** Where it shows on the statements: revenue, cogs, other_income, staff, opex, other_expense, da, interest_income,
   * interest_expense, tax, nci_pl; ca, nca, cl, ncl, equity. */
  line: string;
  /** Its place on the cash flow statement: cash, op (working capital), inv, fin; none for income statement accounts. */
  cf?: 'cash' | 'op' | 'inv' | 'fin' | 'none';
  /** An income statement account the cash flow adds back (depreciation) and the asset whose movement it belongs with. */
  noncash?: number;
  /** What the consolidation uses it for: cash, investments, nci, nci_pl, diff, capital, retained, distributions, cost_of_sale. */
  role?: string;
}

export interface GroupEntity {
  code: number;
  name: string;
  parent: number | null;
  /** The share its parent holds (1 for wholly owned). */
  held: number;
  gst: boolean;
  status?: 'Actual' | 'Planned';
  /** The month it joins the group (1 = the model's first month; 0 or less: before the model starts). */
  from?: number;
  /** The name of the group it heads, when it produces group accounts. */
  group?: string;
  role?: string;
}

export interface GroupSite {
  code: string;
  name: string;
  units: number;
  /** By month: the entity holding the asset, the account it sits in, and the homes still held. */
  holder: number[];
  account: number[];
  held: number[];
}

export interface GroupInvestment { parent: number; sub: number; cost: number; period: number }

export interface GroupTransaction {
  period: number;
  type: string;
  seller: number;
  buyer: number;
  site?: string | null;
  sellerAcct: number;
  buyerAcct: number;
  costAcct?: number | null;
  sellerAmt: number;
  buyerAmt: number;
  gst?: number;
  margin?: number;
  desc?: string;
}

/** A balance pair: the receivable side (seller) and the payable side as its negative (buyer), opening then by month. */
export interface GroupBalance {
  seller: number;
  buyer: number;
  sellerAcct: number;
  buyerAcct: number;
  sellerOpening?: number;
  buyerOpening?: number;
  sellerBal: number[];
  buyerBal: number[];
  desc?: string;
}

export interface GroupFigure { entity: number; account: number; opening?: number; values: number[] }

/** What the consolidation module's instance holds (Home Hub hands it over; a recipe brings it for a demo). */
export interface GroupData {
  entities: GroupEntity[];
  sites?: GroupSite[];
  investments?: GroupInvestment[];
  transactions?: GroupTransaction[];
  balances?: GroupBalance[];
  figures?: GroupFigure[];
}

export const IC_TYPES = ['Trading, expensed', 'Trading, capitalised', 'Asset sale', 'At cost, expensed', 'At cost, capitalised',
  'Moved at cost', 'Distribution'] as const;

/** The sheets the consolidation lays out, in workbook order after its own (Group statements). */
export const CONS_SHEETS = {
  byGroup: 'By group', entities: 'Entities', data: 'Entity data', ic: 'Intercompany', bal: 'Intergroup balances', sites: 'Sites',
  margin: 'Unrealised margin', elim: 'Eliminations', nci: 'NCI',
} as const;

export interface ConsInput {
  block: string;
  chart: GroupAccount[];
  data: GroupData;
  periods: number;
  /** The Group shown drop-down's name (Sel_), from the module's settings. */
  groupSetting: string;
}

export interface ConsOut {
  main: LRow[];
  sheets: [string, LRow[]][];
  kinds: Record<string, SheetKind>;
  widths: Record<string, Record<number, number>>;
  heads: Record<string, [number, number, unknown][]>;
  /** Timeline sheets whose column I is not a total: what it holds (the opening). */
  totalHeads: Record<string, string>;
  names: [string, string, number][];
  ranges: [string, RangeName][];
  lists: ListSpec[];
}

interface Tier { head: number; name: string; node: boolean; accounts: boolean }

const IS_CLASSES = new Set(['income', 'expense']);
/** A name inside a sentence: its first letter in lower case, unless the first word is an acronym (GST payable). */
const lower = (s: string) => (/^[A-Z][a-z]/.test(s) ? s[0].toLowerCase() + s.slice(1) : s);
const round = (f: string) => `ROUND(${f},6)`;

export function consolidationSheets(inp: ConsInput): ConsOut {
  const { block, chart, data, periods } = inp;
  const P = (s: string) => `${block}/${s}`;
  const sel = inp.groupSetting;
  const ents = data.entities;
  if (!ents.length) throw new AssemblyError('the group consolidation has no entities');
  const byCode = new Map(ents.map(e => [e.code, e]));
  if (byCode.size !== ents.length) throw new AssemblyError('two entities share a code');
  for (const e of ents) {
    if (e.parent !== null && !byCode.has(e.parent)) throw new AssemblyError(`entity ${e.code}: no entity ${e.parent} to be its parent`);
  }
  const acct = new Map(chart.map(a => [a.code, a]));
  const role = (r: string) => {
    const a = chart.find(x => x.role === r);
    if (!a) throw new AssemblyError(`the group chart has no account for ${r}`);
    return a.code;
  };
  const CASH = role('cash'), INVEST = role('investments'), NCI = role('nci'), NCIPL = role('nci_pl'), DIFF = role('diff');
  const CAPITAL = role('capital'), RE = role('retained'), DIST = role('distributions'), COS = role('cost_of_sale');
  const consOnly = new Set([NCI, NCIPL, DIFF]);
  const ledger = [...chart].sort((a, b) => a.code - b.code);   // trial balances run in code order; statements in the chart's

  // Groups: entities that produce accounts, and NCI nodes, in tree order.
  const tiers: Tier[] = [];
  for (const e of ents) {
    const node = e.parent !== null && e.held < 1;
    if (e.group || node) tiers.push({ head: e.code, name: e.group ?? e.name, node, accounts: !!e.group });
  }
  if (!tiers.length) throw new AssemblyError('no entity produces group accounts');
  const ancestors = (c: number) => {
    const out: number[] = [];
    for (let p = byCode.get(c)!.parent; p !== null; p = byCode.get(p)!.parent) out.push(p);
    return out;
  };
  const inTier = (c: number, h: number) => c === h || ancestors(c).includes(h);
  const nodesBelow = (t: Tier) => tiers.filter(n => n.node && n.head !== t.head && inTier(n.head, t.head));

  // Past the timeline (after a gap column): each row's keys (entity and account), then its flags by group, then any extras.
  const K0 = FIRST_PERIOD_COL + periods + 1;
  const KE = K0, KA = K0 + 1;
  const flagCols = tiers.map((_, k) => K0 + 2 + k);
  const XC = K0 + 2 + tiers.length;
  const heads: Record<string, [number, number, unknown][]> = {};
  const widths: Record<string, Record<number, number>> = {};
  const names: [string, string, number][] = [];
  const ranges: [string, RangeName][] = [];
  const range = (nm: string, from: string, to: string, col: number, toCol?: number | 'timeline') =>
    ranges.push([nm, toCol === undefined ? { from, to, col } : { from, to, col, toCol }]);
  const totalHeads: Record<string, string> = {};
  const timelineHeads = (sheet: string, keys: [string, string] = ['Entity', 'Account'], extra: string[] = [], total = 'Opening') => {
    heads[sheet] = [[5, KE, keys[0]], [5, KA, keys[1]],
      ...tiers.map((t, k): [number, number, unknown] => [5, flagCols[k], `In ${t.head}`]),
      ...extra.map((x, j): [number, number, unknown] => [5, XC + j, x])];
    widths[sheet] = Object.fromEntries([KE, KA, ...flagCols, ...extra.map((_, j) => XC + j)].map(c => [c, 9]));
    totalHeads[sheet] = total;
  };
  const codeStyles = (e: boolean, a: boolean, input = false): Record<number, string> => {
    const st: Record<number, string> = {};
    if (e) st[KE] = input ? 'in.count' : 'count';
    if (a) st[KA] = input ? 'in.count' : 'count';
    return st;
  };
  const flagStyles = (extra = 0) => Object.fromEntries([...flagCols, ...Array.from({ length: extra }, (_, j) => XC + j)].map(c => [c, 'count']));

  // ------------------------------------------------------------------ Entities (a table)
  const EC = { parent: 9, held: 10, gst: 11, status: 12, from: 13, fromLabel: 14, node: 15, anc: 16, path: 22, inFirst: 23 };
  const ANC = 6;
  const tree = EC.inFirst + tiers.length, depth = tree + 1, join = tree + 2;
  const ent: LRow[] = [];
  ent.push(new LRow(P('ent/heading'), 'heading', 'Entities'));
  ent.push(new LRow(P('ent/note'), 'text', 'Home Hub is the master for actual entities; Add entity puts a planned one under its parent, '
    + 'in tree order, until Home Hub has it. Each entity sits in every group above it.', { indent: 1, role: 'note' }));
  const entHead: Record<number, unknown> = {
    2: 'Code', [EC.parent]: 'Parent', [EC.held]: 'Held', [EC.gst]: 'GST', [EC.status]: 'Status',
    [EC.from]: 'From', [EC.fromLabel]: 'Joins', [EC.node]: 'NCI node', [EC.path]: 'Path',
    [tree]: 'Order', [depth]: 'Depth', [join]: 'Early',
  };
  for (let j = 0; j < ANC; j++) entHead[EC.anc + j] = `${j + 1} up`;
  tiers.forEach((t, k) => { entHead[EC.inFirst + k] = `In ${t.head}`; });
  ent.push(new LRow(P('ent/head'), 'table', 'Name', { indent: 1, role: 'r.head', cells: entHead, colStyles: { 2: 'colHead' } }));
  const entId = (c: number) => P(`ent/${c}`);
  for (const e of ents) {
    const rid = entId(e.code);
    const C = (col: number) => `«C${col}|${rid}»`;
    const cells: Record<number, unknown> = {
      2: e.code, [EC.parent]: e.parent ?? '', [EC.held]: e.held, [EC.gst]: e.gst ? 'Yes' : 'No', [EC.status]: e.status ?? 'Actual',
      [EC.from]: e.from ?? 1,
      [EC.fromLabel]: `=IF(${C(EC.from)}<=0,"Before the start",TEXT(INDEX(List_Months,${C(EC.from)}),"mmm yyyy"))`,
      [EC.node]: `=IF(AND(${C(EC.parent)}<>"",${C(EC.held)}<1),"Yes","No")`,
    };
    for (let j = 0; j < ANC; j++) {
      const prev = j === 0 ? C(EC.parent) : C(EC.anc + j - 1);
      cells[EC.anc + j] = j === 0
        ? `=IF(${prev}="","",${prev})`
        : `=IF(${prev}="","",IF(IFERROR(INDEX(Ent_Parent,MATCH(${prev},Ent_Codes,0)),0)=0,"",INDEX(Ent_Parent,MATCH(${prev},Ent_Codes,0))))`;
    }
    cells[EC.path] = '="|"&' + Array.from({ length: ANC }, (_, j) => `IF(${C(EC.anc + ANC - 1 - j)}="","",${C(EC.anc + ANC - 1 - j)}&"|")`).join('&')
      + `&${C(2)}&"|"`;
    tiers.forEach((t, k) => { cells[EC.inFirst + k] = `=IF(ISNUMBER(SEARCH("|${t.head}|",${C(EC.path)})),1,0)`; });
    cells[tree] = `=IF(${C(EC.parent)}="",0,IF(ISNA(MATCH(${C(EC.parent)},Ent_Codes,0)),1,IF(MATCH(${C(EC.parent)},Ent_Codes,0)<MATCH(${C(2)},Ent_Codes,0),0,1)))`;
    cells[depth] = `=IF(${C(EC.anc + ANC - 1)}="",0,IF(IFERROR(INDEX(Ent_Parent,MATCH(${C(EC.anc + ANC - 1)},Ent_Codes,0)),0)=0,0,1))`;
    cells[join] = `=IF(${C(EC.from)}<=1,0,IF(SUMPRODUCT((ED_Ent=${C(2)})*(COLUMN(ED_Mov)-COLUMN(INDEX(ED_Mov,1,1))+1<${C(EC.from)})*ABS(ED_Mov))`
      + `+SUMPRODUCT((ED_Ent=${C(2)})*ABS(ED_Open))>0.001,1,0))`;
    const st: Record<number, string> = {
      2: 'in.count', [EC.parent]: 'in.count', [EC.held]: 'in.pct', [EC.gst]: 'in.text', [EC.status]: 'in.text', [EC.from]: 'in.count',
      [EC.fromLabel]: 'text', [EC.node]: 'text', [EC.path]: 'text', [tree]: 'check', [depth]: 'check', [join]: 'check',
    };
    for (let j = 0; j < ANC; j++) st[EC.anc + j] = 'count';
    tiers.forEach((_, k) => { st[EC.inFirst + k] = 'count'; });
    ent.push(new LRow(rid, 'table', e.name, { indent: 1, role: 'r.cols', cells, colStyles: st }));
  }
  const e0 = entId(ents[0].code), e1 = entId(ents[ents.length - 1].code);
  for (const [nm, col] of [['Ent_Codes', 2], ['Ent_Parent', EC.parent], ['Ent_Held', EC.held], ['Ent_GST', EC.gst],
    ['Ent_Status', EC.status], ['Ent_From', EC.from], ['Ent_Path', EC.path], ['Ent_Tree', tree], ['Ent_Depth', depth],
    ['Ent_Join', join]] as const) range(nm, e0, e1, col);
  tiers.forEach((t, k) => range(`Ent_In_${t.head}`, e0, e1, EC.inFirst + k));
  ent.push(new LRow(P('ent/end'), 'blank', '', { space: 9, level: 0 }));

  // Groups that consolidate
  ent.push(new LRow(P('grp/heading'), 'heading', 'Groups'));
  ent.push(new LRow(P('grp/note'), 'text', 'A pair is eliminated in every group holding both sides and is a related party below that. '
    + 'NCI is worked out once at each node and carried into the groups above it.', { indent: 1, role: 'note' }));
  ent.push(new LRow(P('grp/head'), 'table', 'Group', { indent: 1, role: 'r.head', colStyles: { 2: 'colHead' },
    cells: { 2: 'Head', 9: 'Number', 10: 'Members', 11: 'Accounts', 12: 'NCI node' } }));
  tiers.forEach((t, k) => {
    const rid = P(`grp/${t.head}`);
    ent.push(new LRow(rid, 'table', t.name, { indent: 1, role: 'r.cols', colStyles: { 2: 'count', 9: 'count', 10: 'count', 11: 'text', 12: 'text' },
      cells: { 2: t.head, 9: k + 1, 10: `=SUM(Ent_In_${t.head})`, 11: t.accounts ? 'Yes' : 'No', 12: t.node ? 'Yes' : 'No' } }));
  });
  ent.push(new LRow(P('grp/end'), 'blank', '', { space: 9, level: 0 }));

  // Investments in group entities
  const invs = data.investments ?? [];
  const IV = { sub: 9, period: 10, label: 11, cost: 12, held: 13 };
  if (invs.length) {
    ent.push(new LRow(P('inv/heading'), 'heading', 'Investments in group entities'));
    ent.push(new LRow(P('inv/note'), 'text', 'Every entity here was set up by the group, so there is no goodwill: a cost that is not the '
      + "parent's share of capital is an error.", { indent: 1, role: 'note' }));
    ent.push(new LRow(P('inv/head'), 'table', 'Sub', { indent: 1, role: 'r.head', colStyles: { 2: 'colHead' },
      cells: { 2: 'Parent', [IV.sub]: 'Sub', [IV.period]: 'Month', [IV.label]: 'Made', [IV.cost]: 'Cost', [IV.held]: 'Share' } }));
    invs.forEach((v, i) => {
      const rid = P(`inv/${i + 1}`);
      const C = (col: number) => `«C${col}|${rid}»`;
      ent.push(new LRow(rid, 'table', byCode.get(v.sub)?.name ?? String(v.sub), { indent: 1, role: 'r.cols',
        colStyles: { 2: 'in.count', [IV.sub]: 'in.count', [IV.period]: 'in.count', [IV.label]: 'text', [IV.cost]: 'in.int', [IV.held]: 'pct' },
        cells: { 2: v.parent, [IV.sub]: v.sub, [IV.period]: v.period, [IV.cost]: v.cost,
          [IV.label]: `=IF(${C(IV.period)}<=0,"Before the start",TEXT(INDEX(List_Months,${C(IV.period)}),"mmm yyyy"))`,
          [IV.held]: `=INDEX(Ent_Held,MATCH(${C(IV.sub)},Ent_Codes,0))` } }));
    });
    ent.push(new LRow(P('inv/end'), 'blank', '', { space: 9, level: 0 }));
  }
  widths[CONS_SHEETS.entities] = {
    [EC.parent]: 8, [EC.held]: 8, [EC.gst]: 9, [EC.status]: 9, [EC.from]: 7, [EC.fromLabel]: 12, [EC.node]: 9, [EC.path]: 26,
    ...Object.fromEntries(Array.from({ length: ANC }, (_, j) => [EC.anc + j, 7])),
    ...Object.fromEntries(tiers.map((_, k) => [EC.inFirst + k, 8])), [tree]: 8, [depth]: 8, [join]: 8,
  };

  // ------------------------------------------------------------------ Entity data (timeline)
  const figs = new Map((data.figures ?? []).map(f => [`${f.entity}|${f.account}`, f]));
  const dataAccts = ledger.filter(a => !consOnly.has(a.code));
  const ed: LRow[] = [];
  const edSheet = CONS_SHEETS.data;
  timelineHeads(edSheet);
  ed.push(new LRow(P('ed/heading'), 'heading', 'Figures by entity'));
  ed.push(new LRow(P('ed/note'), 'text', "Each entity's opening balance (column I) and the month's movement by group account, debit "
    + 'positive, as Home Hub hands them over. Each entity and month must net to nil.', { indent: 1, role: 'note' }));
  const edId = (e: number, a: number) => P(`ed/${e}/${a}`);
  const inFlagCells = (entityExpr: string) => Object.fromEntries(tiers.map((t, k) => [flagCols[k],
    `=INDEX(Ent_In_${t.head},MATCH(${entityExpr},Ent_Codes,0))`]));
  for (const e of ents) {
    ed.push(new LRow(P(`ed/${e.code}`), 'subheading', `${e.code} ${e.name}`, { level: 0 }));
    for (const a of dataAccts) {
      const f = figs.get(`${e.code}|${a.code}`);
      const rid = edId(e.code, a.code);
      ed.push(new LRow(rid, 'series', a.name, {
        indent: 2, unit: '$', role: 'opening', value: f?.opening ?? 0, input: 'all', total: 'none',
        values: Array.from({ length: periods }, (_, t) => f?.values[t] ?? 0),
        cells: { [KE]: e.code, [KA]: a.code, ...inFlagCells(`«C${KE}|${rid}»`) },
        colStyles: { ...codeStyles(true, true), ...flagStyles() },
      }));
    }
  }
  const edFirst = edId(ents[0].code, dataAccts[0].code), edLast = edId(ents[ents.length - 1].code, dataAccts[dataAccts.length - 1].code);
  range('ED_Ent', edFirst, edLast, KE);
  range('ED_Acct', edFirst, edLast, KA);
  range('ED_Open', edFirst, edLast, TOTAL_COL);
  range('ED_Mov', edFirst, edLast, FIRST_PERIOD_COL, 'timeline');
  tiers.forEach((t, k) => range(`ED_In_${t.head}`, edFirst, edLast, flagCols[k]));
  ed.push(new LRow(P('ed/tb/sp'), 'blank', '', { space: 9, level: 0 }));
  ed.push(new LRow(P('ed/tb/heading'), 'heading', 'Trial balances (each must be nil)'));
  for (const e of ents) {
    ed.push(new LRow(P(`ed/tb/${e.code}`), 'series', `${e.code} ${e.name}`, {
      indent: 1, unit: '$', style: 'check', total: 'none',
      formula: `=${round(`SUMIFS(INDEX(ED_Mov,0,«N»),ED_Ent,${e.code})`)}`,
      cells: { [TOTAL_COL]: `=${round(`SUMIFS(ED_Open,ED_Ent,${e.code})`)}` },
    }));
  }
  range('ED_TB', P(`ed/tb/${ents[0].code}`), P(`ed/tb/${ents[ents.length - 1].code}`), TOTAL_COL, 'timeline');

  // ------------------------------------------------------------------ Intercompany (a table of transactions)
  const txs = data.transactions ?? [];
  const IC = { month: 9, monthLabel: 10, type: 11, typeNo: 12, seller: 13, buyer: 14, site: 15, sAcct: 16, bAcct: 17, cAcct: 18,
    sAmt: 19, bAmt: 20, gst: 21, margin: 22, diff: 23, status: 24, counted: 25, lowest: 26, inFirst: 27 };
  const upFirst = IC.inFirst + tiers.length;
  const nodes = tiers.filter(t => t.node);
  const ic: LRow[] = [];
  ic.push(new LRow(P('ic/heading'), 'heading', 'Intercompany register'));
  ic.push(new LRow(P('ic/note'), 'text', "Home Hub's matching: each intergroup transaction with both sides, its type and the month it falls in. "
    + 'Eliminations come from these rows; nobody types an elimination journal.', { indent: 1, role: 'note' }));
  const icHead: Record<number, unknown> = {
    2: 'Id', [IC.month]: 'Month', [IC.monthLabel]: '', [IC.type]: 'Type', [IC.typeNo]: 'Type no', [IC.seller]: 'Seller',
    [IC.buyer]: 'Buyer', [IC.site]: 'Site', [IC.sAcct]: 'Seller GL', [IC.bAcct]: 'Buyer GL', [IC.cAcct]: 'Cost GL',
    [IC.sAmt]: 'Seller amt', [IC.bAmt]: 'Buyer amt', [IC.gst]: 'GST', [IC.margin]: 'Margin', [IC.diff]: 'Difference',
    [IC.status]: 'Status', [IC.counted]: 'Margin held', [IC.lowest]: 'Eliminated from',
  };
  tiers.forEach((t, k) => { icHead[IC.inFirst + k] = `In ${t.head}`; });
  nodes.forEach((n, j) => { icHead[upFirst + j] = `Up ${n.head}`; });
  ic.push(new LRow(P('ic/head'), 'table', 'Description', { indent: 1, role: 'r.head', cells: icHead, colStyles: { 2: 'colHead' } }));
  const icId = (n: number) => P(`ic/${n}`);
  const pathOf = (expr: string) => `INDEX(Ent_Path,MATCH(${expr},Ent_Codes,0))`;
  const bothIn = (s: string, b: string, h: number) => `INDEX(Ent_In_${h},MATCH(${s},Ent_Codes,0))*INDEX(Ent_In_${h},MATCH(${b},Ent_Codes,0))`;
  txs.forEach((x, i) => {
    const n = i + 1;
    const rid = icId(n);
    const C = (col: number) => `«C${col}|${rid}»`;
    const cells: Record<number, unknown> = {
      2: n, [IC.month]: x.period, [IC.type]: x.type, [IC.seller]: x.seller, [IC.buyer]: x.buyer, [IC.site]: x.site ?? '',
      [IC.sAcct]: x.sellerAcct, [IC.bAcct]: x.buyerAcct, [IC.cAcct]: x.costAcct ?? '', [IC.sAmt]: x.sellerAmt, [IC.bAmt]: x.buyerAmt,
      [IC.gst]: x.gst ?? 0, [IC.margin]: x.margin ?? 0,
      [IC.monthLabel]: `=IF(${C(IC.month)}<=0,"Before",TEXT(INDEX(List_Months,${C(IC.month)}),"mmm yyyy"))`,
      [IC.typeNo]: `=MATCH(${C(IC.type)},List_IC_Types,0)`,
      [IC.diff]: `=${round(`${C(IC.sAmt)}-${C(IC.bAmt)}`)}`,
      [IC.status]: `=IF(ABS(${C(IC.diff)})<0.001,"Matched","Break")`,
      [IC.counted]: `=IF(OR(${C(IC.typeNo)}=2,${C(IC.typeNo)}=3),${C(IC.margin)},0)`,
    };
    let lowest = '"None"';
    tiers.forEach((t, k) => {
      cells[IC.inFirst + k] = `=${bothIn(C(IC.seller), C(IC.buyer), t.head)}`;
      lowest = `IF(${C(IC.inFirst + k)}=1,"${t.name.replaceAll('"', '""')}",${lowest})`;
    });
    cells[IC.lowest] = `=${lowest}`;
    nodes.forEach((t, j) => {
      cells[upFirst + j] = `=INDEX(Ent_In_${t.head},MATCH(${C(IC.seller)},Ent_Codes,0))*(1-INDEX(Ent_In_${t.head},MATCH(${C(IC.buyer)},Ent_Codes,0)))`;
    });
    const st: Record<number, string> = {
      2: 'count', [IC.month]: 'in.count', [IC.monthLabel]: 'text', [IC.type]: 'in.text', [IC.typeNo]: 'count', [IC.seller]: 'in.count',
      [IC.buyer]: 'in.count', [IC.site]: 'in.text', [IC.sAcct]: 'in.count', [IC.bAcct]: 'in.count', [IC.cAcct]: 'in.count',
      [IC.sAmt]: 'in.int', [IC.bAmt]: 'in.int', [IC.gst]: 'in.int', [IC.margin]: 'in.int', [IC.diff]: 'check', [IC.status]: 'text',
      [IC.counted]: 'int', [IC.lowest]: 'text',
    };
    tiers.forEach((_, k) => { st[IC.inFirst + k] = 'count'; });
    nodes.forEach((_, j) => { st[upFirst + j] = 'count'; });
    ic.push(new LRow(rid, 'table', x.desc ?? x.type, { indent: 1, role: 'r.cols', cells, colStyles: st }));
  });
  if (txs.length) {
    const a = icId(1), b = icId(txs.length);
    for (const [nm, col] of [['IC_Month', IC.month], ['IC_Site', IC.site], ['IC_Status', IC.status], ['IC_Margin', IC.counted],
      ['IC_Type_No', IC.typeNo]] as const) range(nm, a, b, col);
    tiers.forEach((t, k) => range(`IC_In_${t.head}`, a, b, IC.inFirst + k));
  }
  ic.push(new LRow(P('ic/end'), 'blank', '', { space: 9, level: 0 }));
  widths[CONS_SHEETS.ic] = {
    [IC.month]: 7, [IC.monthLabel]: 10, [IC.type]: 20, [IC.typeNo]: 8, [IC.seller]: 8, [IC.buyer]: 8, [IC.site]: 6, [IC.sAcct]: 10,
    [IC.bAcct]: 10, [IC.cAcct]: 10, [IC.sAmt]: 11, [IC.bAmt]: 11, [IC.gst]: 8, [IC.margin]: 9, [IC.diff]: 11, [IC.status]: 9,
    [IC.counted]: 12, [IC.lowest]: 18, ...Object.fromEntries([...tiers.keys(), ...nodes.map((_, j) => tiers.length + j)].map(k => [IC.inFirst + k, 9])),
  };

  // ------------------------------------------------------------------ Intergroup balances (timeline)
  const bals = data.balances ?? [];
  const bal: LRow[] = [];
  const balSheet = CONS_SHEETS.bal;
  timelineHeads(balSheet);
  bal.push(new LRow(P('bal/heading'), 'heading', 'Balance pairs'));
  bal.push(new LRow(P('bal/note'), 'text', 'Each pair as the matching hands it over: the receivable side as it stands and the payable side '
    + 'as its negative, at the opening (column I) and every month end. The two must agree.', { indent: 1, role: 'note' }));
  const nameOf = (c: number) => byCode.get(c)?.name ?? String(c);
  const acctName = (c: number) => acct.get(c)?.name ?? String(c);
  bals.forEach((x, j) => {
    const n = j + 1;
    bal.push(new LRow(P(`bal/${n}`), 'subheading', `${x.desc ?? 'Balance pair'}: ${nameOf(x.seller)} and ${nameOf(x.buyer)}`, { level: 0 }));
    bal.push(new LRow(P(`bal/${n}/seller`), 'series', `${nameOf(x.seller)}, ${acctName(x.sellerAcct)}`, {
      indent: 2, unit: '$', role: 'opening', value: x.sellerOpening ?? 0, input: 'all', total: 'none', values: x.sellerBal,
      cells: { [KE]: x.seller, [KA]: x.sellerAcct }, colStyles: codeStyles(true, true, true) }));
    bal.push(new LRow(P(`bal/${n}/buyer`), 'series', `${nameOf(x.buyer)}, ${acctName(x.buyerAcct)} (as its negative)`, {
      indent: 2, unit: '$', role: 'opening', value: x.buyerOpening ?? 0, input: 'all', total: 'none', values: x.buyerBal,
      cells: { [KE]: x.buyer, [KA]: x.buyerAcct }, colStyles: codeStyles(true, true, true) }));
  });
  if (bals.length) {
    bal.push(new LRow(P('bal/diff/sp'), 'blank', '', { space: 9, level: 0 }));
    bal.push(new LRow(P('bal/diff/heading'), 'heading', 'Differences (each must be nil)'));
    bals.forEach((x, j) => {
      const n = j + 1;
      const s = P(`bal/${n}/seller`), b = P(`bal/${n}/buyer`);
      bal.push(new LRow(P(`bal/${n}/diff`), 'series', `${nameOf(x.seller)} and ${nameOf(x.buyer)}, ${acctName(x.sellerAcct)}`, {
        indent: 1, unit: '$', style: 'check', total: 'none', formula: `=${round(`«R|${s}»-«R|${b}»`)}`,
        cells: { [TOTAL_COL]: `=${round(`«V|${s}»-«V|${b}»`)}` } }));
    });
    range('BAL_Diff', P('bal/1/diff'), P(`bal/${bals.length}/diff`), TOTAL_COL, 'timeline');
    const netting = bals.map((x, j) => [x, j + 1] as const).filter(([x]) => acct.get(x.sellerAcct)?.role === 'netting');
    if (netting.length) {
      bal.push(new LRow(P('bal/net/sp'), 'blank', '', { space: 9, level: 0 }));
      bal.push(new LRow(P('bal/net/heading'), 'heading', 'Netting accounts left with a balance'));
      bal.push(new LRow(P('bal/net/note'), 'text', 'An on-charge not raised by the cut-off: the receiver accrues it until it is invoiced.',
        { indent: 1, role: 'note' }));
      for (const [x, n] of netting) {
        bal.push(new LRow(P(`bal/${n}/net`), 'series', `${nameOf(x.seller)}, owed by ${nameOf(x.buyer)}`, {
          indent: 1, unit: '$', total: 'none', formula: `=«R|${P(`bal/${n}/seller`)}»`, cells: { [TOTAL_COL]: `=«V|${P(`bal/${n}/seller`)}»` } }));
      }
      range('BAL_Netting', P(`bal/${netting[0][1]}/net`), P(`bal/${netting[netting.length - 1][1]}/net`), TOTAL_COL, 'timeline');
    }
  }

  // ------------------------------------------------------------------ Sites (timeline)
  const sites = data.sites ?? [];
  const st: LRow[] = [];
  const siteSheet = CONS_SHEETS.sites;
  totalHeads[siteSheet] = 'At the start';
  const codeInputs = Object.fromEntries(Array.from({ length: periods }, (_, t) => [FIRST_PERIOD_COL + t, 'in.count']));
  st.push(new LRow(P('site/heading'), 'heading', 'Sites'));
  st.push(new LRow(P('site/note'), 'text', 'Where each site\'s asset sits at every month end, and how many homes are still held (column I: '
    + 'the homes in the site). Margin stays out of the asset while it is held inside a group.', { indent: 1, role: 'note' }));
  const siteId = (s: string, k: string) => P(`site/${s}/${k}`);
  for (const s of sites) {
    st.push(new LRow(P(`site/${s.code}`), 'subheading', `${s.code} ${s.name}`, { level: 0 }));
    st.push(new LRow(siteId(s.code, 'holder'), 'series', 'Held by (entity)', { indent: 2, input: 'all', total: 'none', values: s.holder,
      colStyles: codeInputs }));
    st.push(new LRow(siteId(s.code, 'account'), 'series', 'In account', { indent: 2, input: 'all', total: 'none', values: s.account,
      colStyles: codeInputs }));
    st.push(new LRow(siteId(s.code, 'homes'), 'series', 'Homes still held', { indent: 2, unit: '#', role: 'opening', value: s.units,
      input: 'all', total: 'none', values: s.held }));
    tiers.forEach(t => {
      const h = siteId(s.code, 'holder');
      const test = (x: string) => `IF(ISNUMBER(SEARCH("|${t.head}|",${pathOf(x)})),1,0)`;
      st.push(new LRow(siteId(s.code, `in/${t.head}`), 'series', `In the ${t.name}`, { indent: 2, unit: 'flag', total: 'none',
        role: 'working', formula: `=${test(`«R|${h}»`)}`, cells: { [TOTAL_COL]: `=${test(`«C${FIRST_PERIOD_COL}|${h}»`)}` } }));
    });
  }

  // ------------------------------------------------------------------ Unrealised margin (timeline)
  const mg: LRow[] = [];
  const mgSheet = CONS_SHEETS.margin;
  timelineHeads(mgSheet, ['Register row', 'Site'], nodes.map(n => `Up ${n.head}`));
  const marginRows = txs.map((x, i) => [x, i + 1] as const).filter(([x]) => x.site && (x.type === IC_TYPES[1] || x.type === IC_TYPES[2]));
  const marginSites = sites.filter(s => marginRows.some(([x]) => x.site === s.code));
  for (const [x] of marginRows) if (!sites.some(s => s.code === x.site)) throw new AssemblyError(`register row for ${x.desc}: no site ${x.site}`);
  if (marginRows.length) {
    mg.push(new LRow(P('mg/heading'), 'heading', 'Margin by register row'));
    mg.push(new LRow(P('mg/note'), 'text', "Margin on intergroup charges and sales still held in the asset, at the share of the site's homes "
      + 'still held; column I is the opening.', { indent: 1, role: 'note' }));
    for (const [x, n] of marginRows) {
      const ic0 = icId(n);
      const C = (col: number) => `«C${col}|${ic0}»`;
      const homes = siteId(x.site!, 'homes');
      const start = `IF(${C(IC.month)}<=1,«V|${homes}»,INDEX(«A|${homes}»,${C(IC.month)}-1))`;
      const cells: Record<number, unknown> = {
        [KE]: n, [KA]: x.site!,
        [TOTAL_COL]: `=IF(${C(IC.month)}<=0,${C(IC.counted)},0)`,
        ...Object.fromEntries(tiers.map((_, k) => [flagCols[k], `=${C(IC.inFirst + k)}`])),
        ...Object.fromEntries(nodes.map((_, j) => [XC + j, `=${C(upFirst + j)}`])),
      };
      mg.push(new LRow(P(`mg/a/${n}`), 'series', `${n} ${x.desc ?? x.type}`, { indent: 2, unit: '$', total: 'none', cells,
        formula: `=IF(OR(${C(IC.counted)}=0,«N»<${C(IC.month)}),0,${C(IC.counted)}*«R|${homes}»/${start})`,
        colStyles: { [KE]: 'count', [KA]: 'text', ...flagStyles(nodes.length) } }));
    }
    const a = P(`mg/a/${marginRows[0][1]}`), b = P(`mg/a/${marginRows[marginRows.length - 1][1]}`);
    range('MG_Site', a, b, KA);
    range('MG_Open', a, b, TOTAL_COL);
    range('MG_Val', a, b, FIRST_PERIOD_COL, 'timeline');
    tiers.forEach((t, k) => range(`MG_In_${t.head}`, a, b, flagCols[k]));
    nodes.forEach((t, j) => range(`MG_Up_${t.head}`, a, b, XC + j));
    mg.push(new LRow(P('mg/b/sp'), 'blank', '', { space: 9, level: 0 }));
    mg.push(new LRow(P('mg/b/heading'), 'heading', 'By group and site: still held, added and released'));
    for (const t of tiers) {
      mg.push(new LRow(P(`mg/b/${t.head}`), 'subheading', t.name, { level: 0 }));
      for (const s of marginSites) {
        const inK = siteId(s.code, `in/${t.head}`);
        const held = P(`mg/b/${t.head}/${s.code}/held`), added = P(`mg/b/${t.head}/${s.code}/added`);
        mg.push(new LRow(held, 'series', `${s.code}: still held`, { indent: 2, unit: '$', total: 'last',
          formula: `=«R|${inK}»*SUMIFS(INDEX(MG_Val,0,«N»),MG_Site,"${s.code}",MG_In_${t.head},1)`,
          cells: { [TOTAL_COL]: `=«V|${inK}»*SUMIFS(MG_Open,MG_Site,"${s.code}",MG_In_${t.head},1)` } }));
        mg.push(new LRow(added, 'series', `${s.code}: added since the start`, { indent: 2, unit: '$', total: 'last',
          formula: `=SUMIFS(IC_Margin,IC_Site,"${s.code}",IC_In_${t.head},1,IC_Month,"<="&«N»)`,
          cells: { [TOTAL_COL]: `=SUMIFS(IC_Margin,IC_Site,"${s.code}",IC_In_${t.head},1,IC_Month,"<=0")` } }));
        mg.push(new LRow(P(`mg/b/${t.head}/${s.code}/released`), 'series', `${s.code}: released since the start`, { indent: 2, unit: '$',
          total: 'last', formula: `=«R|${added}»-«R|${held}»`, cells: { [TOTAL_COL]: `=«V|${added}»-«V|${held}»` } }));
      }
    }
    const upRows = tiers.flatMap(t => nodesBelow(t).map(n => [t, n] as const));
    if (upRows.length) {
      mg.push(new LRow(P('mg/c/sp'), 'blank', '', { space: 9, level: 0 }));
      mg.push(new LRow(P('mg/c/heading'), 'heading', 'Upstream margin by group and NCI node'));
      mg.push(new LRow(P('mg/c/note'), 'text', "Margin on sales from a partly owned node's group to the rest of the group, still held: the "
        + "node's NCI bears its share.", { indent: 1, role: 'note' }));
      for (const [t, n] of upRows) {
        const term = (val: string, cond: (s: GroupSite) => string) => marginSites.map(s =>
          `${cond(s)}*SUMIFS(${val},MG_Site,"${s.code}",MG_In_${t.head},1,MG_Up_${n.head},1)`).join('+') || '0';
        mg.push(new LRow(P(`mg/c/${t.head}/${n.head}`), 'series', `${t.name}: from ${n.name}`, { indent: 2, unit: '$', total: 'last',
          formula: `=${term('INDEX(MG_Val,0,«N»)', s => `«R|${siteId(s.code, `in/${t.head}`)}»`)}`,
          cells: { [TOTAL_COL]: `=${term('MG_Open', s => `«V|${siteId(s.code, `in/${t.head}`)}»`)}` } }));
      }
    }
  }

  // ------------------------------------------------------------------ Eliminations (timeline)
  const el: LRow[] = [];
  const elSheet = CONS_SHEETS.elim;
  timelineHeads(elSheet);
  el.push(new LRow(P('el/heading'), 'heading', 'Elimination lines'));
  el.push(new LRow(P('el/note'), 'text', 'Every line in the groups it applies in (the flags past the timeline), cumulative from the start '
    + 'as the trial balances are: from the register, the balance pairs, investments and unrealised margin.', { indent: 1, role: 'note' }));
  let elCount = 0;
  const elFirst: string[] = [];
  const line = (id: string, label: string, entity: unknown, account: unknown, open: string, val: string, flags: Record<number, unknown>,
    first?: string) => {
    elCount += 1;
    el.push(new LRow(id, 'series', label, { indent: 2, unit: '$', total: 'last', formula: val, first: first ?? null,
      cells: { [KE]: entity, [KA]: account, [TOTAL_COL]: open, ...flags }, colStyles: { ...codeStyles(true, true), ...flagStyles() } }));
    if (!elFirst.length) elFirst.push(id);
    elFirst[1] = id;
  };
  const flagsFrom = (f: (h: number, k: number) => string) => Object.fromEntries(tiers.map((t, k) => [flagCols[k], `=${f(t.head, k)}`]));
  const onlyIn = (k0: number) => Object.fromEntries(tiers.map((_, k) => [flagCols[k], k === k0 ? 1 : 0]));
  // (a) the register: two sides and any difference
  if (txs.length) el.push(new LRow(P('el/ic'), 'subheading', 'From the intercompany register', { level: 0 }));
  txs.forEach((x, i) => {
    const n = i + 1;
    const C = (col: number) => `«C${col}|${icId(n)}»`;
    const t = C(IC.typeNo);
    const [sa, ba, ca, se, be, sv, bv, m] = [C(IC.sAcct), C(IC.bAcct), C(IC.cAcct), C(IC.seller), C(IC.buyer), C(IC.sAmt), C(IC.bAmt), C(IC.margin)];
    const at = (amount: string) => [`=IF(${C(IC.month)}<=0,${amount},0)`, `=IF(«N»>=${C(IC.month)},${amount},0)`] as const;
    const flags = flagsFrom((_, k) => C(IC.inFirst + k));
    const sideA = `CHOOSE(${t},${sv},${sv},${sv},0,0,0,${bv})`;
    const sideB = `CHOOSE(${t},-${bv},-(${sv}-${m}),-(${sv}-${m}),0,0,0,-${bv})`;
    const diff = `IF(${t}=1,${bv}-${sv},0)`;
    line(P(`el/ic/${n}/a`), `${n} ${x.desc ?? x.type}: first side`, `=CHOOSE(${t},${se},${se},${se},"","","",${be})`,
      `=CHOOSE(${t},${sa},${sa},${sa},"","","",${ba})`, ...at(sideA), flags);
    line(P(`el/ic/${n}/b`), `${n} ${x.desc ?? x.type}: other side`, `=CHOOSE(${t},${be},${se},${se},"","","",${se})`,
      `=CHOOSE(${t},${ba},${ca},${ca},"","","",${sa})`, ...at(sideB), flags);
    line(P(`el/ic/${n}/d`), `${n} ${x.desc ?? x.type}: difference`, '', `=IF(${t}=1,${DIFF},"")`, ...at(diff), flags);
  });
  // (b) balance pairs against each other
  if (bals.length) el.push(new LRow(P('el/bal'), 'subheading', 'From the balance pairs', { level: 0 }));
  bals.forEach((x, j) => {
    const n = j + 1;
    const s = P(`bal/${n}/seller`), b = P(`bal/${n}/buyer`);
    const flags = flagsFrom(h => bothIn(String(x.seller), String(x.buyer), h));
    line(P(`el/bal/${n}/a`), `${nameOf(x.buyer)}: ${acctName(x.buyerAcct)}`, x.buyer, x.buyerAcct, `=«V|${b}»`, `=«R|${b}»`, flags);
    line(P(`el/bal/${n}/b`), `${nameOf(x.seller)}: ${acctName(x.sellerAcct)}`, x.seller, x.sellerAcct, `=-«V|${s}»`, `=-«R|${s}»`, flags);
    line(P(`el/bal/${n}/d`), `${nameOf(x.seller)} and ${nameOf(x.buyer)}: difference`, '', DIFF,
      `=«V|${s}»-«V|${b}»`, `=«R|${s}»-«R|${b}»`, flags);
  });
  // (c) investments against capital; the outside investors' share to NCI
  if (invs.length) el.push(new LRow(P('el/inv'), 'subheading', 'Investments against capital', { level: 0 }));
  invs.forEach((v, i) => {
    const n = i + 1;
    const inv = P(`inv/${n}`);
    const C = (col: number) => `«C${col}|${inv}»`;
    const flags = flagsFrom(h => bothIn(C(2), C(IV.sub), h));
    const cap = P(`el/inv/${n}/cap`), held = P(`el/inv/${n}/inv`);
    const mv = `SUMIFS(INDEX(ED_Mov,0,«N»),ED_Ent,${C(IV.sub)},ED_Acct,${CAPITAL})`;
    line(cap, `${nameOf(v.sub)}: its capital`, `=${C(IV.sub)}`, CAPITAL, `=-SUMIFS(ED_Open,ED_Ent,${C(IV.sub)},ED_Acct,${CAPITAL})`,
      `=«W|${cap}»-${mv}`, flags);
    line(held, `${nameOf(v.parent)}: its investment`, `=${C(2)}`, INVEST, `=-IF(${C(IV.period)}<=0,${C(IV.cost)},0)`,
      `=-IF(«N»>=${C(IV.period)},${C(IV.cost)},0)`, flags);
    line(P(`el/inv/${n}/nci`), `${nameOf(v.sub)}: outside investors' share of capital`, `=${C(IV.sub)}`, NCI,
      `=-(«V|${cap}»+«V|${held}»)`, `=-(«R|${cap}»+«R|${held}»)`, flags);
  });
  // (d) unrealised margin: out of the asset where it sits, released to cost of sale
  const assetAccounts = [...new Set(marginSites.flatMap(s => s.account))];
  if (marginRows.length) el.push(new LRow(P('el/mg'), 'subheading', 'Unrealised margin', { level: 0 }));
  tiers.forEach((t, k) => {
    for (const s of marginSites) {
      const held = P(`mg/b/${t.head}/${s.code}/held`), rel = P(`mg/b/${t.head}/${s.code}/released`);
      const acc = siteId(s.code, 'account');
      for (const a of assetAccounts) {
        line(P(`el/mg/${t.head}/${s.code}/${a}`), `${t.name}, ${s.code}: margin out of ${acctName(a)}`, '', a,
          `=-«V|${held}»*(«C${FIRST_PERIOD_COL}|${acc}»=${a})`, `=-«R|${held}»*(«R|${acc}»=${a})`, onlyIn(k));
      }
      line(P(`el/mg/${t.head}/${s.code}/cos`), `${t.name}, ${s.code}: margin released to cost of sale`, '', COS,
        `=-«V|${rel}»`, `=-«R|${rel}»`, onlyIn(k));
    }
  });
  if (!elCount) {   // nothing to eliminate: the names still need a row to point at
    line(P('el/none'), 'No eliminations', '', '', '=0', '=0', onlyIn(-1));
  }
  range('EL_Acct', elFirst[0], elFirst[1] ?? elFirst[0], KA);
  range('EL_Open', elFirst[0], elFirst[1] ?? elFirst[0], TOTAL_COL);
  range('EL_Val', elFirst[0], elFirst[1] ?? elFirst[0], FIRST_PERIOD_COL, 'timeline');
  tiers.forEach((t, k) => range(`EL_In_${t.head}`, elFirst[0], elFirst[1] ?? elFirst[0], flagCols[k]));
  if (invs.length) {
    el.push(new LRow(P('el/invchk/sp'), 'blank', '', { space: 9, level: 0 }));
    el.push(new LRow(P('el/invchk/heading'), 'heading', "Investments less the parent's share of capital (each must be nil)"));
    invs.forEach((v, i) => {
      const n = i + 1;
      const cap = P(`el/inv/${n}/cap`), held = P(`el/inv/${n}/inv`);
      const share = `«C${IV.held}|${P(`inv/${n}`)}»`;
      el.push(new LRow(P(`el/invchk/${n}`), 'series', `${nameOf(v.parent)} in ${nameOf(v.sub)}`, { indent: 1, unit: '$', style: 'check',
        total: 'none', formula: `=${round(`-«R|${held}»-${share}*«R|${cap}»`)}`, cells: { [TOTAL_COL]: `=${round(`-«V|${held}»-${share}*«V|${cap}»`)}` } }));
    });
    range('EL_InvChk', P('el/invchk/1'), P(`el/invchk/${invs.length}`), TOTAL_COL, 'timeline');
  }

  // ------------------------------------------------------------------ NCI (timeline)
  const nci: LRow[] = [];
  const nciSheet = CONS_SHEETS.nci;
  timelineHeads(nciSheet);
  const bg = (h: number, s: string) => P(`bg/${h}/${s}`);
  if (nodes.length) {
    nci.push(new LRow(P('nci/heading'), 'heading', 'At each node'));
    nci.push(new LRow(P('nci/note'), 'text', "From the node's own consolidation (By group), cumulative from the start; the outside investors' "
      + 'share of capital comes through the investment eliminations.', { indent: 1, role: 'note' }));
    for (const n of nodes) {
      const id = (s: string) => P(`nci/${n.head}/${s}`);
      nci.push(new LRow(P(`nci/${n.head}`), 'subheading', `${n.head} ${n.name}`, { level: 0 }));
      nci.push(new LRow(id('share'), 'series', 'Outside investors\' share', { indent: 2, unit: '%', total: 'none', formula: `=«V|${id('share')}»`,
        cells: { [TOTAL_COL]: `=1-INDEX(Ent_Held,MATCH(${n.head},Ent_Codes,0))` } }));
      for (const [k, label, src] of [['re', 'Retained earnings at the start', bg(n.head, `cons/${RE}`)],
        ['surplus', 'Surplus since the start', bg(n.head, 'surplus')], ['dist', 'Distributions paid', bg(n.head, `cons/${DIST}`)],
        ['na', 'Net assets', bg(n.head, 'na')]] as const) {
        nci.push(new LRow(id(k), 'series', label, { indent: 2, unit: '$', total: 'last', formula: `=«R|${src}»`, cells: { [TOTAL_COL]: `=«V|${src}»` } }));
      }
    }
  }
  const nciRanges = new Set<number>();
  for (const t of tiers) {
    const below = nodesBelow(t);
    if (!below.length) continue;
    nciRanges.add(t.head);
    nci.push(new LRow(P(`nci/l/${t.head}/sp`), 'blank', '', { space: 9, level: 0 }));
    nci.push(new LRow(P(`nci/l/${t.head}`), 'heading', `Lines for the ${t.name}`));
    const ids: string[] = [];
    for (const n of below) {
      const s = `«V|${P(`nci/${n.head}/share`)}»`;
      const v = (k: string, f: 'R' | 'V') => `«${f}|${P(`nci/${n.head}/${k}`)}»`;
      const mk = (k: string, label: string, account: number, f: (x: 'R' | 'V') => string) => {
        const id = P(`nci/l/${t.head}/${n.head}/${k}`);
        ids.push(id);
        nci.push(new LRow(id, 'series', `${n.name}: ${label}`, { indent: 2, unit: '$', total: 'last', formula: `=${f('R')}`,
          cells: { [KE]: n.head, [KA]: account, [TOTAL_COL]: `=${f('V')}` }, colStyles: codeStyles(true, true) }));
      };
      mk('re', 'NCI share of retained earnings at the start', RE, x => `-${s}*${v('re', x)}`);
      mk('pl', 'NCI share of the surplus', NCIPL, x => `${s}*${v('surplus', x)}`);
      mk('dist', 'NCI share of distributions', DIST, x => `-${s}*${v('dist', x)}`);
      mk('nci', 'Non-controlling interests', NCI, x => `${s}*(${v('re', x)}-${v('surplus', x)}+${v('dist', x)})`);
      if (marginRows.length) {
        const up = P(`mg/c/${t.head}/${n.head}`);
        mk('up', 'NCI share of upstream margin still held', NCI, x => `${s}*«${x}|${up}»`);
        mk('uppl', 'NCI share of upstream margin, in the surplus', NCIPL, x => `-${s}*«${x}|${up}»`);
      }
    }
    range(`NCI_${t.head}_Acct`, ids[0], ids[ids.length - 1], KA);
    range(`NCI_${t.head}_Open`, ids[0], ids[ids.length - 1], TOTAL_COL);
    range(`NCI_${t.head}_Val`, ids[0], ids[ids.length - 1], FIRST_PERIOD_COL, 'timeline');
  }

  // ------------------------------------------------------------------ By group (timeline)
  const by: LRow[] = [];
  const bySheet = CONS_SHEETS.byGroup;
  heads[bySheet] = [[5, KA, 'Account']];
  widths[bySheet] = { [KA]: 9 };
  totalHeads[bySheet] = 'Opening';
  const isAccts = chart.filter(a => IS_CLASSES.has(a.class));
  const bsAccts = chart.filter(a => a.class === 'asset' || a.class === 'liability' || a.class === 'equity');
  const cfAccts = (c: 'op' | 'inv' | 'fin') => chart.filter(a => a.cf === c);
  const noncash = chart.filter(a => a.noncash !== undefined);
  // Summary first: each must be nil
  by.push(new LRow(P('bg/sum/heading'), 'heading', 'Summary (each must be nil)'));
  const sums: [string, string, (t: Tier) => [string, string] | null][] = [
    ['elim', 'Eliminations', t => [`«R|${bg(t.head, 'elim/total')}»`, `«V|${bg(t.head, 'elim/total')}»`]],
    ['tb', 'Consolidated trial balance', t => [`«R|${bg(t.head, 'cons/total')}»`, `«V|${bg(t.head, 'cons/total')}»`]],
    ['diff', 'Intercompany differences', t => [`«R|${bg(t.head, `cons/${DIFF}`)}»`, `«V|${bg(t.head, `cons/${DIFF}`)}»`]],
    ['nci', "NCI less the nodes' share", t => {
      const parts = (f: 'R' | 'V') => nodesBelow(t).map(n => `«V|${P(`nci/${n.head}/share`)}»*(«${f}|${P(`nci/${n.head}/na`)}»`
        + (marginRows.length ? `-«${f}|${P(`mg/c/${t.head}/${n.head}`)}»` : '') + ')').join('+') || '0';
      return [`-«R|${bg(t.head, `cons/${NCI}`)}»-(${parts('R')})`, `-«V|${bg(t.head, `cons/${NCI}`)}»-(${parts('V')})`];
    }],
    ['cf', 'Cash flow less the movement in cash', t => [`«R|${bg(t.head, 'cf/diff')}»`, '0']],
  ];
  for (const [key, label, f] of sums) {
    by.push(new LRow(P(`bg/sum/${key}`), 'section', label, { indent: 1 }));
    for (const t of tiers) {
      const [r, v] = f(t)!;
      by.push(new LRow(P(`bg/sum/${key}/${t.head}`), 'series', t.name, { indent: 2, unit: '$', style: 'check', total: 'none',
        formula: `=${round(r)}`, cells: { [TOTAL_COL]: `=${round(v)}` } }));
    }
    range(`BG_${key}`, P(`bg/sum/${key}/${tiers[0].head}`), P(`bg/sum/${key}/${tiers[tiers.length - 1].head}`), TOTAL_COL, 'timeline');
  }
  // The checks the module's check rows read (each a count of what fails)
  by.push(new LRow(P('bg/chk/sp'), 'blank', '', { space: 9, level: 0 }));
  by.push(new LRow(P('bg/chk/heading'), 'heading', 'Checks'));
  const top = tiers[0];
  const count = (rng: string) => `SUMPRODUCT(--(ABS(${rng})>0.001))`;
  const checks: [string, string, string][] = [
    ['Cn_Chk_TB', 'Entity months not netting to nil', `=${count('ED_TB')}`],
    ['Cn_Chk_Tree', 'Entities out of order, too deep, or two tops', '=SUM(Ent_Tree)+SUM(Ent_Depth)+IF(COUNTIF(Ent_Parent,"")<>1,1,0)'],
    ['Cn_Chk_Join', 'Entities with figures before they join', '=SUM(Ent_Join)'],
    ['Cn_Chk_Pairs', 'Register rows and pairs not matching', `=${txs.length ? 'COUNTIF(IC_Status,"Break")' : '0'}+${bals.length ? count('BAL_Diff') : '0'}`],
    ['Cn_Chk_Inv', "Investment months off the parent's share", `=${invs.length ? count('EL_InvChk') : '0'}`],
    ['Cn_Chk_Elim', 'Group months: eliminations not nil', `=${count('BG_elim')}`],
    ['Cn_Chk_Cons', 'Group months: trial balance not nil', `=${count('BG_tb')}+${count('BG_diff')}`],
    ['Cn_Chk_NCI', "Group months: NCI not the nodes' share", `=${count('BG_nci')}`],
    ['Cn_Chk_Top', 'Months the top group holds investments',
      `=${count(`«A|${bg(top.head, `cons/${INVEST}`)}»`)}+IF(ABS(«V|${bg(top.head, `cons/${INVEST}`)}»)>0.001,1,0)`],
    ['Cn_Chk_CF', 'Group months: cash flow not reconciled', `=${count('BG_cf')}`],
    ['Cn_Alert_Planned', 'Planned entities not in Home Hub yet', '=COUNTIF(Ent_Status,"Planned")'],
    ['Cn_Alert_Netting', 'Netting account balances at month ends', bals.some(x => acct.get(x.sellerAcct)?.role === 'netting') ? `=${count('BAL_Netting')}` : '=0'],
  ];
  for (const [nm, label, f] of checks) {
    const id = P(`bg/chk/${nm}`);
    by.push(new LRow(id, 'fixed', label, { indent: 1, unit: '#', role: 'count', cells: { [TOTAL_COL]: f } }));
    names.push([nm, id, TOTAL_COL]);
  }
  // Each group: combined, eliminations, consolidated, then its cash flow
  for (const t of tiers) {
    const h = t.head;
    const nciTerm = (open: boolean, a: number) => nciRanges.has(h)
      ? (open ? `+SUMIFS(NCI_${h}_Open,NCI_${h}_Acct,${a})` : `+SUMIFS(INDEX(NCI_${h}_Val,0,«N»),NCI_${h}_Acct,${a})`) : '';
    by.push(new LRow(P(`bg/${h}/sp`), 'blank', '', { space: 9, level: 0 }));
    by.push(new LRow(bg(h, 'heading'), 'heading', `${t.name} (head ${h})`));
    const part = (key: 'comb' | 'elim' | 'cons', label: string, f: (a: GroupAccount) => { first?: string; formula: string; open: string }) => {
      by.push(new LRow(bg(h, key), 'section', label, { indent: 1 }));
      for (const a of ledger) {
        const x = f(a);
        by.push(new LRow(bg(h, `${key}/${a.code}`), 'series', a.name, { indent: 2, unit: '$', total: 'last', first: x.first ?? null,
          formula: x.formula, cells: { [KA]: a.code, [TOTAL_COL]: x.open }, colStyles: codeStyles(false, true) }));
      }
      const fst = bg(h, `${key}/${ledger[0].code}`), lst = bg(h, `${key}/${ledger[ledger.length - 1].code}`);
      by.push(new LRow(bg(h, `${key}/total`), 'series', 'Total (must be nil)', { indent: 2, unit: '$', style: 'check', total: 'last',
        formula: `=${round(`SUM(«R|${fst}»:«R|${lst}»)`)}`, cells: { [TOTAL_COL]: `=${round(`SUM(«V|${fst}»:«V|${lst}»)`)}` } }));
    };
    part('comb', 'Combined', a => {
      const self = bg(h, `comb/${a.code}`);
      const mv = `SUMIFS(INDEX(ED_Mov,0,«N»),ED_Acct,${a.code},ED_In_${h},1)`;
      return { formula: `=«W|${self}»+${mv}`, open: `=SUMIFS(ED_Open,ED_Acct,${a.code},ED_In_${h},1)` };
    });
    part('elim', 'Eliminations', a => ({
      formula: `=SUMIFS(INDEX(EL_Val,0,«N»),EL_Acct,${a.code},EL_In_${h},1)${nciTerm(false, a.code)}`,
      open: `=SUMIFS(EL_Open,EL_Acct,${a.code},EL_In_${h},1)${nciTerm(true, a.code)}`,
    }));
    part('cons', 'Consolidated', a => ({
      formula: `=«R|${bg(h, `comb/${a.code}`)}»+«R|${bg(h, `elim/${a.code}`)}»`,
      open: `=«V|${bg(h, `comb/${a.code}`)}»+«V|${bg(h, `elim/${a.code}`)}»`,
    }));
    const cons = (a: number, f: 'R' | 'V' | 'W' | 'P' = 'R') => `«${f}|${bg(h, `cons/${a}`)}»`;
    const sumOf = (list: GroupAccount[], f: 'R' | 'V') => list.length ? list.map(a => cons(a.code, f)).join('+') : '0';
    const results: [string, string, string, string, string][] = [
      ['surplus', 'Surplus since the start', `=-(${sumOf(isAccts, 'R')})`, `=-(${sumOf(isAccts, 'V')})`, 'last'],
      ['nci_pl', 'Of which attributable to non-controlling interests', `=${cons(NCIPL)}`, `=${cons(NCIPL, 'V')}`, 'last'],
      ['owners', 'Of which attributable to the owners', `=«R|${bg(h, 'surplus')}»-«R|${bg(h, 'nci_pl')}»`,
        `=«V|${bg(h, 'surplus')}»-«V|${bg(h, 'nci_pl')}»`, 'last'],
      ['retained', "Owners' retained earnings", `=-(${cons(RE)}+${cons(DIST)}+${sumOf(isAccts, 'R')}+${cons(NCIPL)})`,
        `=-(${cons(RE, 'V')}+${cons(DIST, 'V')}+${sumOf(isAccts, 'V')}+${cons(NCIPL, 'V')})`, 'last'],
      ['na', 'Net assets', `=${sumOf(bsAccts.filter(a => a.class !== 'equity'), 'R')}`, `=${sumOf(bsAccts.filter(a => a.class !== 'equity'), 'V')}`, 'last'],
    ];
    by.push(new LRow(bg(h, 'results'), 'section', 'Results', { indent: 1 }));
    for (const [k, label, f, open, total] of results) {
      by.push(new LRow(bg(h, k), 'series', label, { indent: 2, unit: '$', total, formula: f, cells: { [TOTAL_COL]: open } }));
    }
    // Cash flow, from the consolidated balance sheet's movements (indirect)
    by.push(new LRow(bg(h, 'cf'), 'section', 'Cash flow for the month', { indent: 1 }));
    const mv = (a: number) => `(${cons(a)}-${cons(a, 'W')})`;
    const cfLine = (k: string, label: string, f: string) => {
      by.push(new LRow(bg(h, `cf/${k}`), 'series', label, { indent: 2, unit: '$', total: 'sum', formula: f }));
      return bg(h, `cf/${k}`);
    };
    const opIds = [cfLine('surplus', 'Surplus for the month', `=-(${isAccts.map(a => mv(a.code)).join('+') || '0'})`)];
    for (const a of noncash) opIds.push(cfLine(`add/${a.code}`, `Add back ${lower(a.name)}`, `=${mv(a.code)}`));
    for (const a of cfAccts('op')) opIds.push(cfLine(`op/${a.code}`, `Movement in ${lower(a.name)}`, `=-${mv(a.code)}`));
    const sumIds = (ids: string[]) => `=SUM(«R|${ids[0]}»:«R|${ids[ids.length - 1]}»)`;
    const opNet = cfLine('op', 'Net cash from operating activities', sumIds(opIds));
    const invIds = cfAccts('inv').map(a => {
      const pair = noncash.filter(x => x.noncash === a.code).map(x => `+${mv(x.code)}`).join('');
      return cfLine(`inv/${a.code}`, a.name, `=-(${mv(a.code)}${pair})`);
    });
    const invNet = cfLine('inv', 'Net cash from investing activities', invIds.length ? sumIds(invIds) : '=0');
    const finIds = cfAccts('fin').map(a => cfLine(`fin/${a.code}`, a.name,
      a.code === NCI ? `=-(${mv(NCI)}+${mv(NCIPL)})` : `=-${mv(a.code)}`));
    const finNet = cfLine('fin', 'Net cash from financing activities', finIds.length ? sumIds(finIds) : '=0');
    const net = cfLine('net', 'Net cash flow', `=«R|${opNet}»+«R|${invNet}»+«R|${finNet}»`);
    by.push(new LRow(bg(h, 'cf/diff'), 'series', 'Net cash flow less the movement in cash (must be nil)', { indent: 2, unit: '$',
      style: 'check', total: 'sum', formula: `=${round(`«R|${net}»-${mv(CASH)}`)}` }));
  }
  const groupList: ListSpec = { name: 'List_Groups', title: 'Groups that consolidate', group: 'Group consolidation',
    items: tiers.map(t => ({ value: t.name, style: 'lu.text' })) };
  const typeList: ListSpec = { name: 'List_IC_Types', title: 'Intercompany types', group: 'Group consolidation',
    items: IC_TYPES.map(value => ({ value, style: 'lu.text' })) };

  // ------------------------------------------------------------------ Group statements (the module's own sheet)
  const main: LRow[] = [];
  const choose = (key: (h: number) => string, f: 'R' | 'P' | 'V' = 'R') => `CHOOSE(${sel},${tiers.map(t => `«${f}|${key(t.head)}»`).join(',')})`;
  const groupLabel = (id: string, label: string) => main.push(new LRow(P(`gs/${id}`), 'section', label, { indent: 1, role: 'group' }));
  const lineRows = (prefix: string, accts: GroupAccount[], f: (a: GroupAccount) => string, total: string) => accts.map(a => {
    const id = P(`gs/${prefix}/${a.code}`);
    main.push(new LRow(id, 'series', a.name, { indent: 2, unit: '$', total, formula: f(a) }));
    return id;
  });
  const totalRow = (id: string, label: string, f: string, total: string, style = 'total') => {
    main.push(new LRow(P(`gs/${id}`), 'series', label, { indent: 1, unit: '$', style, total, formula: f }));
    return P(`gs/${id}`);
  };
  const sumRows = (ids: string[]) => (ids.length ? `=SUM(«R|${ids[0]}»:«R|${ids[ids.length - 1]}»)` : '=0');
  const R = (id: string) => `«R|${id}»`;
  main.push(new LRow(P('gs/note'), 'text', 'The group chosen above: the income statement and cash flow for each month (the Total column adds '
    + 'the months up), the balance sheet at each month end.', { indent: 1, role: 'note' }));
  // Income statement
  main.push(new LRow(P('gs/is'), 'heading', 'Income statement'));
  const isLine = (a: GroupAccount) => `=${a.class === 'income' ? '-' : ''}(${choose(h => bg(h, `cons/${a.code}`))}-${choose(h => bg(h, `cons/${a.code}`), 'P')})`;
  const sec = (line: string, label: string, totalLabel: string | null) => {
    const accts = isAccts.filter(a => a.line === line);
    if (!accts.length) return null;
    if (accts.length === 1) {   // one account: its line alone, without a group label over it
      const id = P(`gs/is/${line}/${accts[0].code}`);
      main.push(new LRow(id, 'series', accts[0].name, { indent: 1, unit: '$', total: 'sum', formula: isLine(accts[0]) }));
      return id;
    }
    groupLabel(`is/${line}`, label);
    const ids = lineRows(`is/${line}`, accts, isLine, 'sum');
    return totalLabel ? totalRow(`is/${line}/total`, totalLabel, sumRows(ids), 'sum', '') : ids[ids.length - 1];
  };
  const z = (id: string | null) => (id ? R(id) : '0');
  const rev = sec('revenue', 'Revenue', 'Total revenue');
  const cogs = sec('cogs', 'Cost of sales', 'Total cost of sales');
  const gross = totalRow('is/gross', 'Gross margin', `=${z(rev)}-${z(cogs)}`, 'sum');
  const oi = sec('other_income', 'Other income', 'Total other income');
  const staff = sec('staff', 'Staff costs', 'Total staff costs');
  const opex = sec('opex', 'Operating expenses', 'Total operating expenses');
  const oe = sec('other_expense', 'Other expenses', 'Total other expenses');
  const ebitda = totalRow('is/ebitda', 'EBITDA', `=${R(gross)}+${z(oi)}-${z(staff)}-${z(opex)}-${z(oe)}`, 'sum');
  const da = sec('da', 'Depreciation and amortisation', 'Total depreciation and amortisation');
  const ebit = totalRow('is/ebit', 'EBIT', `=${R(ebitda)}-${z(da)}`, 'sum');
  const ii = sec('interest_income', 'Interest income', 'Total interest income');
  const ie = sec('interest_expense', 'Interest expense', 'Total interest expense');
  const npbt = totalRow('is/npbt', 'Surplus before tax', `=${R(ebit)}+${z(ii)}-${z(ie)}`, 'sum');
  const tax = sec('tax', 'Income tax', null);
  const npat = totalRow('is/npat', 'Surplus for the month', `=${R(npbt)}-${z(tax)}`, 'sum');
  const nciPl = P('gs/is/nci_pl');
  main.push(new LRow(nciPl, 'series', 'Attributable to non-controlling interests', { indent: 2, unit: '$', total: 'sum',
    formula: `=${choose(h => bg(h, `cons/${NCIPL}`))}-${choose(h => bg(h, `cons/${NCIPL}`), 'P')}` }));
  main.push(new LRow(P('gs/is/owners'), 'series', 'Attributable to the owners', { indent: 2, unit: '$', total: 'sum', formula: `=${R(npat)}-${R(nciPl)}` }));
  // Balance sheet
  main.push(new LRow(P('gs/bs/sp'), 'blank', '', { space: 9, level: 0 }));
  main.push(new LRow(P('gs/bs'), 'heading', 'Balance sheet'));
  const bsLine = (sign: 1 | -1) => (a: GroupAccount) => `=${sign < 0 ? '-' : ''}${choose(h => bg(h, `cons/${a.code}`))}`;
  const bsSec = (line: string, label: string, totalLabel: string, sign: 1 | -1) => {
    const accts = bsAccts.filter(a => a.line === line);
    groupLabel(`bs/${line}`, label);
    const ids = lineRows(`bs/${line}`, accts, bsLine(sign), 'last');
    return totalRow(`bs/${line}/total`, totalLabel, sumRows(ids), 'last', '');
  };
  const ca = bsSec('ca', 'Current assets', 'Total current assets', 1);
  const nca = bsSec('nca', 'Non-current assets', 'Total non-current assets', 1);
  const assets = totalRow('bs/assets', 'Total assets', `=${R(ca)}+${R(nca)}`, 'last');
  const cl = bsSec('cl', 'Current liabilities', 'Total current liabilities', -1);
  const ncl = bsSec('ncl', 'Non-current liabilities', 'Total non-current liabilities', -1);
  const liab = totalRow('bs/liabilities', 'Total liabilities', `=${R(cl)}+${R(ncl)}`, 'last');
  const netAssets = totalRow('bs/na', 'Net assets', `=${R(assets)}-${R(liab)}`, 'last');
  groupLabel('bs/equity', 'Equity');
  const eqAccts = bsAccts.filter(a => a.line === 'equity' && ![RE, DIST, NCI].includes(a.code));
  const eqIds = lineRows('bs/equity', eqAccts, bsLine(-1), 'last');
  const ret = P('gs/bs/retained');
  main.push(new LRow(ret, 'series', 'Retained earnings', { indent: 2, unit: '$', total: 'last', formula: `=${choose(h => bg(h, 'retained'))}` }));
  const nciRow = P('gs/bs/nci');
  main.push(new LRow(nciRow, 'series', 'Non-controlling interests', { indent: 2, unit: '$', total: 'last', formula: `=-${choose(h => bg(h, `cons/${NCI}`))}` }));
  const equity = totalRow('bs/equity/total', 'Total equity', `=${eqIds.map(R).join('+') || '0'}+${R(ret)}+${R(nciRow)}`, 'last');
  main.push(new LRow(P('gs/bs/check'), 'series', 'Net assets less equity (must be nil)', { indent: 1, unit: '$', style: 'check', total: 'last',
    formula: `=${round(`${R(netAssets)}-${R(equity)}`)}` }));
  // Cash flow statement
  main.push(new LRow(P('gs/cf/sp'), 'blank', '', { space: 9, level: 0 }));
  main.push(new LRow(P('gs/cf'), 'heading', 'Cash flow statement'));
  const cfRow = (k: string, label: string, style = '', indent = 2) => {
    const id = P(`gs/cf/${k}`);
    main.push(new LRow(id, 'series', label, { indent, unit: '$', style, total: 'sum', formula: `=${choose(h => bg(h, `cf/${k}`))}` }));
    return id;
  };
  groupLabel('cf/g/op', 'Operating activities');
  cfRow('surplus', 'Surplus for the month');
  for (const a of noncash) cfRow(`add/${a.code}`, `Add back ${lower(a.name)}`);
  for (const a of cfAccts('op')) cfRow(`op/${a.code}`, `Movement in ${lower(a.name)}`);
  cfRow('op', 'Net cash from operating activities', 'total', 1);
  groupLabel('cf/g/inv', 'Investing activities');
  for (const a of cfAccts('inv')) cfRow(`inv/${a.code}`, a.name);
  cfRow('inv', 'Net cash from investing activities', 'total', 1);
  groupLabel('cf/g/fin', 'Financing activities');
  for (const a of cfAccts('fin')) cfRow(`fin/${a.code}`, a.name);
  cfRow('fin', 'Net cash from financing activities', 'total', 1);
  const netCf = cfRow('net', 'Net cash flow', 'total', 1);
  const cashStart = P('gs/cf/start');
  main.push(new LRow(cashStart, 'series', 'Cash at the start of the month', { indent: 1, unit: '$', total: 'none',
    formula: `=${choose(h => bg(h, `cons/${CASH}`), 'P')}`, first: `=${choose(h => bg(h, `cons/${CASH}`), 'V')}`,
    cells: { [TOTAL_COL]: `=${choose(h => bg(h, `cons/${CASH}`), 'V')}` } }));
  const cashEnd = P('gs/cf/end');
  main.push(new LRow(cashEnd, 'series', 'Cash at the end of the month', { indent: 1, unit: '$', style: 'total', total: 'last',
    formula: `=${R(cashStart)}+${R(netCf)}` }));
  const cashBs = P(`gs/bs/ca/${CASH}`);
  main.push(new LRow(P('gs/cf/check'), 'series', 'Cash at the end less cash on the balance sheet (must be nil)', { indent: 1, unit: '$',
    style: 'check', total: 'last', formula: `=${round(`${R(cashEnd)}-${R(cashBs)}`)}` }));

  return {
    main,
    sheets: [[CONS_SHEETS.byGroup, by], [CONS_SHEETS.entities, ent], [CONS_SHEETS.data, ed], [CONS_SHEETS.ic, ic],
      [CONS_SHEETS.bal, bal], [CONS_SHEETS.sites, st], [CONS_SHEETS.margin, mg], [CONS_SHEETS.elim, el], [CONS_SHEETS.nci, nci]],
    kinds: { [CONS_SHEETS.entities]: 'table', [CONS_SHEETS.ic]: 'table' },
    widths, heads, totalHeads, names, ranges, lists: [groupList, typeList],
  };
}
