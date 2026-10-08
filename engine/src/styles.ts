// HFG's style catalogue: named cell styles on theme slots only, so a model takes its entity's look
// from the theme and rebrands with it. Names are HFG's own ("HFG Heading 1", "HFG Input Number").
// Body text is 9pt Segoe UI in Text 1 lighter 25% (#404040), vertically centred; headings 10pt
// where they are titles, the sheet title 11pt. A cell's format is a style plus optional modifiers
// (a total's rule above, a list's dashed rule beneath, bold, italic), written as cell formats that
// keep the style's name.

import type { Brand } from './model.ts';
import { barTextSlot, SLOT, THEMES, TINT } from './theme.ts';

export type Color = { theme: number; tint?: number } | { rgb: string };

export interface FontSpec {
  size: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color: Color;
  /** A font other than the theme's (the navigation symbols). */
  name?: string;
}

export type BorderStyle = 'thin' | 'dashed' | 'hair';
export interface BorderSide { style: BorderStyle; color: Color }
export interface BorderSpec { left?: BorderSide; right?: BorderSide; top?: BorderSide; bottom?: BorderSide }

export interface StyleSpec {
  name: string;
  font: FontSpec;
  fill?: Color;
  border?: BorderSpec;
  numFmt?: string;
  h?: 'left' | 'center' | 'right';
  wrap?: boolean;
  /** Inputs and cell links are unlocked, so a protected model still takes them. */
  unlocked?: boolean;
}

export type Modifier = 'total' | 'last' | 'bold' | 'italic' | 'band';

const MOD_WORDS: Record<Modifier, string> = { total: 'Total', last: 'Last Item', bold: 'Bold', italic: 'Italic', band: 'Band' };

/**
 * A style with modifiers is a named style of its own ("HFG Number Total", "HFG Number Last Item"),
 * so the live writer sets formats by style name only: Office.js cannot set theme colours, and a
 * named style keeps the rule's colour on the theme.
 */
export function styleName(base: string, mods: readonly Modifier[] = []): string {
  const m = [...new Set(mods)].sort();
  return m.length ? `${base} ${m.map(x => MOD_WORDS[x]).join(' ')}` : base;
}

const TEXT: Color = { theme: SLOT.dk1, tint: TINT.lighter25 };
const MUTED: Color = { theme: SLOT.dk1, tint: TINT.lighter50 };
const LINK: Color = { theme: SLOT.hlink };
const DK2: Color = { theme: SLOT.dk2 };
const WHITE: Color = { theme: SLOT.lt1 };

/** Number formats: one decimal, negatives in brackets, zero as a dash, padded so figures line up. */
export const NUMBER_FORMATS = {
  num: '_(#,##0.0_);(#,##0.0);_("-"_);_(@_)',
  int: '_(#,##0_);(#,##0);_("-"_);_(@_)',
  pct: '_(0.0%_);(0.0%);_("-"_);_(@_)',
  mult: '_(0.00"x"_);(0.00"x");_("-"_);_(@_)',
  flag: '_(0_);(0);_("-"_);_(@_)',
  signed: '_(+#,##0_);_(-#,##0_);_("-"_);_(@_)',
  date: 'd mmm yyyy',
  month: 'mmm yy',
  monthYear: 'mmm yyyy',
  year: '0',
  count: '0',
  text: '@',
} as const;

/** Which number format a unit takes. */
const UNIT_FORMATS: Record<string, keyof typeof NUMBER_FORMATS> = {
  '$': 'int', '%': 'pct', flag: 'flag', '#': 'int', months: 'int', days: 'int', period: 'int', date: 'date', x: 'mult', year: 'year',
  text: 'text', switch: 'text',
};

export function formatForUnit(unit: string): keyof typeof NUMBER_FORMATS {
  return Object.hasOwn(UNIT_FORMATS, unit) ? UNIT_FORMATS[unit] : 'num';
}

const body = (extra: Partial<FontSpec> = {}): FontSpec => ({ size: 9, color: TEXT, ...extra });
const grid = (): BorderSpec => {
  const side: BorderSide = { style: 'thin', color: TEXT };
  return { left: side, right: side, top: side, bottom: side };
};
const inputBorder = (): BorderSpec => {
  const side: BorderSide = { style: 'thin', color: { theme: SLOT.accent1, tint: TINT.lighter40 } };
  return { left: side, right: side, top: side, bottom: side };
};

/** The catalogue, keyed by the engine's short names. */
export function catalogue(brand: Brand): Record<string, StyleSpec> {
  const barText: Color = { theme: barTextSlot(THEMES[brand]) };
  const input = (name: string, fmt: keyof typeof NUMBER_FORMATS, h?: StyleSpec['h']): StyleSpec => ({
    name, font: body(), fill: { theme: SLOT.accent1, tint: TINT.lighter80 }, border: inputBorder(),
    numFmt: NUMBER_FORMATS[fmt], unlocked: true, h,
  });
  const calc = (name: string, fmt: keyof typeof NUMBER_FORMATS, h?: StyleSpec['h']): StyleSpec => ({
    name, font: body(), numFmt: NUMBER_FORMATS[fmt], h,
  });
  return {
    normal: { name: 'Normal', font: body() },
    title: { name: 'HFG Sheet Title', font: body({ size: 11, bold: true }) },
    modelName: { name: 'HFG Model Name', font: body({ size: 10, bold: true }) },
    entity: { name: 'HFG Entity', font: body({ color: MUTED }) },
    h1: { name: 'HFG Heading 1', font: body({ bold: true, color: barText }), fill: { theme: SLOT.accent1 } },
    h2: {
      name: 'HFG Heading 2', font: body({ bold: true, color: DK2 }), fill: { theme: SLOT.lt1, tint: TINT.darker5 },
      border: { bottom: { style: 'thin', color: DK2 } },
    },
    h3: { name: 'HFG Heading 3', font: body({ bold: true }) },
    label: { name: 'HFG Label', font: body() },
    muted: { name: 'HFG Background Label', font: body({ color: { theme: SLOT.lt1, tint: TINT.darker25 } }) },
    period: { name: 'HFG Period', font: body({ color: WHITE }), fill: { theme: SLOT.dk1, tint: TINT.lighter25 }, numFmt: NUMBER_FORMATS.month, h: 'right' },
    periodLabel: { name: 'HFG Period Label', font: body({ color: WHITE }), fill: { theme: SLOT.dk1, tint: TINT.lighter25 } },
    period2: { name: 'HFG Period 2', font: body(), fill: { theme: SLOT.lt1, tint: TINT.darker15 }, h: 'right' },
    period2Label: { name: 'HFG Period 2 Label', font: body(), fill: { theme: SLOT.lt1, tint: TINT.darker15 } },
    unit: { name: 'HFG Unit', font: body(), h: 'center' },
    num: calc('HFG Number', 'num'),
    int: calc('HFG Whole Number', 'int'),
    pct: calc('HFG Percent', 'pct'),
    mult: calc('HFG Multiple', 'mult'),
    flag: calc('HFG Flag', 'flag'),
    date: calc('HFG Date', 'date', 'right'),
    month: calc('HFG Month', 'month', 'right'),
    year: calc('HFG Year', 'year', 'right'),
    count: calc('HFG Count', 'count', 'right'),
    text: calc('HFG Text', 'text'),
    key: calc('HFG Key', 'text', 'right'),
    'in.num': input('HFG Input Number', 'num'),
    'in.int': input('HFG Input Whole Number', 'int'),
    'in.pct': input('HFG Input Percent', 'pct'),
    'in.mult': input('HFG Input Multiple', 'mult'),
    'in.flag': input('HFG Input Flag', 'flag'),
    'in.date': input('HFG Input Date', 'date', 'right'),
    'in.month': input('HFG Input Month', 'month', 'right'),
    'in.year': input('HFG Input Year', 'year', 'right'),
    'in.text': input('HFG Input Text', 'text', 'left'),
    'in.count': input('HFG Input Count', 'count', 'right'),
    'in.switch': { ...input('HFG Input Switch', 'text', 'center'), numFmt: undefined },
    monthYear: calc('HFG Month Year', 'monthYear', 'right'),
    cellLink: { name: 'HFG Cell Link', font: body({ color: WHITE }), h: 'center', unlocked: true },   // hidden behind its control
    luHead: { name: 'HFG Lookup Heading', font: body({ bold: true }), fill: { theme: SLOT.lt1, tint: TINT.darker5 }, border: grid() },
    'lu.text': { name: 'HFG Lookup Text', font: body(), border: grid(), numFmt: NUMBER_FORMATS.text },
    'lu.monthYear': { name: 'HFG Lookup Month', font: body(), border: grid(), numFmt: NUMBER_FORMATS.monthYear, h: 'left' },
    'lu.int': { name: 'HFG Lookup Number', font: body(), border: grid(), numFmt: NUMBER_FORMATS.int },
    check: { name: 'HFG Check', font: body(), numFmt: NUMBER_FORMATS.flag },
    link: { name: 'HFG Link', font: body({ color: LINK }) },
    linkU: { name: 'HFG Link Underlined', font: body({ color: LINK, underline: true }) },
    toc1: { name: 'HFG Contents 1', font: body({ color: LINK, bold: true, underline: true }) },
    toc2: { name: 'HFG Contents 2', font: body({ color: LINK, underline: true }) },
    toc3: { name: 'HFG Contents 3', font: body({ color: LINK }) },
    nav: { name: 'HFG Navigation', font: body({ color: LINK, name: 'Segoe UI Symbol' }), h: 'center' },
    sectionNo: { name: 'HFG Section Number', font: body({ size: 10, bold: true }) },
    note: { name: 'HFG Note', font: body() },
    colHead: { name: 'HFG Column Heading', font: body({ bold: true }), numFmt: NUMBER_FORMATS.text, h: 'right' },
    signed: calc('HFG Signed Number', 'signed'),
    mutedNum: { name: 'HFG Background Number', font: body({ color: { theme: SLOT.lt1, tint: TINT.darker25 } }), numFmt: NUMBER_FORMATS.int, h: 'right' },
    logoTile: { name: 'HFG Logo Tile', font: body({ color: WHITE }), fill: DK2 },
    // The Scenarios sheet's band: each scenario's number and name on the dark band, the active
    // column's tab in the entity's colour, and a marker under the active scenario.
    scnBand: { name: 'HFG Scenario Band', font: body({ color: WHITE }), fill: { theme: SLOT.dk1, tint: TINT.lighter25 } },
    scnBandHead: { name: 'HFG Scenario Band Heading', font: body({ color: WHITE }), fill: { theme: SLOT.dk1, tint: TINT.lighter25 }, h: 'center' },
    scnBandName: { name: 'HFG Scenario Band Name', font: body({ color: WHITE, italic: true }), fill: { theme: SLOT.dk1, tint: TINT.lighter25 }, h: 'center' },
    scnActive: { name: 'HFG Scenario Active', font: body({ color: barText }), fill: { theme: SLOT.accent1 }, h: 'center' },
    scnMarker: { name: 'HFG Scenario Marker', font: body({ color: { theme: SLOT.lt1, tint: TINT.darker25 }, name: 'Wingdings 3' }), h: 'center' },
    // Dashboard tables: a block's heading over its columns, and the actual or forecast line under the months.
    blockHead: {
      name: 'HFG Block Heading', font: body({ bold: true, color: DK2 }), fill: { theme: SLOT.lt1, tint: TINT.darker5 },
      border: { bottom: { style: 'thin', color: DK2 } }, h: 'center',
    },
    colSub: { name: 'HFG Column Subheading', font: body({ color: MUTED }), numFmt: NUMBER_FORMATS.text, h: 'right' },
    // The value a scenario line uses: boxed, as it is what the model reads.
    selPct: { name: 'HFG Selected Percent', font: body(), border: grid(), numFmt: NUMBER_FORMATS.pct },
    selText: { name: 'HFG Selected Text', font: body(), border: grid(), numFmt: NUMBER_FORMATS.text, h: 'center' },
  };
}

/** The bold red a check that is not clear turns (conditional format). */
export const CHECK_RED = 'FFCB2840';

/** Cell formats: each style plus modifiers becomes one entry of styles.xml's cellXfs. */
export class StyleBook {
  readonly brand: Brand;
  readonly styles: Record<string, StyleSpec>;
  private keys: string[];
  private fonts: string[] = [];
  private fills: string[] = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  private borders: string[] = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  private numFmts = new Map<string, number>();
  private styleXfs: string[] = [];
  private cellXfs: string[] = [];
  private xfIndex = new Map<string, number>();
  private dxfs: string[] = [];

  constructor(brand: Brand) {
    this.brand = brand;
    this.styles = catalogue(brand);
    this.keys = Object.keys(this.styles);
    for (const key of this.keys) this.styleXfs.push(this.xfXml(this.styles[key], null));
    this.xf('normal');   // cell format 0 is Normal
  }

  /** The cell format index for a style with modifiers; a combination is registered as a named style the first time. */
  xf(style: string, mods: Modifier[] = []): number {
    const spec = this.styles[style];
    if (!spec) throw new Error(`no style ${style} in the catalogue`);
    const m = [...new Set(mods)].sort();
    const key = m.length ? `${style}+${m.join('+')}` : style;
    const known = this.xfIndex.get(key);
    if (known !== undefined) return known;
    if (m.length && !this.styles[key]) {
      const eff: StyleSpec = { ...spec, name: styleName(spec.name, m), font: { ...spec.font }, border: { ...(spec.border || {}) } };
      for (const mod of m) {
        if (mod === 'bold' || mod === 'total') eff.font.bold = true;
        if (mod === 'italic') eff.font.italic = true;
        if (mod === 'total') eff.border!.top = { style: 'thin', color: TEXT };
        if (mod === 'last') eff.border!.bottom = { style: 'dashed', color: MUTED };
      }
      this.styles[key] = eff;
      this.keys.push(key);
      this.styleXfs.push(this.xfXml(eff, null));
    }
    const idx = this.cellXfs.length;
    this.cellXfs.push(this.xfXml(this.styles[key], this.keys.indexOf(key)));
    this.xfIndex.set(key, idx);
    return idx;
  }

  /** A differential format for conditional formatting; returns its index. */
  dxf(xml: string): number {
    const i = this.dxfs.indexOf(xml);
    if (i >= 0) return i;
    this.dxfs.push(xml);
    return this.dxfs.length - 1;
  }

  private color(c: Color): string {
    if ('rgb' in c) return `<color rgb="${c.rgb}"/>`;
    return `<color theme="${c.theme}"${c.tint ? ` tint="${c.tint}"` : ''}/>`;
  }

  private add(list: string[], xml: string): number {
    const i = list.indexOf(xml);
    if (i >= 0) return i;
    list.push(xml);
    return list.length - 1;
  }

  private fontId(f: FontSpec): number {
    const xml = '<font>' + (f.bold ? '<b/>' : '') + (f.italic ? '<i/>' : '') + (f.underline ? '<u/>' : '')
      + `<sz val="${f.size}"/>${this.color(f.color)}<name val="${f.name || 'Segoe UI'}"/>`
      + (f.name?.startsWith('Wingdings') ? '<family val="1"/><charset val="2"/>' : '<family val="2"/>')   // a symbol font
      + (f.name ? '' : '<scheme val="minor"/>') + '</font>';
    return this.add(this.fonts, xml);
  }

  private fillId(c?: Color): number {
    if (!c) return 0;
    return this.add(this.fills, `<fill><patternFill patternType="solid"><fgColor${this.color(c).slice(6)}<bgColor indexed="64"/></patternFill></fill>`);
  }

  private borderId(b?: BorderSpec): number {
    if (!b || !Object.keys(b).length) return 0;
    const side = (tag: string, s?: BorderSide) => (s ? `<${tag} style="${s.style}">${this.color(s.color)}</${tag}>` : `<${tag}/>`);
    return this.add(this.borders, `<border>${side('left', b.left)}${side('right', b.right)}${side('top', b.top)}${side('bottom', b.bottom)}<diagonal/></border>`);
  }

  private numFmtId(code?: string): number {
    if (!code) return 0;
    if (code === '@') return 49;
    let id = this.numFmts.get(code);
    if (id === undefined) {
      id = 164 + this.numFmts.size;
      this.numFmts.set(code, id);
    }
    return id;
  }

  private xfXml(s: StyleSpec, styleIndex: number | null): string {
    const font = this.fontId(s.font);
    const fill = this.fillId(s.fill);
    const border = this.borderId(s.border);
    const num = this.numFmtId(s.numFmt);
    const align = `<alignment${s.h ? ` horizontal="${s.h}"` : ''} vertical="center"${s.wrap ? ' wrapText="1"' : ''}/>`;
    const prot = s.unlocked ? '<protection locked="0"/>' : '';
    const apply = styleIndex === null ? '' : ` xfId="${styleIndex}" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1"`;
    return `<xf numFmtId="${num}" fontId="${font}" fillId="${fill}" borderId="${border}"${apply}>${align}${prot}</xf>`;
  }

  /** styles.xml. Call after every cell format the workbook uses has been asked for. */
  xml(): string {
    const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
    const numFmts = [...this.numFmts].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`);
    const cellStyles = this.keys.map((k, i) => (k === 'normal'
      ? `<cellStyle name="Normal" xfId="${i}" builtinId="0"/>`
      : `<cellStyle name="${esc(this.styles[k].name)}" xfId="${i}"/>`));
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + (numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.join('')}</numFmts>` : '')
      + `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>`
      + `<fills count="${this.fills.length}">${this.fills.join('')}</fills>`
      + `<borders count="${this.borders.length}">${this.borders.join('')}</borders>`
      + `<cellStyleXfs count="${this.styleXfs.length}">${this.styleXfs.join('')}</cellStyleXfs>`
      + `<cellXfs count="${this.cellXfs.length}">${this.cellXfs.join('')}</cellXfs>`
      + `<cellStyles count="${cellStyles.length}">${cellStyles.join('')}</cellStyles>`
      + `<dxfs count="${this.dxfs.length}">${this.dxfs.join('')}</dxfs>`
      + '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
      + '</styleSheet>';
  }
}
