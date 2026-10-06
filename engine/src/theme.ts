// Entity themes. A model takes its look from its entity's theme: accent 1 is the entity's
// signature colour (section bars, input fills and borders, the first chart series), accents 2 to 6
// its chart sequence, the hyperlink colour its link colour and Segoe UI both theme fonts. Styles and
// charts refer to theme slots only, so no colour is typed into a model. Colours come from the HFG
// brand palettes (December 2025 master templates).

import type { Brand } from './model.ts';

export interface BrandTheme {
  brand: Brand;
  name: string;
  /** Dark 2: headings on light fills, the logo tile and the text on a light section bar. */
  dk2: string;
  /** Light 2: a subtle alternate fill. */
  lt2: string;
  /** Accent 1 (signature) to accent 6, in chart order. */
  accents: [string, string, string, string, string, string];
  hlink: string;
  /** The logo file in assets/logos, and whether it is a white mark that needs a dark tile behind it. */
  logo: { file: string; onDark: boolean };
}

export const THEMES: Record<Brand, BrandTheme> = {
  HF: {
    brand: 'HF', name: 'Home Foundation', dk2: '09122C', lt2: 'D7DAE1',
    accents: ['679DB5', '09122C', '566A89', '90B6C8', '3C4E60', 'CEDCE5'], hlink: '3C4E60',
    logo: { file: 'HF_white.png', onDark: true },
  },
  HCP: {
    brand: 'HCP', name: 'Home Capital Partners', dk2: '0A132E', lt2: 'E8E8E8',
    accents: ['575F46', '415166', 'B4B7AC', 'D9E6C5', 'B5BBC4', '0A132E'], hlink: '415166',
    logo: { file: 'HCP_olive.png', onDark: false },
  },
  HCL: {
    brand: 'HCL', name: 'Home Construction Limited', dk2: '0A132E', lt2: 'E8E8E8',
    accents: ['20376C', '7EA5EA', '415166', 'CFDBF5', 'B5BBC4', '0A132E'], hlink: '20376C',
    logo: { file: 'HCL_white.png', onDark: true },
  },
  KM: {
    brand: 'KM', name: 'Kāinga Maha', dk2: '09122C', lt2: 'E2E4E7',
    accents: ['829795', '415166', '9EABAA', 'C8DDDB', 'B5BBC4', '3D4256'], hlink: '415166',
    logo: { file: 'KM_navy.png', onDark: false },
  },
  TWK: {
    brand: 'TWK', name: 'Te Wawata Kāinga', dk2: '0A132E', lt2: 'E8E8E8',
    accents: ['FFCB05', '0A132E', '415166', 'E7E0D0', 'B5BBC4', 'FBE8AC'], hlink: '415166',
    logo: { file: 'TWK_navy.png', onDark: false },
  },
};

/** Excel's theme colour indexes in cell formats (light 1 and dark 1 come first, swapped). */
export const SLOT = { lt1: 0, dk1: 1, lt2: 2, dk2: 3, accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9, hlink: 10 } as const;

/** Tints Excel uses for "lighter 80%", "darker 5%" and so on. */
export const TINT = {
  lighter80: 0.7999816888943144,
  lighter60: 0.5999938962981048,
  lighter40: 0.3999755851924192,
  lighter25: 0.249977111117893,
  lighter50: 0.499984740745262,
  darker5: -0.0499893185216834,
  darker15: -0.1499984740745262,
  darker25: -0.249977111117893,
} as const;

function luminance(hex: string): number {
  const ch = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16) / 255)
    .map(c => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
  return 0.2126 * ch[0] + 0.7152 * ch[1] + 0.0722 * ch[2];
}

/** WCAG contrast ratio between two colours given as six hex digits. */
export function contrast(a: string, b: string): number {
  const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x);
  return (hi + 0.05) / (lo + 0.05);
}

/** The text slot for a section bar filled with accent 1: white where it reads (4.5:1), otherwise dark 2. */
export function barTextSlot(t: BrandTheme): number {
  return contrast('FFFFFF', t.accents[0]) >= 4.5 ? SLOT.lt1 : SLOT.dk2;
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

/**
 * The theme part. Dark 1 stays black so the body text (Text 1 lighter 25%) is #404040, HFG's text
 * colour; light 1 stays white so sheets keep a white canvas.
 */
export function themeXml(brand: Brand): string {
  const t = THEMES[brand];
  const srgb = (tag: string, hex: string) => `<a:${tag}><a:srgbClr val="${hex}"/></a:${tag}>`;
  const font = (tag: string) => `<a:${tag}><a:latin typeface="Segoe UI"/><a:ea typeface=""/><a:cs typeface=""/></a:${tag}>`;
  const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  const line = (w: number) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>`;
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${esc(`HFG ${t.brand}`)}"><a:themeElements>`
    + `<a:clrScheme name="${esc(`HFG ${t.brand}`)}">${srgb('dk1', '000000')}${srgb('lt1', 'FFFFFF')}${srgb('dk2', t.dk2)}${srgb('lt2', t.lt2)}`
    + t.accents.map((c, i) => srgb(`accent${i + 1}`, c)).join('')
    + `${srgb('hlink', t.hlink)}${srgb('folHlink', t.hlink)}</a:clrScheme>`
    + `<a:fontScheme name="HFG">${font('majorFont')}${font('minorFont')}</a:fontScheme>`
    + `<a:fmtScheme name="HFG"><a:fillStyleLst>${solid}${solid}${solid}</a:fillStyleLst>`
    + `<a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>`
    + '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
    + `<a:bgFillStyleLst>${solid}${solid}${solid}</a:bgFillStyleLst></a:fmtScheme>`
    + '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';
}
