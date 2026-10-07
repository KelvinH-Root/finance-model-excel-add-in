// The package writer: a whole workbook from a layout, built part by part (workbook, sheets, styles,
// theme, names, drawings, charts, the logo and the model metadata) and zipped. It runs in Node for
// builds and in the task pane for New model, where Excel.createWorkbook opens the result.

import { zipSync, strToU8 } from 'fflate';
import { colLetter, CONTENTS, FIRST_PERIOD_COL, META_NS } from '../frame.ts';
import type { Layout, RChart } from '../layout.ts';
import type { Brand, Model } from '../model.ts';
import { chartRefs, rchartRefs, sheetPrefix } from '../render.ts';
import { StyleBook } from '../styles.ts';
import { themeXml, THEMES } from '../theme.ts';
import { CHART_SIZE, chartXml, drawingXml, EMU_PER_CM, EMU_PER_PX, rchartXml, type DrawingItem } from './charts.ts';
import { GRID, RCHART_CM } from '../reports.ts';
import { REPORT_COL_WIDTH } from './dress.ts';
import { dress } from './dress.ts';
import { controlAnchor, esc, type ControlOut, type SheetOut } from './sheet.ts';

export interface Logo {
  png: Uint8Array;
  width: number;
  height: number;
}

export interface BuildOptions {
  /** The entity's logo for the contents; New model passes the one for the chosen entity. */
  logo?: Logo;
  /** Written as the file's created and modified time; fixed in tests so builds compare byte for byte. */
  created?: Date;
  /** The model metadata part (custom XML). On by default. */
  metadata?: boolean;
}

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = {
  workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  sheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  customProps: 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
  ctrlProp: 'application/vnd.ms-excel.controlproperties+xml',
};
const REL = {
  doc: `${NS_REL}/officeDocument`, sheet: `${NS_REL}/worksheet`, styles: `${NS_REL}/styles`, theme: `${NS_REL}/theme`,
  drawing: `${NS_REL}/drawing`, chart: `${NS_REL}/chart`, image: `${NS_REL}/image`, customXml: `${NS_REL}/customXml`,
  customProps: `${NS_REL}/customXmlProps`, vml: `${NS_REL}/vmlDrawing`, ctrlProp: `${NS_REL}/ctrlProp`, core: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  app: `${NS_REL}/extended-properties`,
};

const rels = (list: [string, string, string][]) => XML + `<Relationships xmlns="${NS_PKG}">`
  + list.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${esc(target)}"/>`).join('') + '</Relationships>';

/** The model metadata: the model itself, its link records and the row ids on each sheet. */
export function metadataXml(model: Model, layout: Layout): string {
  const payload = { model: model.toDict(), records: layout.records, rows: Object.fromEntries(layout.sheets.map(([s, rows]) => [s, rows.map(r => r.id)])) };
  return XML + `<hfgModel xmlns="${META_NS}" version="0">${esc(JSON.stringify(payload))}</hfgModel>`;
}

export function buildWorkbook(layout: Layout, model: Model, opts: BuildOptions = {}): Uint8Array {
  const brand: Brand = model.info?.entity.brand ?? 'HF';
  const book = new StyleBook(brand);
  const sheets = dress(layout, book);
  const files: Record<string, Uint8Array> = {};
  const put = (path: string, text: string) => { files[path] = strToU8(text); };
  const overrides: [string, string][] = [];
  const pos = layout.positions();

  // Drawings: module charts on their sheets, the logo on the contents.
  const drawings = new Map<string, { items: DrawingItem[]; rels: [string, string, string][] }>();
  const drawingFor = (sheet: string) => {
    let d = drawings.get(sheet);
    if (!d) drawings.set(sheet, d = { items: [], rels: [] });
    return d;
  };
  let chartNo = 0;
  for (const c of layout.charts) {
    const refs = chartRefs(layout, c, 'excel');
    chartNo += 1;
    put(`xl/charts/chart${chartNo}.xml`, chartXml(refs));
    overrides.push([`/xl/charts/chart${chartNo}.xml`, CT.chart]);
    const d = drawingFor(c.sheet);
    const rel = `rId${d.rels.length + 1}`;
    d.rels.push([rel, REL.chart, `../charts/chart${chartNo}.xml`]);
    d.items.push({ kind: 'chart', rel, name: c.title, at: { col: refs.anchor.col - 1, row: refs.anchor.row - 1, ...CHART_SIZE } });
  }
  // Report charts: a grid of three across under each report module's selections.
  for (const c of layout.rcharts) {
    chartNo += 1;
    put(`xl/charts/chart${chartNo}.xml`, rchartXml(rchartRefs(layout, c, 'excel')));
    overrides.push([`/xl/charts/chart${chartNo}.xml`, CT.chart]);
    const d = drawingFor(c.sheet);
    const rel = `rId${d.rels.length + 1}`;
    d.rels.push([rel, REL.chart, `../charts/chart${chartNo}.xml`]);
    const a = rchartAnchor(c, pos);
    d.items.push({ kind: 'chart', rel, name: c.id, at: { col: a.col - 1, colOff: a.offPx * EMU_PER_PX, row: a.row - 1,
      rowOff: 4 * EMU_PER_PX, cx: Math.round(RCHART_CM.w * EMU_PER_CM), cy: Math.round(RCHART_CM.h * EMU_PER_CM) } });
  }
  const contents = sheets.find(s => s.name === CONTENTS);
  if (opts.logo && contents && layout.frame.id === 'standard') {
    files['xl/media/logo.png'] = opts.logo.png;
    const theme = THEMES[brand];
    if (theme.logo.onDark) for (let r = 1; r <= 3; r++) contents.style(r, 10, book.xf('logoTile'));
    const cy = 30 * 12700;   // 30 points tall, within the three header rows
    const cx = Math.round(cy * opts.logo.width / opts.logo.height);
    const d = drawingFor(CONTENTS);
    const rel = `rId${d.rels.length + 1}`;
    d.rels.push([rel, REL.image, '../media/logo.png']);
    d.items.push({ kind: 'picture', rel, name: 'Logo', descr: `${theme.name} logo`,
      at: { col: 9, row: 0, colOff: 6 * EMU_PER_PX, rowOff: 5 * 12700, cx, cy } });
  }

  // Classic controls: a properties part each, the sheet's VML shapes, and a hidden shape in its drawing.
  let ctrlNo = 0;
  const vmlSheets = new Map<string, string>();
  sheets.forEach((s, i) => {
    if (!s.controls.length) return;
    const n = i + 1;
    const shapes: string[] = [];
    s.controls.forEach((c, k) => {
      ctrlNo += 1;
      const shapeId = 1024 * n + k + 1;
      const name = `${c.control.kind === 'drop' ? 'Drop' : 'Check'}_${c.link}`;   // found by its link's name
      put(`xl/ctrlProps/ctrlProp${ctrlNo}.xml`, ctrlPropXml(c));
      overrides.push([`/xl/ctrlProps/ctrlProp${ctrlNo}.xml`, CT.ctrlProp]);
      s.controlRels.push({ rel: `rId${3 + k}`, shapeId, name, part: `ctrlProp${ctrlNo}.xml` });
      const [from, to] = controlAnchor(c);
      drawingFor(s.name).items.push({ kind: 'control', shapeId, name, from, to });
      shapes.push(vmlShape(s, c, shapeId, name));
    });
    vmlSheets.set(s.name, vmlXml(n, shapes));
  });

  const sheetRels: [string, string, string][] = [];
  sheets.forEach((s, i) => {
    const n = i + 1;
    const d = drawings.get(s.name);
    const own: [string, string, string][] = [];
    if (d) {
      put(`xl/drawings/drawing${n}.xml`, drawingXml(d.items));
      if (d.rels.length) put(`xl/drawings/_rels/drawing${n}.xml.rels`, rels(d.rels));
      overrides.push([`/xl/drawings/drawing${n}.xml`, CT.drawing]);
      s.drawingRel = 'rId1';
      own.push(['rId1', REL.drawing, `../drawings/drawing${n}.xml`]);
    }
    const vml = vmlSheets.get(s.name);
    if (vml) {
      put(`xl/drawings/vmlDrawing${n}.vml`, vml);
      s.legacyRel = 'rId2';
      own.push(['rId2', REL.vml, `../drawings/vmlDrawing${n}.vml`]);
      for (const c of s.controlRels) own.push([c.rel, REL.ctrlProp, `../ctrlProps/${c.part}`]);
    }
    if (own.length) put(`xl/worksheets/_rels/sheet${n}.xml.rels`, rels(own));
    put(`xl/worksheets/sheet${n}.xml`, s.toXml());
    overrides.push([`/xl/worksheets/sheet${n}.xml`, CT.sheet]);
    sheetRels.push([`rId${n}`, REL.sheet, `worksheets/sheet${n}.xml`]);
  });

  const names = [...layout.names].map(([nm, rid]) => {
    const at = pos.get(rid);
    if (!at) throw new Error(`name ${nm} points at ${rid}, which is not in the layout`);
    return [nm, `${sheetPrefix(at[0], 'excel')}$${colLetter(layout.nameCol(nm))}$${at[1]}`];
  });
  for (const [nm, rg] of layout.ranges) {
    const a = pos.get(rg.from);
    const b = pos.get(rg.to);
    if (!a || !b) throw new Error(`range ${nm} points at rows not in the layout`);
    const L = colLetter(rg.col);
    const R = rg.toCol === undefined ? L : colLetter(rg.toCol === 'timeline' ? FIRST_PERIOD_COL + layout.periods - 1 : rg.toCol);
    names.push([nm, `${sheetPrefix(a[0], 'excel')}$${L}$${a[1]}:$${R}$${b[1]}`]);
  }
  names.sort((a, b) => a[0].toLowerCase().localeCompare(b[0].toLowerCase()));
  const k = sheets.length;
  put('xl/workbook.xml', XML + `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><workbookPr defaultThemeVersion="164011"/>`
    + '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15600" activeTab="0"/></bookViews><sheets>'
    + sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets>'
    + (names.length ? `<definedNames>${names.map(([n, r]) => `<definedName name="${esc(n)}">${esc(r)}</definedName>`).join('')}</definedNames>` : '')
    + '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
  const wbRels: [string, string, string][] = [...sheetRels, [`rId${k + 1}`, REL.styles, 'styles.xml'], [`rId${k + 2}`, REL.theme, 'theme/theme1.xml']];
  if (opts.metadata !== false) {
    wbRels.push([`rId${k + 3}`, REL.customXml, '../customXml/item1.xml']);
    put('customXml/item1.xml', metadataXml(model, layout));
    put('customXml/itemProps1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n'
      + '<ds:datastoreItem ds:itemID="{6B1F3C2A-0D7E-4C55-9A61-2F0E5B7C8D11}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">'
      + `<ds:schemaRefs><ds:schemaRef ds:uri="${META_NS}"/></ds:schemaRefs></ds:datastoreItem>`);
    put('customXml/_rels/item1.xml.rels', rels([['rId1', REL.customProps, 'itemProps1.xml']]));
    overrides.push(['/customXml/itemProps1.xml', CT.customProps]);
  }
  put('xl/_rels/workbook.xml.rels', rels(wbRels));
  put('xl/styles.xml', book.xml());   // after dressing and drawings, so every cell format is in it
  put('xl/theme/theme1.xml', themeXml(brand));
  overrides.push(['/xl/workbook.xml', CT.workbook], ['/xl/styles.xml', CT.styles], ['/xl/theme/theme1.xml', CT.theme],
    ['/docProps/core.xml', CT.core], ['/docProps/app.xml', CT.app]);

  const when = (opts.created ?? new Date()).toISOString().replace(/\.\d{3}Z$/, 'Z');
  const title = model.info?.title ?? 'HFG model';
  put('docProps/core.xml', XML + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
    + 'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    + `<dc:title>${esc(title)}</dc:title><dc:creator>HFG model engine</dc:creator>`
    + `<dcterms:created xsi:type="dcterms:W3CDTF">${when}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${when}</dcterms:modified></cp:coreProperties>`);
  put('docProps/app.xml', XML + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>');
  put('_rels/.rels', rels([['rId1', REL.doc, 'xl/workbook.xml'], ['rId2', REL.core, 'docProps/core.xml'], ['rId3', REL.app, 'docProps/app.xml']]));
  put('[Content_Types].xml', XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
    + (vmlSheets.size ? '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' : '')
    + overrides.map(([p, t]) => `<Override PartName="${p}" ContentType="${t}"/>`).join('') + '</Types>');

  const order = ['[Content_Types].xml', '_rels/.rels', ...Object.keys(files).filter(f => f !== '[Content_Types].xml' && f !== '_rels/.rels').sort()];
  const ordered: Record<string, Uint8Array> = {};
  for (const f of order) ordered[f] = files[f];
  return zipSync(ordered, { level: 6, mtime: opts.created ?? new Date() });
}

/**
 * Where a report chart sits: the row and column (1-based) of its top-left cell and its offset into
 * that column in pixels. Three charts across from column B, a row of charts every GRID.rows rows.
 */
export function rchartAnchor(c: RChart, pos: Map<string, [string, number]>): { row: number; col: number; offPx: number } {
  const px = (w: number) => Math.trunc(((256 * w + Math.trunc(128 / 7)) / 256) * 7);
  const widths = (col: number) => (col === 1 ? 3.75 : col <= 6 ? 2.5 : col === 7 ? 34 : col === 8 ? 7 : col === 9 ? 14 : REPORT_COL_WIDTH);
  const chartPx = Math.round(RCHART_CM.w / 2.54 * 96);
  const at = pos.get(c.grid);
  if (!at) throw new Error(`chart ${c.id}: no grid row ${c.grid}`);
  let x = (c.slot % GRID.cols) * (chartPx + 14);   // from the left of column B
  let col = 2;
  while (x >= px(widths(col))) { x -= px(widths(col)); col += 1; }
  return { row: at[1] + Math.floor(c.slot / GRID.cols) * GRID.rows, col, offPx: x };
}

/** The column after the timeline, where module charts sit. */
export const chartColumn = (layout: Layout) => FIRST_PERIOD_COL + layout.periods + 1;

const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';

/** A control's properties part: a combo box reading a List_ range, or a check box. */
export function ctrlPropXml(c: ControlOut): string {
  if (c.control.kind === 'drop') {
    const sel = typeof c.value === 'number' ? c.value : 1;
    return XML + `<formControlPr xmlns="${NS_X14}" objectType="Drop" dropLines="${c.control.lines ?? 20}" dropStyle="combo" dx="26" `
      + `fmlaLink="${esc(c.link)}" fmlaRange="${esc(c.control.list)}" noThreeD="1" sel="${sel}" val="0"/>`;
  }
  return XML + `<formControlPr xmlns="${NS_X14}" objectType="CheckBox"${c.value === true ? ' checked="Checked"' : ''} `
    + `fmlaLink="${esc(c.link)}" lockText="1" noThreeD="1"/>`;
}

/** Points from the sheet's left edge to a column, and from the top to a row (for the VML shape's position). */
function offsets(s: SheetOut, col: number, row: number): { left: number; top: number } {
  const width = (c: number) => s.cols.find(x => c >= x.min && c <= x.max)?.width ?? 8.43;
  let px = 0;
  for (let c = 1; c < col; c++) px += Math.round(width(c) * 7 + 5);
  let top = 0;
  for (let r = 1; r < row; r++) top += s.rows.get(r)?.ht ?? s.defaultHeight;
  return { left: Math.round(px * 0.75 * 10) / 10, top: Math.round(top * 10) / 10 };
}

function vmlShape(s: SheetOut, c: ControlOut, shapeId: number, name: string): string {
  const [from, to] = controlAnchor(c);
  const { left, top } = offsets(s, c.col, c.row);
  const w = Math.round((s.cols.find(x => c.col >= x.min && c.col <= x.max)?.width ?? 8.43) * 7 + 5) * 0.75;
  const h = s.rows.get(c.row)?.ht ?? s.defaultHeight;
  const anchor = `${from.col}, ${from.colOff}, ${from.row}, 0, ${to.col}, ${to.colOff}, ${to.row}, 0`;
  const style = `position:absolute;margin-left:${left}pt;margin-top:${top}pt;width:${w}pt;height:${h}pt;z-index:${shapeId % 1024}`;
  if (c.control.kind === 'drop') {
    const sel = typeof c.value === 'number' ? c.value : 1;
    return `<v:shape id="${esc(name)}" o:spid="_x0000_s${shapeId}" type="#_x0000_t201" style='${style}' filled="f" fillcolor="windowText [64]" `
      + 'strokecolor="windowText [64]" o:insetmode="auto"><v:fill color2="window [65]" o:detectmouseclick="t"/><o:lock v:ext="edit" rotation="t" text="t"/>'
      + `<x:ClientData ObjectType="Drop"><x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill>`
      + `<x:FmlaLink>${esc(c.link)}</x:FmlaLink><x:Val>0</x:Val><x:Min>0</x:Min><x:Max>0</x:Max><x:Inc>1</x:Inc><x:Page>${c.control.lines ?? 20}</x:Page><x:Dx>26</x:Dx>`
      + `<x:FmlaRange>${esc(c.control.list)}</x:FmlaRange><x:Sel>${sel}</x:Sel><x:NoThreeD2/><x:SelType>Single</x:SelType><x:LCT>Normal</x:LCT>`
      + `<x:DropStyle>Combo</x:DropStyle><x:DropLines>${c.control.lines ?? 20}</x:DropLines></x:ClientData></v:shape>`;
  }
  return `<v:shape id="${esc(name)}" o:spid="_x0000_s${shapeId}" type="#_x0000_t201" style='${style};mso-wrap-style:tight' filled="f" fillcolor="window [65]" `
    + 'stroked="f" strokecolor="windowText [64]" o:insetmode="auto"><v:path shadowok="t" strokeok="t" fillok="t"/><o:lock v:ext="edit" rotation="t"/>'
    + '<v:textbox style=\'mso-direction-alt:auto\' o:singleclick="f"><div style=\'text-align:left\'></div></v:textbox>'
    + `<x:ClientData ObjectType="Checkbox"><x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill><x:AutoLine>False</x:AutoLine>`
    + `<x:TextVAlign>Center</x:TextVAlign>${c.value === true ? '<x:Checked>1</x:Checked>' : ''}<x:FmlaLink>${esc(c.link)}</x:FmlaLink><x:NoThreeD/></x:ClientData></v:shape>`;
}

function vmlXml(n: number, shapes: string[]): string {
  return '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">'
    + `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${n}"/></o:shapelayout>`
    + '<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
    + '<v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/><o:lock v:ext="edit" shapetype="t"/></v:shapetype>'
    + shapes.join('') + '</xml>';
}
