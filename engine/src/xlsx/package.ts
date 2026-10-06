// The package writer: a whole workbook from a layout, built part by part (workbook, sheets, styles,
// theme, names, drawings, charts, the logo and the model metadata) and zipped. It runs in Node for
// builds and in the task pane for New model, where Excel.createWorkbook opens the result.

import { zipSync, strToU8 } from 'fflate';
import { colLetter, CONTENTS, FIRST_PERIOD_COL, META_NS } from '../frame.ts';
import type { Layout } from '../layout.ts';
import type { Brand, Model } from '../model.ts';
import { chartRefs, sheetPrefix } from '../render.ts';
import { StyleBook } from '../styles.ts';
import { themeXml, THEMES } from '../theme.ts';
import { CHART_SIZE, chartXml, drawingXml, EMU_PER_PX, type DrawingItem } from './charts.ts';
import { dress } from './dress.ts';
import { esc } from './sheet.ts';

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
};
const REL = {
  doc: `${NS_REL}/officeDocument`, sheet: `${NS_REL}/worksheet`, styles: `${NS_REL}/styles`, theme: `${NS_REL}/theme`,
  drawing: `${NS_REL}/drawing`, chart: `${NS_REL}/chart`, image: `${NS_REL}/image`, customXml: `${NS_REL}/customXml`,
  customProps: `${NS_REL}/customXmlProps`, core: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
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

  const sheetRels: [string, string, string][] = [];
  sheets.forEach((s, i) => {
    const n = i + 1;
    const d = drawings.get(s.name);
    if (d) {
      put(`xl/drawings/drawing${n}.xml`, drawingXml(d.items));
      put(`xl/drawings/_rels/drawing${n}.xml.rels`, rels(d.rels));
      overrides.push([`/xl/drawings/drawing${n}.xml`, CT.drawing]);
      s.drawingRel = 'rId1';
      put(`xl/worksheets/_rels/sheet${n}.xml.rels`, rels([['rId1', REL.drawing, `../drawings/drawing${n}.xml`]]));
    }
    put(`xl/worksheets/sheet${n}.xml`, s.toXml());
    overrides.push([`/xl/worksheets/sheet${n}.xml`, CT.sheet]);
    sheetRels.push([`rId${n}`, REL.sheet, `worksheets/sheet${n}.xml`]);
  });

  const names = [...layout.names].map(([nm, rid]) => {
    const at = pos.get(rid);
    if (!at) throw new Error(`name ${nm} points at ${rid}, which is not in the layout`);
    return [nm, `${sheetPrefix(at[0], 'excel')}$${colLetter(layout.nameCol(nm))}$${at[1]}`];
  }).sort((a, b) => a[0].toLowerCase().localeCompare(b[0].toLowerCase()));
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
    + overrides.map(([p, t]) => `<Override PartName="${p}" ContentType="${t}"/>`).join('') + '</Types>');

  const order = ['[Content_Types].xml', '_rels/.rels', ...Object.keys(files).filter(f => f !== '[Content_Types].xml' && f !== '_rels/.rels').sort()];
  const ordered: Record<string, Uint8Array> = {};
  for (const f of order) ordered[f] = files[f];
  return zipSync(ordered, { level: 6, mtime: opts.created ?? new Date() });
}

/** The column after the timeline, where module charts sit. */
export const chartColumn = (layout: Layout) => FIRST_PERIOD_COL + layout.periods + 1;
