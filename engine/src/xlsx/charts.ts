// Chart and drawing parts. A module chart is stacked columns with any line series drawn over them,
// every series coloured by theme slot (accent 1, 2, 3 in order) so the chart takes the entity's
// theme; series names and categories read cells, so legends follow the model.

import type { ChartRefs } from '../render.ts';
import { esc, type AnchorMark } from './sheet.ts';

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';

/** 1 cm in drawing units. */
export const EMU_PER_CM = 360000;
export const EMU_PER_PX = 9525;
export const CHART_SIZE = { cx: 16 * EMU_PER_CM, cy: 7.5 * EMU_PER_CM };

const accent = (i: number) => `accent${(i % 6) + 1}`;

export function chartXml(c: ChartRefs): string {
  const columns = c.series.filter(s => s.kind === 'column');
  const lines = c.series.filter(s => s.kind === 'line');
  let idx = 0;
  const ser = (s: ChartRefs['series'][number], fill: string) => {
    const i = idx++;
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>${esc(s.label)}</c:f></c:strRef></c:tx>${fill}`
      + `<c:cat><c:strRef><c:f>${esc(c.categories)}</c:f></c:strRef></c:cat>`
      + `<c:val><c:numRef><c:f>${esc(s.values)}</c:f></c:numRef></c:val>`;
  };
  const bars = columns.map((s, i) => ser(s, `<c:spPr><a:solidFill><a:schemeClr val="${accent(i)}"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>`) + '</c:ser>');
  const lineSers = lines.map((s, i) => ser(s, `<c:spPr><a:ln w="22225" cap="rnd"><a:solidFill><a:schemeClr val="${accent(columns.length + i)}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>`) + '<c:smooth val="0"/></c:ser>');
  const axes = '<c:axId val="50010001"/><c:axId val="50010002"/>';
  const text = (sz: number, bold = false) => `<a:defRPr sz="${sz}"${bold ? ' b="1"' : ''}><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr>`;
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:roundedCorners val="0"/><c:chart>`
    + `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr>${text(1000, true)}</a:pPr><a:r><a:rPr lang="en-NZ" sz="1000" b="1"/><a:t>${esc(c.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
    + '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>'
    + (bars.length ? `<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:varyColors val="0"/>${bars.join('')}<c:gapWidth val="60"/><c:overlap val="100"/>${axes}</c:barChart>` : '')
    + (lineSers.length ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${lineSers.join('')}<c:marker val="1"/>${axes}</c:lineChart>` : '')
    + '<c:catAx><c:axId val="50010001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>'
    + '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>'
    + '<c:spPr><a:ln w="6350"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="50000"/><a:lumOff val="50000"/></a:schemeClr></a:solidFill></a:ln></c:spPr>'
    + '<c:crossAx val="50010002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
    + '<c:valAx><c:axId val="50010002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>'
    + '<c:majorGridlines><c:spPr><a:ln w="3175"><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="85000"/></a:schemeClr></a:solidFill></a:ln></c:spPr></c:majorGridlines>'
    + '<c:numFmt formatCode="#,##0" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
    + '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:crossAx val="50010001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>'
    + '</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
    + `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>${text(800)}</a:pPr><a:endParaRPr lang="en-NZ"/></a:p></c:txPr></c:chartSpace>`;
}

export interface Anchor {
  /** Zero-based column and row of the top-left corner. */
  col: number;
  row: number;
  colOff?: number;
  rowOff?: number;
  cx: number;
  cy: number;
}

export type DrawingItem =
  | { kind: 'chart'; rel: string; name: string; at: Anchor }
  | { kind: 'picture'; rel: string; name: string; descr: string; at: Anchor }
  /** A classic control's shape: hidden here, drawn from the sheet's VML. */
  | { kind: 'control'; shapeId: number; name: string; from: AnchorMark; to: AnchorMark };

export function drawingXml(items: DrawingItem[]): string {
  const from = (a: Anchor) => `<xdr:from><xdr:col>${a.col}</xdr:col><xdr:colOff>${a.colOff ?? 0}</xdr:colOff><xdr:row>${a.row}</xdr:row><xdr:rowOff>${a.rowOff ?? 0}</xdr:rowOff></xdr:from><xdr:ext cx="${a.cx}" cy="${a.cy}"/>`;
  const mark = (m: AnchorMark) => `<xdr:col>${m.col}</xdr:col><xdr:colOff>${m.colOff}</xdr:colOff><xdr:row>${m.row}</xdr:row><xdr:rowOff>0</xdr:rowOff>`;
  const body = items.map((it, i) => {
    const id = i + 2;
    if (it.kind === 'control') {
      return '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        + '<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">'
        + `<xdr:twoCellAnchor><xdr:from>${mark(it.from)}</xdr:from><xdr:to>${mark(it.to)}</xdr:to>`
        + `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${it.shapeId}" name="${esc(it.name)}" hidden="1"><a:extLst>`
        + `<a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}"><a14:compatExt spid="_x0000_s${it.shapeId}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr>`
        + '<xdr:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        + '<a:noFill/><a:ln w="9525"><a:miter lim="800000"/><a:headEnd/><a:tailEnd/></a:ln></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>'
        + '</mc:Choice><mc:Fallback/></mc:AlternateContent>';
    }
    if (it.kind === 'chart') {
      return `<xdr:oneCellAnchor>${from(it.at)}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(it.name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
        + '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
        + `<a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" r:id="${it.rel}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`;
    }
    return `<xdr:oneCellAnchor>${from(it.at)}<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${esc(it.name)}" descr="${esc(it.descr)}"/>`
      + '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
      + `<xdr:blipFill><a:blip r:embed="${it.rel}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>`
      + `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${it.at.cx}" cy="${it.at.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
  });
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + `<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">${body.join('')}</xdr:wsDr>`;
}

// --- Report charts --------------------------------------------------------------------------------
// Charts that summary and report modules bring: bars, lines, combinations, pies and waterfalls,
// coloured by theme slot so they take the entity's theme. Titles, series names and categories read
// cells, so a chart follows the selections on its sheet.

/** A report chart's parts as cell references in one dialect. */
export interface RChartRefs {
  title: { ref: string; cache: string };
  cats: string;
  series: (RSeriesLook & { tx: string; values: string })[];
  type: 'bar' | 'pie';
  dir: 'col' | 'bar';
  grouping: 'clustered' | 'stacked';
  gap: number;
  overlap: number;
  valueAxis: boolean;
  reverse: boolean;
  yFmt: string;
  legend: 'b' | 'r' | null;
}

export interface RSeriesLook {
  as: 'bar' | 'line';
  colour: string;
  width?: number;
  dash?: boolean;
  marker?: boolean;
  hatch?: boolean;
  outline?: boolean;
  labels?: { fmt: string; pos?: string; pct?: boolean };
  points?: string[];
}

const TEXT_CLR = '<a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr>';

/** A theme colour: accent1 to accent6, tx2 (dark 2), grey or light (background shades). */
function clr(colour: string): string {
  if (colour === 'grey') return '<a:schemeClr val="bg1"><a:lumMod val="65000"/></a:schemeClr>';
  if (colour === 'light') return '<a:schemeClr val="bg1"><a:lumMod val="85000"/></a:schemeClr>';
  return `<a:schemeClr val="${colour}"/>`;
}

const rText = (sz: number, bold = false) => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${Math.round(sz * 100)}"${bold ? ' b="1"' : ''}>`
  + `<a:solidFill>${TEXT_CLR}</a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:pPr><a:endParaRPr lang="en-NZ"/></a:p></c:txPr>`;

function fillXml(s: RSeriesLook): string {
  if (s.colour === 'none') return '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>';
  const fill = s.hatch
    ? `<a:pattFill prst="upDiag"><a:fgClr>${clr(s.colour)}</a:fgClr><a:bgClr><a:schemeClr val="bg1"/></a:bgClr></a:pattFill>`
    : `<a:solidFill>${clr(s.colour)}</a:solidFill>`;
  const ln = s.outline || s.hatch ? `<a:ln w="6350"><a:solidFill>${s.hatch ? clr(s.colour) : TEXT_CLR}</a:solidFill></a:ln>` : '<a:ln><a:noFill/></a:ln>';
  return `<c:spPr>${fill}${ln}</c:spPr>`;
}

function lineXml(s: RSeriesLook): string {
  const w = Math.round((s.width ?? 1.5) * 12700);
  const marker = s.marker
    ? `<c:marker><c:symbol val="circle"/><c:size val="4"/><c:spPr><a:solidFill>${clr(s.colour)}</a:solidFill><a:ln><a:solidFill>${clr(s.colour)}</a:solidFill></a:ln></c:spPr></c:marker>`
    : '<c:marker><c:symbol val="none"/></c:marker>';
  return `<c:spPr><a:ln w="${w}" cap="rnd"><a:solidFill>${clr(s.colour)}</a:solidFill>${s.dash ? '<a:prstDash val="dash"/>' : ''}<a:round/></a:ln></c:spPr>${marker}`;
}

function labelsXml(l: NonNullable<RSeriesLook['labels']>): string {
  return `<c:dLbls><c:numFmt formatCode="${esc(l.fmt)}" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${rText(7)}`
    + `${l.pos ? `<c:dLblPos val="${l.pos}"/>` : ''}<c:showLegendKey val="0"/><c:showVal val="${l.pct ? 0 : 1}"/><c:showCatName val="0"/>`
    + `<c:showSerName val="0"/><c:showPercent val="${l.pct ? 1 : 0}"/><c:showBubbleSize val="0"/></c:dLbls>`;
}

export function rchartXml(c: RChartRefs): string {
  const head = (s: RChartRefs['series'][number], i: number) =>
    `<c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>${esc(s.tx)}</c:f></c:strRef></c:tx>`;
  const data = (s: RChartRefs['series'][number]) =>
    `<c:cat><c:strRef><c:f>${esc(c.cats)}</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>${esc(s.values)}</c:f></c:numRef></c:val>`;
  const ax = '<c:axId val="50020001"/><c:axId val="50020002"/>';
  let plot = '';
  if (c.type === 'pie') {
    const s = c.series[0];
    const pts = (s.points ?? []).map((p, j) => `<c:dPt><c:idx val="${j}"/><c:bubble3D val="0"/><c:spPr><a:solidFill>${clr(p)}</a:solidFill>`
      + '<a:ln w="9525"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:ln></c:spPr></c:dPt>').join('');
    plot = `<c:pieChart><c:varyColors val="1"/><c:ser>${head(s, 0)}${pts}${s.labels ? labelsXml(s.labels) : ''}${data(s)}</c:ser><c:firstSliceAng val="0"/></c:pieChart>`;
  } else {
    let i = 0;
    const bars = c.series.filter(s => s.as === 'bar').map(s => `<c:ser>${head(s, i++)}${fillXml(s)}<c:invertIfNegative val="0"/>`
      + `${s.labels ? labelsXml(s.labels) : ''}${data(s)}</c:ser>`);
    const lines = c.series.filter(s => s.as === 'line').map(s => `<c:ser>${head(s, i++)}${lineXml(s)}`
      + `${s.labels ? labelsXml(s.labels) : ''}${data(s)}<c:smooth val="0"/></c:ser>`);
    if (bars.length) {
      plot += `<c:barChart><c:barDir val="${c.dir}"/><c:grouping val="${c.grouping}"/><c:varyColors val="0"/>${bars.join('')}`
        + `<c:gapWidth val="${c.gap}"/>${c.grouping === 'stacked' || c.overlap ? `<c:overlap val="${c.grouping === 'stacked' ? 100 : c.overlap}"/>` : ''}${ax}</c:barChart>`;
    }
    if (lines.length) plot += `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${lines.join('')}<c:marker val="1"/>${ax}</c:lineChart>`;
    const catPos = c.dir === 'bar' ? 'l' : 'b';
    const valPos = c.dir === 'bar' ? 'b' : 'l';
    plot += `<c:catAx><c:axId val="50020001"/><c:scaling><c:orientation val="${c.reverse ? 'maxMin' : 'minMax'}"/></c:scaling><c:delete val="0"/>`
      + `<c:axPos val="${catPos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>`
      + `<c:tickLblPos val="low"/><c:spPr><a:ln w="6350"><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill></a:ln></c:spPr>`
      + `${rText(7.5)}<c:crossAx val="50020002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`
      + `<c:valAx><c:axId val="50020002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${c.valueAxis ? 0 : 1}"/><c:axPos val="${valPos}"/>`
      + (c.valueAxis ? '<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="85000"/></a:schemeClr></a:solidFill></a:ln></c:spPr></c:majorGridlines>' : '')
      + `<c:numFmt formatCode="${esc(c.yFmt)}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>`
      + `<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${rText(7.5)}<c:crossAx val="50020001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  }
  const legend = c.legend ? `<c:legend><c:legendPos val="${c.legend}"/><c:overlay val="0"/>${rText(7.5)}</c:legend>` : '';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:roundedCorners val="0"/><c:chart>`
    + `<c:title><c:tx><c:strRef><c:f>${esc(c.title.ref)}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${esc(c.title.cache)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`
    + `<c:overlay val="0"/>${rText(9.5, true)}</c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${plot}</c:plotArea>${legend}`
    + '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
    + '<c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="6350"><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="85000"/></a:schemeClr></a:solidFill></a:ln></c:spPr>'
    + `${rText(8)}</c:chartSpace>`;
}
