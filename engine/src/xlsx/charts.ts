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
