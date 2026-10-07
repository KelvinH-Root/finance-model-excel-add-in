// One worksheet as the package writer holds it before writing: cells with their cell formats,
// row heights and outline levels, columns, freeze panes, hyperlinks to defined names, conditional
// formats, validations and the drawing (charts and pictures). toXml writes SpreadsheetML in the
// element order Excel requires.

import { colLetter } from '../frame.ts';
import type { Control } from '../layout.ts';

export const esc = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

export const ref = (r: number, c: number): string => `${colLetter(c)}${r}`;

export interface CellOut {
  /** A value; strings are written inline. */
  v?: string | number | boolean | null;
  /** A formula, without the leading "=". */
  f?: string;
  /** Cell format index (styles.xml cellXfs). */
  s: number;
  /** The first cell of a one-variable data table: its results range and its input cell (inputs across the row above). */
  dt?: { ref: string; r1: string };
}

export interface RowOut {
  ht?: number;
  level?: number;
  hidden?: boolean;
  collapsed?: boolean;
}

export interface ColOut {
  min: number;
  max: number;
  width: number;
  level?: number;
  hidden?: boolean;
}

export interface LinkOut {
  row: number;
  col: number;
  /** Defined name the link goes to. */
  to: string;
  tip: string;
}

export interface CondOut {
  sqref: string;
  /** cfRule elements with dxfId set and priority left as {p}. */
  rules: string[];
}

/** A classic control over a cell (drawn by the package writer), with the defined name it links to. */
export interface ControlOut {
  row: number;
  col: number;
  control: Control;
  link: string;
  /** The linked cell's value when written: the position chosen, or TRUE or FALSE. */
  value: unknown;
}

/** A zero-based cell corner with an offset in drawing units. */
export interface AnchorMark { col: number; colOff: number; row: number }

/** Where a control sits: over its linked cell (a drop-down fills it; a check box sits at its left edge). */
export function controlAnchor(c: ControlOut): [AnchorMark, AnchorMark] {
  return [{ col: c.col - 1, colOff: 0, row: c.row - 1 }, { col: c.col, colOff: 0, row: c.row }];
}

export class SheetOut {
  name: string;
  cells = new Map<number, Map<number, CellOut>>();
  rows = new Map<number, RowOut>();
  cols: ColOut[] = [];
  defaultHeight = 11.4;
  freeze: { row: number; col: number } | null = null;
  summaryBelow = true;
  links: LinkOut[] = [];
  conds: CondOut[] = [];
  validations: string[] = [];
  controls: ControlOut[] = [];
  /** Relationship ids of the sheet's legacy drawing (VML) and of each control's properties, set by the package writer. */
  legacyRel: string | null = null;
  controlRels: { rel: string; shapeId: number; name: string; part: string }[] = [];
  /** Relationship id of the sheet's drawing, set by the package writer. */
  drawingRel: string | null = null;
  zoom = 100;

  constructor(name: string) {
    this.name = name;
  }

  set(r: number, c: number, cell: CellOut): void {
    let row = this.cells.get(r);
    if (!row) this.cells.set(r, row = new Map());
    row.set(c, cell);
  }

  get(r: number, c: number): CellOut | undefined {
    return this.cells.get(r)?.get(c);
  }

  /** Give a cell a format, keeping its value; a cell that does not exist yet is created blank. */
  style(r: number, c: number, s: number): void {
    const cell = this.get(r, c);
    if (cell) cell.s = s;
    else this.set(r, c, { s });
  }

  row(r: number, props: RowOut): void {
    this.rows.set(r, { ...(this.rows.get(r) || {}), ...props });
  }

  private cellXml(r: number, c: number, cell: CellOut): string {
    const a = `r="${ref(r, c)}"${cell.s ? ` s="${cell.s}"` : ''}`;
    if (cell.dt) return `<c ${a}><f t="dataTable" ref="${cell.dt.ref}" dt2D="0" dtr="1" r1="${cell.dt.r1}"/></c>`;
    if (cell.f !== undefined) return `<c ${a}><f>${esc(cell.f)}</f></c>`;
    const v = cell.v;
    if (v === undefined || v === null || v === '') return `<c ${a}/>`;
    if (typeof v === 'boolean') return `<c ${a} t="b"><v>${v ? 1 : 0}</v></c>`;
    if (typeof v === 'number') {
      if (!Number.isFinite(v)) throw new Error(`${this.name}!${ref(r, c)}: ${v} is not a number Excel can store`);
      return `<c ${a}><v>${v}</v></c>`;
    }
    const space = v !== v.trim() ? ' xml:space="preserve"' : '';
    return `<c ${a} t="inlineStr"><is><t${space}>${esc(v)}</t></is></c>`;
  }

  toXml(): string {
    const out: string[] = [];
    const w = (s: string) => out.push(s);
    const rowNums = [...new Set([...this.cells.keys(), ...this.rows.keys()])].sort((a, b) => a - b);
    let maxCol = 1;
    for (const row of this.cells.values()) for (const c of row.keys()) maxCol = Math.max(maxCol, c);
    const maxRow = rowNums.length ? rowNums[rowNums.length - 1] : 1;
    const maxLevel = Math.max(0, ...[...this.rows.values()].map(p => p.level || 0));
    const maxColLevel = Math.max(0, ...this.cols.map(c => c.level || 0));
    w('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n');
    w('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
      + (this.controls.length ? ' xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
        + ' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"'
        + ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' : '') + '>');
    w(`<sheetPr><outlinePr summaryBelow="${this.summaryBelow ? 1 : 0}" summaryRight="1"/><pageSetUpPr fitToPage="1"/></sheetPr>`);
    w(`<dimension ref="A1:${ref(maxRow, maxCol)}"/>`);
    let view = `<sheetView showGridLines="0" zoomScale="${this.zoom}" zoomScaleNormal="${this.zoom}" workbookViewId="0">`;
    if (this.freeze) {
      const { row, col } = this.freeze;
      const xs = col - 1;
      const ys = row - 1;
      const tl = ref(row, col);
      const pane = xs && ys ? 'bottomRight' : xs ? 'topRight' : 'bottomLeft';
      view += `<pane${xs ? ` xSplit="${xs}"` : ''}${ys ? ` ySplit="${ys}"` : ''} topLeftCell="${tl}" activePane="${pane}" state="frozen"/>`;
      view += `<selection pane="${pane}" activeCell="${tl}" sqref="${tl}"/>`;
    } else {
      view += '<selection activeCell="B4" sqref="B4"/>';
    }
    w(`<sheetViews>${view}</sheetView></sheetViews>`);
    w(`<sheetFormatPr defaultRowHeight="${this.defaultHeight}" customHeight="1"`
      + `${maxLevel ? ` outlineLevelRow="${maxLevel}"` : ''}${maxColLevel ? ` outlineLevelCol="${maxColLevel}"` : ''}/>`);
    if (this.cols.length) {
      w('<cols>');
      for (const c of [...this.cols].sort((a, b) => a.min - b.min)) {
        w(`<col min="${c.min}" max="${c.max}" width="${c.width}" customWidth="1"`
          + `${c.hidden ? ' hidden="1"' : ''}${c.level ? ` outlineLevel="${c.level}"` : ''}/>`);
      }
      w('</cols>');
    }
    w('<sheetData>');
    for (const r of rowNums) {
      const p = this.rows.get(r) || {};
      let attrs = `r="${r}"`;
      if (p.ht !== undefined) attrs += ` ht="${p.ht}" customHeight="1"`;
      if (p.level) attrs += ` outlineLevel="${p.level}"`;
      if (p.hidden) attrs += ' hidden="1"';
      if (p.collapsed) attrs += ' collapsed="1"';
      const cells = this.cells.get(r);
      if (!cells || !cells.size) {
        w(`<row ${attrs}/>`);
        continue;
      }
      w(`<row ${attrs}>`);
      for (const c of [...cells.keys()].sort((a, b) => a - b)) w(this.cellXml(r, c, cells.get(c)!));
      w('</row>');
    }
    w('</sheetData>');
    let priority = 1;
    for (const cf of this.conds) {
      w(`<conditionalFormatting sqref="${cf.sqref}">${cf.rules.map(rule => rule.replace('{p}', String(priority++))).join('')}</conditionalFormatting>`);
    }
    if (this.validations.length) w(`<dataValidations count="${this.validations.length}">${this.validations.join('')}</dataValidations>`);
    if (this.links.length) {
      w('<hyperlinks>');
      for (const l of this.links) w(`<hyperlink ref="${ref(l.row, l.col)}" location="${esc(l.to)}" tooltip="${esc(l.tip)}"/>`);
      w('</hyperlinks>');
    }
    w('<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>');
    w('<pageSetup paperSize="9" orientation="landscape" fitToHeight="0"/>');
    if (this.drawingRel) w(`<drawing r:id="${this.drawingRel}"/>`);
    if (this.legacyRel) w(`<legacyDrawing r:id="${this.legacyRel}"/>`);
    if (this.controlRels.length) {
      w('<mc:AlternateContent><mc:Choice Requires="x14"><controls>');
      this.controls.forEach((c, i) => {
        const { rel, shapeId, name } = this.controlRels[i];
        const [from, to] = controlAnchor(c);
        const mark = (m: AnchorMark) => `<xdr:col>${m.col}</xdr:col><xdr:colOff>${m.colOff}</xdr:colOff><xdr:row>${m.row}</xdr:row><xdr:rowOff>0</xdr:rowOff>`;
        w(`<mc:AlternateContent><mc:Choice Requires="x14"><control shapeId="${shapeId}" r:id="${rel}" name="${esc(name)}">`
          + `<controlPr defaultSize="0" autoFill="0" autoLine="0" autoPict="0"><anchor moveWithCells="1" sizeWithCells="1">`
          + `<from>${mark(from)}</from><to>${mark(to)}</to></anchor></controlPr></control></mc:Choice></mc:AlternateContent>`);
      });
      w('</controls></mc:Choice></mc:AlternateContent>');
    }
    w('</worksheet>');
    return out.join('');
  }
}
