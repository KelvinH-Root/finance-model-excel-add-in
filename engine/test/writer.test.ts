// The package writer: the parts a workbook needs, names that resolve, deterministic output, and
// the theme and style rules (deeper checks, with LibreOffice and against the Python writer, are in
// tests/test_engine_writer.py).

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { strFromU8, unzipSync } from 'fflate';
import {
  assemble, barTextSlot, buildWorkbook, catalogue, contrast, Library, SLOT, STD, THEMES, type Brand,
} from '../src/index.ts';
import { loadLogo } from '../src/node/assets.ts';
import { demoModel, LIBRARY } from './demo.ts';
import { loadLibrary } from '../src/node/library.ts';

const lib = loadLibrary(LIBRARY);
const created = new Date('2026-10-07T00:00:00Z');
const BRANDS: Brand[] = ['HF', 'HCP', 'HCL', 'KM', 'TWK'];

function build(brand: Brand | null) {
  const model = demoModel(lib, brand);
  const layout = assemble(model);
  const bytes = buildWorkbook(layout, model, { created, logo: brand ? loadLogo(brand) : undefined });
  const parts = unzipSync(bytes);
  const text = (p: string) => strFromU8(parts[p]);
  return { model, layout, bytes, parts, text };
}

test('a standard build has every part and declares each one', () => {
  const { parts, text, layout } = build('HF');
  for (const p of ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/styles.xml', 'xl/theme/theme1.xml',
    'docProps/core.xml', 'customXml/item1.xml', 'xl/media/logo.png']) assert.ok(parts[p], p);
  const types = text('[Content_Types].xml');
  for (const p of Object.keys(parts)) {
    if (p.endsWith('.rels') || p.endsWith('.png') || p.endsWith('.vml') || p === '[Content_Types].xml' || p === 'customXml/item1.xml') continue;
    assert.ok(types.includes(`PartName="/${p}"`), `${p} is not in [Content_Types].xml`);
  }
  assert.equal(Object.keys(parts).filter(p => /^xl\/worksheets\/sheet\d+\.xml$/.test(p)).length, layout.sheets.length);
  assert.equal(Object.keys(parts).filter(p => p.startsWith('xl/charts/')).length, layout.charts.length);
  if (Object.keys(parts).some(p => p.endsWith('.vml'))) assert.ok(types.includes('Extension="vml"'), 'VML is declared');
});

test('the Settings sheet draws a control over each drop-down and check box cell, reading the Lookups lists', () => {
  const { parts, text, layout } = build('HF');
  const n = layout.sheets.findIndex(([s]) => s === 'Settings') + 1;
  const sheet = text(`xl/worksheets/sheet${n}.xml`);
  const rels = text(`xl/worksheets/_rels/sheet${n}.xml.rels`);
  const props = Object.keys(parts).filter(p => p.startsWith('xl/ctrlProps/')).map(p => text(p));
  assert.equal(props.filter(p => p.includes('objectType="Drop"')).length, 5);
  assert.equal(props.filter(p => p.includes('objectType="CheckBox"')).length, 2);
  for (const list of ['List_Month_Names', 'List_Start_Months', 'List_Denominations', 'List_Last_Actual', 'List_Months']) {
    assert.ok(props.some(p => p.includes(`fmlaRange="${list}"`)), list);
  }
  assert.equal((sheet.match(/<control shapeId=/g) || []).length, 7);
  assert.match(sheet, /<legacyDrawing r:id="rId2"\/>/);
  assert.equal((rels.match(/ctrlProp"/g) || []).length, 7);
  const wb = text('xl/workbook.xml');
  assert.match(wb, /<definedName name="List_Month_Names">Lookups!\$D\$\d+:\$D\$\d+<\/definedName>/);
  assert.match(wb, /<definedName name="Sel_FY_End_Month">Settings!\$I\$\d+<\/definedName>/);
  // the drop-down opens on the setting chosen: a March year end is the third month
  assert.ok(props.some(p => p.includes('fmlaLink="Sel_FY_End_Month"') && p.includes('sel="3"')));
});

test('every defined name and hyperlink resolves', () => {
  const { text, layout } = build('HF');
  const wb = text('xl/workbook.xml');
  const sheets = new Set([...wb.matchAll(/<sheet name="([^"]+)"/g)].map(m => m[1].replace(/&amp;/g, '&')));
  const names = new Set<string>();
  for (const m of wb.matchAll(/<definedName name="([^"]+)">([^<]+)<\/definedName>/g)) {
    names.add(m[1]);
    const sheet = m[2].replace(/&apos;|'/g, "'").match(/^'?([^'!]+)'?!/)![1];
    assert.ok(sheets.has(sheet), `${m[1]} points at ${sheet}`);
  }
  for (const [nm] of layout.names) assert.ok(names.has(nm), nm);
  layout.sheets.forEach((_, i) => {
    for (const m of text(`xl/worksheets/sheet${i + 1}.xml`).matchAll(/<hyperlink ref="[^"]+" location="([^"]+)" tooltip="([^"]*)"/g)) {
      assert.ok(names.has(m[1]), `link to ${m[1]}`);
      assert.ok(m[2].length > 3, 'every link has a screen tip');
    }
  });
});

test('builds are byte for byte repeatable', () => {
  assert.deepEqual(build('KM').bytes, build('KM').bytes);
});

test('section bar text reads on every entity colour', () => {
  const expected: Record<Brand, number> = { HF: SLOT.dk2, HCP: SLOT.lt1, HCL: SLOT.lt1, KM: SLOT.dk2, TWK: SLOT.dk2 };
  for (const b of BRANDS) {
    const t = THEMES[b];
    assert.equal(barTextSlot(t), expected[b], b);
    const text = barTextSlot(t) === SLOT.lt1 ? 'FFFFFF' : t.dk2;
    assert.ok(contrast(text, t.accents[0]) >= 4.5, `${b} bar text contrast`);
  }
});

test('style names are HFG\'s own and unique', () => {
  const names = Object.values(catalogue('HF')).map(s => s.name);
  assert.equal(new Set(names).size, names.length);
  for (const n of names) assert.ok(n === 'Normal' || n.startsWith('HFG '), n);
  assert.ok(names.every(n => !n.endsWith('.')));
});

test('the standard frame puts content below the timeline block and the contents below its header', () => {
  const { layout } = build('TWK');
  assert.equal(layout.firstRowOf('Revenue'), STD.contentRow);
  assert.equal(layout.firstRowOf('Contents'), STD.plainRow);
  assert.equal(layout.kindOf('Settings'), 'settings');
  const order = layout.sheets.map(([s]) => s);
  assert.ok(order.indexOf('Settings') < order.indexOf('Checks'));
  assert.equal(layout.names.get('Model_Name'), '@Contents/2');
});

test('a proof build carries no standard frame parts', () => {
  const { parts, text } = build(null);
  assert.ok(!parts['xl/media/logo.png']);
  assert.ok(!text('xl/workbook.xml').includes('Go_Contents'));
});
