// New model wizard: defaults, checks, the model it makes and the file it builds.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import { loadLibrary } from '../../engine/src/node/library.ts';
import { loadLogo } from '../../engine/src/node/assets.ts';
import {
  buildFile, fileName, initialState, periodMonth, preview, problems, STEPS, toBase64, toModel, yearStart,
} from '../src/wizard/core.ts';

const LIBRARY = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'prototypes', 'assembly', 'library');
const lib = loadLibrary(LIBRARY);
const today = new Date(2026, 9, 7);   // 7 October 2026

test('defaults start at the financial year the date falls in', () => {
  const s = initialState(today);
  assert.equal(s.start, '2026-04');
  assert.equal(yearStart(new Date(2026, 1, 10), 3), '2025-04');
  assert.equal(yearStart(new Date(2026, 5, 1), 6), '2025-07');
  assert.equal(yearStart(new Date(2026, 6, 1), 6), '2026-07');
  assert.equal(yearStart(new Date(2026, 0, 1), 12), '2026-01');
  assert.equal(s.brand, 'HF');
});

test('each step says what is missing before moving on', () => {
  const s = initialState(today);
  assert.deepEqual(problems(s, 0), []);
  assert.deepEqual(problems(s, 1), ['Give the model a title.']);
  s.title = 'Operating budget FY2027';
  s.lastActual = 13;
  assert.match(problems(s, 2)[0], /0 \(none\) to 12/);
  s.lastActual = 6;
  s.entityName = '  ';
  assert.deepEqual(problems(s, STEPS.length - 1), ['Type the entity\'s name.']);
});

test('a blank model is the frame: contents, a cover, Settings and Lookups', () => {
  const s = { ...initialState(today), title: 'Blank test', recipe: 'blank' };
  const p = preview(s, lib);
  assert.deepEqual(p.layout.sheets.map(([n]) => n), ['Contents', 'Appendices', 'Settings', 'Lookups']);
  assert.equal(p.timeline, '12 months, April 2026 to March 2027; financial year ends in March; no actuals yet; budget April 2026 to March 2027; in $.');
});

test('the demo recipe brings its modules and the Settings values come from the answers', () => {
  const s = { ...initialState(today), title: 'Demo', recipe: 'demo', brand: 'TWK' as const, entityName: 'Te Wawata Kāinga', months: 24, lastActual: 6 };
  const m = toModel(s, lib);
  assert.equal(m.instances.length, 8);
  assert.equal(m.info?.timeline.lastActual, 6);
  assert.equal(periodMonth(s.start, 6), 'September 2026');
  const p = preview(s, lib);
  assert.deepEqual(p.sections.map(x => x.title), ['Dashboards', 'Financial Model', 'Appendices']);
  assert.match(p.timeline, /actuals to September 2026/);
});

test('Create builds a workbook with the answers in it', () => {
  const s = { ...initialState(today), title: 'Operating budget FY2027', recipe: 'demo', brand: 'KM' as const, entityName: 'Kāinga Maha',
    notes: 'First note\n\nSecond note' };
  const bytes = buildFile(s, lib, { KM: loadLogo('KM') }, new Date('2026-10-07T00:00:00Z'));
  const parts = unzipSync(bytes);
  const wb = strFromU8(parts['xl/workbook.xml']);
  assert.match(wb, /<sheet name="Settings"/);
  assert.match(wb, /<definedName name="Model_Title">Settings!\$I\$6<\/definedName>/);
  const meta = strFromU8(parts['customXml/item1.xml']);
  assert.match(meta, /Operating budget FY2027/);
  assert.match(meta, /Kāinga Maha/);
  assert.ok(parts['xl/media/logo.png']);
  assert.throws(() => buildFile({ ...s, title: '' }, lib, {}), /Give the model a title/);
});

test('file names and base64', () => {
  assert.equal(fileName({ ...initialState(today), title: 'Budget: FY2027 / v2' }), 'Budget FY2027 v2.xlsx');
  assert.equal(fileName({ ...initialState(today), title: '' }), 'HFG model.xlsx');
  const bytes = new Uint8Array(70000).map((_, i) => i % 251);
  assert.deepEqual(new Uint8Array(Buffer.from(toBase64(bytes), 'base64')), bytes);
});
