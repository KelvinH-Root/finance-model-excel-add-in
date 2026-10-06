// The demo model the writer tests and tools/build.ts use: the assembly proof's base model, with
// New model's choices when a brand is given. Fictional demo data only.

import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Model, type Brand, type Library, type ModelInfo } from '../src/index.ts';
import { THEMES } from '../src/theme.ts';

export const LIBRARY = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'prototypes', 'assembly', 'library');

export function demoInfo(brand: Brand): ModelInfo {
  return {
    title: 'Demo operating model', entity: { name: THEMES[brand].name, brand }, preparedBy: 'Prepared by Group Finance',
    notes: ['Fictional demo data, built by the HFG model engine.', 'Inputs are the shaded cells; everything else is calculated.'],
    timeline: { start: '2026-04', fyEndMonth: 3, lastActual: 0, denomination: '$' },
    display: { errors: true, alerts: true },
  };
}

export function demoModel(lib: Library, brand: Brand | null, periods = 12): Model {
  const m = new Model(lib, periods);
  if (brand) m.info = demoInfo(brand);
  m.insert('demo.statements');
  m.insert('demo.checks');
  m.insert('demo.revenue_line', { base: 100, growth: 0.01 });
  m.insert('demo.revenue_line', { base: 60, growth: 0.03 });
  m.insert('demo.cost_line', { amount: 90, inflation: 0.002 });
  m.insert('demo.debtors');
  m.insert('demo.facility', { amount: 1000, rate: 0.06, instalment: 50 });
  m.insert('demo.dashboard', { first: 1 });
  return m;
}
