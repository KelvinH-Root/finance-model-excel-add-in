// Build the demo model with the package writer.
//
//   node tools/build.ts [out.xlsx] [--brand HF|HCP|HCL|KM|TWK] [--proof]
//
// --proof builds it in the Phase 0 proof's frame, for comparison with the Python writer.
// Fictional demo data only.

import { writeFileSync } from 'node:fs';
import { assemble, buildWorkbook, Library, Model, type Brand } from '../src/index.ts';
import { loadLogo } from '../src/node/assets.ts';
import { demoModel, LIBRARY } from '../test/demo.ts';
import { loadLibrary } from '../src/node/library.ts';

const args = process.argv.slice(2);
const out = args.find(a => a.endsWith('.xlsx')) ?? 'build/demo.xlsx';
const brand = (args[args.indexOf('--brand') + 1] ?? 'HF') as Brand;
const proof = args.includes('--proof');
const lib = loadLibrary(LIBRARY);
const model: Model = demoModel(lib, proof ? null : brand);
const layout = assemble(model);
const bytes = buildWorkbook(layout, model, { logo: proof ? undefined : loadLogo(brand), created: new Date('2026-10-07T00:00:00Z') });
writeFileSync(out, bytes);
console.log(`wrote ${out} (${bytes.length} bytes, ${layout.sheets.length} sheets)`);
