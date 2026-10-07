// Build a model from a recipe with the package writer.
//
//   node tools/build-recipe.ts library/hfg/recipes/full_model.json out.xlsx
//
// Fictional demo data only.

import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { assemble, buildWorkbook, Model, type Recipe } from '../src/index.ts';
import { loadLogo } from '../src/node/assets.ts';
import { loadLibrary } from '../src/node/library.ts';

const [recipePath, out = 'build/model.xlsx'] = process.argv.slice(2);
if (!recipePath) throw new Error('usage: node tools/build-recipe.ts <recipe.json> [out.xlsx]');
const recipe = JSON.parse(readFileSync(recipePath, 'utf8')) as Recipe;
const lib = loadLibrary(join(dirname(recipePath), '..'));
const model = Model.fromRecipe(lib, recipe);
const layout = assemble(model);
const bytes = buildWorkbook(layout, model, { logo: loadLogo(recipe.entity.brand), created: new Date('2026-10-07T00:00:00Z') });
writeFileSync(out, bytes);
console.log(`wrote ${out} (${bytes.length} bytes, ${layout.sheets.length} sheets${layout.warnings.length ? `; warnings: ${layout.warnings.join('; ')}` : ''})`);
