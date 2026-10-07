// The Phase 1 add-in features, bundled for the Phase 0 probe: its ribbon commands draw them in the
// task pane through the same hooks as the probe's other designed views (window.HfgWizard for
// New model, window.HfgInsert for Modules > Insert).

import bundle from 'virtual:library';
import logoData from 'virtual:logos';
import { assemble, Library, type Brand } from '../../engine/src/index.ts';
import { mountInsert } from './insert/view.ts';
import { addRecipes, initialState, toModel, type Libraries } from './wizard/core.ts';
import { mountWizard, type WizardAssets } from './wizard/view.ts';

const fromBase64 = (b64: string) => Uint8Array.from(atob(b64), c => c.charCodeAt(0));

const libs: Libraries = Object.fromEntries(Object.entries(bundle.libraries).map(([id, b]) => [id, Library.fromBundle(b)]));
addRecipes(bundle.recipes.map(r => ({ ...r, modules: [] })));
const logos: WizardAssets['logos'] = {};
for (const [brand, l] of Object.entries(logoData)) {
  logos[brand as Brand] = { png: fromBase64(l.base64), width: l.width, height: l.height, dataUrl: `data:image/png;base64,${l.base64}` };
}

/** The demo model, for running the views outside Excel. */
function sample() {
  const model = toModel({ ...initialState(new Date()), title: 'Demo operating model', recipe: 'demo' }, libs);
  return { model, layout: assemble(model) };
}

type View = { render(key: string, host: HTMLElement): boolean };
const w = window as unknown as { HfgWizard: View; HfgInsert: View };

w.HfgWizard = {
  render(key, host) {
    if (key !== 'model-new') return false;
    mountWizard(host, { lib: libs, logos });
    return true;
  },
};

w.HfgInsert = {
  render(key, host) {
    if (key !== 'mod-insert') return false;
    mountInsert(host, { libs, sample });
    return true;
  },
};
