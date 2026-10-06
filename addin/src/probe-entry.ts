// The New model wizard, bundled for the Phase 0 probe: the probe's ribbon command "New model"
// draws it in the task pane (window.HfgWizard, the same hook as its other designed views).

import bundle from 'virtual:library';
import logoData from 'virtual:logos';
import { Library, type Brand } from '../../engine/src/index.ts';
import { mountWizard, type WizardAssets } from './wizard/view.ts';

const fromBase64 = (b64: string) => Uint8Array.from(atob(b64), c => c.charCodeAt(0));

const lib = Library.fromBundle(bundle);
const logos: WizardAssets['logos'] = {};
for (const [brand, l] of Object.entries(logoData)) {
  logos[brand as Brand] = { png: fromBase64(l.base64), width: l.width, height: l.height, dataUrl: `data:image/png;base64,${l.base64}` };
}

(window as unknown as { HfgWizard: unknown }).HfgWizard = {
  render(key: string, host: HTMLElement): boolean {
    if (key !== 'model-new') return false;
    mountWizard(host, { lib, logos });
    return true;
  },
};
