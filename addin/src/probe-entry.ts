// The Phase 1 add-in features, bundled for the Phase 0 probe. The ribbon decides where each
// command opens (addin-probe/src/commands.js gives each its surface): the add-in's own views
// (New model, Insert module, Versions, Explorer, Remove, Rename, Delete, Properties) open in
// dialogs that reach the workbook through the model service in the runtime; Add category, Add
// above and Duplicate run on the selection with no window.
// Where dialogs are not available the same views draw in the task pane (window.HfgWizard and the
// rest, the hooks the probe's pane already calls).

import bundle from 'virtual:library';
import logoData from 'virtual:logos';
import { assemble, Library, Model, type Brand } from '../../engine/src/index.ts';
import { mountChecks } from './checks/view.ts';
import { mountExplorer } from './explorer/view.ts';
import { mountInsert } from './insert/view.ts';
import { mountProperties } from './properties/view.ts';
import { mountRemove } from './remove/view.ts';
import { mountRename } from './rename/view.ts';
import { askClose, askOpen, closeDialog, DIALOG_SIZES, dialogsSupported, openDialog, remote, tellDone } from './rpc.ts';
import { ModelService, type ServiceApi } from './service.ts';
import type { Exit } from './ui/kit.ts';
import { mountVersions } from './versions/view.ts';
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

/** The budget model demo, for the versions view outside Excel. */
function sampleFull() {
  const r = bundle.recipes.find(x => x.id === 'budget_model') ?? bundle.recipes[0];
  const model = Model.fromRecipe(libs[r.library ?? 'hfg'], r.model);
  return { model, layout: assemble(model) };
}

const service = new ModelService({ libs, logos, sample, sampleFull });

/** The add-in view a command opens, by command key. */
const VIEWS: Record<string, string> = {
  'model-new': 'wizard', 'mod-insert': 'insert', 'model-explorer': 'explorer', 'b-explorer': 'explorer', 'mod-links': 'explorer',
  'an-ver-save': 'versions', 'an-ver-approve': 'versions', 'an-ver-lock': 'versions', 'an-ver-manage': 'versions', 'cat-remove': 'remove',
  'model-checks': 'checks', 'mod-rename': 'rename', 'mod-delete': 'delete', 'model-props': 'properties',
};

/** Dialog sizes for the probe's own designed views and the generic pages. */
const OTHER_SIZES: Record<string, string> = {
  'grp-structure': 'group', 'model-group': 'group', 'grp-add': 'entity', 'grp-ownership': 'entity', 'grp-remove': 'entity',
  'imp-live': 'impacts', 'imp-sheets': 'impacts', 'ctx-impacts': 'impacts', 'an-impacts': 'impacts', message: 'message',
};

/** Draw a command's add-in view into host; false when the command has none here. */
function mount(key: string, host: HTMLElement, svc: ServiceApi, exit: Exit): boolean {
  switch (VIEWS[key]) {
    case 'wizard':
      mountWizard(host, {
        lib: libs, logos, framed: { cancel: () => exit.close() },
        create: async s => {
          const r = await svc.createModel(s);
          if (r.opened) setTimeout(() => exit.close(), 1800);
          return r;
        },
      });
      return true;
    case 'insert': mountInsert(host, svc, exit); return true;
    case 'versions': mountVersions(host, svc, exit); return true;
    case 'explorer': mountExplorer(host, svc, exit, key === 'mod-links' ? 'links' : 'composition'); return true;
    case 'remove': mountRemove(host, svc, exit); return true;
    case 'delete': mountRemove(host, svc, exit, 'delete'); return true;
    case 'rename': mountRename(host, svc, exit); return true;
    case 'properties': mountProperties(host, svc, exit); return true;
    case 'checks': mountChecks(host, svc); return true;
    default: return false;
  }
}

/** Open a command's dialog (from the ribbon, a menu, the right-click menu or another dialog). */
async function open(key: string, arg?: string): Promise<void> {
  const size = DIALOG_SIZES[VIEWS[key] ?? OTHER_SIZES[key] ?? 'notbuilt'];
  const url = new URL(`hfg-dialog.html?view=${encodeURIComponent(key)}${arg ? `&arg=${encodeURIComponent(arg)}` : ''}`, location.href).href;
  await openDialog(url, size, {
    call: (m, a) => {
      const fn = (service as unknown as Record<string, unknown>)[m];
      if (typeof fn !== 'function' || m === 'constructor' || m === 'load') throw new Error(`The add-in has no ${m}.`);
      return (fn as (...x: unknown[]) => Promise<unknown>).apply(service, a);
    },
    open: (k, a) => void open(k, a),
  });
}

/** The commands that run on the selection with no window, by command key. */
const ON_SELECTION: Record<string, [string, () => Promise<{ ok: boolean; message: string }>]> = {
  'cat-add': ['Add category', () => service.addCategoryAtSelection()],
  'cat-above': ['Add above', () => service.addAboveAtSelection()],
  'mod-duplicate': ['Duplicate', () => service.duplicateAtSelection()],
};

/** Run a selection command (no window). A problem (nothing selected to act on) opens a short message instead. */
async function onSelection(key: string): Promise<{ ok: boolean; message: string }> {
  const cmd = ON_SELECTION[key];
  if (!cmd) return { ok: false, message: `No selection command ${key}.` };
  let r: { ok: boolean; message: string };
  try {
    r = await cmd[1]();
  } catch (e) {
    r = { ok: false, message: `${cmd[0]} stopped: ${(e as Error).message}. Undo in Excel (Ctrl+Z) takes back what was written.` };
  }
  if (!r.ok && dialogsSupported()) await open('message', r.message);
  return r;
}

const addCategory = () => onSelection('cat-add');

const paneExit = (show?: (k: string | null) => void): Exit => ({ close: () => show?.(null), open: k => show?.(k) });

type PaneView = { render(key: string, host: HTMLElement, show?: (k: string | null) => void): boolean };
const w = window as unknown as { HfgAddinViews: PaneView; HfgAddin: Record<string, unknown> };

// In the task pane (Checks, and every view where dialogs are not available): the same views, on the service directly.
w.HfgAddinViews = { render: (key, host, show) => key in VIEWS && mount(key, host, service, paneExit(show)) };

w.HfgAddin = {
  service, libs, VIEWS,
  hasView: (key: string) => key in VIEWS,
  runsOnSelection: (key: string) => key in ON_SELECTION,
  mount, open, onSelection, addCategory, dialogsSupported, closeDialog,
  remote: () => remote<ServiceApi>(), askClose, askOpen, tellDone,
};
