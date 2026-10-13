// Explorer, as a dialog: the open model's tree on the left (sections, sheets, categories and
// modules, coloured by area, with check status) and the module chosen on the right, with its
// Composition, Links (drawn), Properties and Checks. Choosing a module selects it in the workbook.

import type { ExplorerTree, ModuleInfo } from './core.ts';
import type { ServiceApi, Status } from '../service.ts';
import { frame, h, useCss, type Exit } from '../ui/kit.ts';

type Svc = Pick<ServiceApi, 'explorerTree' | 'explorerModule' | 'select'>;
type Tab = 'composition' | 'links' | 'properties' | 'checks';

// HF palette first, then neutrals: areas are told apart by name as well as colour.
const AREA_COLOURS = ['#679db5', '#09122c', '#90b6c8', '#566a89', '#3c4e60', '#cedce5', '#a6a6a6', '#404040'];
const SVG = 'http://www.w3.org/2000/svg';

const CSS = `
.x-tree { padding: 6px 0 10px; }
.x-sec { font-weight: 600; padding: 8px 12px 2px; }
.x-sheet { color: var(--accent-text); padding: 4px 12px 2px 22px; }
.x-row { display: flex; align-items: center; gap: 7px; width: 100%; margin: 0; padding: 3px 12px; border: 0; border-radius: 0; background: transparent;
  font: inherit; color: var(--ink); text-align: left; cursor: pointer; }
.x-row:hover { background: var(--panel); }
.x-row.sel { background: var(--accent-tint); box-shadow: inset 3px 0 0 var(--accent-line); }
.x-row .sq { width: 9px; height: 9px; border-radius: 2px; flex: none; }
.x-row .n { font-size: 11px; color: var(--quiet); }
.x-row .end { margin-left: auto; display: flex; align-items: center; gap: 6px; }
.x-cv { width: 10px; flex: none; color: var(--quiet); font-size: 9px; }
.x-head { display: flex; gap: 10px; align-items: flex-start; margin: 0 0 10px; }
.x-head .sq { width: 13px; height: 13px; border-radius: 3px; margin-top: 3px; flex: none; }
.x-tabs { display: flex; border-bottom: 1px solid var(--line); margin: 0 0 10px; }
.x-tabs button { border: 0; border-radius: 0; margin: 0; background: transparent; color: var(--quiet); padding: 6px 10px;
  border-bottom: 2px solid transparent; margin-bottom: -1px; }
.x-tabs button.on { color: var(--ink); font-weight: 600; border-bottom-color: var(--accent-line); }
.x-what { font-size: 11px; color: var(--quiet); white-space: nowrap; }
tr.x-h td { font-weight: 600; }
.x-links { max-width: 520px; }
.x-links rect.box { fill: var(--panel); stroke: var(--line); }
.x-links g.go { cursor: pointer; }
.x-links g.go:hover rect.box, .x-links g.go:focus rect.box { stroke: var(--accent-line); }
.x-links rect.self { fill: var(--accent); }
.x-links text { font: 12px 'Segoe UI', system-ui, sans-serif; fill: var(--ink); }
.x-links text.t { font-weight: 600; }
.x-links text.w { font-size: 11px; fill: var(--quiet); }
.x-links text.self { fill: var(--on-accent); font-weight: 600; }
.x-links text.lab { font-size: 11px; font-weight: 600; fill: var(--accent-text); }
.x-links path.arrow { stroke: var(--accent-text); fill: none; }
.x-props { display: grid; grid-template-columns: max-content 1fr; gap: 4px 14px; margin: 0 0 10px; }
.x-props dt { color: var(--quiet); }
.x-props dd { margin: 0; }
`;

function svg(tag: string, attrs: Record<string, unknown> = {}, ...kids: (Node | string)[]): SVGElement {
  const e = document.createElementNS(SVG, tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === null || v === undefined) continue;
    if (k.startsWith('on') && typeof v === 'function') e.addEventListener(k.slice(2), v as EventListener);
    else e.setAttribute(k, String(v));
  }
  kids.forEach(k => e.append(k));
  return e;
}

const clip = (s: string, n: number) => (s.length > n ? s.slice(0, n - 1) + '…' : s);

export function mountExplorer(host: HTMLElement, svc: Svc, exit: Exit, startTab: Tab = 'composition'): void {
  useCss('hfg-explorer-css', CSS);
  let tree: (ExplorerTree & { source: string; status: Record<string, Status> }) | null = null;
  let mod: (ModuleInfo & { raised: (boolean | null)[] }) | null = null;
  let tab: Tab = startTab;
  const open = new Set<string>();
  let message = 'Reading the model...';
  let bad = false;

  const colour = (area: string) => AREA_COLOURS[Math.max(0, tree?.areas.indexOf(area) ?? 0) % AREA_COLOURS.length];
  const dot = (s?: Status) => (s ? h('span', { class: 'k-stat ' + s, title: s === 'error' ? 'An error check is raised' : s === 'alert' ? 'An alert is raised' : 'Checks clear' }) : null);

  async function load() {
    try {
      tree = await svc.explorerTree();
      message = `${tree.summary}, read from ${tree.source}.`;
      bad = false;
      const first = tree.sections.flatMap(s => s.sheets.flatMap(sh => sh.entries)).find(e => e.kind === 'category') ?? tree.sections[0]?.sheets[0]?.entries[0];
      if (first) await pick(first.modules[0].uid, false);
    } catch (e) {
      message = (e as Error).message;
      bad = true;
    }
    render();
  }

  async function pick(uid: string, selectInWorkbook = true) {
    try {
      mod = await svc.explorerModule(uid);
      for (const s of tree?.sections ?? []) for (const sh of s.sheets) for (const e of sh.entries) if (e.modules.some(m => m.uid === uid)) open.add(e.id);
      if (selectInWorkbook) void svc.select(mod.goName);
    } catch (e) {
      message = (e as Error).message;
      bad = true;
    }
    render();
  }

  function treeView(): HTMLElement {
    const box = h('div', { class: 'left' });
    const t = h('div', { class: 'x-tree', role: 'tree', 'aria-label': 'Model' });
    tree?.sections.forEach((sec, i) => {
      t.append(h('div', { class: 'x-sec' }, `${i + 1}. ${sec.title}`));
      sec.sheets.forEach(sh => {
        t.append(h('div', { class: 'x-sheet' }, sh.sheet));
        for (const e of sh.entries) {
          const sel = (uid: string) => mod?.uid === uid;
          if (e.kind === 'single') {
            const m = e.modules[0];
            t.append(h('button', { class: 'x-row' + (sel(m.uid) ? ' sel' : ''), role: 'treeitem', 'aria-selected': String(sel(m.uid)), style: 'padding-left:32px', onclick: () => void pick(m.uid) },
              h('span', { class: 'sq', style: `background:${colour(m.area)}` }), m.title, h('span', { class: 'end' }, dot(tree!.status[m.uid]))));
            continue;
          }
          const isOpen = open.has(e.id);
          const worst = e.modules.map(m => tree!.status[m.uid]).reduce<Status | undefined>((w, s) => (s === 'error' || w === 'error' ? 'error' : s === 'alert' || w === 'alert' ? 'alert' : s ?? w), undefined);
          t.append(h('button', { class: 'x-row', role: 'treeitem', 'aria-expanded': String(isOpen), style: 'padding-left:22px',
            onclick: () => { if (isOpen) open.delete(e.id); else open.add(e.id); render(); } },
          h('span', { class: 'x-cv', 'aria-hidden': 'true' }, isOpen ? '▾' : '▸'), h('span', { class: 'sq', style: `background:${colour(e.area)}` }), e.title,
          h('span', { class: 'n' }, String(e.modules.length)), h('span', { class: 'end' }, dot(worst))));
          if (isOpen) {
            for (const m of e.modules) {
              t.append(h('button', { class: 'x-row' + (sel(m.uid) ? ' sel' : ''), role: 'treeitem', 'aria-selected': String(sel(m.uid)), style: 'padding-left:54px', onclick: () => void pick(m.uid) },
                m.title, h('span', { class: 'end' }, dot(tree!.status[m.uid]))));
            }
          }
        }
      });
    });
    box.append(t);
    return box;
  }

  function links(m: ModuleInfo): SVGElement {
    const W = 460, H = 38, GAP = 8;
    let y = 4;
    const kids: SVGElement[] = [];
    const label = (text: string) => { kids.push(svg('text', { x: 4, y: y + 11, class: 'lab' }, text)); y += 18; };
    const box = (r: { uid: string; title: string; carries: string[] }) => {
      const words = clip(r.carries.join(', '), 70);
      kids.push(svg('g', { class: 'go', tabindex: 0, role: 'button', 'aria-label': `${r.title}: ${r.carries.join(', ')}`,
        onclick: () => void pick(r.uid), onkeydown: (ev: KeyboardEvent) => { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); void pick(r.uid); } } },
      svg('rect', { x: 4, y, width: W - 8, height: H, rx: 5, class: 'box' }),
      svg('text', { x: 14, y: y + 15, class: 't' }, clip(r.title, 60)),
      svg('text', { x: 14, y: y + 30, class: 'w' }, words)));
      y += H + GAP;
    };
    const arrow = () => { kids.push(svg('path', { d: `M${W / 2} ${y} V${y + 20} M${W / 2 - 5} ${y + 14} L${W / 2} ${y + 20} L${W / 2 + 5} ${y + 14}`, class: 'arrow' })); y += 26; };
    label(m.takes.length ? 'Takes from' : 'Takes nothing from other modules');
    m.takes.forEach(box);
    if (m.takes.length) arrow();
    kids.push(svg('rect', { x: 60, y, width: W - 120, height: H, rx: 6, class: 'self' }));
    kids.push(svg('text', { x: W / 2, y: y + 24, class: 'self', 'text-anchor': 'middle' }, clip(m.title, 50)));
    y += H + 4;
    if (m.sends.length) arrow();
    label(m.sends.length ? 'Sends to' : 'Sends nothing to other modules');
    m.sends.forEach(box);
    return svg('svg', { viewBox: `0 0 ${W} ${y}`, width: '100%', role: 'img', class: 'x-links', 'aria-label': `Links of ${m.title}` }, ...kids);
  }

  function moduleView(): HTMLElement {
    const box = h('div', { class: 'right' });
    if (!mod) return box;
    const m = mod;
    box.append(h('div', { class: 'x-head' }, h('span', { class: 'sq', style: `background:${colour(m.area)}` }),
      h('div', {}, h('p', { class: 'k-title' }, m.title), h('p', { class: 'k-quiet', style: 'margin:0' }, `${m.kind}; ${m.where}`))));
    const tabs: [Tab, string][] = [['composition', 'Composition'], ['links', 'Links'], ['properties', 'Properties'], ['checks', `Checks (${m.checks.length})`]];
    box.append(h('div', { class: 'x-tabs', role: 'tablist' }, tabs.map(([k, l]) =>
      h('button', { type: 'button', role: 'tab', class: tab === k ? 'on' : null, 'aria-selected': String(tab === k), onclick: () => { tab = k; render(); } }, l))));
    if (tab === 'composition') {
      for (const c of m.components) {
        if (m.components.length > 1) box.append(h('p', { class: 'k-title', style: 'font-size:12px' }, c.sheet));
        box.append(h('table', { class: 'k-table' }, c.rows.map(r => h('tr', { class: r.head ? 'x-h' : null },
          h('td', { class: 'n' }, String(r.row)), h('td', {}, r.label), h('td', { class: 'x-what' }, r.head ? '' : r.what)))));
      }
      box.append(h('p', { class: 'k-note' }, 'Workings are grouped out of view on the sheet. Choosing a module selects its heading in the workbook.'));
    } else if (tab === 'links') {
      box.append(links(m));
      box.append(h('p', { class: 'k-note' }, 'Select a box to move to that module. Each link is named by the rows that carry it.'));
    } else if (tab === 'properties') {
      box.append(h('dl', { class: 'x-props' },
        h('dt', {}, 'Library module'), h('dd', {}, `${m.module}, version ${m.version}`),
        h('dt', {}, 'Instance'), h('dd', {}, `${m.uid} (number ${m.number}, never reused)`),
        h('dt', {}, 'Area'), h('dd', {}, m.area),
        h('dt', {}, 'Go to'), h('dd', {}, m.goName)));
      if (m.description) box.append(h('p', { class: 'k-quiet' }, m.description));
      if (m.settings.length) {
        box.append(h('table', { class: 'k-table' }, h('tr', {}, h('th', {}, 'Setting'), h('th', {}, 'Value')),
          m.settings.map(s => h('tr', {}, h('td', {}, s.label), h('td', {}, s.value)))));
      } else box.append(h('p', { class: 'k-quiet' }, 'No settings: everything on this module is worked out from its links.'));
    } else {
      if (!m.checks.length) box.append(h('p', { class: 'k-quiet' }, 'No checks of its own; what it sends is checked where it lands.'));
      m.checks.forEach((c, k) => {
        const r = m.raised[k];
        const s: Status | undefined = r === null ? undefined : r ? (c.kind === 'error' ? 'error' : 'alert') : 'ok';
        box.append(h('p', { style: 'display:flex;gap:8px;align-items:center' }, dot(s) ?? h('span', { class: 'k-stat' }),
          `${c.label} (${c.kind})`, h('span', { class: 'k-quiet' }, r === null ? '' : r ? 'raised' : 'clear')));
      });
    }
    return box;
  }

  function render() {
    frame(host, h('div', { class: 'dlg-cols' }, treeView(), moduleView()), message, bad,
      [h('button', { type: 'button', class: 'primary', onclick: () => exit.close() }, 'Close')]);
  }

  render();
  void load();
}
