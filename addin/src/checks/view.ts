// Checks, in the task pane (it stays beside the sheet while the model is worked on): how many
// errors and alerts are raised, then every check by module, failing ones first, each with Show to
// select its row on the sheet. Refresh reads them again after a change.

import type { ChecksInfo, ServiceApi } from '../service.ts';
import { h, useCss, CSS as KIT } from '../ui/kit.ts';

type Svc = Pick<ServiceApi, 'checksState' | 'selectRow' | 'select'>;
type Filter = 'failing' | 'error' | 'alert' | 'all';

const CSS = `
.c-tiles { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin: 6px 0 10px; }
.c-tile { border: 1px solid var(--line); border-radius: 6px; padding: 7px 9px; }
.c-tile b { display: block; font-size: 19px; font-weight: 600; line-height: 1.1; }
.c-tile span { font-size: 11px; color: var(--quiet); }
.c-chips { display: flex; flex-wrap: wrap; gap: 4px; margin: 0 0 8px; }
.c-chips button { border-radius: 12px; font-size: 11px; padding: 2px 10px; margin: 0; }
.c-chips button.on { background: var(--accent); border-color: var(--accent); color: var(--on-accent); }
.c-grp { margin: 0 0 8px; }
.c-gh { display: flex; justify-content: space-between; gap: 6px; font-weight: 600; padding: 5px 0 3px; border-bottom: 1px solid var(--line); }
.c-gh a, .c-ck a { font-weight: 400; font-size: 11px; white-space: nowrap; color: var(--accent-text); }
.c-ck { display: flex; gap: 8px; align-items: flex-start; padding: 5px 0; border-bottom: 1px solid var(--line); }
.c-ck .k-stat { margin-top: 4px; }
.c-ck .t { flex: 1; line-height: 1.35; }
.c-ck .t small { display: block; color: var(--quiet); font-size: 11px; }
`;

export function mountChecks(host: HTMLElement, svc: Svc): void {
  useCss('hfg-kit-css', KIT);
  useCss('hfg-checks-css', CSS);
  let info: ChecksInfo | null = null;
  let filter: Filter = 'failing';
  let message = 'Reading the checks...';
  let at = '';

  async function load() {
    message = 'Reading the checks...';
    render();
    try {
      info = await svc.checksState();
      at = new Date().toLocaleTimeString('en-NZ', { hour: '2-digit', minute: '2-digit' });
      message = info.read ? `Read from ${info.source} at ${at}.` : `Read from ${info.source}; whether each is raised shows only in Excel.`;
      if (info.read && !info.errors && !info.alerts) filter = 'all';
    } catch (e) {
      message = (e as Error).message;
    }
    render();
  }

  function render() {
    host.innerHTML = '';
    const top = h('div', { style: 'display:flex;justify-content:space-between;align-items:center;gap:8px' },
      h('span', { class: 'k-quiet', role: 'status' }, message),
      h('button', { type: 'button', onclick: () => void load() }, 'Refresh'));
    host.append(top);
    if (!info) return;
    const total = info.groups.reduce((n, g) => n + g.checks.length, 0);
    host.append(h('div', { class: 'c-tiles' },
      h('div', { class: 'c-tile' }, h('b', { style: 'color:var(--fail)' }, info.read ? String(info.errors) : '-'), h('span', {}, 'Errors')),
      h('div', { class: 'c-tile' }, h('b', { style: 'color:var(--warn)' }, info.read ? String(info.alerts) : '-'), h('span', {}, 'Alerts')),
      h('div', { class: 'c-tile' }, h('b', { style: 'color:var(--accent-text)' }, info.read ? String(info.passing) : '-'), h('span', {}, 'Passing'))));
    const chips: [Filter, string][] = [['failing', `Failing (${info.errors + info.alerts})`], ['error', 'Errors'], ['alert', 'Alerts'], ['all', `All (${total})`]];
    host.append(h('div', { class: 'c-chips', role: 'radiogroup', 'aria-label': 'Show' }, chips.map(([k, l]) =>
      h('button', { type: 'button', role: 'radio', 'aria-checked': String(filter === k), class: filter === k ? 'on' : null, onclick: () => { filter = k; render(); } }, l))));
    const keep = (c: ChecksInfo['groups'][number]['checks'][number]) =>
      filter === 'all' || (!info!.read ? true : c.raised && (filter === 'failing' || c.kind === filter));
    let shown = 0;
    for (const g of info.groups) {
      const list = g.checks.filter(keep);
      if (!list.length) continue;
      shown += list.length;
      host.append(h('div', { class: 'c-grp' },
        h('div', { class: 'c-gh' }, h('span', {}, g.title), h('a', { href: '#', onclick: (e: Event) => { e.preventDefault(); void svc.select(g.goName); } }, 'Go to module')),
        list.map(c => h('div', { class: 'c-ck' },
          h('span', { class: 'k-stat ' + (c.raised === null ? '' : c.raised ? c.kind : 'ok') }),
          h('span', { class: 't' }, c.label, h('small', {}, `${c.kind === 'error' ? 'Error' : 'Alert'} · ${g.sheet}${c.raised === null ? '' : c.raised ? ' · raised' : ' · clear'}`)),
          h('a', { href: '#', onclick: (e: Event) => { e.preventDefault(); void svc.selectRow(c.row); } }, 'Show')))));
    }
    if (!shown) host.append(h('p', { class: 'k-quiet' }, filter === 'failing' ? 'No check is failing.' : 'Nothing to show.'));
    host.append(h('p', { class: 'k-note' }, 'Error and alert counts also show in the model name on every sheet. Show selects the check\'s row.'));
  }

  void load();
}
