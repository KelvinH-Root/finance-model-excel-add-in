// The model's properties (Model > Properties): its title, the Prepared by line, the notes on the
// contents and what the model name line shows, with what saving them changes; beside them what
// the model was built with (entity, timeline, library) and its modules with the version each was
// built from. The title, Prepared by and the status line are also inputs on the Settings sheet.

import type { Preview, PropertiesInfo, ServiceApi } from '../service.ts';
import { frame, h, useCss, type Exit } from '../ui/kit.ts';

type Svc = Pick<ServiceApi, 'propertiesState' | 'propertiesPreview' | 'propertiesApply'>;

const CSS = `
.p-dl { display: grid; grid-template-columns: max-content 1fr; gap: 3px 12px; margin: 4px 0 12px; }
.p-dl dt { color: var(--quiet); }
.p-dl dd { margin: 0; }
.p-left { padding: 12px 14px; }
.p-mods td.v { white-space: nowrap; text-align: right; color: var(--quiet); }
.p-check { display: flex; align-items: center; gap: 6px; margin: 0 0 6px; font-weight: 400; }
textarea.k-inp { min-height: 76px; resize: vertical; }
`;

export function mountProperties(host: HTMLElement, svc: Svc, exit: Exit): void {
  useCss('hfg-props-css', CSS);
  let info: PropertiesInfo | null = null;
  let edit = { title: '', preparedBy: '', notes: [] as string[], display: { errors: true, alerts: true } };
  let preview: Preview | null = null;
  let message = 'Reading the model...';
  let bad = false;
  let busy = false;
  let done = false;
  let timer: ReturnType<typeof setTimeout> | null = null;
  let ask = 0;
  let card: HTMLElement | null = null;

  async function load() {
    try {
      info = await svc.propertiesState();
      edit = { title: info.title, preparedBy: info.preparedBy, notes: [...info.notes], display: { ...info.display } };
      message = `Read from ${info.source}.`;
      bad = false;
      done = false;
      preview = null;
    } catch (e) {
      message = (e as Error).message;
      bad = true;
    }
    render();
  }

  const clean = (notes: string[]) => notes.map(n => n.trim()).filter(Boolean).join('\n');
  const unchanged = () => !!info && edit.title.trim() === info.title && edit.preparedBy.trim() === info.preparedBy
    && clean(edit.notes) === clean(info.notes) && edit.display.errors === info.display.errors && edit.display.alerts === info.display.alerts;

  function schedule(ms = 450) {
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => void refresh(), ms);
  }

  async function refresh() {
    if (!info) return;
    const mine = ++ask;
    if (unchanged()) {
      preview = null;
      message = `Read from ${info.source}.`;
      bad = false;
      return update();
    }
    message = 'Working out what it changes...';
    bad = false;
    update();
    try {
      const p = await svc.propertiesPreview(edit);
      if (mine !== ask) return;
      preview = p;
      message = 'Undo in Excel (Ctrl+Z) takes back the change.';
    } catch (e) {
      if (mine !== ask) return;
      preview = null;
      message = (e as Error).message;
      bad = true;
    }
    update();
  }

  async function save() {
    if (!preview || busy) return;
    if (timer) { clearTimeout(timer); timer = null; }
    busy = true;
    message = 'Saving...';
    update();
    try {
      message = (await svc.propertiesApply()).message;
      bad = false;
      done = true;
    } catch (e) {
      message = `Stopped: ${(e as Error).message}. Undo in Excel (Ctrl+Z) takes back what was done.`;
      bad = true;
    }
    busy = false;
    render();
  }

  function buttons(): HTMLElement[] {
    if (done || !info) return [h('button', { type: 'button', class: 'primary', onclick: () => exit.close() }, 'Close')];
    return [h('button', { type: 'button', onclick: () => exit.close() }, 'Cancel'),
      h('button', { type: 'button', class: 'primary', disabled: !preview || busy, onclick: () => void save() }, 'Save')];
  }

  function update() {
    if (card) {
      card.innerHTML = '';
      card.hidden = !preview;
      if (preview) card.append(h('strong', {}, 'What saving changes'), h('ul', {}, preview.lines.map(l => h('li', {}, l))));
    }
    const foot = host.querySelector('.dlg-foot');
    if (!foot) return;
    const msg = foot.querySelector('.msg') as HTMLElement;
    msg.textContent = message;
    msg.className = 'msg' + (bad ? ' bad' : '');
    const btns = foot.querySelector('.btns') as HTMLElement;
    btns.innerHTML = '';
    btns.append(...buttons());
  }

  function about(i: PropertiesInfo): HTMLElement {
    const box = h('div', { class: 'p-left' });
    box.append(h('p', { class: 'k-title' }, 'Built with'));
    box.append(h('dl', { class: 'p-dl' }, i.facts.map(([k, v]) => [h('dt', {}, k), h('dd', {}, v)])));
    const outdated = i.modules.filter(m => m.held < m.library).length;
    box.append(h('p', { class: 'k-title' }, `Modules (${i.modules.length})`));
    box.append(h('table', { class: 'k-table p-mods' },
      h('thead', {}, h('tr', {}, h('th', {}, 'Module'), h('th', { style: 'text-align:right' }, 'Version'))),
      h('tbody', {}, i.modules.map(m => h('tr', {},
        h('td', {}, m.title, m.count > 1 ? h('span', { class: 'k-quiet' }, ` × ${m.count}`) : null),
        h('td', { class: 'v' }, m.held < m.library ? h('span', { class: 'k-pill' }, `${m.held}, library has ${m.library}`) : String(m.held)))))));
    box.append(h('p', { class: 'k-note' }, outdated
      ? `${outdated} module${outdated > 1 ? 's were' : ' was'} built from an older library version. Bringing a module up to the library's version is not built yet.`
      : 'Every module is at the library\'s version.'));
    return box;
  }

  function form(): HTMLElement {
    const box = h('div', {});
    const text = (label: string, value: string, set: (v: string) => void, max: number) =>
      h('label', { class: 'k-field' }, h('span', {}, label), h('input', {
        class: 'k-inp', type: 'text', value, maxlength: max, 'aria-label': label,
        oninput: (e: Event) => { set((e.target as HTMLInputElement).value); schedule(); },
      }));
    box.append(text('Model title', edit.title, v => { edit.title = v; }, 80));
    box.append(text('Prepared by line', edit.preparedBy, v => { edit.preparedBy = v; }, 120));
    box.append(h('label', { class: 'k-field' }, h('span', {}, 'Notes on the contents (one per line)'), h('textarea', {
      class: 'k-inp', 'aria-label': 'Notes on the contents',
      oninput: (e: Event) => { edit.notes = (e.target as HTMLTextAreaElement).value.split('\n'); schedule(); },
    }, edit.notes.join('\n'))));
    const tick = (label: string, key: 'errors' | 'alerts') => h('label', { class: 'p-check' }, h('input', {
      type: 'checkbox', checked: edit.display[key],
      onchange: (e: Event) => { edit.display[key] = (e.target as HTMLInputElement).checked; schedule(0); },
    }), label);
    box.append(h('div', { class: 'k-field' }, h('span', {}, 'Model name line on every sheet'),
      tick('Show the error count', 'errors'), tick('Show the alert count', 'alerts')));
    card = h('div', { class: 'k-card', hidden: true });
    box.append(card);
    return box;
  }

  function render() {
    card = null;
    let body: HTMLElement;
    if (info && !done) {
      body = h('div', { class: 'dlg-cols' }, h('div', { class: 'left' }, about(info)), h('div', { class: 'right' }, form()));
    } else {
      body = h('div', { class: 'dlg-body' }, h('p', {}, message));
    }
    frame(host, body, done || !info ? '' : message, bad && !!info, buttons());
    update();
  }

  render();
  void load();
}
