// The model's properties (Model > Properties): its title, the Prepared by line, the notes on the
// contents and what the model name line shows, with what saving them changes; beside them what
// the model was built with (entity, timeline, library) and its modules with the version each was
// built from. The title, Prepared by and the status line are also inputs on the Settings sheet.
// Below them, the change log's newest entries (what each change did to the key outputs) and, when
// a change stopped part way, the ways to put the workbook right.

import type { Preview, PropertiesInfo, ServiceApi } from '../service.ts';
import { frame, h, useCss, type Exit } from '../ui/kit.ts';

type Svc = Pick<ServiceApi, 'propertiesState' | 'propertiesPreview' | 'propertiesApply' | 'restoreCopy' | 'markRepaired'>;

const CSS = `
.p-dl { display: grid; grid-template-columns: max-content 1fr; gap: 3px 12px; margin: 4px 0 12px; }
.p-dl dt { color: var(--quiet); }
.p-dl dd { margin: 0; }
.p-left { padding: 12px 14px; }
.p-mods td.v { white-space: nowrap; text-align: right; color: var(--quiet); }
.p-check { display: flex; align-items: center; gap: 6px; margin: 0 0 6px; font-weight: 400; }
.p-log { list-style: none; margin: 4px 0 0; padding: 0; }
.p-log li { margin: 0 0 6px; }
.p-log small { display: block; color: var(--quiet); }
.p-broken { border-color: var(--bad, #CB2840); }
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
      message = `Stopped: ${(e as Error).message}`;
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

  async function recover(how: 'restoreCopy' | 'markRepaired') {
    busy = true;
    try {
      message = (await svc[how]()).message;
      bad = false;
      if (how === 'markRepaired') { busy = false; await load(); message = 'Marked as put right. Check Model > Checks before going on.'; }
    } catch (e) {
      message = (e as Error).message;
      bad = true;
    }
    busy = false;
    update();
  }

  function stopped(i: PropertiesInfo): HTMLElement | null {
    if (!i.broken) return null;
    const when = new Date(i.broken.at).toLocaleString('en-NZ', { dateStyle: 'medium', timeStyle: 'short' });
    return h('div', { class: 'k-card p-broken', role: 'alert', style: 'margin-top:0' },
      h('strong', {}, 'A change stopped part way'),
      h('p', {}, `${i.broken.title} stopped at ${when}, so this workbook no longer matches its model. The add-in will not change it until it is put right.`),
      h('div', { style: 'display:flex;gap:8px;flex-wrap:wrap' },
        i.broken.copy ? h('button', { type: 'button', disabled: busy, onclick: () => void recover('restoreCopy') }, 'Open the copy from before') : null,
        h('button', { type: 'button', disabled: busy, onclick: () => void recover('markRepaired') }, 'I have undone it (Ctrl+Z)')));
  }

  function changes(i: PropertiesInfo): HTMLElement[] {
    if (!i.changes.length) return [h('p', { class: 'k-title' }, 'Changes'), h('p', { class: 'k-note' }, 'No change made by the add-in is logged in this model yet.')];
    return [h('p', { class: 'k-title' }, `Changes (newest ${i.changes.length})`), h('ul', { class: 'p-log' }, i.changes.slice(0, 8).map(c => h('li', {},
      `${c.command}: ${c.title}`,
      h('small', {}, new Date(c.when).toLocaleString('en-NZ', { dateStyle: 'medium', timeStyle: 'short' })),
      c.reason ? h('small', {}, `Reason: ${c.reason}`) : null,
      ...c.lines.slice(0, 3).map(l => h('small', {}, l)))))];
  }

  function about(i: PropertiesInfo): HTMLElement {
    const box = h('div', { class: 'p-left' });
    const warn = stopped(i);
    if (warn) box.append(warn);
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
    box.append(...changes(i));
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
