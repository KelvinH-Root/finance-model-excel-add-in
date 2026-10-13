// Insert module, as a dialog: the library by area on the left; on the right the module chosen, its
// settings and, kept up to date as they change, what inserting it will write. Insert applies the
// plan to the open workbook and selects the new module there. It talks to the workbook only
// through the model service, so the same view runs in a dialog or in the task pane.

import type { InsertChoiceInfo, Preview, ServiceApi, SettingInfo } from '../service.ts';
import { frame, h, type Exit } from '../ui/kit.ts';

type Svc = Pick<ServiceApi, 'insertChoices' | 'insertPreview' | 'insertApply'>;

const isPct = (s: SettingInfo) => s.unit === '%';
const shown = (s: SettingInfo, v: unknown) => (v === null || v === undefined ? '' : isPct(s) && typeof v === 'number' ? String(Math.round(v * 1e6) / 1e4) : String(v));

export function mountInsert(host: HTMLElement, svc: Svc, exit: Exit): void {
  let list: InsertChoiceInfo[] = [];
  let source = '';
  let summary = '';
  let picked: InsertChoiceInfo | null = null;
  let values: Record<string, unknown> = {};
  let name = '';
  let filter = '';
  let preview: Preview | null = null;
  let message = 'Reading the model...';
  let bad = false;
  let busy = false;
  let done = false;
  let timer: ReturnType<typeof setTimeout> | null = null;
  let ask = 0;

  async function load() {
    try {
      const r = await svc.insertChoices();
      list = r.choices;
      source = r.source;
      summary = r.summary;
      message = `${summary}, read from ${source}.`;
      bad = false;
      done = false;
      picked = null;
      preview = null;
    } catch (e) {
      message = (e as Error).message;
      bad = true;
    }
    render();
  }

  function pick(c: InsertChoiceInfo) {
    if (c.blocked) return;
    picked = c;
    values = Object.fromEntries(c.settings.map(s => [s.key, s.value]));
    name = '';
    preview = null;
    schedule(0);
    render();
  }

  function schedule(ms = 450) {
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => void refresh(), ms);
  }

  async function refresh() {
    if (!picked) return;
    const mine = ++ask;
    message = 'Working out what it writes...';
    bad = false;
    renderFoot();
    try {
      const p = await svc.insertPreview(picked.id, values, picked.category && name.trim() ? name.trim() : null);
      if (mine !== ask) return;
      preview = p;
      message = `${summary}, read from ${source}.`;
    } catch (e) {
      if (mine !== ask) return;
      preview = null;
      message = (e as Error).message;
      bad = true;
    }
    render();
  }

  async function insert() {
    if (!picked || !preview || busy) return;
    busy = true;
    message = 'Inserting...';
    renderFoot();
    try {
      const r = await svc.insertApply();
      message = r.message;
      bad = false;
      done = true;
    } catch (e) {
      message = `The insert stopped: ${(e as Error).message}. Undo in Excel (Ctrl+Z) takes back what was written.`;
      bad = true;
    }
    busy = false;
    render();
  }

  function field(s: SettingInfo): HTMLElement {
    const v = values[s.key];
    const dim = s.when ? Number(values[s.when[0]]) !== s.when[1] : false;
    let control: HTMLElement;
    if (s.check) {
      control = h('input', { type: 'checkbox', checked: !!v, onchange: (e: Event) => { values[s.key] = (e.target as HTMLInputElement).checked; schedule(); } });
    } else if (s.options) {
      control = h('select', { class: 'k-inp', onchange: (e: Event) => { values[s.key] = Number((e.target as HTMLSelectElement).value); render(); schedule(); } },
        s.options.map((o, k) => h('option', { value: String(k + 1), selected: Number(v) === k + 1 }, o)));
    } else {
      const inp = h('input', { class: 'k-inp', type: 'number', step: 'any', value: shown(s, v), 'aria-label': s.label,
        oninput: (e: Event) => {
          const t = (e.target as HTMLInputElement).value;
          values[s.key] = t === '' ? null : Number(t) / (isPct(s) ? 100 : 1);
          schedule();
        } });
      control = h('span', { class: 'k-unit' }, inp, h('small', {}, s.unit === '$' ? '' : s.unit));
    }
    return h('label', { class: 'k-field' + (dim ? ' dim' : '') }, h('span', {}, s.label), control);
  }

  function left(): HTMLElement {
    const box = h('div', { class: 'left' });
    box.append(h('div', { class: 'k-search' }, h('input', { class: 'k-inp', type: 'search', placeholder: 'Find a module', value: filter, 'aria-label': 'Find a module',
      oninput: (e: Event) => { filter = (e.target as HTMLInputElement).value; drawList(); } })));
    const items = h('div', { role: 'listbox', 'aria-label': 'Modules' });
    box.append(items);
    function drawList() {
      items.innerHTML = '';
      let area = '';
      const f = filter.trim().toLowerCase();
      for (const c of list) {
        if (f && !`${c.title} ${c.area}`.toLowerCase().includes(f)) continue;
        if (c.area !== area) { area = c.area; items.append(h('div', { class: 'k-area' }, area)); }
        items.append(h('button', {
          class: 'k-item' + (picked?.id === c.id ? ' sel' : '') + (c.blocked ? ' off' : ''), role: 'option', 'aria-selected': String(picked?.id === c.id),
          title: c.blocked ?? '', disabled: !!c.blocked, onclick: () => pick(c),
        }, c.title, c.blocked ? h('small', {}, c.blocked.startsWith('Already') ? 'In the model' : 'Update first') : null));
      }
    }
    drawList();
    return box;
  }

  function right(): HTMLElement {
    const box = h('div', { class: 'right' });
    if (!picked) {
      box.append(h('p', { class: 'k-quiet' }, list.length ? 'Choose a module on the left. Modules that go in once are greyed while they are in the model.' : ''));
      return box;
    }
    box.append(h('p', { class: 'k-title' }, picked.title));
    if (picked.description) box.append(h('p', { class: 'k-quiet' }, picked.description));
    const grid = h('div', { class: 'k-grid' });
    if (picked.category) {
      grid.append(h('label', { class: 'k-field' }, h('span', {}, 'Title'),
        h('input', { class: 'k-inp', type: 'text', value: name, placeholder: `${picked.title} (numbered)`,
          oninput: (e: Event) => { name = (e.target as HTMLInputElement).value; schedule(); } })));
    }
    picked.settings.forEach(s => grid.append(field(s)));
    box.append(grid);
    if (preview) {
      box.append(h('div', { class: 'k-card' }, h('strong', {}, `What inserting ${preview.title} writes`),
        h('ul', {}, preview.lines.map(l => h('li', {}, l)))));
    }
    return box;
  }

  let footHost: HTMLElement | null = null;
  function buttons(): HTMLElement[] {
    if (done) {
      return [h('button', { type: 'button', onclick: () => void load() }, 'Insert another'),
        h('button', { type: 'button', class: 'primary', onclick: () => exit.close() }, 'Close')];
    }
    return [h('button', { type: 'button', onclick: () => exit.close() }, 'Cancel'),
      h('button', { type: 'button', class: 'primary', disabled: !preview || busy, onclick: () => void insert() }, 'Insert')];
  }

  function renderFoot() {
    if (!footHost) return render();
    const msg = footHost.querySelector('.msg') as HTMLElement;
    msg.textContent = message;
    msg.className = 'msg' + (bad ? ' bad' : '');
    const btns = footHost.querySelector('.btns') as HTMLElement;
    btns.innerHTML = '';
    btns.append(...buttons());
  }

  function render() {
    const body = h('div', { class: 'dlg-cols' }, left(), right());
    frame(host, body, message, bad, buttons());
    footHost = host.querySelector('.dlg-foot');
  }

  render();
  void load();
}
