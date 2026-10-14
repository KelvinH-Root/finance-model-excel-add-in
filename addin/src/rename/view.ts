// Rename the line under the selection (Modules > Rename): its labels, contents entry, the
// drop-down item other lines choose it by and every label that reads it follow. What the rename
// writes shows as the title is typed; Enter or Rename applies it. A single module keeps its
// library title, so the view says so instead.

import type { Preview, RenameInfo, ServiceApi } from '../service.ts';
import { frame, h, type Exit } from '../ui/kit.ts';

type Svc = Pick<ServiceApi, 'renameState' | 'renamePreview' | 'renameApply'>;

export function mountRename(host: HTMLElement, svc: Svc, exit: Exit): void {
  let info: RenameInfo | null = null;
  let name = '';
  let preview: Preview | null = null;
  let message = 'Reading the selection...';
  let bad = false;
  let busy = false;
  let done = false;
  let timer: ReturnType<typeof setTimeout> | null = null;
  let ask = 0;
  let card: HTMLElement | null = null;

  async function load() {
    try {
      info = await svc.renameState();
      name = info.title;
      message = `A ${info.category} line on ${info.sheet}.`;
      bad = false;
    } catch (e) {
      message = (e as Error).message;
      bad = true;
    }
    render();
  }

  function schedule(ms = 450) {
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => void refresh(), ms);
  }

  async function refresh() {
    if (!info) return;
    const mine = ++ask;
    if (name.trim() === info.title || !name.trim()) {
      preview = null;
      message = name.trim() ? `A ${info.category} line on ${info.sheet}.` : 'Type the new title.';
      bad = false;
      return update();
    }
    message = 'Working out what it changes...';
    bad = false;
    update();
    try {
      const p = await svc.renamePreview(name);
      if (mine !== ask) return;
      preview = p;
      message = 'Undo in Excel (Ctrl+Z) takes back a rename.';
    } catch (e) {
      if (mine !== ask) return;
      preview = null;
      message = (e as Error).message;
      bad = true;
    }
    update();
  }

  async function apply() {
    if (!preview || busy) return;
    if (timer) { clearTimeout(timer); timer = null; }
    busy = true;
    message = 'Renaming...';
    update();
    try {
      message = (await svc.renameApply()).message;
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
      h('button', { type: 'button', class: 'primary', disabled: !preview || busy, onclick: () => void apply() }, 'Rename')];
  }

  /** The preview card and the footer, without drawing the field again (so typing keeps its place). */
  function update() {
    if (card) {
      card.innerHTML = '';
      if (preview) {
        card.hidden = false;
        card.append(h('strong', {}, `What renaming to ${preview.title} changes`), h('ul', {}, preview.lines.map(l => h('li', {}, l))));
      } else card.hidden = true;
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

  function render() {
    const body = h('div', { class: 'dlg-body' });
    card = null;
    let input: HTMLInputElement | null = null;
    if (info && !done) {
      body.append(h('p', { class: 'k-title' }, `Rename ${info.title}`));
      input = h('input', {
        class: 'k-inp', type: 'text', value: name, maxlength: 80, 'aria-label': 'New title',
        oninput: (e: Event) => { name = (e.target as HTMLInputElement).value; schedule(); },
        onkeydown: (e: KeyboardEvent) => { if (e.key === 'Enter') { e.preventDefault(); void apply(); } },
      });
      body.append(h('label', { class: 'k-field', style: 'margin-top:8px' }, h('span', {}, 'New title'), input));
      body.append(h('p', { class: 'k-quiet', style: 'margin:0' },
        'Its labels, its contents entry, the drop-downs that choose it and every label that reads it take the new title. Its names stay as they are, so formulas that use them keep working.'));
      card = h('div', { class: 'k-card', hidden: true });
      body.append(card);
    } else {
      body.append(h('p', {}, message));
    }
    frame(host, body, done || !info ? '' : message, bad && !!info, buttons());
    update();
    if (input) { input.focus(); input.select(); }
  }

  render();
  void load();
}
