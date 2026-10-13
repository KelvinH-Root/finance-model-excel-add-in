// Remove a category line, asking first: what goes with it (its rows, names and checks, and what
// it sent elsewhere), then Remove. The line is the one under the selection in the workbook.

import type { Preview, ServiceApi } from '../service.ts';
import { frame, h, type Exit } from '../ui/kit.ts';

type Svc = Pick<ServiceApi, 'removePreview' | 'removeApply'>;

export function mountRemove(host: HTMLElement, svc: Svc, exit: Exit): void {
  let preview: Preview | null = null;
  let message = 'Reading the selection...';
  let bad = false;
  let busy = false;
  let done = false;

  async function load() {
    try {
      preview = await svc.removePreview();
      message = 'Undo in Excel (Ctrl+Z) takes back a removal.';
      bad = false;
    } catch (e) {
      message = (e as Error).message;
      bad = true;
    }
    render();
  }

  async function remove() {
    if (!preview || busy) return;
    busy = true;
    message = 'Removing...';
    render();
    try {
      message = (await svc.removeApply()).message;
      bad = false;
      done = true;
    } catch (e) {
      message = `Stopped: ${(e as Error).message}. Undo in Excel (Ctrl+Z) takes back what was done.`;
      bad = true;
    }
    busy = false;
    render();
  }

  function render() {
    const body = h('div', { class: 'dlg-body' });
    if (preview && !done) {
      body.append(h('p', { class: 'k-title' }, `Remove ${preview.title}?`));
      body.append(h('div', { class: 'k-card', style: 'margin-top:6px' }, h('strong', {}, 'What goes'), h('ul', {}, preview.lines.map(l => h('li', {}, l)))));
    } else if (done) {
      body.append(h('p', {}, message));
    } else if (bad) {
      body.append(h('p', {}, message));
    }
    const buttons = done || !preview
      ? [h('button', { type: 'button', class: 'primary', onclick: () => exit.close() }, 'Close')]
      : [h('button', { type: 'button', onclick: () => exit.close() }, 'Cancel'),
        h('button', { type: 'button', class: 'primary', disabled: busy, onclick: () => void remove() }, 'Remove')];
    frame(host, body, done || (!preview && bad) ? '' : message, bad && !!preview, buttons);
  }

  render();
  void load();
}
