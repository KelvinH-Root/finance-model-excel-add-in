// Versions, as a dialog: Save version (this month's reforecast or the budget being built), the
// saved versions with Approve and Lock or unlock, and for each a preview of what it writes before
// Apply. It talks to the workbook only through the model service.

import type { Preview, ServiceApi, VersionRequest, VersionsInfo } from '../service.ts';
import { frame, h, type Exit } from '../ui/kit.ts';

type Svc = Pick<ServiceApi, 'versionsState' | 'versionsPreview' | 'versionsApply'>;

export function mountVersions(host: HTMLElement, svc: Svc, exit: Exit): void {
  let info: VersionsInfo | null = null;
  let option = 1;
  let savedBy = 'Group Finance';
  let preview: Preview | null = null;
  let message = 'Reading the model...';
  let bad = false;
  let busy = false;

  async function load() {
    try {
      info = await svc.versionsState();
      option = Math.min(option, info.options.length - 1);
      message = `${info.register.length} saved versions, read from ${info.source}.`;
      bad = false;
    } catch (e) {
      message = (e as Error).message;
      bad = true;
    }
    render();
  }

  async function ask(req: VersionRequest) {
    busy = true;
    message = req.kind === 'save' ? 'Reading the statements...' : 'Working out what it writes...';
    bad = false;
    render();
    try {
      preview = await svc.versionsPreview(req);
      message = 'Check what it writes, then Apply.';
    } catch (e) {
      preview = null;
      message = (e as Error).message;
      bad = true;
    }
    busy = false;
    render();
  }

  async function apply() {
    if (!preview || busy) return;
    busy = true;
    message = 'Writing...';
    render();
    try {
      const r = await svc.versionsApply();
      preview = null;
      busy = false;
      await load();
      message = r.message;
      bad = false;
    } catch (e) {
      message = `Stopped: ${(e as Error).message}. Undo in Excel (Ctrl+Z) takes back what was written.`;
      bad = true;
      busy = false;
    }
    render();
  }

  function render() {
    const body = h('div', { class: 'dlg-body' });
    if (info) {
      const save = h('div', { class: 'k-card', style: 'margin-top:0' }, h('strong', {}, 'Save version'));
      info.options.forEach((o, k) => save.append(h('label', { class: 'k-opt' + (k === option ? ' sel' : '') },
        h('input', { type: 'radio', name: 'ver-opt', checked: k === option, onchange: () => { option = k; preview = null; render(); } }),
        h('span', {}, o.label, h('small', {}, o.note)))));
      save.append(h('div', { style: 'display:flex;gap:8px;align-items:center;margin-top:6px' },
        h('span', {}, 'Saved by'),
        h('input', { class: 'k-inp', type: 'text', value: savedBy, style: 'flex:1', oninput: (e: Event) => { savedBy = (e.target as HTMLInputElement).value; } }),
        h('button', { type: 'button', disabled: busy, onclick: () => void ask({ kind: 'save', option, savedBy }) }, 'Preview')));
      body.append(save);
      if (preview) {
        body.append(h('div', { class: 'k-card' }, h('strong', {}, `What "${preview.title}" writes`), h('ul', {}, preview.lines.map(l => h('li', {}, l)))));
      }
      const table = h('table', { class: 'k-table' }, h('tr', {}, h('th', {}, 'Version'), h('th', {}, 'Status'), h('th', {}, '')));
      for (const v of info.register) {
        table.append(h('tr', {},
          h('td', {}, v.label),
          h('td', {}, h('span', { class: 'k-pill' + (v.status === 'Approved' ? ' ok' : '') }, v.status), v.locked ? ' locked' : ''),
          h('td', { style: 'text-align:right;white-space:nowrap' },
            v.canApprove ? h('button', { type: 'button', disabled: busy, onclick: () => void ask({ kind: 'approve', id: v.id }) }, 'Approve') : null,
            h('button', { type: 'button', disabled: busy, onclick: () => void ask({ kind: 'lock', id: v.id, locked: !v.locked }) }, v.locked ? 'Unlock' : 'Lock'))));
      }
      body.append(h('p', { class: 'k-title', style: 'margin-top:14px;font-size:12px' }, 'Saved versions'), table);
      body.append(h('p', { class: 'k-note' }, 'A version saves only when no error check is failing. Compare versions goes to the Version comparison sheet.'));
    }
    frame(host, body, message, bad, [
      h('button', { type: 'button', onclick: () => exit.close() }, 'Close'),
      h('button', { type: 'button', class: 'primary', disabled: !preview || busy, onclick: () => void apply() }, 'Apply'),
    ]);
  }

  render();
  void load();
}
