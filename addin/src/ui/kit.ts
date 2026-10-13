// Small DOM helpers and the shared look of the add-in's dialogs (and of the same views when they
// fall back to the task pane). Colours come from the page's variables (HF palette, taskpane.css).

type Kid = Node | string | null | undefined | false | Kid[];

/** An element with attributes (on* functions become listeners) and children. */
export function h<K extends keyof HTMLElementTagNameMap>(tag: K, attrs: Record<string, unknown> = {}, ...kids: Kid[]): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === null || v === undefined || v === false) continue;
    if (k.startsWith('on') && typeof v === 'function') e.addEventListener(k.slice(2), v as EventListener);
    else e.setAttribute(k, v === true ? '' : String(v));
  }
  const add = (k: Kid) => {
    if (k === null || k === undefined || k === false) return;
    if (Array.isArray(k)) k.forEach(add);
    else e.append(k);
  };
  kids.forEach(add);
  return e;
}

export const CSS = `
.dlg { display: flex; flex-direction: column; height: 100%; min-height: 0; }
.dlg-body { flex: 1 1 auto; min-height: 0; overflow: auto; padding: 12px 16px; }
.dlg-foot { display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 8px 16px;
  border-top: 1px solid var(--line); background: var(--panel); }
.dlg-foot .msg { color: var(--quiet); flex: 1 1 auto; }
.dlg-foot .msg.bad { color: var(--fail); }
.dlg-foot .btns { display: flex; gap: 6px; flex: none; }
.dlg-cols { display: grid; grid-template-columns: minmax(230px, 34%) 1fr; flex: 1 1 auto; min-height: 0; }
.dlg-cols > .left { border-right: 1px solid var(--line); overflow: auto; min-height: 0; }
.dlg-cols > .right { overflow: auto; min-height: 0; padding: 12px 16px; }
@media (max-width: 600px) {
  .dlg-cols { grid-template-columns: 1fr; }
  .dlg-cols > .left { border-right: 0; border-bottom: 1px solid var(--line); max-height: 280px; }
}
.k-search { padding: 10px 12px 6px; }
.k-inp { width: 100%; box-sizing: border-box; font: inherit; padding: 5px 7px; border: 1px solid var(--line); border-radius: 4px;
  color: var(--ink); background: transparent; }
.k-area { font-size: 11px; font-weight: 600; color: var(--accent-text); text-transform: uppercase; letter-spacing: .04em; padding: 9px 12px 3px; }
.k-item { display: flex; align-items: center; gap: 8px; width: 100%; padding: 4px 12px; margin: 0; border: 0; border-radius: 0;
  background: transparent; text-align: left; font: inherit; color: var(--ink); cursor: pointer; }
.k-item:hover { background: var(--panel); }
.k-item.sel { background: var(--accent-tint); box-shadow: inset 3px 0 0 var(--accent-line); }
.k-item.off { color: var(--quiet); cursor: default; }
.k-item small { margin-left: auto; font-size: 11px; color: var(--quiet); }
.k-title { font-size: 14px; font-weight: 600; margin: 0 0 2px; }
.k-field { display: flex; flex-direction: column; gap: 3px; margin: 0 0 10px; }
.k-field > span { font-weight: 600; }
.k-field.dim { opacity: .5; }
.k-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(190px, 1fr)); gap: 0 14px; }
.k-unit { display: flex; align-items: center; gap: 6px; }
.k-unit .k-inp { text-align: right; }
.k-unit small { color: var(--quiet); min-width: 14px; }
.k-card { border: 1px solid var(--accent-tint); border-radius: 6px; background: var(--panel); padding: 10px 12px; margin: 12px 0 0; }
.k-card strong { display: block; margin-bottom: 4px; }
.k-card ul { margin: 4px 0 0; padding-left: 16px; }
.k-card li { margin: 0 0 4px; line-height: 1.4; }
.k-opt { display: flex; gap: 8px; align-items: flex-start; padding: 6px 8px; margin: 0 0 4px; border-radius: 4px; cursor: pointer; }
.k-opt.sel { background: var(--accent-tint); }
.k-opt small { display: block; color: var(--quiet); font-size: 11px; }
.k-table { width: 100%; border-collapse: collapse; margin: 4px 0; }
.k-table th { font-size: 11px; color: var(--accent-text); text-align: left; padding: 3px 4px; border-bottom: 1px solid var(--line); }
.k-table td { padding: 4px; border-bottom: 1px solid var(--line); vertical-align: middle; }
.k-table td.n { text-align: right; color: var(--quiet); width: 34px; font-variant-numeric: tabular-nums; }
.k-table button { font-size: 11px; padding: 1px 8px; }
.k-pill { display: inline-block; font-size: 11px; padding: 0 7px; border-radius: 9px; background: var(--panel); color: var(--accent-text); }
.k-pill.ok { background: var(--accent-tint); }
.k-stat { display: inline-block; width: 8px; height: 8px; border-radius: 50%; flex: none; }
.k-stat.error { background: var(--fail); }
.k-stat.alert { background: var(--warn); }
.k-stat.ok { background: var(--accent-line); }
.k-quiet { color: var(--quiet); }
.k-note { font-size: 11px; color: var(--accent-text); background: var(--panel); border-radius: 6px; padding: 7px 10px; margin: 10px 0 0; }
`;

export function useCss(id: string, css: string): void {
  if (!document.getElementById(id)) document.head.append(h('style', { id }, css));
}

/** The frame of a dialog view: a body that scrolls and a footer with a message and buttons. */
export function frame(host: HTMLElement, body: Node, message: string, bad: boolean, buttons: HTMLElement[]): void {
  useCss('hfg-kit-css', CSS);
  host.innerHTML = '';
  host.append(h('div', { class: 'dlg' }, body,
    h('div', { class: 'dlg-foot' }, h('span', { class: 'msg' + (bad ? ' bad' : ''), role: 'status' }, message), h('span', { class: 'btns' }, buttons))));
}

/** How a dialog view leaves: close the dialog, or (in the task pane) do nothing. */
export interface Exit {
  close(): void;
  /** Open another command in this one's place (a dialog asks the runtime; the pane just switches). */
  open(key: string, arg?: string): void;
}
