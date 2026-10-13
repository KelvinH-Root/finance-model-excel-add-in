/*
  The page every HFG dialog opens on: hfg-dialog.html?view=<command key>[&arg=...]. It draws the
  title and where the command lives, then the command's view: the add-in's own views (New model,
  Insert, Versions, Explorer, Remove) through window.HfgAddin, which reach the workbook by
  messages to the runtime that opened the dialog; the probe's designed views (the group
  structure chart, Add entity, Change ownership, Impacts) on their fictional samples; a short
  message; or, for a command not built yet, what it will do.
  Opened outside Excel (straight in a browser) it runs on the demo models, for a look.
*/
(function (root) {
  'use strict';
  const $ = id => document.getElementById(id);
  const params = new URLSearchParams(location.search);
  const C = root.HfgCommands;
  const A = root.HfgAddin;

  function el(tag, attrs, ...kids) {
    const n = document.createElement(tag);
    Object.entries(attrs || {}).forEach(([k, v]) => {
      if (v == null || v === false) return;
      if (k === 'class') n.className = v;
      else if (k.startsWith('on') && typeof v === 'function') n.addEventListener(k.slice(2), v);
      else n.setAttribute(k, v === true ? '' : String(v));
    });
    kids.flat(Infinity).forEach(k => { if (k != null && k !== false) n.appendChild(typeof k === 'string' ? document.createTextNode(k) : k); });
    return n;
  }

  function inOffice() {
    return typeof Office !== 'undefined' && Office.context && Office.context.ui && typeof Office.context.ui.messageParent === 'function'
      && Office.context.requirements && Office.context.requirements.isSetSupported('DialogApi', '1.2');
  }

  function heading(key) {
    const c = C && C.find(key);
    if (key === 'message') { $('d-title').textContent = params.get('title') || 'HFG Model'; $('d-crumb').textContent = ''; document.title = 'HFG Model'; return; }
    if (!c) return;
    const tab = c.tab === 'build' ? C.BUILD_TAB.label : c.tab === 'main' ? C.MAIN_TAB.label : 'Right-click';
    $('d-crumb').textContent = `${tab} › ${c.group}${c.parent ? ' › ' + C.find(c.parent).label : ''}`;
    $('d-title').textContent = c.ctxLabel && c.tab === 'context' ? c.ctxLabel : c.label;
    document.title = c.label;
  }

  function scroll(host) {
    host.innerHTML = '';
    const box = el('div', { class: 'd-pad' });
    host.appendChild(box);
    return box;
  }

  function footer(host, exit, text) {
    host.appendChild(el('div', { class: 'd-foot' }, el('span', { class: 'msg' }, text || ''),
      el('span', {}, el('button', { class: 'primary', onclick: () => exit.close() }, 'Close'))));
  }

  // The group views share one state in this page, so Add entity comes back to the chart with the new entity in it.
  function groupView(key, host, exit) {
    const show = k => {
      heading(k);
      host.innerHTML = '';
      if (k === 'grp-structure' || k === 'model-group') {
        const box = el('div', { class: 'gc', style: 'flex:1 1 auto;min-height:0;display:flex;flex-direction:column' });
        host.appendChild(box);
        root.HfgGroupChart.render(box, show);
        footer(host, exit, 'Fictional sample from the consolidation proof. Nothing here changes the workbook yet.');
      } else {
        const box = scroll(host);
        root.HfgGroup.render(k, box, show);
        footer(host, exit, 'Fictional sample. Add entity and Change ownership change this sample only.');
      }
    };
    show(key);
  }

  function impactsView(key, host, exit) {
    const box = scroll(host);
    root.HfgImpacts.render(key, box, () => {});
    footer(host, exit, 'Designed view on the probe\'s fictional samples; nothing is written to the workbook.');
  }

  function message(host, exit, text) {
    const box = scroll(host);
    box.appendChild(el('p', { style: 'font-size:13px;line-height:1.5' }, text || ''));
    footer(host, exit, '');
  }

  function notBuilt(key, host, exit) {
    const c = C && C.find(key);
    const box = scroll(host);
    if (!c) { box.appendChild(el('p', {}, `No command ${key}.`)); footer(host, exit); return; }
    const surface = C.surfaceOf(key);
    $('d-title').appendChild(el('span', { class: 'nb-badge' }, 'Not built yet'));
    box.appendChild(el('div', { class: 'nb' },
      el('p', {}, c.tip || ''),
      el('h3', {}, 'What it will do'),
      el('p', {}, c.view || ''),
      el('dl', {},
        el('dt', {}, 'Opens as'), el('dd', {}, C.SURFACE_TEXT[surface]),
        el('dt', {}, 'Lands in'), el('dd', {}, c.phase ? `Phase ${c.phase}` : '-'))));
    footer(host, exit, 'Each command not built yet opens a page like this, in the place it will open when built.');
  }

  function start() {
    const key = params.get('view') || '';
    const arg = params.get('arg') || '';
    const host = $('d-host');
    const office = inOffice();
    const exit = {
      close: () => { if (office) A.askClose(); else { try { window.close(); } catch (e) { /* not a window we opened */ } } },
      open: (k, a) => { if (office) A.askOpen(k, a); else location.search = `?view=${encodeURIComponent(k)}${a ? `&arg=${encodeURIComponent(a)}` : ''}`; }
    };
    heading(key);
    try {
      if (key === 'message') return message(host, exit, arg);
      if (A && A.hasView(key)) {
        const svc = office ? A.remote() : A.service;
        host.innerHTML = '';
        A.mount(key, host, svc, exit);
        return;
      }
      if (['grp-structure', 'model-group', 'grp-add', 'grp-ownership'].includes(key) && root.HfgGroup) return groupView(key, host, exit);
      if (['imp-live', 'imp-sheets', 'ctx-impacts', 'an-impacts'].includes(key) && root.HfgImpacts) return impactsView(key, host, exit);
      notBuilt(key, host, exit);
    } catch (e) {
      message(host, exit, `This view could not open: ${(e && e.message) || e}`);
    }
  }

  if (typeof Office !== 'undefined' && typeof Office.onReady === 'function') Office.onReady(() => start());
  else start();
})(typeof globalThis !== 'undefined' ? globalThis : this);
