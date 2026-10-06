/* Task pane wiring for the HFG probe. */
(function () {
  'use strict';
  const P = window.HfgProbe;
  const $ = id => document.getElementById(id);
  const STATUS_TEXT = { pass: 'Pass', warn: 'Check', fail: 'Fail', manual: 'Your turn', skip: 'Skipped', running: 'Running', waiting: 'Not run' };

  function message(text) { $('message').textContent = text; }

  function render() {
    const body = document.querySelector('#results tbody');
    body.innerHTML = '';
    P.PROBES.forEach(p => {
      const r = P.state.results.get(p.id) || { status: 'waiting', detail: '' };
      const tr = document.createElement('tr');
      const test = document.createElement('td');
      const title = document.createElement('div');
      title.textContent = p.title;
      const detail = document.createElement('div');
      detail.className = 'detail';
      detail.textContent = r.detail || '';
      test.appendChild(title);
      test.appendChild(detail);
      const status = document.createElement('td');
      const chip = document.createElement('span');
      chip.className = 'chip ' + r.status;
      chip.textContent = STATUS_TEXT[r.status] || r.status;
      status.appendChild(chip);
      const time = document.createElement('td');
      time.className = 'time';
      time.textContent = r.ms == null ? '' : (r.ms / 1000).toFixed(1) + ' s';
      tr.appendChild(test); tr.appendChild(status); tr.appendChild(time);
      body.appendChild(tr);
    });
    const link = $('download-json');
    if (link.dataset.url) URL.revokeObjectURL(link.dataset.url);
    const url = URL.createObjectURL(new Blob([P.resultsJson()], { type: 'application/json' }));
    link.href = url;
    link.dataset.url = url;
  }

  function busy(button, fn) {
    return async () => {
      button.disabled = true;
      try { await fn(); } catch (e) { message((e && e.message) || String(e)); }
      finally { button.disabled = false; render(); }
    };
  }

  function bind(id, fn) { const b = $(id); b.addEventListener('click', busy(b, fn)); }

  async function onFileChosen(event) {
    const file = event.target.files[0];
    const list = $('sheet-list');
    list.innerHTML = '';
    $('insert-sheets').disabled = true;
    if (!file) return;
    try {
      const sheets = await P.readChosenFile(file);
      sheets.forEach(s => {
        const label = document.createElement('label');
        const box = document.createElement('input');
        box.type = 'checkbox';
        box.value = s.name;
        label.appendChild(box);
        label.appendChild(document.createTextNode(' ' + s.name + (s.state !== 'visible' ? ` (${s.state}: Excel will not insert it)` : '')));
        list.appendChild(label);
      });
      $('insert-sheets').disabled = false;
      message(`${file.name}: ${sheets.length} sheets. Tick the ones to insert.`);
    } catch (e) { message('Could not read the file: ' + e.message); }
  }

  function renderView(key) {
    const C = window.HfgCommands;
    const c = key && C ? C.find(key) : null;
    $('cmd-view').classList.toggle('hidden', !c);
    $('probe-body').classList.toggle('hidden', !!c);
    if (!c) return;
    const tab = c.tab === 'build' ? C.BUILD_TAB.label : c.tab === 'main' ? C.MAIN_TAB.label : 'Right-click menu';
    $('cmd-where').textContent = `${tab} > ${c.group}${c.parent ? ' > ' + C.find(c.parent).label : ''}`;
    $('cmd-title').textContent = c.label;
    $('cmd-what').textContent = c.view;
    $('cmd-modano').textContent = c.modano || 'New';
    $('cmd-phase').textContent = c.phase ? `Phase ${c.phase}` : '';
    window.scrollTo(0, 0);
  }

  function wire() {
    $('version').textContent = 'v' + P.VERSION;
    P.onView = renderView;
    $('cmd-back').addEventListener('click', () => { P.state.view = null; renderView(null); });
    bind('build-show', () => P.toggleBuildTab(true));
    bind('build-hide', () => P.toggleBuildTab(false));
    bind('group-hide', () => P.runProbe('ribbon-hide', () => P.hideGroup(false)));
    bind('group-show', () => P.runProbe('ribbon-hide', () => P.hideGroup(true)));
    P.state.listeners.push(() => render());
    render();

    bind('run-auto', async () => {
      const auto = P.PROBES.filter(p => p.auto);
      let done = 0;
      await P.runAutomatic(() => { done++; $('auto-progress').textContent = `${done} of ${auto.length} done`; });
      const env = P.state.env || {};
      $('platform').textContent = `${env.platform || ''} Excel ${env.version || ''}`;
      message('Automatic tests finished. Now work through section 2.');
    });

    $('file').addEventListener('change', onFileChosen);
    bind('insert-sheets', async () => {
      const names = Array.from(document.querySelectorAll('#sheet-list input:checked')).map(b => b.value);
      if (!names.length) { message('Tick at least one sheet.'); return; }
      await P.runProbe('file-insert', () => P.insertFromFile(names));
    });
    bind('copy-sheet', () => P.runProbe('sheet-copy', P.copyInsertedSheet));
    bind('sim-run', () => P.runProbe('sim-run', P.simulationRun));
    bind('sim-insert', () => P.runProbe('sim-insert', P.simulationRun));
    bind('file-read', () => P.runProbe('file-read', P.exportCompressed));
    bind('file-open', () => P.runProbe('file-open', P.openCopy));
    bind('pdf', async () => {
      await P.runProbe('pdf', P.exportPdf);
      if (P.state.pdfUrl) { const a = $('pdf-link'); a.href = P.state.pdfUrl; a.classList.remove('hidden'); }
    });
    bind('undo', () => P.runProbe('undo', P.undoSetup));
    bind('ribbon-off', () => P.runProbe('ribbon', () => P.ribbonDisable(false)));
    bind('ribbon-on', () => P.runProbe('ribbon', () => P.ribbonDisable(true)));
    bind('context-off', () => P.runProbe('context', () => P.contextMenuDisable(false)));
    bind('context-on', () => P.runProbe('context', () => P.contextMenuDisable(true)));
    bind('dialog', () => P.runProbe('dialog', P.openDialog));
    bind('sign-in', () => P.runProbe('naa', () => P.signIn($('client-id').value.trim(), $('tenant-id').value.trim())));
    bind('metadata-again', () => P.runProbe('metadata'));

    document.querySelectorAll('.confirm').forEach(box => {
      box.querySelectorAll('button').forEach(b => b.addEventListener('click', () => {
        P.confirm(box.dataset.probe, b.dataset.answer === 'yes', b.dataset.note);
        render();
      }));
    });

    bind('write-sheet', async () => { const name = await P.writeResultsSheet(); message(`Results written to ${name}. Save the workbook or copy the sheet to send it.`); });
    bind('copy-json', async () => {
      const text = P.resultsJson();
      try { await navigator.clipboard.writeText(text); message('Results copied.'); }
      catch (e) {
        const area = document.createElement('textarea');
        area.value = text; document.body.appendChild(area); area.select();
        const ok = document.execCommand('copy'); area.remove();
        message(ok ? 'Results copied.' : 'Copy is blocked here: use Write results to a sheet instead.');
      }
    });
    bind('clean-up', async () => { message(await P.cleanUp()); });
  }

  Office.onReady(info => {
    P.registerCommands();
    if (info.host === Office.HostType.Excel) {
      const d = Office.context.diagnostics || {};
      wire();
      $('platform').textContent = `${d.platform || ''} Excel ${d.version || ''}, session ${P.SESSION}`;
    } else {
      $('platform').textContent = 'Open this add-in in Excel.';
    }
  });
})();
