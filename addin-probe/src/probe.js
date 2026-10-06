/*
 * HFG add-in probe.
 *
 * Runs the Office.js tests that desk research cannot settle (see "How sure we are"
 * in the requirements spec). Plain JavaScript with no build step, so it can be
 * served from any HTTPS folder. Everything it creates in the workbook starts with
 * "zProbe_" and is removed by Clean up.
 */
(function (root) {
  'use strict';

  const VERSION = '0.2.0';
  const NS = 'urn:hfg:addin-probe:v1';
  const PREFIX = 'zProbe_';
  const STYLE = 'HFG Probe Input';
  const SESSION = Math.random().toString(36).slice(2, 10);

  // Requirement sets reported by the environment probe, and the ones the spec relies on.
  const SETS = {
    ExcelApi: ['1.1', '1.4', '1.5', '1.7', '1.8', '1.9', '1.10', '1.12', '1.13', '1.14', '1.15', '1.16', '1.17', '1.18', '1.19', '1.20', '1.21'],
    ExcelApiDesktop: ['1.1'],
    ExcelApiOnline: ['1.1'],
    SharedRuntime: ['1.1', '1.2'],
    RibbonApi: ['1.1', '1.2', '1.3'],
    ContextMenuApi: ['1.1'],
    KeyboardShortcuts: ['1.1'],
    AddinCommands: ['1.1', '1.3'],
    NestedAppAuth: ['1.1'],
    DialogApi: ['1.1', '1.2'],
    TaskPaneApi: ['1.1'],
    IdentityAPI: ['1.3'],
    CompressedFile: ['1.1'],
    PdfFile: ['1.1']
  };
  const REQUIRED = [
    ['ExcelApi', '1.18'], ['SharedRuntime', '1.1'], ['RibbonApi', '1.1'], ['ContextMenuApi', '1.1'],
    ['NestedAppAuth', '1.1'], ['DialogApi', '1.1']
  ];

  const state = {
    results: new Map(),
    env: null,
    file: null,            // {name, buffer, sheets}
    inserted: [],          // sheet names inserted from a file
    exported: null,        // {bytes, base64} of the open workbook from getFileAsync
    pdfUrl: null,
    listeners: []
  };

  // ---------- helpers ----------

  const now = () => (typeof performance !== 'undefined' ? performance.now() : Date.now());
  const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
  const avg = list => list.reduce((s, v) => s + v, 0) / Math.max(1, list.length);
  const secs = ms => (ms / 1000).toFixed(ms < 10000 ? 2 : 1) + ' s';
  const num = n => Number(n).toLocaleString('en-NZ');

  function supports(name, version) {
    try { return Office.context.requirements.isSetSupported(name, version); } catch (e) { return false; }
  }

  function failure(e) {
    const code = (e && (e.code || e.name)) || 'Error';
    const where = e && e.debugInfo && e.debugInfo.errorLocation ? ` at ${e.debugInfo.errorLocation}` : '';
    return { status: 'fail', detail: `${code}: ${(e && e.message) || e}${where}` };
  }

  function record(id, outcome) {
    const probe = PROBES.find(p => p.id === id) || { id, area: 'Command', title: id };
    state.results.set(id, Object.assign({ id, area: probe.area, title: probe.title, at: new Date().toISOString() }, outcome));
    state.listeners.forEach(fn => { try { fn(id); } catch (e) { /* ignore */ } });
  }

  async function freshSheet(context, suffix) {
    const name = PREFIX + suffix;
    const old = context.workbook.worksheets.getItemOrNullObject(name);
    await context.sync();
    if (!old.isNullObject) { old.delete(); await context.sync(); }
    const ws = context.workbook.worksheets.add(name);
    await context.sync();
    return ws;
  }

  async function dropNames(context, names) {
    const items = names.map(n => context.workbook.names.getItemOrNullObject(n));
    await context.sync();
    items.forEach(item => { if (!item.isNullObject) item.delete(); });
    await context.sync();
  }

  async function withManualCalc(context, fn) {
    const app = context.workbook.application;
    app.load('calculationMode');
    await context.sync();
    const before = app.calculationMode;
    if (supports('ExcelApi', '1.8')) app.calculationMode = Excel.CalculationMode.manual;
    await context.sync();
    try { return await fn(); }
    finally {
      if (supports('ExcelApi', '1.8')) { app.calculationMode = before; await context.sync(); }
    }
  }

  function saveSettings() {
    return new Promise((resolve, reject) => {
      Office.context.document.settings.saveAsync(r => (r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error)));
    });
  }

  function getFile(type) {
    return new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(type, { sliceSize: 4194304 }, res => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) { reject(res.error); return; }
        const file = res.value;
        const parts = [];
        const next = i => file.getSliceAsync(i, r => {
          if (r.status !== Office.AsyncResultStatus.Succeeded) { file.closeAsync(); reject(r.error); return; }
          parts.push(r.value.data);
          if (i + 1 < file.sliceCount) { next(i + 1); return; }
          file.closeAsync();
          const total = parts.reduce((s, p) => s + p.length, 0);
          const bytes = new Uint8Array(total);
          let at = 0;
          parts.forEach(p => { bytes.set(p, at); at += p.length; });
          resolve({ size: file.size, slices: parts.length, bytes });
        });
        next(0);
      });
    });
  }

  // ---------- automatic probes ----------

  async function pEnvironment() {
    const d = Office.context.diagnostics || {};
    const supported = {};
    Object.keys(SETS).forEach(name => { supported[name] = SETS[name].filter(v => supports(name, v)); });
    const excel = supported.ExcelApi;
    const top = excel.length ? excel[excel.length - 1] : 'none';
    const t = Office.context.officeTheme;
    state.env = {
      probe: VERSION, session: SESSION, host: d.host, platform: d.platform, version: d.version,
      supported, theme: t ? { dark: !!t.isDarkTheme, body: t.bodyBackgroundColor } : null,
      agent: typeof navigator !== 'undefined' ? navigator.userAgent : ''
    };
    const missing = REQUIRED.filter(([n, v]) => !supports(n, v)).map(([n, v]) => `${n} ${v}`);
    return {
      status: missing.length ? 'warn' : 'pass',
      detail: `${d.platform || '?'} Excel ${d.version || '?'}; ExcelApi up to ${top}. ` +
        (missing.length ? `Missing for the add-in: ${missing.join(', ')}.` : 'Every set the spec relies on is present.'),
      data: state.env
    };
  }

  async function pWriteSpeed() {
    const cols = 60;
    const runs = [];
    for (const rows of [200, 1000, 2000]) {
      const run = await Excel.run(async context => {
        const ws = await freshSheet(context, 'Write');
        return withManualCalc(context, async () => {
          const block = 500;
          const t0 = now();
          for (let s = 0; s < rows; s += block) {
            const n = Math.min(block, rows - s);
            const formulas = [];
            const formats = [];
            for (let i = 0; i < n; i++) {
              const row = [s + i + 1];
              const fmt = ['#,##0'];
              for (let c = 1; c < cols; c++) { row.push('=RC[-1]*1.01'); fmt.push('#,##0.00;(#,##0.00);-'); }
              formulas.push(row); formats.push(fmt);
            }
            const range = ws.getRangeByIndexes(s, 0, n, cols);
            range.formulasR1C1 = formulas;
            range.numberFormat = formats;
            await context.sync();
          }
          return { cells: rows * cols, ms: now() - t0 };
        });
      });
      runs.push(run);
    }
    let broadcast;
    try {
      broadcast = await Excel.run(async context => {
        const ws = context.workbook.worksheets.getItem(PREFIX + 'Write');
        return withManualCalc(context, async () => {
          const t0 = now();
          ws.getRangeByIndexes(0, 62, 2000, 58).formulasR1C1 = '=RC[-1]*1.01';
          await context.sync();
          return { cells: 2000 * 58, ms: now() - t0 };
        });
      });
    } catch (e) { broadcast = { error: failure(e).detail }; }
    const biggest = runs[runs.length - 1];
    const perModule = biggest.ms / biggest.cells * 1800;
    const detail = runs.map(r => `${num(r.cells)} cells in ${secs(r.ms)}`).join('; ') +
      (broadcast.error ? `. One formula broadcast to a block failed (${broadcast.error}).` : `. One formula broadcast to ${num(broadcast.cells)} cells took ${secs(broadcast.ms)}.`) +
      ` A 30-row module on 60 months (1,800 cells) would take about ${secs(perModule)}; target under 3 s.`;
    return { status: perModule < 3000 ? 'pass' : 'warn', detail, data: { runs, broadcast, perModuleMs: perModule } };
  }

  async function pRowInserts() {
    return Excel.run(async context => {
      await dropNames(context, [PREFIX + 'Block']);
      const ws = await freshSheet(context, 'Insert');
      const values = [];
      for (let i = 1; i <= 20; i++) values.push(['Line ' + i, i]);
      ws.getRange('A2:B21').values = values;
      ws.getRange('A22:B22').formulas = [['Total', '=SUM(B2:B21)']];
      context.workbook.names.add(PREFIX + 'Block', ws.getRange('B2:B21'));
      ws.names.add(PREFIX + 'Total', ws.getRange('B22'));
      await context.sync();
      ws.getRange('10:14').insert(Excel.InsertShiftDirection.down);
      await context.sync();
      const total = ws.getRange('B27');
      total.load('formulas');
      const block = context.workbook.names.getItem(PREFIX + 'Block').getRange();
      block.load('address');
      const totalName = ws.names.getItem(PREFIX + 'Total').getRange();
      totalName.load('address');
      await context.sync();
      const sumOk = /SUM\(B2:B26\)/i.test(String(total.formulas[0][0]));
      const blockOk = /!\$?B\$?2:\$?B\$?26$/.test(block.address);
      const totalOk = /!\$?B\$?27$/.test(totalName.address);
      const times = [];
      for (let i = 0; i < 40; i++) {
        const t0 = now();
        ws.getRange('5:5').insert(Excel.InsertShiftDirection.down);
        await context.sync();
        times.push(now() - t0);
      }
      const first = avg(times.slice(0, 10));
      const last = avg(times.slice(-10));
      const slowing = last > first * 2;
      const ok = sumOk && blockOk && totalOk;
      return {
        status: ok ? (slowing ? 'warn' : 'pass') : 'fail',
        detail: `After inserting 5 rows inside a block: SUM ${sumOk ? 'grew' : 'did not grow'}, workbook name ${blockOk ? 'grew' : 'did not grow (' + block.address + ')'}, sheet name ${totalOk ? 'moved' : 'did not move'}. ` +
          `40 single-row inserts: first ten ${first.toFixed(0)} ms each, last ten ${last.toFixed(0)} ms each${slowing ? ' (slowing, as in office-js issue 3320)' : ''}.`,
        data: { sumOk, blockOk, totalOk, times }
      };
    });
  }

  async function pNames() {
    return Excel.run(async context => {
      await dropNames(context, [PREFIX + 'Formula']);
      const ws = await freshSheet(context, 'Names');
      const count = 500;
      let t0 = now();
      for (let i = 0; i < count; i++) context.workbook.names.add(`${PREFIX}N${i}`, ws.getRange(`A${i + 1}`));
      await context.sync();
      const addMs = now() - t0;
      context.workbook.names.add(PREFIX + 'Formula', '=10*3');
      const formula = context.workbook.names.getItem(PREFIX + 'Formula');
      formula.load('type,value');
      await context.sync();
      let hiddenSet = true;
      try { context.workbook.names.getItem(`${PREFIX}N0`).visible = false; await context.sync(); } catch (e) { hiddenSet = false; }
      const all = context.workbook.names;
      all.load('items/name,items/visible');
      await context.sync();
      const hiddenListed = all.items.some(n => n.name === `${PREFIX}N0` && n.visible === false);
      t0 = now();
      all.items.filter(n => n.name.indexOf(PREFIX + 'N') === 0).forEach(n => n.delete());
      await context.sync();
      const deleteMs = now() - t0;
      return {
        status: String(formula.value) === '30' ? 'pass' : 'warn',
        detail: `Added ${count} names in ${secs(addMs)} and deleted them in ${secs(deleteMs)}. Named formula returned ${formula.value} (type ${formula.type}). ` +
          (hiddenSet ? `A hidden name is ${hiddenListed ? 'still listed' : 'not listed'} by names.load.` : 'Could not hide a name.'),
        data: { addMs, deleteMs, hiddenSet, hiddenListed }
      };
    });
  }

  async function pCustomXml() {
    return Excel.run(async context => {
      const parts = context.workbook.customXmlParts.getByNamespace(NS);
      parts.load('items/id');
      await context.sync();
      const xmls = parts.items.map(p => p.getXml());
      await context.sync();
      const earlier = xmls.map(x => x.value).find(x => {
        const m = /<session>([^<]*)<\/session>/.exec(x);
        return m && m[1] !== SESSION;
      });
      if (earlier) {
        const written = (/<written>([^<]*)<\/written>/.exec(earlier) || [])[1];
        const where = (/<platform>([^<]*)<\/platform>/.exec(earlier) || [])[1];
        let marker = 'no settings marker found (run the automatic tests before saving)';
        try {
          const raw = Office.context.document.settings.get('hfgProbeMarker');
          const m = raw ? JSON.parse(raw) : null;
          if (m && m.session !== SESSION) marker = `the settings marker from ${m.written} survived too`;
        } catch (e) { /* ignore */ }
        return { status: 'pass', detail: `Metadata written on ${where || 'another platform'} at ${written} survived save, close and reopen; ${marker}.` };
      }
      parts.items.forEach(p => p.delete());
      const platform = (Office.context.diagnostics || {}).platform || 'unknown';
      const small = `<probe xmlns="${NS}"><session>${SESSION}</session><written>${new Date().toISOString()}</written><platform>${platform}</platform><pad>${'x'.repeat(100000)}</pad></probe>`;
      let t0 = now();
      const part = context.workbook.customXmlParts.add(small);
      part.load('id');
      await context.sync();
      const smallMs = now() - t0;
      let big = 'not tried';
      try {
        t0 = now();
        const bigPart = context.workbook.customXmlParts.add(`<probeBig xmlns="${NS}:big"><pad>${'y'.repeat(1000000)}</pad></probeBig>`);
        const back = bigPart.getXml();
        await context.sync();
        big = back.value.length > 1000000 ? `a 1 MB part read back in ${secs(now() - t0)}` : 'a 1 MB part came back short';
        bigPart.delete();
        await context.sync();
      } catch (e) { big = 'a 1 MB part failed: ' + failure(e).detail; }
      return {
        status: 'manual',
        detail: `Wrote a 100 KB metadata part in ${secs(smallMs)}; ${big}. Now save, close and reopen the workbook (try another platform too), then run this probe again.`
      };
    });
  }

  async function pSettings() {
    const settings = Office.context.document.settings;
    const out = [];
    for (const size of [100000, 1000000, 3000000]) {
      try {
        const t0 = now();
        settings.set('hfgProbePad', 'x'.repeat(size));
        await saveSettings();
        out.push(`${num(size)} characters saved without error in ${secs(now() - t0)}`);
      } catch (e) { out.push(`${num(size)} characters failed (${failure(e).detail})`); }
    }
    settings.remove('hfgProbePad');
    settings.set('hfgProbeMarker', JSON.stringify({ session: SESSION, written: new Date().toISOString() }));
    await saveSettings();
    const failed = out.some(s => /failed/.test(s));
    return {
      status: failed ? 'warn' : 'pass',
      detail: 'Document settings: ' + out.join('; ') + '. A small marker is left for the save and reopen check.'
    };
  }

  async function pStyles() {
    return Excel.run(async context => {
      const ws = await freshSheet(context, 'Styles');
      if (supports('ExcelApi', '1.14')) {
        const old = context.workbook.styles.getItemOrNullObject(STYLE);
        await context.sync();
        if (!old.isNullObject) { old.delete(); await context.sync(); }
      }
      context.workbook.styles.add(STYLE);
      const style = context.workbook.styles.getItem(STYLE);
      style.font.name = 'Segoe UI';
      style.font.size = 9;
      style.font.color = '#404040';
      style.fill.color = '#FFF2CC';
      style.numberFormat = '#,##0;(#,##0);-';
      style.locked = false;
      await context.sync();
      const range = ws.getRange('B2:M30');
      range.style = STYLE;
      range.values = 1234.5;
      range.load('style');
      await context.sync();
      return {
        status: range.style === STYLE ? 'pass' : 'fail',
        detail: `Named style created and applied to 348 cells; the range reports style "${range.style}". Excel cannot undo style changes, so the add-in warns first.`
      };
    });
  }

  async function pControls() {
    return Excel.run(async context => {
      await dropNames(context, [PREFIX + 'List']);
      const ws = await freshSheet(context, 'Controls');
      ws.getRange('E2:E4').values = [['Base'], ['Upside'], ['Downside']];
      context.workbook.names.add(PREFIX + 'List', ws.getRange('E2:E4'));
      const pick = ws.getRange('B2');
      pick.values = [['Base']];
      pick.dataValidation.rule = { list: { inCellDropDown: true, source: '=' + PREFIX + 'List' } };
      const box = ws.getRange('B4');
      const checkbox = supports('ExcelApi', '1.18');
      if (checkbox) { box.values = [[true]]; box.control = { type: Excel.CellControlType.checkbox }; }
      const cfRange = ws.getRange('B6:B10');
      cfRange.values = [[1], [-2], [3], [-4], [5]];
      const cf = cfRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      cf.custom.rule.formula = '=B6<0';
      cf.custom.format.font.color = '#C00000';
      let note = 'not tried (needs ExcelApi 1.18)';
      if (checkbox) {
        try { ws.notes.add(ws.getRange('B2'), 'Probe note: scenario picker'); await context.sync(); note = 'added'; } catch (e) { note = 'failed: ' + failure(e).detail; }
      }
      await context.sync();
      const dv = pick.dataValidation;
      dv.load('type');
      const cfCount = cfRange.conditionalFormats.getCount();
      if (checkbox) box.load('control');
      await context.sync();
      const boxOk = checkbox && box.control && box.control.type === 'Checkbox';
      return {
        status: dv.type === 'List' && cfCount.value > 0 && (boxOk || !checkbox) ? 'pass' : 'warn',
        detail: `Validation list: ${dv.type}. In-cell checkbox: ${checkbox ? (boxOk ? 'works' : 'did not stick') : 'not available'}. ` +
          `Conditional format rules: ${cfCount.value}. Note: ${note}. Check sheet zProbe_Controls: the drop-down in B2 and the checkbox in B4 should both work.`
      };
    });
  }

  async function pOutline() {
    return Excel.run(async context => {
      const ws = await freshSheet(context, 'Outline');
      const rows = [];
      for (let i = 1; i <= 20; i++) rows.push(['Row ' + i, i]);
      ws.getRange('A1:B20').values = rows;
      ws.getRange('3:8').group(Excel.GroupOption.byRows);
      ws.getRange('D:F').group(Excel.GroupOption.byColumns);
      ws.showOutlineLevels(1, 1);
      ws.freezePanes.freezeAt(ws.getRange('A1:I15'));
      ws.tabColor = '#575F46';
      ws.showGridlines = false;
      await context.sync();
      const hidden = ws.getRange('A5');
      hidden.load('rowHidden');
      const frozen = ws.freezePanes.getLocationOrNullObject();
      frozen.load('address');
      await context.sync();
      return {
        status: hidden.rowHidden && !frozen.isNullObject ? 'pass' : 'warn',
        detail: `Grouped rows and columns, collapsed to level 1 (row 5 hidden: ${hidden.rowHidden}). Frozen at ${frozen.isNullObject ? 'nothing' : frozen.address}. Tab colour and gridlines set.`
      };
    });
  }

  async function pTrace() {
    return Excel.run(async context => {
      const a = await freshSheet(context, 'TraceA');
      const b = await freshSheet(context, 'TraceB');
      const values = [];
      const formulas = [];
      for (let i = 2; i <= 1001; i++) { values.push([i]); formulas.push([`='${PREFIX}TraceA'!B${i}*2`]); }
      a.getRange('B2:B1001').values = values;
      b.getRange('B2:B1001').formulas = formulas;
      b.getRange('B1002').formulas = [['=SUM(B2:B1001)']];
      await context.sync();
      const out = {};
      let t0 = now();
      const direct = b.getRange('B1002').getDirectPrecedents();
      direct.load('addresses');
      await context.sync();
      out.direct = { ms: now() - t0, addresses: direct.addresses };
      if (supports('ExcelApi', '1.14')) {
        t0 = now();
        const all = b.getRange('B1002').getPrecedents();
        all.load('addresses');
        await context.sync();
        out.all = { ms: now() - t0, addresses: all.addresses };
      }
      if (supports('ExcelApi', '1.15')) {
        t0 = now();
        const deps = a.getRange('B2').getDependents();
        deps.load('addresses');
        await context.sync();
        out.deps = { ms: now() - t0, addresses: deps.addresses };
      }
      t0 = now();
      const many = b.getRange('B2:B1001').getDirectPrecedents();
      many.load('addresses');
      await context.sync();
      out.many = { ms: now() - t0 };
      const crossSheet = out.all && out.all.addresses.join(' ').indexOf('TraceA') >= 0;
      return {
        status: crossSheet || !out.all ? 'pass' : 'warn',
        detail: `Direct precedents of a total: ${out.direct.addresses.join(', ')} in ${secs(out.direct.ms)}. ` +
          (out.all ? `All precedents ${crossSheet ? 'reach the other sheet' : 'stay on one sheet'} (${secs(out.all.ms)}). ` : '') +
          (out.deps ? `Dependents of an input: ${out.deps.addresses.join(', ')} (${secs(out.deps.ms)}). ` : '') +
          `Precedents of 1,000 cells in ${secs(out.many.ms)}.`,
        data: out
      };
    });
  }

  async function pErrorScan() {
    return Excel.run(async context => {
      const ws = await freshSheet(context, 'Errors');
      ws.getRange('B2:B3').formulas = [['=1/0'], ['=NA()']];
      await context.sync();
      const t0 = now();
      const sheets = context.workbook.worksheets;
      sheets.load('items/name');
      await context.sync();
      const found = sheets.items.map(s => {
        const cells = s.getRange().getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas, Excel.SpecialCellValueType.errors);
        cells.load('address,cellCount');
        return { name: s.name, cells };
      });
      await context.sync();
      const ms = now() - t0;
      const withErrors = found.filter(f => !f.cells.isNullObject);
      const total = withErrors.reduce((s, f) => s + f.cells.cellCount, 0);
      const ours = withErrors.find(f => f.name === PREFIX + 'Errors');
      return {
        status: ours && ours.cells.cellCount === 2 ? 'pass' : 'fail',
        detail: `Scanned ${sheets.items.length} sheets in ${secs(ms)}: ${total} error cells on ${withErrors.length} sheets, including the 2 planted ones.`,
        data: withErrors.map(f => ({ sheet: f.name, cells: f.cells.cellCount }))
      };
    });
  }

  async function pFormulaRead() {
    return Excel.run(async context => {
      let ws = context.workbook.worksheets.getItemOrNullObject(PREFIX + 'Write');
      await context.sync();
      if (ws.isNullObject) return { status: 'skip', detail: 'Run the write speed probe first.' };
      const used = ws.getUsedRange(true);
      used.load('rowCount,columnCount');
      await context.sync();
      const t0 = now();
      const blocks = [];
      const step = 1000;
      for (let r = 0; r < used.rowCount; r += step) {
        const rng = ws.getRangeByIndexes(r, 0, Math.min(step, used.rowCount - r), used.columnCount);
        rng.load('formulasR1C1');
        blocks.push(rng);
      }
      await context.sync();
      const ms = now() - t0;
      let breaks = 0;
      let cells = 0;
      blocks.forEach(rng => rng.formulasR1C1.forEach(row => {
        cells += row.length;
        for (let c = 2; c < 60 && c < row.length; c++) if (row[c] !== row[1]) breaks++;
      }));
      return {
        status: 'pass',
        detail: `Read ${num(cells)} formulas in ${secs(ms)} and compared each row: ${breaks} break${breaks === 1 ? '' : 's'} in the pattern. This is the consistency scan's raw speed.`
      };
    });
  }

  async function pEvents() {
    return Excel.run(async context => {
      const ws = await freshSheet(context, 'Events');
      ws.getRange('A1:A5').values = [[1], [2], [3], [4], [5]];
      await context.sync();
      const seen = [];
      const handler = ws.onChanged.add(async e => { seen.push({ type: e.changeType, address: e.address, source: e.source, trigger: e.triggerSource }); });
      await context.sync();
      ws.getRange('3:3').insert(Excel.InsertShiftDirection.down);
      await context.sync();
      ws.getRange('A1').values = [[10]];
      await context.sync();
      await sleep(1500);
      handler.remove();
      await context.sync();
      const rowInsert = seen.find(s => /RowInserted/i.test(String(s.type)));
      return {
        status: rowInsert ? 'pass' : 'warn',
        detail: `Saw ${seen.length} change events: ${seen.map(s => `${s.type} at ${s.address}${s.trigger ? ' from ' + s.trigger : ''}`).join('; ') || 'none'}.`,
        data: seen
      };
    });
  }

  async function pChartsShapes() {
    return Excel.run(async context => {
      const ws = await freshSheet(context, 'Charts');
      const data = [['Month', 'Cost', 'Budget']];
      for (let m = 1; m <= 12; m++) data.push([`M${m}`, 100 + m * 7, 150]);
      ws.getRange('A1:C13').values = data;
      const chart = ws.charts.add(Excel.ChartType.columnClustered, ws.getRange('A1:C13'), Excel.ChartSeriesBy.columns);
      chart.setPosition('E2', 'L18');
      chart.title.text = 'Probe chart';
      let waterfall = 'added';
      try {
        const wf = ws.charts.add(Excel.ChartType.waterfall, ws.getRange('A1:B13'), Excel.ChartSeriesBy.columns);
        wf.setPosition('E20', 'L36');
        await context.sync();
      } catch (e) { waterfall = 'failed: ' + failure(e).detail; }
      const shape = ws.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
      shape.left = 600; shape.top = 20; shape.width = 120; shape.height = 30;
      shape.textFrame.textRange.text = 'Probe shape';
      await context.sync();
      let layoutResult = 'set: landscape, print area, print titles, fit to width, footer';
      try {
        const layout = ws.pageLayout;
        layout.orientation = Excel.PageOrientation.landscape;
        layout.setPrintArea('A1:L36');
        layout.setPrintTitleRows('$1:$1');
        layout.zoom = { horizontalFitToPages: 1 };
        layout.headersFooters.defaultForAllPages.centerFooter = '&P of &N';
        await context.sync();
      } catch (e) { layoutResult = 'failed: ' + failure(e).detail; }
      ws.charts.load('items/name');
      ws.shapes.load('items/type');
      await context.sync();
      const ok = waterfall === 'added' && layoutResult.indexOf('failed') < 0;
      return {
        status: ok ? 'pass' : 'warn',
        detail: `${ws.charts.items.length} charts and ${ws.shapes.items.length} shapes on the sheet; waterfall chart ${waterfall}. Page layout ${layoutResult}.`
      };
    });
  }

  async function pProtection() {
    return Excel.run(async context => {
      const ws = await freshSheet(context, 'Protect');
      ws.protection.protect({ allowFormatCells: false, allowInsertRows: false }, 'probe-pass');
      ws.protection.load('protected');
      await context.sync();
      const on = ws.protection.protected;
      ws.protection.unprotect('probe-pass');
      ws.protection.load('protected');
      await context.sync();
      return {
        status: on && !ws.protection.protected ? 'pass' : 'fail',
        detail: `Protected with a password (${on}) and unprotected again (${!ws.protection.protected}).`
      };
    });
  }

  async function pPaneWidth() {
    if (!supports('TaskPaneApi', '1.1')) return { status: 'skip', detail: 'TaskPaneApi 1.1 is not available here.' };
    try {
      Office.extensionLifeCycle.taskpane.setWidth(480);
      await sleep(500);
      const width = typeof root.innerWidth === 'number' ? root.innerWidth : null;
      return {
        status: 'pass',
        detail: `Asked for a 480 px task pane (web allows 350 to 500 px); the pane now reports ${width == null ? 'an unknown' : width + ' px'} width.`
      };
    } catch (e) { return failure(e); }
  }

  // ---------- interactive probes (driven from the task pane) ----------

  async function readChosenFile(file) {
    const buffer = await file.arrayBuffer();
    const sheets = await HfgZip.listSheets(buffer);
    state.file = { name: file.name, buffer, sheets };
    return sheets;
  }

  async function insertFromFile(names) {
    if (!state.file) throw new Error('Choose a file first');
    const base64 = HfgZip.base64FromBuffer(state.file.buffer);
    return Excel.run(async context => {
      const wb = context.workbook;
      wb.names.load('items/name');
      await context.sync();
      const namesBefore = wb.names.items.length;
      const t0 = now();
      const res = wb.insertWorksheetsFromBase64(base64, { sheetNamesToInsert: names, positionType: Excel.WorksheetPositionType.end });
      await context.sync();
      const ms = now() - t0;
      const added = (res.value || []).map(id => wb.worksheets.getItem(id));
      added.forEach(s => s.load('name'));
      await context.sync();
      const inserted = added.map(s => s.name);
      state.inserted = state.inserted.concat(inserted);
      const checks = inserted.map(name => {
        const ws = wb.worksheets.getItem(name);
        const shapes = ws.shapes; shapes.load('items/type,items/name');
        const charts = ws.charts; charts.load('items/name');
        const own = ws.names; own.load('items/name');
        const cf = ws.getRange().conditionalFormats.getCount();
        const dv = ws.getRange().getSpecialCellsOrNullObject(Excel.SpecialCellType.dataValidations); dv.load('cellCount');
        const comments = supports('ExcelApi', '1.10') ? ws.comments : null; if (comments) comments.load('items/id');
        const notes = supports('ExcelApi', '1.18') ? ws.notes : null; if (notes) notes.load('items');
        ws.protection.load('protected');
        return { name, ws, shapes, charts, own, cf, dv, comments, notes };
      });
      wb.names.load('items/name');
      await context.sync();
      const sheetLines = checks.map(c => {
        const types = {};
        c.shapes.items.forEach(s => { types[s.type] = (types[s.type] || 0) + 1; });
        const shapeText = Object.keys(types).map(t => `${types[t]} ${t}`).join(', ') || 'no shapes';
        return `${c.name}: ${shapeText}; ${c.charts.items.length} charts; ${c.own.items.length} sheet names; ${c.cf.value} conditional formats; ` +
          `${c.dv.isNullObject ? 0 : c.dv.cellCount} validated cells; ${c.comments ? c.comments.items.length : '?'} comments; ${c.notes ? c.notes.items.length : '?'} notes; protected ${c.ws.protection.protected}`;
      });
      return {
        status: inserted.length === names.length ? 'manual' : 'warn',
        detail: `Inserted ${inserted.length} of ${names.length} sheets from ${state.file.name} (${num(state.file.buffer.byteLength)} bytes) in ${secs(ms)}; workbook names went from ${namesBefore} to ${wb.names.items.length}. ` +
          sheetLines.join(' | ') + '. Look at the sheets: did the form controls (drop-downs and check boxes) come across and still work? Answer below.',
        data: { inserted, ms }
      };
    });
  }

  // Monte Carlo: calculation mode and the native data table ------------------

  async function pCalcMode() {
    return Excel.run(async context => {
      const app = context.workbook.application;
      app.load('calculationMode');
      await context.sync();
      const before = app.calculationMode;
      app.calculationMode = Excel.CalculationMode.automaticExceptTables;
      await context.sync();
      app.load('calculationMode');
      await context.sync();
      const set = app.calculationMode;
      app.calculationMode = before;
      await context.sync();
      const ok = set === Excel.CalculationMode.automaticExceptTables;
      return {
        status: ok ? 'pass' : 'fail',
        detail: `Calculation mode was ${before}; set it to automatic except tables and read back ${set}; put it back. ` +
          'Excel keeps one mode for the whole session, so the add-in checks it whenever a simulation model is opened.'
      };
    });
  }

  const SIM_NAMES = ['MC_Trials', 'MC_Ran', 'Chk_Errors', 'MC_Breach_Prob'];

  async function simulationRun() {
    return Excel.run(async context => {
      const wb = context.workbook;
      const items = SIM_NAMES.map(n => wb.names.getItemOrNullObject(n));
      const run = wb.worksheets.getItemOrNullObject('Run');
      await context.sync();
      if (run.isNullObject) {
        return { status: 'fail', detail: 'No Run sheet here. Open the Monte Carlo demo workbook (or insert all its sheets) first.' };
      }
      const missing = SIM_NAMES.filter((n, i) => items[i].isNullObject);
      if (missing.length) {
        return { status: 'fail', detail: `The Run sheet is here but these names are missing: ${missing.join(', ')}. If the sheets were inserted from the file, names did not come across.` };
      }
      const ranges = items.map(n => n.getRange());
      ranges.forEach(r => r.load('values'));
      const cell = run.getRange('C7');
      cell.load('formulas');
      const app = wb.application;
      app.load('calculationMode');
      await context.sync();
      const mode = app.calculationMode;
      const t0 = now();
      app.calculate(Excel.CalculationType.full);
      await context.sync();
      const ms = now() - t0;
      ranges.forEach(r => r.load('values'));
      await context.sync();
      const [trials, ran, errors, breach] = ranges.map(r => r.values[0][0]);
      const ok = ran === 1 && errors === 0;
      return {
        status: ok ? 'pass' : 'fail',
        detail: `Calculation mode ${mode}. Run!C7 holds ${JSON.stringify(cell.formulas[0][0])}. A full calculation took ${secs(ms)} for ` +
          `${trials} trials (${(ms / Math.max(1, trials)).toFixed(1)} ms a trial). Data table filled: ${ran === 1 ? 'yes' : 'no'}; ` +
          `error checks failing: ${errors} (one of them compares the results with the Python reference); breach probability ${breach}.` +
          (ran === 1 ? '' : ' If the table is still empty, press F9 in the workbook and run this again, then note that a full calculation from Office.js did not run the table.'),
        data: { ms, trials, ran, errors, breach, mode }
      };
    });
  }

  async function copyInsertedSheet() {
    const name = state.inserted[0] || PREFIX + 'Controls';
    return Excel.run(async context => {
      const ws = context.workbook.worksheets.getItem(name);
      ws.shapes.load('items/type');
      const copy = ws.copy(Excel.WorksheetPositionType.end);
      copy.load('name');
      copy.shapes.load('items/type');
      await context.sync();
      state.inserted.push(copy.name);
      return {
        status: copy.shapes.items.length === ws.shapes.items.length ? 'pass' : 'warn',
        detail: `Copied ${name} to ${copy.name}: shapes ${ws.shapes.items.length} before, ${copy.shapes.items.length} after.`
      };
    });
  }

  async function exportPdf() {
    const t0 = now();
    const file = await getFile(Office.FileType.Pdf);
    const head = String.fromCharCode.apply(null, file.bytes.subarray(0, 5));
    const text = new TextDecoder('latin1').decode(file.bytes);
    const pages = (text.match(/\/Type\s*\/Page[^s]/g) || []).length;
    if (state.pdfUrl) URL.revokeObjectURL(state.pdfUrl);
    state.pdfUrl = URL.createObjectURL(new Blob([file.bytes], { type: 'application/pdf' }));
    return {
      status: head === '%PDF-' ? 'manual' : 'fail',
      detail: `PDF of ${num(file.size)} bytes in ${file.slices} slice(s), about ${pages} pages, in ${secs(now() - t0)}. Download it and check which sheets it holds and whether print areas and page setup were respected.`
    };
  }

  async function exportCompressed() {
    const t0 = now();
    const file = await getFile(Office.FileType.Compressed);
    const sheets = await HfgZip.listSheets(file.bytes.buffer);
    const parts = HfgZip.listParts(file.bytes.buffer);
    state.exported = { bytes: file.bytes, base64: HfgZip.base64FromBuffer(file.bytes.buffer) };
    const live = await Excel.run(async context => {
      const ws = context.workbook.worksheets;
      ws.load('items/name');
      await context.sync();
      return ws.items.map(w => w.name);
    });
    const same = live.length === sheets.length && live.every(n => sheets.some(s => s.name === n));
    const custom = parts.filter(p => p.indexOf('customXml/') === 0 && /item\d+\.xml$/.test(p)).length;
    const controls = parts.filter(p => p.indexOf('xl/ctrlProps/') === 0).length;
    return {
      status: same ? 'pass' : 'warn',
      detail: `Read the open workbook as a file: ${num(file.size)} bytes, ${parts.length} parts (${custom} custom XML, ${controls} form control parts) in ${secs(now() - t0)}. ` +
        `Its sheet list ${same ? 'matches' : 'does not match'} the open workbook. This is what lets the package writer edit a live model.`
    };
  }

  async function openCopy() {
    if (!state.exported) throw new Error('Run "Read the workbook as a file" first');
    await Excel.createWorkbook(state.exported.base64);
    return { status: 'manual', detail: 'Asked Excel to open a copy built from the exported file. Did a new workbook open with the probe sheets in it? Answer below.' };
  }

  async function undoSetup() {
    await Excel.run(async context => {
      const ws = await freshSheet(context, 'Undo');
      ws.getRange('A2:A4').values = [['First write'], ['Second write'], ['Third write']];
      await context.sync();
    });
    if (!supports('ExcelApi', '1.20')) return { status: 'skip', detail: 'ExcelApi 1.20 (one-step undo) is not available here.' };
    await Excel.run({ mergeUndoGroup: true }, async context => {
      const ws = context.workbook.worksheets.getItem(PREFIX + 'Undo');
      ws.getRange('B2').values = [[1]]; await context.sync();
      ws.getRange('B3').values = [[2]]; await context.sync();
      ws.getRange('B4').values = [[3]]; await context.sync();
    });
    return { status: 'manual', detail: 'Wrote B2, B3 and B4 on zProbe_Undo in three steps inside one undo group. Click in the sheet, press Ctrl+Z (Cmd+Z on Mac) once, then answer below.' };
  }

  async function ribbonDisable(enabled) {
    if (!supports('RibbonApi', '1.1')) return { status: 'skip', detail: 'RibbonApi 1.1 is not available here.' };
    await Office.ribbon.requestUpdate({ tabs: [{ id: 'HFG.Probe.Tab', groups: [{ id: 'HFG.Probe.Group', controls: [{ id: 'HFG.Probe.Toggle', enabled }] }] }] });
    return { status: 'manual', detail: `Asked the ribbon to ${enabled ? 'enable' : 'grey out'} the Toggle test button. Is it ${enabled ? 'enabled' : 'greyed out'}? Answer below.` };
  }

  async function contextMenuDisable(enabled) {
    if (!supports('ContextMenuApi', '1.1')) return { status: 'skip', detail: 'ContextMenuApi 1.1 is not available here.' };
    await Office.contextMenu.requestUpdate({ controls: [{ id: 'HFG.Probe.Ctx.Mark', enabled }] });
    return { status: 'manual', detail: `Asked the right-click menu to ${enabled ? 'enable' : 'grey out'} HFG Probe > Mark this cell. Right-click a cell: is it ${enabled ? 'enabled' : 'greyed out'}? Answer below.` };
  }

  function openDialog() {
    return new Promise(resolve => {
      const url = new URL('dialog.html', root.location.href).href;
      let done = false;
      Office.context.ui.displayDialogAsync(url, { height: 30, width: 30 }, res => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) { done = true; resolve(failure(res.error)); return; }
        const dialog = res.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, arg => {
          if (done) return;
          done = true;
          dialog.close();
          resolve({ status: 'pass', detail: `Dialog opened and sent back "${arg.message}".` });
        });
      });
      setTimeout(() => { if (!done) { done = true; resolve({ status: 'fail', detail: 'No message from the dialog within 20 seconds.' }); } }, 20000);
    });
  }

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src; s.onload = resolve; s.onerror = () => reject(new Error('Could not load ' + src + ' (run npm install in addin-probe)'));
      document.head.appendChild(s);
    });
  }

  async function signIn(clientId, tenantId) {
    if (!supports('NestedAppAuth', '1.1')) return { status: 'skip', detail: 'Nested app authentication is not available here (on the web it needs a file in SharePoint or OneDrive).' };
    if (!clientId || !tenantId) return { status: 'skip', detail: 'Enter the application (client) ID and the directory (tenant) ID first.' };
    if (!root.msal) await loadScript('../node_modules/@azure/msal-browser/lib/msal-browser.min.js');
    const t0 = now();
    const pca = await root.msal.createNestablePublicClientApplication({
      auth: { clientId, authority: 'https://login.microsoftonline.com/' + tenantId },
      cache: { cacheLocation: 'localStorage' }
    });
    let loginHint;
    try { loginHint = (await Office.auth.getAuthContext()).userPrincipalName; } catch (e) { /* not available on this platform */ }
    const request = { scopes: ['User.Read'], loginHint };
    let result;
    let how = 'silently';
    try { result = await pca.ssoSilent(request); }
    catch (e) {
      if (root.msal.InteractionRequiredAuthError && !(e instanceof root.msal.InteractionRequiredAuthError)) throw e;
      how = 'after a prompt';
      result = await pca.acquireTokenPopup(request);
    }
    const response = await fetch('https://graph.microsoft.com/v1.0/me', { headers: { Authorization: 'Bearer ' + result.accessToken } });
    if (!response.ok) return { status: 'fail', detail: `Got a token ${how}, but Microsoft Graph returned ${response.status}.` };
    const me = await response.json();
    const domain = String(me.userPrincipalName || '').split('@')[1] || 'unknown';
    return { status: 'pass', detail: `Signed in ${how} in ${secs(now() - t0)} and called Microsoft Graph as a user in ${domain}. The Home Hub call will work the same way.` };
  }

  // ---------- commands (ribbon, menu, right-click, shortcut) ----------

  function command(id, label) {
    return async function (event) {
      try {
        await Excel.run(async context => {
          const cell = context.workbook.getActiveCell();
          cell.getOffsetRange(0, 1).values = [[`${label} ${new Date().toLocaleTimeString('en-NZ')}`]];
          await context.sync();
        });
        record(id, { status: 'pass', detail: `${label} ran and wrote next to the selected cell.` });
      } catch (e) { record(id, failure(e)); }
      if (event && typeof event.completed === 'function') event.completed();
    };
  }

  function registerCommands() {
    if (!Office.actions || typeof Office.actions.associate !== 'function') return;
    Office.actions.associate('ribbonMark', command('cmd-ribbon', 'Ribbon button'));
    Office.actions.associate('menuMark', command('cmd-menu', 'Ribbon menu item'));
    Office.actions.associate('contextMark', command('cmd-context', 'Right-click item'));
    Office.actions.associate('SHORTCUTMARK', command('cmd-shortcut', 'Keyboard shortcut'));
    Office.actions.associate('ribbonToggle', command('cmd-toggle', 'Toggle test button'));
    Office.actions.associate('openView', openView);
    const C = CMD();
    if (C) C.SHORTCUTS.forEach(sc => Office.actions.associate(sc.action, shortcutView(sc.key)));
  }


  // ---------- the designed ribbon (src/commands.js) ----------

  const CMD = () => root.HfgCommands;
  state.view = null;          // the command whose placeholder view is open
  state.clicks = [];          // {key, sourceId, at} for every designed command clicked
  state.buildTab = { created: false, visible: false };

  function keyFromSource(id) {
    if (!id) return null;
    if (id.indexOf('HFG.ctx.') === 0) return id.slice(8);
    if (id.indexOf('HFG.') === 0) return id.slice(4);
    return id;
  }

  function showView(key, sourceId) {
    const c = CMD() && CMD().find(key);
    state.view = c ? key : null;
    state.clicks.push({ key, sourceId: sourceId || '', at: new Date().toISOString() });
    const seen = Array.from(new Set(state.clicks.map(x => x.key)));
    const missing = state.clicks.filter(x => !x.sourceId).length;
    record('ribbon-full', {
      status: 'manual',
      detail: `${state.clicks.length} clicks on ${seen.length} designed commands (${seen.slice(-8).join(', ')}). ` +
        (missing ? `${missing} clicks arrived without a control id. ` : 'Every click carried its control id. ') +
        'Did each one open the matching view in the pane? Answer below.'
    });
    if (typeof api.onView === 'function') { try { api.onView(state.view); } catch (e) { /* ignore */ } }
    if (Office.addin && typeof Office.addin.showAsTaskpane === 'function') {
      Office.addin.showAsTaskpane().catch(() => {});   // not awaited: see office-js issue 3250
    }
  }

  async function openView(event) {
    const sourceId = event && event.source ? event.source.id : '';
    const key = keyFromSource(sourceId);
    try {
      if (key === 'sys-builder') await toggleBuildTab();
      showView(key, sourceId);
    } catch (e) { record('ribbon-full', failure(e)); }
    if (event && typeof event.completed === 'function') event.completed();
  }

  function shortcutView(key) {
    return function (event) {
      showView(key, 'shortcut');
      record('cmd-shortcut', { status: 'pass', detail: `Shortcut for ${key} ran.` });
      if (event && typeof event.completed === 'function') event.completed();
    };
  }

  function iconSet(key, origin) {
    return [16, 32, 80].map(size => ({ size, sourceLocation: `${origin}/assets/cmd/${key}-${size}.png` }));
  }

  // The Build tab as a runtime contextual tab: the XML manifest allows one custom tab.
  function buildTabDefinition(origin) {
    const C = CMD();
    const tip = c => ({ title: c.label, description: c.tip });
    const control = c => c.type === 'menu'
      ? { type: 'Menu', id: C.controlId(c.key), label: c.label, superTip: tip(c), icon: iconSet(c.key, origin),
          items: c.items.map(i => ({ type: 'MenuItem', id: C.controlId(i.key), label: i.label, superTip: tip(i),
            icon: iconSet(c.key, origin), actionId: 'hfgOpenView', enabled: true })) }
      : { type: 'Button', id: C.controlId(c.key), label: c.label, superTip: tip(c), icon: iconSet(c.key, origin),
          actionId: 'hfgOpenView', enabled: true };
    return {
      version: '1.0',
      actions: [{ id: 'hfgOpenView', type: 'ExecuteFunction', functionName: 'openView' }],
      tabs: [{
        id: C.BUILD_TAB.id, label: C.BUILD_TAB.label, visible: false,
        groups: C.BUILD.map(g => ({ id: 'HFG.BG.' + g.id, label: g.label, icon: iconSet(g.controls[0].key, origin), controls: g.controls.map(control) }))
      }]
    };
  }

  async function toggleBuildTab(force) {
    if (!supports('RibbonApi', '1.2')) {
      const outcome = { status: 'skip', detail: 'RibbonApi 1.2 is not available here, so the Build tab cannot be created; the HFG add-in would show the builder commands in the task pane instead.' };
      record('build-tab', outcome);
      return outcome;
    }
    const origin = (root.location && root.location.origin) || 'https://localhost:3000';
    if (!state.buildTab.created) {
      const t0 = now();
      await Office.ribbon.requestCreateControls(buildTabDefinition(origin));
      state.buildTab.created = true;
      state.buildTab.createMs = now() - t0;
    }
    const visible = typeof force === 'boolean' ? force : !state.buildTab.visible;
    await Office.ribbon.requestUpdate({ tabs: [{ id: CMD().BUILD_TAB.id, visible }] });
    state.buildTab.visible = visible;
    const groups = CMD().BUILD.length;
    const controls = CMD().BUILD.reduce((n, g) => n + g.controls.length, 0);
    const outcome = {
      status: 'manual',
      detail: `Created the ${CMD().BUILD_TAB.label} tab at runtime (${groups} groups, ${controls} controls) in ${secs(state.buildTab.createMs || 0)} and asked Excel to ${visible ? 'show' : 'hide'} it. ` +
        `Is it ${visible ? 'showing at the right of the ribbon, with Charts among its groups' : 'gone'}? Answer below.`
    };
    record('build-tab', outcome);
    return outcome;
  }

  async function hideGroup(visible) {
    if (!supports('RibbonApi', '1.3')) return { status: 'skip', detail: 'RibbonApi 1.3 is not available here, so groups and buttons can only be greyed out, not hidden.' };
    await Office.ribbon.requestUpdate({ tabs: [{ id: CMD().MAIN_TAB.id, groups: [{ id: 'HFG.G.analysis', visible }] }] });
    return { status: 'manual', detail: `Asked Excel to ${visible ? 'show' : 'hide'} the Analysis group on the ${CMD().MAIN_TAB.label} tab. Is it ${visible ? 'back' : 'hidden'}? Answer below.` };
  }

  // ---------- charts the Build tab's Charts group needs ----------

  const ZDATA = [
    ['Month', 'AC', 'FC', 'Budget', 'AC cumulative', 'FC cumulative', 'Budget cumulative', 'MAT'],
    ['Apr', 8, null, 6, 8, null, 6, 64], ['May', 8, null, 6, 16, null, 12, 64.5], ['Jun', 13, null, 5, 29, null, 17, 69],
    ['Jul', 7, null, 10, 36, null, 27, 69], ['Aug', 6, null, 8, 42, null, 35, 67.5], ['Sep', 2, null, 3, 44, 44, 38, 62],
    ['Oct', null, 6, 8, null, 50, 46, 60.5], ['Nov', null, 9, 7, null, 59, 53, 62], ['Dec', null, 7, 6, null, 66, 59, 61.5],
    ['Jan', null, 7, 7, null, 73, 66, 61], ['Feb', null, 4, 4, null, 77, 70, 57.5], ['Mar', null, 5, 6, null, 82, 76, 55]
  ];

  async function step(log, label, fn, context) {
    try { await fn(); await context.sync(); log.push(`${label}: yes`); return true; }
    catch (e) { log.push(`${label}: no (${failure(e).detail})`); return false; }
  }

  async function pChartZ() {
    return Excel.run(async context => {
      const ws = await freshSheet(context, 'ZChart');
      ws.getRange('A1:H13').values = ZDATA.map(r => r.map(v => (v === null ? '' : v)));
      await context.sync();
      const log = [];
      const chart = ws.charts.add(Excel.ChartType.columnClustered, ws.getRange('A1:H13'), Excel.ChartSeriesBy.columns);
      chart.setPosition('J2', 'T24');
      chart.title.text = 'Revenue (Z chart)';
      await context.sync();
      const s = i => chart.series.getItemAt(i);
      await step(log, 'cumulative and MAT series switched to lines', async () => { [3, 4, 5, 6].forEach(i => { s(i).chartType = Excel.ChartType.lineMarkers; }); }, context);
      await step(log, 'forecast cumulative dashed', async () => { s(4).format.line.lineStyle = Excel.ChartLineStyle.dash; }, context);
      await step(log, 'budget cumulative grey', async () => { s(5).format.line.color = '#A6A6A6'; }, context);
      await step(log, 'actual bars solid dark', async () => { s(0).format.fill.setSolidColor('#262626'); }, context);
      await step(log, 'forecast bars outlined (no hatch through Office.js)', async () => { s(1).format.fill.clear(); s(1).format.border.color = '#262626'; s(1).format.border.lineStyle = Excel.ChartLineStyle.continuous; }, context);
      await step(log, 'budget bars grey on the secondary axis', async () => { s(2).format.fill.setSolidColor('#D9D9D9'); s(2).axisGroup = Excel.ChartAxisGroup.secondary; }, context);
      await step(log, 'bar overlap and gap width', async () => { s(0).overlap = 100; s(0).gapWidth = 60; }, context);
      await step(log, 'data labels on actual bars', async () => { s(0).hasDataLabels = true; }, context);
      const ok = log.filter(l => l.indexOf(': yes') > 0).length;
      return { status: ok === log.length ? 'pass' : 'warn', detail: `Z chart built with Office.js: ${log.join('; ')}. Hatched forecast bars need the package writer.` };
    });
  }

  async function pChartIbcs() {
    return Excel.run(async context => {
      const ws = await freshSheet(context, 'IBCS');
      const rows = [['Line', 'PY', 'PL', 'AC', 'ΔPL'], ['Revenue', 920, 1000, 1085, 85], ['Materials', 330, 350, 380, -30],
        ['Labour', 190, 200, 215, -15], ['Overheads', 140, 150, 140, 10], ['EBITDA', 260, 300, 350, 50]];
      ws.getRange('A1:E6').values = rows;
      await context.sync();
      const log = [];
      const col = ws.charts.add(Excel.ChartType.columnClustered, ws.getRange('A1:D6'), Excel.ChartSeriesBy.columns);
      col.setPosition('G2', 'N18');
      col.title.text = 'IBCS: PY, PL, AC';
      await context.sync();
      const s = i => col.series.getItemAt(i);
      await step(log, 'PY grey', async () => { s(0).format.fill.setSolidColor('#A6A6A6'); }, context);
      await step(log, 'PL outlined', async () => { s(1).format.fill.clear(); s(1).format.border.color = '#262626'; }, context);
      await step(log, 'AC solid dark', async () => { s(2).format.fill.setSolidColor('#262626'); }, context);
      await step(log, 'overlap', async () => { s(2).overlap = 40; s(2).gapWidth = 80; }, context);
      ws.getRange('A9:B14').values = rows.map(r => [r[0], r[4]]);
      await context.sync();
      const vb = ws.charts.add(Excel.ChartType.barClustered, ws.getRange('A9:B14'), Excel.ChartSeriesBy.columns);
      vb.setPosition('G20', 'N36');
      vb.title.text = 'IBCS: ΔPL';
      await context.sync();
      const v = vb.series.getItemAt(0);
      const signs = rows.slice(1).map(r => (r[0] === 'Materials' || r[0] === 'Labour' || r[0] === 'Overheads' ? -1 : 1) * Math.sign(r[4]));
      await step(log, 'variance bars green or red by point (cost lines inverted)', async () => {
        signs.forEach((sg, i) => v.points.getItemAt(i).format.fill.setSolidColor(sg >= 0 ? '#8CB400' : '#FF0000'));
      }, context);
      await step(log, 'variance data labels', async () => { v.hasDataLabels = true; }, context);
      await step(log, 'value axis hidden, as IBCS charts label the bars instead', async () => { vb.axes.valueAxis.visible = false; }, context);
      await step(log, 'gridlines off', async () => { vb.axes.valueAxis.majorGridlines.visible = false; }, context);
      const ok = log.filter(l => l.indexOf(': yes') > 0).length;
      return { status: ok === log.length ? 'pass' : 'warn', detail: `IBCS charts built with Office.js: ${log.join('; ')}.` };
    });
  }

  // ---------- results ----------

  function resultRows() {
    const env = state.env || {};
    return Array.from(state.results.values()).map(r => [
      r.id, r.area, r.title, r.status, r.ms == null ? '' : Math.round(r.ms), r.detail, env.platform || '', env.version || '', r.at
    ]);
  }

  async function writeResultsSheet() {
    const header = ['Probe', 'Area', 'Test', 'Status', 'Time (ms)', 'Detail', 'Platform', 'Excel version', 'Run at'];
    const rows = resultRows();
    return Excel.run(async context => {
      const ws = await freshSheet(context, 'Results');
      const all = [header].concat(rows);
      const range = ws.getRangeByIndexes(0, 0, all.length, header.length);
      range.values = all;
      range.format.font.name = 'Segoe UI';
      range.format.font.size = 9;
      range.format.font.color = '#404040';
      range.format.rowHeight = 15;
      const head = ws.getRangeByIndexes(0, 0, 1, header.length);
      head.format.font.bold = true;
      head.format.font.size = 10;
      ws.getRange('A:E').format.autofitColumns();
      ws.getRange('F:F').format.columnWidth = 600;
      ws.getRange('G:I').format.autofitColumns();
      ws.freezePanes.freezeRows(1);
      ws.activate();
      await context.sync();
      return ws.name;
    });
  }

  function resultsJson() {
    return JSON.stringify({ env: state.env, results: Array.from(state.results.values()) }, null, 2);
  }

  async function cleanUp() {
    return Excel.run(async context => {
      const wb = context.workbook;
      wb.worksheets.load('items/name');
      wb.names.load('items/name');
      await context.sync();
      let sheets = 0;
      let names = 0;
      wb.names.items.filter(n => n.name.indexOf(PREFIX) === 0).forEach(n => { n.delete(); names++; });
      await context.sync();
      const drop = wb.worksheets.items.filter(w => w.name.indexOf(PREFIX) === 0 || state.inserted.indexOf(w.name) >= 0);
      if (drop.length === wb.worksheets.items.length) wb.worksheets.add('Sheet1');
      drop.forEach(w => { w.delete(); sheets++; });
      await context.sync();
      if (supports('ExcelApi', '1.14')) {
        const style = wb.styles.getItemOrNullObject(STYLE);
        await context.sync();
        if (!style.isNullObject) { style.delete(); await context.sync(); }
      }
      const parts = wb.customXmlParts.getByNamespace(NS);
      parts.load('items/id');
      await context.sync();
      parts.items.forEach(p => p.delete());
      await context.sync();
      state.inserted = [];
      try { Office.context.document.settings.remove('hfgProbeMarker'); await saveSettings(); } catch (e) { /* ignore */ }
      return `Removed ${sheets} sheets, ${names} names, the probe style, the probe metadata and its settings marker.`;
    });
  }

  // ---------- registry ----------

  const PROBES = [
    { id: 'env', area: 'Platform', title: 'Excel version and requirement sets', auto: true, run: pEnvironment },
    { id: 'write', area: 'Live writer', title: 'Write speed for formulas and formats', auto: true, needs: ['ExcelApi', '1.8'], run: pWriteSpeed },
    { id: 'insert-rows', area: 'Live writer', title: 'Row inserts inside a block', auto: true, needs: ['ExcelApi', '1.4'], run: pRowInserts },
    { id: 'names', area: 'Live writer', title: 'Defined names in bulk', auto: true, needs: ['ExcelApi', '1.7'], run: pNames },
    { id: 'styles', area: 'Live writer', title: 'Named styles', auto: true, needs: ['ExcelApi', '1.7'], run: pStyles },
    { id: 'controls', area: 'Controls', title: 'Validation list, in-cell checkbox, conditional format, note', auto: true, needs: ['ExcelApi', '1.8'], run: pControls },
    { id: 'outline', area: 'Views', title: 'Outline, freeze panes, tab colour, gridlines', auto: true, needs: ['ExcelApi', '1.10'], run: pOutline },
    { id: 'charts', area: 'Reports', title: 'Charts, shapes and page layout', auto: true, needs: ['ExcelApi', '1.9'], run: pChartsShapes },
    { id: 'chart-z', area: 'Charts', title: 'Z chart built with Office.js', auto: true, needs: ['ExcelApi', '1.9'], run: pChartZ },
    { id: 'chart-ibcs', area: 'Charts', title: 'IBCS column and variance charts built with Office.js', auto: true, needs: ['ExcelApi', '1.9'], run: pChartIbcs },
    { id: 'protect', area: 'Finishing', title: 'Sheet protection with a password', auto: true, needs: ['ExcelApi', '1.7'], run: pProtection },
    { id: 'trace', area: 'Review', title: 'Precedents and dependents', auto: true, needs: ['ExcelApi', '1.12'], run: pTrace },
    { id: 'errors', area: 'Review', title: 'Error scan across every sheet', auto: true, needs: ['ExcelApi', '1.9'], run: pErrorScan },
    { id: 'scan', area: 'Review', title: 'Formula read speed for the consistency scan', auto: true, needs: ['ExcelApi', '1.9'], run: pFormulaRead },
    { id: 'events', area: 'Live writer', title: 'Change events, including row inserts', auto: true, needs: ['ExcelApi', '1.7'], run: pEvents },
    { id: 'metadata', area: 'Metadata', title: 'Custom XML part survives save and reopen', auto: true, needs: ['ExcelApi', '1.5'], run: pCustomXml },
    { id: 'settings', area: 'Metadata', title: 'Document settings size', auto: true, run: pSettings },
    { id: 'pane', area: 'Platform', title: 'Task pane width', auto: true, run: pPaneWidth },
    { id: 'calc-mode', area: 'Simulation', title: 'Set calculation to automatic except tables', auto: true, needs: ['ExcelApi', '1.8'], run: pCalcMode },
    { id: 'file-insert', area: 'Package', title: 'Insert sheets from a file', auto: false, needs: ['ExcelApi', '1.13'] },
    { id: 'sheet-copy', area: 'Package', title: 'Copy a sheet with its shapes', auto: false, needs: ['ExcelApi', '1.9'] },
    { id: 'file-read', area: 'Package', title: 'Read the open workbook as a file', auto: false, needs: ['CompressedFile', '1.1'] },
    { id: 'file-open', area: 'Package', title: 'Open a copy with createWorkbook', auto: false, needs: ['ExcelApi', '1.8'] },
    { id: 'pdf', area: 'Reports', title: 'Export the workbook to PDF', auto: false, needs: ['PdfFile', '1.1'] },
    { id: 'sim-run', area: 'Simulation', title: 'Monte Carlo data table runs and matches the reference', auto: false, needs: ['ExcelApi', '1.4'] },
    { id: 'sim-insert', area: 'Simulation', title: 'Data table survives sheet insertion from a file', auto: false, needs: ['ExcelApi', '1.13'] },
    { id: 'undo', area: 'Live writer', title: 'One-step undo for a command', auto: false, needs: ['ExcelApi', '1.20'] },
    { id: 'ribbon', area: 'Commands', title: 'Ribbon button enable and disable', auto: false, needs: ['RibbonApi', '1.1'] },
    { id: 'context', area: 'Commands', title: 'Right-click item enable and disable', auto: false, needs: ['ContextMenuApi', '1.1'] },
    { id: 'dialog', area: 'Commands', title: 'Dialog window', auto: false, needs: ['DialogApi', '1.1'] },
    { id: 'naa', area: 'Sign-in', title: 'Entra sign-in through nested app authentication', auto: false, needs: ['NestedAppAuth', '1.1'] },
    { id: 'cmd-ribbon', area: 'Commands', title: 'Ribbon button runs code without the pane', auto: false },
    { id: 'cmd-menu', area: 'Commands', title: 'Ribbon menu item runs code', auto: false },
    { id: 'cmd-context', area: 'Commands', title: 'Right-click item runs code', auto: false },
    { id: 'cmd-shortcut', area: 'Commands', title: 'Keyboard shortcut runs code', auto: false, needs: ['SharedRuntime', '1.1'] },
    { id: 'cmd-toggle', area: 'Commands', title: 'Toggle test button clicked', auto: false },
    { id: 'ribbon-full', area: 'Ribbon', title: 'Designed ribbon and right-click commands open their views', auto: false },
    { id: 'build-tab', area: 'Ribbon', title: 'Build tab created at runtime as a contextual tab', auto: false, needs: ['RibbonApi', '1.2'] },
    { id: 'ribbon-hide', area: 'Ribbon', title: 'Hide and show a ribbon group', auto: false, needs: ['RibbonApi', '1.3'] }
  ];

  async function runProbe(id, fn) {
    const probe = PROBES.find(p => p.id === id);
    const run = fn || (probe && probe.run);
    if (probe && probe.needs && !supports(probe.needs[0], probe.needs[1])) {
      record(id, { status: 'skip', detail: `Needs ${probe.needs[0]} ${probe.needs[1]}, which this Excel does not have.` });
      return state.results.get(id);
    }
    record(id, { status: 'running', detail: 'Running...' });
    const t0 = now();
    let outcome;
    try { outcome = await run(); } catch (e) { outcome = failure(e); }
    outcome.ms = now() - t0;
    record(id, outcome);
    return state.results.get(id);
  }

  async function runAutomatic(onEach) {
    for (const p of PROBES.filter(x => x.auto)) {
      await runProbe(p.id);
      if (onEach) onEach(p.id);
    }
  }

  function confirm(id, ok, note) {
    const before = state.results.get(id);
    record(id, {
      status: ok ? 'pass' : 'fail',
      ms: before ? before.ms : null,
      detail: `${before ? before.detail + ' ' : ''}User confirmed: ${note}.`
    });
  }

  const api = {
    VERSION, NS, PREFIX, SESSION, PROBES, state, supports,
    runProbe, runAutomatic, confirm, record, registerCommands,
    readChosenFile, insertFromFile, copyInsertedSheet, simulationRun, toggleBuildTab, hideGroup, buildTabDefinition, showView, openView, exportPdf, exportCompressed, openCopy,
    undoSetup, ribbonDisable, contextMenuDisable, openDialog, signIn,
    writeResultsSheet, resultsJson, cleanUp
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.HfgProbe = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
