/*
  Explorer: the model in one pane. A tree of the model (sections, sheets and the
  modules on them, coloured by area) above four tabs for the module selected: Composition (its
  rows on each sheet), Links (what it takes from and sends to, named, as a diagram you can click
  through), Properties (its settings and named inputs) and Checks. Selecting a module selects it
  in the workbook when the model is open in Excel (its HL_Toc_ name).

  In the probe the view runs on the assembly proof's demo model (src/model-sample.js), with its
  own calculation (src/model-calc.js) standing in for Excel for the check status. The model
  functions are DOM free so the Node tests can run them.
*/
(function (root) {
  'use strict';

  const LINK_WORDS = {
    'is.revenue': 'revenue', 'is.opex': 'operating costs', 'is.interest': 'interest', 'cf.receipts': 'receipts',
    'cf.financing': 'financing cash flows', 'bs.debtors': 'closing debtors', 'bs.debt': 'debt balance',
    'check.error': 'error checks', 'check.alert': 'alerts'
  };
  // HF chart palette (steel blue first), then two neutrals: areas are told apart by name as well as colour.
  const AREA_COLOURS = ['#679db5', '#09122c', '#90b6c8', '#566a89', '#3c4e60', '#cedce5', '#a6a6a6', '#404040'];

  const instanceOf = block => block.split('/')[0];
  const moduleOf = (M, uid) => M.modules.find(m => m.uid === uid) || null;

  // Sections, then sheets, then the modules on each sheet in row order.
  function tree(M) {
    return M.sections.map(sec => ({
      title: sec.title,
      sheets: sec.sheets.map(sheet => ({
        sheet,
        modules: M.modules.filter(m => m.components.some(c => c.sheet === sheet))
          .sort((a, b) => first(a, sheet) - first(b, sheet)).map(m => m.uid)
      }))
    }));
    function first(m, sheet) { return Math.min(...m.components.find(c => c.sheet === sheet).rows.map(r => r.row)); }
  }

  // What a module takes from and sends to, by module, with the links between them.
  function relations(M, uid) {
    const takes = new Map(), sends = new Map();
    M.links.forEach(l => {
      const a = instanceOf(l.from), b = instanceOf(l.to);
      if (a === b) return;
      if (b === uid) (takes.get(a) || takes.set(a, new Set()).get(a)).add(l.link);
      if (a === uid) (sends.get(b) || sends.set(b, new Set()).get(b)).add(l.link);
    });
    const list = mp => Array.from(mp, ([u, links]) => ({ uid: u, title: moduleOf(M, u).title, links: Array.from(links) }));
    return { takes: list(takes), sends: list(sends) };
  }

  function checkStatus(M, uid, result) {
    const m = moduleOf(M, uid);
    return m.checks.map(c => {
      const flags = (result && result.checks[c.id]) || [];
      const raised = flags.filter(Boolean).length;
      return { label: c.label, kind: c.kind, raised, periods: flags.length };
    });
  }

  // ---------------------------------------------------------------- view
  const ui = { uid: null, tab: 'links' };
  const K = () => root.HfgViewKit;

  function render(key, host, onShow) {
    if (!['model-explorer', 'b-explorer', 'mod-links'].includes(key) || !root.HfgModelSample) return false;
    if (key === 'mod-links') ui.tab = 'links';
    draw(host, key === 'b-explorer');
    return true;
  }

  function draw(host, builder) {
    const { el } = K();
    const M = root.HfgModelSample;
    const result = root.HfgModelCalc ? root.HfgModelCalc.calc(M) : null;
    if (!ui.uid) ui.uid = M.modules.find(m => m.kind === 'category').uid;
    const redraw = () => draw(host, builder);
    const pick = uid => { ui.uid = uid; redraw(); K().selectName('HL_Toc_' + K().code(uid)); };
    host.innerHTML = '';
    const areaColour = a => AREA_COLOURS[M.areas.indexOf(a) % AREA_COLOURS.length];
    const status = uid => {
      const cs = checkStatus(M, uid, result);
      if (!cs.length) return null;
      const bad = cs.some(c => c.raised && c.kind === 'error'), alert = cs.some(c => c.raised);
      return el('span', { class: 'xstat ' + (bad ? 'bad' : alert ? 'alert' : 'ok'), title: bad ? 'An error check is raised' : alert ? 'An alert is raised' : 'Checks clear' },
        bad || alert ? '!' : '✓');
    };
    host.appendChild(el('p', { class: 'quiet' }, `${M.name}: ${M.modules.length} modules on ${M.sections.reduce((n, s) => n + s.sheets.length, 0)} sheets in ${M.sections.length} sections. Select a module; it is selected in the workbook too.`));
    const t = el('div', { class: 'xtree', role: 'tree' });
    tree(M).forEach((sec, i) => {
      t.appendChild(el('div', { class: 'xsec' }, `${i + 1}. ${sec.title}`));
      sec.sheets.forEach((sh, j) => {
        t.appendChild(el('div', { class: 'xsheet' }, `${String.fromCharCode(97 + j)}. ${sh.sheet}`));
        sh.modules.forEach(uid => {
          const m = moduleOf(M, uid);
          t.appendChild(el('button', { class: 'xmod' + (uid === ui.uid ? ' sel' : ''), role: 'treeitem', 'aria-selected': String(uid === ui.uid), onclick: () => pick(uid) },
            el('span', { class: 'xdot', style: `background:${areaColour(m.area)}` }), el('span', { class: 'xname' }, m.title),
            m.kind === 'mirror' ? el('span', { class: 'gchip' }, `mirror ×${m.blocks.filter(b => b.id !== m.uid).length}`) : null, status(uid)));
          if (m.kind === 'mirror' && uid === ui.uid) {
            m.blocks.filter(b => b.id !== uid).forEach(b => t.appendChild(el('div', { class: 'xblock' }, b.title.split(': ').pop())));
          }
        });
      });
    });
    host.appendChild(t);
    // The selected module
    const m = moduleOf(M, ui.uid);
    const rowsAll = m.components.flatMap(c => c.rows.map(r => Object.assign({ sheet: c.sheet }, r)));
    host.appendChild(el('div', { class: 'xhead' },
      el('span', { class: 'xdot big', style: `background:${areaColour(m.area)}` }),
      el('div', {}, el('h3', {}, m.title), el('p', { class: 'quiet' }, `${m.area} area, ${m.kind} module; ` +
        m.components.map(c => `${c.sheet} rows ${Math.min(...c.rows.map(r => r.row))} to ${Math.max(...c.rows.map(r => r.row))}`).join('; ')))));
    const tabs = [['composition', 'Composition'], ['links', 'Links'], ['properties', 'Properties'], ['checks', 'Checks']];
    host.appendChild(el('div', { class: 'gseg', role: 'tablist' }, tabs.map(([k, label]) =>
      el('button', { class: ui.tab === k ? 'on' : null, role: 'tab', 'aria-selected': String(ui.tab === k), onclick: () => { ui.tab = k; redraw(); } }, label))));
    const body = el('div', { class: 'xbody' });
    host.appendChild(body);
    if (ui.tab === 'composition') {
      m.components.forEach(c => {
        body.appendChild(el('h4', {}, `${c.sheet}`));
        body.appendChild(el('table', { class: 'xrows' }, el('tbody', {}, c.rows.filter(r => r.kind !== 'section' || builder).map(r =>
          el('tr', { class: r.kind }, el('td', { class: 'n' }, String(r.row)), el('td', {}, r.label),
            el('td', { class: 'quiet' }, r.name || (r.style === 'check' ? 'check' : r.kind === 'setting' ? 'input' : r.style === 'total' ? 'total' : '')))))));
      });
      if (builder) body.appendChild(el('p', { class: 'quiet' }, `${rowsAll.length} rows written by the engine from the module's definition; each row's id is kept in the model's metadata.`));
    } else if (ui.tab === 'links') {
      body.appendChild(linksDiagram(M, m, pick));
      body.appendChild(el('p', { class: 'quiet' }, 'Click a module to move to it. Links are named by what they carry; each one is a record the engine wrote when it linked the module in.'));
    } else if (ui.tab === 'properties') {
      const settings = m.settings.map(s => el('tr', {}, el('td', {}, s.label), el('td', { class: 'quiet' }, s.name),
        el('td', { class: 'n' }, String(result ? result.settings[s.name] : s.value))));
      body.appendChild(el('dl', { class: 'xprops' },
        el('dt', {}, 'Library module'), el('dd', {}, m.module),
        el('dt', {}, 'Instance'), el('dd', {}, `${m.uid}${m.kind === 'category' ? ` (number ${m.number}, never reused)` : ''}`),
        el('dt', {}, 'Kind'), el('dd', {}, m.kind === 'mirror' ? 'Mirror: one block for every module that sends what it mirrors' : m.kind === 'category' ? 'Category: can be inserted more than once' : 'Single: once per model'),
        el('dt', {}, 'Area'), el('dd', {}, m.area)));
      if (m.settings.length) body.appendChild(el('table', { class: 'xrows' }, el('thead', {}, el('tr', {}, el('th', {}, 'Setting'), el('th', {}, 'Name'), el('th', { class: 'n' }, 'Value'))), el('tbody', {}, settings)));
      else body.appendChild(el('p', { class: 'quiet' }, 'No settings: everything on this module is worked out from its links.'));
    } else {
      const cs = checkStatus(M, m.uid, result);
      if (!cs.length) body.appendChild(el('p', { class: 'quiet' }, 'No checks of its own; what it sends is checked where it lands (the Financial statements and Checks).'));
      cs.forEach(c => body.appendChild(el('p', {}, el('span', { class: 'xstat ' + (c.raised ? (c.kind === 'error' ? 'bad' : 'alert') : 'ok') }, c.raised ? '!' : '✓'), ' ',
        `${c.label} (${c.kind}): ` + (c.raised ? `raised in ${c.raised} of ${c.periods} months` : 'clear'))));
    }
  }

  function linksDiagram(M, m, pick) {
    const { el } = K();
    const rel = relations(M, m.uid);
    const W = 300, H0 = 34, GAP = 8;
    let y = 4;
    const kids = [];
    const box = (r, kind) => {
      const words = r.links.map(l => LINK_WORDS[l] || l).join(', ');
      const g = el('g', { class: 'xbox ' + kind, tabindex: '0', role: 'button', 'aria-label': `${r.title}: ${words}`, onclick: () => pick(r.uid),
        onkeydown: ev => { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); pick(r.uid); } } },
      el('title', {}, r.links.join(', ')),
      el('rect', { x: 6, y, width: W - 12, height: H0, rx: 5 }),
      el('text', { x: 16, y: y + 14, class: 'xt' }, r.title),
      el('text', { x: 16, y: y + 27, class: 'xw' }, words));
      kids.push(g);
      y += H0 + GAP;
    };
    const label = text => { kids.push(el('text', { x: 6, y: y + 10, class: 'xlab' }, text)); y += 16; };
    const arrow = text => {
      kids.push(el('path', { d: `M${W / 2} ${y} V${y + 22}`, class: 'xarrow', 'marker-end': 'url(#xhead)' }));
      if (text) kids.push(el('text', { x: W / 2 + 8, y: y + 15, class: 'xw' }, text));
      y += 28;
    };
    label(rel.takes.length ? 'Takes from' : 'Takes nothing from other modules');
    rel.takes.forEach(r => box(r, 'pre'));
    if (rel.takes.length) arrow('');
    kids.push(el('rect', { x: 40, y, width: W - 80, height: H0, rx: 6, class: 'xself' }));
    kids.push(el('text', { x: W / 2, y: y + 21, class: 'xt mid' }, m.title));
    y += H0 + 4;
    if (rel.sends.length) arrow('');
    label(rel.sends.length ? 'Sends to' : 'Sends nothing to other modules');
    rel.sends.forEach(r => box(r, 'dep'));
    return el('svg', { viewBox: `0 0 ${W} ${y}`, width: '100%', role: 'img', class: 'xlinks', 'aria-label': `Links of ${m.title}` },
      el('defs', {}, el('marker', { id: 'xhead', viewBox: '0 0 10 10', refX: '9', refY: '5', markerWidth: '6', markerHeight: '6', orient: 'auto' },
        el('path', { d: 'M0 0L10 5L0 10z', class: 'xarrowhead' }))), kids);
  }

  const api = { tree, relations, checkStatus, render, LINK_WORDS, reset: () => { ui.uid = null; ui.tab = 'links'; } };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.HfgExplorer = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
