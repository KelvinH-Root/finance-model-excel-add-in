/*
  Impacts: what a change or a transaction does to the income statement, balance sheet and cash
  flow (from Analysis > Impacts or right-click > Show impacts).

  Impact of a change (live): pick an input, a new value and a month. The add-in will change the
  input, recalculate, read the statements and put the input back in one step (proven in
  prototypes/impacts/impact_live.py); here the sample model's own calculation (model-calc.js)
  stands in for Excel. Every line that moved, the ties (surplus to equity, net cash flow to
  cash, the balance sheet still balancing) and the chain of links that carried it.

  Impacts sheets: what Add to model writes, one sheet per kind of transaction the model holds,
  in its own accounts and entities (prototypes/impacts/impact_sheets.py); intergroup items show
  the eliminations and the group. Nothing is written to the workbook in the probe.
*/
(function (root) {
  'use strict';
  const STMT = 'demo.statements#1';
  const SECTIONS = ['Income statement', 'Cash flow', 'Balance sheet'];

  // ---------------------------------------------------------------- live mode
  function liveImpact(M, name, value, month) {
    const C = root.HfgModelCalc;
    const base = C.calc(M), changed = C.calc(M, { [name]: Number(value) });
    const n = M.periods, m = Math.min(Math.max(1, month), n) - 1;
    const rows = M.lines.map(l => {
      const d = changed.lines[l.id].map((v, t) => v - base.lines[l.id][t]);
      const balance = l.section === 'Balance sheet' || l.id === `${STMT}/cash`;
      return { id: l.id, label: l.label, section: l.section, total: l.total, delta: d, month: d[m],
               year: balance ? d[n - 1] : d.reduce((a, b) => a + b, 0), moved: d.some(x => Math.abs(x) > 1e-9) };
    });
    const D = id => rows.find(r => r.id === `${STMT}/${id}`).delta;
    const toDate = (arr, t) => arr.slice(0, t + 1).reduce((a, b) => a + b, 0);
    const ties = {
      surplus: { carried: toDate(D('surplus'), m), equity: D('equity')[m] },
      cash: { flows: toDate(D('net_cash'), m), cash: D('cash')[m] },
      balance: D('assets')[m] - D('debt')[m] - D('equity')[m]
    };
    return { rows, ties, month: m + 1, was: root.HfgModelCalc.settingsOf(M)[name], value: Number(value) };
  }

  function chain(M, uid) {
    const reached = new Set(Object.keys(M.blocks).filter(b => b === uid || b.startsWith(uid + '/')));
    const edges = [];
    let frontier = new Set(reached);
    while (frontier.size) {
      const next = new Set();
      M.links.forEach(l => {
        if (l.link.startsWith('check.') || !frontier.has(l.from)) return;
        const e = [M.blocks[l.from], l.link, M.blocks[l.to]];
        if (!edges.some(x => x.join() === e.join())) edges.push(e);
        if (!reached.has(l.to)) { reached.add(l.to); next.add(l.to); }
      });
      frontier = next;
    }
    return edges;
  }

  // ---------------------------------------------------------------- sheets mode
  function evaluate(expr, values) {
    const names = Object.keys(values);
    return Number(new Function(...names, `return (${expr});`)(...names.map(k => Number(values[k]))));   // expressions come from the sample
  }

  function panels(item, ctx, values, cashClasses) {
    const v = {};
    item.inputs.forEach(i => { v[i.name] = i.value; });
    item.switches.forEach(s => { v[s.name] = s.value; });
    Object.assign(v, values || {});
    const legs = item.legs.map(l => ({ entity: l.entity, role: l.role, amount: evaluate(l.expr, v), cash: l.cash }));
    const out = {};
    item.columns.filter(c => c !== 'Group').forEach(col => {
      const amt = r => legs.filter(l => l.entity === col && l.role === r).reduce((a, l) => a + l.amount, 0);
      const p = {};
      item.lines.revenue.forEach(r => { p[r] = -amt(r); });
      item.lines.expense.forEach(r => { p[r] = amt(r); });
      p.surplus = item.lines.revenue.reduce((a, r) => a + p[r], 0) - item.lines.expense.reduce((a, r) => a + p[r], 0);
      item.lines.asset.forEach(r => { p[r] = amt(r); });
      p.assets = item.lines.asset.reduce((a, r) => a + p[r], 0);
      item.lines.liability.forEach(r => { p[r] = -amt(r); });
      p.liabilities = item.lines.liability.reduce((a, r) => a + p[r], 0);
      item.lines.equity.forEach(r => { p[r] = -amt(r); });
      p.retained = p.surplus;
      p.equity_total = item.lines.equity.reduce((a, r) => a + p[r], 0) + p.retained;
      p.balance = p.assets - p.liabilities - p.equity_total;
      cashClasses.forEach(cc => { p[cc] = legs.filter(l => l.entity === col && l.role === 'cash' && l.cash === cc).reduce((a, l) => a + l.amount, 0); });
      p.net_cash = cashClasses.reduce((a, cc) => a + p[cc], 0);
      p.cash_tie = p.net_cash - (p.cash || 0);
      out[col] = p;
    });
    if (item.group) {
      const keys = Object.keys(Object.values(out)[0]);
      out.Group = {};
      keys.forEach(k => { out.Group[k] = Object.values(out).reduce((a, p) => a + (p[k] || 0), 0); });
    }
    return { legs, panels: out };
  }

  // ---------------------------------------------------------------- views
  const ui = { tab: 'live', name: null, value: null, month: 3, ctx: 1, item: 'claim', values: {} };
  const K = () => root.HfgViewKit;

  function render(key, host, onShow) {
    if (!['an-impacts', 'imp-live', 'imp-sheets', 'ctx-impacts'].includes(key) || !root.HfgModelSample || !root.HfgImpactsSample) return false;
    if (key === 'imp-sheets') ui.tab = 'sheets';
    if (key === 'imp-live' || key === 'ctx-impacts') ui.tab = 'live';
    draw(host);
    return true;
  }

  function draw(host) {
    const { el } = K();
    host.innerHTML = '';
    const redraw = () => draw(host);
    host.appendChild(el('div', { class: 'gseg', role: 'tablist' },
      [['live', 'Impact of a change'], ['sheets', 'Impacts sheets']].map(([k, label]) =>
        el('button', { class: ui.tab === k ? 'on' : null, role: 'tab', 'aria-selected': String(ui.tab === k), onclick: () => { ui.tab = k; redraw(); } }, label))));
    if (ui.tab === 'live') liveView(host, redraw); else sheetsView(host, redraw);
  }

  function liveView(host, redraw) {
    const { el, money } = K();
    const M = root.HfgModelSample;
    const settings = M.modules.filter(m => !m.module.includes('dashboard')).flatMap(m => m.settings.map(s => Object.assign({ module: m.title, uid: m.uid }, s)));
    if (!ui.name) { ui.name = 'Rev1_Base'; ui.value = 110; }
    const s = settings.find(x => x.name === ui.name);
    const res = liveImpact(M, ui.name, ui.value, ui.month);
    host.appendChild(el('div', { class: 'gform' },
      el('div', { class: 'gfield' }, el('label', { for: 'i-name' }, 'Input'), el('select', { id: 'i-name', onchange: ev => {
        ui.name = ev.target.value; const x = settings.find(y => y.name === ui.name); ui.value = Math.round(x.value * 1.1 * 10000) / 10000; redraw(); } },
      settings.map(x => el('option', { value: x.name, selected: x.name === ui.name }, `${x.module}: ${x.label} (${x.name})`)))),
      el('div', { class: 'grow' },
        el('div', { class: 'gfield' }, el('label', { for: 'i-value' }, `New value (now ${res.was})`),
          el('input', { id: 'i-value', type: 'number', step: 'any', value: String(ui.value), onchange: ev => { ui.value = Number(ev.target.value); redraw(); } })),
        el('div', { class: 'gfield' }, el('label', { for: 'i-month' }, 'Month shown'), el('select', { id: 'i-month', onchange: ev => { ui.month = Number(ev.target.value); redraw(); } },
          Array.from({ length: M.periods }, (_, t) => el('option', { value: String(t + 1), selected: t + 1 === ui.month }, `Month ${t + 1}`)))))));
    const table = el('table', { class: 'gbridge' }, el('thead', {}, el('tr', {}, el('th', {}, 'Change in'), el('th', { class: 'n' }, `Month ${res.month}`), el('th', { class: 'n' }, 'Year'))));
    const tbody = el('tbody', {});
    SECTIONS.forEach(sec => {
      const rows = res.rows.filter(r => r.section === sec && (r.moved || r.total));
      if (!rows.some(r => r.moved)) { tbody.appendChild(el('tr', { class: 'sec' }, el('td', { colspan: '3' }, `${sec}: no change`))); return; }
      tbody.appendChild(el('tr', { class: 'sec' }, el('td', { colspan: '3' }, sec)));
      rows.forEach(r => tbody.appendChild(el('tr', { class: r.total ? 'tot' : null }, el('td', {}, r.label),
        el('td', { class: 'n' }, money(r.month)), el('td', { class: 'n' }, money(r.year)))));
    });
    table.appendChild(tbody);
    host.appendChild(table);
    const ok = x => Math.abs(x) < 1e-6;
    const tick = b => el('span', { class: 'xstat ' + (b ? 'ok' : 'bad') }, b ? '✓' : '!');
    host.appendChild(el('div', { class: 'gcard' },
      el('p', {}, tick(ok(res.ties.surplus.carried - res.ties.surplus.equity)), ` Surplus to month ${res.month} (${money(res.ties.surplus.carried)}) is carried to the accumulated surplus.`),
      el('p', {}, tick(ok(res.ties.cash.flows - res.ties.cash.cash)), ` Net cash flow to month ${res.month} (${money(res.ties.cash.flows)}) equals the change in cash.`),
      el('p', {}, tick(ok(res.ties.balance)), ' The balance sheet still balances.')));
    const edges = chain(M, s.uid);
    host.appendChild(el('h4', {}, 'Why it moved'));
    host.appendChild(el('ol', { class: 'ichain' }, edges.map(([a, link, b]) => el('li', {}, `${a} sends ${root.HfgExplorer ? root.HfgExplorer.LINK_WORDS[link] || link : link} to ${b}`))));
    host.appendChild(el('div', { class: 'gactions' }, el('button', { disabled: true, title: 'Not in the probe' }, 'Keep the change'),
      el('button', { disabled: true, title: 'Not in the probe' }, 'Save as an Impacts sheet')));
    host.appendChild(el('p', { class: 'quiet' }, 'In Excel the add-in writes the new value, recalculates, reads the statements and puts the value back in one step, so the model is never left changed. Here the sample model\'s own calculation stands in for Excel.'));
  }

  function sheetsView(host, redraw) {
    const { el, money } = K();
    const S = root.HfgImpactsSample;
    const ctx = S.contexts[ui.ctx];
    if (!ctx.items.some(i => i.key === ui.item)) ui.item = ctx.items[0].key;
    const item = ctx.items.find(i => i.key === ui.item);
    const vals = ui.values[ui.item] || (ui.values[ui.item] = {});
    host.appendChild(el('div', { class: 'gfield' }, el('label', { for: 'i-ctx' }, 'Accounts and entities from'),
      el('select', { id: 'i-ctx', onchange: ev => { ui.ctx = Number(ev.target.value); redraw(); } },
        S.contexts.map((c, k) => el('option', { value: String(k), selected: k === ui.ctx }, `${c.label}: ${c.items.length} sheets`)))));
    host.appendChild(el('div', { class: 'ichips' }, ctx.items.map(i => el('button', { class: 'gchip' + (i.key === ui.item ? ' first' : ''), onclick: () => { ui.item = i.key; redraw(); } }, i.title))));
    host.appendChild(el('h3', {}, item.title));
    host.appendChild(el('p', { class: 'quiet' }, item.purpose));
    const form = el('div', { class: 'gform' });
    item.inputs.forEach(i => form.appendChild(el('div', { class: 'gfield' }, el('label', { for: 'in-' + i.name }, i.label),
      el('input', { id: 'in-' + i.name, type: 'number', step: 'any', value: String(vals[i.name] ?? i.value), onchange: ev => { vals[i.name] = Number(ev.target.value); redraw(); } }))));
    item.switches.forEach(sw => form.appendChild(el('label', { class: 'gcheck' },
      el('input', { type: 'checkbox', checked: (vals[sw.name] ?? sw.value) ? true : null, onchange: ev => { vals[sw.name] = ev.target.checked ? 1 : 0; redraw(); } }), ' ' + sw.label)));
    host.appendChild(form);
    const { legs, panels: P } = panels(item, ctx, vals, S.cashClasses);
    const label = k => (ctx.accounts[k] ? ctx.accounts[k][0].replace(/^\d{4} /, '') : S.fixed[k].replace(' (must be nil)', ''));
    const cols = item.columns;
    const short = c => (c === 'Eliminations' ? 'Elims' : c.replace(/^\d{4} /, '').replace(/ (Ltd|LP)$/, ''));
    const table = el('table', { class: 'gbridge icols' }, el('thead', {}, el('tr', {}, el('th', {}, ''), cols.map(c => el('th', { class: 'n' }, short(c))))));
    const tb = el('tbody', {});
    const row = (k, cls) => tb.appendChild(el('tr', { class: cls }, el('td', {}, label(k)), cols.map(c => el('td', { class: 'n' }, money(P[c][k] || 0)))));
    const sec = t => tb.appendChild(el('tr', { class: 'sec' }, el('td', { colspan: String(cols.length + 1) }, t)));
    sec('Income statement');
    item.lines.revenue.concat(item.lines.expense).forEach(k => row(k));
    row('surplus', 'tot');
    sec('Balance sheet');
    item.lines.asset.forEach(k => row(k));
    row('assets', 'tot');
    item.lines.liability.forEach(k => row(k));
    row('liabilities', 'tot');
    item.lines.equity.forEach(k => row(k));
    row('retained');
    row('equity_total', 'tot');
    sec('Cash flow');
    S.cashClasses.forEach(cc => row(cc));
    row('net_cash', 'tot');
    table.appendChild(tb);
    host.appendChild(el('div', { class: 'iscroll' }, table));
    const okAll = cols.every(c => Math.abs(P[c].balance) < 1e-6 && Math.abs(P[c].cash_tie) < 1e-6);
    host.appendChild(el('p', {}, el('span', { class: 'xstat ' + (okAll ? 'ok' : 'bad') }, okAll ? '✓' : '!'),
      ' Surplus carried to retained surplus; net cash flow equals the change in cash; every column balances.'));
    item.notes.forEach(n => host.appendChild(el('p', { class: 'quiet' }, n)));
    host.appendChild(el('details', {}, el('summary', {}, `The entries (${legs.length})`),
      el('table', { class: 'xrows' }, el('tbody', {}, legs.map(l => el('tr', {}, el('td', {}, short(l.entity)), el('td', {}, ctx.accounts[l.role][0]),
        el('td', { class: 'n' }, money(l.amount)), el('td', { class: 'quiet' }, l.cash || '')))))));
    host.appendChild(el('div', { class: 'gcard' }, el('h3', {}, 'Add to model'),
      el('p', {}, `Writes an Impacts section into ${ctx.model}: a cover, ${ctx.items.length} sheets (${ctx.items.map(i => i.sheet.replace(/^Imp /, '')).join(', ')}), contents entries and one error check. The sheets are formulas, labelled with this model's accounts and entities, and work without the add-in.`),
      el('button', { disabled: true, title: 'Not in the probe' }, 'Add to model')));
  }

  const api = { liveImpact, chain, panels, render, reset: () => { ui.name = null; } };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.HfgImpacts = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
