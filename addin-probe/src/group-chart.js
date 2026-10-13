/*
  Group structure drawn as a chart, for its dialog (Kelvin, 13 October 2026: it should be visual,
  not just text). Each entity is a box under its parent with the share held, its tags (heads a
  group, outside investors, planned) and a shaded outline round every group below the top, so
  the groups that consolidate can be seen. Selecting a box shows its details beside the chart; the
  Roll-up tab shows a group's figures for the year shown. Same fictional sample and the same state
  as the other group views (group.js), so Add entity and Change ownership change this chart too.
*/
(function (root) {
  'use strict';
  const SVG = 'http://www.w3.org/2000/svg';
  const W = 150, H = 64, GAP = 14, LEVEL = 112, PAD = 24;

  function svg(tag, attrs, ...kids) {
    const n = document.createElementNS(SVG, tag);
    Object.entries(attrs || {}).forEach(([k, v]) => {
      if (v == null || v === false) return;
      if (k.startsWith('on') && typeof v === 'function') n.addEventListener(k.slice(2), v);
      else n.setAttribute(k, String(v));
    });
    kids.flat(Infinity).forEach(k => { if (k != null && k !== false) n.appendChild(typeof k === 'string' ? document.createTextNode(k) : k); });
    return n;
  }
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
  const clip = (s, n) => (s.length > n ? s.slice(0, n - 1) + '…' : s);

  // Tidy layout: leaves take the next slot from the left, a parent sits over the middle of its children.
  function layout(G, st) {
    const pos = new Map();
    let slot = 0, deepest = 0;
    function place(code, depth) {
      deepest = Math.max(deepest, depth);
      const kids = G.childrenOf(st, code);
      let x;
      if (!kids.length) x = slot++;
      else {
        const xs = kids.map(k => place(k.code, depth + 1));
        x = (xs[0] + xs[xs.length - 1]) / 2;
      }
      pos.set(code, { x, depth });
      return x;
    }
    st.entities.filter(e => e.parent == null).forEach(e => place(e.code, 0));
    const px = x => PAD + x * (W + GAP);
    const py = d => PAD + 18 + d * LEVEL;
    const box = new Map();
    pos.forEach((p, code) => box.set(code, { x: px(p.x), y: py(p.depth), w: W, h: H }));
    return { box, width: PAD * 2 + slot * (W + GAP) - GAP, height: py(deepest) + H + PAD };
  }

  function chart(G, st, ui, rerender) {
    const L = layout(G, st);
    const layers = [];
    // groups below the top, outermost first, each round its members' boxes
    const top = st.entities.find(e => e.parent == null).code;
    const groups = st.groups.filter(g => g.head !== top)
      .map(g => ({ g, members: G.membersOf(st, g.head) }))
      .sort((a, b) => b.members.length - a.members.length);
    groups.forEach(({ g, members }, k) => {
      const bs = members.map(c => L.box.get(c));
      const nest = groups.filter(o => o !== groups[k] && o.members.length > members.length && members.every(c => o.members.includes(c))).length;
      const pad = 12 - Math.min(nest, 2) * 3;
      const x1 = Math.min(...bs.map(b => b.x)) - pad, x2 = Math.max(...bs.map(b => b.x + b.w)) + pad;
      const y1 = Math.min(...bs.map(b => b.y)) - pad - 14, y2 = Math.max(...bs.map(b => b.y + b.h)) + pad;
      const nci = G.byCode(st, g.head).held < 1;
      layers.push(svg('g', { class: 'gc-group' + (nci ? ' nci' : '') },
        svg('rect', { x: x1, y: y1, width: x2 - x1, height: y2 - y1, rx: 10 }),
        svg('text', { x: x1 + 8, y: y1 + 12 }, g.name)));
    });
    // connectors
    st.entities.forEach(e => {
      if (e.parent == null) return;
      const p = L.box.get(e.parent), c = L.box.get(e.code);
      const mid = p.y + p.h + (c.y - p.y - p.h) / 2;
      layers.push(svg('path', { class: 'gc-link', d: `M${p.x + p.w / 2} ${p.y + p.h} V${mid} H${c.x + c.w / 2} V${c.y}` }));
      layers.push(svg('text', { class: 'gc-held', x: c.x + c.w / 2 + 5, y: c.y - 5 }, G.pct(e.held)));
    });
    // boxes
    st.entities.forEach(e => {
      const b = L.box.get(e.code);
      const heads = !!G.groupOf(st, e.code);
      const tags = [];
      if (heads) tags.push(G.membersOf(st, e.code).length > 1 ? 'Group' : 'Node');
      if (e.held < 1) tags.push(`NCI ${G.pct(1 - e.held)}`);
      if (e.status === 'Planned') tags.push('Planned');
      const sel = ui.selected === e.code;
      const pick = () => { ui.selected = e.code; rerender(); };
      layers.push(svg('g', { class: 'gc-node' + (sel ? ' sel' : '') + (e.status === 'Planned' ? ' planned' : '') + (ui.flash === e.code ? ' flash' : ''),
        tabindex: 0, role: 'button', 'aria-pressed': String(sel), 'aria-label': `${e.code} ${e.name}${tags.length ? ', ' + tags.join(', ') : ''}`,
        onclick: pick, onkeydown: ev => { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); pick(); } } },
      svg('rect', { x: b.x, y: b.y, width: b.w, height: b.h, rx: 6, class: 'box' }),
      heads ? svg('rect', { x: b.x, y: b.y, width: b.w, height: 4, rx: 2, class: 'bar' }) : null,
      svg('text', { x: b.x + 10, y: b.y + 21, class: 'name' }, clip(e.name, 20)),
      svg('text', { x: b.x + 10, y: b.y + 37, class: 'code' }, String(e.code)),
      svg('text', { x: b.x + 10, y: b.y + 53, class: 'tags' }, tags.join(' · '))));
    });
    return svg('svg', { class: 'gc-chart', width: L.width, height: L.height, viewBox: `0 0 ${L.width} ${L.height}`, role: 'img',
      'aria-label': `Group structure: ${st.entities.length} entities` }, layers);
  }

  function details(G, st, ui, show, rerender) {
    const e = G.byCode(st, ui.selected);
    if (!e) return el('div', {});
    const p = ui.year;
    const parent = e.parent == null ? null : G.byCode(st, e.parent);
    const into = G.groupsOf(st, e.code);
    const row = (k, v) => [el('dt', {}, k), el('dd', {}, v)];
    const g = G.groupOf(st, e.code);
    const f = g ? G.figures(st, e.code, p) : null;
    return el('div', { class: 'gc-detail' },
      el('p', { class: 'k-title' }, `${e.code}  ${e.name}`),
      e.role ? el('p', { class: 'quiet' }, e.role) : null,
      el('dl', { class: 'gc-dl' },
        row('Parent', parent ? `${parent.code} ${parent.name}` : 'None (top of the group)'),
        row('Held by parent', parent ? G.pct(e.held) : '-'),
        row('Owned by the top', G.pct(G.ownedByTop(st, e.code))),
        row('Outside investors', e.held < 1 ? G.pct(1 - e.held) + ' (NCI worked out here)' : '-'),
        row('Status', e.status === 'Planned' ? 'Planned (not in Home Hub yet)' : 'Actual (from Home Hub)'),
        row('Member from', e.from),
        row('Rolls into', el('span', {}, into.map((x, i) => el('span', { class: 'gchip' + (i === 0 ? ' first' : '') }, x.name)))),
        row(`Own surplus, ${st.years[p]}`, G.money(e.own.surplus[p])),
        row(`Own net assets, ${st.years[p]}`, G.money(e.own.na[p])),
        f ? row(`${g.name}, surplus`, G.money(f.surplus)) : null,
        f ? row(`${g.name}, net assets`, G.money(f.na)) : null),
      el('div', { class: 'gc-actions' },
        el('button', { onclick: () => { G.ui.draft = { parent: e.code }; show('grp-add'); } }, 'Add entity under it'),
        e.parent != null ? el('button', { onclick: () => { G.ui.draft = { entity: e.code }; show('grp-ownership'); } }, 'Change ownership') : null,
        g ? el('button', { onclick: () => { ui.chartTab = 'rollup'; ui.group = e.code; rerender(); } }, 'Roll-up') : null));
  }

  function rollupTable(G, st, ui, rerender) {
    const p = ui.year;
    const r = G.rollup(st, ui.group, p);
    const g = G.groupOf(st, ui.group);
    const trail = [];
    for (let x = g; x; x = G.parentGroup(st, x.head)) trail.unshift(x);
    const tr = (cells, cls) => el('tr', { class: cls || null }, cells);
    return el('div', { class: 'gc-rollup' },
      el('div', { style: 'display:flex;gap:10px;align-items:center;flex-wrap:wrap' },
        el('label', { for: 'gc-pick', style: 'margin:0;font-weight:600' }, 'Group'),
        el('select', { id: 'gc-pick', class: 'k-inp', style: 'width:auto', onchange: ev => { ui.group = Number(ev.target.value); rerender(); } },
          st.groups.map(x => el('option', { value: String(x.head), selected: x.head === ui.group }, x.name))),
        el('span', { class: 'quiet' }, trail.map((x, i) => [i ? ' › ' : '', i === trail.length - 1 ? el('strong', {}, x.name)
          : el('a', { href: '#', onclick: ev => { ev.preventDefault(); ui.group = x.head; rerender(); } }, x.name)]))),
      el('table', { class: 'k-table gc-table' },
        tr([el('th', {}, `${st.years[p]}, $000`), el('th', { class: 'n' }, 'Surplus'), el('th', { class: 'n' }, 'Net assets')]),
        r.lines.map(l => tr([
          el('td', {}, l.kind === 'group' ? el('a', { href: '#', onclick: ev => { ev.preventDefault(); ui.group = l.head; rerender(); } }, `${l.label}, consolidated`) : l.label),
          el('td', { class: 'n' }, G.money(l.surplus)), el('td', { class: 'n' }, G.money(l.na))], l.kind === 'elim' ? 'elim' : '')),
        tr([el('td', {}, `${g.name}, consolidated`), el('td', { class: 'n' }, G.money(r.total.surplus)), el('td', { class: 'n' }, G.money(r.total.na))], 'total'),
        tr([el('td', {}, 'Non-controlling interests'), el('td', { class: 'n' }, G.money(r.nci.surplus)), el('td', { class: 'n' }, G.money(r.nci.na))], 'sub'),
        tr([el('td', {}, 'Owners'), el('td', { class: 'n' }, G.money(r.owners.surplus)), el('td', { class: 'n' }, G.money(r.owners.na))], 'sub')),
      el('p', { class: 'quiet' }, 'Sub-groups are already consolidated, so this group adds only the eliminations made in it: pairs whose lowest common group is this one, its investments, its share of the sub-groups\' NCI and margin still held in it.'));
  }

  // Draws the dialog's body into host; show(key) moves to another group view in the same dialog.
  function render(host, show) {
    const G = root.HfgGroup;
    if (!G || !root.HfgGroupSample) return false;
    const st = G.state();
    const ui = G.ui;
    if (ui.selected == null) ui.selected = st.entities[0].code;
    if (ui.year == null) ui.year = st.years.length - 1;
    if (ui.group == null) ui.group = st.groups[0].head;
    if (!ui.chartTab) ui.chartTab = 'chart';
    const rerender = () => render(host, show);
    host.innerHTML = '';
    const planned = st.entities.filter(e => e.status === 'Planned').length;
    const bar = el('div', { class: 'gc-bar' },
      el('div', { class: 'gseg', role: 'tablist' },
        el('button', { class: ui.chartTab === 'chart' ? 'on' : null, role: 'tab', onclick: () => { ui.chartTab = 'chart'; rerender(); } }, 'Structure'),
        el('button', { class: ui.chartTab === 'rollup' ? 'on' : null, role: 'tab', onclick: () => { ui.chartTab = 'rollup'; rerender(); } }, 'Roll-up')),
      el('div', { class: 'gseg', role: 'group', 'aria-label': 'Year shown' },
        st.years.map((y, i) => el('button', { class: i === ui.year ? 'on' : null, 'aria-pressed': String(i === ui.year), onclick: () => { ui.year = i; rerender(); } }, y))),
      el('span', { class: 'quiet' }, `${st.entities.length} entities, ${st.groups.length} groups that consolidate${planned ? `, ${planned} planned` : ''}. Fictional sample, $000.`));
    const body = ui.chartTab === 'chart'
      ? el('div', { class: 'gc-cols' }, el('div', { class: 'gc-scroll' }, chart(G, st, ui, rerender)), details(G, st, ui, show, rerender))
      : el('div', { class: 'gc-scroll pad' }, rollupTable(G, st, ui, rerender));
    host.appendChild(el('div', { class: 'gc' }, bar, body));
    ui.flash = null;
    return true;
  }

  root.HfgGroupChart = { render, layout };
})(typeof globalThis !== 'undefined' ? globalThis : this);
