/*
  Group views for the probe: the group as a tree with ownership, each group's figures rolled up
  the tree, and Add entity with its share held. Designed views on the fictional sample group
  (src/group-sample.js, generated from prototypes/consolidation); in the probe, Add entity
  changes the sample in the pane only, never the workbook.

  The model functions are plain and DOM free so the Node tests can run them. They follow the
  consolidation proof: Home Hub is the master for actual entities; an entity added here is
  Planned; it goes under its parent after the parent's last descendant, so the register stays in
  tree order; an entity its parent holds less than 100% of has outside investors, so it becomes
  an NCI node and a group of its own.
*/
(function (root) {
  'use strict';

  // ---------------------------------------------------------------- model
  function makeState(sample) {
    const s = JSON.parse(JSON.stringify(sample));
    return {
      years: s.years, yearStart: s.yearStart, note: s.note,
      entities: s.entities.map(e => Object.assign({}, e, { own: e.own || { surplus: s.years.map(() => 0), na: s.years.map(() => 0) } })),
      groups: s.groups.map(g => Object.assign({}, g))
    };
  }

  const byCode = (st, code) => st.entities.find(e => e.code === code) || null;
  const groupOf = (st, head) => st.groups.find(g => g.head === head) || null;

  function ancestors(st, code) {
    const out = [];
    let e = byCode(st, code);
    while (e && e.parent != null) { out.push(e.parent); e = byCode(st, e.parent); }
    return out;
  }
  const depth = (st, code) => ancestors(st, code).length;
  const childrenOf = (st, code) => st.entities.filter(e => e.parent === code);
  function ownedByTop(st, code) {
    let share = 1;
    for (let e = byCode(st, code); e; e = e.parent == null ? null : byCode(st, e.parent)) share *= e.held;
    return share;
  }
  function membersOf(st, head) {
    return st.entities.filter(e => e.code === head || ancestors(st, e.code).includes(head)).map(e => e.code);
  }
  // Groups holding the entity, smallest first: the first is where it is consolidated first.
  function groupsOf(st, code) {
    return st.groups.filter(g => membersOf(st, g.head).includes(code))
      .sort((a, b) => membersOf(st, a.head).length - membersOf(st, b.head).length);
  }
  function parentGroup(st, head) {
    return groupsOf(st, head).find(g => g.head !== head) || null;
  }
  const childGroups = (st, head) => st.groups.filter(g => { const p = parentGroup(st, g.head); return p && p.head === head; });
  function directMembers(st, head) {
    return membersOf(st, head).filter(c => groupsOf(st, c)[0].head === head);
  }

  // Each group's figures for year p: its members' own figures, plus every elimination made in it or below it.
  function figures(st, head, p) {
    const g = groupOf(st, head);
    const own = k => membersOf(st, head).reduce((s, c) => s + byCode(st, c).own[k][p], 0);
    const pick = (k, dflt) => (g && g[k] ? g[k][p] : dflt);
    const elimS = pick('elimSurplus', 0), elimN = pick('elimNa', 0);
    return {
      surplus: own('surplus') + elimS, na: own('na') + elimN, elimSurplus: elimS, elimNa: elimN,
      nciSurplus: pick('nciSurplus', 0), nci: pick('nci', 0)
    };
  }

  // The roll-up of one group for year p: direct members' own figures, each sub-group consolidated, and the
  // eliminations made in this group (its eliminations less those already made in the groups below it).
  function rollup(st, head, p) {
    const lines = [];
    directMembers(st, head).forEach(c => {
      const e = byCode(st, c);
      lines.push({ kind: 'entity', code: c, label: e.name + (c === head ? ' (own figures)' : ''), surplus: e.own.surplus[p], na: e.own.na[p] });
    });
    const kids = childGroups(st, head);
    kids.forEach(g => { const f = figures(st, g.head, p); lines.push({ kind: 'group', head: g.head, label: g.name, surplus: f.surplus, na: f.na }); });
    const f = figures(st, head, p);
    const below = kids.reduce((s, g) => { const k = figures(st, g.head, p); return { s: s.s + k.elimSurplus, n: s.n + k.elimNa }; }, { s: 0, n: 0 });
    lines.push({ kind: 'elim', label: 'Eliminations made in this group', surplus: f.elimSurplus - below.s, na: f.elimNa - below.n });
    const total = lines.reduce((s, l) => ({ surplus: s.surplus + l.surplus, na: s.na + l.na }), { surplus: 0, na: 0 });
    return { head, lines, total, consolidated: { surplus: f.surplus, na: f.na }, nci: { surplus: f.nciSurplus, na: f.nci },
             owners: { surplus: f.surplus - f.nciSurplus, na: f.na - f.nci } };
  }

  function validate(st, d) {
    const errors = [];
    if (!/^\d{4}$/.test(String(d.code))) errors.push('Code: four digits, as in Home Hub.');
    else if (byCode(st, Number(d.code))) errors.push(`Code ${d.code} is already in the group.`);
    if (!d.name || !String(d.name).trim()) errors.push('Name: required.');
    else if (st.entities.some(e => e.name.toLowerCase() === String(d.name).trim().toLowerCase())) errors.push('Name: another entity has it.');
    if (!byCode(st, Number(d.parent))) errors.push('Parent: pick an entity in the group.');
    const held = Number(d.held);
    if (!(held > 0 && held <= 100)) errors.push('Share held: more than 0% and at most 100%.');
    if (!st.yearStart.includes(d.from)) errors.push('Member from: the start of one of the model\'s years.');
    return errors;
  }

  // What Add entity will do, before it does it.
  function preview(st, d) {
    const parent = byCode(st, Number(d.parent));
    const held = Number(d.held) / 100;
    if (!parent || !(held > 0 && held <= 1)) return null;
    const into = groupsOf(st, parent.code).map(g => g.name);
    const node = held < 1;
    const writes = [
      `Entities: a row under ${parent.name}, after its last entity, status Planned, member from ${d.from}`,
      'Entity data: its rows, nil until it has figures',
      'Group structure: its row in the tree and in each roll-up'
    ];
    if (node) writes.push(`A group of its own (NCI node, outside investors ${Math.round((1 - held) * 100)}%): a column in By group, NCI and the roll-ups`);
    if (Number(d.capital) || Number(d.external)) writes.push('Investments: the parent\'s investment, eliminated against its capital');
    writes.push('Checks: the planned entity alert until Home Hub has it');
    return { parent: parent.name, into: node ? [d.name || 'its own group'].concat(into) : into, owned: held * ownedByTop(st, parent.code),
             node, outside: 1 - held, writes };
  }

  function addEntity(st, d) {
    const errors = validate(st, d);
    if (errors.length) return { ok: false, errors };
    const code = Number(d.code), parent = Number(d.parent), held = Number(d.held) / 100;
    const lastIndex = st.entities.reduce((m, e, i) => (e.code === parent || ancestors(st, e.code).includes(parent) ? i : m), -1);
    const e = { code, name: String(d.name).trim(), parent, held, gst: !!d.gst, role: d.role || '', status: 'Planned', from: d.from,
                own: { surplus: st.years.map(() => 0), na: st.years.map(() => 0) } };
    st.entities.splice(lastIndex + 1, 0, e);
    if (held < 1) st.groups.push({ head: code, name: e.name });
    return { ok: true, entity: e };
  }

  const pct = x => `${Math.round(x * 1000) / 10}%`.replace('.0%', '%');
  const money = x => {
    if (Math.abs(x) < 0.05) return '-';
    const s = Math.abs(x).toLocaleString('en-NZ', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
    return x < 0 ? `(${s})` : s;
  };

  // ---------------------------------------------------------------- views
  const ui = { st: null, year: null, selected: null, tab: 'tree', group: null, flash: null, draft: null };

  function el(tag, attrs, ...kids) {
    const n = document.createElement(tag);
    Object.entries(attrs || {}).forEach(([k, v]) => {
      if (v == null || v === false) return;
      if (k === 'class') n.className = v;
      else if (k.startsWith('on')) n.addEventListener(k.slice(2), v);
      else n.setAttribute(k, v === true ? '' : v);
    });
    kids.flat(Infinity).forEach(k => { if (k != null && k !== false) n.appendChild(typeof k === 'string' ? document.createTextNode(k) : k); });
    return n;
  }

  function ensure() {
    if (!ui.st) {
      ui.st = makeState(root.HfgGroupSample);
      ui.year = ui.st.years.length - 1;
      ui.group = ui.st.groups[0].head;
    }
    return ui.st;
  }

  function yearPicker(rerender) {
    return el('div', { class: 'gseg', role: 'group', 'aria-label': 'Year shown' },
      ui.st.years.map((y, i) => el('button', { class: i === ui.year ? 'on' : null, 'aria-pressed': String(i === ui.year),
        onclick: () => { ui.year = i; rerender(); } }, y)));
  }

  function chips(st, e) {
    const out = [];
    const g = groupOf(st, e.code);
    if (g) out.push(el('span', { class: 'gchip group', title: 'Heads a group that consolidates' }, membersOf(st, e.code).length > 1 ? 'Group' : 'Node'));
    if (e.held < 1) out.push(el('span', { class: 'gchip nci', title: 'Outside investors' }, `NCI ${pct(1 - e.held)}`));
    if (e.status === 'Planned') out.push(el('span', { class: 'gchip planned', title: 'Not in Home Hub yet' }, 'Planned'));
    return out;
  }

  function treeList(st, code, rerender) {
    const e = byCode(st, code);
    const node = el('button', { class: 'gnode' + (ui.selected === code ? ' sel' : '') + (e.status === 'Planned' ? ' planned' : '') +
                                       (ui.flash === code ? ' flash' : ''),
                                'aria-pressed': String(ui.selected === code), onclick: () => { ui.selected = code; rerender(); } },
      el('span', { class: 'gtitle' }, el('span', { class: 'gname' + (groupOf(st, code) ? ' head' : '') }, e.name), ' ',
        el('span', { class: 'gcode' }, String(code))),
      el('span', { class: 'gheld', title: 'Share held by its parent' }, e.parent == null ? '' : pct(e.held)),
      chips(st, e).length ? el('span', { class: 'gchips' }, chips(st, e)) : null);
    const kids = childrenOf(st, code);
    return el('li', {}, node, kids.length ? el('ul', {}, kids.map(k => treeList(st, k.code, rerender))) : null);
  }

  function detailCard(st, code, rerender) {
    const e = byCode(st, code);
    if (!e) return null;
    const p = ui.year;
    const parent = e.parent == null ? null : byCode(st, e.parent);
    const into = groupsOf(st, code);
    const row = (k, v) => [el('dt', {}, k), el('dd', {}, v)];
    return el('div', { class: 'gcard' },
      el('h3', {}, `${e.code}  ${e.name}`),
      e.role ? el('p', { class: 'quiet' }, e.role) : null,
      el('dl', {},
        row('Parent', parent ? `${parent.code} ${parent.name}` : 'None (top of the group)'),
        row('Held by parent', parent ? pct(e.held) : '-'),
        row('Owned by the top', pct(ownedByTop(st, code))),
        row('Outside investors', e.held < 1 ? pct(1 - e.held) + ' (NCI worked out here)' : '-'),
        row('Status', e.status === 'Planned' ? 'Planned (not in Home Hub yet)' : 'Actual (from Home Hub)'),
        row('Member from', e.from),
        row('Rolls into', el('span', {}, into.map((g, i) => el('span', { class: 'gchip' + (i === 0 ? ' first' : '') }, g.name)))),
        row(`Surplus, ${st.years[p]}`, money(e.own.surplus[p])),
        row(`Net assets, ${st.years[p]}`, money(e.own.na[p]))),
      el('div', { class: 'gactions' },
        el('button', { onclick: () => { ui.draft = { parent: code }; show('grp-add'); } }, 'Add entity under it'),
        el('button', { onclick: () => { ui.draft = { entity: code }; show('grp-ownership'); } }, 'Change ownership'),
        groupOf(st, code) ? el('button', { onclick: () => { ui.tab = 'rollup'; ui.group = code; rerender(); } }, 'Roll-up') : null));
  }

  function rollupView(st, rerender) {
    const p = ui.year;
    const r = rollup(st, ui.group, p);
    const g = groupOf(st, ui.group);
    const trail = [];
    for (let x = g; x; x = parentGroup(st, x.head)) trail.unshift(x);
    const crumbs = el('p', { class: 'gcrumbs' }, trail.map((x, i) => [i ? ' › ' : '',
      i === trail.length - 1 ? el('strong', {}, x.name) : el('a', { href: '#', onclick: ev => { ev.preventDefault(); ui.group = x.head; rerender(); } }, x.name)]));
    const tr = (cells, cls) => el('tr', { class: cls }, cells);
    const body = r.lines.map(l => tr([
      el('td', {}, l.kind === 'group'
        ? el('a', { href: '#', title: 'Open this group\'s roll-up', onclick: ev => { ev.preventDefault(); ui.group = l.head; rerender(); } }, l.label + ', consolidated')
        : l.label),
      el('td', { class: 'n' }, money(l.surplus)), el('td', { class: 'n' }, money(l.na))], l.kind));
    return el('div', {},
      el('label', { for: 'g-pick' }, 'Group'),
      el('select', { id: 'g-pick', onchange: ev => { ui.group = Number(ev.target.value); rerender(); } },
        st.groups.map(x => el('option', { value: String(x.head), selected: x.head === ui.group }, x.name))),
      crumbs,
      el('table', { class: 'gbridge' },
        el('thead', {}, tr([el('th', {}, st.years[p]), el('th', { class: 'n' }, 'Surplus'), el('th', { class: 'n' }, 'Net assets')])),
        el('tbody', {}, body,
          tr([el('td', {}, `${g.name}, consolidated`), el('td', { class: 'n' }, money(r.total.surplus)), el('td', { class: 'n' }, money(r.total.na))], 'total'),
          tr([el('td', {}, 'Non-controlling interests'), el('td', { class: 'n' }, money(r.nci.surplus)), el('td', { class: 'n' }, money(r.nci.na))], 'sub'),
          tr([el('td', {}, 'Owners'), el('td', { class: 'n' }, money(r.owners.surplus)), el('td', { class: 'n' }, money(r.owners.na))], 'sub'))),
      el('p', { class: 'quiet' }, 'Sub-groups are already consolidated, so this group adds only the eliminations made in it: pairs whose lowest common group is this one, its investments, its share of the sub-groups\' NCI and margin still held in it.'));
  }

  function structureView(host) {
    const st = ensure();
    const rerender = () => structureView(host);
    host.innerHTML = '';
    if (ui.selected == null) ui.selected = st.entities[0].code;
    host.appendChild(el('div', { class: 'gbar' },
      el('div', { class: 'gseg', role: 'tablist' },
        el('button', { class: ui.tab === 'tree' ? 'on' : null, role: 'tab', 'aria-selected': String(ui.tab === 'tree'), onclick: () => { ui.tab = 'tree'; rerender(); } }, 'Tree'),
        el('button', { class: ui.tab === 'rollup' ? 'on' : null, role: 'tab', 'aria-selected': String(ui.tab === 'rollup'), onclick: () => { ui.tab = 'rollup'; rerender(); } }, 'Roll-up')),
      yearPicker(rerender)));
    if (ui.tab === 'tree') {
      const nodes = st.entities.length, grps = st.groups.length, planned = st.entities.filter(e => e.status === 'Planned').length;
      host.appendChild(el('p', { class: 'quiet' }, `${nodes} entities, ${grps} groups that consolidate` + (planned ? `, ${planned} planned` : '') + '. Click an entity for its details.'));
      const top = st.entities.filter(e => e.parent == null);
      host.appendChild(el('ul', { class: 'gtree' }, top.map(t => treeList(st, t.code, rerender))));
      host.appendChild(detailCard(st, ui.selected, rerender));
      host.appendChild(el('div', { class: 'gactions' },
        el('button', { class: 'primary', onclick: () => { ui.draft = { parent: ui.selected }; show('grp-add'); } }, 'Add entity'),
        el('button', { onclick: () => readWorkbook(host) }, 'Read this workbook')));
    } else {
      host.appendChild(rollupView(st, rerender));
    }
    ui.flash = null;
  }

  function addView(host) {
    const st = ensure();
    const d = Object.assign({ code: '', name: '', parent: st.entities[0].code, held: 100, from: st.yearStart[st.yearStart.length - 1],
                              gst: false, role: 'Limited partnership', capital: '', external: '' }, ui.draft || {});
    ui.draft = d;
    host.innerHTML = '';
    const field = (id, label, input, hint) => el('div', { class: 'gfield' }, el('label', { for: id }, label), input, hint ? el('span', { class: 'quiet' }, hint) : null);
    const set = k => ev => { d[k] = ev.target.type === 'checkbox' ? ev.target.checked : ev.target.value; refresh(); };
    const pv = el('div', { class: 'gcard' });
    const errs = el('ul', { class: 'gerrors', role: 'alert' });
    function refresh() {
      const v = preview(st, d);
      pv.innerHTML = '';
      if (!v) { pv.appendChild(el('p', { class: 'quiet' }, 'Pick a parent and a share held to see where it goes.')); return; }
      pv.appendChild(el('h3', {}, 'What Add entity will do'));
      pv.appendChild(el('p', {}, `Goes under ${v.parent}. Owned by the top: ${pct(v.owned)}.` +
        (v.node ? ` Outside investors hold ${pct(v.outside)}, so it becomes an NCI node and a group of its own.` : '')));
      pv.appendChild(el('p', {}, 'Rolls into: ', v.into.map((n, i) => el('span', { class: 'gchip' + (i === 0 ? ' first' : '') }, n))));
      pv.appendChild(el('ul', {}, v.writes.map(w => el('li', {}, w))));
    }
    const form = el('form', { class: 'gform', novalidate: true, onsubmit: ev => {
      ev.preventDefault();
      const res = addEntity(st, d);
      errs.innerHTML = '';
      if (!res.ok) { res.errors.forEach(x => errs.appendChild(el('li', {}, x))); return; }
      ui.selected = ui.flash = res.entity.code;
      ui.tab = 'tree';
      ui.draft = null;
      show('grp-structure');
    } },
      field('g-code', 'Code', el('input', { id: 'g-code', type: 'text', inputmode: 'numeric', value: d.code, oninput: set('code'), placeholder: '9011' }), 'Four digits, as Home Hub will number it'),
      field('g-name', 'Name', el('input', { id: 'g-name', type: 'text', value: d.name, oninput: set('name'), placeholder: 'Dev LP C' })),
      field('g-parent', 'Parent', el('select', { id: 'g-parent', onchange: set('parent') },
        st.entities.map(e => el('option', { value: String(e.code), selected: String(e.code) === String(d.parent) },
          ' '.repeat(depth(st, e.code) * 3) + `${e.code} ${e.name}`)))),
      field('g-held', 'Share held by parent (%)', el('input', { id: 'g-held', type: 'number', min: '0', max: '100', step: '0.1', value: String(d.held), oninput: set('held') })),
      field('g-from', 'Member from', el('select', { id: 'g-from', onchange: set('from') },
        st.yearStart.map((y, i) => el('option', { value: y, selected: y === d.from }, `${y} (${st.years[i]})`)))),
      el('div', { class: 'grow' },
        field('g-capital', 'Capital from parent ($000)', el('input', { id: 'g-capital', type: 'number', min: '0', value: d.capital, oninput: set('capital') })),
        field('g-external', 'From outside investors ($000)', el('input', { id: 'g-external', type: 'number', min: '0', value: d.external, oninput: set('external') }))),
      field('g-role', 'Type', el('select', { id: 'g-role', onchange: set('role') },
        ['Limited partnership', 'Limited company', 'Charitable trust', 'Unit trust'].map(t => el('option', { selected: t === d.role }, t)))),
      el('label', { class: 'gcheck' }, el('input', { type: 'checkbox', checked: d.gst, onchange: set('gst') }), ' GST registered'),
      el('p', { class: 'quiet' }, 'Status: Planned. Home Hub is the master for actual entities; this one stays flagged in the checks until Home Hub has it, and Refresh from Home Hub matches it then.'),
      errs, pv,
      el('div', { class: 'gactions' }, el('button', { type: 'submit', class: 'primary' }, 'Add entity'),
        el('button', { type: 'button', onclick: () => { ui.draft = null; show('grp-structure'); } }, 'Cancel')));
    host.appendChild(form);
    refresh();
  }

  function ownershipView(host) {
    const st = ensure();
    const d = Object.assign({ entity: st.entities.find(e => e.parent != null).code }, ui.draft || {});
    const e = byCode(st, Number(d.entity));
    if (d.held == null) d.held = Math.round(e.held * 1000) / 10;
    if (d.parent == null) d.parent = e.parent;
    if (d.from == null) d.from = st.yearStart[st.yearStart.length - 1];
    ui.draft = d;
    host.innerHTML = '';
    const out = el('div', { class: 'gcard' });
    function refresh() {
      const now = byCode(st, Number(d.entity));
      const held = Number(d.held) / 100;
      const before = groupsOf(st, now.code).map(g => g.name);
      const trial = makeState({ years: st.years, yearStart: st.yearStart, entities: st.entities.map(x => Object.assign({}, x)), groups: st.groups.map(x => Object.assign({}, x)) });
      const t = byCode(trial, now.code);
      t.parent = Number(d.parent); t.held = held;
      const bad = ancestors(trial, Number(d.parent)).includes(now.code) || Number(d.parent) === now.code;
      out.innerHTML = '';
      out.appendChild(el('h3', {}, 'What changes from ' + d.from));
      if (bad) { out.appendChild(el('p', { role: 'alert' }, 'An entity cannot sit under itself or one of its own subsidiaries.')); return; }
      if (held < 1 && !groupOf(trial, now.code)) trial.groups.push({ head: now.code, name: now.name });
      if (held === 1 && groupOf(trial, now.code) && !childrenOf(trial, now.code).length) trial.groups = trial.groups.filter(g => g.head !== now.code);
      const after = groupsOf(trial, now.code).map(g => g.name);
      out.appendChild(el('dl', {},
        el('dt', {}, 'Owned by the top'), el('dd', {}, `${pct(ownedByTop(st, now.code))} → ${pct(ownedByTop(trial, now.code))}`),
        el('dt', {}, 'Outside investors'), el('dd', {}, `${pct(1 - now.held)} → ${pct(1 - held)}`),
        el('dt', {}, 'NCI node'), el('dd', {}, `${now.held < 1 ? 'Yes' : 'No'} → ${held < 1 ? 'Yes' : 'No'}`),
        el('dt', {}, 'Rolls into'), el('dd', {}, before.join(', ') + (before.join() === after.join() ? ' (no change)' : ` → ${after.join(', ')}`))));
      out.appendChild(el('p', { class: 'quiet' }, 'Groups, eliminations and NCI follow from that date; earlier years keep the old structure. ' +
        'A change in share held that keeps control is a transaction with the outside investors, in equity. Preview only in the probe.'));
    }
    const set = k => ev => { d[k] = ev.target.value; if (k === 'entity') { const x = byCode(st, Number(d.entity)); d.held = Math.round(x.held * 1000) / 10; d.parent = x.parent; ownershipView(host); return; } refresh(); };
    host.appendChild(el('form', { class: 'gform', onsubmit: ev => ev.preventDefault() },
      el('div', { class: 'gfield' }, el('label', { for: 'o-entity' }, 'Entity'), el('select', { id: 'o-entity', onchange: set('entity') },
        st.entities.filter(x => x.parent != null).map(x => el('option', { value: String(x.code), selected: x.code === Number(d.entity) }, `${x.code} ${x.name}`)))),
      el('div', { class: 'grow' },
        el('div', { class: 'gfield' }, el('label', { for: 'o-parent' }, 'Parent'), el('select', { id: 'o-parent', onchange: set('parent') },
          st.entities.map(x => el('option', { value: String(x.code), selected: x.code === Number(d.parent) }, `${x.code} ${x.name}`)))),
        el('div', { class: 'gfield' }, el('label', { for: 'o-held' }, 'Share held (%)'), el('input', { id: 'o-held', type: 'number', min: '0', max: '100', step: '0.1', value: String(d.held), oninput: set('held') }))),
      el('div', { class: 'gfield' }, el('label', { for: 'o-from' }, 'From'), el('select', { id: 'o-from', onchange: set('from') },
        st.yearStart.map((y, i) => el('option', { value: y, selected: y === d.from }, `${y} (${st.years[i]})`)))),
      out,
      el('div', { class: 'gactions' }, el('button', { type: 'button', onclick: () => { ui.draft = null; show('grp-structure'); } }, 'Back to the structure'))));
    refresh();
  }

  async function readWorkbook(host) {
    const msg = text => { const p = host.querySelector('.gread') || host.appendChild(el('p', { class: 'gread quiet', role: 'status' })); p.textContent = text; };
    if (typeof Excel === 'undefined') { msg('Open the consolidation demo in Excel to read its register.'); return; }
    try {
      const names = ['Ent_Codes', 'Ent_Names', 'Ent_Parent', 'Ent_Held', 'Ent_Status', 'Ent_From', 'Tier_Heads', 'Tier_Names'];
      const got = await Excel.run(async ctx => {
        const rs = names.map(n => ctx.workbook.names.getItem(n).getRange().load('values,text'));
        await ctx.sync();
        return rs.map(r => ({ values: r.values, text: r.text }));
      });
      const col = i => got[i].values.map(r => r[0]);
      const st = ensure();
      const years = st.years;
      ui.st = {
        years, yearStart: st.yearStart, note: 'Read from this workbook (structure only; figures stay nil here).',
        entities: col(0).map((c, i) => ({ code: Number(c), name: String(col(1)[i]), parent: col(2)[i] === '' || col(2)[i] === 0 ? null : Number(col(2)[i]),
          held: Number(col(3)[i]), status: String(col(4)[i]), from: got[5].text[i][0], role: '',
          own: { surplus: years.map(() => 0), na: years.map(() => 0) } })),
        groups: got[6].values[0].map((h, k) => ({ head: Number(h), name: String(got[7].values[k][0]) }))
      };
      ui.selected = ui.st.entities[0].code;
      ui.group = ui.st.groups[0].head;
      structureView(host);
      msg(`Read ${ui.st.entities.length} entities and ${ui.st.groups.length} groups from this workbook.`);
    } catch (e) {
      msg('This workbook has no entity register (Ent_ names): ' + ((e && e.message) || e));
    }
  }

  let show = () => {};
  const VIEWS = { 'model-group': structureView, 'grp-structure': structureView, 'grp-add': addView, 'grp-ownership': ownershipView };

  // Renders the view for a command into host; false when the command has no designed view here.
  function render(key, host, onShow) {
    const v = VIEWS[key];
    if (onShow) show = onShow;
    if (!v || !root.HfgGroupSample) return false;
    v(host);
    return true;
  }

  const api = { makeState, ancestors, depth, childrenOf, ownedByTop, membersOf, groupsOf, parentGroup, childGroups, directMembers,
                figures, rollup, validate, preview, addEntity, render, VIEWS: Object.keys(VIEWS), reset: () => { ui.st = null; ui.draft = null; } };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.HfgGroup = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
