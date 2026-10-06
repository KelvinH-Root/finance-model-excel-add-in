/* Small helpers shared by the probe's designed views (Explorer, Impacts). */
(function (root) {
  'use strict';
  function el(tag, attrs, ...kids) {
    const n = tag === 'svg' || ['g', 'rect', 'text', 'path', 'line', 'title', 'defs', 'marker', 'tspan'].includes(tag)
      ? document.createElementNS('http://www.w3.org/2000/svg', tag) : document.createElement(tag);
    Object.entries(attrs || {}).forEach(([k, v]) => {
      if (v == null || v === false) return;
      if (k === 'class') n.setAttribute('class', v);
      else if (k.startsWith('on') && typeof v === 'function') n.addEventListener(k.slice(2), v);
      else n.setAttribute(k, v === true ? '' : v);
    });
    kids.flat(Infinity).forEach(k => { if (k != null && k !== false) n.appendChild(typeof k === 'string' || typeof k === 'number' ? document.createTextNode(String(k)) : k); });
    return n;
  }
  const money = x => {
    if (Math.abs(x) < 0.05) return '-';
    const s = Math.abs(x).toLocaleString('en-NZ', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
    return x < 0 ? `(${s})` : s;
  };
  const pct = x => `${Math.round(x * 1000) / 10}%`;
  const code = s => String(s).replace(/\W/g, '_');
  // Select a defined name's range in the open workbook, when there is one (the probe also runs without Excel).
  async function selectName(name) {
    if (typeof Excel === 'undefined') return false;
    try {
      return await Excel.run(async ctx => {
        const nm = ctx.workbook.names.getItemOrNullObject(name);
        await ctx.sync();
        if (nm.isNullObject) return false;
        nm.getRange().select();
        await ctx.sync();
        return true;
      });
    } catch (e) { return false; }
  }
  root.HfgViewKit = { el, money, pct, code, selectName };
})(typeof globalThis !== 'undefined' ? globalThis : this);
