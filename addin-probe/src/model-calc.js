/*
  The sample model's own calculation, standing in for Excel's recalculation in the probe's
  Explorer and Impacts views. It reproduces the assembly demo's economics (revenue lines, cost
  lines, debtors one month in arrears, debt facilities and the statements) for any settings,
  line by line, keyed by the statement row ids in src/model-sample.js. The Node tests compare
  its totals with the assembly reference written into the sample.
*/
(function (root) {
  'use strict';
  const STMT = 'demo.statements#1';

  function settingsOf(M, over) {
    const S = {};
    M.modules.forEach(m => m.settings.forEach(s => { S[s.name] = s.value; }));
    return Object.assign(S, over || {});
  }

  function calc(M, over) {
    const n = M.periods, S = settingsOf(M, over);
    const set = m => { const o = {}; m.settings.forEach(s => { o[s.key] = Number(S[s.name]); }); return o; };
    const series = {};
    const range = f => Array.from({ length: n }, (_, t) => f(t));
    M.modules.forEach(m => {
      const s = set(m);
      if (m.module === 'demo.revenue_line') series[m.uid] = { 'is.revenue': range(t => s.base * Math.pow(1 + s.growth, t)) };
      if (m.module === 'demo.cost_line') series[m.uid] = { 'is.opex': range(t => s.amount * Math.pow(1 + s.inflation, t)) };
      if (m.module === 'demo.facility') {
        let bal = 0;
        const interest = [], flow = [], closing = [];
        for (let t = 0; t < n; t++) {
          const draw = t === 0 ? s.amount : 0;
          const repay = t > 0 ? Math.min(bal, s.instalment) : 0;
          interest.push(bal * s.rate / 12);
          bal = bal + draw - repay;
          flow.push(draw - repay);
          closing.push(bal);
        }
        series[m.uid] = { 'is.interest': interest, 'cf.financing': flow, 'bs.debt': closing };
      }
    });
    M.modules.filter(m => m.module === 'demo.debtors').forEach(m => m.blocks.filter(b => b.id !== m.uid).forEach(b => {
      const rev = series[b.id.slice(m.uid.length + 1)]['is.revenue'];
      const receipts = range(t => (t ? rev[t - 1] : 0));
      let c = 0;
      const closing = range(t => (c += rev[t] - receipts[t]));
      series[b.id] = { 'cf.receipts': receipts, 'bs.debtors': closing };
    }));
    const lines = {};
    const sum = link => range(t => M.lines.filter(l => l.id.startsWith(`${STMT}/in/${link}/`)).reduce((a, l) => a + lines[l.id][t], 0));
    M.lines.filter(l => l.id.startsWith(STMT + '/in/')).forEach(l => {
      const rest = l.id.slice(STMT.length + 4).split('/');
      lines[l.id] = series[rest.slice(1).join('/')][rest[0]];
    });
    const k = key => `${STMT}/${key}`;
    lines[k('revenue')] = sum('is.revenue');
    lines[k('opex')] = sum('is.opex');
    lines[k('interest')] = sum('is.interest');
    lines[k('surplus')] = range(t => lines[k('revenue')][t] - lines[k('opex')][t] - lines[k('interest')][t]);
    lines[k('receipts')] = sum('cf.receipts');
    lines[k('payments')] = lines[k('opex')];
    lines[k('interest_paid')] = lines[k('interest')];
    lines[k('financing')] = sum('cf.financing');
    lines[k('net_cash')] = range(t => lines[k('receipts')][t] - lines[k('payments')][t] - lines[k('interest_paid')][t] + lines[k('financing')][t]);
    let cash = 0, eq = 0;
    lines[k('cash')] = range(t => (cash += lines[k('net_cash')][t]));
    lines[k('bs_cash')] = lines[k('cash')];
    lines[k('assets')] = range(t => lines[k('bs_cash')][t] + sum('bs.debtors')[t]);
    lines[k('debt')] = sum('bs.debt');
    lines[k('equity')] = range(t => (eq += lines[k('surplus')][t]));
    const checks = {};
    checks[k('chk_balance')] = range(t => (Math.abs(lines[k('assets')][t] - lines[k('debt')][t] - lines[k('equity')][t]) > 0.005 ? 1 : 0));
    checks[k('chk_cash')] = range(t => (lines[k('cash')][t] < 0 ? 1 : 0));
    M.modules.filter(m => m.module === 'demo.facility').forEach(m => {
      checks[`${m.uid}/chk_negative`] = series[m.uid]['bs.debt'].map(v => (v < 0 ? 1 : 0));
    });
    return { lines, checks, settings: S };
  }

  const api = { calc, settingsOf, STMT };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.HfgModelCalc = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
