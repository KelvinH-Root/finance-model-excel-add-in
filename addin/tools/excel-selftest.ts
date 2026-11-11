// The Excel self-test harness: builds fictional test workbooks with the package writer, opens each
// in Excel desktop on this PC with the self-test task pane (addin-probe/selftest-manifest.xml),
// lets the pane run the add-in's commands (src/selftest/run.ts), then compares what Excel held
// after every step with a fresh build of the model the workbook's metadata describes.
//
//   node tools/excel-selftest.ts [--books full_model,budget_model] [--smoke] [--steps n] [--diagnose] [--timeout 900] [--keep]
//
// Results go to build/selftest/<run>/ (report.md, report.json, each step's snapshot, and a copy of
// each workbook as Excel saved it, without the test pane, to look over). Windows with Excel only.
// Uses the Office development certificate (npm run certs in addin-probe, once).

import { execFileSync, spawn } from 'node:child_process';
import { X509Certificate } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { createServer } from 'node:https';
import { homedir } from 'node:os';
import { dirname, extname, join, normalize, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { assemble, buildWorkbook, Model, type Brand, type Recipe } from '../../engine/src/index.ts';
import { loadLogo } from '../../engine/src/node/assets.ts';
import { loadLibrary } from '../../engine/src/node/library.ts';
import { readModel } from '../src/insert/core.ts';
import { compare, expectedOf, withoutTaskPane, withTaskPane, type Comparison } from '../src/selftest/compare.ts';
import type { Config, Snapshot, Step, StepResult } from '../src/selftest/run.ts';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '..', '..');
const PROBE = join(ROOT, 'addin-probe');
const MANIFEST = join(PROBE, 'selftest-manifest.xml');
const ADDIN_ID = /<Id>([^<]+)<\/Id>/.exec(readFileSync(MANIFEST, 'utf8'))![1];
const DEV_KEY = 'HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer';
const PORT = 3100;

const args = process.argv.slice(2);
const opt = (name: string, fallback: string) => (args.includes(name) ? args[args.indexOf(name) + 1] : fallback);
const smoke = args.includes('--smoke');
const keep = args.includes('--keep');
/** Apply each operation on its own and time each kind, so a refusal names the operation. */
const diagnose = args.includes('--diagnose');
/** Run only the first n steps of each scenario. */
const stepsLimit = Number(opt('--steps', '999'));
const timeoutMs = Number(opt('--timeout', '900')) * 1000;

/** What each test workbook is built from and the steps the pane runs on it. */
const SCENARIOS: Record<string, { recipe: string; steps: Step[] }> = {
  full_model: {
    recipe: 'full_model.json',
    steps: [
      { op: 'add', module: 'fm.revenue' },
      { op: 'above', module: 'fm.revenue', nth: 2 },
      { op: 'duplicate', module: 'fm.opex' },
      { op: 'rename', module: 'fm.revenue', last: true, title: 'Self-test fees' },
      { op: 'remove', module: 'fm.revenue', last: true },
      { op: 'insert', module: 'fm.debt' },
      { op: 'delete', module: 'fm.debt', last: true },
      { op: 'save-version' },
      { op: 'properties', title: 'Full model (self-test)', preparedBy: 'Prepared by the Excel self-test' },
    ],
  },
  budget_model: {
    recipe: 'budget_model.json',
    steps: [{ op: 'add', module: 'fm.opex' }, { op: 'save-version' }],
  },
  consolidation_model: { recipe: 'consolidation_model.json', steps: [] },
};

const books = smoke ? ['full_model'] : opt('--books', Object.keys(SCENARIOS).join(',')).split(',');
const run = new Date().toISOString().replace(/[-:]/g, '').replace('T', '-').slice(0, 15);
const out = join(ROOT, 'build', 'selftest', run);
mkdirSync(out, { recursive: true });

const libs = { hfg: loadLibrary(join(ROOT, 'library', 'hfg')), demo: loadLibrary(join(ROOT, 'prototypes', 'assembly', 'library')) };
const fileOf = (book: string) => `selftest-${book}.xlsx`;
const build = (model: Model) => buildWorkbook(assemble(model), model, {
  logo: loadLogo((model.info?.entity.brand ?? 'HF') as Brand), created: new Date('2026-11-11T00:00:00Z'),
});

// ------------------------------------------------------------------ the test workbooks

for (const book of books) {
  const sc = SCENARIOS[book];
  if (!sc) throw new Error(`No scenario ${book}; choose from ${Object.keys(SCENARIOS).join(', ')}.`);
  const recipe = JSON.parse(readFileSync(join(ROOT, 'library', 'hfg', 'recipes', sc.recipe), 'utf8')) as Recipe;
  writeFileSync(join(out, fileOf(book)), withTaskPane(build(Model.fromRecipe(libs.hfg, recipe)), ADDIN_ID));
}

// ------------------------------------------------------------------ the server the pane talks to

interface BookRun { snaps: Snapshot[]; done?: { host: unknown; results: StepResult[]; saved: boolean }; stopped?: string }
const state = new Map<string, BookRun>(books.map(b => [fileOf(b), { snaps: [] }]));
const waiters = new Map<string, () => void>();
const started = new Map<string, () => void>();

/** Close a test workbook Excel still has open (the pane never started, or stopped), without saving. */
function closeInExcel(file: string): void {
  const ps = `try { $x = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); $x.Workbooks.Item('${file}').Close($false) } catch {}`;
  try { execFileSync('powershell.exe', ['-NoProfile', '-Command', ps], { stdio: 'ignore', timeout: 30000 }); } catch { /* Excel has closed it already */ }
}
const TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.png': 'image/png', '.svg': 'image/svg+xml', '.json': 'application/json',
};

const certs = join(homedir(), '.office-addin-dev-certs');
if (!existsSync(join(certs, 'localhost.crt'))) throw new Error('No Office development certificate: run npm run certs in addin-probe once.');
// Excel blocks the pane ("not signed by a valid security certificate") once the certificate has expired
const validTo = new Date(new X509Certificate(readFileSync(join(certs, 'localhost.crt'))).validTo);
if (validTo < new Date()) {
  throw new Error(`The Office development certificate expired on ${validTo.toDateString()}, so Excel will block the test pane. `
    + 'Renew it in addin-probe with: npx office-addin-dev-certs uninstall, then npx office-addin-dev-certs install --days 365 (Windows asks you to trust the new certificate).');
}

const server = createServer({ cert: readFileSync(join(certs, 'localhost.crt')), key: readFileSync(join(certs, 'localhost.key')) }, (req, res) => {
  const url = new URL(req.url ?? '/', `https://localhost:${PORT}`);
  const chunks: Buffer[] = [];
  req.on('data', c => chunks.push(c));
  req.on('end', () => {
    try {
      const body = Buffer.concat(chunks).toString('utf8');
      const m = /^\/selftest\/([^/]+)\/([^/]+)\/(snapshot|done)$/.exec(url.pathname);
      if (url.pathname === '/selftest/config') {
        const file = url.searchParams.get('book') ?? '';
        const book = books.find(b => fileOf(b) === file);
        if (!book) { res.writeHead(404).end(); return; }
        started.get(file)?.();
        const cfg: Config = { run, book: file, steps: SCENARIOS[book].steps.slice(0, stepsLimit), close: true, diagnose };
        res.writeHead(200, { 'Content-Type': 'application/json' }).end(JSON.stringify(cfg));
        return;
      }
      if (m && req.method === 'POST' && decodeURIComponent(m[1]) === run && state.has(decodeURIComponent(m[2]))) {
        const st = state.get(decodeURIComponent(m[2]))!;
        if (m[3] === 'snapshot') {
          const snap = JSON.parse(body) as Snapshot;
          st.snaps.push(snap);
          writeFileSync(join(out, `${decodeURIComponent(m[2])}.${String(st.snaps.length).padStart(2, '0')}.json`), body);
          console.log(`  ${decodeURIComponent(m[2])}: ${snap.label}`);
        } else {
          st.done = JSON.parse(body);
          waiters.get(decodeURIComponent(m[2]))?.();
        }
        res.writeHead(204).end();
        return;
      }
      if (url.pathname === '/selftest/stopped') {
        const st = state.get(url.searchParams.get('book') ?? '');
        if (st) { st.stopped = body; waiters.get(url.searchParams.get('book')!)?.(); }
        res.writeHead(204).end();
        return;
      }
      const file = normalize(join(PROBE, url.pathname));
      if (!file.startsWith(PROBE) || !existsSync(file)) { res.writeHead(404).end(); return; }
      res.writeHead(200, { 'Content-Type': TYPES[extname(file)] ?? 'application/octet-stream', 'Cache-Control': 'no-store' });
      res.end(readFileSync(file));
    } catch (e) {
      res.writeHead(500).end(String(e));
    }
  });
});

await new Promise<void>((ok, fail) => { server.once('error', fail); server.listen(PORT, ok); });

// ------------------------------------------------------------------ run each workbook in Excel

execFileSync('reg', ['add', DEV_KEY, '/v', ADDIN_ID, '/t', 'REG_SZ', '/d', MANIFEST, '/f'], { stdio: 'ignore' });
console.log(`Self-test ${run}: ${books.length} workbook(s) in ${out}`);
try {
  for (const book of books) {
    const file = fileOf(book);
    const path = join(out, file);
    console.log(`Opening ${file} in Excel`);
    const finished = new Promise<'done' | 'timeout' | 'no pane'>(ok => {
      const timer = setTimeout(() => ok('timeout'), timeoutMs);
      // the pane asks for its steps within a minute or two of the workbook opening, or it is not running
      const noPane = setTimeout(() => { clearTimeout(timer); ok('no pane'); }, 180_000);
      started.set(file, () => clearTimeout(noPane));
      waiters.set(file, () => { clearTimeout(timer); clearTimeout(noPane); ok('done'); });
    });
    spawn('cmd.exe', ['/d', '/s', '/c', `start "" "${path}"`], { windowsVerbatimArguments: true, detached: true, stdio: 'ignore' }).unref();
    const how = await finished;
    if (how !== 'done') {
      console.log(`  ${file}: ${how === 'no pane' ? 'the self-test pane did not start (is the add-in blocked in Excel?)' : `no result after ${timeoutMs / 1000} seconds`}`);
      state.get(file)!.stopped ??= how === 'no pane' ? 'The self-test pane did not start.' : 'Timed out.';
      closeInExcel(file);
    }
    // give Excel a moment to close the workbook before the next one opens
    await new Promise(r => setTimeout(r, 3000));
  }
} finally {
  server.close();
  if (!keep) execFileSync('reg', ['delete', DEV_KEY, '/v', ADDIN_ID, '/f'], { stdio: 'ignore' });
}

// ------------------------------------------------------------------ compare and report

const report: { run: string; books: Record<string, { host: unknown; results: StepResult[]; comparisons: Comparison[]; stopped?: string; saved: boolean }> } = { run, books: {} };
const lines: string[] = [`# Excel self-test ${run}`, ''];
for (const book of books) {
  const file = fileOf(book);
  const st = state.get(file)!;
  const comparisons = st.snaps.map(s => {
    try {
      const open = readModel(s.metadata, libs);
      return compare(s, expectedOf(build(open.model)));
    } catch (e) {
      return { label: s.label, error: (e as Error).message } as unknown as Comparison;
    }
  });
  report.books[book] = { host: st.done?.host ?? null, results: st.done?.results ?? [], comparisons, stopped: st.stopped, saved: !!st.done?.saved };
  if (st.done?.saved) writeFileSync(join(out, `${book} for review.xlsx`), withoutTaskPane(new Uint8Array(readFileSync(join(out, file)))));
  const host = st.done?.host as { version?: string; platform?: string; excelApi?: string[] } | undefined;
  lines.push(`## ${book}`, '', host ? `Excel ${host.version ?? '?'} on ${host.platform ?? '?'}; ExcelApi ${host.excelApi?.join(', ') ?? '?'}` : 'No result from Excel.', '');
  if (st.stopped) lines.push(`Stopped: ${st.stopped.split('\n')[0]}`, '');
  lines.push('| Step | Result | ms | Cells compared | Mismatches | Live drop-downs | Names missing / extra / different | Error cells | Errors / alerts |', '|---|---|---:|---:|---:|---:|---|---:|---|');
  st.snaps.forEach((s, k) => {
    const r = st.done?.results[k];
    const c = comparisons[k] as Comparison & { error?: string };
    const names = c.error ? '' : `${c.names.missing.length} / ${c.names.extra.length} / ${c.names.different.length}`;
    lines.push(`| ${s.label} | ${r ? (r.ok ? 'ok' : 'failed') : '?'} | ${r?.ms ?? ''} | ${c.error ? `compare failed: ${c.error}` : c.compared} | ${c.error ? '' : c.mismatches} | ${c.error ? '' : c.liveDropdowns} | ${names} | ${s.errorCells.length} | ${s.checks.errors ?? '?'} / ${s.checks.alerts ?? '?'} |`);
  });
  lines.push('', 'What each step said:', '');
  for (const r of st.done?.results ?? []) lines.push(`- ${r.label}: ${r.message}`);
  lines.push('');
}
writeFileSync(join(out, 'report.json'), JSON.stringify(report, null, 1));
writeFileSync(join(out, 'report.md'), lines.join('\n'));
console.log(lines.join('\n'));
console.log(`\nReport: ${join(out, 'report.md')}`);
