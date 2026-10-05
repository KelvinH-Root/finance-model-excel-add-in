// Serves the probe over HTTPS on https://localhost:3000 using the Office
// development certificate (run `npm run certs` once first).
import { createServer } from 'node:https';
import { readFile, stat } from 'node:fs/promises';
import { extname, join, normalize, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import devCerts from 'office-addin-dev-certs';

const ROOT = resolve(fileURLToPath(new URL('.', import.meta.url)));
const PORT = Number(process.env.PORT || 3000);
const TYPES = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8', '.xml': 'application/xml; charset=utf-8',
  '.png': 'image/png', '.svg': 'image/svg+xml', '.md': 'text/plain; charset=utf-8', '.map': 'application/json'
};

const options = await devCerts.getHttpsServerOptions();

createServer(options, async (req, res) => {
  try {
    const path = decodeURIComponent(new URL(req.url, 'https://localhost').pathname);
    const file = normalize(join(ROOT, path === '/' ? '/src/taskpane.html' : path));
    if (!file.startsWith(ROOT)) { res.writeHead(403).end(); return; }
    const info = await stat(file);
    if (!info.isFile()) { res.writeHead(404).end(); return; }
    res.writeHead(200, {
      'Content-Type': TYPES[extname(file)] || 'application/octet-stream',
      'Cache-Control': 'no-store',
      'Access-Control-Allow-Origin': '*'
    });
    res.end(await readFile(file));
  } catch {
    res.writeHead(404).end();
  }
}).listen(PORT, () => console.log(`HFG probe served at https://localhost:${PORT}/src/taskpane.html`));
