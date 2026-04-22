/**
 * serve.js — minimal static server for local development.
 * Serves dist/ on http://localhost:3000
 * No dependencies beyond Node built-ins.
 *
 * Usage: node scripts/serve.js
 *        node scripts/serve.js 8080   (custom port)
 */

import { createServer } from 'http';
import { readFile, stat } from 'fs/promises';
import { join, extname } from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dirname, '..', 'dist');
const PORT = parseInt(process.argv[2] ?? '3000', 10);

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.css':  'text/css; charset=utf-8',
  '.js':   'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg':  'image/svg+xml',
  '.png':  'image/png',
  '.ico':  'image/x-icon',
  '.map':  'application/json',
  '.txt':  'text/plain; charset=utf-8',
};

const server = createServer(async (req, res) => {
  let urlPath = new URL(req.url, `http://localhost:${PORT}`).pathname;

  // Strip trailing slash
  if (urlPath !== '/' && urlPath.endsWith('/')) urlPath = urlPath.slice(0, -1);

  let filePath = join(ROOT, urlPath);

  try {
    const info = await stat(filePath);
    if (info.isDirectory()) filePath = join(filePath, 'index.html');
  } catch {
    // Not found as-is — SPA fallback to index.html
    filePath = join(ROOT, 'index.html');
  }

  const ext = extname(filePath).toLowerCase();
  const mime = MIME[ext] ?? 'application/octet-stream';

  try {
    const data = await readFile(filePath);
    res.writeHead(200, {
      'Content-Type': mime,
      'Cache-Control': 'no-cache',
      // Required for SharedArrayBuffer / browser isolation (not needed here but good practice)
      'Cross-Origin-Opener-Policy': 'same-origin',
      'Cross-Origin-Embedder-Policy': 'require-corp',
    });
    res.end(data);
  } catch {
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('404 Not Found: ' + urlPath);
  }
});

server.listen(PORT, () => {
  console.log(`\n  Lex Assistant — dev server`);
  console.log(`  http://localhost:${PORT}\n`);
  console.log(`  Serving: dist/`);
  console.log(`  Stop:    Ctrl+C\n`);
});
