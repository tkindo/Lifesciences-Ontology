// Zero-dependency static file server for the Process Atlas production build.
// Uses only Node.js built-in modules — no `npm install` required.
// CommonJS (.cjs) so it runs standalone regardless of any package.json "type" field.
'use strict';

const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');

const ROOT = path.join(__dirname, 'dist');
const PORT = Number(process.env.PORT) || 8080;
const HOST = process.env.HOST || '0.0.0.0';

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.ico': 'image/x-icon',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
};

function safeResolve(urlPath) {
  const decoded = decodeURIComponent(urlPath.split('?')[0]);
  const normalized = path.normalize(decoded).replace(/^(\.\.[/\\])+/, '');
  return path.join(ROOT, normalized);
}

const server = http.createServer((req, res) => {
  // --- Live Neo4j-backed data for the Pharma sector ---
  // Everything else in this file stays untouched -- this is the one new
  // route. neo4j-driver is only required here, lazily, so the rest of
  // the static server still works even before `npm install` is run.
  if (req.url.startsWith('/api/sector/pharma')) {
    Promise.resolve()
      .then(() => require('./neo4j-adapter.cjs').getPharmaSector())
      .then((data) => {
        const body = JSON.stringify(data);
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(body);
      })
      .catch((err) => {
        console.error('Failed to load pharma sector from Neo4j:', err.message);
        res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ error: err.message }));
      });
    return;
  }

  let filePath = safeResolve(req.url === '/' ? '/index.html' : req.url);

  // Reject any path that escapes the dist root.
  if (!filePath.startsWith(ROOT)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      // SPA fallback: unknown routes serve index.html.
      filePath = path.join(ROOT, 'index.html');
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';

    fs.readFile(filePath, (readErr, data) => {
      if (readErr) {
        res.writeHead(404);
        res.end('Not found');
        return;
      }
      res.writeHead(200, { 'Content-Type': contentType });
      res.end(data);
    });
  });
});

server.listen(PORT, HOST, () => {
  console.log(`Process Atlas is running at http://localhost:${PORT}`);
  console.log('Press Ctrl+C to stop.');
});
