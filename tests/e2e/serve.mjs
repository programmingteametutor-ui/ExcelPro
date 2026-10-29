// Tiny static server for dist/ (used by the browser smoke test).
import http from 'node:http'
import fs from 'node:fs'
import path from 'node:path'
const root = path.resolve(process.argv[2] ?? 'dist')
const port = +(process.argv[3] ?? 4173)
const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.svg': 'image/svg+xml' }
http.createServer((req, res) => {
  let p = decodeURIComponent(new URL(req.url, 'http://x').pathname)
  if (p.endsWith('/')) p += 'index.html'
  const f = path.join(root, p)
  if (!f.startsWith(root) || !fs.existsSync(f)) { res.writeHead(404); return res.end('not found') }
  res.writeHead(200, { 'content-type': types[path.extname(f)] ?? 'application/octet-stream' })
  fs.createReadStream(f).pipe(res)
}).listen(port, () => console.log(`serving ${root} on http://localhost:${port}`))
