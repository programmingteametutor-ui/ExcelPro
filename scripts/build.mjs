// Build script: esbuild for TS/React + Tailwind v4 compile API for CSS.
//   node scripts/build.mjs           -> dist/ (index.html + assets)
//   node scripts/build.mjs --single  -> dist/excelpro.html (everything inlined, one file)
//   node scripts/build.mjs --serve   -> dev server on http://localhost:5173 with rebuild on change
import { fileURLToPath } from 'node:url'
import * as esbuild from 'esbuild'
import { compile } from 'tailwindcss'
import fs from 'node:fs'
import path from 'node:path'
import { createRequire } from 'node:module'

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const require = createRequire(path.join(root, 'package.json'))
const args = new Set(process.argv.slice(2))
const single = args.has('--single')
const serve = args.has('--serve')
const outdir = path.join(root, 'dist')

function walk(dir) {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) =>
    e.isDirectory() ? walk(path.join(dir, e.name)) : [path.join(dir, e.name)],
  )
}

async function buildCss() {
  const entry = path.join(root, 'src/styles.css')
  const twDir = path.dirname(require.resolve('tailwindcss/package.json'))
  const compiler = await compile(fs.readFileSync(entry, 'utf8'), {
    base: path.dirname(entry),
    loadStylesheet: async (id, base) => {
      let file
      if (id === 'tailwindcss') file = path.join(twDir, 'index.css')
      else if (id.startsWith('tailwindcss/')) file = path.join(twDir, id.slice('tailwindcss/'.length))
      else file = path.resolve(base, id)
      return { path: file, base: path.dirname(file), content: fs.readFileSync(file, 'utf8') }
    },
  })
  const candidates = new Set()
  for (const f of walk(path.join(root, 'src'))) {
    if (!/\.(tsx?|html)$/.test(f)) continue
    for (const m of fs.readFileSync(f, 'utf8').matchAll(/[^\s"'`{}<>]+/g)) candidates.add(m[0])
  }
  const css = compiler.build([...candidates])
  const min = await esbuild.transform(css, { loader: 'css', minify: !serve })
  return min.code
}

const jsOptions = {
  entryPoints: [path.join(root, 'src/main.tsx')],
  bundle: true,
  format: 'esm',
  target: 'es2020',
  jsx: 'automatic',
  minify: !serve,
  sourcemap: serve ? 'inline' : false,
  define: { 'process.env.NODE_ENV': serve ? '"development"' : '"production"' },
  alias: { '@': path.join(root, 'src') },
  logLevel: 'warning',
  legalComments: 'none',
}

function html({ css, js, inline }) {
  const head = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>ExcelPro</title>
<link rel="icon" href="data:image/svg+xml,${encodeURIComponent(
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32"><rect width="32" height="32" rx="7" fill="#0f766e"/><path d="M8 9h16M8 16h16M8 23h16M14 6v20" stroke="#fff" stroke-width="2.4" stroke-linecap="round"/></svg>',
  )}" />
${inline ? `<style>${css}</style>` : '<link rel="stylesheet" href="assets/app.css" />'}
</head>
<body>
<div id="root"></div>
${inline ? `<script type="module">${js.replace(/<\/script/gi, '<\\/script')}</script>` : '<script type="module" src="assets/app.js"></script>'}
</body>
</html>
`
  return head
}

async function buildOnce() {
  fs.mkdirSync(path.join(outdir, 'assets'), { recursive: true })
  const css = await buildCss()
  const result = await esbuild.build({ ...jsOptions, write: false })
  const js = result.outputFiles[0].text
  if (single) {
    fs.writeFileSync(path.join(outdir, 'excelpro.html'), html({ css, js, inline: true }))
    console.log(`dist/excelpro.html  ${(fs.statSync(path.join(outdir, 'excelpro.html')).size / 1024).toFixed(0)} KB`)
  } else {
    fs.writeFileSync(path.join(outdir, 'assets/app.css'), css)
    fs.writeFileSync(path.join(outdir, 'assets/app.js'), js)
    fs.writeFileSync(path.join(outdir, 'index.html'), html({ css, js, inline: false }))
    console.log(`dist/ built: app.js ${(js.length / 1024).toFixed(0)} KB, app.css ${(css.length / 1024).toFixed(0)} KB`)
  }
}

if (serve) {
  await buildOnce()
  let timer
  fs.watch(path.join(root, 'src'), { recursive: true }, () => {
    clearTimeout(timer)
    timer = setTimeout(() => buildOnce().catch((e) => console.error(e.message)), 120)
  })
  const ctx = await esbuild.context({ entryPoints: [], write: false })
  const { port } = await ctx.serve({ servedir: outdir, port: 5173 })
  console.log(`ExcelPro dev server: http://localhost:${port}`)
} else {
  await buildOnce()
}
