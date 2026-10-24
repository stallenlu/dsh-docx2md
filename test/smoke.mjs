#!/usr/bin/env node
/**
 * 本地冒烟测试：不启动 DSH，直接调用 py/ 下的脚本，验证插件逻辑。
 * 用法: node test/smoke.mjs <file.docx>
 */
import { spawn } from 'node:child_process'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = fileURLToPath(new URL('../', import.meta.url))
const PY_DIR = join(ROOT, 'py')
const DEPS = ['--from', 'olefile', '--with', 'pypdf']

function run(cmd, args, cwd) {
  return new Promise((done) => {
    const c = spawn(cmd, args, { cwd, stdio: ['ignore', 'pipe', 'pipe'] })
    let out = ''
    let err = ''
    c.stdout.on('data', (b) => { out += b })
    c.stderr.on('data', (b) => { err += b })
    c.on('close', (code) => done({ code, out, err }))
    c.on('error', (e) => done({ code: -1, out, err: String(e) }))
  })
}

const py = (script, args, cwd) =>
  run('uvx', [...DEPS, 'python', join(PY_DIR, script), ...args], cwd)

const src = process.argv[2]
if (!src) {
  console.error('用法: node test/smoke.mjs <file.docx>')
  process.exit(2)
}

const work = '/tmp/dsh_docx2md_smoke'
await mkdir(work, { recursive: true })

console.log('== 1. inspect ==')
const insp = await py('docx_scan.py', [src], work)
if (insp.code !== 0) { console.error(insp.err); process.exit(1) }
const info = JSON.parse(insp.out.trim())
console.log(JSON.stringify(info, null, 1))

console.log('\n== 2. extract ==')
const ext = await py('extract.py', [src, work], work)
if (ext.code !== 0) { console.error(ext.err); process.exit(1) }
const ex = JSON.parse(ext.out.trim())
console.log(JSON.stringify(ex, null, 1))

console.log('\n== 3. render ==')
const ren = await py('render.py', [join(work, 'media'), join(work, 'png'), '--max-width', '1600'], work)
console.log(ren.out.trim())

console.log('\n== 4. build (suggest) ==')
const sug = await py('build.py', [work, '--latex', '/tmp/latex_by_stream.json',
  '--out', join(work, 'out.md'), '--suggest'], work)
console.log(sug.out.trim())

console.log('\n== 5. build ==')
const bld = await py('build.py', [work, '--latex', '/tmp/latex_by_stream.json',
  '--out', join(work, 'out.md'), '--use-suggested'], work)
if (bld.code !== 0) { console.error(bld.err); process.exit(1) }
const b = JSON.parse(bld.out.trim())
console.log(JSON.stringify(b, null, 1))

console.log('\n== 6. validate ==')
const val = await py('validate.py', [join(work, 'out.md'),
  '--expect-formulas', String(info.formulas), '--json'], work)
const v = JSON.parse(val.out.trim())
console.log(JSON.stringify(v, null, 1))

console.log('\n== 判定 ==')
const checks = [
  ['段落数一致', b.paragraphs === ex.paragraphs_nonempty],
  ['公式全部解析', b.unresolved_placeholders === 0],
  ['公式全部有 LaTeX', b.formulas_missing === 0],
  ['公式数吻合', (b.formulas_display + b.formulas_inline) === info.formulas],
  ['校验通过', v.status === 'PASS'],
]
let fail = 0
for (const [label, ok] of checks) {
  console.log((ok ? '  PASS  ' : '  FAIL  ') + label)
  if (!ok) fail += 1
}
process.exit(fail === 0 ? 0 : 1)
