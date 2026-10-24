#!/usr/bin/env node
/**
 * hygiene.mjs —— 临时目录与残留文件的行为约束。
 *
 * 这个插件不建系统临时目录（不用 mkdtemp / tempfile），所有中间产物都写在
 * 调用方指定的 outDir 里。本测试把这条约束钉死：
 *
 *   1. 坏输入不留下空目录（否则每失败一次就在磁盘上多一个空壳）
 *   2. 渲染报告 _render_report.json 不在 png/ 里残留
 *   3. outDir 之外不产生任何文件（默认 outDir 就落在 docx 同级）
 *   4. 源码里没有 mkdtemp / tempfile 之类的系统临时目录调用
 *
 * 用法: node test/hygiene.mjs <真实.docx>     退出码 0 = 全绿
 */
import { spawn } from 'node:child_process'
import { mkdtemp, mkdir, writeFile, readFile, readdir, rm, stat } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = fileURLToPath(new URL('../', import.meta.url))
const PY_DIR = join(ROOT, 'py')
const DOC = process.argv[2]

let pass = 0
let fail = 0
const rec = async (name, fn) => {
  try {
    await fn()
    console.log(`  ok   ${name}`)
    pass += 1
  } catch (error) {
    console.log(`  FAIL ${name}\n       ${error.message}`)
    fail += 1
  }
}
const assert = (cond, msg) => { if (!cond) throw new Error(msg) }

function runPy(script, args) {
  return new Promise((done) => {
    const c = spawn('python3', [join(PY_DIR, script), ...args], {
      cwd: PY_DIR,
      env: { ...process.env, PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8' },
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    const o = []
    const e = []
    c.stdout.on('data', (b) => o.push(b))
    c.stderr.on('data', (b) => e.push(b))
    c.on('error', (err) => done({ code: -1, out: '', err: String(err) }))
    c.on('close', (code) => done({
      code,
      out: Buffer.concat(o).toString('utf8'),
      err: Buffer.concat(e).toString('utf8'),
    }))
  })
}

const work = await mkdtemp(join(tmpdir(), 'd2m-hyg-'))
try {
  console.log('== A. 源码：不使用系统临时目录 ==')
  await rec('py/ 与 lib/ 里没有 mkdtemp / tempfile / gettempdir', async () => {
    const bad = []
    for (const dir of ['py', 'lib']) {
      for (const f of await readdir(join(ROOT, dir))) {
        if (!/\.(py|js)$/.test(f)) continue
        const src = await readFile(join(ROOT, dir, f), 'utf8')
        for (const pat of ['mkdtemp', 'tempfile', 'gettempdir', 'os.tmpnam']) {
          if (src.includes(pat)) bad.push(`${dir}/${f}: ${pat}`)
        }
      }
    }
    assert(bad.length === 0, `发现系统临时目录调用: ${bad.join(', ')}`)
  })

  await rec('render.py 读取报告后会删除它', async () => {
    const src = await readFile(join(PY_DIR, 'render.py'), 'utf8')
    const i = src.indexOf('_render_report.json')
    assert(i > 0, '找不到 _render_report.json')
    const tail = src.slice(i)
    assert(/os\.remove\(report\)/.test(tail), '读完报告没有 os.remove(report)')
  })

  console.log('\n== B. 坏输入不留空目录 ==')
  await rec('非 zip 的 .docx：退出码 3 且不建 outDir', async () => {
    const broken = join(work, 'broken.docx')
    await writeFile(broken, 'not a zip at all', 'utf8')
    const outDir = join(work, 'should_not_exist')
    const r = await runPy('extract.py', [broken, outDir])
    assert(r.code === 3, `期望退出码 3，实际 ${r.code}`)
    assert(!existsSync(outDir), `坏输入仍创建了目录 ${outDir}`)
  })

  await rec('不存在的 .docx：退出码 3 且不建 outDir', async () => {
    const outDir = join(work, 'also_should_not_exist')
    const r = await runPy('extract.py', [join(work, 'nope.docx'), outDir])
    assert(r.code === 3, `期望退出码 3，实际 ${r.code}`)
    assert(!existsSync(outDir), `不存在的输入仍创建了目录 ${outDir}`)
  })

  console.log('\n== C. 真实文档：产物只在 outDir 内 ==')
  if (!DOC || !existsSync(DOC)) {
    console.log('  skip 未提供真实 .docx')
  } else {
    const outDir = join(work, 'conv')
    await rec('extract + render 只写 outDir', async () => {
      // 用 outDir 的父目录做快照，事后确认没有多出别的条目。
      const before = new Set(await readdir(work))
      const a = await runPy('extract.py', [DOC, outDir])
      assert(a.code === 0, `extract 失败: ${a.err.slice(0, 300)}`)
      const b = await runPy('render.py', [join(outDir, 'media'),
        join(outDir, 'png'), '--max-width', '1600'])
      assert(b.code === 0, `render 失败: ${b.err.slice(0, 300)}`)
      const after = (await readdir(work)).filter((n) => !before.has(n))
      assert(after.length === 1 && after[0] === 'conv',
        `outDir 之外多出了: ${after.join(', ') || '(无)'}`)
    })

    await rec('png/ 里没有 _render_report.json', async () => {
      const names = await readdir(join(outDir, 'png'))
      assert(!names.includes('_render_report.json'), '渲染报告残留在 png/')
    })

    await rec('outDir 顶层只有已知产物', async () => {
      const names = (await readdir(outDir)).sort()
      const allowed = ['eq_index.json', 'media', 'paras.json', 'png', 'summary.json']
      const extra = names.filter((n) => !allowed.includes(n))
      assert(extra.length === 0, `多出未知产物: ${extra.join(', ')}`)
    })

    await rec('png/ 里除位图外不该有别的临时文件', async () => {
      const names = await readdir(join(outDir, 'png'))
      const odd = names.filter((n) => !/\.(png|jpe?g)$/i.test(n))
      assert(odd.length === 0, `png/ 里有非位图文件: ${odd.join(', ')}`)
    })
  }
} finally {
  await rm(work, { recursive: true, force: true })
}

console.log(`\n${fail === 0 ? 'ALL PASS' : 'FAILED'}  pass=${pass} fail=${fail}`)
process.exit(fail === 0 ? 0 : 1)
