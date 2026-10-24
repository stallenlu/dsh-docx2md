#!/usr/bin/env node
/**
 * encoding.mjs —— 中文编码回归测试。
 *
 * 背景：中文 Windows 上 Python 的 sys.stdout.encoding 默认是 gbk，非 ASCII 字符
 * 会被编成 GBK 字节；调用方按 UTF-8 解码就得到 "����"。这个 bug 在 Linux 上
 * 永远不会出现（Linux 恒为 UTF-8），所以在开发机上是隐形的 —— 它正是在真机
 * （Windows 桌面版）上才暴露出来的。
 *
 * 对策：用 PYTHONIOENCODING=gbk 在 Linux 上伪造同一种输出环境。实测该设置会让
 * python3 把「位错」写成 GBK 字节 316 273 264 355，与 Windows 中文版一致。
 * 于是这条路径变得可在 CI 里复现。
 *
 * 两层防御，各测一遍：
 *   1. 脚本层  —— docxfile.force_utf8_stdio() 无需任何环境变量即把输出钉成 UTF-8
 *   2. 进程层  —— lib/index.js 的 pyEnv() 给子进程塞 PYTHONUTF8/PYTHONIOENCODING
 *
 * 用法: node test/encoding.mjs [docx]     退出码 0 = 全绿
 */
import { spawn } from 'node:child_process'
import { mkdtemp, writeFile, readFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = fileURLToPath(new URL('../', import.meta.url))

const TINY_DOCX_B64 =
  'UEsDBBQAAAAIAMuUWF1FCOEt7QAAAK4BAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbH2QzU7DMBCE' +
  '7zyF5SuKHTgghJL0wM8ROJQHWNmbxMJ/8rqlfXs2bekBVT3as9/MaLrVLnixxUIuxV7eqVYKjCZZ' +
  'F6defq3fmkcpqEK04FPEXu6R5Gq46db7jCQYjtTLudb8pDWZGQOQShkjK2MqASo/y6QzmG+YUN+3' +
  '7YM2KVaMtamLhxy6Fxxh46t43fH3sQjjUjwf75aoXkLO3hmoLOtF1Re5gp6ugNto/7VrTs0Uk4cb' +
  'ml2m21PCBy9TnEXxCaW+Q2A7/ZOK1TaZTeAIdb3ohbw0js7gmV/cckkGiXjy4NVZCeDiXw99mHv4' +
  'BVBLAwQUAAAACADLlFhduYFEcbAAAAAqAQAACwAAAF9yZWxzLy5yZWxzjc87DsIwDAbgnVNE3mla' +
  'BoRQky4IqSsqB4gSN41oHkrCo7cnAwMgBkbbvz/LbfewM7lhTMY7Bk1VA0EnvTJOMzgPx/UOSMrC' +
  'KTF7hwwWTNDxVXvCWeSykyYTEimISwymnMOe0iQntCJVPqArk9FHK3Ipo6ZByIvQSDd1vaXx3QD+' +
  'YZJeMYi9aoAMS8B/bD+ORuLBy6tFl3+c+EoUWUSNmcHdR0XVq10VFihv6ceL/AlQSwMEFAAAAAgA' +
  'y5RYXaCbzZ/cAAAAFwEAABEAAAB3b3JkL2RvY3VtZW50LnhtbLOxr8jNUShLLSrOzM+zVTLUM1BS' +
  'SM1Lzk/JzEu3VQoNcdO1UFIoLknMS0nMyc9LtVWqTC1Wsrfjsim3SslPLs1NzStRAJqQV2xVbquU' +
  'UVJSYKWvX5yckZqbWKyXX5CaB5RLyy/KTSwBcovS9cvzi1IKivKTU4uLgRbk5ugbGRiY6ecmZuYp' +
  '2QGNTMpPqQTRBSCiCESU2D3Z2/tyysynXSueds1+unbZs63dL9ZPfbZu64sJe230QQpAZBGYLEDX' +
  'e765okLhyY4+hXMbHzU1P2pqxKJDH2atPsJLdgBQSwECFAMUAAAACADLlFhdRQjhLe0AAACuAQAA' +
  'EwAAAAAAAAAAAAAAgAEAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQIUAxQAAAAIAMuUWF25gURx' +
  'sAAAACoBAAALAAAAAAAAAAAAAACAAR4BAABfcmVscy8ucmVsc1BLAQIUAxQAAAAIAMuUWF2gm82f' +
  '3AAAABcBAAARAAAAAAAAAAAAAACAAfcBAAB3b3JkL2RvY3VtZW50LnhtbFBLBQYAAAAAAwADALkA' +
  'AAACAwAAAAA='
const PY_DIR = join(ROOT, 'py')

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
const recAsync = rec
const assert = (cond, msg) => { if (!cond) throw new Error(msg) }

/** 跑一个进程，返回 (stdout 字节, 按 UTF-8 严格解码) */
function runRaw(cmd, args, env, cwd) {
  return new Promise((done) => {
    const child = spawn(cmd, args, { env, cwd, stdio: ['ignore', 'pipe', 'pipe'] })
    const out = []
    const err = []
    child.stdout.on('data', (b) => out.push(b))
    child.stderr.on('data', (b) => err.push(b))
    child.on('error', (e) => done({ code: -1, buf: Buffer.alloc(0), errBuf: Buffer.from(String(e)), spawnError: true }))
    child.on('close', (code) => done({ code, buf: Buffer.concat(out), errBuf: Buffer.concat(err) }))
  })
}

/** 严格 UTF-8 解码；非法则抛，不静默替换成 U+FFFD。 */
function strictUtf8(buf) {
  return new TextDecoder('utf-8', { fatal: true }).decode(buf)
}

const PY = process.env.DOCX2MD_TEST_PYTHON || 'python3'
const GBK_ENV = () => ({ ...process.env, PYTHONIOENCODING: 'gbk', PYTHONUTF8: '0' })

console.log('== A. 伪造环境确实是 gbk（否则本测试无意义）==')
await recAsync('PYTHONIOENCODING=gbk 时 python 的 stdout 编码是 gbk', async () => {
  const r = await runRaw(PY, ['-c', 'import sys;print(sys.stdout.encoding)'], GBK_ENV(), PY_DIR)
  assert(r.code === 0, `python 未正常退出: ${strictUtf8(r.errBuf).slice(0, 200)}`)
  assert(strictUtf8(r.buf).trim() === 'gbk', `实际编码是 ${strictUtf8(r.buf).trim()}，伪造失败`)
})
await recAsync('gbk 环境下中文会变成非 UTF-8 字节（对照组）', async () => {
  const r = await runRaw(PY, ['-c', 'print("位错")'], GBK_ENV(), PY_DIR)
  let isUtf8 = true
  try { strictUtf8(r.buf) } catch { isUtf8 = false }
  assert(!isUtf8, '对照组本应产生非 UTF-8 字节，但解码成功了 —— 伪造环境无效')
})

console.log('\n== B. 脚本层防御：force_utf8_stdio() 不依赖环境变量 ==')
const SCRIPTS = ['docx_scan.py', 'extract.py', 'build.py', 'render.py', 'validate.py']
await rec('五个脚本都调用 force_utf8_stdio()', () => {
  return (async () => {
    for (const s of SCRIPTS) {
      const src = await readFile(join(PY_DIR, s), 'utf8')
      assert(src.includes('force_utf8_stdio()'), `${s} 的 __main__ 未调用 force_utf8_stdio()`)
      assert(/from docxfile import .*force_utf8_stdio/.test(src), `${s} 未 import force_utf8_stdio`)
    }
  })()
})
await recAsync('gbk 环境下 force_utf8_stdio() 把编码改成 utf-8', async () => {
  const r = await runRaw(PY,
    ['-c', 'import sys,docxfile;docxfile.force_utf8_stdio();print(sys.stdout.encoding)'],
    GBK_ENV(), PY_DIR)
  assert(r.code === 0, `退出码 ${r.code}: ${strictUtf8(r.errBuf).slice(0, 300)}`)
  assert(strictUtf8(r.buf).trim() === 'utf-8', `实际是 ${strictUtf8(r.buf).trim()}`)
})
await recAsync('gbk 环境下中文输出是合法 UTF-8 且内容正确', async () => {
  const r = await runRaw(PY,
    ['-c', 'import docxfile;docxfile.force_utf8_stdio();print("位错动力学 α=∇×β")'],
    GBK_ENV(), PY_DIR)
  const text = strictUtf8(r.buf)   // 解码失败即断言失败
  assert(text.trim() === '位错动力学 α=∇×β', `内容被改变: ${JSON.stringify(text.trim())}`)
  assert(!text.includes('\uFFFD'), '出现了 U+FFFD 替换字符')
})

console.log('\n== C. 进程层防御：pyEnv() 覆盖父进程的 gbk 设置 ==')
await recAsync('父进程设 PYTHONIOENCODING=gbk，插件仍拿到 UTF-8 中文', async () => {
  const saved = { ...process.env }
  process.env.PYTHONIOENCODING = 'gbk'
  process.env.PYTHONUTF8 = '0'
  try {
    const mod = await import('../lib/index.js')
    const tools = new Map()
    const ctx = {
      tools: { register: (t) => tools.set(t.name, t) },
      skills: { registerProvider: () => {} },
      workingDirectory: { ensure: async () => ROOT },
      logger: { warn: () => {}, info: () => {} },
    }
    mod.apply(ctx, {})
    const inspect = tools.get('docx2md_inspect')
    assert(inspect, '未注册 docx2md_inspect')

    // 用最小 docx 触发一次真实扫描，检查返回的中文是否干净。
    const dir = await mkdtemp(join(tmpdir(), 'd2m-enc-'))
    const sample = join(dir, 'sample.docx')
    await writeFile(sample, buildTinyDocx())
    const r = await inspect.execute({ path: sample }, {})
    const text = JSON.stringify(r)
    assert(!text.includes('\uFFFD'), '返回里出现 U+FFFD —— pyEnv() 未生效')
    assert(r.paragraphs >= 1, `段落数异常: ${r.paragraphs}`)
  } finally {
    process.env = saved
  }
})

console.log('\n== D. 真实文档（可选）==')
const doc = process.argv[2]
if (doc && existsSync(doc)) {
  await recAsync(`真实文档中文段落可读: ${doc.split(/[\\/]/).pop()}`, async () => {
    process.env.PYTHONIOENCODING = 'gbk'
    process.env.PYTHONUTF8 = '0'
    const mod = await import('../lib/index.js')
    const tools = new Map()
    mod.apply({
      tools: { register: (t) => tools.set(t.name, t) },
      skills: { registerProvider: () => {} },
      workingDirectory: { ensure: async () => ROOT },
      logger: { warn: () => {} },
    }, {})
    const r = await tools.get('docx2md_inspect').execute({ path: doc }, {})
    const text = JSON.stringify(r)
    assert(!text.includes('\uFFFD'), '出现 U+FFFD 乱码')
    // 整个返回里必须有汉字：中文文档若全是乱码，这里一个汉字都不会有。
    assert(/[\u4e00-\u9fff]/.test(text), '返回里没有任何汉字 —— 中文输出已损坏')
    const ctxs = r.missing_preview_context ?? []
    if (ctxs.length > 0) {
      // 缺图公式的上下文是模型判断「图注还是正文」的依据，必须可读。
      const withHan = ctxs.filter((c) => /[\u4e00-\u9fff]/.test(c.paragraph_text || ''))
      assert(withHan.length > 0,
        `missing_preview_context 共 ${ctxs.length} 条，没有一条含汉字 —— 段落文字仍是乱码`)
    }
  })
} else {
  console.log('  skip 未传真实 .docx（可传一个中文文档做实测）')
}

/** 造一个只含中文段落的最小合法 docx。 */
function buildTinyDocx() {
  // 最小 OOXML 包；用 Python 的 zipfile 生成以免手写二进制。
  return Buffer.from(TINY_DOCX_B64, 'base64')
}

// 一个 1 段中文的 docx（由 python-docx 结构裁剪而来），base64 内联避免依赖。
/** 一个 2 段中文的最小合法 docx（base64），避免测试依赖外部文件。 */


console.log(`\n${fail === 0 ? 'ALL PASS' : 'FAILED'}  pass=${pass} fail=${fail}`)
process.exit(fail === 0 ? 0 : 1)
