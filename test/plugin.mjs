#!/usr/bin/env node
/**
 * 插件层测试：用 mock ctx 直接跑 apply()，验证工具注册与真实执行。
 * 用法: node test/plugin.mjs <file.docx>
 */
import { mkdir, writeFile, readFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { join } from 'node:path'

const mod = await import('../lib/index.js')

// ---- mock ctx ----
const tools = new Map()
const providers = []
const ctx = {
  logger: { warn: (m) => console.log('[warn]', m), info: () => {} },
  tools: {
    register(tool) {
      tools.set(tool.name, tool)
      return () => tools.delete(tool.name)
    },
  },
  skills: {
    registerProvider(fn) { providers.push(fn()) },
  },
  workingDirectory: {
    async ensure() { return process.cwd() },
  },
}

mod.apply(ctx, { maxImageWidth: 1600, allowUvx: true })
console.log('注册的工具:', [...tools.keys()])
console.log('注册的技能:', providers.map((p) => p.name))

const src = process.argv[2]
if (!src) { console.log('\n(未给 docx 路径，仅验证注册)'); process.exit(0) }

const work = '/tmp/plugin_test_work'
await mkdir(work, { recursive: true })

async function call(name, args) {
  const tool = tools.get(name)
  if (!tool) throw new Error('no such tool: ' + name)
  const t0 = Date.now()
  const out = await tool.execute(args, { signal: undefined })
  console.log(`\n== ${name} (${Date.now() - t0} ms) ==`)
  return out
}

// 1. inspect
const info = await call('docx2md_inspect', { path: src })
console.log(JSON.stringify(info, null, 1).slice(0, 900))

// 2. extract
const ex = await call('docx2md_extract', { path: src, outDir: work })
console.log('paragraphs:', ex.paragraphs, 'formulas:', ex.formulas)
console.log('render:', JSON.stringify(ex.render))

// 3. 造一份 LaTeX 表（用已知的映射，验证 build 链路）
const latexPath = join(work, 'latex.json')
const eqIndex = JSON.parse(await readFile(join(work, 'eq_index.json'), 'utf8'))
const latex = {}
for (const row of eqIndex) latex[row.stream] = `\\alpha_{${row.stream.replace(/\D/g, '')}}`
await writeFile(latexPath, JSON.stringify(latex, null, 1))

const sug = await call('docx2md_build', {
  workDir: work, latex: latexPath, out: join(work, 'out.md'), suggestHeadings: true,
})
console.log(JSON.stringify(sug))

const bld = await call('docx2md_build', {
  workDir: work, latex: latexPath, out: join(work, 'out.md'),
  useSuggestedHeadings: true, expectFormulas: info.formulas,
})
console.log('headings:', bld.headings, 'display:', bld.formulas_display,
  'inline:', bld.formulas_inline, 'missing:', bld.formulas_missing)
console.log('validation:', JSON.stringify(bld.validation))

const v = bld.validation
const checks = [
  ['工具全部注册', tools.size === 3],
  ['技能已注册', providers.length === 1],
  ['提取到公式', info.formulas > 0],
  ['渲染成功', ex.render?.ok === true],
  ['无未解析占位', bld.unresolved_placeholders === 0],
  ['公式数吻合', (bld.formulas_display + bld.formulas_inline) === info.formulas],
  ['校验通过', v?.status === 'PASS'],
]
console.log('\n== 判定 ==')
let fail = 0
for (const [label, ok] of checks) {
  console.log((ok ? '  PASS  ' : '  FAIL  ') + label)
  if (!ok) fail += 1
}
process.exit(fail ? 1 : 0)
