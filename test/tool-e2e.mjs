#!/usr/bin/env node
/**
 * 工具层端到端契约测试：自己造一份合成 .docx，不依赖任何真实文件/fixture。
 *
 * 合成文档故意把 rels 写成 Target 在 Id 之前（OOXML 合法的属性反序），并带
 * 一个 External 超链接和一张图片，用来守住三处历史回归：
 *   1. rels 正则要求 Id 在 Target 前 → 图片/超链接静默丢失（media_extracted=0）
 *   2. rich.read_rels 跳过 External → 外链 URL 丢失
 *   3. lib 宿主侧 resolvePython / extractJson / validate 回传链路断裂
 *
 * 直接跑 lib/index.js 的 apply() → 三个工具的 execute()，不经 DSH 宿主。
 * 退出码 0 = 通过（含「无 Python/uvx 明确跳过」）；1 = 契约失败。
 * 用法: node test/tool-e2e.mjs
 */
import { mkdtemp, writeFile, readFile, rm } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = fileURLToPath(new URL('../', import.meta.url))
const mod = await import('../lib/index.js')

// 最小 .docx（deflate zip）：[Content_Types].xml、_rels/.rels、
// word/document.xml（4 段）、word/_rels/document.xml.rels（rId7 图片 Target
// 在 Id 前 / rId8 External 超链接）、word/media/image1.png。由
// _d2m_testdocs_patch.py 生成并 base64 内嵌，运行时无需 Python 造 fixture。
const B64 = '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'

const tools = new Map()
const providers = []
const ctx = {
  logger: { warn: () => {}, info: () => {} },
  tools: {
    register(t) { tools.set(t.name, t); return () => tools.delete(t.name) },
  },
  skills: {
    registerProvider(fn) { providers.push(fn()) },
  },
  workingDirectory: {
    async ensure() { return ROOT },
  },
}
mod.apply(ctx, undefined)

const checks = []
const ok = (label, cond) => checks.push([label, !!cond])

function noPython(e) {
  return /找不到 Python|no-python|缺少依赖|uvx|missing-deps/.test(String(e))
}

async function main() {
  ok('注册 3 个工具', tools.size === 3)
  ok('注册技能 provider', providers.length === 1)

  const work = await mkdtemp(join(tmpdir(), 'd2m-tool-e2e-'))
  try {
    const src = join(work, 'mini.docx')
    await writeFile(src, Buffer.from(B64, 'base64'))

    let inspect
    try {
      inspect = await tools.get('docx2md_inspect').execute(
        { path: src }, { signal: undefined })
    } catch (e) {
      if (noPython(e)) {
        console.log('SKIP: 本机没有可用的 Python/uvx，跳过工具层 e2e（非失败）。')
        return 0
      }
      throw e
    }

    ok('inspect 返回 paragraphs', typeof inspect.paragraphs === 'number' && inspect.paragraphs > 0)
    ok('inspect 返回 formulas_total', typeof inspect.formulas_total === 'number')

    const outDir = join(work, 'out')
    const ex = await tools.get('docx2md_extract').execute(
      { path: src, outDir }, { signal: undefined })
    ok('extract 回传 warnings 数组', Array.isArray(ex.warnings))
    ok('属性反序 rels 未丢图片（media_extracted >= 1）', ex.media_extracted >= 1)
    ok('extract 产出 eq_index.json', existsSync(join(outDir, 'eq_index.json')))

    const latex = join(work, 'latex.json')
    await writeFile(latex, '{}', 'utf8')
    const bld = await tools.get('docx2md_build').execute({
      workDir: outDir, latex, out: join(work, 'final.md'),
      expectFormulas: inspect.formulas_total,
    }, { signal: undefined })
    ok('build 校验 PASS', bld.validation?.status === 'PASS')

    const md = await readFile(join(work, 'final.md'), 'utf8')
    ok('External 外链 URL 保留', md.includes('https://example.com/x'))
    ok('图片引用可见（![]）', /!\[/.test(md))
  } finally {
    await rm(work, { recursive: true, force: true })
  }

  let fail = 0
  for (const [label, pass] of checks) {
    console.log((pass ? '  PASS  ' : '  FAIL  ') + label)
    if (!pass) fail += 1
  }
  console.log(fail === 0 ? 'ALL PASS' : `FAILED (${fail})`)
  return fail === 0 ? 0 : 1
}

process.exit(await main())
