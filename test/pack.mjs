#!/usr/bin/env node
/**
 * pack.mjs —— 发布前的打包内容检查。
 *
 * 回答一个问题：`npm publish` 出去的那个 tarball 里，到底会有什么？
 *
 * 为什么需要它：实测发现 `files: ["py"]` 会把 `py/__pycache__/*.pyc` 一起打进去
 * （3 个 Python 版本的字节码，且 `.npmignore` 在 files 白名单模式下对已纳入的
 * 目录不生效——这一条是实测出来的，与直觉相反）。后果是包体积虚增、用户看到
 * 自己 Python 版本的陌生二进制文件会起疑。改成 `py/*.py` 才真正排除。
 *
 * 检查项：
 *   A. tarball 里只有预期的文件（不多不少，package.json 由 npm 强制包含）
 *   B. 没有 .pyc / __pycache__ / test/ / 编辑器残留
 *   C. 解包后 lib/index.js 能被 import，工具与技能都能注册（不依赖源码目录）
 *   D. 解包后 py/*.py 能独立运行（不依赖源码目录的 sys.path）
 *
 * 用法: node test/pack.mjs     退出码 0 = 可发布
 */
import { spawn } from 'node:child_process'
import { mkdtemp, readFile, rm, mkdir, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = fileURLToPath(new URL('../', import.meta.url))

let pass = 0
let fail = 0
const rec = async (name, fn) => {
  try {
    await fn()
    console.log(`  ok   ${name}`)
    pass += 1
  } catch (e) {
    console.log(`  FAIL ${name}\n       ${e.message}`)
    fail += 1
  }
}
const assert = (c, m) => { if (!c) throw new Error(m) }

/** 期望出现在包里的文件（package.json 由 npm 强制包含，另算）。 */
const EXPECTED = new Set([
  'LICENSE',
  'README.md',
  'cordis.patch.yml',
  'lib/index.js',
  'lib/toolkit.js',
  'package.json',
  'py/build.py',
  'py/docx_scan.py',
  'py/docxfile.py',
  'py/extract.py',
  'py/render.py',
  'py/validate.py',
  'skills/docx2md/SKILL.md',
])

function run(cmd, args, opts = {}) {
  return new Promise((done) => {
    const c = spawn(cmd, args, { cwd: opts.cwd ?? ROOT, stdio: ['ignore', 'pipe', 'pipe'] })
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

const work = await mkdtemp(join(tmpdir(), 'd2m-pack-'))
const pkgDir = await mkdtemp(join(tmpdir(), 'd2m-unpack-'))
try {
  console.log('== A. 打包（npm pack --dry-run --json）==')
  let files
  await rec('npm pack 可执行并返回清单', async () => {
    const r = await run('npm', ['pack', '--dry-run', '--json'])
    assert(r.code === 0, `npm pack 失败: ${r.err.slice(0, 300)}`)
    const data = JSON.parse(r.out)
    files = (data[0]?.files ?? []).map((f) => f.path)
    assert(files.length > 0, '打包清单为空')
  })

  await rec('文件清单与预期完全一致', () => {
    const got = new Set(files)
    const missing = [...EXPECTED].filter((f) => !got.has(f))
    const extra = [...got].filter((f) => !EXPECTED.has(f))
    assert(missing.length === 0, `缺少: ${missing.join(', ')}`)
    assert(extra.length === 0, `多出: ${extra.join(', ')}`)
  })

  console.log('\n== B. 不能进包的东西 ==')
  await rec('没有 .pyc / __pycache__', () => {
    const bad = files.filter((f) => f.includes('__pycache__') || f.endsWith('.pyc'))
    assert(bad.length === 0, `混入了字节码: ${bad.join(', ')}`)
  })
  await rec('没有 test/', () => {
    const bad = files.filter((f) => f.startsWith('test/'))
    assert(bad.length === 0, `混入了测试: ${bad.join(', ')}`)
  })
  await rec('没有编辑器/系统残留', () => {
    const bad = files.filter((f) =>
      /(^|\/)(\.DS_Store|Thumbs\.db|\.vscode|node_modules|\.git)(\/|$)/.test(f)
      || /~$|\.swp$|\.bak$/.test(f))
    assert(bad.length === 0, `混入了残留: ${bad.join(', ')}`)
  })

  console.log('\n== C. 真实打包 + 解包 ==')
  let unpacked = ''
  await rec('npm pack 生成 tarball 并解包', async () => {
    const r = await run('npm', ['pack', '--pack-destination', work])
    assert(r.code === 0, `打包失败: ${r.err.slice(0, 300)}`)
    const tgz = r.out.trim().split('\n').pop().trim()
    const tar = join(work, tgz)
    assert(existsSync(tar), `tarball 不存在: ${tar}`)
    const x = await run('tar', ['xzf', tar, '-C', pkgDir])
    assert(x.code === 0, `解包失败: ${x.err.slice(0, 300)}`)
    unpacked = join(pkgDir, 'package')
  })

  await rec('解包后的 lib/index.js 能注册工具与技能', async () => {
    const probe = join(pkgDir, 'probe.mjs')
    await writeFile(probe, `
const mod = await import(${JSON.stringify('file://' + join(unpacked, 'lib/index.js'))})
const tools = new Map()
let prov = null
mod.apply({
  tools: { register: (t) => tools.set(t.name, t) },
  skills: { registerProvider: (create) => {
    prov = (async () => {
      const p = await create({})
      const list = await p.list()
      const def = await p.get(list[0])
      return { name: list[0].name, len: def.content.length }
    })()
  } },
  workingDirectory: { ensure: async () => ${JSON.stringify(pkgDir)} },
  logger: { warn: (m) => console.error('WARN', m) },
}, {})
const skill = await prov
console.log(JSON.stringify({
  tools: [...tools.keys()].sort(),
  skill,
}))
`, 'utf8')
    const r = await run(process.execPath, [probe])
    assert(r.code === 0, `加载失败: ${(r.err || r.out).slice(0, 400)}`)
    const got = JSON.parse(r.out.trim().split('\n').pop())
    assert(got.tools.length === 3, `工具数 ${got.tools.length}，应为 3`)
    assert(got.tools.includes('docx2md_inspect'), '缺少 docx2md_inspect')
    assert(got.skill?.name === 'docx2md', '技能未注册')
    assert(got.skill.len > 500, `技能内容过短: ${got.skill.len}`)
  })

  await rec('解包后的 py 脚本能独立运行', async () => {
    // 造一个最小 docx，验证 py 目录能自带运行（不依赖源码 sys.path）
    const docx = join(work, 'tiny.docx')
    const py = join(unpacked, 'py')
    const mk = await run('python3', ['-c', `
import zipfile, sys
p = sys.argv[1]
doc = '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>测试</w:t></w:r></w:p></w:body></w:document>'
ct = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
rels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
with zipfile.ZipFile(p, 'w') as z:
    z.writestr('[Content_Types].xml', ct)
    z.writestr('_rels/.rels', rels)
    z.writestr('word/document.xml', doc)
`, docx])
    assert(mk.code === 0, `造样例失败: ${mk.err.slice(0, 200)}`)

    const r = await run('python3', ['docx_scan.py', docx], { cwd: py })
    assert(r.code === 0, `docx_scan.py 失败: ${(r.err || r.out).slice(0, 300)}`)
    const j = JSON.parse(r.out)
    assert(j.paragraphs === 1, `段落数 ${j.paragraphs}，应为 1`)
  })

  console.log('\n== D. package.json 元数据 ==')
  await rec('files 用 py/*.py 而非 py（否则 .pyc 会进包）', async () => {
    const pkg = JSON.parse(await readFile(join(ROOT, 'package.json'), 'utf8'))
    const files = pkg.files ?? []
    assert(!files.includes('py'), 'files 里有裸目录 "py" —— __pycache__ 会跟着进包')
    assert(files.includes('py/*.py'), 'files 里应显式写 "py/*.py"')
  })
  await rec('version / main / engines 齐备', async () => {
    const pkg = JSON.parse(await readFile(join(ROOT, 'package.json'), 'utf8'))
    assert(/^\d+\.\d+\.\d+$/.test(pkg.version), `version 不合规: ${pkg.version}`)
    assert(pkg.main === 'lib/index.js', `main 应为 lib/index.js，实际 ${pkg.main}`)
    assert(pkg.engines?.node, '缺少 engines.node')
    assert(pkg.dsh?.bundle?.patch === './cordis.patch.yml', '缺少 dsh.bundle.patch')
  })
} finally {
  await rm(work, { recursive: true, force: true })
  await rm(pkgDir, { recursive: true, force: true })
}

console.log(`\n${fail === 0 ? 'ALL PASS' : 'FAILED'}  pass=${pass} fail=${fail}`)
process.exit(fail === 0 ? 0 : 1)
