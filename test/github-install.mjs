#!/usr/bin/env node
/**
 * github-install.mjs —— 验证「从 tarball 安装」这条 GitHub 发布路径。
 *
 * 为什么单独测这一条：`file:` 本地目录安装和 tarball 安装是不一样的。
 *   - `file:` 目录 → pnpm 建硬链接，整个目录（含 test/）都在 node_modules 里
 *   - tarball    → 走 npm 的 files 过滤，只有包内文件
 * GitHub 安装最终产出的就是 tarball，所以用户拿到的东西必须靠这条路径验证。
 *
 * 实测中发现的两件事：
 *   1. 在 UNC 路径（\\wsl.localhost\...）下 pnpm 建符号链接会失败
 *      （os error 1，函数不正确）。测试环境必须放在原生 Windows 路径。
 *   2. pnpm 会提示 "Agent skills awaiting approval"。这是 pnpm 对包内 skill
 *      的安全机制，不影响插件功能，但用户需要知道有这个提示。
 *
 * 检查项：
 *   A. tarball 不含 test/ 与 .pyc
 *   B. tarball 安装后 node_modules 内容与 tarball 一致
 *   C. 安装后能注册 3 工具 + 1 技能
 *   D. 安装后能真的处理一份 docx（中文不乱码）
 *   E. cleanup 参数在安装后的版本里存在
 *
 * 用法: node test/github-install.mjs [真实.docx]   退出码 0 = 通过
 */
import { spawn } from 'node:child_process'
import { mkdtemp, mkdir, writeFile, readFile, readdir, rm, copyFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = fileURLToPath(new URL('../', import.meta.url))
const DOC = process.argv[2]

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

function run(cmd, args, opts = {}) {
  return new Promise((done) => {
    const c = spawn(cmd, args, {
      cwd: opts.cwd ?? ROOT,
      stdio: ['ignore', 'pipe', 'pipe'],
      shell: opts.shell === true,
      env: { ...process.env, ...(opts.env ?? {}) },
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

// pnpm 在 UNC 路径下无法建符号链接，测试环境必须放在原生 Windows 路径。
// 在 WSL 里就用 /mnt/c（对应 C:\），在 Windows 上就用 TEMP。
const IS_WIN = process.platform === 'win32'

// Windows 上 node 不能直接 spawn .cmd / .bat（CVE-2024-27980 之后加了限制，报 EINVAL，
// 找不到时是 ENOENT）。npm 与 pnpm 在 Windows 上都是 .cmd，所以两者都要包一层
// cmd.exe /c。实测：cmd.exe /c 不会重写参数里的引号，比 shell:true 更稳
// （shell:true 会把带空格/中文的路径再套一层引号，pnpm 就认不出路径了）。
const NPM_CMD = process.env.DOCX2MD_NPM
  || (IS_WIN ? 'C:\\Program Files\\nodejs\\npm.cmd' : 'npm')
const PNPM_CMD = process.env.DOCX2MD_PNPM
  || (IS_WIN
    ? (process.env.LOCALAPPDATA || 'C:\\Users\\Public\\AppData\\Local')
      + '\\pnpm\\bin\\pnpm.cmd'
    : 'pnpm')

/** 拼出真正要 spawn 的 (cmd, args)：Windows 下统一经 cmd.exe /c */
function tool(winCmd, args) {
  return IS_WIN ? ['cmd.exe', ['/c', winCmd, ...args]] : [winCmd, args]
}
// 与本机路径无关：Windows 上用 %TEMP%，WSL 上映射到 /mnt/c/.../Temp
// （必须落在 NTFS：pnpm 无法在 UNC 路径建符号链接）。
const WIN_TEMP = process.env.TEMP || process.env.TMP
const base = process.env.DOCX2MD_INSTALL_BASE || (IS_WIN
  ? (WIN_TEMP || 'C:\\Windows\\Temp')
  : (WIN_TEMP
      ? '/mnt/' + WIN_TEMP.charAt(0).toLowerCase() + WIN_TEMP.slice(2).replace(/\\/g, '/')
      : '/tmp'))
if (!existsSync(base)) {
  console.log(`  skip 找不到可用的原生路径: ${base}`)
  process.exit(0)
}

// 这条测试需要 Windows 的 pnpm：tarball 安装要在 NTFS 路径下建符号链接。
// WSL 的 Linux node 调 Windows pnpm 会 ENOENT，与其报一堆错不如直接说清楚。
if (!IS_WIN) {
  const where = await run('which', ['pnpm'])
  if (where.code !== 0) {
    console.log('  skip 当前不是 Windows node，且没有可用的 Linux pnpm。')
    console.log('       这条测试请在 Windows 的 node 下运行：')
    console.log('       "C:/Program Files/nodejs/node.exe" test/github-install.mjs <docx>')
    process.exit(0)
  }
}

const work = await mkdtemp(join(base, 'd2m-github-'))
const profile = join(work, 'profile')
const srv = join(work, 'srv')
await mkdir(profile, { recursive: true })
await mkdir(srv, { recursive: true })

/** WSL 路径 → Windows 路径，供 pnpm（Windows 版）使用 */
const toWin = (p) => (IS_WIN ? p : 'C:' + p.replace(/^\/mnt\/c/, '').replace(/\//g, '\\'))

try {
  console.log('== A. 打包 ==')
  let tgz = ''
  await rec('npm pack 成功', async () => {
    const [cn, an] = tool(NPM_CMD, ['pack', '--pack-destination', toWin(srv)])
    const r = await run(cn, an)
    assert(r.code === 0, `打包失败: ${r.err.slice(0, 300)}`)
    tgz = r.out.trim().split('\n').pop().trim()
    assert(existsSync(join(srv, tgz)), `tarball 不存在: ${tgz}`)
  })

  await rec('tarball 不含 test/ 与 .pyc', async () => {
    const r = await run('tar', ['tzf', join(srv, tgz)])
    assert(r.code === 0, '列 tarball 失败')
    const names = r.out.split('\n').map((s) => s.trim()).filter(Boolean)
    assert(!names.some((n) => n.includes('test/')), 'tarball 里混入了 test/')
    assert(!names.some((n) => n.endsWith('.pyc') || n.includes('__pycache__')),
      'tarball 里混入了 .pyc')
  })

  console.log('\n== B. 从 tarball 安装 ==')
  await rec('建干净 profile 并 pnpm add tarball', async () => {
    await writeFile(join(profile, 'package.json'), JSON.stringify({
      name: 'dsh-docx2md-install-test',
      private: true,
      dependencies: {},
      dsh: { profile: { bundles: [], patchReload: 'live' } },
    }, null, 2), 'utf8')
    await writeFile(join(profile, 'pnpm-workspace.yaml'), 'packages: [.]\n', 'utf8')
    await writeFile(join(profile, '.npmrc'), 'nodeLinker: hoisted\n', 'utf8')
    await writeFile(join(profile, 'cordis.patch.yml'), '[]\n', 'utf8')
    await writeFile(join(profile, 'cordis.yml'), 'plugins: []\n', 'utf8')
    await copyFile(join(srv, tgz), join(profile, tgz))

    const [c1, a1] = tool(PNPM_CMD, ['add', `./${tgz}`])
    const r = await run(c1, a1, { cwd: toWin(profile) })
    assert(r.code === 0, `pnpm add 失败:\n${(r.err || r.out).slice(-600)}`)
  })

  await rec('node_modules 里只有包内文件（无 test/、无 .pyc）', async () => {
    // 通过 require.resolve 拿到真实安装位置（.pnpm 里是符号链接）
    const probe = join(profile, '_where.mjs')
    await writeFile(probe, `
import { createRequire } from 'node:module'
const r = createRequire(import.meta.url)
console.log(r.resolve('dsh-docx2md'))
`, 'utf8')
    const r = await run(process.execPath, [probe], { cwd: profile })
    assert(r.code === 0, `解析失败: ${r.err.slice(0, 300)}`)
    const entry = r.out.trim()
    const dir = entry.replace(/[\\/]lib[\\/]index\.js$/, '')
    const names = await readdir(dir)
    assert(!names.includes('test'), 'node_modules 里有 test/')
    const pyc = names.filter((n) => n.endsWith('.pyc'))
    assert(pyc.length === 0, `node_modules 里有 .pyc: ${pyc.join(', ')}`)
    for (const need of ['lib', 'py', 'skills', 'package.json', 'cordis.patch.yml']) {
      assert(names.includes(need), `缺 ${need}`)
    }
  })

  console.log('\n== C. 安装后加载 ==')
  await rec('能注册 3 工具 + 1 技能，且 cleanup 参数存在', async () => {
    const probe = join(profile, '_load.mjs')
    await writeFile(probe, `
import { createRequire } from 'node:module'
const require = createRequire(import.meta.url)
const mod = await import('file:///' + require.resolve('dsh-docx2md').replace(/\\\\/g, '/'))
const tools = new Map()
let prov = null
mod.apply({
  tools: { register: (t) => tools.set(t.name, t) },
  skills: { registerProvider: (create) => {
    prov = (async () => {
      const p = await create({})
      const list = await p.list()
      const d = await p.get(list[0])
      return { name: list[0].name, len: d.content.length }
    })()
  } },
  workingDirectory: { ensure: async () => process.cwd() },
  logger: { warn: () => {} },
}, {})
const skill = await prov
const b = tools.get('docx2md_build')
console.log(JSON.stringify({
  tools: [...tools.keys()].sort(),
  skill,
  cleanup: 'cleanup' in (b.parameters.properties ?? {}),
  cleanupType: b.parameters.properties?.cleanup?.type,
}))
`, 'utf8')
    const r = await run(process.execPath, [probe], { cwd: profile })
    assert(r.code === 0, `加载失败: ${(r.err || r.out).slice(0, 400)}`)
    const got = JSON.parse(r.out.trim().split('\n').pop())
    assert(got.tools.length === 3, `工具数 ${got.tools.length}`)
    assert(got.skill?.name === 'docx2md', '技能未注册')
    assert(got.skill.len > 500, `技能内容过短: ${got.skill?.len}`)
    assert(got.cleanup === true, '安装后的版本里没有 cleanup 参数')
    assert(got.cleanupType === 'boolean', `cleanup 类型应为 boolean，实际 ${got.cleanupType}`)
  })

  console.log('\n== D. 安装后真的能处理文档 ==')
  if (!DOC || !existsSync(DOC)) {
    console.log('  skip 未提供真实 .docx')
  } else {
    await rec('inspect 一份中文 docx，结果正确且不乱码', async () => {
      const dst = join(base, `d2m-doc-${Date.now()}.docx`)
      await copyFile(DOC, dst)
      const winDoc = toWin(dst).replace(/\\/g, '/')
      const probe = join(profile, '_run.mjs')
      await writeFile(probe, `
import { createRequire } from 'node:module'
const require = createRequire(import.meta.url)
const mod = await import('file:///' + require.resolve('dsh-docx2md').replace(/\\\\/g, '/'))
const tools = new Map()
mod.apply({
  tools: { register: (t) => tools.set(t.name, t) },
  skills: { registerProvider: () => {} },
  workingDirectory: { ensure: async () => process.cwd() },
  logger: { warn: () => {} },
}, {})
const r = await tools.get('docx2md_inspect').execute({ path: ${JSON.stringify(winDoc)} }, {})
console.log(JSON.stringify({
  paragraphs: r.paragraphs,
  formulas: r.formulas,
  runtime: r.runtime,
  bad: JSON.stringify(r).includes(String.fromCharCode(0xFFFD)),
  ctxHasHan: (r.missing_preview_context ?? []).every((c) =>
    !c.paragraph_text || /[\\u4e00-\\u9fff]/.test(c.paragraph_text)),
}))
`, 'utf8')
      const r = await run(process.execPath, [probe], { cwd: profile })
      assert(r.code === 0, `运行失败: ${(r.err || r.out).slice(0, 500)}`)
      const got = JSON.parse(r.out.trim().split('\n').pop())
      assert(got.paragraphs > 0, `段落数 ${got.paragraphs}`)
      assert(got.formulas >= 0, '公式数异常')
      assert(got.bad === false, '返回里出现 U+FFFD 乱码 —— 编码防御在安装版里失效')
      assert(got.ctxHasHan === true, '缺图公式的上下文不是可读中文')
      console.log(`       实测: ${got.paragraphs} 段 / ${got.formulas} 公式 / runtime ${got.runtime}`)
      await rm(dst, { force: true })
    })
  }

  console.log('\n== E. 卸载干净 ==')
  await rec('pnpm remove 后 node_modules 无残留', async () => {
    const [c2, a2] = tool(PNPM_CMD, ['remove', 'dsh-docx2md'])
    const r = await run(c2, a2, { cwd: toWin(profile) })
    assert(r.code === 0, `pnpm remove 失败:\n${(r.err || r.out).slice(-400)}`)
    const nm = join(profile, 'node_modules')
    if (existsSync(nm)) {
      const names = await readdir(nm)
      assert(!names.includes('dsh-docx2md'), 'node_modules 里仍有 dsh-docx2md')
    }
    const pkg = JSON.parse(await readFile(join(profile, 'package.json'), 'utf8'))
    const deps = pkg.dependencies ?? {}
    assert(!('dsh-docx2md' in deps), 'package.json 里仍声明依赖')
  })
} finally {
  await rm(work, { recursive: true, force: true })
}

console.log(`\n${fail === 0 ? 'ALL PASS' : 'FAILED'}  pass=${pass} fail=${fail}`)
process.exit(fail === 0 ? 0 : 1)
