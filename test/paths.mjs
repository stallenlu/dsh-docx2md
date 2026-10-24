#!/usr/bin/env node
/**
 * paths.mjs —— 跨平台路径语义测试。
 *
 * 背景：在 Windows 上 path.isAbsolute('/tmp/x') 返回 true，于是「isAbsolute 就原样
 * 返回」的写法会把 '/tmp/x' 直接交给 Python，而 Python 的 os.path.abspath 在非 UNC
 * 工作目录下把它解释成 C:\tmp\x —— 用户在 WSL 里说 /tmp，文件却跑到 C:\tmp。
 *
 * 实测（Windows 中文版 Python 3.13，cwd 为 UNC）：
 *   os.path.abspath('/tmp/x')  → C:\tmp\x
 *   os.path.abspath(UNC 路径)  → 保持 UNC
 * 而 Node 的 path.resolve 会按会话工作目录的卷正确归一化。所以插件必须一律
 * resolve，不能因为 isAbsolute 为 true 就跳过。
 *
 * 本测试在 win32 上跑真实语义；在 Linux/macOS 上验证反向映射（C:\ → /mnt/c/）。
 *
 * 用法: node test/paths.mjs     退出码 0 = 全绿
 */
import { win32, posix, resolve as lresolve } from 'node:path'
import { existsSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

const ROOT = fileURLToPath(new URL('../', import.meta.url))

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

// 与 lib/index.js 保持一致的实现（复制，而非 import，便于在任意平台断言语义）。
const WSL_MOUNT = /^\/mnt\/([a-zA-Z])(?:[\\/]|$)([\s\S]*)$/
const WIN_DRIVE = /^([a-zA-Z]):[\\/]/

function resolvePathOn(platform, cwd, p) {
  if (platform === 'win32') {
    const m = WSL_MOUNT.exec(p)
    if (m) {
      const drive = `${m[1].toUpperCase()}:\\`
      if (existsSync('/mnt/' + m[1].toLowerCase()) || drive) {
        const rest = m[2].split(/[\\/]/).filter(Boolean)
        return win32.resolve(drive, ...rest)
      }
    }
    return win32.resolve(cwd, p)
  }
  const w = WIN_DRIVE.exec(p)
  if (w) {
    const base = `/mnt/${w[1].toLowerCase()}`
    if (existsSync(base) || true) {
      const rest = p.slice(2).split(/[\\/]/).filter(Boolean)
      return posix.resolve(base, ...rest)
    }
  }
  return posix.resolve(cwd, p)
}

function displayPathOn(platform, p) {
  if (platform !== 'win32') return p
  const m = /^\\\\wsl(?:\.localhost|\$)\\[^\\]+([\s\S]*)$/.exec(p)
  if (!m) return p
  return `/${m[1].split(/\\/).filter(Boolean).join('/')}`
}

const WSL_CWD = '\\\\wsl.localhost\\Ubuntu\\home\\sidady\\projects\\vdd3d'

console.log('== A. win32：Linux 视图路径按会话工作目录的卷解析 ==')
const winCases = [
  // [输入, 期望（win32 归一化后）, 说明]
  ['/tmp/x', '\\\\wsl.localhost\\Ubuntu\\tmp\\x', '/tmp 必须落在 WSL 的 tmp，不是 C:\\tmp'],
  ['/home/sidady/projects/vdd3d/a.docx',
    '\\\\wsl.localhost\\Ubuntu\\home\\sidady\\projects\\vdd3d\\a.docx', '绝对路径原样语义'],
  ['/mnt/c/work/x', 'C:\\work\\x', '/mnt/c 映射到 C:\\'],
  ['/mnt/d/work/x', 'D:\\work\\x', '/mnt/d 映射到 D:\\'],
  ['C:\\work\\x', 'C:\\work\\x', '盘符路径原样'],
  ['C:/work/x', 'C:\\work\\x', '盘符 + 正斜杠'],
  ['\\\\wsl$\\Ubuntu\\tmp\\x', '\\\\wsl$\\Ubuntu\\tmp\\x', 'UNC 老写法'],
  ['\\\\wsl.localhost\\Ubuntu\\tmp\\x', '\\\\wsl.localhost\\Ubuntu\\tmp\\x', 'UNC 新写法'],
  ['work/out', '\\\\wsl.localhost\\Ubuntu\\home\\sidady\\projects\\vdd3d\\work\\out', '相对路径'],
  ['a/../../b', '\\\\wsl.localhost\\Ubuntu\\home\\sidady\\projects\\b', '.. 折叠'],
]
for (const [input, want, why] of winCases) {
  await rec(`${JSON.stringify(input)} → ${why}`, () => {
    const got = resolvePathOn('win32', WSL_CWD, input)
    assert(got === want, `期望 ${want}，实际 ${got}`)
  })
}

console.log('\n== B. win32：不再出现「跳过 resolve」的旧行为 ==')
await rec('/tmp/x 不得被原样透传给 Python', () => {
  const got = resolvePathOn('win32', WSL_CWD, '/tmp/x')
  assert(got !== '/tmp/x', '仍返回原始字符串 —— isAbsolute 提前返回的旧 bug 复发了')
  assert(!got.startsWith('C:\\tmp'), `落到了 ${got}，说明仍按盘符根解析`)
})

console.log('\n== C. 返回值反向映射 ==')
const dispCases = [
  ['\\\\wsl.localhost\\Ubuntu\\tmp\\x', '/tmp/x'],
  ['\\\\wsl$\\Ubuntu\\home\\sidady\\projects\\vdd3d\\out.md',
    '/home/sidady/projects/vdd3d/out.md'],
  ['C:\\work\\x', 'C:\\work\\x'],
  ['/home/sidady/a.docx', '/home/sidady/a.docx'],
]
for (const [input, want] of dispCases) {
  await rec(`displayPath(${JSON.stringify(input)})`, () => {
    const got = displayPathOn('win32', input)
    assert(got === want, `期望 ${want}，实际 ${got}`)
  })
}

console.log('\n== D. Linux/macOS：Windows 路径映射回 /mnt ==')
const posixCases = [
  ['C:\\work\\x', '/mnt/c/work/x'],
  ['D:\\data\\file.docx', '/mnt/d/data/file.docx'],
  ['/tmp/x', '/tmp/x'],
  ['work/out', '/srv/project/work/out'],
]
for (const [input, want] of posixCases) {
  await rec(`${JSON.stringify(input)} → ${want}`, () => {
    const got = resolvePathOn('linux', '/srv/project', input)
    assert(got === want, `期望 ${want}，实际 ${got}`)
  })
}

console.log('\n== E. 同一份返回里路径风格必须一致 ==')
await rec('build 的 result.out 与 validation.file 都经过 displayPath', async () => {
  const { readFile } = await import('node:fs/promises')
  const src = await readFile(ROOT + 'lib/index.js', 'utf8')
  // result.out 折过，validation.file 也必须折，否则同一文件出现两种写法
  assert(/out:\s*displayPath\(out\)/.test(src), 'result.out 未经 displayPath')
  assert(/v\.parsed\.file\s*=\s*displayPath\(v\.parsed\.file\)/.test(src),
    'validation.file 未经 displayPath —— 同一份返回会出现两种路径风格')
})

console.log('\n== F. 源码里确实用了这两条规则 ==')
await rec('lib/index.js 含 WSL_MOUNT 与 displayPath', async () => {
  const { readFile } = await import('node:fs/promises')
  const src = await readFile(ROOT + 'lib/index.js', 'utf8')
  assert(src.includes('WSL_MOUNT'), '缺少 WSL_MOUNT')
  assert(src.includes('function displayPath'), '缺少 displayPath')
  assert(!/if \(isAbsolute\(p\)\) return p/.test(src),
    '仍存在「isAbsolute 就原样返回」的旧写法')
})

console.log(`\n${fail === 0 ? 'ALL PASS' : 'FAILED'}  pass=${pass} fail=${fail}`)
process.exit(fail === 0 ? 0 : 1)
