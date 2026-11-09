#!/usr/bin/env node
/**
 * prepack.mjs —— npm pack / npm publish 前的静态门禁。
 *
 * 关键约束：本脚本绝不能调用 `npm pack`（连 --dry-run 也不行）——
 * `npm pack --dry-run` 同样会触发 prepack，那样会无限递归。真正的
 * tarball 清单核对放在 test/pack.mjs，由 CI 单独跑。
 *
 * **所有输出必须走 stderr**：npm 会把 prepack 的 stdout 转发进 `npm pack --json`
 * 的 stdout，那样会污染 JSON，test/pack.mjs 就解析不了了。
 *
 * 用法: node test/prepack.mjs     退出码 0 = 可打包
 */
import { readdirSync, readFileSync, existsSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = fileURLToPath(new URL('../', import.meta.url))
let fail = 0
const rec = (name, fn) => {
  try { fn(); console.error('  ok   ' + name) }
  catch (e) { console.error('  FAIL ' + name + '\n       ' + e.message); fail += 1 }
}
const assert = (c, m) => { if (!c) throw new Error(m) }

console.error('== prepack 静态门禁 ==')
await rec('package.json 元数据齐备', () => {
  const pkg = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8'))
  assert(/^\d+\.\d+\.\d+$/.test(pkg.version), 'version 不合规')
  assert(pkg.main === 'lib/index.js', 'main 不是 lib/index.js')
  assert(pkg.engines?.node, '缺 engines.node')
  assert(pkg.dsh?.bundle?.patch === './cordis.patch.yml', '缺 dsh.bundle.patch')
  assert((pkg.files ?? []).includes('py/*.py'), 'files 必须是 py/*.py（不能是裸 py）')
  assert(!(pkg.files ?? []).includes('py'), 'files 里有裸 py，__pycache__ 会进包')
})
await rec('打包清单不会带进字节码', () => {
  const pkg = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8'))
  const files = pkg.files ?? []
  // 真正的保证在 files 字段：py/*.py 只收 py/ 下的 .py，子目录里的
  // __pycache__/*.pyc 永远进不了 tarball。若有人把这里改成裸 py，
  // 字节码就会漏进去 —— 这正是要拦的回归。
  assert(files.some((f) => f === 'py/*.py'), 'files 缺 py/*.py')
  assert(!files.some((f) => f === 'py' || f === 'py/' || f === 'py/**'),
    'files 含裸 py 目录，py/__pycache__/*.pyc 会被打进 tarball')
  // 本地 / CI 跑过 Python 测试后 py/__pycache__ 常驻是正常现象，不该让
  // prepack 假失败；上面已确认它不会被打包，这里只提示。
  const py = join(ROOT, 'py')
  if (existsSync(join(py, '__pycache__'))) {
    console.error('       note: 本地有 py/__pycache__（files: py/*.py 不会打进包）')
  }
})
await rec('入口与资源文件齐全', () => {
  for (const f of ['lib/index.js', 'lib/toolkit.js', 'cordis.patch.yml',
                   'icon.svg', 'README.md', 'LICENSE']) {
    assert(existsSync(join(ROOT, f)), '缺文件: ' + f)
  }
})
console.error(fail === 0 ? 'PREPACK OK' : `PREPACK FAILED (${fail})`)
process.exit(fail === 0 ? 0 : 1)
