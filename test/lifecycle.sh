#!/usr/bin/env bash
# lifecycle.sh —— 插件的安装/使用/卸载闭环验证。
#
# 用一个**独立的一次性 profile** 做测试，不碰用户真实的 desktop profile。
# 一旦验证通过，同样的步骤可以套用到任何 profile。
#
# 用法: ./test/lifecycle.sh [论文.docx]
# 退出码 0 = 闭环成立。
#
# 路径约定：全部从 lib/discover.sh 来，没有写死任何用户目录。一次性 profile 用
# mktemp -d 生成，所以别人 clone 仓库后也能跑。需要 Windows 程序（node.exe /
# pnpm.cmd）时要求本机是 WSL 或 Windows；否则明确跳过并返回 0（跳过不是失败）。
#
# 一个踩过的坑：内嵌的 node 脚本原先写在 <<'EOF' heredoc 里。为了给里面的路径做
# 变量替换而去掉引号后，JS 模板字符串里的 ${...} 会被 shell 抢先展开，把脚本弄坏
# （表现为 "bad substitution"、"command not found"）。所以内嵌脚本一律落到独立
# 文件再执行——shell 只负责把路径通过环境变量传进去。
set -u
cd "$(dirname "$0")/.." || exit 1

ROOT="$(pwd)"
. "$(dirname "$0")/lib/discover.sh"

if [ -z "${WIN_TMP:-}" ]; then
  echo "skip 这条测试需要 Windows 侧的 node 与 pnpm（当前环境未检测到）。"
  echo "     Linux/macOS 上请改用：node test/plugin.mjs <docx>"
  exit 0
fi

NODE="$(find_win_node)" || { echo "skip 找不到 Windows node"; exit 0; }
PNPM_CMD="${DOCX2MD_PNPM:-$(win_pnpm)}" || { echo "skip 找不到 pnpm"; exit 0; }

TMPBASE="${TMPBASE%/}"
PROFILE_WSL="$(mktemp -d "$TMPBASE/d2m-lifecycle-profile-XXXXXX")"
PKG_WSL="$(mktemp -d "$TMPBASE/d2m-lifecycle-pkg-XXXXXX")"
PROFILE="$(to_win_path "$PROFILE_WSL")"
PKG="$(to_win_path "$PKG_WSL")"
TMPW="$(to_win_path "$TMPBASE")"

DOC="${1:-}"
DOC_WIN=""
if [ -n "$DOC" ] && [ -f "$DOC" ]; then
  cp "$DOC" "$PROFILE_WSL/_demo.docx"
  DOC_WIN="$(to_win_path "$PROFILE_WSL/_demo.docx")"
fi

pass=0; fail=0
step() { printf '\n\033[1m── %s\033[0m\n' "$1"; }
ok()   { pass=$((pass+1)); printf '  \033[32m✓\033[0m %s\n' "$1"; }
no()   { fail=$((fail+1)); printf '  \033[31m✗\033[0m %s\n' "$1"; }
check(){ if eval "$2" >/dev/null 2>&1; then ok "$1"; else no "$1"; fi; }

# 调用 Windows 的 pnpm：win_run 已封装 cmd.exe /c + cd /d + 路径换算

# ---------------------------------------------------------------- 0. 准备
step "0. 准备一次性 profile"
mkdir -p "$PROFILE_WSL"
cp -r lib py skills cordis.patch.yml package.json README.md LICENSE "$PKG_WSL"/
cat > "$PROFILE_WSL/package.json" <<'JSON'
{
  "name": "dsh-profile-lifecycle",
  "private": true,
  "dependencies": {},
  "dsh": { "profile": { "bundles": [], "patchReload": "live" } }
}
JSON
printf 'packages: [.]\n' > "$PROFILE_WSL/pnpm-workspace.yaml"
printf 'nodeLinker: hoisted\n' > "$PROFILE_WSL/.npmrc"
printf '[]\n' > "$PROFILE_WSL/cordis.patch.yml"
printf 'plugins: []\n' > "$PROFILE_WSL/cordis.yml"

check "profile 骨架已建立" "[ -f '$PROFILE_WSL/package.json' ]"
check "包副本已放置" "[ -f '$PKG_WSL/lib/index.js' ]"

# ---------------------------------------------------------------- 1. 安装前
step "1. 安装前：插件不应存在"
check "bundles 为空" "python3 -c \"import json,sys;d=json.load(open('$PROFILE_WSL/package.json'));sys.exit(0 if not d['dsh']['profile']['bundles'] else 1)\""
check "node_modules 无此包" "[ ! -d '$PROFILE_WSL/node_modules/dsh-docx2md' ]"
out=$(cd "$PROFILE_WSL" && "$NODE" -e "import('file:///$PROFILE/node_modules/dsh-docx2md/lib/index.js').then(()=>console.log('LOADED')).catch(()=>console.log('NOT_FOUND'))" 2>/dev/null | tail -1)
check "模块不可加载" "echo '$out' | grep -q NOT_FOUND"

# ---------------------------------------------------------------- 2. 安装
step "2. 安装（pnpm add file:）"
inst=$(win_run "$PROFILE" "$PNPM_CMD" add "file:$PKG" 2>&1; echo "RC=$?")
check "pnpm add 成功" "echo '$inst' | grep -q 'RC=0'"
check "lock 已记录" "grep -q 'dsh-docx2md' '$PROFILE_WSL/pnpm-lock.yaml'"
check "node_modules 出现" "[ -d '$PROFILE_WSL/node_modules/dsh-docx2md' ]"
check "package.json 写入依赖" "grep -q 'dsh-docx2md' '$PROFILE_WSL/package.json'"

# ---------------------------------------------------------------- 3. bundles
step "3. 加入 bundles"
python3 - "$PROFILE_WSL/package.json" <<'PY'
import json, sys
p = sys.argv[1]
d = json.load(open(p, encoding='utf-8'))
b = d.setdefault('dsh', {}).setdefault('profile', {}).setdefault('bundles', [])
if 'dsh-docx2md' not in b:
    b.append('dsh-docx2md')
json.dump(d, open(p, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
PY
check "bundles 已含插件" "python3 -c \"import json,sys;d=json.load(open('$PROFILE_WSL/package.json'));sys.exit(0 if 'dsh-docx2md' in d['dsh']['profile']['bundles'] else 1)\""

# ---------------------------------------------------------------- 4. 加载
step "4. 加载与注册"
cat > "$PROFILE_WSL/_probe.mjs" <<'JSEOF'
import { createRequire } from 'node:module'
const require = createRequire(import.meta.url)
const mod = await import('file:///' + require.resolve('dsh-docx2md').replace(/\\/g, '/'))
const tools = new Map()
let skillP = null
mod.apply({
  tools: { register: (t) => tools.set(t.name, t) },
  skills: { registerProvider: (create) => {
    skillP = (async () => {
      const p = await create({})
      const list = await p.list()
      return list[0].name
    })()
  } },
  workingDirectory: { ensure: async () => process.cwd() },
  logger: { warn: () => {} },
}, {})
console.log(JSON.stringify({
  tools: [...tools.keys()].sort(),
  skill: [await skillP],
  exports: Object.keys(mod).sort(),
}))
JSEOF
res=$(cd "$PROFILE_WSL" && "$NODE" _probe.mjs 2>&1 | tail -1)
check "3 个工具已注册" "echo '$res' | grep -q 'docx2md_build'"
check "技能已注册" "echo '$res' | grep -q 'docx2md'"
check "Config 可 resolve" "echo '$res' | grep -q exports"

# ---------------------------------------------------------------- 5. 使用
step "5. 实际使用（inspect / extract / build）"
if [ -z "$DOC_WIN" ]; then
  printf '  \033[33m—\033[0m 未提供 docx，跳过实际转换（用法: ./test/lifecycle.sh 论文.docx）\n'
else
cat > "$PROFILE_WSL/_use.mjs" <<'JSEOF'
import { createRequire } from 'node:module'
import { writeFileSync, readFileSync } from 'node:fs'
const require = createRequire(import.meta.url)
const mod = await import('file:///' + require.resolve('dsh-docx2md').replace(/\\/g, '/'))
const tools = new Map()
mod.apply({
  tools: { register: (t) => tools.set(t.name, t) },
  skills: { registerProvider: () => {} },
  workingDirectory: { ensure: async () => process.cwd() },
  logger: { warn: () => {} },
}, {})

const DOC = process.argv[2]
const WORK = process.argv[3]

const insp = await tools.get('docx2md_inspect').execute({ path: DOC }, {})
const ex = await tools.get('docx2md_extract').execute({ path: DOC, outDir: WORK }, {})

// 最小 latex 表：只为验证链路，不求数学正确
const eq = JSON.parse(readFileSync(WORK + '\\eq_index.json', 'utf8'))
const map = {}
for (const r of eq) map[r.stream] = '\\alpha'
writeFileSync(WORK + '\\latex.json', JSON.stringify(map), 'utf8')

const b = await tools.get('docx2md_build').execute({
  workDir: WORK, latex: WORK + '\\latex.json', out: WORK + '\\out.md',
  expectFormulas: eq.length, useSuggestedHeadings: true,
}, {})

console.log(JSON.stringify({
  formulas: insp.formulas,
  rendered: ex.render?.rendered,
  unresolved: b.unresolved_placeholders,
  validation: b.validation?.status,
}))
JSEOF
  use=$(cd "$PROFILE_WSL" && "$NODE" _use.mjs "$DOC_WIN" "$TMPW\\d2m-use-out" 2>&1 | tail -1)
  check "inspect 得到公式" "echo '$use' | grep -qE '\"formulas\":[1-9]'"
  check "渲染全部成功" "echo '$use' | grep -qE '\"rendered\":[1-9]'"
  check "build 无未解析公式" "echo '$use' | grep -q '\"unresolved\":0'"
  check "校验 PASS" "echo '$use' | grep -q '\"validation\":\"PASS\"'"
fi

# ---------------------------------------------------------------- 6. 卸载
step "6. 卸载"
un=$(win_run "$PROFILE" "$PNPM_CMD" remove dsh-docx2md 2>&1; echo "RC=$?")
check "pnpm remove 成功" "echo '$un' | grep -q 'RC=0'"

# ---------------------------------------------------------------- 7. 干净
step "7. 卸载后：应完全干净"
python3 - "$PROFILE_WSL/package.json" <<'PY'
import json, sys
p = sys.argv[1]
d = json.load(open(p, encoding='utf-8'))
b = d.setdefault('dsh', {}).setdefault('profile', {}).setdefault('bundles', [])
if 'dsh-docx2md' in b:
    b.remove('dsh-docx2md')
d.setdefault('dependencies', {}).pop('dsh-docx2md', None)
json.dump(d, open(p, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
PY
check "node_modules 已移除" "[ ! -d '$PROFILE_WSL/node_modules/dsh-docx2md' ]"
check "lock 已无记录" "! grep -q dsh-docx2md '$PROFILE_WSL/pnpm-lock.yaml'"
check "bundles 已清空" "python3 -c \"import json,sys;d=json.load(open('$PROFILE_WSL/package.json'));sys.exit(0 if 'dsh-docx2md' not in d['dsh']['profile']['bundles'] else 1)\""
check "dependencies 已清空" "python3 -c \"import json,sys;d=json.load(open('$PROFILE_WSL/package.json'));sys.exit(0 if 'dsh-docx2md' not in d['dependencies'] else 1)\""
out=$(cd "$PROFILE_WSL" && "$NODE" -e "import('file:///$PROFILE/node_modules/dsh-docx2md/lib/index.js').then(()=>console.log('LOADED')).catch(()=>console.log('NOT_FOUND'))" 2>/dev/null | tail -1)
check "模块不可加载（已彻底移除）" "echo '$out' | grep -q NOT_FOUND"

# ---------------------------------------------------------------- 8. 重装
step "8. 重装：证明可反复安装"
re=$(win_run "$PROFILE" "$PNPM_CMD" add "file:$PKG" 2>&1; echo "RC=$?")
check "可重装" "echo '$re' | grep -q 'RC=0'"
check "重装后模块可加载" "[ -f '$PROFILE_WSL/node_modules/dsh-docx2md/lib/index.js' ]"

# ---------------------------------------------------------------- 9. 清理
step "9. 清理一次性环境"
rm -rf "$PROFILE_WSL" "$PKG_WSL"
check "profile 已删除" "[ ! -d '$PROFILE_WSL' ]"
check "包副本已删除" "[ ! -d '$PKG_WSL' ]"

printf '\n'
if [ $fail -eq 0 ]; then
  printf '\033[32m闭环成立：%d 项全部通过（0 失败）\033[0m\n' "$pass"
  exit 0
fi
printf '\033[31m%d 通过 / %d 失败\033[0m\n' "$pass" "$fail"
exit 1
