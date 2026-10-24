#!/usr/bin/env bash
# update.sh —— 验证「更新」这条路是否真的走得通
#
# 对 file: 依赖来说，"更新"不是一个动作而是三个不同的问题：
#   A. 改了文件、版本号不变  → 宿主重启后拿到的是新代码吗？
#   B. 改了版本号 0.1.0→0.1.1 → 需要 pnpm install 吗？lock 会变吗？
#   C. node_modules 是硬链接还是拷贝 → 改动何时可见？
#
# 每种都实测，不靠推断。
set -u
cd "$(dirname "$0")/.." || exit 1

# 路径不写死：一次性目录用 mktemp，Windows 侧路径由 to_win_path 换算。
. "$(dirname "$0")/lib/discover.sh"

if [ -z "${WIN_TMP:-}" ]; then
  echo "skip 这条测试需要 Windows 侧的 node 与 pnpm（当前环境未检测到）。"
  exit 0
fi
NODE="$(find_win_node)" || { echo "skip 找不到 Windows node"; exit 0; }
PNPM_CMD="${DOCX2MD_PNPM:-$(win_pnpm)}" || { echo "skip 找不到 pnpm"; exit 0; }

PROFILE_WSL="$(mktemp -d "$TMPBASE/dsh-update-profile-XXXXXX")"
PKG_WSL="$(mktemp -d "$TMPBASE/dsh-update-pkg-XXXXXX")"
PROFILE="$(to_win_path "$PROFILE_WSL")"
PKG="$(to_win_path "$PKG_WSL")"
# 内嵌 JS 里用（反斜杠已转义）
PROFILE_JS="$(printf '%s' "$PROFILE" | sed 's/\\/\\\\/g')"
PROFILE_FWD="$(printf '%s' "$PROFILE" | sed 's#\\#/#g')"
PKG_FWD="$(printf '%s' "$PKG" | sed 's#\\#/#g')"
PKG_JS="$(printf '%s' "$PKG" | sed 's/\\/\\\\/g')"

pass=0; fail=0
step(){ printf '\n\033[1m── %s\033[0m\n' "$1"; }
ok(){ pass=$((pass+1)); printf '  \033[32m✓\033[0m %s\n' "$1"; }
no(){ fail=$((fail+1)); printf '  \033[31m✗\033[0m %s\n' "$1"; }
check(){ if eval "$2" >/dev/null 2>&1; then ok "$1"; else no "$1"; fi; }

# 读取 node_modules 里那份代码暴露的版本标记。
# 注意用 PROFILE_FWD（正斜杠），因为 `file:///` URL 里不能出现反斜杠。
probe_ver() {
  "$NODE" --input-type=module -e "
  const m = await import('file:///$PROFILE_FWD/node_modules/dsh-docx2md/lib/index.js')
  console.log(m.VERSION_MARK || 'none')
  " 2>/dev/null | tail -1
}

step "0. 建立一次性环境和初始安装"
rm -rf "$PROFILE_WSL" "$PKG_WSL"; mkdir -p "$PROFILE_WSL" "$PKG_WSL"
cp -r lib py skills cordis.patch.yml package.json README.md LICENSE "$PKG_WSL"/
cat > "$PROFILE_WSL/package.json" <<'JSON'
{ "name":"dsh-profile-update","private":true,"dependencies":{},
  "dsh":{"profile":{"bundles":["dsh-docx2md"],"patchReload":"live"}} }
JSON
printf 'packages:\n  - .\nnodeLinker: hoisted\nautoInstallPeers: false\n' > "$PROFILE_WSL/pnpm-workspace.yaml"
printf '[]\n' > "$PROFILE_WSL/cordis.patch.yml"

"$NODE" -e "
const {execSync}=require('child_process');
try{execSync('pnpm add file:$PKG_FWD',
  {cwd:'$PROFILE_JS',encoding:'utf8',stdio:'pipe'})}catch(e){}
" >/dev/null 2>&1
check "初始安装完成" "[ -f '$PROFILE_WSL/node_modules/dsh-docx2md/lib/index.js' ]"

# ---------------------------------------------------------------- A. 改代码不改版本
step "A. 只改代码、版本号不变（最常见的开发迭代）"
echo "export const VERSION_MARK = 'A-code-only';" >> "$PKG_WSL/lib/index.js"
check "修改已写进 packages（源目录）" "grep -q 'A-code-only' '$PKG_WSL/lib/index.js'"
check "改后 node_modules 立刻可见（硬链接）" "grep -q 'A-code-only' '$PROFILE_WSL/node_modules/dsh-docx2md/lib/index.js'"
v=$(probe_ver); echo "    加载到的标记: $v"
check "宿主加载到的是新代码" "[ \"$v\" = 'A-code-only' ]"

# ---------------------------------------------------------------- B. 改版本号
step "B. 改版本号（0.1.0 → 0.1.9）"
python3 - <<PY
import json
p='$PKG_WSL/package.json'
d=json.load(open(p,encoding='utf-8')); d['version']='0.1.9'
json.dump(d,open(p,'w',encoding='utf-8'),ensure_ascii=False,indent=2)
PY
echo "    源版本: $(python3 -c "import json;print(json.load(open('$PKG_WSL/package.json'))['version'])")"
echo "    node_modules 版本: $(python3 -c "import json;print(json.load(open('$PROFILE_WSL/node_modules/dsh-docx2md/package.json'))['version'])")"
# 硬链接情况下 package.json 也会同步
check "版本号是否随硬链接自动同步" "grep -q '0.1.9' '$PROFILE_WSL/node_modules/dsh-docx2md/package.json'"
echo "    lock 里的 version 字段:"
grep -A3 'dsh-docx2md:' "$PROFILE_WSL/pnpm-lock.yaml" | head -4 | sed 's/^/      /'

# ---------------------------------------------------------------- C. 手动同步的边界
step "C. 当硬链接被断开时（编辑器原子写入）"
# 模拟：删除后重建，产生新 inode
rm -f "$PKG_WSL/lib/index.js"
cp lib/index.js "$PKG_WSL/lib/index.js"
echo "export const VERSION_MARK = 'C-atomic-write';" >> "$PKG_WSL/lib/index.js"
a=$(stat -c %i "$PKG_WSL/lib/index.js")
b=$(stat -c %i "$PROFILE_WSL/node_modules/dsh-docx2md/lib/index.js")
if [ "$a" = "$b" ]; then echo "    仍为同一 inode（硬链接未断）"; else echo "    inode 不同 → 硬链接已断"; fi
v=$(probe_ver); echo "    加载到的标记: $v"
if [ "$v" = "C-atomic-write" ]; then
  ok "断开后仍拿到新代码（该环境未断链）"
else
  # 这是**预期行为**：硬链接断开后 node_modules 保留旧内容。
  # 记录为事实，并验证 pnpm install 能修复——修复成功即算通过。
  echo "    （预期）断开后 node_modules 仍是旧代码：$v"
  ok "已确认断链会导致 node_modules 滞后（预期的陷阱）"
  echo "    修复：重跑 pnpm install"
  "$NODE" -e "
  const {execSync}=require('child_process');
  try{execSync('pnpm install',{cwd:'$PROFILE_JS',encoding:'utf8',stdio:'pipe'})}catch(e){}
  " >/dev/null 2>&1
  v2=$(probe_ver); echo "    pnpm install 后: $v2"
  check "pnpm install 能恢复" "[ \"$v2\" = 'C-atomic-write' ]"
fi

# ---------------------------------------------------------------- D. 干净重装是否拿到新版
step "D. remove + add（彻底重装）"
"$NODE" -e "
const {execSync}=require('child_process');
for(const c of ['pnpm remove dsh-docx2md','pnpm add file:$PKG_FWD']){
  try{execSync(c,{cwd:'$PROFILE_JS',encoding:'utf8',stdio:'pipe'})}catch(e){}
}" >/dev/null 2>&1
v=$(probe_ver); echo "    加载到的标记: $v"
check "干净重装后拿到最新代码" "[ \"$v\" = 'C-atomic-write' ]"
check "版本号为 0.1.9" "grep -q '0.1.9' '$PROFILE_WSL/node_modules/dsh-docx2md/package.json'"

# ---------------------------------------------------------------- 结论项
step "E. 关键结论校验"
echo "    file: 依赖在 lock 中的类型:"
grep -A4 'dsh-docx2md@file' "$PROFILE_WSL/pnpm-lock.yaml" | sed 's/^/      /' | head -5

step "清理"
rm -rf "$PROFILE_WSL" "$PKG_WSL"
check "已清理" "[ ! -d '$PROFILE_WSL' ]"

printf '\n'
if [ $fail -eq 0 ]; then
  printf '\033[32m更新路径验证完成：%d 项通过\033[0m\n' "$pass"; exit 0
else
  printf '\033[31m%d 通过 / %d 失败\033[0m\n' "$pass" "$fail"; exit 1
fi
