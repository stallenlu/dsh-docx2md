#!/usr/bin/env bash
# sync-check.sh —— 确认「源码 / 安装副本 / node_modules」三处代码一致。
#
# 为什么需要这个脚本：
#   插件开发时改的是源码目录，但真正被 DSH 加载的是 node_modules 里的那份
#   （pnpm 用硬链接指向 ~/.dsh/packages/）。如果只改源码不复制，跑真机验证时
#   用的还是旧代码，测试结论就是假的。这个坑在本次开发中踩了三次，每次都
#   表现为「改了没生效」，浪费了一整轮排查。
#
# 判定三件事：
#   1. 源码 → 安装副本：package.json files 字段里的每个文件哈希是否一致
#   2. 安装副本 → node_modules：是否仍是硬链接（inode 相同）
#   3. 已加载进程：DSH 进程启动时间是否早于文件修改时间（早 = 需重启）
#
# 用法:
#   ./test/sync-check.sh            # 只检查
#   ./test/sync-check.sh --fix      # 发现不一致时自动同步源码 → 安装副本
#
# 退出码: 0 = 三处一致且无待重启提示；1 = 有不一致（或需要重启）

set -u

FIX=0
[ "${1:-}" = "--fix" ] && FIX=1

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SRC="$(cd "$HERE/.." && pwd)"

# 本机 DSH 的安装位置是**天然本机特有**的，不可能可移植。所以不写死路径，而是
# 按「显式覆盖 > 自动发现 > 明确跳过」处理。对 clone 仓库的其他人来说，这个
# 脚本的检查项本来就不适用（他们不是在本机开发这个插件），此时应当成功跳过，
# 而不是失败——否则别人的 CI 会莫名变红。
. "$HERE/lib/discover.sh"

# 着色函数要在跳过分支之前定义——否则跳过时会出现「bold: command not found」。
red()  { printf '\033[31m%s\033[0m\n' "$1"; }
grn()  { printf '\033[32m%s\033[0m\n' "$1"; }
wrn()  { printf '\033[33m%s\033[0m\n' "$1"; }
bold() { printf '\033[1m%s\033[0m\n' "$1"; }

PROFILE="${DSH_PROFILE:-}"
DST="${DSH_PACKAGE:-}"

if [ -z "$PROFILE" ] || [ -z "$DST" ]; then
  bold "== 本机开发环境未就绪，跳过 =="
  echo "  没找到本机 DSH 的 desktop profile 或已安装的插件副本。"
  echo "  这个脚本检查的是「源码 → 安装副本 → node_modules」三处是否同步，"
  echo "  只对**在本机开发这个插件**的人有意义；普通使用者不需要它。"
  echo ""
  echo "  如需指定位置："
  echo "    DOCX2MD_PROFILE=/path/to/.dsh/profiles/desktop \\"
  echo "    DOCX2MD_PACKAGE=/path/to/.dsh/packages/dsh-docx2md \\"
  echo "    ./test/sync-check.sh"
  exit 0
fi

NM="$PROFILE/node_modules/dsh-docx2md"

fail=0          # 代码不同步（硬失败）
need_restart=0  # 代码已同步，但进程加载的是旧版本（软提示）

# ---------------------------------------------------------------- 0. 前置
bold "== 0. 路径检查 =="
for d in "$SRC" "$DST" "$NM"; do
  if [ -d "$d" ]; then
    echo "  ok   $d"
  else
    red   "  缺失 $d"
    fail=1
  fi
done
[ $fail -ne 0 ] && { echo; red "路径不全，无法比对"; exit 1; }

# ---------------------------------------------------------------- 1. 清单
# 以 package.json 的 files 字段为准，而不是「源码目录里有什么」——
# test/ 不随包安装，拿它比对会一直报差异。
bold ""
bold "== 1. 运行时文件清单（取自 package.json files） =="
FILES=$(python3 - "$SRC/package.json" <<'PY'
import glob
import json
import os
import sys

pkg = json.load(open(sys.argv[1], encoding="utf-8"))
files = pkg.get("files") or []
out = []
missing = []
for entry in files:
    entry = entry.rstrip("/")
    if any(ch in entry for ch in "*?["):
        # 通配条目（如 py/*.py）。files 之所以用通配而不是裸目录，
        # 是为了把 __pycache__ 排除在包外；这里必须按同样语义展开，
        # 否则整个 py 目录会被当成「找不到」。
        hits = [p for p in glob.glob(entry) if os.path.isfile(p)]
        if hits:
            out.extend(hits)
        else:
            missing.append(entry)
    elif os.path.isfile(entry):
        out.append(entry)
    elif os.path.isdir(entry):
        for root, dirs, names in os.walk(entry):
            dirs[:] = [d for d in dirs if d != "__pycache__"]
            for n in names:
                out.append(os.path.join(root, n))
    else:
        missing.append(entry)
for m in missing:
    sys.stderr.write("MISSING %s\n" % m)
print("\n".join(sorted(set(out))))
PY
)
if [ -z "$FILES" ]; then
  red "  无法从 package.json 解析出文件清单"
  exit 1
fi
# package.json 本身不在 files 字段里（它由 npm 强制包含），但它是插件的
# 核心配置：版本、main、files、dsh.bundle 都靠它。漏比它就会出现「脚本说
# 一致、实际副本还指向旧配置」的情况。补进清单。
case " $FILES " in
  *" package.json "*) ;;
  *) FILES=$(printf '%s\npackage.json' "$FILES") ;;
esac
FILES=$(printf '%s\n' "$FILES" | sort -u)
echo "$FILES" | sed 's/^/    /'

# ---------------------------------------------------------------- 2. 哈希
bold ""
bold "== 2. 源码 vs 安装副本 =="
diffs=()
for f in $FILES; do
  if [ ! -f "$SRC/$f" ]; then
    red   "  缺源文件   $f"
    diffs+=("$f")
    continue
  fi
  if [ ! -f "$DST/$f" ]; then
    wrn   "  副本缺失   $f"
    diffs+=("$f")
    continue
  fi
  a=$(sha256sum "$SRC/$f" | cut -d' ' -f1)
  b=$(sha256sum "$DST/$f" | cut -d' ' -f1)
  if [ "$a" = "$b" ]; then
    echo "  ok   $f  ${a:0:12}"
  else
    red   "  不一致 $f  src=${a:0:12} dst=${b:0:12}"
    diffs+=("$f")
  fi
done

if [ ${#diffs[@]} -gt 0 ]; then
  fail=1
  if [ $FIX -eq 1 ]; then
    bold ""
    bold "== 3. 自动同步（原地写，保住硬链接） =="
    for f in "${diffs[@]}"; do
      if [ ! -f "$SRC/$f" ]; then
        red "  跳过 $f（源文件不存在）"
        continue
      fi
      mkdir -p "$(dirname "$DST/$f")"
      # 用 cat > 重定向而不是 cp：cp 到已存在的文件可能替换 inode，
      # 一旦换掉，node_modules 里的硬链接就指向旧 inode，node 读到的还是旧代码。
      if cat "$SRC/$f" > "$DST/$f"; then
        echo "  已同步 $f"
      else
        red "  同步失败 $f"
      fi
    done
    # 同步后重新比对
    bad=0
    for f in "${diffs[@]}"; do
      [ -f "$SRC/$f" ] || continue
      a=$(sha256sum "$SRC/$f" | cut -d' ' -f1)
      b=$(sha256sum "$DST/$f" 2>/dev/null | cut -d' ' -f1)
      [ "$a" = "$b" ] || { red "  仍不一致 $f"; bad=1; }
    done
    [ $bad -eq 0 ] && grn "  同步完成，全部一致" && fail=0
  else
    bold ""
    wrn "提示：加 --fix 可自动把源码复制到安装副本"
  fi
fi

# ---------------------------------------------------------------- 4. 硬链接
bold ""
bold "== 4. 安装副本 vs node_modules（硬链接） =="
if [ ! -e "$NM" ]; then
  red "  node_modules 里没有插件：$NM"
  fail=1
else
  link_bad=0
  for f in $FILES; do
    [ -f "$DST/$f" ] || continue
    [ -f "$NM/$f" ]  || { wrn "  node_modules 缺 $f"; link_bad=1; continue; }
    i1=$(stat -c '%i' "$DST/$f" 2>/dev/null)
    i2=$(stat -c '%i' "$NM/$f"  2>/dev/null)
    if [ "$i1" = "$i2" ]; then
      echo "  ok   硬链接 $f  inode=$i1"
    else
      red   "  断链   $f  inode $i1 != $i2"
      link_bad=1
    fi
  done
  if [ $link_bad -ne 0 ]; then
    fail=1
    bold ""
    wrn "硬链接已断，node_modules 不会跟随 packages/ 的改动。修复："
    echo "    cd \"$PROFILE\" && pnpm install"
    echo "  （pnpm 会按 file: 依赖重建链接；不要手工往 node_modules 里拷）"
  fi
fi

# ---------------------------------------------------------------- 5. 重启
bold ""
bold "== 5. 是否需要重启 DSH =="
if ! command -v powershell.exe >/dev/null 2>&1; then
  echo "  skip 非 WSL，无法查询宿主进程启动时间"
else
  # 取最新的源码修改时间（只算运行时文件）
  newest=$(for f in $FILES; do
    [ -f "$DST/$f" ] && stat -c '%Y' "$DST/$f"
  done | sort -n | tail -1)

  # 取 DSH 主进程启动时间（最早的那个，因为主进程先起）
  started=$(powershell.exe -NoProfile -Command \
    "(Get-Process -Name 'DeepSeek Harness' -ErrorAction SilentlyContinue | Sort-Object StartTime | Select-Object -First 1).StartTime.ToString('yyyy-MM-ddTHH:mm:ss')" \
    2>/dev/null | tr -d '\r' | head -1)

  if [ -z "$started" ]; then
    echo "  skip 未找到运行中的 DSH 进程"
  else
    start_epoch=$(date -d "$started" +%s 2>/dev/null || echo 0)
    if [ "$start_epoch" -eq 0 ]; then
      echo "  skip 无法解析进程启动时间: $started"
    elif [ "$newest" -gt "$start_epoch" ]; then
      # 「需要重启」与「代码不一致」是两件事，退出码要分开：
      # 前者说明三处代码已经同步好、只差重启；后者说明代码本身不同步。
      # 混成一个码会让 CI 在「代码已同步」时误报失败。
      need_restart=1
      wrn "  需要重启：插件文件在进程启动之后被修改过"
      echo "    进程启动 : $started"
      echo "    最新改动 : $(date -d "@$newest" '+%Y-%m-%dT%H:%M:%S')"
      echo ""
      echo "    原因：ESM 模块在一个进程里只加载一次，patchReload=live 只作用于配置，"
      echo "          不重载插件代码。请完全退出 DSH 再启动。"
    else
      echo "  ok   进程启动于 $started，晚于最新文件改动，无需重启"
    fi
  fi
fi

# ---------------------------------------------------------------- 结论
bold ""
if [ $fail -ne 0 ]; then
  red "代码不一致——此时跑真机验证得到的结论不可信。"
  echo "  修复：./test/sync-check.sh --fix  再 cd \"$PROFILE\" && pnpm install"
  exit 1
fi

if [ $need_restart -ne 0 ]; then
  wrn "三处代码已一致，但运行中的 DSH 加载的是旧版本。"
  echo "  请完全退出 DSH 再启动，之后真机验证才有意义。"
  exit 2
fi

grn "三处一致，且进程加载的是最新代码，可以跑真机验证。"
exit 0
