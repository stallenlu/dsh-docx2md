#!/usr/bin/env bash
# discover.sh —— 测试共用的环境发现。被其它测试 source，不单独运行。
#
# 为什么要有这个文件：
#   仓库要公开，但测试里原先写死了 C:/Users/GT14/... 这类路径。别人 clone 下来
#   就跑不起来。这些路径其实分两类，处理方式完全不同：
#
#   1. 「一个临时目录」——没有任何理由绑定到某个人。用 mktemp 即可。
#   2. 「本机 DSH 的安装位置」——它**本质上**是本机特有的，不可能可移植。
#      对这类路径，正确的做法不是硬编码，也不是假装可移植，而是：
#          显式覆盖（环境变量） > 自动发现 > 明确跳过
#      跳过必须是「成功跳过」而不是「失败」——否则 CI 里会出现假红。
#
# 于是约定：
#   $TMPBASE     一个可写的临时根目录（自动）
#   $DSH_PROFILE 本机 desktop profile 目录；找不到时为空
#   $DSH_PACKAGE 本机已安装的包副本目录；找不到时为空
#   $WIN_TMP     Windows 风格的临时路径（用于传给 Windows 程序）；非 WSL 时为空
#
# 用法：
#   . "$(dirname "$0")/lib/discover.sh"

set -u

# ---------------------------------------------------------------- 平台
case "$(uname -s)" in
  Linux)
    if grep -qiE 'microsoft|wsl' /proc/version 2>/dev/null; then
      IS_WSL=1
    else
      IS_WSL=0
    fi
    ;;
  Darwin) IS_WSL=0 ;;
  *)      IS_WSL=0 ;;
esac

# ---------------------------------------------------------------- 临时根目录
# 关键约束：在 WSL 上，临时目录必须**两边都看得见**。
#   /tmp                        → Windows 侧是 \\wsl.localhost\...\tmp，cmd.exe
#                                 在 UNC 工作目录下会退化到 Windows 目录，
#                                 于是 pnpm 收到 \tmp\x 这种无效路径。
#   C:\Users\x\AppData\Local\Temp → /mnt/c/Users/x/AppData/Local/Temp，两边都真
#                                 实存在，pnpm 与 python 都能用。
# 所以 WSL 上优先选后者；Linux/macOS 原生就用 /tmp。
WIN_TMP=""
if [ "$IS_WSL" -eq 1 ] && command -v powershell.exe >/dev/null 2>&1; then
  WIN_TMP=$(powershell.exe -NoProfile -Command \
    "[System.IO.Path]::GetTempPath()" 2>/dev/null | tr -d '\r\n' | sed 's/[[:space:]]*$//')
  WIN_TMP="${WIN_TMP%\\}"
fi

TMPBASE="${TMPDIR:-}"
if [ -z "$TMPBASE" ]; then
  if [ -n "$WIN_TMP" ]; then
    # C:\Users\x\AppData\Local\Temp → /mnt/c/Users/x/AppData/Local/Temp
    TMPBASE=$(printf '%s' "$WIN_TMP" | sed -E 's#^([A-Za-z]):\\#/mnt/\L\1/#' | sed 's#\\#/#g')
    [ -d "$TMPBASE" ] || TMPBASE="/mnt/c/Users/$(whoami)/AppData/Local/Temp"
  else
    TMPBASE="/tmp"
  fi
fi
TMPBASE="${TMPBASE%/}"
[ -d "$TMPBASE" ] || TMPBASE=/tmp

# ---------------------------------------------------------------- DSH 安装位置
# 自动发现顺序：环境变量 → Windows 用户目录 → WSL 侧 ~/.dsh
# 注意：WSL 里的 ~/.dsh 与 Windows 的 %USERPROFILE%\.dsh 是**两个不同目录**，
# 桌面版 DSH 读的是后者。所以优先探测 Windows 那一份，最后才退回 ~/.dsh。
find_dsh_profile() {
  # 1. 显式覆盖
  if [ -n "${DOCX2MD_PROFILE:-}" ] && [ -d "${DOCX2MD_PROFILE}" ]; then
    printf '%s\n' "${DOCX2MD_PROFILE%/}"
    return 0
  fi

  # 2. Windows 用户目录下的 .dsh（桌面版的真实位置）
  local winhome=""
  if [ -n "${WIN_TMP:-}" ]; then
    # C:\Users\x\AppData\Local\Temp → C:\Users\x
    winhome=$(printf '%s' "$WIN_TMP" | sed -E 's#\\AppData\\Local\\Temp$##')
  elif command -v powershell.exe >/dev/null 2>&1; then
    winhome=$(powershell.exe -NoProfile -Command '$env:USERPROFILE' 2>/dev/null \
      | tr -d '\r\n' | sed 's/[[:space:]]*$//')
  fi
  if [ -n "$winhome" ]; then
    local cand
    cand=$(printf '%s' "$winhome" \
      | sed -E 's#^([A-Za-z]):\\#/mnt/\L\1/#' \
      | sed 's#\\#/#g')
    if [ -d "$cand/.dsh/profiles/desktop" ]; then
      printf '%s\n' "$cand/.dsh/profiles/desktop"
      return 0
    fi
  fi

  # 3. 当前用户 home 下的 .dsh（Linux 原生安装）
  if [ -d "$HOME/.dsh/profiles/desktop" ]; then
    printf '%s\n' "$HOME/.dsh/profiles/desktop"
    return 0
  fi

  return 1
}

find_dsh_package() {
  if [ -n "${DOCX2MD_PACKAGE:-}" ] && [ -d "${DOCX2MD_PACKAGE}" ]; then
    printf '%s\n' "${DOCX2MD_PACKAGE%/}"
    return 0
  fi
  local prof
  prof=$(find_dsh_profile) || return 1
  # profile 形如 <root>/.dsh/profiles/desktop；包副本在 <root>/.dsh/packages/
  local root="${prof%/profiles/*}"
  if [ -d "$root/packages/dsh-docx2md" ]; then
    printf '%s\n' "$root/packages/dsh-docx2md"
    return 0
  fi
  return 1
}

DSH_PROFILE="$(find_dsh_profile 2>/dev/null || true)"
DSH_PACKAGE="$(find_dsh_package 2>/dev/null || true)"

# ---------------------------------------------------------------- 便捷函数

# Windows 侧 node（跑需要原生路径的测试用）；没有则为空
find_win_node() {
  for p in "/mnt/c/Program Files/nodejs/node.exe" "C:/Program Files/nodejs/node.exe"; do
    [ -x "$p" ] && { printf '%s\n' "$p"; return 0; }
  done
  command -v node.exe >/dev/null 2>&1 && { command -v node.exe; return 0; }
  return 1
}

# WSL 路径 → Windows 路径（只处理 /mnt/<drive>/... 形式）；非 WSL 原样返回
to_win_path() {
  local p="$1"
  if [ -z "${WIN_TMP:-}" ]; then
    printf '%s\n' "$p"
    return 0
  fi
  printf '%s' "$p" | sed -E 's#^/mnt/([a-zA-Z])/#\U\1:/#' | sed 's#/#\\#g'
}

# ---------------------------------------------------------------- Windows 程序调用
# 从 WSL 调 Windows 的 .cmd 有一串连带约束，全部实测过：
#   1. node 不能直接 spawn .cmd（EINVAL/ENOENT），必须经 cmd.exe /c
#   2. cmd.exe 不支持 UNC 作为工作目录，会退化到 Windows 目录，导致
#      pnpm 在错的地方找 package.json —— 所以要显式 `cd /d <dir>`
#   3. .cmd 的路径本身也必须是 Windows 形式；传 /mnt/c/... 会「系统找不到
#      指定的路径」。所以要把命令路径也过一遍 to_win_path
# win_run 把这些一次封装好，调用方只管给目录和参数。
win_run() {
  local dir_win="$1"; shift
  if [ "$IS_WSL" -eq 1 ]; then
    cmd.exe /c "cd /d $dir_win && $*" 2>/dev/null
  else
    ( cd "$dir_win" && "$@" ) 2>/dev/null
  fi
}

# 找一个可直接在 cmd.exe 里用的 pnpm 命令路径（Windows 形式）
win_pnpm() {
  local p
  p="$(command -v pnpm 2>/dev/null || true)"
  if [ -n "$p" ]; then
    to_win_path "$p"
    return 0
  fi
  [ -n "${WIN_TMP:-}" ] || return 1
  # %LOCALAPPDATA%\pnpm\bin\pnpm.cmd，从 WIN_TMP 反推用户目录
  local home_win
  home_win=$(printf '%s' "$WIN_TMP" | sed -E 's#\\AppData\\Local\\Temp$##')
  printf '%s\\AppData\\Local\\pnpm\\bin\\pnpm.cmd\n' "$home_win"
}

export IS_WSL TMPBASE WIN_TMP DSH_PROFILE DSH_PACKAGE
