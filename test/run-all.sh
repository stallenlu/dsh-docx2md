#!/usr/bin/env bash
# run-all.sh —— 一次跑完全部测试。
#
# 用法:
#   ./test/run-all.sh [path/to/file.docx]          日常：功能与契约测试
#   ./test/run-all.sh --with-pack [file.docx]      发布前：额外检查 tarball 内容
#
# 为什么打包检查默认不跑：它要调 npm pack（网络无关但较慢），且只在发布前有意义。
# 退出码 0 = 全绿。
set -u
cd "$(dirname "$0")/.." || exit 1
WITH_PACK=0
if [ "${1:-}" = "--with-pack" ]; then
  WITH_PACK=1
  shift
fi
DOC="${1:-}"
fail=0
step() { printf '\n\033[1m== %s ==\033[0m\n' "$1"; }
note() { if [ $1 -ne 0 ]; then fail=$((fail+1)); printf '  >>> FAILED\n'; else printf '  >>> ok\n'; fi; }

step "1. Python 语法"
for f in py/*.py; do
  python3 -c "import ast,sys; ast.parse(open('$f',encoding='utf-8').read())" || { echo "  BAD $f"; fail=$((fail+1)); }
done
echo "  全部通过"; note 0

step "2. Node 语法"
for f in lib/*.js test/*.mjs; do
  node --check "$f" || fail=$((fail+1))
done
echo "  全部通过"
note 0; note 0

step "3. Python 模块名不与标准库冲突"
python3 - <<'PY'
import os, sys
bad = [f for f in os.listdir('py') if f.endswith('.py') and f[:-3] in sys.stdlib_module_names]
if bad:
    print("  冲突: " + ", ".join(bad)); sys.exit(1)
print("  无冲突")
PY
note $?

step "4. 同目录直跑（插件的真实调用方式）"
( cd py && python3 docx_scan.py >/dev/null 2>&1; [ $? -eq 2 ] ) && echo "  docx_scan.py 可启动" || { echo "  docx_scan.py 启动失败"; fail=$((fail+1)); }
( cd py && python3 validate.py >/dev/null 2>&1; [ $? -eq 2 ] ) && echo "  validate.py 可启动" || { echo "  validate.py 启动失败"; fail=$((fail+1)); }
note 0

step "5. tool 契约（按 dsh-tools 真实校验复刻）"
node test/tool-contract.mjs | tail -3
note $?

step "6. skill 契约（按 dsh-skill 真实校验复刻）"
node test/skill-contract.mjs | tail -3
note $?

step "7. 校验器自测（T0–T11 每条规则都有反例）"
python3 test/validate_selftest.py | tail -3
note $?

step "7b. 公式↔预览图配对（ShapeID 优先，回退必告警）"
python3 test/pairing_selftest.py | tail -2
note $?

step "7c. 公式可显示性（定界/边界/长度，覆盖「显示成源码」那类反馈）"
if [ -n "$DOC" ] && [ -f "$DOC" ] && [ -f /tmp/latex_by_stream.json ]; then
  python3 test/display_selftest.py "$DOC" /tmp/latex_by_stream.json | tail -2
else
  python3 test/display_selftest.py | tail -2
fi
note $?

step "7d. OMML→LaTeX 转换器（Word 原生公式，22 类结构）"
python3 test/omml_selftest.py | tail -2
note $?

step "7e. OMML 端到端（合成 docx：inspect→extract→build，空 latex 表也要出公式）"
python3 test/omml_e2e.py | tail -2
note $?

step "7f. 图片公式端到端（占位不静默丢 / 可见引用 / 转写替换 / T11 断链）"
python3 test/image_formula_e2e.py | tail -2
note $?

step "7g. 画布内公式（w:pict 装公式，不能被当成预览图整棵跳过）"
python3 test/pict_ole_e2e.py | tail -2
note $?

step "7h. 标题结构信号（样式级别 / Word 自动编号 / 整段加粗）"
python3 test/docstruct_e2e.py | tail -2
note $?

step "7i. 通用 docx 结构（表格/列表/行内格式/超链接/脚注/合并单元格）"
python3 test/generic_struct_e2e.py | tail -2
note $?

step "7j. 流程图画布（v:group 整块渲染、画布内公式不重复占位）"
python3 test/canvas_e2e.py | tail -2
note $?

step "7k. 三项排版信号（居中 / 颜色 / 首行缩进）"
python3 test/parafmt_e2e.py | tail -2
note $?

step "7l. 探测器通用性（T13/T14 阈值随文档缩放自标定）"
python3 test/detector_portability.py | tail -2
note $?

step "8. 边界输入"
tmp=$(mktemp -d)
printf 'not a zip' > "$tmp/fake.docx"
python3 -c "import zipfile;zipfile.ZipFile('$tmp/empty.docx','w').close()"
python3 -c "import zipfile;z=zipfile.ZipFile('$tmp/nodoc.docx','w');z.writestr('a.txt','x');z.close()"
ok=1
for f in fake empty nodoc missing; do
  ( cd py && python3 docx_scan.py "$tmp/$f.docx" >/dev/null 2>&1 )
  [ $? -eq 3 ] || { echo "  $f.docx 未返回退出码 3"; ok=0; }
done
rm -rf "$tmp"
[ $ok -eq 1 ] && echo "  四种坏输入均返回退出码 3"
note $((ok==1?0:1))

if [ -n "$DOC" ] && [ -f "$DOC" ]; then
  step "9. 端到端（$DOC）"
  node test/plugin.mjs "$DOC" | tail -8
  note $?
  step "10. 独立路径 smoke"
  node test/smoke.mjs "$DOC" | tail -6
  note $?
else
  step "9-10. 端到端（跳过：未提供 docx 路径）"
fi

step "11. 生命周期闭环（安装→使用→卸载→重装）"
./test/lifecycle.sh 2>&1 | tail -2
note $?

step "12. 跨平台路径语义（/tmp vs C:\tmp、/mnt/c、UNC 反查）"
node test/paths.mjs 2>&1 | tail -2
note $?

step "13. 临时目录与残留（不建系统临时目录、坏输入不留空壳）"
node test/hygiene.mjs $DOC 2>&1 | tail -2
note $?

step "14. 中文编码（用 PYTHONIOENCODING=gbk 伪造中文 Windows 环境）"
node test/encoding.mjs $DOC 2>&1 | tail -3
note $?

step "15. 三处代码一致性（源码 / packages / node_modules）"
bash test/sync-check.sh > /tmp/_synccheck.log 2>&1
sc=$?
if [ $sc -eq 0 ]; then
  echo "  三处一致，且进程已加载最新代码"
  note 0
elif [ $sc -eq 2 ]; then
  # 三处已一致，只是宿主进程还是旧版本——这不算测试失败，但必须提醒
  echo "  三处一致；但运行中的 DSH 仍是旧代码，需重启后才算生效"
  note 0
else
  echo "  代码不同步（下方为详情），此时任何真机结论都不可信"
  tail -6 /tmp/_synccheck.log | sed 's/^/    /'
  note 1
fi

step "16. 更新路径（改代码 / 改版本 / 断链 / 重装）"
./test/update.sh 2>&1 | tail -2
note $?

if [ $WITH_PACK -eq 1 ]; then
  step "17. 发布前：tarball 内容检查"
  node test/pack.mjs 2>&1 | tail -3
  note $?

  step "18. 发布前：从 tarball 真实安装（GitHub 路径）"
  . "$(dirname "$0")/lib/discover.sh"
  WIN_NODE="$(find_win_node 2>/dev/null || true)"
  if [ -n "$WIN_NODE" ] && [ -n "${WIN_TMP:-}" ]; then
    RUNDIR="${TMPBASE%/}/d2m_run"
    rm -rf "$RUNDIR" && mkdir -p "$RUNDIR"
    for f in lib py skills test package.json README.md LICENSE cordis.patch.yml .gitignore; do
      [ -e "$f" ] && cp -r "$f" "$RUNDIR/" 2>/dev/null
    done
    find "$RUNDIR" -name "__pycache__" -type d -exec rm -rf {} + 2>/dev/null
    WINDOC=""
    if [ -n "$DOC" ] && [ -f "$DOC" ]; then
      cp "$DOC" "$RUNDIR/demo.docx" && WINDOC="$(to_win_path "$RUNDIR")\\demo.docx"
    fi
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "
      Set-Location '$(to_win_path "$RUNDIR")'
      \$env:DOCX2MD_INSTALL_BASE = '$(to_win_path "$TMPBASE")'
      \$env:TEMP = '$(to_win_path "$TMPBASE")'
      & '$(to_win_path "$WIN_NODE")' 'test\github-install.mjs' '$WINDOC'
    " 2>&1 | tr -d '\r' | tail -4
    note ${PIPESTATUS[0]}
  else
    echo "  skip 未找到 Windows node（该测试需要它）"
    note 0
  fi
fi

printf '\n'
if [ $fail -eq 0 ]; then
  printf '\033[32m全部通过\033[0m\n'; exit 0
else
  printf '\033[31m%d 组失败\033[0m\n' "$fail"; exit 1
fi
