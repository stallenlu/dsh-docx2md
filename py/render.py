#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""render.py —— 把公式预览图（EMF/WMF）批量转成 PNG，跨平台。

用法:
    python render.py <media_dir> <png_dir> [--max-width 1600]
    python render.py --probe          只探测可用引擎

引擎（按顺序探测，第一个成功者胜出）:
    1. Windows GDI+     System.Drawing 原生支持 WMF/EMF；WSL 下自动调宿主
                        powershell.exe。结果写文件而非 stdout —— Windows 的
                        stdout 编码（GBK/UTF-16 CLIXML）不可靠。
    2. LibreOffice      soffice --headless --convert-to png，Linux/macOS 主力。
    3. Pillow           Linux 上若装了 libwmf 可用（少见）。

为什么必须渲染而不能解 MTEF:
    渲染位图是本流程唯一的"最终权威"。MTEF 字符解码会丢结构——实测
    151 字符的公式只解出 2 个字符（见 METHODOLOGY § 三条路径可靠性）。
    所以渲染失败时必须显式失败，绝不能用字符解码结果冒充公式。
"""
import glob
import json
import os
import platform
import shutil
import subprocess
import sys
from docxfile import force_utf8_stdio

WMF_EXT = (".emf", ".wmf", ".EMF", ".WMF")


def which(names):
    """在 PATH 中找第一个存在的可执行文件。"""
    for n in names:
        p = shutil.which(n)
        if p:
            return p
    return None


def is_wsl():
    if platform.system() != "Linux":
        return False
    try:
        with open("/proc/version", "r", encoding="utf-8", errors="replace") as f:
            return "microsoft" in f.read().lower()
    except OSError:
        return False


# PowerShell 侧脚本：把结果写到 JSON 文件，stdout 只留一个 done 标记。
# 参数以字面量拼进脚本正文（-EncodedCommand 不接受额外的位置参数）。
PS_TEMPLATE = r"""
$Src    = '@@SRC@@'
$Dst    = '@@DST@@'
$Report = '@@REPORT@@'
$MaxW   = @@MAXW@@
$ErrorActionPreference = 'Stop'
try {
    Add-Type -AssemblyName System.Drawing
    $rows = New-Object System.Collections.ArrayList
    $files = @(Get-ChildItem -LiteralPath $Src -File | Where-Object {
        $_.Extension -eq '.emf' -or $_.Extension -eq '.wmf'
    })
    foreach ($f in $files) {
        $outName = $f.BaseName + '.png'
        $outPath = Join-Path $Dst $outName
        try {
            $img = [System.Drawing.Image]::FromFile($f.FullName)
            $w = [int]$img.Width; $h = [int]$img.Height
            $nw = $w; $nh = $h
            if ($w -gt $MaxW) {
                $nw = $MaxW
                $nh = [int][Math]::Round(($h * $MaxW) / $w)
            }
            if ($nw -lt 1) { $nw = 1 }
            if ($nh -lt 1) { $nh = 1 }
            $bmp = New-Object System.Drawing.Bitmap($nw, $nh)
            $g = [System.Drawing.Graphics]::FromImage($bmp)
            $g.Clear([System.Drawing.Color]::White)
            $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $g.DrawImage($img, 0, 0, $nw, $nh)
            $bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
            $g.Dispose(); $bmp.Dispose(); $img.Dispose()
            [void]$rows.Add([pscustomobject]@{ src = $f.Name; png = $outName; w = $nw; h = $nh; ok = $true })
        } catch {
            [void]$rows.Add([pscustomobject]@{ src = $f.Name; png = $null; ok = $false })
        }
    }
    $json = ConvertTo-Json -InputObject @($rows) -Depth 4 -Compress
    [System.IO.File]::WriteAllText($Report, $json, (New-Object System.Text.UTF8Encoding($false)))
    Write-Output ('done ' + $rows.Count)
} catch {
    $msg = 'fatal: ' + $_.Exception.Message
    [System.IO.File]::WriteAllText($Report, ('{"error":"' + $msg.Replace('"','') + '"}'),
        (New-Object System.Text.UTF8Encoding($false)))
    Write-Output $msg
}
"""


def to_win_path(p):
    """Linux 路径 → Windows 可访问路径（WSL 下用 UNC，/mnt/X 用盘符）。"""
    p = os.path.abspath(p)
    if sys.platform == "win32":
        return p
    if p.startswith("/mnt/") and len(p) > 6 and p[6:7] == "/":
        return p[5].upper() + ":" + p[6:].replace("/", "\\")
    host = os.environ.get("WSL_DISTRO_NAME") or platform.node()
    return "\\\\wsl.localhost\\" + host + p.replace("/", "\\")


def render_windows(media_dir, png_dir, max_width):
    """Windows GDI+（原生 Windows 或 WSL 调宿主）。"""
    powershell = which(["powershell.exe", "powershell", "pwsh"])
    if not powershell:
        return None

    # 报告文件放在 png_dir 内（已确认可写），避免依赖系统临时目录。
    report = os.path.join(png_dir, "_render_report.json")
    if os.path.exists(report):
        os.remove(report)

    script = (PS_TEMPLATE
              .replace("@@SRC@@", to_win_path(media_dir).replace("'", "''"))
              .replace("@@DST@@", to_win_path(png_dir).replace("'", "''"))
              .replace("@@REPORT@@", to_win_path(report).replace("'", "''"))
              .replace("@@MAXW@@", str(int(max_width))))
    import base64
    b64 = base64.b64encode(script.encode("utf-16-le")).decode("ascii")

    cmd = [powershell, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
           "-EncodedCommand", b64]

    proc = subprocess.run(cmd, capture_output=True, timeout=1800)
    out = (proc.stdout or b"").decode("utf-8", "replace").strip()
    err = (proc.stderr or b"").decode("utf-8", "replace").strip()

    if not os.path.exists(report):
        detail = (err or out)[:600]
        # 去掉 CLIXML 噪音，只留有信息量的部分
        if "#< CLIXML" in detail:
            detail = "PowerShell 返回 CLIXML 而未产出报告"
        return {"ok": False, "engine": "windows-gdi+", "reason": detail}

    with open(report, "r", encoding="utf-8-sig") as f:
        rows = json.load(f)
    # 报告只是渲染过程的中间产物，读进内存后立刻删掉，别把它留在 png/ 里。
    try:
        os.remove(report)
    except OSError:
        pass
    if isinstance(rows, dict):
        rows = [rows]
    good = [r for r in rows if r.get("ok")]
    return {
        "ok": len(good) > 0,
        "engine": "windows-gdi+",
        "total": len(rows),
        "rendered": len(good),
        "failed": [r["src"] for r in rows if not r.get("ok")][:10],
    }


def render_libreoffice(media_dir, png_dir, max_width):
    soffice = which(["soffice", "libreoffice"])
    if not soffice:
        return None
    srcs = sorted([p for p in glob.glob(os.path.join(media_dir, "*"))
                   if os.path.splitext(p)[1] in WMF_EXT])
    if not srcs:
        return None
    ok, errs = 0, []
    for s in srcs:
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "png", "--outdir", png_dir, s],
                capture_output=True, timeout=180)
            produced = os.path.splitext(os.path.basename(s))[0] + ".png"
            if os.path.exists(os.path.join(png_dir, produced)):
                ok += 1
            else:
                errs.append(os.path.basename(s))
        except subprocess.TimeoutExpired:
            errs.append(os.path.basename(s) + " (timeout)")
    return {"ok": ok > 0, "engine": "libreoffice", "total": len(srcs),
            "rendered": ok, "failed": errs[:10]}


def render_pillow(media_dir, png_dir, max_width):
    try:
        from PIL import Image
    except ImportError:
        return None
    srcs = sorted([p for p in glob.glob(os.path.join(media_dir, "*"))
                   if os.path.splitext(p)[1] in WMF_EXT])
    if not srcs:
        return None
    ok, errs = 0, []
    for s in srcs:
        try:
            with Image.open(s) as im:
                im.load()
                out = os.path.splitext(os.path.basename(s))[0] + ".png"
                im.convert("RGBA").save(os.path.join(png_dir, out))
            ok += 1
        except Exception:
            errs.append(os.path.basename(s))
    return {"ok": ok > 0, "engine": "pillow", "total": len(srcs),
            "rendered": ok, "failed": errs[:10]}


ENGINES = [render_windows, render_libreoffice, render_pillow]


def do_probe():
    out = {
        "platform": platform.system(),
        "is_wsl": is_wsl(),
        "powershell": which(["powershell.exe", "powershell", "pwsh"]),
        "libreoffice": which(["soffice", "libreoffice"]),
        "uvx": which(["uvx"]),
        "python": sys.executable,
    }
    try:
        from PIL import Image  # noqa: F401
        out["pillow"] = True
    except ImportError:
        out["pillow"] = False

    # 只报告探测到哪些引擎，不排序、不推荐——选哪个取决于用户环境与偏好。
    detected = []
    if out["powershell"]:
        detected.append("windows-gdi+")
    if out["libreoffice"]:
        detected.append("libreoffice")
    if out["pillow"]:
        detected.append("pillow")
    out["engines_available"] = detected
    out["engine_available"] = bool(detected)
    print(json.dumps(out, ensure_ascii=False, indent=1))


def main():
    if "--probe" in sys.argv:
        do_probe()
        return

    positional = []
    i = 1
    while i < len(sys.argv):
        a = sys.argv[i]
        if a == "--max-width":
            i += 2
            continue
        if not a.startswith("--"):
            positional.append(a)
        i += 1

    if len(positional) < 2:
        print(__doc__)
        sys.exit(2)
    media_dir, png_dir = positional[0], positional[1]
    max_width = 1600
    if "--max-width" in sys.argv:
        max_width = int(sys.argv[sys.argv.index("--max-width") + 1])

    os.makedirs(png_dir, exist_ok=True)

    # 已经是位图的预览图（PNG/JPEG）不需要渲染，直接拷进 png/，
    # 这样下游"读 png/ 目录里所有文件"就是一条统一指令。
    copied = []
    for p in glob.glob(os.path.join(media_dir, "*")):
        ext = os.path.splitext(p)[1].lower()
        if ext in (".png", ".jpg", ".jpeg", ".gif"):
            dst = os.path.join(png_dir, os.path.basename(p))
            try:
                shutil.copyfile(p, dst)
                copied.append(os.path.basename(p))
            except OSError:
                pass

    srcs = [p for p in glob.glob(os.path.join(media_dir, "*"))
            if os.path.splitext(p)[1] in WMF_EXT]
    if not srcs:
        print(json.dumps({"ok": True, "rendered": 0, "copied": len(copied),
                          "note": "no EMF/WMF in media/, nothing to render"},
                         ensure_ascii=False))
        return

    attempts = []
    for engine in ENGINES:
        try:
            res = engine(media_dir, png_dir, max_width)
        except Exception as exc:
            res = {"ok": False, "engine": engine.__name__, "reason": str(exc)[:400]}
        if res is None:
            continue
        attempts.append(res)
        if res.get("ok"):
            out = dict(res)
            out["attempts"] = [dict(a) for a in attempts]
            out["copied"] = len(copied)
            print(json.dumps(out, ensure_ascii=False, indent=1))
            return

    print(json.dumps({
        "ok": False,
        "rendered": 0,
        "attempts": attempts,
        "note": (
            "No available engine could convert the vector previews. The original "
            "EMF/WMF files remain in the media directory."
        ),
    }, ensure_ascii=False, indent=1))
    sys.exit(1)


if __name__ == "__main__":
    force_utf8_stdio()
    main()
