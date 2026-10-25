# DOCX → Markdown 科学文档转换方法论

> 适用范围：以 **OLE 公式对象**（MathType / Equation Editor）嵌入的 Word 文档，
> 典型来源为 WPS 或旧版 Word（.doc / .docx），含大量带上下标、分式、矩阵的数学公式。
>
> 本文档由 `基于矢量位错密度的位错动力学及其数理方案(VDDD3D)` 的实际转换经验总结而成
> （145 个公式、308 段正文、约 45,000 字符）。
>
> **另有一族公式不在本文讨论范围内**：Word 原生公式（OMML，`m:oMath` 结构化 XML，
> Word 2007+ 的默认格式）。它不是私有二进制，结构完整可解析，插件用 `py/omml.py`
> 在 extract 阶段**确定性**转成 LaTeX，不需要读图——与本文反复强调的「以位图为准」
> 是两条路。「以位图为准」只适用于 OLE 这一族，别把它当成所有公式的通用军规。

> **本文档与实现的关系（重要）**
>
> 这份方法论来自那次转换的**实验记录**，描述了三条并行的公式来源。其中：
>
> | 路径 | 本文档 | 插件实现 |
> |---|---|---|
> | ① 渲染 WMF/EMF 后读图 | 详述 | ✅ `py/render.py` + `py/extract.py` |
> | ② EMF 内嵌 PDF 文字层（pypdf） | 详述 | ⚠️ **未实现** |
> | ③ MTEF 直接解码 | 详述 | ⚠️ 仅 `docx_scan.py` 做变体判定，不产出字符 |
>
> 路径②是那次实验里手工做的一步，没有固化成脚本。`pypdf` 目前只在
> `lib/index.js` 的依赖声明里出现（为将来实现预留），`py/` 下零引用。
> 因此插件当前对**变体 B**（15/145 个公式）只能靠路径①读图，
> 读不清时 `summary.formulas_without_media` 会明确报出数量，不会假装成功。
>
> 补上路径②是可做的工作，但需要先解决 EMF 内的 GDICOMMENT/PDF 流提取
> 与文本层对齐——本文档 §2–§3 记录了这一思路。若只需要跑通转换，
> 路径①足够；这条注记是为了避免把「实验里做到过」误当成「插件已具备」。

---

## 1. 核心结论（先看这一节）

### 1.1 三条路径的可靠性排序

| 路径 | 可得内容 | 可靠性 | 适用 |
|---|---|---|---|
| **① WMF/EMF 图像渲染 + 读图转写** | 完整排版（上下标/分式/矩阵全在） | ★★★★★ | **最终依据** |
| ② EMF 内嵌 PDF 文本层 | 字符 + 部分结构 | ★★★★ | 矩阵类公式 |
| ③ MTEF 记录流解码 | 仅字符序列 | ★★ | 快速清点、交叉校验 |

**最重要的经验：不要相信字符级提取，必须看渲染图。**

实测反差（同一个公式）：

| 公式 | MTEF 解码结果 | 真实内容 |
|---|---|---|
| #6 | `α=-∇×βP`（7 字符） | **5 段等式链**，含 `∫`、`∂S`、`δ(∂S)b` |
| #27 | `βP`（2 字符） | **6 行推导**，共 151 个字符 |
| #30 | `u=dr`（4 字符） | **5 行推导** |

字符数对比（WMF 实际字符 vs MTEF 解出）：

```
公式 #27:  WMF 151 字  →  MTEF  2 字   丢失 149
公式 #30:  WMF 150 字  →  MTEF 40 字   丢失 110
公式 #33:  WMF 149 字  →  MTEF 23 字   丢失 126
```

### 1.2 环境要求

| 能力 | 用途 | 检查命令 |
|---|---|---|
| `olefile` | 解析 OLE2 复合文档 | `uvx --from olefile python -c "import olefile"` |
| `pypdf` | 抽取内嵌 PDF 文本 | `uvx --from pypdf python -c "import pypdf"` |
| **Windows GDI+**（WSL 下经 `powershell.exe` 调用） | **渲染 WMF/EMF** | 见 §4 |
| `matplotlib`（可选） | 重绘示意图 | `uvx --from matplotlib python -c "import matplotlib"` |

> **关键**：WSL 环境下可直接调用 `/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe`，
> 用 .NET 的 `System.Drawing` 原生渲染 WMF/EMF——**无需安装 LibreOffice/ImageMagick**。

---

## 2. 文档结构解剖

### 2.1 .docx 的 zip 结构

```
word/
├── document.xml          ← 正文，含 <w:object> 公式占位
├── _rels/document.xml.rels  ← rId ↔ 文件路径映射
├── embeddings/           ← 145 个 oleObjectN.bin（公式 OLE 包）
└── media/                ← 公式预览图（wmf/emf/png）
```

### 2.2 三层嵌套

```
docx (zip)
 └── word/embeddings/oleObjectN.bin      ← OLE2 复合文档（魔数 d0cf11e0）
      └── "Equation Native" 流            ← 真正的公式数据
           ├── 前 28 字节：EQ 头
           └── 之后：MTEF 记录流
```

### 2.3 两种 MTEF 变体（易踩坑）

| 变体 | 特征 | 数量 | 可解析性 |
|---|---|---|---|
| **A** | body 以 `03 01 01 03`（TMPL）开头 | 130 | 可解出字符 |
| **B** | body 以 `05 00 00 06 07 "DSMT6"` 字体表开头 | 15 | **MTEF 不可解**，须走 EMF→PDF |

变体 B 的判断代码：

```python
if body[:5] == b"\x05\x00\x00\x06\x07":
    # 变体 B：改用 EMF 内嵌 PDF 路径
```

---

## 3. 推荐流程

### 步骤 0：环境探测

```bash
# 检查 Windows PowerShell 是否可用（WSL 关键一步）
PS="/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"
[ -x "$PS" ] && echo "GDI+ 渲染可用" || echo "需另寻方案"
```

### 步骤 1：提取正文与公式位置

用 `xml.etree` 解析 `document.xml`，把 `<w:object>` 替换为占位标记 `⟦EQ:oleObjectN.bin⟧`：

```python
def para_text(p):
    out = []
    for node in p.iter():
        if node.tag == W + "t":
            out.append(node.text or "")
        elif node.tag == O + "OLEObject":
            rid = node.get(R + "id")
            out.append("⟦EQ:%s⟧" % rel_map[rid].split("/")[-1])
    return "".join(out)
```

**产出**：`paras.json`（段落文本 + 公式占位），这是后续所有处理的骨架。

### 步骤 2：三路径并行提取公式

```
                  ┌─→ MTEF 解码 ──→ 字符序列（快，结构差）
oleObjectN.bin ───┤
                  └─→ 字体/变体判定
                           │
                           └─→ 变体B: EMF → GDICOMMENT → %PDF → pypdf → 文本
```

### 步骤 3：渲染全部预览图为 PNG（**最关键**）

```powershell
Add-Type -AssemblyName System.Drawing
$files = Get-ChildItem -Path $src -File | Where-Object { $_.Extension -eq ".wmf" -or $_.Extension -eq ".emf" }
foreach ($f in $files) {
    $img = [System.Drawing.Image]::FromFile($f.FullName)
    $scale = [Math]::Min(1.0, 1400 / [Math]::Max(1, $img.Width))
    $w = [int]($img.Width * $scale); $h = [int]($img.Height * $scale)
    $bmp = New-Object System.Drawing.Bitmap($w), ($h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::White)
    $g.DrawImage($img, 0, 0, $w, $h)
    $bmp.Save((Join-Path $dst ($f.BaseName + ".png")), [System.Drawing.Imaging.ImageFormat]::Png)
}
```

**踩坑记录**：
- PowerShell 脚本**必须纯 ASCII**（中文会导致 `ParserError`）
- `Get-ChildItem -Include` 不加 `-Recurse` 无效，改用 `Where-Object`
- WSL 路径需 `wslpath -w` 转成 `\\wsl.localhost\...` 形式

### 步骤 4：逐图读图转写 LaTeX

这是唯一可靠的环节，需人工（或视觉模型）逐个读图。建议：

1. **分批处理**（如每批 10–30 个），每批后核对，避免一路错到底
2. 转写结果存 JSON（`{编号: latex}`），便于回填与修订
3. 长公式（含 `aligned`/`cases`）单独标记，生成时必须用 `$$` 显示块

### 步骤 5：生成 Markdown

回填公式 + 加编号：

```python
def render(p):
    def rep(m):
        name = m.group(1)
        n = EQ[name]
        tex = LATEX[n]
        # 多行公式必须用显示块，不能内联！
        if MULTILINE.search(tex):
            return "\n\n$$\n%s\n$$\n" % tex
        return "$%s$" % tex
    return re.sub(r"⟦EQ:([^⟧]+)⟧", rep, p)
```

### 步骤 6：排版规范化

| 项 | 处理 |
|---|---|
| 标题层级 | `1．` → `## 1`；`A）` → `### 2.1`；`i)` → `#### 2.1.1` |
| 公式编号 | 用 `\tag{n}`，置于公式内容末尾 |
| 多行公式 | `\end{aligned}\tag{n}` 放同一行 |
| 图注 | `**图 N**　说明。(a) …；(b) …` |
| 列表 | 连续并列项改有序列表 |
| 附录 | 转换说明等元信息移至附录 |

---

## 4. Markdown + LaTeX 的六个必须遵守的规则

以下每一条都是实际踩过的坑，违反会导致**渲染错乱**：

### 规则 1：`$$` 必须独立成行

```markdown
❌ 错：  $$\begin{aligned}...\end{aligned}$$
✅ 对：  $$
         \begin{aligned}...\end{aligned}
         $$
```

**后果**：`$$` 与内容同行时多数渲染器不识别为显示公式，会与后续 `$$` 错配，
从该点起**整段崩坏**（本次转换中曾导致「从这里到文件结尾都显示不正确」）。

### 规则 2：多行公式不能内联

```markdown
❌ 错：  $\begin{aligned} a&=b \\ &=c \end{aligned}$
✅ 对：  $$
         \begin{aligned} a&=b \\ &=c \end{aligned}
         $$
```

**后果**：内联模式不解析 `\\` 和 `&`，这些字符原样显示 → 看起来像乱码。

### 规则 3：相邻公式之间必须有分隔

```markdown
❌ 错：  $\sigma_{xx}$$\sigma_{zz}$$\sigma_{xy}$
✅ 对：  各自独立成 $$ 块，或中间留文字
```

**后果**：`$A$$B$` 中间的 `$$` 被解析为显示公式定界符，导致后续全部错位。

### 规则 4：`\tag` 的位置

```markdown
❌ 错：  \begin{aligned}...\end{aligned}
         \tag{8}          ← 换行后
✅ 对：  \begin{aligned}...\end{aligned}\tag{8}
```

### 规则 5：表格单元格内的公式

含 `|` 的公式需转义；超长公式（含 `\frac`、多项）**不宜放表格**，改用列表。

### 规则 6：图片占位用 HTML 注释

```markdown
<!-- 图 1：待插入图片（说明） -->

**图 1**　图注文字…
```

渲染时注释不可见，图注保持可见，图片到位后直接在注释处插入 `![](...)`。

---

## 5. 质量校验清单

转换完成后必须逐项检查：

```python
import re

d = open("out.md", encoding="utf-8").read()

# 1. $$ 配对（必须偶数）
assert d.count("$$") % 2 == 0, "$$ 不配对"

# 2. 编号连续
tags = [int(t) for t in re.findall(r"\\tag\{(\d+)\}", d)]
assert tags == list(range(1, len(tags) + 1)), "编号不连续"

# 3. aligned/cases 全在显示块内
spans = []  # 所有 $$..$$ 区间
pos = 0
while True:
    a = d.find("$$", pos)
    if a < 0: break
    b = d.find("$$", a + 2)
    if b < 0: break
    spans.append((a, b)); pos = b + 2

for m in re.finditer(r"\\begin\{(aligned|cases)\}", d):
    assert any(a <= m.start() <= b for a, b in spans), "多行公式逃逸显示块"

# 4. 行内 $ 配对（屏蔽显示块后逐行检查）
masked = list(d)
for a, b in spans:
    for i in range(a, min(b + 2, len(masked))):
        masked[i] = " "
masked = "".join(masked)
assert all(l.count("$") % 2 == 0 for l in masked.split("\n")), "行内 $ 不配对"

# 5. 无空公式块
assert all(d[a + 2:b].strip() for a, b in spans), "存在空公式块"

# 6. 编码干净
assert d.count("\ufffd") == 0, "存在替换字符"
```

### 内容完整性审计（**最容易忽略**）

转换过程中的正则替换可能**静默丢内容**。必须与备份对比：

```python
old = open("备份.md", encoding="utf-8").read()
new = open("out.md", encoding="utf-8").read()

# 对比关键标记的计数
for kw in ["关键术语1", "关键术语2", ...]:
    assert old.count(kw) == new.count(kw), "「%s」数量变化：%d → %d" % (
        kw, old.count(kw), new.count(kw))
```

> **真实案例**：某次表格化转换中，正则要求「标签后 2 个以上空格」，
> 而一条数据的标签后只有 1 个空格 → **该条被静默丢弃**。
> 靠备份对比审计才发现。

---

## 6. 常见问题速查

| 现象 | 原因 | 解决 |
|---|---|---|
| `ParserError: 字符串缺少终止符` | PowerShell 脚本含中文 | 改为纯 ASCII |
| `Get-ChildItem -Include` 返回空 | 缺 `-Recurse` 或路径无通配 | 用 `Where-Object` 过滤扩展名 |
| `Pillow cannot find loader for WMF` | 缺 libwmf | 改用 Windows GDI+ |
| 公式块内多出 `$` | 行内公式转块时未删定界符 | 转块时剥离 `$...$` 的 `$` |
| 编号重复/跳号 | 先删后加的顺序错误 | 用哨兵标记区域，分段编号 |
| 中文文档路径找不到 | WSL/Windows 路径混淆 | `wslpath -w` 转换 |
| `uvx --from A --from B` 报错 | uvx 只允许一个 `--from` | 用 `--from A --with B` |

---

## 7. 编号管理的正确做法

多次增删公式编号时，**不要「删一个、补一个」**，而应：

```
① 清除全部 \tag
② 用哨兵标记「不编号」区间（如 @@NOEQ@@ ... @@END@@）
③ 分段按顺序重编号
④ 移除哨兵
```

代码骨架：

```python
import re

d = open(P, encoding="utf-8").read()
d = re.sub(r"\\tag\{\d+\}", "", d)                 # ① 清除

# ② 标记不编号区间
d = d[:i_a] + "@@NOEQ@@" + d[i_a:i_b] + "@@END@@" + d[i_b:]

cnt = [0]
def add_tag(m):
    body = m.group(1)
    if not body.strip():
        return m.group(0)
    cnt[0] += 1
    n = cnt[0]
    if re.search(r"\\begin\{(aligned|cases)\}", body):
        # 多行：tag 放 \end 之后（同一行）
        ia = body.rfind("\\end{aligned}")
        ic = body.rfind("\\end{cases}")
        idx, ln = (ia, len("\\end{aligned}")) if ia > ic else (ic, len("\\end{cases}"))
        body = body[:idx + ln] + ("\\tag{%d}" % n) + body[idx + ln:]
    else:
        s = body.rstrip()
        body = s + ("\\tag{%d}" % n) + body[len(s):]
    return "$$" + body + "$$"

pre, rest = d.split("@@NOEQ@@", 1)
mid, post = rest.split("@@END@@", 1)
pre = re.sub(r"\$\$(.*?)\$\$", add_tag, pre, flags=re.S)
post = re.sub(r"\$\$(.*?)\$\$", add_tag, post, flags=re.S)
d = pre + mid + post
```

---

## 8. 对「图示重绘」的警告

如果需要重绘原文的示意图，**先确认读清了图形结构**。

真实教训：本次重绘「位错环与曲面」示意图时，前 3 版都与原图差异巨大，原因：

1. **漏掉了核心元素**——原图中间那条粗点划线（表示环穿过曲面的背侧部分），
   漏掉后画出的只是一个普通椭圆
2. **搞错了轮廓形状**——原图是「眼形」（上下两弧在左右收成尖点），不是椭圆
3. 只看文字标注、没看清图形结构

**正确做法**：先建立归一化坐标（以图形宽度为 1），逐元素测量相对位置，
再按比例绘制。即便如此，若原图是手绘风格，复现也难达「一致」，不如直接用原图。

---

## 9. 复用资产清单

下表是**本仓库里**对应的实现（原始实验脚本在另一项目，未随本仓库发布）：

| 资产 | 本仓库位置 | 说明 |
|---|---|---|
| 段落与公式提取 | `py/extract.py` | 生成 `paras.json`、`eq_index.json`、`media/` |
| 文档结构与变体判定 | `py/docx_scan.py` | 段落数、MTEF 变体分布、渲染引擎探测 |
| WMF/EMF 渲染 | `py/render.py` | Windows GDI+ / LibreOffice 两路 |
| Markdown 生成 | `py/build.py` | 回填公式 + 编号 + 标题分级 |
| 结构校验 | `py/validate.py` | 六条铁律 + T8 公式数 + T9 未解析公式 |
| 坏输入处理 | `py/docxfile.py` | 非 zip / 空 zip / 缺 document.xml / 加密 |
| 工具封装 | `lib/index.js` | 三个模型侧工具 + 一个技能 |
| 本方法论 | `docs/methodology.md` | 全流程说明书 |

> 这些脚本由插件调用，也可以直接跑：
> `python py/docx_scan.py 论文.docx`。
> 依赖 `olefile`；缺依赖时插件会经 `uvx` 自动准备。
> **注意**：本文档 §2 描述的 EMF→PDF 文本提取（`pypdf`）在本仓库里尚未实现，
> 见文首注记。

---

## 10. 流程清单（可打印）

```
□ 0. 探测环境：olefile / pypdf / PowerShell
□ 1. 解压 docx，统计 embeddings 与 media 数量
□ 2. 解析 document.xml → paras.json（段落 + ⟦EQ:xxx⟧ 占位）
□ 3. 判定每个公式的 MTEF 变体（A / B）
□ 4. 变体 B → EMF 提取内嵌 PDF → pypdf 抽文本
□ 5. 全部预览图 → PowerShell 渲染为 PNG
□ 6. 逐图读图 → 转写 LaTeX → 存 latex_all.json
□ 7. 生成 Markdown（多行公式用 $$ 块！）
□ 8. 排版规范化（标题层级 / 编号 / 图注 / 列表）
□ 9. 六项结构校验 + 内容完整性审计
□ 10. 备份原文件，人工抽查 10% 公式
```
