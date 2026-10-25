# Changelog

## 0.2.0

### 新增：支持 Word 原生公式（OMML）

docx 里的公式有两族，此前只支持一族：

| | OLE 公式 | OMML（Word 原生） |
|---|---|---|
| 谁在用 | Equation Editor / MathType、WPS 默认 | Word 2007+ 默认；2016 后手敲 LaTeX 存下来也是它 |
| 存储 | 私有二进制 MTEF（`oleObjectN.bin`） | 结构化 XML（`m:oMath`） |
| 此前行为 | 渲染 PNG → 读图转写 | **静默丢失** |

修复前的实测：一份含 3 个 OMML 公式的文档，`inspect` 报 `formulas: 0`，
`extract` 的正文变成「设比值，且平方项为」——公式连同结构一起无声消失。
这是本插件最坏的一类失败（不报错、不完整、还看起来正常）。

现在：
- **`py/omml.py`**：OMML → LaTeX 确定性转换器。分式（含 linear）、上下标、
  根式（含 n 次）、定界符、大运算符（∑∏∫…）、函数名、重音（hat/vec/…）、
  横线、矩阵、方程组、underset/overset、纯文本样式（`\text{}`）、希腊字母
  与常用符号的 Unicode→命令映射。未知结构降级为拼接子节点——宁可少一层
  装饰，绝不丢内容、绝不抛异常（`limLow` 的 `m:lim` 双重身份曾实测触发
  崩溃，已修）。
- **`extract.py`**：OMML 生成 `⟦EQ:ommlK⟧` 占位符，LaTeX 直接写进
  `eq_index.json` 的 `latex` 字段（`kind:"omml"`、`pairing:"xml"`）。
- **`build.py`**：自动预填 OMML 的 LaTeX；Agent 的 `latex.json` 给同名键时
  以 Agent 为准（留了纠正的余地）。纯 OMML 文档用**空的** latex 表即可
  产出完整 Markdown。
- **`docx_scan.py`**：`formulas`（OLE）/ `formulas_omml` / `formulas_total`
  分开报，附 `omml_note`；两族都为 0 时给 `no_formulas_note` 提示可能是
  图片式公式。
- **计数口径统一**：`inspect` 与 `extract` 的 `formulas` 都只数 OLE，
  OMML 一律走 `formulas_omml`。混排文档实测两边一致（1+1=2）。
  `expectFormulas` 传的是 `formulas_total`。

### 测试

- `test/omml_selftest.py`：22 类结构逐一断言（含 linear 分式加括号、
  命令后紧跟 `_` 不补空格这类边角）。
- `test/omml_e2e.py`：合成 OMML docx 走 inspect→extract→build→validate
  全链路，钉死「公式绝不静默消失」。
- `run-all.sh` 接入为 7d / 7e 两组。
- `pack.mjs` 清单加入 `py/omml.py`（14 个文件）。

### 文档

- README 开头改为「两族公式，两条路」对照表；技能文档同步——
  「以位图为准」明确限定为 OLE 族的军规，OMML 走确定性路径、
  第 3 步（读图）对它不适用。

## 0.1.5

### 修复

- **段落被超长公式劈开后，前半段的行内公式漏补中文边界**。`build.py` 在遇到
  需要独立成块的长公式时，会先把已攒的文字 flush 出去；这条 flush 路径没有
  经过 `_boundary_after_formula`，于是前半段里的公式全都留下 `$公式$中文` 的
  形态。阈值越小（劈开越多）暴露越多：同一份 145 公式文档在阈值 40 时冒出
  3 条 T10 警告，阈值 80/150 时恰好一条都没有——这也是它从一开始没被发现的
  原因，默认值刚好没触发。

  修好后阈值 40 / 80 / 150 / 300 全部 0 警告。

- **`display_selftest.py` 只测默认阈值，抓不到上面这个 bug**。现在扫
  40 / 80 / 150 / 300 四个阈值，并新增一段「前面有短式紧贴中文 + 后面跟超长式」
  的合成段落——把修复回退后实测该测试会报 D1 失败，确认它有淘汰力。

- **D5 的公式计数口径错了**。它拿 `eq_index.json` 的条数当分母，但同一公式可以
  在正文里被多次引用，build 会把每处引用都算一个公式，导致合成用例误报。
  改为数 `paras.json` 里的占位符个数。

### 文档

- README 与 SKILL.md 补充「参数到底写在哪」：**不在插件管理界面**（那个页面只有
  插件信息与开关），而在 profile 的 `cordis.patch.yml` 里按 `id: docx2md` 加
  config 覆盖。同时写明 `patchReload: live` 下这条配置会热加载、不必重启。

## 0.1.4

### 新增

- **排版阈值改为可配置**。0.1.3 把 `LONG_TEX_CHARS` 从 150 降到 80，那是一次
  有取舍的改动——调小让长公式独立成块、显示更稳，但句子会被切得更碎。这类
  取舍没有唯一正确值，取决于文档给谁看，所以不该由插件单方面决定。现在两个
  阈值都开放：

  | 参数 | 默认 | 作用 |
  |---|---|---|
  | `longTexChars` | 80 | 混排段里超过这个长度就独立成块 |
  | `t5MaxCellFormula` | 60 | 表格单元格内公式的字符上限 |

  两种传法，优先级 **单次调用 > 插件配置 > Python 默认**：
  `docx2md_build(longTexChars=150)` 只对这一篇生效；改 profile 的插件配置则是
  全局默认。`build` 的返回值回报实际生效的 `long_tex_chars`，用它确认参数真的
  传进去了——不要只看自己传了什么。

  同一篇 145 公式文档实测：`longTexChars` 40 → 行内 28 / 独立 117；80 → 43 / 102；
  150 → 57 / 88；300 → 68 / 77。**公式总数始终 145**，改编排不会丢公式。

### 修复

- **`test/display_selftest.py` 把阈值写死成 80**（0.1.3 引入的问题）。阈值一旦
  变成用户可配，这个测试在用户改配置后就会误报 D3 失败——那是测试的错，不是
  产物的错。现在从 `build` 回报的 `long_tex_chars` 读取实际值；已用阈值 150
  实测，不再误报。

### 测试

- `tool-contract.mjs` 增补四条：`docx2md_build` 的 parameters 里确实暴露了这两个
  整数参数、Config 填默认值、拒绝非法值（0 / 小数 / 负数）、接受合法覆盖。
  参数是用户的调节旋钮，它消失或行为变化必须有测试拦住。

## 0.1.3

### 修复

- **行内公式紧贴中文，导致部分渲染器不显示公式**。生成的 md 里出现过
  `粒子（如$\alpha$粒子）`、`应力场$\sigma_{xy}$` 这样的写法。不少渲染器
  （Pandoc 的 `tex_math_dollars`、部分 KaTeX 集成）只在 `$` 与前后字符有明确
  边界时才把它识别为公式；紧贴中文时按普通文本处理，用户看到的就是
  **LaTeX 源码原样出现**，或者公式整块空白。实测一份 145 公式的真实文档里
  有 21 处属于这种情况。

  `build.py` 现在在行内公式与 CJK 字符相邻时补一个空格；纯 ASCII 上下文
  保持原样，`$a$b` 这类正常写法不受影响。

- **校验器新增 T10**：逐处报出「公式紧贴中文」。这是警告级——它不该把文档
  判死，但必须在交付前被看见。修复前的产物会报出 48 条，修复后为 0。

- **行内长公式阈值 150 → 80**。实测这份文档里有 115–119 字符的公式留在行内，
  在中文字里几乎必然折行或溢出，窄栏与移动端更容易出问题——用户反馈的
  「只有长公式坏、短的正常」正落在这个区间。降到 80 后，这份文档的行内公式
  从 57 个减到 43 个、最长 80 字符，display 从 88 升到 102，公式总数不变
  （145）。`a+b`、`\alpha` 这类真正的短式仍留在句内，不破坏句子连贯性。

- **新增 `test/display_selftest.py`**：把这次的判据固化成五条（D1 边界 / D2 行内
  禁块级环境 / D3 长度阈值 / D4 `$$` 独占且块内无裸 `$` / D5 公式总数），
  合成段落与真实文档都跑。这是唯一能拦住「KaTeX 渲染得出来、但渲染器不认」
  这类回归的测试。

### 说明

- 这次排查还确认了一件事，值得写进文档：**`katex.renderToString` 能渲染，
  不等于 Markdown 解析器会把它识别成公式**。这是两层独立判定。本轮把两层
  分开测过：145 条公式在 KaTeX 严格模式下 0 失败，但仍有 21 处因为定界符
  边界问题而不显示。只测前者会给出虚假的信心。

## 0.1.2

### 修复

- **公式配错了预览图，而且是静默的**。旧实现按「同段落第 k 个 `OLEObject`
  ↔ 第 k 张 `v:imagedata`」配对。只要公式前面夹着一张插图，下标就整体右移，
  于是把插图的位图当成公式预览图交给模型去读——不报错、不告警。实测在一份
  145 个公式的真实文档里错了 10 个（配成了 VDDD3D 程序截图与位错环三维图），
  而 `summary` 只统计「没图」的公式，看不见「配错图」。

  现在按 `OLEObject` 的 `ShapeID` 精确配对到带 `o:ole="t"` 的 `v:shape`，
  再由该 shape 的 `imagedata` 取图。同一份文档实测 145/145 精确命中，
  `formulas_without_media` 从 8 降到 0。文档不写 `ShapeID` 时才退回下标配对，
  并在 `summary.formulas_paired_by_ordinal`、`pairing_warning` 与 `next` 里
  明确报出「这些图未经确认」。

- **校验器对铁律 2、5 的两个条款根本没查**。`validate.py` 的 docstring 写着
  T2 要查「行内公式不得含 `\\`」，铁律 2 原文也列了 `\\` 与 `&`，但正则只查了
  `\begin{`。于是 `$a \\ b$` 与 `$a & b$` 都拿到 PASS。现已补齐，并排除
  `\&`、`\{`、`\}` 这类转义，避免误报。

- **T5 的阈值与文档不符**。铁律 5 说「超过约 60 字符就移出表格」，实现查的是
  行长度 400，于是 100 字符的公式塞进单元格也通过。现在按单元格内公式自身的
  LaTeX 长度与 60 比较，并在告警里给出实际字符数。

- **复用 `outDir` 时 `headings.json` 不被清理**。它是上一份文档人工确认过的
  「段落索引 → 标题级别」，段落号对另一份文档毫无意义；而 `build.py` 只要看到
  它存在就直接采用（`heading_source="headings.json"`），于是新文档的标题会
  挂到错误的段落上，且无声。现已纳入清理清单。

- **`test/smoke.mjs` 依赖仓库外文件**。它写死 `--latex /tmp/latex_by_stream.json`，
  这个文件只在开发机上手工存在过，仓库里没有任何东西会创建它；别人 clone 后
  跑该测试必然在第 5 步失败。现在由测试自己从 `eq_index.json` 生成映射表。

- 清掉 `py/render.py` 末行的 `# drift-test` 与 `py/validate.py` 末行的 `# x`。

### 文档

- README 的 `formulas_without_preview` 更正为实际字段名
  `formulas_without_media`；「失败时不会留下空目录」限定为 Python 脚本层
  （工具层会先按 `outDir` 建目录再校验），并补充配对规则与清理 `headings.json`
  两条已知行为。

### 测试

- 新增 `test/pairing_selftest.py`：用合成 docx 钉死配对规则——有 `ShapeID` 时
  段落内夹插图也能配到真正的公式图；无 `ShapeID` 时退回下标配对且必然告警。
- `validate_selftest.py` 增补 T2（`\\`、`&`、转义放行）与 T5（警告级）反例。
- `run-all.sh` 接入配对自测，并把「T0–T8」更正为「T0–T9」。

## 0.1.1

### 修复

- **中文输出在某些环境下变成 `����`**。中文 Windows 上 Python 的
  `sys.stdout.encoding` 是 gbk，非 ASCII 字符被编成 GBK 字节，接收方按 UTF-8
  解码就得到乱码。这不只是显示问题：`missing_preview_context[].paragraph_text`
  是判断「缺图公式属于图注还是正文」的唯一依据，乱码等于这个判断失效。
  现在三层防御：进程层注入 `PYTHONUTF8`/`PYTHONIOENCODING`、脚本层每个入口调
  `force_utf8_stdio()`、接收层严格 UTF-8 解码（非法字节直接报错，不静默替换）。

- **WSL 里写 `/tmp/x` 会落到 `C:\tmp\x`**。`path.isAbsolute('/tmp/x')` 在
  Windows 上为 true，旧代码因此跳过归一化，把字符串原样交给 Python，Python
  按盘符根解析。现在一律由 Node 按会话工作目录的卷归一化；`/mnt/c` 自动映射成
  `C:\`，返回的路径也会折回调用方习惯的写法。

- **公式没解析出来时校验仍报 PASS**。`build.py` 把缺失的公式写成
  `<!-- 未解析公式: xxx -->`，这种文档六条铁律全过、注释渲染后也看不见，但正文
  里少了公式。新增 **T9**：残留 `⟦EQ:...⟧` 与上述注释形态都判 FAIL。

- **坏输入会在磁盘上留下空目录**。`extract.py` 原先先建 `outDir` 再校验文档，
  文档损坏时留下空目录。已改为先校验。

- **渲染报告 `_render_report.json` 残留在 `png/`**。读完即删。

### 新增

- `docx2md_build` 的 **`cleanup`** 参数：构建并校验通过后删除 `media/` 与
  `png/`，保留 `paras.json`、`eq_index.json` 供复现。校验失败时不删——那时
  `png/` 正是排查「哪张图读错了」的唯一依据。

- `test/sync-check.sh`：检查「源码 → 安装副本 → node_modules」三处是否一致、
  宿主是否需重启。退出码 0/1/2 分别表示可验证 / 代码不同步 / 需重启。
  这是**开发本插件专用**的工具，使用者不需要。

- `test/pack.mjs`：发布前检查 tarball 内容。修掉一个实测发现的问题——
  `files: ["py"]` 会把 `py/__pycache__/*.pyc` 打进包，而 `.npmignore` 在
  `files` 白名单模式下对已纳入的目录**不生效**。改用 `py/*.py` 后包体积
  从 121.9 kB 降到 107.1 kB。

- `test/github-install.mjs`：从 tarball 真实安装、加载、跑一份文档、卸载。
  GitHub 安装产出的就是 tarball，与 `file:` 本地目录安装是两条不同的路径。

- `.gitignore`：没有它时 `.pyc` 会被提交进版本历史。

### 测试可移植性

原先测试脚本里写死了 `C:/Users/<某用户>/...`。现在全部走
`test/lib/discover.sh`：一次性目录用 `mktemp -d`，Windows 路径由
`to_win_path` 换算，本机 DSH 位置按「环境变量 → 自动发现 → 明确跳过」处理。
别人 clone 后跑 `test/run-all.sh`，需要 Windows 的组会明确 skip 并返回 0，
不会误报失败。

## 0.1.0

首个版本。

- 三个工具：`docx2md_inspect` / `docx2md_extract` / `docx2md_build`
- 一个技能：`docx2md`
- 三种公式来源按可靠性排序：渲染 PNG 读数（★★★★★）> EMF 内嵌 PDF 文本
  （★★★★，变体 B 的唯一路径）> MTEF 解码（★★，仅作交叉校验）
