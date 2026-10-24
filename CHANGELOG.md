# Changelog

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
