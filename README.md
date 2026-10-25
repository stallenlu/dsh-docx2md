# dsh-docx2md

把含公式的 Word 文档转成带 LaTeX 的 Markdown —— 跨平台，保留公式编号与结构。
**OLE 公式（Equation Editor / MathType）、Word 原生公式（OMML，新版 Word 的
默认格式）、图片形式的公式（截图/扫描）三种载体都处理。**

A DeepSeek Harness plugin that converts Word documents with embedded equations —
legacy OLE (MTEF), native OMML, and image-embedded formulas — into Markdown with
LaTeX, preserving numbering and document structure.

---

## 三种载体，三条路

| | OLE 公式 | OMML（Word 原生） | 图片公式 |
|---|---|---|---|
| 谁在用 | Equation Editor / MathType；WPS 公式的默认载体 | Word 2007+ 插入公式的默认格式；2016 后手敲 LaTeX 存下来也是它 | 截图粘贴、扫描页、从 PDF 拷贝 |
| 存在哪 | 私有二进制（MTEF）包在 `oleObjectN.bin` | 结构化 XML（`m:oMath`） | 一张位图（`w:drawing` / `v:pict`） |
| 怎么转 | 渲染成 PNG → 视觉模型逐张读图转写 | **extract 阶段确定性转成 LaTeX**，不读图 | **无法自动转换**（像素无结构可解）；落成占位符，读 `media/` 原件转写，键 `"img:文件名"` |
| 质量 | 取决于读图；有校验器与配对告警兜底 | 22 类结构逐一测试（分式/上下标/根式/矩阵/大运算符/重音/…） | 取决于读图；候选按形状启发式（扁条行内图），清单带证据供人过目 |
| 需要人工 | 是（第 3 步） | 否（不满意可覆盖，见下） | 是（第 3.5 步） |
| 不处理会怎样 | 报 T9 未解析 | —— | **不丢**：build 落成可见图片引用（`![图片公式 …]`），T11 查断链 |

纯 OMML 文档全程无需读图：`docx2md_build` 用**空的** latex 表也能产出完整
Markdown（OMML 自动预填）。混排文档里只有 OLE 与图片那部分需要视觉转写。

## 为什么 OLE 需要看图

普通的 `pandoc doc.docx -o doc.md` 遇到 Word 的 OLE 公式基本无能为力：那是
WPS / Word 的私有二进制格式（MTEF），解出来必然失真。

本插件对 OLE 公式的方案只有一句话：

> **以渲染出的位图为准，不以解析出的字符为准。**

实测同一份 145 个公式的文档，直接解 MTEF：

| 公式 | 看图应有的字符量 | MTEF 解出 |
|---|---|---|
| #27 | 151（5 行矩阵式） | `βP` — 2 个字符 |
| #30 | 150（5 行） | `u=dr` — 4 个字符 |
| #6 | 5 个等式链 | `α=-∇×βP` |

**它不会报错，只是安静地给你一个残缺结果。** 这是最危险的一类失败。
所以本插件的设计原则是：渲染失败就明确失败，绝不降级到字符解码。

---

## 安装

DSH 的 profile 用 **pnpm** 管理，`node_modules` 是 pnpm 管的（`nodeLinker: hoisted`）。
**不要手工往 `node_modules` 里拷目录**——下次 `pnpm install` 会把没有写进
lockfile 的目录删掉。走 pnpm 正规安装：

```bash
# 1. 把包放到 profile 外部（file: 依赖不能指向 node_modules 内部，会成自环）
cp -r dsh-docx2md ~/.dsh/packages/dsh-docx2md

# 2. 在 profile 里用 pnpm 安装（会写进 pnpm-lock.yaml）
cd ~/.dsh/profiles/desktop
pnpm add file:~/../packages/dsh-docx2md
```

Windows 上路径要写成 `pnpm add file:C:/Users/<你>/.dsh/packages/dsh-docx2md`。

### 声明的两处

`package.json` 里有两件事都要做：

```jsonc
{
  "dependencies": { "dsh-docx2md": "file:../../packages/dsh-docx2md" },
  "dsh": { "profile": { "bundles": [ /* ... */, "dsh-docx2md" ] } }
}
```

**只需加进 `bundles`**。本包的 `package.json` 已声明
`dsh.bundle.patch: ./cordis.patch.yml`，DSH 会自动读取它并把插件插入。
**不要再往 profile 的 `cordis.patch.yml` 里手写 `insert`** —— 那会导致同一插件被
注册两次。

### 让配置改完立即生效

profile 的 `package.json` 里加上：

```jsonc
"dsh": { "profile": { "patchReload": "live" } }
```

没有这一项时，`bundles` 的变更要**重启 DSH 进程**才生效。
注意：技能列表走另一套机制（`dsh-skill-filesystem` 用 chokidar 监听文件），
所以可能出现"技能已出现、工具还没出现"的中间状态——这不是 bug，重启即可。

### 从 Git 仓库安装

```bash
cd ~/.dsh/profiles/desktop
pnpm add github:<你的账号>/dsh-docx2md
```

再把包名加进 `bundles`。本地开发时用 `file:` 装一次，之后改
`~/.dsh/packages/dsh-docx2md` 里的文件即可——pnpm 建的是**硬链接**，改动
会立即反映到 `node_modules`（若用编辑器原子写入导致硬链接断开，
重跑一次 `pnpm install` 即可）。

---

### 更新

`file:` 依赖在 lock 里记为 `type: directory`（不是版本号），所以 pnpm **不按版本判断
是否需要更新**——它只是指向那个目录。因此“更新”要分三种情况：

| 情况 | 需要做什么 | 实测 |
|---|---|---|
| 只改代码，版本号不变 | `pnpm install` + 重启 DSH | node_modules 立即同步（硬链接） |
| 改了版本号 | 同上 | package.json 随硬链接自动同步 |
| 文件被原子写入（编辑器保存、`git checkout`、`cp` 覆盖） | **必须 `pnpm install`** | 硬链接断开，node_modules 保留**旧代码** |

**第三条是最容易踩的坑**：inode 变了，硬链接断了，`node_modules` 里还是旧代码。
没有报错，工具照常运行，只是行为没变——你在源目录看到的是新代码，实际加载的不是。

标准更新流程：

```bash
# 1. 改代码（改 ~/.dsh/packages/dsh-docx2md 或你的源码目录后同步过去）
# 2. 修复可能断开的硬链接（1-2 秒，不要省）
cd ~/.dsh/profiles/desktop && pnpm install
# 3. 重启 DSH
# 4. 验证：跑一次 docx2md_inspect 确认行为符合预期
```

`test/update.sh` 会把这四种情况都在一次性环境里复现一遍。

#### 改完先跑 sync-check.sh

上面三步（同步 → 修链接 → 重启）只要漏一步，跑出来的“验证通过”就是假的。
开发时改的是源码目录，DSH 加载的是 `node_modules` 里那份，中间隔着复制和硬链接
两道；任何一道断了都不报错，只是行为不变。

```bash
./test/sync-check.sh          # 检查
./test/sync-check.sh --fix    # 不一致时自动同步源码 → 安装副本
```

它检查四件事，并用退出码区分严重程度：

| 退出码 | 含义 | 该怎么办 |
|---|---|---|
| 0 | 三处一致，且进程加载的是最新代码 | 可以跑真机验证 |
| 1 | 代码不同步（副本落后 / 硬链接断裂） | `--fix`，必要时 `pnpm install` |
| 2 | 三处代码已一致，但 DSH 进程是旧版本 | 重启 DSH，之后才有意义 |

分开 1 和 2 是有意的：退出码 2 表示「代码没问题，只差重启」，CI 不该因此报失败；
退出码 1 才是真的代码不一致。

脚本以 `package.json` 的 `files` 字段为准来列清单（`test/` 不随包安装，不参与比对）。
它也会读宿主进程启动时间，与最新文件修改时间对比，因此能直接告诉你「要不要重启」。

---

### 卸载

```bash
cd ~/.dsh/profiles/desktop
pnpm remove dsh-docx2md
```

然后从 `package.json` 的 `dsh.profile.bundles` 里删掉 `"dsh-docx2md"`，重启 DSH。
`~/.dsh/packages/dsh-docx2md` 可以留着（下次重装直接用），也可以删掉。

验证是否干净：`node_modules/dsh-docx2md` 不存在、`pnpm-lock.yaml` 无记录、
`bundles` 无条目。本仓库 `test/lifecycle.sh` 就是跑这套闭环的。

---

## 运行环境

| 依赖 | 用途 | 必需 |
|---|---|---|
| Python 3.9+ | 跑 py/ 下的脚本 | ✅ |
| `olefile` | 解 OLE 容器、取 Equation Native 流 | 自动经 uvx 安装 |
| `pypdf` | **预留，当前未使用**（见「已知限制」）| 自动经 uvx 安装 |
| `uv`（可选） | 免手动装 Python 依赖 | 推荐 |

**渲染引擎**（用于把 WMF/EMF 公式图转 PNG），按顺序自动探测：

| 平台 | 引擎 | 说明 |
|---|---|---|
| Windows（桌面 App） | GDI+ (`System.Drawing`) | 原生支持；**实测 127 张图 4.9 s** |
| WSL | 调宿主 Windows 的 GDI+ | 自动转换路径为 UNC |
| Linux / macOS | LibreOffice (`soffice`) | 需自行安装 |
| 兜底 | Pillow + libwmf | 少见，有就用 |

全都没有时，`docx2md_inspect` 会明确告诉你，`docx2md_extract` 会保留原始
矢量图、把 `render.ok` 置为 false——不会假装成功。

---

## 用法

### 在 Agent 里

直接说：

> 把 `论文.docx` 转成 Markdown，公式要保留。

Agent 会加载 `docx2md` 技能，按 inspect → extract → 读图 → build 的顺序执行。

### 三个工具

```
docx2md_inspect(path)
```
只读扫描。返回段落数、公式数（分 `formulas`＝OLE 与 `formulas_omml`＝Word 原生
三类：OLE / OMML / 图片公式候选）、变体分布（A 可解 / B 必须看图）、预览图格式分布、本机渲染引擎。

**它只输出事实，不给结论。** 例如缺预览图的公式，`inspect` 会给出
`missing_preview_context`（所在段落、段落文字、前后段），但不会替你判断
「这是图注标注还是正文公式」——那需要看全文，属于调用方的判断。

```
docx2md_extract(path, outDir?, render?)
```
产出：
- `paras.json` —— 正文段落，公式位置为 `⟦EQ:oleObject12.bin⟧`（OLE）、
  `⟦EQ:omml3⟧`（Word 原生）或 `⟦EQ:img:文件名⟧`（图片公式候选）
- `eq_index.json` —— 每个公式的 stream / 预览图 / 字节数 / 变体 / `pairing`；
  **OMML 条目直接带 `latex` 字段**（从 XML 结构确定性转出，无需读图）
- `media/` —— 原始 EMF/WMF 与图片公式的位图原件
- `png/` —— 可读位图（**读这个目录**）
- `image_inventory.json` —— 正文全部行内图片的形状证据（ratio / 尺寸 /
  上下文段落），图片公式候选的判定依据；示意图也在这里，供人核对

```
docx2md_build(workDir, latex, out, suggestHeadings?, expectFormulas?)
```
把 `{"oleObject1.bin": "\\alpha", ...}` 合成 Markdown 并校验。
`expectFormulas` 会核对公式总数，能抓住漏转。OMML 公式无需出现在 latex 表里
（自动预填；给了同名键则以你给的为准）。图片公式不转写也不丢——落成
`![图片公式 …](media/…)` 可见引用；转写时键用 `"img:文件名"`。

### 产物落在哪、会不会留垃圾

插件**不创建系统临时目录**（不用 `mkdtemp` / `tempfile` / `TMPDIR`），所有中间产物
都在你给的 `outDir` 里；不给就落在 docx 同级的 `<名字>_docx2md/`。

一次转换产生 6 个条目（145 公式的论文实测 4.3 MB）：

| 条目 | 内容 | 大小 |
|---|---|---|
| `paras.json` | 正文段落，公式位置为 `⟦EQ:oleObjectN.bin⟧` | ~20 KB |
| `eq_index.json` | 公式索引（stream / 预览图 / MTEF 变体） | ~24 KB |
| `image_inventory.json` | 行内图片的形状证据与上下文（图片公式候选清单） | ~1 KB |
| `media/` | 原始矢量预览图（wmf/emf）+ 文档自带位图 | ~1.3 MB |
| `png/` | 渲染后的可读位图，Agent 靠它转写 LaTeX | ~3.0 MB |
| `summary.json` | 本次提取的统计 | ~0.3 KB |

**不会残留的东西**：渲染用的 `_render_report.json` 读完即删；坏输入不会留下空目录。
`png/` 里出现的 `.jpeg` 不是残留，而是原文档自带的位图预览（`inspect` 的
`media_formats` 会报出 jpeg 数量）。

**清理**：插件没有 `--clean` 参数，也不需要——`outDir` 是唯一写入点，`rm -rf` 该目录即可。
同一 `outDir` 复用时旧产物会被自动清掉，不会混入上一份文档的图。

**清理**：默认什么都不删（`png/` 是你事后核对公式的凭据）。想要自动清理就传
`cleanup: true` 给 `docx2md_build`：

```js
await docx2md_build({
  workDir: 'work/', latex: 'work/latex.json', out: '论文.md',
  expectFormulas: 145, cleanup: true,
})
```

`cleanup` 只在**构建完成且校验 PASS** 时才删 `media/` 与 `png/`，保留
`paras.json`、`eq_index.json`、`summary.json`（复现用）。校验没通过时不删——
那时 `png/` 正是排查「哪张图读错了」的唯一依据。返回值里的 `cleanup` 字段会
说明删了什么、或为什么没删。

也可以不给 `cleanup`，直接 `rm -rf` 该目录，效果一样。

### 校验器查什么

`docx2md_build` 内部会调 `validate.py`，除六条 Markdown 铁律外还查：

| 检查 | 抓住什么 |
|---|---|
| T0 | 空文件 |
| T1–T3 | `$$` 未独占一行、行内公式含块级环境、相邻公式块紧贴 |
| T7 | `$$` 奇数个（未闭合） |
| **T9** | **公式没解析出来**（残留 `⟦EQ:...⟧`，或 build 写的 `<!-- 未解析公式 -->`） |
| T8 | `expectFormulas` 给了时，公式总数不符 |

T9 值得单独说：latex 表里缺的公式，build 会写成 HTML 注释。这种文档六条铁律
全过、注释渲染后也看不见，但正文里确实少了公式。早期版本只查铁律，会给出
PASS——现在判 FAIL。

### 完整流程示例

```js
// 1. 看清有什么
await docx2md_inspect({ path: '论文.docx' })

// 2. 提取 + 渲染
await docx2md_extract({ path: '论文.docx', outDir: 'work/' })

// 3. ← 这里由 Agent 逐张读 work/png/*.png，转写成 latex.json
//    （这一步无法自动化：需要视觉理解公式结构）

// 4. 先看标题初稿
await docx2md_build({
  workDir: 'work/', latex: 'work/latex.json',
  out: '论文.md', suggestHeadings: true,
})
// 人工修订 work/headings_suggested.json → 改名 headings.json

// 5. 正式构建 + 校验
await docx2md_build({
  workDir: 'work/', latex: 'work/latex.json',
  out: '论文.md', expectFormulas: 145,
})
```

### 路径怎么写都行

三个工具都接受你从当前环境出发自然写出的路径，不需要考虑宿主是 Windows 还是 WSL：

| 你写的 | 在 Windows 宿主上落到 | 说明 |
|---|---|---|
| `论文.docx` | 会话工作目录下 | 相对路径 |
| `/home/sidady/proj/论文.docx` | WSL 视图的同一位置 | Linux 绝对路径 |
| `/tmp/work` | **WSL 的 `/tmp/work`** | 不是 `C:\tmp` |
| `/mnt/c/work` | `C:\work` | WSL 挂载点会自动换算 |
| `C:\work\cdot.docx` | `C:\work` | 盘符路径原样 |
| `\\wsl.localhost\Ubuntu\tmp` | 原样 | UNC 路径原样 |

返回的路径会自动折回你更习惯的写法：WSL 里的位置返回成 `/tmp/work`，Windows 位置返回成 `C:\work`，可以直接粘进 shell 或 `read` 工具。

> 早期版本在这里有个坑：`path.isAbsolute('/tmp/x')` 在 Windows 上返回 `true`，插件就把
> 字符串原样交给了 Python，而 Python 按盘符根解析成了 `C:\tmp\x`。现在一律先由 Node
> 按会话工作目录的卷归一化，不再透传未归一化的字符串。回归测试见 `test/paths.mjs`。

---

## 为什么拆成三步

因为 **第 3 步不可能自动化**。把 145 张公式图转成 LaTeX，需要视觉模型逐张看懂
矩阵有几行、上下标在哪、括号层级如何。

如果做成一个大工具，Agent 就必须在一次调用里完成读图——它做不到。
拆开之后，Agent 可以在 extract 和 build 之间插入自己的读图逻辑，
结果通过一个 JSON 文件交接。

---

## 配置项

| 键 | 默认 | 说明 |
|---|---|---|
| `maxImageWidth` | 1600 | 渲染 PNG 的最大宽度 |
| `allowUvx` | true | 允许经 uvx 自动装 Python 依赖（需联网） |
| `python` | `''` | 指定解释器路径；留空自动探测 |
| `timeoutMs` | 600000 | 单次调用的超时 |
| `longTexChars` | 80 | 混排段里 LaTeX 源串超过这个长度就独立成块 |
| `t5MaxCellFormula` | 60 | 表格单元格内公式的字符上限（超过给 T5 警告） |

### 排版阈值怎么调

后两个是**排版取舍**，没有唯一正确的值——取决于文档给谁看：

| `longTexChars` | 效果 | 适合 |
|---|---|---|
| 小（如 40） | 大量公式独立成块，句子被切得更碎 | 移动端、窄栏、公式极长的文档 |
| 80（默认） | 长式独立、短式留句内 | 大多数情况 |
| 大（如 150–300） | 更多公式留在句内，句子更连贯 | 桌面 / A4 单栏、看重段落流畅的读者 |

同一篇 145 公式文档实测：40 → 行内 28、300 → 行内 68；**公式总数始终 145**，
改编排不会丢公式。

### 在哪改：两种传法

优先级 **单次调用 > profile 配置 > 默认值**。

**① 单次调用**——只对这一篇生效，最省事：

```js
await docx2md_build({ workDir: 'work/', latex: 'latex.json', out: '论文.md', longTexChars: 150 })
```

**② 改 profile 配置**——对所有转换生效。

> ⚠️ **不在插件管理界面。** 那个页面只显示插件信息与开关，没有参数表单。
> 参数写在 profile 的 `cordis.patch.yml` 里（`~/.dsh/profiles/desktop/cordis.patch.yml`），
> 按插件 `id` 加一条 config 覆盖：

```yaml
# ~/.dsh/profiles/desktop/cordis.patch.yml，追加到文件末尾
- id: docx2md
  name: dsh-docx2md
  config:
    longTexChars: 150
    t5MaxCellFormula: 200
```

`id` 必须是 `docx2md`（插件的 `cordis.patch.yml` 里声明的那一个），不是包名
`dsh-docx2md`。

profile 的 `package.json` 里若有 `"dsh": { "profile": { "patchReload": "live" } }`，
这条配置**会热加载**，不必重启 DSH。改完直接跑一次 `docx2md_build` 看返回值里的
`long_tex_chars` 就知道有没有生效。

**验证参数真的生效**：`docx2md_build` 的返回值里有 `long_tex_chars`，那是 Python
实际收到的值。不要只看自己传了什么——配置没写对时它不会报错，只是静默用默认值。

---

## 产出的 Markdown 遵循六条铁律

每条都来自真实踩坑：

1. `$$` 必须独占一行 —— 否则与后文某个 `$$` 错配，**从那里到文末全乱**，且不报错
2. 行内公式不能跨行 —— 不能含 `\begin{aligned}`、`\\`、`&`
3. 相邻公式之间要隔开 —— `$a$$b$` 会被解析成一个 `$$`
4. `\tag` 紧跟 `\end{aligned}`
5. 超长公式不要塞进表格单元格
6. 图片占位用 HTML 注释 —— `![]()` 空链接会显示破图
7. **行内公式与中文之间要有空格** —— `中文$公式$中文` 在不少渲染器里不被识别为
   公式，用户看到的是 LaTeX 源码原样出现。`build.py` 会自动补，校验器用 **T10** 兜底。

`docx2md_build` 内置 `validate.py` 会逐条检查（T1–T10）。

---

## 测试

```bash
# 一次跑完全部功能与契约测试
./test/run-all.sh                     # 跳过需要真实 docx 的组
./test/run-all.sh 论文.docx            # 含端到端

# 发布前：额外检查 tarball 内容
./test/run-all.sh --with-pack 论文.docx

# 单跑某项
node test/smoke.mjs   /path/to/file.docx
node test/plugin.mjs  /path/to/file.docx
node test/pack.mjs                    # 打包内容
```

所有脚本都输出 PASS/FAIL 判定表，退出码可用于 CI。

## 发布

上 GitHub 前先跑：

```bash
./test/run-all.sh --with-pack 论文.docx
```

它会额外验两组：`test/pack.mjs`（tarball 里有什么）和 `test/github-install.mjs`
（拿 tarball 真装一次、真跑一次、再卸载）。后者是关键——GitHub 安装最终产出的
就是 tarball，`file:` 本地目录安装走的却是另一条路（硬链接、整个目录都在），
两者不能互证。

### 打包内容

| 检查 | 目的 |
|---|---|
| 文件清单与预期完全一致 | 不多不少 |
| 无 `.pyc` / `__pycache__` / `test/` / 编辑器残留 | 干净 |
| 解包后能 import、能注册工具与技能 | 不依赖源码目录结构 |
| 解包后 `py/*.py` 能独立运行 | 同上 |
| `files` 写的是 `py/*.py` 而非 `py` | 见下 |

**`files` 为什么不能写裸目录。** 实测：`files: ["py"]` 会把
`py/__pycache__/*.pyc` 一起打进去（本机有 3 个 Python 版本的字节码）；
而且 `.npmignore` 在 `files` 白名单模式下对已纳入的目录**不生效**——这一条与
直觉相反，是实测结论。改成 `py/*.py` 才真正排除，包体积从 121.9 kB 降到 107.1 kB。

当前包的 15 个文件：`LICENSE`、`README.md`、`cordis.patch.yml`、`package.json`、
`lib/index.js`、`lib/toolkit.js`、`py/*.py`（8 个）、`skills/docx2md/SKILL.md`。

### 从 tarball 安装这条路径

`test/github-install.mjs` 验证：打包 → 干净 profile 里 `pnpm add` → 检查
`node_modules` 内容 → 加载并注册 → 真跑一次 `inspect` → `pnpm remove` 后无残留。

三个只在实测中才会暴露的坑，都写进脚本注释了：

| 坑 | 现象 | 解法 |
|---|---|---|
| UNC 路径下建符号链接失败 | `os error 1（函数不正确）` | 测试环境放原生 Windows 路径 |
| Node 不能直接 spawn `.cmd` | `EINVAL` / `ENOENT` | 经 `cmd.exe /c` 调用 npm 与 pnpm |
| pnpm 拦截包内 skill | `Agent skills awaiting approval` | 功能不受影响；可 `pnpm approve` 复核 |

第三条是 pnpm 对第三方包声明 skill 的安全机制，不是插件缺陷，但用户会看到这条
提示，值得在 README 里说明。

### 安装时看到「Agent skills awaiting approval」怎么办

pnpm 12 起会对第三方包内声明的 skill 做一次人工确认，安装后打印：

```
Agent skills awaiting approval: dsh-docx2md@...
Run "pnpm approve" to review them.
```

这是 pnpm 的安全机制，**不是安装失败**，工具与技能都已经装好了（`test/github-install.mjs`
在出现这条提示的情况下仍实测通过：3 个工具 + 1 个技能注册正常，能处理真实文档）。
这条提示是 pnpm **每个依赖**（包括其他插件）一共汇总，不是本包专属。

想消掉它，在该 profile 目录执行一次：

```bash
cd ~/.dsh/profiles/desktop
pnpm approve
```

它会列出待批准项，按提示确认即可。若不想交互，也可在 `pnpm-workspace.yaml` 里
把本包加入 `onlyBuiltDependencies` / 相应的允许列表（具体字段随 pnpm 版本变化，
以 `pnpm approve` 的提示为准）。

### 测试的可移植性

`test/` 里的脚本原先写死了 `C:/Users/<某用户>/...`。现在全部通过
`test/lib/discover.sh` 解析：

| 需求 | 做法 |
|---|---|
| 一次性目录 | `mktemp -d`，落在系统临时目录 |
| Windows 路径 | `to_win_path` 把 `/mnt/c/...` 换算成 `C:\...` |
| 本机 DSH 位置 | 环境变量 → 自动发现 → **明确跳过** |
| Windows node / pnpm | `find_win_node` / `win_pnpm` |

**「明确跳过」是成功而不是失败。** 需要 Windows 环境的测试在 Linux/macOS 上会
打印 `skip ...` 并返回 0——否则别人 clone 后跑测试会莫名变红。

`sync-check.sh` 尤其如此：它检查本机 DSH 的三处代码是否同步，对使用者本来就
不适用，找不到目标时直接跳过。

### 仓库卫生

`.gitignore` 排除 `__pycache__/`、`*.pyc`、`*.tgz` 等。没有它时 `.pyc` 会被提交
进版本历史——之后再加 `.gitignore` 也只是停止追踪新文件，已有记录仍在。

### 开发用的 sync-check.sh 不进包

`test/sync-check.sh` 检查「源码 → 安装副本 → node_modules」三处是否一致，以及
宿主进程是否需要重启。它依赖本机的 pnpm `file:` 硬链接、宿主进程名和 `~/.dsh`
路径，**是开发环境专用的**，不出现在 tarball 里（`test/` 整目录都不进包）。

发布者用不到它；改代码后用它避免「改了没生效」的假验证：

```bash
./test/sync-check.sh          # 退出码 0=可验 / 1=代码不同步 / 2=需重启
./test/sync-check.sh --fix    # 自动同步源码到安装副本
```

### 中文编码：为什么单测要在伪造的 GBK 下跑

中文 Windows 上 Python 的 `sys.stdout.encoding` 默认是 **gbk**，非 ASCII 字符会被编成
GBK 字节；接收方按 UTF-8 解码就得到 `����`。这个故障是**静默**的，而且只在中文
Windows 上出现 —— 在 Linux/macOS 上永远不会复现（那里恒为 UTF-8）。

它的后果不只是「显示难看」：`docx2md_inspect` 的
`missing_preview_context[].paragraph_text` 是判断「缺图公式属于图注还是正文」的
唯一依据，乱码等于这个判断失效。

三层防御：

| 层 | 做法 | 位置 |
|---|---|---|
| 进程层 | 给子进程注入 `PYTHONUTF8=1` / `PYTHONIOENCODING=utf-8` | `lib/index.js` `pyEnv()` |
| 脚本层 | 每个脚本入口调 `force_utf8_stdio()`，不依赖环境变量 | `py/docxfile.py` |
| 接收层 | 严格 UTF-8 解码，非法字节直接报错而非替换成 `U+FFFD` | `lib/index.js` `decodeStrict()` |

回归测试用 `PYTHONIOENCODING=gbk` 在 Linux 上伪造同一环境（实测该设置会让
`python3` 把「位错」写成 GBK 字节，与 Windows 中文版一致），因此这条路径可在
任意平台复现：

```bash
node test/encoding.mjs                 # 用内联的微型 docx
node test/encoding.mjs 真实文档.docx    # 再验一份中文文档
```

实测基线（145 公式的论文）：

| 阶段 | 耗时 |
|---|---|
| inspect | ~0.4 s |
| extract | ~13 s（含 127 张图渲染） |
| build + validate | ~0.7 s |

---

## 性能实测

在 Windows 桌面版 DSH + Python 3.13 + uvx 兜底的组合下：

| 操作 | 输出到 Windows 本地 | 输出到 WSL 路径 |
|---|---|---|
| inspect（145 公式） | 5.2 s | — |
| extract + 渲染 127 张图 | **4.9 s** | 9.9 s |
| build + validate | 0.75 s | — |

**把工作目录放在 Windows 本地比放在 WSL 路径快一倍**（GDI+ 经 UNC 写图有固定开销）。
文档本身放 WSL 没问题，但 `outDir` 建议给个 Windows 本地路径。

---

## 已知限制

- **EMF 内嵌 PDF 文字层未实现**。「三种手段」一节里提到这条路径（★★★★，
  变体 B 的解法），但它在原始实验里是手工做的一步，没有固化进本插件；`pypdf`
  目前只是预留依赖，`py/` 下零引用。因此 15/145 的变体 B 公式只能靠渲染图读，
  读不清时 `extract` 的 `summary.formulas_without_media` 会报出数量，不会假装成功。
  详见 [docs/methodology.md](docs/methodology.md) 开头的说明。

- **只吃 `.docx`**。旧 `.doc` 请先用 Word / WPS 另存。非 zip、空 zip、缺
  `word/document.xml`、加密文档都会给出可读错误并以退出码 3 结束。
  **Python 脚本失败时不会留下空目录**（文档校验在建目录之前）；但工具层
  `docx2md_extract` 会先按 `outDir` 建目录再调用脚本，所以经工具调用时可能
  留下一个空目录——里面不会有任何产物。
- **同一 outDir 会被清理**。换文档复用目录时，旧的 `paras.json`/`media/`/`png/`
  以及**人工确认过的 `headings.json`** 都会被删除——段落号对另一份文档没有意义，
  留着会让新文档的标题挂到错误的段落上。
- **无预览图的公式无法自动转换**（仅 OLE 族；OMML 从 XML 转出，不需要预览图）。
  `inspect` 会报出数量，需要看原 Word 手工补。
- **OMML 转换覆盖常用结构，未知结构降级拼接**。分式、上下标、根式、定界符、
  大运算符、函数、重音、横线、矩阵、方程组、underset/overset、纯文本样式、
  希腊字母与常用符号（22 类逐一测试）。遇到不认识的节点时拼接其子节点内容——
  宁可少一层装饰，绝不丢内容、绝不抛异常。降级发生时该公式仍会出现在产物里，
  但可能缺外层结构；对关键公式建议抽查。
- **图片公式无法自动转换，但绝不静默丢**。截图/扫描进来的公式在 docx 里
  就是一张位图，像素没有结构可解。工具做的是：按形状启发式（扁条状行内图：
  宽高比≥2.5 且高≤120px）识别**候选**、生成 `⟦EQ:img:文件名⟧` 占位、把原件
  抽进 `media/`；Agent 读图转写后以 `"img:文件名"` 为键回填 LaTeX。什么都不做
  也不丢——build 会落成可见的图片引用（不是藏进注释），校验器 T11 保证引用
  不断链。候选是启发式不是结论：行内小图标也可能被算进来，
  `image_inventory.json` 带 ratio/尺寸/上下文供人逐个过目。
- **预览图必须按 ShapeID 配对，不能按顺序**。公式的预览图是带 `o:ole="t"` 的
  `v:shape`，其 `id` 等于 `OLEObject` 的 `ShapeID`。早期实现按「同段落第 k 个
  公式配第 k 张图」，在公式前面夹着插图的段落里会整体错位，把插图位图当成公式
  预览图交给模型去读（实测 145 个公式里错 10 个，且不报错）。现在优先按
  `ShapeID` 配对；只有文档不写 `ShapeID` 时才退回下标配对，并在
  `summary.formulas_paired_by_ordinal` 与 `pairing_warning` 里明确报出来。
- **标题层级需要人工确认**。中文文档 `1．` / `A）` / `i）` / 无编号「引言」混用是常态，
  自动识别只作为初稿。
- **插图不在处理范围**。文档里的示意图需要在 Markdown 里留 `<!-- 图 N -->` 占位，
  由用户提供图片文件。不要试图用文字描述重画——实测 7 轮全失败。
- **只保证结构合规，不保证每个渲染器都认**。`docx2md_build` 的 PASS 意味着
  六条铁律与 T9/T10 都过了，但不同 Markdown 预览器对数学语法的支持差异很大。
  交付前**务必在目标渲染器里打开看一次**，尤其是长公式与矩阵。

  | 现象 | 原因 | 查什么 |
  |---|---|---|
  | 公式显示成 LaTeX 源码 / 整块空白 | `$` 与中文紧贴，渲染器没识别成公式 | `validation.warnings` 里的 **T10** |
  | 产物里出现 `![图片公式 …]` | 该公式在 Word 里就是张图，未转写（不丢机制在起作用） | `build` 返回的 `image_placeholders` |
  | 图片公式引用是破图 | 引用目标文件不存在 | `validation.errors` 里的 **T11** |
  | 显示成 `\begin{aligned}` 原文 | 块级环境写进了行内公式 | **T2** |
  | 只有长的、矩阵的坏 | 同上，或该渲染器不支持某些命令 | T2 + 该渲染器的 KaTeX 版本 |

  > 实测教训：`katex.renderToString` 能渲染，**不等于** Markdown 解析器会把
  > 它识别成公式。这是两层独立判定，必须分开测。

---

## 目录结构

```
dsh-docx2md/
├── lib/
│   ├── index.js       插件主体：注册 3 个工具 + 1 个技能
│   └── toolkit.js     defineTool 的零依赖实现
├── py/
│   ├── docxfile.py    .docx 打开与结构校验（坏输入转可读错误）
│   ├── docx_scan.py   只读扫描（OLE / OMML / 图片候选三类分开计数）
│   ├── extract.py     段落 + 公式 + 预览图（OMML 在此转出 LaTeX，图片公式落占位）
│   ├── omml.py        OMML → LaTeX 确定性转换器
│   ├── imgformula.py  图片公式候选检测（形状启发式 + OLE 预览排除）
│   ├── render.py      跨平台 WMF/EMF → PNG
│   ├── build.py       段落 + LaTeX 表 → Markdown（OMML 自动预填）
│   └── validate.py    六条铁律 + T9/T10/T11 校验
├── skills/docx2md/
│   └── SKILL.md       技能说明（给 Agent 读）
├── docs/
│   └── methodology.md 完整方法论（转换实验总结）
├── test/              23 组测试（21 日常 + 2 发布前；不进发布包）
├── cordis.patch.yml   插件注册
└── CHANGELOG.md
```

`lib/` 只用 Node 内建模块。这不是洁癖——第三方插件装在 profile 的
`node_modules` 下，而 DSH 自身的包在 `app.asar` 里，标准 Node 解析不到
`@deepseek-ai/*`。详见 `lib/toolkit.js` 顶部注释。

---

## License

MIT
