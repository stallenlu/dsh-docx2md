/**
 * dsh-docx2md —— 把含 OLE 公式的 Word 文档转成带 LaTeX 的 Markdown。
 *
 * 三个模型侧工具：
 *   docx2md_inspect  只读扫描：段落数、公式数、变体分布、渲染能力
 *   docx2md_extract  提取正文 + 公式索引 + 预览图，并把图渲染成 PNG
 *   docx2md_build    用给定的 LaTeX 表生成 Markdown 并校验
 *
 * 为什么拆三步：中间的「读图转写 LaTeX」不可能自动化——它需要视觉模型
 * 逐张看图。拆开后 Agent 可以在 extract 与 build 之间插入自己的读图逻辑，
 * 而不必把上百张图塞进一次工具调用。
 *
 * 平台支持：Windows / WSL / Linux / macOS。渲染引擎自动探测：
 *   Windows/WSL → GDI+（原生支持 WMF/EMF）
 *   Linux/macOS → LibreOffice
 *   都不可用    → 显式失败并给出可操作提示（绝不回退到 MTEF 字符解码）
 *
 * 依赖：本文件与 toolkit.js 只用 Node 内建模块。原因见 toolkit.js 顶部注释。
 *
 * @module dsh-docx2md
 */
import { spawn } from 'node:child_process'
import { existsSync } from 'node:fs'
import { mkdir, readFile, rm } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { config, defineTool } from './toolkit.js'

/** Cordis 插件标识。 */
export const name = 'docx2md'

/** 依赖的宿主服务。 */
export const inject = ['tools', 'skills', 'workingDirectory']

/** 插件配置。必须是 Standard Schema——cordis 读 Config["~standard"].validate。 */
export const Config = config(
  {
    maxImageWidth: 1600,
    allowUvx: true,
    python: '',
    timeoutMs: 600000,
    longTexChars: 80,
    t5MaxCellFormula: 60,
  },
  (v) => {
    const issues = []
    if (!Number.isFinite(v.maxImageWidth) || v.maxImageWidth < 200) {
      issues.push('maxImageWidth must be a number >= 200')
    }
    if (!Number.isFinite(v.timeoutMs) || v.timeoutMs < 1000) {
      issues.push('timeoutMs must be a number >= 1000')
    }
    if (typeof v.allowUvx !== 'boolean') issues.push('allowUvx must be a boolean')
    if (typeof v.python !== 'string') issues.push('python must be a string')
    // 排版类阈值：必须为正整数。给 0 或负数会让每条公式都独立成块/都报警，
    // 是配置失误而不是有意义的取值。
    if (!Number.isInteger(v.longTexChars) || v.longTexChars < 1) {
      issues.push('longTexChars must be an integer >= 1')
    }
    if (!Number.isInteger(v.t5MaxCellFormula) || v.t5MaxCellFormula < 1) {
      issues.push('t5MaxCellFormula must be an integer >= 1')
    }
    return issues
  },
)

/** 包内资源根目录（lib/ 的上一级）。 */
const PKG_ROOT = fileURLToPath(new URL('../', import.meta.url))
const PY_DIR = join(PKG_ROOT, 'py')
const SKILL_DIR = join(PKG_ROOT, 'skills')

/** 三份 Python 脚本用到的第三方包。缺了就经 uvx 临时装。 */
const PY_DEPS = ['olefile', 'pypdf']

/** 合并用户配置与默认值。 */
function resolveConfig(raw) {
  const cfg = raw ?? {}
  return {
    maxImageWidth: Number.isFinite(cfg.maxImageWidth) ? cfg.maxImageWidth : 1600,
    allowUvx: cfg.allowUvx !== false,
    python: typeof cfg.python === 'string' ? cfg.python : '',
    timeoutMs: Number.isFinite(cfg.timeoutMs) ? cfg.timeoutMs : 600000,
    longTexChars: Number.isFinite(cfg.longTexChars) ? cfg.longTexChars : 80,
    t5MaxCellFormula: Number.isFinite(cfg.t5MaxCellFormula) ? cfg.t5MaxCellFormula : 60,
  }
}

// ------------------------------------------------------------------ 进程

/**
 * 子进程环境：强制 Python 以 UTF-8 输出。
 *
 * 中文 Windows 上 Python 的 sys.stdout.encoding 是 gbk，非 ASCII 字符会被编成
 * GBK 字节；我们按 UTF-8 解码就得到乱码（且是静默的）。那不只是「显示难看」——
 * missing_preview_context.paragraph_text 是模型判断「缺图公式是图注还是正文」的
 * 唯一依据，乱码等于这个判断失效。UTF-8 模式对 3.7+ 全平台生效，Linux/macOS 上
 * 本来就是 UTF-8，设置它不会改变行为。
 */
function pyEnv() {
  const env = { ...process.env }
  env.PYTHONUTF8 = '1'
  env.PYTHONIOENCODING = 'utf-8'
  // 防止父进程遗留的旧设置反过来把编码钉死。
  delete env.PYTHONLEGACYWINDOWSSTDIO
  return env
}

/** 严格按 UTF-8 解码；不是合法 UTF-8 时返回 ok:false，而不是静默替换成乱码。 */
function decodeStrict(buf) {
  try {
    return { ok: true, text: new TextDecoder('utf-8', { fatal: true }).decode(buf) }
  } catch {
    return { ok: false, text: buf.toString('utf8') }
  }
}

/** 跑一个命令，收 stdout/stderr，超时则杀掉。 */
function run(cmd, args, { cwd, timeoutMs = 600000, env } = {}) {
  return new Promise((done) => {
    let child
    try {
      child = spawn(cmd, args, { cwd, env, stdio: ['ignore', 'pipe', 'pipe'], shell: false })
    } catch (error) {
      done({ code: -1, stdout: '', stderr: String(error) })
      return
    }
    const outChunks = []
    const errChunks = []
    let settled = false
    const finish = (payload) => {
      const o = decodeStrict(Buffer.concat(outChunks))
      const e = decodeStrict(Buffer.concat(errChunks))
      done({
        ...payload,
        stdout: o.text,
        stderr: payload.stderrSuffix ? `${e.text}${payload.stderrSuffix}` : e.text,
        encodingError: !o.ok || !e.ok,
        stderrSuffix: undefined,
      })
    }
    const timer = setTimeout(() => {
      if (settled) return
      settled = true
      try { child.kill('SIGKILL') } catch { /* 已退出 */ }
      finish({ code: -1, stderrSuffix: `[超时 ${timeoutMs}ms]`, timedOut: true })
    }, timeoutMs)

    child.stdout?.on('data', (b) => { outChunks.push(b) })
    child.stderr?.on('data', (b) => { errChunks.push(b) })
    child.on('error', (error) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      finish({ code: -1, stderrSuffix: String(error) })
    })
    child.on('close', (code) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      finish({ code: code ?? -1 })
    })
  })
}

/** 在 PATH 里找可执行文件。 */
async function findOnPath(candidates) {
  const which = process.platform === 'win32' ? 'where' : 'which'
  for (const c of candidates) {
    const r = await run(which, [c], { timeoutMs: 5000 })
    if (r.code === 0 && r.stdout.trim()) return r.stdout.trim().split(/\r?\n/)[0]
  }
  return null
}

/** uvx 的依赖参数：--from 第一个包 --with 其余包（--from 不能重复）。 */
function uvxDeps() {
  return ['--from', PY_DEPS[0], ...PY_DEPS.slice(1).flatMap((d) => ['--with', d])]
}

/**
 * 解析 Python 运行方式。
 * 优先用带依赖的直连解释器；否则经 uvx 临时注入依赖。
 */
async function resolvePython(config) {
  let py = config.python || null
  if (!py) py = await findOnPath(['python3', 'python'])
  const uvx = config.allowUvx ? await findOnPath(['uvx']) : null

  if (py) {
    const probe = await run(py, ['-c', `import ${PY_DEPS.join(', ')}`], { timeoutMs: 20000, env: pyEnv() })
    if (probe.code === 0) return { mode: 'direct', python: py, prefix: [] }
    if (uvx) return { mode: 'uvx', python: py, prefix: uvxDeps(), launcher: uvx }
    return { mode: 'missing-deps', python: py, missing: PY_DEPS, uvx: null }
  }
  if (uvx) return { mode: 'uvx-only', python: null, prefix: uvxDeps(), launcher: uvx }
  return { mode: 'no-python', python: null, uvx: null }
}

/** 按解析结果拼出可执行的 argv。 */
function pythonArgv(resolved, script, args) {
  if (resolved.launcher) {
    return { cmd: resolved.launcher, argv: [...resolved.prefix, 'python', join(PY_DIR, script), ...args] }
  }
  return { cmd: resolved.python, argv: [join(PY_DIR, script), ...args] }
}

/** 从输出里挑出最后一个完整的 JSON 对象或数组。 */
function extractJson(text) {
  if (!text) return null
  const lastBrace = Math.max(text.lastIndexOf('}'), text.lastIndexOf(']'))
  if (lastBrace < 0) return null
  const closeCh = text[lastBrace]
  const openCh = closeCh === '}' ? '{' : '['
  let depth = 0
  let inStr = false
  let esc = false
  for (let i = lastBrace; i >= 0; i -= 1) {
    const ch = text[i]
    if (esc) { esc = false; continue }
    if (ch === '\\') { esc = true; continue }
    if (ch === '"') { inStr = !inStr; continue }
    if (inStr) continue
    if (ch === closeCh) depth += 1
    else if (ch === openCh) {
      depth -= 1
      if (depth === 0) {
        try { return JSON.parse(text.slice(i, lastBrace + 1)) } catch { return null }
      }
    }
  }
  return null
}

/** 统一调用一个 py 脚本并把最后一段 JSON 解析出来。 */
async function runPy(config, resolved, script, args, extra = {}) {
  if (resolved.mode === 'no-python') {
    throw new Error(
      'docx2md: 找不到 Python。请安装 python3（或 uv），或在插件配置里用 python 指定解释器。')
  }
  if (resolved.mode === 'missing-deps') {
    throw new Error(
      `docx2md: Python (${resolved.python}) 缺少依赖 ${resolved.missing.join(', ')}，`
      + '也没找到 uvx 可自动准备。请 `pip install olefile pypdf`，或安装 uv。')
  }
  const { cmd, argv } = pythonArgv(resolved, script, args)
  const r = await run(cmd, argv, { timeoutMs: config.timeoutMs, cwd: extra.cwd, env: pyEnv() })
  const text = (r.stdout || '').trim()
  const parsed = extractJson(text)
  // 输出不是合法 UTF-8 时明确报错。静默替换成 U+FFFD 会让中文段落变成
  // 「����」，那正好破坏了 missing_preview_context 的判断依据。
  if (r.encodingError && parsed === null) {
    throw new Error(
      `docx2md ${script}: 子进程输出不是合法 UTF-8，无法解析 JSON。`
      + '通常是 Python 未按 UTF-8 输出（见 pyEnv）。')
  }
  if (r.code !== 0 && parsed === null) {
    const tail = (r.stderr || text).trim().slice(-1200)
    throw new Error(`docx2md ${script} 失败（exit ${r.code}）：\n${tail}`)
  }
  return { code: r.code, parsed, stdout: text, stderr: (r.stderr || '').trim(), encodingError: r.encodingError }
}

/** WSL 风格的挂载路径：/mnt/c/... → C:\... */
const WSL_MOUNT = /^\/mnt\/([a-zA-Z])(?:[\\/]|$)([\s\S]*)$/
/** Windows 盘符路径：C:\... 或 C:/... */
const WIN_DRIVE = /^([a-zA-Z]):[\\/]/

/**
 * 把工具参数里的路径归一化成绝对路径。
 *
 * 为什么不能「isAbsolute 就原样返回」：
 * 在 Windows 上 path.isAbsolute('/tmp/x') 返回 true（它表示「当前卷根目录下的
 * tmp/x」），于是 '/tmp/x' 会被原样交给 Python，而 Python 的 os.path.abspath
 * 在非 UNC 工作目录下把它解释成 C:\tmp\x —— 用户在 WSL 里说 /tmp，文件却跑到
 * C:\tmp。实测：Node 的 resolve 会正确归一到 \\wsl.localhost\Ubuntu\tmp\x，
 * 只有跳过 resolve 才会出错。所以一律归一化，让 Node 按会话工作目录的卷来解释。
 *
 * 归一化后的语义（Windows 上）：
 *   /tmp/x                     → 会话工作目录所在 Linux 视图的 /tmp
 *   /mnt/c/work/x              → C:\work\x
 *   C:\work\x                  → C:\work\x
 *   \\wsl.localhost\Ubuntu\x   → 原样
 * 即：以 / 开头的按 Linux 视图解释，盘符与 UNC 路径原样使用。
 *
 * 在 Linux/macOS 上反向补一条：C:\work\x 若存在对应挂载则映射成 /mnt/c/work/x，
 * 免得把 Windows 路径当成相对路径拼到工作目录后面。
 */
async function resolvePath(ctx, exec, p) {
  let cwd = process.cwd()
  if (ctx.workingDirectory?.ensure) {
    cwd = await ctx.workingDirectory.ensure(exec?.agent, exec?.signal)
  }

  if (process.platform === 'win32') {
    const m = WSL_MOUNT.exec(p)
    if (m) {
      const drive = `${m[1].toUpperCase()}:\\`
      if (existsSync(drive)) {
        const rest = m[2].split(/[\\/]/).filter(Boolean)
        return resolve(drive, ...rest)
      }
    }
    return resolve(cwd, p)
  }

  const w = WIN_DRIVE.exec(p)
  if (w) {
    const base = `/mnt/${w[1].toLowerCase()}`
    if (existsSync(base)) {
      const rest = p.slice(2).split(/[\\/]/).filter(Boolean)
      return resolve(base, ...rest)
    }
  }
  return resolve(cwd, p)
}

/**
 * 把 Windows 绝对路径反向映射成调用方习惯的写法，用于**返回值**（不是内部使用）。
 *
 * 输入路径接受两种写法，但输出路径只能给一种。宿主是 DSH，它在 Windows 上跑
 * node、在 WSL 里看工作区，所以：
 *   \\wsl.localhost\Ubuntu\tmp\x        → /tmp/x        （WSL 侧可直接 ls）
 *   \\wsl$\Ubuntu\home\sidady\a.docx    → /home/...     （同上）
 *   C:\work\x                            → C:\work\x     （本来就是 Windows 路径）
 *
 * 只在平台是 win32 时做；Linux/macOS 原样返回。
 */
function displayPath(p) {
  if (process.platform !== 'win32') return p
  const m = /^\\\\wsl(?:\.localhost|\$)\\[^\\]+([\s\S]*)$/.exec(p)
  if (!m) return p
  const rest = m[1].split(/\\/).filter(Boolean).join('/')
  return `/${rest}`
}

// ------------------------------------------------------------------ 工具

/**
 * 注册三个模型侧工具与一个技能。
 * @param ctx 携带 tools / skills / workingDirectory 的上下文
 * @param rawConfig 插件配置
 */
export function apply(ctx, rawConfig) {
  const config = resolveConfig(rawConfig)
  if (!existsSync(PY_DIR)) {
    ctx.logger?.warn?.(`docx2md: 找不到 Python 脚本目录 ${PY_DIR}`)
  }

  // ---------------- inspect ----------------
  ctx.tools.register(defineTool({
    name: 'docx2md_inspect',
    description:
      '只读扫描一个 .docx：统计段落数、OLE 公式数、MTEF 变体分布，并探测本机有哪些 '
      + 'WMF/EMF 渲染引擎。用于确认这个文档能不能转、以及要走哪条路。不写任何文件。',
    parameters: {
      path: { type: 'string', required: true, description: '.docx 路径（绝对路径，或相对会话工作目录）。' },
    },
    async execute(args, exec) {
      const src = await resolvePath(ctx, exec, args.path)
      // 先看扩展名再看存在性：.doc / .txt 这类输入应当得到「格式不对」的
      // 准确提示，而不是误导性的「文件不存在」。
      const ext = src.split('.').pop().toLowerCase()
      if (ext === 'doc') {
        throw new Error(
          'docx2md_inspect: 这是旧版 .doc，请先用 Word / WPS 另存为 .docx 再转换。'
          + `路径: ${displayPath(src)}`)
      }
      if (ext !== 'docx') {
        throw new Error(
          `docx2md_inspect: 只接受 .docx（当前是 .${ext}）。路径: ${displayPath(src)}`)
      }
      if (!existsSync(src)) throw new Error(`docx2md_inspect: 文件不存在 ${displayPath(src)}`)
      const resolved = await resolvePython(config)
      const raw = await runPy(config, resolved, 'docx_scan.py', [src])
      return { file: displayPath(src), runtime: resolved.mode, ...(raw.parsed ?? {}) }
    },
  }))

  // ---------------- extract ----------------
  ctx.tools.register(defineTool({
    name: 'docx2md_extract',
    description:
      '提取 .docx 的正文段落与全部 OLE 公式，把公式预览图（EMF/WMF）抽出来并渲染成 PNG。'
      + '产出 paras.json（正文，公式位置为 ⟦EQ:oleObjectN.bin⟧ 占位）、eq_index.json、'
      + 'media/（原始矢量图）、png/（可读位图）。下一步是用视觉逐张读 png/ 转写 LaTeX。',
    parameters: {
      path: { type: 'string', required: true, description: '.docx 路径。' },
      outDir: { type: 'string', description: '工作目录（放中间产物）。默认是 docx 同级的 <名字>_docx2md/。' },
      render: { type: 'boolean', description: '是否渲染 PNG（默认 true）。false 只抽原始 EMF/WMF。' },
    },
    async execute(args, exec) {
      const src = await resolvePath(ctx, exec, args.path)
      if (!existsSync(src)) throw new Error(`docx2md_extract: 文件不存在 ${displayPath(src)}`)

      const outDir = args.outDir
        ? await resolvePath(ctx, exec, args.outDir)
        : `${src.replace(/\.docx$/i, '')}_docx2md`
      await mkdir(outDir, { recursive: true })

      const resolved = await resolvePython(config)
      const ex = await runPy(config, resolved, 'extract.py', [src, outDir])

      const result = {
        outDir: displayPath(outDir),
        runtime: resolved.mode,
        ...(ex.parsed ?? {}),
        mediaDir: displayPath(join(outDir, 'media')),
      }

      if (args.render !== false) {
        try {
          const rn = await runPy(config, resolved, 'render.py',
            [join(outDir, 'media'), join(outDir, 'png'), '--max-width', String(config.maxImageWidth)])
          result.render = rn.parsed
          result.pngDir = displayPath(join(outDir, 'png'))
        } catch (error) {
          // 如实报告失败；不替调用方决定下一步该怎么做。
          result.render = { ok: false, error: String(error).slice(0, 800) }
        }
        // 渲染没成功时，明确说明哪些产物仍可用。
        if (!result.render || result.render.ok !== true) {
          result.renderIncomplete = {
            message: 'PNG 渲染未完成；media/ 中的原始矢量图已保存。',
            mediaDir: displayPath(join(outDir, 'media')),
            availableEngines: '可用 docx2md_inspect 查看本机 detect 到的渲染引擎。',
          }
        }
      }
      return result
    },
  }))

  // ---------------- build ----------------
  ctx.tools.register(defineTool({
    name: 'docx2md_build',
    description:
      '用「公式名 → LaTeX」映射表把提取结果组织成 Markdown，并按六条 Markdown+LaTeX '
      + '铁律校验。可先传 suggestHeadings=true 生成标题初稿，人工修订后再正式构建。',
    parameters: {
      workDir: { type: 'string', required: true, description: 'docx2md_extract 的 outDir。' },
      latex: {
        type: 'string',
        required: true,
        description: '公式映射的 JSON 文件路径，形如 {"oleObject1.bin": "\\\\alpha", ...}。键用 eq_index.json 的 stream 字段。',
      },
      out: { type: 'string', required: true, description: '输出 .md 路径。' },
      suggestHeadings: { type: 'boolean', description: '只生成 headings_suggested.json 供人工修订，不写 Markdown。' },
      useSuggestedHeadings: { type: 'boolean', description: '直接用自动识别的标题（跳过人工确认）。' },
      expectFormulas: { type: 'number', description: '期望的公式总数；给了就校验数量是否吻合。' },
      longTexChars: {
        type: 'integer',
        description: '混排段里 LaTeX 源串超过这个长度就独立成块。不传则用插件配置的长值'
          + '（默认 80）。调大 = 更多公式留在句内、句子更连贯；调小 = 更多独立成块、'
          + '长式更好读也更不易在窄栏溢出。',
      },
      t5MaxCellFormula: {
        type: 'integer',
        description: '表格单元格内公式的字符上限，超过则校验器给 T5 警告。不传则用插件配置'
          + '（默认 60）。',
      },
      cleanup: {
        type: 'boolean',
        description: '构建且校验通过后删除 media/ 与 png/（保留 paras.json、eq_index.json，便于复现）。'
          + '校验未通过时不删，保留现场供排查。',
      },
    },
    async execute(args, exec) {
      const workDir = await resolvePath(ctx, exec, args.workDir)
      const latex = await resolvePath(ctx, exec, args.latex)
      const out = await resolvePath(ctx, exec, args.out)
      if (!existsSync(join(workDir, 'paras.json'))) {
        throw new Error(`docx2md_build: ${workDir} 里没有 paras.json，先跑 docx2md_extract`)
      }
      if (!existsSync(latex)) throw new Error(`docx2md_build: LaTeX 表不存在 ${latex}`)

      const resolved = await resolvePython(config)
      // 阈值优先级：本次调用显式传参 > 插件配置 > Python 侧默认。
      // 显式传参让人可以针对某一篇文档试不同的排版，而不必改全局配置。
      const longTexChars = Number.isInteger(args.longTexChars)
        ? args.longTexChars
        : config.longTexChars
      const t5MaxCellFormula = Number.isInteger(args.t5MaxCellFormula)
        ? args.t5MaxCellFormula
        : config.t5MaxCellFormula
      const argv = [workDir, '--latex', latex, '--out', out,
        '--long-tex-chars', String(longTexChars)]
      if (args.suggestHeadings) argv.push('--suggest')
      if (args.useSuggestedHeadings) argv.push('--use-suggested')

      const b = await runPy(config, resolved, 'build.py', argv)
      const result = { out: displayPath(out), ...(b.parsed ?? {}) }
      if (args.suggestHeadings) return result

      const vargs = [out, '--t5-max-cell-formula', String(t5MaxCellFormula)]
      if (typeof args.expectFormulas === 'number') {
        vargs.push('--expect-formulas', String(args.expectFormulas))
      }
      const v = await runPy(config, resolved, 'validate.py', [...vargs, '--json'])
      // validate.py 回传的是它收到的原始路径（在 Windows 上是 UNC 形式）。
      // 同一份返回里 result.out 已经折成调用方习惯的写法，这里也要一致，
      // 否则使用者会看到同一文件的两个不同路径。
      if (v.parsed && typeof v.parsed.file === 'string') {
        v.parsed.file = displayPath(v.parsed.file)
      }
      result.validation = v.parsed

      // 只在构建完成且校验 PASS 时清理。校验失败就别删——此时 png/ 与 media/
      // 正是排查「哪张图读错了」的唯一依据，删掉等于毁掉现场。
      if (args.cleanup === true) {
        const ok = b.parsed?.unresolved_placeholders === 0
          && (b.parsed?.formulas_missing ?? 0) === 0
          && v.parsed?.status === 'PASS'
        if (!ok) {
          result.cleanup = {
            done: false,
            reason: '构建或校验未通过，保留 media/ 与 png/ 供排查；通过后重跑本工具即会清理。',
          }
        } else {
          const removed = []
          for (const sub of ['media', 'png']) {
            const dir = join(workDir, sub)
            if (!existsSync(dir)) continue
            try {
              await rm(dir, { recursive: true, force: true })
              removed.push(sub)
            } catch (error) {
              result.cleanup = {
                done: false,
                reason: `删除 ${sub}/ 失败: ${String(error).slice(0, 200)}`,
                removed,
              }
            }
          }
          if (!result.cleanup) {
            result.cleanup = {
              done: true,
              removed,
              kept: ['paras.json', 'eq_index.json', 'summary.json'],
              note: 'media/ 与 png/ 已删除；如需重新读图，重跑 docx2md_extract 即可。',
            }
          }
        }
      }
      return result
    },
  }))

  // ---------------- skill ----------------
  registerSkill(ctx)
}

/** 把 skills/docx2md 注册成一个 bundled skill。 */
function registerSkill(ctx) {
  const skillPath = join(SKILL_DIR, 'docx2md', 'SKILL.md')
  if (!existsSync(skillPath)) return
  if (!ctx.skills?.registerProvider) return

  const provider = {
    name: 'dsh-docx2md',
    // 必须是数组或 { candidates, complete }；可同步或 Promise，宿主会 await。
    list: async () => [{
      name: 'docx2md',
      description:
        '把含 OLE 公式的 Word 文档转成带 LaTeX 的 Markdown：公式渲染成 PNG 后逐张读图，'
        + '再合成为符合六条铁律的 Markdown。当用户要求把 .doc/.docx 转 Markdown、'
        + '要求保留公式、或提到「公式转换」「Word 转 md」时使用。',
      invocation: { modelInvocable: true, userInvocable: true },
      provider: 'dsh-docx2md',
      source: 'bundled',
      // rank 越小越优先。600 = BUNDLED_SKILL_RANK（与 DSH 内置技能同级），
      // 用户自己 ~/.dsh/skills 下的同名技能优先级更高，可覆盖本技能。
      rank: 600,
      resourceBase: { kind: 'directory', path: SKILL_DIR.split('\\').join('/') },
      locator: skillPath,
    }],
    // 返回的 definition 必须含 name / description / invocation / source /
    // provider / content；rank 与 locator 不再回传（宿主自己保留）。
    get: async (candidate) => {
      const raw = await readFile(candidate.locator, 'utf8')
      const m = /^---\r?\n([\s\S]*?)\r?\n---(?:\r?\n|$)/u.exec(raw)
      const content = m ? raw.slice(m[0].length).trim() : raw.trim()
      const { locator, rank, ...rest } = candidate
      return { ...rest, content }
    },
  }
  ctx.skills.registerProvider(() => provider)
}
