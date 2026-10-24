/**
 * toolkit.js —— 不依赖任何 @deepseek-ai/* 包的最小工具构造器。
 *
 * 为什么自己实现而不用 `@deepseek-ai/dsh-tools` 的 defineTool：
 *   第三方插件以 npm 包形式装在 profile 的 node_modules 下，而 DSH 自身的
 *   包在 asar 里。标准 Node 的模块解析从插件真实路径向上找，找不到那些包，
 *   于是 `import { defineTool } from '@deepseek-ai/dsh-tools'` 会直接
 *   ERR_MODULE_NOT_FOUND，插件连加载都失败。
 *
 *   宿主确实有一套 internal import 重载，但它作用于插件入口的 *模块说明符*，
 *   不改变插件内部裸导入的解析规则。因此把依赖降到零是最稳的做法。
 *
 * 这里复刻 defineTool 的可观察行为：
 *   - 把简化 schema（{ key: { type, required, description, items } }）转成
 *     JSON Schema
 *   - 执行前按必填项与类型做校验，违规时抛带明细的错误
 *   - 结果对象原样返回（宿主自己序列化）
 */

/** 简化 schema → JSON Schema 对象。 */
export function toJsonSchema(params) {
  const properties = {}
  const required = []
  for (const [key, spec] of Object.entries(params ?? {})) {
    properties[key] = toSchemaNode(spec)
    if (spec?.required === true) required.push(key)
  }
  const schema = { type: 'object', additionalProperties: false, properties }
  if (required.length) schema.required = required
  return schema
}

function toSchemaNode(spec) {
  if (!spec || typeof spec !== 'object') return {}
  const node = {}
  if (spec.type) node.type = spec.type
  if (spec.description) node.description = spec.description
  if (spec.enum) node.enum = spec.enum
  if (spec.items) node.items = toSchemaNode(spec.items)
  if (spec.properties) {
    node.type = node.type ?? 'object'
    node.additionalProperties = false
    node.properties = {}
    for (const [k, v] of Object.entries(spec.properties)) {
      node.properties[k] = toSchemaNode(v)
      if (v?.required === true) {
        node.required = node.required ?? []
        node.required.push(k)
      }
    }
  }
  return node
}

/** 按 JSON Schema 的顶层必填与基本类型校验参数。 */
export function validateArgs(schema, args) {
  const violations = []
  if (args === null || typeof args !== 'object' || Array.isArray(args)) {
    return ['arguments must be an object']
  }
  for (const key of schema.required ?? []) {
    if (args[key] === undefined) violations.push(`missing required property "${key}"`)
  }
  for (const [key, node] of Object.entries(schema.properties ?? {})) {
    const value = args[key]
    if (value === undefined) continue
    const want = node.type
    if (!want) continue
    const ok = Array.isArray(want)
      ? want.some((t) => typeOk(t, value))
      : typeOk(want, value)
    if (!ok) violations.push(`property "${key}" must be ${Array.isArray(want) ? want.join('|') : want}`)
  }
  return violations
}

function typeOk(type, value) {
  switch (type) {
    case 'string': return typeof value === 'string'
    case 'number': return typeof value === 'number' && Number.isFinite(value)
    case 'integer': return Number.isInteger(value)
    case 'boolean': return typeof value === 'boolean'
    case 'array': return Array.isArray(value)
    case 'object': return value !== null && typeof value === 'object' && !Array.isArray(value)
    case 'null': return value === null
    default: return true
  }
}

/**
 * 定义一个可注册的工具。
 * @param {object} options
 * @param {string} options.name
 * @param {string} options.description
 * @param {object} options.parameters  简化 schema
 * @param {(args: object, exec: object) => Promise<any>} options.execute
 * @param {(args: object, value: any) => Array<{type: string, text: string}>} [options.render]
 */
export function defineTool(options) {
  const schema = toJsonSchema(options.parameters)
  return {
    name: options.name,
    description: options.description,
    parameters: schema,
    output: {
      schema: { type: 'object', additionalProperties: true, properties: {} },
      render: options.render
        ?? ((_args, value) => [{ type: 'text', text: stringify(value) }]),
    },
    async execute(args, exec) {
      const violations = validateArgs(schema, args)
      if (violations.length) {
        throw new Error(`${options.name}: ${violations.join('; ')}`)
      }
      return options.execute(args, exec)
    },
  }
}

function stringify(value) {
  if (typeof value === 'string') return value
  try { return JSON.stringify(value, null, 1) } catch { return String(value) }
}

/**
 * 极简 Config schema，实现 Standard Schema v1 的最小面。
 *
 * 为什么必须是这个形状：cordis 的 resolveConfig 直接读
 *   runtime.Config["~standard"].validate(config)
 * 没有 ~standard 就抛 TypeError，插件在加载阶段就崩，且是启动期失败、
 * 不产生任何工具——用户只会看到"插件没生效"。
 *
 * validate 的职责：把缺省值填上，并把明显的非法值按默认值纠正，
 * 返回 { value }（成功）或 { issues }（失败）。
 */
export function config(defaults, check) {
  return {
    '~standard': {
      version: 1,
      vendor: 'dsh-docx2md',
      validate(raw) {
        const input = raw && typeof raw === 'object' && !Array.isArray(raw) ? raw : {}
        const value = {}
        for (const [k, fallback] of Object.entries(defaults)) {
          value[k] = input[k] === undefined ? fallback : input[k]
        }
        if (check) {
          const issues = check(value)
          if (issues.length) {
            return { issues: issues.map((message) => ({ message, path: [] })) }
          }
        }
        return { value }
      },
    },
  }
}
