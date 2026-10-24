#!/usr/bin/env node
/**
 * 按 @deepseek-ai/dsh-tools 的真实 schema 校验子集复刻，验证本插件注册的
 * 三个工具定义能否被 tools.register() 接受。
 *
 * 规则逐条抄自 dsh-tools/lib/index.js：
 *   SCHEMA_TYPES / CONSTRAINT_KEYWORDS / checkSchemaNode / checkObjectSchemaTail
 *   register(definition) 的前置检查
 *
 * 用法: node test/tool-contract.mjs
 */
const SCHEMA_TYPES = ['object', 'array', 'string', 'number', 'integer', 'boolean', 'null']
const CONSTRAINT_KEYWORDS = new Set(['type', 'oneOf', 'properties', 'required', 'additionalProperties', 'items', 'enum', 'const'])
const ANNOTATION_KEYWORDS = new Set(['description', 'title', 'default', 'examples', 'deprecated', 'readOnly', 'writeOnly'])

const isObj = (v) => typeof v === 'object' && v !== null && !Array.isArray(v)
const isJsonValue = (v) => {
  if (v === null || typeof v === 'string' || typeof v === 'boolean') return true
  if (typeof v === 'number') return Number.isFinite(v)
  if (Array.isArray(v)) return v.every(isJsonValue)
  if (isObj(v)) return Object.values(v).every(isJsonValue)
  return false
}

function checkSchemaNode(root, rootPath, violations) {
  const stack = [{ node: root, path: rootPath }]
  const seen = new Set()
  while (stack.length) {
    const { node, path } = stack.pop()
    if (!isObj(node)) { violations.push(`${path} must be a schema object`); continue }
    if (seen.has(node)) { violations.push(`${path} is circular`); continue }
    seen.add(node)

    for (const key of Object.keys(node)) {
      if (CONSTRAINT_KEYWORDS.has(key)) continue
      if (ANNOTATION_KEYWORDS.has(key)) {
        if (!isJsonValue(node[key])) violations.push(`${path}.${key} annotation must be lossless JSON data`)
        continue
      }
      violations.push(`${path}.${key} is not a supported keyword`)
    }
    if (node.description !== undefined && typeof node.description !== 'string') {
      violations.push(`${path}.description must be a string`)
    }
    const hasType = node.type !== undefined
    const hasOneOf = node.oneOf !== undefined
    if (hasType && hasOneOf) { violations.push(`${path} cannot declare both type and oneOf`); continue }
    if (!hasType && !hasOneOf) { continue }
    if (hasOneOf) continue

    const type = node.type
    if (typeof type !== 'string' || !SCHEMA_TYPES.includes(type)) {
      violations.push(`${path}.type must be one of ${SCHEMA_TYPES.join('/')}`)
      continue
    }
    for (const [key, types] of Object.entries({
      properties: ['object'], required: ['object'], additionalProperties: ['object'],
      items: ['array'], enum: ['string', 'number', 'integer', 'boolean', 'null'],
      const: ['string', 'number', 'integer', 'boolean', 'null'],
    })) {
      if (node[key] !== undefined && !types.includes(type)) {
        violations.push(`${path}.${key} is not supported on type "${type}"`)
      }
    }

    if (type === 'object') {
      if (node.additionalProperties !== undefined && typeof node.additionalProperties !== 'boolean') {
        violations.push(`${path}.additionalProperties must be a boolean`)
      }
      if (node.properties !== undefined) {
        if (!isObj(node.properties)) violations.push(`${path}.properties must be an object of schemas`)
        else for (const [k, v] of Object.entries(node.properties)) {
          stack.push({ node: v, path: `${path}.properties.${k}` })
        }
      }
      if (node.required !== undefined) {
        if (!Array.isArray(node.required) || !node.required.every((r) => typeof r === 'string')) {
          violations.push(`${path}.required must be an array of strings`)
        }
      }
    } else if (type === 'array') {
      if (node.items !== undefined) stack.push({ node: node.items, path: `${path}.items` })
    }
  }
}

function assertSupportedJsonSchema(schema) {
  const violations = []
  checkSchemaNode(schema, 'schema', violations)
  if (violations.length) throw new Error(`unsupported JSON schema: ${violations.join('; ')}`)
}

function assertObjectJsonSchema(schema) {
  if (schema.type !== 'object') throw new Error('schema root must be type "object"')
  assertSupportedJsonSchema(schema)
}

// ---- 取真实工具定义 ----
const mod = await import('../lib/index.js')
const registered = []
const ctx = {
  logger: { warn() {}, info() {} },
  tools: { register(d) { registered.push(d); return () => {} } },
  skills: { registerProvider() {} },
  workingDirectory: { async ensure() { return process.cwd() } },
}
mod.apply(ctx, {})

// ---- Config 契约：cordis 的 resolveConfig 逐行复刻 ----
// 原样抄自 cordis/lib/index.js:956-962
function cordisResolveConfig(runtime, config) {
  if (!runtime.Config) return config
  const result = runtime.Config['~standard'].validate(config)
  if ('then' in result) throw new TypeError('Async config validation is not supported')
  if (result.issues) throw new Error('ValidationError: ' + JSON.stringify(result.issues))
  return result.value
}

const checks = []
const rec = (name, fn) => {
  try { fn(); checks.push([name, true, '']) } catch (e) { checks.push([name, false, e.message.slice(0, 160)]) }
}

rec('注册了 3 个工具', () => {
  if (registered.length !== 3) throw new Error(`实际 ${registered.length} 个`)
})

for (const t of registered) {
  rec(`${t.name}: output 形状符合 register() 要求`, () => {
    const o = t.output
    if (!isObj(o) || typeof o.render !== 'function') {
      throw new Error('必须声明 output { schema, render }')
    }
  })
  rec(`${t.name}: output.schema 在支持子集内`, () => {
    assertSupportedJsonSchema(t.output.schema)
  })
  rec(`${t.name}: parameters 在支持子集内`, () => {
    assertSupportedJsonSchema(t.parameters)
  })
  rec(`${t.name}: parameters 根为 object 且 required 是数组`, () => {
    assertObjectJsonSchema(t.parameters)
    if (t.parameters.required !== undefined && !Array.isArray(t.parameters.required)) {
      throw new Error('required 必须是数组')
    }
  })
  rec(`${t.name}: 每个属性的 required 标记已提升为数组`, () => {
    for (const [k, v] of Object.entries(t.parameters.properties ?? {})) {
      if (Object.hasOwn(v, 'required')) {
        throw new Error(`属性 ${k} 残留了属性级 required，DSH 子集不支持`)
      }
    }
  })
  rec(`${t.name}: description 非空`, () => {
    if (typeof t.description !== 'string' || !t.description.trim()) throw new Error('description 为空')
  })
}

rec('Config 实现了 Standard Schema（cordis 硬性要求）', () => {
  const C = mod.Config
  if (!C || typeof C !== 'object') throw new Error('Config 必须是对象')
  const std = C['~standard']
  if (!std || typeof std.validate !== 'function') {
    throw new Error('缺少 Config["~standard"].validate —— cordis 会抛 TypeError')
  }
})
rec('Config 能通过 cordis resolveConfig（空配置）', () => {
  const v = cordisResolveConfig(mod, {})
  if (typeof v !== 'object' || v === null) throw new Error('resolve 结果不是对象')
  if (!Number.isFinite(v.maxImageWidth)) throw new Error('maxImageWidth 未填默认值')
})
rec('Config 能通过 cordis resolveConfig（部分覆盖）', () => {
  const v = cordisResolveConfig(mod, { maxImageWidth: 800 })
  if (v.maxImageWidth !== 800) throw new Error('用户配置未生效')
  if (v.allowUvx !== true) throw new Error('未覆盖项未填默认值')
})
rec('Config 拒绝非法值', () => {
  let threw = false
  try { cordisResolveConfig(mod, { maxImageWidth: 1 }) } catch { threw = true }
  if (!threw) throw new Error('非法 maxImageWidth 未被拦截')
})
rec('Config resolve 是同步的（cordis 不支持 async）', () => {
  const r = mod.Config['~standard'].validate({})
  if (r && typeof r.then === 'function') throw new Error('validate 返回了 Promise')
})

// 排版阈值是用户可配的。这两个参数一旦从 parameters 里消失，用户就没法
// 在不改全局配置的情况下针对单篇文档调排版——要钉住。
rec('docx2md_build 暴露了 longTexChars / t5MaxCellFormula', () => {
  const b = registered.find((t) => t.name === 'docx2md_build')
  if (!b) throw new Error('找不到 docx2md_build')
  for (const k of ['longTexChars', 't5MaxCellFormula']) {
    const node = b.parameters.properties?.[k]
    if (!node) throw new Error(`parameters 里缺 ${k}`)
    if (node.type !== 'integer') throw new Error(`${k} 应为 integer，实际 ${node.type}`)
  }
})
rec('Config 填充排版阈值默认值', () => {
  const v = cordisResolveConfig(mod, {})
  if (v.longTexChars !== 80) throw new Error(`longTexChars 默认应为 80，实际 ${v.longTexChars}`)
  if (v.t5MaxCellFormula !== 60) throw new Error(`t5MaxCellFormula 默认应为 60，实际 ${v.t5MaxCellFormula}`)
})
rec('Config 拒绝非法阈值（0 / 小数 / 负值）', () => {
  for (const bad of [{ longTexChars: 0 }, { longTexChars: 1.5 }, { t5MaxCellFormula: 0 }]) {
    let threw = false
    try { cordisResolveConfig(mod, bad) } catch { threw = true }
    if (!threw) throw new Error(`${JSON.stringify(bad)} 未被拦截`)
  }
})
rec('Config 接受合法阈值覆盖', () => {
  const v = cordisResolveConfig(mod, { longTexChars: 150, t5MaxCellFormula: 200 })
  if (v.longTexChars !== 150) throw new Error('longTexChars 未生效')
  if (v.t5MaxCellFormula !== 200) throw new Error('t5MaxCellFormula 未生效')
})

console.log('== tool 契约测试 ==')
let fail = 0
for (const [name, ok, msg] of checks) {
  console.log((ok ? '  PASS  ' : '  FAIL  ') + name + (msg ? ' — ' + msg : ''))
  if (!ok) fail += 1
}
console.log(fail ? `\n${fail} 项失败` : '\n全部通过')
process.exit(fail ? 1 : 0)
