#!/usr/bin/env node
/**
 * 按 @deepseek-ai/dsh-skill 的真实校验逻辑复刻，验证本插件注册的 skill
 * candidate 与 definition 是否合法。
 *
 * 校验规则逐条抄自 dsh-skill/lib/index.js：
 *   SKILL_NAME = /^[a-z0-9]+(?:-[a-z0-9]+)*$/
 *   validateCandidate(candidate, providerName)
 *   validateDefinition(skill)
 *
 * 用法: node test/skill-contract.mjs
 */
const SKILL_NAME = /^[a-z0-9]+(?:-[a-z0-9]+)*$/

function validateInvocation(invocation, subject) {
  if (invocation === undefined) return
  if (typeof invocation !== 'object' || invocation === null || Array.isArray(invocation)) {
    throw new TypeError(`${subject} with a non-object invocation policy`)
  }
  if (typeof invocation.modelInvocable !== 'boolean') {
    throw new TypeError(`${subject} with a non-boolean invocation.modelInvocable`)
  }
  if (typeof invocation.userInvocable !== 'boolean') {
    throw new TypeError(`${subject} with a non-boolean invocation.userInvocable`)
  }
}

function validateCandidate(candidate, providerName) {
  if (typeof candidate.name !== 'string') throw new TypeError('non-string skill name')
  if (!SKILL_NAME.test(candidate.name)) throw new Error(`invalid skill name "${candidate.name}"`)
  if (typeof candidate.description !== 'string') throw new TypeError('non-string description')
  if (candidate.description.length === 0) throw new Error('without a description')
  validateInvocation(candidate.invocation, 'candidate')
  if (candidate.whenToUse !== undefined && typeof candidate.whenToUse !== 'string') {
    throw new TypeError('non-string whenToUse')
  }
  if (typeof candidate.source !== 'string') throw new TypeError('non-string source')
  if (typeof candidate.rank !== 'number' || !Number.isFinite(candidate.rank)) {
    throw new Error(`skill "${candidate.name}" with an invalid rank`)
  }
  if (typeof candidate.provider !== 'string') throw new TypeError('non-string provider')
  if (candidate.provider !== providerName) {
    throw new Error(`skill "${candidate.name}" for provider "${candidate.provider}"`)
  }
  if (candidate.path !== undefined && typeof candidate.path !== 'string') {
    throw new TypeError('non-string path')
  }
}

function validateDefinition(skill) {
  const { name, description, whenToUse, invocation, source, provider, content, path } = skill
  if (typeof name !== 'string') throw new TypeError('loaded skill name must be a string')
  if (!SKILL_NAME.test(name)) throw new Error(`loaded skill has invalid name "${name}"`)
  if (typeof description !== 'string') throw new TypeError('description must be a string')
  if (description.length === 0) throw new Error('requires a description')
  validateInvocation(invocation, 'loaded skill')
  if (whenToUse !== undefined && typeof whenToUse !== 'string') throw new TypeError('whenToUse must be a string')
  if (typeof source !== 'string') throw new TypeError('source must be a string')
  if (typeof provider !== 'string') throw new TypeError('provider must be a string')
  if (typeof content !== 'string') throw new TypeError('content must be a string')
  if (path !== undefined && typeof path !== 'string') throw new TypeError('path must be a string')
}

// ---- 建 mock ctx 并注册 ----
const mod = await import('../lib/index.js')
let provider

const ctx = {
  logger: { warn() {}, info() {} },
  tools: { register() { return () => {} } },
  skills: {
    registerProvider(create) { provider = create() },
  },
  workingDirectory: { async ensure() { return process.cwd() } },
}
mod.apply(ctx, {})

if (!provider) { console.error('FAIL: 没有注册 skill provider'); process.exit(1) }

const checks = []
const record = (name, fn) => {
  try { fn(); checks.push([name, true, '']) } catch (e) { checks.push([name, false, e.message]) }
}

let candidates
record('provider 有 name', () => {
  if (typeof provider.name !== 'string' || !provider.name) throw new Error('provider.name 缺失')
})
let definitions = []
await (async () => {
  try {
    candidates = await provider.list({})
    if (!Array.isArray(candidates)) throw new Error('list 未返回数组')
    checks.push(['list() 返回数组', true, ''])
  } catch (e) { checks.push(['list() 返回数组', false, e.message]); return }

  try {
    for (const c of candidates) validateCandidate(c, provider.name)
    checks.push(['candidate 通过 validateCandidate', true, ''])
  } catch (e) { checks.push(['candidate 通过 validateCandidate', false, e.message]) }

  try {
    definitions = await Promise.all(candidates.map((c) => provider.get(c, {})))
    for (const d of definitions) validateDefinition(d)
    checks.push(['get() 返回 definition 且通过 validateDefinition', true, ''])
  } catch (e) { checks.push(['get() 返回 definition 且通过 validateDefinition', false, e.message]) }
})()
if (definitions.length === candidates.length) record('definition.name 与 candidate 一致', () => {
  candidates.forEach((c, i) => {
    if (definitions[i].name !== c.name) throw new Error(`name 不一致: ${definitions[i].name} vs ${c.name}`)
  })
})
record('content 非空且已剥离 frontmatter', () => {
  for (const d of definitions) {
    if (!d.content.trim()) throw new Error('content 为空')
    if (d.content.startsWith('---')) throw new Error('frontmatter 未剥离')
  }
})

console.log('== skill 契约测试 ==')
let fail = 0
for (const [name, ok, msg] of checks) {
  console.log((ok ? '  PASS  ' : '  FAIL  ') + name + (msg ? ' — ' + msg : ''))
  if (!ok) fail += 1
}
console.log(fail ? `\n${fail} 项失败` : '\n全部通过')
process.exit(fail ? 1 : 0)
