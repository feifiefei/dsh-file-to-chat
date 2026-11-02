// 回归用例：injectText 用哪一份「会话上下文」。
//
// 背景（2026-11-02 实测）：在聊天里划词点「批注」→ 按 Enter → 输入框空着、
// 聊天里也没东西。原因是 injectText 直接信了 `bridge.sessionId`：右栏一关它
// 就冻在上一个会话上，`sessions.scope()` 拿它取不到会话 → chip 跳过、座位桥
// 不存在、最后拿同一个过期 id 抛错。修法与 1.1.3 的粘贴折叠一致：
// **候选逐个验证，能取出会话的才算数**。
//
// 这里盯的就是这段挑选逻辑（resolveScoped）本身。
const fs = require('fs')
const path = require('path')
const src = fs.readFileSync(path.join(__dirname, '..', 'client.js'), 'utf8')

function extract(marker) {
  const i = src.indexOf(marker)
  if (i < 0) throw new Error('marker not found: ' + marker)
  let depth = 0
  for (let k = src.indexOf('{', i); k < src.length; k++) {
    if (src[k] === '{') depth++
    else if (src[k] === '}') { depth--; if (depth === 0) return src.slice(i, k + 1) }
  }
  throw new Error('unbalanced: ' + marker)
}

const code = [
  'var G = { sessionId: "", readSessionId: null }',
  'var bridge = { sessionId: null }',
  extract('function sessionCandidates()'),
  extract('function resolveScoped(ctx, prefer)'),
  'return {',
  '  resolveScoped: resolveScoped,',
  '  sessionCandidates: sessionCandidates,',
  '  G: G, bridge: bridge,',
  '}',
].join('\n')

const api = new Function(code)()

let pass = 0, fail = 0
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('  ok  ' + name) }
  else { fail++; console.log('  FAIL ' + name + (extra === undefined ? '' : '  → ' + JSON.stringify(extra))) }
}

function reset() {
  api.G.sessionId = ''
  api.G.readSessionId = null
  api.bridge.sessionId = null
}

/** 造一个假 ctx：只有 known 里的 id 能拿到 scope；记下每次被问到的 id。 */
function ctxOf(known, opts) {
  opts = opts || {}
  const asked = []
  const ctx = {
    sessions: {
      scope: function (id) {
        asked.push(id)
        if (opts.throwsOn === id) throw new Error('boom')
        if (known.indexOf(id) < 0) return undefined
        return { id: id }
      },
    },
  }
  return { ctx, asked }
}

const LONG_A = 'aaaaaaaa-1111-2222-3333-444444444444'
const LONG_B = 'bbbbbbbb-1111-2222-3333-444444444444'
const LONG_C = 'cccccccc-1111-2222-3333-444444444444'

// ① 活读取器给出的 id 有效 → 直接用它
{
  reset()
  const f = ctxOf([LONG_A])
  const r = api.resolveScoped(f.ctx, function () { return LONG_A })
  ok('① 用活读取器的会话', r.id === LONG_A && r.scoped !== null, r)
  ok('① 只问了一次', f.asked.length === 1 && f.asked[0] === LONG_A, f.asked)
}

// ② prefer 是字符串（引用篮那条路径传的就是现成的 id）
{
  reset()
  const f = ctxOf([LONG_B])
  const r = api.resolveScoped(f.ctx, LONG_B)
  ok('② 字符串 id 也能用', r.id === LONG_B, r)
}

// ③ prefer 过期 → 自动换到下一个候选（座位桥上的活 id）
{
  reset()
  api.bridge.sessionId = LONG_C
  const f = ctxOf([LONG_C])
  const r = api.resolveScoped(f.ctx, function () { return LONG_A })
  ok('③ 过期 id 被跳过', r.id === LONG_C && r.scoped !== null, r)
  ok('③ 先试过期那个再试活的', f.asked[0] === LONG_A && f.asked[1] === LONG_C, f.asked)
  ok('③ tried 记下被跳过的前缀', r.tried.length === 1 && r.tried[0] === LONG_A.slice(0, 8), r.tried)
}

// ④ 全过期 → 明确返回「没有」，并留下试过谁
{
  reset()
  api.bridge.sessionId = LONG_B
  api.G.sessionId = LONG_C
  const f = ctxOf([])
  const r = api.resolveScoped(f.ctx, function () { return LONG_A })
  ok('④ 一个都不成立', r.id === '' && r.scoped === null, r)
  ok('④ tried 有三个前缀', r.tried.length === 3, r.tried)
  ok('④ 每个都真去问过', f.asked.length === 3, f.asked)
}

// ⑤ 某个候选让 scope() 抛错，不影响后面的候选
{
  reset()
  api.bridge.sessionId = LONG_B
  const f = ctxOf([LONG_B], { throwsOn: LONG_A })
  const r = api.resolveScoped(f.ctx, function () { return LONG_A })
  ok('⑤ 抛错的那个被跳过', r.id === LONG_B, r)
  ok('⑤ 抛错也算试过', r.tried[0] === LONG_A.slice(0, 8), r.tried)
}

// ⑥ 同一个 id 不重复验证
{
  reset()
  api.bridge.sessionId = LONG_C
  api.G.sessionId = LONG_C
  const f = ctxOf([LONG_C])
  const r = api.resolveScoped(f.ctx, function () { return LONG_C })
  ok('⑥ 去重后只问一次', f.asked.length === 1 && r.id === LONG_C, f.asked)
}

// ⑦ 一个候选都没有
{
  reset()
  const f = ctxOf([])
  const r = api.resolveScoped(f.ctx, '')
  ok('⑦ 没候选时返回空', r.id === '' && r.scoped === null && r.tried.length === 0, r)
}

// ⑧ prefer 自己抛错（会话正在切换）时不能把整个流程带崩
{
  reset()
  api.G.sessionId = LONG_C
  const f = ctxOf([LONG_C])
  let r = null
  let threw = false
  try {
    r = api.resolveScoped(f.ctx, function () { throw new Error('switching') })
  } catch (err) { threw = true }
  ok('⑧ 读取器抛错不致命', threw === false && r !== null && r.id === LONG_C, r)
}

// ⑨ 返回的 scoped 就是宿主给的那个对象（不做包装，调用方直接拿去用）
{
  reset()
  const only = { id: LONG_A }
  const ctx = { sessions: { scope: function () { return only } } }
  const r = api.resolveScoped(ctx, LONG_A)
  ok('⑨ scoped 原样返回', r.scoped === only, r)
}

console.log('\ninject: ' + pass + ' passed, ' + fail + ' failed')
if (fail > 0) process.exit(1)
