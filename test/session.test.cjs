// 回归用例：composerShell 的**候选会话**探测。
//
// 背景（2026-10-28 实测）：输入框里粘长文本没折，浮条写着
// 「拿不到 composer(scope-null)」—— 会话 id 有值，只是**过期**的：
// 右栏一关，`bridge.sessionId` 就冻在上一个会话上，`sessions.scope()` 拿
// 它取不到会话。这条路径以前只试一个 id，失败就静默放行。
//
// 现在的契约：候选按新鲜度排（活读取器 → 座位桥 → 缓存），逐个用
// `sessions.scope()` 验证，**能取出会话的那个才算数**。
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
  'var G = { sessionId: "" }',
  'var bridge = { sessionId: null }',
  'var COMPOSER_WHY = "-"',
  'var LAST_SCOPED = null',
  extract('function sessionCandidates()'),
  extract('function composerShell(ctx)'),
  'return {',
  '  sessionCandidates: sessionCandidates,',
  '  composerShell: composerShell,',
  '  G: G, bridge: bridge,',
  '  why: function () { return COMPOSER_WHY },',
  '  scoped: function () { return LAST_SCOPED },',
  '  resetScoped: function () { LAST_SCOPED = null },',
  '}',
].join('\n')

const api = new Function(code)()

let pass = 0, fail = 0
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('  ok  ' + name) }
  else { fail++; console.log('  FAIL ' + name + (extra === undefined ? '' : '  → ' + JSON.stringify(extra))) }
}

/** 造一个假 ctx：只有 known 里的 id 能拿到 scope。 */
function ctxOf(known, opts) {
  opts = opts || {}
  const scopes = []
  const ctx = {
    sessions: {
      scope: function (id) {
        scopes.push(id)
        if (opts.throwsOn === id) throw new Error('boom')
        if (known.indexOf(id) < 0) return undefined
        return { id: id }
      },
    },
    conversation: {
      input: {
        for: function (scoped) {
          return opts.noInsert === true
            ? { state: {}, caretSpan: function () { return null } }
            : { insertReference: function () { return true }, scoped: scoped }
        },
      },
    },
  }
  ctx.scopes = scopes
  return ctx
}

function reset(readerId, bridgeId, cachedId) {
  api.resetScoped()
  api.G.sessionId = cachedId === undefined ? '' : cachedId
  api.bridge.sessionId = bridgeId === undefined ? null : bridgeId
  delete api.G.readSessionId
  if (readerId !== undefined) {
    api.G.readSessionId = function () {
      if (readerId === '#throw') throw new Error('switching')
      return readerId
    }
  }
}

console.log('— 候选顺序与去重 —')
{
  reset('live-1', 'bridge-1', 'cached-1')
  const ids = api.sessionCandidates()
  ok('活读取器排第一', ids[0] === 'live-1', ids)
  ok('座位桥排第二', ids[1] === 'bridge-1', ids)
  ok('缓存排第三', ids[2] === 'cached-1', ids)

  reset('same', 'same', 'same')
  ok('同一个 id 只留一份', api.sessionCandidates().length === 1, api.sessionCandidates())

  reset('#throw', 'from-bridge', '')
  ok('读取器抛错时仍有候选', api.sessionCandidates().join(',') === 'from-bridge', api.sessionCandidates())

  reset(undefined, null, '')
  ok('什么都没有 → 空数组', api.sessionCandidates().length === 0, api.sessionCandidates())
}

console.log('— 过期 id 要能跳过 —')
{
  // 活读取器给的 id 是过期的（取不到会话），座位桥的是对的 —— 必须继续试。
  reset('stale-live', 'good', '')
  const ctx = ctxOf(['good'])
  const shell = api.composerShell(ctx)
  ok('拿得到 composer', shell !== null && shell !== undefined)
  ok('试了两个 id', ctx.scopes.length === 2, ctx.scopes)
  ok('第一个是过期的那个', ctx.scopes[0] === 'stale-live', ctx.scopes)
  ok('用的是能取到会话的那个', ctx.scopes[1] === 'good', ctx.scopes)
  ok('G.sessionId 被校正', api.G.sessionId === 'good', api.G.sessionId)
  ok('LAST_SCOPED 记的是好的 scoped', api.scoped().id === 'good', api.scoped())
  ok('诊段为 ok', api.why() === 'ok', api.why())
}

console.log('— 全都取不到时留下可读原因 —')
{
  reset('x1', 'x2', 'x3')
  const ctx = ctxOf([])
  ok('返回 null', api.composerShell(ctx) === null)
  ok('原因写明 scope-null 与试过的 id', /^scope-null\(x1,x2,x3\)$/.test(api.why()), api.why())
  ok('三个都试过', ctx.scopes.length === 3, ctx.scopes)
}

console.log('— 拿不到 id 时的原因 —')
{
  reset(undefined, null, '')
  ok('返回 null', api.composerShell(ctxOf([])) === null)
  ok('原因是 no-sid', api.why() === 'no-sid', api.why())
}

console.log('— scope 抛错不影响后面的候选 —')
{
  reset('bad-throw', 'good', '')
  const ctx = ctxOf(['good'], { throwsOn: 'bad-throw' })
  const shell = api.composerShell(ctx)
  ok('仍然拿得到', shell !== null && shell !== undefined)
  ok('第二个候选生效', api.G.sessionId === 'good', api.G.sessionId)
}

console.log('— shell 能力缺失时说明是哪一项 —')
{
  reset('good', null, '')
  ok('没有 insertReference → null', api.composerShell(ctxOf(['good'], { noInsert: true })) === null)
  ok('原因是 no-insertReference', api.why() === 'no-insertReference', api.why())
}

console.log('')
console.log(pass + ' passed, ' + fail + ' failed')
process.exit(fail === 0 ? 0 : 1)
