// 回归用例：输入框里粘贴长文本 → 折成胶囊（foldPastedText）。
//
// 这条路径的危险点是**吞掉粘贴**：我们 preventDefault 之后，一旦胶囊没插成，
// 用户粘的东西就消失了。所以用例里一半是「失败必须返回 false」——
// 调用方据此放行默认粘贴。
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
  'var G = { refBodies: new Map(), refSeq: 0 }',
  'var QUOTE_SOURCE = "file-to-chat-ref"',
  'var PASTE_STAT = { seen: 0, long: 0, folded: 0, passed: 0 }',
  'var dbg = function () {}',
  extract('function isLongText(text)'),
  extract('function longTextLabel(text)'),
  extract('function longTextPreview(text, max)'),
  extract('function foldPastedText(shell, raw)'),
  'return { foldPastedText: foldPastedText, G: G, STAT: PASTE_STAT }',
].join('\n')

const api = new Function(code)()
const long = new Array(20).fill('第一段会议记录').join('\n')

let pass = 0, fail = 0
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('  ok  ' + name) }
  else { fail++; console.log('  FAIL ' + name + (extra === undefined ? '' : '  → ' + JSON.stringify(extra))) }
}

function shellOf(opts) {
  opts = opts || {}
  const calls = []
  const shell = {
    state: { getSnapshot: () => ({ draft: opts.draft === undefined ? '已有草稿' : opts.draft, draftRev: 9 }) },
    insertReference: function (info, span) {
      calls.push({ info: info, span: span })
      if (opts.throws === true) throw new Error('phase not plain')
      return opts.applied !== false
    },
  }
  if (opts.caret !== false) shell.caretSpan = function () { return opts.caret || null }
  shell.calls = calls
  return shell
}

console.log('— 正常折叠 —')
{
  const sh = shellOf({ caret: { start: 3, end: 3 }, draft: 'abc' })
  const rawBefore = api.G.refBodies.size
  ok('返回 true', api.foldPastedText(sh, long) === true)
  ok('插了一次胶囊', sh.calls.length === 1)
  const call = sh.calls[0]
  ok('span 用当场的光标', call.span.start === 3 && call.span.end === 3 && call.span.draftRev === 9, call.span)
  ok('source 是引用那一个', call.info.source === 'file-to-chat-ref')
  ok('ref 以 p 开头', /^p\d+$/.test(call.info.ref), call.info.ref)
  ok('正文存的是原文', api.G.refBodies.get(call.info.ref) === long)
  ok('clipboardText 也是原文', call.info.clipboardText === long)
  ok('胶囊标签含行数与字数',
    call.info.label.indexOf('长文本 · 20 行 · ' + long.length + ' 字') === 0, call.info.label)
  ok('胶囊标签带首行摘要', call.info.label.indexOf('· 第一段会议记录') > 0, call.info.label)
  ok('计数加了 1', api.STAT.folded === 1 && api.G.refBodies.size === rawBefore + 1)
}

console.log('— 没有光标时落到草稿末尾 —')
{
  const sh = shellOf({ caret: false, draft: '一段草稿' })
  ok('还是能插', api.foldPastedText(sh, long) === true)
  ok('span 落在末尾', sh.calls[0].span.start === 4 && sh.calls[0].span.end === 4, sh.calls[0].span)
}

console.log('— 失败必须放行默认粘贴 —')
{
  const before = api.G.refBodies.size
  const sh1 = shellOf({ applied: false })
  ok('insertReference 返回 false → false', api.foldPastedText(sh1, long) === false)
  ok('不留垃圾正文', api.G.refBodies.size === before)
  const sh2 = shellOf({ throws: true })
  ok('insertReference 抛错 → false', api.foldPastedText(sh2, long) === false)
  ok('抛错也不留垃圾', api.G.refBodies.size === before)
  const sh3 = { state: { getSnapshot: function () { throw new Error('没快照') } } }
  ok('拿不到快照 → false', api.foldPastedText(sh3, long) === false)
  const sh4 = { state: { getSnapshot: function () { return { draft: 'x' } } }, insertReference: function () { return true } }
  ok('snapshot 没 draftRev → false', api.foldPastedText(sh4, long) === false)
}

console.log('— 胶囊提交时展开成原文 —')
{
  const key = Array.from(api.G.refBodies.keys()).pop()
  ok('能按 ref 取回原文（codec.serialize 走的就是这张表）', api.G.refBodies.get(key) === long)
  ok('存进来的不是引用块（没有 > 前缀）', !/^>/.test(api.G.refBodies.get(key)))
}

console.log('')
console.log(pass + ' passed, ' + fail + ' failed')
process.exit(fail === 0 ? 0 : 1)
