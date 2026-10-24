// 回归用例：对话内容批注（composeChatBlock / parseChatTarget / describeQuote）。
//
// 场景：用户在聊天区里划中一段助手回答或用户消息，点「批注」发出去。
// 发出去的那块正文长这样（首行是自己的签名，折叠器靠它分辨）：
//
//   > 💬 引用对话：助手回答（第 2 轮）
//   > 正文第一行
//
// 这里只测纯字符串处理的部分（拼块、解析、胶囊标签）。
//
// 跑法：node test/chat.test.cjs   （或 npm test）
const fs = require('fs')
const path = require('path')
const src = fs.readFileSync(path.join(__dirname, '..', 'client.js'), 'utf8')

function extract(marker) {
  const i = src.indexOf(marker)
  if (i < 0) throw new Error('marker not found: ' + marker)
  let depth = 0
  const j = src.indexOf('{', i)
  for (let k = j; k < src.length; k++) {
    if (src[k] === '{') depth++
    else if (src[k] === '}') { depth--; if (depth === 0) return src.slice(i, k + 1) }
  }
  throw new Error('unbalanced: ' + marker)
}

// MAX_EXCERPT_LINES / MAX_EXCERPT_CHARS 是 var 声明，不是函数 —— 按行取。
const capLine = src.split('\n').find((l) => l.includes('var MAX_EXCERPT_LINES ='))
if (capLine === undefined) throw new Error('marker not found: MAX_EXCERPT_LINES')
const charLine = src.split('\n').find((l) => l.includes('var MAX_EXCERPT_CHARS ='))
if (charLine === undefined) throw new Error('marker not found: MAX_EXCERPT_CHARS')

const code = [
  capLine.trim(),
  charLine.trim(),
  extract('function chatRoleOf(kind)'),
  extract('function composeChatBlock(snap)'),
  extract('function parseChatTarget(body)'),
  extract('function describeQuote(quoteText)'),
  extract('function baseName(p)'),
  'return { chatRoleOf: chatRoleOf, composeChatBlock: composeChatBlock,'
    + ' parseChatTarget: parseChatTarget, describeQuote: describeQuote }',
].join('\n')

const api = new Function(code)()

let pass = 0, fail = 0
function eq(name, got, want) {
  const g = JSON.stringify(got), w = JSON.stringify(want)
  if (g === w) { pass++; console.log('  ok  ' + name) }
  else { fail++; console.log('  FAIL ' + name + '\n    got  ' + g + '\n    want ' + w) }
}

console.log('— 角色映射 —')
eq('① user → 用户消息', api.chatRoleOf('user'), '用户消息')
eq('② steering → 用户插话', api.chatRoleOf('steering'), '用户插话')
eq('③ assistant → 助手回答', api.chatRoleOf('assistant'), '助手回答')
eq('④ assistant-step 也算助手回答', api.chatRoleOf('assistant-step'), '助手回答')
eq('⑤ 工具过程行不引用', api.chatRoleOf('turn-process'), '')
eq('⑥ 空值不引用', api.chatRoleOf(null), '')

console.log('— 拼对话引用块 —')
const block = api.composeChatBlock({ role: '助手回答', turn: 2, text: '第一行\n第二行' })
eq('① 首行带 💬 签名与轮次', block[0], '> 💬 引用对话：助手回答（第 2 轮）')
eq('② 正文每行都加 > 前缀', block.slice(1), ['> 第一行', '> 第二行'])
eq('③ 没有轮次时不写括号',
  api.composeChatBlock({ role: '用户消息', turn: null, text: 'x' })[0], '> 💬 引用对话：用户消息')
eq('④ 超长选区按 MAX_EXCERPT_LINES 截断并留提示',
  /^> …（选区共 45 行 \/ \d+ 字，仅内联前 40 行）$/.test(
    api.composeChatBlock({ role: '助手回答', turn: 1, text: Array.from({ length: 45 }, (_, i) => 'L' + i).join('\n') }).slice(-1)[0]),
  true)
eq('⑤ 正好等于上限时不加提示',
  api.composeChatBlock({ role: '助手回答', turn: 1, text: Array.from({ length: 40 }, (_, i) => 'L' + i).join('\n') }).length,
  41)
// 行数上限挡不住一整段不换行的长回答：那一行本身就要按字符预算截。
const huge = api.composeChatBlock({ role: '助手回答', turn: 1, text: 'x'.repeat(5000) })
eq('⑥ 单行超长按字符预算截断（补省略号）',
  huge[1].length, 2 + 1200 + 1)
eq('⑦ 截在预算上时结尾是省略号', huge[1].slice(-1), '…')
eq('⑧ 只有一行、已内联到头 → 不再多写统计行', huge.length, 2)

console.log('— 解析对话引用首行（胶囊 ↗ 定位用）—')
eq('① 角色 + 轮次', api.parseChatTarget('💬 引用对话：助手回答（第 2 轮）'), { role: '助手回答', turn: 2 })
eq('② 没有轮次', api.parseChatTarget('💬 引用对话：用户消息'), { role: '用户消息', turn: null })
eq('③ 多行正文只看首行',
  api.parseChatTarget('💬 引用对话：用户插话（第 7 轮）\n正文'), { role: '用户插话', turn: 7 })
eq('④ 文件引用不是对话引用', api.parseChatTarget('📄 引用来源：C:\\a.js（第 1 行）'), null)
eq('⑤ 普通文本 → null', api.parseChatTarget('随便一段话'), null)

console.log('— 胶囊标签 —')
eq('① 对话引用显示角色与轮次', api.describeQuote('💬 引用对话：助手回答（第 2 轮）\n正文'),
  { label: '助手回答 · 第 2 轮' })
eq('② 文件引用不受影响', api.describeQuote('📄 引用来源：C:\\w\\a.js（第 12-20 行）'),
  { label: 'a.js · 第 12-20 行' })

console.log('— 定位到「当初引用的那一条」（同轮多条 assistant 时）—')
// 一行只有两个属性有用：kind 和 turn；这里给个最小替身。
// `hidden` 模拟 React 留在树上、用 display:none 藏起来的 transcript 副本。
function fakeRow(id, kind, turn, hidden) {
  return {
    id: id,
    getAttribute: function (n) {
      if (n === 'data-chat-flow-kind') return kind
      if (n === 'data-chat-turn') return String(turn)
      return null
    },
    getClientRects: function () { return hidden === true ? [] : [{}] },
  }
}
const rows = [
  fakeRow('u1', 'user', 1),
  fakeRow('a1-step', 'assistant-step', 1),
  fakeRow('a1', 'assistant', 1),
  fakeRow('u2', 'user', 2),
  fakeRow('a2', 'assistant', 2),
  // 同一轮的隐藏副本排在后面：跳上去等于「点了没反应」。
  fakeRow('a1-hidden', 'assistant', 1, true),
]
global.document = { querySelectorAll: function () { return rows } }

const locate = new Function([
  src.split('\n').find((l) => l.includes('var CHAT_ROW_SELECTOR =')).trim(),
  extract('function chatRoleOf(kind)'),
  extract('function chatTargetOf(row)'),
  extract('function isRendered(el)'),
  extract('function findChatRow(target)'),
  'return { findChatRow: findChatRow }',
].join('\n'))()

eq('① kind 精确匹配优先于「同轮最后一条」',
  locate.findChatRow({ role: '助手回答', flowKind: 'assistant', turn: 1 }).id, 'a1')
eq('② kind 对不上时退回角色匹配（取同角色最后一条）',
  locate.findChatRow({ role: '助手回答', flowKind: 'nope', turn: 1 }).id, 'a1')
eq('③ 没有 kind 时按角色',
  locate.findChatRow({ role: '用户消息', flowKind: null, turn: 2 }).id, 'u2')
eq('④ 轮次不存在 → null',
  locate.findChatRow({ role: '助手回答', flowKind: 'assistant', turn: 9 }), null)
eq('⑤ turn 为 null → null',
  locate.findChatRow({ role: '助手回答', flowKind: 'assistant', turn: null }), null)
eq('⑥ 隐藏副本排在后面也不跳过去（看得见的优先）',
  locate.findChatRow({ role: '助手回答', flowKind: 'assistant', turn: 1 }).id, 'a1')

console.log(pass + ' passed, ' + fail + ' failed')
process.exit(fail === 0 ? 0 : 1)
