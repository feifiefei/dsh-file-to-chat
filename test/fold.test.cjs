// 回归用例：折叠切段（splitQuoteRuns）与序列化补换行（serialize）。
//
// 这两个函数只依赖 G（globalThis.__dshFileToChat 的本地别名），所以可以从 client.js 里
// 抠出来单独跑 —— 不需要浏览器、不需要 DSH。哪天它们被重构掉，这里会当场炸。
//
// 跑法：node test/fold.test.cjs   （或 npm test）
const fs = require('fs')
const path = require('path')
const src = fs.readFileSync(path.join(__dirname, '..', 'client.js'), 'utf8')

// 按花括号配平，从源码里剪出一段完整定义（从 marker 起，到与之配平的那个 } 为止）。
function extract(marker) {
  const i = src.indexOf(marker)
  if (i < 0) throw new Error('marker not found: ' + marker)
  let depth = 0, j = src.indexOf('{', i)
  for (let k = j; k < src.length; k++) {
    if (src[k] === '{') depth++
    else if (src[k] === '}') { depth--; if (depth === 0) return src.slice(i, k + 1) }
  }
  throw new Error('unbalanced: ' + marker)
}

const G = { refBodies: new Map() }
const splitSrc = [
  extract('var isSource = function (line)'),
  extract('var bodyOf = function (run)'),
  extract('var gluedBlock = function (run)'),
  extract('function splitQuoteRuns(text)'),
  'return splitQuoteRuns',
].join('\n')
const splitQuoteRuns = new Function('G', splitSrc)(G)
const serialize = new Function('G', 'return ({ ' + extract('serialize: function (ref)') + ' }).serialize')(G)

const B1 = '> \uD83D\uDCC4 引用来源：D:\\wendang\\wind指标-ods版.xlsx（仅路径）\n> （未内联正文，请自行读取上面的路径）'
G.refBodies.set('q1', B1)

let pass = 0, fail = 0
function eq(name, got, want) {
  const g = JSON.stringify(got), w = JSON.stringify(want)
  if (g === w) { pass++; console.log('  ok  ' + name) }
  else { fail++; console.log('  FAIL ' + name + '\n    got  ' + g + '\n    want ' + w) }
}

console.log('— 折叠切段 —')
eq('① 老消息：用户正文黏在末行 → 切出来',
  splitQuoteRuns('根据当前的需求样表和\n' + B1 + '，整理一个SmartBI的配置方案，用.md文档即可，步骤详细'),
  [{ quote: false, text: '根据当前的需求样表和' },
   { quote: true, text: B1, body: '📄 引用来源：D:\\wendang\\wind指标-ods版.xlsx（仅路径）\n（未内联正文，请自行读取上面的路径）' },
   { quote: false, text: '，整理一个SmartBI的配置方案，用.md文档即可，步骤详细' }])

eq('② 连插两枚胶囊（换行隔开）→ 整段照折，不拆碎',
  splitQuoteRuns(B1 + '\n' + B1),
  [{ quote: true, text: B1 + '\n' + B1, body: '📄 引用来源：D:\\wendang\\wind指标-ods版.xlsx（仅路径）\n（未内联正文，请自行读取上面的路径）\n📄 引用来源：D:\\wendang\\wind指标-ods版.xlsx（仅路径）\n（未内联正文，请自行读取上面的路径）' }])

eq('③ 单块、后面没字 → 照旧一块',
  splitQuoteRuns('' + B1), [{ quote: true, text: B1, body: '📄 引用来源：D:\\wendang\\wind指标-ods版.xlsx（仅路径）\n（未内联正文，请自行读取上面的路径）' }])

eq('④ 别人写的 Markdown 引用 → 绝不折',
  splitQuoteRuns('> 这只是一句引用'),
  [{ quote: false, text: '> 这只是一句引用' }])

eq('⑤ 批注块（空行隔开）→ 引用块 + 批注行',
  splitQuoteRuns(B1 + '\n\n批注：看一下这个'),
  [{ quote: true, text: B1, body: '📄 引用来源：D:\\wendang\\wind指标-ods版.xlsx（仅路径）\n（未内联正文，请自行读取上面的路径）' },
   { quote: false, text: '\n批注：看一下这个' }])

console.log('— 序列化补换行 —')
serialize('q1').then(function (s) {
  eq('⑥ 普通块 → 末尾补换行', s, B1 + '\n')
  G.refBodies.set('q2', B1 + '\n\n批注：')
  return serialize('q2')
}).then(function (s) {
  eq('⑦ 「批注：」前缀 → 不补（批注要接在同一行）', s, B1 + '\n\n批注：')
  G.refBodies.set('q3', '')
  return serialize('q3')
}).then(function (s) {
  eq('⑧ 空正文 → 不补', s, '')
  console.log('\n' + pass + ' passed, ' + fail + ' failed')
  process.exit(fail === 0 ? 0 : 1)
})
