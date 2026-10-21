// 回归用例：胶囊「打开并定位」的地址与解析（fileAddressOf / parseQuoteTarget）。
//
// 这两个函数是纯字符串处理，抠出来直接跑。
// 地址规则与 @deepseek-ai/dsh-util-workspace-path 的 fileAddressFor 同构：
//   dsh-resource://file/session/<sid>/<path>   会话内（相对 cwd 或绝对路径都行）
//   dsh-resource://file/absolute/<path>        没有会话 id 时
//
// 跑法：node test/jump.test.cjs   （或 npm test）
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

const code = [
  extract('function isAbsPath(p)'),
  extract('function encodeSegment(s)'),
  extract('function fileAddressOf(path, sessionId, cwd)'),
  extract('function parseQuoteTarget(body)'),
  'return { fileAddressOf: fileAddressOf, parseQuoteTarget: parseQuoteTarget }',
].join('\n')

const api = new Function(code)()

let pass = 0, fail = 0
function eq(name, got, want) {
  const g = JSON.stringify(got), w = JSON.stringify(want)
  if (g === w) { pass++; console.log('  ok  ' + name) }
  else { fail++; console.log('  FAIL ' + name + '\n    got  ' + g + '\n    want ' + w) }
}

console.log('— 解析引用块首行 —')
// 真实形态：composeTextBlock 拼的是 `> 📄 引用来源：…`，折起来后 body 去掉了 `> `。
const t1 = api.parseQuoteTarget('📄 引用来源：C:\\Users\\me\\a.js（第 12-20 行）')
eq('① 行区间 → 取起始行', t1, { path: 'C:\\Users\\me\\a.js', line: 12 })
eq('② 单行', api.parseQuoteTarget('📄 引用来源：C:\\a.js（第 7 行）'), { path: 'C:\\a.js', line: 7 })
eq('③ 整份文件 → 不带行号', api.parseQuoteTarget('📄 引用来源：C:\\a.js（整份文件）'), { path: 'C:\\a.js', line: null })
eq('④ Excel 区域 → 不带行号',
  api.parseQuoteTarget('📄 引用来源：D:\\表.xlsx（工作表 Sheet1 · 区域 B3:D7）'),
  { path: 'D:\\表.xlsx', line: null })
eq('⑤ PDF 页码 → 不当作行号',
  api.parseQuoteTarget('📄 引用来源：D:\\稿.pdf（第 3 页）'), { path: 'D:\\稿.pdf', line: null })
eq('⑥ 多行正文只看首行',
  api.parseQuoteTarget('📄 引用来源：C:\\a.js（第 9 行）\nfunction foo() {\n  return 1'),
  { path: 'C:\\a.js', line: 9 })
eq('⑦ 首行不是引用来源 → null', api.parseQuoteTarget('📄 引用来源：C:\\a.js（第 1 行）', null).path, 'C:\\a.js')
eq('⑧ 未识别来源 → null', api.parseQuoteTarget('📄 引用来源：（未识别来源文件）'), null)
eq('⑨ 相对路径不跳（没法定位到具体工作区）', api.parseQuoteTarget('📄 引用来源：src/a.js（第 1 行）'), null)
eq('⑩ 没有「引用来源」字样 → null', api.parseQuoteTarget('这就是一段普通引用'), null)
eq('⑪ 带空格与全角括号的路径', api.parseQuoteTarget('📄 引用来源：C:\\my docs\\a b.js（第 2 行）'),
  { path: 'C:\\my docs\\a b.js', line: 2 })

console.log('— 拼地址 —')
eq('① Windows 绝对路径 + 会话',
  api.fileAddressOf('C:\\Users\\me\\a.js', 's1', ''), 'dsh-resource://file/session/s1/C:/Users/me/a.js')
eq('② 在工作区内 → 转成相对路径',
  api.fileAddressOf('C:\\w\\src\\a.js', 's1', 'C:\\w'), 'dsh-resource://file/session/s1/src/a.js')
eq('③ 工作区根自己',
  api.fileAddressOf('C:\\w', 's1', 'C:\\w'), 'dsh-resource://file/session/s1/')
eq('④ 没有会话 id → absolute 作用域',
  api.fileAddressOf('C:\\Users\\me\\a.js', '', ''), 'dsh-resource://file/absolute/C:/Users/me/a.js')
eq('⑤ 空格要转义', api.fileAddressOf('C:\\my docs\\a b.js', 's1', ''),
  'dsh-resource://file/session/s1/C:/my%20docs/a%20b.js')
eq('⑥ 盘符的冒号保持字面量', api.fileAddressOf('D:\\x.js', 's2', ''),
  'dsh-resource://file/session/s2/D:/x.js')
eq('⑦ 反斜杠统一成斜杠', api.fileAddressOf('C:\\a\\b\\c.js', 's', ''), 'dsh-resource://file/session/s/C:/a/b/c.js')
eq('⑧ 相对路径', api.fileAddressOf('src/a.js', 's', ''), 'dsh-resource://file/session/s/src/a.js')

console.log('')
console.log(pass + ' passed, ' + fail + ' failed')
process.exit(fail === 0 ? 0 : 1)
