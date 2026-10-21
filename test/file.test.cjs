// 回归用例：大文件续发（fileRangeLabel 的范围描述 + 续发游标的推进规则）。
//
// 跑法：node test/file.test.cjs   （或 npm test）
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
  'var G = {}',
  extract('function fileRangeLabel(snap)'),
  extract('function cursorOf(path)'),
  extract('function rememberCursor(path, meta)'),
  'return { fileRangeLabel: fileRangeLabel, cursorOf: cursorOf, rememberCursor: rememberCursor, G: G }',
].join('\n')

const api = new Function(code)()

let pass = 0, fail = 0
function eq(name, got, want) {
  const g = JSON.stringify(got), w = JSON.stringify(want)
  if (g === w) { pass++; console.log('  ok  ' + name) }
  else { fail++; console.log('  FAIL ' + name + '\n    got  ' + g + '\n    want ' + w) }
}

console.log('— 范围描述 —')
const snap = (meta, label) => ({ kind: 'file', label: label === undefined ? '整份文件' : label, meta: meta })

eq('① 读到头 → 敢说共 N 行',
  api.fileRangeLabel(snap({ startLine: 1, endLine: 42, lines: 42, eof: true })), '整份文件，共 42 行')
eq('② 宿主不报总行数 → 只说行区间，不编总数',
  api.fileRangeLabel(snap({ startLine: 1, endLine: 200, lines: null, eof: false })), '第 1-200 行')
eq('③ 续发第三块',
  api.fileRangeLabel(snap({ startLine: 401, endLine: 500, lines: null, eof: false })), '第 401-500 行')
eq('④ 续发正好读到头',
  api.fileRangeLabel(snap({ startLine: 201, endLine: 260, lines: 260, eof: true })), '第 201-260 行（共 260 行）')
eq('⑤ 没有 meta（仅路径/选区）→ 退回 label',
  api.fileRangeLabel(snap(null, '仅路径')), '仅路径')
eq('⑥ meta 里没有行号 → 退回 label',
  api.fileRangeLabel(snap({ lines: null, eof: true }, '第 9 行起')), '第 9 行起')

console.log('— 续发游标 —')
eq('⑦ 没记过 → 没有入口', api.cursorOf('a.md'), null)
api.rememberCursor('a.md', { startLine: 1, endLine: 200, eof: false })
eq('⑧ 没读完 → 下次从 201 行起', api.cursorOf('a.md'), 201)
api.rememberCursor('a.md', { startLine: 201, endLine: 260, eof: true })
eq('⑨ 读完了 → 入口消失', api.cursorOf('a.md'), null)
api.rememberCursor('b.md', { startLine: 401, endLine: 900, eof: false })
eq('⑩ 每个文件各自记', api.cursorOf('b.md'), 901)
api.rememberCursor('c.md', null)
eq('⑪ meta 缺失时不动游标（不误删）', api.cursorOf('b.md'), 901)
api.rememberCursor('d.md', { lines: null, eof: true, from: 'text' })
eq('⑫ 没有 endLine 也不记', api.cursorOf('d.md'), null)

console.log('\n' + (fail === 0 ? 'ALL PASS' : 'FAILED') + ' — ' + pass + ' passed, ' + fail + ' failed')
process.exit(fail === 0 ? 0 : 1)
