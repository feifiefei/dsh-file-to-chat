// 回归用例：引用块的定位上下文（symbolOfLine / headingPathOf / enclosingSymbolOf）。
//
// 三个函数都只读**已经渲染出来的 DOM**，所以这里手搓一个够用的假 DOM：
// closest / querySelectorAll / contains / compareDocumentPosition 四个能力就够了，
// 不需要 jsdom。
//
// 跑法：node test/spot.test.cjs   （或 npm test）
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

// ---- 假 DOM -------------------------------------------------------------
let SEQ = 0
function matchesOne(node, sel) {
  if (!node || node.nodeType !== 1) return false
  let m = sel.match(/^\[([\w-]+)\]$/)
  if (m !== null) return node.hasAttribute(m[1])
  m = sel.match(/^\[([\w-]+)="([^"]*)"\]$/)
  if (m !== null) return node.getAttribute(m[1]) === m[2]
  m = sel.match(/^\.([\w-]+)$/)
  if (m !== null) return String(node.getAttribute('class') || '').split(/\s+/).indexOf(m[1]) >= 0
  return node.tagName === sel.toUpperCase()
}
function matches(node, sel) {
  return String(sel).split(',').some((one) => matchesOne(node, one.trim()))
}

function el(tag, attrs) {
  const node = {
    nodeType: 1,
    tagName: String(tag).toUpperCase(),
    attrs: attrs || {},
    childNodes: [],
    parentElement: null,
    _id: ++SEQ,
    get textContent() {
      let s = ''
      for (const k of this.childNodes) s += typeof k === 'string' ? k : k.textContent
      return s
    },
    getAttribute(n) {
      return Object.prototype.hasOwnProperty.call(this.attrs, n) ? this.attrs[n] : null
    },
    hasAttribute(n) {
      return Object.prototype.hasOwnProperty.call(this.attrs, n)
    },
    append(...kids) {
      for (const k of kids) {
        if (typeof k === 'string') this.childNodes.push(k)
        else { k.parentElement = this; this.childNodes.push(k) }
      }
      return this
    },
    contains(o) {
      for (let p = o; p; p = p.parentElement) if (p === this) return true
      return false
    },
    compareDocumentPosition(o) {
      if (this === o) return 0
      if (this.contains(o)) return 4 | 16
      if (o.contains && o.contains(this)) return 2 | 8
      return this._id < o._id ? 4 : 2
    },
    closest(sel) {
      for (let p = this; p; p = p.parentElement) if (matches(p, sel)) return p
      return null
    },
    querySelectorAll(sel) {
      const out = []
      const walk = (n) => {
        for (const k of n.childNodes) {
          if (typeof k === 'string') continue
          if (matches(k, sel)) out.push(k)
          walk(k)
        }
      }
      walk(this)
      return out
    },
  }
  return node
}

function extractArr(marker) {
  const i = src.indexOf(marker)
  if (i < 0) throw new Error('marker not found: ' + marker)
  let depth = 0
  const j = src.indexOf('[', i)
  for (let k = j; k < src.length; k++) {
    if (src[k] === '[') depth++
    else if (src[k] === ']') { depth--; if (depth === 0) return src.slice(i, k + 1) }
  }
  throw new Error('unbalanced: ' + marker)
}

const code = [
  extract('function nodeToElement(node)'),
  extractArr('var SYMBOL_PATTERNS'),
  src.slice(src.indexOf('var NOT_A_NAME'), src.indexOf('\n', src.indexOf('var NOT_A_NAME'))),
  extract('function symbolOfLine(text)'),
  extract('function lineNumberOf(node)'),
  extract('function headingPathOf(node)'),
  extract('function enclosingSymbolOf(node)'),
  'return { symbolOfLine: symbolOfLine, headingPathOf: headingPathOf, enclosingSymbolOf: enclosingSymbolOf }',
].join('\n')

const api = new Function(code)()

let pass = 0, fail = 0
function eq(name, got, want) {
  const g = JSON.stringify(got), w = JSON.stringify(want)
  if (g === w) { pass++; console.log('  ok  ' + name) }
  else { fail++; console.log('  FAIL ' + name + '\n    got  ' + g + '\n    want ' + w) }
}

console.log('— 声明识别 —')
const sym = api.symbolOfLine
eq('JS 函数声明', sym('function readRange(workbook, sheetName, rg) {'), { kind: 'function', name: 'readRange' })
eq('导出 + 生成器', sym('export function* gen() {'), { kind: 'function', name: 'gen' })
eq('箭头函数赋值', sym('var spotOf = (node) => {'), { kind: 'function', name: 'spotOf' })
eq('const 函数表达式', sym('const go = async function () {'), { kind: 'function', name: 'go' })
eq('类', sym('class Foo extends Bar {'), { kind: 'class', name: 'Foo' })
eq('Python def', sym('    def render(self, x):'), { kind: 'def', name: 'render' })
eq('Go 方法（带接收者）', sym('func (s *Server) Handle(w http.ResponseWriter) {'), { kind: 'func', name: 'Handle' })
eq('Rust fn', sym('pub async fn build() {'), { kind: 'fn', name: 'build' })
eq('同类方法体', sym('  toString() {'), { kind: 'method', name: 'toString' })
eq('控制语句不是声明', sym('if (x) {'), null)
eq('for 循环不是声明', sym('  for (var i = 0; i < 3; i++) {'), null)
eq('调用不是声明', sym('  doSend()'), null)
eq('超长行（打包产物）直接放弃', sym('x'.repeat(500)), null)
eq('空行', sym(''), null)

console.log('— Markdown 标题路径 —')
const md = el('div', { 'data-document-markdown': '' })
const h1 = el('h1').append('需求')
const p1 = el('p').append('正文')
const h2 = el('h2').append('表结构')
const p2 = el('p').append('选中的文字')
md.append(h1, p1, h2, p2)
eq('① 二级标题下 → 带上父级', api.headingPathOf(p2), '# 需求 › ## 表结构')
eq('② 选中标题本身也把它算进去', api.headingPathOf(h2), '# 需求 › ## 表结构')
eq('③ 一级标题之下、二级之前 → 只有一级', api.headingPathOf(p1), '# 需求')

const md2 = el('div', { 'data-document-markdown': '' })
md2.append(el('h2').append('A'), el('h3').append('A.1'), el('h2').append('B'), el('p').append('t'))
eq('④ 同级标题要顶掉上一个（h2 → h2）', api.headingPathOf(md2.querySelectorAll('p')[0]), '## B')

const md3 = el('div', { 'data-document-markdown': '' })
md3.append(el('p').append('开头'))
eq('⑤ 选区在所有标题之前 → 无路径', api.headingPathOf(md3), null)
eq('⑥ 不在 Markdown 预览里 → 无路径', api.headingPathOf(el('p').append('x')), null)

console.log('— 纯文本预览所在函数 —')
const plain = el('div', { 'data-textpreview-plain': '' })
const page = el('pre', { 'data-textpreview-page': '0' })
const lineEl = (n, text) => el('div', { 'data-textpreview-line': String(n) }).append(text + '\n')
page.append(
  lineEl(100, 'function outer() {'),
  lineEl(101, '  var x = 1'),
  lineEl(102, '  function inner() {'),
  lineEl(103, '    var y = 2'),
  lineEl(104, '  }'),
  lineEl(105, '}'),
)
plain.append(page)
eq('① 取最近的那个声明，并带上起始行号',
  api.enclosingSymbolOf(page.childNodes[3]), 'function inner()，第 102 行起')
eq('② 声明自己那一行也算（选中的就是函数头）',
  api.enclosingSymbolOf(page.childNodes[0]), 'function outer()，第 100 行起')
eq('③ 上方没有声明 → null', api.enclosingSymbolOf(el('div', { 'data-textpreview-line': '1' }).append('x')), null)

console.log('— 代码预览（.line，无行号属性）—')
const cbody = el('div', { 'data-code-block-content': '' })
const pre = el('pre')
const codeLine = (t) => el('div', { class: 'line' }).append(t)
pre.append(codeLine('class Foo {'), codeLine('  bar() {'), codeLine('    return 1'), codeLine('  }'))
cbody.append(pre)
eq('① 代码按 index 推行号', api.enclosingSymbolOf(pre.childNodes[2]), 'method bar()，第 2 行起')
eq('② 类头（class 后面不加括号）', api.enclosingSymbolOf(pre.childNodes[0]), 'class Foo，第 1 行起')

console.log('')
console.log(pass + ' passed, ' + fail + ' failed')
process.exit(fail === 0 ? 0 : 1)
