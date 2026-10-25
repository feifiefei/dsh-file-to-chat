// 回归用例：Excel 网格命中判定 + 活动工作表名。
//
// 这两件事都只读**已经渲染出来的 DOM**，所以手搓一个够用的假 DOM：
// closest / querySelector / querySelectorAll 三个能力，加上一个能认
// `.class`、`[attr]`、`[attr="v"]`、`[class*="v"]` 和后代组合的选择器。
//
// 为什么值得单独盯住：这两个函数读的都是 FortuneSheet 的**内部类名**，
// 而它是第三方组件、类名随版本换过好几次（fortune-sheet-list-item-* →
// luckysheet-sheets-item-*）。产物一升级，它们就静默失效 —— 没有异常、
// 没有日志，只是「点了编辑栏会冒批注按钮」「引用里少了工作表名」。
//
// 跑法：node test/excel.test.cjs   （或 npm test）
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

// 取一行或多行「字符串拼接」的 var 定义。FORTUNE_CHROME 是跨行 4 段的拼接。
function varText(name) {
  const m = src.match(new RegExp('var ' + name + ' = [^\\n]*(?:\\n\\s*\\+ [^\\n]*)*'))
  if (m === null) throw new Error('var not found: ' + name)
  return m[0]
}

// ---- 假 DOM -------------------------------------------------------------
let SEQ = 0

function splitTop(s, sep) {
  const out = []
  let depth = 0
  let cur = ''
  for (const ch of s) {
    if (ch === '[') depth++
    else if (ch === ']') depth--
    if (ch === sep && depth === 0) { out.push(cur); cur = '' } else cur += ch
  }
  out.push(cur)
  return out.map((x) => x.trim()).filter((x) => x !== '')
}

function matchesSimple(node, sel) {
  if (!node || node.nodeType !== 1) return false
  let m = sel.match(/^\[([\w-]+)\]$/)
  if (m !== null) return node.hasAttribute(m[1])
  m = sel.match(/^\[([\w-]+)="([^"]*)"\]$/)
  if (m !== null) return node.getAttribute(m[1]) === m[2]
  m = sel.match(/^\[class\*="([^"]*)"\]$/)
  if (m !== null) return String(node.getAttribute('class') || '').indexOf(m[1]) >= 0
  m = sel.match(/^\.([\w-]+)$/)
  if (m !== null) return String(node.getAttribute('class') || '').split(/\s+/).indexOf(m[1]) >= 0
  return node.tagName === sel.toUpperCase()
}

function matches(node, sel) {
  return splitTop(sel, ',').some((one) => {
    const parts = splitTop(one, ' ')
    if (!matchesSimple(node, parts[parts.length - 1])) return false
    let i = parts.length - 2
    let p = node.parentElement
    while (i >= 0) {
      if (!p) return false
      if (matchesSimple(p, parts[i])) i--
      p = p.parentElement
    }
    return true
  })
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
    querySelector(sel) {
      const all = this.querySelectorAll(sel)
      return all.length === 0 ? null : all[0]
    },
  }
  return node
}

const code = [
  extract('function nodeToElement(node)'),
  varText('FORTUNE_HOST'),
  varText('FORTUNE_CHROME'),
  extract('function previewRootOf(node)'),
  extract('function excelHostOf(node)'),
  extract('function activeSheetName(host)'),
  extract('function colIndexOf(letters)'),
  extract('function colNameOf(index)'),
  extract('function parseAddress(text)'),
  extract('function addressOf(rg)'),
  'return { excelHostOf: excelHostOf, activeSheetName: activeSheetName,'
    + ' colIndexOf: colIndexOf, colNameOf: colNameOf,'
    + ' parseAddress: parseAddress, addressOf: addressOf }',
].join('\n')

const api = new Function(code)()

let pass = 0
let fail = 0
// 断言失败时最想看的是「命中了哪个元素」，而不是循环引用报错 —— 假 DOM 有
// parentElement 反向指针，直接 JSON.stringify 会炸掉整个用例。
function show(v) {
  if (v !== null && typeof v === 'object' && v.nodeType === 1) {
    const cls = v.getAttribute('class')
    return '<' + String(v.tagName).toLowerCase() + (cls ? '.' + String(cls).split(/\s+/).join('.') : '') + '>'
  }
  return JSON.stringify(v)
}
function eq(name, got, want) {
  const g = show(got)
  const w = show(want)
  if (g === w) { pass++; console.log('  ok  ' + name) }
  else { fail++; console.log('  FAIL ' + name + '\n    got  ' + g + '\n    want ' + w) }
}

// ---- 复现 FortuneSheet 在 DSH 0.2.1 里的真实结构 -------------------------
// 关键点：workarea（工具栏 + 编辑栏）与 sheet-container（网格）、
// luckysheet-sheet-area（标签条）是**兄弟**，不是祖先关系。
// 把标签条塞进 workarea 里做夹具，就再也测不出「标签条被误判成网格」。
function buildSheet(opts) {
  const o = opts || {}
  const grid = el('div', { class: 'fortune-sheet-canvas-placeholder' })
  const canvas = el('canvas', { class: 'fortune-sheet-canvas' })
  const sheet = el('div', { class: 'fortune-sheet-container' }).append(grid, canvas)

  const button = el('button', { class: 'fortune-toolbar-btn tip' })
  const toolbar = el('div', { class: 'fortune-toolbar' }).append(button)
  const nameBox = el('div', { class: 'fortune-name-box' }).append('A1')
  const nameBoxBox = el('div', { class: 'fortune-name-box-container' }).append(nameBox)
  const fxInput = el('div', { class: 'fortune-fx-input', contenteditable: 'true' })
  const fx = el('div', { class: 'fortune-fx-editor' }).append(nameBoxBox, fxInput)
  const workarea = el('div', { class: 'fortune-workarea' }).append(toolbar, el('aside').append(fx))

  const tabName = el('span', { class: o.tabClass || 'luckysheet-sheets-item-name' }).append('Sheet1')
  const tabItem = el('div', {
    class: o.tabItemClass || 'luckysheet-sheets-item luckysheet-sheets-item-active',
  }).append(tabName)
  const tabBarClass = o.tabBarClass || 'luckysheet-sheet-area luckysheet-noselected-text'
  const tabBar = el('div', { class: tabBarClass }).append(el('div', { id: 'luckysheet-sheet-content' }).append(tabItem))

  const statArea = el('div', { class: 'fortune-stat-area' })
  const container = el('div', { class: 'fortune-container' })
    .append(el('section', { id: 'shortcut-list' }), workarea, sheet, tabBar, statArea)

  const root = el('div', { 'data-document-preview': 'doc-1' })
    .append(el('div', { class: 'header' }), container)

  return { root, container, workarea, sheet, canvas, grid, button, nameBox, fxInput, tabBar, tabItem, tabName }
}

console.log('— 网格命中（点在哪儿算「点表格」）—')

const d = buildSheet({})
const hit = api.excelHostOf(d.canvas)
eq('① 点在网格上 → 命中', hit !== null, true)
eq('② 命中的是文档序最外层的 workarea（名称框在它里面）',
  hit === d.workarea, true)
eq('③ 点在工具栏按钮上 → 不算点网格（.fortune-toolbar 在排除表里）',
  api.excelHostOf(d.button), null)
eq('④ 点在名称框上 → 不算（.fortune-name-box / -container 都在排除表里）',
  api.excelHostOf(d.nameBox), null)
// 这一条是 1.1.1 修的：rc.3 的编辑栏是 div.fortune-fx-input[contenteditable]，
// 它是 workarea 的后代 —— 不排除它，点一下编辑栏就冒批注按钮。
eq('⑤ 点在编辑栏上 → 不算（.fortune-fx-editor / .fortune-fx-input）',
  api.excelHostOf(d.fxInput), null)
eq('⑥ 点在工作表标签上 → 不算（.luckysheet-sheet-area / -sheets-item）',
  api.excelHostOf(d.tabItem), null)
eq('⑦ 点在工作表标签文字上 → 不算', api.excelHostOf(d.tabName), null)
eq('⑧ 预览区之外的元素 → null', api.excelHostOf(el('div', { class: 'outside' })), null)
eq('⑨ 文本节点按父元素算（nodeToElement 的老约定）',
  api.excelHostOf(d.canvas) !== null, true)

// 旧世代的标签条类名也要能被认出并排除 —— 否则回退到旧宿主时又会误弹。
const legacyTab = buildSheet({ tabBarClass: 'fortune-sheettab-container' })
eq('⑩ 旧结构 .fortune-sheettab-container 里的标签同样被排除',
  api.excelHostOf(legacyTab.tabItem), null)

console.log('— 活动工作表名 —')
eq('① 从 workarea 找 → 拿到标签条上的名字（标签条是它的兄弟，必须往外找）',
  api.activeSheetName(d.workarea), 'Sheet1')
eq('② 从预览根找 → 同样拿到', api.activeSheetName(d.root), 'Sheet1')
eq('③ 从网格容器找 → 同样拿到', api.activeSheetName(d.sheet), 'Sheet1')
eq('④ 标签条不在传入元素的作用域里、也没有 fortune-container → null',
  api.activeSheetName(el('div', { class: 'lonely' })), null)

// 旧世代（fortune-sheet-list-item-*）仍要兼容：它们出现在「工作表列表」浮层里，
// 浮层挂在 fortune-container 下，所以从 workarea 出发也能找到。
const legacy = buildSheet({})
legacy.container.append(el('div', { class: 'fortune-sheet-list-item fortune-sheet-list-item-active' })
  .append(el('span', { class: 'fortune-sheet-list-item-name' }).append('旧世代表')))
eq('⑤ 旧世代选择器仍兼容', api.activeSheetName(legacy.workarea), 'Sheet1')

const onlyLegacy = el('div', { class: 'fortune-container' })
  .append(el('div', { class: 'fortune-workarea' }),
    el('div', { class: 'fortune-sheet-list-item fortune-sheet-list-item-active' })
      .append(el('span', { class: 'fortune-sheet-list-item-name' }).append('旧表')))
eq('⑥ 只有旧世代标签时也能读到', api.activeSheetName(onlyLegacy.querySelector('.fortune-workarea')), '旧表')

const blank = el('div', { class: 'fortune-container' })
  .append(el('div', { class: 'fortune-workarea' }),
    el('div', { class: 'luckysheet-sheet-area' })
      .append(el('div', { class: 'luckysheet-sheets-item luckysheet-sheets-item-active' })
        .append(el('span', { class: 'luckysheet-sheets-item-name' }).append('   '))))
eq('⑦ 标签文字是空白 → 视为没读到（不能返回空串）',
  api.activeSheetName(blank.querySelector('.fortune-workarea')), null)

console.log('— 地址换算（名称框兜底路径要用）—')
eq('① A → 0', api.colIndexOf('A'), 0)
eq('② Z → 25', api.colIndexOf('Z'), 25)
eq('③ AA → 26', api.colIndexOf('AA'), 26)
eq('④ 0 → A / 26 → AA', [api.colNameOf(0), api.colNameOf(26)], ['A', 'AA'])
eq('⑤ 双向一致性',
  [api.colIndexOf(api.colNameOf(51)), api.colNameOf(api.colIndexOf('AZ'))], [51, 'AZ'])
eq('⑥ 带表名前缀 + 绝对引用', api.parseAddress("Sheet1!$B$3:D7"),
  { sheet: 'Sheet1', r1: 2, r2: 6, c1: 1, c2: 3 })
eq('⑦ 单格', api.parseAddress('C5'), { sheet: null, r1: 4, r2: 4, c1: 2, c2: 2 })
eq('⑧ 表名带引号与转义', api.parseAddress("'My ''S''!A1").sheet, "My 'S'")
eq('⑨ 非法输入', [api.parseAddress(''), api.parseAddress('A'), api.parseAddress('1A')],
  [null, null, null])
eq('⑩ 反写', api.addressOf({ r1: 2, r2: 6, c1: 1, c2: 3 }), 'B3:D7')
eq('⑪ 单格反写', api.addressOf({ r1: 0, r2: 0, c1: 0, c2: 0 }), 'A1')

console.log('\n' + (fail === 0 ? 'ALL PASS' : 'FAILED') + ' — ' + pass + ' passed, ' + fail + ' failed')
process.exit(fail === 0 ? 0 : 1)
