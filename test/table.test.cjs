// 回归用例：Excel 选区自动带表头（headerRowOf + readRange 的表头字段 + composeTableBlock 的拼装）。
//
// 这几个函数只依赖工作簿数组与几个常量，可以从 client.js 里抠出来单独跑。
//
// 跑法：node test/table.test.cjs   （或 npm test）
const fs = require('fs')
const path = require('path')
const src = fs.readFileSync(path.join(__dirname, '..', 'client.js'), 'utf8')

// 按花括号配平，从源码里剪出一段完整定义（从 marker 起，到与之配平的那个 } 为止）。
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

function constOf(name) {
  const m = src.match(new RegExp('var ' + name + ' = (\\d+)'))
  if (m === null) throw new Error('const not found: ' + name)
  return 'var ' + name + ' = ' + m[1]
}

const code = [
  constOf('MAX_ROWS'), constOf('MAX_COLS'), constOf('MAX_CELLS'), constOf('MAX_EXCERPT_ROWS'),
  'var CELL_CACHE = new WeakMap()',
  extract('function cellMapOf(workbook)'),
  extract('function cellTextOf(cell)'),
  extract('function activeSheetIndex(workbook)'),
  extract('function sheetNameAt(workbook, index)'),
  extract('function headerRowOf(map, sheetIndex, rg, cols)'),
  extract('function readRange(workbook, sheetName, rg)'),
  extract('function mdCell(s)'),
  extract('function composeTableBlock(snap)'),
  'return { readRange: readRange, composeTableBlock: composeTableBlock }',
].join('\n')

const api = new Function(code)()

let pass = 0, fail = 0
function eq(name, got, want) {
  const g = JSON.stringify(got), w = JSON.stringify(want)
  if (g === w) { pass++; console.log('  ok  ' + name) }
  else { fail++; console.log('  FAIL ' + name + '\n    got  ' + g + '\n    want ' + w) }
}

// 一张会撞见所有边角的表（0-based 行列）：
//   r1: 指标 | 2024 | 2025   ← 真表头
//   r2: 营业收入 | 100 | 120
//   r3: 净利润 | 10 | 12
//   r6: 只有A列的大标题       ← 合并单元格的典型长相，不能当表头
//   r7: x | y
const wb = [{
  name: 'Sheet1',
  status: 1,
  celldata: [
    { r: 1, c: 0, v: { m: '指标' } }, { r: 1, c: 1, v: { m: '2024' } }, { r: 1, c: 2, v: { m: '2025' } },
    { r: 2, c: 0, v: { m: '营业收入' } }, { r: 2, c: 1, v: { m: '100' } }, { r: 2, c: 2, v: { m: '120' } },
    { r: 3, c: 0, v: { m: '净利润' } }, { r: 3, c: 1, v: { m: '10' } }, { r: 3, c: 2, v: { m: '12' } },
    { r: 6, c: 0, v: { m: '只有A列的大标题' } },
    { r: 7, c: 0, v: { m: 'x' } }, { r: 7, c: 1, v: { m: 'y' } },
  ],
}]

console.log('— 表头探测 —')

const rg = { r1: 2, r2: 3, c1: 0, c2: 2, sheet: 'Sheet1' }
const got = api.readRange(wb, 'Sheet1', rg)
eq('① 选数据区（r2..r3）→ 带上 r1 那行做表头', got.header, ['指标', '2024', '2025'])
eq('① headerRow 指向选区上方那一行（0-based，显示时 +1）', got.headerRow, 1)

eq('② 上方只有 A 列有字（大标题）→ 不当表头',
  api.readRange(wb, 'Sheet1', { r1: 7, r2: 7, c1: 0, c2: 2, sheet: 'Sheet1' }).header, null)

eq('③ 从第一行开始选 → 没有上方一行',
  api.readRange(wb, 'Sheet1', { r1: 0, r2: 1, c1: 0, c2: 2, sheet: 'Sheet1' }).header, null)

eq('④ 单列选区放宽到 1 个非空格',
  api.readRange(wb, 'Sheet1', { r1: 2, r2: 3, c1: 0, c2: 0, sheet: 'Sheet1' }).header, ['指标'])

eq('⑤ 上方整行是空的 → 不硬凑',
  api.readRange(wb, 'Sheet1', { r1: 6, r2: 6, c1: 0, c2: 2, sheet: 'Sheet1' }).header, null)

console.log('— 拼装 —')

const snap = {
  kind: 'excel', path: 'D:\\wendang\\p.xlsx', sheet: 'Sheet1', address: 'A3:C4',
  table: got.table, header: got.header, headerRow: got.headerRow, useHeader: true,
}
eq('⑥ 表头写进 Markdown 表头行，并注明来源行',
  api.composeTableBlock(snap), [
    '> \uD83D\uDCC4 引用来源：D:\\wendang\\p.xlsx（工作表 Sheet1 · 区域 A3:C4）',
    '> （表头取自选区上方第 2 行，已一并带上）',
    '> | 指标 | 2024 | 2025 |',
    '> | --- | --- | --- |',
    '> | 营业收入 | 100 | 120 |',
    '> | 净利润 | 10 | 12 |',
  ])

eq('⑦ 卡片上取消勾选 → 原样发选区，不塞表头',
  api.composeTableBlock(Object.assign({}, snap, { useHeader: false })), [
    '> \uD83D\uDCC4 引用来源：D:\\wendang\\p.xlsx（工作表 Sheet1 · 区域 A3:C4）',
    '> | 营业收入 | 100 | 120 |',
    '> | --- | --- | --- |',
    '> | 净利润 | 10 | 12 |',
  ])

eq('⑧ 没有表头候选时输出与改动前一致',
  api.composeTableBlock({
    kind: 'excel', path: 'D:\\wendang\\p.xlsx', sheet: 'Sheet1', address: 'A8:C8',
    table: [['x', 'y', '']], header: null, headerRow: 6, useHeader: false,
  }), [
    '> \uD83D\uDCC4 引用来源：D:\\wendang\\p.xlsx（工作表 Sheet1 · 区域 A8:C8）',
    '> x | y',
  ])

console.log('\n' + (fail === 0 ? 'ALL PASS' : 'FAILED') + ' — ' + pass + ' passed, ' + fail + ' failed')
process.exit(fail === 0 ? 0 : 1)
