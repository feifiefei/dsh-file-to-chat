// 回归用例：长文本折叠的判定与文案（isLongText / longTextLabel / longTextPreview）。
//
// 这三支是纯函数，抠出来直接跑。判定边界都写进用例里：
//   - 行数 ≥ 12 或 字数 ≥ 600 才算「长」；
//   - 标题里给出真实的「几行 / 几字」，模型与用户看的是同一组数；
//   - 摘要取第一个非空行、超 60 字截断（tooltip 会用到）。
//
// 跑法：node test/long.test.cjs   （或 npm test）
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

function numConst(name) {
  const m = new RegExp('var ' + name + ' = (\\d+)').exec(src)
  if (m === null) throw new Error('const not found: ' + name)
  return m[1]
}

const code = [
  'var G = { refBodies: new Map() }',
  'var LONG_TEXT_LINES = ' + numConst('LONG_TEXT_LINES'),
  'var LONG_TEXT_CHARS = ' + numConst('LONG_TEXT_CHARS'),
  extract('function isLongText(text)'),
  extract('function longTextLabel(text)'),
  extract('function longTextPreview(text, max)'),
  extract('function pasteSplit(text)'),
  'return { isLongText: isLongText, longTextLabel: longTextLabel, longTextPreview: longTextPreview, pasteSplit: pasteSplit, G: G }',
].join('\n')

const api = new Function(code)()

let pass = 0, fail = 0
function eq(name, got, want) {
  const g = JSON.stringify(got), w = JSON.stringify(want)
  if (g === w) { pass++; console.log('  ok  ' + name) }
  else { fail++; console.log('  FAIL ' + name + '\n    got  ' + g + '\n    want ' + w) }
}

const line = (n) => new Array(n).fill('一行').join('\n')

console.log('— 判定 —')
eq('① 短句不折', api.isLongText('帮我看下这个函数'), false)
eq('② 空串不折', api.isLongText(''), false)
eq('③ null 不折', api.isLongText(null), false)
eq('④ 11 行仍不折', api.isLongText(line(11)), false)
eq('⑤ 12 行开始折', api.isLongText(line(12)), true)
eq('⑥ 599 字不折', api.isLongText('字'.repeat(599)), false)
eq('⑦ 600 字开始折', api.isLongText('字'.repeat(600)), true)
eq('⑧ 单行超长也折（无换行）', api.isLongText('a'.repeat(900)), true)
eq('⑨ 带 CRLF 的 12 行', api.isLongText(new Array(12).fill('x').join('\r\n')), true)

console.log('— 标题 —')
eq('⑩ 行数与字数都是真实的', api.longTextLabel('a\nb\nc'), '长文本 · 3 行 · 5 字')
eq('⑪ 单行', api.longTextLabel('abc'), '长文本 · 1 行 · 3 字')

console.log('— 摘要 —')
eq('⑫ 跳过开头空行', api.longTextPreview('\n\n  第一行有内容  \n第二行'), '第一行有内容')
eq('⑬ 超 60 字截断', api.longTextPreview('x'.repeat(80)), 'x'.repeat(60) + '…')
eq('⑭ 全空 → 空串', api.longTextPreview('\n \n'), '')
eq('⑮ 刚好 60 字不加省略号', api.longTextPreview('y'.repeat(60)), 'y'.repeat(60))
eq('⑯ 卡片上只用 30 字', api.longTextPreview('z'.repeat(50), 30), 'z'.repeat(30) + '…')

// 这一组是为「胶囊吃掉我后边打的字」加的：整条消息折成一枚卡片时，用户自己打的
// 那些字必须还能看见。所以折之前得先认出「哪一段是粘进来的」。
console.log('— 粘贴段的边界（pasteSplit）—')
{
  const body = new Array(22).fill('select b.project_name, a.xxx from t').join('\n')
  const M = api.G.refBodies
  const reset = (...pairs) => { M.clear(); for (const [k, v] of pairs) M.set(k, v) }

  reset(['p1', body])
  eq('⑰ 整条就是粘贴的', api.pasteSplit(body), { before: '', body: body, after: '' })

  reset(['p1', body])
  eq('⑱ 粘贴后边又打了一句', api.pasteSplit(body + '\n这段 SQL 帮我看下性能'),
    { before: '', body: body, after: '\n这段 SQL 帮我看下性能' })

  reset(['p1', body])
  eq('⑲ 粘贴前边也打了字', api.pasteSplit('看看这个：\n' + body),
    { before: '看看这个：\n', body: body, after: '' })

  reset()
  eq('⑳ 认不出（刷新过、没有正文表）→ null', api.pasteSplit(body), null)

  reset(['q1', body])
  eq('㉑ 只认粘贴进来的（p 键），引用正文不算', api.pasteSplit(body), null)

  const short = '短短的粘贴'
  reset(['p1', short], ['p2', body])
  eq('㉒ 两段粘贴取最长的那段', api.pasteSplit(short + '\n中间\n' + body),
    { before: short + '\n中间\n', body: body, after: '' })

  reset(['p1', body.replace(/\n/g, '\r\n')])
  eq('㉓ CRLF 归一后仍认得出', api.pasteSplit(body + '\n后面的话'),
    { before: '', body: body, after: '\n后面的话' })

  reset(['p1', body])
  eq('㉔ 正文被宿主改过一个字 → 不猜', api.pasteSplit(body.replace('project_name', 'project_name ')), null)
}

console.log('')
console.log(pass + ' passed, ' + fail + ' failed')
process.exit(fail === 0 ? 0 : 1)
