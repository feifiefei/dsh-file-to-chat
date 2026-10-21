// 回归用例：元数据自洽（版本号两处一致、发布包里没有测试、cordis 补丁还在）。
//
// 这几条不测逻辑，只测「改了一处忘了另一处」这类漂移 —— 版本号两处不一致
// 会让体检卡片报错的版本，是最难解释的一种 bug。
//
// 跑法：node test/meta.test.cjs   （或 npm test）
const fs = require('fs')
const path = require('path')
const root = path.join(__dirname, '..')
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'))
const src = fs.readFileSync(path.join(root, 'client.js'), 'utf8')

let pass = 0, fail = 0
function eq(name, got, want) {
  const g = JSON.stringify(got), w = JSON.stringify(want)
  if (g === w) { pass++; console.log('  ok  ' + name) }
  else { fail++; console.log('  FAIL ' + name + '\n    got  ' + g + '\n    want ' + w) }
}

console.log('— 元数据 —')

const m = src.match(/var F2C_VERSION = '([^']+)'/)
eq('① client.js 里有 F2C_VERSION', m !== null, true)
eq('② client.js 的版本号与 package.json 一致', m === null ? null : m[1], pkg.version)

eq('③ 发布白名单不含 test/', Array.isArray(pkg.files) && pkg.files.includes('test'), false)
eq('④ 测试脚本列全了所有用例文件',
  pkg.scripts.test.split('&&').map(s => s.trim().replace(/^node /, '')).sort(),
  fs.readdirSync(path.join(root, 'test')).filter(f => f.endsWith('.cjs')).map(f => 'test/' + f).sort())

eq('⑤ 客户端平台仍是 web', pkg.dsh.client.platform, 'web')
eq('⑥ cordis 补丁挂在 bundle 上', pkg.dsh.bundle.patch, './cordis.patch.yml')
eq('⑦ 补丁文件真的在', fs.existsSync(path.join(root, 'cordis.patch.yml')), true)
eq('⑧ 入口文件真的在', fs.existsSync(path.join(root, pkg.main)), true)

console.log('\n' + (fail === 0 ? 'ALL PASS' : 'FAILED') + ' — ' + pass + ' passed, ' + fail + ' failed')
process.exit(fail === 0 ? 0 : 1)
