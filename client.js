// dsh-file-to-chat 的浏览器端 half（client bundle）。
//
// 手写 CJS + ModuleLoader 包装（零构建步骤）：pure DOM 自渲染，
// 无任何 @deepseek-ai 值导入；cordis 服务经 exports.inject 的字符串名接入。
//
// ============================ 功能范围 ============================
// 划词/划区批注：在右侧「文档预览」窗格里的三类渲染器上取选区 ->
//   选区上方浮出「批注」按钮 -> 点击 -> 批注卡片 -> 发送
//   -> 批注 + 选区 + 来源(路径 + 行区间/页码/单元格区域) 一起写进当前会话 composer。
//   ① P2a 纯文本类：.txt/.md/.json/代码…（text-pages 渲染器），定位到「第 N 行」
//   ② P2b 表格类：.xlsx/.xls/.csv/.tsv（FortuneSheet，canvas），定位到「工作表 X · 区域 B3:D7」
//   ③ P2c Word/PPT/PDF：宿主 LibreOffice 转 PDF 后由 PDF.js 渲染，定位到「第 N 页」
//
// ==================== 右侧文档预览窗格的 DOM 契约 ====================
// 事实来源：dsh-file-quote@0.1.1 lib/client.js:993-1094（已在实机验证过的兼容层）
//   + dsh-client-ui-sidebar-documentpreview/lib/client.js @offset 37537（实机微出）
//
// 关键：那个类名叫 TextPreview 的组件其实是**所有渲染器共用的外壳**——
//   div.preview[data-document-preview="<渲染器 id>"][data-textpreview-state="text"]
//     ├─ header: PathLabel[data-textpreview-path] + renderSlot("…document.actions")
//     └─ div.body[data-textpreview-body] -> <选中的渲染器>
//   所以 .txt / .xlsx / .docx 三种预览**都**有 [data-document-preview] 与
//   [data-textpreview-path]；差别只在里面挂的是哪个渲染器。
//   · 窗格根          [data-document-preview]      ← 三种渲染器都有
//   · 正文            [data-textpreview-body]      ← 三种渲染器都有
//   · 文件路径标签    [data-textpreview-path]      title = 绝对/展示路径，三种都有
//   · 纯文本体        [data-textpreview-plain]     ← dsh-annotation 的 DOCUMENT_TEXT
//   · 纯文本行        <div data-textpreview-line="N">  N = 1-based 源码行号
//   · 代码高亮        [data-code-preview] -> [data-code-block-content] -> pre > span.line
//   · Markdown        [data-document-markdown]
//   · 表格（P2b）     .fortune-sheet-container / .fortune-workarea（FortuneSheet）
//                     选区地址在 .fortune-name-box；真区域在 React fiber 的 context
//   · PDF（P2c）      section[data-pdf-preview] > div[data-pdf-page="N"]（自绘，非官方 viewer）
//                     > div[data-document-zoom-surface] > div[data-pdf-text] > .textLayer > span…
//                     文字是真实文本层 ⇒ getSelection() 可选中；只有页、没有源码行号
//
// 定位教训（dsh-annotation client.js:864-881）：原生选中菜单锚定在选区上沿，且原生 UI
// 恒绘制在页面内容之上（z-index 对它无效）。桌面端没有这个浮层，所以按用户要求「上方优先」，
// 放不下时才回落到下方。
//
// ==================== composer 注入 ====================
// 座位 standardProps 直接提供 inputActions / sessionId / useInput
// （client Slots Inspect listSubTree 已核实），优先走它；
// 兜底走 ctx.conversation.input.for(sessions.scope(id))，
// 来源：@changfenhuang/dsh-annotation v1.4.11 client.js:1945-1963。
window.__ModuleLoader__.load({
  // 必须与 package.json "name" 完全一致，否则 client-modules 报：
  // bundle loaded without registering "dsh-file-to-chat"
  id: 'dsh-file-to-chat',
  factory: (require) => {
    'use strict'
    var module = { exports: {} }
    var exports = module.exports

    var React = require('react')

    var Z_BAR = 1200
    var Z_CARD = 1201
    var Z_MENU = 1202
    var Z_NOTE = 1203
    var MAX_EXCERPT_LINES = 40
    var MAX_EXCERPT_ROWS = 60    // Excel 选区里最多渲染成表格的行数

    // ================================================================
    // 诊断开关
    // ================================================================
    // 默认关闭：正常使用时控制台是干净的，只有真正的错误才 warn。
    // 需要排查时在控制台执行一次
    //     localStorage.setItem('dsh-file-to-chat:debug', '1')
    // 再刷新页面即可看到全部细节日志；执行 removeItem 关回去。
    var DEBUG = false
    try {
      if (typeof localStorage !== 'undefined' && localStorage !== null) {
        DEBUG = localStorage.getItem('dsh-file-to-chat:debug') === '1'
      }
    } catch (err) { void err }

    /** 细节日志。只有 DEBUG 打开时才产生输出。 */
    function dbg() {
      if (!DEBUG) return
      try {
        var args = Array.prototype.slice.call(arguments)
        console.log.apply(console, ['[file-to-chat]'].concat(args))
      } catch (err) { void err }
    }

    // ================================================================
    // 一、无依赖的小工具
    // ================================================================
    function nodeToElement(node) {
      if (node === null || node === undefined) return null
      return node.nodeType === 1 ? node : (node.parentElement || null)
    }

    function truncate(s, n) {
      return s.length > n ? s.slice(0, n) + '…' : s
    }

    function isAbsPath(p) {
      return /^([a-zA-Z]:[\\/]|\\\\|\/)/.test(p)
    }

    /** 把绝对路径尽量压成工作区相对路径；做不到就原样返回。 */
    function relPathOf(cwd, abs) {
      if (cwd === null || cwd === undefined || cwd === '') return abs
      var a = String(abs).replace(/\\/g, '/')
      var c = String(cwd).replace(/\\/g, '/').replace(/\/+$/, '')
      if (a.toLowerCase().indexOf(c.toLowerCase() + '/') === 0) return a.slice(c.length + 1)
      return abs
    }

    // ================================================================
    // 二、文档预览窗格的 DOM 读取
    // ================================================================
    /** 选区所在节点 -> 预览窗格根元素（配方同 dsh-file-quote client.js:999-1009） */
    function previewRootOf(node) {
      var start = nodeToElement(node)
      if (start === null || !start.closest) return null
      var root = start.closest('[data-document-preview]')
      if (root !== null) return root
      var body = start.closest('[data-textpreview-body]')
      if (body !== null) return body.parentElement || body
      var label = start.closest('[data-textpreview-path]')
      if (label !== null && label.parentElement !== null && label.parentElement.parentElement !== null) {
        return label.parentElement.parentElement
      }
      return null
    }

    /** 选区是否落在「文档预览」窗格内（划词的作用域限定） */
    function isInDocumentPreview(node) {
      var start = nodeToElement(node)
      if (start === null || !start.closest) return false
      if (start.closest('[data-document-preview]') !== null) return true
      if (start.closest('[data-textpreview-body]') !== null) return true
      if (start.closest('[data-textpreview-plain]') !== null) return true
      if (start.closest('[data-textpreview-line]') !== null) return true
      // Word/PPT/PDF：Office 那份 PDF 预览不带上面任何属性，只有 PDF.js 的文本层
      if (pdfPageOf(start) !== null) return true
      return false
    }

    /** 预览窗格显示的文件路径（绝对或相对）。 */
    function pathFromPreview(node) {
      try {
        var root = previewRootOf(node)
        if (root === null || !root.querySelector) return null
        var label = root.querySelector('[data-textpreview-path]')
        if (label === null) return null
        // PathLabel 的 title 是绝对路径/展示路径；原样带上即可 —— 对模型而言
        // 绝对路径同样无歧义，且省掉一次 cwd 查询。
        var raw = label.getAttribute('title') || label.textContent || ''
        var p = String(raw).trim()
        return p === '' ? null : p
      } catch (_) { return null }
    }

    /** 某一节点所属的 1-based 源码行号。 */
    function lineNumberOf(node) {
      try {
        var el = nodeToElement(node)
        if (el === null || !el.closest) return null
        var plain = el.closest('[data-textpreview-line]')
        if (plain !== null) {
          var n = Number(plain.getAttribute('data-textpreview-line'))
          if (isFinite(n) && n >= 1) return n
        }
        var line = el.closest('.line')
        if (line !== null) {
          var scope = line.closest('[data-code-block-content]') || line.closest('pre') || line.parentElement
          if (scope !== null && scope.querySelectorAll) {
            var all = scope.querySelectorAll('.line')
            var index = Array.prototype.indexOf.call(all, line)
            if (index >= 0) return index + 1
          }
        }
        return null
      } catch (_) { return null }
    }

    /** 选区端点常落在容器元素上（跨块选择）：就近取一个子节点的行号。 */
    function lineNear(node, offset) {
      var n = lineNumberOf(node)
      if (n !== null) return n
      if (node === null || node === undefined || node.nodeType !== 1 || !node.childNodes) return null
      var kids = node.childNodes
      var at = Math.min(typeof offset === 'number' ? offset : kids.length, kids.length)
      var i
      for (i = at - 1; i >= 0; i--) { n = lineNumberOf(kids[i]); if (n !== null) return n }
      for (i = at; i < kids.length; i++) { n = lineNumberOf(kids[i]); if (n !== null) return n }
      return null
    }

    /** 选区的源码行区间 [lo, hi]，取不到返回 null。 */
    function lineRangeOf(range) {
      try {
        var a = lineNear(range.startContainer, range.startOffset)
        var b = lineNear(range.endContainer, range.endOffset)
        if (a === null && b === null) return null
        var lo = a === null ? b : a
        var hi = b === null ? a : b
        return [Math.min(lo, hi), Math.max(lo, hi)]
      } catch (_) { return null }
    }

    // ---- Word / PPT / PDF：走 LibreOffice→PDF，只有页、没有源码行号 ----
    // .docx/.doc/.ppt/.pptx 先由宿主 @deepseek-ai/dsh-office-to-pdf 转成 PDF，
    // 再交给 PDF.js 渲染；.pdf 直接走同一条路。
    //
    // DSH 的 PdfBody **不是** PDF.js 官方 viewer 那套 .page[data-page-number]，
    // 是自绘的：每页一个 canvas + 一个独立的文本层覆盖框。实测（client.pdf.js
    // 尾部 PdfPage 的 JSX，偏移 ≈7106610）结构是：
    //   section[data-pdf-preview]
    //     └ div[data-pdf-page="N"]                     ← 页码就挂在这儿（稳定标记）
    //         └ div[data-document-zoom-surface]
    //             ├ canvas[role=img]
    //             └ div[data-pdf-text]                 ← 文本层宿主
    //                 └ div.textLayer > span…          ← 真正可选的文字
    // 类名（ZhMWpq_page 等）是 CSS-module 哈希，不能依赖；上面三个 data-* 才是契约。
    // 文字是真实文本层（不是 canvas 画的）⇒ window.getSelection() 能选到；
    // 但页面是按排版切的，跟源文件行号无从对应 —— 只能定位到「第 N 页」。
    var PDF_PAGE = '[data-pdf-page]'
    var PDF_PAGE_ATTR = 'data-pdf-page'

    function pdfPageOf(node) {
      try {
        var el = nodeToElement(node)
        if (el === null || !el.closest) return null
        var layer = el.closest('.textLayer')
        if (layer === null) return null
        return layer.closest(PDF_PAGE)
      } catch (_) { return null }
    }

    /** 某一节点所属的 1-based 页码（非 PDF 预览返回 null）。 */
    function pageNumberOf(node) {
      try {
        var page = pdfPageOf(node)
        if (page === null) return null
        var n = Number(page.getAttribute(PDF_PAGE_ATTR))
        return isFinite(n) && n >= 1 ? n : null
      } catch (_) { return null }
    }

    function pageNear(node, offset) {
      var n = pageNumberOf(node)
      if (n !== null) return n
      if (node === null || node === undefined || node.nodeType !== 1 || !node.childNodes) return null
      var kids = node.childNodes
      var at = Math.min(typeof offset === 'number' ? offset : kids.length, kids.length)
      var i
      for (i = at - 1; i >= 0; i--) { n = pageNumberOf(kids[i]); if (n !== null) return n }
      for (i = at; i < kids.length; i++) { n = pageNumberOf(kids[i]); if (n !== null) return n }
      return null
    }

    /** 选区的页码区间 [lo, hi]，取不到返回 null。 */
    function pageRangeOf(range) {
      try {
        var a = pageNear(range.startContainer, range.startOffset)
        var b = pageNear(range.endContainer, range.endOffset)
        if (a === null && b === null) return null
        var lo = a === null ? b : a
        var hi = b === null ? a : b
        return [Math.min(lo, hi), Math.max(lo, hi)]
      } catch (_) { return null }
    }

    // ================================================================
    // 三、选区快照
    // ================================================================
    /**
     * 取一次选区快照。**必须在任何点击之前调用**：点击会清掉原生选区，
     * 快照里的文本/路径/行号/坐标之后都还能用。
     * @param fallbackPath - 预览窗格没有 [data-textpreview-path] 时用的路径。
     *   Word/PPT/PDF 预览正是这种情况，靠 `document.actions` 座位发给工具栏的
     *   absolutePath 兜底（见 bridge.absolutePath）。
     */
    function captureSelection(fallbackPath) {
      var sel
      try { sel = window.getSelection() } catch (_) { return null }
      if (sel === null || sel === undefined) return null
      if (sel.isCollapsed === true || sel.rangeCount === 0) return null
      var text = ''
      try { text = sel.toString() } catch (_) { return null }
      if (text === null || text.replace(/\s+/g, '') === '') return null

      var range
      try { range = sel.getRangeAt(0) } catch (_) { return null }
      var startEl = nodeToElement(range.startContainer)
      if (startEl === null) return null
      if (!isInDocumentPreview(startEl)) return null
      if (startEl.closest && startEl.closest('button, input, textarea, [contenteditable="true"]') !== null) return null

      // 多行选区用首/末行的矩形定位，比并集矩形准
      var first = null
      var last = null
      try {
        var rects = range.getClientRects()
        if (rects !== null && rects.length > 0) {
          first = rects[0]
          last = rects[rects.length - 1]
        }
      } catch (_) { /* 落到并集矩形 */ }
      if (first === null) {
        try {
          var r = range.getBoundingClientRect()
          if (r.width === 0 && r.height === 0 && r.top === 0) return null
          first = r
          last = r
        } catch (_) { return null }
      }

      var fromPreview = pathFromPreview(startEl)
      var fb = fallbackPath === null || fallbackPath === undefined ? '' : String(fallbackPath).trim()

      return {
        kind: 'text',
        text: text,
        range: range,
        path: fromPreview !== null && fromPreview !== '' ? fromPreview : (fb === '' ? null : fb),
        lineRange: lineRangeOf(range),
        pageRange: pageRangeOf(range),
        top: first.top,
        bottom: last.bottom,
        left: Math.min(first.left, last.left),
        right: Math.max(first.right, last.right),
      }
    }

    // ================================================================
    // 三·B、Excel 预览的选区
    // ================================================================
    // 为什么不能沿用上面那套划词：Excel 预览是 FortuneSheet，**单元格画在
    // canvas 上**，文字根本不在 DOM 里 —— window.getSelection() 永远读不到
    // 单元格，也没有 [data-textpreview-url] 可挂。这正是同类插件一律把
    // Office/Excel 归为「只能整文件」的物理原因。
    //
    // 但 FortuneSheet 自己维护着选区，并且把它渲染进一个**真实 DOM 节点**：
    // 名称框 .fortune-name-box（源码 getRangetxt(context, sheetId, lastSelection)）。
    // 拖选时 React 实时更新它的文本，形如 "B3:D7" 或 "Sheet1!B3:D7"。
    // 所以我们只读这一个节点的 textContent 就拿到了选区范围 —— 不做 canvas
    // 命中测试（那要靠猜列宽/行高/滚动偏移，必漂）。
    //
    // 数值则直接从 React 的 workbook 数据里取（见 readRange）：不走网络、
    // 不解析文件、不触发重渲染，点开卡片时才算一次，并按 workbook 缓存。

    var FORTUNE_HOST = '.fortune-sheet-container, .fortune-workarea'
    var MAX_COLS = 60            // 单次选区允许的最大列数（性能护栏）
    var MAX_ROWS = 300           // 单次选区允许的最大行数
    var MAX_CELLS = 3000         // 单元格总数上限，超出即拒绝取值

    // 工具栏 / 工作表标签 / 名称框 / 编辑栏 / 状态栏都点在「不是网格」的地方。
    // 注意：这里只列 .fortune-* 的精确类名，不用 [class*="toolbar"] 这种模糊
    // 匹配 —— 预览窗格自己的头部工具条也可能命中，那会把整个网格误排除。
    var FORTUNE_CHROME = '.fortune-toolbar, .fortune-sheet-list, .fortune-name-box,'
      + ' .fortune-formula-box, .fortune-stat-area, .fortune-sheet-tab,'
      + ' input, textarea, select, button'

    /**
     * 一次指针释放是否落在表格网格上；是则返回 fortune 容器。
     *
     * 关键：返回**文档序最外层**的容器，而不是 closest() 给的最近祖先。
     * 实测结构（client.excel.js:4171171 / 4468250）：
     *   <div class="fortune-workarea">          ← Toolbar（含名称框）在这层
     *     <div class="fortune-sheet-container">
     *       <canvas class="fortune-sheet-canvas" aria-hidden="true">   ← 纯绘制
     *       <SheetOverlay/>                                            ← 真正命中目标
     * 用 closest() 会命中内层 container，而名称框在外层 workarea 里，
     * 在那里 querySelector('.fortune-name-box') 恒为 null —— 表格就永远不弹按钮。
     */
    function excelHostOf(node) {
      var el = nodeToElement(node)
      if (el === null || el.closest === undefined) return null
      if (el.closest(FORTUNE_CHROME) !== null) return null
      var near = el.closest(FORTUNE_HOST)
      if (near === null) return null
      var root = previewRootOf(near)
      if (root === null || !root.querySelector) return near
      return root.querySelector(FORTUNE_HOST) || near
    }

    function activeExcelHost() {
      var root = document.querySelector('[data-document-preview]')
      if (root === null) return null
      return root.querySelector(FORTUNE_HOST)
    }

    function colIndexOf(letters) {
      var n = 0
      for (var i = 0; i < letters.length; i++) {
        var c = letters.charCodeAt(i)
        if (c >= 97) c -= 32
        if (c < 65 || c > 90) return -1
        n = n * 26 + (c - 64)
      }
      return n - 1
    }

    function colNameOf(index) {
      var s = ''
      var n = index
      while (n >= 0) {
        s = String.fromCharCode(65 + (n % 26)) + s
        n = Math.floor(n / 26) - 1
      }
      return s
    }

    /** "Sheet1!B3:D7" / "B3:D7" / "$B$3" → { sheet, r1, c1, r2, c2 }，行列均 0 基。 */
    function parseAddress(text) {
      if (typeof text !== 'string') return null
      var s = text.trim()
      if (s === '') return null
      var sheet = null
      var bang = s.lastIndexOf('!')
      if (bang >= 0) {
        sheet = s.slice(0, bang).replace(/^'/, '').replace(/'$/, '').replace(/''/g, "'")
        s = s.slice(bang + 1)
      }
      var m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?$/.exec(s.trim())
      if (m === null) return null
      var c1 = colIndexOf(m[1])
      var c2 = m[3] === undefined ? c1 : colIndexOf(m[3])
      var r1 = parseInt(m[2], 10) - 1
      var r2 = m[4] === undefined ? r1 : parseInt(m[4], 10) - 1
      if (c1 < 0 || c2 < 0 || r1 < 0 || r2 < 0) return null
      return {
        sheet: sheet,
        r1: Math.min(r1, r2), r2: Math.max(r1, r2),
        c1: Math.min(c1, c2), c2: Math.max(c1, c2),
      }
    }

    function addressOf(rg) {
      var a = colNameOf(rg.c1) + (rg.r1 + 1)
      var b = colNameOf(rg.c2) + (rg.r2 + 1)
      return a === b ? a : a + ':' + b
    }

    function activeSheetName(host) {
      var el = host.querySelector('.fortune-sheet-list-item-active .fortune-sheet-list-item-name')
        || host.querySelector('[class*="sheet-list-item-active"]')
      if (el === null) return null
      var t = (el.textContent || '').trim()
      return t === '' ? null : t
    }

    // ---- 直接向 FortuneSheet 要「真实选区」 ----
    // 名称框（.fortune-name-box）在**选区的焦点格正好是合并单元格的锚点**时，会
    // 故意只显示那一个格子。源码（bundle 里的 LocationBox）：
    //     var rf = lastSelection.row_focus, cf = lastSelection.column_focus;
    //     if (config.merge != null && (rf + "_" + cf) in config.merge)
    //       return getRangetxt(ctx, currentSheetId, { column: [cf, cf], row: [rf, rf] });
    // 于是「从合并单元格起拖」在名称框里会塌成一个锚点地址，真实拖出来的范围被丢掉。
    // 真实选区一直完好地存在 context.luckysheet_select_save 里 —— 去那里拿。
    //
    // 取法：名称框所在的 LocationBox 组件用 useContext(WorkbookContext) 拿到
    // { context, setContext, refs }。React 把内部实例挂在 DOM 节点上，顺着 fiber
    // 向上找「值里带 .context.luckysheet_select_save」的那一层即可。全程只读，
    // 不改 React 状态、不触发重渲染。

    /** 认一认「WorkbookContext 的值」这个形状，认出来就返回里面的 context。 */
    function fortuneContextBag(v) {
      if (v === null || typeof v !== 'object') return null
      var ctx = v.context
      if (ctx === null || typeof ctx !== 'object') return null
      if (ctx.luckysheet_select_save === undefined) return null
      return ctx
    }

    function fortuneContextOf(host) {
      var box = host.querySelector('.fortune-name-box')
        || host.querySelector('.fortune-name-box-container')
        || host
      var node = reactFiberOf(box)
      if (node === null) node = reactFiberOf(host)
      for (var depth = 0; node !== null && node !== undefined && depth < 200; depth++) {
        // Context.Provider 把值放在 memoizedProps.value 上
        var bag = fortuneContextBag(node.memoizedProps !== null && node.memoizedProps !== undefined
          ? node.memoizedProps.value
          : undefined)
        if (bag !== null) return bag
        // 函数组件的话，值挂在 useContext 那个 hook 的 memoizedState 上
        for (var h = node.memoizedState, n = 0; h !== null && h !== undefined && n < 60; h = h.next, n++) {
          bag = fortuneContextBag(h.memoizedState)
          if (bag !== null) return bag
        }
        node = node.return
      }
      return null
    }

    /** 活动工作表名：以 context.currentSheetId 为准，去 luckysheetfile 里查名字。 */
    function fortuneSheetName(ctx) {
      var files = ctx.luckysheetfile
      if (!Array.isArray(files)) return null
      for (var i = 0; i < files.length; i++) {
        var f = files[i]
        if (f !== null && f !== undefined && f.id === ctx.currentSheetId) {
          return typeof f.name === 'string' && f.name !== '' ? f.name : null
        }
      }
      return null
    }

    /** 真实选区 → { sheet, r1, c1, r2, c2 }（0 基、含端点）。拿不到返回 null。 */
    function fortuneSelection() {
      var host = activeExcelHost()
      if (host === null) return null
      var ctx = fortuneContextOf(host)
      if (ctx === null) return null
      var save = ctx.luckysheet_select_save
      if (!Array.isArray(save) || save.length === 0) return null
      var last = save[save.length - 1]          // 与名称框取法一致：最后一段
      if (last === null || typeof last !== 'object') return null
      var row = last.row
      var col = last.column
      if (!Array.isArray(row) || !Array.isArray(col) || row.length < 2 || col.length < 2) return null
      var r1 = Math.min(row[0], row[1]), r2 = Math.max(row[0], row[1])
      var c1 = Math.min(col[0], col[1]), c2 = Math.max(col[0], col[1])
      if (!(r1 >= 0 && c1 >= 0 && r2 >= r1 && c2 >= c1)) return null
      return { sheet: fortuneSheetName(ctx), r1: r1, c1: c1, r2: r2, c2: c2 }
    }

    /**
     * 读一次 Excel 预览的选区。没开表格预览 / 读不到 / 解析失败 → null。
     * `fallbackPath` 是工具栏座位传下来的 absolutePath，路径拿不到时兜底。
     */
    function captureExcelSelection(pointer, fallbackPath) {
      var host = activeExcelHost()
      if (host === null) return null

      // 首选 FortuneSheet 的真实选区 —— 合并单元格起拖也不会塌成锚点。
      var rg = fortuneSelection()
      if (rg === null) {
        // 兜底：名称框。普通选区它是对的，只有「焦点格是合并单元格锚点」时会塌。
        var box = host.querySelector('.fortune-name-box')
        if (box === null) return null
        rg = parseAddress((box.textContent || '').trim())
        if (rg === null) return null
        if (rg.sheet === null) rg.sheet = activeSheetName(host)
      }

      var path = pathFromPreview(host)
      if ((path === null || path === '') && fallbackPath !== null && fallbackPath !== undefined) {
        path = fallbackPath
      }
      var x = pointer !== null && pointer !== undefined ? pointer.x : 0
      var y = pointer !== null && pointer !== undefined ? pointer.y : 0
      return {
        kind: 'excel',
        host: host,
        path: path,
        sheet: rg.sheet,
        range: rg,
        address: addressOf(rg),
        top: y, bottom: y, left: x, right: x,
      }
    }

    // ---- 从 React 数据里取单元格数值 ----
    // FortuneSheet 把整本工作簿放在 React 组件的 props.data 里；React 会在
    // 宿主 DOM 节点上挂内部字段 __reactFiber$xxx。顺着它向上找最近的、
    // props.data 是数组的那一层，就是工作簿。纯读，不改，不触发重渲染。

    function reactFiberOf(el) {
      if (el === null || el === undefined) return null
      var keys = Object.keys(el)
      for (var i = 0; i < keys.length; i++) {
        var k = keys[i]
        if (k.indexOf('__reactFiber$') === 0 || k.indexOf('__reactInternalInstance$') === 0) {
          return el[k]
        }
      }
      return null
    }

    /** 向上找工作簿数据（props.data）。找不到返回 null。 */
    function workbookOf(host) {
      var el = host
      var hops = 0
      while (el !== null && el !== undefined && hops < 40) {
        var node = reactFiberOf(el)
        var depth = 0
        while (node !== null && node !== undefined && depth < 80) {
          var p = node.memoizedProps
          if (p !== null && p !== undefined && Array.isArray(p.data) && p.data.length > 0
              && p.data[0] !== null && typeof p.data[0] === 'object') {
            return p.data
          }
          node = node.return
          depth++
        }
        el = el.parentElement
        hops++
      }
      return null
    }

    var CELL_CACHE = new WeakMap()   // workbook → { "sheet|r:c" → 单元格对象 }

    function cellMapOf(workbook) {
      var cached = CELL_CACHE.get(workbook)
      if (cached !== undefined) return cached
      var map = Object.create(null)
      for (var i = 0; i < workbook.length; i++) {
        var sheet = workbook[i]
        if (sheet === null || sheet === undefined) continue
        var cd = sheet.celldata
        if (Array.isArray(cd)) {
          for (var j = 0; j < cd.length; j++) {
            var it = cd[j]
            if (it === null || it === undefined) continue
            map[i + '|' + it.r + ':' + it.c] = it.v
          }
        }
        var grid = sheet.data
        if (Array.isArray(grid)) {
          for (var r = 0; r < grid.length; r++) {
            var row = grid[r]
            if (!Array.isArray(row)) continue
            for (var c = 0; c < row.length; c++) {
              var cell = row[c]
              if (cell === null || cell === undefined) continue
              var key = i + '|' + r + ':' + c
              if (map[key] === undefined) map[key] = cell
            }
          }
        }
      }
      CELL_CACHE.set(workbook, map)
      return map
    }

    /** 单元格的显示值：优先 m（格式化后的显示串），再退回 v。 */
    function cellTextOf(cell) {
      if (cell === null || cell === undefined) return ''
      var t = typeof cell
      if (t === 'string' || t === 'number' || t === 'boolean') return String(cell)
      if (t === 'object') {
        if (cell.m !== undefined && cell.m !== null && cell.m !== '') return String(cell.m)
        if (cell.v !== undefined && cell.v !== null) return String(cell.v)
      }
      return ''
    }

    /** 活动工作表的下标。Luckysheet / FortuneSheet 用 status === 1 标记当前表。 */
    function activeSheetIndex(workbook) {
      if (!Array.isArray(workbook)) return 0
      for (var i = 0; i < workbook.length; i++) {
        var s = workbook[i]
        if (s !== null && s !== undefined && (s.status === 1 || s.active === true)) return i
      }
      for (var j = 0; j < workbook.length; j++) {
        var t = workbook[j]
        if (t !== null && t !== undefined && t.order === 0) return j
      }
      return 0
    }

    function sheetNameAt(workbook, index) {
      if (!Array.isArray(workbook)) return null
      var s = workbook[index]
      if (s === null || s === undefined) return null
      return typeof s.name === 'string' && s.name !== '' ? s.name : null
    }

    /**
     * 按范围取值，返回 { table, sheet } 或 { error }。
     * 护栏：行列数与总数都封顶 —— 用户框选整列整行时不至于把浏览器拖死。
     */
    function readRange(workbook, sheetName, rg) {
      if (workbook === null) return { error: '没读到表格数据（预览可能刚切换，重开一次预览再试）' }

      var rows = rg.r2 - rg.r1 + 1
      var cols = rg.c2 - rg.c1 + 1
      if (rows > MAX_ROWS) return { error: '选了 ' + rows + ' 行，超过上限 ' + MAX_ROWS + ' 行' }
      if (cols > MAX_COLS) return { error: '选了 ' + cols + ' 列，超过上限 ' + MAX_COLS + ' 列' }
      if (rows * cols > MAX_CELLS) return { error: '选了 ' + (rows * cols) + ' 格，超过上限 ' + MAX_CELLS + ' 格' }

      var sheetIndex = -1
      if (sheetName !== null && sheetName !== undefined) {
        for (var i = 0; i < workbook.length; i++) {
          var s = workbook[i]
          if (s !== null && s !== undefined && s.name === sheetName) { sheetIndex = i; break }
        }
      }
      // 名字对不上时不能直接退回第 0 张：用户在第 3 张表上框选，会把第 1 张
      // 的数据**静默**发出去。改为按活动表定位。
      if (sheetIndex < 0) sheetIndex = activeSheetIndex(workbook)
      if (workbook[sheetIndex] === undefined) return { error: '工作表不存在' }

      var map = cellMapOf(workbook)
      var table = []
      for (var r = 0; r < rows; r++) {
        var line = []
        for (var c = 0; c < cols; c++) {
          line.push(cellTextOf(map[sheetIndex + '|' + (rg.r1 + r) + ':' + (rg.c1 + c)]))
        }
        table.push(line)
      }
      return { table: table, sheet: sheetNameAt(workbook, sheetIndex) }
    }

    // ================================================================
    // 四、样式
    // ================================================================
    var STYLE_ID = 'dsh-file-to-chat-style'

    // 主题 token —— 全部取自 DSH 自己的样式表，不是猜的：
    //   · ui-primitives/lib/MenuSurface.module.css  → 菜单材质
    //   · ui-primitives/lib/Menu.module.css         → 卡片/行/分隔/底栏
    //   · ui-primitives/lib/Input.module.css        → 输入框
    //   · ui-primitives/lib/Button.module.css       → 按钮
    //   · ui-theme/README.md                        → 材质与抬升的语义说明
    // 教训：曾经用过 --dsw-alias-bg-elevated / -border-secondary / -bg-sunken
    // 这些**不存在**的名字，后果是「浅色模式下 color 取到真实 token 的近黑
    // 值、background 回落到硬编码深色」→ 黑字压黑底。所以 fallback 必须是
    // **自洽的一对**（深底 + 浅字），即使一个 token 都没解析到也只是观感
    // 朴素，不会不可读。
    var T = {
      // 表面与材质
      surface: '--dsw-menu-surface-fill',        // 共享菜单材质填充（半透明）
      blur: '--dsw-menu-backdrop-filter',        // 配套背景模糊
      layer1: '--dsw-alias-bg-layer-1',          // 一级抬升面（引用块底 / 输入框底）
      layer2: '--dsw-alias-bg-layer-2',          // 二级嵌套面（折叠注释底）
      bgOverlay: '--dsw-alias-bg-overlay',       // overlay/popover 底（提示条用不透明底，比菜单材质清楚）
      hover: '--dsw-alias-interactive-bg-hover', // 行悬停填充（= 菜单项 hover）
      // 描边
      border1: '--dsw-alias-border-l1',          // 浅色菜单表面的抬升描边
      border2: '--dsw-alias-border-l2',          // 分隔线 / 底栏上边线
      border4: '--dsw-alias-border-l4',          // 输入框描边
      // 文字
      label: '--dsw-alias-label-primary',
      label2: '--dsw-alias-label-secondary',
      label3: '--dsw-alias-label-tertiary',      // 菜单分组标题 / 提示文字
      labelDim: '--dsw-alias-label-dimmed',      // placeholder
      menuIcon: '--dsw-alias-menu-icon',         // 菜单图标（中性偏蓝 800）
      fg: '--dsw-alias-label-primary-foreground',
      brand: '--dsw-alias-brand-primary',
      // 按钮
      btnFill: '--dsw-alias-button-primary-fill',
      btnHover: '--dsw-alias-button-primary-hover',
      // 圆角
      radiusLg: '--dsw-radius-lg',
      radiusMd: '--dsw-radius-md',
      radiusSm: '--dsw-radius-sm',
    }
    var FALLBACK = {
      surface: '#2b2b31',
      blur: 'none',
      layer1: '#3a3a42',
      layer2: '#232328',
      bgOverlay: '#2b2b31',
      hover: 'rgba(255,255,255,.07)',
      border1: 'rgba(255,255,255,.10)',
      border2: 'rgba(255,255,255,.16)',
      border4: 'rgba(255,255,255,.18)',
      label: '#ececf1',
      label2: '#a8a8b0',
      label3: '#7c7c86',
      labelDim: '#6a6a74',
      menuIcon: '#b9c2d8',
      fg: '#ffffff',
      brand: '#4d6bfe',
      btnFill: '#4d6bfe',
      btnHover: '#5f7bff',
      radiusLg: '12px',
      radiusMd: '8px',
      radiusSm: '6px',
    }
    function tok(key) {
      return 'var(' + T[key] + ', ' + FALLBACK[key] + ')'
    }
    // 行悬停填充。菜单里的 hover 是**叠在菜单表面之上**的一层半透明色
    // （.item 坐在 .list 的材质上），所以不能拿它直接替换 background ——
    // 那样会把材质一起抹掉。用 background-image 的渐变层压在
    // background-color 上面，得到和菜单项一模一样的叠色效果。
    function hoverFill() {
      var c = tok('hover')
      return 'linear-gradient(' + c + ', ' + c + ')'
    }
    // 抬升：--dsw-elevation-prominent 是「0.5px 细线 + 两层柔和阴影」的合成，
    // 细线颜色从 --dsw-elevation-stroke-color 重绑读取。所以表面必须自己
    // border:0、把 stroke-color 指到 border-l1 —— 否则会多出一圈占位描边。
    var ELEV = 'var(--dsw-elevation-prominent, 0 8px 28px rgba(0,0,0,.12), 0 2px 6px rgba(0,0,0,.06))'
    var SURFACE = [
      '  border: 0; background: ' + tok('surface') + ';',
      '  backdrop-filter: ' + tok('blur') + ';',
      '  -webkit-backdrop-filter: ' + tok('blur') + ';',
      '  --dsw-elevation-stroke-color: ' + tok('border1') + ';',
      '  box-shadow: ' + ELEV + ';',
    ].join('\n')

    var CSS = [
      // 浮层容器：全屏 fixed 但不拦事件，只有子元素 pointer-events:auto
      '.f2c-layer { position: fixed; inset: 0; pointer-events: none; z-index: ' + Z_BAR + '; }',
      '.f2c-layer .f2c-card { z-index: ' + Z_CARD + '; }',
      // 共享菜单材质：填充 + 背景模糊 + prominent 抬升
      '.f2c-card { position: fixed; pointer-events: auto; box-sizing: border-box;',
      '  width: 348px; padding: 4px; color: ' + tok('label') + ';',
      '  font-family: inherit; border-radius: ' + tok('radiusLg') + ';',
      SURFACE,
      '}',
      '.f2c-card.is-table { width: auto; min-width: 320px; max-width: 560px; }',

      // 文件树右键菜单：照 Menu.module.css 的 .list / .item 配方
      '.f2c-menu { position: fixed; pointer-events: auto; box-sizing: border-box;',
      '  z-index: ' + Z_MENU + '; min-width: 132px; padding: 4px;',
      '  display: flex; flex-direction: column;',
      '  border-radius: ' + tok('radiusLg') + ';',
      SURFACE,
      '}',
      '.f2c-item { display: flex; align-items: center; gap: 6px; width: 100%;',
      '  min-height: 34px; padding: 6px 8px; border: none;',
      '  border-radius: ' + tok('radiusMd') + '; background: transparent;',
      '  cursor: pointer; font-family: inherit; font-size: 13px; line-height: 20px;',
      '  color: ' + tok('label') + '; text-align: left; }',
      '.f2c-item:hover { background-image: ' + hoverFill() + '; }',
      // 「打开失败」这类一次性提示：浮在点击处附近，不挡鼠标、几秒后自己消失。
      // 用不透明底而不是菜单材质 —— 它可能压在任意内容上，模糊材质会糊成一片。
      '.f2c-note-pop { position: fixed; z-index: ' + Z_NOTE + '; max-width: 320px;',
      '  padding: 6px 10px; border: 0.5px solid ' + tok('border4') + ';',
      '  border-radius: ' + tok('radiusMd') + ';',
      '  background: ' + tok('bgOverlay') + '; color: ' + tok('label') + ';',
      '  font-family: inherit; font-size: 12px; line-height: 18px;',
      '  overflow-wrap: anywhere; pointer-events: none; }',

      // 选区上方的「批注」按钮 —— 形状照菜单项（Menu.module.css .item）
      '.f2c-bar { position: fixed; pointer-events: auto; display: inline-flex;',
      '  align-items: center; gap: 6px; min-height: 30px; padding: 5px 10px;',
      '  border-radius: ' + tok('radiusMd') + ';',
      SURFACE,
      '  color: ' + tok('label') + '; font-family: inherit;',
      '  font-size: 13px; line-height: 20px; cursor: pointer; white-space: nowrap; }',
      '.f2c-bar:hover { background-image: ' + hoverFill() + '; }',
      // 图标：与菜单项图标同规格 —— 14px 线性图标，取 DSH 的图标色。
      '.f2c-ic { flex: none; display: inline-flex; align-items: center; justify-content: center; }',
      '.f2c-ic svg { display: block; width: 14px; height: 14px; }',
      '.f2c-bar .f2c-ic, .f2c-sent .f2c-ic { color: ' + tok('menuIcon') + '; }',
      '.f2c-chip .f2c-ic { color: inherit; }',

      // 批注卡片 —— 表面见上面的 .f2c-card，这里是三段式内容
      '.f2c-card-head { display: flex; align-items: center; gap: 6px;',
      '  padding: 5px 8px 6px; color: ' + tok('label3') + ';',
      '  font-size: 11px; line-height: 15px; }',
      '.f2c-src { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }',
      '.f2c-x { flex: none; display: inline-flex; align-items: center; justify-content: center;',
      '  width: 20px; height: 20px; margin-right: -3px; padding: 0; border: 0;',
      '  border-radius: ' + tok('radiusSm') + '; background: transparent;',
      '  color: ' + tok('menuIcon') + '; cursor: pointer; font-size: 12px; line-height: 1; }',
      '.f2c-x:hover { background-image: ' + hoverFill() + '; color: ' + tok('label') + '; }',

      // 引用正文：套输入框观感（Input.module.css .wrap）
      '.f2c-quote { max-height: 108px; overflow: auto; margin: 0 4px 4px; padding: 6px 8px;',
      '  border: 0.5px solid ' + tok('border4') + ';',
      '  border-radius: ' + tok('radiusMd') + ';',
      '  background: ' + tok('layer1') + ';',
      '  color: ' + tok('label2') + '; white-space: pre-wrap;',
      '  word-break: break-word; font-size: 12px; line-height: 18px; }',
      '.f2c-note { display: block; box-sizing: border-box; width: calc(100% - 8px);',
      '  margin: 0 4px; min-height: 56px; max-height: 140px; resize: vertical;',
      '  padding: 6px 8px; border: 0.5px solid ' + tok('border4') + ';',
      '  border-radius: ' + tok('radiusMd') + ';',
      '  background: ' + tok('layer1') + ';',
      '  color: ' + tok('label') + '; font-family: inherit;',
      '  font-size: 13px; line-height: 20px; outline: none; }',
      '.f2c-note::placeholder { color: ' + tok('labelDim') + '; }',
      '.f2c-note:focus { border-color: ' + tok('brand') + '; }',

      // Excel 选区预览：小表格，横向可滚，同引用块一套框
      '.f2c-grid { max-height: 132px; overflow: auto; margin: 0 4px 4px;',
      '  border: 0.5px solid ' + tok('border4') + ';',
      '  border-radius: ' + tok('radiusMd') + ';',
      '  background: ' + tok('layer1') + '; }',
      '.f2c-grid table { border-collapse: collapse;',
      '  font-size: 12px; line-height: 18px; color: ' + tok('label2') + '; }',
      '.f2c-grid td, .f2c-grid th { padding: 4px 8px; text-align: left;',
      '  border-right: 0.5px solid ' + tok('border2') + ';',
      '  border-bottom: 0.5px solid ' + tok('border2') + ';',
      '  white-space: nowrap; max-width: 180px; overflow: hidden;',
      '  text-overflow: ellipsis; }',
      '.f2c-grid tr:last-child td { border-bottom: 0; }',
      '.f2c-grid td:last-child, .f2c-grid th:last-child { border-right: 0; }',
      '.f2c-grid th { position: sticky; top: 0; background: ' + tok('layer2') + ';',
      '  color: ' + tok('label') + '; font-weight: 500; }',
      '.f2c-empty { color: ' + tok('label3') + '; font-style: italic; }',
      '.f2c-note2 { margin: 0 8px 4px; color: ' + tok('label3') + ';',
      '  font-size: 12px; line-height: 18px; }',

      // 底栏：Menu.module.css .footer 的配方（3px 间距 + 0.5px l2 上边线）
      '.f2c-foot { display: flex; align-items: center; justify-content: space-between;',
      '  gap: 8px; margin-top: 3px; padding: 5px 8px 1px;',
      '  border-top: 0.5px solid ' + tok('border2') + '; }',
      '.f2c-hint { color: ' + tok('label3') + '; font-size: 11px; line-height: 16px; }',
      '.f2c-send { flex: none; height: 28px; padding: 0 10px; border: 0;',
      '  border-radius: ' + tok('radiusSm') + '; cursor: pointer;',
      '  background: ' + tok('btnFill') + '; color: ' + tok('fg') + ';',
      '  font-family: inherit; font-size: 12px; line-height: 18px; }',
      '.f2c-send:hover:not(:disabled) { background: ' + tok('btnHover') + '; }',
      '.f2c-send:disabled { cursor: not-allowed; opacity: .4; }',
      '.f2c-send.is-ghost { background: transparent; color: ' + tok('label2') + '; }',
      '.f2c-send.is-ghost:hover:not(:disabled) { background: ' + tok('hover') + '; }',

      // 已发送消息里被折叠起来的引用块 chip
      '.f2c-sent { display: inline-flex; align-items: center; max-width: 100%;',
      '  margin: 2px 0; padding: 1px 8px; cursor: pointer;',
      '  border-radius: ' + tok('radiusSm') + ';',
      '  background: ' + tok('layer2') + ';',
      '  border: 0.5px solid ' + tok('border2') + ';',
      '  color: ' + tok('label2') + '; font-family: inherit;',
      '  font-size: 12px; line-height: 18px; vertical-align: baseline;',
      '  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;',
      '  user-select: none; }',
      '.f2c-sent:hover { border-color: ' + tok('border4') + '; color: ' + tok('label') + '; }',
      '.f2c-sent:focus-visible { outline: 2px solid ' + tok('brand') + '; outline-offset: 1px; }',

      // 折叠后**必须留着**的批注。刻意做成 inline 的裸文字，不加方框 ——
      // 目标是和输入框里「胶囊后面直接跟批注」的观感一致；加框会让聊天
      // 看起来比输入框重得多，那正是用户嫌「乱」的原因。
      '.f2c-sent-note { color: ' + tok('label') + '; font-family: inherit;',
      '  font-size: inherit; line-height: inherit;',
      '  white-space: pre-wrap; overflow-wrap: anywhere; }',

      // 「封膜」：输入框里挂着胶囊期间（= 马上就要发出去），还没被我们处理过的
      // 消息行一律不可见。于是新消息从进入 DOM 到被折成胶囊之间的那个「全文」
      // 形态永远不会被绘制出来 —— 这层保险不依赖我们比绘制快多少，只依赖
      // 「早于绘制」这个微任务窗口，而 MutationObserver 天然就在那个窗口里。
      // 用 visibility 而非 display：保住布局，不引起回流跳动。
      'body[data-f2c-seal] [data-chat-flow-kind="user"]:not([data-f2c-seen]),',
      'body[data-f2c-seal] [data-chat-flow-kind="steering"]:not([data-f2c-seen]),',
      'body[data-f2c-seal] [data-pending-steering]:not([data-f2c-seen]),',
      // 乐观回显（见 FOLD_SEL 处的说明）同样要在折叠完成前藏起来。
      'body[data-f2c-seal] [data-submission-echo]:not([data-f2c-seen]) { visibility: hidden; }',

      // 工具栏上的常驻提示（静态摆放，不参与浮层定位）
      '.f2c-chip { display: inline-flex; align-items: center; gap: 4px; height: 26px; padding: 0 8px;',
      '  color: ' + tok('label3') + '; font-family: inherit; font-size: 12px;',
      '  line-height: 18px; white-space: nowrap; pointer-events: none; user-select: none; }',
    ].join('\n')

    function ensureStyle() {
      // 注意：**不能**「已存在就 return」。
      // 这个 <style> 挂在 document.head 上、不属于本模块的 fiber，client HMR
      // 热换插件时它会被留下来。若这里直接返回，新加的规则永远进不来 ——
      // 表现为「JS 是新的、CSS 还是上一世代的」，而且没有任何报错。
      // 每次覆写；内容相同则跳过，避免无谓的样式重算。
      var style = document.getElementById(STYLE_ID)
      if (style === null) {
        style = document.createElement('style')
        style.id = STYLE_ID
        document.head.appendChild(style)
      }
      if (style.textContent !== CSS) style.textContent = CSS
    }

    // ---- 主题变量搬运 ----
    // 浮层挂在 document.body 上（为了绕开侧边栏的 overflow 裁剪与 stacking
    // context），但主题变量未必定义在 body 这一层。所以每次显示浮层时，
    // 从一个**确定在主题作用域内**的元素上把 token 的**计算值**读出来，
    // 作为内联自定义属性写到浮层自己身上 —— var() 于是从浮层自身解析，
    // 不再依赖继承链，浅色/深色切换也都跟着对。
    function themeHost() {
      var probe = T.label
      var candidates = [
        document.querySelector('[data-document-preview]'),
        document.querySelector('[data-textpreview-body]'),
        document.getElementById('root'),
        document.body,
        document.documentElement,
      ]
      for (var i = 0; i < candidates.length; i++) {
        var el = candidates[i]
        if (el === null || el === undefined) continue
        try {
          if (window.getComputedStyle(el).getPropertyValue(probe).trim() !== '') return el
        } catch (_) { /* 下一个 */ }
      }
      return document.body
    }

    function syncTheme(target) {
      if (target === null || target === undefined) return
      try {
        var cs = window.getComputedStyle(themeHost())
        var keys = Object.keys(T)
        for (var i = 0; i < keys.length; i++) {
          var v = cs.getPropertyValue(T[keys[i]])
          if (v !== null && v.trim() !== '') target.style.setProperty(T[keys[i]], v.trim())
        }
      } catch (_) { /* 保留 CSS 里自洽的 fallback 对 */ }
    }

    // ================================================================
    // 五、注入桥：React 侧把 inputActions / 草稿读法交给命令式浮层
    // ================================================================
    // 工具栏组件在预览标签页开着时是常驻挂载的，因此桥上的值随渲染刷新。
    var bridge = {
      inputActions: null,
      sessionId: null,
      absolutePath: null,
      draft: '',
    }

    // ================================================================
    // 五·B、输入框里的引用 chip
    // ================================================================
    // 为什么需要这一层：直接把整段引用块 setDraft 进去，输入框里就会铺开
    // 一大坨 `> 📄 引用来源…`，看着很乱。DSH 官方对这种情况有专门的机制 ——
    // 「引用 source」：草稿里放一个 chip 节点，**提交时**由 codec.serialize
    // 展开成完整正文。于是输入框里只有一枚小胶囊，模型收到的仍是全文。
    //
    // 三件事配合：
    //   1. registerSource() 注册 source（提供 codec，chip↔文本的编解码）
    //   2. refBodies  存 ref → 正文（codec 回调时按 ref 取回全文）
    //   3. shell.insertReference(refInfo, span) 把 chip 插进草稿
    //
    // 注意：插入 chip 之后**绝不能再 setDraft** —— 那会用纯文本重写整份草稿，
    // 把刚插进去的 chip 抹掉。批注因此并入 codec 的正文里，而不是单独发。
    // 这份状态必须挂在 globalThis 上，不能留在模块作用域里。
    //
    // 原因：client 侧是 HMR 热更新的 —— 每次改写 client.js，浏览器都会丢掉旧
    // 模块、用新 factory 重新执行一遍，于是模块级变量全部重置；但提供
    // inputTriggers 的服务是**另一个插件的单例**，上一世代的注册还挂在它身上。
    // 结果就是新世代再注册同一个 name 会撞 `slash source "…" is already registered`，
    // 而旧注册的 codec 闭包读的又是**旧**的 refBodies，两边都对不上。
    //
    // 挂到 globalThis 后，新旧世代的 codec 读的是同一张表：撞名时直接沿用旧注册，
    // 序列化照样能取到全文。
    var G = (function () {
      var g = typeof globalThis !== 'undefined' ? globalThis : window
      if (g.__dshFileToChat === undefined || g.__dshFileToChat === null) {
        g.__dshFileToChat = { refBodies: new Map(), refSeq: 0, ready: false }
      }
      return g.__dshFileToChat
    })()

    // 当前**真正生效**的 source 名。放在 G 上而不是模块作用域：换名绕开残留注册
    // 之后，下一个世代得从新名字接着用，不能又退回撞名的那个。
    var QUOTE_SOURCE = G.source || 'file-to-chat-ref'

    /** 注册引用 source（幂等，且跨 HMR 世代安全）。svc = inputTriggers 服务实例。 */
    function registerQuoteSource(svc) {
      if (G.ready === true && G.source === QUOTE_SOURCE) return true
      if (svc === null || svc === undefined || typeof svc.registerSource !== 'function') return false
      var spec = {
        trigger: '@',
        name: QUOTE_SOURCE,
        order: 102,
        candidates: function () { return Promise.resolve([]) },
        onPick: function () { return undefined },
        codec: {
          // 一律经 G 读表：这样即使这段闭包属于上一个代码世代，也仍然取得到
          // 新世代写进去的正文。
          clipboardText: function (ref) { return G.refBodies.get(ref) || '' },
          serialize: function (ref) { return Promise.resolve(G.refBodies.get(ref) || '') },
        },
        // 点输入框里的 chip：不接管跳转，返回 false 让它走默认行为。
        openReference: function () { return false },
      }
      try {
        G.dispose = svc.registerSource(spec)
        G.ready = true
        G.source = QUOTE_SOURCE
        dbg('引用 source 已注册：输入框将显示为 chip，提交时展开全文')
        return true
      } catch (e) {
        if (!/already registered/i.test(String(e && e.message))) {
          console.warn('[file-to-chat] 注册引用 source 失败（输入框将退化为纯文本）', e)
          return false
        }
      }
      // 撞名：上一世代的注册还挂在服务上。若手里有它的 disposer，就撤掉换成
      // 本世代的 codec —— 只有这样才能保证 chip 提交时**真的**取得到正文。
      if (typeof G.dispose === 'function') {
        try { G.dispose() } catch (_) { /* 已失效就算了 */ }
        G.dispose = null
        try {
          G.dispose = svc.registerSource(spec)
          G.ready = true
          G.source = QUOTE_SOURCE
          dbg('引用 source 已换成本世代注册')
          return true
        } catch (e2) {
          console.warn('[file-to-chat] 重新注册失败，沿用旧注册', e2)
        }
      }
      // 没有 disposer（globalThis 改造之前的老世代留下的注册）—— 换一个名字注册。
      // 那种旧 codec 读的是它自己世代的旧表，序列化会得到空串；换名之后本世代的
      // codec 才真正生效，用户不必硬刷新。
      var alt = QUOTE_SOURCE + '-g' + (G.gen = (G.gen || 1) + 1)
      try {
        G.dispose = svc.registerSource(Object.assign({}, spec, { name: alt }))
        QUOTE_SOURCE = alt
        G.ready = true
        G.source = alt
        dbg('引用 source 换名注册为 ' + alt + '（绕开旧世代的残留注册）')
        return true
      } catch (e2) {
        // 退无可退：至少让旧注册继续挂着，用户硬刷新后自然恢复。
        G.ready = true
        console.warn('[file-to-chat] 换名注册也失败，只能沿用旧注册；'
          + '若发出的消息为空，请 Ctrl+Shift+R 硬刷新一次', e2)
        return true
      }
    }

    /**
     * 拿到 inputTriggers 并注册。
     *
     * 激活顺序不保证：本插件随 slots 就绪就激活，而提供 inputTriggers 的插件
     * 此刻常常还没 ACTIVE，ctx.get 只能拿到 undefined —— 那样输入框里就永远
     * 不会出现 chip（这正是 dsh-file-quote 踩过并写在注释里的坑）。所以先即时
     * 试一次，拿不到就交给 ctx.inject 等该服务就绪后再注册（cordis 的迟到依赖
     * 用法），再加一层定时重试兜底。
     */
    /**
     * 服务上当前生效的注册，是不是**本世代这个** source？
     *
     * 不能只测 `G.ready`：残留注册（别的世代留下的）也会让它变成 true，而那种
     * codec 读的是旧表，序列化会得到空串 —— 表现就是「发出去之后内容没了」。
     * 必须连名字一起对上，才算真的可用。
     */
    function quoteSourceLive() {
      return G.ready === true && G.source === QUOTE_SOURCE
    }

    function wireQuoteSource(ctx) {
      if (quoteSourceLive()) return
      var svc = null
      try { svc = ctx.get('inputTriggers') } catch (_) { svc = null }
      if (registerQuoteSource(svc)) return
      try {
        ctx.inject(['inputTriggers'], function (child) {
          if (quoteSourceLive()) return
          var s = null
          try { s = child.get('inputTriggers') } catch (_) { s = null }
          registerQuoteSource(s)
        })
      } catch (e) {
        console.warn('[file-to-chat] ctx.inject(inputTriggers) 不可用', e)
      }
      var tries = 0
      var timer = setInterval(function () {
        if (quoteSourceLive()) { clearInterval(timer); return }
        if (++tries > 20) {
          clearInterval(timer)
          console.warn('[file-to-chat] inputTriggers 始终未就绪：输入框将只能显示纯文本引用块')
          return
        }
        var s2 = null
        try { s2 = ctx.get('inputTriggers') } catch (_) { s2 = null }
        registerQuoteSource(s2)
      }, 1000)
    }

    /**
     * 把引用块送进 composer。
     * 优先级：官方 chip（输入框里只有小胶囊） → 座位 setDraft → shell.setDraft。
     * @param label - chip 上显示的短标签；缺省时自动编号。
     * @param note  - 用户写的批注。批注**内容**不塞进胶囊的序列化正文，而是作为
     *                普通草稿文本补在胶囊右侧 —— 否则输入框里只看得到一枚胶囊，
     *                批注等于被藏起来了（用户实测反馈）。序列化正文只携带
     *                「空行 + 批注：」这个分隔前缀，所以输入框里是「[胶囊] 内容」
     *                同一行、无「批注」二字，而发出去的正文格式完全不变。
     */
    function injectText(ctx, readSessionId, text, label, note) {
      var sid = bridge.sessionId
      if (sid === undefined || sid === null || sid === '') sid = readSessionId()
      var noteBody = note === null || note === undefined ? '' : String(note).trim()
      // 纯文本路径（②③）没法把批注和引用拆开显示，直接把两者拼成完整正文。
      var full = noteBody === '' ? text : text + '\n\n批注：' + noteBody

      // ① 首选：插成 composer 里的引用 chip
      if (sid !== undefined && sid !== null && sid !== '') {
        try {
          wireQuoteSource(ctx)
          if (quoteSourceLive()) {
            var scoped = ctx.sessions.scope(sid)
            var shell = scoped !== undefined && scoped !== null
              ? ctx.conversation.input.for(scoped) : null
            if (shell !== null && shell !== undefined && typeof shell.insertReference === 'function') {
              var st = shell.state.getSnapshot()
              var at = typeof st.draft === 'string' ? st.draft.length : 0
              var cs = null
              try { cs = typeof shell.caretSpan === 'function' ? shell.caretSpan() : null } catch (_) { cs = null }
              // `insertReference` 的语义是「用胶囊**替换**这个 span」。光标若落在
              // 一个选区上，被选中的草稿内容会被替换掉甚至删掉 —— 那不是用户要的。
              // 一律收敛成零宽点（取选区末端），宁可插偏也不删东西。
              var csStart = cs !== null && cs !== undefined && typeof cs.start === 'number' ? cs.start : at
              var csEnd = cs !== null && cs !== undefined && typeof cs.end === 'number' ? cs.end : at
              if (csStart !== csEnd) csStart = csEnd
              var span = {
                start: csStart,
                end: csEnd,
                draftRev: st.draftRev,
              }
              // 草稿里已经有内容时先断一行再插胶囊 —— 否则第二条引用会紧贴在
              // 上一行末尾（用户实测反馈）。这里刻意**不用 setDraft**：它会剥掉
              // U+FFFC 占位符后重建整份草稿，把草稿里已有的胶囊一并抹掉。
              // 输入域 CSS 是 white-space:pre-wrap，所以裸 \n 就能渲染成换行。
              var draftNow = typeof st.draft === 'string' ? st.draft : ''
              var prevChar = span.start > 0 ? draftNow.charAt(span.start - 1) : ''
              // 只在**零宽光标**上断行。光标若是一个选区（start !== end），
              // insertText 会把选中的内容整个替换成换行 —— 那等于吃掉用户已经
              // 插好的胶囊。宁可让两条引用挤在一行，也不能删东西。
              if (span.start === span.end && span.start > 0 && draftNow !== '' &&
                  prevChar !== '\n' && bridge.inputActions) {
                var broke = false
                try {
                  broke = bridge.inputActions.insertText('\n', {
                    start: span.start, end: span.end, draftRev: st.draftRev,
                  }) !== false
                } catch (_) { broke = false }
                if (broke) {
                  try {
                    var sp = bridge.inputActions.captureInsertion()
                    if (sp !== null && sp !== undefined) {
                      span = { start: sp.start, end: sp.end, draftRev: sp.draftRev }
                    }
                  } catch (_) { /* 拿不到新光标就沿用旧 span，胶囊会退化成同行追加 */ }
                }
              }
              var n = ++G.refSeq
              var ref = 'q' + n
              var text2 = label !== undefined && label !== null && String(label) !== ''
                ? String(label) : ('引用 #' + n)
              // 胶囊的序列化正文 = 引用块 + （有批注时）「空行 + 批注：」这个
              // 分隔前缀。批注**内容本身**不进正文，而是作为普通草稿文本补在
              // 胶囊右侧 —— 这样输入框里就是 `[胶囊] 内容`，同一行、无「批注」
              // 二字；而拼起来发出去的正文仍是「引用块 + 空行 + 批注：内容」。
              var chipBody = noteBody === '' ? text : text + '\n\n批注：'
              G.refBodies.set(ref, chipBody)
              if (G.refBodies.size > 200) G.refBodies.delete(G.refBodies.keys().next().value)
              var refInfo = { source: QUOTE_SOURCE, ref: ref, label: text2, clipboardText: chipBody }
              var applied = false
              try { applied = shell.insertReference(refInfo, span) === true } catch (_) { applied = false }
              // 事件通道兜底：shell 方法因 phase/rev 拒绝时再试一次官方通道。
              if (!applied && scoped !== null && scoped !== undefined && typeof scoped.bail === 'function') {
                try {
                  applied = scoped.bail(scoped, 'slash/input-insert-reference', { reference: refInfo, span }) === true
                } catch (_) { applied = false }
              }
              if (applied) {
                // 胶囊已落地。此刻光标被 Lexical 放在最后一个插入节点之后
                // （$replaceDetectSpanWithNodes 的明文契约），正好用来补批注。
                var noteInline = false
                if (noteBody !== '') noteInline = appendNoteAfterChip(noteBody)
                // 补不进去就把批注并回序列化正文 —— 输入框里看不到，但发出去的
                // 内容仍然完整。注意正文此时只有「批注：」这个前缀，得把内容
                // 一并写回去，否则会留下一个悬空的前缀。宁可少看，不可少发。
                if (!noteInline && noteBody !== '') {
                  G.refBodies.set(ref, text + '\n\n批注：' + noteBody)
                }
                // 封膜：胶囊真正落地的这一刻，几乎可以断定用户接下来就是按发送。
                // 直接在这里开膜，不依赖任何关于占位符码位的假设 —— 这是最可靠
                // 的触发点。
                sealOn()
                return {
                  via: 'insertReference(chip)', sessionId: sid, ref: ref,
                  label: text2, noteInline: noteInline,
                }
              }
              // 被拒时把守卫现场打出来：insertReference 只在 phase 为
              // plain/claimed 且 span.draftRev === shell.rev 时才动手。
              var why = '?'
              try { why = shell.core.state.phase } catch (_) { why = '?' }
              var srev = '?'
              try { srev = shell.rev } catch (_) { srev = '?' }
              console.warn('[file-to-chat] chip 插入被拒，回退纯文本', {
                phase: why, spanDraftRev: span.draftRev, shellRev: srev,
              })
              G.refBodies.delete(ref)
              G.refSeq--
            }
          }
        } catch (e) {
          console.warn('[file-to-chat] chip 插入失败，回退纯文本', e)
        }
      }

      // ② 座位标准 prop（已实测可用）
      //
      // 这里有个**必须绕开的坑**：`setDraft` 会用纯文本重写整份草稿 —— 它会先
      // 把占位符（`/[\uE100-\uE11D\uFFFC]/`）剥掉，再重建每一个段落。草稿里
      // 已经插好的引用胶囊本身就靠占位符表示，于是**一并被抹掉**。
      // 这正是「同时引用多个会吃掉后边的和备注」的成因：第二条引用在 chip 插入
      // 被拒时落到这条兜底，把第一条连同它的批注一起清了。
      // 所以：草稿里只要有胶囊，就绝不走 setDraft，改为在光标处追加纯文本。
      var ia = bridge.inputActions
      if (ia !== null && ia !== undefined) {
        var before = typeof bridge.draft === 'string' ? bridge.draft : ''
        if (CHIP_MARK_RE.test(before) &&
            typeof ia.captureInsertion === 'function' && typeof ia.insertText === 'function') {
          var sp2 = null
          try { sp2 = ia.captureInsertion() } catch (_) { sp2 = null }
          if (sp2 !== null && sp2 !== undefined) {
            var appended = false
            try { appended = ia.insertText('\n' + full, sp2) !== false } catch (_) { appended = false }
            if (appended) {
              return { via: 'inputActions.insertText(append)', sessionId: bridge.sessionId }
            }
          }
        }
        if (typeof ia.setDraft === 'function') {
          ia.setDraft(before !== '' ? before + '\n' + full : full)
          return { via: 'inputActions.setDraft', sessionId: bridge.sessionId }
        }
      }

      // ③ 兜底：内核 conversation 服务
      if (sid === undefined || sid === null || sid === '') throw new Error('无法确定当前会话 id')
      var scoped2 = ctx.sessions.scope(sid)
      if (scoped2 === undefined || scoped2 === null) throw new Error('会话 scope 不可用')
      var shell2 = ctx.conversation.input.for(scoped2)
      if (shell2 === undefined || shell2 === null) throw new Error('conversation.input.for() 未返回 shell')
      var b2 = ''
      try { b2 = shell2.state.getSnapshot().draft || '' } catch (_) { b2 = '' }
      shell2.setDraft(b2 !== '' ? b2 + '\n' + full : full)
      return { via: 'conversation.input.setDraft', sessionId: sid }
    }

    /**
     * 胶囊插入后，把批注当**普通草稿文本**补在它右侧，同一行、不带任何前缀。
     *
     * 依据 `$replaceDetectSpanWithNodes` 的明文契约：「The caret lands after the
     * last inserted node.」所以此刻 `captureInsertion()` 拿到的就是胶囊右侧的
     * 合法插入点。必须用 `insertText`，**不能用 setDraft** —— setDraft 会用纯文本
     * 重写整份草稿，把刚插进去的胶囊抹掉（dsh-file-quote 踩过并写在注释里）。
     *
     * 这里只插**批注内容本身**：`\n\n批注：` 这个分隔前缀由胶囊的序列化正文
     * 携带（见 injectText 的 chipBody）。于是输入框里看到的是
     * `[胶囊] 批注内容`，而发出去的正文仍然是「引用块 + 空行 + 批注：内容」。
     * @returns 是否真的插进去了；false 时调用方会把批注并回引用正文。
     */
    function appendNoteAfterChip(noteBody) {
      var ia = bridge.inputActions
      if (ia === null || ia === undefined) return false
      if (typeof ia.captureInsertion !== 'function' || typeof ia.insertText !== 'function') return false
      try {
        var span = ia.captureInsertion()
        if (span === null || span === undefined) return false
        // 只插内容：不换行、不要「批注」二字。胶囊插入时自带一个尾随空格，
        // 正好把胶囊和批注隔开。
        return ia.insertText(noteBody, span) !== false
      } catch (e) {
        console.warn('[file-to-chat] 批注未能作为草稿文本补入，改为并进引用正文', e)
        return false
      }
    }

    /** 输入框胶囊上的短标签：够认出是哪一个引用，又不撑爆一行。 */
    // 整份文件的内联上限。比选区的 MAX_EXCERPT_LINES 宽得多 —— 选区是「看这段」，
    // 整份文件是「把这份东西交给模型」，太窄就没意义；但也不能不设，否则一次
    // 发送就能把上下文顶满。
    var MAX_FILE_LINES = 200
    /** 一次向宿主取多少行；取满就够拼引用块，`eof === false` 时再补一句提示。 */
    var FILE_READ_LINES = MAX_FILE_LINES
    // Office 抽包与字节回退共用的读取上限：readBytes 硬顶 32 MiB，这里再收一收。
    var OFFICE_MAX_BYTES = 24 * 1024 * 1024
    var FILE_MAX_BYTES = OFFICE_MAX_BYTES
    // 抽 xlsx 时最多内联多少行 —— 一整张表全塞进消息会吃掉几万 token。
    var MAX_OFFICE_ROWS = 200
    // 目录清单最多列几项。
    var MAX_DIR_ENTRIES = 300

    function chipLabel(snap) {
      var short = ''
      if (snap.kind === 'excel') {
        var sheet = snap.sheet !== null && snap.sheet !== undefined && snap.sheet !== '' ? String(snap.sheet) + ' ' : ''
        short = sheet + String(snap.address !== undefined && snap.address !== null ? snap.address : '')
      } else if (snap.kind === 'file') {
        // 整份文件只给文件名。没有行号、没有页码 —— 「没有位置」本身就是签名，
        // 和选区、表格的胶囊一眼能分开。
        short = snap.path ? String(snap.path).split(/[\\/]/).pop() : '文件'
      } else {
        var base = snap.path ? String(snap.path).split(/[\\/]/).pop() : '选区'
        if (Array.isArray(snap.lineRange)) {
          short = base + ':' + snap.lineRange.join('-')
        } else if (Array.isArray(snap.pageRange)) {
          short = base + ' · 第' + (snap.pageRange[0] === snap.pageRange[1]
            ? snap.pageRange[0]
            : snap.pageRange[0] + '-' + snap.pageRange[1]) + '页'
        } else {
          short = base
        }
      }
      short = truncate(short, 30)
      return short === '' ? '引用' : '引用 · ' + short
    }

    // ---------------- 整份文件的读取（文件树右键用） ----------------
    /**
     * 把宿主抛回来的错误翻成人话。宿主侧约定是「只分支 on `code`，绝不看 message」，
     * 所以这里先取 `err.code` 再退到 message。
     */
    function fileFailureLabel(err) {
      var code = err !== null && err !== undefined && typeof err.code === 'string' ? err.code : ''
      if (code === 'workspace-file/not-text') return '该格式不是纯文本'
      if (code === 'workspace-file/too-large') return '文件超过单页上限'
      if (code === 'workspace-file/not-found') return '文件不存在'
      if (code === 'workspace-file/not-regular-file') return '不是一个普通文件'
      if (code === 'workspace-file/outside-workspace') return '不在工作区内'
      if (code !== '') return code
      var msg = err !== null && err !== undefined && typeof err.message === 'string' ? err.message : ''
      return msg === '' ? '读取失败' : msg
    }

    // ================================================================
    // 五·B、Office 正文抽取（docx / xlsx）
    // ================================================================
    // 为什么不走宿主 half：宿主能读字节，但抽取逻辑要写 Node 侧，而宿主 half
    // **不热重载** —— 每改一次都得退出应用重启。而 docx/xlsx 本质就是 ZIP，
    // 浏览器自带 `DecompressionStream('deflate-raw')` 就能解，全在客户端做，
    // 改完 HMR 立刻生效。
    //
    // 范围只到 docx / xlsx：pdf 的文本抽取要处理字体编码和 ToUnicode CMap，
    // 手写必错，而那一步 DSH 已经用 pdf.js 做好了 —— 预览时文本层就在 DOM 里。
    // 所以在文件区右键一个 pdf，仍然只发来源行。

    function decodeUtf8(bytes) {
      try { return new TextDecoder('utf-8').decode(bytes) } catch (err) { return '' }
    }

    /** 二进制探测：UTF-8 文本里出现 NUL 就基本可以断定不是给人读的。 */
    function indexOfNul(bytes) {
      for (var i = 0; i < bytes.length; i++) if (bytes[i] === 0) return i
      return -1
    }

    /** XML 实体还原。`&amp;` 必须最后做，否则 `&amp;lt;` 会被二次还原。 */
    function decodeEntities(s) {
      return String(s)
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&#x([0-9a-fA-F]+);/g, function (_, h) { return String.fromCodePoint(parseInt(h, 16)) })
        .replace(/&#(\d+);/g, function (_, d) { return String.fromCodePoint(parseInt(d, 10)) })
        .replace(/&amp;/g, '&')
    }

    function u16At(view, off) { return view.getUint16(off, true) }
    function u32At(view, off) { return view.getUint32(off, true) }

    /**
     * 最小 ZIP 目录读取：只读中央目录，拿到每个条目的名字/压缩方式/位置。
     * 从尾部往前找 EOCD 签名（注释最长 65535，所以最多回退 66000 字节）。
     */
    function zipEntries(bytes) {
      if (bytes.length < 22) return null
      var view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
      var eocd = -1
      var floor = Math.max(0, bytes.length - 66000)
      for (var i = bytes.length - 22; i >= floor; i--) {
        if (u32At(view, i) === 0x06054b50) { eocd = i; break }
      }
      if (eocd < 0) return null
      var count = u16At(view, eocd + 10)
      var p = u32At(view, eocd + 16)
      var out = []
      for (var n = 0; n < count && p + 46 <= bytes.length; n++) {
        if (u32At(view, p) !== 0x02014b50) break
        var nameLen = u16At(view, p + 28)
        var name = ''
        for (var k = 0; k < nameLen; k++) name += String.fromCharCode(bytes[p + 46 + k])
        out.push({
          name: name,
          method: u16At(view, p + 10),
          compSize: u32At(view, p + 20),
          localOff: u32At(view, p + 42)
        })
        p += 46 + nameLen + u16At(view, p + 30) + u16At(view, p + 32)
      }
      return out
    }

    /** 解压一个条目；method 0 直取，8 走 DecompressionStream，其余放弃。 */
    function zipRead(bytes, entry) {
      var view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
      if (u32At(view, entry.localOff) !== 0x04034b50) return Promise.resolve(null)
      var dataOff = entry.localOff + 30 + u16At(view, entry.localOff + 26) + u16At(view, entry.localOff + 28)
      var raw = bytes.subarray(dataOff, dataOff + entry.compSize)
      if (entry.method === 0) return Promise.resolve(raw)
      if (entry.method !== 8 || typeof DecompressionStream !== 'function') return Promise.resolve(null)
      try {
        var stream = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'))
        return new Response(stream).arrayBuffer().then(function (buf) {
          return new Uint8Array(buf)
        }, function () { return null })
      } catch (err) { return Promise.resolve(null) }
    }

    /** `word/document.xml` → 纯文本。段落收行、`<w:tab/>` 收制表符。 */
    function docxText(xml) {
      var s = String(xml)
        .replace(/<w:tab\b[^>]*\/?>/g, '\t')
        .replace(/<w:br\b[^>]*\/?>/g, '\n')
        .replace(/<\/w:p>/g, '\n')
        .replace(/<[^>]+>/g, '')
      return decodeEntities(s)
    }

    /** `xl/sharedStrings.xml` → 字符串表（`<si>` 内的多个 `<t>` 要拼起来）。 */
    function xlsxShared(xml) {
      var out = []
      var re = /<si\b[^>]*>([\s\S]*?)<\/si>/g
      var m
      while ((m = re.exec(xml)) !== null) {
        var inner = m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, '')
        var ts = inner.match(/<t\b[^>]*>[\s\S]*?<\/t>/g) || []
        var text = ''
        for (var i = 0; i < ts.length; i++) text += decodeEntities(ts[i].replace(/<[^>]+>/g, ''))
        out.push(text)
      }
      return out
    }

    function xlsxCellText(attrs, body, shared) {
      var t = /t="([^"]+)"/.exec(attrs)
      var type = t !== null ? t[1] : 'n'
      var v
      if (type === 's') {
        v = /<v>([\s\S]*?)<\/v>/.exec(body)
        var idx = v !== null ? parseInt(v[1], 10) : -1
        return idx >= 0 && idx < shared.length ? shared[idx] : ''
      }
      if (type === 'inlineStr') {
        var is = /<is\b[^>]*>([\s\S]*?)<\/is>/.exec(body)
        if (is === null) return ''
        var ts = is[1].match(/<t\b[^>]*>[\s\S]*?<\/t>/g) || []
        var out = ''
        for (var i = 0; i < ts.length; i++) out += decodeEntities(ts[i].replace(/<[^>]+>/g, ''))
        return out
      }
      // 数字、布尔、日期序列号、以及公式的缓存结果 —— 都在 <v> 里
      v = /<v>([\s\S]*?)<\/v>/.exec(body)
      return v !== null ? decodeEntities(v[1]) : ''
    }

    /** 一张工作表 → 二维文本数组。空尾列丢掉，整行空的也丢掉。 */
    function xlsxSheetRows(xml, shared) {
      var rows = []
      var rowRe = /<row\b[^>]*>([\s\S]*?)<\/row>/g
      var rm
      while ((rm = rowRe.exec(xml)) !== null) {
        var cells = []
        var cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g
        var cm
        while ((cm = cellRe.exec(rm[1])) !== null) {
          var attrs = cm[1]
          var body = cm[2] !== undefined ? cm[2] : ''
          var refM = /r="([A-Z]+)\d+"/.exec(attrs)
          var col = refM !== null ? colIndexOf(refM[1]) : cells.length
          if (col < 0) col = cells.length
          while (cells.length < col) cells.push('')
          cells[col] = xlsxCellText(attrs, body, shared)
        }
        while (cells.length > 0 && cells[cells.length - 1] === '') cells.pop()
        if (cells.length > 0) rows.push(cells)
      }
      return rows
    }

    function xlsxSheets(entries, bytes) {
      // 工作表顺序与显示名来自 workbook.xml + 它的 rels；取不到就退化成目录里的顺序。
      var names = []
      var rels = {}
      var wb = null
      var wbRels = null
      var i
      for (i = 0; i < entries.length; i++) {
        if (entries[i].name === 'xl/workbook.xml') wb = entries[i]
        if (entries[i].name === 'xl/_rels/workbook.xml.rels') wbRels = entries[i]
      }
      var sheets = []
      for (i = 0; i < entries.length; i++) {
        if (/^xl\/worksheets\/sheet\d+\.xml$/.test(entries[i].name)) sheets.push(entries[i])
      }
      if (wb !== null) {
        return zipRead(bytes, wb).then(function (b) {
          var xml = b === null ? '' : decodeUtf8(b)
          var re = /<sheet\b[^>]*>/g
          var m
          while ((m = re.exec(xml)) !== null) {
            var nm = /name="([^"]*)"/.exec(m[0])
            var rid = /r:id="([^"]*)"/.exec(m[0])
            names.push({
              name: nm !== null ? decodeEntities(nm[1]) : '',
              rid: rid !== null ? rid[1] : ''
            })
          }
          var relP = wbRels === null ? Promise.resolve(null) : zipRead(bytes, wbRels)
          return relP.then(function (rb) {
            if (rb !== null) {
              var rxml = decodeUtf8(rb)
              var rre = /<Relationship\b[^>]*>/g
              var rmm
              while ((rmm = rre.exec(rxml)) !== null) {
                var id = /Id="([^"]*)"/.exec(rmm[0])
                var tg = /Target="([^"]*)"/.exec(rmm[0])
                if (id !== null && tg !== null) {
                  rels[id[1]] = tg[1].replace(/^\/?xl\//, '').replace(/^\//, '')
                }
              }
            }
            var ordered = []
            for (var j = 0; j < names.length; j++) {
              var target = rels[names[j].rid]
              if (target === undefined) continue
              for (var k = 0; k < sheets.length; k++) {
                if (sheets[k].name === 'xl/' + target) {
                  ordered.push({ entry: sheets[k], title: names[j].name })
                  sheets[k].__used = true
                }
              }
            }
            for (var q = 0; q < sheets.length; q++) {
              if (sheets[q].__used !== true) ordered.push({ entry: sheets[q], title: '' })
            }
            return ordered
          })
        })
      }
      return Promise.resolve(sheets.map(function (e) { return { entry: e, title: '' } }))
    }

    /**
     * 解包 docx / xlsx 并抽出文本。失败一律 resolve(null) —— 抽取是「加分项」，
     * 不值得为它把整条发送流程搞崩。
     */
    function extractOfficeText(bytes, kind, path) {
      var entries = zipEntries(bytes)
      if (entries === null) return Promise.resolve(null)

      if (kind === 'docx') {
        var doc = null
        for (var i = 0; i < entries.length; i++) {
          if (entries[i].name === 'word/document.xml') { doc = entries[i]; break }
        }
        if (doc === null) return Promise.resolve(null)
        return zipRead(bytes, doc).then(function (b) {
          return b === null ? null : docxText(decodeUtf8(b))
        })
      }

      var sharedEntry = null
      for (var j = 0; j < entries.length; j++) {
        if (entries[j].name === 'xl/sharedStrings.xml') { sharedEntry = entries[j]; break }
      }
      var sharedP = sharedEntry === null ? Promise.resolve('') : zipRead(bytes, sharedEntry)
      return Promise.all([xlsxSheets(entries, bytes), sharedP]).then(function (both) {
        var sheets = both[0]
        var shared = both[1] === null || both[1] === undefined ? '' : decodeUtf8(both[1])
        var list = xlsxShared(shared)
        if (sheets.length === 0) return null
        // 只抽第一张工作表 —— 一份工作簿全量内联会吃掉几万 token，不值得。
        var first = sheets[0]
        return zipRead(bytes, first.entry).then(function (b) {
          if (b === null) return null
          var rows = xlsxSheetRows(decodeUtf8(b), list)
          if (rows.length === 0) return null
          var out = ['（工作表：' + (first.title === '' ? '第一张' : first.title) + '）']
          var cap = Math.min(rows.length, MAX_OFFICE_ROWS)
          for (var r = 0; r < cap; r++) {
            var line = []
            for (var c = 0; c < rows[r].length; c++) line.push(String(rows[r][c]).replace(/\s+/g, ' '))
            out.push(line.join('\t'))
          }
          if (rows.length > cap) out.push('…（本表还有 ' + (rows.length - cap) + ' 行未内联）')
          if (sheets.length > 1) {
            out.push('…（本工作簿还有 ' + (sheets.length - 1) + ' 张工作表未内联）')
          }
          return out.join('\n')
        })
      })
    }

    /**
     * Remote 调用返回的是**信封**：`{ ok: true, value: <真正的结果> }`，失败时是
     * `{ ok: false, error: … }`。
     *
     * 实测（`read` / `readBytes` / `list` 全都一样）—— 直接取 `.text` / `.data`
     * 恒为 undefined，于是插件一路判定「读不到正文」。这就是真因：
     * 数据一直在 `value` 里，只是被包了一层。
     */
    function unwrapRemote(res) {
      var guard = 0
      while (guard < 4
        && res !== null && res !== undefined && typeof res === 'object'
        && !(res instanceof Uint8Array) && !Array.isArray(res)
        && typeof res.ok === 'boolean' && ('value' in res || res.ok === false)) {
        if (res.ok === false) {
          var e = res.error !== undefined ? res.error : res
          return { error: fileFailureLabel(e) + '（' + describeValue(e) + '）' }
        }
        res = res.value
        guard += 1
      }
      return { value: res }
    }

    /** Remote 传回来的字节可能是 Uint8Array / 数字数组 / base64 字符串。 */
    function normalizeBytes(res) {
      if (res === null || res === undefined) return null
      var d = res.data !== undefined ? res.data : res
      if (d instanceof Uint8Array) return d
      if (Array.isArray(d)) return new Uint8Array(d)
      if (d !== null && typeof d === 'object' && Array.isArray(d.data)) return new Uint8Array(d.data)
      if (typeof d === 'string') {
        try {
          var bin = atob(d)
          var out = new Uint8Array(bin.length)
          for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
          return out
        } catch (err) { return null }
      }
      return null
    }

    /** 把任意返回值压成一句可读的形状描述 —— 排查「拿到了但不是预期形状」用。 */
    function describeValue(v) {
      if (v === null) return 'null'
      if (v === undefined) return 'undefined'
      var t = typeof v
      if (t === 'string') return '字符串(' + v.length + ' 字)'
      if (t === 'number' || t === 'boolean' || t === 'function') return t
      if (t !== 'object') return t
      try {
        if (v instanceof Uint8Array) return 'Uint8Array(' + v.length + ' 字节)'
        if (Array.isArray(v)) return '数组(' + v.length + ' 项)'
        var keys = Object.keys(v)
        return '对象{' + (keys.length === 0 ? '空' : keys.slice(0, 10).join(', ')) + (keys.length > 10 ? ', …' : '') + '}'
      } catch (err) { return '对象(读不出键)' }
    }

    function officeKindOf(path) {
      var s = String(path).toLowerCase()
      if (/\.docx$/.test(s)) return 'docx'
      if (/\.xlsx$/.test(s)) return 'xlsx'
      return null
    }

    /**
     * 读一份工作区文件，回调 `done(text, failure, meta)`，前两者恒有其一非 null。
     *
     * 走宿主 `remote.workspaceFiles.read`：`range.offset` 是 **1 基行号**，
     * `limit` 上限 `maxLines = 5000`；返回 `{ offset, text, lines, eof, version, bytes }`，
     * `text` 用 `\n` 连接、末尾不带终止符。
     *
     * 读不动不是异常路径 —— docx/xlsx/pdf 这类会被 `workspace-file/not-text` 拒掉，
     * 那就 `done(null, 原因)`，由 composeTextBlock 只发来源行，让模型自己去读。
     */
    function readFileForChat(ctx, sessionId, path, done) {
      // 两条路（文本 / 字节抽包）可能都会回落，done 只允许生效一次。
      var settled = false
      function finish(text, failure, meta) {
        if (settled) return
        settled = true
        done(text, failure, meta)
      }

      // 拿 workspaceFiles 命名空间。两条路都试：`ctx.remote.*` 与 `ctx.get('remote').*`
      // —— 只要有一条能拿到就行，另外 read / readBytes 有哪个用哪个。
      var api = null
      var candidates = []
      function consider(v) {
        if (v === null || v === undefined) return
        if (typeof v.read === 'function' || typeof v.readBytes === 'function') candidates.push(v)
      }
      try {
        if (ctx !== null && ctx !== undefined && ctx.remote !== undefined && ctx.remote !== null) {
          consider(ctx.remote.workspaceFiles)
        }
      } catch (err) { void err }
      try {
        var viaGet = typeof ctx.get === 'function' ? ctx.get('remote') : null
        if (viaGet !== null && viaGet !== undefined) consider(viaGet.workspaceFiles)
      } catch (err) { void err }
      for (var ci = 0; ci < candidates.length; ci++) {
        if (typeof candidates[ci].read === 'function') { api = candidates[ci]; break }
        if (api === null) api = candidates[ci]
      }
      if (api === null) {
        finish(null, '当前连接没有 workspaceFiles 接口', null)
        return
      }
      if (typeof sessionId !== 'string' || sessionId === '') {
        finish(null, '拿不到会话 id', null)
        return
      }

      /**
       * 第二条路：按字节读回来自己解码/解包。
       *
       * 这条**对任何文件都走**，不只是 Office —— 文本路的 `read` 在实机上
       * 出现过「成功但正文为空」，而 `readBytes` 是另一个方法，不受影响。
       * 与其猜原因，不如把字节路做成常规回退：拿到字节之后
       *   · docx / xlsx → 解包抽正文
       *   · 其余       → UTF-8 解码 + NUL 二进制探测 + 行数截断
       */
      function viaBytes(why) {
        var kind = officeKindOf(path)
        if (typeof api.readBytes !== 'function') {
          finish(null, why + '；这个接口也没有 readBytes（' + typeof api.readBytes + '）', null)
          return
        }
        var bp = null
        try {
          // **同样一个参数都不传**：`readBytes` 的 `range.length` 会被宿主的
          // `maxBytes`（默认 2 MiB）卡住，传大了直接抛 `workspace-file/too-large`
          // —— 这正是上一版字节回退「也失败」的原因。不传就是读整份文件，
          // 上限由 `maxFileBytes`（默认 32 MiB）管。
          bp = api.readBytes(sessionId, path, {})
        } catch (err) {
          finish(null, why + '；字节回退也失败：' + fileFailureLabel(err), null)
          return
        }
        Promise.resolve(bp).then(function (raw) {
          var un = unwrapRemote(raw)
          if (un.error !== undefined) {
            finish(null, why + '；readBytes 失败：' + un.error, null)
            return
          }
          var bytes = normalizeBytes(un.value)
          if (bytes === null || bytes.length === 0) {
            finish(null, why + '；readBytes 返回 ' + describeValue(raw), null)
            return
          }

          if (kind === null) {
            var cut = bytes.length > FILE_MAX_BYTES ? bytes.subarray(0, FILE_MAX_BYTES) : bytes
            if (indexOfNul(cut) >= 0) {
              finish(null, '读不到纯文本（文件含 NUL 字节，是二进制）', null)
              return
            }
            var text = decodeUtf8(cut)
            if (text === '') { finish(null, why, null); return }
            var all = text.split('\n')
            var truncated = bytes.length > FILE_MAX_BYTES || all.length > FILE_READ_LINES
            if (all.length > FILE_READ_LINES) text = all.slice(0, FILE_READ_LINES).join('\n')
            if (truncated) {
              text += '\n…（文件还有更多内容，此处只引用前 ' + FILE_READ_LINES + ' 行）'
            }
            finish(text, null, { lines: all.length, bytes: bytes.length, eof: !truncated, from: 'bytes' })
            return
          }

          return extractOfficeText(bytes, kind, path).then(function (t) {
            if (t === null || t === '') { finish(null, why + '；解包也没抽出正文', null); return }
            finish(t, null, { lines: null, bytes: bytes.length, eof: true, from: kind })
          })
        }).then(null, function (err) {
          // 回退也失败时，把真实原因带上 —— 不要用第一段的失败原因把第二段的
          // 原因盖掉，否则排查时看不到真正卡在哪一步。
          finish(null, why + '；字节回退也失败：' + fileFailureLabel(err), null)
        })
      }

      var pending = null
      try {
        // **一个参数都不传**：`limit` 传大了会被宿主以 `gateway/bad-request`
        // 挡掉（`limit must be at most <maxLines>`），而那个错误在实机上表现
        // 成「成功但正文为空」。用宿主自己的默认页大小最稳，行数截断我们在
        // 客户端自己再做一次。
        pending = api.read(sessionId, path, {})
      } catch (err) {
        viaBytes(fileFailureLabel(err))
        return
      }
      Promise.resolve(pending).then(function (raw) {
        // 先拆信封：`{ ok, value }` → 真正的 WorkspaceFileText。
        var un = unwrapRemote(raw)
        if (un.error !== undefined) { viaBytes('read 失败：' + un.error); return }
        var page = un.value
        // 返回形状放宽容一点：万一没包信封、直接是字符串也认。
        var text = ''
        if (typeof page === 'string') text = page
        else if (page !== null && page !== undefined && typeof page.text === 'string') text = page.text
        var meta = {
          lines: page !== null && page !== undefined && typeof page.lines === 'number' ? page.lines : null,
          bytes: page !== null && page !== undefined && typeof page.bytes === 'number' ? page.bytes : null,
          eof: page !== null && page !== undefined ? page.eof === true : null,
          from: 'text'
        }
        if (text === '') {
          // 宿主说成功、正文却空 —— 切到字节路，并把原始返回留给 DEBUG。
          dbg('workspaceFiles.read 返回了空正文，原始返回：', raw)
          viaBytes('read 返回 ' + describeValue(raw) + '，没有可用的 text')
          return
        }
        // 行数截断一律在客户端做，不管宿主的页大小是多少。
        var allLines = text.split('\n')
        if (allLines.length > FILE_READ_LINES) {
          text = allLines.slice(0, FILE_READ_LINES).join('\n') + '\n'
        }
        if (meta.eof === false || allLines.length > FILE_READ_LINES) {
          text += '…（文件还有更多内容，此处只引用前 ' + FILE_READ_LINES + ' 行）'
        }
        finish(text, null, meta)
      }, function (err) { viaBytes(fileFailureLabel(err)) })
    }

    /** 目录清单：只列名字，不读内容 —— 一个目录几十上百项，读内容没意义。 */
    function listDirForChat(ctx, sessionId, path, done) {
      var api = null
      try {
        api = ctx !== null && ctx !== undefined && ctx.remote !== undefined && ctx.remote !== null
          ? ctx.remote.workspaceFiles : null
      } catch (err) { api = null }
      if (api === null || api === undefined || typeof api.list !== 'function') {
        done(null, '当前连接没有目录列举接口', null)
        return
      }
      if (typeof sessionId !== 'string' || sessionId === '') {
        done(null, '拿不到会话 id', null)
        return
      }
      var pending = null
      try {
        pending = api.list(sessionId, path)
      } catch (err) {
        done(null, fileFailureLabel(err), null)
        return
      }
      Promise.resolve(pending).then(function (raw) {
        var un = unwrapRemote(raw)
        if (un.error !== undefined) { done(null, un.error, null); return }
        var res = un.value
        var entries = res !== null && res !== undefined && Array.isArray(res.entries) ? res.entries : null
        if (entries === null) { done(null, '目录列举返回了意外形状：' + describeValue(raw), null); return }
        var dirs = []
        var files = []
        for (var i = 0; i < entries.length; i++) {
          var e = entries[i]
          if (e === null || e === undefined || typeof e.name !== 'string') continue
          if (e.type === 'directory') dirs.push(e.name + '/')
          else files.push(e.name)
        }
        dirs.sort()
        files.sort()
        var all = dirs.concat(files)
        if (all.length === 0) { done('（空目录）', null, { lines: 0, bytes: null, eof: true, from: 'dir' }); return }
        var cap = Math.min(all.length, MAX_DIR_ENTRIES)
        var text = all.slice(0, cap).join('\n')
        if (all.length > cap) {
          text += '\n…（共 ' + all.length + ' 项，此处只列前 ' + cap + ' 项）'
        }
        done(text, null, { lines: all.length, bytes: null, eof: all.length <= cap, from: 'dir' })
      }, function (err) { done(null, fileFailureLabel(err), null) })
    }

    // ================================================================
    // 六、引用块拼装
    // ================================================================
    function composeBlock(snap, note) {
      var out = snap.kind === 'excel' ? composeTableBlock(snap) : composeTextBlock(snap)

      var n = note === null || note === undefined ? '' : String(note).trim()
      if (n !== '') {
        out.push('')
        out.push('批注：' + n)
      }
      return out.join('\n')
    }

    function composeTextBlock(snap) {
      var where = snap.path !== null && snap.path !== undefined && snap.path !== ''
        ? snap.path : '（未识别来源文件）'
      if (Array.isArray(snap.lineRange)) {
        where += snap.lineRange[0] === snap.lineRange[1]
          ? '（第 ' + snap.lineRange[0] + ' 行）'
          : '（第 ' + snap.lineRange[0] + '-' + snap.lineRange[1] + ' 行）'
      } else if (Array.isArray(snap.pageRange)) {
        // Word/PPT/PDF 走 PDF 渲染，没有源码行号，只有页
        where += snap.pageRange[0] === snap.pageRange[1]
          ? '（第 ' + snap.pageRange[0] + ' 页）'
          : '（第 ' + snap.pageRange[0] + '-' + snap.pageRange[1] + ' 页）'
      } else if (snap.kind === 'file') {
        where += '（' + (typeof snap.label === 'string' && snap.label !== '' ? snap.label : '整份文件') + '）'
      }

      var out = ['> \uD83D\uDCC4 引用来源：' + where]

      // 只发路径：正文留给模型自己去读。刻意不写成「正文未内联：<失败原因>」，
      // 因为这里根本不是失败 —— 是我们要的行为。
      if (snap.pathOnly === true) {
        out.push('> （未内联正文，请自行读取上面的路径）')
        return out
      }

      // 取不到正文：docx/xlsx/pdf 不是 UTF-8 文本（宿主回 `workspace-file/not-text`），
      // 或者超了单页上限。这种只发来源行 —— 模型自己有读文件的工具，拿着路径能
      // 自己去读；硬塞一段乱码或者假装有正文都比这糟。
      if (snap.text === null || snap.text === undefined) {
        var why = snap.failure !== null && snap.failure !== undefined && String(snap.failure) !== ''
          ? String(snap.failure) : '该格式不是纯文本'
        out.push('> （正文未内联：' + why + '）')
        return out
      }
      if (String(snap.text) === '') {
        out.push(snap.kind === 'file'
          ? '> （宿主返回的正文为空，请自行打开该文件查看）'
          : '> （选区里没有内容）')
        return out
      }

      var cap = snap.kind === 'file' ? MAX_FILE_LINES : MAX_EXCERPT_LINES
      var lines = String(snap.text).replace(/\r\n/g, '\n').split('\n')
      var clipped = lines.length > cap
      var shown = clipped ? lines.slice(0, cap) : lines

      for (var i = 0; i < shown.length; i++) out.push('> ' + shown[i])
      if (clipped) {
        out.push(snap.kind === 'file'
          ? '> …（文件更长，此处只引用前 ' + cap + ' 行）'
          : '> …（选区共 ' + lines.length + ' 行，此处截断）')
      }
      return out
    }

    /** 单元格文本在 Markdown 表格里的安全形式。 */
    function mdCell(s) {
      return String(s === null || s === undefined ? '' : s)
        .replace(/\r?\n/g, ' ')
        .replace(/\|/g, '\\|')
        .trim()
    }

    /**
     * Excel 选区 → Markdown 表格（嵌在引用块里）。
     * 首行当表头 —— 从表格里框选时，第一行几乎总是列名，这样模型读起来最省。
     */
    function composeTableBlock(snap) {
      var where = snap.path !== null && snap.path !== undefined && snap.path !== ''
        ? snap.path : '（未识别来源文件）'
      var sheet = snap.sheet !== null && snap.sheet !== undefined && snap.sheet !== ''
        ? snap.sheet : '?'
      where += '（工作表 ' + sheet + ' · 区域 ' + snap.address + '）'

      var out = ['> \uD83D\uDCC4 引用来源：' + where]
      var rows = snap.table
      if (!Array.isArray(rows) || rows.length === 0) {
        out.push('> （未能读取单元格数值，仅发送了区域坐标）')
        return out
      }

      var shown = rows.length > MAX_EXCERPT_ROWS ? rows.slice(0, MAX_EXCERPT_ROWS) : rows
      var width = 0
      for (var i = 0; i < shown.length; i++) {
        if (Array.isArray(shown[i]) && shown[i].length > width) width = shown[i].length
      }
      if (width === 0) {
        out.push('> （选区里没有内容）')
        return out
      }
      // 单行 / 单列不套 Markdown 表格：Markdown 的表必须有表头，把唯一一行当
      // 表头会渲染成一张没有数据行的空表 —— 用户实测选单格 B8 得到的就是
      // 「> | 值 |」加「> | --- |」。单列同理：一列竖线排下来，不如每行一个值。
      if (rows.length === 1) {
        var one = []
        for (var c1 = 0; c1 < width; c1++) {
          var s1 = mdCell(rows[0] === undefined ? '' : rows[0][c1])
          if (s1 !== '') one.push(s1)
        }
        out.push('> ' + (one.length > 0 ? one.join(' | ') : '（空单元格）'))
        return out
      }
      if (width === 1) {
        for (var r1 = 0; r1 < shown.length; r1++) {
          out.push('> ' + mdCell(shown[r1] === undefined ? '' : shown[r1][0]))
        }
        if (rows.length > MAX_EXCERPT_ROWS) {
          out.push('> …（选区共 ' + rows.length + ' 行，此处截断）')
        }
        return out
      }
      for (var r = 0; r < shown.length; r++) {
        var cells = []
        for (var c = 0; c < width; c++) cells.push(mdCell(shown[r] === undefined ? '' : shown[r][c]))
        out.push('> | ' + cells.join(' | ') + ' |')
        if (r === 0) {
          var sep = []
          for (var k = 0; k < width; k++) sep.push('---')
          out.push('> | ' + sep.join(' | ') + ' |')
        }
      }
      if (rows.length > MAX_EXCERPT_ROWS) {
        out.push('> …（选区共 ' + rows.length + ' 行，此处截断）')
      }
      return out
    }

    // ================================================================
    // 六·B、已发送消息的展示层折叠
    // ================================================================
    // 发给模型的文本一个字都不改（引用块全文照旧），只在**渲染出来的 DOM** 上把
    // 那几行 `> …` 收成一个 chip：悬停看全文，点击展开原文、再点收回。
    //
    // 为什么不注册自己的消息渲染器：`conversation.chat.node` 的 user 档已被内核
    // 占用，接管它等于重写整条用户消息的渲染（附件、引用、投影都得自己来），
    // 而这里要的只是「换个显示」。DOM 后处理是最小、也最容易整个摘掉的做法。
    //
    // 代价：React 重渲染会把我插进去的节点抹掉，所以每次 DOM 变动都要重扫。
    // 开销靠三层过滤压住 —— ① WeakSet 记住已折好的行（一次 O(1) 命中就跳过）；
    // ② 行级 `> ` 粗筛，不必遍历子树；③ 只在「自身含引用行、子元素不含」的
    // 最深容器上动手，不假设正文容器是 div 还是 span。

    // `[data-submission-echo]` 是**必需**的一项，不是补充：
    // 点发送之后、宿主的正式消息节点落地之前，DSH 会先渲染一条「本地乐观回显」，
    // 把草稿原文整条摆出来。这在 `_asar_x/dsh-client-ui-chat/lib/client.js` 里
    // 的 `PendingSubmissionBubble` 注释里写得很直白 ——
    //   "visible from the submit click until the durable `user/message` or
    //    steering occurrence renders"
    // 那条回显**不是 chat 节点**（它来自 `useSession(s => s.pendingSubmissions)`），
    // 行上没有 `data-chat-flow-kind` —— 只按 flow-kind 选就必然漏掉它。用户看到的
    // 「先显示全部、再缩成胶囊」正是它：回显是全文，随后被我们折好的正式气泡替换。
    // 队列停靠栏（QueueDock，`_asar_x/dsh-client-ui-conversation/lib/client.js`）里
    // 排队的提交渲染时也带同一个属性，一并覆盖。
    var FOLD_SEL = '[data-chat-flow-kind="user"], [data-chat-flow-kind="steering"],'
      + ' [data-pending-steering], [data-submission-echo]'
    var FOLD_CHIP = 'data-f2c-chip'
    var FOLD_DONE = 'data-f2c-done'
    var FOLD_NOTE = 'data-f2c-note'
    var FOLDED_ROWS = new WeakSet()

    // ---------------- 「封膜」：让未折叠的形态根本没有机会被绘制 ----------------
    //
    // 问题：发出消息时，新那一行进入 DOM 到我们把它折成胶囊之间，只要浏览器
    // 抢到一次绘制，用户就会看到「先全文明晃晃地弹出来、再缩成胶囊」。
    //
    // 单靠「折得够快」是有风险的（我们不知道 React 会怎么改这棵子树）。所以再加
    // 一层倒过来的保险：**在没被我们处理过之前，那一行本来就是不可见的**。
    //   · 处理过的行打 data-f2c-seen；
    //   · 封膜开启期间，没有 seen 的用户/引导行在 CSS 上 visibility:hidden；
    //   · 于是「全文」形态既使存在也画不出来，用户直接看到折好的样子。
    // 封膜的开关挂在「输入框里是否挂着引用胶囊」上 —— 那正是「马上要发出去」的
    // 精确信号；发送后草稿清空，再宽限两秒收膜，绝不留后患。
    var SEAL_ATTR = 'data-f2c-seal'
    // 「这一行是被封膜按住的」。它是**收膜的唯一凭据**：被按住的行一旦露面，
    // 就说明那条要发送的消息已经落地，封膜的使命完成了。
    var HOLD_ATTR = 'data-f2c-held'
    // 草稿里的胶囊占位符。**不能想当然地只查 U+FFFC** —— 内核自己的
    // REFERENCE_PLACEHOLDER_RE 是 /[\uE100-\uE11D\uFFFC]/gu，胶囊究竟占哪个码位
    // 属于内核私有实现。早期版本只判断 U+FFFC，于是封膜一次都没打开过，那条
    // visibility:hidden 规则形同虚设（这正是「先全文、再折叠」始终没修掉的原因）。
    var CHIP_MARK_RE = /[\uE100-\uE11D\uFFFC]/
    var SEAL_OFF_TIMER = 0
    var SEAL_MAX_TIMER = 0
    var SEAL_GRACE_MS = 4000
    // 绝对上限，且**每次 sealOn 都重新计时**，所以它的语义是「距最后一次确认
    // 已经过去多久」，而不是「总时长」。只有这样，用户在输入框里慢慢写批注
    // 也不会把膜等没。
    var SEAL_MAX_MS = 600000
    function sealOn() {
      if (SEAL_OFF_TIMER !== 0) { clearTimeout(SEAL_OFF_TIMER); SEAL_OFF_TIMER = 0 }
      try {
        if (document.body.getAttribute(SEAL_ATTR) === null) {
          document.body.setAttribute(SEAL_ATTR, '')
          dbg('封膜开启：接下来的新消息行在折叠完成前不可见')
        }
        // 绝对上限：胶囊组件万一不在（右侧栏那个标签页没激活就不挂载），
        // 就没有人来收膜，之后的新消息会一直「先隐身一瞬」。给一条无条件兜底。
        if (SEAL_MAX_TIMER !== 0) clearTimeout(SEAL_MAX_TIMER)
        SEAL_MAX_TIMER = setTimeout(function () {
          SEAL_MAX_TIMER = 0
          sealOffNow()
        }, SEAL_MAX_MS)
      } catch (err) { void err }
    }
    function sealOffNow() {
      if (SEAL_OFF_TIMER !== 0) { clearTimeout(SEAL_OFF_TIMER); SEAL_OFF_TIMER = 0 }
      if (SEAL_MAX_TIMER !== 0) { clearTimeout(SEAL_MAX_TIMER); SEAL_MAX_TIMER = 0 }
      try { document.body.removeAttribute(SEAL_ATTR) } catch (err) { void err }
    }
    function sealOffSoon() {
      if (SEAL_OFF_TIMER !== 0) return
      // 目前没有调用点：收膜改由「被按住的新行已经露面」驱动（见 markSeen），
      // 那是一个不依赖草稿占位符、也不会被打字期间的全量扫描误触发的信号。
      // 保留这个函数是因为它仍然是「按时间宽限收膜」的现成实现，将来若需要
      // 一条基于时间的收膜路径可以直接用。
      try { if (document.body.getAttribute(SEAL_ATTR) === null) return } catch (err) { return }
      SEAL_OFF_TIMER = setTimeout(function () {
        SEAL_OFF_TIMER = 0
        sealOffNow()
      }, SEAL_GRACE_MS)
    }
    /** 标记这一行已被我们处理过（折过，或确认无需折叠）。属性写入不被观察，不会自激。 */
    function markSeen(row) {
      try {
        if (row.getAttribute('data-f2c-seen') === null) row.setAttribute('data-f2c-seen', '')
        // 放开 holdNewRows 按下的那一行。按住用的是内联样式，放开也必须用
        // 内联样式 —— 否则会被自己写下的那条 inline 规则一直压着。
        if (row.style !== undefined && row.style.visibility === 'hidden') {
          row.style.visibility = ''
          dbg('放开 文本长度=' + ((row.textContent || '').length))
        }
        // 刚到达的那一行已经露面 = 那条要发送的消息落地了 -> 收膜。
        // 收膜只认这一个信号：它既不依赖草稿占位符，也不会被「打字期间的全量
        // 扫描」误触发（打字时不会新增消息行）。
        if (row.getAttribute(HOLD_ATTR) !== null) {
          row.removeAttribute(HOLD_ATTR)
          sealOffNow()
        }
      } catch (err) { void err }
      // 这里**故意不叫 sealOffSoon**。markSeen 在用户打字期间就会被反复调到
      // （编辑器每次变更都会触发一次全量扫描），收膜倒计时会在用户还没按发送
      // 时就走完，膜提前落下，行落地时又退化成「先全文、再折叠」。
      // 收膜只由「草稿失去胶囊」（真正的发送信号）驱动，外加 sealOn 里的绝对上限。
    }

    // 「批注：…」那一行。折叠只收引用正文，批注要原样留着给用户看。
    var NOTE_LINE = /^\s*(?:批注|备注)\s*[：:]\s*/

    function hashText(text) {
      var h = 5381
      for (var i = 0; i < text.length; i++) h = ((h << 5) + h + text.charCodeAt(i)) >>> 0
      return h.toString(36)
    }

    function baseName(p) {
      var s = String(p)
      var cut = Math.max(s.lastIndexOf('/'), s.lastIndexOf('\\'))
      return cut >= 0 ? s.slice(cut + 1) : s
    }

    /**
     * 文件树拼子路径用的是 `父路径 + '/' + 名字`，所以 Windows 上会拿到
     * `C:\Users\…\MyDemo/_asar_list.js` 这种**混用分隔符**的串。宿主两种都认，
     * 但写进引用块给模型看很难看 —— 统一成反斜杠。
     * 只规整盘符开头的本机路径，UNC / POSIX 原样放行。
     */
    function tidyPath(p) {
      var s = String(p)
      if (!/^[A-Za-z]:[\\/]/.test(s)) return s
      return s.replace(/\//g, '\\')
    }

    /** 引用块首行 → 给用户看的短标签，如 `📄 套餐对比.xlsx · 工作表 Sheet1 · 区域 B8`。 */
    function describeQuote(quoteText) {
      var first = String(quoteText).split('\n')[0].trim()
      var m = /^\uD83D\uDCC4\s*引用来源[：:]\s*(.*)$/.exec(first)
      var tail = m !== null ? m[1].trim() : first
      // 末尾那对括号里是「第 N-M 行」或「工作表 X · 区域 B8」，长路径只留文件名
      var pm = /^(.*)（([^（）]*)）$/.exec(tail)
      if (pm === null) return { label: baseName(tail) }
      return { label: baseName(pm[1]) + ' · ' + pm[2] }
    }

    /**
     * 把一段文本切成「引用段 / 普通段」交替的序列。
     * 引用段 = 连续以 `> ` 开头的行，普通段原样保留换行（别把多行正文挤成一行）。
     *
     * **引用段一旦开始就不在中间断开** —— 这一点很关键：当引用的正是一份会话
     * 日志、而日志里记着我们自己发过的引用块时，正文里会出现 `> 📄 引用来源：`
     * 行。若在那里切开，一整块正文会被撕成好几段、各挂一枚胶囊（实测踩过）。
     * 新块只能由**非引用行**（空行 / 批注行）隔开。
     */
    function splitQuoteRuns(text) {
      var lines = String(text).split('\n')
      var parts = []
      var buf = []
      var plain = []

      var isSource = function (line) { return /^\s*>\s*\uD83D\uDCC4\s*引用来源[：:]/.test(line) }

      var flushQuote = function () {
        if (buf.length === 0) return
        var raw = buf.join('\n')
        // 只有**以来源行开头**的引用段才是本插件发出去的引用块。普通的 Markdown
        // 引用（`> 别人说的话`）绝不能被折成胶囊 —— 那不是我们的东西，
        // 折了就是把用户自己写的正文藏起来。
        var ours = isSource(buf[0])
        var body = buf.map(function (l) { return l.replace(/^\s*>\s?/, '') }).join('\n').trim()
        parts.push(ours && body !== '' ? { quote: true, text: raw, body: body } : { quote: false, text: raw })
        buf = []
      }
      var flushPlain = function () {
        if (plain.length === 0) return
        parts.push({ quote: false, text: plain.join('\n') })
        plain = []
      }

      for (var i = 0; i < lines.length; i++) {
        var line = lines[i]
        if (/^\s*>\s?/.test(line)) {
          flushPlain()
          buf.push(line)
        } else {
          flushQuote()
          plain.push(line)
        }
      }
      flushQuote()
      flushPlain()
      return parts
    }

    function clearChildren(el) {
      while (el.firstChild !== null) el.removeChild(el.firstChild)
    }

    // ---------------- DSH 自带线性图标 ----------------
    // 原样取自 @deepseek-ai/dsh-client-ui-primitives 的 IconListPenOutlineArtwork
    // （16×16 viewBox、fill:none、stroke-width 1 的 "Regular one-pixel" 变体）：
    // 一份带笔的清单，正是「在文档上做批注」。
    // 之所以内联而不 require 那个包：浮层是命令式 DOM，为一个静态图标挂 React 根
    // 不划算；图标本身是静态资源，DSH 自己也是编译期内联的。
    // width/height 属性是**兜底**，不是主要尺寸来源（正常由 .f2c-ic svg 的 CSS 定）。
    // 原因：这个 <svg> 只有 viewBox、没有内在尺寸，一旦 CSS 没到位（例如样式表
    // 因某种原因陈旧），它会被当成无固有尺寸的替换元素而塌缩成看不见 —— 表现为
    // 「图标消失 + 文字看上去不居中」。实测复现过：浮层按钮变成一条偏心的空胶囊。
    var ICON_ANNOTATE = '<svg viewBox="0 0 16 16" width="14" height="14" fill="none"'
      + ' xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false"'
      + ' stroke-width="1">'
      + '<path d="M4.9375 5.90295H11.0625" stroke="currentColor"/>'
      + '<path d="M4.9375 9.02991H8.27841" stroke="currentColor"/>'
      + '<path d="M12.5 1.32617C13.3039 1.32617 14 1.95171 14 2.77637V7.61328L13 8.68164V2.77637'
      + 'C13 2.55186 12.8007 2.32617 12.5 2.32617H3.5C3.1993 2.32617 3 2.55186 3 2.77637V13.2246'
      + 'C3.00044 13.4489 3.19963 13.6738 3.5 13.6738H8.32812L7.39258 14.6738H3.5'
      + 'C2.69637 14.6738 2.00042 14.0489 2 13.2246V2.77637C2 1.95171 2.69613 1.32617 3.5 1.32617H12.5Z"'
      + ' fill="currentColor"/>'
      + '<path d="M8.97212 14.3693C9.17511 14.5723 9.37811 14.7753 9.5811 14.9783'
      + 'C9.67012 14.8953 9.75914 14.8123 9.84815 14.7293C11.4505 13.2352 13.0528 11.7411 14.6551 10.247'
      + 'C14.7441 10.164 14.8331 10.081 14.9221 9.99803C14.5989 9.6748 14.2756 9.35157 13.9524 9.02834'
      + 'C13.8694 9.11736 13.7864 9.20637 13.7034 9.29539C12.2093 10.8977 10.7152 12.5 9.22113 14.1023'
      + 'C9.13813 14.1913 9.05513 14.2803 8.97212 14.3693Z" fill="currentColor"/>'
      + '<path d="M11.6323 13.7841C11.6323 14.0395 11.6323 14.295 11.6323 14.5504'
      + 'C11.6812 14.5523 11.7301 14.5543 11.779 14.5562C12.659 14.5913 13.539 14.6263 14.419 14.6614'
      + 'C14.4679 14.6633 14.5168 14.6653 14.5657 14.6672C14.5657 14.3339 14.5657 14.0006 14.5657 13.6672'
      + 'C14.5168 13.6692 14.4679 13.6711 14.419 13.6731C13.539 13.7081 12.659 13.7432 11.779 13.7783'
      + 'C11.7301 13.7802 11.6812 13.7821 11.6323 13.7841Z" fill="currentColor"/>'
      + '</svg>'

    /** 造一个图标元素。尺寸与配色交给 CSS 的 .f2c-ic，不必每处传参。 */
    function makeIcon() {
      var span = document.createElement('span')
      span.className = 'f2c-ic'
      span.innerHTML = ICON_ANNOTATE
      return span
    }

    // ---------------- 展开之后的自动合上 ----------------
    // 点开胶囊是想瞄一眼原文；合上的时机是「点到别的地方去了」。
    //
    // 判据是「这次 pointerdown 的目标落没落在展开的那一块里」。host 就是被我们
    // 换掉内容的那个容器，展开后它的矩形正好等于摊开的原文范围 —— 点在它里面
    // （点胶囊自己、或者去选原文）一律不动作，点在它外面就折回去。
    //
    // 监听挂在**捕获阶段**：胶囊自己的 click 里有 stopPropagation，冒泡阶段会漏；
    // 捕获阶段先于一切处理器，而且此刻拿到的 e.target 就是真实的点击目标。
    var OPEN_CHIP = null            // { host, close }
    var AUTO_CLOSE_ARMED = false

    function autoCloseNow() {
      var open = OPEN_CHIP
      if (open === null) return
      OPEN_CHIP = null
      try { open.close() } catch (err) { void err }
    }

    function onAutoCloseDown(e) {
      var open = OPEN_CHIP
      if (open === null) return
      var node = nodeToElement(e.target)
      if (node !== null && open.host.contains !== undefined && open.host.contains(node)) return
      autoCloseNow()
    }

    /** 登记「当前展开的那一个」。同一时刻只允许开一个，新开先合旧的。 */
    function armAutoClose(host, close) {
      autoCloseNow()
      OPEN_CHIP = { host: host, close: close }
      if (AUTO_CLOSE_ARMED) return
      AUTO_CLOSE_ARMED = true
      document.addEventListener('pointerdown', onAutoCloseDown, true)
    }

    function makeQuoteChip(host, body, total, seq) {
      var info = describeQuote(body)
      var chip = document.createElement('span')
      chip.className = 'f2c-sent'
      chip.setAttribute(FOLD_CHIP, '')
      chip.setAttribute('role', 'button')
      chip.setAttribute('tabindex', '0')
      chip.title = body
      chip.appendChild(makeIcon())
      chip.appendChild(document.createTextNode(info.label + (total > 1 ? ' #' + seq : '')))
      // 合上：只在确实开着的时候动手，避免无谓的重排。
      var close = function () {
        var raw = host.__f2cRaw
        if (typeof raw !== 'string') return
        if (host.getAttribute(FOLD_DONE) === 'open') renderFold(host, raw)
      }
      var open = function () {
        var raw = host.__f2cRaw
        if (typeof raw !== 'string') return
        if (host.getAttribute(FOLD_DONE) === 'open') return
        renderPlain(host, raw)
        armAutoClose(host, close)
      }
      var toggle = function (e) {
        e.preventDefault()
        e.stopPropagation()
        if (host.getAttribute(FOLD_DONE) === 'open') {
          autoCloseNow()
          close()
        } else open()
      }
      chip.addEventListener('click', toggle)
      chip.addEventListener('keydown', function (e) {
        if (e.key === 'Enter' || e.key === ' ') toggle(e)
      })
      return chip
    }

    /**
     * 折：把引用块换成 chip，紧随其后的批注并到**同一行**。
     *
     * 为什么必须在这里重排行，而不是继续顺序铺普通段：
     * 「批注：xxx」在正文里是引用块**后面**独立的一段，顺序铺就必然落到 chip
     * 的下一行、还带一个方框。而输入框里的观感是 `[chip] 批注` 挤在一行 ——
     * 折叠后要还原成那个样子，用户才不会觉得「发出去变了样」。
     *
     * 同时把 `批注：` 前缀丢掉：输入框里也没有它，留着只会让聊天显得比输入框重。
     * 展开（renderPlain）仍然还原成完整原文，一个字不少。
     */
    function renderFold(el, raw) {
      var parts = splitQuoteRuns(raw)
      var total = 0
      for (var i = 0; i < parts.length; i++) if (parts[i].quote) total++
      if (total === 0) return false

      // 先把 parts 规整成「行」：引用块 + 紧跟其后的批注合成一行，其余普通文本各自成行。
      var rows = []
      var seq = 0
      var last = null                    // 最近一个还能接收批注的引用行
      var pushText = function (s) {
        var t = s.replace(/^\n+/, '').replace(/\n+$/, '')
        if (t.trim() !== '') { rows.push({ text: t }); last = null }
      }
      for (var j = 0; j < parts.length; j++) {
        var part = parts[j]
        if (part.quote) {
          seq++
          last = { body: part.body, seq: seq, note: null }
          rows.push(last)
          continue
        }
        var segs = String(part.text).split('\n')
        var buf = []
        for (var k = 0; k < segs.length; k++) {
          if (!NOTE_LINE.test(segs[k])) { buf.push(segs[k]); continue }
          pushText(buf.join('\n'))
          buf = []
          var note = segs[k].replace(NOTE_LINE, '').trim()
          // 紧跟引用块的批注并进那一行；孤立的批注行原样保留（带「批注：」前缀）
          if (last !== null && last.note === null) last.note = note
          else rows.push({ text: segs[k] })
          last = null
        }
        pushText(buf.join('\n'))
      }

      if (rows.length === 0) return false
      clearChildren(el)
      for (var r = 0; r < rows.length; r++) {
        var row = rows[r]
        if (r > 0) el.appendChild(document.createElement('br'))
        if (row.body !== undefined) {
          el.appendChild(makeQuoteChip(el, row.body, total, row.seq))
        }
        if (row.note !== null && row.note !== undefined && row.note !== '') {
          var noteEl = document.createElement('span')
          noteEl.className = 'f2c-sent-note'
          noteEl.setAttribute(FOLD_NOTE, '')
          noteEl.textContent = ' ' + row.note
          el.appendChild(noteEl)
        }
        if (row.text !== undefined && row.text !== '') {
          el.appendChild(document.createTextNode(row.text))
        }
      }
      el.setAttribute(FOLD_DONE, hashText(raw))
      return true
    }

    /** 展开：还原成原样的文本。标记成 open，扫描器就不会再折回去。 */
    function renderPlain(el, raw) {
      clearChildren(el)
      el.appendChild(document.createTextNode(raw))
      el.setAttribute(FOLD_DONE, 'open')
    }

    /**
     * @param scope 文档，或任意元素。元素**自己**命中 FOLD_SEL 时也算进去 ——
     *              MutationObserver 递进来的常常正是一整行。
     */
    function foldSentQuotes(scope) {
      var root = scope === undefined || scope === null ? document : scope
      var rows = []
      if (root.nodeType === 9) {
        rows = root.querySelectorAll(FOLD_SEL)
      } else {
        if (root.matches !== undefined && root.matches(FOLD_SEL)) rows.push(root)
        var found = root.querySelectorAll(FOLD_SEL)
        for (var fi = 0; fi < found.length; fi++) rows.push(found[fi])
      }
      var folded = 0
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i]
        try {
          // ⓪ 揭膜。**空行先别揭**：React 往往先挂一个空壳、下一批变更才把正文
          //    写进来；提前揭膜等于把「正文到货之前那一帧」暴露出去，那恰恰就是
          //    闪动的来源。这里下面的处理是同步的，所以揭膜与折叠仍属同一次微任务。
          if ((row.textContent || '') !== '') markSeen(row)
          // ① 折过、且 chip 仍在 → 跳过。省掉下面 textContent 这种要拷贝整段
          //    文本的操作 —— 流式回答期间这个函数每帧都可能被叫醒。
          if (FOLDED_ROWS.has(row) && row.querySelector('[' + FOLD_CHIP + ']') !== null) {
            continue
          }
          // ② 行级粗筛：整条气泡里一个 `> ` 行都没有就直接走人
          if (!/^\s*>\s?/m.test(row.textContent || '')) continue

          // ③ 找「自身含引用行、但子元素都不含」的最深容器
          var targets = []
          var all = row.querySelectorAll('*')
          for (var j = 0; j < all.length; j++) {
            var el = all[j]
            if (el.getAttribute(FOLD_CHIP) !== null) continue
            var text = el.textContent
            if (text === null || text === '' || !/^\s*>\s?/m.test(text)) continue
            var childHit = false
            for (var k = 0; k < el.children.length; k++) {
              if (/^\s*>\s?/m.test(el.children[k].textContent || '')) { childHit = true; break }
            }
            if (childHit) continue
            targets.push(el)
          }
          if (targets.length === 0) continue

          var did = 0
          for (var n = 0; n < targets.length; n++) {
            var host = targets[n]
            // 用户手动展开过就别折回去（React 重渲染会清掉这个标记，那时重折是对的）
            if (host.getAttribute(FOLD_DONE) === 'open') continue
            var raw = host.textContent
            if (raw === null || !/^\s*>\s?/m.test(raw)) continue
            host.__f2cRaw = raw
            if (renderFold(host, raw)) did++
          }
          if (did > 0) FOLDED_ROWS.add(row)
          folded += did
        } catch (err) {
          console.warn('[file-to-chat] 折叠已发送引用块失败：', err)
        }
      }
      if (folded > 0) dbg('折 ' + folded + ' 处 候选=' + rows.length)
      return folded
    }

    /** 诊断用：第一次**同步**折叠成功时记一笔（DEBUG 打开时才输出）。 */
    var FOLD_SYNC_LOGGED = false
    function noteSyncFold(n) {
      if (FOLD_SYNC_LOGGED || n <= 0) return
      FOLD_SYNC_LOGGED = true
      dbg('同步折叠生效（' + n + ' 处，发生在绘制之前）')
    }

    /**
     * 顺着祖先（含自己）找最近的一条「用户消息行」，找到就当场折它。
     *
     * 这条路径是**闪动的关键**：React 重新渲染一条消息时，绝大多数情况下并
     * 不替换那一行本身，只替换行**里面**的文本节点/span。此时新增的节点既
     * 不是 FOLD_SEL、子孙里也没有 FOLD_SEL —— 只按「新子树」判断会全部漏掉，
     * 只能等下一帧的 rAF 全量扫描，用户看到的就是「先显示全文、再快速折叠」。
     */
    function foldRowNear(node) {
      var el = node
      if (el === null || el === undefined) return
      if (el.nodeType === 3) el = el.parentElement          // 文本节点（characterData）
      if (el === null || el === undefined || el.nodeType !== 1) return
      if (typeof el.closest !== 'function') return
      try {
        var row = el.matches !== undefined && el.matches(FOLD_SEL) ? el : el.closest(FOLD_SEL)
        if (row !== null) noteSyncFold(foldSentQuotes(row))
      } catch (err) {
        console.warn('[file-to-chat] 折叠增量扫描失败：', err)
      }
    }

    /**
     * 对一棵「刚被插进 DOM 的子树」做一次同步折叠。
     *
     * 为什么必须在 MutationObserver 的回调里当场折，而不是丢给 rAF：
     * MO 回调是微任务，紧跟 DOM 变更、稳在浏览器绘制之前跑完；rAF 却要等
     * 下一次渲染步 —— 一旦 React 这次提交本身就发生在某一帧的渲染步里，
     * 「未折叠的全文」这一形态就会被真真切切画出来一帧，用户看到的就是
     * 「发出去先弹一下全文、再缩成胶囊」。同步折掉之后，那个中间形态
     * 从来没有机会被绘制。
     */
    /**
     * 把「刚刚出现的」消息行当场按住（内联 visibility:hidden），等折叠完成
     * （markSeen）再放开。
     *
     * 这一手不依赖任何时机判断 —— 行一进 DOM 就按。封膜那套
     * （`body[data-f2c-seal]` + `:not([data-f2c-seen])`）需要我们先知道
     * 「马上就要发送」，而这一步只需要「有个新行出现了」。两者互为冗余，
     * 任一生效都不会有闪动。
     */
    function holdNewRows(node) {
      var list = []
      try {
        if (node.matches !== undefined && node.matches(FOLD_SEL)) list.push(node)
        var found = node.querySelectorAll(FOLD_SEL)
        for (var i = 0; i < found.length; i++) list.push(found[i])
      } catch (err) { return }
      for (var j = 0; j < list.length; j++) {
        var row = list[j]
        // 早就处理过的行不要动 —— 它可能正是当前正在显示的那条消息。
        if (row.getAttribute('data-f2c-seen') !== null) continue
        try {
          if (row.style.visibility !== 'hidden') row.style.visibility = 'hidden'
          row.setAttribute(HOLD_ATTR, '')
          dbg('按住 文本长度=' + ((row.textContent || '').length))
        } catch (err) { continue }
        // 兜底：万一折叠那一路抛异常、没来得及放开，1.5 秒后无条件当作已处理。
        // 注意必须调 markSeen 而不是只清内联样式 —— 清样式只解开了「按住」那一道，
        // 封膜那条 CSS 规则还压着它（行上没有 seen 属性）。宁可闪一下，也绝不把
        // 用户的消息藏起来。
        ;(function (r) {
          setTimeout(function () {
            try { dbg('兜底放开'); markSeen(r) } catch (err) { void err }
          }, 1500)
        })(row)
      }
    }

    function foldIn(node) {
      if (node === null || node.nodeType !== 1) return
      if (typeof node.querySelector !== 'function') return
      // ① 这棵新子树自己就是目标行、或含有目标行 → 折它（覆盖面最大）
      var isRow = node.matches !== undefined && node.matches(FOLD_SEL)
      if (isRow || node.querySelector(FOLD_SEL) !== null) {
        holdNewRows(node)
        try {
          noteSyncFold(foldSentQuotes(node))
        } catch (err) {
          console.warn('[file-to-chat] 折叠增量扫描失败：', err)
        }
        return
      }
      // ② 新节点落在某一行**内部** → 顺祖先找那一行（见 foldRowNear 的说明）
      foldRowNear(node)
    }

    // ================================================================
    // 七、划词浮层（命令式 DOM，挂到 document.body）
    // ================================================================
    function createOverlay(ctx, readSessionId) {
      var layer = document.createElement('div')
      layer.className = 'f2c-layer'
      document.body.appendChild(layer)
      // 把主题 token 的计算值搬到浮层自己身上：浮层在 body 上，未必处在
      // 主题变量的作用域里；写死在 layer 上则子元素直接继承，稳。
      syncTheme(layer)

      var bar = null          // 「批注」按钮
      var card = null         // 批注卡片
      var menu = null         // 文件树右键菜单
      var snap = null         // 当前选区 / 文件快照
      var flashTimer = null
      var noteDraft = ''      // 卡片里那行批注的草稿，跨重画保留

      function clearBar() {
        if (bar !== null && bar.parentNode !== null) bar.parentNode.removeChild(bar)
        bar = null
      }

      function clearCard() {
        if (card !== null && card.parentNode !== null) card.parentNode.removeChild(card)
        card = null
      }

      function clearMenu() {
        if (menu !== null && menu.parentNode !== null) menu.parentNode.removeChild(menu)
        menu = null
      }

      function clearAll() {
        clearBar()
        clearCard()
        clearMenu()
        noteDraft = ''
        snap = null
      }

      /** 上方优先（用户要求）；上方放不下才回落下方。 */
      function placeAbove(top, bottom, left, right, width, height) {
        var w = width
        var l = (left + right) / 2 - w / 2
        l = Math.max(8, Math.min(l, window.innerWidth - w - 8))
        var t = top - height - 8
        if (t < 8) t = bottom + 8
        if (t + height > window.innerHeight - 8) t = Math.max(8, window.innerHeight - height - 8)
        return { left: Math.round(l), top: Math.round(t) }
      }

      /** 选区值 → 只读小表格。首行当表头。 */
      function buildGrid(rows) {
        var t = document.createElement('table')
        var width = 0
        for (var i = 0; i < rows.length; i++) {
          if (Array.isArray(rows[i]) && rows[i].length > width) width = rows[i].length
        }
        for (var r = 0; r < rows.length; r++) {
          var tr = document.createElement('tr')
          for (var c = 0; c < width; c++) {
            var v = rows[r] === undefined || rows[r][c] === undefined ? '' : String(rows[r][c])
            var cell = document.createElement(r === 0 ? 'th' : 'td')
            cell.textContent = v === '' ? '·' : v
            if (v === '') cell.className = 'f2c-empty'
            cell.title = v
            tr.appendChild(cell)
          }
          t.appendChild(tr)
        }
        return t
      }

      // ---------------- 批注按钮 ----------------
      function showBar() {
        clearBar()
        syncTheme(layer)   // 每次浮现都重取一次，跟随浅色/深色切换
        bar = document.createElement('button')
        bar.type = 'button'
        bar.className = 'f2c-bar'
        bar.appendChild(makeIcon())
        bar.appendChild(document.createTextNode('批注'))
        bar.title = '为选中的内容写一条批注，连同选区一起发送到当前聊天'

        // 关键：preventDefault 保住原生选区，否则 mousedown 会先清空它
        bar.addEventListener('mousedown', function (e) { e.preventDefault() })
        bar.addEventListener('click', function (e) {
          e.preventDefault()
          e.stopPropagation()
          openCard()
        })
        layer.appendChild(bar)
        positionBar()
      }

      function positionBar() {
        if (bar === null) return
        if (snap !== null && snap.kind === 'excel') {
          // Excel 快照没有 DOM Range，锚点就是当时按下的那个坐标，固定不动。
          var e = placeAbove(snap.top, snap.bottom, snap.left, snap.right, 92, 28)
          bar.style.left = e.left + 'px'
          bar.style.top = e.top + 'px'
          return
        }
        var r = liveRect()
        if (r === null) { clearBar(); return }
        snap.top = r.top
        snap.bottom = r.bottom
        snap.left = r.left
        snap.right = r.right
        var p = placeAbove(r.top, r.bottom, r.left, r.right, 92, 28)
        bar.style.left = p.left + 'px'
        bar.style.top = p.top + 'px'
      }

      /** 从仍然有效的 range 里重新量一次坐标；range 失效返回 null。 */
      function liveRect() {
        if (snap === null || snap.kind === 'excel' || snap.range === null) return null
        try {
          var rects = snap.range.getClientRects()
          if (rects === null || rects.length === 0) return null
          var f = rects[0]
          var l = rects[rects.length - 1]
          if (f.width === 0 && f.height === 0) return null
          return {
            top: f.top,
            bottom: l.bottom,
            left: Math.min(f.left, l.left),
            right: Math.max(f.right, l.right),
          }
        } catch (_) { return null }
      }

      // ---------------- 批注卡片 ----------------
      // ---------------- 文件树的右键菜单 ----------------
      // 内核没有给「行级菜单」留座位（`sidebar.right.tab.files.actions` 是工具栏
      // 那一条），而文件树的行上带着稳定的 `data-files-entry` / `data-files-path`
      // （`@deepseek-ai/dsh-client-ui-sidebar-files` 的 FilesBody 输出），所以走 DOM
      // 接管右键：只在认得的那一行上 preventDefault，别处一律放行原生菜单。
      function menuItem(label, onClick) {
        var item = document.createElement('button')
        item.type = 'button'
        item.className = 'f2c-item'
        item.textContent = label
        item.addEventListener('mousedown', function (e) { e.preventDefault() })
        item.addEventListener('click', function (e) {
          e.preventDefault()
          e.stopPropagation()
          onClick()
        })
        return item
      }

      function copyToClipboard(text) {
        try {
          if (navigator.clipboard !== undefined && navigator.clipboard !== null) {
            navigator.clipboard.writeText(text)
            return true
          }
        } catch (err) { void err }
        return false
      }

      /**
       * 一次性提示：在 (x, y) 附近浮一小条字，几秒后自己消失。
       *
       * 存在的理由和工具栏那颗自诊断 chip 一样 —— 把**静默失败**变成看得见的
       * 信号。打开文件、写剪贴板这类动作一旦被宿主或系统拒绝，控制台里没人看。
       */
      var noticeEl = null
      var noticeTimer = 0
      function notice(x, y, text) {
        if (noticeEl !== null) { noticeEl.remove(); noticeEl = null }
        if (noticeTimer !== 0) { clearTimeout(noticeTimer); noticeTimer = 0 }
        var el = document.createElement('div')
        el.className = 'f2c-note-pop'
        el.textContent = text
        el.style.visibility = 'hidden'
        layer.appendChild(el)
        var w = el.offsetWidth
        var h = el.offsetHeight
        el.style.left = Math.min(Math.max(8, x), Math.max(8, window.innerWidth - w - 8)) + 'px'
        el.style.top = Math.min(Math.max(8, y), Math.max(8, window.innerHeight - h - 8)) + 'px'
        el.style.visibility = ''
        noticeEl = el
        noticeTimer = setTimeout(function () {
          noticeTimer = 0
          if (noticeEl === el) { el.remove(); noticeEl = null }
        }, 3200)
      }

      /**
       * 用系统的默认关联程序打开一个工作区路径（目录则是文件管理器）。
       *
       * 走的是 **DSH 自己的 Session Remote**：`remote.session.openWorkspacePath`，
       * 也就是右侧文档预览头部那个「打开」按钮用的同一个后端。宿主在 Windows 上
       * 把路径交给 Explorer 的关联解析 —— 和双击完全同一条路。因此这里**不需要**
       * 我们自带宿主路由（那还得重启应用才能挂上）。
       *
       * `request` 的形状照抄出厂实现：`{ path }` 打开、`{ path, action: 'reveal' }`
       * 定位、`{ path, application }` 指定程序；返回 `{ ok, value }` 信封。
       *
       * @returns Promise<true | string> —— true 表示宿主收下了，字符串是失败原因。
       */
      function openWithDefaultApp(path, action) {
        var api = null
        try {
          if (ctx.remote !== undefined && ctx.remote !== null) api = ctx.remote.session
        } catch (err) { void err }
        if (api === null || api === undefined) {
          try {
            var viaGet = typeof ctx.get === 'function' ? ctx.get('remote') : null
            if (viaGet !== null && viaGet !== undefined) api = viaGet.session
          } catch (err) { void err }
        }
        if (api === null || api === undefined || typeof api.openWorkspacePath !== 'function') {
          return Promise.resolve('当前连接没有 openWorkspacePath 接口')
        }
        var request = action === 'reveal' ? { path: path, action: 'reveal' } : { path: path }
        return Promise.resolve()
          .then(function () { return api.openWorkspacePath(request) })
          .then(function (result) {
            if (result !== null && result !== undefined && result.ok === false) {
              return '系统里没有能打开它的程序'
            }
            return true
          })
          .catch(function (err) { return String((err && err.message) || err) })
      }

      /** 相对工作区根的路径 —— 发给模型时比绝对路径省 token，也更好读。 */
      function relPathUnder(path, root) {
        if (typeof root !== 'string' || root === '') return null
        var a = String(path).replace(/\\/g, '/')
        var b = String(root).replace(/\\/g, '/').replace(/\/+$/, '')
        if (b === '' || a === b) return null
        if (a.slice(0, b.length + 1) !== b + '/') return null
        return a.slice(b.length + 1)
      }

      function openFileMenu(x, y, path, entry, root) {
        clearMenu()
        menu = document.createElement('div')
        menu.className = 'f2c-menu'
        var rel = relPathUnder(path, root)

        if (entry === 'directory') {
          menu.appendChild(menuItem('发送目录清单', function () {
            clearMenu()
            openDirCard(path, x, y)
          }))
        } else if (entry === 'file') {
          // 只发路径是**默认**动作：模型自己就能读文件，内联一份正文既费 token
          // 又容易过期。要正文的场合（想固定住当时的内容）留成第二项。
          menu.appendChild(menuItem('发送路径到聊天', function () {
            clearMenu()
            openPathCard(path, x, y)
          }))
          menu.appendChild(menuItem('发送内容到聊天', function () {
            clearMenu()
            openFileCard(path, x, y, '整份文件')
          }))
        }
        // 文件与目录都适用：文件交给它的关联程序（等同双击），目录交给文件管理器。
        menu.appendChild(menuItem('用默认程序打开', function () {
          clearMenu()
          openWithDefaultApp(path, 'open').then(function (r) {
            if (r !== true) notice(x, y, '打不开：' + r)
          })
        }))
        menu.appendChild(menuItem('复制完整路径', function () {
          clearMenu()
          copyToClipboard(path)
        }))
        if (rel !== null) {
          menu.appendChild(menuItem('复制相对路径', function () {
            clearMenu()
            copyToClipboard(rel)
          }))
        }
        menu.addEventListener('mousedown', function (e) { e.preventDefault() })
        menu.addEventListener('click', function (e) { e.stopPropagation() })
        // 先藏起来再量尺寸，免得在 (0,0) 闪一下。
        menu.style.visibility = 'hidden'
        layer.appendChild(menu)
        var w = menu.offsetWidth
        var h = menu.offsetHeight
        var left = x
        var top = y
        if (left + w > window.innerWidth - 8) left = Math.max(8, window.innerWidth - w - 8)
        if (top + h > window.innerHeight - 8) top = Math.max(8, y - h)
        menu.style.left = left + 'px'
        menu.style.top = top + 'px'
        menu.style.visibility = ''
      }

      function onContextMenu(e) {
        if (menu !== null) clearMenu()
        var node = nodeToElement(e.target)
        if (node === null || node.closest === undefined) return
        if (node.closest('[data-files-body]') === null) return
        var row = node.closest('[data-files-entry][data-files-path]')
        if (row === null) return
        var rawPath = row.getAttribute('data-files-path')
        if (typeof rawPath !== 'string' || rawPath === '') return
        var path = tidyPath(rawPath)
        var entry = row.getAttribute('data-files-entry')
        // `data-files-root` 挂在树的根节点上，是本次会话的工作区绝对路径。
        var rootEl = row.closest('[data-files-root]')
        var root = rootEl === null ? '' : tidyPath(rootEl.getAttribute('data-files-root') || '')
        e.preventDefault()
        e.stopPropagation()
        openFileMenu(e.clientX, e.clientY, path, entry, root)
      }

      /**
       * 右键「发送到聊天」：先按文件建快照、把卡片画出来（正文位显示「读取中…」），
       * 读完再画一次。
       *
       * 分两趟是有理由的：文本文件本地读很快，但 docx/xlsx/pdf 会立刻被
       * `workspace-file/not-text` 拒掉 —— 先画出来，用户才不会以为卡住了。
       * 重画时草稿靠 noteDraft 留住。
       */
      /**
       * 三种来源共用同一条「先占位、再填内容」的路：卡片立刻出现（正文位显示
       * 「读取中…」），内容到了再画一遍。草稿靠 noteDraft 跨重画保留。
       *
       * 分两趟是有理由的：文本文件本地读很快，但 docx/xlsx 要走抽包、
       * 目录要走列举 —— 先画出来，用户才不会以为卡住了。
       */
      function openLoadedCard(label, path, x, y, load, extra) {
        var base = {
          kind: 'file',
          path: path,
          label: label,
          text: null,
          failure: null,
          meta: null,
          pending: true,
          top: y, bottom: y, left: x, right: x
        }
        if (extra !== null && extra !== undefined) {
          for (var k in extra) {
            if (Object.prototype.hasOwnProperty.call(extra, k)) base[k] = extra[k]
          }
        }
        clearAll()
        snap = base
        openCard()
        load(function (text, failure, meta) {
          // 这期间用户可能已经取消、或者开了别的来源 —— 认人不认值。
          if (snap !== base) return
          snap.text = text
          snap.failure = failure
          snap.meta = meta === undefined ? null : meta
          snap.pending = false
          openCard()
        })
      }

      function openFileCard(path, x, y, label) {
        openLoadedCard(label, path, x, y, function (done) {
          readFileForChat(ctx, readSessionId(), path, done)
        })
      }

      /**
       * 不发正文，只把路径本身交给模型 —— 让它自己按需去读。
       *
       * 这是文件引用的**默认**动作：内联正文只在「要固定住当时那份内容」时才有
       * 意义，平时既费 token 又和磁盘上的实际内容可能不一致。`pathOnly` 让
       * `composeTextBlock` 走一条不输出任何正文的支路（而不是把「读不到正文」
       * 当成失败来报）。
       */
      function openPathCard(path, x, y) {
        openLoadedCard('仅路径', path, x, y, function (done) {
          done(null, null, { lines: null, bytes: null, eof: true, from: 'path' })
        }, { pathOnly: true })
      }

      function openDirCard(path, x, y) {
        openLoadedCard('目录清单', path, x, y, function (done) {
          listDirForChat(ctx, readSessionId(), path, done)
        })
      }

      document.addEventListener('contextmenu', onContextMenu)

      function openCard() {
        clearBar()
        clearCard()
        if (snap === null) return

        card = document.createElement('div')
        card.className = snap.kind === 'excel' ? 'f2c-card is-table' : 'f2c-card'

        // Excel：先把工作簿快照取出来。一来要读数值，二来要补工作表名 ——
        // 名称框多数时候只有 "B8" 而没有 "Sheet1!" 前缀，DOM 里那个高亮的
        // 工作表标签又读不到，所以表名只能从快照里拿。放在建 DOM 之前，
        // 因为下面的来源行就要用它。
        var wb = null
        if (snap.kind === 'excel') {
          wb = workbookOf(snap.host)
          if (snap.sheet === null || snap.sheet === undefined || snap.sheet === '') {
            snap.sheet = sheetNameAt(wb, activeSheetIndex(wb))
          }
        }

        var head = document.createElement('div')
        head.className = 'f2c-card-head'
        var src = document.createElement('span')
        src.className = 'f2c-src'
        var label = snap.path !== null && snap.path !== undefined && snap.path !== ''
          ? snap.path : '（未识别来源文件）'
        if (snap.kind === 'excel') {
          var sheetLabel = snap.sheet !== null && snap.sheet !== undefined && snap.sheet !== ''
            ? snap.sheet : '?'
          label += '：工作表 ' + sheetLabel + ' · ' + snap.address
        } else if (Array.isArray(snap.lineRange)) {
          label += snap.lineRange[0] === snap.lineRange[1]
            ? '：' + snap.lineRange[0] + ' 行'
            : '：' + snap.lineRange[0] + '-' + snap.lineRange[1] + ' 行'
        } else if (Array.isArray(snap.pageRange)) {
          label += snap.pageRange[0] === snap.pageRange[1]
            ? '：第 ' + snap.pageRange[0] + ' 页'
            : '：第 ' + snap.pageRange[0] + '-' + snap.pageRange[1] + ' 页'
        } else if (snap.kind === 'file') {
          label += '：' + (typeof snap.label === 'string' && snap.label !== '' ? snap.label : '整份文件')
        }
        src.textContent = label
        src.title = label
        var x = document.createElement('button')
        x.type = 'button'
        x.className = 'f2c-x'
        x.textContent = '\u00D7'
        x.title = '取消 (Esc)'
        x.addEventListener('mousedown', function (e) { e.preventDefault() })
        x.addEventListener('click', function (e) { e.preventDefault(); clearAll() })
        head.appendChild(src)
        head.appendChild(x)

        var quote = document.createElement('div')
        if (snap.kind === 'excel') {
          quote.className = 'f2c-grid'
          // 取值只在打开卡片时做一次；失败就退化成「只发坐标」，不阻塞发送。
          var got = readRange(wb, snap.sheet, snap.range)
          if (got.error !== undefined) {
            snap.table = null
            var warn = document.createElement('div')
            warn.className = 'f2c-note2'
            warn.textContent = got.error
            quote.appendChild(warn)
          } else {
            snap.table = got.table
            quote.appendChild(buildGrid(got.table))
          }
        } else if (snap.kind === 'file' && (snap.text === null || snap.text === undefined)) {
          quote.className = 'f2c-quote'
          quote.textContent = snap.pending === true
            ? '读取中…'
            : '读不到纯文本（' + String(snap.failure === null || snap.failure === undefined ? '未知原因' : snap.failure)
              + '），将只发送引用来源行，正文留给模型自己去读。'
        } else if (snap.kind === 'file' && String(snap.text) === '') {
          // 宿主说读成功、正文却是空的 —— 这种事不该悄悄发一条空引用出去，
          // 所以把矛盾直接摆到卡片上（带上宿主报的文件大小，便于判断是哪一边）。
          quote.className = 'f2c-quote'
          quote.textContent = '宿主返回了空正文'
            + (snap.meta !== null && snap.meta !== undefined && snap.meta.bytes !== null
              ? '（宿主报文件 ' + snap.meta.bytes + ' 字节）' : '')
            + '。仍可发送，正文留给模型自己去读。'
        } else {
          quote.className = 'f2c-quote'
          quote.textContent = truncate(String(snap.text), 600)
        }

        var note = document.createElement('textarea')
        note.className = 'f2c-note'
        note.placeholder = snap.kind === 'file'
          ? '写点什么…（可留空，只发送文件）'
          : '写点什么…（可留空，只发送选区）'
        // 卡片会因为「文件读完」而重画一次，草稿要跨重画留住。
        note.value = noteDraft
        note.addEventListener('input', function () { noteDraft = note.value })

        var foot = document.createElement('div')
        foot.className = 'f2c-foot'
        var hint = document.createElement('span')
        hint.className = 'f2c-hint'
        hint.textContent = 'Enter 发送 · Shift+Enter 换行'
        var send = document.createElement('button')
        send.type = 'button'
        send.className = 'f2c-send'
        send.textContent = '发送到聊天'
        foot.appendChild(hint)
        foot.appendChild(send)

        card.appendChild(head)
        card.appendChild(quote)
        card.appendChild(note)
        card.appendChild(foot)

        // 点卡片本体不丢选区；textarea 需要真正获得焦点，故放行
        card.addEventListener('mousedown', function (e) {
          if (e.target !== note) e.preventDefault()
        })
        card.addEventListener('click', function (e) { e.stopPropagation() })

        // 一张卡片只允许发送一次。发送成功后卡片还会停 900ms 显示「已发送 ✓」，
        // 这段时间里再敲一次 Enter 就会**再插一份**引用+批注（`send.disabled`
        // 只挡按钮，挡不住 textarea 的 Enter 分支）。失败时解锁，允许重试。
        var sent = false
        function doSend() {
          if (sent) return
          if (snap === null) return
          sent = true
          send.disabled = true
          try {
            var noteText = String(note.value === null || note.value === undefined ? '' : note.value).trim()
            // 批注不进序列化正文：它作为普通草稿文本补在胶囊后面，这样输入框里
            // 也看得见（用户反馈：批注在输入框里被藏起来了）。
            var block = composeBlock(snap, '')
            var r = injectText(ctx, readSessionId, block, chipLabel(snap), noteText)
            send.textContent = '已发送 ✓'
            hint.textContent = '经 ' + r.via
              + (noteText === '' ? '' : (r.noteInline === true ? ' · 批注已随文本补入' : ' · 批注并入正文'))
            dbg('已注入引用块', r)
            clearTimeout(flashTimer)
            flashTimer = setTimeout(clearAll, 900)
          } catch (err) {
            sent = false
            send.disabled = false
            var msg = err !== null && err !== undefined && err.message !== undefined
              ? err.message : String(err)
            hint.textContent = '失败：' + msg
            console.warn('[file-to-chat] 注入失败：', err)
          }
        }

        send.addEventListener('mousedown', function (e) { e.preventDefault() })
        send.addEventListener('click', function (e) { e.preventDefault(); doSend() })
        note.addEventListener('keydown', function (e) {
          // 卡片是挂在 body 上的浮层，按键会一路冒泡到 document。虽然正常情况
          // 下输入焦点在我们的 textarea 里、composer 收不到，但把冒泡掐断成本
          // 极低，能彻底排除「一次 Enter 同时被两边处理」这类耦合。
          e.stopPropagation()
          if (e.key === 'Enter' && e.shiftKey !== true) {
            e.preventDefault()
            doSend()
          } else if (e.key === 'Escape') {
            e.preventDefault()
            clearAll()
          }
        })

        layer.appendChild(card)
        // 量一次实际尺寸再定位：表格卡片的宽度由内容决定，写死 340 会偏。
        var box = { width: 340, height: 240 }
        try {
          var rect = card.getBoundingClientRect()
          if (rect.width > 0 && rect.height > 0) box = { width: rect.width, height: rect.height }
        } catch (_) { /* 用兜底尺寸 */ }
        var p = placeAbove(snap.top, snap.bottom, snap.left, snap.right, box.width, box.height)
        card.style.left = p.left + 'px'
        card.style.top = p.top + 'px'
        try { note.focus() } catch (_) { /* 忽略 */ }
      }

      // ---------------- 选区监听 ----------------
      var scheduled = false
      function schedule() {
        if (scheduled) return
        scheduled = true
        setTimeout(function () {
          scheduled = false
          onSelectionMaybeChanged()
        }, 10)
      }

      // 拖选时 `selectionchange` 会一路连发 —— 若每次都弹按钮，按钮就在拖拽
      // 途中反复闪现。约定：**指针按住期间一律不弹**，等松手（pointerup）后由
      // 上面那条 `mouseup -> schedule` 补弹一次，也就是「停止选择后再出现」。
      // 键盘选择（Shift+方向键）没有指针事件，仍走 10ms 去抖，保持跟手。
      // 兜底：pointercancel 与在窗口外松手也会解除，避免卡在「永不弹出」。
      var pointerSelecting = false

      function onSelectionMaybeChanged() {
        if (card !== null) return          // 卡片开着时不打扰
        if (pointerSelecting) return       // 还在拖选，等松手再弹
        // 刚在表格里按过：canvas 上没有原生选区，别让这次手势把它清掉
        if (snap !== null && snap.kind === 'excel' && Date.now() - excelAt < 250) return
        var next = null
        try { next = captureSelection(bridge.absolutePath) } catch (_) { next = null }
        if (next === null) {
          if (bar !== null) clearBar()
          snap = null
          return
        }
        snap = next
        showBar()
      }

      // Excel 预览是 canvas —— 单元格文字不在 DOM 里，window.getSelection()
      // 永远读不到，只能听指针释放，再去读名称框里的地址。
      var excelAt = 0
      function onPointerUp(e) {
        pointerSelecting = false          // 松手即视为选择结束（在预览外松手也算）
        if (card !== null) return
        // 点的是浮层自己（「批注」按钮）—— 绝不能当成「离开表格」处理。
        // 时序：pointerup 早于 click，若在这里 clearBar() 就把按钮的 DOM 节点
        // 删了，click 随之不再派发 —— 表现正是「按钮在，点了没反应」。
        if (bar !== null && (e.target === bar || bar.contains(e.target))) return
        // pointerup 和 mouseup 都挂了这个处理器：相隔极近的算同一次手势。
        if (snap !== null && snap.kind === 'excel' && Date.now() - excelAt < 60) return
        if (excelHostOf(e.target) === null) {
          if (snap !== null && snap.kind === 'excel') { clearBar(); snap = null }
          return
        }
        var next = null
        try {
          next = captureExcelSelection({ x: e.clientX, y: e.clientY }, bridge.absolutePath)
        } catch (err) {
          console.warn('[file-to-chat] 表格选区读取异常：', err)
          next = null
        }
        if (next === null) {
          // 命中网格却读不出地址 —— 最可能是名称框的选择器变了。留一条日志，
          // 省得又变成一次「什么都不发生」。
          dbg('命中表格但读不到选区', e.target)
          if (snap !== null && snap.kind === 'excel') { clearBar(); snap = null }
          return
        }
        dbg('表格选区', next.address, '| 工作表', next.sheet, '| 文件', next.path)
        excelAt = Date.now()
        snap = next
        showBar()
      }

      // 指针一按下就进入「选择中」，期间不弹按钮。
      function onPointerDown(e) {
        if (card !== null) return
        if (bar !== null && (e.target === bar || bar.contains(e.target))) return
        if (!isInDocumentPreview(nodeToElement(e.target))) return
        pointerSelecting = true
        if (bar !== null) clearBar()       // 别让上一轮的按钮挂在拖选途中
      }

      document.addEventListener('selectionchange', schedule)
      document.addEventListener('mouseup', schedule)
      document.addEventListener('keyup', schedule)
      document.addEventListener('pointerdown', onPointerDown)
      document.addEventListener('pointerup', onPointerUp)
      // 指针被系统取消（例如拖出窗口后松手）也要解除「选择中」，否则按钮永不出现
      document.addEventListener('pointercancel', onPointerUp)
      // mouseup 再兜一次：某些拖拽实现会 setPointerCapture，pointerup 的
      // target 可能被重定向；mouseup 是最后的保险。
      document.addEventListener('mouseup', onPointerUp)

      function onScrollOrResize() {
        if (card !== null) { clearAll(); return }
        if (bar !== null) positionBar()
      }
      window.addEventListener('scroll', onScrollOrResize, true)
      window.addEventListener('resize', onScrollOrResize)

      function onKeyDown(e) {
        if (e.key === 'Escape') clearAll()
      }
      document.addEventListener('keydown', onKeyDown)

      function onDocMouseDown(e) {
        // 菜单先处理：点在菜单里放行；点在别处就收起。右键也会先走 mousedown，
        // 紧接着 contextmenu 再开一个新的 —— 顺序正好。
        if (menu !== null) {
          if (e.target === menu || menu.contains(e.target)) return
          clearMenu()
        }
        if (card === null && bar === null) return
        var t = e.target
        if (t === bar || (bar !== null && bar.contains(t))) return
        if (t === card || (card !== null && card.contains(t))) return
        clearAll()
      }
      document.addEventListener('mousedown', onDocMouseDown)

      // 命令式浮层的卸载
      return function dispose() {
        document.removeEventListener('selectionchange', schedule)
        document.removeEventListener('mouseup', schedule)
        document.removeEventListener('keyup', schedule)
        document.removeEventListener('pointerdown', onPointerDown)
        document.removeEventListener('pointerup', onPointerUp)
        document.removeEventListener('pointercancel', onPointerUp)
        document.removeEventListener('mouseup', onPointerUp)
        document.removeEventListener('keydown', onKeyDown)
        document.removeEventListener('mousedown', onDocMouseDown)
        document.removeEventListener('contextmenu', onContextMenu)
        window.removeEventListener('scroll', onScrollOrResize, true)
        window.removeEventListener('resize', onScrollOrResize)
        clearAll()
        if (layer.parentNode !== null) layer.parentNode.removeChild(layer)
      }
    }

    // ================================================================
    // 八、工具栏提示组件（同时把 inputActions 桥接给命令式浮层）
    // ================================================================
    function makeHintButton(seat) {
      // 只有「文档预览」那个座位发下来的 absolutePath 才是被预览文件的路径；
      // 文件树座位发的是当前浏览的目录 —— 拿它当来源会把目录写进引用块。
      var isDocSeat = seat === 'sidebar.right.tab.document.actions'
      return function FileToChatHint() {
        // 这个组件挂在右侧工具栏上，只要对应标签页开着就是常驻挂载的 ——
        // 桥上的值随渲染保持新鲜。
        var props = arguments.length > 0 ? arguments[0] : {}
        bridge.inputActions = props.inputActions !== undefined ? props.inputActions : null
        bridge.sessionId = props.sessionId !== undefined ? props.sessionId : null
        bridge.absolutePath = isDocSeat && props.absolutePath !== undefined ? props.absolutePath : null

        var draft = ''
        try {
          if (typeof props.useInput === 'function') {
            draft = props.useInput(function (s) {
              return s !== undefined && s !== null && typeof s.draft === 'string' ? s.draft : ''
            }) || ''
          }
        } catch (_) { draft = '' }
        bridge.draft = draft

        // 封膜：**只开不收**。
        // 收膜若按草稿内容判断（草稿里没有胶囊占位符就收），就会在用户还在
        // 写批注的时候把膜撤掉 —— 草稿文本里到底带不带占位符属于内核实现细节，
        // 一旦不带，4 秒宽限一到膜就落了，消息落地时又退回「先全文、再折叠」。
        // 真正的收膜时机是「刚到达的那一行已经露面」，见 holdNewRows / markSeen。
        if (CHIP_MARK_RE.test(draft)) sealOn()

        return React.createElement('span', {
          className: 'f2c-chip',
          title: '文本预览：划选文本，选区上方出现「批注」按钮。\n'
            + '表格预览：在表格里拖选区域，随后点「批注」即可按区域取值发送。\n'
            + 'Word/PPT/PDF：在 PDF 文字层里划选，按「第 N 页」定位（这类预览没有源码行号）。',
        }, React.createElement('span', {
          className: 'f2c-ic',
          dangerouslySetInnerHTML: { __html: ICON_ANNOTATE },
        }), '批注')
      }
    }

    // ================================================================
    // 九、插件主体
    // ================================================================
    function makeSessionIdReader(ctx) {
      return function readCurrentSessionId() {
        try {
          var face = ctx.get('uiSession')
          var adapter = face !== undefined && face !== null ? face.adapter : undefined
          var source = adapter !== undefined && adapter !== null ? adapter.current : undefined
          if (source !== undefined && source !== null && typeof source.getSnapshot === 'function') {
            var binding = source.getSnapshot()
            if (binding !== undefined && binding !== null && binding.key !== undefined) return binding.key
          }
        } catch (_) { /* 旧版内核无 uiSession */ }
        try {
          if (typeof localStorage !== 'undefined') {
            var raw = localStorage.getItem('dsh.sessions.current')
            if (raw !== null) {
              var parsed = JSON.parse(raw)
              if (parsed !== null && typeof parsed === 'object'
                && typeof parsed.sessionId === 'string' && parsed.sessionId !== '') return parsed.sessionId
            }
          }
        } catch (_) { /* 落到下一级 */ }
        try {
          var snap = ctx.sessions.list.getSnapshot()
          if (snap !== undefined && snap !== null) {
            if (typeof snap.current === 'string' && snap.current !== '') return snap.current
            var byId = snap.byId
            if (byId !== undefined && byId !== null && typeof byId === 'object') {
              var ids = Array.isArray(snap.ids) ? snap.ids : Object.keys(byId)
              for (var i = 0; i < ids.length; i++) {
                var row = byId[ids[i]]
                var retained = row !== undefined && row !== null ? row.retainedBy : null
                if (retained !== null && typeof retained === 'object'
                  && typeof retained.number === 'number') return ids[i]
                if (retained !== null && typeof retained === 'object'
                  && typeof retained.mainView === 'number' && retained.mainView > 0) return ids[i]
              }
            }
          }
        } catch (_) { /* 落到下一级 */ }
        return undefined
      }
    }

    // ================================================================
    // 九、锚点自检
    // ================================================================
    // 这个插件整套都挂在 DSH 的内部约定上：`data-chat-flow-kind`（折叠）、
    // `data-document-preview` / `data-textpreview-line` / `data-pdf-page`（选区）、
    // `.fortune-name-box` + `__reactFiber$`（Excel）、`data-files-*`（右键菜单）。
    // DSH 一升级这些就可能改名，而失效的表现是「功能悄悄没了」，不是报错。
    //
    // 所以做一份锚点普查：某个锚点**曾经**命中、后来找不到了，就立刻告警一次。
    // 只对「命中过」的锚点告警 —— 否则「这个界面本来就没打开」会全是误报。
    var F2C_ANCHORS = [
      ['会话行', '[data-chat-flow-kind]', '折叠已发送的引用块'],
      ['输入框', '[data-lexical-editor="true"]', '把批注补到胶囊后面'],
      ['文档预览', '[data-document-preview]', '划词批注'],
      ['文本预览行', '[data-textpreview-line]', '行号定位'],
      ['PDF 页', '[data-pdf-page]', '页码定位'],
      ['Excel 名称框', '.fortune-name-box', 'Excel 选区'],
      ['文件树', '[data-files-body]', '文件区右键菜单'],
      ['文件行', '[data-files-entry]', '文件区右键菜单'],
      ['引用胶囊', '[data-f2c-chip]', '折叠后的胶囊']
    ]

    function anchorCensus() {
      var present = []
      var missing = []
      for (var i = 0; i < F2C_ANCHORS.length; i++) {
        var ok = false
        try { ok = document.querySelector(F2C_ANCHORS[i][1]) !== null } catch (err) { ok = false }
        if (ok) present.push(F2C_ANCHORS[i][0])
        else missing.push(i)
      }
      return { present: present, missing: missing }
    }

    function startAnchorWatch() {
      var everSeen = {}
      var lastLine = null
      function tick() {
        var c = anchorCensus()
        var i
        for (i = 0; i < c.present.length; i++) everSeen[c.present[i]] = true
        for (i = 0; i < c.missing.length; i++) {
          var def = F2C_ANCHORS[c.missing[i]]
          var name = def[0]
          if (everSeen[name] !== true || everSeen[name + '\u0000warned'] === true) continue
          everSeen[name + '\u0000warned'] = true
          console.warn('[file-to-chat] 锚点失效：「' + name + '」曾经命中，现在找不到了。'
            + '受影响的功能是「' + def[2] + '」，选择器 `' + def[1] + '` —— '
            + '多半是 DSH 改了 DOM 约定，这个插件需要跟着更新。')
        }
        var line = c.present.join(' · ')
        if (line !== lastLine) {
          lastLine = line
          dbg('锚点：' + (line === '' ? '（当前一个都没命中）' : line))
        }
      }
      tick()
      var timer = setInterval(tick, 5000)
      return function () { clearInterval(timer) }
    }

    function apply(ctx) {
      ensureStyle()
      var readCurrentSessionId = makeSessionIdReader(ctx)
      // 尽早注册引用 source（composer 的 chip 编解码）。提供 inputTriggers 的
      // 插件激活顺序不保证，因此 wireQuoteSource 内部有 ctx.inject + 定时重试。
      wireQuoteSource(ctx)

      // 锚点自检：把「DSH 改约定导致功能静默失效」变成控制台里的一条告警。
      ctx.effect(function () {
        return startAnchorWatch()
      }, 'file-to-chat: anchor watch')

      // 划词浮层是全局单例：挂在 document.body 上，不随标签页切换卸载。
      ctx.effect(function () {
        return createOverlay(ctx, readCurrentSessionId)
      }, 'file-to-chat: selection overlay')

      // 工具栏提示按钮 —— 同时承担把 inputActions 桥接到浮层的职责。
      var SEATS = [
        'sidebar.right.tab.document.actions',
        'sidebar.right.tab.files.actions',
      ]
      for (var i = 0; i < SEATS.length; i++) {
        (function (seat) {
          var SeatHint = makeHintButton(seat)
          ctx.effect(function () {
            return ctx.slots.inject(seat, function () {
              return ctx.slots.register({
                name: seat,
                id: 'dsh-file-to-chat.hint',
                order: 50,
              }, SeatHint)
            })
          }, 'file-to-chat: hint @ ' + seat)
        })(SEATS[i])
      }

      // 已发送引用块的展示层折叠。挂 MutationObserver + rAF 合并：新消息挂载
      // 或 React 重渲染之后重扫一遍。rAF 合并是必须的 —— 流式回答期间
      // childList 变动极其频繁，不合并会变成每个 token 扫一次整棵聊天区。
      ctx.effect(function () {
        var raf = 0
        var micro = false
        var scanOnce = function () {
          try {
            foldSentQuotes()
          } catch (err) {
            console.warn('[file-to-chat] 折叠扫描失败：', err)
          }
        }
        var scan = function () { raf = 0; scanOnce() }
        // 微任务版的同一次全量扫描。它和 foldIn 是**互相独立的两道保险**：
        //   · foldIn  是快路 —— 直接命中刚变更的那一行，代价最低；
        //   · scanMic 是慢路 —— 不看变更类型，整表重扫。
        // 两者都跑在「本次 DOM 变更之后、浏览器绘制之前」这个窗口里，所以只要
        // 任意一条生效，用户就看不到「全文 → 胶囊」的中间形态。rAF 那一路保留
        // 只为兜住 React 之外（例如样式驱动）的变更。
        // 终止性：scanMic 只会改动那些尚未折叠的行；折完之后 FOLDED_ROWS 命中，
        // 不再产生 DOM 变更，于是不会自我触发。
        var scanMic = function () {
          micro = false
          scanOnce()
        }
        var schedule = function () {
          if (raf === 0) raf = requestAnimationFrame(scan)
          if (!micro) { micro = true; queueMicrotask(scanMic) }
        }
        var mo = new MutationObserver(function (records) {
          // 先把这一批变更**同步**折掉（见 foldIn 的说明：抢在绘制之前，这样
          // 「发出去先弹一下全文」的闪动不会再发生），再排一次全量扫描兜底 ——
          // React 重渲染可能把已有行的 chip 抹掉，那属于「没有新增节点」的变更，
          // 只有全量扫描才看得到。
          for (var i = 0; i < records.length; i++) {
            var added = records[i].addedNodes
            for (var j = 0; j < added.length; j++) foldIn(added[j])
            // 变更**发生的位置**所在的行同样要折。React 重渲染一条消息时常常
            // 只替换行内的文本节点（不新增整行），光看 addedNodes 会漏掉。
            foldRowNear(records[i].target)
          }
          schedule()
        })
        // characterData 也观察：React 更新文本时若直接改 nodeValue（而不是换节点），
        // 那就只有 characterData 记录能捕获到。代价是流式回答期间每来一个 token
        // 多一次 closest 上溯 —— 常数极小，换来的是闪动路径被彻底堵死。
        mo.observe(document.body, { childList: true, subtree: true, characterData: true })
        schedule()
        return function () {
          if (raf !== 0) cancelAnimationFrame(raf)
          mo.disconnect()
        }
      }, 'file-to-chat: fold sent quotes')

      dbg('已装载：划词批注浮层 + 已发送引用折叠')
    }

    exports.name = 'dsh-file-to-chat'
    exports.inject = ['slots', 'sessions', 'conversation', 'remote', 'remote.workspaceFiles', 'remote.session']
    exports.apply = apply

    return module.exports
  },
})
