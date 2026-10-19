# Changelog

本项目遵循 [语义化版本](https://semver.org/lang/zh-CN/)。日期为 `YYYY-MM-DD`。

## [1.0.1] — 2026-10-19

适配 **DSH Web `0.2.1-rc.2`** 并做了一轮自检修复。

### 修复

- **防静默数据丢失**：折叠已发送引用块时新增「换行保护」。DSH 0.2.1 给用户正文的展示层加了一套由 `facts` 驱动的空白规则（`collapseWhitespaceRule` / `trimWhitespaceRule`）；聊天气泡目前传的是 `EMPTY_TEXT_FACTS = {}`，这两条规则没有生效、换行照旧保留。但一旦内核改了这行配置，换行会在 DOM 里塌成空格、整条消息变成一行，而「该行以 `> 📄 引用来源` 开头」依然成立 —— 旧逻辑就会把**用户整条消息**折进一枚胶囊、正文被藏起来。现在要求目标行「含引用行**且**含换行」，宁可退回「不折、显示全文」。
- **路径以宿主为准**：划词批注的来源路径改为**优先取 `document.actions` 座位发下来的 `absolutePath`**，`[data-textpreview-path]` 标签只作兜底。标签是给眼睛看的展示路径（可能被缩写），座位给的才是宿主认定的绝对路径 —— 这份引用块最终是给模型读的。

### 新增

- **能力普查**：锚点自检从「DOM 锚点」扩展到「cordis 服务能力」。启动 5 秒后一次性核对 8 项依赖（`sessions` / `slots` / `conversation` / `conversation.input.for` / `uiSession` / `remote` / `remote.workspaceFiles.read` / `remote.session.openWorkspacePath`），缺失项在控制台 `warn` 一次并写明受影响的功能。只查同步可判的项；激活顺序不保证的服务不查，避免误报。

### 适配性核对（0.2.0-rc.2 → 0.2.1-rc.2，无需改动）

逐段对照两个版本的出厂源码，以下契约**逐字节一致**：

- `inputTriggers.registerSource` 契约、重名报错文案、返回 disposer。
- `insertReference(span, ref, tail)` —— 连尾随空格规则都一样。
- `$createReferenceChipNode` / `$replaceDetectSpanWithNodes` / `insertAsyncText` 与占位符正则 `REFERENCE_PLACEHOLDER_RE`。
- `serializeReference` / `codec.serialize` / `conversation.input.for` / `caretSpan` / `slash/input-insert-reference` 事件通道。
- `remote.session.openWorkspacePath`、`remote.workspaceFiles.read` 的分页与错误码（`not-text` / `too-large` …）。
- 输入域与消息气泡的 `white-space: pre-wrap`。
- 23 个依赖的 DOM 锚点全部健在（`data-document-preview`、`data-textpreview-*`、`data-pdf-page`、`.fortune-name-box`、`data-files-*`、`data-chat-flow-kind` …）。
- 座位 `sidebar.right.tab.document.actions` / `files.actions` 的 `standardProps`（新增 `useAccess` 系列，`inputActions` 与 `sessionId` 仍在）。

## [1.0.0] — 2026-10-15

首个公开版本。开发与验证均在 **DSH Web `0.2.0-rc.2`**（profile `desktop`）上完成。

### 文档预览划词批注

- 选区上方浮出「批注」按钮，拖选期间不弹、松手才出现；选区太靠顶部时自动回落到下方。
- 批注卡片：显示来源路径与位置、选区预览（上限 600 字符）、批注输入框；`Enter` 发送、`Shift+Enter` 换行、`Esc` 取消。
- 覆盖三种渲染器：
  - **文本 / 代码 / Markdown** —— 读 `[data-textpreview-line]` 换算真实源码行号，引用块写作 `（第 N 行）` / `（第 N-M 行）`。
  - **Excel / CSV / TSV** —— 单元格是 canvas 绘制的，无法用 `getSelection()`；改从 FortuneSheet 的 React 工作簿快照读真实选区（`luckysheet_select_save`），名称框 `.fortune-name-box` 只作兜底。合并单元格起拖时名称框会故意塌缩成单个格子，走快照即可拿到完整范围。区域内的值取成 Markdown 表格或单行值。
  - **Word / PPT / PDF** —— 走「Office → PDF → PDF.js」链路，有真实文本层可选；引用块定位到页码 `（第 N 页）`。
- 引用块正文超 40 行自动截断并注明；批注以「空行 + `批注：`」跟在引用块之后。

### 文件树右键菜单

- **发送到聊天** —— 读正文开卡片。文本直接内联（上限 200 行）；`.docx` / `.xlsx` 在浏览器里解包抽取正文（`DecompressionStream('deflate-raw')`，`.docx` 取 `word/document.xml`，`.xlsx` 取首张工作表的共享字符串与单元格值）；其余类型按字节回退做 UTF-8 文本探测。读不到就在卡片里写明原因，不假装成功。
- **只发送路径** —— 不读正文。
- **发送目录清单** —— 文件夹专用，上限 300 项。
- **用默认程序打开** —— 交给系统关联程序；目录则交给文件管理器。
- **复制完整路径 / 复制相对路径**。
- 菜单点在文件树以外不拦截系统右键菜单。

### 输入框胶囊与聊天折叠

- 通过 `inputTriggers.registerSource` 注册引用 source，发出时由 codec 序列化成完整引用块 —— **模型收到的一字不减**，输入框只显示一枚胶囊。
- 已发送消息里的引用块自动折成胶囊，点击展开、移开鼠标合上，展开态点击别处也会合上。
- 折叠在 `MutationObserver` 的微任务里同步完成（绘制之前），并配合「封膜 + 逐行按住」避免「先全文再折叠」的闪动。

### 稳健性

- **锚点自检**：9 个依赖的 DSH 内部 DOM 锚点每 5 秒巡检一次，曾经出现过又消失的锚点会在控制台 `warn` 一次，避免 DSH 升级后静默失效。
- **诊断开关**：`localStorage.setItem('dsh-file-to-chat:debug', '1')` 打开细节日志，默认控制台干净。
- Excel 选区读 React fiber 属内部约定，一旦失效会安静回退到名称框，不报错、不崩。

### 已知限制

- 只有右侧栏能触发（未激活的标签页内容不挂载）。
- Word / PPT / PDF 只能定位到页码，没有行号。
- Excel 只取当前工作表且上限 200 行；`.xlsx` 公式取缓存值。
- 折叠与 Excel 取值属展示层后处理，非内核官方扩展点。
