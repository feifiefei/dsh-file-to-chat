# dsh-file-to-chat

**把右侧栏里的内容，带着你的批注，直接送进当前对话。**

在 DeepSeek Harness Web 的右侧栏里划选一段文字、框选一块 Excel 区域，或者在文件树上右键一个文件 —— 选区（或整份文件）连同你写的批注一起进入输入框，发出去之后自动收成一枚小胶囊，聊天记录不再被大段引用撑爆。

> Annotate a selection or a whole file in the DSH right-sidebar and send it into the current chat with your note. The quote reaches the model in full; the transcript only shows a small capsule.

```
输入框：  [📄 引用 · 套餐对比.xlsx · 结论速览 B8]  这条再确认下

发出去：  > 📄 引用来源：<绝对路径>（工作表 结论速览 · 区域 B8）
          > 速率上限：官方账户级并发 2,500，可免费扩容
          
          批注： 这条再确认下

聊天里：  [📄 套餐对比.xlsx · 结论速览 B8]   ← 点一下可以展开全文
```

---

## 功能

### 1. 文档预览里划词批注

支持右侧栏文档预览的**全部三种渲染器**：

| 预览类型 | 怎么选 | 引用块里定位到 |
|---|---|---|
| 文本 / 代码 / Markdown | 直接划选文字 | `（第 N 行）` 或 `（第 N-M 行）` |
| Excel / CSV / TSV | 在表格里拖选区域 | `（工作表 <名> · 区域 B3:D7）`，并把区域内的值取成 Markdown 表格 |
| Word / PPT / PDF | 在文字层里划选 | `（第 N 页）` 或 `（第 N-M 页）` |

选区上方会浮出一枚「批注」按钮（选区太靠页面顶部时自动落到下方），点开写批注，`Enter` 发送，`Shift+Enter` 换行，`Esc` 取消。**拖选过程中不会弹**，松手才出现。

### 2. 文件树右键菜单

在右侧栏「文件」标签页的文件或文件夹上右键：

| 菜单项 | 作用 |
|---|---|
| 发送到聊天 | 读正文并开卡片（文本直接内联，`.docx` / `.xlsx` 在浏览器里解包抽正文，其余按字节回退） |
| 只发送路径 | 不读正文，只发路径与「未读取正文」说明 —— 大文件/二进制文件的安全选项 |
| 发送目录清单 | 文件夹专用，列出目录内容 |
| **用默认程序打开** | 交给系统关联程序（等同双击），文件夹则交给文件管理器 |
| 复制完整路径 / 复制相对路径 | 复制到剪贴板 |

### 3. 已发送引用块的折叠

发出去的消息里那段 `> 📄 引用来源：…` 会自动折成一枚胶囊：**模型收到的一字不减，聊天窗口只显示标题**。点胶囊展开，移开鼠标自动合上。

---

## 安装

需要 **DSH Web 0.2.x**（开发与验证均在 `0.2.0-rc.2` 上完成）。

### 方式一：从 Git 仓库装（推荐）

```bash
dsh plugin --profile <你的 profile> add https://gitee.com/mofeo/dsh-file-to-chat.git
```

`<你的 profile>` 通常是 `desktop`（桌面端）或 `web`。装完在 **设置 → 插件** 里确认它是启用状态，然后**硬刷新一次浏览器**（`Ctrl+Shift+R`）。

### 方式二：手动放进 profile

```bash
git clone https://gitee.com/mofeo/dsh-file-to-chat.git
# 把目录链接/复制到：~/.dsh/profiles/<profile>/node_modules/dsh-file-to-chat
```

再把 `dsh-file-to-chat` 追加到 `~/.dsh/profiles/<profile>/package.json` 的 `dsh.profile.bundles` 数组里，重启应用。

### 卸载

```bash
dsh plugin --profile <你的 profile> remove dsh-file-to-chat
```

> **注意**：有会话正在跑的时候卸载会被拒绝（`refused while agents are running`），先停掉再卸。

---

## 用法

1. 在右侧栏打开「文件」标签页，点开任意文件进入预览。
2. 划选（文本/Word/PDF）或在表格里拖选区域（Excel）。
3. 选区上方出现「📝 批注」→ 点开 → 写批注（可留空）。
4. `Enter` 或点「发送到聊天」。
5. 输入框里出现一枚胶囊，**按你自己的节奏按发送键** —— 插件只负责填草稿，不替你发送。

---

## 已知限制

- **只有右侧栏能触发。** 右侧栏未激活的标签页内容不挂载，所以预览没打开时工具栏上不会有入口。
- **Word/PPT/PDF 没有行号**，只能定位到页码 —— 这类文件走的是「Office → PDF → PDF.js」渲染链，页内没有源码行这个概念。
- **Excel 只取当前工作表**，且内联上限 200 行；`.xlsx` 公式取缓存值而非重算。
- **读不到正文的文件不会假装成功**：`.docx`/`.xlsx` 靠浏览器解包抽取，抽不出来时卡片会明说原因，只发引用来源行，让模型自己去读文件。
- **依赖若干 DSH 内部 DOM 约定**（见下节）。DSH 升级后如果这些锚点变了，插件内置的锚点自检会在控制台 `warn` 一次，而不是静默失效。
- 折叠与 Excel 取值走的是**展示层后处理**，不是内核的官方扩展点。功能正常，但属于「DSH 改版时最先需要复查」的部分。

---

## 实现说明（给维护者和审核者）

这个插件**没有宿主路由、没有额外的 Node 依赖**，宿主半边（`lib/index.js`）是空的 —— 它的存在只是为了在 Loader 里占一行。全部逻辑在 `client.js`（浏览器半边，手写 CJS + `window.__ModuleLoader__` 包装，只有 `require('react')` 一个依赖）。

用到的官方扩展点：

- `sidebar.right.tab.files.actions` / `sidebar.right.tab.document.actions` —— 工具栏入口
- `inputTriggers.registerSource` —— 输入框胶囊的注册与提交时的序列化编解码
- `conversation.input.for(scope)` → `insertReference(span, ref)` —— 把胶囊插进 Lexical 草稿
- `remote.workspaceFiles.read` / `readBytes` —— 读工作区文件
- `remote.session.openWorkspacePath` —— 「用默认程序打开」，直接复用 DSH 自带的后端

依赖的 DOM 锚点（会自检，失效时 `console.warn`）：

`[data-document-preview]`、`[data-textpreview-body]`、`[data-textpreview-line]`、`[data-textpreview-path]`、`[data-code-preview]`、`[data-document-markdown]`、`[data-pdf-page]`、`.fortune-name-box`、`[data-files-body]`、`[data-files-entry]`、`[data-files-path]`、`[data-files-root]`、`[data-chat-flow-kind]`、`[data-submission-echo]`

两个值得记下来的实现细节：

1. **Excel 选区不能靠 `window.getSelection()`。** 单元格是画在 canvas 上的，文字不在 DOM 里。但 FortuneSheet 会把当前选区写进名称框 `.fortune-name-box` —— 读它就够了。另外它的名称框在「选区焦点格正好是合并单元格左上角」时会**故意塌缩成单个格子**，所以真实范围改从 React 工作簿快照的 `luckysheet_select_save` 里读，名称框只作兜底。
2. **`setDraft` 会抹掉胶囊。** 它会先 `replace(REFERENCE_PLACEHOLDER_RE, '')` 再重建整份草稿，而胶囊正是用这些占位符表示的。所以补文字一律走 `insertText` / `insertReference`，绝不整份重写。

### 调试开关

默认控制台是干净的。需要排查时，在浏览器控制台执行一次：

```js
localStorage.setItem('dsh-file-to-chat:debug', '1')   // 刷新页面后打开细节日志
localStorage.removeItem('dsh-file-to-chat:debug')     // 关回去
```

---

## 开发

源码即产物：`client.js` 是手写的，没有构建步骤。改完保存即可 —— DSH 的 `dsh-client-hmr` 每 500 ms 轮询一次客户端 bundle，会**自动热重载，不用刷新页面**。（宿主半边 `lib/index.js` 不热重载，改了要重启应用；本插件正常情况下不需要动它。）

```
dsh-file-to-chat/
├── package.json         dsh.bundle / dsh.client 声明
├── cordis.patch.yml     往 Loader 里插一行
├── lib/index.js         宿主半边（空实现，只为占位）
└── client.js            浏览器半边 —— 全部功能
```

---

## 许可

[MIT](LICENSE)
