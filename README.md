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

引用块除了来源和正文，还会带一行**定位上下文**：代码带所在函数（`> 位置：function readRange()，第 681 行起`），Markdown 带标题路径（`> 位置：# 需求 › ## 表结构`）。模型不用自己翻文件就知道这段在哪儿。

### 2. Excel 选区自动带表头

框选数据区（`B3:F20`）时，插件会**多看选区上方一行**：如果那一行填得够满（至少两列有字，大标题那种只填 A 列的不算），就把它当作表头一并带上，并在表头开关里注明来源行。预览表格的表头即来自这一行，卡片上的勾选框可以关掉。

```
> 📄 引用来源：…\套餐对比.xlsx（工作表 结论速览 · 区域 B3:F20）
> （表头取自选区上方第 2 行，已一并带上）
> | 指标 | 2024 | 2025 |
> | --- | --- | --- |
> | 营业收入 | 100 | 120 |
```

### 3. 文件树右键菜单

在右侧栏「文件」标签页的文件或文件夹上右键：

| 菜单项 | 作用 |
|---|---|
| 发送路径到聊天 | **默认动作**：不读正文，只发路径 —— 模型自己就能读，省 token 也不会过期 |
| 发送内容到聊天 | 读正文并开卡片（文本直接内联，`.docx` / `.xlsx` 在浏览器里解包抽正文，其余按字节回退） |
| 发送剩余部分 | 上次只发了前 200 行时出现，从断点接着发（每份文件各自记住位置） |
| 发送目录清单 | 文件夹专用，列出目录内容 |
| **用默认程序打开** | 交给系统关联程序（等同双击），文件夹则交给文件管理器 |
| 复制完整路径 / 复制相对路径 | 复制到剪贴板 |

### 4. 大文件续发

正文一次最多 200 行 —— 但不会默默截断：卡片和引用块都会写明这次带的是哪一段（`第 1-200 行`、`第 201-400 行`），并且留下「发送剩余部分」的入口。截断提示是引用块自己的一行，**不掺进正文**，模型不会把它当成文件内容。

### 5. 已发送引用块的折叠与跳回

发出去的消息里那段 `> 📄 引用来源：…` 会自动折成一枚胶囊：**模型收到的一字不减，聊天窗口只显示标题**。点胶囊展开，移开鼠标自动合上。

胶囊右上角悬停会浮出 `↗`：点它（或 `Ctrl`/`⌘` + 点击胶囊）直接在右侧栏打开那个文件并**跳到引用所在的行**。展开/收起的老行为不变。

### 6. 长文本自动折成卡片

粘进来的一大段（会议记录、报错栈、别人甩过来的 Markdown）发出去后会自动折成一枚卡片：`长文本 · 24 行 · 1208 字`，鼠标悬停看首行摘要，点一下展开全文、再点（或点到别处）收起。

- 判定：**行数 ≥ 12 或 字数 ≥ 600**，两者任一满足就折；短消息一个都不动。
- **只改显示**：发给模型的内容一个字没少，复制按钮拿到的仍是全文。
- 只折用户自己发的行；助手回复、引用块、只带图片的消息都不参与。
- 认不出正文容器（宿主 DOM 变了）时**不折** —— 宁可按原样铺着，也不把时间戳、按钮折进去。
- **输入框里也一样**：长文本一粘进去就变成一枚同样的胶囊（`长文本 · 24 行 · 1208 字 · 首行摘要`），编辑区不再被撑到半屏；提交时胶囊展开成**原样的那段文本**，不加引用头、不加批注前缀。只在聊天输入框里生效，文件搜索框/重命名框不碰。
- 插胶囊失败（服务没就绪、光标拒绝）时**退化成普通粘贴**，绝不吞掉你粘的东西。

### 7. 引用篮

要一次引好几份材料时，点卡片上的「加入引用篮」先把它们攒起来：右下角出现一枚 `引用篮 · N` 的小胶囊，点开可以删、可以上下调序，最后「全部插入」按顺序一次落进草稿 —— 仍然不会替你发送。

### 8. 插件体检

文件树右键 →「插件体检」不再提供（右键菜单只放日常动作）。需要时在控制台执行 `__dshFileToChat.openHealth()`：一张卡片列出插件依赖的 8 项服务能力与全部 DOM 锚点，逐项 ✓/✕，附「缺了会影响哪个功能」和「复制诊断信息」。DSH 升级把某个服务改了名，这里第一眼就能看出来，不用猜为什么功能不见了。

---

## 安装

需要 **DSH Web 0.2.x**（开发与验证均在 `0.2.0-rc.2` 与 `0.2.1-rc.2` 上完成，两个版本均实测可用）。

### 方式一：从 Git 仓库装（推荐）

桌面端（DeepSeek Harness 应用）的 profile 是 `desktop`：

```bash
dsh plugin --profile desktop add https://gitee.com/mofeo/dsh-file-to-chat.git
```

`dsh web` 起的 Web 端把 profile 换成 `web`：

```bash
dsh plugin --profile web add https://gitee.com/mofeo/dsh-file-to-chat.git
```

装完在 **设置 → 插件** 里确认它是启用状态，然后**硬刷新一次浏览器**（`Ctrl+Shift+R`）。

### 方式二：手动放进 profile

```bash
git clone https://gitee.com/mofeo/dsh-file-to-chat.git
# 把目录链接/复制到：~/.dsh/profiles/desktop/node_modules/dsh-file-to-chat
```

再把 `dsh-file-to-chat` 追加到 `~/.dsh/profiles/desktop/package.json` 的 `dsh.profile.bundles` 数组里，重启应用。Web 端把路径里的 `desktop` 换成 `web`。

### 卸载

```bash
dsh plugin --profile desktop remove dsh-file-to-chat
```

Web 端把 `desktop` 换成 `web`。

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

- **只有右侧栏能触发。** 右侧栏未激活的标签页内容不挂载，所以预览没打开时不会有入口。工具栏上**没有按钮**：入口是选区上方浮出的那一枚（两个座位上挂的都是不渲染任何东西的桥接组件）。
- **Word/PPT/PDF 没有行号**，只能定位到页码 —— 这类文件走的是「Office → PDF → PDF.js」渲染链，页内没有源码行这个概念。
- **Excel 只取当前工作表**；`.xlsx` 公式取缓存值而非重算。表头是靠「选区上方那一行填得够满」猜的，猜错时卡片上可以关掉。
- **大文件一次只带 200 行**，剩下的靠「发送剩余部分」接着发 —— 不会静默丢内容。
- **跳到第 N 页不是官方能力**：`↗` 能定位到源码行（文本/代码/Markdown），但 PDF/Word 只认页码，打开后落在第一页。
- **读不到正文的文件不会假装成功**：`.docx`/`.xlsx` 靠浏览器解包抽取，抽不出来时卡片会明说原因，只发引用来源行，让模型自己去读文件。
- **依赖若干 DSH 内部 DOM 约定**（见下节）。DSH 升级后如果这些锚点变了，插件内置的锚点自检会在控制台 `warn` 一次，而不是静默失效。
- 折叠与 Excel 取值走的是**展示层后处理**，不是内核的官方扩展点。功能正常，但属于「DSH 改版时最先需要复查」的部分。

---

## 实现说明（给维护者和审核者）

这个插件**没有宿主路由、没有额外的 Node 依赖**，宿主半边（`lib/index.js`）是空的 —— 它的存在只是为了在 Loader 里占一行。全部逻辑在 `client.js`（浏览器半边，手写 CJS + `window.__ModuleLoader__` 包装），**连 `require('react')` 都不需要** —— 界面上唯一的 React 组件是不渲染任何东西的座位桥。

用到的官方扩展点：

- `sidebar.right.tab.document.actions` / `sidebar.right.tab.files.actions` —— 两个座位都只挂桥接组件，不摆按钮；借它们拿 `inputActions` / `sessionId` / `absolutePath` / `draft`
- `inputTriggers.registerSource` —— 输入框胶囊的注册与提交时的序列化编解码
- `conversation.input.for(scope)` → `insertReference(span, ref)` —— 把胶囊插进 Lexical 草稿
- `remote.workspaceFiles.read` / `readBytes` —— 读工作区文件（`read` 的 `range.offset` 用来续发）
- `remote.session.openWorkspacePath` —— 「用默认程序打开」，直接复用 DSH 自带的后端
- `sidebarRight.openResource(url, { params: { line } })` —— 胶囊跳回时打开文件并定位到行，地址按 `fileAddressFor` 的规则拼

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

功能整体不对劲（锚点变了、服务改名了）时，先看体检卡片：

```js
__dshFileToChat.openHealth()
```

---

## 开发

源码即产物：`client.js` 是手写的，没有构建步骤。改完保存即可 —— DSH 的 `dsh-client-hmr` 每 500 ms 轮询一次客户端 bundle，会**自动热重载，不用刷新页面**。（宿主半边 `lib/index.js` 不热重载，改了要重启应用；本插件正常情况下不需要动它。）

```
dsh-file-to-chat/
├── package.json         dsh.bundle / dsh.client 声明
├── cordis.patch.yml     往 Loader 里插一行
├── lib/index.js         宿主半边（空实现，只为占位）
├── client.js            浏览器半边 —— 全部功能
└── test/                纯 Node 回归用例（不进发布包）
```

回归用例不依赖浏览器：`client.js` 里的纯函数用花括号配平抠出来单独跑，

```bash
npm test    # meta / fold / table / spot / file / jump 六组，全部断言 83 条
```

改动 `client.js` 里那些纯函数（选区解析、表头探测、折行切段、地址拼装、续发游标）之后跑一遍。**`F2C_VERSION` 必须与 `package.json` 的 `version` 同步** —— `test/meta.test.cjs` 会盯着。`jump.test.cjs` 里还有一条源码级守卫：**不许出现 `ctx.sidebarRight` 这种直接读取**（未写进 `exports.inject` 的服务会被 cordis 严格访问拦下、当场抛错），只能走 `ctx.get('sidebarRight')` —— 1.0.3~1.0.9 的「引用块折不起来」就是它。

---

## 仓库与镜像

| 位置 | 角色 | 说明 |
|---|---|---|
| [gitee.com/mofeo/dsh-file-to-chat](https://gitee.com/mofeo/dsh-file-to-chat) | **开发主仓（权威）** | 提交先落这里；版本号、CHANGELOG 与发布版本以主仓为准 |
| [github.com/feifiefei/dsh-file-to-chat](https://github.com/feifiefei/dsh-file-to-chat) | 镜像 | dsh-plugin.org 的收录页与安装命令指向它，**可能滞后于主仓**，仅供同步与市场发现 |

`package.json` 的 `repository` / `bugs` 指向主仓（npm 的 `repository` 只接受单一仓库），GitHub 镜像另记在 `mirrors.github` 里 —— 这是本项目自定义的字段（和本项目的 `dsh` 字段一样），npm 不会解析它，只供人和脚本读取。

### 从插件市场进来，怎么拿到最新版

[dsh-plugin.org 收录页](https://dsh-plugin.org/zh/plugins/feifiefei/dsh-file-to-chat) 给的是 GitHub 源码安装命令；如果镜像还没同步到最新提交，改用主仓命令装（或把已有安装切到主仓）：

```bash
dsh plugin --profile desktop add https://gitee.com/mofeo/dsh-file-to-chat.git
```

---

## 许可

[MIT](LICENSE)
