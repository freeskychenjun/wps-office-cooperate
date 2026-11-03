---
name: wps-office-cooperate
description: 在 WPS 里实时改 Office 文档——Word（.docx，修订模式改写+批注+LaTeX 原生公式+脚注+版式核对+总览图）、Excel（.xlsx/.xls/.csv，批量写值公式+锚点校验+快照回滚+批注+数据清洗/透视表/条件格式/图表次轴）、PowerPoint（.pptx，预设版式建页+形状级改文字+表格图片+套版裁切+框图连线+SmartArt+逐页渲染核对），改动当场出现在用户正开着的 WPS 窗口里。当用户要求：在 WPS 里改这个文档/表格/演示文稿/PPT、在 WPS 里做修订、实时改 Word/Excel/PPT、一边看着一边改、做汇报 PPT、加批注、写公式/脚注、清洗表格/做透视表、核对排版/版式/表格格式，或提到 wps word、wps excel、wps ppt、live edit word/excel/ppt 时使用。不适用：纯离线批量处理（改副本、不要求用户在场，Word 用 docx-paragraph-polish，Excel 用 openpyxl/minimax-xlsx，PPT 用 pptx-generator 等离线方案）、用户明确不要动原文件（只出旁挂报告）、微软 Office 窗口里的文档（本 skill 只走 WPS）。
---

# wps-office-cooperate：在 WPS 里实时改 Word / Excel / PowerPoint


**版本 v1.32（2026-11-03）**：**基础设施两项（忙重试 + 剪贴板不再被吞）＋ PPT 成品版式层 phase 2**。① **OLE MessageFilter**：三端 helper 的 `[STAThread] Main` 首行注册 `MessageFilter.Register()`——WPS 拒绝调用时以 **150 ms 间隔自动重试、最多 6 s**（仍忙才抛成 `BUSY`），以前是立刻失败或等到 180 s 超时；公寓级注册，一次性/常驻模式都覆盖；忙码补 `0x800AC472`（Word 另补原来缺的 `0x80010001`）。真机取证 `CoRegisterMessageFilter -> 0x00000000`，策略值用同源探针逐项验证；**真实忙场景未能在本机构造**（F2 编辑态、模态保存提示下 COM 仍 4 ms 返回），"拒绝→重试→BUSY"属"机制已装、未实弹"。② **ClipboardKeeper**：Excel `render`、PPT `render`/`render -Overview`/`export png|jpg` 前后保存并恢复用户剪贴板（序列号判据：期间用户又复制了新东西就不覆盖；恢复内容打"不进 Win+V/云剪贴板"标记；失败只 Trace 不影响出图）。Lead 复跑：设置文本 → `excel render`/`ppt render` → 剪贴板原文不变且 PNG 正常。③ **PPT phase 2**：`slide -Preset canvas`（自由版面，`items[{type,box}]` + `edges` **用 `AddConnector` 真连到形状边**，`connected.begin/end` 回读，越界夹取并报）、`-Preset grid`（`rows` + `span`，宽度按 span 均分、行高按内容估、`stack` 竖叠，**不用给坐标**）、**形状绘制图表**（`chart` item：scatter 带 OLS 拟合线+方程+R²、curve 末端标系列名、pie/doughnut 真扇区，刻度 1/2/5×10^k 取整）、`slide -Action reuse`（复制整页按形状名换字换图，页码自动改成新页号）。④ **修掉两条真问题**：`reuse`/`dup` 内部 `Duplicate()` 把幻灯片留在剪贴板，v1.32 的剪贴板快照去读这个 OLE 格式会**直接把 helper 进程带崩**（`INTERNAL: helper exited without writing a response`）——改为只快照简单格式 + `Duplicate` 包在 Keep 里，修后 `reuse` → `render` 不用再手动清剪贴板；另实测 **WPS 拒绝 `Presentation.Tags`**（`Add` 不报错但 `Count` 恒 0）⇒ 主题仍存本机文件，`theme -Probe` 可复核，返回 `tags.supported:false` 如实说明。⑤ **回归**：`tests/word/check_daemon.py` **40/40 全过**（三端 daemon 与一次性逐字节一致、daemon 仍更快、`serve start/stop` 与回退正常）——消息过滤器与剪贴板改动没有破坏传输层。

**版本 v1.31（2026-11-02）**：**Excel 补"文字更正/日期分组"，PPT 补"成品版式层"**。① **Excel `replace`（内存块替换）**：`-Find/[-Replace]/-Sheet/-Range/-Column/-Whole/-MatchCase/-Regex/-DryRun/-Max/-All`；**绝不走 `Range.Replace`**（WPS 会在刚写进去的文本上二次替换，`P→P-` 得 `P--`）——我们一次块读 → 本地单遍替换（游标移到插入串之后）→ 只写真正变化的格，实测 `PP→P-P-`；返回 `scanned/changed/byColumn/samples/formulaCellsSkipped/numericMatchesLeftAlone/truncated` + `preSnapshot`（可 `diff -Since`/`restore -Since`），**公式格永不写**、**数字/日期格不改成文本**、`-DryRun` 与实跑同一套判据、不自匹配的规则幂等（实测 5→0）。② **Excel `pivot -Group` 定案**：WPS 12.1 **原生支持**按日期分月/季/年（24 行实测月 `110…1320` 总计 8580、季 `660/1650/2640/3630`、年 `780/7800`，全与手算一致），返回 `groupMode`（`none\|native\|failed`）并**回读 `RowRange` 行项数**校验（不再"没报错就算成功"）；文本日期抛「选定区域不能分组。」→ `failed` + 辅助列降级配方。③ **PPT 成品版式层 phase 1**：`slide -Preset` 新增 **8 档**（`cards`/`stats`/`timeline`/`process`/`compare`/`summary`/`gallery`/`split`，数据走 `-Items <json>`，960×540 设计画布按实际页尺寸缩放，标题/色杠/页脚/页码统一，字号阶梯自动回退并返回 `overflowWarnings`/`fontSizeDrops`）；`theme` 新增 `accent2`/`surface`/`cardFill`/`cardRadius`/`sectionStyle`（旧主题文件缺字段取默认）。④ **修掉图片裁切的错误路线**：WPS 的 `PictureFormat.Crop*` 按**点**解释、且改宽高会把裁切洗掉（实测 1200×600 进 480×540 变成 **552×540、x=-36 出页**，`check` 报 blocker）——改为 **GDI+ 预裁临时副本再按框精确插入**，`split`/`gallery` 的 cover 都走这条，实测回读精确 480×540、`check -MinSeverity blocker` = 0。

**版本 v1.30（2026-11-02）**：**Word 排版收尾三件 + 两处"报成功其实没落上"的修正**。① **新命令 `fillgaps`**：图/表放不下整块跳到下一页时，把图/表后面的正文段提到前面填掉上一页底部的留白（照 dsh-office 的 `fill_gaps` 移植；`-DryRun` 只报不动、只读文档可跑，`-MaxMoves`/`-MinGapPt` 可调；不越过标题、冒号引导句连带搬、题注/图段/表内段不搬；表格块搬运会在表前留一个空段——实现代价已记 quirks）。实测自造 3 页稿：`-DryRun` 报 figure gap 124.2pt / table gap 210.7pt，实跑 `moved:2`，`render` 前后对照可见空白被正文填上、图表原地不动、标题未被越过。② **`noSplitWords` 极性修正**：`ParagraphFormat.WordWrap` 是**反直觉**的——实测 `-1` 才是整词、`0` 是允许拆词（页宽 240pt/边距 90pt/20pt Arial/"aa 3700000"：`-1` → 首行只有 `aa`，`0` → 首行 `aa 37`）。我方 v1.29 及更早写成 `v ? 0 : -1`，即**传 `-NoSplitWords` 反而把拆词打开**；现已改正并**回读**（`noSplitWords=true (readBack=-1)`，不符即报错），且 `insert` 新段/`table` 单元格/`caption` 题注**默认整词**；新增 `-SplitWords` 反向开关（`-NoSplitWords:$false` 过不了 `woc.ps1` 的参数往返）。③ **`table -KeepTogether` 真落上**：旧实现走 `Rows.Item(r).Range.ParagraphFormat`（WPS 上半实现对象，会抛 `0x800A1767`）且异常被吞 ⇒ 报成功但可能一行没落上；现在逐行 `Select()`+`MoveEnd(10,1)` 走 `Selection` 写 `KeepWithNext`/`AllowBreakAcrossPages` 并**逐行回读**，失败则整条命令 `ok:false`（`KEEPTOGETHER_FAILED` + `error.detail` 点名行号）。实测 9 行跨页表：改前第 3 行被页底劈开，改后 `rowsSet 8/8`、`allowBreak 9/9`、`readBackFailures []`，每页一行完整。

**版本 v1.29（2026-11-02）**：**对标开源项目 [dsh-office](https://github.com/copylee711/dsh-office) 补齐三处能力缺口**（研究结论与实测台账在仓库 `docs/borrow/`，不随包分发）。① **Word 原生公式**：`formula` 把 LaTeX 写成真 OMML（`Flat OPC 包 + Range.InsertXML`，WPS 只认这个形状；裸 OMML 片段无效），插后校验 `OMaths.Count` 增长否则回滚，子集外命令点名 `MATH_UNSUPPORTED`。② **Word 原生脚注**：`footnote` list/insert/edit/delete/numbering（`Document.Footnotes`；`Document.Footnotes.NumberingRule` 在 WPS 可读可写）。③ **Word/Excel `render -Overview`**：联系表总览图（Word 6 页/张、≤18 页；Excel 整簿每表一张缩略图），交付前一次看完不用逐页要图。④ **Excel 数据操作层**（最大空白）：`clean`（`¥1,299`/`89元`/`10%`→数字、七种日期写法→日期）、`fillblanks`、`dedupe`、`deletewhere`、`sort`、`filter`、`copy`、`addcol`、`pivot`（结果直接回读）、`freeze`、`cformat`、`validate`、`table`；`read -Profile` 加 **LOOK 段/日期识别/错误码**，`fill` 回报错误格，`chart` 加 **组合图+次坐标轴**/数据标签/图例/占地区域与重叠提示。⑤ **PPT 套版能力**：`crop`（满版/人像裁切，focus 保头）、`connect`（形状间连接线，框图）、`smartart`、`bg`（纹理/渐变/图片背景）、`anim`（切换/入场），`chart` 加趋势线/数据标签/图例位置。⑥ **上下文瘦身**：v1.1–v1.28 版本记录拆到 [reference/changelog.md](reference/changelog.md)，主文档每次加载少读 1.93 万字符。**全部新增命令均在 WPS 12.1.0.29161 上实测**；WPS 硬限制（SmartArt 节点不可写、`radar` 名义支持、裁切单位是点、组合图标题重试等）如实写进各端 quirk 文档与下方坑表。

> **版本历史 v1.1–v1.28 → [reference/changelog.md](reference/changelog.md)**（2026-11-02 起从主文档拆出，避免每次加载都读 2 万字历史）。本节只保留**当前版本要点**。

**能力边界（一句话）**：只走 **WPS**（Word=Writer `KWPS.Application`，Excel=表格 `KET.Application`，PPT=演示 `KWPP.Application`，没装微软 Office 也照跑）。Word 侧所有文字改动都是**修订（track changes）**，用户能逐条接受/拒绝；Excel 与 PPT 侧改动**直写**，靠**锚点校验 + 快照**保证可核对（**这两个 app 的 Ctrl+Z 撤不回 Agent 的改动**，外部 COM 修改不进撤销栈）。

> 下文 `<skill>` = 本 skill 的安装目录。命令一律按**完整可跑**形式给出：
> `powershell -NoProfile -ExecutionPolicy Bypass -File <skill>\scripts\office.ps1 <word|excel|ppt> <命令> [参数]`
> （`word.ps1` / `excel.ps1` / `ppt.ps1` 也可直接调用，参数完全一样。装有 PowerShell 7 的机器可把 `powershell` 换成 `pwsh`。）
> 输出只有一段 JSON：`{"ok":true,"result":{...}}` 或 `{"ok":false,"error":{"code":"...","message":"..."}}`。
> 环境：Windows + WPS；首次调用会用系统自带 `csc.exe` 编译 helper（约 1 秒，之后按源码哈希缓存，不再编译）。

## 铁律（三个 app 通用，一条都不能省）

> **调用方式铁律（v1.17 更新）**：命令行调用优先 **`scripts\woc.ps1 <app> <命令> ...`**（v1.17 新增：不经过 cmd.exe，公式/`^`/`&`/`$` 原样到达）；只有在 cmd.exe 里调用且参数简单时才用 `scripts\woc.cmd`——**cmd.exe 会吃掉以 `=` 开头的参数并把 `^` 当转义符**，`fill/write -Formula` 走它必坏。两者都等价于 `powershell -Mta -NoProfile -ExecutionPolicy Bypass -File <app>.ps1 ...`。**必须带 `-Mta`**：
> 本机 STA 公寓初始化要 ~2.2 s（实测 `-Sta` 2.2 s vs `-Mta` 0.13 s），而包装器并不使用 COM，
> 带 `-Mta` 不会影响任何功能。**helper 自己是 `[STAThread]`，不要改**（WPS 自动化对象需要 STA）。
> 不按这条走一切照样能跑，只是每条命令白付 2 秒。


0. **一份文档，一条通道（最重要）。** 一旦本 skill 打开了某份文档、或接管了用户正开着的那份，**在这次协作期间，这份文档的一切处理都走本 skill**：
   - Word 校审 → `word audit` + `comment -Items`（**不要**加载 `design-report-review`）
   - Word 润色 / 改写 → `revise`（**不要**加载 `docx-paragraph-polish`）
   - Word 排版 → `format` / `style` / `page`（**不要**加载 `design-report-format`）
   - Excel 改表 → `write`/`fill`（**不要**再用 openpyxl/pandas/minimax-xlsx 直接改磁盘文件）
   - PPT 改稿 → `slide`/`text`/`textbox`（**不要**再用 pptx-generator/python-pptx 生成副本替换它）

   理由不是"体验"，是**会坏**：那些 skill/工具改的是磁盘文件，而这份文档正被 WPS 独占（实测写盘 `PermissionError`）；它们还产出副本，会让你眼前这份和副本分成两条线。
   唯一例外：用户明确要"**不动原件、出一份新稿**"或"**无人值守批量跑**"——那时先 `save` → `close` 交出去，回来再 `open` 产出的文件、`snapshot` 重建基线（Excel 用 `save -As` 另出一份，或在用户关掉文件后离线处理）。
1. **先 `status`，再动手，并说出在哪儿动手。** 每次开工第一条命令是 `status`，把"几个实例在跑 / 开着哪些文档 / 哪份是 active / 有没有未保存修改"念给用户，并明确报出**即将修改的那份文档的完整路径**（Excel 再加工作表名）。不碰用户没指定的文件。
   接着**让用户看得见接管**：`open` 已自动把标签/标题改成「〔AI接管中〕 文件名」，再 `hint -Doc <路径> -Toast` 弹一句人话（"已接管《X》…"）——同一份文档一轮协作只需公告一次，后续进度用 `hint -Text "正在做…" -Quiet` 刷状态栏（表格可见；Writer 不渲染，别指望）。**PPT 的标签页是自绘的、改标题只有任务栏看得见**：只有用户明确要"在页面上/投屏时看得见接管"时才加 `-SlideMark`（逐页贴「〔AI接管中〕」药丸，这是**内容**——Ctrl+S 会存进文件，所以默认不贴；收尾必须 `finish`/`hint -Clear`，万一被存进去，下一条命令会自动修复并返回 `badgeRepaired:true`）。要边干边看进度就再加 `-Tray`（托盘图标，悬停看当前步骤、右键可结束并还原；新图标默认在「^」里，提醒用户拖出来）。`finish`/`close` 会自动还原，别留着标记过夜。
2. **`readOnly` 就是不许改。** `open`/`read` 返回 `readOnly: true`（文件被别处占用或有只读属性）时，写命令会直接报 `READONLY`。**不要**用另存、改文件属性、杀占用进程等方式绕过去——先问用户。
3. **不替用户关文档。** `close` 只关本次明确要关的那一份；用户开着的其它文档、用户自己的 WPS 窗口一律不动。关之前先 `status` 确认这份是本次打开的。
   反过来，**用户在你看不见的时候关掉了文档**：Word 的下一条命令会**静默重开**并返回 `reopened:true`（若日志显示关闭前未保存，还会带 `unsavedAtClose:true` 警告，提醒可能丢了改动），Excel/PPT 则返回 `DOC_CLOSED`——必须**显式 `open`** 才能继续，别假装还在协作；托盘图标会在约 10 s 内自动消失。
4. **结束必须 `close`，并汇报进程数。** 收尾命令 `close -Doc <路径>` 会返回 `processes`（Word 侧 wps.exe / Excel 侧 et.exe / PPT 侧 wpp.exe 计数）、`documentsLeft`。把这两个数报给用户；只清理自己起的无窗口 `/Automation` 实例，**绝不**按映像名批量杀进程。（注：WPS 退出后会自己预拉起 1–2 个隐藏进程，约 1 分钟自灭，别把它当成泄漏。）
5. **交付前提醒"落盘可能是密文"。** 装了单位透明加密的机器上，WPS `save` 之后文件头变成 `17DA5FA01633CD9A`，直接发给同事打不开。保存后必须提这一句，并按 [reference/encryption.md](reference/encryption.md) 处理。
6. **落盘前先想清楚反悔手段，别改完就存。**
   - **Word 修订类**（`revise`/`comment`）：拒绝修订不受 save 影响，可以及时存；
   - **Word 格式类**（`format`/`style`/`page`/`table`/`font`/`picture`/`caption`）：不进修订，唯一反悔手段是撤销栈，而 **save 会清空撤销栈（实测）**——必须在**用户看过/认可**或 `render` 复核通过**之后**才 save；
   - **Excel 一切写操作**：COM 修改不进撤销栈，改前 `snapshot`（`write`/`fill` 已自动落 pre-batch 快照）、改后比对 before/after 与 `diff`，别指望 Ctrl+Z；
   - **PPT 一切写操作**：同 Excel，无修订层、无撤销；批量改前 `snapshot`、改后 `diff` + `render` 目检；
   - `snapshot`/`diff`/`handoff` 都走 COM 读活文档，不依赖 save，不构成提前落盘的理由。
7. **先分清"只连接"和"要动手"，别把连接做成体检。** 用户只说"连上 / 接管 / 打开这份文档 / 先看看有什么"时，本轮做到 **`status` → `open` → 报一句摘要**（文件名 / 页数或段落数 / 纸张或尺寸 / `readOnly` / 有无未保存改动）**就停下等指令**：不 `read` 全稿、不 `render`、不 `audit`、不 `snapshot`、不 `close`（文档留给用户继续用，等他下一次给出任务时再按铁律 4 收尾）。这些动作只在用户给出具体任务时才做，且只做该任务需要的那一部分——大文档 `read` 一次吐上万字符、`audit` 一次吐上百条（人工老稿误报居多），对"只是连上"的用户是纯等待成本。
   反过来同样成立：用户给了"改"的任务，别在没改完、没 `render` 复核之前就宣布完成。
8. **公式与脚注是"原生对象"，不是文字（v1.29）。** `formula` 写的是真 OMML 公式、`footnote` 写的是真脚注——它们**不随 `revise` 的 `replace` 改字**（公式内部文字不能用 find/replace 改，脚注正文也不在正文段落里）。要改公式就重插（`formula` 无删除命令，用 `undo` 或让用户手工删）；要改脚注一律 `footnote -Action edit`。同理，**含脚注的段落不要整段 `set_text`**——旧实现会把脚注一起删掉；改成段内 `replace` 并且不跨过引用标记。

### 接管提示（`hint`，v1.11）

用户的原话是"agent 接管 wps 文档后可否有比较友好的提示"——所以**接管那一步必须看得见**，而且不能打扰：

- **`open` 自动打标签**：窗口标签/标题变成 `〔AI接管中〕 报告.docx`（Writer / 表格都有；WPS 12 的标签页从右截断，所以标记放前面，长文件名也不会被吃掉）。这是纯窗口状态：不进修订、不写文件、不改 `saved`，`-NoHint` 可关。
- **`hint` 补一句人话**：默认弹 6 s 托盘气泡——"已接管《报告.docx》：我会把改动做成修订或批注，你随时可以接受或拒绝。"`-Text` 换说法，`-HintTag` 换标题里的标签文字，`-Bar` 单独控制状态栏那行（表格会显示，Writer 不渲染），`-Quiet` 静默（不弹气泡，适合刷新标题/状态栏）。
- **还原**：`hint -Clear`、`finish`、`close` 都会把标题恢复成文件名、把状态栏交还 WPS；`status` 每份文档带 `caption` 与 `hinted`，随时可核对当前有没有挂着标记。
- **PPT 的标题标记（v1.26）**：COM 的 `ActiveWindow.Caption` 只读、`Application.Caption` 不重绘标题栏、`StatusBar` 属性不存在，但窗口框（`PP12FrameClass`）接受 `SetWindowText`——`open`/`hint` 给标题加「〔AI接管中〕」前缀，`hint -Clear`/`finish`/`close` 去掉；用户切标签页时 WPS 自己重置（不会挂错文档），下一条命令自动补回。返回里 `captionSupported:false` 仍表示"COM 标题不可写"，`titleMarked:true` 表示 Win32 标记已生效；气泡照旧（`-Toast`）。
- **PPT 的页面标记（v1.27 起，v1.28 改为按需开启）**：WPS 演示的标签页是自绘的（从文件名画），改窗口标题只有任务栏 / Alt+Tab 看得见；唯一"界面里看得见"的通道是在**每一页右上角**贴一枚「〔AI接管中〕」橙色药丸（形状名 `woc-ai-takeover-mark`，最上层，放映时也在）。**它是内容不是窗口状态**，所以：默认**不贴**（`hint` 零改动，响应里 `inAppHint` 提示要加 `-SlideMark`）；`-SlideMark` 显式开启时返回 `slideMarkWarning`；`save`/`export` 落盘前剥离、写完补回；`hint -Clear`/`finish`/`close` 逐页删除；**用户手动 Ctrl+S 造成的污染会自动修复**（下一条命令：剥离→重存→贴回，返回 `badgeRepaired:true`）。读侧零干扰：`read`/`search`/`snapshot`/`diff`/`check`/`replace`/`outline` 都跳过它（`read` 每页给 `takeoverBadge:true`）。
- **常驻托盘图标（v1.12，`-Tray`）**：`hint -Doc X -Tray` 在系统托盘挂一个橙色「AI」图标，悬停显示「〔AI接管中〕 文档 · 当前步骤」，右键可**结束协作并还原标记**；`finish`/`close`/`-Clear` 自动移除，最长 4 小时自灭。**新图标默认在「^」溢出区**（Windows 行为，进程无法自我提升），要提醒用户拖到任务栏常显；PPT 侧这是唯一常驻通道。用户中途关掉文档，图标约 10 s 内自动消失（不会再挂着"AI 协作中"骗人）。
- **不做悬浮窗（用户明确决策，2026-10-30）**：不加置顶悬浮条/浮层——它会遮正文、抢焦点，而且要常驻 daemon 跟随 WPS 窗口（移动/最小化/切标签/异常退出都得收拾）。常驻信息由「标签页前缀标记 + 托盘图标」承担；**不要再向用户推荐，也不要把它做成"可选模式"偷偷加回来**。

---

## 多任务模式（一次布置多个任务，v1.20）

用户经常一次给好几件事：多章节、同一章好几个任务、或"第 3 章润色 + 全文单位上标"。规则是 **并行规划、合并写回**：

1. **冻结**：`status` → `open` → `snapshot -Tag pre` → `read -Count 0 -SkipTableCells`，一次拿全稿段表（`i` / `text` / `style`）。
2. **切分 + 并行规划**：`tasks.ps1 ranges -Read read.json -Out ranges.json -Styles "标题 1"` 自动切章（也可手写），再 `tasks.ps1 split -Read read.json -Ranges ranges.json -Out work` 得到章节片；每个任务派一个子代理，只喂切片 + 规则，**不调 skill 命令、不碰 COM**，产出 `task-*.json`（字段见 reference/multi-task.md）。
3. **合并**：`tasks.ps1 merge -Read read.json -TaskFiles work -Out work`——按段合并、做 `expect` / 数字多重集 / 相似度 guard，输出 `merged.json`。**同一段的多任务必须合成一条 item**：一批 `revise` 里同段放两条，第 2 条会 `could not place the edit` 并整批回滚。
4. **提交**：`tasks.ps1 plan -Merged merged.json -Ranges ranges.json -Out work -Doc <绝对路径> -BatchOps`，按 **表格结构 → 文本批 → 格式 → `unit` → 批注** 的顺序写；多章节**从后往前**；可用 `word.ps1 batch -Ops work\batch-ops.json` 一次进程启动跑完整批。

三条铁律：

- **写回只有一个写者**（父代理）；子代理只做规划。因此不需要锁，也不会互相踩段号。
- **同段一条 item**；优先 `replace`——`set_text` 会丢段内 run 级格式（单位上标）。
- **段号会变的先跑**：表格结构 / 插删段放在文本批之前，跑完重读，再提交文本与批注。

汇报：回执写清任务清单与顺序理由（例如"T2 在 T1 之后，避免上标被整段重写冲掉"）；完成后分**局部改动 / 全局改动**两栏 + `diff` 摘要 + 待裁决项；**不要 `save`**，等用户确认（铁律 6）。

细节（任务卡字段、三种组合的排法、实测坑与验收清单）见 [reference/multi-task.md](reference/multi-task.md)。

**跨 app 编排**：多份文档/多个 app 的任务先用 `tasks-session.ps1 plan` 排波次（同文档串行、跨文档并行≤3），每波内部再用对应 app 的 `tasks-*.ps1`。

**Excel / PPT 版**：Excel 用 `tasks-excel.ps1`（split/merge/plan，任务卡与坑见 [reference/multi-task-excel.md](reference/multi-task-excel.md)）；PPT 用 `tasks-ppt.ps1`（见 [reference/multi-task-ppt.md](reference/multi-task-ppt.md)）。三端共用同一条纪律：**并行规划、单写者提交**。

---

# Word：在 WPS Writer 里实时改文档

命令前缀：`office.ps1 word <命令>`（或直接 `word.ps1 <命令>`）。

## Word 命令速查

| 命令 | 参数 | 作用 |
|---|---|---|
| `doctor` | — | 自检：哪个套件可用/版本/是否在跑/开着的文档/渲染是否可用/各类进程数 |
| `status` | — | 在跑的实例与文档清单（name / path / saved / readOnly / active / paragraphs） |
| `open` | `-Doc <路径> [-Show] [-ReadOnly] [-NoHint]` | 已开则复用，未开则打开；返回 readOnly、段落数、页数；发现可能开在另一实例时给 `multiInstanceWarning`。**v1.11 起自动给窗口标签打「〔AI接管中〕」接管标记**（`hint.caption`/`hint.captionSet`，纯窗口状态；`-NoHint` 关掉） |
| `read` | `-Doc <路径> [-From 1] [-Count 60]`；大文档加 `[-SkipTableCells] [-Only <正则>] [-Slow]` | 段落（编号/样式/自动编号/文本/inTable/**表号**）+ 表格清单 + 页数/修订数/批注数。**凡是带 `-SkipTableCells`/`-Only`，`-Count` 就变成"返回多少段"**（不是"走多少个段落号"），`-Count 0` = 一直扫到底；返回里 `scanned`/`skippedTableCells`/`filter` 说明跳过了什么。**v1.10 起过滤扫默认走快路径**（一次取全文+本地过滤，`path:fast`；4479 段/20 表实测 38.6 s → 8.4 s，与旧路径逐字段一致；有未处理修订或校验失败自动回退，`-Slow` 强制旧路径）。**v1.13 起块模型失效时自动改用 `xml` 路径**（一次 `WordOpenXML`+本地解析，7334 段全篇扫 8.2 s；`path` 字段说明走了哪条）。`-Only` 的收益取决于选择性：命中大半文档时快慢持平 |
| `search` | `-Doc <路径> -Text <词\|词2\|词3> [-Style 样式名] [-Max 50] [-Case]` | **全文找词**（WPS 原生 Find，大文档秒级；多词用 `\|` 分隔一次查完，返回每词计数）——定位用这个，别一页页 `read` |
| `selection` | `-Doc <路径>` | **读用户当前选中/光标**：返回选中文本与覆盖的段落号（`fromPara/toPara`）；在表格里时再给 `table`（表号，直通 `table -Table n`/`cell -Table n`）与 `row`/`col`；覆盖**图片**时给 `images`（图号，直通 `picture -Image n`）——只读不改 |
| `revise` | `-Doc <路径> -Items <json文件>` | **修订模式**批量改写（见下），返回改动与修订计数 |
| `comment` | 列表：`-Action list`；单条：`-Para N [-Anchor "文字"] (-TextFile <txt文件> \| -Text "短文本")`；**批量：`-Items <json>`**；删除：`-Delete -Index N \| -All` | 加批注（中文正文走文件，避开引号转义）；`list` 返回每条的编号/作者/锚文/内容——编号就是 `-Delete -Index` 用的 |
| `render` | `-Doc <路径> -Page 1 [-To 4] -Out <目录> [-Width 1100]`；**总览：`-Overview [-Page 1] [-To 18]`** | 页面 EMF → PNG（宽 320–2400），返回每页路径与尺寸。**`-Overview`（v1.29）= 联系表总览**：一张图 **6 页、3 列**、格宽 **640**、浅灰底 + 每格上方「第 N 页」，**一次最多 18 页**（超出返回 `capped:true` + `more`）；一次调用看清半空页/图表出框/表格断裂，可疑页再用 `-Page n -Width 1100` 放大复核 |
| `formula` | 单条：`-Doc -Para N [-Expect "<该段开头>"] -Latex "<LaTeX>" [-Display] [-Number "1"] [-Anchor "文字" -Where after\|before]`；**批量：`-Items <json>`**（`[{"para":N,"expect":"...","latex":"...","display":true,"number":"1"}]`） | **LaTeX → WPS 原生公式**（真 OMML，不是图片、不是线性文字）。行内式插到段末（或 `-Anchor` 指定处），`-Display` 独占一段并居中，`-Number`/`\tag{n}` 生成右对齐编号 `(n)`。**插入后校验 `OMaths.Count` 必须增加**，否则回滚并返回 `MATH_FAILED`；LaTeX 超出子集返回 `MATH_UNSUPPORTED`（点名不支持的命令），两者都**不留残留**。一次 ≤12 条 |
| `footnote` | `-Action list`；`insert -Para N [-Expect] [-Find "词" -Where after\|before] (-Text "…" \| -TextFile <txt>)`；`edit -Index N [-Expect "…"] (-Text\|-TextFile)`；`delete -Index N`；`numbering -Number continuous\|section\|page` | **原生脚注 CRUD**。脚注是真脚注（WPS 自己编号），`list` 给 `i/label/para/page/text`（`i` **永远是 `Document.Footnotes` 的序号**，不是 `Range.Footnotes` 的）；`edit/delete` 带 `-Expect`（该脚注正文开头）不符报 `ANCHOR_MISMATCH` 而不是改错条 |
| `font` | `-Doc <路径> -Name <字体名> [-Track]` | 把**正文、页眉、页脚、文本框**里的字体统一改成指定字体（中英文一起改，`-Name 仿宋`）；返回每个区段的结果与回读值。默认**不产生修订**，整批算一个撤销单元 |
| `format` | `-Doc <路径> [-Para N \| -From N -To M \| -Style 样式名 \| -Match <正则>] [-Slow]` + 任意格式项 | 对**整篇/一段/一段区间/某样式的所有段/文本匹配正则的所有段**套字符与段落格式（`-Match` 例：`"^\s*(图\|附图)\s*\d"` 一次选中全部图名；返回命中数与前 5 个样例）。`-Style` 与 `-Match` 的定位扫描都在 WPS 进程内完成（v1.10 起 `-Style` 改 Find-by-style，往返 O(命中) 而非 O(段数)） |
| `style` | `-Doc <路径> -Style <样式名> [-NoNumbering]` + 任意格式项 | 改**样式本身**（正文、标题 1…），所有用它的段落一起变 |
| `page` | `-Doc <路径>`（**不带任何选项=查询各节设置**）`[-Paper A4] [-Orientation portrait\|landscape] [-Top/-Bottom/-Left/-Right 厘米] [-HeaderText] [-FooterText] [-PageNumber center\|right\|left\|none] [-PageNumberStart N\|none] [-All]` | 纸张/方向/页边距/页眉页脚/页码（逐节应用）；查询模式报每节纸张/方向/边距/页眉页脚 **+ `pageNumbers`（起始号/是否节内重启/该节首段的实际显示页码——页码断号一眼看出）**；`-PageNumberStart none` 试图清掉「本节重新编号」（WPS 常做不到，会**明说失败**并给 XML 路线） |
| `table` | 插入：`-Para N -Data <json> [-Borders grid\|thick\|three-line\|none] [-Widths "2,3,2"] [-HeaderRow] [-Caption]`；改表：`-Table n [-Borders ...] [-Autofit] [-RowHeight] [-KeepTogether] [-Font/-Size/-Align]`（**省略 -Table = 全文档所有表，一次 COM 遍历 + 单个撤销记录**）；**体检 `-Action scan`（每行带 `para` = 表起始段号）**；**表头跨页 `-Action header-repeat [-Table n] [-Row 2] [-Passes 3]`**；**题名出框 `-Action caption [-Table n] [-Borders thick] [-Force]`**；**删行 `-Action rows -Table n [-Row 1]`**；**v1.33 一键整备 `-Action plan`**（只读侦察：每表 pages/rows/columns/题注位置与页码/边框/tblStyle/tblW/gridCols/pt 列宽/fontSize + `risks[]`（`captionInFirstRow`/`bordersNotThick`/`breakBefore`/`noStyle`/`noBorders`/`pctWidths`/`gridOverflow`/`multipageWithoutHeaderRepeat`）+ `captions[]`/`references[]`/`issues[]`（复用 captions 同一解析器）+ `renumber` 提议）与 **`-Action normalize`**（**默认只改 plan 判定未达标的表**：题注外提→空白规范化→可选重编号→加粗边框（WordOpenXML 列宽快照/恢复）→跨页表头→自检；逐表隔离、单条命令一个撤销单元、不 save；`-DryRun`/`-Table n`/`-NoCaptions`/`-NoBorders`/`-NoHeaderRepeat`/`-Gap 2`/`-Borders thick`/`-Fit keep\|window\|content`/`-Plan <plan.json>`/`-Verify`）。**表格整备先跑 plan 再 normalize，不要逐条小命令拼**（2026-11-03 实测：同一份 14 表报告用旧做法 127 次调用/88.6 min，其中 43 步是勘探与返工） | 建表 / 改表 / 批量体检 / 表头重复 / 把做进第 1 行的题名提到框外并删该行 / 删指定行（后三者内部绕过 WPS 的 `Rows.Item(n)` 残缺对象，走选区写法）。`scan` 的 `para` 可直接 `read -From <para-3> -Count 6` 看题注，或喂给 `-Action caption -Table n`。**v1.30**：① 新建表格的单元格默认**整词换行**（长数字不被拆）；② `-KeepTogether` 改为**逐行走选区写法 + 逐行回读**，返回 `keepTogether:{rowsExpected,rowsSet,readBackFailures,allowBreak}`，WPS 没接受的行走**整条命令 `ok:false`（`KEEPTOGETHER_FAILED`）**并点名行号——不再"报成功其实没落上"。**v1.33 两条硬结论**：③ **表名外提不是本 skill 的可靠能力**——题名做进表首行时，`normalize` 只在「回读题注页码 == 表起始页」成立时才外提（普通段落后/手动分页符前通常可以）；**表前是分节符时 WPS 无法在断点后造正文段，一律报 `needsHuman`，不重试**（彻底解决请人工把分节符剪到题注之后，或用微软 Word 的 `-CutPaste`；「删表+InsertXML」在 WPS 上会把 Writer 转死，默认禁用）；④ 给**XML 里没有 `w:tblBorders` 元素**的表写边框会让 WPS **重建整表布局、按渲染宽度重推 gridCol**（列宽漂移；无 tblStyle + dxa 固定的表会"崩"）——所以写前快照、写后回读、漂移即恢复，且已达标的表**跳过不写** |
| `fillgaps` | `-Doc <路径> [-DryRun] [-MaxMoves 8] [-MinGapPt 0]` | **填掉图/表前面的页面留白**（v1.30；照 dsh-office 的 `fill_gaps` 移植）：图/表因放不下整块跳到下一页、上一页底部空出一截时，把**图/表后面的正文段**提到图/表之前补上。`-DryRun` 只报不动（**只读文档也能跑**），返回 `moves:[{block,kind,para,page,gap,text}]`；空白阈值默认 `max(90pt, 文本高度×12%)`，`-MinGapPt` 可覆盖。**保守判据**：块前一段必须是正文（不越标题）、块首行页码必须晚于前一段、块后段不能是空段/题注/图段/表内段、冒号结尾的引导句连带搬。搬完会**在表格前留一个空段**（实现代价，已记 quirks）；`moved:0` 是正常结果（说明这稿没有值得填的留白） | **常驻 helper 控制**（Word/Excel/PPT 三端都有，v1.15 起）：start 预热/查看管道、status 报 `running`/`pipe`、stop 让它答完即退。正常流程**不用手动调**——`open` 会自动拉起、`close` 关完文档自动退出；设 `WOC_NO_DAEMON=1` 可整轮退回一次性模式（也是 A/B 用的开关）。设计见 `reference/workflows.md` §15 |
| `textdiag` | `-Doc <路径> [-Slow]` | **读侧慢/回退的体检**（只读，一次调用）：① 段落数/修订数/表数/表行总数；② `Content.Text` 的块数与尾块；③ 抽样段落号在块模型下的文本 vs 真值；④ 表格行数 vs 段落数的关系（WPS 每行多算一个空段落——块模型对不上的根因）；⑤ `WordOpenXML` 是否可取、字符数、`w:p` 数、`w:p`+表行是否等于段落数、解析耗时；⑥ 最终 `fastPath`/`fastFail` 结论。**遇到「按说几秒却几十秒」先跑它** |
| `captions` | `-Doc <路径> [-Label 表\|图] [-Max 200]` | **题注清单 + 编号体检**（只读，一条命令替掉"全篇扫段落配对"）：每张表 → 题注段号 `para` / 位置 `pos`（`above` 表上方、`in-table` 做在第 1 行、`below` 在表下方）/ 表号 `number` / 格式 `format`（`dot` 2.4.1、`hyphen` 3.3-1、`plain` 表1）/ 空白 `gap` / `title`；`summary` 汇总重号、同节跳号（`missing` 列出缺的号）、格式混用、空格不一、题注位置混用、缺题注、正文引用与表号的对应（`references` 里 `resolved`）+ 悬空引用；`issues` 每条给 `rule`+`para`+`sample`+`suggest`，直接可照着改。实测 10521 段 / 32 表 **6.4 s**；编号问题**不要**用 `audit` 或全篇 `read` 去查 |
| `caption` | `-Para N -Label 表\|图 -Title "..." [-Above] [-Align center]` | 生成**域**式题注「表 1 xxx」（自动编号，增删后重排） |
| `picture` | 插图：`-Path <图片> -Para N [-Width "12cm"] [-Caption] [-Align]`；调整：`-Image n [-Width/-Height] [-CropLeft/-CropRight/-CropTop/-CropBottom 百分比]` | 插入图片（可带图题）/ 调尺寸与裁剪 |
| `toc` | `-Doc <路径> [-Para N] [-Title 目录] [-Levels 3]` | 插入并刷新目录（域），可从文首或指定段之前 |
| `reference` | 列表：`-Action list -Items <json 数组> [-AtPara N] [-Heading 参考文献]`；引用：`-Action cite -Para N -Numbers "1,2,5"\|"3-6"` | 生成参考文献列表 / 在段末插入上标引用 |
| `view` | `-Doc <路径> -Para N` 或 `-Page N` | 把 WPS 窗口滚到那一段/那一页（唯一能做的"跟随"，不动文档内容） |
| `unit` / `units` | `-Doc <路径> [-Apply] [-Units "m3,km2,s2"] [-Normalize] [-Track]` | **单位上标批量修**：把掉下来的 `m3/s`、`km2`、`万m3` 里的指数改成 Unicode 上标（`m³/s`、`km²`）。不带 `-Apply` = 只体检，带 `-Apply` = 修并报逐单位处数与跳过项。**只修"挂在数值上"的单位**（前字符＝数字/万/亿/汉字/左括号，或"数字 空格 单位"）；**已正确的两种形态自动跳过**（Unicode ²/³ 与 run 级上标）；**英文/数字贴连跳过**（`item3`、`m23`、`km20`）。三类只报不改：`fragments`（孤立 token／后跟 `=`，多半是模板变量名）、`looseFragments`（`n2`/`n3` 缺基名）、`suspects`（`km3` 这类疑似笔误）。300 页/13k 段实测 ~20 s |
| `audit` | `-Doc <路径> [-Max N] [-Action 规则名,逗号分隔] [-Slow]` | **机械检查**：单位未用上标、重复词、括号混用、连续空格、全角数字、段末空格、标题层级跳级、图表缺题注、连续空段、**题名做进表格第 1 行**（`table-caption-in-row`，直接给出该用哪条 `table -Action caption`）；返回段落号 + 建议，供你判断后落批注。**v1.9 起默认走快路径**（一次取全文+本地规则+Find 定位段号，几千段文档从几十秒降到几秒；有未处理修订或定位校验失败时自动整体回退旧逐段扫描，`path` 字段说明走了哪条；`-Slow` 强制旧路径）。题注判据与 `captions` 共用同一个解析器（**标签必须在段首**，"如表3.3-2所示"这类正文引用不算题注），题名在表内时只报一条、不再同时报 `caption-missing`；**编号本身的问题（重号/跳号/格式混用/悬空引用）不在这里查，用 `captions`** |
| `fields` | `-Doc <路径>` | **刷新全部域**：页码、目录、题注、引用（改完正文、交付前跑一次） |
| `cell` | `-Doc <路径> (-Table n -Row r -Col c \| -Para N) -Text "..."` + 任意格式项 | 改一个表格单元格的文字（可带字体字号对齐） |
| `insert` | `-Doc <路径> [-Para N] [-Where after\|before\|end] -Items <json> [-Track]` | **批量插入段落**（默认干净写入、不进修订；`-Track` 才走修订），可带样式与格式。**锚点是表格前一段时自动改用"段末标记之前插入 CR"的写法**，不会掉进表格第一格。**v1.30：新段落默认整词换行**（`wordWrap:["noSplitWords=true (readBack=-1)"]`）；要允许拆词给 `"noSplitWords":false` 或 `format -SplitWords` |
| `revisions` | `-Doc <路径> [-Action list\|accept\|reject] [-Index N \| -All]` | 列出**未处理**的修订（类型/作者/所在段/文本），或按人要求接受/拒绝 |
| `snapshot` | `-Doc <路径> [-Tag 名字] [-Out <路径>] [-Slow]` | 存一份段落文本快照（给 `diff` 用，默认存 %LOCALAPPDATA% 不污染文档目录；`-Out` 写到指定路径，跨机器/交接时随文档走）。**v1.10 起带表文档也走快路径**（一次取全文；4479 段实测 22.4 s → 3.2 s，`textsPath` 字段说明走了哪条）。**v1.13 起改走 `WordOpenXML`**：7334 段真实报告 **48.5 s → 6.1 s**，且**带未处理修订也能走快路径**（旧块模型遇到修订必然回退） |
| `diff` | `-Doc <路径> [-Since 快照路径\|last] [-Slow]` | 和快照比：哪些段被改了（before→after）、哪些增删了——**人改的和 Agent 改的都算**（当前侧同样走快路径），v1.13 起 7334 段实测 48.6 s → 5.5 s |
| `handoff` | `-Doc <路径> [-Note "本轮做了什么"] [-As 输出路径]` | 在文档旁生成 `.handoff.md` 交接单：状态/待办/操作记录/最新快照路径 |
| `undo` | `-Doc <路径> [-Count N]`（默认 1，上限 20） | **Agent 自己的 Ctrl+Z**：撤回最近 N 步（排版批是一步）。注意 Undo 不分作者——只在自己刚写完且无 `changedOutside` 时用 |
| `save` | `-Doc <路径> [-As <另存路径>]` | 保存 / 另存 |
| `close` | `-Doc <路径> [-Save]` | 关闭；无其它文档时退出实例，并回报进程数 |
| `quit` | — | 关掉该实例的**全部**文档并退出（不保存）。**慎用**：会波及用户在同一实例里开的文档，只在自己的独占实例上用 |
| `finish` | `[-Note "做了什么"] [-Toast] [-Sound] [-Quiet]` | **任务完成信号**（agent 无关，零依赖）：默认**托盘气泡 + 提示音**（原生 .NET），**不触碰文档任何内容**；`-Quiet` 静音。**任务收尾默认调用** |
| `hint` | `-Doc <路径> [-Text "正在做什么"] [-HintTag "〔AI接管中〕"] [-Bar "..."] [-Toast] [-Sound] [-Quiet] [-Clear] [-Tray]` | **接管提示**（v1.11 / v1.12 加 `-Tray`）：改窗口标题 + 写状态栏一行 + 默认弹托盘气泡，**三者都是窗口状态，不写文档、不落盘**。`open` 已自动打标签；`-Text` 换气泡/状态栏里的话（如"正在核对第 3 节"）；`-Clear` 还原（`finish`/`close` 也会自动还原）。Writer 的 `Application.StatusBar` 只写不读、且 WPS 界面不渲染（实测状态栏仍是页面/字数），`statusBar:true` 只代表写入被接受；表格侧写入会真的显示。**`-Tray`**（v1.12）：常驻托盘图标，悬停=标记+当前步骤，右键可还原，`finish`/`close`/`-Clear` 自动移除 |

所有命令都可用 `-TimeoutSec <秒>`（默认 180）控制等待；超时返回 `TIMEOUT`，说明 WPS 大概率弹了对话框。

### `revise` 的 items 格式

```json
{"items":[
  {"para":12,"expect":"3.2 工程地质","op":"replace","find":"孔口高程","replace":"孔口地面高程"},
  {"para":15,"expect":"根据地质勘察","op":"set_text","text":"根据地质勘察成果，坝址区基岩为……"},
  {"para":20,"expect":"4.1 施工组织","op":"insert_after","text":"4.1.1 施工条件"},
  {"para":31,"expect":"本阶段不再","op":"delete"}
]}
```

- `op` 四种：`replace`（字符级，**优先用**）、`set_text`（整段删+插）、`insert_after`、`delete`。
- `expect` 必填，校验该段文本是否**以它开头**，不符即 `ANCHOR_MISSING`（防改错位置）；`replace` 的 `find` 找不到也报同一码。
- `replace` 默认只替换该段首个命中；同一段内要多处替换，加 `"all":true`。
- 批内自动按 `para` **从大到小**执行，编号不会互相串；要插多段时按原始编号一次写全。
- **整批是原子的**：动手前先把每一条都核一遍（段落存在、`expect` 对得上、`find` 找得到），有一条不过就**一条都不改**并报 `ANCHOR_MISSING`；万一执行中途出错，已产生的修订会被逐条 `Reject` 回滚，错误信息里会写"rolled back N revision(s)"。
- 每条结果带 `textAfter`（该段改完后的开头 90 字）——**立刻比对它**，能第一时间发现"替换后变成「进行了了」"这类找错位置的问题，不必再 `read`。
- 返回 `applied`、`revisionsBefore`、`revisionsAfter`、`revisionsAdded`。
- **段落编号的坑**：表格里每个单元格算一段，且每行末尾**再多算一段**（3×4 表占 15 个段落号），所以表格附近的段落号不能按"行×列"推算；`read` 会给表内段落打 `inTable: true` 与 `table: n`。找段号优先用 `search`（v1.8 起段号经字符区间校验）；**表格与题注的段号**直接用 `captions`/`table -Action scan` 的 `para`，别盲试。

## Word 格式项（`format` / `style` / `table` / `caption` / `cell` 通用）

| 参数 | 取值 | 说明 |
|---|---|---|
| `-Font` | `仿宋` | 中英文一起换（`NameFarEast`/`Name`/`NameAscii`/`NameOther`） |
| `-LatinFont` | `"Times New Roman"` | 只换西文，不动中文字体 |
| `-Size` | `12` / `"10.5pt"` | 字号（磅） |
| `-Bold` `-Italic` `-Underline` | 开关 | 关掉用 `-Bold:$false` |
| `-Super` `-Sub` | 开关 | 上标 / 下标 |
| `-Color` | `"#C00000"` | 字体颜色 |
| `-Align` | `left\|center\|right\|justify` | 对齐（中文也认 `居左/居中/居右/两端对齐`） |
| `-IndentChars` | `2` | **左缩进 2 字符**（中文单位，比磅值可靠） |
| `-LeftIndent` | `"0.74cm"` / `"21pt"` | 左缩进（长度） |
| `-FirstLineIndent` | `2` / `"0.74cm"` / 0 | **首行缩进**：纯数字=字符数（2=两字），带单位=长度，0=取消 |
| `-LineSpacing` | `1.5` / `"20pt"` / `"at least 20pt"` / `single` / `double` | 数字=倍数，`pt`=固定值，`at least`=最小值 |
| `-SpaceBefore` `-SpaceAfter` | `"6pt"` / `0` | 段前 / 段后间距 |
| `-PageBreakBefore` | 开关 | 该段另起一页 |
| `-Track` | 开关 | 让**格式改动也进修订**（默认不进；整篇排版不要开） |
| `-NoSplitWords` | 开关 | **整词换行**（v1.30 修正极性）：西文单词/长数字不在词中断行。实测 `WordWrap=-1` 才是整词、`0` 是允许拆词——v1.29 及更早写反了（传 `-NoSplitWords` 实际把拆词打开）。现在我们写入的新段落/表格单元格/题注**默认就是整词**，一般不用传 |
| `-SplitWords` | 开关 | **允许拆词**（`-NoSplitWords` 的反向开关，v1.30 新增）：`-NoSplitWords:$false` 过不了 `woc.ps1` 的参数往返（会变成字符串 `"false"` 绑不上 switch），要放开拆词用这个 |

## Word 典型工作流（详见 reference/workflows.md）

1. **读+核对版式**：`status` → `open -Show` → `read` → `render`（看表格出框/图表贴题/空段半空页）。
2. **定位**：`search -Text` 拿段落号与上下文（全文找词，取代一页页 `read` 翻）；用户说"就改我选中的那段"时用 `selection` 直接拿他选中的段落号。
3. **润色**：`search`/`read` 拿锚点 → 造 items（`replace` 优先）→ `revise` → 比对 `textAfter` → `render` 抽查 → 用户看过再 `save`。
4. **校审批注**：`audit` 机械检查 → `read` 被点名段自己判断 → `comment -Items` 一次落批注 → `render` 复核气泡 → `save`。
5. **统一字体**：`font -Name 仿宋`（看 `readBack` 回读是否落上）。
6. **一键排版**：先 `style` 后 `format`，再 `page`；排版前后各 `render`；别开 `-Track`。
7. **表格/题注/图片**：`table`（三线表+表头跨页）/ `caption`（SEQ 域，表上图下）/ `picture`（插后必 render）。
8. **目录/引用**：`toc`（依赖标题样式）/ `reference list+cite` / 交付前 `fields` 刷全部域。
9. **协作接力（核心用法）**：每轮 `snapshot` → 干活 → 人裁决 → `diff`+`revisions` 读结果 → `handoff` 交接单。跨机器接力用 `snapshot -Out` 把快照随文档带走。写命令结果里的 `changedOutside` = 文档被人动过，先 `diff` 再继续。
10. **表格编号/题注体检（先查再改）**：`captions -Doc X` 一条命令出清单 + 判定，**不要**用 `audit`/全篇 `read` 去凑（实测差 40 倍）。
    - 先看 `summary`：`duplicates`（重号）、`sequenceGaps`（同节跳号，`missing` 给缺的号）、`formats`（点分/短横混用）、`spaces`（表号与题名之间空格不一）、`positions`（表内/表外混用）、`uncaptioned`（缺题注）、`unresolvedReferences`（悬空引用）。
    - 再看 `issues[].suggest` 照改：改号用 `revise`（题注段号就在 `issues[].para`）；表格改动本身用 `table -Action caption` / `-Action rows`；改完 `fields` 刷域式题注。
    - `tables[].para` 是每张表的起始段号，可以直接 `read -From <para-3> -Count 6` 看题注上下文。
11. **批量整理表格（中文报告高频）——v1.33 起一条命令**：**先 `table -Action plan`，再 `table -Action normalize`**，最后 `render -Table n` 看一张表级图后 `save`。
    - `plan`（只读，一条命令给全）：每表题注位置/页码、边框是否达标、`tblStyle`/`tblW`/`gridCols`/列宽/字号、`risks[]`（表内题名 `captionInFirstRow`、边框未达标 `bordersNotThick`、表前断点 `breakBefore`、无样式 `noStyle`、无 `tblBorders` `noBorders`、宽度是百分比 `pctWidths`、列宽超版心 `gridOverflow`、跨页缺表头）+ 题注清单/正文表号引用/重编号提议。**这些字段以前只能自己解 XML 才看到，别再去解密读 XML**。
    - 编号语义要人拍板：`renumber.items[]` 给的是"按出现顺序+引用指向"的**候选**映射，`duplicates`/`ambiguous` 非空时不要自动改；人确认后把 `approved:true` 的 plan 文件喂 `normalize -Plan <file>`。
    - `normalize`（默认只改 plan 判定未达标的表）：题注外提（**尽力而为、不承诺**：回读「题注页码 == 表起始页」通过才保留；普通段落后、手动分页符前的表通常可提；**表前是分节符时 WPS 无法在断点后造正文段 → 报 `needsHuman`，不重试**）→ 题注空白统一（默认 `-Gap 2`，只统一「编号与题名」之间，不擅改「表2.1.1」这种标签-编号间距）→ 可选重编号 → 加粗边框（WordOpenXML 列宽快照/恢复，本来达标的表跳过）→ 跨页表头 → 自检；**逐表隔离**（一张失败不回滚其它）、整条命令一个撤销单元、**绝不隐式 save**。先 `-DryRun` 看一眼要改哪些表。
    - 复核：`render -Table n`（只渲该表所跨页，不再整页翻）+ `plan` 再跑一次比 before/after；`verification` 里 `captionInFirstRow`/`bordersNotThick`/`multipageWithoutHeaderRepeat`/`captionPageMismatch`/`widthsChanged` 都应归零。
    - 只有"单点小活"（就一张表、只提题名）才用逐条命令：`table -Action caption -Table n -Borders thick` / `-Action header-repeat` / `-Action rows`。
    - 细节、失败码与实测数字见 [reference/workflows.md](reference/workflows.md) 的「表格整备（v1.33）」一节。
12. **单位上标批量修（几百页报告的常见活）**：`unit -Doc X` 先看体量 → `unit -Doc X -Apply` 一键修（已正确的跳过）→ `render` 抽 1–2 页确认 → `save`。
    - 判据：`unit` 的 `plainDigitHits` − `unicodeOk` 不等于坏例数，**坏例数以 `-Apply` 返回的 `fixed` 为准**（只有 Find 能看到 run 格式，文本层分不出"真掉"和"格式上标"）。
    - `skippedSamples` 里若出现疑似真单位，用 `-Units` 补上重跑；若是英文词（`item3`、`com2`）则不用管。
    - **`scriptPages` 先看**：命中脚本页（公文模板/脚本说明）时命令已整段跳过，若它恰好是你要修的那几页，只能人工处理（那里的 `m2`/`km2` 是变量名）。
    - **`fragments` / `looseFragments` / `suspects` 三样都必须看一遍**：`fragments` 多半是**模板/脚本示例里的变量名**（`km2 = 1`、`<m1，m2，m3>`），`looseFragments` 是缺基名的形态（`n2`、`n3`，常是 `m2`/`m3` 打错），`suspects` 是疑似笔误（面积表里的 `km3` 多半该是 `km2`）。**这三类一律只报不改**——自动改会把变量名/笔误变成"看起来是对的"上标，比不改更糟。报给作者确认后再用 `revise`/`cell` 处理。
    - 若 `fixed` 明显少于预期（例如只剩个位数），先看 `skipped.assignment/isolated`：说明命中多半在模板示例或表格标题里，符合预期。
    - 全文见 [reference/workflows.md](reference/workflows.md) 的「单位上标批量修」一节。
13. **反悔**：自己刚写完且无 `changedOutside` 时 `undo -Count N`（排版批=1 步）。
14. **写公式（LaTeX，v1.29）**：`read` 定位段号 → `formula -Para N -Expect "<段首>" -Latex "..."`（行内式）；要编号的独立公式用 `-Display -Number "3"`。**含 `^` `{}` `\` 的 LaTeX 一定用单引号**，或整批走 `-Items <json>`。插完 `read` 一眼（公式在 `text` 里显示为数学字符，另有 `omaths` 计数）→ `render` 抽一页看分式/根号/上下标是否成形。支持的子集与不支持清单见 [reference/word-quirks.md](reference/word-quirks.md) 末节——**不支持的命令会直接报 `MATH_UNSUPPORTED` 并点名**，别反复试。
15. **加脚注（v1.29）**：`footnote -Action insert -Para N -Expect "<段首>" [-Find "<锚词>" -Where after] -TextFile <txt>`（中文正文走文件）→ `footnote -Action list` 拿 `i` → 要改 `edit -Index i -Expect "<原文开头>"`，要删 `delete -Index i`。**不要**用"上标数字 + 页脚手写"伪造脚注：那不是原生脚注，增删后编号必错、也不会随页面重排。
16. **交付前过全稿（v1.29）**：`render -Overview -Page 1 -To 18`（≤18 页/次，一张图 6 页）→ 看总览图挑可疑页 → `render -Page n -Width 1100` 放大 → 再 `save`。
17. **排版收尾：填留白（v1.30）**：全稿写完、图表都插好之后跑一次 `fillgaps -DryRun` 看它会搬哪几段（返回 `moves[].gap` 是那一页底部的空白大小）→ 没意见就 `fillgaps` 实跑 → `render` 看 `moves[].page` 那几页。表格跨页被劈开、行落单时先 `table -KeepTogether -Table n`（它会逐行回读，`ok:false` 就是 WPS 没接受）。**注意**：`fillgaps` 只搬正文段、绝不越过标题；表格前会留一个空段；`moved:0` 说明这稿没有值得填的留白，不是失败。

## Word 长任务的上下文纪律（v1.33，2026-11-03 实测）

> 背景：一份 29 页/14 表的报告做「题注编号 + 表名外提 + 加边框」，旧做法用了 **127 次工具调用 / 88.6 分钟**。拆开看：**工具只跑了 8.8 分钟（10%），79.7 分钟（90%）在等模型**；其中 57.8 分钟是每步重读 21.8 万 token 上下文的**首字延迟**（prompt <150k 时 TTFT 4.6 s，>300k 时 60 s，峰值 94 s；相关系数 0.83）。**真正贵的是「步数 × 每步上下文」，不是 COM 慢。** 文档涨 1 万 token，后面每一步都变慢。
> 纪律比命令更重要：把 127 次调用里 56 次（42.2 分钟）的试错绕路消掉，比把命令改快 2 倍更值钱。下面每条都带本次实测代价，违反哪条会在复盘里被点名。

1. **先 `table -Action plan` 再动手**，plan 的结论直接驱动后续入参。本次 13:20 的 `scan` 已报出「6 张边框未达标、T7 内线 9999999、T9 题名在首行」，13:46 仍 14 张一起改——"侦察结论没被消费"代价 **38.2 分钟（43% 墙钟）**。
2. **别读 helper 源码猜参数**。本次 31 次 `Select-String/read/grep/Get-ChildItem` 打在 `.cs`/`word.ps1` 上 = **15.4 分钟**；答案本来就在命令返回码里（`ANCHOR_MISSING` 等错误已写清原因与出路）。要新能力就提需求改命令，不要在会话里考古。
3. **别为拿列宽/样式/分节符自己解 XML**。本次 3 次 `decrypt-files.ps1`（92.3 s，占工具时间 17%）+ 4 个自写脚本 = **26.4 分钟**；这些现在是 `plan` 的返回字段（`gridCols`/`tblW`/`tblStyle`/`risks[].breakBefore`）。
4. **`E_FAIL` 出现一次就换机制，不要试第三种写法**。表 2.1.9 题注外提失败后连试 5 条路径 = **14.7 分钟**；正确动作是查 plan 的 `breakBefore`/`captionMoveRisk`，能外提就走 `normalize`，不能就报工单。
5. **别全量下笔，只改"未达标清单"**。本次只需改 6 张却改了 14 张 → 无法定位/无法单表回退 → 整体回滚重做，**37.6 分钟排查 + 17.9 分钟作废**。`normalize` 默认就只改未达标表，别再手工 `-Borders thick` 一把梭。
6. **别反复整页 render 看同一页**。本次 18 次 `read_image`、`page-17` 看了 7 次 ≈ **8.4 分钟**；现在用 `render -Table n` 只出该表那几页，一次一张图看完就够。
7. **回滚用「close 不保存 + 重开」或单表快照，不要叠 `undo`**。本次 4 次 `undo` 后原稿 2 条批注丢了 1 条、段落数 4069→4066；`undo` 不分作者。
8. **返回体与 reasoning 都会进下一次 prompt**：本次工具结果占上下文增量 46.8%、自己的 reasoning 回灌 41.4%。所以：命令返回要精简（必要时 `-DryRun` 先看摘要）、别写长推理、同一事实别重复查。

## Word 故障排查

| 错误码 / 现象 | 怎么办 |
|---|---|
| `NO_APP` | 没装 WPS 或 COM 没注册。先手动打开一次 WPS Writer 再重试；`doctor` 里 `apps.wps.installed=false` 说明这台机器确实没有 WPS。 |
| `NOT_OPEN` | `-Doc` 与实际打开的文档不是同一路径（另存过/短路径/大小写/相对路径）。先 `status`，把 `path` 原样复制过来。 |
| `READONLY` | 文件被别处占用或是只读属性。关掉占用它的程序，或请用户另存一份再改；不要强行覆盖保存。 |
| `ANCHOR_MISSING` | 段落编号漂移（表格也占段落号，且每行多一段）或 `expect`/`find` 与原文不符。整批**一条都没改**，先 `search`/`read` 该段核对编号与原文再来；并发协作时先怀疑"人刚改了这段"。 |
| `BUSY` / `TIMEOUT` | 用户正在打字、或有对话框/弹窗挡着。**v1.32 起 helper 装了 OLE 消息过滤器**：WPS 拒绝调用时会以 **150 ms 间隔自动重试、最多 6 s**，仍不行才报 `BUSY`（可重试）——所以看到 `BUSY` 说明"已经等过 6 秒还忙"，等用户停手再发，别连发。`TIMEOUT` 是等你设的 `-TimeoutSec`。 |
| `COM_ERROR: RPC 服务器不可用 (0x800706BA)`（v1.29 实测） | 有人**在 daemon 还活着时把 WPS 进程杀了**（或 WPS 崩溃），常驻 helper 握着的是死引用，之后每条命令都报这个 | 三端各跑一次 `serve -Action stop` 把 daemon 停掉，下一条命令会重新拉起并重连；**不要**反复重试同一条命令 |
| 打开即 `COM_ERROR` / 文件读不出来 | 多半是密文且不归 WPS 管（例如离线脚本写出的）。让用户在 WPS 里打开一次（驱动会透明解密），或按 [reference/encryption.md](reference/encryption.md) 先解密。注意：helper **不产出** `ENCRYPTED` 码。 |
| `open` 返回 `multiInstanceWarning` | 文档可能同时开在另一个 WPS 实例里（本 skill 只能接到第一个注册实例）。请用户在 WPS 里关掉另一份再继续，否则两份会互相覆盖。 |
| 结果里出现 `reopened` | 文档刚才被人在 WPS 里关掉了，本命令已从磁盘**静默重开**（未弹窗）。若同时有 `unsavedAtClose: true`，说明关闭时未保存——关闭前的改动（含本 skill 的修订/批注）**可能已丢失**：先 `diff` 快照对账、向用户确认后再继续。 |
| `MATH_UNSUPPORTED` + 命令名（v1.29） | 该 LaTeX 用了 `Omml.cs` 子集外的命令（`\substack`、`\mathop`、`\stackrel`、`\overset/\underset`、`\boxed`、`\begin{array}{cc}` 的列格式串等） | 换写法：矩阵用 `matrix/pmatrix/cases` 且**不写列格式**；分段函数用 `cases`；实在要 `array` 就改成手写多行公式。报告类公式优先 ASCII，中文只在 `\text{}` 里且字体不保证 |
| `MATH_FAILED`（"inserted the XML but built no equation"） | WPS 没把 Flat OPC 包认成公式：多发生在域代码/文本框/特殊段落上下文里 | 命令已按 `Content.End` 增量回滚，文档没脏。换段落下手，或先插一个空段再往空段里 `formula`；仍失败就退回 Unicode 上标文字并告知用户 |
| 脚注 `ANCHOR_MISMATCH`（v1.29） | `-Index` 指向的脚注正文与 `-Expect` 不符（前面插/删过脚注，序号已移位） | 先 `footnote -Action list` 重取 `i` 再改，**不要**跳过后硬重试 |
| 脚注里要写公式（v1.29） | 目前不支持（返回 `BAD_ARGS` 明说） | 公式放正文用 `formula`；脚注里用 Unicode（`x²`、`α ≤ β`） |
| `KEEPTOGETHER_FAILED`（v1.30） | `table -KeepTogether` 逐行回读时 WPS 没接受某些行的 `KeepWithNext`/`AllowBreakAcrossPages`（合并单元格、复杂行高、表格在文本框里等） | 看 `error.detail.readBackFailures` / `.allowBreak.readBackFailures` 点名的行号；这几行改用"调段前文字让表格整块落下"或先取消合并单元格再跑。**不要**重跑同一条命令期待不同结果 |
| `fillgaps` 返回 `moved:0`（v1.30） | 正常结果：这稿没有"图表被挤到下一页、上一页底部还空着"的情形（空白 < `max(90pt,12%)`，或图表本来就在同页） | 不需要处理。若你**肉眼看到**明显空白却没搬，用 `-MinGapPt 60` 再试一次；仍为 0 就把那页 `render` 出来给人看 |
| `fillgaps` 搬完表格上方多出一个空段（v1.30，已知代价） | 表格块的搬运必须在表格前一段的段末标记**之前**插一个 `\r` 才写得进去（直接写 `table.Range.Start` 会落进第一个单元格） | 需要极干净版面时，可先 `fillgaps` 再 `table -Action rows`/手工删掉那个空段，或接受它（视觉上是一行空白） |
| `render -Overview` 只给 1–3 张图且 `capped:true` | 一次最多 18 页（3 张 × 6 页/张） | 接着 `-Page 19 -Overview` 看下一页组；**不要**一次要 60 页（图会大且糊） |
| 中文乱码 | 控制台代码页不是 UTF-8。先 `[Console]::OutputEncoding=[Text.Encoding]::UTF8`，或把 JSON 落成文件按 UTF-8 读。 |
| `wps.exe` 变多 | 中途异常没 `close`，或别处 `New-Object` 开了新实例。先 `status` 看实例数，只结束无窗口的 `/Automation` 实例，用户窗口留着，然后汇报数字。 |
| Ctrl+Z 撤不干净批量改动 | 三个前提：①改动必须走本 skill 的命令（自带 UndoRecord 分组，实测 WPS 支持；**绕过 skill 手写 COM 循环 = 每段一条撤销记录**）；②**save 会清空撤销栈**，要反悔必须在保存前；③格式套成与原值相同的段不产生撤销记录（再撤也不会变）。 |
| `undo` 撤不动 | 撤销栈已被 save 清空或已见底；文本类改动优先走修订（`revise`），随时可拒绝，不依赖撤销栈。 |
| 每条命令都要等 5–8 s | v1.14 起 Word 侧默认走常驻 helper（`serve`），热命令 ~2.6 s；若返回很慢，先 `serve -Action status` 看 `running`：false 说明常驻没起来（多半被 `WOC_NO_DAEMON=1` 关掉、或上次 `close` 后没再 `open`），再跑一次 `open` 即可重新预热。注意**剩下的 ~2 s 是调用方 PowerShell 启动**，skill 侧压不掉 |
| `search` 给的段号看着对、改下去错位 | v1.8 起段号按**字符区间**校验（结果带 `verified:true`），同一段多个命中各有自己的 `at`；若 `atFallbacks` 偏大，说明修订遮挡了文本层偏移（已按段内递推纠正）。仍怀疑时：`read -From <para-1> -Count 3` 核对该段原文，并用 `revise` 的 `expect` 兜底。**别**再据一次 `search` 就断言"缺题注"——用 `captions` 复核。 |
| 题注重号/跳号查不出来 | `audit` 只查"缺题注/题名在框内"，不查编号体系；`search 表` 只能给命中，不能配对。跑 `captions`：重号、同节跳号、格式混用、悬空引用一次全出。 |
| `captions` 漏了某条题注 | 解析器要求**标签在段首**（`表` / `附表` / `图` / `附图` + 数字）。若题注写在表格第 2 行、前面加了"续表"、或题注是插入的域而不是文字，会被漏掉：`read -From <tables[].para-3> -Count 6` 看清楚后手工处理。 |
| `inTable` 与所见不符 | 已修（`Table.Range.End` 是排他端，判据改 `start < end`）；回归用例在 `tests/word/captions/check_captions.py` 的 B1 两条断言。若仍见到紧随表格的正文段带 `inTable:true`，先 `undo` 再报告。 |
| `table` 报 `0x800A1767` / 行操作炸 | WPS 的 `Row` 对象残缺（`Rows.Item(n).Delete/.HeadingFormat/.Range/.Cells` 都不可靠）。**不要自己写 `Rows.Item(n)` 循环**——用 `table -Action rows/header-repeat/caption`，内部走"选中整行 → `Selection.Rows`"的写法 |
| `insert -Where after` 的段落进了表格 | 已修：锚点后紧跟表格时，helper 改在段末标记之前插 CR。若仍见到 `inTable:true` 而锚点不在表格里，先 `undo` 再报告 |
| **表名外提：本 skill 不承诺**（v1.33 结论） | 把「做在表格第 1 行的题名」提到框外，**WPS 上不可靠**：①表前是**分节符**时（题名行在横向/新节首），分节符的段末标记紧邻表格，任何"在表前插段落"都会落回上一节、题注留上一页——WPS 没有可用原语，`normalize` 直接报 `needsHuman{captionMoved:"wps-section-start"}`，**不重试**；②「删表 + InsertXML 放回」的 cut-paste 在 WPS 上会把 Writer **转死**（100% CPU、COM 全挂，2026-11-03 实测），只在微软 Word 或显式 `-CutPaste` 时尝试；③能可靠外提的只有普通段落后/手动分页符前的表（回读题注页 == 表页才算成功）。**要彻底解决请人工把分节符剪到题注之后，或在微软 Word 里处理。** |
| 加边框后**列宽/版式崩**（v1.33 实测根因） | 触发条件不是"没样式"或"宽度是百分比"，而是**该表 XML 里没有 `w:tblBorders` 元素**：第一次写边框会让 WPS 重建整表布局、按渲染宽度重推 `gridCol`（无 `tblStyle` + dxa 固定宽度的表会直接崩）。`plan` 的 `risks[].noBorders`/`noStyle`/`gridOverflow` 会预判；`normalize` 写前快照、写后回读、漂移即恢复，本来达标的表跳过 |
| 想读列宽/字号/表样式，`table -Action format` 回的全是 `9999999` | WPS 对 autofit 列 `Column.Width` 返回 `wdUndefined`（即 9999999），COM 侧读不出真实列宽。用 `table -Action plan` 的 `gridCols`（来自一次 `Content.WordOpenXML`，整篇约 0.16 s），**不要**自己解密读 XML |
| `doctor` 显示 Word 16 但实际操作行为像 WPS / `version` 只有 12.0 | 本机 `Word.Application` 的官方 CLSID 被**用户级注册表覆盖**到 `wps.exe`（`HKCU\Software\Classes\WOW6432Node\CLSID\{000209FF-…}\LocalServer32`），`CurVer=Word.Application.12`；两家 `Application.Name` 都返回 "Microsoft Word"，**判据只看 `doctor`/`status` 的 `app.path` + `app.build`** |

---

# Excel：在 WPS 表格里实时改工作簿

命令前缀：`office.ps1 excel <命令>`（或直接 `excel.ps1 <命令>`）。

## Excel 命令速查

| 命令 | 参数 | 作用 |
|---|---|---|
| `doctor` | — | 自检：ProgID 安装/在跑、进程数、渲染能力、打开的工作簿 |
| `status` | — | 在跑的实例与工作簿清单（name / path / saved / readOnly / active），工作表概况 |
| `open` | `-Doc <路径> [-Show] [-ReadOnly] [-NoHint]` | 已开则复用，未开则打开；返回工作表清单与各自 usedRange。**v1.11 起自动给窗口标签打「〔AI接管中〕」接管标记**（`hint.caption`；用 `wb.Windows(1).Caption`，WPS 表格里 `wb.ActiveWindow` 常为 null） |
| `new` | `-Doc <路径> [-Show] [-Overwrite] [-NoHint]` | **新建空白工作簿并落盘**（v1.16）：`Workbooks.Add` + `SaveAs`（.xlsx/.xls/.csv/.xlsm 按扩展名选格式），已打开并打接管标记。文件已存在时拒绝（`EXISTS`），`-Overwrite` 才替换；**目标正开在 WPS 里则拒绝**（`IN_USE`，先 `close`）。`open` 对不存在的路径报 `NOT_OPEN`，要造新表用这条，**不要**再自己写 openpyxl |
| `read` | `-Doc` `[(-Sheet 名 \| -Range "Sheet1!A1:F40")] [-Profile]` | 不带选项=工作簿概览；带 Sheet=读该表 used range（值 + 公式）；`>60 行或 >15 列`自动改列概况，要局部全量用 `-Range`。**v1.29 概况增强**：每列多了 `letter/header`、`numbers/dates/texts/blanks/errors`、`dateFrom/dateTo`、`farAbove`（离群 q3+3·IQR、命中 ≤5% 才报）、`distinct/topValues/repeats`，以及 **`look:[…]`**（`n with stray spaces`、`n number(s) written as text`、`n date(s) written as text`、`mixed numbers and text`、`n error value(s)`）；顶层加 `dataRows/emptyRows/duplicateRows`；**日期列按 `NumberFormat` 判定、值直接从序列号转成 `yyyy-MM-dd`**；错误值改成 `#DIV/0!`/`#N/A`/`#REF!` 等可读码。**拿到概况先看 `look`，再决定跑哪条 `clean`** |
| `write` | `-Doc -Items <json>`（或单格 `-Cell B3 -Value x / -Formula "=..."`） | **批量写**（见下）：原子批、`expect` 锚校验、中途失败回滚，每格返回 before→after；**批内自动挂起重算+屏幕刷新，批末一次重算**（大表写几千格不产生重算风暴） |
| `fill` | `-Doc -Range D2:D5000 [-Sheet] -Formula "=B2*C2"` | **整列填公式**：首格写一次 + 原生 FillDown（2000 行约 2 秒，含进程启动）；比逐格 write 快几个量级，大表必用。**v1.29 起返回 `errors:[{cell,code}]`**（写完即告 `#DIV/0!` 这类错误格，不用等用户发现） |
| `replace` | `-Find "旧" [-Replace "新"]`（省略 = 删除）`[-Sheet] [-Range A1:F200] [-Column 表头\|字母] [-Whole] [-MatchCase] [-Regex] [-DryRun] [-Max N] [-All]` | **内存块替换**（v1.31）：一次块读 → 本地逐格替换 → 只写真正变化的格，**绝不走 `Range.Replace`**（WPS 会在刚写进去的文本上二次替换，`P→P-` 得 `P--`；我们单遍扫描，实测 `PP→P-P-`）。返回 `scanned/changed/byColumn/samples/formulaCellsSkipped/numericMatchesLeftAlone/truncated` + `preSnapshot`（可 `diff -Since`/`restore -Since`）。**公式格永不写**；**数字/日期格不改成文本**（`-Find 0` 想改数字列会给 `changed:0` + `numericMatchesLeftAlone:N`）；`-DryRun` 与实跑同一套判据；规则不自匹配时幂等（第二次 `changed:0`） |
| `clean` | `-Column 表头\|C`（或 `-Columns "A,C"`）`[-To number\|date\|text] [-Casing upper\|lower\|proper] [-Remove "元,["] [-Map '{"是":"Y"}'] [-Year 2025] [-NoSpaces] [-NumberFormat]` | **清洗列**（v1.29）：去多余空格（含全角/制表符）、去指定字样、大小写、取值改名；`-To number` 认 `¥1,299`/`89元`/`10%`/`12 件`，`-To date` 认 `2025/3/1`/`2025.03.01`/`20250301`/`2025年3月1日`/`3月1日`（**年份取整列众数**，可用 `-Year` 指定）。返回每列 `changed`/`notReadable`（读不成的格 + 前 5 个地址与原文）/`valuesNow`；**有公式的列只写非公式格** |
| `fillblanks` | `-Column 水位`（或 `-Columns`）`-With 0\|above\|mean\|median\|mode` | **补空格**（v1.29）：填固定值 / 上一格 / 均值 / 中位数 / 众数（报告附表"同上"高频） |
| `dedupe` | `[-Columns "站码,日期"] [-Range]` | **去重复行**（v1.29）：键列相同（省略 = 整行相同）只留第一行；**真删行**，删前 `snapshot` |
| `deletewhere` | `-Column 流量 -Cond "<" -Value 0`；`-Cond blank`（省略 `-Column` = 删整行为空的行）；`-Cond in -InValues "a,b"` | **按条件删行**（v1.29）：`blank\|=\|<>\|>\|<\|>=\|<=\|contains\|text\|error\|in`；**自底向上合并区间一次 `Delete`**（不是逐行删），返回删了哪些行 |
| `sort` | `-By "流量:desc,站码"`（或 `-Column 流量 -Order desc`）`[-NoHeader]` | **就地表排序**（v1.29）：公式与格式随行走，不破坏联动 |
| `filter` | `-Column 站点 -InValues "外洲,李家渡"`；`-Column 流量 -Criteria ">1000"`；`-Action clear` | **自动筛选**（v1.29）：只隐藏不删除，返回 `shown/total`；`-Action clear` 复原 |
| `copy` | `-Range A1:F21 -To "备份!A1" [-Values]` | **复制区域**（v1.29）：公式+格式一起；`-Values` 固化数值且**不重解析文本**（20250301 不会变数字）。只做同簿 |
| `addcol` | `-Header 流量2 -Formula "=D2*2"`（或 `-Value 0`）`[-After 流量] [-NumberFormat "0.0"] [-Values]` | **加一列并填满**（v1.29）：同名表头 = 重填；`-Values` 固化结果；返回错误格 |
| `pivot` | `-Range A1:F2000 -Rows 站点 [-Columns 月份] -Agg "流量:sum,水位:average" [-Filters] [-Group "日期:month\|quarter\|year"] [-SortBy "合计：流量:desc"] [-Name 透视表] [-To "某表!A3"]` | **透视表**（v1.29）：结果网格**随返回值一起回**（`rows:[[行号,值…]]`，前 40 行 × 14 列），不用再 `read`；默认新建工作表「透视表」落在 A3。源区域每列必须有表头。**v1.31**：`-Group` 已实测——WPS 12.1 **原生支持**按日期分月/季/年（24 行实测：月 110…1320 总计 8580、季 660/1650/2640/3630、年 780/7800，全与手算一致），返回 `groupMode`（`none\|native\|failed`）+ `group{field,by,mode}`；**文本日期会失败**（WPS 抛「选定区域不能分组。」）→ `failed` + `notes` 给降级配方（加 `=YEAR()`/`TEXT(日期,"yyyy-mm")`/`"Q"&ROUNDUP(MONTH()/3,0)` 辅助列再按辅助列透视；辅助列是数据源、建完不能删） |
| `sheet` | `-Doc -Action list\|add\|rename\|delete\|activate [-Name] [-NewName]` | 工作表增删改名 + **`activate`（v1.17）把某张表变成活动表**（用户看得见、且依赖 ActiveSheet 的老操作也跟它走）；以前只能裸 COM `Worksheets.Item(n).Activate()` 逃生 |
| `comment` | `-Action list`；加：`-Cell D5 (-Text "短文" \| -TextFile <txt>)`；删：`-Action delete -All` | 单元格批注（校审指认用；中文长文走文件避开引号转义） |
| `search` | `-Doc -Text <词> [-Sheet 名] [-Max 50] [-Case]` | **找格**：跨表搜值和公式，返回命中格/值/公式（定位用这个，别整表 read） |
| `selection` | `-Doc` | **读用户当前选中的区域**：sheet/range/行列数/值（≤400 格直接带值）——用户说"就改我选的这块"时用；选在别的窗口报 NOT_OPEN |
| `rows` / `cols` | `-Doc [-Sheet] -Action insert\|delete -Index N [-Count M]` | 插入/删除 N 起的 M 行（列）——结构编辑；**先 snapshot**（可回滚），公式由表格自动调整 |
| `render` | `-Doc [-Sheet] [-Range] -Out <目录> [-Width 1100]` **`[-IncludeCharts]`** | 区域 `CopyPicture(xlScreen)` → PNG。**语义（v1.17 实测修正）**：图里**会**包含与该区域相交的浮动图表（A1:H52 有两张图 → 都拍到；A1:H10 无重叠 → 一张也没有），但**不含区域之外的**图表。`-IncludeCharts` 把视图自动扩到覆盖该表**全部**图表再出图（返回 `chartsIncluded` 与每张图 `left/top`）——有图表的表用它，**一次成图，不要再截桌面**。**v1.32：render 前后会保存/恢复你的剪贴板**（`CopyPicture` 用完不吞你复制的内容，且恢复的内容不进 Win+V 历史） |
| `chart` | 加：`-Range 数据区 [-Type column\|bar\|line\|pie\|scatter\|area\|combo] [-Title] [-XTitle] [-YTitle] [-Y2Title] [-Secondary "系列名/序号"] [-Labels value\|percent] [-Legend bottom\|top\|left\|right] [-AxisMin/-AxisMax] [-NumberFormat] [-Anchor E2] [-Width 480] [-Height 300]`；`-Action list [-All]`；`-Action delete\|move -Index n [-Anchor X] [-Width] [-Height]` | 建原生图表。**v1.17**：① 第 1 列=x，**第 2..N 列各一条系列**（表头作系列名，多系列自动出图例）——「Q–H 族曲线」一张图即可；② 尺寸可调（默认 480×300 pt）；③ 建完可 `move` 挪位/缩放；④ `list -All` 跨表列出全部图表与几何信息；⑤ `-Sheet` 已正确转发。**v1.29**：⑥ `-Type combo` + `-Secondary` = **组合图/次坐标轴**（过程线+雨量柱，水文刚需）；⑦ `-Labels/-Legend/-AxisMin/-AxisMax/-NumberFormat`；⑧ 返回 **`occupancy:"G22:N41"`（占地区域）与 `overlaps:[…]`（与已有图重叠硬提示）**——放下张图前先看它；⑨ WPS 组合图刚改完系列类型时设标题会被拒，helper 已按 dsh-office 口径**重试 3 次再退 `SetElement`** |
| `autofit` | `-Doc [-Sheet] [-Range]` | 列宽行高自适应 |
| `freeze` | `-Cell A2`（冻结其上方与左侧）；`-Action remove` 取消 | **冻结窗格**（v1.29）：长表滚动看表头用 |
| `cformat` | `-Range D2:D500 -Rule ">" -Value 1000 [-FillColor "#FFC7CE"]`；`-Rule colorscale\|databar\|top\|bottom\|duplicate\|unique\|blank\|formula\|between\|clear` | **条件格式**（v1.29）：阈值着色/色阶/数据条/前 N/重复值/空值/公式；`-Rule clear` 清空该区域规则 |
| `validate` | `-Range H2:H200 -List "是,否,待定"` | **数据验证下拉**（v1.29，≤255 字符、值里不能含逗号；区域源未做） |
| `table` | `-Range A1:F21 [-Name 站点表] [-Style TableStyleMedium2]` | **建成 ListObject 表**（v1.29）：纹行 + 筛选箭头 + 列名可当名字用；已是表则只改样式/名字。⚠️ 表内不能 `filter`（见下方限制表） |
| `format` | `-Doc -Range A1:D1 [-Sheet] + 格式项`；**多区域：`-Range "A4:E4;A13:E13;A27:E27"`（分号，v1.17）**；**每块不同格式：`-Items <json>`（v1.18，见下）** | 字体/字号/加粗/颜色/填充/对齐/数字格式/列宽/行高/框线/换行。一次调用排完整版，别再逐块串行（实测 26 次 ≈ 20 s）。**关掉某项用 `-NoBold/-NoItalic/-NoUnderline/-NoWrap`**（v1.17：`-Bold:$false` 过不了 `powershell -File` 的参数解析）。**v1.18 起 `-Color`/`-FillColor` 的 RGB 通道已修正**（之前红蓝互换） |
| `snapshot` | `-Doc [-Sheet 名1,名2] [-Out <路径>]` | 存工作表的值+公式快照（默认全部表；大表用 `-Sheet` 限定范围，diff/restore 也只在该范围内比对/回滚，不碰未覆盖的表；默认存 %LOCALAPPDATA%，`-Out` 随文档带走） |
| `diff` | `-Doc [-Since 快照路径\|last]` | 和快照比：哪些格改了（before→after，含公式联动）、哪些增删——**人改的和 Agent 改的都算** |
| `restore` | `-Doc [-Since 快照路径\|last] [-Cells "Sheet1!C2,Sheet1!E2"] [-KeepExtras]` | **一键撤回**：把工作簿还原到快照状态（只改有差异的格；快照后新增的非空格默认清空，`-KeepExtras` 保留）。`-Cells` = **选择性回滚**，只还原点名的格（快照中没有的点名格清空），其余一切——包括人改的——原样保留。只还原内存，`diff` 验证后再决定 `save` 或不保存直接 `close` |
| `calc` | `-Doc` | 全簿重算 |
| `save` | `-Doc [-As <路径>]` | 保存 / 另存（.xlsx/.xls/.csv/.xlsm；`.pdf` = 整簿导出 PDF） |
| `close` | `-Doc [-Save]` | 关闭；无其它工作簿时退出实例，回报进程数 |
| `quit` | — | 关掉该实例**全部**工作簿并退出（不保存）。**慎用** |
| `finish` | `[-Note "改了什么"] [-Cell A1] [-Toast] [-Sound] [-Quiet] [-Force]` | **任务完成信号**（agent 无关，零依赖）：默认**托盘气泡 + 提示音**（原生 .NET，任何机器都响），**不动工作簿任何内容**。`-Cell` 显式指定时才额外写一行状态（目标格非空拒绝，`-Force` 覆盖）；`-Quiet` 静音。**任务收尾默认调用** |
| `hint` | `-Doc <路径> [-Text "正在做什么"] [-HintTag "〔AI接管中〕"] [-Bar "..."] [-Toast] [-Sound] [-Quiet] [-Clear] [-Tray]` | **接管提示**（v1.11 / v1.12 加 `-Tray`）：改窗口标题 + 写状态栏一行 + 默认弹托盘气泡，**不写单元格、不落盘**。`open` 自动打标签，`finish`/`close` 自动还原；表格的 `Application.StatusBar` 可读写且实测会显示这行字，`-Clear` 时写 `false` 把状态栏交还 WPS。`-Tray` 加常驻托盘图标（悬停看进度、右键还原） |

所有命令都可用 `-TimeoutSec <秒>`（默认 180）；超时返回 `TIMEOUT`，说明 WPS 大概率弹了对话框。

### Excel 侧已知限制（v1.17 实测，先读再动手）

| 限制 | 实测表现 | 正确做法 |
|---|---|---|
| `render` 的图表语义 | `Range.CopyPicture(xlScreen)` 画的是「这块区域在屏幕上的样子」：**与该区域相交的浮动图表会被拍进去**（2026-11-02 实测 A1:H52 两张图都在、A1:H10 一张都没有），**区域之外的图表永远不在图里** | 图表的表一律 `render -IncludeCharts`（把视图扩到覆盖全部图表）；只渲染表格时**别把图表圈进 `-Range`**，否则会重复画 |
| `chart` 数据**必须同表** | `dynamic data = ws.Range(dataRef)` 不走跨表解析，`"其他表!A1:B9"` 不可靠 | 辅助数据放在**同一工作表**（例如右侧 AG:AQ，字号 8 弱化并标注"勿删"） |
| `chart` 的系列约定 | 第 1 列 = x 轴，第 2..N 列各一条系列，**第 1 行是系列名**（v1.17 起多系列；v1.16 及更早只取最后一列） | 单列数据 = 一条系列（无 x 轴）；要「一族曲线」就按 x + 多值列排好 |
| `format -Borders` 只认 `grid` / `thick` / `none` | 其他值会静默当 `grid`；**v1.16 起对角线（XlBordersIndex 5/6）恒被清空**，不会再把表格画成大 X（v1.15 及更早有此 BUG） | 表格框线用 `-Borders grid`（细）/`thick`（粗）/`none`（清） |
| `render` 的宽度上限 | `-Width` 低于 100 不缩放；超大区域按 `-Width` 等比缩放 | 想看全貌用 `-Width 1200~1600`，别一次要 4000 |
| **一个 ET 实例同时开两个工作簿会挂起**（v1.20 实测，非本次改动引入） | `open` 第二个工作簿时 WPS 12 的 `Workbooks.Open` 不返回（等 30 s+ 无响应；之后实例里剩一个空名幽灵工作簿，后续命令报 `DOC_CLOSED`/`E_FAIL`，只能结束 `et.exe` 重来）；用 `-Show`、走 daemon/一次性、先关再开都复现，部署版 v1.20b 同样复现 | **一个实例只开一个工作簿**：多工作簿任务必须串行（`close` 前一份再 `open` 下一份），不要把两个工作簿并存当并行 |
| **`Range.NumberFormat` 整块多列返回 Null**（v1.29 实测） | `ws.Range("A1:C9").NumberFormat` 在 WPS 上得 Null，取 `.GetType()` 直接崩；只有**单列**能拿到 String | 判日期/格式一律**按列取**（整列格式不一致时再回退抽样单元格）；`read -Profile` 已按这个口径实现 |
| **块写必须 0 基、尺寸精确、单列要含表头格**（v1.29 实测） | 数组尺寸差一格 WPS 静默不落位；单列块不含表头格会把表头写空 | 新命令（`clean`/`addcol`/`fillblanks`…）内部已遵守；自己写裸 COM 时同样规矩 |
| **表（ListObject）内不能 `Range.AutoFilter`**（v1.29 实测） | 在 `table` 建出的表区域上调 `AutoFilter` 报 `E_ACCESSDENIED`；同区域 `sort`/`delete`/`fill` 正常 | 要在表内筛选：先 `table` 不建表、用普通区域 `filter`；已建表则删表后再筛（命令里已把这句翻成明确报错） |
| **`pivot` 的 `Create` 只能两参**（v1.29 实测） | `PivotCaches().Create(1, range)` 可用；照 VBA 写四参/三参会报「找不到 Create 的重载」 | 命令内部已固定两参；`CreatePivotTable(dest,name)` 同样两参 |
| **`pivot -Group` / 次轴刻度未实测**（v1.29） | 按月/季/年分组代码按 dsh-office 口径写了但未实测（失败只记 `notes` 不抛错）；图表次坐标轴的刻度/数字格式只作用主轴 | 需要分组就先小样本试一次、看 `notes`；次轴刻度目前只能接受默认或手工调 |
| **组合图刚改完系列类型时设标题会被拒**（v1.29 实测） | WPS 在图表还在排版时拒绝 `HasTitle/ChartTitle`，直接 `SetElement` 也有概率失败 | helper 已**重试 3 次再退 `SetElement`**（照 dsh-office 的做法）；调用方不需要做任何事，但看到 `title` 为空就重跑一次 |

### 调用姿势（v1.17，踩过的坑）

1. **优先 `woc.ps1`，不是 `woc.cmd`**：`woc.cmd` 走 cmd.exe，`=A15^1.5` 会被改成 `A151.5`、前导 `=` 会被吃掉，`fill/write -Formula` 基本全废（另一台机器实测，浪费 3 min）。`woc.ps1` 是 PowerShell→PowerShell，`= ^ & | > $` 原样到达：
   ```powershell
   & <skill>\scripts\woc.ps1 excel fill -Doc X.xlsx -Sheet Sheet1 -Range "B15:B25" -Formula "=A15^1.5"
   ```
2. **公式里的 `$` 在 PowerShell 双引号里会被插值**（`"=$B$5*2"` → `=*5*2`，WPS 报 `0x8FE3000B` 这种诡异 HRESULT）。**含 `$` 的公式用单引号** `-Formula '=$B$5*2*SQRT(2*$B$6)*B15'`；含 `=`/`^` 又不方便转义时，走 `-Items <json>` 文件。
3. **中文一律走 JSON 文件**（`write -Items`、`comment -TextFile`），不要塞进命令行——`woc.cmd` 那条路还受 OEM 代码页影响。
4. **参数被静默忽略会告警（三端通用，v1.18）**：`scripts\param-check.ps1` 被 Word/Excel/PPT 三个包装器共用，对「绑定了却没进 request」的参数在 stderr 打 `WARN ignored parameter -X for command "Y"`。看到就说明这条命令不吃该参数（v1.16 前是静默丢弃，`chart -Sheet` 就是这么丢的）。显式传默认值（`-From 1`、`-Index 0`）不会误报；`WOC_NO_PARAMCHECK=1` 可整轮关掉。
5. **颜色通道（Excel 侧 v1.18 修）**：`-Color`/`-FillColor` 之前红蓝互换（COM 要 BGR，旧代码给 RGB），现在 `#FF0000` 就是红色。若历史脚本里颜色看起来"反了"，按新行为重跑即可。
6. **多条命令用 `batch`，别一条一个工具调用（v1.19）**：`woc.ps1 <app> batch -Ops ops.json`，一次进程启动跑完整批。实测本机 5 条命令 1982 ms → 891 ms（每条 op 35–204 ms）；命令越多差距越大。
   - `ops.json` 里是**helper 级请求**（不是命令行参数名），每条形如 `{"cmd":"...", <请求字段>}`，`doc` 可写相对路径（会转绝对）。
   - 常用映射：写单元格 `{"cmd":"write","doc":X,"items":[{"sheet":"S","cell":"A1","value":1}]}`；格式 `{"cmd":"format","doc":X,"range":"A1:A2","bold":true}`；建图 `{"cmd":"chart","doc":X,"action":"add","range":"A1:B2","type":"line","anchor":"D2"}`；读 `{"cmd":"read","doc":X,"sheet":"S","range":"A1:A2"}`。
   - 返回 `{"result":{"count":n,"totalMs":...,"ops":[{i,ms,ok,result|error}]}}`，逐条带耗时；某条失败不影响后续（但**不会回滚**已成功的那些，重要批次仍要 `snapshot`）。
   - 适合"搭一页 PPT / 排一张表"这种 10–30 条小命令的场景；单条命令用普通 CLI 就行。

### skill 没覆盖时：合法逃生通道

「一份文档一条通道」针对的是**改内容**，不是**看一眼**。以下操作允许直接 `Marshal.GetActiveObject('KET.Application')`，因为不碰数据：

- 只读探测（读 `ActiveWindow.ScrollRow`、窗口标题、可见性等 skill 不回报的状态）；
- 纯窗口操作（滚动到某行、截图、置前）——但优先用 `render -IncludeCharts`；
- **任何写单元格 / 改结构 / 改格式仍然必须走 skill**（否则撤销栈、快照、diff 全断，且与用户的操作打架）。

v1.17 已经把最常触发的两处逃生需求做成命令（`sheet -Action activate`、`chart -Action list -All`），如果还遇到必须裸 COM 的场景，请当作 skill 缺口反馈。

### 视觉核对纪律（省步数，也是 v1.16 的由来）

2026-11-01 一次「新建 Excel + 讲宽顶堰公式」的会话跑了 80 步 / 53.7 min，其中约 24 步（30%）耗在**截图 → 裁剪 → 放大 → 再看**上，还有一批 `SetForegroundWindow`/`mouse_event` 抢窗口的 P/Invoke。归因是当时的 `render` 不含区域外的图表，只能整屏截图。现在按下面做：

1. **有图表的表：一条 `render -IncludeCharts` 就够**，它同时给出单元格与全部图表；**不要**再截桌面、不要滚轮翻页、不要用 Win32 API 抢前台。
2. **目检最多 2~3 次**：写完内容看一次、改完版式看一次、交付前看一次。**不要**每改一处就重截一次。
3. 需要看细节时用 **`-Range` 框小区域 + `-Width` 放大**，不要截 1920×1080 整桌面——整屏截图会让后续每一步都变慢（图像累积进上下文，实测每步从 27 s 涨到 40~80 s）。
4. 没有图表的表：`render -Range` 直接出区域图，比截图清晰且不含桌面噪音。

### 数据活儿怎么干（v1.29：概况 → 清洗 → 分析 → 成品感）

```powershell
# 1) 先看概况：look 告诉你哪列脏、dateColumns 哪列是日期（值已转 yyyy-MM-dd）
& scripts\woc.ps1 excel read -Doc 数据.xlsx -Sheet 数据 -Profile

# 2) 清洗（先 save -As 备份或 snapshot；每步都会报 changed/notReadable）
& scripts\woc.ps1 excel clean -Doc 数据.xlsx -Sheet 数据 -Column 金额 -To number
& scripts\woc.ps1 excel clean -Doc 数据.xlsx -Sheet 数据 -Column 日期 -To date -Year 2025
& scripts\woc.ps1 excel clean -Doc 数据.xlsx -Sheet 数据 -Columns "站点" -NoSpaces
& scripts\woc.ps1 excel fillblanks -Doc 数据.xlsx -Sheet 数据 -Column 水位 -With above
& scripts\woc.ps1 excel dedupe -Doc 数据.xlsx -Sheet 数据 -Columns "站码,日期"
# 文字更正/统一叫法：内存块替换（先 -DryRun 看清单），公式格与数字格不会被改动
& scripts\woc.ps1 excel replace -Doc 数据.xlsx -Sheet 数据 -Column 单位 -Find "水利局" -Replace "水利水电勘测设计院" -Whole -DryRun
& scripts\woc.ps1 excel replace -Doc 数据.xlsx -Sheet 数据 -Column 单位 -Find "水利局" -Replace "水利水电勘测设计院" -Whole

# 3) 分析：透视表结果直接回，不用再 read
& scripts\woc.ps1 excel pivot -Doc 数据.xlsx -Sheet 数据 -Range A1:F2000 `
    -Rows "站点" -Columns "月份" -Agg "流量:sum,水位:average" -Name 站点月汇总

# 4) 成品感：冻结表头 / 条件格式 / 下拉 / 建表
& scripts\woc.ps1 excel freeze -Doc 数据.xlsx -Sheet 数据 -Cell A2
& scripts\woc.ps1 excel cformat -Doc 数据.xlsx -Sheet 数据 -Range D2:D500 -Rule ">" -Value 1000 -FillColor "#FFC7CE"
& scripts\woc.ps1 excel validate -Doc 数据.xlsx -Sheet 数据 -Range H2:H200 -List "是,否,待定"
& scripts\woc.ps1 excel table -Doc 数据.xlsx -Sheet 数据 -Range A1:F21 -Name 站点表
```

**纪律**：`clean`/`deletewhere`/`dedupe` 是**真改真删**（`deletewhere`/`dedupe` 行号会变），先 `snapshot`，需要留底先 `save -As`；每步之后看返回的 `changed/notReadable/deleted/errors`，不要连跑五条再回来看；`pivot` 新建表放分析结果（默认名「透视表」，A3 起），**图表数据要同表**（跨表不可靠）。
**参数命名约定**（避免与既有开关撞车）：大小写用 `-Casing`（`-Case` 是搜索开关）、条件用 `-Cond`、in 列表用 `-InValues`、图表轴上下限用 `-AxisMin/-AxisMax`（`-Max` 已存在）。

### `write` 的 items 格式

```json
{"items":[
  {"sheet":"数据","cell":"E1","value":"avg"},
  {"sheet":"数据","cell":"E2","formula":"=AVERAGE(C2:C3)"},
  {"cell":"'汇总'!A1","value":"合计"},
  {"sheet":"数据","cell":"F1:G2","values":[[1,2],[3,4]]},
  {"sheet":"数据","cell":"C2","value":1875.5,"expect":"1850.5"}
]}
```

- 三种写法：`value`（标量）、`formula`（以 `=` 开头，交给表格算）、`values`（二维块，尺寸必须与 `range` 一致）。
- 定位：`cell` 单格或 `range` 区块；工作表用 `sheet` 字段，或直接写 `"Sheet名!A1"` 引用（名字带空格/特殊字符用单引号包住）。
- `expect` 可选但强烈建议：该格**当前显示文本**必须与它完全相等才动手，不符即 `ANCHOR_MISSING`、**整批一条不改**——防编号漂移、防"人刚改了这格"。
- 整批原子：全部校验通过才开写；万一中途 COM 报错，已写的格用捕获的旧 Formula **逐格还原**，错误信息里写"rolled back N cell(s)"。

### `format` 的 items 格式（v1.18，每块不同格式）

分号多区域（`-Range "A1:E1;A5:E5"`）只能共用**一套**格式项；要「表头加粗蓝底、数值 0.00、第一列红字」这种**每块不同**的排版，用 `-Items`：

```json
{"items":[
  {"sheet":"Sheet1","range":"A1:C1","bold":true,"size":"16","fillColor":"#D9E1F2","halign":"center","borders":"grid"},
  {"sheet":"Sheet1","range":"A2:C5","numberFormat":"0.00","borders":"grid"},
  {"sheet":"Sheet1","range":"A2:A5","color":"#C00000","bold":true},
  {"sheet":"Sheet1","range":"A7:C7","font":"黑体","size":"12","italic":true}
]}
```

- 每块字段与命令行格式项同名（`font/size/bold/italic/underline/color/fillColor/halign/valign/wrap/numberFormat/colWidth/rowHeight/borders`），可写 `sheet`，`range` 里还能再用分号并多个区域。
- 关掉某项用 `false`（JSON 里是布尔，不受 `powershell -File` 参数解析限制）；命令行侧对应 `-NoBold/-NoItalic/-NoUnderline/-NoWrap`。
- 返回 `{"items":[{sheet, ranges:[{range, applied}]}], count:n}`；某块没有格式项直接 `BAD_ARGS`。
- 每条返回 `before`/`after`（写入前后该格显示文本）——**立刻比对**，能第一时间发现写错位置。
- **撤销**：表格的 COM 修改不进撤销栈且会清掉用户自己的撤销记录——写前 `snapshot`、写后看 before/after，别指望 Ctrl+Z。

### 大表效率纪律（几千行以上必读）

1. **连续数据拼成 `values` 块**一次写（`{"cell":"A1:F5000","values":[...]}`），绝不逐格 item；
2. **整列公式用 `fill`**，不用 write 逐行；
3. **快照限定范围** `snapshot -Sheet 数据表`，避免几十 MB 的全簿快照；
4. 读大表交给自动概况（>60 行自动列概况+头尾样本），要局部全量用 `-Range`；
5. write/restore 已自动挂起重算+屏幕刷新、批末一次重算，无需手动干预。

## Excel 格式项（`format`）

| 参数 | 取值 | 说明 |
|---|---|---|
| `-Font` / `-Size` | `仿宋` / `12` | 字体 / 字号（磅） |
| `-Bold` `-Italic` `-Underline` | 开关 | `-Bold:$false` 关闭 |
| `-Color` `-FillColor` | `"#C00000"` / `"#DDEBF7"` | 字色 / 底色 |
| `-HAlign` | `left\|center\|right\|justify` | 水平对齐 |
| `-VAlign` | `top\|center\|bottom` | 垂直对齐 |
| `-Wrap` | 开关 | 自动换行 |
| `-NumberFormat` | `"0.00"` / `"0.0%"` / `"yyyy-mm-dd"` | 数字格式 |
| `-ColWidth` / `-RowHeight` | `12` / `20` | 列宽（字符）/ 行高（磅） |
| `-Borders` | `grid\|thick\|none` | 全边框 / 粗边框 / 去边框 |

## Excel 典型工作流

1. **接管**：`status` → `open -Show` → `read`（概览）→ 需要数据时 `read -Sheet X` 或 `read -Range "X!A1:F200"`。
2. **定位**：`search -Text` 拿命中格与上下文（全文找格，取代整表 read 翻找）；用户说"就改我选中的这块"时用 `selection` 直接拿他选中的区域。
3. **改表**：`write`/`fill` **每次自动先落 pre-batch 快照**（返回值 `preSnapshot.path`，无需手动 snapshot）→ 比对每条 before/after → `diff -Since <path>` = 本批改动归因 → `render` 抽查 → 用户看过再 `save`。改坏了 `restore -Since <path>` 整批回滚，或 `restore -Since <path> -Cells "表!C2"` 只撤点名格（人改的其它格不受影响）。
4. **改结构**：`rows`/`cols` 插删行列（**先 snapshot**，结构命令暂无自动快照）→ `fill` 补公式 → `read` 复核。
5. **校审指认**：`comment -Cell D5 -Text "..."`（或批量逐条加）→ `render` 看角标 → `comment -Action list` 复核。
6. **排版**：`format`（表头加粗+填充+居中+边框）→ `autofit` → `render` 复看。
7. **新表**：`sheet -Action add -Name 汇总` → `write`（引用写 `'汇总'!A1`）→ `calc` → `render`。
8. **图表**：`chart -Action add -Type line -Range 数据 -Title ... -Anchor E2` → `render` 目检 → 不满意 `-Action delete` 重来。

## Excel 故障排查

| 错误码 / 现象 | 怎么办 |
|---|---|
| `NO_APP` | 没装 WPS 或 KET COM 没注册。手动开一次 WPS 表格再试；`doctor` 里 `apps.wpsEt.installed=false` 说明这台机器没有。 |
| `NOT_OPEN` | `-Doc` 与实际打开的工作簿不是同一路径（另存过身份会变：`save -As` 之后按新路径操作）。先 `status` 把 `path` 原样复制过来。 |
| `DOC_CLOSED` | **工作簿在协作期间被关闭了**（很可能是用户手动关的，未保存的改动随关闭丢失）。**不要静默继续**：先向用户确认，确需继续就显式 `open -Doc` 重新接管。skill 记住了接管过的工作簿（attachment marker），从关闭到显式 open 之间的任何命令都会得到这个错误，而不是悄悄把文件复活 |
| `READONLY` | 文件被占用或只读打开。让用户关掉占用程序或另存，不要强写。 |
| `ANCHOR_MISSING` | `expect` 与该格当前显示不符（人刚改了 / 定位错表）。整批没改，重新 `read`/`write` 核对。 |
| `BUSY` / `TIMEOUT` | 用户正在单元格里输入、或有对话框挡着。**v1.32 起 helper 装了 OLE 消息过滤器**：WPS 拒绝调用时自动以 150 ms 间隔重试、最多 **6 s**，仍不行才报 `BUSY`——看到它就说明"等过 6 秒还忙"，等用户停手再发。Excel 忙码另补了 `0x800AC472`。 |
| 文件打不开/读不出 | 多半是密文且无白名单进程可读。让用户在 WPS 里打开一次（驱动透明解密），或按 [reference/encryption.md](reference/encryption.md) 处理。 |
| `write` 报 rolled back | 批中途 COM 失败已自动还原；看错误信息，通常用户正好在编辑。重试或分小批。 |
| 中文乱码 | 控制台代码页问题；本 skill 输出已是 UTF-8，若外层管道乱码，先 `[Console]::OutputEncoding=[Text.Encoding]::UTF8`。 |
| et.exe 变多 | 中途异常没 `close`。`status` 看实例，只清理无窗口实例，汇报数字。 |
| `replace` 返回 `changed:0` 但有 `numericMatchesLeftAlone`（v1.31） | 命中的是数字/日期格（如 `-Find 0` 想改 `10/20/30`）。**这是有意的**：把数字列悄悄变成文本列比"没改到"更坏 | 要改数值就用 `clean -To number`、`write`，或对整列做 `format`；确实要改成文本，先 `clean -To text` 再 `replace` |
| `pivot -Group` 返回 `groupMode:"failed"` + 「选定区域不能分组。」（v1.31） | 日期列是**文本**（`@` 格式的 `2024-01-15`），WPS 拒绝按日期分组——这是 WPS 行为，不是本 skill 的问题 | 按返回 `notes` 的降级配方：`addcol -Header 年份 -Formula '=YEAR(A2)'` / `=TEXT(A2,"yyyy-mm")` / `="Q"&ROUNDUP(MONTH(A2)/3,0)`，再用辅助列做 `pivot -Rows 年份`。辅助列是数据源，**建完不能删** |
| `replace` 的规则第二次还命中（v1.31） | 替换串里包含查找串（`P`→`P-`、`0`→`00`）时，再跑一次必然再命中——这是规则性质 | 一次调用只插一遍（`PP→P-P-`，不会 `P--`）；要幂等就把规则写成不自匹配的形式（如 `表1`→`表一`），或先 `snapshot`/`diff` 对账 |

---

# PPT：在 WPS 演示里实时改演示文稿

命令前缀：`office.ps1 ppt <命令>`（或直接 `ppt.ps1 <命令>`）。ProgID `KWPP.Application`；`open` 一个不存在的路径会**新建并立即保存**一份空稿到该路径（新稿身份靠这次落盘确立）。

**关键限制（先读）**：PPT 无修订层、外部 COM 改动不进撤销栈（同 Excel）；**WPS 演示不接受公式**（不能建也不能贴）；`render` 走剪贴板（`Slide.Copy` → EMF → PNG，helper 自己写盘故是明文；`Slide.Export` 写的文件会被透明加密、读不回来，所以不用它）。

## PPT 实测坑（2026-11-02，另一台机器 70 条命令的实测报告）

| 坑 | 实测表现 | 做法 |
|---|---|---|
| **页码占位符渲染成重复 "1"** | `master -Action footer -PageNum` 在本机 WPS 12 上每页都显示 "1"（占位符文本不随页更新） | 别用页码占位符；页脚写纯文字，或干脆不放页码 |
| **`open` 不存在的路径 = 新建** | 这是特性（新稿靠这次落盘确立身份），但路径写错会留下垃圾文件；旧返回与普通打开无法区分 | v1.19 起返回 `how:"created"` + `note`，看到 `created` 而你以为在打开已有文件 ⇒ 立刻检查路径 |
| **未转义的 Windows 路径** | 手拼 JSON 时 `C:\Users\...` 里的 `\U`/`\A`/`\D` 被当非法转义，旧解析器**静默丢反斜杠** ⇒ `C:Users...`，再叠加"路径不存在就新建"就凭空多一份稿 | v1.19 起非法转义直接 `BAD_ARGS`；拼请求一律走 JSON 序列化器（`batch` 的 ops 文件也是） |
| **渲染成本 ~355 ms/页** | 3 页 1471 ms；`check` 全规则 970 ms/次；`read` 3 页 632 ms | 设计阶段渲一次 `-Overview`（6 页/图），之后**只渲改动页**；`check` 过程别跑，收尾跑一遍 `-MinSeverity blocker`；`read` 一次读全稿并缓存形状名 |
| **接管后窗口不可见**（v1.25 修） | 用 `open`（不带 `-Show`）接管用户已打开的演示时，窗口会被置为隐藏（`Windows.Count=0`/`Visible=false`），用户在 WPS 里看不到任何改动 | v1.25 起 `open` 默认显示并 `Activate()`；确实要后台跑才传 `-Show:$false` |
| **标题标记会被切标签页清掉**（v1.26 实测） | WPS 12 用共享窗口框（`PP12FrameClass`，属 `wps.exe`），`SetWindowText` 写的「〔AI接管中〕」在 COM 编辑/翻页/空闲后保持，但**用户切到别的标签页时 WPS 会重置标题** | 不是 bug：这样不会把标记挂到别的文档；下一条针对该稿的命令会自动补回（helper 里有 `TitleState` 记录，`hint -Clear`/`finish`/`close` 会停用） |
| **母版上的标记会被整幅大图盖住**（v1.27 实测） | 第一版把「〔AI接管中〕」药丸加在 `SlideMaster`，实测封面这类**满版图片页**把母版形状整片遮住（母版在页面形状之下），渲染图上完全看不见 | v1.27 起改为**逐页贴、置顶**（`woc-ai-takeover-mark`）；如果自己做类似标记，别指望母版 |
| **页面标记是内容，不是窗口状态**（v1.28 定案） | 与标题标记不同，药丸是**形状**：用户自己按 Ctrl+S 会把它存进文件（实测：裸 COM `Save()` 后重开，徽章确实在文件里）；`save`/`export` 已在落盘前自动剥离、写完补回 | 默认**不贴**（`hint` 零改动）；要界面可见才 `-SlideMark`，收尾 `finish`/`hint -Clear`；万一被 Ctrl+S 存进去，**下一条命令自动修复**（`badgeRepaired:true`），只读文件给 `badgeDirty` |
| **daemon 必须保温** | 一次性模式 746 ms vs 保温 422 ms；`close` 掉最后一份文档后 daemon 自己退，下一条命令重新付 ~0.5 s | 一轮任务里别频繁 `close` 又重开；批量命令用 `batch`（一次启动跑一批） |
| **图片裁切单位是点，不是比例**（v1.29） | `PictureFormat.CropLeft=0.14` 会被存下、也读得回来，但**渲染完全不裁**（图被拉伸填满框）；且**裁完再改 `Width/Height` 会把裁切洗掉** | `crop` 内部已换算成「比例 × 缩放后尺寸（点）」并保证顺序（先定框→写裁切→只调 Left/Top）；自己写裸 COM 时照这个顺序。裁完 `boxActual` 可能与请求差几十点（480×270→480×224），**永远 `render` 目检** |
| **SmartArt 只能建不能改**（v1.29） | `Shapes.AddSmartArt` 成功（"图示 1"），但 `Shape.SmartArt.AllNodes` 读或写全部 `E_FAIL`/`RPC_E_SERVERFAULT`；`AddNode` 也不通 | 别承诺"在 SmartArt 里逐条写文字"。要可控框图：`shape -Name` 摆框 + `connect` 连线；`smartart` 只用来快速占位（返回 `nodes:0` + note） |
| **`-Arrow both` / `Fill.OneColorGradient` 回读不可信**（v1.29） | 双向箭头的起始箭头回读 `begin=1` 但实际不生效；`OneColorGradient` 写进去 `Fill.Type` 回读仍是 1（solid） | 需要渐变/纹理一律走 `bg -Texture`（生成位图再 `UserPicture`，fillType 回读 6 才是真生效）；箭头只信 `-Arrow end` |
| **`anim` 只证明写入，不证明放映**（v1.29） | `EntryEffect` 回读一致、`MainSequence` 条数增加，但**未在放映里验证** | 交付前 F5 人工看一眼；返回 `unsupported:true` 或 `matches:false` 时如实告诉用户"没设上"，别写"已加动画" |
| **成品页的 `check` 误报 pagenum / fontsize**（v1.31） | 新版式页码是**普通文本框**（WPS 页码占位符每页显示 "1"，不可用），`check` 的 `pagenum` 规则按占位符判 → 每页报 1 条；`gallery` 图注 10pt < 正文 14pt 门槛也报 `fontsize` | 交付前用 `check -Config '{"suppress":[{"rule":"pagenum"}],"bodyMin":12}'`（或把图注调到 12pt）；这两类不是缺陷，别为此改版式 |
| **`TextFrame.AutoSize` 别碰**（v1.31 实测） | 给文本框写一次 `AutoSize` 会把框**压成约 29pt 高**（WPS 按文字重算，宽度信息丢失），随后 `check` 满屏 overflow | 写文字**不要**设 AutoSize；写完回写 `Left/Top/Width/Height` 保持几何（`Kinds.cs` 的实现即如此） |

## PPT 命令速查

| 命令 | 参数 | 作用 |
|---|---|---|
| `doctor` | — | 自检：ProgID 安装/在跑、wpp 进程数、开着的演示文稿 |
| `status` | — | 实例与文稿清单（name / path / saved / readOnly / slides / active） |
| `open` | `-Doc <路径> [-Show] [-ReadOnly] [-NoHint]` | 已开则复用；未开则打开；**路径不存在则新建空稿并保存到该路径**；返回 slides 与 slideSize。接管标记：`open` 只打窗口标题（v1.26），**每页的页面标记由 `hint` 贴**（v1.27） |
| `read` | `-Doc [-From 1] [-Count 30]` | 逐页：layout + 形状清单（编号/名称/类型/占位符/l/t/w/h/文字前 160 字）+ 备注；>30 页自动分页 |
| `search` | `-Doc -Text <词> [-Max 50] [-Case]` | 跨页找文字：返回 页号/形状名/片段 |
| `replace` | `-Doc -Find <词> [-Replace <新词>] [-Scope all\|slides\|groups\|tables\|notes\|footer\|master] [-Slides "1-5,8"] [-Regex] [-Case] [-Max 500] [-DryRun]`；**批量：`-Items <json>`** | **全稿查找替换**（Word 侧 `revise`、Excel 侧 `write` 的对等能力）：扫**正文 / 组合内部 / 表格单元格 / 备注页 / 页眉页脚**；**母版默认不动**，要改必须显式 `-Scope master`（一改波及全稿）。默认**忽略大小写**（`-Case` 区分）、`-Regex` 支持 `$1`；**每处写入前用 `Characters()` 回读校验**，不符进 `skipped` 绝不盲写；返回逐处 `before→after`+`context`，并自动存 `prereplace` 快照，`restore -Since prereplace` 一键盘回滚。`-Replace` 省略 = 删除匹配 |
| `selection` | `-Doc` | 用户当前所在页与选中的形状名（需窗口开着） |
| `slide` | `-Action add [-Preset <档位>] [-Title ...] [-Items <json>] [-Fit contain\|cover] [-L -T -W -H] [-At N]`；`-Action reuse -From <源页> [-At N] -Items <json>`；`-Action dup\|move\|delete -Index N [-At\|-To M]` | **预设版式建页**（几何随文稿尺寸自适应；预设文本框已命名 title/subtitle/number/body/left/right，后续 `-Shape` 直达）。**v1.31 新增 8 种成品版式**（数据走 `-Items <json>`，中文走 JSON 文件；标题可放 `-Title` 或 JSON 里的 `title`）：`cards`（2–6 张等高卡片，`[{head,text?,value?}]`）、`stats`（2–4 个大数字，`{stats:[{value,unit?,label}]}`）、`timeline`（3–5 步横向时间线、文字上下交错，`{steps:[{when,head,text?}]}`）、`process`（3–5 步流程块+箭头，`{steps:[{head,text?}]}`）、`compare`（左右对比，`{left:{head,points},right:{…}}`）、`summary`（2–4 条编号总结，`{points:[{head,text?}]}`）、`gallery`（2–6 图带题注，`{images:[{path,caption?,fit?}]}`，默认 `contain` 整图不裁）、`split`（半页满版图+半页文字，`{title,text?,points?,image,side}`）。统一带标题、色杠、页脚与**普通文本框页码**（不用 WPS 页码占位符——那玩意儿每页显示 "1"）；返回 `{preset,shapes[],overflowWarnings[],fontSizeDrops[]}`，**`overflowWarnings` 非空就按提示改数据或换档位**。**v1.32 再补两类**：`canvas`（自由版面，`{title,items:[{type,box:[x,y,w,h],…}],edges:[…]}`，item 支持 text/heading/card/panel/line/arrow/image/stat/bullets/steps/callout/quote/node/chart；**`node`+`edges` 用 `AddConnector` 真连到形状边**，`connected.begin/end` 回读，连不上必进 `warnings`；越界自动夹取并逐条报）与 `grid`（结构化版面，`{title,rows:[[part,…],…]}`，`part` 有 `span`，宽度按 span 均分、行高按内容估、chart/image/panel/stack 吸收余量，支持 `stack` 竖叠；**不用给坐标**）。`-Action reuse`（v1.32）= 复制一页后按**形状名**换文字/换图（`{texts:{名:新文字},images:{名:新图}}`），保留版式与连线；替换不到的键进 `missing` 并警告；**复制页的页码会自动改成「新页号 / 总页数」**。旧的 `title\|section\|bullets\|two\|end` 仍可用 |
| `split` | `-Doc -Chunks "1-20,21-45,46-80" -Out <目录> [-Prefix 名字] [-DryRun]` | **拆分册**：每个区间存成一个独立文稿（`<前缀>-part01.pptx`…），**源稿不动**。区间必须从 1 起连续覆盖全稿——**有空隙直接报错**（静默丢页比报错更贵）。新稿的页尺寸先按源稿设好再粘贴，否则整册被缩放 |
| `merge` | `-Doc <目标稿> -Sources "a.pptx,b.pptx"` 或 `-Items <json: {"sources":[…]}>` `[-After N] [-DryRun]` | **合并多稿**：把来源稿的页依次插到目标稿 `-After N` 之后（默认追加到末尾），插完关掉来源稿。**每页保留它原来的版式**；要全篇统一外观再跑 `theme -Apply`，交付前跑 `check`。返回 `from/to` 页码范围，撤销就删这段再保存 |
| `text` | 单条：`-Slide N [-Shape 名\|-ShapeIdx n] (-Text \| -TextFile) [-Expect 原文开头]`；**批量：`-Items <json>`** | 改形状文字；`expect` 锚校验不符则**整批一条不改**（`ANCHOR_MISSING`）；默认解析到标题占位符，否则第一个有字形状——**强烈建议显式 -Shape**；每条返回 before/textAfter，**立刻比对** |
| `textbox` | `-Slide N (-Text\|-TextFile) [-Name 名] [-L -T -W -H pt]` + 格式项 | 加文本框（`-Name` 命名，后续定位用；中文字体/字号/对齐一起给） |
| `image` | `-Slide N -Path <图片> [-L -T] [-W/-H pt]` | 插图（缺省居中）；返回最终尺寸；图占页宽 <20% 时提醒字母可能太小。**路径含中文也行**：内部先拷到 ASCII 临时目录再插（`AddPicture` 不认非 ASCII 路径） |
| `picture` | `-Action list [-Slides "1-5"]`；`-Action replace -New <图> [-Match "name:X"\|"size:56x57"\|"pos:0,0" \| -Ids "12,3"] [-Slides] [-DryRun] [-FitInside] [-KeepCrop]` | **全稿换 logo / 换效果图**：`list` 先看清有哪些图（页/形状/路径/几何/是否链接/是否在组合里），`replace` 按 `-Match` 或 `-Ids` 选中后**沿用原几何**换图（先插新图→还原位置尺寸→再删旧图），并**默认清掉旧裁剪**（新图通常不该继承旧图的裁切）。`-FitInside` 改成"等比缩进原框内"、`-KeepCrop` 保留旧裁剪。链接图走 `LinkFormat` 直接换源、几何不动。⚠️ **组合内的图只列不换**：子形状坐标是组合局部坐标，COM 拿不到 chOff/chExt，硬换会把 logo 丢到别处——这类会进 `skipped` 并说明原因 |
| `comment` | `-Action add -Slide N [-Shape 名] (-Text\|-TextFile) [-Author AI]`；**批量：`-Items <json>`**；`-Action list`；`-Action delete -Slide N (-Index n \| -All)` | **PPT 批注**（与 Word 侧对齐）：锚到形状时自动在正文前加 `@形状名：`。⚠️ WPS 会自己摆批注钉位置（传 100,100 读回 5,5），**所以锚点信息必须写在文字里**，判据就是 `list` 里能读到 `@形状名`。批量中途失败**整批回滚**（`rolled back N`） |
| `export` | `-Format pdf\|png\|jpg\|md [-Slides "1-5,8"] [-Out <路径或目录>] [-Width 1600] [-Quality 88]` | **交付导出**：`pdf` 走 `SaveCopyAs(...,32)`（真 PDF，且**不改打开中文稿的身份**）；`png/jpg` 走剪贴板→helper 写盘（**明文**，页范围生效）；`md` 直接出大纲（等于 `outline`）。⚠️ **本机 PDF 落盘是密文**，返回里 `encrypted:true` 会明确提示，发人前先解密 |
| `table` | `-Slide N -Data <json: {"rows":[[...]]}> [-L -T -W -H] [-Size 14]` | 建表并填内容（首行默认加粗表头） |
| `shape` | `-Action add -Type rect\|roundrect\|ellipse [-L -T -W -H] [-Fill 颜色] [-Line 颜色] [-Text ...] [-Name 名]`；`-Action delete -Name <名>` | 画色块/圆角矩形/椭圆（可带字）；或按名删形状。**`-Name` 是 `connect` 的前提** |
| `crop` | `-Slide -Shape <图片名> [-Path <图>] [-Focus top\|center\|bottom\|left\|right] [-Fit cover\|contain] [-L -T -W -H] [-Name]` | **满版套图/人像裁切**（v1.29）：`cover`=铺满裁溢出（默认），`contain`=完整显示居中；默认 `focus` = 水平居中 + **垂直保留上 1/5**（人像保头），裁掉 >60% 时返回 `cut` 点名。⚠️ **单位是点不是比例**（写 0.14 会被存下但渲染不裁）；**裁完不要再改 W/H**（会洗掉裁切）；返回 `box`/`boxActual`/`cropReadBack`，**必须 `render` 目检** |
| `smartart` | `-Action list [-Match <关键字>]`；`-Action add -Slide N -Preset <版式> -Items <json> [-L -T -W -H]` | 原生 SmartArt（v1.29）：`-Preset` 认 process/cycle/hierarchy/org/pyramid/venn/list/matrix/funnel/target/gear/timeline/arrow 等英文关键字、中文关键字或版式全名（134 种，中文名）；`-Items` 每项 `{text,children}`。⚠️ **本机 WPS 12.1 能建图形但节点文字写不进**（`SmartArt.AllNodes` E_FAIL）⇒ 返回 `nodes:0` + `note`；要可控框图用 `shape` + `connect` |
| `connect` | `-Slide N -FromShape <名> -ToShape <名> [-Kind straight\|elbow\|curve] [-Arrow none\|end\|both] [-Label <文字>] [-FromSide/-ToSide top\|left\|bottom\|right]` | **形状之间拉连接线**（v1.29，框图用）：边自动选（左右相邻→右接左，上下相邻→下接上），`BeginConnected/EndConnected` 回读判真连；返回 `connected.begin/end` 与 `warning`（未真连时线不跟随框移动）。⚠️ **`-From/-To` 是页码范围**，连线两端只能用 `-FromShape/-ToShape`；`-Arrow both` 的起始箭头在 WPS 未必生效 |
| `bg` | `-Slide N\|-All [-Color #RRGGBB] [-Gradient [-Color]] [-Texture paper\|grid\|dot\|gradient] [-Image <图>] [-Reset]` | **页面背景/纹理**（v1.29）：纹理=生成一张 1600×900 PNG 再 `Background.Fill.UserPicture`（不碰 COM 纹理填充，WPS 上真生效）；`-Gradient` 写入不报错但 `Fill.Type` 回读仍 1 ⇒ **不可信，优先 `-Texture gradient`**；`-Reset` 交还母版 |
| `anim` | `-Slide N\|-All [-Transition none\|fade\|push\|wipe\|split] [-Effect fade\|fly-in\|wipe\|zoom\|appear] [-OnClick] [-Duration 0.6]` | **切换与入场动画**（v1.29）：返回 `transition.matches`（写入=回读）、`MainSequence` 前后条数、`unsupported`。⚠️ **只能证明写进去了，不能证明放映真播**——要交付就 F5 人看一眼；`unsupported:true` 别当成生效 |
| `format` | `-Slide N -Shape <名> `+ 格式项 | 改文字格式（字体/字号/加粗/斜体/颜色/对齐/填充），返回回读值 |
| `font` | `-Doc [-Font 宋体] [-From N] [-To M]` | **全稿统一字体**：逐页扫所有文字（**含组合内部与表格单元格**），只写 `Font.Name`+`NameFarEast`，**不动字号/颜色/加粗/版式**；返回 `shapes`/`groups`/`tables`/`failed` 与样例。⚠️ 换字体常使**字身变宽而多折行**，跑完必须 `render -Overview` 复核，个别挤爆的框用 `format -Size` 收一档 |
| `notes` | `-Slide N (-Text\|-TextFile\|-Clear)` | 写/清演讲者备注 |
| `view` | `-Slide N` | 把 WPS 窗口切到那一页（唯一"跟随"手段，不动内容） |
| `render` | `-Doc [-From 1] [-To 4] -Out <目录> [-Width 1100]`；**总览：`-Overview`（每次 ≤24 页，6 页拼一张并带页码）** | 逐页 EMF→PNG（**每次最多 4 页**）；写盘者是 helper，**明文可看**；`-Overview` 是"交付前过一遍全稿"的经济方式。**v1.32：render / export png\|jpg 前后会保存/恢复你的剪贴板**（`Slide.Copy` 用完不吞你复制的内容，恢复内容不进 Win+V 历史） |
| `theme` | `-Set <json 文件>` / `-Apply` / `-Reset` / 不带选项=查看 | **整稿风格打底**：primary/accent/title/text/muted 五色 + font/titleFont + footer（页脚文字）+ pageNumber + skipFooterOnCover；**v1.31 新增可选 token**：`accent2`（次强调色）、`surface`（卡片底色基色，默认 `#F2F5F9`）、`cardFill`（卡片底色，缺省跟随 `surface`）、`cardRadius`（卡片圆角 pt，默认 12）、`sectionStyle`（章节页样式 `solid\|side\|band\|number`，缺省 `side`）；旧主题文件缺字段自动取默认值。`-Set` 存到该文稿主题文件（%LOCALAPPDATA%\wps-office-cooperate\ppt-themes），之后所有预设页自动跟随；`-Apply` 把已有页的字体/配色刷一遍并补页脚页码（**落在深色色块上的文字自动改白**） |
| `master` | `-Action list`；`-Action footer [-Text 页脚] [-PageNum] [-Date] [-SkipCover] [-Slides] [-Clear]`；`-Action layout -Layout 1..11 [-Slides] [-DryRun]`；`-Action rebase -Path <模板稿> -Out <新稿> [-Layout n] [-KeepTemplatePages] [-Slides]` | **版式 / 页脚 / 套外来模板**。`list` 给出设计、版式清单与每页现状；`footer` 批量页脚+页码+日期（封面可跳过，`-Clear` 全清）；`layout` 批量换版式；`rebase` **以外来模板为底，把本稿的页搬进去**，内容页继承模板的母版与版式，模板自带的示例页默认删掉。⚠️ 换版式**只能用 `-Layout 1..11` 旧常量**（`Slide.CustomLayout = 对象` 在 WPS 里静默无效，`ApplyTemplate` 返回成功但什么都不做）；换完必须 `render` 看有没有丢占位符内容 |
| `chart` | `-Action add -Slide N -Data <json> [-Type column\|barh\|line\|pie\|area\|doughnut\|scatter\|radar] [-Title] [-XTitle] [-YTitle] [-Legend] [-LegendPos bottom\|top\|left\|right\|corner] [-Trendline] [-DataLabels] [-L -T -W -H]`；`-Action list` | **原生可编辑图表**（不是图，是图表对象）：数据经图表**内嵌工作簿**写入，所以读者能就地改数、换类型、重设样式。⚠️ **必须有窗口**：`open -Show` 之后再画，无窗口时 `AddChart` 直接 E_FAIL（`NO_WINDOW`）。WPS 只收**字面值**（不放公式）；种子是 3 系列×4 点，多余系列会删、点位靠**增删工作表行**对齐（`SetSourceData` 被拒）；超过 3 个系列会明确报"未写入"而不是静默丢。**v1.29 扩展**：`-Trendline`（线性趋势线，`equationShown` 看方程/R² 是否真显现）、`-DataLabels`、`-LegendPos`；**`-Type radar` 在 WPS 是"名义支持"**（回读 -415 但渲染成柱形）——要雷达图用形状自绘或换 `line` |
| `audit` | `-Doc [-From/-To] [-MinFont 12]` | **机械体检**：文字溢出文本框、字号过小、形状越出页面、形状重叠（装饰色块除外）、图片 <20% 页宽、空页、无标题页。返回到 页/形状/类型/细节。⚠️ **人工制作的老稿误报率高**：整幅背景色块会报"重叠/越界"，标题用普通矩形会报"无标题页"，页面 logo 会报"图过小"——这些是设计意图，**结论以 `render` 目检为准**，audit 只当线索清单 |
| `check` | `-Doc [-Profile report\|review\|bid\|internal] [-Rules all\|overflow,outframe,…] [-MinSeverity info\|minor\|major\|blocker] [-Config 规则.json] [-Out 报告.md\|.json] [-ExitCode]` | **交付前门禁**（audit 的判据版，**只读零风险**）：每条发现带 rule + 严重度（blocker/major/minor/info）+ fix 建议，返回 `summary`/{`issues`}/{`suppressed`}/`fonts`/`verdict`。规则库：`overflow` 文字撑破框、`outframe` 出页面、`placeholder` 残留"单击此处添加"、`fontsize` 标题<28/正文<14/脚注<10（按 profile 变）、`pagenum` 非封面缺页码/封面有页码、`title` 标题字号不一致、`contrast` 对比度<4.5:1、`density` 单页>300 字或>40 形状、`image-small` 图<20% 页宽、`empty` 空页。`-Config` 可覆盖阈值 + 按 `slide`/`shape`/`rule` 白名单抑制（抑制项**单列**在 `suppressed`，不消失）；`-ExitCode` 有 blocker 时返回码 3（给 CI/脚本用）；`-Out` 出 markdown 或 json 报告 |
| `outline` | `-Doc [-Out 大纲.md\|.json]` | **导出大纲**：每页标题 + 要点 + 备注（按上→下、左→右的阅读序）。给"AI 写讲稿"这条链路供料：命令只负责抽取，写讲稿交给模型，`notes` 写回。⚠️ **中文参数别走命令行**（PS 5.1 会把 `-Profile 投标` 变成乱码），profile 用 ASCII 别名 `report/review/bid/internal`；中文一律走 JSON 文件 |
| `align` | `-Slide N -Mode left\|right\|top\|bottom\|center-h\|same-width\|same-height\|distribute-h\|distribute-v [-Shapes "a,b,c"]` | 对齐/等距分布/统一尺寸（`-Shapes` 省略=该页全部形状）；返回每个形状的最终几何 |
| `zorder` | `-Slide N (-Name 名 \| -ShapeIdx n) [-Where front\|back]` | 置顶/置底（图片压字时用） |
| `restore` | `-Doc [-Since 快照\|tag] [-Shapes "3!备注,4!title"]` | **文本级回滚**：把文字退回快照状态（版式/新增形状不还原）；`-Since prereplace` = 退回上一次 `replace` 之前（含组合/表格/备注/页脚）。改坏了先回滚再想别的 |
| `snapshot` | `-Doc [-Tag 名]` | 存文本级快照（每页形状名+文字，默认存 %LOCALAPPDATA%\wps-office-cooperate\ppt-snapshots） |
| `diff` | `-Doc [-Since 快照\|tag]` | 与快照比：哪页哪个形状的文字变了（before→after）、页增删 |
| `save` | `-Doc [-As <路径>]` | 保存 / 另存（`.pdf` = 导出 PDF） |
| `close` | `-Doc [-Save]` | 关闭；`-Save` 会先保存（有改动才存，返回 `saved:true`）；最后一份时退出实例，回报 wpp 进程数。**关闭后同一路径可以再 `open`** |
| `quit` | — | 全部关闭并退出（不保存）。**慎用** |
| `finish` | `[-Note "做了什么"] [-Toast] [-Sound] [-Quiet]` | 完成信号（托盘气泡+提示音）；**不碰文稿任何内容**。任务收尾默认调用 |
| `hint` | `-Doc <路径> [-Text "正在做什么"] [-SlideMark\|-NoSlideMark] [-Toast] [-Sound] [-Quiet] [-Clear] [-Tray]` | **接管提示**（v1.11）：气泡/托盘图标 + 窗口标题标记（v1.26）+ **按需的页面标记（v1.27 引入、v1.28 改为 opt-in）**。**默认零改动**，响应 `inAppHint` 说明"要界面可见加 `-SlideMark`"；`-SlideMark` 逐页贴「〔AI接管中〕」并返回 `slideMarkCount` + `slideMarkWarning`；`-NoSlideMark`/`-Clear` 移除；`save`/`export` 落盘前自动剥离；**用户 Ctrl+S 造成的污染由下一条命令自动修复**（`badgeRepaired:true`，只读文件给 `badgeDirty`）。COM 标题仍不可写 → `captionSupported:false` |

所有命令都可用 `-TimeoutSec <秒>`（默认 180）；超时返回 `TIMEOUT`，多半是 WPS 弹了对话框。

### 套版图 / 框图 / 背景（v1.29，实测可跑）

```powershell
# 满版人像（保头）+ 右侧半页图（保下沿）+ 图表完整显示
ppt crop -Doc deck.pptx -Slide 1 -Path portrait.jpg -L 0 -T 0 -W 960 -H 540 -Focus top
ppt crop -Doc deck.pptx -Slide 2 -Shape "图片 3" -L 480 -T 0 -W 480 -H 540 -Focus bottom
ppt crop -Doc deck.pptx -Slide 2 -Path chart.png -L 40 -T 360 -W 480 -H 150 -Fit contain

# 框图：先 shape -Name 摆框，再 connect 拉线（边自动选，也可 -FromSide/-ToSide 强制）
ppt shape -Doc deck.pptx -Slide 3 -Type roundrect -L 80 -T 80 -W 200 -H 90 -Text 数据层 -Name BoxA -Fill #1F4E79
ppt shape -Doc deck.pptx -Slide 3 -Type roundrect -L 620 -T 320 -W 200 -H 90 -Text 应用层 -Name BoxB
ppt connect -Doc deck.pptx -Slide 3 -FromShape BoxA -ToShape BoxB -Kind elbow -Arrow end -Label 接口

# 背景/纹理（-All 整稿；不给 -Slide 也默认整稿，与其它命令不同）
ppt bg -Doc deck.pptx -All -Texture paper
ppt bg -Doc deck.pptx -Slide 1 -Texture grid -Color #F2F6FB
ppt bg -Doc deck.pptx -All -Reset
```

**纪律**：三类改动都要 `snapshot` 打底（PPT 无撤销栈）；每步之后 `render` 抽页目检；`crop` 的 `cut` 字段点名"裁掉超过 60% 的图"必须看，必要时改 `-Focus` 或 `-Fit contain`。

### 成品版式建页（v1.31，8 档，数据走 JSON）

```powershell
# 一次定主题（含新 token：cardFill / cardRadius / sectionStyle）
ppt theme -Doc deck.pptx -Set theme.json
# 逐档加页：-Items 指向 JSON 文件（中文只在文件里，别塞命令行）
ppt slide -Doc deck.pptx -Action add -Preset cards    -Items cards.json      # {"title":"…","cards":[{head,text,value}]}
ppt slide -Doc deck.pptx -Action add -Preset stats    -Items stats.json      # {"title":"…","stats":[{value,unit,label}]}
ppt slide -Doc deck.pptx -Action add -Preset timeline -Items timeline.json   # {"title":"…","steps":[{when,head,text}]}
ppt slide -Doc deck.pptx -Action add -Preset process  -Items process.json    # {"title":"…","steps":[{head,text}]}
ppt slide -Doc deck.pptx -Action add -Preset compare  -Items compare.json    # {"title":"…","left":{head,points},"right":{…}}
ppt slide -Doc deck.pptx -Action add -Preset summary  -Items summary.json    # {"title":"…","points":[{head,text}]}
ppt slide -Doc deck.pptx -Action add -Preset gallery  -Items gallery.json    # {"title":"…","images":[{path,caption}]}
ppt slide -Doc deck.pptx -Action add -Preset split    -Items split.json      # {"title":"…","image":"…jpg","side":"left"}
ppt render -Doc deck.pptx -Overview -Out shots      # 60 秒看完全稿
```

- **数据放 JSON、标题也放 JSON**（`title` 字段）：七个内容档位都要求标题，缺了会报 `BAD_ARGS … needs -Title (or "title" in the -Items object)`。
- 统一外观：标题 28pt（`**强调**` 会用 accent 加粗）、色杠 + 细线、页脚、页码「n / N」；正文 14pt，放不下会**自动降一档字号**并记进 `fontSizeDrops`。
- **`overflowWarnings` 非空就别交**：要么删字、要么换档（例如 5 条要点拆成 two/compare）。形状是 `gallery` 默认 `contain`（整图不裁）；要统一比例裁切用 `-Fit cover`。
- 图片裁切走 **GDI+ 预裁再按框插入**（不碰 `PictureFormat.Crop*`——WPS 上它按点解释、且改宽高会把裁切洗掉，实测 1200×600 进 480×540 会变成 552×540 出页）。

### 自由版面 / 结构化版面 / 形状图表 / 复用页（v1.32）

```powershell
# canvas：给坐标，节点之间自动连线（AddConnector 真连到形状边）
ppt slide -Doc deck.pptx -Action add -Preset canvas -Items canvas.json
# canvas.json = {"title":"…","items":[{"type":"node","id":"A","box":[60,150,220,80],"text":"采集层"},…],
#                "edges":[{"from":"A","to":"B","kind":"elbow","label":"原始数据"}]}

# grid：不给坐标，按行/span 排（左图右文、2×2 卡片、上表下结论）
ppt slide -Doc deck.pptx -Action add -Preset grid -Items grid.json
# grid.json = {"title":"…","rows":[[{"type":"image","span":1,"path":"…","fit":"cover"},
#                                   {"type":"stack","span":1,"items":[{"type":"heading","text":"…"},{"type":"text","text":"…"}]}],
#                                  [{"type":"card","head":"…","value":"12"},{"type":"card","…":…}],
#                                  [{"type":"callout","text":"结论一"},{"type":"callout","text":"结论二"}]]}

# 形状绘制图表（不建原生图表对象，塞进任意框；canvas/grid 的 item {"type":"chart","chart":{…}}）
#   scatter（可 fit:true → 最小二乘线 + 方程 + R²）、curve（末端标系列名）、pie、doughnut

# 复用一页：复制后按形状名换文字/换图，版式与连线保留，页码自动改成「新页 / 总页」
ppt read -Doc deck.pptx -From 4 -Count 1                     # 先看形状名
ppt slide -Doc deck.pptx -Action reuse -From 4 -At 7 -Items reuse.json
# reuse.json = {"texts":{"title":"新标题"},"images":{"pic1":"C:\\…\\new.png"}}
```

- **canvas 不会替你避让**：框重叠、文字压边不会自己挪（只报越界与连接失败），重叠要自己调 `box`；`grid` 才是"不用管坐标"的那条路。
- **形状图表一张约 45–60 个形状**：`check` 的 `density` 规则会报，交付时按规则抑制（`-Config`）或把点数控制在 10 个内。要"读者能改数"的原生图表仍用 `chart` 命令（代价：必须有窗口、最多 3 系列）。
- **`reuse -At` 是尽力而为**：WPS 对刚 `Duplicate()` 出来的页 `MoveTo` 被静默忽略（独立 `slide -Action move` 有效），所以返回里 `added` 是实际落点、`requestedAt` 是你要的位置；要精确位置再补一条 `slide -Action move -Index <added> -To <目标>`。
- **主题写不进 pptx**：实测 WPS 拒绝 `Presentation.Tags`（`Tags.Add` 不报错但 `Count` 恒 0）⇒ 主题仍存 `%LOCALAPPDATA%\wps-office-cooperate\ppt-themes\`，**换机要把这个 JSON 一起带走**再 `theme -Set`；`theme -Probe` 可复核该结论。
- `slide -Action reuse` / `dup` 内部会**保存并恢复你的剪贴板**（`Duplicate` 借剪贴板，v1.32 修掉了"复用后再 render 必崩"的问题）。

### `text -Items` 批量格式

```json
{"items":[
  {"slide":4,"shape":"title","expect":"谢谢","text":"谢谢聆听"},
  {"slide":3,"shape":"备注","text":"集水面积 1520 平方公里"}
]}
```

- `shape` 用 `read` 里看到的名称（预设页的形状已命名）；不填则解析到标题占位符，否则第一个有字形状。
- `expect` 可选但强烈建议：该形状当前文字必须**以它开头**，不符即整批拒绝。
- 中途 COM 出错会按序回滚已改项（"rolled back N item(s)"）。

### `replace` 用法与边界（**批量改字必读**）

```powershell
# 一条规则
ppt replace -Doc 稿.pptx -Find "2024" -Replace "2025" -DryRun      # 先只看清单
ppt replace -Doc 稿.pptx -Find "2024" -Replace "2025"
# 多条规则一次跑（批内按顺序，后一条看到前一条的结果）
ppt replace -Doc 稿.pptx -Items 规则.json
```
```json
{"items":[
  {"find":"2024","replace":"2025","label":"年份"},
  {"find":"XX水利水电勘测设计院","replace":"中铁水利水电规划设计集团"},
  {"find":"(\\d{4})年(\\d{1,2})月","replace":"$1年$2月","regex":true,"label":"日期规范化"}
]}
```

- **返回什么**：`scanned`（扫了多少页/形状/组合/表格单元格/备注/页脚/文字槽）、`changed`、`shapesTouched`、`hits[]`（`slide/where/shape/path/start/before/after/context`）、`skipped[]`（带 `reason`：`cross-paragraph` / `unverified-span`）、`preSnapshot`。
- **`-DryRun` 先跑**：`changed` 就是"会改几处"，`applied:false`，**不落盘、不写快照**，`diff` 为空。给人看清单就是这个模式。
- **跨 run 命中能改**（"2024"被拆成两个格式不同的 run 也能替换），靠的是写入前 `Characters(start,len)` 回读比对；**跨段落命中一律跳过**（`cross-paragraph`），因为替换会把两段并成一段，宁可不改。
- **保格式**：替换后把 span 首字符的字体（字号/latin/中文字体/粗体/斜体/下划线/颜色）回填到新文字，且**只回填真正不同的项**。原因：WPS 在 span 覆盖整段 run 时会拿**后面**那个 run 的字体，实测 26pt 绿字会被写成 16pt 灰字——不回填就是肉眼可见的掉格式。
- **母版默认不动**：`-Scope` 默认 `slides,groups,tables,notes,footer`。改母版就 `-Scope master`（会扫 SlideMaster + 全部 CustomLayouts），**改完必须 `render -Overview`**，因为措辞会出现在每一页。
- **页脚只在已显示时改**：`HeadersFooters` 里隐藏的页脚会被跳过——为了改一个词而把页脚翻出来显示，等于偷偷改了版式。
- **幂等**：同一条命令跑第二次 `changed:0`；`changed:0` 且 `skipped` 为空时返回里会提示先查 `-Case`/`-Scope`/全角半角。
- **回滚**：`restore -Doc -Since prereplace` 把**整批**退回（含组合/表格/备注/页脚）。快照是"上一次 replace 之前"的那份，连跑两批只会留最后一批的。

### 原生图表（`chart`）

```powershell
ppt open -Doc 稿.pptx -Show            # 必须有窗口，否则 AddChart 报 NO_WINDOW
ppt chart -Doc 稿.pptx -Slide 3 -Data 数据.json -Type column -Title "实测流量" -YTitle "m³/s" -XTitle "年份"
ppt chart -Doc 稿.pptx -Action list    # 已有图表清单（类型/系列/几何）
```
```json
{"categories":["2020","2021","2022"],
 "series":[{"name":"流量","values":[1200,1580,2248]},
           {"name":"设计值","values":[1100,1500,2100]}]}
```
- **为什么用图表而不是贴图**：数据留在稿子里，读者能改数、切类型、重设样式；`outline`/`check` 也照常工作。
- **WPS 的硬约束**（都实测过）：① 只在**有窗口**的稿里能建图表；② 只收**字面值**，公式不生效；③ `Series.Values` 写了不落盘，只有写内嵌工作簿才算数；④ `SetSourceData` 被拒，点数/系列数靠**增删工作表行**和删多余系列来对齐；⑤ `NewSeries` 不可用，超过种子（3）的系列会被告知未写入。
- 建完 **必须 `render` 那一页**：默认几何是"页宽 55%×页高 45% 居中"，字号偏小是常态，必要时 `-W/-H` 调大或 `format` 调字号。

### 套外来模板与批量页脚（`master`）

```powershell
ppt master -Doc 稿.pptx                                  # list：设计/版式清单 + 每页版式与页脚状态
ppt master -Doc 稿.pptx -Action footer -Text "中铁水利 · 2025" -PageNum -SkipCover
ppt master -Doc 稿.pptx -Action footer -Clear            # 全清（页码、页脚、日期）
ppt master -Doc 稿.pptx -Action layout -Layout 6 -Slides "2-9"   # 换版式（旧常量）
ppt master -Doc 外单位稿.pptx -Action rebase -Path 我们的模板.pptx -Out 套好模板.pptx -Layout 6
```
- **套模板的方向很重要**：把**内容页粘进模板稿**，内容才继承模板的母版；反过来往空白稿里粘只会把你的设计丢掉（实测）。`rebase` 走的就是前者，源稿不动。
- **`-Layout` 只认 1..11 旧常量**（1 标题页、2 标题和文本、3 节标题、4 两栏、5 比较、6 仅标题、7 空白、8 带标题内容、9 带题注图片、10 竖排标题、11 竖排标题和文本）。传 `CustomLayout` 对象在 WPS 里**静默不生效**。
- **页脚/页码是"每页占位符"**：开了以后每页会多出 `页脚占位符`／`灯片编号占位符` 两个形状（WPS 不写 `<p:hf>`）。所以开页脚会改变每页的形状数，`check -density` 会跟着变。
- **换版式会重排占位符**：`layout` 返回每页 `shapesBefore/shapesAfter` 与 `layoutBefore/layoutAfter`，形状变少的页会进 `warnings`——那正是内容可能掉的地方，必须 `render` 复核。
- 模板与源稿尺寸不一致时，`rebase` 会给 `warning`：形状只按绝对坐标搬，大页内容塞进小页会溢出。

### 分册拆分与多稿合并（`split` / `merge`）

```powershell
# 80 页拆 5 个分册：区间必须连续覆盖全稿，有缺口直接报错
ppt split -Doc 总报告.pptx -Chunks "1-16,17-32,33-48,49-64,65-80" -Out 分册\ -Prefix 上册
ppt split -Doc 总报告.pptx -Chunks "1-16,17-32" -Out 分册\ -DryRun     # 先看计划

# 三个县的稿合成一份（路径含逗号或不好写时走 JSON）
ppt merge -Doc 汇总.pptx -Sources "a县.pptx,b县.pptx,c县.pptx"
ppt merge -Doc 汇总.pptx -Items sources.json      # {"sources":["a.pptx","b.pptx"]}
ppt merge -Doc 汇总.pptx -Sources a.pptx -After 1 # 插到第 1 页之后，而不是末尾
```
- **别用 `Slides.InsertFromFile`**：本机实测它会**把目标稿清成 0 页**。这里走的是 `Slide.Copy()` + `Slides.Paste()`。
- **拆分册不能靠粘贴**：跨文稿粘贴**不带设计**（母版标记会丢）——所以 `split` 是"整份复制再删区间外的页"，设计、版式、备注母版全部保留（测试用母版标记逐份验过）。
- **页尺寸是"每份文稿"的属性**：合并时新页进目标稿会按目标稿尺寸摆放，源稿与目标稿尺寸不同要留神溢出。
- 拆/合都是**新建或追加**，不做破坏性覆盖：`split` 不动源稿，`merge` 只在目标稿插入，返回 `from/to` 页码，撤销就是删这一段。
- 合出来的稿**每页保留原版式**，所以外观可能花——要统一再 `theme -Apply`，交付前 `check`。

### `picture` / `comment` / `export` 三条常用链路

```powershell
# 全稿换 logo：先看清有哪些图，再按需替换
ppt picture -Doc 稿.pptx                       # list：页/形状/几何/链接/是否在组合里
ppt picture -Doc 稿.pptx -Action replace -New 新logo.png -DryRun
ppt picture -Doc 稿.pptx -Action replace -New 新logo.png -Match "size:56x57"

# 校审意见落到页面上（批量，失败整批回滚）
ppt comment -Doc 稿.pptx -Items 意见.json -Author AI
ppt comment -Doc 稿.pptx -Action list
ppt comment -Doc 稿.pptx -Action delete -Slide 4 -All

# 交付
ppt export -Doc 稿.pptx -Format pdf -Out 稿.pdf
ppt export -Doc 稿.pptx -Format png -Slides "1-2" -Out 图\ -Width 1600
ppt export -Doc 稿.pptx -Format md  -Out 稿.md
```
```json
{"items":[
  {"slide":4,"shape":"note1","text":"数据来源需补充：《江西省水文年鉴》2020—2024。"},
  {"slide":1,"text":"封面色带建议加深一档。"}
]}
```

- **换图的正确姿势**：`list` → 挑 `-Match`/`-Ids` → `-DryRun` 看清单 → 真跑 → `render` 目检。新图纵横比不同时，**几何沿用旧图**意味着画面被拉伸，这点只有肉眼能判。
- **组合内的图只列不换**（`skipped.reason = "inside a group"`）：子形状坐标是组合局部坐标，COM 不给 chOff/chExt，硬换必然丢位置。要换请先在 WPS 里把它移出组合。
- **批注锚点写在文字里**（`@形状名：…`）：WPS 会自己摆钉位置，`Left/Top` 传 100,100 读回 5,5，靠坐标对不上的东西不算锚点。
- **导出 PDF 会落密文**：返回 `encrypted:true` 就是提醒，发人前先解密（本机用 FastDycrypt，见 `reference/encryption.md`）。

### `check` 用法（交付前必跑，只读）

```powershell
ppt check -Doc 稿.pptx                                  # 全规则，出分级清单
ppt check -Doc 稿.pptx -MinSeverity blocker -ExitCode    # 门禁：有 blocker 就失败（返回码 3）
ppt check -Doc 稿.pptx -Profile bid                      # 投标场合：正文下限抬到 18pt
ppt check -Doc 稿.pptx -Config 院标规则.json -Out check.md
```
```json
{"titleMin":28,"bodyMin":14,"footMin":10,"densityChars":300,"densityShapes":40,
 "contrastMin":4.5,"overflowFactor":1.05,"imageWidthPct":20,
 "suppress":[{"slide":7,"shape":"背景 1"},{"rule":"contrast","slide":12}]}
```

- **严重度分级**，交付前只看 `blocker`：`blocker` = 撑破框 / 出页面 / 残留占位符；`major` = 字号过小 / 缺页码 / 空页 / 标题字号不一致；`minor` = 对比度 / 密度 / 图过小。
- **误报控制靠 `-Config` 白名单**：设计意图（整幅背景块、页脚小字、故意的小 logo）用 `slide`/`shape`/`rule` 抑制，被抑制的**列在 `suppressed` 里**而不是消失——下一篇同模板稿就能复用同一份配置。
- **`verdict`** 直接给结论：`not ready:` / `readable but rough:` / `mechanically clean`。
- **别把它当审美判据**：它只说"机械上能不能交"，内容对不对、好不好看仍要 `render` 目检。
- ⚠️ **查不出"缺字体"**：实测 WPS 在打开时就把不存在的字体替换成已装的（fixture 里 `Definitely-Not-Installed-Font` 经 COM 读回来是 `Arial`），信息在我们看到之前就没了。`fonts` 只报"WPS 解析成了什么"，换机前请在有目标字体的机器上打开看一眼——这条限制作为 `limits` 字段随结果返回，不假装能查。

### PPT 视觉自检闭环（**本轮真的改动过这份稿才走**）
前提是本轮确实动过这份稿（新建、改文字、换主题、调版式都算）。**只是连上、只是读、只是查，不走这一步**（见铁律 7）。顺序固定：

1. **`audit -Doc`** —— 机械筛查（溢出/字号/越界/重叠/空页），便宜但**对人工老稿误报多**（背景色块 / logo / 非占位符标题），先挑出真问题再修，别按条数硬修。
   要"能不能交"的判据而不是线索清单时用 `check -MinSeverity blocker`（分级 + `-ExitCode` 门禁）。
2. **`render -Overview -Out <目录>`** —— 每 6 页拼一张图，**逐张看**：标题层级是否统一、色块对不齐、要点与表格图挤在一起、半空页。
3. 对可疑页 **`render -From N -To M`** 单页放大确认。
4. 修（`text` / `align` / `table` / `theme -Apply`）→ **再 `render` 确认** → 才 `save`。

**能力不同时的降级**：当前模型能读图 → 第 2 步必须真看（`read_image`）；读不了图（某些模型不声明图像输入）→ 只能依赖 `audit` 并把 `render` 的结果路径交给用户看，**不要**声称"已核对版式"。

### PPT 典型工作流

1. **新建汇报稿**：`status` → `open -Doc 新路径 -Show`（自动建空稿）→ 可选 `theme -Set 主题.json`（一次定型：主色/字体/页脚）→ `slide -Preset title ...`（封面）→ `slide -Preset bullets -Items 提纲.json`（目录页）→ 逐节 `slide -Preset section -Number 01`。
2. **填充内容页**：`slide -Preset bullets` 打骨架（正文占 25%–75%，下方留给表格/图）→ `table -Data 表.json`（默认 90% 页宽、y=62%）→ `image -Path 图.png` → 每页 `render` 目检。
3. **改已有稿**：`read` 拿形状名 → `text -Items`（带 expect）批量改 → 比对 textAfter → `render` 抽查。
   **同一个词散落在几十页时不要用 `text` 逐条改**：`replace -Find ... -Replace ... -DryRun` 先出清单，确认后再去掉 `-DryRun` 跑一遍（一次覆盖正文/组合/表格/备注/页脚），`diff` 验收，改坏了 `restore -Since prereplace`。
4. **收拾版式**：`align` 对齐/等距、`zorder` 修压字、`audit` 兜底；要统一字体用 `font -Font <字体>`（一次全稿，含组合/表格），**换字体后必 `render -Overview` 查折行溢出**，挤爆的框用 `format -Size` 收一档。
5. **人机接力**：`snapshot -Tag r1` → 干活 → `diff -Since r1` 看人改了什么；改坏 `restore -Since r1`。
6. **收尾**：`theme -Apply`（统一配色页脚）→ `check -MinSeverity blocker`（交不交得出去）→ `audit` → `render -Overview` **看一眼** → `save`（提醒落盘可能是密文）→ `close`（汇报 wpp 数）→ `finish`。

## PPT 故障排查

| 错误码 / 现象 | 怎么办 |
|---|---|
| `NO_APP` | 没装 WPS 演示或 KWPP COM 没注册；手动开一次 WPS 演示再试。 |
| `NOT_OPEN` / `DOC_CLOSED` | 语义同 Excel：`-Doc` 路径与实际不符先 `status`；协作中被关闭不要静默重开，先问用户。**用户自己关掉后想继续**，直接 `open -Doc <同一路径>` 即可（显式 open 会重新打开并刷新标记） |
| `replace` 返回 `changed:0` | 依次排查：`-Case`（默认忽略大小写，你反而可能想要区分）、`-Scope`（母版/备注/页脚默认是否在范围内）、全角半角与原稿是否一致、`-Slides` 是否限窄了。返回里的 `note` 会给出这条提示 |
| `replace` 的 `skipped[].reason=cross-paragraph` | 匹配跨了段落标记（替换会把两段并成一段）；缩小 `-Find` 或手工改这两处 |
| `replace` 的 `skipped[].reason=unverified-span` | 写入前回读到的不是原匹配文字（通常是合并单元格被重复寻址，或用户同时手改了）；**没写入是好事**，`read` 后手工补 |
| `ANCHOR_MISSING` | 形状当前文字与 `expect` 不符（人刚改了/定位错形状）。整批没改，重新 `read` 核对形状名。 |
| `render` 报 `COM_ERROR`（clipboard） | 多半是弹了对话框占住剪贴板，或刚 Copy 完还没就绪。关掉对话框重试。 |
| 同页形状重名（旧稿"文本框 1"满天飞） | `-Shape` 会命中第一个；用 `-ShapeIdx`（read 里的 i 列）或先 `format` 不行就 `textbox -Name` 加命名框替代。 |
| 图片插入后糊/字小 | `render` 目检；图小于页宽 20% 时结果里会点名提醒。 |
| wpp.exe 变多 | 同其它 app：`status` 看实例，只清无窗口实例，汇报数字。 |

---

## 与其它 skill / 工具的关系

- **`office_*` 工具（dsh-office 插件）**：装了那个插件时，`office_open/office_read/office_edit/office_render` 与本 skill 能力重叠、互不冲突——插件适合"顺手改一份"，本 skill 适合"要分发给同事、不想让每人装插件"的场合。两者都可用时**选一个用到底**，别交替（会各开实例）。
- **`docx-paragraph-polish`**（离线润色）：**本机有 WPS 且用户要在场逐条过**时，改用本 skill Word 侧——修订直接出现在 WPS 窗口里、可当场接受/拒绝、顺手能核对版式；无人值守、上千段一次跑完的批量润色仍用 `docx-paragraph-polish`（不依赖 WPS，也不受"落盘即密文"影响）。
- **`office_xlsx` / `minimax-xlsx` / openpyxl 等**：它们改磁盘文件。工作簿正开在 WPS 里时**必须**用本 skill Excel 侧；无人值守批量跑、或文件没开且不需要人在场时，用离线方案更快。
- **`pptx-generator` / python-pptx 等**：同理——演示文稿正开在 WPS 里（或用户要在场边看边改）时用本 skill PPT 侧；从零生成、无人值守的整份 PPT 用离线方案。
- **`design-report-review` / `design-report-format`**：见铁律 0——同一份文档在本 skill 手里时，用本 skill 的 `audit` + `comment` / `format`/`style`/`page` 在原稿处理，不另起离线流程；用户明确只要旁挂报告、不动原文时才另说。

## 参考

- [reference/workflows.md](reference/workflows.md) —— Word 全部典型工作流分步教程 + 排版实测坑详解。
- [reference/multi-task.md](reference/multi-task.md)
- [reference/multi-task-excel.md](reference/multi-task-excel.md)
- [reference/multi-task-ppt.md](reference/multi-task-ppt.md)
- [reference/multi-task-session.md](reference/multi-task-session.md) —— 跨 app 任务板与波次调度（`tasks-session.ps1`）、物理约束与验收。 —— PPT 多任务：页号/形状锚点、全局任务优先、结构先做、`tasks-ppt.ps1` 三条命令与验收。 —— Excel 多任务：任务卡字段、`tasks-excel.ps1` 三条命令、合并/冲突规则、提交顺序与验收。 —— 多任务协作（多章节 / 同章多任务 / 章节+全文）：任务卡字段、`tasks.ps1` 三条命令、提交顺序、合并规则、实测坑与验收清单。
- [reference/word-quirks.md](reference/word-quirks.md) —— WPS Writer 与 Word 的实测差异表（冒充版本号、样式中文名、`Range` 是方法、关最后一个文档静默杀实例……）。
- [reference/excel-quirks.md](reference/excel-quirks.md) —— WPS 表格 COM 的实测差异与坑（冒充 Excel、SaveAs 身份变化、块数组 0 基、渲染走剪贴板……）。
- [reference/encryption.md](reference/encryption.md) —— 落盘即密文的现象、为什么 COM 路线不受影响、交付与解密链路。
- [reference/word-acceptance.md](reference/word-acceptance.md) / [reference/excel-acceptance.md](reference/excel-acceptance.md) —— 各版本验收记录。

