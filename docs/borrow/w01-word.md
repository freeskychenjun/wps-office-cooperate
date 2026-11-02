# w01 · dsh-office Word 域深读与可借鉴清单

> 对象：`D:\DSH工作区\_research\dsh-office`（v1.2.0，MIT，浅克隆；本次**只读**，未改其任何文件，也未启动/操作任何 WPS 实例）。
> 对照：本仓库 `SKILL.md`（**v1.28**，Word 段 163–302 行）、`scripts/word.ps1`、`scripts/helper/word/*.cs`、`DESIGN-word.md`、`reference/word-quirks.md`。
> 引用格式：`相对路径:行号`；**dsh-office 侧路径相对其仓库根**（`helper/OfficeHelper.cs` = 9893 行单文件），**我方侧路径相对本仓库根**。
> 环境事实：本机 WPS **12.1.0.29161**、无微软 Office（`WPS` 接管 `Word.Application`）；dsh-office 的 WPS 实测机是 **12.1.0.28505**（`docs/verification.md:688`）——跨版本结论一律按「需实测」看。
> 一个必须先说清的前提：dsh-office 是**给模型用的插件**（工具 = 6 个，Word 全部操作塞进一个 `office_edit` 的 ops 数组，参数校验靠一段自然语言 GUIDE + helper 侧兜底）；我方是**给 agent 用的命令行 skill**（每条命令一次进程/管道往返 + 单个 JSON）。所以它的**能力边界与 WPS 适配经验可直接照抄，实现方式必须按我方「一次块读、本地算、一次块写」的纪律重写**，否则会退化成几千次 COM 往返（我方 v1.13–v1.15 已付过这笔学费）。
> 另：dsh-office 的 `docs/verification.md` 只记录到 **1.1.2**，而**脚注（1.2.0 主打功能）在 verification.md 里没有任何条目**——脚注的验收证据只有 `scripts/footnotes-smoke.mjs`（含 `--wps` 分支）与 README 的叙述，这一点在评估「WPS 脚注可用」时要打折。
> **交叉核对的警告（很重要）**：我方已在**同一台机器、更新版本**上**推翻**过 dsh-office 的一条 WPS 结论——它说"WPS 没有 `Range.Cells`"（`docs/verification.md:697`），我方实测 `Table.Range.Cells.Count=9` 正常（`reference/word-quirks.md:22`，并留了一句"别因为别人的注释就绕路，先实测再决定绕法"）。所以下文所有 dsh-office 的 **WPS 专项结论都按"作者那台机（12.1.0.28505）+ 那条代码路径"的样本**看待：**凡是要投大工作量的（尤其 P1-1 公式），必须先在我方 12.1.0.29161 上做一次否定性验证**（§4.3 步骤 0）。

---

## §0 校准：Lead 在 WPS 12.1.0.29161 的实测（2026-11-02）vs 本文的证据

> 本节是全篇的**结论校准层**：凡与下表冲突的推断，一律以下表为准；下表未覆盖的，才用 dsh-office 的代码/文档推断。

| Lead 实测（WPS 29161） | 证据等级 | 对本文的影响 |
|---|---|---|
| `Footnotes.Add(range, '', text)` **可行**，`count` 与脚注文本可回读 | 已实测 | P1-5 脚注**可以做**；§1.3 的 COM 序列基本可用 |
| `FootnoteOptions.NumberingRule` **在 WPS 不存在** | 已实测 | **dsh-office 的 `footnote_settings`（`:2196-2205`）在 WPS 上必定失败**；其 `footnotes-smoke.mjs --wps` 里有两处 `footnote_settings` 断言（`:109-112`、`:122-132`），而 `verification.md` 里 1.2.0 **零记录** → 可判定：**WPS 分支的验收从未跑通（或从未跑）**，"WPS 支持脚注"实际只到"增删读"，不含改编号规则。我方实现必须把"编号规则"降级为**只读报告文档原设置**（`numberingRule` 读不到就报 `unsupported`，绝不写入） |
| `doc.OMaths` 对象存在但 `Add(range)` 报「值不在预期范围内」 | 已实测 | 与 `verification.md:695` 的"WPS 的 `OMaths.Add` 不按给定位置插入、`BuildUp` 不生效"**方向一致**（都是"WPS 自己建公式"失败）→ §4.3 的探针 A **可判定为不通**，不必再花半天 |
| `Range.InsertXML(OMML)` **调用成功但 `OMaths.Count` 仍为 0** | 已实测（但"没转成公式"是**推断**，未验证 XML 是否落进文档） | ★**这是唯一的关键未知**。dsh-office 恰恰用 `OMaths.Count` 判成功（`:1710`），而它在 28505 上验过 11 个公式（`verification.md:742`）→ 差异只可能在**插入的 XML 形态**或**版本**。§4.3 已改写为按此收敛，先做"三判据验真"（见 §4.3 步骤 1） |
| `Fields.Add(EQ \o(2,3))` 返回 `Result` 为空 | 已实测 | EQ 域路线**不可用**；顺带说明"用 EQ 域模拟公式"这条退路也没了 |
| `Shapes.AddShape` **可行**；`Shapes.AddConnector` + `BeginConnect` **E_FAIL** | 已实测 | 与 dsh-office 的做法**完全吻合**（`:3991-4011` 自己画线段；`verification.md:838` 明说"WPS 文字里连接符接不上形状，两个套件于是用同一套画法"）→ P2 的 `insert_diagram` 可做，但**要用 `Shapes.AddCanvas` + `CanvasItems` 那一层**（见 §1.5 最后一行），不是文档级 `Shapes.AddShape`；`AddCanvas` 是否可用**需再测一条** |
| `Application.UndoRecord.Start/EndCustomRecord` 调用不报错 | 已实测（"能整批撤销"待确认） | 我方 `SKILL.md:293` 已记"实测 WPS 支持分组"，两侧互证；`save` 清栈的坑另算 |
| `SEQ` 题注域可行（`Result=1`） | 已实测 | 与我方 `Tables.cs:1114-1119` 同路；dsh-office 用的是 `Application.CaptionLabels`+`InsertCaption`（`:2651-2656`），**不必改**，我方的 SEQ 写法在 WPS 上更稳 |

---

## §1 功能清单（op/字段 → 语义 → 实现位置 → 技巧/坑）

### 1.1 契约形态：一段 GUIDE + helper 侧字段点名

| 项 | 语义 | 实现位置 | 技巧/坑 |
|---|---|---|---|
| `GUIDES.word` | Word 全部 op 与字段写在一段约 25 行的自然语言里，第一次 `office_open/office_read` 随结果下发一次，`office_help {app}` 可再取 | `src/tools.ts:107-131`（下发给模型的时机见 `:230`、`:424`） | 上下文经济：**用不到的应用不占 token**；我方 SKILL.md 常驻，可考虑把 Word 长教程继续下沉 `reference/` |
| `Bag.Unread()` | 每个 op 的入参包一层 `Bag`，**凡是给了但没有任何 `Raw/Has` 读过的字段名，结果里点名** | `helper/OfficeHelper.cs:23-33`；使用点 `:9669-9670`（`— IGNORED field(s) a,b: set_text does not take them`） | 机制比"参数白名单"便宜：不需要 schema，只要"读过就记一笔" |
| 错误码 | `BAD_ARGS / ANCHOR_MISSING / ANCHOR_MISMATCH / STALE / NOT_FOUND / STYLE_MISSING / BUSY` | 分散：`:2137`、`:2140`、`:3260`、`:2458`、`:3726` | 我方错误码集合见 `DESIGN-word.md:58`，两边高度重合（可直接映射） |
| `office_edit` 批处理 | 一批 ops 顺序执行，**第一个失败就停**并报告已完成哪些 | `:9659-9683` | 与我方"整批原子"不同：**它是"做到哪报到哪"，我方 revise 是"一条不过全批不改"**——两条路线各有适用场景 |

### 1.2 可靠性与撤销

| op/字段 | 语义 | 实现位置 | 技巧/坑 |
|---|---|---|---|
| `para` + `expect` | 段号 + 段首文字双重锚点；`expect=""` 表示空段 | `:2052-2063`（`Squash(MarkNotes(...))` 比较，等价忽略空白差异） | 比对前先 `MarkNotes` 把脚注引用换成 `[footnote:i]`，否则脚注段永远对不上 |
| 整批一次撤销 | `app.UndoRecord.StartCustomRecord("AI edit")` ... `EndCustomRecord()`，**只对 Word 开** | `:9654-9656`、`:9716`；限制声明 `README.md:140` | Excel/PPT 明确放弃（Excel 外部修改会清空撤销栈）。我方三端都做了（`WpsHelper.cs:1233`、`Tables.cs:69`、`Layout.cs:23`、`Toc.cs:30`、`Text.cs:105/195`），并额外处理"WPS 不暴露 UndoRecord"（`WpsHelper.cs:826-828`） |
| 不拆词 | `range.ParagraphFormat.WordWrap = -1` | `:2573-2576`，调用点 `:2664`（题注）、`:3167`（插入段）、`:3292`、`:3699`（参考文献） | **关键实测：`WordWrap = -1` 才是整词换行，`= 0` 会拆**（`docs/verification.md:146`，在 Word 里量过 `37000` 的首尾字符同行）。这条与 MS 文档字面相反 |
| 标题自带编号 vs 样式自动编号 | 文本已匹配 `Numbered` 正则 → `ListFormat.RemoveNumbers()` | `:2579-2584`（单段，插入时用）；`:3748-3758`（`style_format {numbering:false}`：`style.LinkToListTemplate(null)` + **全文档扫描**该样式的段逐段 `RemoveNumbers`） | 我方 `style -NoNumbering` 已有逐段那半（`Layout.cs:247-284`），缺的是**插入时按文本形态自动去重** |
| 纯文本公式提醒 | 文本里出现 `d_k` / `10^18` 这类形态 → 结果里点名"请写成 `$d_k$`" | `:2560-2568`（`Lint`），调用点 `:3316`、`:3350`、`:3090`、`:3160` | 正则先剔掉 `$...$` 区间再匹配（`:2563`），否则公式内会误报 |

### 1.3 原生脚注（1.2.0 主打；README.md:104-118）

| op/字段 | 语义 | 实现位置 | 技巧/坑 |
|---|---|---|---|
| `insert_footnote {para, expect, text, find?, occurrence?, where?}` | 在段末/段内第 n 次命中前后插真正的页底脚注 | `:2206-2236` | `find` 存在时用 `Find.Execute` 循环跳 `occurrence` 次（`:2219-2223`）；`^` 必须转义成 `^^`（`:2221`） |
| `edit_footnote` / `delete_footnote` / `format_text {footnote, expect,...}` | 以**集合序号**（不是显示编号）定位，`expect` = 脚注正文开头 | `:2132-2142`（`FindNote`：`expect` 必填，不符报 `ANCHOR_MISMATCH`）、`:2237-2245` | 序号越界报 `ANCHOR_MISSING` 并提示"重新读"（`:2136`） |
| `footnote_settings {numbering: continuous\|section\|page, section?}` | 全篇或单节编号规则 | `:2196-2205` | `doc.Footnotes.NumberingRule = 0/1/2`；单节走 `doc.Sections[n].Range.FootnoteOptions`。**Lead 已实测 WPS 无 `FootnoteOptions.NumberingRule` → 此 op 在 WPS 上做不了**，只能"保留文档原设置 + 只读报告" |
| 正文内联语法 `[fn:…]` / `\footnote{…}` | 插入段/改写段/替换文本里内联写脚注 | `:2082-2108`（`ParseNotes`）、`:2160-2189`（`InlineNotes`） | ①**先解析后碰 COM**（`:2101` 未闭合直接拒绝，`:2097` 禁止嵌套）；②括号/花括号可嵌套（LaTeX、书目），`Escaped()`（`:2110-2115`）认转义；③**倒序插入**（`:2165`）保证前面的字符位置不漂；④解析完再把转义符还原（`:2174-2187`） |
| 脚注体读写 | `NoteBody()` 去掉首 `\u0002`（native reference mark）与尾 `\r` | `:2122-2130` | 注释直说 "including on WPS"——说明这两个哨兵在 WPS 文字里同样存在 |
| 读脚注 | `office_read {footnotes:true, full:true}` → 每条 `{i, label, numberStyle, para, offset, page, section, text}`；正文里显示 `[footnote:i]` | `:2296-2329`（`ReadNotes`）、`:2248-2262`（`MarkNotes`） | **`label` 是自己算的**：先 `Repaginate()`，再按 `NumberingRule`（0 全篇 / 1 每节 / 2 每页）分桶累加 `StartingNumber`，最后 `NumberLabel(number, style)`（`:2264-2294`）转罗马/字母/带圈/全角。**算不出的样式返回 null，绝不拿集合序号冒充显示号**（`:2293`）——这条设计值得直接学 |
| 编号样式判据 | `reference.Text` 含 `\u0002` → 自动编号；否则原始文本即标签 | `:2318-2321` | 手工编号的脚注（用户自己敲的 `*`）也能读对 |
| 段落含脚注时禁止 `set_text` | 显式拒绝并提示改用小范围 `replace_text` | `:3188-3193` | 理由写得很清楚：`set_text` 会丢脚注。**我方 revise 的 `set_text` 是"整段删+插"，等于静默删脚注**，见 §2/§3 |
| `replace_text` 跨脚注引用标记 → 拒绝 | 删/换的区间若跨过脚注引用标记，报 `BAD_ARGS` | `:3062-3070`（先全文档预检一遍再动手）、`:3075` | "**先预检再改**"与我方 revise 两轮思路同源（`DESIGN-word.md:73`） |
| `delete_range` 删除含脚注的区间 | 连脚注一起删，并报数量 | `README.md:118`、`src/tools.ts:119` | — |
| 脚注内可写 `**粗体**` 与 `$公式$` | `FillNote` 里 `Bold()` + `Mathify()` | `:2150-2158` | 脚注与公式两个特性**互相嵌套**，验收脚本 `scripts/footnotes-smoke.mjs:118` 断言重开文件后脚注里没有 `$` |

### 1.4 LaTeX → 原生公式（OMML）：两条完全不同的路

**A. 微软 Word 路（Word 自己读 LaTeX）**

| 步 | 做法 | 实现位置 |
|---|---|---|
| 剥 `\tag{n}` | `Tag` 正则取出编号，源串移除 | `:1538`、`:1972-1973` |
| 预处理 | `\mathrm{…}` 里的字母改写成 `\text{…}`（单位要正体，否则 Word 的线性式会排成斜体、`m/s` 变分式）；`\displaystyle`/`\limits`/`\dfrac`/`\lVert` 等替换；`''` → `″` | `:1794`、`:1800-1830`、`:1833-1841` |
| 能不能交给 Word | 只有 matrix/pmatrix/bmatrix/vmatrix/Vmatrix/cases 六种环境才算"Word 会读"；`aligned/array` 走自带转换器 | `:1843-1854` |
| **借一个不可见的 Word 当解析器** | 自己 `Activator.CreateInstance("Word.Application")` 开一个隐藏实例，把 `CommandBars` 切到 `EquationLaTexFormat`，逐条 `OMaths.Add(line).BuildUp()`，再 `ExecuteMso("EquationUnicodeFormat")` + `Linearize()` 取 Word 的线性式，**缓存源串→线性式** | `:5694-5766`（`MathSource`：写 pid 文件，崩溃后下次启动 `Sweep()` 清残留，`:5714-5735`）、`:1868-1932`（`ToLinear`） |
| 写进用户文档 | `spot.Text = linear(+"#(n)")` → `spot.OMaths.Add(spot).BuildUp()` | `:1986-2003` |
| 失败判据 | 建完文本里仍含 `\`、`&`、`{` → `Unbuilt` = 失败 → 计入 `MathFailed` 并在结果点名 | `:1857-1864`、`:1780-1788` |
| 用户把公式输入切成 LaTeX 时 | 批开始时切回线性、批结束还原（`UserMode` 状态机） | `:1934-1949`、`:2016-2024` |
| 占位 Padding | 建公式时 Word 会把"函数名后的字符"当参数吃掉，先塞 `Pad` 字符再删掉 | `:1951-1966`、`:1988-1998` |

**B. WPS 路（WPS 不建公式）** —— 这是本研究对**我方**最有价值的一段

| 步 | 做法 | 实现位置 / 证据 |
|---|---|---|
| 前提事实 | **WPS 的 `OMaths.Add` 不按给定位置插入、`BuildUp` 不生效** | `docs/verification.md:695`（0.9.6 实测结论）；代码分支 `:1548-1558` |
| 在**插件侧（TS）**预先把整批 ops 里所有 `$...$` 转成 OMML | `temml.__renderToMathMLTree` → MathML 树 → 手写 OMML 序列化；key = `"d:"+源`（独立成行）/`"i:"+源`（行内） | `src/omml.ts:1-7`、`:8`、`:18`、`:96-153`（节点映射）、`:156-174`（`\sum` 类 n-ary 吃后面一个操作数）、`:177-185`、`:187-218`（`formulasOf` 深度遍历 ops，遇 `latex` 字段也收） |
| 随调用把字典交给 helper | `office_edit {…, math: formulasOf(ops)}` | `src/tools.ts` 调用侧；验收脚本同样写法 `scripts/footnotes-smoke.mjs:61` |
| helper 把文本换成公式 | 找到 `$...$` 原文 → `spot.InsertXML(FlatDocument(omml))`；Flat OPC 是**手工拼的** `pkg:package` + `/_rels/.rels` + `/word/document.xml`（`w`+`m` 命名空间） | `:1672-1709`、`:1663-1669` |
| 行间公式 | 包 `<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>…` | `:1709` |
| `\tag{n}` 编号靠右 | 不靠"制表位"，而是把 OMML 改写成 Word 自己导出"公式#(1)"的那个结构：`m:eqArr` + `m:maxDist=1` + 一段 `<m:t>#</m:t>` + 正体 `(n)` | `:1704-1708`；WPS 实测 `docs/verification.md:888` |
| 插入后收尾 | ①`InsertXML` 会多带段落标记 → 删掉把行合回去（`:1712-1721`）；②插入前记下段落样式与对齐/缩进/段距/行距，插入后还原（否则列表项变正文，`:1691-1699`、`:1722-1736`）；③行间公式居中 + 上下 ≥6pt + 首行缩进清零（`:1737-1741`） | WPS 实测 `docs/verification.md:815` |
| 编号贴太近时逐步缩字号 | 量"公式末尾"到"编号"的水平距离（`Information[5]`），不足 14pt 或折行就每次 −0.5pt，最多减到 70% | `:1744-1767`；`docs/verification.md:918` |
| 失败兜底 | 逐条 `MathSkipped++`，批末统一提示"这些公式没变成公式、仍是 `$…$` 文本，请改写更简单的 LaTeX 或告诉用户" | `:1550-1557`、`:1772-1778`、`src/tools.ts` GUIDE 侧同义说明 |
| 验收证据 | 一份含 11 个公式（分式/求和/三次根/矩阵/cases/aligned/范数/上划线/花括号/连乘/表格内公式）的文档在 WPS 里**全部成为原生公式**，段落数不变 | `docs/verification.md:742` |

**B2. 逐 API 复刻：它在 WPS 文字里到底做了什么（Lead 实测冲突点的唯一解释入口）**

> 结论先说：**WPS 文字里没有剪贴板路线、没有 EMF 路线、没有 EQ 域路线**（EMF/剪贴板是 dsh-office 给 **WPS 演示**用的：`PptMath` 走剪贴板 `:6171`、WPS 演示失败后取 EMF 图片，`verification.md:760`）。WPS **文字**只有一条路：**helper 把现成 OMML 字符串包成 Flat OPC 文档，`Range.InsertXML` 一次性插入**。所以 Lead 的 `InsertXML` 实验"没转成公式"，差异只可能在 ①插的 XML 形态（Flat 包 vs 裸片段）②OMML 片段本身 ③版本 ④判真方式。

按 `helper/OfficeHelper.cs` 的执行顺序（`WpsMath`，`:1672-1770`）：

| # | 代码动作 | 行号 | 我方复刻时必须注意 |
|---|---|---|---|
| 1 | 正则找 `$…$` / `$$…$$`（`Dollars`） | `:1536`、`:1546` | 从**后往前**替换（`:1552`） |
| 2 | 从插件传来的 OMML 字典里按键 `"d:"+源`（行间）/`"i:"+源`（行内）取串 | `:1680` | **转换发生在插件侧，不在 COM 里**——WPS 侧零解析 |
| 3 | 定位：短文本（≤250 字符）先 `Find.Execute` 找原文；否则按字符区间 `InStory` | `:1682-1687` | 字符偏移在有域的段落里不可靠 → 一律先 Find |
| 4 | **记下该段的样式名 + 对齐/缩进/段距/行距/首行缩进** | `:1691-1699` | 插入会被套上"公式段落"的样子，事后必须还原 |
| 5 | 行间+`\tag{n}`：把 `<m:oMath>…</m:oMath>` 改写成 `m:eqArr`+`m:maxDist=1`+`#`+正体 `(n)` | `:1704-1708` | 见 §5.2 |
| 6 | **`spot.InsertXML(FlatDocument(...))`**：行间包 `<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>…</m:oMathPara>`；行内直接 `<m:oMath>…</m:oMath>` | `:1709` | ★★**`FlatDocument` 是 `pkg:package` 整包**（§5.3），不是裸 XML 片段。**Lead 的实验最可能就是差这一步**：`Range.InsertXML` 在 Word 里两种都收，但 WPS 很可能只认 Flat 包 |
| 7 | **验真：`if ((int)range.OMaths.Count <= maths) return false;`** | `:1710` | 作者在 28505 上靠"计数增长"判成功；Lead 在 29161 上计数不增长 → **要么 XML 形态不对（回到了 6），要么 WPS 29161 换了行为**。这一步的作者写法本身不够硬，我方应改成三判据（§4.3） |
| 8 | `InsertXML` 会多带一个段落标记：删掉它把行合回去 | `:1712-1721` | 不删 → 每个公式多一个空段（`verification.md:742` 特别断言"段落数不变"） |
| 9 | 还原步骤 4 记的样式与段落格式 | `:1722-1736` | 不做 → 列表项变正文（`verification.md:815` 的实锤） |
| 10 | 行间公式：居中 + 段前/段后 ≥6pt + 首行缩进清零 | `:1737-1741` | — |
| 11 | 编号挤压：量 `Information[5]`（水平位置）差 <14pt 或折行 → 字号 −0.5pt，最多到 70% | `:1744-1767` | `Information[5]` 我方未用过，需探 |
| 12 | 失败 → `MathSkipped++`，批末统一点名"仍是 `$…$` 文本" | `:1554-1557`、`:1772-1778` | 失败必须可见 |

**三条可直接执行的验真判据（WPS 上比 `OMaths.Count` 可靠，且我方已有 `WordOpenXML` 读取能力，见 `SKILL.md:186` 的 `textdiag`）**
```
① 真·公式：Content.WordOpenXML 里出现 <m:oMath 或 <m:oMathPara   → 是原生公式
② 被当图片：WordOpenXML 里出现 <w:drawing> 且 InlineShapes.Count 增加 → WPS 把 OMML 渲染成图了（不可编辑）
③ 被当域/文本：Fields.Count 增加 或 Range.Text 仍是 "<m:oMath…"（字面 XML） → 插入成了纯文本
```
→ **`OMaths.Count` 在 WPS 上可能滞后**（需 `doc.Repaginate()` 或重新打开文档后再读），这与"集合计数不刷新"是两件事，必须用 ①②③ 分开定性。

**C. 定长行距裁切公式/图片（`Roomy`）** —— 最容易漏、修起来最便宜的一条

| 项 | 语义 | 实现位置 | 技巧 |
|---|---|---|---|
| `Roomy(doc)` | 扫 `doc.OMaths` + `doc.InlineShapes`，把所在段落的**固定行距（Rule=4）改成最小值（Rule=3）且行距数值不变** | `:1615-1649` | 只在"本批建过公式、或设过固定行距"（`Cramped` 标志）时跑，`:9705-9709`；按段首字符位置去重（`:1634-1636`）；结果里点名改了 N 段 |
| 行间公式留白 | `SetOff()`：段前/段后 ≥6pt、清首行缩进 | `:1592-1613` | 只用段落间距、不插空段（`:2630-2632` 同一哲学） |

### 1.5 题注、三线表、跨页、图片

| op/字段 | 语义 | 实现位置 | 技巧/坑 |
|---|---|---|---|
| `insert_table {caption, borders, header}` | 建表时同给题注与框线 | `:3249-3339` | ①表格段落先 `Plain(p,"Normal")` 再 `Font.Reset()`，否则继承下一个标题的字体（`:3273-3284`）；②`borders=three-line`：上下 `LineWidth=12`（=1.5pt）、表头下 `6`（=0.75pt）、其余全关（`:3460-3480`）；③`Rows[1].HeadingFormat=-1` 表头跨页重复、`Rows.AllowBreakAcrossPages=0` 行不拆页（`:3321-3327`）；④题注/无题注都给 8pt `Air` 段距（`:3328-3335`） |
| 题注（三条硬规矩） | 自动编号（**Word 题注域**）、居中、**表上图下** | `:2646-2676` | ①`Application.CaptionLabels` 先补一个中文标签 `表`/`图`（`:2651-2654`）；②`target.InsertCaption(Label, Title, Position: 表?0:1)`（`:2656`）——用**内置题注域**而不是自己拼 SEQ；③题注段 `KeepWithNext=-1`（表题不落单，`:2670`）；④题注里的公式也转（`:2673`） |
| `format_table {keepTogether}` | 表格被切到两页且"只切出单独一行"时整体挪一页 | `:2889-2906`、`:2851-2885` | 判据 `Stranded`：`OnOnePage` 否 → 首页只站 ≤2 行且还有 ≥1 行在下一页 = 1；下一页只剩 ≤1 行 = 2（`:2837-2846`）；修法①正文里的表：**除末行外每行 `KeepWithNext=-1`** 让 Word 自己挪（`:2898`）；②修法②推空行下推（`:2876-2884`），**下推会让上一页空出 >1/3 就放弃并明说**（`:2873`、`:2902-2904`） |
| `fill_gaps {}` | 图表整块移到下一页、页脚空出一截 → 把**图表后面的正文段**提到图表前面（像排版软件那样填缝） | `:2967-3016`；说明 `README.md:41`、`src/tools.ts:129` | 见 §5 摘录 4。四个判据缺一不可：①图/表块（图=前空段+图+后题注，表=前题注+表，`:2732-2758`）②块前的页与块的首页不同 ③页脚空隙 ≥ `max(90pt, 页高×12%)`（`:2994`）④**块前是正文（`OutlineLevel==10`）才搬，标题绝不越过**（`:2987`、`:3000`）；另：相邻图表先 `Absorb` 成一块（`:2760-2785`）再搬；块后那段若以 `:`/`：` 结尾，会连带把下一段一起搬（`:3010`） |
| `fill_gaps` 的第二半 | 顺带修"跨页表格切出单独一行" | `:3394-3398`（先 `FillGaps` 再 `KeepWhole`，因为填缝会改变表格落点） | `KeepWhole`（`:2914-2965`）：正文表走 KeepWithNext；**在别人单元格里的表（报告模板外框）Word 不认 KeepWithNext、单元格里也放不了分页符**，只能连题注+上面小标题一起用空行推下去（`:2911-2913`、`:2949-2951`）；>12 行的嵌套表只点名不动（`:2943-2948`）；没修好的逐条点名并给建议（`:2962-2963`） |
| `insert_diagram {nodes, edges}` | 用 Word 自己的形状在文档里画流程图/结构图（可逐个编辑） | `:3857-4070` | ①画在**行内绘图画布** `Shapes.AddCanvas(...Anchor: ref range)`（`:3941`）；②节点按 `row/col` 落格，`cellWidth/cellHeight` 自适应（`:3866-3872`、`:3906`）；③菱形按文字长度加宽（`:3896-3912`）；④**连线不是连接符，是自己算的直角折线**：`Route()` 按"上下/左右 + 同侧/异侧"分三种走法（`:3820-3855`），接哪边由两节点相对位置推（`:3925-3929`，可 `fromSide/toSide` 覆盖）——**Lead 已实测 WPS 的 `AddConnector`+`BeginConnect` E_FAIL，正好证明 dsh-office 这条"自己画线"的选择不是风格而是唯一出路**；⑤按页宽整体缩放并把比例写进结果（`:3939`、`:4069`）；⑥WPS 专项：线段必须**从左上方画起**（否则 WPS 会翻转、箭头跑到另一头，`:3991-3997`）、行内画布在 WPS 里不占字符 → 插一个零宽 `\u200B` 占位（`:4056-4058`）、画布 in-line 后若锚点跑偏要搬回（`:4041-4055`） |
| `set_image` / 图片段 | 改尺寸；**插图段落不被固定行距裁切** | `:3352-3372`（插入）、`:3432-3453`（改尺寸）、`Cramped` → `Roomy` | 插图后 `p.Range.ParagraphFormat` 居中、清首行缩进、8pt 段距；缩小过多的提示在 `Lettering`（`:7159-7174`，图里文字实际多少磅） |
| `style_format` | 改样式本身（字体/字号/行距/首行缩进/是否自动编号/另起一页） | `:3730-3761` | `numbering:false` 时**同时**做"摘掉样式的编号模板"与"逐段 RemoveNumbers"（`:3748-3758`），两道都做才在已有文档上干净 |
| `page_setup / page_numbers / header / insert_toc / update_fields / page_break` | 页面与域 | `:3763-3774`、`:4072-4079`、`:4081-4096`、`:4099-4103`、`:3380-3385` | `update_fields` = `TablesOfContents.Update()` + `Fields.Update()`（`:4099-4103`）；目录"保存时自动刷新"由宿主保存流程保证（`src/tools.ts:131`） |

### 1.6 引用与参考文献

| op/字段 | 语义 | 实现位置 | 技巧/坑 |
|---|---|---|---|
| `insert_references {items}` | 生成 `[n] 文献` 段落，**并在编号数字上建 `cite_n` 书签** | `:3681-3715` | 悬挂缩进：`LeftIndent=21pt` + `FirstLineIndent=-21pt`（`:3704`）；`items` 里自带 `[k]` 前缀时以它为准（`:3685-3693`） |
| `\cite{1}` / `{2,5}` / `{3-6}` | 变成上标 `[1]`，**是 `REF cite_n \h \* MERGEFORMAT` 交叉引用域**（Ctrl+点击跳转） | `:3497-3514`（`CiteNumber`） | 目标书签**还不存在**时先降级成 `Hyperlink{SubAddress: cite_n}`，等参考文献列表生成后 `Relink` 升级成域（`:3507-3513`、`:3606`）——"先引用后建列表"的写入顺序因此无所谓 |
| 手写 `[1]` 自动转域 | 正文里的 `[2]`、`[1,3]` 改写为 `\cite{...}`；段首的 `[n] …` 认成参考文献条目并自动打书签 | `:3516-3559`（`Handwritten`） | 用通配 Find `\[[0-9,，–\-]@\]` 一次扫全篇（`:3529`）；**公式里 `$[1,2]$` 不当引用**（数 `$` 奇偶，`:3549-3553`） |
| 无人引用提醒 | 写了参考文献列表却在正文一个 `REF cite_` 都没有 → 批末点名 | `:3565-3582`、`:9710-9715` | 把"生成的列表必须被引用"变成机器可查的约束 |

---

## §2 我方现状（已有 / 半有 / 没有）

> 判定依据：`SKILL.md` 命令表 167–207 行 + `reference/word-quirks.md:36`（边界与未验证）+ `DESIGN-word.md:115`（**作者自己列的"P4 未做"**）+ helper 源码实读。

| # | 能力 | 判定 | 依据 |
|---|---|---|---|
| 1 | 原生脚注（增/改/删/读） | **没有（WPS 能做，Lead 已实测 `Footnotes.Add` 可行）** | 命令表 `SKILL.md:167-207` 无 footnotes；helper 全目录 grep `Footnotes` **零命中**；`DESIGN-word.md:115` 未列此项 |
| 2 | 脚注编号规则切换（`footnote_settings`） | **没有（WPS 做不了）** | Lead 实测 WPS 无 `FootnoteOptions.NumberingRule` → 即使实现也只能"保留原设置+只读报告" |
| 2b | 内联 `[fn:…]` / `\footnote{…}` 语法 | **没有** | 同上（纯 helper 侧解析，与 WPS 能力无关，可直接照 `:2082-2108` 实现） |
| 3 | 含脚注段落改写护栏 | **没有**（且有数据丢失风险） | `revise` 的 `set_text` 是"整段删+插"（`SKILL.md:222`），helper 无 `Range.Footnotes.Count` 检查。用户原稿若带脚注，`set_text`/`delete` 会**静默删掉脚注**且不报告 |
| 4 | 读文档时标出脚注引用 | **没有** | `read` 返回文本直接取 `Range.Text`，脚注引用标记（`\u0002`）会被 `CleanPara` 当控制字符处理掉，模型看不出"这段有脚注" |
| 5 | LaTeX → 原生公式 | **没有（且被正经声明为范围外）** | `DESIGN-word.md:115`「**P4（未做，按需再议）：公式 LaTeX→OMML**、图示（形状+连线）、fill_gaps 重排」；`reference/word-quirks.md:41`「公式（OMML）：不在本 skill 范围，未验证」；helper 无 `OMaths/InsertXML` |
| 6 | 公式编号（`\tag`）+ 编号靠右 + 缩字号适配 | **没有** | 同上 |
| 7 | 公式/图片段落固定行距裁切防护（`Roomy`） | **没有** | helper 只在 `SetLineSpacing` 里**设置**行距（`WpsHelper.cs:542-561`），无"含公式/图片段把 Rule 4→3"的收尾扫描；我方 `picture` 插入（`Tables.cs:1303`）只设 6pt 段距 |
| 8 | `\cite` → REF 域 + 书签 | **半有** | 列表+上标**纯文本**已实现（`Toc.cs:431-496`），且**明确拒绝**建 REF 域并说明理由（`Toc.cs:494`：无书签时 REF 会显示 "Error! Reference source not found."） |
| 9 | 参考文献列表（悬挂缩进） | **已有** | `Toc.cs:289-300`（`reference -Action list`）；缩进值需实测比对（我方用 21pt 悬挂？见落地建议） |
| 10 | 题注域（自动编号、表上图下、不落单） | **已有（更细）** | `Tables.cs:1080-1149`：SEQ 域 `表 \* ARABIC`（`:1114-1119`）、**先写题名再在题名前插域**（`:1105-1111`，注释记录了"往域结果尾部写字会被 F9 冲掉"的实测）、`KeepWithNext`（`:1140`）、题注位置默认表上图下（`:1002`） |
| 11 | 题注编号体检（重号/跳号/格式混用/悬空引用） | **已有（dsh-office 没有）** | `SKILL.md:187`、`Captions.cs`（`CaptionRx` 单点解析器、`captionParas` 去重、按节跳号、`references.resolved`） |
| 12 | 三线表 | **已有（口径一致）** | `Tables.cs:686-696`：上下 1.5pt、表头下 0.75pt、其余全关；另加 `grid/thick/none`（`:659-685`）——`thick`（1.5pt 外框+0.5pt 内线）是 dsh-office 没有的中文报告口径 |
| 13 | 表头跨页重复 | **已有（更硬）** | `Tables.cs:931-957`：绕开 WPS 残缺的 `Rows.Item(n).HeadingFormat`，走 `Selection.Rows.HeadingFormat` **并回读校验**（`:941-944`）；dsh-office 直接 `Rows[1].HeadingFormat` 且无回读 |
| 14 | 行不拆页 / 表整体同页 | **半有（且有静默失效疑点）** | `table -KeepTogether`（`SKILL.md:184`）→ `Tables.cs:959-980`：`AllowBreakAcrossPages=0` 用 `Rows.Item(r)`（`:968-969`）、`KeepWithNext` 用 `table.Rows.Item(r).Range.ParagraphFormat`（`:975`）。而本文件自己写明 **WPS 的 `Rows.Item(n).Range` 会抛 0x800A1767**（`Tables2.cs:5`）→ `:975` 大概率抛异常被 `catch` 吞掉，**"同页"实际没落上却报成功**。需实测确认 |
| 15 | 跨页表单行落单的**自动判定+修复** | **没有** | 无 `Stranded/PushDown` 等价物；`table -Action header-repeat` 只重复表头，`Table2.cs:456-457` 只是用 `Information(3)` 判断跨页（能力已具备，逻辑未写） |
| 16 | `fill_gaps` 消空白 | **没有** | `DESIGN-word.md:115`；`SKILL.md:250-277` 工作流只写"`render` 看空段半空页"。`save` 前也无自动跑一遍（`word.ps1:598-601`） |
| 17 | 保存前自动跑 fill_gaps / 强制视觉核对 | **没有** | dsh-office 在 `office_save` 里自动跑（`docs/verification.md:952`）；我方 `save` 只保存 |
| 18 | `insert_diagram`（形状流程图/结构图） | **没有** | `DESIGN-word.md:115`；helper 无 `AddCanvas/CanvasItems` |
| 19 | `insert` 里直接放图片/表格（items 混排） | **半有** | `Text.cs:162-173`：items **只认 `text`**（+style/格式）；图片要另起 `picture` 命令、表格要另起 `table`（两者都按段号插入，混排时段号漂移需自己算） |
| 20 | `expect` 锚点 + 整批原子 + 回滚 | **已有（更强）** | `SKILL.md:222-226`、`DESIGN-word.md:73`（两轮：全量预检→降序执行→中途出错逐条 Reject 回滚）；dsh-office 是"做到哪报到哪"（`:9659-9683`） |
| 21 | 整批一次撤销（UndoRecord） | **已有（三端都有）** | `WpsHelper.cs:1233`、`Tables.cs:69`、`Layout.cs:23`、`Toc.cs:30`、`Text.cs:105/195`、`Review.cs`；并有"WPS 不暴露 UndoRecord"的降级提示（`WpsHelper.cs:826-828`）；局限见 `SKILL.md:293`（save 清栈、同值格式不进撤销栈） |
| 22 | 未知字段/参数点名 | **已有（等价，通道不同）** | `scripts/param-check.ps1:22-52`：v1.18 起"绑定了但没进 request 的参数"打 `WARN ignored parameter -X for command Y`（stderr）；dsh-office 写进结果正文（`:9669-9670`） |
| 23 | 不拆词（整词换行） | **半有（极性疑似反了，且不默认）** | 开关存在：`word.ps1:53`（`-NoSplitWords`）→ `WpsHelper.cs:632` 写 `pf.WordWrap = v ? 0 : -1`。**dsh-office 实测 `-1` 才是整词、`0` 会拆**（`docs/verification.md:146`）→ 我方语义疑似**相反**；且只在 `format` 传该开关时生效，`insert`/`table`/`caption` 新写入的内容不设 |
| 24 | 标题自带编号自动去重 | **半有** | `style -NoNumbering` 改样式 + 逐段 `RemoveNumbers`（`Layout.cs:247-284`）在；缺"插入时按文本形态 `一、`/`2.1` 自动去掉自动编号"（dsh-office `:2579-2584`） |
| 25 | 纯文本公式提醒 | **没有** | 无 `Lint` 等价物；`audit` 的规则集（`SKILL.md:194`）不含此项 |
| 26 | 表格内锚点防护 | **半有** | 有"锚点段后紧跟表格 → 改在段末标记前插 CR"（`Text.cs:208-224`），无 dsh-office 的"给的表内段号会插进单元格 → `STALE` 拒绝"（`:3255-3261`） |
| 27 | 样式/版式/目录/页眉页脚/页码/域刷新 | **已有（部分更细）** | `SKILL.md:181-195`；`page` 有逐节查询 + `firstDisplayedPage` 断号检查（`Layout.cs:368-374`） |
| 28 | 渲染核对 | **已有（同源技法）** | 我方 `render` 也是 `Pages.EnhMetaFileBits` → `Metafile`（`WpsHelper.cs:1576-1601`），并同样采用"**先大图再缩小**"以免表格细线落在像素缝里（`:1596-1601`；dsh-office 同一坑 `:9726-9737`） |
| 29 | 批注 / 修订 / 快照 / diff / 交接单 / 单位上标 / 题注体检 | **已有（dsh-office 没有）** | `SKILL.md:177-207` |

---

## §3 差距与价值（面向"中文水利工程报告 + WPS 实时改稿"）

### P0 — 正确性缺陷，成本极低，**建议立刻做**

| 项 | 为什么是 P0 | 理由/代价 |
|---|---|---|
| **P0-1 含脚注段落的改写护栏** | 现在 `revise set_text`（整段删+插）会把用户原稿里的脚注一起删掉，**不报错、不提示**，属于静默数据丢失。水利报告正文少脚注，但**评审意见、引用的规范条文、图注出处**常用脚注；且我方是"在用户正开着的文档里改"，用户不会自己发现 | helper 里 3 行：`if ((int)p.Range.Footnotes.Count > 0) throw new Fail("BAD_ARGS", …)`（照 `:3193` 抄），再加 `read` 里把脚注引用标成 `[footnote:n]`（照 `:2248-2262`）。半天 |
| **P0-2 不拆词的极性与默认** | 中文报告正文常夹 `37000 m³`、`SL/T 4—2020`、英文单位；模板样式带"允许西文在单词中间换行"时会被拆成 `37 / 000`。我方开关语义疑似**反的**，且默认关 | 先花 10 分钟实测极性（§4.1 脚本），然后把 `insert`/`table`/`caption`/`reference` 新写入的内容**一律**设成整词（dsh-office `:2573`、`:2664`、`:3167`、`:3299`、`:3699`）。半天 |
| **P0-3 固定行距裁切公式/图片** | 水利报告正文标配"固定值 20 磅/28 磅"；一旦在正文段插图片（现有 `picture` 命令已经允许插在正文段后）或将来插公式，**图片/公式会被行高裁掉一半**，渲染核对若没看到那一页就流到成品 | `Roomy` 等价物 20 行（`OMaths` + `InlineShapes` → `LineSpacingRule 4→3`）。半天（含验证） |
| **P0-4 `table -KeepTogether` 的静默失效** | `Tables.cs:975` 用了自己文档里声明"WPS 会抛异常"的 `Rows.Item(r).Range`，异常被吞 → 命令报 keepTogether 成功、实际没落上。**报告排版最常被要求的就是"表格别断开"** | 改用 `Selection.Rows`（`Tables2.cs:187` 已有先例）+ **回读校验**；`AllowBreakAcrossPages` 也一并回读。半天 |

### P1 — 真实能力缺口，报告高频，**值得排期**

| 项 | 价值 | 理由 |
|---|---|---|
| **P1-1 LaTeX → 原生公式（OMML）** | ★★★★★ | 水利/水力学报告公式密度极高（水面线、曼宁、频率计算、堰流），历来靠人工插公式或截图。dsh-office 已经**给出了一条在 WPS 上验过 11 个公式的路径**（预生成 OMML + `Range.InsertXML`，`docs/verification.md:742`），也给出了"WPS 自己建公式失败"的否定结论（`:695`）——后者需按 §4.3 步骤 0 在我方机器上复核。这是本研究里**可复用度最高**的一条：路已探明，剩下的是"LaTeX→OMML 子集转换器"的工程量 |
| **P1-2 公式编号 + 编号靠右 + 缩字号适配** | ★★★★☆ | 报告规范要求 `(1)` 右对齐、公式居中；dsh-office 用 `m:eqArr`+`#` 的结构化写法已解决（`:1704-1708`），并在 WPS 上实测（`verification.md:888`/`:918`）。不做这一步，"公式"只能算半成品 |
| **P1-3 `fill_gaps` 消空白 + 保存前自动跑** | ★★★★☆ | 报告图多、图大，图表整块挪页后页脚空半页极其常见，人肉调整最费时间。算法已完整可抄（§5 摘录 4），且**保存前自动跑一次**（dsh-office `verification.md:948-952`）能把"没人看图就交付"的风险兜住。技术依赖 `Information(6)`，我方尚未用过，需先探一次 |
| **P1-4 跨页表单行落单的自动修复** | ★★★★☆ | 与 P0-4 同族但是"体检+修复"级：`stranded` 判据（首页只站 ≤2 行 / 下一页只剩 ≤1 行）+ `KeepWithNext` 修 + 推空行兜底 + **修不动就点名**。报告规范禁止"续表单行"，人工找漏极难 |
| **P1-5 原生脚注** | ★★★★☆（Lead 实测后上调） | **Lead 已实测 `Footnotes.Add` 在 WPS 可行** → 从"受众窄、待验"变成"确定能做"。增/删/读三件事都能上；只有"编号规则切换"做不了（WPS 无 `FootnoteOptions.NumberingRule`）。它同时是 P0-1 的正解（有了读写能力才能改而不是删） |

### P2 — 收益/成本比一般，**先记录不排期**

| 项 | 判断 |
|---|---|
| `\cite` → REF 域 + 书签 | 我方纯上标文本对**打印稿**够用；REF 域的唯一增量是"Ctrl+点击跳转"（WPS 里还要实测 `REF \h` 是否可跳）。若做，成本约半天（书签 + 域 + `Relink` 两步走） |
| `insert_diagram` 形状流程图 | **Lead 已实测：`Shapes.AddShape` 可行、`Shapes.AddConnector`+`BeginConnect` E_FAIL** → 与 dsh-office 的做法一致（它本来就自己画线段、不用连接符）。**待测一条：`Shapes.AddCanvas` + `CanvasItems.AddShape/AddLine`（这才是它真正用的那一层，也是"图形随正文走"的关键）**。通路明确后成本 ≈ 1 天 |
| `insert` 的 items 混排图片/表格 | 现流程（先 `picture` 再 `table`，各自按段号插）能用，只是需要 agent 重算段号；收益是省几轮 |
| 纯文本公式提醒（`Lint`） | 有了 P1-1 之后才有意义（否则"提醒"了也做不了），20 行 |
| 表格内锚点 `STALE` 拒绝 | 一次误插会把表塞进单元格且位置全乱；但我方 `insert` 已有"表格前陷阱"防护，剩的是"目标段在表内却要插表"的显式拦截 |
| 标题自带编号自动去重（插入时） | `style -NoNumbering` 已能一键根治整篇，逐段判定的增量小 |

---

## §4 落地建议（PowerShell/COM 怎么接、WPS 支持性实测线索、风险、工作量）

### 4.0 共用的前置结论

1. 我方 helper 已经是"WPS 优先、微软 Word 兜底"（`WpsHelper.cs:290-293`、`:684`），**但本机 `Word.Application` 被 WPS 接管**（`reference/excel-quirks.md:3` 同类事实）→ dsh-office 那条"借隐藏 Word 当 LaTeX 解析器"的路在我方**不可用**，P1-1 必须自带转换器。
2. 我方必须坚持"**一次块读、本地算、一次块写**"：公式走 `InsertXML` 是一次调用（天然满足）；`fill_gaps` 的循环（每挪一段重算块边界）会放大往返，需按 dsh-office 的上限（`step<8`、`guard<6`）**加硬上限**，并在结果里报告"这一轮动了几段、还有几处没填"。
3. 所有新能力都要**失败可见**（照 `MathSkipped` 模式）：宁可报"3 个公式没转成"，也不要静默留 `$…$` 文本。

### 4.1 P0-2 先做：不拆词极性实测（10 分钟，不改代码）

```
新临时文档 → 造一个带"允许西文在单词中间换行"的段样式（或直接把该段 WordWrap 分别设 0 / -1）
→ 段落文字写 "SL/T 4—2020 与 37000 m³s-1 的取值…" 并让列宽刚好放不下 → render 该页看是否拆成 37/000
→ 读回 pf.WordWrap 的真值（WPS 可能不保存该属性，需 TryGet）
```
判据：**若 `-1` 不拆、`0` 拆**，则 `WpsHelper.cs:632` 取反。若 WPS 表现与 Word 相反（理论上不应），则需在 helper 里按 `Suite` 分派。

### 4.2 P0-1 / P1-5 脚注

**COM 序列（照 dsh-office 抄，全部是标准对象模型）**

```
insert:  spot = 已 collapse 的插入点（段末 = p.Range.End-1 再 Collapse(0)）
         note = doc.Footnotes.Add(Range: spot, Text: " ")
         body = note.Range.Duplicate
         去首 \u0002、去尾 \r            → body.Text = 正文（可再走 Bold/Mathify）
         ※ 不能用 doc.Range 造脚注坐标：脚注是另一个 story（InStory 概念，`:2065-2072`）
edit/delete: 序号 = 集合序号；expect 必填（`:2132-2142`）
read:    doc.Repaginate() → 逐条 note.Reference.Information[2]=节 [3]=页；
         label 自己算（rule 0/1/2 分桶累计 StartingNumber + NumberStyle），算不出就返回 null
settings: doc.Footnotes.NumberingRule = 0/1/2；单节 → doc.Sections[n].Range.FootnoteOptions
```
**WPS 实测线索**
- **Lead 已实测（29161）**：`Footnotes.Add(range, '', text)` 可行，`Footnotes.Count` 与脚注文本可回读 → 增/删/读这条路**能做**。
- **Lead 已实测**：`FootnoteOptions.NumberingRule` / `NumberStyle` 一类在 WPS **不存在** → `footnote_settings` 一律降级：只**读并报告**文档现有设置，读不到就返回 `numberingRule: "unsupported"`，**绝不写入**；`read` 的 `label` 就只能取"WPS 自己给的显示编号"或退回集合序号（**并在结果里写明这是集合序号**，照 `:2293` 的"算不出就不冒充"原则）。
- 推论（**据代码推断，非实测**）：dsh-office 的 `footnotes-smoke.mjs --wps` 里有两处 `footnote_settings` 断言（`:109-112`、`:122-132`），而 `verification.md` 对 1.2.0 零记录 → 该 WPS 分支**很可能从未跑通**；而 `README.md:31` 的"支持连续/每节/每页重编号"是**Word 侧**事实，不能直接搬到 WPS。
- dsh-office 代码里另有三处 WPS 专项（说明"增删读"确实在 WPS 上跑过）：`:2233-2235`（新增后 WPS 返回"追加序号"而非真实序号 → 绝不能拿 `Add` 返回值当标识符，一律重读集合）、`:2251-2252`（WPS 的 `note.Index` 相对 `Range.Footnotes` 而非 `Document.Footnotes`）、`:2126`/`:2302`（脚注体首尾哨兵在 WPS 同样存在；拍摄版式 Reading Layout 下读写被拒，先关后还原）。
- 验收脚本可借其**结构**（`scripts/footnotes-smoke.mjs`）：造夹具走 `kwps.Application` 且**只在微软 Office 分支 `Quit`**（`:37`，绝不可 Quit 用户的 WPS）；断言清单：一条 undo 撤掉整批脚注（`:85-88`）、含未闭合脚注被拒（`:97`）、重开文件后脚注内无 `$`（`:118`）。**编号规则那两段断言要删掉**（WPS 做不到）。
**风险**：修订模式下 `Footnotes.Add` 是否留修订、拒绝修订是否回滚脚注（**需再测**）；脚注正文里的公式要走 P1-1，两个特性互相耦合。
**工作量**：helper 150–200 行 + `word.ps1` 4 个命令（`footnote` 增/改/删/读，去掉了 settings 的写入口）+ 文档 ≈ 0.5–1 天。

### 4.3 P1-1/P1-2 公式（本次最有价值，也最重）

**步骤 0（Lead 已完成，结论已定）：WPS 不会自己建公式**
- 已实测：`doc.OMaths` 对象在、`OMaths.Add(range)` 抛「值不在预期范围内」；`Fields.Add(EQ \o(2,3))` 的 `Result` 为空 → **A 路线（线性式 + BuildUp）与 EQ 域路线都排除**（与 `docs/verification.md:695` 方向一致）。
- 已实测：`Range.InsertXML(...)` **返回成功但 `OMaths.Count` 不变** → 只剩"InsertXML 到底插进去了什么"这一个问题。

**步骤 1（30 分钟，决定性）：用三判据验 `InsertXML` 的真伪，而不是看 `OMaths.Count`**
```
只改一个变量：把插入的串从「裸 OMML 片段」换成 dsh-office 的 Flat OPC 整包（§5.3 逐字节照抄，
  段落内容换成 <m:oMathPara>…<m:oMath><m:f><m:num>2</m:num><m:den>3</m:den></m:f></m:oMath>…</m:oMathPara>）
插完立刻取 Content.WordOpenXML，按 §1.4-B2 末尾的 ① ② ③ 判：
  ① 有 <m:oMath>  → 真公式（后续照 §4.3 步骤 2 起做）
  ② 只有 <w:drawing> → WPS 渲染成图（不可编辑，但"能看"仍有用，可当降级方案收下）
  ③ 是字面 XML 文本 / 只有域 → 不通，停
再补两次对照：doc.Repaginate() 后读 OMaths.Count；close+reopen 后再读（排除计数滞后）。
★ 这一步的结果决定 2–3 天的投入方向，必须先做，且只用临时新文档。
```
**若 ① 成立 → 走 B**：`LaTeX → OMML → range.InsertXML(FlatOpc)`

1. **Flat OPC 包装**（照 `:1663-1669` 原样拼，约 6 行字符串）：
   `pkg:package` → `/_rels/.rels`（指向 `word/document.xml`）→ `/word/document.xml`（`w`+`m` 命名空间）→ `<w:body><w:p>…</w:p></w:body>`。
2. **OMML 生成器**（我方需自研）：把 dsh-office 的 `src/omml.ts`（218 行）作为**规格书**，用 C# 重写一个**子集**转换器，覆盖水利报告实际用到的构造：
   分式 `\frac/\dfrac`、上下标 `^ _`、根 `\sqrt[n]`、求和/连乘/积分 `\sum \prod \int`（含上下限，走 `m:nary`）、括号 `\left( \right)` → `m:d`、矩阵/`cases`/`aligned` → `m:m`/`m:eqArr`、重音 `\bar \hat \vec \overline` → `m:bar`/`m:acc`、函数名 `\sin \lg \ln \max` → `m:func` 或正体 `m:r/m:nor`、希腊字母与常见运算符、`\mathrm{…}`（单位正体）。
   建议**先做 20 个真实公式的回归集**（从 `D:\DSH工作区` 的历史报告里抄），跑通再扩。
3. **扫描与替换**：正则照 `:1536`（`\$\$([^$\r]+?)\$\$|\$([^$\r]+?)\$(?!\$)`）；**从后往前**替换（`:1552`）保证位置不漂；先剥 `\tag{n}`（`:1538`）。
4. **行间公式**：`<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>…</m:oMathPara>`；编号用 `m:eqArr`+`maxDist`+`#`（照 `:1704-1708` 逐字节抄，这是 MS 自己导出的结构，兼容性最好）。
5. **插入后收尾**（照 `:1712-1741`）：删掉 `InsertXML` 多带的段落标记；**先存样式/对齐/缩进/段距/行距、插完还原**（WPS 实锤会丢，`:1715` 注释 + `verification.md:815`）；行间公式居中 + 上下 ≥6pt。
6. **编号挤压**：量 `Information(5)`（水平位置）差 <14pt 或折行 → 字号 −0.5pt，最多 70%（照 `:1744-1767`）。**注意**：`Information(5)` 我方从未用过，需先探一次（WPS 可能返回 0）。
7. **定长行距**：公式落地的段落按 P0-3 的 `Roomy` 修。
8. **失败兜底**：转换失败就保留 `$…$` 原文 + 结果点名 + 建议"改写更简单的写法"（照 `:1772-1778`）。

**WPS 实测线索**
- `docs/verification.md:695`（`OMaths.Add` 不定位、`BuildUp` 无效）是**否定性结论**——但它与已被我方推翻的 `Range.Cells` 结论同源（同上一条警告），所以按 §4.3 步骤 0 复核之前，只能当作"作者机上那条路径失败"。
- `:742`（11 公式全成、段落数不变）、`:888`（编号靠右）、`:918`/`:925`（编号挤压缩字号）、`:815`（行内公式丢样式 → 需还原）——四条都是 WPS 上跑出来的。
- 我方已有的 `render` 正好用来逐页核对渲染图（dsh-office 就是这么验的）。
**风险**
1. `Range.InsertXML` 在 **WPS 12.1.0.29161**（比作者机新）是否仍接受 Flat OPC 与 `m:oMathPara`，**必须自测**；失败特征预计是"插入空内容"或"抛 0x800A 参数错误"。
2. 子集转换器的正确性风险最高：OMML 里少一个 `<m:rPr><m:nor/>` 就变成斜体函数名；`\sum` 的 operand 取错会把整个式子吞进下标（`omml.ts:156-174` 那条"遇到 `=`、`+`、逗号就停"的规则一定要照抄）。
3. 存 `.docx` 后在微软 Word 里打开的效果（作者自己也列为"没验证"，`verification.md:748`）。
**工作量**：走 A（线性式）半天；走 B（OMML）2–3 天（含回归集与 20 个公式的逐条渲染核对）；若砍到"分式/上下标/根/括号/求和/希腊字母"六类，1 天可出可用版。

### 4.4 P1-3 fill_gaps

**先探 3 个原语**（我方只验过 `Information(3)`：`Layout.cs:374`、`Tables2.cs:456-457`、`Units.cs:112`）：
```
① doc.Range(p,p).Information(6)   —— 纵向位置（pt，相对页顶）；WPS 是否返回真值？
② para.OutlineLevel               —— 标题判据（10=正文）；我方 style 链路用过？
③ doc.Range(a,b).FormattedText = doc.Range(c,d).FormattedText  —— 整块搬运；WPS 是否可靠？
```
若 ③ 不可用，备选：`Range.Cut` + 目标处 `Paste`（会碰剪贴板、且有格式损失风险）→ 那就降级为"**只报不改**"（照 `KeepWhole` 的 `left/wide` 点名模式），价值仍然很大。

**算法照抄要点**（§5 摘录 4）：
- 块边界 = 图（前置空段 + 图 + 后题注）/ 表（前题注 + 表），再 `Widen` 吸掉相邻空段；
- 相邻图表先 `Absorb` 成一块（防止只搬半个）；
- 判"真有空档"：块前所在页 ≠ 块首页，且页脚剩余 ≥ `max(90pt, 页高×12%)`；
- **不越标题**：块前段（`before`）与块后段（`after`）都要求 `OutlineLevel==10`；
- 块后段以 `:`/`：` 结尾 → 连带搬下一段（"如下表所示："不会与表分离）；
- 每块最多试 8 次、全局有硬上限，结果报告"移动 N 段 / 覆盖 K 个图表 / 还有 J 处未填"。

**WPS 实测线索**：dsh-office 的 fill_gaps **只在 Word 上验过**（`verification.md:455` 明说"含文本框、分栏、脚注的文档没有试"；`:967` 说保存前自动跑只测了操作本身）——所以这条**完全是"需实测"**，我方的 WPS 拆解风险最高的是 ③（`FormattedText` 搬运）与 ①（`Information(6)`）。
**工作量**：探测 0.5 天 + 实现 1 天（含"保存前自动跑"开关，默认关，因为我们有用户在场）。

### 4.5 P1-4 跨页表单行修复

把 `stranded` 判据（`:2837-2846`）与修法（`:2894-2904`）搬到 helper，但**绕开 `Rows.Item` 残缺**：
```
判：RowsOnFirstPage 用 Range.Information(3)（我方已验证可用）；
修①：SelectRows(k..k) → app.Selection.Rows.ParagraphFormat.KeepWithNext = -1（我方已有 SelectRows/Selection 写法，Tables2.cs:187）
修②：兜底推空行 —— doc.Range(tableStart, tableStart).InsertBefore("\r") × n，最多 45 行，超 1/3 页高就撤回并点名
```
**风险**：`Selection` 写法会**移动用户光标/选区**（我方已有 `view`/`hint` 体系，但选区变动是可见副作用）→ 必须先存 `Selection.Start/End` 后还原。
**工作量**：1 天。

### 4.6 与我方架构对齐的两条纪律

1. **别把 dsh-office 的"逐段循环"直接搬进来**：它的 `Roomy`/`KeepWhole`/`FillGaps` 都是"逐对象 COM 往返"，在 Word 上够快，在我方常驻 helper 里虽然单次很快（1.3 s 级），但 `fill_gaps` 的"挪一段→重算边界→再挪"最坏是 O(n²)。**必须加硬上限 + 报告剩余**，并优先用"一次读全稿（`WordOpenXML`）本地判断 → 少量写"的 v1.13+ 路径。
2. **失败必须可见**：我方现有 `warning`/`note` 字段的用法（如 `Tables.cs:1335` 缩图告警）已经是这个风格，新能力沿用即可。

---

## §5 原文摘录（≤30 行/条）

### 5.1 脚注内联语法解析（先解析、后碰 COM；禁嵌套；倒序插入）
`helper/OfficeHelper.cs:2082-2106`
```csharp
static List<NoteToken> ParseNotes(string text) {
    List<NoteToken> notes = new List<NoteToken>();
    for (int i = 0; i < text.Length; i++) {
        if (Escaped(text, i)) continue;
        bool tex = text.Substring(i).StartsWith("\\footnote{", StringComparison.Ordinal);
        bool bracket = text.Substring(i).StartsWith("[fn:", StringComparison.Ordinal);
        if (!tex && !bracket) continue;
        int content = i + (tex ? 10 : 4), depth = 1, j = content;
        char open = tex ? '{' : '[', close = tex ? '}' : ']';
        for (; j < text.Length; j++) {
            if (Escaped(text, j)) continue;
            if (text.Substring(j).StartsWith("\\footnote{", StringComparison.Ordinal) || text.Substring(j).StartsWith("[fn:", StringComparison.Ordinal))
                throw new Fail("BAD_ARGS", "Nested footnotes are not supported.");
            if (text[j] == open) depth++;
            if (text[j] == close) { depth--; if (depth == 0) break; }
        }
        if (depth != 0) throw new Fail("BAD_ARGS", "Unclosed footnote: close it before editing the document.");
        string body = text.Substring(content, j - content);
        if (string.IsNullOrWhiteSpace(body)) throw new Fail("BAD_ARGS", "A footnote must contain text.");
        notes.Add(new NoteToken { Start = i, Length = j - i + 1, Text = body });
        i = j;
    }
    return notes;
}
```
→ 我方对应：目前**无**。这条的价值是"**在碰 COM 之前就把语法错误全部拒掉**"，与我方 revise 两轮预检同哲学。

### 5.2 WPS 里 `\tag{n}` 编号靠右的 OMML 构造（MS 自己导出的结构）
`helper/OfficeHelper.cs:1702-1709`
```csharp
// The number of an equation stands flush right, the equation itself in the middle: the form Word writes for
// "equation#(1)", an array of one line whose "#" sends what follows to the margin.
if (display && !string.IsNullOrEmpty(tag) && omml.StartsWith("<m:oMath>", StringComparison.Ordinal) && omml.EndsWith("</m:oMath>", StringComparison.Ordinal)) {
    string inner = omml.Substring(9, omml.Length - 19);
    omml = "<m:oMath><m:eqArr><m:eqArrPr><m:maxDist m:val=\"1\"/></m:eqArrPr><m:e>" + inner
         + "<m:r><m:t>#</m:t></m:r><m:r><m:rPr><m:nor/></m:rPr><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/></w:rPr><m:t>("
         + System.Security.SecurityElement.Escape(tag) + ")</m:t></m:r></m:e></m:eqArr></m:oMath>";
}
spot.InsertXML(FlatDocument(display
    ? "<m:oMathPara><m:oMathParaPr><m:jc m:val=\"center\"/></m:oMathParaPr>" + omml + "</m:oMathPara>"
    : omml));
```
→ 我方对应：**无**（`DESIGN-word.md:115`）。这是 P1-2 的核心 8 行。

### 5.3 Flat OPC 包装（WPS 能吃的最小文档）
`helper/OfficeHelper.cs:1663-1669`
```csharp
static string FlatDocument(string paragraph) {
    return "<?xml version=\"1.0\" standalone=\"yes\"?><pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">"
        + "<pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\"><pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships></pkg:xmlData></pkg:part>"
        + "<pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"><pkg:xmlData><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><w:body><w:p>"
        + paragraph + "</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
}
```

### 5.4 `fill_gaps` 主体（判空档 / 不越标题 / 冒号连带）
`helper/OfficeHelper.cs:2967-3016`（节选 2974-3013）
```csharp
foreach (int index in all) {
    int here = 0;
    bool more = false;   // the paragraph just moved announces what follows it (it ends with a colon)
    for (int step = 0; step < 8; step++) {
        int start, end;
        try { if (!Block(doc, index, out start, out end)) break; } catch (Exception) { break; }
        docEnd = (int)doc.Content.End;
        if (start <= 0 || end >= docEnd - 1) break;
        dynamic before = doc.Range(start - 1, start - 1).Paragraphs[1];
        if (!more) {
            if ((int)before.OutlineLevel != 10) break;                 // 不过标题
            dynamic tip = doc.Range(start - 1, start - 1), head = doc.Range(start, start);
            if ((int)head.Information[3] <= (int)tip.Information[3]) break;   // 没有跨页就不算空档
            dynamic setup = before.Range.Sections[1].PageSetup;
            double bottom = (double)setup.PageHeight - (double)setup.BottomMargin;
            double gap = bottom - Convert.ToDouble(tip.Information[6]) - 20;
            if (gap < Math.Max(90, (bottom - (double)setup.TopMargin) * 0.12)) break;   // 空得不够多
        }
        Absorb(doc, ref start, ref end);                               // 相邻图表当一块
        if (end >= docEnd - 1) break;
        dynamic after = doc.Range(end, end).Paragraphs[1];
        string text = ((string)after.Range.Text ?? "").Trim();
        if (text.Length == 0 || (int)after.OutlineLevel != 10 || (int)after.Range.InlineShapes.Count > 0 || IsCaption(doc, after)) break;
        try { if (Truthy(after.Range.Information[12])) break; } catch (Exception) { }   // 表内不搬
        doc.Range(start, start).FormattedText = after.Range.FormattedText;              // ★WPS 需实测
        if (!Block(doc, index, out start, out end)) break;
        Absorb(doc, ref start, ref end);
        doc.Range(end, end).Paragraphs[1].Range.Delete();              // 删掉搬走的那段
        moved++; here++;
        more = text.EndsWith(":", StringComparison.Ordinal) || text.EndsWith("\uff1a", StringComparison.Ordinal);
    }
    if (here > 0) blocks++;
}
```

### 5.5 定长行距 → 最小值（防裁切）
`helper/OfficeHelper.cs:1621-1649`（节选）
```csharp
static int Roomy(dynamic doc) {
    Cramped = false;
    int changed = 0, last = -1;
    try {
        int maths = (int)doc.OMaths.Count, count = maths + (int)doc.InlineShapes.Count;
        for (int i = 1; i <= count; i++) {
            try {
                dynamic paragraph = (i <= maths ? doc.OMaths[i].Range : doc.InlineShapes[i - maths].Range).Paragraphs[1];
                if (i == maths + 1) last = -1;
                int start = (int)paragraph.Range.Start;
                if (start == last) continue;
                last = start;
                dynamic format = paragraph.Format;
                if ((int)format.LineSpacingRule != 4) continue;     // 4 = wdLineSpaceExactly
                float height = (float)format.LineSpacing;
                format.LineSpacingRule = 3;                        // 3 = wdLineSpaceAtLeast，高度不变
                format.LineSpacing = height;
                changed++;
            } catch (COMException) { }
        }
    } catch (COMException) { }
    return changed;
}
```

### 5.6 WPS 差异总表（作者 0.9.6 的适配结论，节选 693–701 行）
`docs/verification.md:693-701`
```
- helper 每条命令先定套件：WPS 用 KWPS.Application、KET.Application、KWPP.Application，对象模型与微软 Office 相同，其余代码共用。
- WPS 下不同的地方：
  - 公式不排版，$...$ 保留为文本，并在结果里告诉模型（WPS 的 OMaths.Add 不按给定位置插入，BuildUp 不生效）。
  - “重置段落格式”在 WPS 里会把样式一并清掉、把新表格拆成多张，改为逐项清除缩进。
  - WPS 没有 Range.Cells：表格的单元格数按行乘列算，垂直对齐逐格设置。
  - 表格里的文本替换在内存里做（WPS 的替换会在刚替换出的内容里再替换一次，“P”换“P-”得到“P--”）。
  - 按名字找不到形状时 WPS 抛的是参数错误，同样报“没有这个形状”并列出现有形状。
  - WPS 演示不给出窗口句柄，迷你卡片按进程和标题找窗口。
```
→ 这段是 0.9.6 时的状态（其中"公式不排版"在 0.9.7/1.0.3 已被 OMML 路线解决）。对我方有效的三条：①`Range.Cells` 缺失——**已被我方推翻**（`word-quirks.md:22`），正好证明"别照抄别人的 WPS 结论"；②`OMaths.Add`/`BuildUp` 无效——**已由 Lead 在 29161 实测确认**（`Add` 抛「值不在预期范围内」），可以写进 quirks 定案；③`ParagraphFormat.Reset()` 会清样式/拆表——对我方无影响，因为 helper 全程**不调用** `ParagraphFormat.Reset()`（grep 零命中），段落格式都是逐项赋值（`WpsHelper.cs` 的 `ApplyParaFormat`）。

### 5.7 不拆词极性的实测记录（与 MS 文档相反）
`docs/verification.md:146`
```
- 西文单词和数字在行尾被从中间拆开（`37 / 000`、`mode / l)`），模型为此花了约 200 秒用鼠标去改段落对话框。原因：用户模板里的标题样式带有“允许西文在单词中间换行”，其后的段落逐段继承。Word 对象模型里这个属性的含义与文档写的相反：实测 `WordWrap = -1` 时 `37000` 在同一行，`= 0` 时被拆到两行。现在插件写入的段落、表格、题注一律设为 -1。修后在 Word 里量过：`37000` 首尾字符同在第 4 行。
```
→ 我方 `scripts/helper/word/WpsHelper.cs:632` 写的是 `pf.WordWrap = v ? 0 : -1`，**语义疑似取反**，见 §4.1。

### 5.8 字段点名（模型一定能看到的那种）
`helper/OfficeHelper.cs:9666-9670`
```csharp
if (card) Card.Hold = true;
if (card) Card.Report("AI 正在编辑 " + docName + (ops.Count > 1 ? " · " + (i + 1) + "/" + ops.Count : ""));
string line = kind == "word" ? WordOp(doc, type, op) : kind == "excel" ? ExcelOp(app, doc, type, op) : PptOp(doc, type, op);
string ignored = op.Unread();
done.Add(ignored.Length == 0 ? line : line + " — IGNORED field(s) " + ignored + ": " + type + " does not take them, check the operation list");
```
→ 我方 `param-check.ps1:50` 走 stderr；建议**同时**把 WARN 汇总进 JSON 的 `result.warnings`，避免模型忽略 stderr。

### 5.9 保存前自动兜底（`fill_gaps` 与视觉检查都挂在保存上）
`docs/verification.md:945`、`:952`
```
- 论文：模型调用过一次 fill_gaps，之后又有编辑，第 8 页页尾在图 3 前重新空出一段。对最终文件再跑一次，移动了 1 段。
- office_save 保存 Word 文档前，如果上次查看之后有过编辑，自动再做一次 fill_gaps，移动了段落就在回复里说明。
```
→ 我方可在 `save`（`word.ps1:598-601`）里加 `-AutoFillGaps` 开关，默认**关**（用户在看着文档改时，静默重排会让他困惑），但 `save` 前若检测到"本会话没 render 过任何页"则在结果里点名提醒。

### 5.10 我方现有实现的两处对照（写清楚了才知道改什么）
`scripts/helper/word/Tables.cs:686-696`（三线表口径与 dsh-office 完全一致，无需改）
```csharp
// three-line: outer top + outer bottom thick, a thin rule under the header row, nothing else.
TabBorder(borders, BT_TOP, LS_SINGLE, LW_150);
TabBorder(borders, BT_BOTTOM, LS_SINGLE, LW_150);
try { TabBorder(table.Rows.Item(1).Borders, BT_TOP, LS_SINGLE, LW_150); } catch (Exception) { }
try { TabBorder(table.Rows.Item(rows).Borders, BT_BOTTOM, LS_SINGLE, LW_150); } catch (Exception) { }
if (rows > 1) { try { TabBorder(table.Rows.Item(1).Borders, BT_BOTTOM, LS_SINGLE, LW_075); } catch (Exception) { } }
done.Add("borders=three-line (1.5 pt top/bottom, 0.75 pt under the header, no other lines)");
```

`scripts/helper/word/Tables.cs:959-980`（P0-4 的疑点：用的正是自己文档里声明会抛异常的成员）
```csharp
private static List<string> TabKeepTogether(dynamic table) {
    ...
    for (int r = 1; r <= rows; r++) {
        try { dynamic row = table.Rows.Item(r); row.AllowBreakAcrossPages = 0; kept++; }
        catch (Exception) { }
        if (r < rows) {
            try { table.Rows.Item(r).Range.ParagraphFormat.KeepWithNext = -1; } catch (Exception) { }   // ← Rows.Item(r).Range 在 WPS 抛 0x800A1767（Tables2.cs:5）
        }
    }
    done.Add("keepTogether: " + kept + " row(s) cannot break across pages, and every row but the last keeps with the next");
```
对照：`scripts/helper/word/Tables2.cs:5` 写着 `table.Rows.Item(n).Delete() / .HeadingFormat / .Range / .Cells throw 0x800A1767`，且 `Tables2.cs:187` 已给出正解 `app.Selection.Rows.HeadingFormat = -1`（带回读校验 `:188`）。

---

## 附：本研究**没有**核实的部分（留给实测/下一手）

| 项 | 为什么没核实 | 怎么核实 |
|---|---|---|
| **Flat OPC 整包的 `Range.InsertXML` 能否在 WPS 29161 变成真公式** | Lead 已测"裸 OMML 片段 → 调用成功但 `OMaths.Count` 不变"；**Flat 包形态没测过**，而 dsh-office 用的正是 Flat 包 | §4.3 步骤 1（三判据看 `WordOpenXML`），30 分钟即可定性 |
| `Shapes.AddCanvas` + `CanvasItems.AddShape/AddLine` 在 WPS 是否可用 | Lead 只测了文档级 `Shapes.AddShape`（可用）与 `AddConnector`（E_FAIL）；dsh-office 用的是画布那一层 | 一条脚本：`Shapes.AddCanvas(0,0,300,200)` → `CanvasItems.AddShape(1,…)` / `AddLine` → `ConvertToInlineShape()` |
| `Information(6)`（纵向位置）在 WPS 的返回值 | 同上 | 3 行探测脚本 + `render` 对照 |
| `Range.FormattedText` 整块赋值在 WPS 的行为 | 同上 | 临时文档：两段 A/B，把 B 搬到 A 前，读回文本+格式+段数 |
| `WordWrap` 在 **WPS** 的极性与 Word 是否一致 | 同上 | §4.1 脚本 |
| `Footnotes.Add` 在 WPS 2003/兼容模式文档上是否同样可行 | Lead 只在 .docx 上验过 | 另存为 .doc 后重试一条 |
| 修订模式下 `Footnotes.Add` 是否留痕、拒绝修订后脚注是否回滚 | dsh-office 未测（README 未提），只能自测 | 临时文档 + `TrackRevisions=true` + `revisions reject -All` 后数脚注 |
| 存成 `.docx` 的 OMML 公式在微软 Word 里打开的效果 | 本机无微软 Word | 交给有 Word 的同事（dsh-office 自己也把它列为未验证，`verification.md:748`） |
| 我方 `-KeepTogether` 是否真的失效 | 需要跑命令 | `table -KeepTogether` → 读回 `Rows.Item(r).AllowBreakAcrossPages` 与整表页码；`render` 看跨页 |
