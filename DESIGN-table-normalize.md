# DESIGN — `table -Action plan|normalize`（v1.33 提案）

> 起因：2026-11-03 中午 DSH 会话 `session-a305706b`（乐平下闸蓄水设计自检报告，14 张表：题注编号问题 + 表名在表内 + 所有表格加粗边框）。任务本身很小，却用了 **127 次工具调用 / 约 88 分钟**。
> 参考实现：`design-report-format`（一条 exe 调用内部完成 scan → format → table 全流程，进度走 stderr，无来回）。
> 本文件是给实现者的规格书。**接口、行为、判据以此为准**；实现细节可优化，但验收标准不能减。

## 1. 目标

一条命令完成整篇 Word 报告的「表格规范化」，把 20–40 次 tool call 压到 **1–2 次**，并把两个已知缺陷在命令内部修掉：

| 缺陷 | 会话里的代价 | 修复要求 |
|---|---|---|
| 题注外提遇「表前的段落带分页/分节符」→ 题注落在上一页（孤儿题注） | 表 2.1.9 提取 E_FAIL + 三种写法试探 + 渲染确认 ≈ 13 min | **结论（2026-11-03 用户定调）：WPS 上不做可靠的表名外提。** 只对「普通段落后 / 手动分页符前」的表尝试，且**回读题注页码 == 表起始页**才接受；**分节符在表前 → 一律 `needsHuman`，不重试、不研究**（彻底解决靠人工挪分节符或微软 Word） |
| 给无表样式的表套边框 → WPS 自动挂默认样式（`tblStyle=77`）、列宽变百分比 → 表被压成一行一字 | 表 2.1.7 版式崩、查因 + 开副本实验 + 换 XML 取证 ≈ 37 min | 套边框前后**快照/回读列宽与字号**，变了就按快照恢复并记录 `widthsRestored`；单表失败不影响其它表 |

另加三项机制化能力：**结构侦察（plan）**、**题注编号规范化（可选）**、**自检（verify）**。

## 2. 命令接口

### 2.1 `table -Action plan`（只读，必须先跑）

```powershell
office.ps1 word table -Doc $DOC -Action plan [-Label 表] [-Max 200]
```

返回（JSON）：`tables[]`（每张表：`table` / `para` / `pageFrom` / `pageTo` / `rows` / `columns` / `captionInFirstRow` / `captionText` / `captionPara` / `captionPos` / `captionPage` / `split` / `headerRepeated` / `outerWidth` / `innerWidth` / `hasStyle` / `styleName` / `widthMode`（points/pct/none）/ `widths`）+
`captions[]`（`table` / `para` / `pos` / `number` / `format` / `gap` / `title` / `page`）+ `references[]`（`para` / `number` / `n`）+ `issues[]`（复用 `captions` 的规则码）+ `plan.renumber`（见 3.3）+ `risks[]`（每表一条：`breakBefore` / `noStyle` / `pctWidths` / `captionInFirstRow` / `multipageWithoutHeaderRepeat`）+ `summary`。

要点：**一次调用把「动手前必须知道的事」全给出来**——会话里这些信息是靠 `captions` + `scan` + 逐段 `read` + 读 helper 源码拼出来的。

### 2.2 `table -Action normalize`（默认 = 全套；逐表隔离）

```powershell
office.ps1 word table -Doc $DOC -Action normalize [-Table n] [-Plan <plan.json>] [-DryRun] [-NoCaptions] [-NoBorders] [-NoHeaderRepeat] [-Gap 2] [-Borders thick] [-Fit keep|window|content] [-Verify] [-Force]
```

默认动作（按顺序，**顺序由命令内部保证**）：

1. **题注外提**：所有 `captionInFirstRow=true` 的表 → 题注段移到表格上方、删表内题名行（复用 `TabCaptionMove` 逻辑 + §3.1 新定位算法）；
2. **题注空白规范化**：题注段内「标签 + 编号 + 题名」之间统一为 `-Gap` 个半角空格（默认 2，与中文报告惯例一致；`-Gap 0` 关闭）；
3. **编号规范化**（仅当给了 `-Plan <plan.json>` 且其中 `renumber.approved=true`）：按 §3.3 的映射改题注编号 + 正文唯一引用；
4. **加粗边框**：`-Borders thick`（1.5 pt 外框 + 0.5 pt 内线），套完按 §3.2 校验并恢复列宽/字号；
5. **跨页表头重复**：复用 `TabHeaderRepeat` 迭代逻辑（题注外提之后再跑）；
6. **自检**：复跑 scan 判据，输出 `verification`。

返回值：

```json
{"ok":true,"cmd":"table","result":{
  "action":"normalize","tables":14,
  "summary":{"captionsMoved":2,"captionsMovedFailed":0,"spacesNormalized":14,
             "renumbered":0,"referencesUpdated":0,"bordersApplied":14,"widthsRestored":1,
             "headerRepeated":1,"pageBefore":29,"pageAfter":29,
             "revisionsBefore":20,"revisionsAfter":20,"commentsBefore":2,"commentsAfter":2},
  "verification":{"captionInFirstRow":0,"bordersNotThick":0,"multipageWithoutHeaderRepeat":0,
                  "captionPageMismatch":0,"widthsChanged":0,"pageDelta":0},
  "tables":[{"table":9,"page":20,"caption":{"before":"表2.1.9 ...","after":"表 2.1.9 ...","para":3158,"page":20,"method":"table-start"},
             "borders":{"before":[1,4],"after":[1,12],"thick":true},
             "widths":{"before":[30.9,28.4,...],"after":[30.9,28.4,...],"restored":false},
             "font":{"sizeBefore":10.5,"sizeAfter":10.5},
             "headerRepeat":{"applied":false},"warnings":[]}],
  "failures":[{"table":3,"code":"CAPTION_MOVE_FAILED","detail":"..."}],
  "saved":false,
  "note":"format-only, one undo step; render the affected pages before saving."}}
```

约束：

- **逐表 try/catch**：单表失败 → 撤销该表已做的改动（`doc.Undo()` 一次，前提是该表包在自己的 UndoRecord/或记录起始撤销计数）+ 记 `failures[]`，**继续处理其它表**；不允许像 v1.32 那样整批回滚。
- 整条命令仍是**一个用户级撤销单元**（`P2Edit` 的 `StartCustomRecord`），Ctrl+Z 一次全撤。
- 不动修订状态（不产生 w:ins/w:del），不动已有批注；`commentsBefore/After` 不一致要 WARN（会话里 undo 误删过批注）。
- `-DryRun`：只算并返回计划与将发生的改动，不写文档（只读文档也能跑）。
- 不 save。`render` 复核由 agent 决定。

### 2.3 `render -Table n`（顺手加）

`render -Doc X -Table 9` → 用 `PageSpanOf` 自动把页码设为该表跨的页（含题注页），只渲这几页。省掉「先 scan 拿 pageFrom/pageTo，再两次 render」的来回。

## 3. 关键算法

### 3.1 题注外提的插入位置（修 bug 1；**最终结论：WPS 上不承诺，分节符场景 needsHuman**）

> 2026-11-03 收口：本文下面的三种插入法 + cut-paste 兜底在微软 Word 上成立；在 WPS 12.1 上「分节符紧邻表格」的表没有任何可用原语（所有插入都落回上一节），cut-paste 还会把 Writer 转死。**该场景不再研究，统一报 `needsHuman{captionMoved:"wps-section-start"}`。** 实际方法链收敛为 `table-para → range-insert → table-start → before-para`（cut-paste 仅真 Word / `-CutPaste`）。

现状（`Tables2.cs:237 TabCaptionMove`）：取表格前一段 `prevPara`，在 `prevPara.Range.End - 1` 插 `"\r" + caption`。**当 `prevPara` 带手动分页符（`\f`, U+000C）或分节符（`sectPr` 在段末标记）时，插进去的段落在断点之前 → 题注留在上一页**（实测表 2.1.9：paragraph 3158 = `caption + "\u000c"`，题注在 19 页底、表在 20 页）。

新算法（按顺序尝试，第一个「回读通过」的即采用）：

1. **判断断点**：`prevPara.Range.Text` 含 `\f`；或 `prevPara.Range.ParagraphFormat.PageBreakBefore`；或 `prevPara` 与表的 `Range.Information(2)`（节号）不同；或 `prevPara` 的文字以空段 + 表首行页码 > prevPara 页码。→ 记 `breakBefore=true`。
2. **方法 A（有断点优先）**：`table.Range.InsertParagraphBefore()` 生成新段，再写题名文本（注意：先插空段、后 `np.Range.Text = caption`，避免把文本塞进首格）。回读：`np.Range.Information(3)`（页码）必须 == 表起始页 `table.Range.Information(3)`。
3. **方法 B（无断点，或 A 失败）**：沿用现状 `prevPara.Range.End - 1` 插 `"\r"+caption`；回读同上。
4. **方法 C（兜底）**：`doc.Range(tableStart, tableStart).InsertBefore("\r")` + 写文本；回读同上。
5. **回读不通过**（题注页码 < 表页码，或段落 `inTable`）：删除刚插的段（`doc.Undo()` 或按段号删）→ 换下一法；三法都失败 → `CAPTION_MOVE_FAILED`，该表回滚，其它表继续。
6. 删除表内题名行：沿用 `DeleteRowViaSelection`（`Tables2.cs:429`，WPS `Rows.Item(n)` 残缺绕行）；回读行数必须 -1。
7. 题注格式：沿用现逻辑（继承题名单元格的 `Font.Name/NameFarEast/Size` + 对齐，缩进清零）。**不额外缩字号**。

### 3.2 边框 + 列宽保持（修 bug 2）

**实测修正（2026-11-03 审计复现，推翻原假设）**：WPS 没有"自动挂样式 77"，也没有把宽度改成百分比（磁盘上 `tblW/tcW` 本来就是 pct）。真正触发条件 = **该表 XML 里缺 `w:tblBorders` 元素**：第一次写边框属性迫使 WPS 补建边框描述并**重建该表布局**，`w:gridCol` 按"当前渲染宽度"重新推导（容量取整）→ 列宽漂移。破坏力由该表有无 `tblStyle`、宽度是 pct 还是 dxa 决定——原稿那张表恰好无样式 + dxa 固定，于是崩。fixture 上 T3（18 列 fixed）肉眼变宽 16%、T9 漂 −30 dxa，其余 12 张有 `tblBorders` 的表纹丝不动。

另：**COM 读列宽不可用**——WPS 对 autofit 列 `Column.Width` 返回 `wdUndefined`（实测 `table -Action format` 的 widths 全是 9999999）。唯一便宜可靠的通道是 `doc.Content.WordOpenXML`（实测 **0.16 s / 4.1 MB**）。

`normalize` 的流程：

1. **开工前 1 次全文 XML**（`Content.WordOpenXML`）：正则解析每表 `tblStyle / tblW(w,type) / tblLayout / gridCol[] / 首行 tcW / tblBorders 是否存在`，存快照；`plan` 也用它出结构字段。
2. **短路**：已达标的表（`tblBorders` 存在且 outer top ≥12、innerH ≤4）跳过，不白改（实测全局 thick 改 14 张只有 2 张真变）。
3. 套边框（复用 `TabBorders`）。
4. **写后 1 次全文 XML**，逐表比对 `gridCol[]` 与 `tblW`：相对误差 >2% 或 sum 变化 >2% 判漂移。
5. **恢复（只对漂移表）**：`table.PreferredWidthType/PreferredWidth` 写回快照类型与值 → 逐列 `table.Columns.Item(i).Width = gridCol_dxa/20`（点）→ 再读 1 次 XML 复核；仍不一致 → `WIDTHS_CHANGED` WARN 并给 before/after 数组，**不静默**。
6. 字号：读 `table.Range.Font.Size` 前后比对，变了按快照恢复（失败 WARN）。`Style` 空→非空也记 `styleAttached`。
7. `-Fit window`：恢复后按原比例把列缩放到 `TextWidth(doc)`；`-Fit keep`（默认）只保持原宽；`-Fit content` 调 `AutoFitBehavior(2)`。
8. 边框回读判据用现有 `TabBorderReadback`（`outer top >= 12 && innerH <= 4`）。

**禁止**：不要逐字符/逐单元格清空格（参考实现开了这个功能后表格阶段 1.6 s → 126.9 s，+79×）。

### 3.3 题注编号规范化（可选，plan → 批准 → apply）

**不自动猜语义**。`plan` 只做机械提议：

1. 按文档顺序取所有题注（`CaptionAbove`/`CaptionBelow`/表内题名，复用 `Captions.cs` 解析器），按 `NumberPrefix()` 分组（如 `2.1`、`2.3`、`2.4`）。
2. 每组内按出现顺序提议 `prefix.1, prefix.2, ...`，生成 `plan.renumber.items[] = [{table, para, from, to, unique:true|false}]`；`from` 在该组内重复时 `unique=false`。
3. 正文引用扫描（`RefRx`）：每个引用按「同组内 `from` 是否唯一」判定——唯一才给自动映射，否则 `ambiguous:true` 并列出候选；**有歧义时 apply 一律拒绝自动改**（`RENUMBER_AMBIGUOUS`），要求人/agent 给显式映射。
4. 显式映射文件（`-Plan` 指向的文件里 `renumber.items[]` 由 agent 改好、`renumber.approved=true`）：apply 时按 `para` 定位题注段改数字（`revise` 式：只改数字 token，保留题名/空格），正文引用按显式 `refs[]` 映射改。**题注改动走修订**（可选 `-Track`，默认走修订——编号属于文字，保留可拒绝性）、格式类不走修订。
5. 判据：apply 后复跑题注解析，`duplicates=[]`、`missing=[]`、`spaces` 统计一致；返回 `renumbered` / `referencesUpdated`。

> 会话里这件事靠模型读 20 段 + 三轮 revise 批次完成（≈10 min）；plan 把「清单 + 提议 + 歧义」一次给全，agent 只需确认/改映射。

## 4. 验收标准

在 `D:\DSH工作区\乐平表格整理\orig-copy.docx`（坏稿基线：14 表 / 29 页 / `captionInFirstRow≈4` / `bordersNotThick≈8` / 20 处修订 / 2 条批注）副本上：

| # | 断言 |
|---|---|
| A1 | `plan` 一条命令返回：14 表逐表风险、4 张表内题注、8 张边框不达标、表前带断点的表 ≥2、正文表号引用 ≥20，且不写文档（`saved:false`、修订/批注数不变）。 |
| A2 | `normalize -DryRun` 不改变段落数/页数/修订数。 |
| A3 | `normalize` 后 `verification.captionInFirstRow==0`、`bordersNotThick==0`、`multipageWithoutHeaderRepeat==0`、`captionPageMismatch==0`、`widthsChanged==0`，页数 29（±1）。 |
| A4 | 表 2.1.9（表前有断点）题注段落页码 == 表起始页；表 2.1.1/2.1.3/2.1.7 题注在框外、行数各 -1。 |
| A5 | 表 2.1.7（14 列 / 无样式）列宽 before≈after（相对误差 ≤2%），字号不变；`widthsRestored` 若为 true 则 `widths.after` 必须 ≈ `before`。 |
| A6 | 修订数、批注数与执行前一致（格式类不改修订；不误删批注）。 |
| A7 | `ok:true` 且执行时间 ≤ 60 s（单条命令，热 helper）；`-Action normalize` 全程一次 tool call。 |
| A8 | 单表故障隔离：用 `-Table` 指向不存在的表 → `BAD_ARGS`；人为构造失败表时其它表仍完成（`failures[]` 有记录、`ok:true`）。 |
| A9 | 回归：既有 `table -Action scan/caption/header-repeat/rows`、`captions`、`render` 行为不变（旧测试通过）。 |

## 5. 文件落点（实现者）

| 内容 | 文件 |
|---|---|
| `plan` / `normalize` 主体 | 新文件 `scripts/helper/word/Tables3.cs`（`internal static partial class Program`） |
| 动作分发 | `scripts/helper/word/Tables.cs:186-197`（`if (action=="plan"...) return TabPlan(a); if (action=="normalize"...) return TabNormalize(a);`） |
| 题注外提定位改造 | `scripts/helper/word/Tables2.cs:237 TabCaptionMove`（抽成 `MoveCaptionAbove(doc, table, out method)` 供两处复用） |
| 边框+列宽快照/恢复 | `scripts/helper/word/Tables.cs:656 TabBorders` 旁新增 `TableWidthSnapshot`/`RestoreTableWidths`；`TabFormat` 也可复用 |
| 题注解析/编号 | 复用 `Captions.cs` 的 `ParseCaption`/`NumberParts`/`NumberPrefix`/`CompareNumbers`/`CaptionAbove`/`CaptionBelow` |
| 分页读回 | 复用 `Tables2.cs:451 PageSpanOf`、`Information(3)` |
| 包装参数 | `scripts/word.ps1`：`table` 分支增加 `plan`/`normalize` 相关参数（`-Plan` 文件路径、`-NoCaptions`、`-NoBorders`、`-NoHeaderRepeat`、`-Gap`、`-Fit`、`-Verify`）；`action` 已有小写化转发 |
| `render -Table n` | `scripts/helper/word/WpsHelper.cs:1604 RenderDoc`（`a.Int("table",0)>0` → 用 `PageSpanOf` 定页码） |
| 测试 | `tests/word/`：新增 `check_normalize.py`（A1–A8 断言）+ 复用 `tests/word/run_acceptance.ps1` 习惯 |
| 文档 | Lead 改 `SKILL.md`（Word 命令速查 `table` 行 + 工作流第 11 条）与 `reference/workflows.md` §12 |

## 6. 非目标

- 不做「自动理解题名语义后重排编号」（歧义时必须人确认，§3.3）。
- 不改字体/字号（`-Size` 走既有 `TabFormat`）。
- 不处理图题（`captions -Label 图` 已有）。
- 不引入 MS Word 依赖；一切走 WPS/Word 通用 COM。
- 不自动 save。

## 7. 实施后记：一次把 WPS 转死的挂起（2026-11-03 17:00–17:15）

**现象**：v1.33 首次回归时，对测试副本执行 `table -Table 1 -Action caption`（旧命令、新实现），WPS Writer 进程 100% CPU、`Responding=false`、COM 全部挂起，4.5 分钟无返回，只能结束进程（用户已在 Microsoft Word 里工作，未受影响；WPS 内那份是会话恢复副本，`saved=true`）。

**定位**：最小文档上复现（`.tmp/hangrepro/min-*.docx` + 带 watchdog 的 `try-kind.ps1`），trace 最后一行停在 `caption move table 1 via cut-paste`。两个叠加的代码缺陷：

1. `DetectBreakBefore` 用 `prev.Range.WordOpenXML` 里找 `<w:sectPr` 判"表前有分节符"——**Range 级 WordOpenXML 会带上该 Range 所属节的 sectPr**，于是普通单节文档的每张表都被判成"表在节首"，方法链第一项变成 cut-paste。
2. cut-paste（`table.Delete()` + `Range.InsertXML(表)`) 在 **WPS 12.1** 上会把 Writer 进程锁死（微软 Word 上可行，参考实现即如此）；实测删除后 InsertXML 放回 → 无限循环。

**修复**（已落地）：
- 分节符判定改为**节号比较**（`prev.Range.Sections(1).Index != table.Range.Sections(1).Index`），不再看 XML；
- 最终实现只保留 **`before-para`**（表前一段段末标记前插 CR）+ 回读"题注页 == 表起始页"；`table-para`/`range-insert`/`table-start` 三种写法在 WPS 上实测都会把题注插到上一节或表格第一格，且清理失败尝试会毁表，已删除（留档见 `reference/workflows.md` §12.3）；
- cut-paste 降级为**真 Word 专属**（按 `app.Path`/`Version` 判 `IsWpsApp`），WPS 上必须显式 `-CutPaste`；
- 三法都不过 → `needsHuman`，**绝不为一张表把整条命令挂住**；
- `cutPaste` 键名大小写要与包装器一致（`Bag.Flag` 大小写敏感）。

**纪律沉淀**：含「删表/重建/InsertXML」的兜底路径，必须在**最小合成文档**上先跑一次并加 watchdog；WPS 调用不能用"超时-重试"救，只能由外部看门狗杀掉进程。

