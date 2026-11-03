# skill-word v1.30：WordWrap 极性、表格 KeepTogether 回读、fillgaps

日期：2026-11-02　环境：WPS 12.1.0.29161（`KWPS.Application`，`doctor.app.build` = `12.1.0.29161`）
范围：**只动 Word 侧**——`scripts/word.ps1`、`scripts/helper/word/*`、`scripts/helper/WpsHelper.cs`。
Excel / PPT / `SKILL.md` / `publish.ps1` / `PipeServer.cs` / `woc-common.ps1` / `param-check.ps1` 未改。

---

## 0. 一句话

三件事：① `noSplitWords` 的 COM 极性写反了（写 `0` 才是拆词），已改正并**逐次回读**；
② `table -KeepTogether` 原来把异常吞掉、报成功，已改成**逐行 `SelectRows` + 逐行回读**，不接受就报失败；
③ 新增 `fillgaps` 命令（图/表被挤到下一页时，把后面的正文段提到前面填掉上一页底部的留白）。

---

## 1. WordWrap 极性（`ApplyParaFormat`）

### 事实

`ParagraphFormat.WordWrap` 与"整词换行"的关系**与直觉相反**，本机实测（页宽 240 pt、左右边距各 90 pt → 文本列 60 pt、20 pt Arial、"aa 3700000"）：

| 写入值 | 读回 | 渲染结果 | 含义 |
|---|---|---|---|
| `-1` | `-1` | 第 1 行 `aa`，第 2 行 `37000`，第 3 行 `00` | **整词**：数字不从词中间起行（超长 token 仍会在行内被折断，见下） |
| `0` | `0` | 第 1 行 `aa 37`，第 2 行 `00000` | **拆词**：西文词从中间断开 |

证据：`%TEMP%\woc-v130\shots\page-1.png`（写入 -1）与 `%TEMP%\woc-v130\shots0\page-1.png`（写入 0）；
同一张图上也能看 A/B：`%TEMP%\woc-v130\shots-insert\page-1.png`（第 1 段 `0` = 拆词态，第 2 段是 `insert` 新建、默认整词）。

> 注意一个**读图的细节**：文本列只有 60 pt 而 "3700000" 需要约 77 pt，所以即使 `WordWrap=-1`，
> WPS 也会把这一整个超长 token 在行内折断（`37000` / `00`）——它保住的是"不从空格处提前拆"，
> 不是"永不折断"。判读整词与否的**判据是首行**：首行只有 `aa` 就是整词（数字另起一行），
> 首行是 `aa 37` 就是拆词。Lead 报的 "aa|3700000" 与本机渲染是同一件事的两种写法。

### 改动

- `ApplyParaFormat`：`pf.WordWrap = v ? -1 : 0`（原来写成 `v ? 0 : -1`，极性反了），
  写入后 `Convert.ToInt32(pf.WordWrap)` **回读**，返回串形如 `noSplitWords=true (readBack=-1)`；
  读回值与要求不符 → 抛 `Fail("COM_ERROR", ...)`，**不再静默**。
- 新工具函数 `SetWordWrap(pf, wholeWord, what)`（`WpsHelper.cs`）：写入 + 回读 + 不符即抛，一处实现三个调用点共用。
- **我方新建的内容默认整词**（可被显式 `noSplitWords:false` 覆盖）：
  - `insert` 新段落（`Text.cs`，每个新段 `DefaultWordWrap`，结果进 `result.wordWrap`）；
  - `table` 新建表格的单元格（`Tables.cs` `TabInsert`，`table.Range.ParagraphFormat` 一次设）；
  - `caption` 新题注段（`Tables.cs` `AddCaption`，结果进 `caption.noSplitWords`）。

### 冒烟

```
word format -Doc wrap.docx -Para 1 -NoSplitWords            -> applied: ["noSplitWords=true (readBack=-1)"]
word format -Doc wrap.docx -Para 1 -NoSplitWords:$false     -> applied: ["noSplitWords=false (readBack=0)"]
word insert -Doc wrap.docx -Para 1 -Items items.json        -> wordWrap: ["noSplitWords=true (readBack=-1)"]
```

---

## 2. `table -KeepTogether` 真落上

### 事实

- `table.Rows.Item(n)` 在 WPS 上是**半实现对象**：`.Delete()` / `.HeadingFormat` / `.Range` / `.Cells` 抛 `0x800A1767`
  （`Tables2.cs` 文件头已记）。老 `TabKeepTogether` 对 `Rows.Item(r).Range.ParagraphFormat.KeepWithNext` 的写入
  **包在 `try { } catch { }` 里**，所以"报成功"但可能一行都没落上。
- 走 `Selection` 才行：`table.Cell(r,1).Range.Select()` + `app.Selection.MoveEnd(10 /*wdRow*/, 1)`，
  然后 `app.Selection.ParagraphFormat.KeepWithNext = -1` / `app.Selection.Rows.AllowBreakAcrossPages = 0`。

### 改动

重写 `TabKeepTogether(app, table)`（`Tables.cs`），返回结构化结果：

```json
{"ok":true,"rowsExpected":8,"rowsSet":8,"readBackFailures":[],
 "allowBreak":{"rowsExpected":9,"rowsSet":9,"readBackFailures":[]},
 "summary":"keepTogether: 8/8 row(s) keep with the next row, and all 9 row(s) refuse to break across pages (each value read back)"}
```

- 每行（除最后一行）写 `KeepWithNext=-1` **并立即读回**；每行（含最后一行）写 `AllowBreakAcrossPages=0` **并读回**。
- 有失败：`summary` 里点名 `WPS 没接受 rows A,B (KeepWithNext) / C (AllowBreakAcrossPages)`，
  并且 `table` 命令**整体返回 `ok:false / error.code=KEEPTOGETHER_FAILED`**（`Fail.Detail` 带完整结构，
  为此在 `WpsHelper.cs` 的 `catch (Fail)` 里补了 `error.detail` 透传）——绝不报成功。
- 结果挂在 `result.tables[].keepTogether`（成功时）或 `error.detail`（失败时）。

### 冒烟

短页（400×330 pt、上下 45 pt）+ 9 行 × 3 列、每行 10 行文字 → 表格跨页、第 3 行被从中间劈开：

- 改前：`%TEMP%\woc-v130\kt-before\page-2.png`（第 2 行完整 + 第 3 行在页底被截断）；
- 跑 `table -Table 1 -KeepTogether` → `rowsSet 8/8`、`allowBreak 9/9`、`readBackFailures []`；
- 改后：`%TEMP%\woc-v130\kt-after\page-2.png`、`page-3.png`（每页一行，行不再被劈开）。

---

## 3. 新命令 `fillgaps`

### 语义

照 dsh-office `helper/OfficeHelper.cs` 的 `FillGaps`（`OfficeHelper.cs:2967-3016`，辅助 `Block` :2754-2772、
`Absorb` :2761、`Widen` :2778、`IsCaption` :2721）移植：**图/表整块放不下跳到下一页 → 上一页底部留白 →
把图/表后面的正文段提到图/表之前填这段空白**。

### 参数

| 参数 | 默认 | 说明 |
|---|---|---|
| `-Doc` | — | 必填（走活文档） |
| `-DryRun` | 关 | 只报不动；**只读文档也能跑**（非 dryRun 时遇到只读抛 `READONLY`） |
| `-MaxMoves N` | 8 | 每个块最多搬几段（1–40） |
| `-MinGapPt N` | 0 | 留白阈值；不给则 `max(90, 文本高度 × 12%)` |

### 判据（保守）

块前一段必须 `OutlineLevel == 10`（正文），块首行页码 > 前一段页码（真被挤下去了），
空白 `gap = (PageHeight - BottomMargin) - 前一段 Information(6) - 20`，不足阈值跳过；
块后第一段须非空、`OutlineLevel == 10`、无 `InlineShapes`、不是题注（复用 `ParseCaption` + 样式名）、
`Information(12)` 不在表内。**不越过标题**（`OutlineLevel != 10` 立即停）。

### 返回

`{moved, blocks, minGapPt, dryRun, moves:[{block, kind:"figure|table", para, gapAtPara, page, gap, text}], notes, summary, pages, saved:false}`
`moved==0` 时 `summary` = 「没有一页在图表之前留白到值得搬段 …」。

### 实现要点 / 踩到的坑

1. `Range.FormattedText` 可赋值搬运（实测），搬完用同一 `Block()` 重新定位块，再删原段。
2. **表格块不能直接往 `table.Range.Start` 里写**：赋值会落进**第一个单元格**
   （实测：`Body paragraph C1…` 出现在 `cell 1,1 value 12345` 上面）。
3. 也不能"删掉表格前那段的段末标记"：把段落标记一起替换 = 删标记 → **表格被并进该段**，文字又进了单元格。
   最终做法：先在该段段末标记**前**插一个 `"\r"`（`InsertBefore("\r")`，与 `insert` 命令同一个技巧），
   得到表格前的空段，再往这段的**文本位置**（折叠 range）赋 `FormattedText`；
   动手前用 `Range.Information(12)` 探测插入点是否落在表内，落在表内就撤销 `"\r"` 并放弃该块（写 note）。
   代价：表格前会留**一个空段**（视觉上是表格上方的空行）。
4. `-DryRun` 每个块只报**一处**（搬第一段），因为不实际搬动就无法准确推演后续；实跑会按 `maxMoves` 连搬，
   并且**实跑的 moves 数可能与 dryRun 不同**（搬完重排后留白变了）。

### 冒烟（3 页文档 → 实跑移了 2 段）

自造文档 `%TEMP%\woc-v130\gaps.docx`（420×400 pt、上下 40 pt、左右 50 pt；标题 + A1–A3 + 自造 PNG 图
`%TEMP%\woc-v130\fig.png`（200 pt 宽）+ 题注「图 1 …」+ B1–B4 + 10 行×3 列跨页表 + C1；
表先跑 `table -KeepTogether` 让它整块跳页）。

`-DryRun` 原文：

```json
{"dryRun":true,"moved":0,"blocks":2,"minGapPt":90,
 "moves":[{"block":1,"kind":"figure","para":5,"gapAtPara":5,"page":1,"gap":124.2,"text":"Body paragraph B1: ...","wouldMove":true},
          {"block":-1,"kind":"table","para":11,"gapAtPara":11,"page":3,"gap":210.7,"text":"Body paragraph C1: text after the table.","wouldMove":true}],
 "summary":"试跑：只报不动 - 2 paragraph(s) could be brought in front of 2 figure(s)/table(s) (minGapPt 90)"}
```

实跑：

```json
{"dryRun":false,"moved":2,"blocks":2,"pages":4,
 "moves":[{"block":1,"kind":"figure","para":5,"page":1,"gap":124.2,"text":"Body paragraph B1: ..."},
          {"block":-1,"kind":"table","para":11,"page":3,"gap":268.3,"text":"Body paragraph C1: text after the table."}],
 "summary":"把正文段提到图表之前，填上它们上一页底部的留白 - 2 paragraph(s) brought in front of 2 figure(s)/table(s); 搬完请目检：图/表是否还在原位、有没有越过标题"}
```

目检（图的路径都是实跑前后各渲一次）：

- 改前 `%TEMP%\woc-v130\gf-final-before\page-1.png`：标题 + A1–A3，下方一大片空白（约 124 pt），图整块在下一页；
- 改后 `%TEMP%\woc-v130\gf-final-after\page-1.png`：**B1 已经补在那片空白里**，标题仍在最上面（没被越过）；
- 改后同目录 `page-2.png`：PNG 图**还在原位**（第 2 页顶部），题注「图 1 …」仍在图下；
- 改后 `page-3.png`：B4 + **C1（从表后搬上来的正文段）**，表格整体在第 4 页；
- 表内第一格仍是 `cell 1,1 value 12345`（读取 `tables[].header` 验证）——没有被搬的正文污染。

---

## 4. 接线（三处，缺一不可）

1. `scripts/word.ps1`：`ValidateSet` 加 `fillgaps`；`param()` 加 `-DryRun`、`-MaxMoves 8`、`-MinGapPt 0`；
   `switch ($Command)` 加 `'fillgaps'` 组装 `dryRun/maxMoves/minGapPt`；`$paramDefaultValues` 补 `MaxMoves=8; MinGapPt=0`。
2. `scripts/helper/WpsHelper.cs`：`Dispatch` 加 `case "fillgaps": return FillGapsDoc(a);`，unknown-cmd 提示串同步。
3. `scripts/helper/word/State.cs`：`WRITE_COMMANDS` 加 `"fillgaps"`（`-DryRun` 也登记，读文档签名不写盘无害）。
4. 新文件 `scripts/helper/word/FillGaps.cs` 由 `word.ps1` 的 `Get-ChildItem *.cs` 自动收进编译，**不改 `publish.ps1`**。

编译：`& scripts\woc.ps1 word doctor`（源码哈希变 → 自动 csc 重编译）。

---

## 5. 未做 / 已知边界

- `fillgaps` 只搬**正文段**；图题/表题、含图段、表内段一律不搬。搬完表格前面可能留下一个空段（见 §3.3）。
- `-DryRun` 的 moves 是**预测**，实跑可能多于/少于它（重排后留白会变）。
- `table -KeepTogether` 现在失败会让整条命令 `ok:false`（不再部分成功却报成功）；同一批里其它表已做的格式不会回滚。
- 段落数是"搬一段 = 拷贝 + 删原段"，若原段是**文档最后一段**，删除时 Word 保留段末标记，
  会留一个空尾段（段数 +1），不影响渲染。
