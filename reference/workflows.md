# 典型工作流（详解）

> 本文件是 SKILL.md 里各工作流的完整版：命令、要点、实测坑。按需读，不必从头到尾。
> 顺序建议：定位（search/read）→ 改（revise/format/…）→ 核对（render）→ 交付（save/close/解密提醒）。

## 1. 读一遍并核对版式

```powershell
<skill>\scripts\word.ps1 status
<skill>\scripts\word.ps1 open -Doc "D:\报告\某报告.docx" -Show
<skill>\scripts\word.ps1 read -Doc "D:\报告\某报告.docx" -From 1 -Count 120
<skill>\scripts\word.ps1 render -Doc "D:\报告\某报告.docx" -Page 1 -To 3 -Out "D:\报告\_pages"
```

看图时按这几项过：表格是否出框/掉到下一页、图题表题是否贴着自己的图表、有没有整段空行或半空页、修订红字与批注气泡是否在预期段落。（`render` 用的是打印视图的页面图元，**不是**截图，窗口在后台也出图。）

## 2. 定位：先 search 再 read

```powershell
# 全文找词，直接拿段落号和上下文（不用一页页 read）
<skill>\scripts\word.ps1 search -Doc "D:\报告\某报告.docx" -Text "孔口高程" -Max 20
# 只在某个样式里找（如只搜标题）
<skill>\scripts\word.ps1 search -Doc "D:\报告\某报告.docx" -Text "概况" -Style "标题 1"
```

- 返回 `hits`（每条带 para/style/inTable/at/context）、`totalHits`、`truncated`；命中多时先看 context 区分，再 `read -From <para> -Count 1` 取全文。
- 命中文本是**接受视图**（与 read/diff 同口径），据此写 `revise` 的 `expect` 不会错位。
- 英文默认不区分大小写，要区分加 `-Case`。

### 用户说"改我选中的这段"：`selection`

```powershell
<skill>\scripts\word.ps1 selection -Doc "D:\报告\某报告.docx"
```

- 返回选中文本 + 覆盖的段落号（`paras`/`fromPara`/`toPara`/`style`/`inTable`），光标未选时报光标所在段（length 0）。
- 拿到段号后照常干活：改文字用 `revise -Items`（`expect` 取选中那段的**开头**，不是选中的片段）、排版用 `format -From <fromPara> -To <toPara>`、贴批注用 `comment -Para <fromPara>`。
- 选中在**别的文档窗口**里时返回 `NOT_OPEN`（本 skill 只驱动 `-Doc` 指的那份）——提醒用户点到目标文档窗口里再选。
- 只读不改；用户选区/光标不会被移动。
- **选中在表格里**（用户说"改这个表"）：额外返回 `table`（表号）、`row`/`col`（光标所在行列）——直接喂给 `table -Table n`（改框线/字号/表头重复）、`cell -Table n -Row r -Col c`（改单元格）、`caption -Para <fromPara>`（补表题）。文档里有多个表时这是唯一不用数的定位法（`read` 的表格清单只有 header 预览，不带表号与位置的对应说明）。
- **选区/光标覆盖图片**（用户说"改这张图"）：额外返回 `images`（图号+宽高）——直接喂给 `picture -Image n`（调宽高/裁剪）。注意只覆盖**嵌入式**图片（`picture` 系列能操作的那种）；浮动图形暂不可寻址。

### 批注与分节：对象清单怎么拿

- **批注**：`comment -Action list` 列出每条的编号/作者/锚文（贴在哪段文字）/内容——用户说"删掉第 3 条批注""把那条关于防洪标准的批注改改"时先 list 再 `-Delete -Index n`（改内容=删了按原锚重加）。`revisions` 本身就带 list，批注此前没有，v0.5 补齐。
- **分节**：`page -Doc <路径>`（不带任何选项）= 查询模式，返回每节的纸张/方向/边距（厘米）/页眉页脚文字——用户说"只改第二节的页边距/横页那节"时先查再改；文档到底几节、哪节是横向，一眼定。

## 3. 逐段润色（`replace` 优先）

1. `search` / `read -From N -Count 40` 分批拿段落编号与原文——编号是唯一锚点。
2. 造 items：能 `replace` 就 `replace`（只产生字符级删/插，最贴近润色）；**只有整段重构才用 `set_text`**（会产生整段删+插，用户不好逐字看）。
3. items 写成 UTF-8 的 JSON 文件，`revise` 一次提交。
4. `render` 抽 1–2 页看修订落点（改动后 WPS 视图会跳回光标，别指望"跟着滚"）。
5. 用户看过再 `save`；保存后按铁律 5 提醒密文。

## 4. 校审批注（`comment` + `render` 复核）

1. `search`/`read` 定位问题段，记下编号与该段开头文字。
2. 每条批注写一个 UTF-8 的 txt（`-TextFile`，避免命令行引号转义把中文弄坏），`-Anchor` 指到具体词句可让批注贴在那几个字上。
3. 批注正文写"问题 + 依据 + 建议"，别写"见上"。作者固定为当前登录用户，要区分来源就在正文里署名（如"（AI 校对）"）。
4. `render` 复核气泡是否落在正确段落（页面边上，跨页易错位），再 `save`。

用 `audit` 拿机械检查候选清单的完整流程（audit → read → comment → render 四步）见 SKILL.md 命令速查与铁律 0；`audit` 只报**能机械判定**的（单位写法、重复词、括号、空格、标题跳级、图表缺题注），语义类问题（强制性条文、数据一致性、结论与图表不符）靠自己 `read` 之后判断——**别指望 audit 替你做判断**。

## 5. 统一字体（`font`）

```powershell
<skill>\scripts\word.ps1 font -Doc "D:\报告\某报告.docx" -Name 仿宋
```

- 覆盖**正文（含表格）、各节页眉页脚、浮动文本框**；中英文与符号一起改（`NameFarEast`/`Name`/`NameAscii`/`NameOther` 都设）。
- 只改**现有文字**：以后新打的字仍随段落样式的字体。要让后续输入也变，得改样式本身（本命令暂不做，用 WPS 的"修改样式"更稳妥）。
- 默认**不进修订**——整篇字体改动若逐字留痕会有成千上万条格式修订，把文字修订淹没；整批包在一个撤销记录里，**按一次 Ctrl+Z 全部撤回**（Agent 自己撤用 `undo -Count 1`）。
- 返回 `ranges`（每个区段的区段名、字符数、成败）与 `readBack`（回读正文的东亚字体名，应等于你给的字体名）。回读为空/不符说明没落上，先别保存。
- 字体名用 Windows 显示名（`仿宋`、`黑体`、`宋体`、`楷体`）；英文名（`FangSong`）在部分环境不认。

## 6. 一键排版（`format` + `style` + `page`）

```powershell
<skill>\scripts\word.ps1 status
# 全文正文：仿宋小四、首行缩进 2 字符、1.5 倍行距、两端对齐
<skill>\scripts\word.ps1 format -Doc "D:\报告\某报告.docx" -Font 仿宋 -Size 12 -FirstLineIndent 2 -LineSpacing 1.5 -Align justify
# 一级标题：黑体三号、加粗、段前后距、左对齐、不分首行缩进、段前分页
<skill>\scripts\word.ps1 style -Doc "D:\报告\某报告.docx" -Style "标题 1" -Font 黑体 -Size 16 -Bold -Align left -FirstLineIndent 0 -SpaceBefore "12pt" -SpaceAfter "6pt"
# 页面：A4、页边距、页脚居中页码
<skill>\scripts\word.ps1 page  -Doc "D:\报告\某报告.docx" -Paper A4 -Top 2.54 -Bottom 2.54 -Left 3.17 -Right 2.54 -PageNumber center
# 复核
<skill>\scripts\word.ps1 render -Doc "D:\报告\某报告.docx" -Page 1 -To 3 -Out "D:\报告\_pages"
```

要点：
- **先 `style` 后 `format`**：样式管"以后新写的字"，`format` 管"现在已经写好的字"，两者一起上才不返工。
- `format` 的三种范围：`-Para N`（一段）、`-From N -To M`（区间）、`-Style "标题 2"`（该样式全部段落）；**不给范围 = 全文**。
- 整篇排版不要开 `-Track`：会生成海量格式修订。排版类改动默认不进修订，一次 Ctrl+Z（或 `undo -Count 1`）全撤。
- 排版前后各 `render` 一次，重点看：标题是否掉行、首行缩进是否一致、表格是否被挤到下一页、页码是否只有一节有。

## 7. 表格、题注、图片（`table` / `caption` / `picture`）

```powershell
# 建表：数据走 JSON 文件（UTF-8），三线表 + 表头跨页重复 + 表题
<skill>\scripts\word.ps1 table -Doc "D:\报告\某报告.docx" -Para 120 -Data "D:\报告\表1.json" -Borders three-line -HeaderRow -Caption "主要水文特征值"
# 改已有表：全框线、按窗口自适应、行不跨页
<skill>\scripts\word.ps1 table -Doc "D:\报告\某报告.docx" -Table 1 -Borders grid -Autofit window -KeepTogether
# 给已有表补题注（域，自动编号）
<skill>\scripts\word.ps1 caption -Doc "D:\报告\某报告.docx" -Para 120 -Label 表 -Title "主要水文特征值" -Above
# 插图（带图题，图题在下）
<skill>\scripts\word.ps1 picture -Doc "D:\报告\某报告.docx" -Path "D:\报告\图1.png" -Para 90 -Width "12cm" -Caption "工程位置示意"
```

- `表1.json` 形如 `{"data":[["断面名称","集水面积(km²)"],["白沙","1560"]]}`；也接受不带外层对象的裸数组。
- `-Borders three-line` = 论文三线表（上下粗、表头下细、其余无线）；`grid` 全框线（0.5pt）；`thick` = **外框 1.5pt 粗线 + 内线 0.5pt 细线**（中文报告"粗边框"的标准样式）；`none` 无线。
- 题注是 **SEQ 域**（`表 1`、`图 1` 自动编号，增删后重排），不要手打编号——手打的不会重排。
- 表题在**表上方**、图题在**图下方**，这是规范；`caption` 有 `-Above` 控制，`table -Caption` / `picture -Caption` 已按此默认。
- 图片宽度用长度（`"12cm"` / `"340pt"`）；插图后**必须 `render` 看**是否出框、是否被压扁。

## 8. 目录与引用（`toc` / `reference` / `view`）

```powershell
# 文首插目录（一级到三级），自动刷新
<skill>\scripts\word.ps1 toc -Doc "D:\报告\某报告.docx" -Title 目录 -Levels 3
# 参考文献列表（JSON 字符串数组）
<skill>\scripts\word.ps1 reference -Doc "D:\报告\某报告.docx" -Action list -Items "D:\报告\refs.json" -Heading 参考文献
# 在正文某段末加引用上标 [1,2,5]
<skill>\scripts\word.ps1 reference -Doc "D:\报告\某报告.docx" -Action cite -Para 208 -Numbers "1,2,5"
# 让 WPS 窗口滚到第 208 段
<skill>\scripts\word.ps1 view -Doc "D:\报告\某报告.docx" -Para 208
```

- 目录**依赖标题样式**：文档里没有「标题 1/2/3」的段落，目录就是空的——先 `style` 或在 WPS 里给标题套样式。
- 目录是域，插完要刷新；`toc` 已内置 `Update()`。改完正文后再跑一次 `toc` 或直接 `fields` 即可更新页码。
- `view` 只滚窗口、不改文档，是"跟随修改位置"能做到的最大程度（不是插件那种跟随滚动）。
- 拉到最后要提醒：**目录、页码、域在 WPS 里刷新过再交给用户**——用 `fields` 一条命令刷新全部域，比逐项找更省事。

## 9. 刷新域 / 改单元格 / 干净插入（`fields` / `cell` / `insert`）

```powershell
# 改完正文、存盘之前：刷新页码、目录、题注、引用
<skill>\scripts\word.ps1 fields -Doc "D:\报告\某报告.docx"
# 改一个表格单元格（带格式）
<skill>\scripts\word.ps1 cell -Doc "D:\报告\某报告.docx" -Table 1 -Row 1 -Col 2 -Text "集水面积(km²)" -Font 宋体 -Size 10.5 -Align center
# 一次插入多段（干净写入，不是修订），带样式与格式
<skill>\scripts\word.ps1 insert -Doc "D:\报告\某报告.docx" -Para 40 -Where after -Items "D:\报告\新增.json"
```

- `insert` 的 items：`[{"text":"第一段","style":"正文","font":"仿宋","size":12,"firstLineIndent":2}, ...]`；`-Text` 可直接插一段。
- **插入位置会受表格影响**：锚点段落紧挨表格（或本身就是单元格）时，新段落会落进表格单元格里——返回里 `inTable: true` 就是提示；要插在表格前面，改用 `-Where before` 指向表格之后的那一段，或改完再 `read` 确认。
- `insert` 默认**不进修订**（写新章节、填模板单元格用），要留痕加 `-Track`。
- `cell` 的 `-Text ""` 清空单元格；合并单元格会报 `BAD_ARGS`（WPS 的 `Cell(r,c)` 取不到）。
- 这三条都不保存；`fields` 只刷新显示，不改文字。

## 10. 和人接力（`snapshot` / `diff` / `revisions` / `handoff`）——本 skill 的核心用法

Agent 接管一段 → 人裁决/手改 → Agent 读人的结果继续。四个动作：

```powershell
# ① 开工前先拍快照（记下"我接手时文档长什么样"）
<skill>\scripts\word.ps1 snapshot -Doc "D:\报告\某报告.docx" -Tag round1
# ② 干活……（revise / format / table …）
# ③ 人介入之后：看他改了什么、还剩多少修订没裁决
<skill>\scripts\word.ps1 diff      -Doc "D:\报告\某报告.docx" -Since round1
<skill>\scripts\word.ps1 revisions -Doc "D:\报告\某报告.docx"
# ④ 收尾：写一份人看得懂的交接单在文档旁边
<skill>\scripts\word.ps1 handoff -Doc "D:\报告\某报告.docx" -Note "本轮润色 12 段，其中 3 段改了数据表述，待你确认"
```

- **`snapshot`**：存每段文本到 `%LOCALAPPDATA%\wps-word-cooperate\snapshots\`（**不写进用户文档目录**）；`-Tag` 起名，`last` 指最新一份。**跨机器/交给别人接力**时加 `-Out <路径>`（如文档旁边 `<名>.snapshot-round1.json`），把快照随文档一起带走；diff 时用 `-Since <该路径>`（带 `-Out` 的快照不参与 `last`）。
- **`diff`**：和快照比，给出 `changed`（某段 before→after）/ `added` / `removed` / `same`。**人改的和 Agent 改的都在里面**。
- **`diff` 比的是"接受视图"**（WPS 的 `Range.Text` 就是接受后的样子：插入计入、删除不计）。所以：人**接受**了 Agent 的修订 → 文本本来如此，`diff` 看不到变化（剩余量看 `revisions`）；人**拒绝**或手改 → `diff` 报成 `changed`。判断"还有多少待裁决"永远看 `revisions`，判断"文本变成什么样"看 `diff`。
- **`revisions`**：列出**未处理**的修订（类型/作者/时间/所在段/文本）。已被人接受或拒绝的不在列表里，所以"列表"天然就是"待办"。
  - `-Action accept -Index N` / `-Action reject -Index N` 处理单条；`-All` 处理全部。
  - **纪律**：`-All` 只在用户明确说"都接受/都拒绝"时才用；否则一律让人自己在 WPS 里裁决。
- **`handoff`**：在文档旁生成 `<文档名>.handoff.md`——文档状态、待处理修订/批注数、这一轮做了什么（你传的 `-Note`）、最近 30 条操作记录、最新快照路径、"接手时建议先做"。下一轮 Agent 或同事直接读这份就能接上。
- **`changedOutside`**：任何写命令的结果里若出现这个字段（还有 `changedOutsideHint`），说明**在你上次写入之后，文档被人动过**。此时不要闷头继续：先 `diff` 看变了什么、`revisions` 看还剩什么，再决定。

**两套接力剧本**

*批次式（A）*：`snapshot` → 批量改（修订）→ 请人处理 → `diff` + `revisions` 读结果 → 第二轮 → `save` → `handoff`。
原则：**一轮只做一类事**（润色 / 排版 / 校审分开），人处理完再开下一轮；每轮开头都 `snapshot`。

*并发式（B）*：人和你同时在看同一份文档。
原则：① 写之前先看 `changedOutside`；② `revise` 的 `expect` 报 `ANCHOR_MISSING` 时，先 `read` 那一段——**很可能是人改了它**，不要当成编号漂移盲目全量重读；③ 不替人裁决修订；④ 人的光标位置别去动（不要 `Select`，`view` 默认也不动光标）。

## 11. 排版实测坑（照做少返工）

1. **`table -Para` 必须指向正文段**，不能指到表格内部：`read` 里带 `inTable: true` 的段落会报 `BAD_ARGS`（要往单元格里塞表就加 `nested: true`，且本 WPS 的嵌套表时灵时不灵，报错时改用"把表放在表格之后"）。
2. **题注是段落在前、表格在后**：`table -Caption` 会自动把「表 N 标题」放在表格上方一段；`picture -Caption` 放在图片下方。别再手打「表 1」。
3. **图片裁剪参数是百分比 0–100**（`-CropLeft 25` = 裁掉左边 25%）；WPS 内部的 `Crop*` 单位是"原图 100% 尺寸的磅"，换算已在命令里处理，别自己再换一次。
4. **`toc` 默认插在真正的文首**（封面之前）。要放在封面/扉页之后，给它 `-Para <正文第一段的编号>`；插之前先确认文档里有「标题 1/2/3」样式，否则目录会是空的（返回 `entries:0` + 提示）。
5. **加粗在渲染图里可能看不出来**：WPS 的 EMF 页面图对 `Bold` 的呈现不可靠（属性、docx 里的 `w:b` 都是对的）。别因为截图里"看着没粗"就反复重设；要看加粗去 WPS 窗口里看，或检查 `read` 的样式。
6. **表格跨页**：`-HeaderRow`（表头重复）+ `-KeepTogether`（行不跨页）一起用；长表就是会跨页，别硬压。
7. **先看文档原来怎么编页码**：中文报告常在**页眉**放页码（奇偶页还可能不同）。`page -PageNumber center` 是往**页脚**插 PAGE 域，文档本来就有页眉页码时会变成两套。动手前先 `status`/`render` 看一眼，或直接不传 `-PageNumber`。
8. **紧挨批注的文字，位置会差 1 个字符**：WPS 里批注锚点占一个位置、却不出现在 `Range.Text` 里，所以对着批注旁边的字做 `replace` 会错位（实测把"对无实测"改成了"对无实测测"）。现在 `replace` 会**先读回目标片段校验**，位置不对时在 ±3 字符内自动纠正，纠正不了就报 `ANCHOR_MISSING` 而**不会写坏**。确实要改那段：`comment -Delete -Index n` 摘掉批注 → 改 → 再 `comment` 按原词补回。
9. **排版类命令都不进修订**（`format`/`style`/`page`/`table`/`caption`/`picture`），整批一次 Ctrl+Z（或 `undo`）可撤；`revise`/`comment` 才是走修订的那两个。
10. **`undo` 是全局撤销、不分作者**：人在你之后改过也会被一起撤掉。只在**自己刚写完、且结果里没有 `changedOutside`** 时用；拿不准就别撤，让人在 WPS 里 Ctrl+Z。

---

## 12. 批量整理表格（2026-10-25 实战，饶河报告 85 张表）

中文报告最常见的四类表格毛病，一条 `table -Action scan` 先体检再动手。**实测提醒：不要凭"中文报告都有这个毛病"预设大面积命中**——那份 85 张表的报告里，题名做进第 1 行的只有 1 张。

```powershell
office.ps1 word table -Doc $DOC -Action scan
```

返回 `summary`（`tables` / `multipage` / `multipageWithoutHeaderRepeat` / `captionInFirstRow` / `bordersNotThick`）+ 每张表一行。只读，不改文档、不进撤销栈，只读打开的文档也能跑。

### 12.1 题名做进了第 1 行（`captionInFirstRow: true`）

```powershell
office.ps1 word table -Doc $DOC -Action caption -Table 1 -Borders thick
```

一条命令完成三件事，**顺序由命令内部保证**（顺序错了会丢题名或白干）：

1. 在表格**前一个段落的段末标记之前**插入「CR + 题名文本」——不能用 `InsertParagraphAfter`，那会落进表格第一格；
2. 按选中整行 → `Selection.Rows.Delete()` 删掉题名行（`Rows.Item(n).Delete()` 在 WPS 里报 `0x800A1767`）；
3. 最后套边框（先删行会重算首行边框，先套等于白套）。

返回 `caption` / `newPara` / `rowsBefore→rowsAfter` / `firstRowText` / `borderRead`。删完一定要看 `firstRowText` 变成真表头（如 `序号`）。

### 12.2 跨页却没重复表头

```powershell
office.ps1 word table -Doc $DOC -Action header-repeat          # 全文档，迭代到稳定
office.ps1 word table -Doc $DOC -Action header-repeat -Table 6 -Row 2   # 强制 2 行表头
```

**必须迭代**：设了重复会让续页多出表头行 → 分页位移 → 原本一页放得下的表变成跨页 → 又冒出新的待补表。实测 22 张 → 又冒 6 张 → 稳定，共 3 轮；命令自动跑到稳定，返回 `passesUsed` 与 `multipageStillMissing`（应为 0，非 0 就加 `-Passes`）。

表头行数自动判断：第 1 行单元格数 < 列数 ⇒ 两级表头 ⇒ 2 行。**仍要抽查**，若第 2 行其实是数据行，重复出去会把数据抄到每页。

顺序：**先提取题名，再设表头重复**，否则重复的是题名行。

### 12.3 全文档边框加粗

```powershell
office.ps1 word table -Doc $DOC -Borders thick        # 不带 -Table = 所有表
```

一次 COM 遍历 + 单个撤销记录（85 张表约十几秒）。**不要**逐张调（85 次进程附着 ≈ 4–5 分钟）。回读：四个外边 `[1,12]`（1.5 pt）、两条内线 `[1,4]`（0.5 pt）。

批量套边框前先确认没有"用表格做的排版/图片框"——`scan` 里 `bordersNotThick` 会把它们一并算进来；如有，按表号排除、不要一刀切。

### 12.4 删行

```powershell
office.ps1 word table -Doc $DOC -Action rows -Table 2 -Row 3
```

返回 `rowsBefore/rowsAfter` 与删后的首行文本。删行会连同文字一起删（等同退格），删前想留内容先 `cell` 读出来。

### 12.5 验证闭环

1. `scan` 复跑：`captionInFirstRow` 归零、`multipageWithoutHeaderRepeat` 归零、`bordersNotThick` 归零；
2. `render` 目标页：题名在框外居中、外框明显粗于内线、续页顶部有列名、数据条数没变；
3. `pages` 基本不变（只有改页面结构时才该变）；
4. 格式类改动不进修订，**save 之前** `render` 让用户看过。

### 12.6 页码断号（另一类坑，COM 做不到）

现象：正文 `1,2,3` 之后某些物理页反复跳回 `1`。根因是 `sectPr` 里带了 `<w:pgNumType w:start="1"/>`。

- `page -Doc $DOC` 的查询模式会报每节 `pageNumbers`（`startNumberRaw` / `restartAtSectionRaw` / `firstDisplayedPage`），断号一眼看出；
- `page -PageNumberStart none` 会尝试清掉"本节重新编号"，但 **WPS 常常做不到**（`RestartNumberingAtSection` 读回 `null`；写 `StartingNumber = 0` 会理解成"从 0 开始"，写 `-1` 会把分页搞乱）——命令会**明说失败**，别硬试；
- 可行路线：关闭文档（不保存）→ 改 `word/document.xml` 删掉多余节的 `pgNumType@w:start`（保留第一节的 `start="1"`）→ 重新 `open` → 导出 PDF 逐页取页码验证连续 → `fields` 刷目录 → `save`。导出 PDF 落盘即密文，要先用 `decrypt-office-pdf` 解密再读。
---

## 13. 单位上标批量修（2026-10-25 实测，300 页 / 13498 段）

**问题**：报告里 `2248m3/s`、`15300km2`、`1.2亿m3` 的指数"掉下来了"（普通数字）。而同一篇文档里**有些人是对的**，且"对"有两种形态：

| 形态 | 文本层看到 | 怎么判 |
|---|---|---|
| Unicode 上标 | `m³`、`km²`（单字符 ²/³） | 文本本身就是 ²/³，`unit` 的 `unicodeOk` 统计 |
| run 级上标 | 文本层仍是 `m3`、`km2`，但数字带 `Font.Superscript` | 只有 Find 的格式过滤能看见；`unit -Apply` 会自动跳过 |
| **掉下来（要修）** | 普通 `3`/`2`，无任何格式 | 格式过滤 `Font.Superscript = 0` 命中 |

```powershell
office.ps1 word unit -Doc $DOC                      # 体检
office.ps1 word unit -Doc $DOC -Apply               # 修（一次撤销记录）
office.ps1 word unit -Doc $DOC -Apply -Units "s2"          # 追加真实使用的单位（笔误别往这里加）
```

### 13.1 为什么不是"查找替换一把梭"

盲目 `ReplaceAll` 会把英文词和非单位数字一起改掉——实测在测试文档里 `item3 → ite³`、`com2 → co²`、`m23 → m²3`、`km20 → km²0`。`unit` 的实现是**逐处 + 上下文守卫**：

1. **格式过滤**：`Find.Font.Superscript = 0`，已经正确的 run 级上标永不出现；
2. **前一个字符不能是英文字母**：`item3` 跳过；`km2` 里的 `m2` 也跳过——因为**按最长单位优先**处理，`km2` 会由它自己那一趟整体修掉；
3. **后一个字符不能是数字**：`m23`、`km20` 跳过（说明数字还在继续）；
4. **必须"挂在数值上"才算单位**：前一个字符是数字、`万`、`亿`、**汉字**、半/全角左括号，或者形如"数字 空格 单位"。否则一律**只报不改**（见 13.4）。

`Execute` 的第 9 个参数（Format）必须是 **true**，否则字体过滤会被丢弃——这是个坑：丢了这个参数，`ReplaceAll` 会把 format-ok 的也一起改掉（实测 598 处被误改）。

### 13.2 实测数字（真实报告副本，300 页 / 36 MB）

| 指标 | 结果 |
|---|---|
| 打开 + 体检 + 修复 全流程 | **~21–24 s**（其中 WPS 打开 36 MB 文档 ~8 s） |
| 文本候选（`plainDigitHits`） | 156 |
| 已 Unicode | 147 |
| **实际修复** | **9–16**（取决于该稿模板示例的占比；本轮最终口径 9 处，其余是模板变量/表格标题被守卫拦下） |
| 只报不改 | `fragments` 4（`km2 = 1`、`<m1，m2，m3>` 等模板变量）、`looseFragments` 2（`n2`、`n3`）、`suspects` 1（`km3` 笔误） |

对比：合成测试文档 34 页 / 1818 候选 / 1220 处损坏 → 全修、1220 全中、format-ok 598 处一处未动、8 个诱饵词全部保留，耗时 21.6 s。**坏得越多越慢（≈17 ms/处），但 300 页文档通常只有十几到几十处，所以通常是秒级。**

### 13.4 "可疑单位碎片"三类报告（只报不改）

| 类别 | 命中长什么样 | 为什么不能自动改 |
|---|---|---|
| `fragments` | `km2 = 1`、`<m1，m2，m3， n1，n2，n3>`、`径流模数（L/s.km2）` | 前两者是**模板/脚本示例里的变量名**；自动加 ³ 会改坏文档（**已实测发生过**：v1.4 无守卫时把那一页的 `m2/m3/km2` 全改成了上标） |
| `looseFragments` | `n2`、`n3`、`k2` | 缺单位基名的形态，可能是 `m2`/`m3` 打错，也可能是变量名——需人判断 |
| `suspects` | `km3`、`km4`、`m4`、`hm3` | 疑似笔误（面积表里"立方公里"几乎必是 `km2`）；自动加 ³ 会让笔误看起来像有意为之 |

判据一句话：**上标是排版，改错单位是内容错误；宁可多报一次，不要改错一处。**

### 13.3 纪律

- `unit` 是**内容改动**（把 `3` 换成 `³`），不是格式改动；默认**不进修订**，整批算**一个撤销记录**，配合 `snapshot` 用。要逐条可审时加 `-Track`（会生成大量修订对，慎用）。
- **先跑不带 `-Apply` 的体检**：`skippedSamples` 里若有疑似真单位，用 `-Units` 补；若是英文词就不管。
- 修完 `render` 抽 1–2 页看（重点看行高有没有被撑开、上下标是否与上下文一致），确认后再 `save`。
- 单位表**只放真实使用**的单位（metric 面积/体积系列：`km2/hm2/cm2/mm2/dm2/um2/nm2/m3/m2`）；`m/s2`（加速度）之类按需 `-Units` 追加，必须写全基名（如 `s2`）。
- **疑似笔误只报不改**：`km3`、`km4`、`m4`、`hm3` 这类会被列进 `suspects` 并附上下文。经验：面积/库容报告里出现"立方公里"，多半是 `km2` 打错；自动加 ³ 会让笔误看起来像有意为之，比不改更糟。确认后手工改数字，或用 `-Units "km3"` 明确表示"这确实是立方单位"。
- 表格、表格单元格、正文都在 `doc.Content` 内，都会被处理；**页眉/页脚/脚注不在范围内**（水利率报告里单位极少出现在页眉）。