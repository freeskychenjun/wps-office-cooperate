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
