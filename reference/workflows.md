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

## 12. 表格整备（v1.33：`plan` + `normalize` 一条命令，取代下面的手工流程）

> 起因：2026-11-03 中午，一份 29 页 / 14 表的水利自检报告做「题注编号修正 + 表名外提 + 所有表格加粗边框」。任务本身 3 件事，旧做法用了 **127 次工具调用 / 88.6 分钟**（第 2 轮 88.6 min，其中 90% 在等模型、10% 在跑命令）。参考 `design-report-format`（一条 exe 内部跑完 scan→format→table）改造成本节的写法后，同样工作实测 **8 条命令 / 约 40 s**（plan 6.5 s + dry-run 4.1 s + normalize 15.4 s + plan-after 5.3 s + render 2.8–3.7 s×3）；其中表名外提只对分页符前的表生效，分节符前的表按平台限制报 `needsHuman`。
>
> 关键结论：慢的根因不是 COM，而是**步数 × 每步上下文**——上下文 21.8 万 token 时每步首字延迟 5–60 s；127 次调用里 **56 次（42.2 分钟）是试错与绕路**，43 步（29 分钟）本可由一条命令替代。

### 12.0 四步流程

```powershell
# 1) 只读侦察：一条命令给全（表结构/题注/边框/列宽/风险/编号提议/正文引用）
office.ps1 word table -Doc $DOC -Action plan
# 2) （可选）人确认编号映射：把 plan 里 renumber.approved 改成 true 再传回
office.ps1 word table -Doc $DOC -Action normalize -Plan approved-plan.json -DryRun   # 先看要改哪些
# 3) 一次改到位：题注外提（尽力而为，见 12.3；分节符在表前会报 needsHuman）→ 空白统一 → 重编号 → 加粗边框（列宽快照/恢复）→ 跨页表头 → 自检
office.ps1 word table -Doc $DOC -Action normalize -Plan approved-plan.json
# 4) 表级复核：只渲该表所跨页（不再整页翻）
office.ps1 word render -Doc $DOC -Table 9 -Out .\check -Width 1300
```

`normalize` **默认只改 plan 判定未达标的表**；整条命令一个撤销单元、逐表隔离、**不 save**（保存永远是人裁决的最后一步）。

### 12.1 `plan` 的返回（只读，一次 6–8 s）

- `tables[]`（每表）：`para` / `pages` / `rows` / `columns` / `caption{pos,para,page,text}` / `split` / `headerRepeated` / `borders[outer,inner]` / `tblStyle`（空 = 无样式）/ `tblW`（`pct/5019`、`dxa/8393`、`auto/0`）/ `gridCols[]`（pt）/ `fontSize` / `risks[]`。
- `risks[].flags`：`captionInFirstRow`（题名在表内）、`bordersNotThick`、`breakBefore`（表前有分页/分节符 → 题注外提要特殊处理）、`noStyle`、`noBorders`（**写边框会触发 WPS 重建布局**）、`pctWidths`、`gridOverflow`（列宽合计超版心）、`multipageWithoutHeaderRepeat`。
- `captions[]` / `references[]` / `issues[]`：与 `captions` 命令同一解析器（两命令永不会打架）。
- `renumber`：`items[{table,para,from,to,unique,basis}]` + `refs[{number,to,ambiguous}]` + `duplicates`/`ambiguous`。**这是"按出现顺序 + 引用指向"的候选映射，不是裁决**——重号时 `unique=false`，人/agent 确认后才能 `approved:true`。
- `summary`：14 个计数（`captionInFirstRow`、`bordersNotThick`、`noStyle`、`noBorders`、`pctWidths`、`breakBefore`、`multipage…`、`gridOverflow`、`textWidth` 等）。

**为什么要它**：这些字段以前只能靠 `scan`（只有 11 个键）+ 自己解密读 XML + 逐段 `read` 拼出来；现场"查不出列宽为什么崩"就是这样烧掉 37 分钟的。

### 12.2 `normalize` 的行为与参数

| 参数 | 说明 |
|---|---|
| `-Action normalize` | 默认全套；只处理未达标表 |
| `-Table n` | 只做这一张（试跑用；仍会自动判风险） |
| `-DryRun` | 只报将改什么，不写文档（只读文档也能跑） |
| `-NoCaptions` / `-NoBorders` / `-NoHeaderRepeat` | 关掉对应阶段 |
| `-Gap 2` | **只统一「编号与题名」之间**的空间（默认 2 个半角空格，`0` 关闭）；「标签与编号」（表/表2.1.1）保持原稿习惯，不擅改 |
| `-Borders thick` | 外框 1.5 pt + 内线 0.5 pt（当前唯一档；`grid`/`three-line` 走旧 `table -Borders`） |
| `-Fit keep\|window\|content` | 边框后列宽策略：保持原宽（默认）/ 按比例撑满版心 / AutoFit 内容 |
| `-Plan <file>` | 带上人确认过的 `renumber`（`approved:true` 才执行） |
| `-Verify` | 收尾复检（默认就带 verification 块）；失败项进 `needsHuman` |

执行顺序（命令内部保证）：**题注外提 → 题注空白规范化 → 可选重编号 → 加粗边框 → 跨页表头 → 自检**（先提题名再设表头，否则重复出去的是题名行；先套边框后删行会让首行边框重算）。

返回：`summary{changed,skipped,needsHuman,captionsMoved,bordersApplied,widthsRestored,headerRepeated,pageBefore/After,revisions/comments before/after}`、`verification{captionInFirstRow,bordersNotThick,multipageWithoutHeaderRepeat,captionPageMismatch,widthsChanged,pageDelta}`、`tables[]`（每表只回关键字段）、`failures[]`（单表失败不中断）。

### 12.3 两个已修的真缺陷（机制级）

1. **表名外提：WPS 上不做可靠承诺（2026-11-03 定调）**。能提的只有「普通段落后 / 手动分页符前」的表：`normalize` 只用 `before-para`（在表前一段的段末标记前插 CR），且**回读题注页码 == 表起始页**才接受（T1/T7 实测可提）。**表前是分节符时（如横向节首的表）无解**——分节符的段末标记紧邻表格，任何"在表前插段落"都会落回上一节，`normalize` 直接报 `needsHuman{captionMoved:"wps-section-start"}`，不重试；要彻底解决只能人工把分节符剪到题注之后，或在微软 Word 里用 `-CutPaste`。表前一段带**手动分页符**（`\f`）或**分节符**（`sectPr`）时，旧写法在"段末标记之前插 CR"会把题注插到断点**之前**。`normalize` 改为固定方法链 **range-insert → table-start → before-para**，每种都**回读题注页码 == 表起始页**才接受：
   - **before-para**：在表前一段的段末标记前插 `"\r" + 题名`，再回读页码；普通段落后/手动分页符前的表用它可成（T1/T7 实测）。
   - 我们**试过并删掉**的写法（留档避免重复踩）：`table-para`（`Table.Range.InsertParagraphBefore`）、`range-insert`（`doc.Range(tableStart,tableStart).InsertBefore`）、`table-start`（`doc.Range(tableStart-1,tableStart-1).InsertParagraphBefore`）——在 WPS 上分节符紧邻表格时全部把题注插进**上一节或表格第一格**，清理失败尝试还会毁表（曾把 18 列表段落数打掉 3、页数 +2），所以按「不挂不脏」原则删除。
   - **cut-paste**（`WordOpenXML` 存表 → 删表 → 节首写题名 → `InsertXML` 放回）在**微软 Word** 上可用；**WPS 12.1 上会把 Writer 进程转死（实测 100% CPU、COM 全部挂起、只能杀进程）**，所以默认只在真 Word 上自动启用，WPS 上必须显式 `-CutPaste` 才尝试。
   - 三法（+可选 cut-paste）都不过 → 该表回滚 + `needsHuman{captionMoved:"section-break"}`，其它表继续。**绝不为了一张表把整条命令或整个文档挂住。**
2. **加边框把列宽搞崩**（表 2.1.7，旧做法排查 37 分钟）：实测根因**不是**"WPS 自动挂默认样式"也不是"宽度被改成百分比"（磁盘上本来就有 `tblStyle` 且 `tblW/tcW` 就是 pct）。触发条件 = **该表 XML 里没有 `w:tblBorders` 元素**：第一次写边框迫使 WPS 补建边框描述并**重建该表布局**，`w:gridCol` 按"当前渲染宽度"重新推导（容量取整）。破坏力由该表有无 `tblStyle`、宽度是 pct 还是 dxa 决定——原稿那张表恰好无样式 + dxa 固定，直接崩；有 `tblBorders` 的 12 张纹丝不动。
   - 而 **COM 读不出列宽**：autofit 列 `Column.Width` 在 WPS 返回 `9999999`（`wdUndefined`）。所以快照走 `doc.Content.WordOpenXML`（整篇 0.16 s）：解析 `tblStyle/tblW/tblLayout/gridCol[]/首行 tcW/tblBorders`；写后比 `gridCol` 漂移（>2%），漂了就 `PreferredWidthType/PreferredWidth` + 逐列 `Width = dxa/20` 写回，再读一次 XML 复核。
   - **短路**：已达标的表跳过（实测全局改 14 张只有 2 张真变）。

### 12.4 验证闭环

1. `normalize` 的 `verification` 五项归零；`failures[]`/`needsHuman[]` 为空或已解释；
2. `plan` 复跑对比 before/after（题注页码、边框、`gridCols` 漂移 ≤2%）；
3. `render -Table n` 目检：题注在框外居中、外框明显粗于内线、续页顶部有列名、数字没被拆行；
4. 段落数变化只来自"删题名行 + 插题注段"；页数基本不变；修订数/批注数不变；
5. **最后才 `save`**，由人裁决。

### 12.5 旧做法（v1.32 及以前，仅"单点小活"还用它）

以下逐条命令仍保留，适合"就一张表 / 只提个题名"的小活；**多表批量一律走 12.0**，不要逐条拼（85 张表逐张调 ≈ 4–5 分钟且无法单表回退）。

### 12.6 题名做进了第 1 行（单表手工）

```powershell
office.ps1 word table -Doc $DOC -Action caption -Table 1 -Borders thick
```

一条命令完成三件事，**顺序由命令内部保证**（顺序错了会丢题名或白干）：

1. 在表格**前一个段落的段末标记之前**插入「CR + 题名文本」——不能用 `InsertParagraphAfter`，那会落进表格第一格；
2. 按选中整行 → `Selection.Rows.Delete()` 删掉题名行（`Rows.Item(n).Delete()` 在 WPS 里报 `0x800A1767`）；
3. 最后套边框（先删行会重算首行边框，先套等于白套）。

返回 `caption` / `newPara` / `rowsBefore→rowsAfter` / `firstRowText` / `borderRead`。删完一定要看 `firstRowText` 变成真表头（如 `序号`）。

### 12.7 跨页却没重复表头（单表手工）

```powershell
office.ps1 word table -Doc $DOC -Action header-repeat          # 全文档，迭代到稳定
office.ps1 word table -Doc $DOC -Action header-repeat -Table 6 -Row 2   # 强制 2 行表头
```

**必须迭代**：设了重复会让续页多出表头行 → 分页位移 → 原本一页放得下的表变成跨页 → 又冒出新的待补表。实测 22 张 → 又冒 6 张 → 稳定，共 3 轮；命令自动跑到稳定，返回 `passesUsed` 与 `multipageStillMissing`（应为 0，非 0 就加 `-Passes`）。

表头行数自动判断：第 1 行单元格数 < 列数 ⇒ 两级表头 ⇒ 2 行。**仍要抽查**，若第 2 行其实是数据行，重复出去会把数据抄到每页。

顺序：**先提取题名，再设表头重复**，否则重复的是题名行。

### 12.8 全文档边框加粗（单表/少量表手工）

```powershell
office.ps1 word table -Doc $DOC -Borders thick        # 不带 -Table = 所有表
```

一次 COM 遍历 + 单个撤销记录（85 张表约十几秒）。**不要**逐张调（85 次进程附着 ≈ 4–5 分钟）。回读：四个外边 `[1,12]`（1.5 pt）、两条内线 `[1,4]`（0.5 pt）。

批量套边框前先确认没有"用表格做的排版/图片框"——`scan` 里 `bordersNotThick` 会把它们一并算进来；如有，按表号排除、不要一刀切。**v1.33 起**：无 `w:tblBorders` 的表写边框会触发 WPS 重建布局（见 12.3），要保列宽就走 `normalize`。

### 12.9 删行

```powershell
office.ps1 word table -Doc $DOC -Action rows -Table 2 -Row 3
```

返回 `rowsBefore/rowsAfter` 与删后的首行文本。删行会连同文字一起删（等同退格），删前想留内容先 `cell` 读出来。

### 12.10 页码断号（另一类坑，COM 做不到）

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

### 13.4 "可疑单位碎片"与脚本段（只报不改）

| 类别 | 命中长什么样 | 为什么不能自动改 |
|---|---|---|
| `scriptPages` | 整页是公文模板/脚本说明：`[for] … [/for]`、`[File]`、`k~a = 1`、成串 `m1 = 1 / n2 = 1` | 那里的 `km2/m2/m3` 是**变量名**。实测该报告第 **256–258 页**，共 9 处命中全部跳过。判定：段落含 `[for]`/`[File]`/`k~a` 或 ≥2 条 `x = 1` |
| `fragments` | `km2 = 1`、`<m1，m2，m3， n1，n2，n3>`、`径流模数（L/s.km2）` | 前两者是变量名；自动加 ³ 会改坏文档（**v1.4 无守卫时真的发生过**） |
| `looseFragments` | `n2`、`n3`、`k2` | 缺单位基名，可能是 `m2`/`m3` 打错，也可能是变量名 |
| `suspects` | `km3`、`km4`、`m4`、`hm3` | 疑似笔误（面积表里"立方公里"几乎必是 `km2`）；自动加 ³ 会让笔误看起来像有意为之 |

页级跳过会误伤"同页既有示例又有真单位"的情况，所以实现上是**页号用于报告、段落特征用于判定**：段落不是脚本段就照修，但若该页在 `scriptPages` 里会额外进 `fixedOnScriptPage` 告警，供人复核。

判据一句话：**上标是排版，改错单位是内容错误；宁可多报一次，不要改错一处。**

### 13.3 纪律

- `unit` 是**内容改动**（把 `3` 换成 `³`），不是格式改动；默认**不进修订**，整批算**一个撤销记录**，配合 `snapshot` 用。要逐条可审时加 `-Track`（会生成大量修订对，慎用）。
- **先跑不带 `-Apply` 的体检**：`skippedSamples` 里若有疑似真单位，用 `-Units` 补；若是英文词就不管。
- 修完 `render` 抽 1–2 页看（重点看行高有没有被撑开、上下标是否与上下文一致），确认后再 `save`。
- 单位表**只放真实使用**的单位（metric 面积/体积系列：`km2/hm2/cm2/mm2/dm2/um2/nm2/m3/m2`）；`m/s2`（加速度）之类按需 `-Units` 追加，必须写全基名（如 `s2`）。
- **疑似笔误只报不改**：`km3`、`km4`、`m4`、`hm3` 这类会被列进 `suspects` 并附上下文。经验：面积/库容报告里出现"立方公里"，多半是 `km2` 打错；自动加 ³ 会让笔误看起来像有意为之，比不改更糟。确认后手工改数字，或用 `-Units "km3"` 明确表示"这确实是立方单位"。
- 表格、表格单元格、正文都在 `doc.Content` 内，都会被处理；**页眉/页脚/脚注不在范围内**（水利率报告里单位极少出现在页眉）。

## 14. 接管提示（`hint`，v1.11）

用户的原话是"当 agent 接管 wps 文档后，可否有比较友好的提示告诉用户？"——接管这一步原本是**不可见的**，所以固定成四步：

1. **`status`** —— 说清在哪个实例、哪份文档上动手。
2. **`open -Doc <路径>`** —— 自动把窗口标签改成 `〔AI接管中〕 报告.docx`（纯窗口状态，不写文档；WPS 12 的标签页从右截断，标记必须放前面——后缀在长文件名上会被整段吃掉）；返回值里 `hint.caption` / `hint.captionSet` 就是取证。只读文档自动标成 `〔AI接管中·只读〕`，因为写命令本来就会被 `READONLY` 拒绝。
3. **`hint -Doc <路径> -Toast`** —— 弹 6 s 托盘气泡："已接管《报告.docx》：我会把改动做成修订或批注，你随时可以接受或拒绝。"（`-Sound` 加提示音、`-HintTag` 换标签文字）。长任务中途报进度用 `hint -Text "正在核对第 5 节的表格" -Quiet`：表格侧状态栏会显示这行字，**Writer 侧 WPS 不渲染状态栏**（实测截图仍是页面/字数），Word 报进度只能靠气泡（别用 -Quiet）与标签页标记。
4. **要常驻进度就加 `-Tray`**（v1.12）—— `hint -Doc X -Text "正在核对第 3 节" -Quiet -Tray`：托盘出现橙色「AI」图标，悬停 `〔AI接管中〕 报告.docx · 正在核对第 3 节`，右键「结束协作：还原标题标记」。`finish`/`close`/`-Clear` 会自动移除，忘关也会 4 小时自灭。
   - 新图标默认在「^」溢出区（Windows 行为）：要告诉用户拖到任务栏一次；`hint` 返回里的 `tray.note` 就是这句话。
   - 用户中途关掉文档：图标约 10 s 内自己消失（RestartManager 判定文件已不被占用，实测 Word 9.6 s / Excel 8.6 s / PPT 7.4 s）；`save -As` 会重指路径。
   - PPT 侧尤其值得开——它没有可写标题、没有状态栏，托盘是常驻信号（v1.27 起还有每页的页面标记）。
5. **收尾** —— `finish` / `close` 自动还原标题与状态栏并移除托盘图标；漏了就 `hint -Clear`。`status` 的 `caption` / `hinted` 可随时核对。

- Word / Excel 的标记全是**窗口状态**：不进修订、不写文件、不改变 `saved`；`tests/word/check_hint.py` 专门断言"段落数不变、无新增修订"。
- **PPT 有两层标记（v1.26 / v1.27，v1.28 定策略）**：① 窗口标题（Win32 `SetWindowText` 打在真实窗口框 `PP12FrameClass` 上）——**只有任务栏 / Alt+Tab 看得见**，因为 WPS 演示的标签页是自绘的；② **每页右上角的「〔AI接管中〕」橙色药丸**（形状名 `woc-ai-takeover-mark`，`hint -SlideMark` 显式开启 / `-NoSlideMark` 移除）——唯一"界面里看得见"的通道，编辑视图和放映都显示。页面标记是**内容**：用户自己 Ctrl+S 会存进文件，所以 ①默认**不贴**（`hint` 零改动，`inAppHint` 提示怎么开）；②`save`/`export` 落盘前剥离、写完补回；③`hint -Clear`/`finish`/`close` 逐页删除；④**万一被 Ctrl+S 存进去，下一条命令自动修复**：helper 记着"贴标记那一刻的文件 mtime/size"，发现文件被保存过就 剥离→重存→贴回，返回 `badgeRepaired:true`（只读文件 `badgeDirty`、重存失败 `badgeRepairFailed`）。
- **修复路径的实测坑**：用户 Ctrl+S 之后，**本进程内存里的形状列表可能已经和文件不同步**（用裸 COM `Presentation.Save()` 复现：保存后 `read` 看不到徽章，但重开文件徽章确实在）。所以修复不能依赖"内存里还能找到徽章"——必须**照常 `Save()` 一次**让文件与内存一致（内存无徽章 → 文件被写成干净版），再重新贴回并更新记录。`tests/ppt/check_hint.py` 的 P6 就是这条链路的回归（裸 COM 保存 → 下一条命令 `badgeRepaired` → 重开文件确认干净）。
- **页面标记不能做在母版上**：实测母版形状在页面形状之下，封面那种满版大图会把它整片盖住（第一版踩过，渲染图上一个像素都看不见）⇒ 逐页贴、置顶。
- **读侧必须跳过页面标记**：`read`/`search`/`snapshot`/`diff`/`check`/`replace`/`outline` 全部按形状名排除，否则它会变成"正文新增文字"、"字号 10.9pt 过小"、"对比度 2.85:1"之类的假问题（实测未排除时 `check` 在 19 页稿上多报 37 条）。`read` 每页用 `takeoverBadge:true` 告知。
- **PPT 的 COM 标题仍然不可写**：`ActiveWindow.Caption` 只读、`Application.Caption` 写了不重绘、`StatusBar` 不存在，响应带 `captionSupported:false`；`titleMarked:true` 指的是 Win32 那一层，别对用户谎称"标题已标记"。
- 实测细节（哪个属性可写、Writer 的 `StatusBar` 只写不读且不渲染、表格侧 `ActiveWindow` 为 null）见 [word-quirks.md](word-quirks.md) 与 [excel-quirks.md](excel-quirks.md)。
- 托盘实现的两个坑：① daemon 用 `--tray <state.json>` 重开自身，状态文件是唯一 IPC，菜单「还原」调本 skill 的 `hint -Clear`（daemon 不懂 COM）；② spawn 必须 `UseShellExecute=true`，否则子进程继承调用方 stdout 管道，`subprocess.communicate`/PowerShell 管道会被图标的 4 小时寿命卡死（2026-10-30 实测踩到）。

### 14.1 租约：接管标记不可能比 agent 活得久（v1.35）

背景（2026-11-05 用户反馈）：agent 接管后**退出 AI 会话**，窗口标签仍长期显示「〔AI接管中〕」。原因是标记只有"显示"、没有生命周期——唯一会清它的是 agent 自己调 `finish` / `close` / `hint -Clear`，会话一断就无人清（现场实测：16:52 打开的 deck 到 17:30 还顶着前缀，`WppHelper.exe --serve` 仍常驻，但已无命令进来）。

现在标记是**租约**：

- 常驻 helper 启动时 + 每服务完一条命令，刷新 `%LOCALAPPDATA%\wps-office-cooperate\hint-lease-<app>.txt`；
- 等请求的空档 **≤5 s** 检查一次，超过 `WOC_HINT_LEASE_MS`（默认 **120 s**）没有新命令 ⇒ 还原所有仍挂着的标记（Word/Excel 标题与状态栏、PPT 标题前缀与页内药丸、托盘图标），三端同一套 `shared/HintLease.cs` + `PipeServer`；
- helper 退出（租约过期、`serve -Action stop`、最后一份文档关闭）也会**先还原再退出**；**v1.36 起租约过期后 helper 直接退出**（不再空挂 20 分钟），包装器在下一条命令前用 ping 探活、没了就重新拉起——所以"AI 走了"与"自动化客户端走了"现在同时发生；
- **业务命令前自愈隐藏窗口（v1.36）**：WPS 自己的关闭路径崩溃后可能留下 `Application.Visible=false`、文档子窗口全隐藏的状态（外框还在、内容全灰 = 用户说的"卡死"）。`HandleRequest` 在 `EnsureHint` 前调 `HealWindow()`：app 被隐藏且有文档 ⇒ `Visible=true` + 激活活动文档/窗口；**AI 下一条命令（甚至 `status`）即可恢复**；
- **陈旧 COM 引用自恢复（v1.36）**：`App()` 复用缓存前先 `StillAlive()` 探活（`Workbooks/Documents/Presentations.Count`），只有断开类 HRESULT 才丢弃重连；`close` 关最后一份文档 / `quit` 后 `_app = null`；
- **只清显示、不清意图**：Word/Excel 的 `hint-state-<app>.json`、PPT 的 `ppt-title-hint.json` 保留，下一条命令开头的 `EnsureHint` / `EnsureTitleMarker` 会自动补标 ⇒ 暂停超过 2 分钟的协作，回来后标记会自己回来；
- 真正结束一轮协作仍用 `finish` / `close` / `hint -Clear`（它们清意图，之后不再自动补标）；`open -NoHint` 同样清意图；
- 覆盖范围：`WOC_NO_DAEMON=1` 的一次性模式没有常驻进程跑定时器，仍只有显式收尾（这也是它只适合 A/B 的原因）；HealWindow 在一次性模式下同样生效；
- 验收：`tests/<app>/check_hint_lease.py`（租约调到 4 s，断言 标记 → 静默 → 消失 → helper 退出（`serve status` 的 `running:false`）→ 命令 → 回来 → `serve stop` → 消失 → `finish` 清意图）；`tests/excel/check_window_heal.py`（隐藏 app → `read` → 可见性恢复）。

## 15. 忙重试与剪贴板礼仪（v1.32 基础设施，用户看不见但会感觉到）

这两件事都**不需要你调任何新参数**，行为变化只体现在"以前会失败 / 以前会丢东西"的地方。

### 15.1 忙重试：`BUSY` 现在多半会自己过去

三个 helper 的 `Main` 第一行都注册了 OLE 消息过滤器（`helper/shared/MessageFilter.cs`）：WPS 拒绝调用（用户在单元格里打字、开着模态对话框、正在重绘）时，OLE 会按 **150 ms 一次重试、最多 6 s**，仍不行才把异常抛出来变成 `BUSY`。

对写命令的实际影响：

- 以前：用户在打字 → 命令**立刻** `BUSY`（或干等到 `TIMEOUT`）→ 只能重试。
- 现在：先自动等 ≤6 s；用户停手了就**直接成功**，返回里看不出区别。
- 仍报 `BUSY` 说明真的忙了 6 s 以上：**先问用户 / 稍等再重试**，不要连发同一批（每次都会白等 6 s）。`batch` 里某条 op 等满 6 s 会让整批顺延——"等"仍比"整批失败"划算。
- 码表补了 `0x800AC472`（Excel/ET 的忙变体）与 `0x80010001`（Word 侧原来缺），所以"忙"基本都归到 `BUSY`；`COM_ERROR` 现在更值得当"真的出错"看。

自检：设 `WOC_MSGFLT_DEBUG=1` 跑任意命令，stderr 会多一行 `woc-msgfilter: CoRegisterMessageFilter -> 0x00000000 (replaced=no)`；`0` 就是注册成功。

### 15.2 剪贴板：render 不再吃掉用户复制的东西

Excel `render` / `render -Overview` / `render -IncludeCharts`、PPT `render` / `render -Overview` / `export -Format png|jpg` 全都经过剪贴板。以前它们画完会 `Clipboard.Clear()`，用户刚复制的文本或图片就没了；现在：

- 进来先**整份存下**（所有格式 + 剪贴板序列号），画完**放回**；
- 只有"这一轮之后没人再动过剪贴板"才放回——用户中途按了 Ctrl+C，**他的新内容留下**，我们不动；
- 放回的数据标了"不进剪贴板历史 / 不上云 / 不参与监控"，所以**不会**把用户原来的内容推进 Win+V。

三条纪律：

1. **不要**为了"渲染后剪贴板空着"再补一次 `Clipboard.Clear()`——恢复是自动的，外面动剪刀反而会把刚放回的内容清掉。
2. 恢复失败不会让 `render` 报错（只写 Trace，`WOC_CLIP_DEBUG=1` 可见）：真遇到时剪贴板上留的是渲染图，等于旧行为。
3. 只保证"内容回来了"，**不保证跨应用的格式细节**（Excel 富文本粘到 Word 里样式仍按目标应用走）。验证就 `Get-Clipboard` 看内容，别对每种格式都作承诺。

取舍说清（别对外承诺过头）：如果用户恰好在我们"复制 → 清空"那几百毫秒里复制，他的内容仍会丢（上游 dsh-office 同样如此）；这条窗口没有消除，只是从"每次渲染必丢"缩到"极窄竞态"。
