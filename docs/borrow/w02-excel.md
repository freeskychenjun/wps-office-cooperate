# w02 · dsh-office Excel 域深读与可借鉴清单

> 对象：`D:\DSH工作区\_research\dsh-office`（v1.2.0，MIT，浅克隆；只读，本次未改其任何文件）。
> 对照：本仓库 `SKILL.md`（v1.28，Excel 段 305–470 行）、`scripts/excel.ps1`、`scripts/helper/excel/EtHelper.cs`、`reference/excel-quirks.md`。
> 全部引用格式 `相对路径:行号`，行号以本次读到的文件内容为准；dsh-office 的 helper 是单文件 `helper/OfficeHelper.cs`（9893 行），Excel 段集中在 4105–5521 与 9636–9722。
> 我方环境 WPS **12.1.0.29161**（`reference/excel-quirks.md:3`），dsh-office 的 WPS 实测机是 **12.1.0.28505**（`docs/verification.md:688`）——跨版本结论一律按「需实测」看待。
> 诚实前提：dsh-office 是**给模型用的插件**（COM 读写 + 自然语言操作说明），我方是**给 agent 用的命令行 skill**（每条命令一次进程/管道往返 + JSON）。它的能力边界可以直接照抄，它的"实现方式"必须按我方"一次块读、本地算、一次块写"的纪律重写，否则会退化成几千次 COM 往返。

---

## §1 功能清单（op/字段 → 语义 → 实现位置 → 技巧/坑）

### 1.1 操作说明即契约（没有 JSON schema）

`office_edit` 的 `ops` 项只有 `additionalProperties: true` 和一个 `op` 字段（`src/tools.ts:385-391`），**每个 op 接受哪些字段完全靠一段自然语言 GUIDE**（`src/tools.ts:132-147`），首次 `office_open/office_read` 时随结果下发，`office_help {app}` 可再取（`src/tools.ts:412-422`）。字段拼错由 helper 侧 `Bag.Unread()` 兜出（见 1.7）。

→ 启示：我方 `SKILL.md` 的"命令速查表 + 参数表"承担同样职能，**但我们的 `param-check.ps1` 是"静默丢弃 → stderr WARN"，它是"写进返回值里点名"**（见 1.7），后者模型一定看得到。

### 1.2 读取与大表概况

| op/字段 | 语义 | 实现位置 | 技巧 |
|---|---|---|---|
| `office_read {sheet, range, profile}` | 读表；`profile:true` 强制出列概况 | `src/tools.ts:351-364` | Excel 专用字段 `sheet/range/profile` 与 Word/PPT 共用一个工具 |
| 自动概况触发 | 无 `range` 时 >60 行或 >26 列才概况；给 `range` 时上限 2000 行×100 列 | `helper/OfficeHelper.cs:4190` | 「先看前 N 行 + 概况」而不是「只给概况」——模型仍能看到真实表头与头几行 |
| 截断说明 | `clipped` 文案里给出总行×列并提示用 `range` | `helper/OfficeHelper.cs:4193` | 让模型知道"你没看全"，且给下一步动作 |
| 概况护栏 | 区域 >3,000,000 格不出概况 | `helper/OfficeHelper.cs:4235-4240` | 大表不至于把 helper 拖死 |
| 日期识别 | 前 12 个数值格中**过半**的 `NumberFormat` 含 `[ymd年月日]` 且不是 `0.0`/`General` → 该列按日期显示 | `helper/OfficeHelper.cs:4486-4501` | 判据是"显示格式"而不是值本身；避免把序数、编号误当日期 |
| 读值转日期 | 命中日期列时 `FromOADate`，当天 0 点只显示 `yyyy-MM-dd`，否则带到分钟 | `helper/OfficeHelper.cs:4212-4220` | 日期列的值不再是 `46319.5` |
| 小数尾差 | 读取时 `double.ToString("G12")` 再 parse | `helper/OfficeHelper.cs:4211` | 去掉 `0.1+0.2` 类二进制尾巴 |
| 错误值 | 错误格在 `Value2` 里表现为**负整数**，helper 按码表映射成 `#DIV/0!`/`#N/A`/`#NAME?`/`#NULL!`/`#NUM!`/`#REF!`/`#VALUE!` | `helper/OfficeHelper.cs:4153-4169` | 不映射就会看到 `-2146826281` 这类整数 |
| 表对象清单 | 每张表列出 chart（名字/标题/位置/系列数）/pivot（名字/区域）/ListObject/筛选状态 | `helper/OfficeHelper.cs:4637-4658` | 让模型知道"这张表上有东西，别乱删" |
| 公式清单 | 值网格之外单列 `formulas: [{cell, formula}]` | `helper/OfficeHelper.cs:4222-4231` | 值/公式分开返回，模型不必二次读 |

### 1.3 概况口径（Profile，最有价值的原创部分）

| 输出 | 语义 | 位置 |
|---|---|---|
| 首行摘要 | `N data row(s) under the header row R`；空行数；**完全重复的行数** | `helper/OfficeHelper.cs:4517-4529` |
| 每列 kinds | `n number(s)` / `n date(s)` / `n text` / `n blank` / `n error(s)` | `4540-4572` |
| 数值列 | `min, median, mean, max, sum`（中位数真排序取中；mean 保留 2 位） | `4578-4581` |
| 日期列 | `from yyyy-MM-dd to yyyy-MM-dd` | `4575` |
| 离群值 | 仅当数值 ≥20 个：上栅栏 `q3 + 3×(q3−q1)`，超栅栏个数 **≤5%** 才报 `n far above the rest (over X)` | `4582-4590` |
| 离散列 | `n distinct`；`distinct≤24` 列 24 个取值、否则列前 10 个（各带次数，文本截 24 字）；取值多而重复多时只说"n repeat an earlier value" | `4595-4608` |
| **LOOK（值得注意）** | 多余空格（`s != Tidy(s)`）、**写成文本的数字**、**写成文本的日期**、数值与文本混杂 | `4609-4614` |
| 日期写法混杂 | 日期列里另有文本时，用"数字改 0"的形状取样 5 种不同写法 | `4594` + `4621-4634` |
| 不看的 | 合并单元格、多行表头、已是 ListObject 的表、几十万行（验证记录自认未测：`docs/verification.md:486`） | — |

> 这套口径的关键不是"统计"，而是**每列都带一个"下一步该干什么"的判据**（LOOK 段），且阈值都写成"少数派"（≤5%、≤12 distinct、≥20 个数值才判离群），避免在几行小表上乱报警。

### 1.4 清洗套件

统一入口：`ExcelData(app, book, ws, type, op)`（`helper/OfficeHelper.cs:4926-5521`），按 type 分派；`range` 省略 = `UsedRange`，首行当表头（`4363-4366`）；列可用**表头名 / 1 基序号 / 列字母**，找不到就报错并列出所有表头（`4397-4416`）。

| op | 语义/字段 | 位置 | 技巧与坑 |
|---|---|---|---|
| `replace` | `find/replace/column?/whole?/matchCase?` | `5029-5068` | 先扫一遍计数，**0 命中直接返回「nothing replaced」**（省一次写）；WPS 的 `Range.Replace` 会在刚替换出的内容里再替换一次（`P→P-` 得 `P--`），WPS 分支改为读 `Formula` 到内存做正则、跳过公式格再整块写回（`5045-5065`） |
| `clean` | `column/columns`、`trim`、`noSpaces`、`remove:[...]`、`case:upper/lower/proper`、`map:{...}`、`to:number/date/text`、`year` | `5069-5190` | 只做 trim 时 `columns` 缺省 = 全部列（`5072`）；`map` 键大小写不敏感且**键会先 Tidy**（`5082-5084`）；`to:date` 未给 `year` 时**取整列年份众数**（`5094-5109`），"3月1日"这类补年；数字型日期 `20250301` 单独走一条分支（`5138-5143`）；失败格计数 + 前 5 个样例的地址与原文（`5133-5134`）；列里有公式时**只写非公式格**（`5150-5165`）；`to:text` 顺手设 `@` 格式（`5152`）；`to:date` 设 `yyyy-mm-dd`（`5167`）；改完若该列只剩 ≤16 个取值就回列一遍现状（`5169-5187`）——"清洗有没有成功"的自检 |
| `fill_blanks` | `with: 值 \| "above" \| "mean" \| "median" \| "mode"` | `5191-5244` | 逐列算一次统计量再填；`mode` 取首个众数；无数字可算 mean 时直接报错而不是填 0 |
| `remove_duplicates` | `columns?`（缺省整行） | `4973-4995` | 键 = `Tidy(值).ToUpper()`，**全空行跳过不算重复**；报告删掉的行号（列 12 个+`..`，`4682-4687`） |
| `delete_rows` | `{blank:true}` 或 `{where:{column, is: blank\|=\|<>\|>\|<\|>=\|<=\|contains\|text\|error, value?}}`、`{column, values:[...]}`、`{row,count}` | `4996-5028` + `Test()` `4892-4923` | 条件语法集中在一个 `Test()`：数值两侧都能解析才比大小，否则按文本（大小写不敏感）比较；`in` 是一组取值；未知条件名直接抛错并列出可选 |
| `sort` | `by:[{column, order}]`、`header` | `4928-4949` | 走 `ws.Sort.SortFields`（不是逐行比较），`Header` 默认 1 = 首行是表头 |
| `filter` | `{column, values:[...]}` 或 `{column, criteria, criteria2?, join?}` | `4950-4966` | `values` 走 `AutoFilter Operator:7`（xlFilterValues）；显示行数用 `SpecialCells(12).Count - 1` 数出来并回报"n of m 行可见（隐藏未删除）" |
| `clear_filter` | `keepArrows?` | `4967-4972` | 先 `ShowAllData()` 再按需关掉 `AutoFilterMode` |
| `copy_range` | `range, to, values?` | `5298-5322` | **`values:true` 必须用 Excel 自己的 `Copy`**，因为整块 `Value2=` 写回会让 Excel 重新解析文本（`20250301` 变数字、`10%` 变 0.1，`docs/verification.md:479` 实测踩过）；再用 `SpecialCells(-4123)` 把公式区换成结果；列宽一并带过去 |
| `insert/delete_columns` | `column, count` | `5288-5297` | 列字母 → `Range[at+"1"].Resize[1,count].EntireColumn` |
| 底层删行 | — | `4661-4680` | **自底向上**删，相邻行合并成区间、每次 `Range["a:b,c:d,..."].Delete()` 最多 18 段——避免"每行一次 COM" |

### 1.5 分析

| op | 语义/字段 | 位置 | 技巧 |
|---|---|---|---|
| `fill_formula` | `range, formula` | `5245-5254` | 整块 `Range.Formula` 一次写（Excel 自己按相对引用平铺），写完立刻报错格 |
| `add_column` | `header, formula\|value, after?, numberFormat?, values:true` | `5255-5287` | **同名表头视为"重填"而非重复加列**（`5260-5263`）；`after` 才真插一列并 `Resize` 扩块；`values:true` 用 `body.Value2 = body.Value2` 把公式固化成值（`5285`）；表头加粗跟随 A1 |
| `Errors(range)` | 写完即查错 | `4690-4711` | 返回值追加 `— n cell(s) show an error (first: F7 #VALUE!)`，把"公式错在哪"第一时间推给模型 |
| `pivot` | `source?, rows, columns?, values:[{field,fn,name?}], filters?, group:{field,by}, sort:{field,by?,order?}, to?, name?` | `5323-5412` | 源有**空表头即拒**（`5327`，透视表要求每列有名字）；默认新建工作表名"透视表"、落 `A3`；`PivotCaches().Create(SourceType:1, SourceData:source)` → `cache.CreatePivotTable(TableDestination:target)`（`5338-5339`）；行/列/筛选 = `PivotField.Orientation` 1/2/3 + `Position`；聚合码 `count -4112 / average -4106 / max -4136 / min -4139 / sum -4157`，默认数字格式 `#,##0.00`（count 不设）；日期分组走 `DataRange.Cells[1,1].Group(Periods: [月/季/年])`，**失败不抛错**而是写 NOTE 并建议"加一列公式月份再透视"（`5377-5380`）；`RowAxisLayout(1)` = 表格布局 + `AutoFit`；**结果不调 `RefreshTable`**，直接读 `TableRange1` 的 `Value2` 网格（前 40 行×14 列，数字保留 2 位）拼成文本返回（`5397-5411`）——模型不用再 read 一次 |

> 关于"结果怎么取回"：新建 `CreatePivotTable` 当场就完成布局，所以**不需要 RefreshTable**；对已存在的透视表改字段才需要。它的取回是"**网格快照 + 行号前缀**"，可直接读，也可给出 `to` 让它落到指定格。

### 1.6 图表

| op/字段 | 语义 | 位置 | 技巧/坑 |
|---|---|---|---|
| `add_chart {type, range \| categories+series, title, xTitle, yTitle, y2Title, labels, legend, numberFormat, min, max, at, width, height, name}` | 建图 | `5501-5517` + `Chart()` `4755-4877` | 先用 `ws.Shapes.AddChart2(-1, kind, ...)`（Excel 2013+ 新样式），失败回退 `ChartObjects().Add`；**建图失败会删掉半成品**（`5516`） |
| 类型表 | column 51 / bar 57 / stacked 52 / stacked bar 58 / line 65 / smooth 4 / pie 5 / doughnut -4120 / scatter -4169 / scatter line 74 / area 1 / radar -4151 / combo 51 | `4713-4731` | 未知类型报错时把可选值全列出来 |
| 系列约定 | `range`=首列分类、首行系列名（`SetSourceData`）；或 `series:[{name, values, x?, chart?, axis?, color?}]` 显式建 | `4760-4784` | 动手前先删掉"在数据区内建图时 Excel 自己猜出来的系列"（`4762-4763`）；`name` 形如单元格引用时写成 `"=绝对地址"`，让系列名跟着格走（`4779`） |
| 组合图/次坐标轴 | `type:"combo"` 且未逐条给 `chart` 时，**最后一条系列自动改折线**（`4795`）；`axis:"secondary"\|"right"` → `Series.AxisGroup = 2`（`4796`） | — | 逐条 `SeriesCollection(index)` 设 `ChartType`/`AxisGroup`，在系列建完之后（`4786-4799`） |
| 数据标签 | `labels:true \| "percent" \| "value"`；饼/圆环默认开 | `4809-4824` | `percent` → `ShowPercentage=true, ShowValue=false`，饼/圆环还带分类名；字号 9 |
| 图例 | `legend: bottom\|right\|top\|left\|none`；缺省=圆类型右、多系列下、单系列无 | `4825-4837` | 默认值按系列数推，模型少写一个字段也好看 |
| 轴标题 | xTitle/yTitle **跳过饼/圆环/雷达**（那类没有坐标轴，`4802-4804`）；`y2Title` 用 `Axes(2,2)`（`4805`） | — | 对无轴图表设轴标题会抛异常，它选择静默跳过而不是失败 |
| 数字格式/上下限/轴刻度 | `numberFormat`、`min`、`max` 改的是**主轴**（`4806-4808`） | — | 次坐标轴的刻度目前没开字段（缺口） |
| 新图观感 | 雅黑 10 号、主网格线 `#E4E6EA`、柱间距 `GapWidth=80`、图区无边框 | `4838-4845` | "开箱不像 2007"靠这几行 |
| **占地区域回报** | 返回 `at A:B` 的 `TopLeftCell:BottomRightCell` | `4852-4853` | 附操作说明"默认尺寸约占 9 列×21 行"（`src/tools.ts:144`），模型知道下一张图放哪 |
| **重叠检测** | 与同表其它 `ChartObject` 做 2px 容差相交判定，命中就点名 + 给改法 | `4864-4875` | 这是"仪表盘最常见的错"，它做成了硬提示 |
| 空图提示 | 0 系列 → `— NOTHING is plotted: check the range` | `4876` | 把"建了张空图"变成一句话 |
| 透视表源 | `chart.PivotLayout != null` → 提示"会用该表全部值并跟随筛选，要局部先把值写到普通格" | `4854-4862` | 另一个高频困惑 |
| `set_chart` / `delete_chart` | 按名字或序号 | `5518-5519` + `ChartByName()` `4879-4890` | 找不到图表时**列出现有图表名** |
| **WPS 专属重试** | 组合图把系列挪到另一类型/坐标轴后，WPS 会短暂拒绝设标题：重试 3 次（间隔 250 ms），第 3 次改用 `chart.SetElement(2)` 再设 `ChartTitle.Text` | `4736-4753` | 见 §5 摘录；这是它唯一为 WPS 写的图表补丁，说明**组合图+次坐标轴在 WPS 下确实建得起来**（否则不需要这个补丁） |

### 1.7 表格对象 / 条件格式 / 冻结 / 验证

| op | 语义 | 位置 | 技巧/坑 |
|---|---|---|---|
| `table {range?, name?, style?}` | ListObject（带纹行+筛选箭头） | `5413-5428` | 先用 `app.Intersect(existing.Range, data)` 判断这块是否**已经是表格**，是则只改样式/名字并说明（`5418-5426`）；`ListObjects.Add(1, data, Missing, 1)`（1=xlSrcRange，末位 1=有表头）；样式默认 `TableStyleMedium2`。坑：跨表区域建表原先报错，已修（`docs/verification.md:667`） |
| `conditional_format {range, rule, value?, value2?, fill?, color?, bold?}` | 规则：`> < >= <= = <> between top bottom duplicate unique blank formula colorscale databar clear` | `5429-5453` | 规则名→枚举：`between 1 / = 3 / <> 4 / > 5 / < 6 / >= 7 / <= 8`，`Add(Type:1, Operator:code, Formula1:, Formula2:)`；`blank` 用 `Type:10`；`formula` 用 `Type:2, Formula1:"=..."`；`AddColorScale(3)`/`AddDatabar()`/`AddUniqueValues()`/`AddTop10()` 各走便捷方法；`eq()` 自动把值包成 `="文本"` 或 `=数字`；默认填充 `#FFC7CE` + 字色 `#9C0006` |
| `freeze {cell}` | 冻结窗格 | `5454-5489` | 过程很讲究：取**本工作簿自己的** `Windows[1]` → `ws.Activate()` → 切普通视图（`View=1`）→ 关旧冻结/拆分 → `ScrollRow/Column=1` 归一 → 设 `SplitRow/SplitColumn` → 再 `FreezePanes`；失败 3 次每次都 `window.Activate()` 后重试，仍失败**只报"这条不影响其余"**（`5487`）。原因记录：Excel 在窗口非活动时拒绝设 `FreezePanes`（`docs/verification.md:496`） |
| `validation {range, list:[...]}` | 数据验证（仅下拉列表） | `5490-5500` | 先 `Validation.Delete()` 再 `Add(Type:3, AlertStyle:1, Operator:1, Formula1:"a,b,c")`（3=xlValidateList，1=xlValidAlertStop）。坑：列表值是**逗号拼接**，值里含逗号/总长 >255 会坏（它没处理，属缺口） |
| 批处理护栏 | — | `9636-9722` | 顺序执行、**首个失败即停**，返回 `done[]`（成功项文案）+ `failed{index, op, error, code}`（`9657-9683`）；**未知字段点名**：`op.Unread()` 把没被任何分支读过的字段列出来 → `— IGNORED field(s) xxx: write_range does not take them, check the operation list`（`9669-9670`）；Word 整批进一条撤销记录，Excel 明确没有撤销（`src/tools.ts:101`） |
| 区域带表名 | `"产品表!A1:E13"` 在所有 op 里都认 | `4124-4144`（`Cells`）+ `4369-4378`（`Ref`） | 早期只有 `copy_range/pivot/图表` 认，其余报"不是有效区域"，已统一（`docs/verification.md:495`） |

### 1.8 渲染

| 能力 | 做法 | 位置 | 备注 |
|---|---|---|---|
| 区域 → PNG | `range.CopyPicture(1, 2)` → 等剪贴板最多 1 s → `Clipboard.GetImage()` → 存盘；**前后保存/恢复剪贴板里的文本** | `9871-9888` | 只恢复文本，图片/文件会丢（`README.md:142`） |
| 单张图表 → PNG | `ChartObject.Chart.Export(path, "PNG")` | `9862-9870` | **我方不可照抄**：WPS 下自己写盘的文件在本机会被透明加密（`reference/excel-quirks.md:19`、`40`），我们已改走 `CopyPicture` 合成 |
| 带图表的区域 | 就是普通区域渲染，说明里写"盖在区域上的图表在图里"（`src/tools.ts:440`） | — | **这条在 WPS 下不成立**：我方实测 `Range.CopyPicture` **不含**浮动图表（`reference/excel-quirks.md:39`）——我方 `-IncludeCharts` 扩视图的方案比它正确 |
| Excel 的 overview | 等同单张区域图 | `src/tools.ts:263` | 没有多页拼接 |

---

## §2 我方现状（已有 / 半有 / 没有）

我方 Excel 命令集：`scripts/excel.ps1:11-12` 的 `ValidateSet`（26 条 + `batch`），实现 `scripts/helper/excel/EtHelper.cs`；`scripts/tasks-excel.ps1` 是**多任务拆分/合并**（split/merge/plan，`scripts/tasks-excel.ps1:5-27`），不是数据操作。

| dsh-office 能力 | 我方对应 | 状态 | 说明 |
|---|---|---|---|
| `office_read` 6 万行概况 + LOOK | `read [-Profile]`（`EtHelper.cs:954` 触发、`ProfileOf` `998-1065`） | **半有** | 我方有 numbers/texts/blanks/others + min/median/max + top8 + 头尾各 5 行样本；**缺**：日期识别（值是序列号，`JsonCellValue` `720-726` 无日期分支）、LOOK 五类、空行/完全重复行、离群值、"取值写法混杂"取样、3M 格护栏；触发条件也不同（我方 >60 行**或** >15 列就只给概况，dsh-office 是前 60 行 + 概况） |
| 错误值识别、写完报错格 | 无 | **没有** | 我方 `fill`/`write` 之后不知道有没有 `#N/A`；`JsonCellValue` 会把错误码当整数返回 |
| `write_range` | `write -Items`（原子批 + `expect` 锚 + 失败回滚 + before/after，`SKILL.md:387-403`） | **已有（更优）** | dsh-office 无锚无回滚 |
| `format_range` | `format`（分号多区域 + `-Items` 每块不同格式 + 颜色通道已修，`SKILL.md:404-421`） | **已有（更优）** | 我方多了"每块不同格式"与参数校验 |
| `autofit` | `autofit` | **已有** | — |
| `add_sheet/rename_sheet/delete_sheet` | `sheet`（+ `activate`） | **已有（更优）** | 我方多 `activate` |
| `insert_rows/delete_rows/insert_columns/delete_columns` | `rows` / `cols`（仅按位置插删，`SKILL.md:324`） | **半有** | 缺"按条件删行"（`delete_rows {where}`、`{blank}`） |
| `fill_formula` | `fill`（首格写一次 + 原生 FillDown，`SKILL.md:319`） | **已有（更优）** | 我方 FillDown 比整块 `.Formula` 更稳（大表实测） |
| `add_column` | 无（可 `write` 表头 + `fill` 组合） | **没有** | 缺"同名表头=重填"与 `values:true` 固化 |
| `clean/replace/fill_blanks/remove_duplicates/sort/filter/clear_filter/copy_range` | 无 | **没有** | **本次最大空白**，8 个 op 全部缺 |
| `pivot` | 无 | **没有** | `reference/excel-quirks.md:31` 明写"未验证：数据透视表" |
| `add_chart/set_chart/delete_chart` | `chart -Action add/list/delete/move`（`SKILL.md:326`，`EtHelper.cs:1268-1406`） | **半有** | 我方：6 类型（column/bar/line/pie/scatter/area）、显式 NewSeries、多系列、轴标题、可移动缩放、跨表 `list -All`。**缺**：combo/次坐标轴、数据标签、图例位置、数字格式/上下限、重叠检测、占地区域回报、0 系列硬提示 |
| `table`（ListObject） | 无 | **没有** | — |
| `conditional_format` | 无 | **没有** | — |
| `freeze` | 无 | **没有** | — |
| `validation` | 无 | **没有** | — |
| 单张图表出图 | 无（只有区域，`EtHelper.cs:1807-1926`） | **半有** | 我方 `render -IncludeCharts` 更强，但"只出一张图"的窄请求还要靠整区域 |
| 批处理 | `batch -Ops`（`SKILL.md:362-366`） | **半有/语义不同** | 我方**失败继续**并逐条报 ms；dsh-office **首个失败即停**。我方缺 `IGNORED field` 级点名（只有 `param-check.ps1` 的 stderr WARN） |
| 快照/回滚 | `snapshot/diff/restore`（含 `-Cells` 选择性回滚，`SKILL.md:329-331`） | **已有（我方独有）** | dsh-office 明确"Excel 改坏只能靠重做"（`README.md:140`） |
| 单元格批注 | `comment` | **已有（更优）** | dsh-office 无批注 |
| 查找/选区 | `search` / `selection` | **已有（更优）** | dsh-office 无 |
| 静默后台实例 / 回合结束自动保存关闭 | 无（我方 `open` 默认可见并激活，`SKILL.md:14`） | **不适用** | 定位相反：我方要"用户在 WPS 里看着"，它是"后台不打扰" |
| 套件选路（office/wps/auto） | 硬钉 `KET.Application`（`reference/excel-quirks.md:8-9`） | **不适用** | 我方只做 WPS，多套件反而危险 |

小计：**已有/更优 8 项，半有 5 项，没有 8 项，不适用 2 项**。缺口全部落在"数据操作"这一层，而不是"读写单元格"这一层——这正是我们面向"历年水文资料 Excel、报告附表、几千行清洗"最缺的一段。

---

## §3 差距与价值（面向本 skill 的真实场景）

场景锚点：① 历年水文资料表（几百到几万行，站码/年份/月/水位/流量，脏在写法而非结构）；② 报告附表（版式要求严、数据量小）；③ 几千行清洗 + 出图（过程线/族曲线/频率曲线）。

### P0（不做就会一直当场外手工干活）

1. **概况的 LOOK 段**（`4511-4614`）：写成文本的数字/日期、多余空格、类型混杂、空行、重复行、离群值。水文资料里"1963 年"写成 `1963` / `1963年` / `63`、水位带全角空格、站码被存成文本，是最常见的三类。我方 `ProfileOf` 只有计数与分位数，**看不出"这列有问题"**。
2. **`ParseNumber` / `ParseDate` + 年份众数推断**（`4451-4483`、`5094-5109`）：`¥1,299`、`89元`、`10%`、`12 件`、`20250301`、`2025.03.01`、`2025年3月1日`、`3月1日`。**纯本地正则，零 COM 风险，收益立刻可见**。
3. **错误值识别 + 写完即报错格**（`4153-4169`、`4690-4711`）：`fill` 一列 `=B2/C2` 之后应当立刻告诉模型"F7 #DIV/0! 共 3 格"，而不是等模型自己发现或用户发现。
4. **`copy_range` 的 `values:true` 用 Excel 自己 Copy**（`5298-5321`）：这条同时解决"复制到副本"和"不要被重新解析"，是我方缺的"保留原表、在副本上清洗"工作流的前提。
5. **`replace` 的 WPS 双替换规避**（`5045-5065`）：我方一旦加 `replace`，WPS 的 `Range.Replace` 必然踩这个坑（`P→P-` 得 `P--`），这是它现场实测出来的，直接省我们一轮返工。

### P1（有了效率数量级不同）

6. `sort` / `filter` / `clear_filter`（`4928-4972`）：用 `ws.Sort` 与 `AutoFilter`，一次 COM 排序几千行，比我方"读出来排完写回"快两三个数量级（且不破坏公式联动）。
7. `delete_rows {where}` + `remove_duplicates` + 自底向上区间删除（`4661-4680`、`4973-5028`）：删除必须一次 `Range.Delete`，逐行删是 O(n) 往返。
8. `clean` 的整列写回策略（`5110-5188`）：**块读一次 → 本地逐格改 → 块写一次**，并且"列里有公式就只写非公式格"；这与我方 v1.19 `batch` 的省时逻辑同构，可复用。
9. `fill_blanks`（`5191-5244`）：报告附表里合并/留空格补"同上"极常用。
10. `add_column` 的"同名=重填 + `values:true` 固化"（`5255-5287`）。
11. 图表的 combo/次坐标轴/数据标签/图例/数字格式 + **占地区域回报与重叠检测**（`4755-4875`）：过程线+雨量柱组合图（次坐标轴）是水文报告刚需；重叠检测是仪表盘类交付的自检。
12. `pivot` 结果直接回读（`5323-5412`）：频率统计/年月汇总的中间表，能省掉"建完再 read 一次"。

### P2（结构性补强，先探测再动手）

13. `table`（ListObject）、`conditional_format`、`validation`、`freeze`（`5413-5500`）：报告附表与看板的"成品感"三件套；WPS 可行性有高有低（§4.2）。
14. 批处理的 `IGNORED field` 点名（`9669-9670`）：我方 `param-check` 已经在 stderr 做了，**把它同时写进返回值**是低成本改进。
15. 单张图表出图（`9862-9870` 的意图，不照抄实现）。

### 不借（会让我们变差）

- **`Chart.Export` 落盘**：本机落盘即密文，读不回来（`reference/excel-quirks.md:19`、`40`）。
- **区域渲染"图表自动在图里"的假设**（`src/tools.ts:440`）：WPS 下不成立，我方 `-IncludeCharts` 是对的。
- **"Excel 里做数据一律让模型写代码"的定位**（`src/prompt.ts:27`、`docs/verification.md:1020`）：它放弃 Excel 窗口做数据分析是**因为它是通用插件**；我方用户就是"在 WPS 里看着改"，照搬会自我否定。
- 无撤销/无锚点/无快照的所有写法：我方已有更强的 `snapshot/diff/restore` + `expect`，不要为了"对齐功能"退回去。

---

## §4 落地建议

### 4.1 接入方式（PowerShell/COM 侧的具体做法）

一律遵守我方既有纪律：**块读一次 → 本地算 → 块写一次**（`SKILL.md:423-429`），不要按 dsh-office 的"每格一次读写"照搬。

| 新增命令 | 建议的 COM 接法 | 关键点 |
|---|---|---|
| `read -Profile`（增强，非新命令） | 沿用 `ReadValues`（`EtHelper.cs:708-718`）读整块 `Value2` + 逐列 `NumberFormat`（只需每列取 `Columns(j).NumberFormat` 一次，或取整块 `.NumberFormat`） | 日期判定按 dsh-office 的"过半 + NumberFormat 含 ymd"（`4486-4501`），但要**整块一次取 NumberFormat**（`range.NumberFormat` 返回 `object[,]`），不要逐格；LOOK 全部本地算 |
| `clean` | 整列 `Value2` 一次读 → 本地 Tidy/ParseNumber/ParseDate → `Range.Value2 = object[,]` **0 基精确尺寸**一次写（WPS 坑，`reference/excel-quirks.md:11`） | 有公式的列要按 `HasFormula` 跳过（可先 `range.Formula` 读一次判定）；`to:date` 顺带 `NumberFormat="yyyy-mm-dd"`；`to:text` 先设 `"@"` |
| `replace` | 走我方已有 `write` 的块写路径，在内存里做替换；**不要用 `Range.Replace`**（WPS 双替换） | 命中计数在内存里算，0 命中直接返回 |
| `fill_blanks` | 整块读 → 本地求 mean/median/mode/above → 只把空格的地址收集起来，按连续区间块写（或整列写回） | `above` 是"上一个非空" |
| `sort` | `ws.Sort.SortFields.Clear/Add(Key:列, Order: 1|2)` + `SetRange` + `Header` + `Apply`（`4928-4948`） | WPS 待探测（多数实现有），失败回退"读-排-写" |
| `filter` / `clear_filter` | `range.AutoFilter(Field, Criteria1, Operator)`；`ws.ShowAllData()` / `ws.AutoFilterMode=false` | `Operator:7`=xlFilterValues 列表；显示行数用 `SpecialCells(12)`（WPS 待探测） |
| `delete_rows -Where` | 条件判定本地做（读一次 `Value2`），收集行号后**自底向上合并区间一次 `Range["a:b,c:d"].Delete()`**（`4661-4680`） | 行号必须换算成**工作表绝对行号**（`row0 + r - 1`）；报告删了哪些行 |
| `copy_range` | `source.Copy(dest.Cells[1,1])` + `app.CutCopyMode=0`（WPS 下要 try/catch，可能没有该属性）；`values:true` 时再 `SpecialCells(-4123)` 转值 | **不要**用 `dest.Value2 = src.Value2` 代替（文本重解析） |
| `add_column` | 同名表头先查 → `body.Formula = "=..."` 整块填 → `values:true` 时 `body.Value2 = body.Value2` | `after` 需 `EntireColumn.Insert()` |
| `pivot` | `wb.PivotCaches().Create(SourceType:1, SourceData:range)` → `CreatePivotTable(TableDestination:cell)` → `PivotFields(name).Orientation/Position` → `AddDataField(field, caption, code)` → 读 `TableRange1.Value2` | WPS 待探测；`Group(Periods:)` 单独探测 |
| 图表扩展（不新增命令，扩 `chart` 参数） | `-Type combo`、`-Secondary`、`-Labels`、`-Legend`、`-NumberFormat/-Min/-Max`；返参加 `range:"F2:M22"` 与 `overlaps:[...]` | 系列先建完再逐条改 `ChartType/AxisGroup`（顺序不能反，`4786-4799`）；**WPS 标题重试逻辑（`4736-4753`）照抄** |
| 单张图表出图 | 扩 `render -Chart <name>`：`ChartObject.CopyPicture(1,2)` → 剪贴板 → GDI+ 存盘（复用我方 `render` 已有的剪贴板互斥 `Local\woc-clipboard`） | **禁止 `Chart.Export`**（密文） |
| 错误值检查 | 在 `fill`/`write` 返回里加 `errors:[{cell, code}]`：读回 `Value2`，`int` 且落在错误码表里即错误（`4153-4169`） | 只需一次读回，成本 1 次往返 |

### 4.2 WPS 支持度评估（逐项）

| 项 | 判断 | 依据 | 验证方法 |
|---|---|---|---|
| **PivotTable** | **能做（证据较强）** | dsh-office 在 WPS 12.1.0.28505 上用同一套代码跑通 **4 张透视表**（含按季度分组），且与微软 Excel 结果一致（销售额合计 2047783、1086 行）（`docs/verification.md:709`）；其透视表代码没有任何 WPS 分支，说明走的是 WPS 兼容的那条老对象模型路径 | 临时新簿 → `PivotCaches().Create(SourceType:1,...)` → `CreatePivotTable` → 读 `TableRange1`；再单独测 `Group(Periods:)` 的月/季/年 |
| **ListObject / `table`** | **需实测（不确定）** | dsh-office 的 WPS 那一轮只跑了"清洗、补列、4 张透视表、3 张图表、看板"（`docs/verification.md:709`），**没有提 ListObject**；`table` 只在微软 Excel 下验过（`docs/verification.md:480`） | 临时簿：`ListObjects.Add(1, range, Missing, 1)` → 回读 `.Name/.Range.Address/.TableStyle`；再测 `Intersect` 判定"已是表格"；失败时回退"不建表，只做格式+自动筛选" |
| **ConditionalFormatting** | **大概率能做（高置信）** | 同上，WPS 那轮的"看板"按 `docs/verification.md:478` 的构成含条件格式；且 `FormatConditions` 是 Excel 97 就有的一代接口，WPS 兼容层通常实现了它 | 临时区域逐条：`Add(Type:1, Operator:5, Formula1:"=100")`、`AddColorScale(3)`、`AddDatabar()`、`AddTop10()`、`AddUniqueValues()`；每条回读 `FormatConditions.Count` 与 `Item(i).Type`，并 `render` 目检 |
| **`FormatConditions.Add` 的便捷方法** | **需实测（比上一项更可疑）** | `AddColorScale/AddDatabar/AddUniqueValues/AddTop10` 是 2007 之后的新接口，WPS 兼容层常缺 | 若缺：退路只有 `Add(Type, Operator, Formula1)` + 手工设 `Interior.Color`（做不了色阶/数据条，只能做阈值着色） |
| **Validation** | **需实测** | dsh-office 只在微软 Excel 下验过（`docs/verification.md:480`），WPS 那轮未提 | 临时区域：`Validation.Delete()` → `Validation.Add(Type:3, AlertStyle:1, Operator:1, Formula1:"甲,乙,丙")` → 回读 `.Type/.Formula1`；再测 >255 字符与值含逗号 |
| **次坐标轴组合图** | **能做（证据较强）** | 两条独立证据：① WPS 轮的"看板"与微软轮构成相同（含组合图+次坐标轴，见 `docs/verification.md:478`），而 WPS 轮明写"看板…结果与微软 Excel 下一致"（`docs/verification.md:709`）；② 更有力的是 dsh-office 专门为 **WPS** 写了"组合图设标题被拒绝要重试"的补丁（`helper/OfficeHelper.cs:4736-4753`、`docs/verification.md:699`）——若组合图建不起来，不需要这个补丁 | 两系列 → 系列1 column、系列2 `ChartType=65; AxisGroup=2` → 读 `Axes(2,2).HasTitle/AxisTitle.Text` → 再走一遍"标题重试"分支；`render` 目检双轴 |
| **`Borders` 常量** | **已有且更优，不需借** | 我方 `format -Borders grid/thick/none` 实现在用，且 v1.16 修过"对角线（XlBordersIndex 5/6）被误画成 X"（`SKILL.md:348`）；dsh-office 只做 `range.Borders.LineStyle = 1`（`OfficeHelper.cs:4269`），比我方差得远 | 不需要新探测 |
| **`freeze`** | **需实测（风险中）** | 我方环境 `wb.ActiveWindow` 常为 null、`Windows(1)` 可用（`reference/excel-quirks.md:24`），dsh-office 的写法正是 `book.Windows[1]`；但"窗口非活动时拒绝 FreezePanes"在 WPS 下的表现未知 | 临时簿：`ws.Activate()` → `Windows(1).SplitRow/SplitColumn` → `FreezePanes=true` → 回读 `FreezePanes/SplitRow/SplitColumn`；给 `-Show`（WPS 12 窗口属 `wps.exe`，`reference/excel-quirks.md:62`） |
| **`SpecialCells(12)` / `SpecialCells(-4123)`** | **需实测** | 用于"数可见行"与"公式格转值"，WPS 兼容层可能只实现部分 `xlSpecialCellsType` | 在筛选后与公式区各测一次，失败则改成本地判断（筛选状态读 `Row.Hidden`） |

> 一条总原则：**先写一个 8 项探测脚本，再决定写哪几个命令**。上表 5 项「需实测」+ 1 项「高置信但本机未验」（ConditionalFormatting），探测成本合计不到 1 小时（一次 `batch` + 一个新簿），但能避免按错误假设写 2–3 天的命令。

### 4.3 探测脚本思路（不碰用户文件）

```
1) 前置：excel.ps1 doctor                              # 确认 KET 在位、渲染可用
2) 新簿：excel.ps1 new -Doc "$env:TEMP\woc-probe-<ts>.xlsx"   # 临时文件，绝不用用户工作簿
3) 造数：write -Items（地区/月份/金额/水位 4 列 × 12 行 + 一张"类型/阈值"小表）
4) 逐项探测：每个 probe 独立 try/catch，输出 {probe, ok, hr, readback}
   P1 pivot      : PivotCaches.Create → CreatePivotTable → 字段摆位 → AddDataField → 读 TableRange1
   P2 pivotgroup : DataRange.Cells(1,1).Group(Start:true, End:true, Periods:...)
   P3 listobject : ListObjects.Add(1, range, Missing, 1) → 读 Name/Range/TableStyle
   P4 cf         : Add(Type:1,Operator:5) / AddColorScale(3) / AddDatabar / AddTop10 / AddUniqueValues
                   → 每步回读 Count/Type
   P5 validation : Validation.Delete → Add(Type:3,...) → 读 Type/Formula1
   P6 combo      : 系列2 ChartType=65 + AxisGroup=2 → Axes(2,2) 标题（含"被拒→重试→SetElement(2)"分支计数）
   P7 freeze     : Activate → SplitRow/SplitColumn → FreezePanes → 回读
   P8 specials   : AutoFilter 后 SpecialCells(12).Count；公式区 SpecialCells(-4123).Areas.Count
5) 目检：render -IncludeCharts -Out "$env:TEMP\probe"   # 色阶/数据条/双轴只能看图
6) 收尾：close（不保存）→ 删临时文件；检查 et 进程数（SKILL.md:469）
7) 落盘：把 {probe,ok,hr} 直接写进 reference/excel-quirks.md 的新一节（一条一行，带日期与 WPS 版本号）
```

判据：`ok=true` 且 `readback` 与预期一致才算"能做"；只 `ok=true` 但读不回（WPS 常见"调用成功、对象残缺"）按**半做**处理，并记录规避写法。

### 4.4 风险与工作量

| 阶段 | 内容 | 预估 | 风险 |
|---|---|---|---|
| 探测 | §4.3 的 8 项 | 0.5 天 | 低（临时文件、只读探测为主） |
| P0-a | 概况增强（日期识别 + LOOK 五类 + 空行/重复行 + 离群值） | 0.5 天 | 低：纯本地计算，只需整块 One-shot `NumberFormat` 读一次 |
| P0-b | `ParseNumber/ParseDate` + 错误值检查（`fill/write` 返参加 `errors`） | 0.5–1 天 | 低 |
| P0-c | `replace`（内存替换，绕开 WPS 双替换）+ `copy_range`（Excel 自己 Copy） | 0.5 天 | 中：`CutCopyMode`/`SpecialCells` 在 WPS 待验 |
| P1 | `clean` / `fill_blanks` / `remove_duplicates` / `delete_rows -Where` / `sort` / `filter` | 2–3 天 | 中：全部依赖"块读+本地算+块写"三段式，注意 0 基数组（`reference/excel-quirks.md:11`）与 `Text` 多格取不到（同文件 `12`） |
| P1 | `add_column`、图表扩展（combo/次轴/标签/图例/重叠/占地） | 1–1.5 天 | 中：WPS 组合图标题重试必须照抄；标签/图例接口在 WPS 逐项试 |
| P2 | `pivot`（+按探测结果决定 `table/conditional_format/validation/freeze`） | 2–4 天 | 高：透视表/ListObject/CF/Validation 任一在 WPS 残缺，就要写降级路径（如"不建透视表，改用 SUMIFS 公式矩阵"） |
| 文档 | `SKILL.md` 命令表 + `reference/excel-quirks.md` 新增实测节 + `tests/excel/` 探测断言 | 0.5 天 | — |

**合计约 7–10 天**（P0 部分 1.5–2 天即可交付并立刻见效）。建议顺序：**先 P0（纯本地、零 WPS 风险）→ 再跑探测脚本 → 再决定 P1/P2 的取舍**，不要先写透视表。

另有一条成本更低的并行收益：把我方 `ProfileOf`（`EtHelper.cs:998-1065`）的字段名与 dsh-office 的文本口径**对齐**（如 `far above the rest`、`LOOK:`、`number(s) written as text`），这样以后看两份实现的心智负担最小。

---

## §5 原文摘录（各 ≤30 行）

### 5.1 `src/tools.ts:132-135` — Excel 的三条全局约定与"读完先看概况"

```
excel: `Excel ops. sheet = sheet name (default: the active sheet). A range may name its sheet: "产品表!A1:E13". Where an op takes "range" for a block of data and you leave it out, it means everything filled on the sheet, the first row being the headers; a column is named by its header ("金额"), or by its letter.
Reading: office_read shows the first 60 rows; for a longer table it adds a profile of every column (kinds of values, min / median / max, the distinct values with counts, and what looks wrong: stray spaces, numbers or dates written as text, mixed kinds, far-off values, empty and repeated rows). Work from the profile, do not page through thousands of rows; office_read {profile: true, range?} gives it for any block. Dates are shown as dates.
- write_range {range: top-left cell, values:[[..],..]} — a string starting with "=" is a formula; also takes the format_range fields.
- fill_formula {range: "F2:F500", formula: "=D2*E2"} — one formula down (or across) a range: write it for the first cell, relative references adjust; use $ to pin. add_column {header, formula | value, after?: column, numberFormat?, values?: true} — a new column at the right of the data (or after a column) with the formula filled down all data rows, written for row 2; values: true leaves the results instead of formulas; an existing header is filled again. Both report cells that show an error.
```

### 5.2 `helper/OfficeHelper.cs:4511-4529` — 概况的表头摘要（空行/重复行）

```csharp
    static List<object> Profile(dynamic data)
    {
        List<object> lines = new List<object>();
        object[,] grid = Grid(data);
        int rows = grid.GetLength(0), cols = grid.GetLength(1), row0 = (int)data.Row, col0 = (int)data.Column;
        if (rows < 2) return lines;
        int emptyRows = 0;
        Dictionary<string, int> whole = new Dictionary<string, int>();
        int sameRows = 0;
        for (int r = 2; r <= rows; r++)
        {
            StringBuilder key = new StringBuilder();
            bool any = false;
            for (int c = 1; c <= cols; c++) { if (!IsBlank(grid[r, c])) any = true; key.Append(Text(grid[r, c])).Append('\u0001'); }
            if (!any) { emptyRows++; continue; }
            string k = key.ToString();
            if (whole.ContainsKey(k)) sameRows++; else whole[k] = 1;
        }
        lines.Add((rows - 1) + " data row(s) under the header row " + row0 + (emptyRows > 0 ? "; " + emptyRows + " entirely empty row(s)" : "") + (sameRows > 0 ? "; " + sameRows + " row(s) repeat an earlier row exactly" : "") + ".");
```

### 5.3 `helper/OfficeHelper.cs:4609-4615` — LOOK 段（每列"值得注意"的判据）

```csharp
            List<string> odd = new List<string>();
            if (spaced > 0) odd.Add(spaced + " with stray spaces");
            if (numberLike > 0) odd.Add(numberLike + " number(s) written as text");
            if (dateLike > 0) odd.Add(dateLike + " date(s) written as text");
            if (numbers > 0 && texts > 0) odd.Add("mixed numbers and text");
            if (odd.Count > 0) line.Append(". LOOK: " + string.Join(", ", odd.ToArray()));
            lines.Add(line.ToString());
```

### 5.4 `helper/OfficeHelper.cs:4451-4460` — `ParseNumber`（货币/百分号/单位）

```csharp
    static bool ParseNumber(string text, out double number)
    {
        number = 0;
        string s = Tidy(text).Replace(",", "").Replace("，", "").Replace(" ", "");
        System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(s, "^[¥￥$€£]?\\s*([-+]?\\d*\\.?\\d+(?:[eE][-+]?\\d+)?)\\s*(%|％)?\\s*(?:[^\\d\\s\\-.,A-Za-z%％]{1,3}|kg|g|km|cm|mm|m|pcs|k|w)?\\z");
        if (!m.Success) return false;
        if (!double.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out number)) return false;
        if (m.Groups[2].Success) number /= 100;
        return true;
    }
```

### 5.5 `helper/OfficeHelper.cs:5045-5065` — WPS `Range.Replace` 双替换的规避

```csharp
            if (Suite == "wps")
            {
                // WPS replaces again inside what it has just put in ("P" to "P-" gives "P--"): done here instead, on
                // the cells that hold text, the formulas left as they are.
                System.Text.RegularExpressions.Regex pattern = new System.Text.RegularExpressions.Regex(whole ? "^" + System.Text.RegularExpressions.Regex.Escape(find) + "\u0024" : System.Text.RegularExpressions.Regex.Escape(find), matchCase ? System.Text.RegularExpressions.RegexOptions.None : System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                string put = with.Replace("\u0024", "\u0024\u0024");
                object raw = target.Formula;
                object[,] cells = raw as object[,];
                if (cells == null) { string one = raw as string; if (one != null && !one.StartsWith("=", StringComparison.Ordinal)) target.Formula = pattern.Replace(one, put); }
                else
                {
                    for (int r = cells.GetLowerBound(0); r <= cells.GetUpperBound(0); r++)
                        for (int c = cells.GetLowerBound(1); c <= cells.GetUpperBound(1); c++)
                        {
                            string text = cells[r, c] as string;
                            if (text == null || text.StartsWith("=", StringComparison.Ordinal) || !(grid[r, c] is string)) continue;
                            cells[r, c] = pattern.Replace(text, put);
                        }
                    target.Formula = cells;
                }
            }
            else target.Replace(What: find.Replace("~", "~~").Replace("*", "~*").Replace("?", "~?"), Replacement: with, LookAt: whole ? 1 : 2, MatchCase: matchCase);
```

### 5.6 `helper/OfficeHelper.cs:4736-4753` — WPS 组合图设标题被拒后的重试与退路

```csharp
    /// The title of a chart. WPS turns the call down while it is still laying out a chart whose series were just
    /// moved to another type or axis: it is asked again, and then by the other way there is to ask.
    static void Titled(dynamic chart, string title)
    {
        for (int attempt = 0; ; attempt++)
        {
            try
            {
                if (attempt == 2) chart.SetElement(2); else chart.HasTitle = true;
                chart.ChartTitle.Text = title;
                return;
            }
            catch (Exception error)
            {
                Trace("chart title, attempt " + attempt + ": " + error.Message.Trim());
                if (Suite != "wps" || attempt >= 3) throw;
                Thread.Sleep(250);
            }
        }
    }
```

### 5.7 `helper/OfficeHelper.cs:5298-5315` — `copy_range values:true` 为什么必须用 Excel 自己的 Copy

```csharp
        if (type == "copy_range")
        {
            dynamic source = Ref(book, ws, op.Need("range"));
            dynamic target = Ref(book, ws, op.Need("to"));
            if (op.Flag("values", false))
            {
                dynamic into = target.Cells[1, 1].Resize[(int)source.Rows.Count, (int)source.Columns.Count];
                // Writing the values over would make Excel read every text again ("20250301" would turn into a
                // number): its own copy keeps each cell as it is, and only the formulas are then replaced by results.
                source.Copy(into.Cells[1, 1]);
                try { app.CutCopyMode = 0; } catch (Exception) { }
                try
                {
                    dynamic formulas = into.SpecialCells(-4123);
                    foreach (dynamic area in formulas.Areas) area.Value2 = area.Value2;
                }
                catch (COMException) { }
                for (int c = 1; c <= (int)source.Columns.Count; c++) { try { into.Columns[c].ColumnWidth = source.Columns[c].ColumnWidth; } catch (Exception) { } }
                return "values of " + (string)source.Address[false, false] + " copied to " + (string)into.Worksheet.Name + "!" + (string)into.Address[false, false];
            }
```

### 5.8 `helper/OfficeHelper.cs:5397-5411` — 透视表结果"直接取回"

```csharp
            // The result, so that it need not be read again.
            object[,] grid = Grid(made2);
            StringBuilder shown = new StringBuilder();
            int limit = Math.Min(grid.GetLength(0), 40);
            for (int r = 1; r <= limit; r++)
            {
                List<string> cells = new List<string>();
                for (int c = 1; c <= Math.Min(grid.GetLength(1), 14); c++)
                {
                    object v = grid[r, c];
                    cells.Add(IsNumber(v) ? Plain(Math.Round(Convert.ToDouble(v), 2)) : Text(v));
                }
                shown.Append("\n").Append((int)made2.Row + r - 1).Append(": ").Append(string.Join(" | ", cells.ToArray()));
            }
```

### 5.9 `helper/OfficeHelper.cs:9669-9670` — 未知字段点名（护栏）

```csharp
                    string ignored = op.Unread();
                    done.Add(ignored.Length == 0 ? line : line + " — IGNORED field(s) " + ignored + ": " + type + " does not take them, check the operation list");
```

### 5.10 WPS 实测证据（文档层）

`docs/verification.md:688`（测试机）与 `:709`（Excel 在 WPS 下的验收）：

```
## 0.9.6
起因：用户要求适配 WPS Office。测试机：WPS 12.1.0.28505，与微软 Office 并存。
...
- 表格：1204 行测试数据的清洗、补列、4 张透视表、3 张图表、看板，结果与微软 Excel 下一致（销售额合计 2047783，1086 行）。
```

`docs/verification.md:699`（WPS 差异清单里唯一与图表有关的一条）：

```
  - 组合图设置标题被拒绝时重试，再不行改用 `SetElement`。
```
