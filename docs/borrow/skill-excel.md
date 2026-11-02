# skill-excel —— v1.20 Excel 补丁片段（数据操作 / 透视表 / 成品感 / 图表次轴 / 概况增强）

> 用途：给 `SKILL.md` 的**补丁片段**（本次不改 SKILL.md，由 Lead 合并）。
> 对应实现：`scripts\excel.ps1`（ValidateSet + param + switch）、`scripts\helper\excel\EtHelper.cs`（概况/图表/render）、
> `scripts\helper\excel\EtData.cs`（新命令，随包编译）。
> 实测环境：WPS 12.1.0.29161，本机 `KET.Application`；新增实测事实见 `reference\excel-quirks.md` 的「v1.20 补充」节。
> 全部命令都是"**块读一次 → 本地算 → 块写一次**"，有公式的格永不写值，返回值里带改了/删了多少格。

## 1. SKILL.md 命令速查表：新增这些行

| 命令 | 一句话 | 常用参数 |
|---|---|---|
| `clean` | 清洗一列/多列：去多余空格、去指定字样、大小写、取值改名、文本转数字/日期 | `-Column 表头\|C` `-Columns "A,C"` `-To number\|date\|text` `-Casing upper\|lower\|proper` `-Remove "元,["` `-Map '{"是":"Y"}'` `-Year 2025` `-NoSpaces` `-NumberFormat` |
| `fillblanks` | 补空单元格：值 / 上一格 / 均值 / 中位数 / 众数 | `-Column 水位` `-Columns "B,C"` `-With 0\|above\|mean\|median\|mode` |
| `dedupe` | 去重复行（按键列或整行，保留每组第一行） | `-Columns "站码,日期"`（省略 = 整行）`-Range` |
| `deletewhere` | 按条件删行（自底向上合并区间一次删） | `-Column 流量 -Cond "<" -Value 0`；`-Cond blank` 省略 `-Column` = 删整行为空的空行；`-Cond in -InValues "a,b"` |
| `sort` | 就地表排序（公式与格式随行走） | `-By "流量:desc,站码"` 或 `-Column 流量 -Order desc` `-NoHeader` |
| `filter` | 自动筛选（隐藏不删），并回报可见行数 | `-Column 站点 -InValues "外洲,李家渡"`；`-Column 流量 -Criteria ">1000"`；`-Action clear` 复原 |
| `copy` | 复制区域（公式+格式；`-Values` 固化数值且不重解析文本） | `-Range A1:F21 -To "备份!A1"` `-Values` |
| `addcol` | 加一列（同名表头 = 重填；`-Values` 固化；回报错误格） | `-Header 流量2 -Formula "=D2*2"` `-Value 0` `-After 流量` `-NumberFormat "0.0"` `-Values` |
| `pivot` | 透视表，**结果网格随返回值一起回**，不用再 read | `-Range A1:F21 -Rows 站点 -Columns 月份 -Agg "流量:sum,水位:average" -Filters 部门 -Group "日期:month" -SortBy "合计：流量:desc" -Name 透视表 -To "另一个表!A3"` |
| `freeze` | 冻结窗格 / 取消 | `-Cell A2`（冻结其上方与左侧）；`-Action remove` 取消 |
| `cformat` | 条件格式：阈值着色 / 色阶 / 数据条 / 前 N / 重复值 / 空值 / 公式 | `-Range D2:D500 -Rule ">" -Value 1000 -FillColor "#FFC7CE"`；`-Rule colorscale\|databar\|top\|bottom\|duplicate\|unique\|blank\|formula\|between` `-Rule clear` |
| `validate` | 数据验证（下拉列表） | `-Range H2:H200 -List "是,否,待定"`（≤255 字符、值内不能含逗号） |
| `table` | 建成 ListObject 表（纹行 + 筛选箭头 + 列名可当名字用） | `-Range A1:F21 -Name 站点表 -Style TableStyleMedium2`（已是表则只改样式/名字） |
| `render -Overview` | 把**整本工作簿的每张表**拼成一张缩略图 | `render -Doc x.xlsx -Overview -Out <目录> [-Width 1200]`（默认最多 12 张表；`-Sheet "A,B"` 指定） |

## 2. 用法细节（值得写进 SKILL.md 正文的）

### 2.1 `read -Profile` 的 LOOK 段（概况增强）

大块（>60 行或 >15 列，或 `-Profile`）返回的 `profile` 现在多了：

- 顶层：`dataRows`（表头下数据行数）、`emptyRows`（整行全空的行数）、`duplicateRows`（与前面某行**完全重复**的行数）。
- 每列：`letter`/`header`、`numbers`/`dates`/`texts`/`blanks`/`errors`、
  `min`/`median`/`mean`/`max`/`sum`（数值列）、`dateFrom`/`dateTo`（日期列）、
  `farAbove:{count,over}`（离群：`q3+3·IQR` 之外且命中 ≤5%，且数值 ≥20 个才判）、
  `distinct`/`topValues`/`repeats`、
  **`look:[...]`** —— 这一列"值得先修"的判据：`n with stray spaces`、`n number(s) written as text`、
  `n date(s) written as text`、`mixed numbers and text`、`n error value(s)`。
- 小表（默认读取路径）多返回 `dateColumns:["E"]`，且该列的值已从序列号转成 `yyyy-MM-dd`（带时间则 `HH:mm`）字符串。
- **错误值**不再是 `-2146826281` 这类负整数，统一映射成 `#DIV/0!`/`#N/A`/`#NAME?`/`#NULL!`/`#NUM!`/`#REF!`/`#VALUE!`。

口径要点：判日期看**列的显示格式**（`NumberFormat`），不看值；`LOOK` 的阈值都取"少数派"，
几行小表不会乱报警。**拿到概况先看 `look`，再决定跑哪条 `clean`。**

### 2.2 `clean` 的两种用途

1. **只去脏东西**（`-Column`/`-Columns` 省略 = 全部列，只在"没有 to/map/case/remove"时）：
   `excel clean -Doc x.xlsx -Sheet 数据` → 全部文本列 trim（全角空格/制表符也清）。
2. **转类型**（必须 `-Column`/`-Columns`）：
   - `-To number`：`¥1,299` / `89元` / `10%` / `12 件` → 数值；
   - `-To date`：`2025/3/1`、`2025.03.01`、`20250301`、`2025年3月1日`、`3月1日`（**年份取整列众数**，可用 `-Year` 指定）→ 日期，并自动设 `yyyy-mm-dd`；
   - `-To text`：转文本并设 `@` 格式。
   返回每列 `changed`/`notReadable`（读不成的格数与前 5 个地址+原文）/`valuesNow`（≤16 个取值时回列现状，自检用）。
   有公式的列只写非公式格。

### 2.3 `pivot` 一次性拿到结果

```powershell
& scripts\woc.ps1 excel pivot -Doc $doc -Sheet 数据 -Range A1:F2000 `
    -Rows "站点" -Columns "月份" -Agg "流量:sum,水位:average" `
    -Group "日期:month" -SortBy "合计：流量:desc" -Name 站点月汇总
```
返回 `{sheet, name, range, rows:[[行号, 值...],...], valueFields, notes}` —— `rows` 就是透视结果网格
（前 40 行 × 14 列，整数字段名随附行号），**不要再 read 一次**。源区域每列都必须有表头（否则报错点名）。
`-Group` 失败不抛错，写进 `notes`（该列必须全是真日期）。

### 2.4 图表：组合图 / 次坐标轴 / 数据标签 / 图例 / 轴刻度

```powershell
& scripts\woc.ps1 excel chart -Doc $doc -Sheet 数据 -Action add -Type combo `
    -Range "A1:C25" -Title "过程线+雨量" -XTitle "时间" -YTitle "水位(m)" -Y2Title "雨量(mm)" `
    -Secondary "雨量" -Labels value -Legend bottom -AxisMin 0 -NumberFormat "0.0" `
    -Anchor G2 -Width 420 -Height 260
```
- `-Type combo`：最后一条系列自动改折线（其余保持柱状）；`-Secondary "名字或序号"`（逗号可多个）把系列挪到**次坐标轴**。
- 返回 `occupancy:"G22:N41"`（占地区域）与 `overlaps:["其它图名 (区域)"]`（与已有图重叠时硬提示，附移动命令）。
- 标题在 WPS 下可能被拒（组合图刚改完系列类型时）：helper 已按 dsh-office 的口径**重试 3 次并退到 `SetElement`**，无需调用方处理。
- `seriesCount:0` 时返回 `note:"NOTHING is plotted"`。

### 2.5 `render -Overview`

每张工作表的 `UsedRange` 各出一张缩略图，2 列排布、上方写表名，合成一张 PNG，返回
`{overview: 路径, sheets:[{sheet,range}], width, height}`。适合"这簿有几张表、都长什么样"的一次性核对；
要看清单张表仍用 `render -Range`（配 `-IncludeCharts` 可把浮动图表一起拍进去）。

## 3. 坑（新命令相关）

- **块写必须 0 基、尺寸精确**，且**单列块要含表头格**（数组第一格留给表头），否则表头被写空（实测踩过）。
- **判日期靠 `NumberFormat`，而整列格式不一致时 `NumberFormat` 返回 Null** → 必须回退到抽样单元格（本实现已做）。
- `deletewhere`/`dedupe` 是**真的删行**，行号会变；删前先 `snapshot`，需要保留时先 `save -As` 备份。
- `filter` 只隐藏不删除；`table` 建成后会接管该区域的自动筛选。
- `pivot` 会**新建一张工作表**（默认名 `透视表`，落在 A3）；指定 `-To` 则落现有表。
- 中文参数值（表头名、`-List` 值等）走 PowerShell 命令行时注意 `cmd.exe` 会改坏引号与 `^`：
  统一用 `scripts\woc.ps1`（PS→PS），复杂结构走 `-Items`/`-Map` 的 JSON。
- 新增参数名与既有参数冲突的地方：**大小写用 `-Casing`**（`-Case` 已是搜索的大小写开关）、
  **条件用 `-Cond`**（避免 `-Is`）、**in 列表用 `-InValues`**、**图表轴上下限用 `-AxisMin/-AxisMax`**（`-Max` 已存在）。

## 4. 本次未做 / 已知边界（截至 2026-11-02）

- **`replace` 命令未做**：`clean -Remove` 覆盖"去掉指定字样"，但"查找替换 + 命中计数"没有单独命令；
  WPS 双替换坑（`P→P-` 得 `P--`）的规避写法已在 `reference\excel-quirks.md` 记录，将来做时按内存替换实现，不用 `Range.Replace`。
- **`copy` 只做同簿**（含 `Sheet2!A1` 形式），跨工作簿复制未支持；`-Values` 走 `SpecialCells(-4123)`，WPS 上已实测可用。
- **`validate` 只做逗号列表**：值本身含逗号、或想引用单元格区域做下拉源，未做（>255 字符已拦截报错）。
- **透视表 `-Group`（月/季/年分组）未实测**（代码按 dsh-office 口径写了，失败只记 `notes` 不抛错）；`-Filters`（筛页字段）已摆位但未做取值筛选。
- **图表次坐标轴的刻度/数字格式不单独设置**（`-AxisMin/-AxisMax/-NumberFormat` 只作用主轴）；`-Labels` 只支持真/假/`value`/`percent`。
- **表（ListObject）内不能 `filter`**：`Range.AutoFilter` 在表区域报 `E_ACCESSDENIED`（见 quirks）；sort/dedupe/deletewhere/fillblanks 在表内正常。
- **`clean` 的 `-Map` 只接受内联 JSON**（如 `-Map '{"Y":"是"}'`），没有 map 文件参数。
- **概况的"取值写法混杂"取样**（dsh-office `Sample()` 那 5 种写法）未移植：`look` 里只给计数，不给样例值。
