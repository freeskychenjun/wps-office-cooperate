# skill-excel —— v1.31 Excel 补丁片段（`replace` 内存块替换 + `pivot -Group` 实测定案）

> 用途：给 `SKILL.md` 的**补丁片段**（本次不改 SKILL.md，由 Lead 合并）。
> 对应实现：`scripts\excel.ps1`（ValidateSet + param + switch）、`scripts\helper\excel\EtReplace.cs`（新命令，随包编译）、
> `scripts\helper\excel\EtHelper.cs`（`Dispatch` 加 `case "replace"`）、`scripts\helper\excel\EtData.cs`（`pivot` 的 `groupMode` 与回读校验）。
> 实测环境：WPS 12.1.0.29161（`KET.Application`，`et.exe` 单实例），2026-11-02 夜。
> 新增实测事实见 `reference\excel-quirks.md` 的「v1.31 补充」节。

## 1. SKILL.md 命令速查表：新增 / 修改这些行

| 命令 | 一句话 | 常用参数 |
|---|---|---|
| `replace` | **内存块替换**（一次块读 → 本地逐格替换 → 只写真正变化的格），绝不走 `Range.Replace` | `-Find "旧" [-Replace "新"]（省略 = 删除）[-Sheet] [-Range A1:F200] [-Column 表头\|字母] [-Whole] [-MatchCase] [-Regex] [-DryRun] [-Max N] [-All]` |

`pivot` 一行的补充：`-Group "日期:month\|quarter\|year"` 现在返回 **`groupMode`**（`native` / `failed` / `none`），失败会明说原因并给出降级配方，不再只是往 `notes` 里丢一句。

## 2. `replace` 用法细节

```powershell
# 先看清单（不写文档）
excel replace -Doc 台账.xlsx -Sheet 明细 -Range A1:F200 -Find "表1" -Replace "表一" -DryRun
# 实际替换；只给 -Find 即"删除匹配"
excel replace -Doc 台账.xlsx -Column 编码 -Find "P" -Replace "P-"
excel replace -Doc 台账.xlsx -Column 备注 -Find "\s+$" -Regex
# 整格相等才命中 + 区分大小写
excel replace -Doc 台账.xlsx -Column 单位 -Find "水利局" -Replace "水利水电勘测设计院" -Whole -MatchCase
```

- **只返回该返回的**：`scanned`（扫了多少格）、`changed`、`byColumn`（每列命中数，键是表头，同名列自动补列字母）、
  `samples:[{cell,before,after}]`（默认 30 条，`-All` 全给）、`formulaCellsSkipped`、`nonTextCellsSkipped`、
  `numericMatchesLeftAlone`（数字/日期格**也含**该串但按规矩不写，附 `numericNote`）、`truncated`（`-Max` 截断）、
  `dryRun`、`applied`，非 dry-run 时还带 `preSnapshot`（可 `diff -Since` / `restore -Since`）。
- **`-DryRun` 与实跑同一套判据**：dry-run 报的 `changed` 就是实跑会改的格数（同一段代码只差最后写回）。
- **公式格永不写**：整块读一次 `Range.Formula` 判断哪些格是公式（`=...` 开头），这些格只计数不参与替换。
- **数字/日期格不改成文本**：`-Find "0"` 想改 `10/20/30` 这种数字列时，`changed:0` + `numericMatchesLeftAlone:5` 并给出说明——
  把数字列悄悄变成文本列是比"没改到"更坏的结局。
- **中文/正则**：`-Regex` 走 .NET 正则（`$1` 反向引用可用），模式非法报 `BAD_ARGS` 并提示去掉 `-Regex`；
  正则与字面替换都是**单次扫描**：替换进去的文本不会再被扫一遍。
- **`-Max` 只在显式给出时生效**（`search` 的 `-Max` 默认 200，若沿用默认值会把大替换悄悄截断）；截断时 `truncated:true`。
- **块范围**：`-Range` 省略 = `UsedRange`；`-Range "表2!A1:F9"` 可跨表；`-Column` 省略 = 范围内全部列。
- **幂等**：规则本身不自匹配时（`表1→表一`、整格 `abc→X`），跑第二次 `changed:0`（实测）。像 `P→P-` 这种
  **替换串包含查找串**的规则，第二次当然还会命中——这是规则的性质，不是实现问题；关键是**一次调用只会插一遍**（`PP→P-P-`，不是 `P--`）。

### 2.1 为什么不用 `Range.Replace`（关键）

WPS 会在**刚替换进去的文本上再替换一次**：`P` → `P-` 得到 `P--`（dsh-office `helper/OfficeHelper.cs:5045-5065` 实测，本机 WPS 12.1.0.29161 同源）。原因是我们无法干预它内部的扫描推进方式，而且：

1. 它**逐格往返**，一次调用扫整张表，无法先给清单再改（没有 dry-run 语义）；
2. 它**不区分公式格**，返回里也拿不到 before/after（只能事后 `read` 全表对比）；
3. 它**没有截断与计数**，改了多少、改了哪列全靠猜；
4. 中文/大小写行为依赖 WPS 的查找参数，不可回读验证。

内存块替换天然免疫第 1 条：我们是**一次遍历原始字符串**、游标移到插入文本**之后**再继续，
所以刚写进去的文本永远不会被二次扫描（`PP` → `P-P-`，实测）。其余三条也在返回里直接给全。

## 3. `pivot -Group` 实测定案（v1.31）

**结论：WPS 12.1.0.29161 原生支持按日期分组**，走 `PivotField.DataRange.Cells(1,1).Group(Start:true, End:true, Periods:[...])`
（`xlMonth=4`、`xlQuarter=3`、`xlYear=2`），不需要辅助列降级。24 行（2024-01 ~ 2025-12，流量 10..120 / 100..1200）实测：

| `-Group "日期:..."` | 结果网格 | 手算核对 |
|---|---|---|
| `month` | 1月 110 … 12月 1320，总计 8580（12 行） | 每月 = 10m + 100m = 110m ✓ |
| `quarter` | 第一季 660、第二季 1650、第三季 2640、第四季 3630，总计 8580 | 110×(1+2+3)=660 ✓ |
| `year` | 2024年 780、2025年 7800，总计 8580 | 10×78=780、100×78=7800 ✓ |
| 无 `-Group`，`-Rows "YEAR"`（辅助列 `=YEAR(A2)`） | 2024 780、2025 7800，总计 8580 | 同上 ✓（降级路线可用） |

**但"没报错"不等于"分好组了"**：对**文本日期**（`@` 格式里的 `2024-01-15`）WPS 直接抛
`选定区域不能分组。`，而某些混合列可能连异常都不抛。所以 v1.31 起一律**回读校验**：
分组前后各取一次 `PivotTable.RowRange.Rows.Count`，行项数没变少就判 `groupMode:"failed"`。

返回值：`groupMode`（`none`/`native`/`failed`）+ `group:{field,by,mode}`；判 `failed` 时 `notes` 里给出降级配方：

```powershell
# 文本日期 → 先加辅助列，再按辅助列做透视（实测可用）
excel addcol -Doc 台账.xlsx -Header 月份 -Formula '=TEXT(A2,"yyyy-mm")' -NumberFormat '@'
excel pivot  -Doc 台账.xlsx -Range 'A1:D25' -Rows 月份 -Agg '流量:sum'
```

辅助列口径：`年份 =YEAR(日期)`、`月份 =TEXT(日期,"yyyy-mm")`、`季度 ="Q"&ROUNDUP(MONTH(日期)/3,0)`。
**辅助列是透视表的数据源，建完不能删**（删了透视表就报错）；多年度数据要"年+季"两列才能避免不同年份的 Q1 合并。

## 4. 冒烟记录（2026-11-02，%TEMP% 自建簿）

`rp.xlsx`（S1：A 名称 / B 编码 / C 表号 / D 数量 / E 合计=`=D2*2` / F 大小写），6 行：

| 步骤 | 命令要点 | 结果 |
|---|---|---|
| 1 | `-Find P -Replace P- -Column 编码 -MatchCase -DryRun` | `changed:4`，样例含 `B5 PP → P-P-`（单遍，无 `P--`） |
| 2 | 同上去掉 `-DryRun` | `applied:4`，回读 `B2 P-` / `B3 P-05` / `B4 XP-` / `B5 P-P-`；E2:E6 仍是 `=D2*2`…，值 20/40/60/80/100 未被写 |
| 3 | `-Find 表1 -Replace 表一 -Column 表号` 连跑两次 | 第一次 `changed:3`（C2/C4/C5，其中 `第表1项 → 第表一项`），第二次 **`changed:0`** |
| 4 | `-Find abc -Replace X -Column 大小写 -Whole -MatchCase` | `changed:2`（只有 F3/F5），再跑 `changed:0` |
| 5 | `-Find "^(甲\|乙)" -Replace '$1项' -Column 名称 -Regex` | `changed:2`（`甲P→甲项P`、`乙P→乙项P`） |
| 6 | `-Find 0 -Replace 9 -Range D2:D6` | `changed:0`，`numericMatchesLeftAlone:5` + `numericNote` |
| 7 | `-Find D -Replace Q -Column 合计 -DryRun` | `changed:0`，`formulaCellsSkipped:5` |
| 8 | `-Find 丙`（省略 `-Replace`） | `changed:1`，`A4 "丙" → ""` |
| 9 | `-Find 表2 -Replace 表二 -Column 表号 -Max 1` | `changed:1` + `truncated:true` |

`pv2.xlsx`（A 日期 / B 流量 + C `=YEAR(A2)` 辅助列 + D 文本日期 / E 序号）：`pivot` 三种原生分组与辅助列路线逐项核对，
数字见上表；文本日期列 `-Group TDATE:month` 得到 `groupMode:"failed"` + `选定区域不能分组。` + 配方。
