# WPS 表格 ≠ Excel：实测差异表（v0.1）

本机实测环境：WPS Office **12.1.0.29161**（无微软 Office），`KET.Application`（注册为 `KET.Application.9`，`Excel.Application.12` 也被 WPS 接管指向同一引擎）能跑通 Excel 对象模型的常用子集，但"脾气"不同。下表每条都是 2026-10-24 在这台机器上跑出来的。

| 现象 | 实测表现 | 正确写法 |
|---|---|---|
| `Application.Name` 冒充 Excel | `Name="Microsoft Excel"`、`Version="12.0"`、`Build="29161"`、`Path=...\Kingsoft\WPS Office\...`（与 Writer 冒充 Word 同款） | 判断套件看 `Path`（含 `Kingsoft` 即 WPS）或 `Build`，**不要**看 `Name`/`Version` |
| ProgID 与附着 | `KET.Application` 可 `Marshal.GetActiveObject` 附着运行实例；`ET.Application` **未注册**（不存在这个 ProgID） | 附着/创建一律 `KET.Application`；本 skill 不用 `Excel.Application` 创建实例（装了真微软 Excel 的机器上会开错程序） |
| `xlSheetVisible = -1` 不是 1 | `ws.Visible` 返回 -1（Excel 枚举习惯），按 `==1` 判会把所有表读成隐藏 | 可见判定写 `vis=="1" || vis=="-1"` |
| `SaveAs` 之后工作簿换身份 | `save -As test.xlsx` 后 `FullName` 变成新路径，原来的 `test.csv` 身份消失，按旧路径 `close` 报 `NOT_OPEN` | 另存后一律按**新路径**继续操作；`status` 可查真身 |
| 写块数组必须 0 基、精确尺寸 | `Range.Value2 = new object[nr+1, nc+1]`（按读回的 1 基习惯造）**不报错但数据不落位** | 写：`new object[nr, nc]` 0 基精确尺寸；读回的 `Value2`/`Formula` 才是 1 基 `object[,]` |
| `Range.Text` 多格区域拿不到东西 | 多格 range 的 `Text` 返回空；单格才有显示文本 | before/after 用单格取；块写入的 after 靠 `read`/`render` 复核 |
| 渲染只能走剪贴板 | 表格没有 Word 的 `Pages(n).EnhMetaFileBits`；`Range.CopyPicture(1,2)` → 剪贴板 → GDI+ 存 PNG 可用（本 skill 的 `render`） | 渲染期间剪贴板内容会被清掉（渲染后主动 `Clipboard.Clear()`）；别在用户刚 Ctrl+C 的瞬间渲染 |
| `Image.FromFile` 持文件锁 | 缩放时在 `using` 内 `File.Delete(原png)` 会失败、返回旧路径 | 先释放源图再动文件 |
| 外部修改不进撤销栈 | COM 写入不产生撤销记录，且会清空用户已有的撤销历史 | 写前 `snapshot`；写后比对 before/after + `diff`；重要数据先 `save -As` 备份 |
| 关最后一份工作簿静默杀实例 | 与 Writer 同款：最后一个 workbook 关掉后实例已死但不抛异常 | 操作前校验 `Workbooks.Count`；`close` 在无工作簿时 `Quit()` 并回报进程数 |
| 每次 `CreateInstance` 都是新实例 | 与 Writer 同款 | 先 `GetActiveObject("KET.Application")` 附着，没有再创建 |
| PS 5.1 单元素数组经函数返回被展开 | `return @($parsed.items)` 单元素时调用方拿到标量，helper 报 "items must be an array" | 函数内 `return ,@(...)`；序列化后再由 helper 校验 |
| 落盘即密文 | `save` 出的 xlsx 头部 `17 DA 5F A0 16 33 CD 9A`，`zipfile` 报 BadZipFile；WPS 自己读写渲染一切正常 | 见 [encryption.md](encryption.md)；交付前解密 |
| 公式跨保存持久、`Text` 反映计算值 | 写 `=AVERAGE(...)` 后立刻读 `Text`/`Value2` 即为算好的值；保存重开后公式与值都在 | 校验公式结果直接看 write 返回的 `after` |
| `SetSourceData` 对单列数据方向猜错 | 单列 range 传给 `Chart.SetSourceData` 会把**每一行当一个系列**（135 行=135 条线）；v0.4 实战首例踩中 | 建图用**显式系列**：`SeriesCollection().NewSeries()` 后分别设 `Values`/`XValues`/`Name`；add 返回 `seriesCount` 供核对；两列数据（时间+值）让横轴成为真时间轴 |
| 写入的日期字符串被自动转序列值 | `"2026-10-24 12:00"`、`"10-18 22:00"` 经 Value2 写入即变日期序列（read 回显 46319.5 之类） | 通常正合适（画图横轴可用）；显示需配 `-NumberFormat "yyyy-mm-dd hh:mm"`；要纯文本得前置 `'` 或事后把该格格式设为文本 |

| `Workbook.ActiveWindow` 常为 null | 簿不是当前活动簿时 `wb.ActiveWindow` 返回空，读 `.Caption` 抛异常（v1.11 探针）；`wb.Windows(1).Caption` 读写都正常 | 窗口标题统一走 `wb.Windows(1).Caption`，`ActiveWindow` 只当兜底 | 2026-10-30 v1.11 探针 |
| `Application.StatusBar` 在表格侧**可读写、且真的显示** | 赋值后能读回，且截图里状态栏出现 `AI 正在协作：…` 那行（Writer 侧只写不读也不渲染，别混）；写 `False` 交还 WPS | 接管提示 `hint` 用它显示"AI 正在做什么"；清空写 `false` | 2026-10-30 v1.11 探针 + 截图 |
| 窗口标题可写，且只影响窗口 | `wb.Windows(1).Caption = "〔AI接管中〕 <名>"`（WPS 12 标签页从右截断，标记放前面） 不改单元格、不进撤销栈、不改变 `saved` | `hint` 的接管标记；`close`/`finish` 写回 `wb.Name` 还原 | 2026-10-30 v1.11 探针 |

## 边界与未验证

- 只在 **12.1.0.29161** 上验证过；其它 WPS 版本先 `doctor` + 小工作簿试。
- 未验证：图表操作（v0.1 未做 chart 命令）、数据透视表、多单元格数组公式、合并单元格上的写行为、用户正在编辑同一格时的并发写入（BUSY 场景只映射了拒绝码）。
- `comment -Action delete` 只支持 `-All`（逐条删除留给下版）。
- `Excel.Application` 附着回退未启用：本机两者同引擎无差别，但装了真微软 Excel 的机器上行为未定，所以创建实例只认 `KET.Application`。

## v1.16 补充（2026-11-01 实测）

| 现象 | 实测表现 | 正确写法 |
|---|---|---|
| `Range.CopyPicture` **不含浮动图表** | 5 张图的表 `render -Range A1:H40`，PNG 里一张图都没有（单元格区域正常） | 用 `render -IncludeCharts`：helper 逐张 `ChartObject.CopyPicture` 取位图，按 `Left/Top/Width/Height`（px/点由基准图宽度自校准）合成到单元格图之上 |
| `ChartObject.CopyPicture(1,2)` 可用 | WPS 12 上直接取图表位图成功（`Chart.CopyPicture` 作兜底；`Chart.Export` 未采用——WPS 自己写盘的文件在本机会被透明加密成密文，读不回来） | 取图一律走剪贴板，落盘由 helper 用 System.Drawing 完成 |
| 图表合成要避开"文件占用" | `Image.FromFile(scratch)` 还开着时往同一路径 `Save` 会报 GDI+ 一般性错误 | 单元格基准图先写临时文件，合成结果写最终路径，最后删临时文件 |
| `Workbooks.Add()` + `SaveAs` 造新簿可行 | `new` 落盘后立即 `saved:true`，后续 write/chart/render 全部正常 | 新簿走 `new`；`open` 对不存在的路径报 `NOT_OPEN`（不会替你新建） |
| `SaveAs` 到已存在且**正打开**的路径会失败 | 文件被 WPS 锁住，`File.Delete` 报"正由另一进程使用" | `new -Overwrite` 先 `FindOpen` 判占用，占用时报 `IN_USE`，让调用方先 `close` |

## v1.17 补充（2026-11-02，来自另一台机器的实战反馈）

| 现象 | 实测表现 | 正确写法 |
|---|---|---|
| `wb.ActiveSheet` **不可靠** | 同一分钟内 `open` 报 Sheet2 active、下一条命令看到 Sheet1；窗口显示与 `ActiveSheet` 也会不一致（反馈方 3 表工作簿实测）。它是 `chart`/`render`/`write` 无 `-Sheet` 时的兜底目标 | v1.17 起统一走 **`wb.Windows(1).ActiveSheet`**，取不到才退回 `wb.ActiveSheet`；要确定性就永远显式带 `-Sheet` |
| 包装层**静默丢弃参数** | v1.16 及更早：`chart` 分支没转发 `$Sheet`，图表永远建到 ActiveSheet，还返回 `ok:true`；`chart list/delete` 同样绑 ActiveSheet ⇒ 落错表后无法自救（反馈方为此耗 ~6 min，最后用裸 COM 删图） | v1.17：chart 全部 action 转发 `-Sheet`；新增 `chart -Action list -All` 跨表扫描；并对所有命令做**参数消费校验**（stderr `WARN ignored parameter ...`） |
| cmd.exe 层**改写参数** | `woc.cmd excel fill -Formula "=A15^1.5"` → 子进程收到 `A151.5`；`"=1+1"` → `1+1`。`^` 是 cmd 转义符，以 `=` 开头的裸 token 也会丢 `=`。**batch 内部无法修复**（拿到时已坏） | 用 **`scripts\woc.ps1`**（PowerShell→PowerShell，不经 cmd 解析）；或直接 `powershell -Mta -File excel.ps1`；公式内容走 `-Items` JSON |
| PowerShell **双引号插值**吃 `$` | `-Formula "=$B$5*2*SQRT(2*$B$6)*B15"` 里 `$B`/`$6` 被当变量 → `=*5*2*SQRT(2**6)*B15`，WPS 报 `0x8FE3000B`（不是语法错，是 HRESULT，很难看出） | 含 `$` 的公式用**单引号** `-Formula '=$B$5*...'` |
| `chart` 建完**无法移动/缩放**（v1.16 及更早） | 尺寸写死 `Add(left, top, 480, 300)`；125% 缩放下 480 pt ≈ 800 px，放在右侧会被窗口裁掉，只能删了重建（反馈方重建 2 轮 4 张图） | v1.17：`-Width/-Height` + `-Action move -Index n [-Anchor X] [-Width] [-Height]`；实操建议图表默认放内容下方，别放右侧 |
| `chart` **只能单系列**（v1.16 及更早） | 取「第一列=x、最后一列=值」，`-Range A14:E25` 只画一列 | v1.17：第 2..N 列各一条系列（表头作系列名、多系列自动出图例），并支持 `-XTitle/-YTitle` |
| 没有「激活工作表」原语（v1.16 及更早） | `sheet` 只有 list/add/rename/delete，agent 只能裸 COM `Worksheets.Item(n).Activate()`（违反铁律 0） | v1.17：`sheet -Action activate -Name X` |

## v1.18 补充（2026-11-02，三端参数校验 + 颜色通道 + 窗口归属）

| 现象 | 实测表现 | 正确写法 |
|---|---|---|
| **`-Color` / `-FillColor` 红蓝互换**（v0.1–v1.17 一直存在） | COM 的 `Font.Color` / `Interior.Color` 低 24 位是 **BGR**（`R + G*256 + B*65536`），而旧 `ColorFrom` 返回的是 System.Drawing 的 ARGB 掩码（RGB），于是 `-Color "#FF0000"` 渲染成**蓝色**、`"#0000FF"` 渲染成**红色**。2026-11-02 用红/蓝/绿三格对照图抓到的；Word 侧 `WpsHelper.cs` 一直是 `r + g*256 + b*65536`（正确），只有 Excel 侧写反 | v1.18 已改为 `r \| (g<<8) \| (b<<16)`；写颜色前想核对可读回 `Range.Font.Color`（纯红应为 **255**，纯蓝 **16711680**） |
| **WPS 12 的窗口属于 `wps.exe`，不是 `et.exe`** | 另一台机器实测（1536×864 / 125%）：`et.exe` 存在但 `MainWindowHandle = 0`；真正带窗口的是 `wps.exe`。想用 `PrintWindow` / `GetWindowRect` 抓整窗必须先按 `wps.exe` 找窗口 | 抓窗：`Get-Process wps | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1`；但**优先用 `render -IncludeCharts`**，不要截屏 |
| 窗口尺寸 ≠ 屏幕尺寸 | 125% 缩放下 480 pt ≈ 800 px；1536 px 宽的屏，A–F 列占 ~1030 px，右侧只剩 ~500 px ⇒ 按 `-Anchor G4` 建 480×300 的图会被窗口右缘裁掉 | 图表默认放**内容下方**（A 列起）；或用 `chart -Width/-Height` 缩到 ≤400 px 再放右侧 |
| 参数被静默丢弃 | 包装层把 `$PSBoundParameters` 转成 JSON，绑定了却没转发进 request 的参数会**无声消失**（`chart -Sheet` 就这样把图建到了活动表） | v1.18 起三端共用 `scripts\param-check.ps1`：未消费参数在 **stderr** 打 `WARN ignored parameter -X for command "Y"`；`WOC_NO_PARAMCHECK=1` 可关 |

## v1.20 补充（2026-11-02，数据操作 / 透视表 / 成品感 / 图表次轴实测；WPS 12.1.0.29161）

| 现象 | 实测表现 | 正确写法 |
|---|---|---|
| **`Range.Columns(c).NumberFormat` 整列格式不一致时返回 Null** | 表头是 `General`、数据格是 `@`/`yyyy-mm-dd` 时整列取回空，于是"按显示格式判日期"整列判定失败、日期被当序列号（45717）返回 | 先取整列；为 Null 就抽样第 2 / 中间 / 末格 `Cells(r,c).NumberFormat`（每列 3 次 COM），过半数命中 `[ymd年月日]` 且非 `0.0`/`General` 才算日期列 |
| **单列块写必须含表头格** | `Columns(c).Value2 = new object[rows,1]` 若 `[0,0]` 留空，**表头被写空**（实测 `clean`/`fillblanks` 曾把 C1 "level" 清掉） | 0 基数块的第一格填回 `grid[1,c]`；每次冒烟都回读 A1/表头 |
| **`PivotCaches().Create` 只能两参** | `Create(1, source)` + `cache.CreatePivotTable(dest, name)` 两参可行；3/4 参报错。结果读 `TableRange1`（实测 wai zhou 合计 **5100** 正确） | 不写具名 4 参；默认落新表 A3；`AddDataField(field, caption, -4157)`（count -4112 / avg -4106 / max -4136 / min -4139） |
| 透视表源区域有空表头即拒 | 点名列字母抛 `BAD_ARGS`（透视表要求每列都有名字） | 先补表头；`DataRange.Cells(1,1).Group(Periods:[…])` 按月/季/年分组失败只记 NOTE 不抛错 |
| **`ws.Sort` 排序可用** | `SortFields.Clear/Add(Key,SortOn:0,Order,DataOption:0)` + `SetRange` + `Header=1` + `Apply()` 实测降序生效；排序**带着公式与格式一起搬**（`=D2*2` 跟着行走） | 用 `ws.Sort`，不要"读出来排完写回"（后者会破坏公式联动） |
| **`AutoFilter` + `SpecialCells(12)` 可用** | `AutoFilter(Field, Criteria1: string[], Operator: 7)` 生效；`SpecialCells(12).Cells.Count - 1` 数可见行正确（实测 4/20） | 筛选是**隐藏**不是删除；`filter -Action clear`（`ShowAllData` + `AutoFilterMode=false`）复原 |
| **`ListObjects.Add` 可用** | `ListObjects.Add(1, range, Type.Missing, 1)` 建表成功；`table.Name`/`TableStyle` 可写；重复调用需 `app.Intersect` 判"这块已是表" | 表名唯一且不含空格；只改样式时不要重建 |
| **`Windows(1)` 冻结窗格可用** | `ws.Activate()` → 关旧冻结/拆分 → `ScrollRow/Column=1` → `SplitRow/SplitColumn` → `FreezePanes`，回读 `splitRow=1, freezePanes=true` | 三道重试（窗口非活动时 WPS 拒绝）；`-Action remove` 置 `FreezePanes=false, Split=false` |
| **`FormatConditions` 便捷方法全部可用** | `Add(1,5,"=1000")`、`AddColorScale(3)`、`AddDatabar()`、`AddTop10()`、`AddUniqueValues()`、`Type:10`(blank)、`between`(Type 1 + Formula2) 逐条实测 `count` 递增 1→6 | 阈值规则的 `Formula1` 要写成 `="文本"` 或 `=数字`；`clear` 走 `FormatConditions.Delete()` |
| **`Validation.Add(3,1,1,"a,b,c")` 可用** | 下拉列表落地；列表值是逗号拼接，**>255 字符或值本身含逗号会坏** | 本命令已拦 255 上限并报错；值含逗号属未做缺口 |
| **组合图 + 次坐标轴可用，`Axes(2,2)` 首次可能为 Null** | 先建完所有系列，再逐条 `SeriesCollection(k).ChartType=65` / `AxisGroup=2`；`chart.Axes(2,2)` 第一次取回 Null，重试 2 次可得；双轴、右轴刻度 0–160、图例位置、数据标签实测全部生效 | 标题按 dsh-office `4736-4753` 重试：第 3 次（attempt==2）改 `SetElement(2)` 再设 `ChartTitle.Text` |
| `render -Overview` 拼图可行 | 每表 `UsedRange.CopyPicture` 出临时 PNG → GDI+ 拼 2 列 + 表名标签；两表实测 **1260×595** | 临时子目录 `out\overview-parts` 拼完即删；缩略图按 `-Width/2` 算格宽 |
| **关掉同一实例里最后一份工作簿会连坐** | 冒烟时先 `close` 无关簿，随后对测试簿的每条命令都报 `DOC_CLOSED`（整个实例已退出） | 同一实例只留一份在用工作簿；先 `save`，或随时 `open` 重新打开 |
| 经 cmd/PowerShell 传**空字符串**会变成"缺参数" | `excel write -Cell C5 -Value ""` 报 `Missing an argument for parameter 'Value'` | 清空格改用 `write -Items`（JSON 里 `"value":""`） |
| **表（ListObject）覆盖区域内 `Range.AutoFilter` 报 `E_ACCESSDENIED`（0x80070005）** | 先 `table` 建了 A1:F21，再 `filter -Column site -InValues …` → `UnauthorizedAccessException: 拒绝访问`；同一区域的 `sort`/`dedupe`/`deletewhere`/`fillblanks` 全部正常（实测 2026-11-02） | 先筛选再建表，或先删表；helper 已把该异常翻成 `COM_ERROR` 并说明原因（不要在表内反复试 filter） |

## v1.31 补充（2026-11-02，内存块替换 / 透视表日期分组实测；WPS 12.1.0.29161）

| 现象 | 实测表现 | 正确写法 |
|---|---|---|
| **`Range.Replace` 会把刚替换进去的文本再替换一遍** | `P` → `P-` 落到 `P--`（同一调用内二次命中）。dsh-office `helper/OfficeHelper.cs:5045-5065` 实测同源；本机 WPS 12.1.0.29161 一样 | **不要用 `Range.Replace`**。`replace` 走"块读一次 → 本地逐格替换 → 块写一次"：遍历原始字符串、游标移到插入文本**之后**，`PP → P-P-`（单遍），并顺带给出 `samples[].before/after`、`byColumn`、`-DryRun` 清单 |
| 逐格写回会触发整簿重算 | 几百格一次一格的 `Value2` 写入，WPS 每写一次重算一次 | 写回期间 `PerfOn()`（`Calculation=-4135` + `ScreenUpdating=false`），一次性重算；连续列合并成一行 run 用 1×N 块写（0 基精确尺寸） |
| **数字列里的"匹配"不能改成文本** | `-Find "0" -Replace "9"` 打 `10/20/30` 这类数字列时，若按文本重写，整列会从数字变文本（后续求和、排序、图表全变味） | v1.31 起**只替换字符串格**：数字/日期/布尔/错误格只计数（`numericMatchesLeftAlone` + `numericNote`），不写 |
| 公式格不能写值 | 块写会把 `=D2*2` 变成静态值 | 一次读 `Range.Formula`（1 次 COM）判断哪些格以 `=` 开头，这些格跳过并计入 `formulaCellsSkipped`（实测 5/5 跳过） |
| **WPS 原生支持透视表按日期分组** | `PivotField.DataRange.Cells(1,1).Group(Start:true,End:true,Periods:[false×4,月,季,年])` **可用**：24 个月数据 → 12 / 4 / 2 个行项，合计 8580 不变（月 110…1320、季 660/1650/2640/3630、年 780/7800，均与手算一致） | 直接用原生 `-Group "日期:month\|quarter\|year"`，**不需要**辅助列降级；`Periods` 是 7 元素布尔数组，`xlMonth=4 / xlQuarter=3 / xlYear=2` 的位置分别对应第 5/6/7 位 |
| 文本日期分组会抛中文异常 | `@` 格式的 `2024-01-15` 走 `Group()` 抛 `选定区域不能分组。`（连异常都不抛的混合列也可能静默不分组） | 分组前后各读一次 `PivotTable.RowRange.Rows.Count`，行项数没变少即判 `groupMode:"failed"` 并给出辅助列配方；`-Group` 省略时 `groupMode:"none"` |
| 辅助列（降级路线）实测可用 | `=YEAR(A2)` 逐行公式列做 `-Rows` → 2024 780 / 2025 7800 / 总计 8580，与手算一致 | 辅助列 `年份 =YEAR(日期)`、`月份 =TEXT(日期,"yyyy-mm")`、`季度 ="Q"&ROUNDUP(MONTH(日期)/3,0)`；**它是透视表的数据源，建完不能删**；多年度"季"要年+季两列 |
| `addcol`/`fill` 的公式是**整块同一份** | `body.Formula = "=YEAR(A2)"` 对整列写同一串，不会像手工填充那样按行相对引用（实测：整列都是 `=YEAR(A2)`） | 需要逐行不同公式时走 `write -Items`（每格一条 `formula`），或先写 `=YEAR(A2)` 再在 WPS 里填充；`replace`/`pivot` 的辅助列配方要按这条来 |

