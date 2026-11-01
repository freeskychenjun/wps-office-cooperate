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
