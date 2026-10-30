# 验收记录

## v0.1（2026-10-24，本机 WPS 12.1.0.29161，无微软 Office）

环境：`KET.Application.9` 注册；helper 单文件 C#5 经 `csc` 编译缓存；全流程经 `et.ps1` 走 UTF-8 JSON 文件往返。

| # | 用例 | 结果 |
|---|---|---|
| 1 | `doctor`（无实例） | ✓ wpsEt installed、excel(接管)、processes 0、undoNote |
| 2 | `open` CSV（新建实例）+ `read` 概览 | ✓ usedRange A1:D5 |
| 3 | `read -Sheet test` 全量 | ✓ 值/数值类型正确（1850.5、2020 为数字） |
| 4 | `write` 含 expect 锚不匹配 | ✓ ANCHOR_MISSING，整批零改动 |
| 5 | `write` 标量+公式+2x2 块 | ✓ E2=1735.25、E3=1882.875（=AVERAGE 公式即算即得） |
| 6 | `comment add` + `list` | ✓ 锚定 D5，作者=登录用户 |
| 7 | `render` used range | ✓ PNG 900x201，非白 27.4%（经剪贴板 CopyPicture） |
| 8 | `format`（Bold+FillColor+HAlign+Borders） | ✓ applied 4 项 |
| 9 | `autofit` | ✓ |
| 10 | `snapshot` → 改 C2 → `diff` | ✓ 抓到直改 C2 **与公式联动 E2**（before→after） |
| 11 | `sheet add summary` → 写 `'summary'!A1` | ✓ 引用写法可用 |
| 12 | `calc` / `save -As test.xlsx` | ✓；另存后身份变更为 xlsx（quirk 已记） |
| 13 | 复开 xlsx 读值 | ✓ 公式与值跨保存持久；文件头 `17DA5FA01633CD9A`（密文，COM 不受影响） |
| 14 | `write` 71 行单列块 | ✓（修 0 基数组 bug 后）；`read` 自动转列概况：71 数字/1 文本/3 空、min 1 median 36 max 71、头尾样本 |
| 15 | `close` → et 进程 0 | ✓ 5 秒后进程归零，无泄漏 |
| 16 | `restore`：snapshot → 改 C2=999 + 新增 H1 → restore | ✓ restored 2（含公式联动 E2）+ cleared 1（H1），随后 `diff` totalChanges=0、值与公式完全复原；仅内存还原，未自动保存 |

## v0.3（2026-10-24）：大表性能三件套

| # | 用例 | 结果 |
|---|---|---|
| 17 | `fill -Range B1:B2000 -Formula "=A1*2+1"` | ✓ 2000 行相对公式一次落位（首格 #VALUE! 属预期：A1 是文本；空 A 行=1 ✓），含进程启动约 2s |
| 18 | write/restore 批内挂起重算 | ✓ PerfOn/PerfOff 包裹写循环：批前 Calculation=Manual + ScreenUpdating=false，批末一次 Calculate 后恢复原设置 |
| 19 | `snapshot -Sheet test`（24 格）→ 同时改 test!C2 与 summary!B5 | ✓ diff 只报 test 表 2 处（含公式联动 E2）；restore 只还原 test 表并报 `extraSheets:["summary"]`，summary!B5=888 **原样保留**——限定范围快照的隔离性验证通过 |
| 20 | .xls 全链路：xlsx 另存 .xls → 复开 → read（公式+值无损）→ write（expect 锚）→ fill 100 行 → snapshot -Sheet → save → close | ✓ 全过；`save -As` 已映射 .xls=56 格式号；et 进程归零 |

## v0.4（2026-10-24）：search / selection / rows+cols / chart

| # | 用例 | 结果 |
|---|---|---|
| 21 | `search -Text AVERAGE` | ✓ 值+公式双命中（2 处，含格号/值/公式原文） |
| 22 | `rows -Action insert -Index 2 -Count 2` | ✓ 第 2 行起插 2 空行，后续数据整体下移 |
| 23 | `cols -Action delete -Index 6` | ✓ 删第 6 列 |
| 24 | `selection`（COM 预设选区 B5:C6 后读取） | ✓ sheet/range/行列数/4 格值全对；跨工作簿选区报 NOT_OPEN 的路径已设防 |
| 25 | `chart add -Type line -Range C3:C6 -Title -Anchor` / `list` / `delete -Index 1` | ✓ 图表落位、命名、列表可寻址、删除干净；渲染区域非白 4.9%（4 点折线图） |
| 26 | 开发期修复 | chart 的 range/delete-index 漏传（wrapper）、StringCompare 笔误、去重 -Formula/-Index 参数冲突 |
| 27 | **实战首例**（外洲站实时水情→库容曲线簿 Sheet3）：查询→写当前值表→136 行逐时过程块→双过程线图 | ✓ 全链路通；实战修复 **SetSourceData 单列方向猜错**（每行一个系列）→ 改显式 NewSeries，`seriesCount:1` 核对通过；`values` 块必须配整块 range（配单格报尺寸不符，属防护）；日期字符串自动转序列值需配 NumberFormat |
| 28 | 多列 chart 取列修复 + **v0.5 协作三件套** | ✓ 多列数据取**最后一列**为数值系列（原固定第 2 列致流量图画成水位）；write/fill 自动 pre-batch 快照（`preSnapshot.path`，仅覆盖受影响表）；`diff -Since <batch文件>` = 批次归因（实测精确命中 fill 的 4 格）；`restore -Cells "test!H2"` 选择性回滚（H2 还原、H3:H6 保留 ✓） |

## v0.6（2026-10-25）：finish 完成信号（agent 无关，零依赖）

| # | 用例 | 结果 |
|---|---|---|
| 29 | `finish` 状态行防覆盖 | ✓ 目标格非空拒绝（提示换 -Cell 或 -Force） |
| 30 | `finish -Cell L1 -Note -Toast -Sound` | ✓ 状态行写入 test!L1、托盘气泡弹出、提示音播放（NotifyIcon + SystemSounds，纯 .NET Framework，无 WinRT/插件依赖）；`-Quiet`/`-Force` 路径在册 |
| 31 | **语义调整（用户决策）**：finish 默认**不写工作簿** | ✓ 默认=气泡+提示音（toast/sound true），前后读取工作簿逐字节一致；状态行改为 `-Cell` 显式选择加入（非空格仍拒绝）；`-Quiet` 静音 |

## v0.6.1（2026-10-25）：DOC_CLOSED——用户关闭文档的感知

| # | 用例 | 结果 |
|---|---|---|
| 32 | 接管后模拟用户绕过 skill 直接关簿（COM Close）→ 下一步 write | ✓ 报 `DOC_CLOSED`（不再静默重开——旧行为会悄悄复活用户刚关的文件） |
| 33 | 显式 `open` 重新接管 → write | ✓ 恢复正常；attachment marker 随 open/接管登记、随 SaveAs 迁移到新路径 |
| 34 | skill 自己 `close` 后 write | ✓ 同样 `DOC_CLOSED`（关闭=需显式 open 才能继续，close 不清标记）；进程归零 |

修复记录（开发当日）：写块数组 1 基→0 基；`SplitRef` 覆盖 `-Sheet` 参数（read 与 write 两处）；visible=-1 判定；`Image.FromFile` 文件锁；et.ps1 缺 `-Save` 开关；PS 函数返回单元素数组展开（`return ,@(...)`）。

### 回归清单（改源码后至少跑这些）

1. `doctor` → 2. `open` → 3. `read -Sheet X` → 4. `write`（1 条带 expect、1 条公式、1 条块）→ 5. `comment add`+`list` → 6. `render` → 7. `snapshot`→改→`diff` → 8. `save -As` → 9. `close` → et 进程数。

## v1.11（2026-10-30）——接管提示 `hint`

用户问"agent 接管 WPS 文档后能不能有个友好的提示"。表格侧与 Writer 同款：窗口标题 + 状态栏 + 托盘气泡，**全部是窗口状态**（不写单元格、不落盘、`saved` 不变）。

实测（WPS 12.0，本机）：

| 能力 | 结论 |
|---|---|
| `wb.Windows(1).Caption` 写 | ✓ 可写；`wb.ActiveWindow` 在这种场景返回 null（簿非活动），所以主路径是 `Windows(1)` |
| `Application.StatusBar` | ✓ 可读写，且截图确认状态栏真的显示 `AI 正在协作：…`（Writer 侧只写不读也不渲染）；写 `false` 交还 |
| 托盘气泡 | ✓ NotifyIcon（与 `finish` 同一机制），`-Quiet` 可静音 |

- `open` 自动打标记（`-NoHint` 关）；只读簿自动标成 `〔AI 只读〕`。
- `hint -Text "正在核对第 3 张表" -Quiet` 只刷状态栏（截图可见）；`finish` / `close` 自动还原（`hintCleared`）。
- 验收：`tests/excel/check_hint.py` 9 条断言全过（caption 标记/`status hinted`/状态栏写入/`finish` 还原/显式 `-Clear`/`close`）。


## v1.12（2026-10-30）——常驻托盘图标 `hint -Tray`

表格侧与 Writer 共用 `scripts/helper/shared/Tray.cs`：`hint -Tray` 写状态文件（`%LOCALAPPDATA%\wps-office-cooperate\tray-excel.json`），再用自身 exe 以 `--tray` 重开一份常驻图标；悬停显示「〔AI〕 工作簿 · 当前步骤」，右键「结束协作：还原标题标记」，`finish`/`close` 自动移除。

- 验收：`tests/excel/check_tray.py` 7 条断言全过（托盘启动 + 进程存活 + 状态文件 + `finish` 报告 trayStopped + 进程 8 s 内退出 + 外部关闭工作簿后图标 8.6 s 自动消失）。
- 坑：spawn 必须 `UseShellExecute=true`，否则子进程继承调用方 stdout 管道，任何读到 EOF 的调用方（python `subprocess.communicate`、PowerShell 管道）会被图标的 4 小时寿命卡死（实测复现）。
