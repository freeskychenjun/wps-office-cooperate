# WPS Writer ≠ Word：实测差异表

本机实测环境：WPS Office **12.1.0.29161**（无微软 Office），COM 由 WPS 接管；`KWPS.Application` 能跑通整套 Word 对象模型，但"脾气"不同。下表每条都是这台机器上跑出来的，照着写能少走一圈。

**出处**：`[研]` = `D:\DSH工作区\wps-skill-research\WPS协作Word-skill-可行性研究.md`（2026-10-23 真机实测，标注 § 章节）；`[设]` = 本目录 `DESIGN.md`；`[未验]` = 未验证。

| 现象 | 实测表现 | 正确写法 | 出处 |
|---|---|---|---|
| `Application.Name` 冒充 Word | `Name="Microsoft Word"`、`Version="12.0"`、`Build="12.1.0.29161"`、`Path=...\Kingsoft\WPS Office\12.1.0.29161\office6`（`status` 的 `app` 字段原样返回过，2026-10-23 复核） | 判断套件看 `Path`（含 `Kingsoft` 即 WPS）或 `Build`，**不要**看 `Name`/`Version`；`doctor`/`status` 返回的 `app` 字段已是真身 | [研] §2.2 |
| 样式名只认中文本地化名或内建常量 | `Paragraph.Style="标题 1"` OK；`="Heading 1"` → PS `E_FAIL`、python `0x80020009` | 用 `Style.NameLocal` 读回的中文名（`标题`/`正文`/`标题 1`），或内建常量（`-2` = 标题 1）；不确定时先 `read` 看 `style` 字段 | [研] §2.2、§2.4 #1/#10 |
| `Document.Range` 是**方法**不是属性 | `$doc.Range.Information(3)` → `[PSMethod] does not contain 'Information'` | C#：`doc.Range(a, b)`；PS：`$doc.Range(0, 0)`。任何"先拿 Range 再调方法"的写法都会踩 | [研] §2.4 #9 |
| 关掉最后一个文档 = 静默杀掉自动化实例 | 关最后一份文档后实例已死（进程 5→3），但**不抛异常**：`$app.Version` 变空串、`Documents.Count` 返回 0、`Range.Text` 返回空 | 留文档保活；每次操作前校验 `Documents.Count`/`Version`，别信"没报错"；本 skill 的 `close` 只在"无文档"时才 `Quit`，并回报 `processes` | [研] §2.4 #4/#8、§5-2 |
| 每次 `New-Object` 都是**新实例**，不是复用 | 进程 5→9；各实例的 `Documents` 互相看不见，A 实例读不到 B 实例打开的文档 | 先附着已运行实例（C# `Marshal.GetActiveObject("KWPS.Application")`，python `win32com.client.GetActiveObject`），没有再创建；本 skill 全命令共用一个实例 | [研] §2.4 #3、§5-1 |
| 没有 `Range.CopyAsPicture` | 方法表里只有 `Copy`，"截一段出来看"没有 API | 版式核对只走 `Windows(1).Panes(1).Pages(n).EnhMetaFileBits` → GDI+ 画 PNG（`render` 已封装；单页 EMF 57–82 KB、非空白像素 4.77%） | [研] §2.2、§2.4 #2 |
| 改动后视图跳回光标 | 每次写入后视口回到光标处，"跟随滚动/自动翻到改动处"做不到 | 不承诺跟随。要核对用 `render` 出图；要让用户自己看，就给出段号与段落开头文字让他定位 | [研] §1.3、§2.4 #5 |
| 批注/修订作者固定为登录用户 | 修订 `Author=陈鋆`；`Comments.Add(range, text)` 只有 2 个参数，传 3 个 `TypeError`；改作者得动 `Application.UserName`（[未验]） | 接受"作者 = 当前登录用户"，AI 改的与用户改的分不开；要留痕就把署名写进批注正文（如"（AI 校对）"） | [研] §2.2、§2.4 #7 |
| 表格会占用段落编号 | 按段落号下"第 3 段之后插入"，实际插进了表格第一格 | 动手前必 `read`；items 必带 `expect`；`read` 标了 `inTable: true` 的段落不要 `insert_after`；插入会改变后续编号，故 items 按 `para` 倒序执行 | [研] §5-7、[设] §3 |
| 刚插入的文本再替换会二次命中 | 参考实现注释里的实测：`"P"`→`"P-"`→`"P--"`（新插入的文字被下一次替换又匹配到） | 插入前把最终文本算好、一次写到位；不对同一段做链式替换 | [研] §1.3、§2.4 #6 |
| 反向纠错：`Range.Cells` 其实**存在** | `Table.Range.Cells.Count=9` 正常；参考仓库注释"WPS 没有 Range.Cells"不成立（或仅特定场景） | 别因为别人的注释就绕路，先实测再决定绕法 | [研] §2.4 #11 |
| 别用 PowerShell 5.1 写 COM | PS 5.1 晚绑定下 `Documents` 返回裸 `__ComObject`（`does not contain a method named 'Open'`）、`RPC server is unavailable`、`Range.Information` 调不到；同一动作 python 一次通过 | COM 逻辑放 C# helper（`csc.exe` 编译），PS 只做参数转发——本 skill 的形态；**不要**在 `.ps1` 里直接 `New-Object -ComObject` | [研] §2.3 ③、[设] §1 |
| 源码含中文会炸 | PS 5.1 按 ANSI 解析含中文的 `.ps1`，中文变乱码并报看似无关的语法错 | `.ps1`/`.cs` 一律纯 ASCII，中文只出现在 JSON 字符串里（`write` 工具产出无 BOM UTF-8，尤要注意） | [设] §4-6、记忆《pwsh 工具实为 PS5.1 与无 BOM 坑》 |
| 批注锚点让同段字符位置错 1 | 批注旁做 `replace` 会多删/少删一个字（实测"对无实测"→"对无实测测"） | helper 已改为读回校验+±3 自动纠正，纠正不了就拒绝；必要时先删批注再改 | 2026-10-23 白塔河第 74 段 |
| `Table.Range.End` 是**排他端** | 判"这段在不在表内"用 `start <= Range.End` 时，**紧接表格之后的那段正文**（`Range.Start` 恰好等于表末）会被算成表内：实测 10521 段报告把段 10297 的"为满足表3.3-2的控制要求…"标成 `inTable:true`，白跑 2 次 `render` 才排除 | 一律 `Range.Start <= start && start < Range.End`（段区间半开且铺满全文，正好只有一个段能命中）；v1.8 起 `read`/`search`/`audit` 都按此判据 | 2026-10-26 新干枢纽报告复盘（§4 缺陷 B） |
| `Find` 的字符偏移 → 段落号**必须按区间校验** | 候选段只用"该段文本包含这个词"裁决，在长表格区域会选到错邻段；错段号≠0，所以连"漂移→退回慢扫"的兜底都不会触发（实测漏 4 条真实题注、同一段被重复报 3 次，据此还写出过错误结论） | 候选段用 `Range.Start <= offset < Range.End` 裁决（候选不变，只换判据）；校验不过才退回逐段慢扫；结果标 `verified:true`，同段多命中用 `foundStart - paraStart` 各自算 `at`（`Range.Text` 遮删修订时会漂，故再按段内递推纠正） | 2026-10-26 复盘（§3 缺陷 A） |
| 表格单元格与行末标记都算段落 | 10521 段的报告里 9000+ 段是单元格；按段号全扫 = 22 次 `read`、4.5 分钟，而真正含题注的段不到 40 个 | 先 `captions` 出清单（与段数几乎无关），再按 `para` 定点读；非要全扫就用 `read -SkipTableCells -Count 0`（一次调用） | 2026-10-26 复盘（§5 缺口 C） |
| 题注判据：标签必须在**段首** | 用"段内含 `表 N`"判题注时，正文句"如表3.3-2所示"也会被算成题注；题名做进表格第 1 行的表又会同时被报"缺题注"和"题名在表内" | 用同一个题注解析器（`^\s*(表\|附表\|图\|附图)\s*数字`），`audit`/`table -Action caption`/`captions` 三处共用；表内题名只报 `table-caption-in-row`；图片题注允许在图片后 1–4 段（实测有图名隔着一个空段和一个"/"残段） | 2026-10-26 复盘（§6 缺陷 D + 现场数据） |

| 窗口标题可改写（Writer） | `doc.ActiveWindow.Caption` 可写，`GetWindowText` 读到真实窗口标题变成 `〔AI〕 文件.docx - WPS Office`。**但 WPS 12 把文档显示成标签页并从右截断（约 22 个半角字符）**：给 `上高县南港水库引调水工程可行性研究报告.docx` 加后缀「〔AI 协作中〕」截图像素与不加完全相同（等于没打），所以 `hint` 把标记放**前面**；文档不在前台时 `doc.Windows(1).Caption` 同样可写 | 用作接管提示（`hint`）。还原 = 写回 `doc.Name`；`save -As` 改名后要重写，否则标题停在旧文件名 | 2026-10-30 v1.11 探针 |
| `Application.StatusBar` 在 Writer **只写不读、且界面不渲染** | 赋值成功且 `DisplayStatusBar=True`，但读回抛 `Cannot get property value because "StatusBar" is a write-only property.`；更关键的是**状态栏看起来没变**（截图仍是 页面/字数/拼写检查，没有写入的文字） | 只报"写入被接受"；**别在 Writer 侧承诺状态栏可见**，报进度用气泡（`-Toast`）或标签页标记；`hint` 返回里的 `statusBarNote` 已写明这一点 | 2026-10-30 v1.11 探针 + 截图 |
| 表格侧 `Workbook.ActiveWindow` 常为 null | 打开的簿不是当前活动簿时 `wb.ActiveWindow` 返回空（读 `.Caption` 抛异常）；`wb.Windows(1).Caption` 读写正常；表格侧 `Application.StatusBar` **可读写**，写 `false` 交还 WPS | 表格一律走 `wb.Windows(1).Caption`；清状态栏用 `StatusBar=false` | 2026-10-30 v1.11 探针 |
| PPT 没有可写的单文档标题 | `ActiveWindow.Caption` 只读（赋值抛 `Cannot find an overload for "Caption"`）；`Application.Caption` 是**应用级**且标题按 `<稿名> - <Caption>` 渲染，写进去会顶掉本地化的"WPS 演示"，安全恢复需持久化状态 | PPT 侧接管提示只给托盘气泡，响应如实返回 `captionSupported:false` | 2026-10-30 v1.11 探针 |

## 边界与未验证

- **版本**：只在 **12.1.0.29161** 上验证。同事若装的是别的 WPS 版本，上表结论**未验证**，先用 `doctor` + 一份小文档试两三条（样式赋值、`render`、`close`）再上真稿。
- **并发**：用户正在打字/有对话框时的写入行为（`BUSY` 场景）——[研] §6 记为待补测，**未验证**。稳妥做法是等用户停手再动。
- **错误码映射**：`BUSY`/`ENCRYPTED` 是 [设] §2 约定的错误码；当前 helper 未单独映射时，会以 `COM_ERROR` + HRESULT（如 app 忙为 `0x8001010A`）出现。见到含义相同的现象，按 [SKILL.md](../SKILL.md) 故障排查表同一条处理。
- **公式（OMML）**：不在本 skill 范围，未验证。
- **Excel / PPT**（`KET.Application` / `KWPP.Application`）：v1.1 起已支持，各自差异见 [excel-quirks.md](excel-quirks.md)；本文件只记 Writer 侧。
