# WPS Writer ≠ Word：实测差异表

本机实测环境：WPS Office **12.1.0.29161**（无微软 Office），COM 由 WPS 接管；`KWPS.Application` 能跑通整套 Word 对象模型，但"脾气"不同。下表每条都是这台机器上跑出来的，照着写能少走一圈。

**出处**：`[研]` = `D:\DSH工作区\wps-skill-research\WPS协作Word-skill-可行性研究.md`（2026-10-23 真机实测，标注 § 章节）；`[设]` = 本目录 `DESIGN.md`；`[未验]` = 未验证。

| 现象 | 实测表现 | 正确写法 | 出处 |
|---|---|---|---|
| **每个表格行被算作一个空段落** | `Document.Paragraphs.Count` = 正文 `w:p` 数 + 表行数：7334 段 / 23 表 / 555 行实测 `6779 + 555 = 7334`；逐个表看：4 个单元格后多一个空段号、8 行表多 8 个（2026-11-01 实测） | 用 `Paragraphs` 段号做映射时必须**给每个表行补一个空段**；`Content.Text` 按 `\r` 切块**不含**这些空段，所以 v1.7–v1.10 的块模型在带表文档上必然错位（实测 7334 段切出 7328 块，从第 2 段起整篇偏） | [隅 v1.13] |
| `Content.Text` 隐藏未接受的删除 | 有修订时 `Content.Text` 给的是"接受视图"（删除的文本不出现），而 `Range.Text` 也同样是接受视图；两者一致，但**修订数≠0 时块模型一律回退**，等于实时协作时读侧只有慢路径 | v1.13 起改走 `doc.Content.WordOpenXML` 一次取全稿 + 本地解析（`w:delText` 不取即可对齐 `Range.Text`），7334 段实测 48 s → 6 s | [隅 v1.13] |
| 内嵌图片在图内文本里是 `/` | 逐段 `Range.Text` 对一张内嵌图返回 `/`（不是 U+FFFC）；XML 里是 `w:drawing` | XML 侧把 `w:drawing`/`w:pict`/`w:object` 映射成 `/` 才能与 `Range.Text` 逐字一致；含 `w:sym`（Wingdings 勾选框）或一段多图的段落**单独回读**最稳 | [隅 v1.13] |
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

| 窗口标题可改写（Writer） | `doc.ActiveWindow.Caption` 可写，`GetWindowText` 读到真实窗口标题变成 `〔AI接管中〕 文件.docx - WPS Office`。**但 WPS 12 把文档显示成标签页并从右截断（约 22 个半角字符）**：给 `上高县南港水库引调水工程可行性研究报告.docx` 加后缀「〔AI 协作中〕」截图像素与不加完全相同（等于没打），所以 `hint` 把标记放**前面**；文档不在前台时 `doc.Windows(1).Caption` 同样可写 | 用作接管提示（`hint`）。还原 = 写回 `doc.Name`；`save -As` 改名后要重写，否则标题停在旧文件名 | 2026-10-30 v1.11 探针 |
| `Application.StatusBar` 在 Writer **只写不读、且界面不渲染** | 赋值成功且 `DisplayStatusBar=True`，但读回抛 `Cannot get property value because "StatusBar" is a write-only property.`；更关键的是**状态栏看起来没变**（截图仍是 页面/字数/拼写检查，没有写入的文字） | 只报"写入被接受"；**别在 Writer 侧承诺状态栏可见**，报进度用气泡（`-Toast`）或标签页标记；`hint` 返回里的 `statusBarNote` 已写明这一点 | 2026-10-30 v1.11 探针 + 截图 |
| 表格侧 `Workbook.ActiveWindow` 常为 null | 打开的簿不是当前活动簿时 `wb.ActiveWindow` 返回空（读 `.Caption` 抛异常）；`wb.Windows(1).Caption` 读写正常；表格侧 `Application.StatusBar` **可读写**，写 `false` 交还 WPS | 表格一律走 `wb.Windows(1).Caption`；清状态栏用 `StatusBar=false` | 2026-10-30 v1.11 探针 |
| PPT 没有可写的单文档标题 | `ActiveWindow.Caption` 只读（赋值抛 `Cannot find an overload for "Caption"`）；`Application.Caption` 是**应用级**且标题按 `<稿名> - <Caption>` 渲染，写进去会顶掉本地化的"WPS 演示"，安全恢复需持久化状态 | PPT 侧接管提示只给托盘气泡，响应如实返回 `captionSupported:false` | 2026-10-30 v1.11 探针 |

## 边界与未验证

- **版本**：只在 **12.1.0.29161** 上验证。同事若装的是别的 WPS 版本，上表结论**未验证**，先用 `doctor` + 一份小文档试两三条（样式赋值、`render`、`close`）再上真稿。
- **并发**：用户正在打字/有对话框时的写入行为（`BUSY` 场景）——[研] §6 记为待补测，**未验证**。稳妥做法是等用户停手再动。
- **错误码映射**：`BUSY`/`ENCRYPTED` 是 [设] §2 约定的错误码；当前 helper 未单独映射时，会以 `COM_ERROR` + HRESULT（如 app 忙为 `0x8001010A`）出现。见到含义相同的现象，按 [SKILL.md](../SKILL.md) 故障排查表同一条处理。
- **公式（OMML）**：v1.29 已实现（`formula` 命令），实测事实见文末「2026-11-02 公式 / 脚注 / 总览图」。
- **Excel / PPT**（`KET.Application` / `KWPP.Application`）：v1.1 起已支持，各自差异见 [excel-quirks.md](excel-quirks.md)；本文件只记 Writer 侧。

## 2026-11-02 公式 / 脚注 / 总览图 实测（v1.29，WPS 12.1.0.29161）

一条一行，全部是这台机器上跑出来的；每条后面是当时的实测证据原文。

| 现象 | 实测表现 | 正确写法 |
|---|---|---|
| **WPS 只认 Flat OPC 包** | `range.InsertXML(...)` 传完整 `pkg:package`（`/_rels/.rels` + `/word/document.xml` 里带 `xmlns:m` 的 `<m:oMath>`）→ `Document.OMaths.Count` 0→1，可回读 `OMaths.Item(1).Range.Text`。**裸 `<m:oMathPara>`/`<m:oMath>` 片段不行**（插进去是空/报错） | 用 `FlatDocument()` 拼包（`scripts/helper/word/Formula.cs`），`<w:p>` 里放 `<m:oMath>`（行内）或 `<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/>` + `<m:oMath>`（独立式） |
| `Document.OMaths.Add(range)` 不可用 | 报「值不在预期范围内」（`0x80020009` 一类），位置参数与命名参数均失败 | 唯一路线是 InsertXML；WPS 不会像微软 Word 那样把线性文本自动转公式 |
| **编号 `\tag{n}` 用 eqArr + `#`** | `<m:oMath><m:eqArr><m:eqArrPr><m:maxDist m:val="1"/></m:eqArrPr><m:e>…<m:r><m:t>#</m:t></m:r><m:r><m:rPr><m:nor/></m:rPr>…<m:t>(2)</m:t></m:r></m:e></m:eqArr></m:oMath>` → render 出来公式居中、`(2)` 靠右页边 | 照 dsh-office `OfficeHelper.cs:1704-1708` 的做法；编号是**普通文字**不是 SEQ 域，增删公式不会自动重排 |
| 行内插入会多一个段落标记 | InsertXML 后 `Paragraphs.Count` +1（包是按 `<w:p>` 进来的） | 插完若段数变多，删「插入点所在段的段末 CR」把两行合并（`doc.Range(lineEnd-1, lineEnd).Delete()`），行内式才不会多空行 |
| 独立式插入会多一个空段 | 先 `host.Range.InsertParagraphAfter()` 造空段再插入，会留下一个空段 | 插完在目标段后 1–3 段内找「文本为空且 `Range.OMaths.Count==0`」的段删掉；再把公式段 `Alignment=1`、段前段后 ≥6pt |
| **校验判据只有 OMaths.Count** | `InsertXML` 成功但 WPS 没建成公式时，`OMaths.Count` 不涨，插进去的东西可能是空段 | 插前记 `mathsBefore`，插后 `<= mathsBefore` 就按 `doc.Content.End` 增量回滚，返回 `MATH_FAILED`；**不要**只看有没有抛异常 |
| **`Footnotes.Add` 可用** | `doc.Footnotes.Add(spot, "", " ")` 成功（三个位置参数；命名参数也可）；脚注是**原生**的：WPS 自己编号，`render` 页脚能看到 `¹ …` | 插入点 `spot` 用正文 range（`p.Range.Duplicate` 去掉段末 CR，或 `-Find` 锚词后 `Collapse`），`FillNote` 时把正文两端剥开 |
| 脚注正文回读要剥两边 | `note.Range.Text` 形如 `"\u0002正文…\r"`（前导 = 脚注引用标记 U+0002，尾部 = 段末 CR）；直接 `.Text` 赋值会把标记冲掉 | `NoteBody()`：`startsWith("\u0002") → Start+1`、`endsWith("\r") → End-1`，再读/写 |
| **索引口径：`Range.Footnotes` ≠ `Document.Footnotes`** | 命中 dsh-office 的注记（`OfficeHelper.cs:2251`）：WPS 的 `Footnote.Index` 相对 `Range.Footnotes`；刚 `Add` 完集合可能未重排 | 对外**只暴露 `Document.Footnotes` 的序号**（`doc.Footnotes.Item(k)`），每次改完重读集合；`insert`/`delete` 后提示用户先 `-Action list` 再按序号改 |
| **`Document.Footnotes.NumberingRule` 可读写**（更正 dsh-office 旁注） | `doc.Footnotes.NumberingRule = 1` → 回读 `1`（`via:"Document.Footnotes"`）；`Range.FootnoteOptions.NumberingRule` 是否同版本可用**未单独验证**，代码先走 Footnotes 集合、失败再试 FootnoteOptions | 脚注编号原子/节/页切换走 `Document.Footnotes.NumberingRule`（0 连续 / 1 按节 / 2 按页），别只盯 `FootnoteOptions` |
| **`render -Overview` 的先决条件** | `Panes.Item(1).Pages.Count` 只在页面视图下可靠：必须先 `win.View.Type = 3`（现有 `render` 已先切），否则 0 页或错页 | 总览与单页共用同一取图路径（`Pages.Item(n).EnhMetaFileBits` → `Metafile` → GDI+），不引入新截图方式 |
| 总览图排版实测常数 | 7 页样张 `-Overview -Page 1 -To 7` → `overview-1-6.png` **1976×1762**（3 列 × 2 行，格宽 640，gap 14 / label 30，底 `(232,232,236)`，每格上方「第 N 页」）+ `overview-7-7.png`；返回 `shown:7, capped:false` | 每张 6 页、单次 ≤18 页；**格宽固定 640**（别跟着 `-Width` 涨，18 页会变十几 MP）；每页自己 `Dispose`，峰值只有 1 张 sheet + 1 页位图 |
| 多实例会随命令变多 | 一轮测试里 `wps.exe` 从 2 涨到 10（`doctor`/`status` 的 `processes.wps` 可见）：一次性模式下 `Marshal.GetActiveObject` 偶发失败就会**新建隐藏实例**，`open` 返回里也会给 `multiInstanceWarning` | 收尾 `close` 每个自建文档 → `close` 后只剩用户的实例；只清**无窗口**的 `/Automation` 实例，绝不 `quit` 有用户文档的实例 |

