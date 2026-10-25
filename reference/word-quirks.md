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

## 边界与未验证

- **版本**：只在 **12.1.0.29161** 上验证。同事若装的是别的 WPS 版本，上表结论**未验证**，先用 `doctor` + 一份小文档试两三条（样式赋值、`render`、`close`）再上真稿。
- **并发**：用户正在打字/有对话框时的写入行为（`BUSY` 场景）——[研] §6 记为待补测，**未验证**。稳妥做法是等用户停手再动。
- **错误码映射**：`BUSY`/`ENCRYPTED` 是 [设] §2 约定的错误码；当前 helper 未单独映射时，会以 `COM_ERROR` + HRESULT（如 app 忙为 `0x8001010A`）出现。见到含义相同的现象，按 [SKILL.md](../SKILL.md) 故障排查表同一条处理。
- **公式（OMML）**：不在本 skill 范围，未验证。
- **Excel / PPT**（`KET.Application` / `KWPP.Application`）：未验证，本 skill 不做。
