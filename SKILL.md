---
name: wps-office-cooperate
description: 在 WPS 里实时改 Office 文档——Word（.docx，修订模式改写+批注+版式核对）、Excel（.xlsx/.xls/.csv，批量写值公式+锚点校验+快照回滚+批注+图表）与 PowerPoint（.pptx，预设版式建页+形状级改文字+表格图片+逐页渲染核对+快照对比），改动当场出现在用户正开着的 WPS 窗口里。当用户要求：在 WPS 里改这个文档/表格/演示文稿/PPT、在 WPS 里做修订、实时改 Word/Excel/PPT、一边看着一边改、做汇报 PPT、加批注、核对排版/版式/表格格式，或提到 wps word、wps excel、wps ppt、live edit word/excel/ppt 时使用。不适用：纯离线批量处理（改副本、不要求用户在场，Word 用 docx-paragraph-polish，Excel 用 openpyxl/minimax-xlsx，PPT 用 pptx-generator 等离线方案）、用户明确不要动原文件（只出旁挂报告）、微软 Office 窗口里的文档（本 skill 只走 WPS）。
---

# wps-office-cooperate：在 WPS 里实时改 Word / Excel / PowerPoint

**版本 v1.7（2026-10-25）**：PPT 新增 **`ppt font`** —— 全稿统一字体（`font -Doc -Font 宋体`），逐页扫**组合内部与表格单元格**，只写字体名、不动字号/颜色/加粗/版式。起因：用户要"所有页面改成宋体"，而已有 `format` 一次只能改一个形状且**进不到组合里**（该稿第 2 页就是 5 个组合）。实测 15 页 / 45 形状 / 5 组合 / 1 表格，0 失败；并记下"换字体 → 字身变宽 → 多折行"的连带风险（第 10 页末条曾被挤出页面，需 `format -Size` 收一档）。

**版本 v1.6（2026-10-25）**：`word unit` 增加**上下文守卫 + "可疑碎片"报告 + 脚本段识别**。起因是真报告里有 3 页（实测第 256–258 页）是**公文模板/脚本语法说明**（`[for] … [/for]`、`[File]`、`k~a = 1`、`m1 = 1`／`n2 = 1` 成串赋值），其中的 `km2/m2/m3` 是**变量名**；无守卫的 v1.4 会把它们改成 `km²/m²/m³`（已在真实文档上复现）。现在：
- **只修"挂在数值上"的单位**（前字符＝数字/万/亿/汉字/半全角左括号，或"数字 空格 单位"）；
- **脚本段整段跳过**：段落命中 `[for]`/`[File]`/`k~a`/≥2 条 `x = 1` 赋值即判为模板示例，报告里单列 `scriptPages`（页码），此段的单位一律不动；
- 三类只报不改：`fragments`（孤立 token、后跟 `=`）、`looseFragments`（`n2`/`n3` 缺基名）、`suspects`（`km3` 疑似笔误）；
- 万一在脚本页上仍修了东西（段落特征没覆盖到），会进 `fixedOnScriptPage` 告警。

**版本 v1.5（2026-10-25）**：把"连接"与"动手"分成两挡——新增 **铁律 7**（用户只要"连上 / 接管"时，`status` + `open` 报一句摘要即停，不 `read` 全稿、不 `render`、不 `audit`、不 `close`）；PPT 视觉自检闭环的前提改为"**本轮真的改动过这份稿**"；`audit` 注明人工制作的老稿误报率高、结论以 `render` 目检为准。起因：一次"连上 PPT"被做成了全稿体检，用户干等。

**版本 v1.4（2026-10-25）**：新增 **`word unit`** —— 单位上标批量修（把掉下来的 `m3/s`/`km2` 改成 `m³/s`/`km²`）。判据来自实测：同一篇文档里"对的上标"有两种形态（Unicode ²/³、run 级上标），只有 `Find.Font.Superscript = 0` 能挑出真正掉下来的；逐处修改 + 三重守卫（格式过滤、前字符非英文字母、后字符非数字），已正确的与 `item3`/`m23` 类英文数字形态自动跳过并回报上下文。**疑似笔误只报不改**：`km3`（面积表里几乎必是 `km2` 笔误）等列进 `suspects` 交人确认——自动加 ³ 会让笔误看起来像有意为之。300 页 / 13498 段真实报告实测 ~21 s，精确修 16 处、零误伤。

**版本 v1.3（2026-10-25）**：按一份真实报告（饶河水系 297 页 / 85 张表）的使用记录补 Word 表格能力——`table -Action scan`（表格体检，只读）、`-Action caption`（题名做进第 1 行的一键提到框外，内部保证"插题名 → 删行 → 加边框"的顺序）、`-Action header-repeat`（跨页表头，自动迭代到稳定）、`-Action rows`（删行）；修掉 `insert -Where after` 在表格前落进表格的坑、`-HeaderRow` 静默失败（WPS 的 `Rows.Item(n)` 是残缺对象）、`audit` 新增"题名在第 1 行"规则、`page` 查询模式报页码断号 + `-PageNumberStart none`。

**版本 v1.2（2026-10-25）**：PPT 侧补齐"能交付"这一层——`theme`（主色/字体/页脚/页码，一套 JSON 打底）、预设版式重做（封面色带、章节页色块+大数字、真项目符号、**深色底自动白字**、正文框留白给表格/图片）、`render -Overview`（6 页拼一张图并带页码）、`audit`（文字溢出/字号过小/越界/重叠/图过小/空页机械体检）、`align`/`zorder`/`restore`。

**版本 v1.1（2026-10-25）**：新增 **PPT（WPS 演示，`KWPP.Application`）** 支持——预设版式建页、形状级锚定改文字、表格/图片/文本框、演讲备注、逐页渲染核对、文本快照对比。v1.0 由 `wps-word-cooperate`（v0.8）与 `wps-excel-cooperate`（v0.6）合并而成。版本历史见 [reference/word-acceptance.md](reference/word-acceptance.md)、[reference/excel-acceptance.md](reference/excel-acceptance.md)。

**能力边界（一句话）**：只走 **WPS**（Word=Writer `KWPS.Application`，Excel=表格 `KET.Application`，PPT=演示 `KWPP.Application`，没装微软 Office 也照跑）。Word 侧所有文字改动都是**修订（track changes）**，用户能逐条接受/拒绝；Excel 与 PPT 侧改动**直写**，靠**锚点校验 + 快照**保证可核对（**这两个 app 的 Ctrl+Z 撤不回 Agent 的改动**，外部 COM 修改不进撤销栈）。

> 下文 `<skill>` = 本 skill 的安装目录。命令一律按**完整可跑**形式给出：
> `powershell -NoProfile -ExecutionPolicy Bypass -File <skill>\scripts\office.ps1 <word|excel|ppt> <命令> [参数]`
> （`word.ps1` / `excel.ps1` / `ppt.ps1` 也可直接调用，参数完全一样。装有 PowerShell 7 的机器可把 `powershell` 换成 `pwsh`。）
> 输出只有一段 JSON：`{"ok":true,"result":{...}}` 或 `{"ok":false,"error":{"code":"...","message":"..."}}`。
> 环境：Windows + WPS；首次调用会用系统自带 `csc.exe` 编译 helper（约 1 秒，之后按源码哈希缓存，不再编译）。

## 铁律（三个 app 通用，一条都不能省）

0. **一份文档，一条通道（最重要）。** 一旦本 skill 打开了某份文档、或接管了用户正开着的那份，**在这次协作期间，这份文档的一切处理都走本 skill**：
   - Word 校审 → `word audit` + `comment -Items`（**不要**加载 `design-report-review`）
   - Word 润色 / 改写 → `revise`（**不要**加载 `docx-paragraph-polish`）
   - Word 排版 → `format` / `style` / `page`（**不要**加载 `design-report-format`）
   - Excel 改表 → `write`/`fill`（**不要**再用 openpyxl/pandas/minimax-xlsx 直接改磁盘文件）
   - PPT 改稿 → `slide`/`text`/`textbox`（**不要**再用 pptx-generator/python-pptx 生成副本替换它）

   理由不是"体验"，是**会坏**：那些 skill/工具改的是磁盘文件，而这份文档正被 WPS 独占（实测写盘 `PermissionError`）；它们还产出副本，会让你眼前这份和副本分成两条线。
   唯一例外：用户明确要"**不动原件、出一份新稿**"或"**无人值守批量跑**"——那时先 `save` → `close` 交出去，回来再 `open` 产出的文件、`snapshot` 重建基线（Excel 用 `save -As` 另出一份，或在用户关掉文件后离线处理）。
1. **先 `status`，再动手，并说出在哪儿动手。** 每次开工第一条命令是 `status`，把"几个实例在跑 / 开着哪些文档 / 哪份是 active / 有没有未保存修改"念给用户，并明确报出**即将修改的那份文档的完整路径**（Excel 再加工作表名）。不碰用户没指定的文件。
2. **`readOnly` 就是不许改。** `open`/`read` 返回 `readOnly: true`（文件被别处占用或有只读属性）时，写命令会直接报 `READONLY`。**不要**用另存、改文件属性、杀占用进程等方式绕过去——先问用户。
3. **不替用户关文档。** `close` 只关本次明确要关的那一份；用户开着的其它文档、用户自己的 WPS 窗口一律不动。关之前先 `status` 确认这份是本次打开的。
4. **结束必须 `close`，并汇报进程数。** 收尾命令 `close -Doc <路径>` 会返回 `processes`（Word 侧 wps.exe / Excel 侧 et.exe / PPT 侧 wpp.exe 计数）、`documentsLeft`。把这两个数报给用户；只清理自己起的无窗口 `/Automation` 实例，**绝不**按映像名批量杀进程。（注：WPS 退出后会自己预拉起 1–2 个隐藏进程，约 1 分钟自灭，别把它当成泄漏。）
5. **交付前提醒"落盘可能是密文"。** 装了单位透明加密的机器上，WPS `save` 之后文件头变成 `17DA5FA01633CD9A`，直接发给同事打不开。保存后必须提这一句，并按 [reference/encryption.md](reference/encryption.md) 处理。
6. **落盘前先想清楚反悔手段，别改完就存。**
   - **Word 修订类**（`revise`/`comment`）：拒绝修订不受 save 影响，可以及时存；
   - **Word 格式类**（`format`/`style`/`page`/`table`/`font`/`picture`/`caption`）：不进修订，唯一反悔手段是撤销栈，而 **save 会清空撤销栈（实测）**——必须在**用户看过/认可**或 `render` 复核通过**之后**才 save；
   - **Excel 一切写操作**：COM 修改不进撤销栈，改前 `snapshot`（`write`/`fill` 已自动落 pre-batch 快照）、改后比对 before/after 与 `diff`，别指望 Ctrl+Z；
   - **PPT 一切写操作**：同 Excel，无修订层、无撤销；批量改前 `snapshot`、改后 `diff` + `render` 目检；
   - `snapshot`/`diff`/`handoff` 都走 COM 读活文档，不依赖 save，不构成提前落盘的理由。
7. **先分清"只连接"和"要动手"，别把连接做成体检。** 用户只说"连上 / 接管 / 打开这份文档 / 先看看有什么"时，本轮做到 **`status` → `open` → 报一句摘要**（文件名 / 页数或段落数 / 纸张或尺寸 / `readOnly` / 有无未保存改动）**就停下等指令**：不 `read` 全稿、不 `render`、不 `audit`、不 `snapshot`、不 `close`（文档留给用户继续用，等他下一次给出任务时再按铁律 4 收尾）。这些动作只在用户给出具体任务时才做，且只做该任务需要的那一部分——大文档 `read` 一次吐上万字符、`audit` 一次吐上百条（人工老稿误报居多），对"只是连上"的用户是纯等待成本。
   反过来同样成立：用户给了"改"的任务，别在没改完、没 `render` 复核之前就宣布完成。

---

# Word：在 WPS Writer 里实时改文档

命令前缀：`office.ps1 word <命令>`（或直接 `word.ps1 <命令>`）。

## Word 命令速查

| 命令 | 参数 | 作用 |
|---|---|---|
| `doctor` | — | 自检：哪个套件可用/版本/是否在跑/开着的文档/渲染是否可用/各类进程数 |
| `status` | — | 在跑的实例与文档清单（name / path / saved / readOnly / active / paragraphs） |
| `open` | `-Doc <路径> [-Show] [-ReadOnly]` | 已开则复用，未开则打开；返回 readOnly、段落数、页数；发现可能开在另一实例时给 `multiInstanceWarning` |
| `read` | `-Doc <路径> [-From 1] [-Count 60]` | 段落（编号/样式/自动编号/文本/inTable）+ 表格清单 + 页数/修订数/批注数 |
| `search` | `-Doc <路径> -Text <词\|词2\|词3> [-Style 样式名] [-Max 50] [-Case]` | **全文找词**（WPS 原生 Find，大文档秒级；多词用 `\|` 分隔一次查完，返回每词计数）——定位用这个，别一页页 `read` |
| `selection` | `-Doc <路径>` | **读用户当前选中/光标**：返回选中文本与覆盖的段落号（`fromPara/toPara`）；在表格里时再给 `table`（表号，直通 `table -Table n`/`cell -Table n`）与 `row`/`col`；覆盖**图片**时给 `images`（图号，直通 `picture -Image n`）——只读不改 |
| `revise` | `-Doc <路径> -Items <json文件>` | **修订模式**批量改写（见下），返回改动与修订计数 |
| `comment` | 列表：`-Action list`；单条：`-Para N [-Anchor "文字"] (-TextFile <txt文件> \| -Text "短文本")`；**批量：`-Items <json>`**；删除：`-Delete -Index N \| -All` | 加批注（中文正文走文件，避开引号转义）；`list` 返回每条的编号/作者/锚文/内容——编号就是 `-Delete -Index` 用的 |
| `render` | `-Doc <路径> -Page 1 [-To 4] -Out <目录> [-Width 1100]` | 页面 EMF → PNG（宽 320–2400），返回每页路径与尺寸 |
| `font` | `-Doc <路径> -Name <字体名> [-Track]` | 把**正文、页眉、页脚、文本框**里的字体统一改成指定字体（中英文一起改，`-Name 仿宋`）；返回每个区段的结果与回读值。默认**不产生修订**，整批算一个撤销单元 |
| `format` | `-Doc <路径> [-Para N \| -From N -To M \| -Style 样式名 \| -Match <正则>]` + 任意格式项 | 对**整篇/一段/一段区间/某样式的所有段/文本匹配正则的所有段**套字符与段落格式（`-Match` 例：`"^\s*(图\|附图)\s*\d"` 一次选中全部图名；返回命中数与前 5 个样例） |
| `style` | `-Doc <路径> -Style <样式名> [-NoNumbering]` + 任意格式项 | 改**样式本身**（正文、标题 1…），所有用它的段落一起变 |
| `page` | `-Doc <路径>`（**不带任何选项=查询各节设置**）`[-Paper A4] [-Orientation portrait\|landscape] [-Top/-Bottom/-Left/-Right 厘米] [-HeaderText] [-FooterText] [-PageNumber center\|right\|left\|none] [-PageNumberStart N\|none] [-All]` | 纸张/方向/页边距/页眉页脚/页码（逐节应用）；查询模式报每节纸张/方向/边距/页眉页脚 **+ `pageNumbers`（起始号/是否节内重启/该节首段的实际显示页码——页码断号一眼看出）**；`-PageNumberStart none` 试图清掉「本节重新编号」（WPS 常做不到，会**明说失败**并给 XML 路线） |
| `table` | 插入：`-Para N -Data <json> [-Borders grid\|thick\|three-line\|none] [-Widths "2,3,2"] [-HeaderRow] [-Caption]`；改表：`-Table n [-Borders ...] [-Autofit] [-RowHeight] [-KeepTogether] [-Font/-Size/-Align]`（**省略 -Table = 全文档所有表，一次 COM 遍历 + 单个撤销记录**）；**体检 `-Action scan`**；**表头跨页 `-Action header-repeat [-Table n] [-Row 2] [-Passes 3]`**；**题名出框 `-Action caption [-Table n] [-Borders thick] [-Force]`**；**删行 `-Action rows -Table n [-Row 1]`** | 建表 / 改表 / 批量体检 / 表头重复 / 把做进第 1 行的题名提到框外并删该行 / 删指定行（后三者内部绕过 WPS 的 `Rows.Item(n)` 残缺对象，走选区写法） |
| `caption` | `-Para N -Label 表\|图 -Title "..." [-Above] [-Align center]` | 生成**域**式题注「表 1 xxx」（自动编号，增删后重排） |
| `picture` | 插图：`-Path <图片> -Para N [-Width "12cm"] [-Caption] [-Align]`；调整：`-Image n [-Width/-Height] [-CropLeft/-CropRight/-CropTop/-CropBottom 百分比]` | 插入图片（可带图题）/ 调尺寸与裁剪 |
| `toc` | `-Doc <路径> [-Para N] [-Title 目录] [-Levels 3]` | 插入并刷新目录（域），可从文首或指定段之前 |
| `reference` | 列表：`-Action list -Items <json 数组> [-AtPara N] [-Heading 参考文献]`；引用：`-Action cite -Para N -Numbers "1,2,5"\|"3-6"` | 生成参考文献列表 / 在段末插入上标引用 |
| `view` | `-Doc <路径> -Para N` 或 `-Page N` | 把 WPS 窗口滚到那一段/那一页（唯一能做的"跟随"，不动文档内容） |
| `unit` / `units` | `-Doc <路径> [-Apply] [-Units "m3,km2,s2"] [-Normalize] [-Track]` | **单位上标批量修**：把掉下来的 `m3/s`、`km2`、`万m3` 里的指数改成 Unicode 上标（`m³/s`、`km²`）。不带 `-Apply` = 只体检，带 `-Apply` = 修并报逐单位处数与跳过项。**只修"挂在数值上"的单位**（前字符＝数字/万/亿/汉字/左括号，或"数字 空格 单位"）；**已正确的两种形态自动跳过**（Unicode ²/³ 与 run 级上标）；**英文/数字贴连跳过**（`item3`、`m23`、`km20`）。三类只报不改：`fragments`（孤立 token／后跟 `=`，多半是模板变量名）、`looseFragments`（`n2`/`n3` 缺基名）、`suspects`（`km3` 这类疑似笔误）。300 页/13k 段实测 ~20 s |
| `audit` | `-Doc <路径> [-Max N] [-Action 规则名,逗号分隔]` | **机械检查**：单位未用上标、重复词、括号混用、连续空格、全角数字、段末空格、标题层级跳级、图表缺题注、连续空段、**题名做进表格第 1 行**（`table-caption-in-row`，直接给出该用哪条 `table -Action caption`）；返回段落号 + 建议，供你判断后落批注 |
| `fields` | `-Doc <路径>` | **刷新全部域**：页码、目录、题注、引用（改完正文、交付前跑一次） |
| `cell` | `-Doc <路径> (-Table n -Row r -Col c \| -Para N) -Text "..."` + 任意格式项 | 改一个表格单元格的文字（可带字体字号对齐） |
| `insert` | `-Doc <路径> [-Para N] [-Where after\|before\|end] -Items <json> [-Track]` | **批量插入段落**（默认干净写入、不进修订；`-Track` 才走修订），可带样式与格式。**锚点是表格前一段时自动改用"段末标记之前插入 CR"的写法**，不会掉进表格第一格 |
| `revisions` | `-Doc <路径> [-Action list\|accept\|reject] [-Index N \| -All]` | 列出**未处理**的修订（类型/作者/所在段/文本），或按人要求接受/拒绝 |
| `snapshot` | `-Doc <路径> [-Tag 名字] [-Out <路径>]` | 存一份段落文本快照（给 `diff` 用，默认存 %LOCALAPPDATA% 不污染文档目录；`-Out` 写到指定路径，跨机器/交接时随文档走） |
| `diff` | `-Doc <路径> [-Since 快照路径\|last]` | 和快照比：哪些段被改了（before→after）、哪些增删了——**人改的和 Agent 改的都算** |
| `handoff` | `-Doc <路径> [-Note "本轮做了什么"] [-As 输出路径]` | 在文档旁生成 `.handoff.md` 交接单：状态/待办/操作记录/最新快照路径 |
| `undo` | `-Doc <路径> [-Count N]`（默认 1，上限 20） | **Agent 自己的 Ctrl+Z**：撤回最近 N 步（排版批是一步）。注意 Undo 不分作者——只在自己刚写完且无 `changedOutside` 时用 |
| `save` | `-Doc <路径> [-As <另存路径>]` | 保存 / 另存 |
| `close` | `-Doc <路径> [-Save]` | 关闭；无其它文档时退出实例，并回报进程数 |
| `quit` | — | 关掉该实例的**全部**文档并退出（不保存）。**慎用**：会波及用户在同一实例里开的文档，只在自己的独占实例上用 |
| `finish` | `[-Note "做了什么"] [-Toast] [-Sound] [-Quiet]` | **任务完成信号**（agent 无关，零依赖）：默认**托盘气泡 + 提示音**（原生 .NET），**不触碰文档任何内容**；`-Quiet` 静音。**任务收尾默认调用** |

所有命令都可用 `-TimeoutSec <秒>`（默认 180）控制等待；超时返回 `TIMEOUT`，说明 WPS 大概率弹了对话框。

### `revise` 的 items 格式

```json
{"items":[
  {"para":12,"expect":"3.2 工程地质","op":"replace","find":"孔口高程","replace":"孔口地面高程"},
  {"para":15,"expect":"根据地质勘察","op":"set_text","text":"根据地质勘察成果，坝址区基岩为……"},
  {"para":20,"expect":"4.1 施工组织","op":"insert_after","text":"4.1.1 施工条件"},
  {"para":31,"expect":"本阶段不再","op":"delete"}
]}
```

- `op` 四种：`replace`（字符级，**优先用**）、`set_text`（整段删+插）、`insert_after`、`delete`。
- `expect` 必填，校验该段文本是否**以它开头**，不符即 `ANCHOR_MISSING`（防改错位置）；`replace` 的 `find` 找不到也报同一码。
- `replace` 默认只替换该段首个命中；同一段内要多处替换，加 `"all":true`。
- 批内自动按 `para` **从大到小**执行，编号不会互相串；要插多段时按原始编号一次写全。
- **整批是原子的**：动手前先把每一条都核一遍（段落存在、`expect` 对得上、`find` 找得到），有一条不过就**一条都不改**并报 `ANCHOR_MISSING`；万一执行中途出错，已产生的修订会被逐条 `Reject` 回滚，错误信息里会写"rolled back N revision(s)"。
- 每条结果带 `textAfter`（该段改完后的开头 90 字）——**立刻比对它**，能第一时间发现"替换后变成「进行了了」"这类找错位置的问题，不必再 `read`。
- 返回 `applied`、`revisionsBefore`、`revisionsAfter`、`revisionsAdded`。
- **段落编号的坑**：表格里每个单元格算一段，且每行末尾**再多算一段**（3×4 表占 15 个段落号），所以表格附近的段落号不能按"行×列"推算；`read` 会给表内段落打 `inTable: true`。找段号优先用 `search`。

## Word 格式项（`format` / `style` / `table` / `caption` / `cell` 通用）

| 参数 | 取值 | 说明 |
|---|---|---|
| `-Font` | `仿宋` | 中英文一起换（`NameFarEast`/`Name`/`NameAscii`/`NameOther`） |
| `-LatinFont` | `"Times New Roman"` | 只换西文，不动中文字体 |
| `-Size` | `12` / `"10.5pt"` | 字号（磅） |
| `-Bold` `-Italic` `-Underline` | 开关 | 关掉用 `-Bold:$false` |
| `-Super` `-Sub` | 开关 | 上标 / 下标 |
| `-Color` | `"#C00000"` | 字体颜色 |
| `-Align` | `left\|center\|right\|justify` | 对齐（中文也认 `居左/居中/居右/两端对齐`） |
| `-IndentChars` | `2` | **左缩进 2 字符**（中文单位，比磅值可靠） |
| `-LeftIndent` | `"0.74cm"` / `"21pt"` | 左缩进（长度） |
| `-FirstLineIndent` | `2` / `"0.74cm"` / 0 | **首行缩进**：纯数字=字符数（2=两字），带单位=长度，0=取消 |
| `-LineSpacing` | `1.5` / `"20pt"` / `"at least 20pt"` / `single` / `double` | 数字=倍数，`pt`=固定值，`at least`=最小值 |
| `-SpaceBefore` `-SpaceAfter` | `"6pt"` / `0` | 段前 / 段后间距 |
| `-PageBreakBefore` | 开关 | 该段另起一页 |
| `-Track` | 开关 | 让**格式改动也进修订**（默认不进；整篇排版不要开） |

## Word 典型工作流（详见 reference/workflows.md）

1. **读+核对版式**：`status` → `open -Show` → `read` → `render`（看表格出框/图表贴题/空段半空页）。
2. **定位**：`search -Text` 拿段落号与上下文（全文找词，取代一页页 `read` 翻）；用户说"就改我选中的那段"时用 `selection` 直接拿他选中的段落号。
3. **润色**：`search`/`read` 拿锚点 → 造 items（`replace` 优先）→ `revise` → 比对 `textAfter` → `render` 抽查 → 用户看过再 `save`。
4. **校审批注**：`audit` 机械检查 → `read` 被点名段自己判断 → `comment -Items` 一次落批注 → `render` 复核气泡 → `save`。
5. **统一字体**：`font -Name 仿宋`（看 `readBack` 回读是否落上）。
6. **一键排版**：先 `style` 后 `format`，再 `page`；排版前后各 `render`；别开 `-Track`。
7. **表格/题注/图片**：`table`（三线表+表头跨页）/ `caption`（SEQ 域，表上图下）/ `picture`（插后必 render）。
8. **目录/引用**：`toc`（依赖标题样式）/ `reference list+cite` / 交付前 `fields` 刷全部域。
9. **协作接力（核心用法）**：每轮 `snapshot` → 干活 → 人裁决 → `diff`+`revisions` 读结果 → `handoff` 交接单。跨机器接力用 `snapshot -Out` 把快照随文档带走。写命令结果里的 `changedOutside` = 文档被人动过，先 `diff` 再继续。
10. **批量整理表格（中文报告高频）**：`table -Action scan` 先出体检表（哪张表题名在框内、哪张跨页没表头、边框粗细）→ 按结果选做：
    - 题名在框内：`table -Action caption -Table n -Borders thick`（提题名 → 删该行 → 加边框，顺序已在命令内部保证）
    - 跨页缺表头：`table -Action header-repeat`（自动迭代到稳定；**先**处理题名，否则重复出去的是题名行）
    - 全文档边框加粗：`table -Borders thick`（**不带 -Table 即全表**，一次 COM 遍历 + 一个撤销记录；85 张表约十几秒，别逐张调）
    - 细节见 [reference/workflows.md](reference/workflows.md) 的「批量整理表格」一节
11. **单位上标批量修（几百页报告的常见活）**：`unit -Doc X` 先看体量 → `unit -Doc X -Apply` 一键修（已正确的跳过）→ `render` 抽 1–2 页确认 → `save`。
    - 判据：`unit` 的 `plainDigitHits` − `unicodeOk` 不等于坏例数，**坏例数以 `-Apply` 返回的 `fixed` 为准**（只有 Find 能看到 run 格式，文本层分不出"真掉"和"格式上标"）。
    - `skippedSamples` 里若出现疑似真单位，用 `-Units` 补上重跑；若是英文词（`item3`、`com2`）则不用管。
    - **`scriptPages` 先看**：命中脚本页（公文模板/脚本说明）时命令已整段跳过，若它恰好是你要修的那几页，只能人工处理（那里的 `m2`/`km2` 是变量名）。
    - **`fragments` / `looseFragments` / `suspects` 三样都必须看一遍**：`fragments` 多半是**模板/脚本示例里的变量名**（`km2 = 1`、`<m1，m2，m3>`），`looseFragments` 是缺基名的形态（`n2`、`n3`，常是 `m2`/`m3` 打错），`suspects` 是疑似笔误（面积表里的 `km3` 多半该是 `km2`）。**这三类一律只报不改**——自动改会把变量名/笔误变成"看起来是对的"上标，比不改更糟。报给作者确认后再用 `revise`/`cell` 处理。
    - 若 `fixed` 明显少于预期（例如只剩个位数），先看 `skipped.assignment/isolated`：说明命中多半在模板示例或表格标题里，符合预期。
    - 全文见 [reference/workflows.md](reference/workflows.md) 的「单位上标批量修」一节。
10. **反悔**：自己刚写完且无 `changedOutside` 时 `undo -Count N`（排版批=1 步）。

## Word 故障排查

| 错误码 / 现象 | 怎么办 |
|---|---|
| `NO_APP` | 没装 WPS 或 COM 没注册。先手动打开一次 WPS Writer 再重试；`doctor` 里 `apps.wps.installed=false` 说明这台机器确实没有 WPS。 |
| `NOT_OPEN` | `-Doc` 与实际打开的文档不是同一路径（另存过/短路径/大小写/相对路径）。先 `status`，把 `path` 原样复制过来。 |
| `READONLY` | 文件被别处占用或是只读属性。关掉占用它的程序，或请用户另存一份再改；不要强行覆盖保存。 |
| `ANCHOR_MISSING` | 段落编号漂移（表格也占段落号，且每行多一段）或 `expect`/`find` 与原文不符。整批**一条都没改**，先 `search`/`read` 该段核对编号与原文再来；并发协作时先怀疑"人刚改了这段"。 |
| `BUSY` / `TIMEOUT` | 用户正在打字、或有对话框/弹窗挡着（等满 `-TimeoutSec` 秒即超时）。等几秒或请用户关掉对话框再重试，别连发。 |
| 打开即 `COM_ERROR` / 文件读不出来 | 多半是密文且不归 WPS 管（例如离线脚本写出的）。让用户在 WPS 里打开一次（驱动会透明解密），或按 [reference/encryption.md](reference/encryption.md) 先解密。注意：helper **不产出** `ENCRYPTED` 码。 |
| `open` 返回 `multiInstanceWarning` | 文档可能同时开在另一个 WPS 实例里（本 skill 只能接到第一个注册实例）。请用户在 WPS 里关掉另一份再继续，否则两份会互相覆盖。 |
| 结果里出现 `reopened` | 文档刚才被人在 WPS 里关掉了，本命令已从磁盘**静默重开**（未弹窗）。若同时有 `unsavedAtClose: true`，说明关闭时未保存——关闭前的改动（含本 skill 的修订/批注）**可能已丢失**：先 `diff` 快照对账、向用户确认后再继续。 |
| 中文乱码 | 控制台代码页不是 UTF-8。先 `[Console]::OutputEncoding=[Text.Encoding]::UTF8`，或把 JSON 落成文件按 UTF-8 读。 |
| `wps.exe` 变多 | 中途异常没 `close`，或别处 `New-Object` 开了新实例。先 `status` 看实例数，只结束无窗口的 `/Automation` 实例，用户窗口留着，然后汇报数字。 |
| Ctrl+Z 撤不干净批量改动 | 三个前提：①改动必须走本 skill 的命令（自带 UndoRecord 分组，实测 WPS 支持；**绕过 skill 手写 COM 循环 = 每段一条撤销记录**）；②**save 会清空撤销栈**，要反悔必须在保存前；③格式套成与原值相同的段不产生撤销记录（再撤也不会变）。 |
| `undo` 撤不动 | 撤销栈已被 save 清空或已见底；文本类改动优先走修订（`revise`），随时可拒绝，不依赖撤销栈。 |
| `table` 报 `0x800A1767` / 行操作炸 | WPS 的 `Row` 对象残缺（`Rows.Item(n).Delete/.HeadingFormat/.Range/.Cells` 都不可靠）。**不要自己写 `Rows.Item(n)` 循环**——用 `table -Action rows/header-repeat/caption`，内部走"选中整行 → `Selection.Rows`"的写法 |
| `insert -Where after` 的段落进了表格 | 已修：锚点后紧跟表格时，helper 改在段末标记之前插 CR。若仍见到 `inTable:true` 而锚点不在表格里，先 `undo` 再报告 |

---

# Excel：在 WPS 表格里实时改工作簿

命令前缀：`office.ps1 excel <命令>`（或直接 `excel.ps1 <命令>`）。

## Excel 命令速查

| 命令 | 参数 | 作用 |
|---|---|---|
| `doctor` | — | 自检：ProgID 安装/在跑、进程数、渲染能力、打开的工作簿 |
| `status` | — | 在跑的实例与工作簿清单（name / path / saved / readOnly / active），工作表概况 |
| `open` | `-Doc <路径> [-Show] [-ReadOnly]` | 已开则复用，未开则打开；返回工作表清单与各自 usedRange |
| `read` | `-Doc` `[(-Sheet 名 \| -Range "Sheet1!A1:F40")] [-Profile]` | 不带选项=工作簿概览；带 Sheet=读该表 used range（值 + 公式）；`>60 行或 >15 列`自动改列概况（类型统计/min/median/max/高频值 + 头尾各 5 行样本），要局部全量用 `-Range` |
| `write` | `-Doc -Items <json>`（或单格 `-Cell B3 -Value x / -Formula "=..."`） | **批量写**（见下）：原子批、`expect` 锚校验、中途失败回滚，每格返回 before→after；**批内自动挂起重算+屏幕刷新，批末一次重算**（大表写几千格不产生重算风暴） |
| `fill` | `-Doc -Range D2:D5000 [-Sheet] -Formula "=B2*C2"` | **整列填公式**：首格写一次 + 原生 FillDown（2000 行约 2 秒，含进程启动）；比逐格 write 快几个量级，大表必用 |
| `sheet` | `-Doc -Action list\|add\|rename\|delete [-Name] [-NewName]` | 工作表增删改名 |
| `comment` | `-Action list`；加：`-Cell D5 (-Text "短文" \| -TextFile <txt>)`；删：`-Action delete -All` | 单元格批注（校审指认用；中文长文走文件避开引号转义） |
| `search` | `-Doc -Text <词> [-Sheet 名] [-Max 50] [-Case]` | **找格**：跨表搜值和公式，返回命中格/值/公式（定位用这个，别整表 read） |
| `selection` | `-Doc` | **读用户当前选中的区域**：sheet/range/行列数/值（≤400 格直接带值）——用户说"就改我选的这块"时用；选在别的窗口报 NOT_OPEN |
| `rows` / `cols` | `-Doc [-Sheet] -Action insert\|delete -Index N [-Count M]` | 插入/删除 N 起的 M 行（列）——结构编辑；**先 snapshot**（可回滚），公式由表格自动调整 |
| `render` | `-Doc [-Sheet] [-Range] -Out <目录> [-Width 1100]` | 区域 `CopyPicture` → PNG（表格没有分页概念，核对用区域图） |
| `chart` | 加：`-Range 数据区 [-Type column\|bar\|line\|pie\|scatter\|area] [-Title] [-Anchor E2]`；`-Action list`；`-Action delete -Index n` | 在工作表上建原生图表（480x300pt，锚在指定格）；建后**必须 render 该区域目检**再保存 |
| `autofit` | `-Doc [-Sheet] [-Range]` | 列宽行高自适应 |
| `format` | `-Doc -Range A1:D1 [-Sheet] + 格式项` | 字体/字号/加粗/颜色/填充/对齐/数字格式/列宽/行高/框线/换行 |
| `snapshot` | `-Doc [-Sheet 名1,名2] [-Out <路径>]` | 存工作表的值+公式快照（默认全部表；大表用 `-Sheet` 限定范围，diff/restore 也只在该范围内比对/回滚，不碰未覆盖的表；默认存 %LOCALAPPDATA%，`-Out` 随文档带走） |
| `diff` | `-Doc [-Since 快照路径\|last]` | 和快照比：哪些格改了（before→after，含公式联动）、哪些增删——**人改的和 Agent 改的都算** |
| `restore` | `-Doc [-Since 快照路径\|last] [-Cells "Sheet1!C2,Sheet1!E2"] [-KeepExtras]` | **一键撤回**：把工作簿还原到快照状态（只改有差异的格；快照后新增的非空格默认清空，`-KeepExtras` 保留）。`-Cells` = **选择性回滚**，只还原点名的格（快照中没有的点名格清空），其余一切——包括人改的——原样保留。只还原内存，`diff` 验证后再决定 `save` 或不保存直接 `close` |
| `calc` | `-Doc` | 全簿重算 |
| `save` | `-Doc [-As <路径>]` | 保存 / 另存（.xlsx/.xls/.csv/.xlsm；`.pdf` = 整簿导出 PDF） |
| `close` | `-Doc [-Save]` | 关闭；无其它工作簿时退出实例，回报进程数 |
| `quit` | — | 关掉该实例**全部**工作簿并退出（不保存）。**慎用** |
| `finish` | `[-Note "改了什么"] [-Cell A1] [-Toast] [-Sound] [-Quiet] [-Force]` | **任务完成信号**（agent 无关，零依赖）：默认**托盘气泡 + 提示音**（原生 .NET，任何机器都响），**不动工作簿任何内容**。`-Cell` 显式指定时才额外写一行状态（目标格非空拒绝，`-Force` 覆盖）；`-Quiet` 静音。**任务收尾默认调用** |

所有命令都可用 `-TimeoutSec <秒>`（默认 180）；超时返回 `TIMEOUT`，说明 WPS 大概率弹了对话框。

### `write` 的 items 格式

```json
{"items":[
  {"sheet":"数据","cell":"E1","value":"avg"},
  {"sheet":"数据","cell":"E2","formula":"=AVERAGE(C2:C3)"},
  {"cell":"'汇总'!A1","value":"合计"},
  {"sheet":"数据","cell":"F1:G2","values":[[1,2],[3,4]]},
  {"sheet":"数据","cell":"C2","value":1875.5,"expect":"1850.5"}
]}
```

- 三种写法：`value`（标量）、`formula`（以 `=` 开头，交给表格算）、`values`（二维块，尺寸必须与 `range` 一致）。
- 定位：`cell` 单格或 `range` 区块；工作表用 `sheet` 字段，或直接写 `"Sheet名!A1"` 引用（名字带空格/特殊字符用单引号包住）。
- `expect` 可选但强烈建议：该格**当前显示文本**必须与它完全相等才动手，不符即 `ANCHOR_MISSING`、**整批一条不改**——防编号漂移、防"人刚改了这格"。
- 整批原子：全部校验通过才开写；万一中途 COM 报错，已写的格用捕获的旧 Formula **逐格还原**，错误信息里写"rolled back N cell(s)"。
- 每条返回 `before`/`after`（写入前后该格显示文本）——**立刻比对**，能第一时间发现写错位置。
- **撤销**：表格的 COM 修改不进撤销栈且会清掉用户自己的撤销记录——写前 `snapshot`、写后看 before/after，别指望 Ctrl+Z。

### 大表效率纪律（几千行以上必读）

1. **连续数据拼成 `values` 块**一次写（`{"cell":"A1:F5000","values":[...]}`），绝不逐格 item；
2. **整列公式用 `fill`**，不用 write 逐行；
3. **快照限定范围** `snapshot -Sheet 数据表`，避免几十 MB 的全簿快照；
4. 读大表交给自动概况（>60 行自动列概况+头尾样本），要局部全量用 `-Range`；
5. write/restore 已自动挂起重算+屏幕刷新、批末一次重算，无需手动干预。

## Excel 格式项（`format`）

| 参数 | 取值 | 说明 |
|---|---|---|
| `-Font` / `-Size` | `仿宋` / `12` | 字体 / 字号（磅） |
| `-Bold` `-Italic` `-Underline` | 开关 | `-Bold:$false` 关闭 |
| `-Color` `-FillColor` | `"#C00000"` / `"#DDEBF7"` | 字色 / 底色 |
| `-HAlign` | `left\|center\|right\|justify` | 水平对齐 |
| `-VAlign` | `top\|center\|bottom` | 垂直对齐 |
| `-Wrap` | 开关 | 自动换行 |
| `-NumberFormat` | `"0.00"` / `"0.0%"` / `"yyyy-mm-dd"` | 数字格式 |
| `-ColWidth` / `-RowHeight` | `12` / `20` | 列宽（字符）/ 行高（磅） |
| `-Borders` | `grid\|thick\|none` | 全边框 / 粗边框 / 去边框 |

## Excel 典型工作流

1. **接管**：`status` → `open -Show` → `read`（概览）→ 需要数据时 `read -Sheet X` 或 `read -Range "X!A1:F200"`。
2. **定位**：`search -Text` 拿命中格与上下文（全文找格，取代整表 read 翻找）；用户说"就改我选中的这块"时用 `selection` 直接拿他选中的区域。
3. **改表**：`write`/`fill` **每次自动先落 pre-batch 快照**（返回值 `preSnapshot.path`，无需手动 snapshot）→ 比对每条 before/after → `diff -Since <path>` = 本批改动归因 → `render` 抽查 → 用户看过再 `save`。改坏了 `restore -Since <path>` 整批回滚，或 `restore -Since <path> -Cells "表!C2"` 只撤点名格（人改的其它格不受影响）。
4. **改结构**：`rows`/`cols` 插删行列（**先 snapshot**，结构命令暂无自动快照）→ `fill` 补公式 → `read` 复核。
5. **校审指认**：`comment -Cell D5 -Text "..."`（或批量逐条加）→ `render` 看角标 → `comment -Action list` 复核。
6. **排版**：`format`（表头加粗+填充+居中+边框）→ `autofit` → `render` 复看。
7. **新表**：`sheet -Action add -Name 汇总` → `write`（引用写 `'汇总'!A1`）→ `calc` → `render`。
8. **图表**：`chart -Action add -Type line -Range 数据 -Title ... -Anchor E2` → `render` 目检 → 不满意 `-Action delete` 重来。

## Excel 故障排查

| 错误码 / 现象 | 怎么办 |
|---|---|
| `NO_APP` | 没装 WPS 或 KET COM 没注册。手动开一次 WPS 表格再试；`doctor` 里 `apps.wpsEt.installed=false` 说明这台机器没有。 |
| `NOT_OPEN` | `-Doc` 与实际打开的工作簿不是同一路径（另存过身份会变：`save -As` 之后按新路径操作）。先 `status` 把 `path` 原样复制过来。 |
| `DOC_CLOSED` | **工作簿在协作期间被关闭了**（很可能是用户手动关的，未保存的改动随关闭丢失）。**不要静默继续**：先向用户确认，确需继续就显式 `open -Doc` 重新接管。skill 记住了接管过的工作簿（attachment marker），从关闭到显式 open 之间的任何命令都会得到这个错误，而不是悄悄把文件复活 |
| `READONLY` | 文件被占用或只读打开。让用户关掉占用程序或另存，不要强写。 |
| `ANCHOR_MISSING` | `expect` 与该格当前显示不符（人刚改了 / 定位错表）。整批没改，重新 `read`/`write` 核对。 |
| `BUSY` / `TIMEOUT` | 用户正在单元格里输入、或有对话框挡着。等几秒再试，别连发。 |
| 文件打不开/读不出 | 多半是密文且无白名单进程可读。让用户在 WPS 里打开一次（驱动透明解密），或按 [reference/encryption.md](reference/encryption.md) 处理。 |
| `write` 报 rolled back | 批中途 COM 失败已自动还原；看错误信息，通常用户正好在编辑。重试或分小批。 |
| 中文乱码 | 控制台代码页问题；本 skill 输出已是 UTF-8，若外层管道乱码，先 `[Console]::OutputEncoding=[Text.Encoding]::UTF8`。 |
| et.exe 变多 | 中途异常没 `close`。`status` 看实例，只清理无窗口实例，汇报数字。 |

---

# PPT：在 WPS 演示里实时改演示文稿

命令前缀：`office.ps1 ppt <命令>`（或直接 `ppt.ps1 <命令>`）。ProgID `KWPP.Application`；`open` 一个不存在的路径会**新建并立即保存**一份空稿到该路径（新稿身份靠这次落盘确立）。

**关键限制（先读）**：PPT 无修订层、外部 COM 改动不进撤销栈（同 Excel）；**WPS 演示不接受公式**（不能建也不能贴）；`render` 走剪贴板（`Slide.Copy` → EMF → PNG，helper 自己写盘故是明文；`Slide.Export` 写的文件会被透明加密、读不回来，所以不用它）。

## PPT 命令速查

| 命令 | 参数 | 作用 |
|---|---|---|
| `doctor` | — | 自检：ProgID 安装/在跑、wpp 进程数、开着的演示文稿 |
| `status` | — | 实例与文稿清单（name / path / saved / readOnly / slides / active） |
| `open` | `-Doc <路径> [-Show] [-ReadOnly]` | 已开则复用；未开则打开；**路径不存在则新建空稿并保存到该路径**；返回 slides 与 slideSize |
| `read` | `-Doc [-From 1] [-Count 30]` | 逐页：layout + 形状清单（编号/名称/类型/占位符/l/t/w/h/文字前 160 字）+ 备注；>30 页自动分页 |
| `search` | `-Doc -Text <词> [-Max 50] [-Case]` | 跨页找文字：返回 页号/形状名/片段 |
| `selection` | `-Doc` | 用户当前所在页与选中的形状名（需窗口开着） |
| `slide` | `-Action add [-Preset title\|section\|bullets\|two\|end] [-Title ...] [-Subtitle ...] [-Number ...] [-Items <json>] [-Layout n] [-At N]`；`-Action dup\|move\|delete -Index N [-At\|-To M]` | **预设版式建页**（几何随文稿尺寸自适应；预设文本框已命名 title/subtitle/number/body/left/right，后续 `-Shape` 直达）；或复制/移动/删除页 |
| `text` | 单条：`-Slide N [-Shape 名\|-ShapeIdx n] (-Text \| -TextFile) [-Expect 原文开头]`；**批量：`-Items <json>`** | 改形状文字；`expect` 锚校验不符则**整批一条不改**（`ANCHOR_MISSING`）；默认解析到标题占位符，否则第一个有字形状——**强烈建议显式 -Shape**；每条返回 before/textAfter，**立刻比对** |
| `textbox` | `-Slide N (-Text\|-TextFile) [-Name 名] [-L -T -W -H pt]` + 格式项 | 加文本框（`-Name` 命名，后续定位用；中文字体/字号/对齐一起给） |
| `image` | `-Slide N -Path <图片> [-L -T] [-W/-H pt]` | 插图（缺省居中）；返回最终尺寸；图占页宽 <20% 时提醒字母可能太小 |
| `table` | `-Slide N -Data <json: {"rows":[[...]]}> [-L -T -W -H] [-Size 14]` | 建表并填内容（首行默认加粗表头） |
| `shape` | `-Action add -Type rect\|roundrect\|ellipse [-L -T -W -H] [-Fill 颜色] [-Line 颜色] [-Text ...]`；`-Action delete -Name <名>` | 画色块/圆角矩形/椭圆（可带字）；或按名删形状 |
| `format` | `-Slide N -Shape <名> `+ 格式项 | 改文字格式（字体/字号/加粗/斜体/颜色/对齐/填充），返回回读值 |
| `font` | `-Doc [-Font 宋体] [-From N] [-To M]` | **全稿统一字体**：逐页扫所有文字（**含组合内部与表格单元格**），只写 `Font.Name`+`NameFarEast`，**不动字号/颜色/加粗/版式**；返回 `shapes`/`groups`/`tables`/`failed` 与样例。⚠️ 换字体常使**字身变宽而多折行**，跑完必须 `render -Overview` 复核，个别挤爆的框用 `format -Size` 收一档 |
| `notes` | `-Slide N (-Text\|-TextFile\|-Clear)` | 写/清演讲者备注 |
| `view` | `-Slide N` | 把 WPS 窗口切到那一页（唯一"跟随"手段，不动内容） |
| `render` | `-Doc [-From 1] [-To 4] -Out <目录> [-Width 1100]`；**总览：`-Overview`（每次 ≤24 页，6 页拼一张并带页码）** | 逐页 EMF→PNG（**每次最多 4 页**）；写盘者是 helper，**明文可看**；`-Overview` 是"交付前过一遍全稿"的经济方式 |
| `theme` | `-Set <json 文件>` / `-Apply` / `-Reset` / 不带选项=查看 | **整稿风格打底**：primary/accent/title/text/muted 五色 + font/titleFont + footer（页脚文字）+ pageNumber + skipFooterOnCover。`-Set` 存到该文稿主题文件（%LOCALAPPDATA%\wps-office-cooperate\ppt-themes），之后所有预设页自动跟随；`-Apply` 把已有页的字体/配色刷一遍并补页脚页码（**落在深色色块上的文字自动改白**） |
| `audit` | `-Doc [-From/-To] [-MinFont 12]` | **机械体检**：文字溢出文本框、字号过小、形状越出页面、形状重叠（装饰色块除外）、图片 <20% 页宽、空页、无标题页。返回到 页/形状/类型/细节。⚠️ **人工制作的老稿误报率高**：整幅背景色块会报"重叠/越界"，标题用普通矩形会报"无标题页"，页面 logo 会报"图过小"——这些是设计意图，**结论以 `render` 目检为准**，audit 只当线索清单 |
| `align` | `-Slide N -Mode left\|right\|top\|bottom\|center-h\|same-width\|same-height\|distribute-h\|distribute-v [-Shapes "a,b,c"]` | 对齐/等距分布/统一尺寸（`-Shapes` 省略=该页全部形状）；返回每个形状的最终几何 |
| `zorder` | `-Slide N (-Name 名 \| -ShapeIdx n) [-Where front\|back]` | 置顶/置底（图片压字时用） |
| `restore` | `-Doc [-Since 快照\|tag] [-Shapes "3!备注,4!title"]` | **文本级回滚**：把文字退回快照状态（版式/新增形状不还原）；改坏了先回滚再想别的 |
| `snapshot` | `-Doc [-Tag 名]` | 存文本级快照（每页形状名+文字，默认存 %LOCALAPPDATA%\wps-office-cooperate\ppt-snapshots） |
| `diff` | `-Doc [-Since 快照\|tag]` | 与快照比：哪页哪个形状的文字变了（before→after）、页增删 |
| `save` | `-Doc [-As <路径>]` | 保存 / 另存（`.pdf` = 导出 PDF） |
| `close` | `-Doc [-Save]` | 关闭；最后一份时退出实例，回报 wpp 进程数 |
| `quit` | — | 全部关闭并退出（不保存）。**慎用** |
| `finish` | `[-Note "做了什么"] [-Toast] [-Sound] [-Quiet]` | 完成信号（托盘气泡+提示音）；**不碰文稿任何内容**。任务收尾默认调用 |

所有命令都可用 `-TimeoutSec <秒>`（默认 180）；超时返回 `TIMEOUT`，多半是 WPS 弹了对话框。

### `text -Items` 批量格式

```json
{"items":[
  {"slide":4,"shape":"title","expect":"谢谢","text":"谢谢聆听"},
  {"slide":3,"shape":"备注","text":"集水面积 1520 平方公里"}
]}
```

- `shape` 用 `read` 里看到的名称（预设页的形状已命名）；不填则解析到标题占位符，否则第一个有字形状。
- `expect` 可选但强烈建议：该形状当前文字必须**以它开头**，不符即整批拒绝。
- 中途 COM 出错会按序回滚已改项（"rolled back N item(s)"）。

### PPT 视觉自检闭环（**本轮真的改动过这份稿才走**）

前提是本轮确实动过这份稿（新建、改文字、换主题、调版式都算）。**只是连上、只是读、只是查，不走这一步**（见铁律 7）。顺序固定：

1. **`audit -Doc`** —— 机械筛查（溢出/字号/越界/重叠/空页），便宜但**对人工老稿误报多**（背景色块 / logo / 非占位符标题），先挑出真问题再修，别按条数硬修。
2. **`render -Overview -Out <目录>`** —— 每 6 页拼一张图，**逐张看**：标题层级是否统一、色块对不齐、要点与表格图挤在一起、半空页。
3. 对可疑页 **`render -From N -To M`** 单页放大确认。
4. 修（`text` / `align` / `table` / `theme -Apply`）→ **再 `render` 确认** → 才 `save`。

**能力不同时的降级**：当前模型能读图 → 第 2 步必须真看（`read_image`）；读不了图（某些模型不声明图像输入）→ 只能依赖 `audit` 并把 `render` 的结果路径交给用户看，**不要**声称"已核对版式"。

### PPT 典型工作流

1. **新建汇报稿**：`status` → `open -Doc 新路径 -Show`（自动建空稿）→ 可选 `theme -Set 主题.json`（一次定型：主色/字体/页脚）→ `slide -Preset title ...`（封面）→ `slide -Preset bullets -Items 提纲.json`（目录页）→ 逐节 `slide -Preset section -Number 01`。
2. **填充内容页**：`slide -Preset bullets` 打骨架（正文占 25%–75%，下方留给表格/图）→ `table -Data 表.json`（默认 90% 页宽、y=62%）→ `image -Path 图.png` → 每页 `render` 目检。
3. **改已有稿**：`read` 拿形状名 → `text -Items`（带 expect）批量改 → 比对 textAfter → `render` 抽查。
4. **收拾版式**：`align` 对齐/等距、`zorder` 修压字、`audit` 兜底；要统一字体用 `font -Font <字体>`（一次全稿，含组合/表格），**换字体后必 `render -Overview` 查折行溢出**，挤爆的框用 `format -Size` 收一档。
5. **人机接力**：`snapshot -Tag r1` → 干活 → `diff -Since r1` 看人改了什么；改坏 `restore -Since r1`。
6. **收尾**：`theme -Apply`（统一配色页脚）→ `audit` → `render -Overview` **看一眼** → `save`（提醒落盘可能是密文）→ `close`（汇报 wpp 数）→ `finish`。

## PPT 故障排查

| 错误码 / 现象 | 怎么办 |
|---|---|
| `NO_APP` | 没装 WPS 演示或 KWPP COM 没注册；手动开一次 WPS 演示再试。 |
| `NOT_OPEN` / `DOC_CLOSED` | 语义同 Excel：`-Doc` 路径与实际不符先 `status`；协作中被关闭不要静默重开，先问用户。 |
| `ANCHOR_MISSING` | 形状当前文字与 `expect` 不符（人刚改了/定位错形状）。整批没改，重新 `read` 核对形状名。 |
| `render` 报 `COM_ERROR`（clipboard） | 多半是弹了对话框占住剪贴板，或刚 Copy 完还没就绪。关掉对话框重试。 |
| 同页形状重名（旧稿"文本框 1"满天飞） | `-Shape` 会命中第一个；用 `-ShapeIdx`（read 里的 i 列）或先 `format` 不行就 `textbox -Name` 加命名框替代。 |
| 图片插入后糊/字小 | `render` 目检；图小于页宽 20% 时结果里会点名提醒。 |
| wpp.exe 变多 | 同其它 app：`status` 看实例，只清无窗口实例，汇报数字。 |

---

## 与其它 skill / 工具的关系

- **`office_*` 工具（dsh-office 插件）**：装了那个插件时，`office_open/office_read/office_edit/office_render` 与本 skill 能力重叠、互不冲突——插件适合"顺手改一份"，本 skill 适合"要分发给同事、不想让每人装插件"的场合。两者都可用时**选一个用到底**，别交替（会各开实例）。
- **`docx-paragraph-polish`**（离线润色）：**本机有 WPS 且用户要在场逐条过**时，改用本 skill Word 侧——修订直接出现在 WPS 窗口里、可当场接受/拒绝、顺手能核对版式；无人值守、上千段一次跑完的批量润色仍用 `docx-paragraph-polish`（不依赖 WPS，也不受"落盘即密文"影响）。
- **`office_xlsx` / `minimax-xlsx` / openpyxl 等**：它们改磁盘文件。工作簿正开在 WPS 里时**必须**用本 skill Excel 侧；无人值守批量跑、或文件没开且不需要人在场时，用离线方案更快。
- **`pptx-generator` / python-pptx 等**：同理——演示文稿正开在 WPS 里（或用户要在场边看边改）时用本 skill PPT 侧；从零生成、无人值守的整份 PPT 用离线方案。
- **`design-report-review` / `design-report-format`**：见铁律 0——同一份文档在本 skill 手里时，用本 skill 的 `audit` + `comment` / `format`/`style`/`page` 在原稿处理，不另起离线流程；用户明确只要旁挂报告、不动原文时才另说。

## 参考

- [reference/workflows.md](reference/workflows.md) —— Word 全部典型工作流分步教程 + 排版实测坑详解。
- [reference/word-quirks.md](reference/word-quirks.md) —— WPS Writer 与 Word 的实测差异表（冒充版本号、样式中文名、`Range` 是方法、关最后一个文档静默杀实例……）。
- [reference/excel-quirks.md](reference/excel-quirks.md) —— WPS 表格 COM 的实测差异与坑（冒充 Excel、SaveAs 身份变化、块数组 0 基、渲染走剪贴板……）。
- [reference/encryption.md](reference/encryption.md) —— 落盘即密文的现象、为什么 COM 路线不受影响、交付与解密链路。
- [reference/word-acceptance.md](reference/word-acceptance.md) / [reference/excel-acceptance.md](reference/excel-acceptance.md) —— 各版本验收记录。
