# wps-word-cooperate skill · 设计契约（v0.1）

> 目标：把「WPS Writer 里实时改 Word + 修订 + 批注 + 版式核对」做成一个**纯文件可分发的 skill**。
> 技术选型（已定）：**自研极简 C# helper**（`csc.exe` 首次编译，.NET Framework 4.8 随 Windows 自带，零第三方依赖）+ PowerShell 薄包装；**只做 Word，只针对 WPS**。
> 分发对象：同事（都是 WPS），通过 skill 文件夹拷贝或 `ghy-skill-manager` 安装。

## 1. 目录结构

```
wps-word-cooperate/
├── SKILL.md                     # 触发词、护栏、命令速查（v0.5 起长教程下沉 reference/workflows.md）
├── README.md
├── scripts/
│   ├── wps.ps1                  # 薄包装：编译缓存 → 喂 JSON → 读 UTF-8 → 打印
│   └── helper/                  # 单文件拆成的 partial class（纯 ASCII 源码，中文全走 JSON 字符串）
│       ├── WpsHelper.cs         # 核心：JSON/COM plumbing/开读改注渲染存关
│       ├── Layout.cs            # P1：format / style / page
│       ├── Tables.cs            # P2：table / caption / picture
│       ├── Toc.cs               # P3：toc / reference / view
│       ├── Text.cs              # v0.3：fields / cell / insert
│       ├── Review.cs            # v0.4：revisions / snapshot / diff
│       ├── State.cs             # v0.4：changedOutside 签名 + handoff
│       ├── Audit.cs             # 机械检查规则
│       └── Search.cs            # v0.5：search / undo
└── reference/
    ├── wps-quirks.md            # WPS 与 Word 的实测差异清单
    ├── encryption.md            # 落盘即密文：交付与解密链路
    ├── workflows.md             # 典型工作流分步教程 + 排版实测坑（v0.5 起）
    └── acceptance.md            # 各版本验收记录（v0.1–v0.5）
```

开发目录另有 `tests/`（语料+回归脚本）、`docs/`、`tools/`、`sandbox-*/`、`.tmp/`——**不进分发包**，用 `publish.ps1` 打包。
编译缓存：`%LOCALAPPDATA%\wps-word-cooperate\helper-<sha256前12>.exe`（源码哈希变则重编）。

## 2. 调用方式（模型看到的就是这个）

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File <skill>\scripts\wps.ps1 doctor
```

| 命令 | 参数 | 作用 |
|---|---|---|
| `doctor` | — | 自检：装了哪个套件、版本、是否在跑、开着的文档、渲染能力、进程数 |
| `status` | — | 在跑的实例与文档清单（name / path / saved / readOnly / active / paragraphs） |
| `open` | `-Doc <path> [-Show]` | 已开则复用，未开则打开；返回 readOnly、段落数、页数 |
| `read` | `-Doc <path> [-From 1] [-Count 60] [-Tables]` | 段落（编号/样式/自动编号/文本）+ 表格尺寸 + 页数 |
| `revise` | `-Doc <path> -Items <jsonfile>` | **修订模式**批量改写（见第 3 节），返回改动与修订计数 |
| `comment` | `-Doc <path> -Para N [-Anchor "文字"] -TextFile <txt>` | 加批注（文本走文件，避开引号转义） |
| `render` | `-Doc <path> -Page 1 [-To 4] -Out <dir> [-Width 1100]` | EMF → PNG，返回每页文件路径与尺寸 |
| `save` | `-Doc <path> [-As <path>]` | 保存 / 另存 |
| `close` | `-Doc <path> [-Save]` | 关闭；无其他文档时退出实例，并汇报进程数 |

**输出**：stdout 只有一个 JSON 对象（UTF-8），形如
```json
{"ok":true,"result":{...}}
{"ok":false,"error":{"code":"ANCHOR_MISSING","message":"..."}}
```
错误码：`NO_APP`（没装/没起来）、`NOT_OPEN`、`READONLY`、`ANCHOR_MISSING`（段落编号与 expect 不符 / find 找不到）、`BUSY`（应用忙/对话框：HRESULT 0x8001010A 或消息含 busy/忙/对话框/dialog）、`BAD_ARGS`、`COM_ERROR`、`INTERNAL`。
（**不产出** `ENCRYPTED`——密文对 WPS 是透明的，别声称能检测。）

## 3. `revise` 的 items 格式

```json
{"items":[
  {"para":12,"expect":"3.2 工程地质","op":"replace","find":"孔口高程","replace":"孔口地面高程"},
  {"para":15,"expect":"根据","op":"set_text","text":"整段替换后的新文本"},
  {"para":20,"expect":"表格前","op":"insert_after","text":"新增一段"},
  {"para":31,"expect":"废弃","op":"delete"}
]}
```
- 全批在 `TrackRevisions = true` 下执行 → 产生**可见修订**，用户可在 WPS 里逐条接受/拒绝。
- `expect` 必填：校验该段文本是否以其开头，不符即 `ANCHOR_MISSING`。
- **两轮执行**（v0.1 实测后的实现）：Pass 1 只校验（段落存在 / `expect` / `find`），任一条不过则**一条都不改**；Pass 2 按 `para` 降序写（插入、删除不影响后面更小的编号）；万一 Pass 2 中途出错，把本批新增的修订逐条 `Reject()` 回滚，错误信息里写 `rolled back N revision(s)`。
- 每条结果带 `textAfter`（改完后的开头 90 字），调用方可立即自查，不必再 `read`。
- 一批结束后报告：`applied`、`revisionsBefore/After/Added`、每条的 `para/op/hits/textAfter`。
- **段落编号**：表格每格算一段，且每行末尾**再多算一段**（3×4 表 = 15 段）；`read` 给表内段落打 `inTable: true`。

## 4. 硬约束（写进 SKILL.md 的护栏）

1. **不碰用户没指定的文件**；操作前先 `status`，报告在哪个实例/哪份文档上动手。
2. **文件被占用 → 只读**：`open` 必须如实返回 `readOnly`，只读时拒绝 `revise`（`READONLY`）。
3. **不替用户关闭文档**；`close` 只关闭本次明确要关的那份。
4. **进程**：结束时汇报 `wps.exe` 数量；只清自己起的自动化实例。
5. **交付提醒**：本机装有透明加密时，保存后文件是密文（头 `17DA5FA01633CD9A`），发给同事前需解密。
6. 源码与脚本**一律纯 ASCII**（中文只出现在 JSON 字符串里），避免 PS 5.1 按 ANSI 解析。

## 5. 验收标准（acceptance）

一份合成的 300 段中文报告：
1. `doctor` 能报出 WPS 版本与"渲染可用"。
2. `read` 段落数与真实一致，样式/自动编号正确。
3. `revise` 10 条（含 5 条 `replace`、3 条 `set_text`、2 条 `insert_after`）→ WPS 里能看到修订；`Revisions.Count` 与预期量级一致（replace≈2/条）。
4. `comment` 2 条 → `Comments.Count == 2`。
5. `render` 3 页 → PNG 非空白（非白像素占比 > 0.5%）。
6. `save` 后文件头为密文（本机特性）且 WPS 能再打开。
7. `close` 后 `wps.exe` 回到调用前基线。

## 2.1 排版命令（P1/P2/P3，v0.2 已实现）

| 命令 | 源文件 | 说明 |
|---|---|---|
| `format` | `Layout.cs` | 整篇/一段/区间/某样式段落的字符+段落格式 |
| `style` | `Layout.cs` | 改样式本身（含 `-NoNumbering` 去自动编号） |
| `page` | `Layout.cs` | 纸张/方向/页边距/页眉页脚/页码（逐节） |
| `table` | `Tables.cs` | 建表（数据/三线表/表头重复/列宽/题注）与改表（框线/字号/行高/不跨页） |
| `caption` | `Tables.cs` | SEQ 域题注「表 N / 图 N」，表上图上 |
| `picture` | `Tables.cs` | 插图（可带图题）/ 调宽高 / 按百分比裁剪 |
| `toc` | `Toc.cs` | 插入并刷新目录域（依赖标题样式） |
| `reference` | `Toc.cs` | 参考文献列表（悬挂缩进）/ 段末上标引用 |
| `view` | `Toc.cs` | 把窗口滚到某段/某页（不动光标） |

共用原语在 `WpsHelper.cs`：`ApplyCharFormat` / `ApplyParaFormat` / `SetLineSpacing` / `Points` / `Alignment`；
格式类命令默认**不进修订**，整批一个 `UndoRecord`（Ctrl+Z 一次撤）。

P4（未做，按需再议）：公式 LaTeX→OMML、图示（形状+连线）、fill_gaps 重排。
## 2.2 v0.3 增补（2026-10-23 晚）

| 命令 | 源文件 | 说明 |
|---|---|---|
| `fields` | `Text.cs` | 刷新全部域（页码/目录/题注/引用），中文报告交付前必跑 |
| `cell` | `Text.cs` | 改单个表格单元格（可按 table/row/col 或单元格内段号），可带格式 |
| `insert` | `Text.cs` | 批量插入段落；默认**干净写入**（不进修订），`-Track` 才留痕；带样式与格式 |

小改进：题注段 `KeepWithNext`（题不落单）；`format -NoSplitWords`（整词换行，治"37000 被拆成 37+000"）；
`picture` 返回 `naturalWidthPt` / `scale`，缩小到 50% 以下给 `warning`（图中文字会印得很小）。
## 2.3 v0.4 协作闭环（2026-10-23 深夜）

| 命令 | 源文件 | 说明 |
|---|---|---|
| `revisions` | `Review.cs` | 列出**未处理**修订（类型/作者/时间/所在段/文本），`-Action accept\|reject -Index N` 或 `-All` |
| `snapshot` | `Review.cs` | 段落文本快照存 `%LOCALAPPDATA%\wps-word-cooperate\snapshots\<hash>\<tag>.json` |
| `diff` | `Review.cs` | 与快照比：changed/added/removed/same；>4M 单元退化为位置比对（`degraded:true`） |
| `handoff` | `State.cs` | 文档旁生成 `<名>.handoff.md`：状态、待办修订/批注、`-Note`、最近 30 条操作记录 |

配套机制（`State.cs`）：写命令前后各取一次**廉价签名**（段数/修订数/保存标志/抽样文本哈希），与上次写入留下的状态比对，不一致就在结果里加 `changedOutside` + 提示；同时维护每文档的写入流水（供 `handoff`）。
关键语义：WPS 的 `Paragraph.Range.Text` 是**接受视图**，所以「已裁决」看 `revisions`，「文本变成什么样」看 `diff`。
## 2.4 v0.5 增补（2026-10-24）

| 命令 | 源文件 | 说明 |
|---|---|---|
| `search` | `Search.cs` | 全文找词 → 段落号/样式/inTable/上下文（`-Style` 过滤、`-Max` 截断、`-Case` 区分大小写）；定位优先于翻页 read |
| `selection` | `Search.cs` | 只读读用户当前选中/光标 → 覆盖段落号 + 表号/行列 + 嵌入图片号（"改我选中的这段/这个表/这张图"的锚点；选在别的窗口报 NOT_OPEN） |
| —（增强） | `WpsHelper.cs` | `comment -Action list`：批注清单（编号/作者/锚文/内容），`-Delete -Index n` 的锚 |
| —（增强） | `Layout.cs` | `page` 无选项=查询模式：报每节纸张/方向/边距/页眉页脚 |
| `undo` | `Search.cs` | Agent 自己的 Ctrl+Z（`-Count` 上限 20）；护栏：不分作者，仅自己刚写完且无 `changedOutside` 时用 |

配套：`snapshot -Out <路径>`（快照随文档走，跨机器接力）；`handoff` 结果与交接单带 `snapshotPath`；`open` 在"WPS 多进程但可接实例无此文档"时返回 `multiInstanceWarning`（GetActiveObject 只接第一个注册实例的已知限制）。
文档：SKILL.md 瘦身重排（教程下沉 `reference/workflows.md`）、版本头 v0.5、`reference/acceptance.md` 验收记录、`publish.ps1` 分发打包。

## 6. 分发（v0.5 起）

`publish.ps1`：把 SKILL.md + README.md + scripts（wps.ps1 与 helper/*.cs）+ reference/*.md 复制到目标（默认部署目录），**不带** tests/docs/tools/DESIGN.md/已编译 exe——首次调用按源码哈希现场编译。