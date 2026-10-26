# 验收记录（v0.1 – v1.8）

> 每个版本发布前在这里记一笔：跑了什么、用什么判据判定通过。回归时按本表逐条复跑。
> 语料与脚本在开发目录 `tests/`（不随分发包发布）：`report-300.docx`（306 段合成报告）+ `make_corpus.py` / `make_edits.py` / `run_acceptance.ps1` / `validate.py`。

## v0.1（2026-10-23）——核心闭环 9 命令

命令：doctor/status/open/read/revise/comment/render/save/close。

- 判据（tests/run_acceptance.ps1 可复跑）：
  1. `doctor` 报出 WPS 版本与"渲染可用"（EMF bytes > 1000）。
  2. `read` 段落数 306 与真实一致，样式/自动编号正确。
  3. `revise` 10 条（5 replace / 3 set_text / 2 insert_after）→ `Revisions.Count` 18–20（replace≈2/条），`AcceptAll` 视图逐条与预期一致。
  4. `comment` 2 条 → `Comments.Count == 2`，锚定正确。
  5. `render` 3 页 PNG 非空白（非白像素 2.3% / 7.0% / 6.6%）。
  6. `save` 后 WPS 能再打开（本机落盘为密文，见 encryption.md，属预期）。
  7. `close` 后 `wps.exe` 回到调用前基线（3→3）。
- 结果：全过。全流程 27s（开→改→批注→渲染→存→关）。

## v0.2（2026-10-23 晚）——排版 P1/P2/P3，+9 命令

命令：format/style/page（P1）、table/caption/picture（P2）、toc/reference/view（P3）。helper 拆 partial class。

- 集成验收（306 段/29 页报告）：format+style+page + table（三线表+表头跨页+题注）+ picture（带图题）+ toc（23 条、页码正确）+ reference（list + cite 上标）+ view，全程跑通，26 页，渲染目视通过。
- 新增实测坑回写 SKILL.md/workflows：WPS `Borders.Item()` 越界画对角框线；`Crop*` 单位是"原图 100% 尺寸的磅"；域 `Result.End` 插字被刷新抹掉；目录 `Add` 留空段；无标题样式时目录域报「错误！未找到目录项。」；`wdFieldRef=3`；EMF 渲染对 Bold 不可靠。

## v0.3（2026-10-23）——fields/cell/insert，+3 命令

- 白塔河 31 页真报告：`fields` 刷新目录+页码通过；`cell`/`insert` 冒烟通过。
- 小改进：题注 KeepWithNext、`format -NoSplitWords`、picture 缩小告警（<50% 给 warning）。

## v0.4（2026-10-23 深夜）——协作闭环，+4 命令

命令：revisions/snapshot/diff/handoff；配套 changedOutside 检测（写命令前后取廉价签名比对）。

- 验收：snapshot→revise→diff 报 changed 段与预期一致；revisions 列表=未处理待办；`-All` accept/reject 行为正确；handoff 生成 `.handoff.md`（状态/待办/操作记录）；人为改动后下一次写命令结果出现 `changedOutside`。

## v0.5（2026-10-24）——search/undo/快照可移植/多实例告警 + 文档重构

新能力：

1. **`search`**：全文找词 → 段落号/样式/inTable/上下文。
   - 冒烟（tests/corpus 报告 306 段）：`search -Text "断面" -Max 5` → 5 hits（para 26/28/36/38/39，context 均含命中词），`totalHits:13`、`truncated:true`；不存在词返回 0 hits 不报错。
2. **`undo`**：`format -Bold` 一批后 `undo -Count 1` → `undone:1`，撤成功（排版批=1 撤销记录）；上限 20；护栏文案随结果返回。
3. **`snapshot -Out`**：快照写到指定路径（跨机器/交接随文档走）；返回 note 说明 diff 须用 `-Since <路径>`（不参与 `last`）。
4. **`handoff` 带 `snapshotPath`**：交接单与结果里都给出最新快照路径（无快照时给"接手先 snapshot"提示）。
5. **`open` 多实例告警**：WPS 多进程但可接实例里没有该文档时，结果带 `multiInstanceWarning`（GetActiveObject 只能拿第一个注册实例的已知限制，先警告不硬来）。
6. **`selection`**：读用户当前选中/光标 → 段落号。
   - 冒烟（306 段语料）：光标在第 1 段 → `paras:[1]`、`style:标题`；外挂 COM 选中 26–28 段 → `paras:[26,27,28]`、选中文本 308 字逐字吻合、`fromPara/toPara` 正确；命令只读、不动选区。
   - 表格内选中（表 1 第 2 行第 2 列单元格）：`inTable:true, table:1, row:2, col:2`，单元格文字正确——"改这个表"场景直通 `table -Table n`/`cell -Table n`。
   - 选中含嵌入式图片的段落（插图于 31 段后）：`images:[{image:1,widthPt:226.5,...}]`——"改这张图"场景直通 `picture -Image n`。
7. **`comment -Action list`**：加 1 条批注后 list → `items:[{i:1,author,scope(锚文),text}]`，i 即 `-Delete -Index` 编号；只读。
8. **`page` 查询模式**：不带选项 → `sections:[{section:1,paperSizeId:7(A4),orientation,marginsCm{...},headerText,footerText}]`，只读不改。
9. **`table -Borders thick`**（外框 1.5pt + 内线 0.5pt，中文报告粗边框）：白塔河报告表 33 实测 `applied: borders=thick`，回读 top/left/bottom/right=[1,12]、insideH/V=[1,4] 正确（首次调用撞上重编译竞态报旧错误码，重跑即过——已知小毛病，缓存按源码哈希命中后不再出现）。
10. **Ctrl+Z 全撤实证（2026-10-24 白塔河报告实测）**：
    - WPS **支持** `app.UndoRecord.StartCustomRecord` 分组（探针：3 次属性设置包 1 记录 → 1 次 Undo 全撤）；`doctor` 现返回 `undoRecord:true/false`。
    - `format -Match` 新能力：按段落文本正则批量套格式（图名 128 段一次命中，返回 samples）。实测一次 `Undo()` 恰好撤掉该命令**实际改变**的全部 32 段（其余 96 段原本已是同值、无撤销记录产生——符合预期）。
    - **重大发现：WPS save 会清空/截断撤销栈**（探针改 128 段、保存后栈里只剩 119 步且只能撤 32 段）。结论：要反悔必须在保存前；文本类改动必须走修订而非靠撤销栈。已写进 SKILL.md 排障表。
    - 反面教材：Agent 绕过 skill 手写 COM 逐段循环改格式 = 每段独立撤销记录（且同样被 save 截断），用户一次 Ctrl+Z 撤不掉——本例即由此触发。修复=`format -Match` 进 skill。

## v0.7（2026-10-24）：大文档定位提速

| 命令 | 之前（14860 段/119 表/271 页，带 31 处修订） | 之后 |
|---|---|---|
| `search` 单词 | ~150 s（逐段 COM 往返 ×2×14860） | **3.9 s**（含 ~2 s helper 启动），38× |
| `search` 4 词 | 4×150 s ≈ 10 min | **4.9 s** 一次查完（每词分项计数） |
| `selection` | 线性扫（光标靠后时分钟级） | 偏移换算 O(1)（~2 s） |

架构（两条路线，各自带正确性闸门）：
1. **`search` 走 WPS 原生 `Find`**：扫描在 WPS 进程内完成，每命中 1~2 次往返（`Range(0,off).Paragraphs.Count` 偏移→段落号，返回 c/c+1 两候选、用"段落文本含命中词"裁决）。**带修订也能用**。
2. **`format -Match` / `DocTexts`（snapshot/diff）走 `Content.Text` 一次拿回 + 本地 `\r` 分块**：仅当**无修订**时启用——实测发现 **Range 坐标把已删除修订文本计入长度而 `Range.Text`/`Content.Text` 隐藏它**（例：段 53 Range 长度 354、可见文本 328），一旦存在删除类修订，本地偏移必然漂移；闸门=`Revisions.Count==0` 且切分段数==段落总数，否则静默退回逐段慢路径。
- 回归：306 段语料 `search 断面` 13 hits/首段 26 与 v0.5 基准一致；白塔河慢/快路径段落号一致（para 59/64/65）。
- 发现记录：`doc.Content.Text` 单次拿全 16.7 万字符仅 556 ms；表格使 `\r` 切分段数+544（119 表的行末标记），故带表格文档的 snapshot/diff 仍走慢路径（正确优先）。
- 文档重构：SKILL.md 瘦身（长教程下沉 reference/workflows.md）、小节重排、去机器特定表述（`pwsh` 说法、示例安装路径）、新增版本头；encryption.md 补"无解密技能机器"的手工路径；新增本验收记录。
- 分发：publish.ps1 只打包 SKILL.md + scripts + reference + README（tests/docs/tools/DESIGN.md 不进包）。

### 回归清单（每次改 helper 后跑）

1. `doctor`（编译 + 自检）；
2. 语料副本上：`open` → `search`（已知词，校验 hits）→ `revise`（2 条 replace，校验 revisionsAdded=4 与 textAfter）→ `comment`（1 条）→ `render` 1 页非白 → `undo -Count 2`（校验回退后 revisions 回 0）→ `snapshot -Out` → `handoff` → `close`，进程数回到基线。

## v0.8（2026-10-25）：finish 完成信号

| # | 用例 | 结果 |
|---|---|---|
| F1 | inish -Note | ✓ 托盘气泡弹出 + 提示音（NotifyIcon/SystemSounds 纯 .NET）；**不触碰文档**（与 Excel 版语义对齐，默认零写入）；wps.ps1 需补 System.Windows.Forms 引用与 [STAThread] |



## v0.8.1（2026-10-25）：文档被用户关闭的可感知性

问题：用户在 WPS 里关掉文档后，下一条命令会**静默重开**继续干活——agent 无感，且"未保存关闭"会丢改动。

- `RequireDoc` 重开路径置 `_reopened` 标志，Main 把 `reopened: true` + `reopenedHint` 写进**result 层**（与 changedOutside 同层；首版误写到顶层是 bug）；重开改为**不弹窗**（`OpenPath(show:false)`），不再吓用户。
- `unsavedAtClose: true`：重开时交叉核对本地账本（State），上次记录为"未保存"才报——**真丢失预警**，提示先 `diff` 对账再动手。
- `close` 命令现在自己在关闭前后记账（关闭前取签名、关闭后按 `-Save` 实际值写 `saved`）——修掉 AfterWrite 在文档已关时无法记账导致的**陈旧"未保存"误报**（实测场景 A 曾误报）。
- 实测：A=带 `-Save` 关闭→重开仅 `reopened`；B=加批注后无 `-Save` 关闭→`reopened`+`unsavedAtClose`，且批注确已丢失（comments=0），预警真实有效。
- 调试副产品：验证 .NET exe 内字符串是 UTF-16 存储（ASCII 子串搜索判"是否编译进"会误判）；重编译竞态再复现一次（首跑旧逻辑、重跑即好）。
## v1.8（2026-10-26）：表格编号检查（captions）+ search 段号校验 + inTable 边界

起因：一份真实报告的复盘（75 页 / 10521 段 / 32 表 / 287 处未处理修订）：一个"检查表格编号有没有问题"的小任务，
**读取 6 分钟、改写 2 秒**。根因不是文档大，而是**命令集缺"清单"能力**——检查被迫退化成 O(全文) 扫描（22 次
`read`、4.5 min），而唯一看起来便宜的 `search` 又给了**静默错段**，据此写下过错误结论、被迫返工。复盘文档：
`WPS表格编号检查-复盘与issue草案-20261026.md`（桌面）。本轮把复盘里的 A/B/C/D 四条全部落地。

### 实测（同一份 10521 段报告，WPS 12.1.0.29161）

| 动作 | 之前 | 之后 |
|---|---|---|
| 题注↔表格配对 + 编号体检 | 22 次 `read` 全扫，**4.5 min**，只得到配对 | **`captions` 一次调用 6.4 s**，另附重号/跳号/格式/空格/位置/悬空引用 |
| `search -Text 表` | 8 s，113 命中但**漏 4 条真实题注**、同段重复报 3 次 | 6.9 s，113 命中/报出 105 条（其余为每段 3 条上限，列在 `cappedParas`），**4 条题注全部命中**、每命中各有自己的 `at` |
| 表头信息 | `table -Action scan` 无段号，只能盲试段号 | `scan` 每行带 `para`，直接定点读题注（本报告 7.3 s / 32 表） |
| 正文全扫 | 22 次调用约 3.5 min | `read -SkipTableCells -Count 0` **一次调用** 111 s，245 段正文（10276 段单元格被跳过） |
| `audit` 题注判据 | 宽松正则：正文"如表3.3-2所示"也算题注；题名在表内的表同时报 `caption-missing` + `table-caption-in-row` | 17 表 `table-caption-in-row`、**0** `caption-missing`（首版收紧过头，会把 9 张图名误报为缺题注——真实排布是"图 / 空段 / 一个 / 残段 / 图2.1-1 xxx"，已改为向后看 1–4 段并遇标题停止） |

`captions` 在本报告上的结论与人工复盘**逐项吻合**：32 表 / 32 题注；`positions` = 17 张题名做在表内第 1 行、
15 张在表外上方；`duplicates` = `3.3-3` 在段 10307 与 10430（正是要找的重号）；`formats` = dot 27 / hyphen 5；
`references` 43 处全部可解析。

### 修复证据（对应复盘 §3/§4/§6）

- **§3 search 静默错段（高）**：`ParaOfHit` 从"候选段文本含该词"改为**字符区间校验**（`Range.Start ≤ offset < Range.End`；
  段区间半开且铺满全文，只有一个候选能通过），失败才 `return null` 退回逐段慢扫（错段≠0 时旧代码不会兜底）；
  命中段内的 `at` 用 `foundStart - paraStart` 各自计算（`Range.Text` 遮掉删除修订时会漂，故再按段内递推纠正并计入 `atFallbacks`）；
  每段 3 条上限显式报 `cappedParas`。
- **§4 inTable 误判（中）**：`read`/`search`/`cell` 的表区间判据由 `<=` 改为 `<`（`Table.Range.End` 是排他端），
  表内段额外带 `table: n`。
- **§5 缺口 C（结构性）**：`table -Action scan` 加 `para`；新增 `captions`；`read` 加 `-SkipTableCells`/`-Only`。
- **§6 缺陷 D（低）**：`audit` 用与 `captions` 共用的题注解析器（标签必须在段首，正文"如表3.3-2所示"不算题注），
  表内题名只报 `table-caption-in-row`；图片题注允许在图片后 1–4 段并在遇标题样式时停止。

### 回归语料与断言（`tests/word/captions/`，不随分发包）

- `make_fixture.py` → `captions-fixture.docx`（20 表 / 19 题注 / 4479 段，其中正文级 180 段；8 张 35×13 大表；埋重号
  `3.3-3`、跳号 `2.4.1→2.4.3`、点分/短横混用、空格混用、`表2.9.9` 悬空引用、表 19 无题注、表 13 后紧跟一段含
  "表3.3-2" 的正文）。幂等（两次生成 SHA256 相同）。
- `check_captions.py`：**30 条断言全过**（C1 清单/配对、C2 六类缺陷且不多报、B1 表后正文 `inTable` 为假、
  C3 `-SkipTableCells`/`-Only` 与全扫结果一致、C4 `scan.para` 真属于该表、A1 每个命中 `at` 指向真实命中位置、
  A2 题注全部命中、audit 只报 1 张"表内也无题注"的表与 6 张"题名在框内"的表）。
- 旧语料 `report-300.docx` 全链回归：`search 断面` 13 命中（与 v0.5 基线一致，hits 段号 26/28/36/38/39/42/43/61/65/101/106/111/151）、
  `revise` 10 条（revisionsAdded 18）、`comment` 1 条、`render` 1 页 96 KB、`undo` 1 步、`snapshot`→`diff` 8 段、
  `handoff` 落盘、`close` 进程回落到基线——全过。

### 已知边界

- `audit` 仍是**逐段**扫描（本报告 89 s）：它要查所有段落（含单元格）的用字用单位。只要"编号体检"就跑 `captions`，
  别用 `audit` 凑。
- `read -SkipTableCells` 的提速来自"少调几次"，单段仍是一次 COM 往返（实测 10521 段 111 s）；要数量级就去用 `captions`。
- `captions` 只认**文字题注**：域式（SEQ）题注、题注在第 2 行、题号前加了"续表"等形态会漏，返回里给 `tables[].para`
  供人工核对。
- 复盘 §7（流程失误：先猜后证、有锚点不用、范围默许放大）写进 SKILL.md 的工作流 10：**编号检查先 `captions`**。