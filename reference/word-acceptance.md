# 验收记录（v0.1 – v0.5）

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