# wps-office-cooperate 项目分析与优化方案

> 生成日期：2026-11-04 · 基线：`master@f4ae0f6`（v1.33）· 状态：**分析定案，未实施任何优化**
>
> 方法：4 路并行深度审查（PowerShell 脚本层 / C# helper 层 / 文档层 / 测试与仓库卫生），关键结论均经二次人工核实（含两处子代理误报的修正）。文中 `文件:行号` 以 v1.33 工作区为准。

---

## 一、项目概况

**是什么**：一个 agent skill——通过 COM 在用户正开着的 WPS 窗口里实时修改 Word/Excel/PPT 文档，改动以修订/批注呈现。skill 的"运行说明书"即文档本身（agent 读 SKILL.md 学会调用）。

**架构**（四层）：

| 层 | 内容 | 规模 |
|---|---|---|
| 入口层 | `office.ps1`（遗留）/ `woc.ps1`+`woc.cmd`（快速入口） | 110 行 |
| 包装层 | `word/excel/ppt.ps1`：param 声明 → 组装 request JSON → 常驻 daemon 命名管道（回落 one-shot） | 3×800-880 行 |
| helper 层 | C#5 单文件 partial class，首跑 csc.exe 编译按源码哈希缓存；`shared/*.cs` 以**源码链接**方式编进三个 exe | 42 文件 / 约 2 万行（最大 `WppHelper.cs` 3606、`EtHelper.cs` 3150、`WpsHelper.cs` 2228） |
| 编排层 | `tasks*.ps1` 多任务离线规划（split/merge/plan，零 COM） | 4 文件 / 约 1580 行 |

文档：SKILL.md 148.7KB（89,521 字符）+ reference/ 11 篇 228.8KB + docs/borrow 14 篇 413KB + DESIGN×3。测试：验收脚本模式（check_*.py 断言 + run_*.ps1 驱动），约 8,000 行。

**总体判断**：这是一个对 WPS COM 怪癖做了大量实测防御的成熟代码库——注释里带测量数据和日期、quirks 台账、错误码体系、验收文化都属于上乘。主要风险不在"乱"，集中在四点：**①三端复制粘贴已发生漂移；②daemon 路径无超时；③SKILL.md 每次加载烧 3 万 token 且版本信息四处双写已现漂移实锤；④测试体系严重偏科（Excel 近乎裸奔）且无统一 runner、无 CI。**

---

## 二、优势（优化时应保留的资产）

1. **JSON 输出契约贯彻好**：所有自身错误走 `Write-Error-Json`，警告走 stderr，stdout 干净；退出码分层（1=调用错、2=ok:false、3=check 门禁）。
2. **性能决策有实测支撑**：`-Mta` 2.2s STA 税（woc.cmd:6-11）、哈希缓存 50ms 税（woc-common.ps1:5-8）、batch 220ms/命令 vs 管道 1-15ms 往返，均注明数据与日期。
3. **`param-check.ps1` 防回归设计**：专治"参数绑定了但没进 request 静默丢失"（excel.ps1:347-349 记录了真实事故）。
4. **shared/ 抽象边界克制正确**：PipeServer 只做传输、MessageFilter/WocLocks/ClipboardKeeper 不依赖 app 类型、TrayDaemon 靠回调脚本还原标题。
5. **协作安全设计成熟**：DOC_CLOSED 防静默重开（excel/ppt）、多实例 fork 警告、badge 被用户 Ctrl+S 覆盖后的自动修复（WppHelper.cs:3226-3330）、changedOutside 检测（word/State.cs:137-168）。
6. **测试 oracle 独立实现**：PPT 断言用 python-pptx 独立算期望（check.py:5-7）；Word 用 zipfile 直解 docx XML 做"静态地面真值"（check_normalize.py:115-163）；语料固定 seed 可复现。
7. **文档有量化自省传统**：v1.29 拆历史出主文档时记录了"1.97 万字符占 25.5%"的依据；SKILL.md:226-238 "上下文纪律"一节自己监控文档开销。

---

## 三、问题清单（按主题，附证据与严重度）

### T1 三端复制粘贴 + 已发生漂移（最大结构性问题）

**PS 包装层**：三个 app 脚本各约 200-210 行（~25%）是逐字相同的复制，合计约 400 行冗余：

| 重复块 | word.ps1 | excel.ps1 | ppt.ps1 | 状态 |
|---|---|---|---|---|
| `Write-Error-Json`/`Resolve-FullPath` | 153-161 | 145-153 | 160-168 | 逐字相同 |
| `Invoke-Daemon`（管道客户端 ~35 行） | 183-216 | 174-207 | 188-221 | 逐字相同，与 woc-common.ps1:58-81 是**同一协议的第 4 份实现** |
| `Start-Daemon`（ping→启动→轮询） | 218-232 | 209-223 | 223-237 | 逐字相同（woc-common 里还有第 4 个变体，循环次数 40→50 已漂移） |
| csc 编译块 | 256-271 | 247-262 | 261-276 | **refs 已分叉：ppt.ps1 缺 `-r:Microsoft.CSharp.dll`**（word:265/excel:256 均有，实测 grep 确认） |
| `batch` 命令块 | 281-297 | 272-288 | 286-302 | word↔excel diff 为 0 |
| `open`/`hint`/`serve` case | — | — | — | word↔excel 逐字相同 |
| 一次性执行尾段 | 780-805 | 769-792 | 851-877 | 仅 app 名不同 |

word.ps1 内部还有自我重复：`Read-Items-Json` 的 5 行逻辑内联 6 遍（367/377/420/522/599/632，excel/ppt 已提成函数）；`unit` 与 `units` 两个 case 函数体完全相同（685-698）。

**C# helper 层**：估算 **1200+ 行 × 3 份**的 plumbing 复制：
- `Fail`/`Json`/`Bag` 三份逐字（WpsHelper.cs:25-284 / EtHelper.cs:26-285 / WppHelper.cs:33-305，diff 为空）；
- COM attach 六件套三份平行（WpsHelper.cs:706-765 / EtHelper.cs:459-517 / WppHelper.cs:509-599）；
- hint 全家桶 word↔excel 两份近乎逐字（WpsHelper.cs:1956-2164 vs EtHelper.cs:2904-3038）；`ShowBalloon/PlayAsterisk` 三份逐字；
- attachment marker 五件套 excel↔ppt 两份（EtHelper.cs:543-598 vs WppHelper.cs:629-686）。

**漂移实锤（复制后单端修改未回同步）**：
- **busy 判定**：word 检查 7 种信号（含中文"忙"/"对话框"，WpsHelper.cs:415-421），excel/ppt 只有 5 种（EtHelper.cs:354-358、WppHelper.cs:406-410）→ 中文环境下 ET/WPP 报忙会被归类 COM_ERROR 而非 BUSY，上层无法重试。**【高】**
- **忙重试策略不对称**：`ComRetry`（5 次×400-1600ms）只在 ppt（WppHelper.cs:524-540），word/excel 完全依赖 MessageFilter 单层。
- PS 层 refs 分叉（见上表）。

**tasks*.ps1 层**：`Read-Json`/`Write-Json`/`FullPath2` 四份（tasks-excel:44-56 与 tasks-ppt:44-55 逐字相同）；task 文件发现块三份逐字；priority 解析四份；selftest 脚手架四份；`New-BatchOp`（tasks.ps1:402）与 `New-BatchEntry`（tasks-excel:296）是同名异写双胞胎。tasks-excel↔tasks-ppt 公共骨架约占各自 30-35%。`Invoke-Merge` 三个均超 110 行（word 版 184 行，tasks.ps1:214-398）。

### T2 可靠性缺口：超时与错误判定

1. **【高】daemon 路径无读超时，双向挂死**：one-shot 有 `WaitForExit($TimeoutSec=180)`（word.ps1:789 等），但 `Invoke-Daemon` 只对 Connect 限时（word.ps1:188），连接后的 `$pipe.Read`（196-209）无限阻塞；`PipeServer.Run` 的 handler 也无看门狗（PipeServer.cs:86）。WPS 弹模态对话框（真实会发生）时 COM 调用阻塞不返回，MessageFilter 只救"拒绝型"忙救不了"挂起"，daemon 主线程与 wrapper 双双永久挂死。`Invoke-WocBatch` 同样无读超时。
2. **【中】`Invoke-WocBatch` 顶层恒 `ok=true`**（woc-common.ps1:134-138）：即使所有 op 都 NO_DAEMON 失败，信封仍 ok:true + exit 0（word.ps1:295-296），只查顶层的调用方会漏判整批失败。
3. **【中】exit 2 靠正则嗅探原始 JSON**：`$text -match '"ok"\s*:\s*false'`（word.ps1:775,801 等 6 处）——编辑一份讲 JSON 的讲义（内容含字面 `"ok": false`）时，ok:true 响应被误判 exit 2。应先 `ConvertFrom-Json` 再判 `.ok`。
4. **【高】PS 参数绑定失败输出非 JSON**：打错命令名（ValidateSet 拒绝）或传未知参数时，PS 宿主向 stderr 打印纯文本、stdout 无 JSON、exit 1——这是 agent 最容易踩的非 JSON 路径，SKILL.md 未提示该形态。
5. **【中】tasks split 的进程回调无超时**：tasks-excel.ps1:100-124 / tasks-ppt.ps1:103-104 `Start-Process powershell` 后 `WaitForExit()` 不限时，WPS 弹框即永久挂起；且逐 chunk 起完整 PowerShell（每 60 行一次），与全仓库省进程启动的原则相悖。
6. **【低】csc 编译块 `& $csc ... 2>&1` + `$ErrorActionPreference='Stop'`**（word.ps1:266 等）：PS 5.1 下 csc 一旦写 stderr 即以非 JSON 形态崩溃（概率低但存在）。
7. **【低】static 残留**：daemon 长寿命进程下 `_chunkStats`（FastText.cs:484）、`_xmlNote`（XmlText.cs:31）跨请求残留，靠"巧合式安全"（下游按本次路径判断）而非显式重置。
8. **【低】COM RCW 从不 `Marshal.ReleaseComObject`**：one-shot 无所谓，daemon 长会话大文档遍历内存单调上升，只能等 GC。

### T3 三端语义不一致（agent 需按 app 记忆的行为分叉）

| 语义点 | word | excel | ppt | 影响 |
|---|---|---|---|---|
| 文档被用户关闭后 | **静默重开** + `reopened:true`（WpsHelper.cs:767-781），无 attachment marker | 拒绝，报 `DOC_CLOSED`（EtHelper.cs:525-533） | 拒绝，`open -reopen` 例外 | 【高】word 可能"从磁盘复活"用户已主动丢弃未保存修改的文档 |
| `QuitAll` | 遍历**用户共享实例**全部文档 `Close(0)` 不保存再 Quit（WpsHelper.cs:2208-2221）——**会连带丢弃用户自己的无关文档** | 仅 `app.Quit()`，但未设 DisplayAlerts=false，有未保存时可能被保存对话框阻塞 | 同 excel | 【高】两个方向各有数据丢失/挂死风险 |
| search 大小写键 | `caseSensitive`（word.ps1:354） | `case`（excel.ps1:323） | `case`（ppt.ps1:360） | 【低】跨 app agent 易写混 |
| `finish` 是否必须 `-Doc` | 否 | 是（excel.ps1:692） | 否 | 【低】 |
| 快照/回滚 | 段落文本快照+diff，**无 restore**（靠修订层） | 值+公式快照 + **唯一带逐格回滚的 restore** | shape 文本快照 + 文本级 restore | 【中】三套格式三种哲学，无共享 |

### T4 文档层：上下文成本与版本多源（skill 的"性能问题"）

1. **【高】SKILL.md 单次加载 ~29K-36K token**（89,521 字符 = 28,596 CJK + 59,509 ASCII）。这与 SKILL.md:228 自己写下的教训（"文档涨 1 万 token，后面每一步都变慢"、>300K prompt 时 TTFT 均值 60s）直接抵触——主文档本身就是它描述的瓶颈的固定组成部分。结构占比：三端速查表 43%（其中最长行 SKILL.md:123 达 1,936 字符、:537 达 1,536 字符，塞进了多版结论/实测数据，成了 changelog 第二现场）＞PPT 章 32%＞Word 章 28%＞Excel 章 23%。
2. **【高】版本头条已漂移一版（实锤）**：SKILL.md:9 头条仍是 "版本 v1.32（2026-11-03）"，而 HEAD 与 changelog.md:5 均为 v1.33（已二次核实）。"唯一权威说明"的版本声明指向旧版；v1.33 内容只散在正文 :123、:206、:226、:272。
3. **【高】v1.29-v1.32 四个版本在 SKILL.md 与 changelog.md 双写**（SKILL.md:9-15 共 4,981 字符 vs changelog.md:15-48），每次发版两处同步，问题 2 即该流程漏更的实证。
4. **【中】版本号无权威单源**："vN.NN" 在 SKILL.md 出现 146 次、changelog 42 次、quirks 三件 46 次、acceptance 15 次；frontmatter 无 version 字段。同一版本历史存在 4 份活记录（SKILL.md 头条 / changelog / word+excel-acceptance / git log），docs/borrow/skill-*-v13x.md 又是 v1.30-1.32 的第 5 份全量报告。
5. **【中】docs/borrow/skill-*.md 是已消费完的补丁片段**（自述"按位置贴进 SKILL.md 即可"，早已贴完），w00-w04 是上游 v1.2.0 研究快照（31 万字符，随上游演进必然失效）——不烧上下文但占 413KB 且三层重复记述。
6. **【低】SKILL.md:836-840 参考列表编辑事故**：multi-task 三条链接裸列，说明文字全部粘连到 multi-task-session.md 条目；参考节漏列 changelog.md；:17 "v1.1–v1.28" 范围说明已过时（changelog 现含至 v1.33）。
7. **【低】README 滞后**：hint 描述仍停在 v1.11/v1.12 语境（"PPT 侧只给气泡"——v1.26 起 PPT 已有标题标记）；能力清单未提 v1.29+ 的 formula/footnote/数据操作层/成品版式层。
8. **【低】tools/build_guide.py 已死**：硬编码旧名 `wps-word-cooperate` 路径（:13）不可跑，HTML 内容写死落后 33 个版本，产物已被 ignore——第五份手写说明层。
9. **正面结论（抽查）**：文档-代码一致性抽查未发现命令/参数漂移（ValidateSet、`render -Table`、`pivot -Group`、`slide -Action dup` 均一致）；仅两处表述歧义非错误。

### T5 测试体系：偏科 + 无 runner + 无 CI

1. **【高】Excel 端近乎裸奔**：`tests/excel/` 仅 check_hint.py(85 行)+check_tray.py(113 行)+multi-task 一个 ps1。核心数据功能 **fill / chart / sort / deletewhere / fillblanks / table / sheet / replace / rows / cols 全部零回归脚本**，v0.3/v0.4 验收只存在于 excel-acceptance.md:26-46 的表格里（一次性跑过，无脚本沉淀）。对照 PPT 的 6 大驱动器 12 场景、Word 的 captions 30 断言，极度失衡。
2. **【高】无统一 runner**：跑全部需手动串 10+ 命令（README:63-71 逐条列举），无 run-all、无 smoke/full 分层标记，新人无从下手。
3. **【高】check_normalize.py 默认 fixture 指向仓库外本机路径**（:43 `D:\DSH工作区\乐平表格整理\orig-copy.docx` 以 unicode 转义写死，:186 无该目录直接退出）——克隆后第一条 Word 测试即挂。
4. **【高】完全没有 CI**：无 .github/、无 yml、无 pre-commit。而仓库里明明存在"无 WPS 也能跑"的 fast 层却没串起来：make_corpus/make_fixture/make_deck（纯 python-docx/pptx）、check_normalize `--static-only`（纯 zipfile）、check.py 期望计算、tasks*.ps1 `selftest` 子命令、param-check.ps1——这些在 Linux+纯 Python 即可断言。
5. **【中】Word 主链半自动 + 引用已失效文件名**：run_acceptance.ps1 只打印手动步骤（:58-63），且引用 v1.0 后已改名的 `wps.ps1`（:8,59-62），照抄必失败；formula/footnote/unit/toc/picture/insert/fillgaps 等命令无回归。
6. **【中】共享层 ClipboardKeeper/WocLocks/MessageFilter 无自动断言**（SKILL.md:9 自述忙重试"机制已装、未实弹"）。
7. **【中】Python 依赖未声明**：pywin32/python-docx/python-pptx/openpyxl/Pillow 无 requirements 文件，无环境说明。
8. **【低】时序脆弱**：全仓库 25 处固定 sleep（check_daemon.py:120/126/166、validate.py:464 等），慢机器 flaky；D3 "daemon 更快"断言依赖本机时序。

### T6 仓库卫生与工具链

1. **【高】publish.ps1 递归删除只护栏 `length >= 10`**（publish.ps1:50,65-71）：`-Target C:\Users\Administrator\` 这类误传（10+ 字符）会**静默递归删除用户 profile 下无关文件**。无目标必须是 skills 目录的校验、无 -WhatIf、无二次确认；逐文件 Copy 无事务性，中途失败留半新半旧部署。
2. **【高】中文进入了自称 Pure-ASCII 的 .ps1 且全部无 BOM**（实测仅 woc.ps1 有 BOM）：word.ps1:481（`仿宋`）、:497（`正文`）——会进 BAD_ARGS 的 JSON message；tasks.ps1:471（throw 消息里的 `标题 1`）、ppt.ps1:452/462（注释）。PS 5.1 按 ANSI 读无 BOM 文件（README:70 自己写明了这条规则）→ 中文系统上 GBK 误解 UTF-8 → JSON 乱码。**规则自定自违。**
3. **【中】被跟踪的运行产物**：`tests/word/validate-report.json`（每次 validate 覆写）与 `tests/word/corpus/_selftest-revised.docx`（selftest 输出）自 v1.0 入库（已核实 git ls-files）；`tests/excel/work/` 8 个历史验收遗留文件（部分输入路径已失效，build_block.py:9 硬编码旧仓库名 `wps-excel-cooperate`）。
4. **【中】office.ps1 半死代码**（:28-31）：`if ($app -eq 'ppt') { Write-Error-Json 'APP_NOT_SUPPORTED' "PPT ... not implemented yet. Available: word, excel." }` 在上一行已放行 ppt 后永不可达，文案还是 ppt 实现前的；分发器本身已被 woc.ps1/woc.cmd 架空，但 SKILL.md:29 与 README:30 仍教 `office.ps1`。
5. **【低】publish.ps1:10 硬编码本机默认目标** `C:\Users\Administrator\.agents\skills\...`。
6. **【低】死代码/残留**：word.ps1:757 空 `$paramCommandSkips`、:763 `Slide=0` 默认值（word 无 -Slide 参数，从 ppt 模板抄来的残留）；tasks.ps1:537 死变量 `$round`、:509 `$out.Count` 应为 `$rangeOut.Count`；excel.ps1:701 两个 case 挤一行；`scripts/helper/word/build/WpsHelper.exe` 磁盘残留（未入库，已核实，可清理）；.tmp/ 438MB + tmp_axp/ 8.4MB 本地磁盘垃圾。
7. **【低】.gitignore `tests/work/` 规则指向不存在目录**（旧 layout 遗留），实际 `tests/excel/work/` 却被跟踪——错位易误导。

---

## 四、优化方案（分期，未实施）

### P0 — 快速修复（每项 ≤ 半天，低风险高价值，先止血）

| # | 事项 | 做法 | 解决的问题 |
|---|---|---|---|
| P0-1 | **版本单源化** | SKILL.md 版本区改为"当前版要点 3-5 行 + 其余见 changelog"，删 v1.29-1.32 四段双写；frontmatter 加 `version:` 字段；建发版 checklist（frontmatter/头条/changelog 首条/README/commit 五处同步） | T4-1/2/3：立省 ~2K token/次加载，消灭已发生的头条漂移 |
| P0-2 | **编码整改** | 全部 .ps1 转存 UTF-8 with BOM；word.ps1:481/497、tasks.ps1:471 的中文字面量改 ASCII（如 `FangSong (GB2312)`）；加一条检查脚本/钩子 `grep -P '[^\x00-\x7F]'`（BOM 除外） | T6-2：规则自定自违的 JSON 乱码隐患 |
| P0-3 | **publish.ps1 安全护栏** | 目标路径必须匹配 `*\.agents\skills\*`（或显式 `-Force`）才允许递归清理；先复制到 `$Target.new` 再原子切换；删除后清空目录 | T6-1：误传目标静默删用户文件的破坏性风险 |
| P0-4 | **exit 2 判定改正** | 6 处正则嗅探改为 `ConvertFrom-Json` 后判 `$obj.ok`（回退才用正则） | T2-3：文档内容含 `"ok": false` 时误判 |
| P0-5 | **batch 整批失败可见** | `Invoke-WocBatch` 所有 op 失败时顶层 ok:false（或 wrapper exit≠0） | T2-2 |
| P0-6 | **参数绑定错误皂化** | woc.ps1 改为先以 `-Command` 捕获绑定异常包成 JSON；SKILL.md 铁律加一条"stderr 含 ParameterBinding/ValidateSet 且无 stdout → 按调用错误处理" | T2-4：agent 最易踩的非 JSON 路径 |
| P0-7 | **小清理一批** | 删 office.ps1:28-31 死分支与过期文案、word.ps1 空 skips/`Slide` 残留、tasks.ps1:537 死变量、修 :509 `$out.Count`；`unit`/`units` 合并 case；word.ps1 6 处内联 `Read-Items-Json` 提函数；`git rm --cached` validate-report.json 与 _selftest-revised.docx 并补 ignore；统一 `case`/`caseSensitive` 与 `finish -Doc` 契约（SKILL.md 同步标注） | T1/T3/T6 零散项 |
| P0-8 | **check_normalize fixture 自给** | 把乐平语料生成器按 make_*.py 模式入库（或 DEFAULT_FIXTURE 改必填参数 + tests/README 说明） | T5-3：克隆即挂 |
| P0-9 | **修 run_acceptance 的 wps.ps1 引用**；requirements-dev.txt + tests/README.md（依赖与 WPS 版本要求） | 照抄必失败 + 依赖声明缺失 | T5-5/7 |

### P1 — 结构性收敛（1-2 天/项，中风险，需回归测试护航）

| # | 事项 | 做法 | 收益 |
|---|---|---|---|
| P1-1 | **PS 包装层去重** | 建 `woc-app-common.ps1`：把 `Invoke-Daemon`/`Start-Daemon`/csc 编译块（refs 参数化，顺手显式化 ppt 的缺 refs 是否为有意）/`batch`/`serve`/`open`/`hint`/`save`/`close` case/`Write-Error-Json`/`Resolve-FullPath`/`Read-Items-Json` 收拢为传 app 名的参数化函数 | 三脚本各瘦 ~200 行（合计 -400 行），管道协议从 4 份实现收敛为 1 份，refs 分叉消除 |
| P1-2 | **C# shared 层扩容**（本方案最大单项收益） | 抽 `Common.cs` 进 shared/：`Fail`/`Json`/`Bag`、COM attach 六件套（参数化 ProgID）、`Ok`/`Err`/`Ping`、`ShowBalloon`/`PlayAsterisk`、marker 五件套（参数化文件名）、`OpenPath` 回退、`ComRetry` 下沉；busy 判定收敛为单一 `IsBusyCom(Exception)`（以 word 的 7 信号版为准） | 可删 **1500+ 行×重复**；M1 busy 漂移、重试不对称一并消除；后续修 bug 改一处生效三端 |
| P1-3 | **daemon 超时看门狗** | wrapper 侧 `Invoke-Daemon`/`Invoke-WocBatch` 读操作加超时（异步 Read + 超时 Task）；helper 侧 PipeServer 对 handler 包看门狗（超时返回 TIMEOUT + `_app=null` 重 attach），与 one-shot 的 180s 对齐 | T2-1：WPS 挂起从"永久挂死"变"可恢复的超时错误" |
| P1-4 | **统一"文档消失"语义** | word 引入 attachment marker（excel/ppt 现成实现可共享）；三端统一为"非 open 命令一律 DOC_CLOSED，open 是唯一复活通道"；`QuitAll` 收敛为只关 marker 记录的本 skill 文档，Quit 前设 `DisplayAlerts=false` | T3-1/2：消灭"复活用户已丢弃文档"与"连带丢弃用户文档"两个数据丢失风险 |
| P1-5 | **tasks 系建 tasks-lib.ps1** | 收拢 Read-Json/Write-Json/FullPath2/文件发现/priority 归一/selftest 脚手架/`New-BatchOp`；三份仲裁骨架统一为"key 提取器 + 冲突策略"通用函数；split 进程回调加 `WaitForExit(timeout)`，逐 chunk 起进程改为一次 `batch -Ops` | -300 行左右；tasks-excel/ppt 骨架单源；挂起免疫 |
| P1-6 | **Excel 回归驱动器** | 建 `tests/excel/run_data.ps1`：把 excel-acceptance.md v0.3/v0.4 的 17-28 号用例脚本化（fill 2000 行/chart add/list/delete/rows/cols/snapshot 隔离/.xls 全链路），断言效仿 check.py 用 openpyxl 独立读回 | T5-1：Excel 从裸奔到有底线 |
| P1-7 | **统一测试 runner + 最小 CI** | `tests/run_all.ps1` 两层编排：fast 层（语料再生成 diff、static-only、check.py --list、tasks selftest、param-check——无 WPS 可跑，`-SkipWps` 即 CI 模式）→ wps 层（三端驱动器汇总 PASS/FAIL/耗时）。CI：GitHub Actions 只跑 fast 层 + 语料可复现 `git diff --exit-code` | T5-2/4：一条命令跑全部；纯逻辑回归进 CI，不再全靠本机人肉 |
| P1-8 | **速查表减重** | SKILL.md:123、:537 等超长行拆为"参数 + 一句话坑位 + 指针"；146 处 vN.NN 注记只留行为变化处；PPT 8 档 preset JSON 示例（:598-649）下沉 reference/ppt-presets.md；修 :836-840 粘连列表并补 changelog 链接；README hint/能力清单同步 v1.33 | T4-1/4/6/7：主文档再省 8-12K token，速查表回归"速查"本职 |

### P2 — 长期改进（按需排期）

| # | 事项 | 说明 |
|---|---|---|
| P2-1 | **巨型文件继续拆分** | WppHelper 的 render(2068-2283)→Render.cs、badge/title-marker(3072-3450)→Takeover.cs；EtHelper 的 snapshot/restore/hint/finish 外移；word 题注编号逻辑（横跨 Captions.cs、Tables.cs:1255、Tables3.cs:515 三处）合并进 Captions.cs，Tables 三件按 scan/plan/normalize 重组或至少加文件头交叉索引。拆分纪律对齐 ppt 的 Kinds/Studio 模式 |
| P2-2 | **快照/回滚三端对齐** | 路径辅助（SnapPath/DocKey/Sha16）与格式收敛到 shared；word 补 restore 命令（或文档明确声明"word 回滚=拒绝修订"是有意设计）；excel 的 selective restore 模式作为三端范本 |
| P2-3 | **共享层实弹测试** | check_clipboard.py（设置文本→render→断言剪贴板原样+PNG 非空白）；锁/忙重试加 helper `--selftest` 单元断言，把 SKILL.md 里"Lead 复跑"固化 |
| P2-4 | **static 状态治理** | `_chunkStats`/`_xmlNote` 等改为随请求重建；daemon 关键路径评估 `Marshal.ReleaseComObject`（长会话内存） |
| P2-5 | **PipeServer/Json 加固** | Escape 补全控制字符（PipeServer.cs:143-147）；`Json.Num` 收紧标准文法（WpsHelper.cs:148）；`openDocuments()` 每请求两次 COM 往返合并（PipeServer.cs:104,109） |
| P2-6 | **docs/borrow 冻结归档** | w00-w04 保留但 README 顶部声明"上游 v1.2.0 快照，落地完成即冻结"；skill-*-v13x.md 移 archive/ 或删除；tools/build_guide.py 删除或重写为从 SKILL.md 机械生成（只要还手写第五份内容就是下一个滞后源） |
| P2-7 | **sleep→条件轮询** | 25 处固定 sleep 改带超时轮询（复用 wait_gone 范式）；D3 时序断言放宽或可跳过 |
| P2-8 | **磁盘卫生** | `tests/clean.ps1` 清 .tmp//out*/；清理 scripts/helper/word/build/ 磁盘残留 exe；.gitignore 修正 tests/work/ 规则 |

---

## 五、预期收益汇总

| 维度 | 现状 | P0 后 | P1 后 |
|---|---|---|---|
| SKILL.md 加载成本 | ~29-36K token/次 | -2K（版本区） | 再 -8~12K（速查表/presets 下沉），合计降 ~40% |
| 三端重复代码 | PS ~400 行 + C# 1200+ 行×3 + tasks ~300 行 | 零散项 | **-2000 行以上**，管道协议/busy 判定/重试单源化 |
| 挂死风险 | daemon 无超时，WPS 弹框=永久挂死 | — | 看门狗+超时，与 one-shot 对齐 |
| 数据丢失风险 | word 静默复活已丢弃文档 / QuitAll 连带关用户文档 | — | marker 统一 + QuitAll 收敛 |
| 测试 | 手动串 10+ 命令；Excel 核心功能零回归 | 克隆即挂修复 | run-all 两层 runner + fast 层 CI + Excel 回归底线 |
| 版本信息 | 4 份活副本，已漂移一版 | 单源 + checklist | — |
| 破坏性操作 | publish 误传目标可递归删用户文件 | 护栏 + 原子切换 | — |

## 六、实施注意事项与明确不做的事

- **实施顺序建议**：P0-1/2/3（文档与安全）可立即做；P1-2（C# shared 扩容）必须在 P1-6/7（回归护栏）就位后动，且每步跑三端 selftest + run_replace + check_normalize `--static-only`；P1-4 改变 word 行为（不再静默重开），需在 SKILL.md 铁律区同步声明并升 minor 版本。
- **不建议做的**：①引入 pytest/框架重写测试——现有"验收脚本 + 编号断言"模式与仓库气质匹配，缺的是 runner 不是框架；②per-app 拆分 SKILL.md——docs/borrow/README.md:106 已论证 skill 无首次注入钩子，拆了反而每次都要多读文件；③把 helper 升级 C# 7+/csproj——csc.exe 零依赖自举是部署优势，C#5 的代价已在可控范围；④为消重复而引入 DLL 发布——源码链接模式零部署依赖，保持。
- **风险最高的单项**是 P1-4（语义变更）与 P2-1（大文件重组）：前者改变 agent 可见行为，后者纯搬运但 touch 面大，都应独立 commit + 全量回归。

---

*本文件为分析产物，不随 publish.ps1 发布。实施任何条目前请以当时代码为准复核行号。*
