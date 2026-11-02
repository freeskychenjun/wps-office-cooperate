# w04 · dsh-office 基础设施层深读（helper / 协议 / 渲染 / 提示 / 验收）

> 研究对象：`D:\DSH工作区\_research\dsh-office`（v1.2.0，MIT，copylee711/dsh-office，浅克隆、未改动）
> 对照对象：本 skill `D:\DSH工作区\wps-office-cooperate`（v1.28）
> 引用写法：**`D:`** = dsh-office 仓库根，**`W:`** = wps-office-cooperate 仓库根；一律 `路径:行号`。
> 阅读口径：OfficeHelper.cs 共 **9,893 行**（574 KB）、tools.ts **546 行**、verification.md **1,075 行**；我方 SKILL.md **733 行 / 77,317 字符**。
> 全部结论均来自本地代码/文档逐行核对；标注「估算」的是按代码里的常数推算、非实测。

---

## §1 组件地图

| 文件 | 职责 | 关键行 |
|---|---|---|
| `D:src/index.ts`（194 行） | Cordis 宿主插件入口：注册 8 个 `office_*` 工具、设置 schema（9 个开关）、消息卡片生命周期、**把模型从宿主自带 office 技能引回本插件**、`/api/dsh-office/status` 路由、系统提示段 | 工具注册 119-128；卡片/settle 142-156；引回 163-174；提示段 191-193 |
| `D:src/helper-client.ts`（208 行） | helper 的**运行时编译 + 哈希缓存 + 常驻子进程 + 逐行 JSON 收发** | 编译 50-90；REFERENCES 22-25；spawn/ready 149-183；超时即 kill 109-123 |
| `D:src/tools.ts`（546 行） | 工具定义（open/status/read/edit/help/render/save/close）＋**按 app 分段的 GUIDES**＋总览图策略＋交付前自动视觉检查 | GUIDES 106-207；guide-once 221-231；picture 233-259；**overview 261-282**；save 的 finalCheck 470-515 |
| `D:src/prompt.ts`（38 行） | 常驻系统提示（13 条纪律）＋**技能引回文案** | promptText 22-37；redirectText 9-19；FILE_SKILLS 5 |
| `D:src/settings.ts`（68 行） | 设置默认值 + 宽容读取（解 volatile 包装、clamp） | DEFAULTS 34-44；resolveConfig 47-67 |
| `D:src/format.ts`（8.2 KB） | helper 返回的 JSON → 模型读的文本（Word 段表 / Excel 行 / PPT 形状 / 状态） | — |
| `D:src/omml.ts`（11.9 KB） | LaTeX → OMML（WPS 不自己建公式时随批次下发） | — |
| `D:src/client/*`（4 文件） | 浏览器侧：设置页（走 `remote.settings`）+ 导航图标 + 主色 | client/index.ts 1-13、33-39 |
| **`D:helper/OfficeHelper.cs`（9,893 行）** | 常驻 C# 5 helper：COM 附着/开文档、op 分发、渲染、卡片、MessageFilter、全套 Word/Excel/PPT 编辑 | 见 §2 逐条 |
| `D:scripts/office-smoke.mjs`（188 行） | **真机冒烟**：只碰自己在 temp 里新建的文档，按 app 参数化 | 1-3、13-16、60-62 |
| `D:scripts/footnotes-smoke.mjs`（140 行） | 脚注专项：用**独立裸 COM** 造 fixture，再测被测命令 | 17-39 |
| `D:tests/*.spec.ts`（4 文件 32 项） | 纯逻辑单测：协议状态机（用 fake HelperLike）、格式、公式、打包约束 | tools.spec.ts 11-29；packaging.spec.ts 19-23 |
| `.github/workflows/ci.yml`（52 行） | 双 job：ubuntu 类型/单测/构建/pack；**windows 编译 helper 并 ping** | 11-26、30-51 |
| `D:docs/verification.md`（1,075 行） | 28 个版本的验收记录：环境→起因→改动→实测→**没有验证的** | 结构见 §4；样例 1043-1053 |
| `D:tsdown.config.ts`（40 行） | 双产物：宿主 ESM（内联宿主库）+ 浏览器 CJS（ModuleLoader 包裹、React 外置） | 13-16、32-37 |
| `D:package.json`（123 行） | 只 ship `.cs` 不 ship exe；`files`／`dsh.bundle.patch`／scripts | 28-36、40-56、57-64 |

我方对应件（对比用）：`W:SKILL.md`（733 行）、`W:scripts/{woc,office,word,excel,ppt}.ps1`、`W:scripts/woc-common.ps1`（139 行，哈希缓存 + 命名管道）、`W:scripts/param-check.ps1`（52 行）、`W:scripts/helper/shared/PipeServer.cs`（148 行）、`W:scripts/helper/{word/WpsHelper.cs(2049),excel/EtHelper.cs,ppt/WppHelper.cs(3360)}`、`W:reference/*-acceptance.md`、`W:tests/**`（python/ps1 检查器）。

---

## §2 六个专题

### §2.1 helper 分发与编译：零依赖 csc + 源码哈希缓存 + 竞态自愈

**机制**（`D:src/helper-client.ts`）

1. `REFERENCES` 只用 .NET Framework 4 自带程序集：`System.dll / System.Core.dll / Microsoft.CSharp.dll / System.Drawing.dll / System.Windows.Forms.dll / System.Web.Extensions.dll`（22-25）。全部是 Win10/11 内置，**零 npm 依赖、零预编译二进制**。
2. `frameworkDir()` 先找 `Framework64\v4.0.30319\csc.exe`，再退 `Framework\`，找不到给出可执行的人话（`启用 ".NET Framework 4.8"`）（37-44）。
3. 缓存目录 `%LOCALAPPDATA%\dsh-office\helper-<sha256前12位>[-debug]`（46-48、56-59）：**哈希只取源码**，`DSH_OFFICE_DEBUG=1` 时用 `-debug:pdbonly` 且目录名带 `-debug` 后缀，避免调试版污染发布版缓存（57-59、66）。
4. 编译输出先写 `build-<pid>-<ts>.exe` 再 `rename` 到最终名；rename 失败判为「另一个宿主进程赢了竞态，它的 exe 与本进程完全相同」→ 删临时文件、只要最终 exe 在就认（64-81）。
5. `building` 是一个进程级的 single-flight Promise：并发调用共享同一次编译；**失败则重置为 undefined 下次重编**（50、54、84-89）。这是唯一的「自愈」——源码改→哈希变→新目录重编；没有「检测 exe 损坏」的逻辑（不臆造：全文无 CRC/签名校验）。
6. 打包侧只 ship 源码：`files` 里有 `helper/OfficeHelper.cs`（`D:package.json:32`），CI 的 windows job 专门验证「它还能在真机上编译并 ping 通」（`D:.github/workflows/ci.yml:43-51`）——把「运行时编译」这条路本身纳入 CI。
7. 语言版本靠一条单测守死：`packaging.spec.ts:19-23` 断言 helper 源码**不含** `$"`（C# 6 插值串）和 `?.[a-z]`（C# 6 null 传播）。文件头也写明「Must stay C# 5」（`D:helper/OfficeHelper.cs:3`）。

**关键代码**（D:src/helper-client.ts:53-90，节选见 §5.1）

**可否移植**：**方法已一致，两个细节值得补**
- 我方 `W:scripts/word.ps1:209-245`（excel/ppt 同构）已经做：源码哈希（`Get-CachedSourceHash`，`W:scripts/woc-common.ps1:16-42`）→ `%LOCALAPPDATA%\wps-office-cooperate\word-helper-<sha>\WpsHelper.exe` → 缺失才 `csc` 编译 → 临时名 `build-<PID>.exe` 再 `Move-Item`。**结构等价**，且我方更进一步：哈希缓存键用「文件数+大小+最新 mtime」避免每条命令 50 ms 的 `Get-FileHash`（`W:scripts/woc-common.ps1:5-14`、22-26）。
- **可借①**：`DSH_OFFICE_DEBUG=1` 式的「调试构建旁路」——我方只有发布构建，遇到 helper 内部异常时没有 pdb/行号。成本：`csc` 多一个 `-debug:pdbonly` + 目录名后缀（我方已是按 sha 分目录，天然隔离）。
- **可借②**：**把「能编译」纳入 CI**。我方仓库没有 CI（`W:` 无 `.github/`），一旦某台机器 csc 版本/引用变动，只有在用户现场才炸。最小实现：windows runner 上 `word.ps1 doctor`（我方 `doctor` 已能报 `renderOk`）作为 smoke。
- **不建议借**：`rename` 竞态处理。我方是 PowerShell `Move-Item -Force`，且 create/compile 本来就是幂等，收益低。

**与我方现状对比**：分发哲学完全一致（系统 csc + 源码哈希缓存 + 首次 ~1 s），差别只在工程化外围（CI、调试旁路）。

---

### §2.2 协议 / 分发 / Bag / 错误码 / MessageFilter / 释放

#### (a) 逐行 JSON、一请求一应答、事件行

`D:helper/OfficeHelper.cs:1-3` 是三行总纲：「One JSON command per stdin line, one JSON reply per stdout line. Must stay C# 5」。

- 启动：`MessageFilter.Register()` → `Send({event:"ready"})`（394-396）；宿主靠这一行解 ready Promise（`D:src/helper-client.ts:154-158、194`）。
- 主循环（401-424）：空行跳过 → 反序列化成 `Bag` → `id = args.Raw("id")` → `Run(args.Need("cmd"), args)` → `{ok:true,result}`；异常 → `{ok:false,error,code}`；`Send` 后 `Flush`（427）。
- 每条命令**串行**：「The helper handles one command at a time; a stuck one blocks the rest」，所以宿主一旦超时就 `kill()` 整个 helper（`D:src/helper-client.ts:116-119`、130-138），下一次调用重新 spawn。超时分档：默认 20 s、open 90 s、edit 5 min、render 50–110 s（tools.ts:313/391/244/272/444）。
- 我方对比：`W:scripts/helper/shared/PipeServer.cs:9-16` 是 **4 字节长度前缀 + UTF-8 JSON**，**一请求一连接**、严格串行、跑在 `[STAThread]` 主线程；没有 `id`（连接即关联）、没有 `ready` 事件（有 `ping`）；空闲 20 分钟自退、最后一份文档关闭即退（36-40、104-116）。**两者都正确，差别是我方多了一个进程级常驻的 `--serve` 生命周期管理**，dsh-office 把「常驻」交给 stdin 管道（进程随宿主死）。
- **可借**：`id` + 每命令超时后「杀掉整个 helper」。我方 daemon 遇到卡死 COM 调用会一直占着单线程循环（只有 idle 超时救命），下一次命令全部排队。建议：管道 handler 里加看门狗（超过 `WOC_DAEMON_WATCHDOG_MS` 就 `Environment.Exit`），代价是丢 daemon 保温（下条命令重付 ~0.5 s），换「不会永久卡住」。

#### (b) op 分发：字符串顺序比较，不用反射/字典

`Run`（866-926）就是十几行 `if (cmd == "...") return ...;`，最后 `throw new Fail("BAD_ARGS", "Unknown command ...")`（925）。每个 app 的 op 分发同样风格：`WordOp(doc, type, op)`（3049）、`ExcelOp(app, doc, type, op)`（4293）、`PptOp(doc, type, op)`（6342），内部也是一串 `if (type == "...")`。**没有 map、没有反射、没有 command 注册表**——代价是文件 9,893 行，收益是 C# 5 可编译、调用链可 grep。

我方对应：`W:scripts/helper/word/WpsHelper.cs:466-494` 是 `switch (cmd)` + 结尾同一句 `unknown cmd ... use doctor|status|...` 提示（494），风格一致。

#### (c) Bag：参数取值 + 「未读字段点名」

`Bag`（22-113）是这套协议的核心设计：
- 每个 getter 都是 `Raw(key)`（58），`Raw` **无条件把 key 记入 `read` 集合**；`Points`（47-57）额外接受 `12cm/厘米/mm/毫米/in/英寸`；`Need`（64-69）缺失即 `BAD_ARGS`；`On`（96-103）把 `true/"thin"/"all"/1` 都算真。
- `Unread()`（28-33）返回「给了但没人读过」的字段名，由 `Edit` 折叠进结果文本（`D:helper/OfficeHelper.cs:9669-9670`）：
  `"… — IGNORED field(s) <names>: <op> does not take them, check the operation list"`。
  这是**幂等、零成本、可被模型自我纠正**的护栏，比抛异常友好（不中断批次）。
- `Copy()`（41-45）把全部 key 标为已读后再转发给下一个 op（样式继承用）。

我方对应 `W:scripts/param-check.ps1:22-51`，判定链一样（别名表 → `$formatFields` → 首字母小写），且处理了「显式传默认值不算丢弃」（41-44）。**差别只有出口**：我方写 **stderr**（50），dsh-office 写 **result 正文**。
**可借（P0，成本 ~20 行）**：我方其它工具都只保证 stdout 一段 JSON（`W:SKILL.md:88`），stderr 不保证被模型吃到；建议把告警塞进 JSON：`{"ok":true,"result":{...,"ignoredParams":["Tables"]}}`，并在 SKILL.md 铁律里加一句「看到 `ignoredParams` 就换参数」。这正好能兑现 `W:SKILL.md:28` 记录的已知缺口（Word 的 `-Tables` 至今无人消费）。

#### (d) 错误码：12 个码 + COM 三码归 BUSY

- 码表（全文件统计）：`BAD_ARGS` 112、`ANCHOR_MISSING` 17、`NOT_FOUND` 8、`STYLE_MISSING` 2、`ANCHOR_MISMATCH` 2、`OFFICE_ERROR`、`NOT_INSTALLED`、`NEEDS_PATH`、`STALE`、`BUSY`、`NO_DOCUMENT`、`NOT_RUNNING`；兜底 `ERROR`（429-449）。
- `Describe`（429-449）：先剥 `TargetInvocationException`（dynamic COM 调用必然包一层）；非 `Fail` 时 `Trace` 到 stderr（只在 `DSH_OFFICE_DEBUG=1` 时输，1186-1189）；`COMException` 的三个 HRESULT **0x80010001（RPC_E_CALL_REJECTED）、0x8001010A（RPC_E_SERVERCALL_RETRYLATER）、0x800AC472（Excel busy）** 一律映射成 `BUSY`，文案固定为「**The Office app is busy: the user is typing in a cell or has a dialog open. Nothing was changed by this call; try again in a moment.**」（439-443）。
- 锚点类错误文案自带「下一步该干什么」：`ANCHOR_MISMATCH` 报「第 N 段现在以 "…" 开头，不是 "…"」（2059）；`STALE` 报「段落号落在表格里，所以这张表会被塞进单元格；如果本批里前面插过表，你读到的编号已过期」（3260）；`STYLE_MISSING` 直接把「永远可用的内置样式名」列出来（2044）。**错误消息当成提示词写**，这点比我方差。

我方对应：`W:` 三端共 20 个码（`BAD_ARGS` 470、`COM_ERROR` 60、`READONLY` 25、`ANCHOR_MISSING` 24、`NOT_OPEN` 10 …），`BUSY` 只在 `0x8001010A`/`0x80010001` 两码上判（`W:scripts/helper/word/WpsHelper.cs:403-408`、`excel/EtHelper.cs:349-353`、`ppt/WppHelper.cs:398-402`）。
**缺 0x800AC472**（Excel/ET 的 busy 变体），**缺少「Nothing was changed by this call」这句关键承诺**（模型据此决定能否重试/要不要重读）。可借：S。

#### (e) MessageFilter：忙时自动重试 ~6 秒（**本次最值得抄的一条**）

`D:helper/OfficeHelper.cs:115-135`：

```csharp
[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IOleMessageFilter { ... }
class MessageFilter : IOleMessageFilter
{
    [DllImport("ole32.dll")] static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
    public static void Register() { IOleMessageFilter old; CoRegisterMessageFilter(new MessageFilter(), out old); }
    public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
    { return dwRejectType == 2 && dwTickCount < 6000 ? 150 : -1; }
    public int MessagePending(...) { return 2; }
}
```

语义：`dwRejectType == 2`（SERVERCALL_RETRYLATER）且累计等待 < 6000 ms 时**返回 150**（=150 ms 后重试），否则返回 **-1**（取消 → 抛 COMException → 被 `Describe` 变成 `BUSY`）。也就是说：用户正在单元格里打字、或开着模态对话框时，调用**自发阻塞重试约 6 秒**，而不是立刻失败；6 秒后仍忙才报错。README 对外就写成「等待约 5 秒后报告应用忙」（`D:README.md:141`）。

注册点在 `Main` 的第一件事（394），在 `SetProcessDPIAware` 之前、任何 COM 调用之前；宿主侧不需要配合。

**可否移植：P0，直接可用。**
- 我方 helper 已经是 `[STAThread]`，且**所有请求都跑在这条 STA 线程上**（`W:scripts/helper/shared/PipeServer.cs:9-12` 明写「server loop runs on the [STAThread] Main thread」），所以在 `Main` 里注册一次即对**全部命令**生效——比 dsh-office 更省事（他们也是一次）。
- 现成的失败现场：`W:SKILL.md:209` 写「超时返回 `TIMEOUT`，说明 WPS 大概率弹了对话框」——即我们今天遇到忙就是干等 180 s 超时。加了过滤器后，绝大多数「打字/对话框」场景会在 6 s 内自动通过。
- 注意（不臆造，按 Win32 语义）：过滤器只管「被 callee 拒绝」的调用；`0x800AC472` 这类**作为返回值直接抛出**的忙错误不受过滤器保护，仍需 `Describe` 里的码表兜底。另外重试期间调用线程被阻塞，我方 `-TimeoutSec` 默认 180 s 远大于 6 s，无冲突。
- 工作量：**S（30 行左右 + 一次人工验收：在 WPS 里打开模态对话框，发 `status`，应等待约 6 s 后返回 `BUSY` 而不是 180 s `TIMEOUT`）**。

**三端落地要点（我方具体改法）**
1. 位置：`W:scripts/helper/word/WpsHelper.cs`、`excel/EtHelper.cs`、`ppt/WppHelper.cs` 的 `static int Main` 里，**在第一次 `Activator.CreateInstance`/`Marshal.GetActiveObject` 之前**、`PipeServer.Run` 之前注册一次。三条 Main 都跑在 `[STAThread]` 线程上，而 `PipeServer` 的请求循环也在这条线程（`W:scripts/helper/shared/PipeServer.cs:9-12`），注册一次即覆盖后续全部请求；一次性模式（`WOC_NO_DAEMON=1`）同样覆盖。
2. 代码照抄 `D:helper/OfficeHelper.cs:115-135` 三方法：`HandleInComingCall → 0`（SERVERCALL_ISHANDLED）；`RetryRejectedCall → dwRejectType == 2 && dwTickCount < 6000 ? 150 : -1`；`MessagePending → 2`（PENDINGMSG_WAITDEFPROCESS）。注意 `[PreserveSig]` 不能省，`CoRegisterMessageFilter` 要 `[DllImport("ole32.dll")]`。
3. **不要动现有码表**：过滤器超时后抛出来的仍是 `0x80010001`/`0x8001010A`，现有 `BUSY` 判定（`W:scripts/helper/word/WpsHelper.cs:403-408` 等三处）继续兜底；顺手补 `0x800AC472`。
4. 与批量/超时的相互作用：`-TimeoutSec` 默认 180 s ≫ 6 s，不会误触发；`W:scripts/woc-common.ps1:44-56` 的 `ConnectMs=1500` 是**连接**超时、不是应答超时，不受影响；但 `batch` 里某条 op 若真的等满 6 s，整批会顺延 6 s——可接受，且比刨掉整批强。
5. 验收两条：① 模态对话框场景 ≤7 s 返回 `BUSY`；② 空闲场景 `status` 连测 5 次取中位数，确认无回归（预期 <50 ms，过滤器不参与无拒绝的调用路径）。

#### (f) 每条命令后释放 COM：不写 `ReleaseComObject`，靠双 GC

`D:helper/OfficeHelper.cs:421-423`：

```csharp
Send(reply);
// Let go of every COM object: a reference kept here would keep a closed Office app alive, hidden.
GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers();
```

同一模式在 `Settle()` 结束处再来一次（760-761）。全文件**没有一处** `Marshal.ReleaseComObject`（统计：0 次）——用「强制终结器」替代逐对象释放，避免漏放导致 RCW 泄漏。

我方现状：**三端 helper 里 `GC.Collect` / `ReleaseComObject` 命中 0 次**。我方靠「daemon 常驻 + COM application 故意保温」活着，所以不能照抄「每请求双 GC」（`snapshot` 这类大返回会二次受伤）；但**在 `close`/`settle` 路径上补一次双 GC 是纯收益**，直接对应我方已记录的痛点「WPS COM 进程泄漏与清理约定」（`W:SKILL.md:113`、`W:SKILL.md:205` 的 `quit` 说明）。
工作量：S；风险：低（只在关闭后跑）。

#### (g) 附着 vs 启动 / 进程归属 / 孤儿清扫

- `Running(kind)` = `Marshal.GetActiveObject(ProgId)`（533-537）→ 附着用户实例；`App(kind,true)` 没有才 `Activator.CreateInstance`（767-779），并 `Started.Add` 记账、Excel 额外 `UserControl = true`（777，否则自动化结束 Excel 会自杀）。
- 套件自动判定 `SuiteFor`（487-512）：命令带 `suite`，或环境变量 `DSH_OFFICE_SUITE`；**「已经打开的文档始终在它所在的套件里改」**——按 `first` 与另一套件的顺序各试一次 `Find`。
- 静默模式用**自己的隐藏实例**（`Own`）、记 `Hidden` 路径集（540-551）；`OpenQuietly`（646-684）：Word/Excel `Visible=false` + `DisplayAlerts=0`，PowerPoint 用 `Presentations.Open(path, 0, 0, 0)`（第 4 个参数是窗口，0=无窗口，665）。`Reveal`（687-702）按需出窗。
- **写了两个我很想要的巧思**：
  1. `OwnNotePath()` = `%LOCALAPPDATA%\dsh-office\background.pid`，`NoteOwn` 把隐藏实例的 PID 追加进去（553-570）；下次 helper 启动时 `SweepOwn()` **只杀「进程名 ∈ {WINWORD,EXCEL,WPS,ET,WPP} 且 MainWindowHandle == 0」的遗留 PID**（572-596）。这正好是我方「不能按映像名批量杀」约束下的安全清扫方案。
  2. `OwnApp` 里的 **decoy Word**（617-624）：当没有 Word 在跑、而我们要起一个隐藏实例时，先起一个又立刻 `Quit` 掉——这样用户随后从资源管理器双击文件时，Explorer 会**新开**一个 Word（发现文件被隐藏实例占用 → 给只读副本），而不是把文档塞进我们的隐藏实例、让用户对着看不见的窗口打字。注释写得很清楚（617-620）。
- 我方现状：主线是**附着用户正开着的 WPS 实例**（`W:SKILL.md:99-112`），`open` 默认 `show=true`（`W:SKILL.md:109`，v1.25 的教训），`close` 只报 `processes` 计数（`W:scripts/helper/word/WpsHelper.cs:1742-1773`、815-816），**没有孤儿清扫**。
  **可借（P1，S）**：`background.pid` 式记账 + 启动时「只杀无主窗口的已知进程名」。收益：daemon 崩溃/重装/`-Show:$false` 场景不再留隐藏 wps.exe（用户对此有硬要求）。

#### (h) 退出条件

- helper：`input.ReadLine() == null`（stdin 关）→ 退出（401）；`AppDomain.ProcessExit` 时 `MathQuit()` 清公式临时状态（397）。
- 应用：`Settle()`（705-765）——遍历两个套件 × 三个 app：自己的隐藏实例里 `Hidden` 的文档，未保存就**先 `FinishMath` 再 `Save`**（724-725，因为没人看得见，丢了就是丢了）；用户此刻正在看（`app.Visible` / PPT 有窗口）就**保存但不关**、转为 `Revealed`（726-735）；否则保存并关闭（736）；文档清空后 `Quit` 自己起的实例（744-757）；最后删掉 `background.pid`（762）。
- 触发点：宿主在 agent `idle` / `disposed` 时调 `settle` + 卡片收尾（`D:src/index.ts:148-156`）。
- 我方对比：`W:` 靠 `close`/`finish`/daemon 空闲 20 分钟/最后一份文档关闭（PipeServer.cs:104-116），**没有「回合结束自动善后」这个概念**（skill 形态没有 agent 生命周期钩子，只能靠铁律要求收尾调用 `close`）。可借的只剩「未保存的隐藏文档必须替用户保存」这条语义（我方 `-Show:$false` 场景适用）。

---

### §2.3 渲染管线与**总览图（contact sheet）**（直接可借鉴项）

#### (a) 三条管线，全部不依赖截图

| app | 取图方式 | 位置 |
|---|---|---|
| Word | `doc.Windows[1].Panes[1].Pages[n].EnhMetaFileBits` → `Metafile` → GDI+ 绘制 | 9726-9753（单页）、9829-9857（Render 内联同一套） |
| PowerPoint | `slide.Export(path, "PNG", width, height)`，高=宽×`SlideHeight/SlideWidth` | 9820-9826；拼图时同一个 `Slide.Export` 到临时文件（9777） |
| Excel | `range.CopyPicture(1, 2)`（xlScreen, xlBitmap）→ 剪贴板 → `Clipboard.GetImage()` 轮询 10×100 ms → **恢复用户原剪贴板文本** | 9871-9887；单图 9865（`holder.Chart.Export`） |

- **抗「细线消失」的超采样**（9734-9750 / 9838-9840 的注释与实现）：直接按目标宽度画 EMF 时，表格边框这种 hairline 会落在像素之间而消失。做法：`factor = max(2, ceil(3000 / width))`，先画 `width*factor` 的大图，再用 `HighQualityBicubic` 缩到目标宽度。
  > 我方 Word 侧已经**一模一样**实现了这段（`W:scripts/helper/word/WpsHelper.cs:1593-1616`，`factor` 在 1599）；PPT 侧仍直接按目标宽度画 EMF（`W:scripts/helper/ppt/WppHelper.cs:1986-1997`），**总览格子只有 ~550 px 时最容易丢线** → 这是可借的一条（见 §3 P1-3）。
- Excel 的剪贴板礼仪：先 `Clipboard.ContainsText()` 存下文本，画完 `SetText` 放回或 `Clear()`（9873-9884）。**我方 `W:scripts/helper/excel/EtHelper.cs:1800` 是直接 `Clipboard.Clear()`**——会吞掉用户剪贴板内容。可借（P2，5 行）。
- 一条已被我方否决、必须写清的路：PPT 侧 dsh-office 用 `Slide.Export`，**我方不能用**——`W:scripts/helper/ppt/WppHelper.cs:988` 明确记录「Slide.Export is not used because wpp.exe-written files are transparently encrypted on this machine family」（本机透明加密驱动会把 wpp 写出的文件变密文）。所以 PPT 侧只能走 `Slide.Copy` → CF_ENHMETAFILE（1876-1888）。

#### (b) 总览图：算法与常数（**核心可借鉴项**）

宿主策略 `D:src/tools.ts:261-282`：

```
overview(app) —— excel 退化为单张 used-range；
  for (from = 1; from <= min(total, 18); from += 6):
      render {sheet:true, from, to: from+5, out, width: 640}   // 超时 110 s
      push 图；total = result.total
  if (total > 18) 追加一行 "Only the first 18 of <total> are shown."
```
→ 硬上限：**每张 6 页、最多 18 页（=3 张图）**；单元格宽度**固定 640**（比单页 1100 小）。

helper 排版 `D:helper/OfficeHelper.cs:9756-9812`：

```csharp
int cell = Math.Max(300, Math.Min(1000, a.Int("width", 640)));         // 9759
int total = word ? pages.Count : slides.Count;                          // 9761
int first = max(1, from), last = min(total, to);                        // 9763
last = Math.Min(last, first + 11);                                      // 9765  单张最多 12 格
int count = last-first+1, columns = Min(count, word ? 3 : 2);           // 9766  Word 3 列 / PPT 2 列
int rows = (count + columns - 1) / columns;
int gap = 14, label = 30, height = max(每格高);                          // 9781
sheet = Bitmap(columns*cell + (columns+1)*gap,
               rows*(height + label + gap) + gap);                      // 9783
g.Clear(Color.FromArgb(232,232,236));                                   // 9789 浅灰底
font = new Font("Microsoft YaHei UI", 14f, Bold, Pixel);                 // 9786
每格: DrawString("page N" / "slide N") 在 (x, y+4)                            // 9793
      DrawImage(pic, x, y+label)                                        // 9794
      DrawRectangle(150 灰, x, y+label, w-1, h-1)                       // 9795
```
返回 `{what:"pages 1–6 of 37", first, last, total, path, width, height}`（9806-9810）——`what` 里的 `of N` 被宿主用正则抠出来当总页数（`D:src/tools.ts:247-248`），并据此**提前截断**循环。

**token 权衡（估算，按 A4 1:1.414、16:9 由代码常数推算）**

| 场景 | 像素 | 说明 |
|---|---|---|
| Word 单页 @1100 | 1100×1556 = **1.71 MP** | 默认 `renderWidth=1100`（`D:src/settings.ts:43`，clamp 600–2000 见 66） |
| Word 总览 1 张（6 页，3×2，cell 640） | 1976×1912 = **3.78 MP** | ≈ 0.63 MP/页 = 单页的 **37%** |
| 6 页各单渲 | 10.3 MP | 总览省 **2.7×**，且只发 1 张图（1 次工具调用 vs 6 次） |
| 全稿 18 页 | 3 张 ≈ **11.3 MP** | 与「6 页单渲」几乎等价，但看的是 18 页 |
| PPT 总览 1 张（6 页，2×3，cell 640） | 1322×1226 = **1.62 MP** | ≈ 0.27 MP/页；单页 @1100 = 0.68 MP → 省 **2.5×** |

**它买到的不是省 token，而是「一次看全 + 分级放大」**：overview 是 640 px/页的**粗筛**（排版错位、半空页、图表出框这类问题在这个分辨率就看得出来），发现可疑页再 `render {page:n}` 用 1100 px 放大看。配合 `finalCheck`（见 §2.5）形成闭环：保存时如果模型自上次「看过」之后又改过，**自动附总览图 + 一段强制过一遍的文案**（`D:src/tools.ts:500-509`）。

#### (c) 我方现状与差距

- **PPT 已有 `-Overview`**：`W:scripts/helper/ppt/WppHelper.cs:1896-1961`，每张 6 页、≤24 页/次、2 列、`pad=18 / label=26`、`cellW = clamp(width/2, 320..720)`、标签「第 N 页」、背景 246 灰、边线 205 灰、字体 微软雅黑 12（非粗体）；SKILL.md 已把「设计阶段渲一次 `-Overview`，之后只渲改动页」写成纪律（`W:SKILL.md:486`、519、682）。**结构上与 dsh-office 同源，且有两处更好**：格宽随 `-Width` 自适应、单次上限 24 页（不必分 3 次）。
- **Word 完全没有总览**：`W:SKILL.md:179` 只有 `-Page 1 [-To 4]`；`W:scripts/helper/word/WpsHelper.cs:1556-1591` 只逐页出图。**这是最大空白**：一份 60 页报告「交付前过一遍」现在要 15 次 `render`（每页 ~355 ms 渲染 + 每次一次工具往返）。
- **Excel 无总览**：`W:SKILL.md:325` 单区域图 + `-IncludeCharts`。Excel 的痛点是"表多"而非"页多"，**建议只做"多表并联"（P2），不做 contact sheet**，理由与做法见 (d)。
- **PPT 总览缺超采样**：格宽 550 px 时细线/细边框会丢（对照 `W:` Word 侧 1599 已有 factor）。**P1，10 行**。

#### (d) 给我方 Word / Excel 加 `-Overview` 的具体做法（含风险）

**Word（值得做，成本最低的一条渲染改动）**

复用现有零件，不需要任何新的取图路径：

1. 参数层（`W:scripts/word.ps1`）：`param` 块加 `[switch]$Overview`；`'render'` 分支（`W:scripts/word.ps1:370-377`）加
   ```powershell
   if ($Overview) { $request['overview'] = $true }
   elseif ($To -gt 0) { $request['to'] = $To }   # 总览与 -To 语义冲突时以总览为准，并在返回里说明
   ```
   `-Width` 对总览**不水涨船高**：helper 侧把 cell 夹在 320–720（默认 640），否则 token 爆炸。
2. helper 层（`W:scripts/helper/word/WpsHelper.cs:1556-1591` 的 `RenderDoc`）：入口先
   `if (a.Flag("overview", false)) return RenderOverviewDoc(a);`，新函数照抄 `D:helper/OfficeHelper.cs:9756-9812` 的骨架：
   - `int cell = clamp(640, 320..720)`；`columns = min(count, 3)`（Word 是竖版，3 列正好一张 A4 横排的观感）；
   - `last = min(last, first + 11)` 单张 ≤12 格；宿主/参数层再加「一次 ≤18 页」的硬上限；
   - `gap = 14 / label = 30`、浅灰底 `(232,232,236)`、每格 `DrawString("第 N 页")` + `DrawImage` + `DrawRectangle(150 灰)`；
   - 每页图直接调**已有的** `SaveEmfAsPng(bits, tmp, cell)`（`W:scripts/helper/word/WpsHelper.cs:1593-1616`，内部已有 `factor = max(2, ceil(3000/width))` 超采样）→ 画进 sheet → `Dispose`，峰值内存只有 1 张 sheet + 1 页临时位图；
   - 返回 `{ overview: [路径...], pages, sheets, capped: bool, note }`，命名沿用我方 PPT 的 `overview-<from>-<to>.png`（`W:scripts/helper/ppt/WppHelper.cs:1933`）。
3. 风险（都已在我方代码/实测里有依据）：
   - **视图类型必须先切**：`Pages.Count` 只在页面视图下可靠，现有 `RenderDoc` 已经 `if ((int)win.View.Type != 3) win.View.Type = 3`（`W:scripts/helper/word/WpsHelper.cs:1566`）；总览同样必须先切，否则 0 页或错页。
   - **每页一次 COM 往返**：60 页全稿总览 = 60 次 `EnhMetaFileBits`；按我方量级（单页渲染 ~百毫秒级，PPT 侧 355 ms/页可参考）18 页约 2–6 s，可接受；**坚持 6/18 上限，不做"一次拼全稿"**。
   - **中文字体**：标签用「第 N 页」时字体必须能显示中文（微软雅黑/宋体），否则方框；dsh-office 用 `Microsoft YaHei UI 14px Bold`，我方 PPT 总览用 微软雅黑 12、非粗体（`W:scripts/helper/ppt/WppHelper.cs:1936`）——建议 Word 与 PPT 统一成同一组常数，别再长第三套。
   - **页码与内容对应**：`Pages.Item(n)` 在含分节/横向节的文档里页序仍是全局页序，但要靠验收确认（造一份含横向节的测试稿）。
   - **不要做"整页原尺寸拼图"**：cell 一旦跟着 `-Width` 变大，18 页的 contact sheet 会变十几 MP，token 立刻失控。

**Excel（建议只做"多表并联"，不做 contact sheet）**

Excel 的真实痛点是"表多"而不是"页多"，而 contact sheet 会强迫不同列宽的表格共用一格宽：

- **推荐做法（路 A，成本 ~40 行）**：`render -Overview` = 对每张工作表各出一张 used-range 图（各自用原 `-Width`，不缩放），最多 6 张/次，返回 `{ sheets: [{name, path, width, height}], more }`。收益：一次调用看清"哪张表有图表/哪张表数据跑偏"，且**数字仍然可读**（Excel 字号小，缩到 640 会看不清）。
- **不推荐（路 B，真拼图）**：`CopyPicture` 逐表取图再 2 列拼接。风险：① 剪贴板必须串行（我方已有 `Local\woc-clipboard` 互斥，`W:SKILL.md:22`），批量拼图会放大串行等待；② 宽高比差异巨大（`A1:K8` 与 `A1:Z300` 混排出现极端留白/挤压）；③ token 收益不如"每表一张"；④ `CopyPicture` 已经把与区域相交的浮动图表拍进去（`W:SKILL.md:345`），拼图后再缩会把这些细节一起压掉。
- 结论：**Word 做 contact sheet（P0），Excel 只做多表并联（P2，可延后）**。当前"每表一张 + `-IncludeCharts`"已经够用，不是阻塞项。

---

### §2.4 体验层：逐字 / 跟随 / 静默 / 迷你卡片 / show:true

| 能力 | dsh-office 实现 | 位置 | skill 形态判断 |
|---|---|---|---|
| **逐字写入** | `Pieces(text)`：切成 **≤14 片**（`size = max(3, ceil(len/14))`），每片 `Sleep(14)` → 一段约 0.2 s（设置页文案写"约多 0.3 s"，`D:src/index.ts:55`；实现对 14 片即 196 ms）。Word 走 `range.InsertAfter` 逐片；PPT 同一套（`PptWrite`）；Excel 按行 `Sleep(25/40)` | 1274-1299（Word）、5680-5690（PPT）、3318/3983/4320/6531（Excel） | **不建议移植**。它是"给人看"的演出效果，代价是每段 +0.2 s；我方单页渲染已 355 ms，逐字只会让批次更慢，且 skill 的读者是模型不是人 |
| **跟随滚动** | Word：`doc.Windows[1].ScrollIntoView(range, true)`；**WPS 特例**：先把 caret 移到 `range.End` 再滚（否则 WPS 每次改动把视口拉回原 caret，画面在编辑点与 caret 之间跳）。Excel：用 `window.VisibleRange` 判断目标格子是否已在视野，不在才 `ScrollRow = row-3` / `ScrollColumn = col-1`（避免多余抖动）。PPT：`deck.Windows[1].View.GotoSlide(slide)` | 1265-1272；4277-4291；6250-6254 | **我方已实现且更细**：`W:scripts/helper/word/Toc.cs:549-608` 的 `ViewDoc` 用 `ScrollIntoView(target)` **不动 caret**，`-Focus` 才 `Select()+"ScrollIntoView(Selection.Range)"`，并回读 `caretBefore` … 无需再借 |
| **静默模式** | 隐藏实例 `Visible=false`+`DisplayAlerts=0`；PPT `Presentations.Open(path,0,0,0)`；`silent` 时强制 `Following=Typing=card=false`；回合结束 `settle` 自动保存并关闭 | 598-630、646-684、9648-9651、705-765 | **部分可借**：我方 `open -Show:$false` 已有；「静默时同时关掉跟随/逐字/卡片」我方没有这三种模式，不适用；`settle` 的「未保存就替用户存」语义可借（见 §3 P1-5） |
| **迷你卡片**（`Card : Form`） | 独立 STA 线程跑 WinForms（343-356），`ShowWithoutActivation`（194）不抢焦点，两枚可点 chip「跟随 / 快速」（276），500 ms 计时器刷新；**用户一旦点过，卡片的选择就覆盖设置**（`FollowChosen/TypingChosen`，145 + 9643-9644）；卡片停留 25 s（`Linger`，1263） | 140-378 | **不做**。理由：① skill 形态没有常驻 UI 进程；② 用户已明确否决任何置顶悬浮窗（`W:SKILL.md:134`，2026-10-30 决策，且写了"不要再做成可选模式偷偷加回来"）。唯一值得记的是**优先级语义**（用户手动选择 > 配置），我方可用一个 state 文件实现，价值低 |
| **`office_open {show:true}`** | 静默模式下临时把文档带到前台，供"只能用鼠标键盘做的事"用（`D:src/tools.ts:310`、97） | tools.ts 310、`OfficeHelper.Reveal` 687-702 | **已有**：`W:` `open -Show`（默认 true，v1.25 起，`W:SKILL.md:109`、`W:scripts/word.ps1:296-305`） |

补充一条非体验但同源的：`D:src/index.ts:142-156` 用 `tools/pre-execute` 把 agent 记进 `WeakSet`、`agent/status === 'idle'` 时收卡片并 `settle`。这是「回合结束」的唯一可靠信号，**skill 形态拿不到**（没有 agent 生命周期钩子），所以我方只能靠铁律要求 `finish`/`close`（`W:SKILL.md:113`、206）。

---

### §2.5 提示注入与上下文成本（GUIDES 分段 / 引回 / 未知字段 / silent）

#### (a) 三层提示，按"何时付 token"分层

| 层 | 内容 | 体积 | 何时进上下文 |
|---|---|---|---|
| 系统提示段 | `promptText()` 13 条纪律（何时用 `office_*`、别加载 office-docx/pptx、公式/引用/样式/交付前 polish） | **5,407 字符** | 每请求常驻（`D:src/index.ts:191-193`，`order: 73`） |
| 工具 schema | 8 个工具 + 参数说明（含 `SILENT` 4 条长文案） | 中等 | 每请求常驻 |
| **按 app 的 GUIDES** | Word 11,195 / Excel 5,388 / PPT 19,742 字符 | 合计 **36,325 字符** | **某个 app 第一次被用时，随 open/read/edit 的结果给一次**；之后 `office_help {app}` 再取 |

`GUIDES` 的实现（`D:src/tools.ts:106-207`）+ 一次性逻辑（221-231）：

```ts
const told = new Map<string, Set<AppKind>>()          // 每个 conversation 一份
const guide = (exec, kind) => {
  const seen = told.get(session) ?? new Set()
  if (seen.has(kind)) return ''
  seen.add(kind)
  return `\n\n${APP_NAMES[kind]} operations for office_edit (this list is given once; office_help {app:"${kind}"} gives it again):\n${GUIDES[kind]}`
}
```
挂载点：`office_open`（333）、`office_read`（374）、`office_edit`（403）三处结果尾部；`office_help` 会**同时**把该 app 标为已给，避免重复（422-423）。README 对外表述见 `D:README.md:49`「用不到的应用不占上下文」。

**这套设计的代价与收益**：模型第一次动手前就知道该 app 的**全部 op 与字段**（PPT 那份 19.7 KB 的版式说明尤其值），但不碰的 app 一个 token 都不花。

**我方现状（这是本次最刺眼的差距）**：`W:SKILL.md` **733 行 / 77,317 字符**，skill 一加载就全部进上下文：

| 段 | 行 | 字符 | 占 |
|---|---|---|---|
| frontmatter（1-7）+ **版本 changelog（8-84，v1.28 往前到 v1.10 及更早）** | 1-90 | **20,750**（其中 changelog 本身 **19,721**） | **26.8%** |
| 调用方式说明（`<skill>` 路径、-Mta、输出契约） | 85-89 | 434 | 0.6% |
| 铁律 + 接管提示 + 多任务 | 91-162 | 6,887 | 8.9% |
| Word 全章 | 163-305 | 15,225 | 19.7% |
| Excel 全章 | 306-472 | 12,435 | 16.1% |
| PPT 全章 | 473-714 | 19,445 | 25.2% |
| 与其它 skill 关系 + 参考 | 715-733 | 1,835 | 2.4% |

（另有 `W:reference/` 163 KB 按需读，不进常驻。）

**可借（P1，判断后折中）**：
1. **把版本 changelog 迁出 SKILL.md**（`W:SKILL.md:8-84`，19,721 字符 → `reference/changelog.md`，SKILL.md 只留一行指针 + 最新版本 3 行摘要）。**省 25.5%，零功能损失**，纯搬运。dsh-office 的历史全部躺在 `docs/verification.md`（1,075 行），从不进上下文——这是它上下文预算能守住的一半原因。
2. **按 app 拆命令表**：Word-only 会话现在要为 Excel/PPT 的 31,880 字符一起付费。折中方案（不照抄"首次注入"，因为 skill 没有这个钩子）：SKILL.md 每章只留「一页速查（命令名 + 一句话）+ 指针」，长的参数详解移到 `reference/app-word.md` 等；模型进入某 app 前 `Read` 一次。收益：**只做 Word 的会话省 ~40%（31.9k chars）**；代价：多用 1 次工具调用、且 reference 的 token 一样要进上下文（只是晚一步）。**建议先做 1（changelog），2 观察一轮再定**——理由见 §3 P2-2 的风险说明。

#### (b) 把模型从宿主自带 office 技能引回本插件

`D:src/prompt.ts:5`：`FILE_SKILLS = { 'office-docx': 'word', 'office-pptx': 'ppt' }`——**`office-xlsx` 故意不拦**（1-4 行注释：表格工作多是数据分析，代码更好；`D:docs/verification.md:1020` 记录了这个决策）。
`redirectText()`（13-19）三段式：① 机器上有 Office/WPS，这件事在 app 里用 `office_*` 做；② 具体路径 `office_open → read → edit → render → save`；③ **"确实做不到（批量转换、app 打不开的文件）就再调一次这个技能，它会加载"**——把"第二次放行"写进给模型看的文案里。

拦截实现 `D:src/index.ts:163-174`（关键：**在 `tools/post-execute` 里 `accept` 并替换 content**）：

```ts
ctx.on('tools/post-execute', async (exec, result, next) => {
  if (exec.name !== 'skill' || result.isError || !settings().preferLive) return next()
  const app = FILE_SKILLS[String(exec.arguments?.name ?? '')]
  if (app === undefined) return next()
  const key = `${session}:${skill}`
  if (redirected.has(key) || !(await installedApps()).has(app)) return next()   // 每会话一次；app 未装不拦
  redirected.add(key)
  return { kind: 'accept', content: [{ type: 'text', text: redirectText(skill, app) }] }
})
```
`installedApps()` 缓存一次 `status` 查询（`D:src/index.ts:134-140`），`preferLive` 设置可整体关（`D:src/settings.ts:27`）。策略动机写在 130-133 的注释：「a loaded skill outweighs the system prompt」。

**最贵的教训写在 verification.md 里，必须抄这一条方法论**：1.0.5 用 `tools/execute` 替换结果**无效**——「执行阶段返回的结果如果不是宿主自己生成的那个对象，宿主会只取其中的 `value`，用技能自己的渲染函数重新生成给模型看的文本，替换的内容被丢弃」（`D:docs/verification.md:1003`）；1.0.7 改到 `post-execute` 返回「接受并替换内容」才成（1007），因为宿主这一支「直接把内容换掉、其余不动」（同一处）。**结论：要拦截宿主内置能力的输出，先读宿主实现里的 `normalizeDispatchResult` / `postExecute` / `applyFinalContent`（1011），别猜钩子。**

**对我方的价值**：`W:SKILL.md:715-722`「与其它 skill / 工具的关系」目前是**纯文字劝说**（"不要加载 design-report-review / docx-paragraph-polish …"）。skill 形态拿不到 post-execute 钩子，但可以直接**搬运它的三段式文案**（尤其是"确有需要再调一次即放行"的出口，能降低模型与文档对抗的概率），并把「会坏」的机械理由（文件被独占、写盘 `PermissionError`）放在最前面——我方 L106 已经这么写了，可以把 dsh-office 的措辞并入。

#### (c) 未知字段点名 + silent 分支提示

- 未知字段：见 §2.2(c)。要点是**文案自解释 + 不中断批次**（`IGNORED field(s) …: <op> does not take them, check the operation list`）。
- `SILENT` 提示词（`D:src/tools.ts:94-97`）4 条：文档在后台无人观看可以照常干；**Excel 的数据活可以改用 Python**（前提是工作簿没在 Excel 里打开）；结束前仍然必须 `office_render` 看结果；只有必须在窗口上做的事才 `office_open {show:true}`（慢，慎用）。每次会话第一次静默 `open` 时给一次（294-301 `quietTold`）。
  这 4 条与我方「无人值守 / 用户不在场」场景高度重合，可整段改写进 `W:SKILL.md` 的无人值守分支（`W:SKILL.md:107` 的"唯一例外"那段）。

---

### §2.6 工程化：打包 / 构建 / CI / 测试三层 / 冒烟

**打包**（`D:package.json`）
- `files` 只放：`lib/index.js`、`lib/client.js`、`lib/types/**/*.d.ts`、**`helper/OfficeHelper.cs`**、`cordis.patch.yml`、README、LICENSE（28-36）——**不 ship 任何二进制**，运行期编译。
- `os: ["win32"]`（37-39）；`dsh.bundle.patch` 指向 `cordis.patch.yml`（40-43）；`dsh.engines.dsh: ">=0.2.0-rc.2"`（44-46）；客户端 `dsh.client.inject` 显式声明 4 个宿主包（47-55）。
- 宿主相关包全在 `peerDependencies` 且 `optional: true`（86-101）——插件不硬绑宿主版本。
- `scripts`：`typecheck`(tsc --noEmit) / `build`(tsc && tsdown) / `test`(vitest run) / `pack:check`(pack --dry-run) / `smoke`（57-64）。

**构建**（`D:tsdown.config.ts`）：双产物。宿主侧 ESM，注释明写「DSH host libraries are inlined on purpose: the plugin must not depend on the exact rc the host ships」（13-16）；浏览器侧 CJS，**用 banner/footer 手工包成 ModuleLoader**（32-37）+ 外置 react（29）+ 烘焙 `process.env.NODE_ENV`（30-31，因为宿主 webview 没有 `process`）。

**CI**（`D:.github/workflows/ci.yml`）：`verify`（ubuntu：install → typecheck → test → build → pack:check，11-26）＋ `helper`（**windows-latest**：build 后 `node -e "ensureHelperExe() → ping → status"`，30-51）。后者是这个项目最有辨识度的 CI 设计——**把"在你的机器上能编译"变成 CI 断言**。

**测试三层**（我方几乎没有第①层）
1. **纯逻辑单测**（vitest，32 项，不需要 Office）：`tools.spec.ts:11-29` 用 **fake `HelperLike`**（记录 calls、按 cmd 返回假结果）测完整状态机：
   - `guide` 只给一次、`office_help` 再取（122-151）；静默模式参数（178-195）；保存时附总览（206-230）；多页渲染在上限处截断（231-244）；无 vision 的模型不给图（256-262）；`close` 只关自己开的（94-109）；save-as 后目标切到新文件（170-177）。
   - `packaging.spec.ts` 三条：包名/公开/files 含 `.cs`；**helper 必须保持 C# 5**（`not.toMatch(/\$"/)`、`not.toMatch(/\?\.[A-Za-z]/)`，19-23）；client bundle 首 200 字符含 `id: "<pkg>"`（29-32）。
   - `packaging.spec.ts` 里的 settings 块（34-40）：默认值、volatile 解包、clamp。
2. **真机冒烟**（`D:scripts/office-smoke.mjs`，188 行，node 直接 fork helper exe）：**只碰自己在 `mkdtemp` 里新建的文档**（1-3）；按 app 参数化 `node scripts/office-smoke.mjs word|excel|ppt`（15-16）、`--wps` 换套件（13-14）、`--keep` 留现场；开跑先 `status`，并**清理上一次崩溃留下的自己目录里的文档**（60-62）；每个断言 `check(name, ok, detail)`（49），失败计数非零则非 0 退出；末尾打印每条命令耗时（44-47）。
3. **专项冒烟**（`D:scripts/footnotes-smoke.mjs`）：**不拿被测插件生成底稿**——先用一段裸 COM PowerShell 脚本（17-39，`kwps.Application` / `Word.Application`）造出带脚注的 fixture，再让 helper 去改，最后用 `assert/strict` 断言编号与文本。这条「独立造底稿」的做法避免「用被测代码验证被测代码」。

**验收记录**：`D:docs/verification.md` 1,075 行 / 28 个版本节，结构见 §4。

**我方差距**：`W:` 有真机检查器（`W:tests/word/check_daemon.py`、`check_hint.py`、`check_tray.py`、`W:tests/ppt/*.py` 等）与验收记录（`W:reference/word-acceptance.md`、`excel-acceptance.md`），**但**：① 没有一层"不需要 Office 的纯逻辑单测"（参数消费校验、JSON 转义、快照 diff 算法、tasks-*.ps1 的 merge/plan 都是纯逻辑，可测）；② 没有 CI；③ 验收记录缺「没有验证的」小节（见 §4）。

---

## §3 优先级建议

| # | 建议 | 优先级 | 理由 | 工作量 | 风险/注意 |
|---|---|---|---|---|---|
| 1 | **MessageFilter（`CoRegisterMessageFilter` + `RetryRejectedCall` 150 ms / 6 s）** | **P0** | 用户正在打字/开对话框时从「180 s 超时」变「≤6 s 自动通过」；三端共 30 行，注册一次覆盖全部命令（我方请求都跑在 `[STAThread]` Main 上） | **S**（半天含验收） | 只覆盖"被 callee 拒绝"；`0x800AC472` 仍需码表兜底；重试期间占用单线程循环约 6 s |
| 2 | **Word `render -Overview`（contact sheet）** | **P0** | 60 页报告「交付前过一遍」从 15 次调用降到 3 张图；算法、常数、标签、边框、上限在 `D:helper/OfficeHelper.cs:9756-9812` 可直接照搬（我方 PPT 侧已有同源代码，`W:scripts/helper/ppt/WppHelper.cs:1896-1961`）。具体改法与风险见 §2.3(d) | **M**（1 天） | 必须先切页面视图再取 `Pages`；`-Width` 对总览不生效（cell 夹 320–720）；坚持 6/18 上限 |
| 2b | **Excel `-Overview` = 多表并联**（每表一张 used-range 图，不拼图） | **P2** | Excel 痛点是"表多"；contact sheet 逼不同列宽共格，数字看不清（§2.3(d)） | **S**（~40 行） | 别做真拼图：剪贴板串行 + 宽高比差异 + 与 `-IncludeCharts` 图表语义叠加会失真 |
| 3 | **`ignoredParams` 进 JSON result**（替代 stderr `WARN`） | **P0** | 我方承诺 stdout 只有一段 JSON（`W:SKILL.md:88`），stderr 不保证被模型读到；这条正好兑现 `-Tables` 无人消费的缺口（`W:SKILL.md:28`） | **S**（20 行 × 3 端） | 保持 `WOC_NO_PARAMCHECK=1`；字段名要短（`ignoredParams`），别把 3 个 app 的措辞写成三套 |
| 4 | **PPT 总览超采样**（`factor = max(2, ceil(3000/cellW))`） | **P1** | 我方 PPT 直接按 550 px 画 EMF（`W:scripts/helper/ppt/WppHelper.cs:1986-1997`），细线在总览里最易消失；Word 侧早已实现（`W:scripts/helper/word/WpsHelper.cs:1599`），可整段搬 | **S**（10 行） | 内存峰值 = `cellW*factor`²；总览 550×2=1100 宽，无压力 |
| 5 | **孤儿隐藏实例清扫**（`background.pid` + 只杀 `MainWindowHandle == 0` 的已知映像） | **P1** | 用户硬要求「不得留下大量 WPS 进程」；我方现在只报数不清扫（`W:scripts/helper/word/WpsHelper.cs:1742-1773`）。方案 `D:helper/OfficeHelper.cs:553-596` 已把「按映像名批量杀」这个禁忌处理干净 | **S–M**（半天） | 判据必须严格（进程名白名单 + 无主窗口 + 由我们记账），**不能**扫所有同名进程 |
| 6 | **Excel render 恢复用户剪贴板**（存文本→放回） | **P1** | 我方 `Clipboard.Clear()`（`W:scripts/helper/excel/EtHelper.cs:1800`）会吞掉用户剪贴板 | **S**（5 行） | 只恢复文本（CF_UNICODETEXT），图片无法无损还原——文档里别承诺"剪贴板不变" |
| 7 | **`save` 的「没看过就别交付」提示**（finalCheck 的 skill 化） | **P1** | dsh-office 在保存时**自动附总览图**（`D:src/tools.ts:500-509`）；我方 CLI 不能让 helper 直接给模型图片，但可以让 `save` 返回 `unseenSinceLastRender: true` + 一句硬话，SKILL.md 铁律接住 | **S**（helper 记 `lastRenderAt`/`editedSince`，~30 行） | 需要 helper 侧一个「上次 render 之后有没有 edit」的计数器；只提醒，不阻断 |
| 8 | **`GC.Collect`×2 放在 `close`/`settle` 路径** | **P1** | 我方 COM 释放完全依赖进程存活；关闭后强制终结器是廉价保险（`D:helper/OfficeHelper.cs:421-423`、760-761） | **S**（2 行 × 3 端 + 复测 snapshot 时延） | **不要**每请求都跑（大返回会二次受伤）；只在 close/quit 后 |
| 9 | **版本 changelog 迁出 SKILL.md** | **P1** | `W:SKILL.md:8-84` = 19,721 字符 = 常驻上下文的 **25.5%**，纯历史信息 | **S**（纯搬运 + 一行指针） | 迁移时保留最新 3 条摘要，防止"看不到最近改了什么" |
| 10 | **验收记录加「没有验证的」小节 + 环境头** | **P1** | 见 §4；这是 dsh-office 记录体系里最不可替代的一节 | **S**（每次改动 +5 行） | 要写具体的"没测什么/为什么"，不是"未充分测试" |
| 11 | **错误文案补「Nothing was changed by this call」** | **P2** | 让模型知道能否安全重试（`D:helper/OfficeHelper.cs:439-443`） | **S** | 要按码区分：`BUSY` 可重试，`ANCHOR_MISSING` 必须重读 |
| 12 | **按 app 拆 SKILL.md 命令表** | **P2** | Word-only 会话可省 ~31.9k 字符常驻 | **M–L**（要把三章拆文件 + 改指针 + 复测模型能否找对） | **风险**：skill 没有"首次注入"钩子，Read 一次 reference 的 token 一样进上下文；只在"整场不用该 app"时省钱。建议先做 #9 再看 |
| 13 | **纯逻辑单测层（Pester 或 python 调 ps1 纯函数）** | **P2** | 参数消费、JSON 转义、`tasks-*.ps1` 的 merge/plan 都是纯逻辑，现在只能靠真机 | **M** | 我方 wrapper 是"边绑定边组 request"，需要先抽出纯函数才可测（重构成本） |
| 14 | **CI（windows runner：`word.ps1 doctor` + `tests/` 里不依赖 GUI 的那部分）** | **P2** | `D:.github/workflows/ci.yml:43-51` 的"能编译"断言 | **M** | 本机没有 CI 习惯；至少要有一键 `run-tests.ps1` |

**明确不抄**（写下来免得后人重复提议）：
- 迷你卡片 / 置顶浮层（用户 2026-10-30 已否决，`W:SKILL.md:134`；形态也不符）。
- 逐字写入（skill 的读者是模型；纯延迟）。
- PPT `Slide.Export`（本机透明加密驱动会把 wpp 写出的文件变密文，`W:scripts/helper/ppt/WppHelper.cs:988`）。
- Word 隐藏实例 + decoy 技巧（我方主线是附着用户实例，只有 `-Show:$false` 极小场景；decoy 反而要起两次 Word 进程）。
- Excel 总览图（一张 used-range 已是全貌；`-IncludeCharts` 已解决图表可见性）。

---

## §4 可复用的验收清单（从 `D:docs/verification.md` 1,075 行提炼）

### §4.1 版本条目模板（每改一轮写一节，7 项，缺一不可）

```
## <版本号>（<日期>）
<环境头：Windows 版本 / Office 或 WPS 版本号 / 宿主 DSH 版本 / 模型名>
### 起因
   用户原话（引用）+ 从会话记录或成品里倒查到的机械原因
### 改动
   逐条：改了什么 → 为什么是这么改（一句）
### 实测 / 验证
   a) 定点复现：先在真实 app 里用临时文档复现问题，再确认修好（含具体输入）
   b) 回归：pnpm typecheck / test（N 项）/ smoke（哪个 app）/ 我方对应 = check_*.py、check_daemon.py
   c) 真实模型端到端：任务 / 方式 / 步数 / 耗时 / token，必要时带对照组
### 没有验证的
   逐条列：哪个分支没跑到、为什么、下次怎么测（这是本模板的核心）
```

出处：`D:docs/verification.md:1-5`（环境头）、`D:docs/verification.md:62`（起因）、`D:docs/verification.md:1043-1053`（先复现再修 + 新增冒烟项 3 条）、`D:docs/verification.md:52-58`、`969-1074`（"没有验证的"写得具体到分支）。

### §4.2 三层验证的金字塔（每层各自的通过判据）

| 层 | 跑什么 | 通过判据 | 出处 |
|---|---|---|---|
| L1 纯逻辑 | 不启动 Office 的单测 | 断言「协议状态」而不是「像素」：清单只给一次、静默参数、保存时附总览、多页截断、无 vision 不给图、`close` 只关自己的 | `D:tests/tools.spec.ts:122-262` |
| L2 真机冒烟 | 直接 fork helper，只碰自建临时文档 | ① 每个动作后的**读回比对**（不信返回值）；② 一次撤销能撤回整批；③ 锚点不符即拒绝且批次停在该步；④ 未知字段被点名；⑤ 打印逐命令耗时 | `D:scripts/office-smoke.mjs:49、52、72-86`；耗时表 `verification.md:22` |
| L3 真实模型端到端 | 隔离宿主 + 固定任务 | ① 任务固定（Word/Excel/PPT 各一条）；② 判据机械（文件存在且 >4 KB）；③ **有对照组**（同样任务用 computer-use 跑一遍：8 步 16.4 s / 104K vs 16 步 167.3 s / 538K） | `D:docs/verification.md:24-37` |
| L4 环境矩阵 | 每类环境一行 | WPS vs 微软 Office、125% 缩放、无微软 Office 的机器、另一台机器（对方实测报告） | `verification.md:986`（两个套件都看过）、`W:SKILL.md:479`（另一台机器 70 条命令实测） |

### §4.3 冒烟/验收的隔离纪律（5 条，直接可抄）

1. **只碰自己在 `mkdtemp` 里新建的东西**；脚本头部把这条写成注释（`office-smoke.mjs:1-3`）。
2. **开跑先清理自己上一次崩溃留下的现场**（按 temp 前缀匹配 + 只关匹配到的路径，`office-smoke.mjs:60-62`）。
3. **fixture 不用被测代码生成**（脚注 fixture 用裸 COM PowerShell 造，`footnotes-smoke.mjs:17-39`）。
4. **可参数化到单个 app / 换套件 / 留现场**：`word|excel|ppt`、`--wps`、`--keep`（`office-smoke.mjs:12-16`）。
5. **失败要能定位**：`check(name, ok, detail)` 每条打一行，`detail` 带上实际值（`office-smoke.mjs:49`）。

### §4.4 从 28 个版本节里反复出现的「验收盲区」，做成自查表

改动前后逐条自问（每条都能在 verification.md 里找到踩过的实例）：

- [ ] **只在"我造的场景"里验证了吗**？（0.1.0 的"没有验证的"第 1 条：用户与 AI 同时编辑没测过）
- [ ] **靠"数量"当判据了吗**？（0.6.0：动画只数了文件里多少个，没放映看过）
- [ ] **修完之后有没有用真实模型重跑触发它的那个任务**？（0.6.0 铁路汇报修完后未重跑；1.1.1 未重跑）
- [ ] **"不移动、只提示"这类分支有实例吗**？（1.1.2：留白>1/3 的分支没造出实例）
- [ ] **另一个套件/WPS 跑过吗**？（1.1.1：WPS 走另一条公式路径，没测）
- [ ] **控件真的点过吗**？（0.2.0：卡片开关"只做了代码走查"；0.6.1："用程序触发同一段代码"，不是真鼠标点）
- [ ] **性能数据是"观察"还是"计时"**？（0.2.0：跟随只读窗口滚动百分比；0.1.0 给了每命令耗时区间）
- [ ] **旧路径回退还在吗**？（我方 `-Slow` / `textdiag` 是同类，`W:SKILL.md:174、186`）
- [ ] **写进文档的"保证"有没有代码兜**？（例：`W:SKILL.md:88` 承诺 stdout 只有 JSON——那就不能依赖 stderr 传告警，见 §3 #3）

### §4.5 反向清单：什么样的验收记录等于没写

- 只写"已测试通过"，不写**测的是什么输入**。
- 只写改动，不写**起因**（下一个 agent 会把它当成可删的复杂度）。
- 不写**没有验证的**（读者会默认全绿）。
- 真实模型只跑一次却给出"稳定"结论（dsh-office 每次都说"各 1 次，数字只说明量级"，`verification.md:35`）。
- 把"代码里有"当成"验证过"（0.1.0 明写"Office 忙时的等待与报错：代码里有，未实测"）。

---

## §5 原文摘录（每条 ≤30 行，均带行号）

### §5.1 零依赖编译 + 哈希缓存 + 竞态（`D:src/helper-client.ts:50-90`）

```ts
let building: Promise<string> | undefined

/** Compile the helper if this source version has no cached exe yet; returns the exe path. */
export async function ensureHelperExe(source = helperSourcePath()): Promise<string> {
  building ??= (async () => {
    const code = await readFile(source)
    const hash = createHash('sha256').update(code).digest('hex').slice(0, 12)
    // With DSH_OFFICE_DEBUG=1 the helper is built with line numbers for the stack traces it then writes to stderr.
    const debug = process.env.DSH_OFFICE_DEBUG === '1'
    const dir = join(cacheRoot(), `helper-${hash}${debug ? '-debug' : ''}`)
    const exe = join(dir, 'office-helper.exe')
    if (existsSync(exe)) return exe
    await mkdir(dir, { recursive: true })
    const fw = frameworkDir()
    const temp = join(dir, `build-${process.pid}-${Date.now()}.exe`)
    const args = [
      '-nologo', ...(debug ? ['-debug:pdbonly'] : ['-optimize+']), '-target:exe', `-out:${temp}`,
      ...REFERENCES.map(ref => `-r:${ref}`), source,
    ]
    await new Promise<void>((resolve, reject) => {
      execFile(join(fw, 'csc.exe'), args, { windowsHide: true, timeout: 120_000 }, (error, stdout) => {
        if (error) reject(new Error(`Compiling the office helper failed: ${String(stdout).trim() || error.message}`))
        else resolve()
      })
    })
    try {
      await rename(temp, exe)
    } catch (error) {
      // Another host process won the race; its exe is identical.
      await rm(temp, { force: true })
      if (!existsSync(exe)) throw error
    }
    return exe
  })()
  try {
    return await building
  } catch (error) {
    building = undefined
    throw error
  }
}
```

### §5.2 MessageFilter：忙时重试约 6 秒（`D:helper/OfficeHelper.cs:115-135`）

```csharp
[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IOleMessageFilter
{
    [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

/// Office rejects calls while the user has a dialog open or is typing: wait for it a few seconds.
class MessageFilter : IOleMessageFilter
{
    [DllImport("ole32.dll")]
    static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
    public static void Register() { IOleMessageFilter old; CoRegisterMessageFilter(new MessageFilter(), out old); }
    public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) { return 0; }
    public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
    {
        return dwRejectType == 2 && dwTickCount < 6000 ? 150 : -1;
    }
    public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) { return 2; }
}
```

### §5.3 三个 HRESULT 归 BUSY + 「本次调用没有改动任何东西」（`D:helper/OfficeHelper.cs:429-449`）

```csharp
static string Describe(Exception error, out string code)
{
    while (error is System.Reflection.TargetInvocationException && error.InnerException != null) error = error.InnerException;
    Fail fail = error as Fail;
    if (fail == null) Trace("error: " + error);
    if (fail != null) { code = fail.Code; return fail.Message; }
    COMException com = error as COMException;
    if (com != null)
    {
        uint hr = unchecked((uint)com.ErrorCode);
        if (hr == 0x80010001 || hr == 0x8001010A || hr == 0x800AC472)
        {
            code = "BUSY";
            return "The Office app is busy: the user is typing in a cell or has a dialog open. Nothing was changed by this call; try again in a moment.";
        }
        code = "OFFICE_ERROR";
        return com.Message.Trim() + " (0x" + hr.ToString("X8") + ")";
    }
    code = "ERROR";
    return error.Message;
}
```

### §5.4 未知字段点名（`D:helper/OfficeHelper.cs:9664-9671`）

```csharp
Bag op = new Bag(ops[i]);
type = op.Need("op");
if (card) Card.Hold = true;
if (card) Card.Report("AI 正在编辑 " + docName + (ops.Count > 1 ? " · " + (i + 1) + "/" + ops.Count : ""));
string line = kind == "word" ? WordOp(doc, type, op) : kind == "excel" ? ExcelOp(app, doc, type, op) : PptOp(doc, type, op);
string ignored = op.Unread();
done.Add(ignored.Length == 0 ? line : line + " — IGNORED field(s) " + ignored + ": " + type + " does not take them, check the operation list");
```

### §5.5 总览图的排版算法（`D:helper/OfficeHelper.cs:9756-9766` + `9781-9799`）

```csharp
/// Several pages or slides side by side in one picture, each with its number: the whole document at a glance.
static object Sheet(string kind, dynamic doc, Bag a)
{
    string path = Path.GetFullPath(a.Need("out"));
    int cell = Math.Max(300, Math.Min(1000, a.Int("width", 640)));
    dynamic pages = kind == "word" ? doc.Windows[1].Panes[1].Pages : null;
    int total = kind == "word" ? (int)pages.Count : (int)doc.Slides.Count;
    if (total == 0) throw new Fail("ANCHOR_MISSING", "There is nothing to picture yet.");
    int first = Math.Max(1, a.Int("from", 1)), last = Math.Min(total, a.Int("to", total));
    if (first > total) throw new Fail("ANCHOR_MISSING", "There is no " + (kind == "word" ? "page " : "slide ") + first + " (there are " + total + ").");
    last = Math.Min(last, first + 11);
    int count = last - first + 1, columns = Math.Min(count, kind == "word" ? 3 : 2), rows = (count + columns - 1) / columns;
```
```csharp
    int gap = 14, label = 30, height = 0;
    foreach (Bitmap picture in pictures) height = Math.Max(height, picture.Height);
    using (Bitmap sheet = new Bitmap(columns * cell + (columns + 1) * gap, rows * (height + label + gap) + gap))
    {
        using (Graphics g = Graphics.FromImage(sheet))
        using (Font font = new Font("Microsoft YaHei UI", 14f, FontStyle.Bold, GraphicsUnit.Pixel))
        using (Pen border = new Pen(Color.FromArgb(150, 150, 150)))
        {
            g.Clear(Color.FromArgb(232, 232, 236));
            for (int k = 0; k < pictures.Count; k++)
            {
                int x = gap + (k % columns) * (cell + gap), y = gap + (k / columns) * (height + label + gap);
                g.DrawString((kind == "word" ? "page " : "slide ") + (first + k), font, Brushes.Black, x, y + 4);
                g.DrawImage(pictures[k], x, y + label);
                g.DrawRectangle(border, x, y + label, pictures[k].Width - 1, pictures[k].Height - 1);
            }
        }
        sheet.Save(path, ImageFormat.Png);
    }
```

### §5.6 抗 hairline 消失的超采样（`D:helper/OfficeHelper.cs:9726-9753`）

```csharp
/// One Word page as a picture of the given width.
static Bitmap WordPage(dynamic pages, int page, int width)
{
    byte[] bits = (byte[])pages[page].EnhMetaFileBits;
    using (MemoryStream stream = new MemoryStream(bits))
    using (Metafile meta = new Metafile(stream))
    {
        int height = (int)Math.Round(width * (double)meta.Height / meta.Width);
        // Drawn straight at the final size, hairlines (table borders) fall between pixels and vanish:
        // draw the page large first, then scale the picture down.
        int factor = Math.Max(2, (int)Math.Ceiling(3000.0 / width));
        Bitmap bitmap = new Bitmap(width, height);
        using (Bitmap large = new Bitmap(width * factor, height * factor))
        {
            using (Graphics g = Graphics.FromImage(large))
            {
                g.Clear(Color.White);
                g.DrawImage(meta, 0, 0, large.Width, large.Height);
            }
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(large, 0, 0, width, height);
            }
        }
        return bitmap;
    }
}
```

### §5.7 总览的宿主侧策略与上限（`D:src/tools.ts:261-282`）

```ts
/** The whole document at a glance: its pages (6 to a picture) or slides tiled with their numbers, up to 18. */
const overview = async (on: { app: AppKind; doc?: string }): Promise<{ lines: string[]; images: AttachmentJson[] }> => {
  if (on.app === 'excel') return picture(on, 0, 0, {}, Math.min(host.settings().renderWidth, 1000))
  const folder = join(tmpdir(), 'dsh-office')
  await mkdir(folder, { recursive: true })
  const lines: string[] = []
  const images: AttachmentJson[] = []
  let total = 1
  for (let from = 1; from <= Math.min(total, 18); from += 6) {
    const out = join(folder, `sheet-${process.pid}-${Date.now()}-${from}.png`)
    try {
      const result = await helper.call<{ what: string; total: number; width: number; height: number }>('render', { ...on, sheet: true, from, to: from + 5, out, width: 640 }, 110_000)
      total = result.total
      images.push(toJson(await host.saveImage(await readFile(out), `${on.app}-overview-${from}.png`)))
      lines.push(`${APP_NAMES[on.app]} ${result.what}, tiled in one picture.`)
    } finally {
      await rm(out, { force: true })
    }
  }
  if (total > 18) lines.push(`Only the first 18 of ${total} are shown.`)
  return { lines, images }
}
```

### §5.8 GUIDES 只给一次（`D:src/tools.ts:221-231`）

```ts
// Which apps' operation lists each conversation has been given already.
const told = new Map<string, Set<AppKind>>()
const sessionOf = (exec: unknown): string => (exec as { agent?: { session?: { id?: string } } } | undefined)?.agent?.session?.id ?? 'default'
/** The operation list of an app, the first time a conversation works with it ('' afterwards). */
const guide = (exec: unknown, kind: AppKind): string => {
  const session = sessionOf(exec)
  let seen = told.get(session)
  if (!seen) { seen = new Set(); told.set(session, seen) }
  if (seen.has(kind)) return ''
  seen.add(kind)
  return `\n\n${APP_NAMES[kind]} operations for office_edit (this list is given once; office_help {app: "${kind}"} gives it again):\n${GUIDES[kind]}`
}
```

### §5.9 孤儿隐藏实例清扫（`D:helper/OfficeHelper.cs:572-596`）

```csharp
/// Hidden instances a helper before this one left behind: ended, if they still show no window.
static void SweepOwn()
{
    try
    {
        string note = OwnNotePath();
        if (!File.Exists(note)) return;
        foreach (string line in File.ReadAllLines(note))
        {
            int pid;
            if (!int.TryParse(line.Trim(), out pid) || pid <= 0) continue;
            try
            {
                using (System.Diagnostics.Process process = System.Diagnostics.Process.GetProcessById(pid))
                {
                    string name = process.ProcessName.ToUpperInvariant();
                    if ((name == "WINWORD" || name == "EXCEL" || name == "WPS" || name == "ET" || name == "WPP") && process.MainWindowHandle == IntPtr.Zero) process.Kill();
                }
            }
            catch (Exception) { }
        }
        File.Delete(note);
    }
    catch (Exception) { }
}
```

### §5.10 交付前的"看过没有"判定（`D:src/tools.ts:500-509`）

```ts
      // The whole-document check before delivery: the model has edited since it last looked, so it is shown the result.
      if (!wasUnseen || !host.settings().finalCheck || !await host.vision(exec)) return { text }
      try {
        const now = moved ? { app: on.app, doc: result.path } : on
        const shown = await overview(now)
        const [image, ...more] = shown.images
        return {
          text: `${text}\nFinal check: you edited after you last looked at the document, so here is the whole of it as it is now. ${shown.lines.join(' ')}\nGo through every page before you tell the user it is done. If anything is off (content outside its frame, uneven fonts or spacing, formulas as plain text, bad tables, leftover hints, blank gaps, a nearly empty last page), look closer with office_render, fix it with office_edit and save again; if all is well, say so.`,
          ...(image === undefined ? {} : { image }), ...(more.length ? { more } : {}),
        }
      } catch {
        return { text }
      }
```

---

## 附：本次核过但**未**采用的判断依据（防止后人重复研究）

| 项 | 结论 | 依据 |
|---|---|---|
| `Slide.Export` 直出 PNG | 不可用于本机 | `W:scripts/helper/ppt/WppHelper.cs:988`：本机家族的透明加密会把 wpp 写出的文件变密文 |
| 每请求 `GC.Collect`×2 | 不照抄 | 我方 daemon 故意保温 COM；只在 close/settle 后加（§3 #8） |
| Word 隐藏实例 / decoy Word | 不做 | 我方主线附着用户实例；decoy 要额外起一次 Word 进程 |
| Excel contact sheet（拼图） | 不做；只做"多表并联" | 不同列宽共格导致数字看不清；`-IncludeCharts` 已解决图表（`W:SKILL.md:325、345`、§2.3(d)） |
| 迷你卡片（`Card : Form`） | 不做 | 用户已否决浮层（`W:SKILL.md:134`）；skill 无常驻 UI 进程 |
| `tools/execute` 阶段替换技能结果 | **不要试** | `D:docs/verification.md:1003`：宿主会用技能自己的渲染函数重建文本，替换被丢弃；只能用 `tools/post-execute` 的 `accept`（1007） |
| `office-xlsx` 拦回 | 与 dsh-office 同判 | 数据分析用代码更合适（`D:src/prompt.ts:1-5`、`D:docs/verification.md:1020`） |
