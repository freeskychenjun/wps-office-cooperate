# skill-infra v1.32：OLE MessageFilter（忙重试）＋ 剪贴板保存/恢复

日期：2026-11-03　环境：WPS 12.1.0.29161（`KET.Application` / `KWPS.Application` / `KWPP.Application`，`doctor.app.build` = `12.1.0.29161`），Windows PowerShell 5.1，系统 `csc.exe`（.NET Framework 4）
范围（只动这些）：
- 新建 `scripts/helper/shared/MessageFilter.cs`、`scripts/helper/shared/ClipboardKeeper.cs`（**随包自动生效**：三端包装器本来就把 `helper/shared/*.cs` 编进每个 helper）
- 三个 helper 的 `Main` 各加一行注册、`BUSY` 码表各加 `0x800AC472`（`word/WpsHelper.cs`、`excel/EtHelper.cs`、`ppt/WppHelper.cs`，外加 PPT 的 `ComRetry`）
- Excel `TryRenderRange`、PPT `RenderPres` / `RenderOverview` / `Deliver.ExportPng` + `CaptureSlide` 的剪贴板段
- 本文与 `reference/workflows.md` §15

**未改**：`SKILL.md`、`publish.ps1`、`scripts/*.ps1`、`reference/` 之外的其它业务代码、`WocLocks.cs`、`PipeServer.cs`。

---

## 0. 一句话

① 三个 helper 现在都装了 OLE 消息过滤器：**WPS 拒绝调用时（用户在打字 / 有对话框）会以 150 ms 为间隔重试，最多 6 s**，仍不行才抛给现有错误映射变成 `BUSY`；
② Excel / PPT 的 `render`、`export -Format png|jpg` 不再吞掉用户剪贴板——**进去先存、出来放回**，且只在"这一轮没人重新复制"时才放回，放回的数据带"不进剪贴板历史"标记。

---

## 1. MessageFilter（忙重试）

### 1.1 上游依据与策略

移植自 dsh-office（MIT）`helper/OfficeHelper.cs:115-135`，策略逐字照抄：

| 方法 | 返回值 | 含义 |
|---|---|---|
| `HandleInComingCall` | `0` | `SERVERCALL_ISHANDLED`：进来的调用照常处理 |
| `RetryRejectedCall` | `dwRejectType == 2`（`SERVERCALL_RETRYLATER`）且 `dwTickCount < 6000` → **`150`**；否则 **`-1`** | 150 ms 后重试；6 s 后仍忙就**取消调用**，让 `COMException` 抛出来，落到 `HandleRequest` 的码表变成 `BUSY`（可重试） |
| `MessagePending` | `2` | `PENDINGMSG_WAITDEFPROCESS`：等默认处理（上游同值） |

`IOleMessageFilter` 用 `[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(InterfaceIsIUnknown)]`，三个方法都必须带 `[PreserveSig]`（否则 HRESULT 会被运行库翻译成异常，返回值语义全丢）。注册用 `[DllImport("ole32.dll")] CoRegisterMessageFilter`。

### 1.2 注册点

三个 helper 的 `Main` **第一条语句**（`[STAThread]`、在任何 COM 调用与 `PipeServer.Run` 之前）：

```csharp
[STAThread]
private static int Main(string[] args)
{
    // v1.32: install the OLE message filter before any COM call ...
    MessageFilter.Register();
    Console.OutputEncoding = new UTF8Encoding(false);
    ...
```

一次注册覆盖全部命令的理由：helper 是 `[STAThread]`，`--serve` 常驻模式的请求循环也跑在这条 Main 线程上（见 `helper/shared/PipeServer.cs` 文件头），消息过滤器属于**公寓**而不属于某次调用，所以一次性模式与常驻模式都被覆盖。

`Register()` 返回注册的 HRESULT（`0` = `S_OK`），并把它记进 `MessageFilter.LastHResult`；同时**静态持有**过滤器实例（COM 持有 CCW 引用，但托管侧被回收就会留下悬空引用）。设 `WOC_MSGFLT_DEBUG=1` 时往 stderr 打一行：

```
woc-msgfilter: CoRegisterMessageFilter -> 0x00000000 (replaced=no)
```

（默认为静默：三端包装器承诺 stdout 只有一段 JSON。）

### 1.3 忙码补 `0x800AC472`

`0x800AC472`（Excel/ET "应用忙"）原来只在 dsh-office 的码表里，我方三端都没有。已补：

| 位置 | 现状 |
|---|---|
| `excel/EtHelper.cs` `HandleRequest` | `0x8001010A` / `0x80010001` / **`0x800AC472`** / "busy" / "对话框" |
| `ppt/WppHelper.cs` `HandleRequest` | 同上（三条码 + 文案） |
| `ppt/WppHelper.cs` `ComRetry<T>` | 重试判据补 `0x800AC472`（原来只认两条 RPC 码） |
| `word/WpsHelper.cs` `HandleRequest` | 原来只认 `0x8001010A`（缺 `0x80010001`），现补齐 `0x80010001` + `0x800AC472` |

---

## 2. ClipboardKeeper（剪贴板保存/恢复）

### 2.1 机制

```csharp
using (ClipboardKeeper.Keep())
{
    range.CopyPicture(1, 2);
    ClipboardKeeper.Sync();      // "剪贴板现在这个状态是我造成的"
    ...
    Clipboard.Clear();
    ClipboardKeeper.Sync();
}
```

- **进入**：记 `GetClipboardSequenceNumber()`（P/Invoke `user32.dll`）+ 把当前剪贴板**每个格式**各取一份放进一个 `DataObject`（`CanIncludeInClipboardHistory` / `CanUploadToCloudClipboard` / `ExcludeClipboardContentFromMonitorProcessing` 三个格式跳过，恢复时重新写）。
- **`Sync()`**：把我们刚写出的状态记为"我们的"序列号。**每当我们自己动了剪贴板就要调一次**，否则恢复判据没有基准。
- **退出**：只有 `GetClipboardSequenceNumber() == 我们最后一次 Sync 记下的值`（即这一轮之后没人再动过）才恢复；否则判定"用户中途复制了新东西，新的更该留"，**不动**并写 Trace。
- **恢复**：照上游做法先给数据打三个格式——`ExcludeClipboardContentFromMonitorProcessing = DWORD 1`、`CanIncludeInClipboardHistory = DWORD 0`、`CanUploadToCloudClipboard = DWORD 0`——再 `Clipboard.SetDataObject(data, true)`。目的是**别把用户原来的内容推进 Win+V / 云剪贴板**（对办公机是真泄漏）。
- **失败不抛**：`catch` + `Trace.WriteLine` + `WOC_CLIP_DEBUG=1` 时 stderr 一行；最坏情况是渲染图留在剪贴板上（等于旧行为），绝不让已经成功的 render 因为恢复失败而报错。
- **嵌套**：`Keep()` 套 `Keep()` 时内层不存不还，只有最外层负责恢复（`Deliver.CaptureSlide` 里留了一层保险，平时它是 no-op）。

### 2.2 覆盖的段

| 位置 | 段 |
|---|---|
| `excel/EtHelper.cs` `TryRenderRange` | `CopyPicture → ContainsImage/GetImage → Save(png) → Clipboard.Clear`（整段在 `using` 内，`WocLocks.Clipboard` 互斥仍在外面，恢复发生在**释放跨进程锁之前**） |
| `ppt/WppHelper.cs` `RenderPres` | 整个 `for` 页循环（一次 Keep 覆盖 N 页，避免每页一次恢复抖动） |
| `ppt/WppHelper.cs` `RenderOverview` | 同上（拼联系表的多页循环） |
| `ppt/Deliver.cs` `ExportPng` | 整个导出循环；`CaptureSlide` 里再套一层（嵌套安全） |

### 2.3 取证（同一份源码单独编成探针，`%TEMP%\woc-infra-test\InfraProbe.exe`）

```
== ClipboardKeeper: text put back ==
clipboard after render = [user-before-123] (expect [user-before-123]) seq 7961 -> 7982
history-control formats present: CanIncludeInClipboardHistory=True CanUploadToCloudClipboard=True ExcludeClipboardContentFromMonitorProcessing=True
== ClipboardKeeper: user copies meanwhile, ours must not win ==
clipboard = [user-after-789] (expect [user-after-789])
== ClipboardKeeper: an image put back ==
after render: ContainsImage=True (expect True)
== ClipboardKeeper: nothing was on the clipboard ==
after render: text=[<none>] ContainsImage=False (expect <none>/False = cleared back to empty)
PROBE RESULT: PASS
```

真机路径（`EtHelper.exe` 一次性直跑，`WOC_CLIP_DEBUG=1`）：

```
BEFORE: user-clip-abc
woc-clip: keep: seq 8069, saved 1 format(s)
woc-clip: restore: 1 format(s) put back
AFTER : user-clip-abc
{"ok":true,"cmd":"render","result":{...,"png":"...\\out-excel-dbg\\Data-A1_F20-w.png"}}
```

PPT 同构：

```
BEFORE: user-clip-ppt-xyz
woc-clip: keep: seq 8108, saved 1 format(s)
woc-clip: restore: 1 format(s) put back
AFTER : user-clip-ppt-xyz
```

---

## 3. 验收台账（2026-11-03 06:46–07:05）

| 项 | 结果 |
|---|---|
| `doctor` × 3 | `ok:true`（三端都重新编译成功；`word` renderOk true、`excel`/`ppt` `ok:true`） |
| Excel 剪贴板 | `Set-Clipboard '用户内容-123'` → `render -Range A1:F20` → `Get-Clipboard` = **用户内容-123**（`\u7528\u6237\u5185\u5BB9-123`，8 字符，`equalToExpected: True`），PNG 156,920 B |
| PPT 剪贴板 | `Set-Clipboard '用户内容-123'` → `render -From 1 -To 1` → `Get-Clipboard` = **用户内容-123**，`slide-1.png` 65,849 B |
| 过滤器注册 | 真机 `EtHelper.exe` 一次性跑：`CoRegisterMessageFilter -> 0x00000000 (replaced=no)` |
| 过滤器策略 | 探针：`HandleInComingCall→0`、`RetryRejectedCall(2,0)→150`、`(2,5000)→150`、`(2,6000)→-1`、`(1,0)→-1`、`MessagePending→2`，与 §1.1 表一致 |
| 不回归（Excel） | `read`（值正确）、`render`（图正常）、`render -Overview`、`render -IncludeCharts`（`chartsIncluded:0` + 提示）、`status`、`snapshot`（117 格）、`diff`（H1 一处 added）、`write -Cell H1`；渲染图目检：20 行数据/表头全在，无空白、无错位 |
| 不回归（PPT） | `read`（2 形状）、`outline`、`check`（`mechanically clean`，blocker 0）、`render`（`slide-1.png` 目检正常：标题 + 副标题）、`render -Overview`、`close` |
| 进程 | 测试后：`et = 0`、`wpp = 0`、`wps = 8`（8 个是开工前就在的 Writer 实例，测试前 `doctor` 也是 8）；`excel serve -Action stop` 已发；`ppt` daemon 随最后一份稿关闭自动退出 |

---

## 4. 没有验证的（说清楚，不假称验过）

1. **真实"用户正在编辑"忙场景没能造出来**。两次尝试都失败，如实记录：
   - **单元格编辑态**：`Application.SendKeys("{F2}")` 之后，同一瞬间 `Workbooks.Count` **4 ms 正常返回**（未拒），说明本机 WPS 12.1 在单元格编辑态**不拒绝**外部 COM 调用；
   - **模态"是否保存"提示**：把工作簿置脏后由后台作业调 `Workbooks.Close()`（作业一直阻塞），此时另一进程的 `Workbooks.Count` 仍 **4 ms 成功**，`EnumWindows` 也**没找到 `#32770` 对话框**（WPS 用的是自己的提示层，或者本实例被设成不提示）；
   - **重载忙（一次大调用占住 STA）**：准备 1200×52 数组反复 `Value2=` 压测的脚本因上一步留下的空工作簿（`Workbooks.Item(1)` 为 null）没跑起来，未再做第三次尝试（时间到）。
   因此本项的结论只能是：**过滤器已注册（HRESULT 0x00000000）且策略值已用同一份源码验证**；"拒绝 → 重试 → 6 s 后 BUSY" 的端到端路径**未在真机触发过**。
2. `0x800AC472` 只是码表新增，**没有真的收到过这个码**（同样的原因：忙碌态造不出来）。
3. 剪贴板恢复只实测了 **纯文本** 与 **位图** 两种内容；Excel 富文本、文件列表（CF_HDROP）、HTML 片段等**没测**。恢复到"整份 `DataObject` 一次性写回"这条路遇到不可序列化格式时可能整次恢复失败（下游是 Trace 而非异常，行为等同旧版）。
4. 恢复判据里"用户在**我们自己最后一次写剪贴板之后**又复制了"才算安全；如果用户恰好在"我们按 CopyPicture 之前 → 我们 Clear 之后"这段极窄窗口里复制，Clear 已经吃掉他的内容，恢复会把旧内容放回去（与上游同一缺陷，窗口是几百毫秒级）。
5. 三端同跑（Excel 与 PPT 同时 render）时的恢复交错**没测**：跨进程互斥 `WocLocks.Clipboard` 语义未改，Keep 在锁内，理论上串行，但没有并发实测。

---

## 5. 上游对照与取舍

- 上游是"只存文本、画完 `SetText` 放回"（`§2.3(a)` 里记的旧结论）；我们照的是上游**公式链路**那套（`SaveClipboard/RestoreClipboard`，`OfficeHelper.cs:5768-5825`）：整份格式快照 + 序列号判据 + 历史抑制，能力更强，代价是恢复时多一次 `SetDataObject`。
- 上游把 `ClipOurs` 记在"粘贴之后"；我们把它抽成显式 `Sync()`，因为渲染段里我们**会写两次**（CopyPicture 与 Clear），只记一次会把第二个状态误判成"别人动的"。
- 上游 `MessagePending` 返回 2，我们照抄（`150` 是 `RetryRejectedCall` 的返回值；`MessagePending` 的合法值是 2/3，150 不是这个方法的语义）。
