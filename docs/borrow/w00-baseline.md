# w00 基线地图：wps-office-cooperate v1.28（只读勘察，供"借鉴 dsh-office"落地直接引用）

> 勘察对象：`D:\DSH工作区\wps-office-cooperate`（git `master`，HEAD `aac25a0 v1.28 (ppt): slide badge is opt-in again…`，**无 tag**，版本号只存在于 commit subject + `SKILL.md` 首段）。
> 勘察方式：只读源码 / 未启动 WPS / 未改动任何脚本。行号以 `HEAD` 实际文件为准（1 基）。
> 路径缩写：`<S>` = `D:\DSH工作区\wps-office-cooperate`，`<D>` = `C:\Users\Administrator\.agents\skills\wps-office-cooperate`（实测存在，非 junction）。

## 0. 一页速览（先看这几行）

| 事实 | 值 |
|---|---|
| 对外命令总数 | **108** 个 ValidateSet 命令名 = Word **39** + Excel **28** + PPT **41**；其中 `batch`/`serve` 各 3 端共 6 个是传输/编排层，Word 的 `units` 是 `unit` 别名 ⇒ **101 个不同的文档能力实现** |
| 入口 | `<S>\scripts\woc.ps1 <word\|excel\|ppt> <cmd> [args]`（v1.17 推荐，见 `woc.ps1:35`）→ `powershell -Mta -NoProfile -ExecutionPolicy Bypass -File <app>.ps1`；另有 `woc.cmd`（cmd.exe 调用，公式会坏）、`office.ps1`（分发器）、`<app>.ps1`（可直呼） |
| 实现形态 | PowerShell 5.1 包装器（只做参数→JSON）+ C# helper（COM 全在这层）+ 常驻命名管道 daemon；输出恒为一段 JSON `{"ok":true,"result":{…}}` / `{"ok":false,"error":{code,message}}` |
| C# 规模 | Word 9 752 行、Excel 2 784 行、PPT 6 932 行、shared 735 行（合计 ≈2.0 万行）；包装层 1 977 行 |
| 编译/缓存 | 首次调用用系统 `csc.exe` 编译（约 1 s），按**源码哈希**缓存到 `%LOCALAPPDATA%\wps-office-cooperate\<app>-helper-<sha12>\`；**改任一 .cs 自动重编 + 新管道名**，旧 daemon 闲置自灭 |
| 性能（SKILL.md 实测记录） | daemon 热路径每条 ≈0.5–1.6 s；`-Sta` 2.2 s vs `-Mta` 0.13 s ⇒ 包装器必须 MTA，helper 必须 `[STAThread]` |
| 版本基线 | WPS Office 12.1.0.29161（无微软 Office），所有 quirks 都在这版实测 |
| 部署 | `.\publish.ps1` → `<D>`；当前部署与 dev **同步**（`SKILL.md` 两侧 mtime 均 2026-11-02 19:58:14，共 52 个文件） |

---

## 1. 实现分层与一条完整调用链

### 1.1 分层表（谁负责什么）

| 层 | 文件（行数） | 职责 | 关键锚点 |
|---|---|---|---|
| 壳入口 | `scripts\woc.ps1`(36) / `woc.cmd`(36) / `office.ps1`(38) | 只做转发：`woc.ps1` 用 `& powershell -Mta … -File <app>.ps1` 避免 cmd.exe 改参数；`office.ps1` 校验 app 名再转发 | woc.ps1:24 拼目标、29 取其余参数、35 起子进程；office.ps1:25-37 |
| 包装层 | `scripts\word.ps1`(685) / `excel.ps1`(552) / `ppt.ps1`(740) | ① `param()` 声明命令行参数；② `ValidateSet` 白名单；③ helper 源码发现+哈希+编译；④ daemon 启停与管道收发；⑤ **命令 switch 把参数翻译成 helper 请求 JSON**；⑥ param-check 告警；⑦ batch 分支；⑧ 一次性回退 | word: param 9-122、sources 139-146、hash 213-224、compile 230-245、batch 255-271、request 273-275、switch 295-627、param-check 629-646、daemon 调用 651-658、回退 660-681 |
| 公共管道 | `scripts\woc-common.ps1`(139) | `Get-CachedSourceHash`（16-42，按"文件数+大小+最新 mtime"缓存 SHA，省每条命令 ~50 ms）、`Invoke-WocBatch`（86-138，一次进程跑 N 个 helper 请求） | 被三端 dot-source（word:213-214） |
| 参数消费校验 | `scripts\param-check.ps1`(52) | `Test-ParamConsumption`（22-51）：**绑定了却没进 request 的参数在 stderr 告警**（stdout JSON 不受影响）；解析顺序 别名表→格式字段表→首字母小写（48）；`WOC_NO_PARAMCHECK=1` 关闭 | 三端共用 |
| helper 传输 | `scripts\helper\shared\PipeServer.cs`(148) | 常驻服务：`--serve <pipe>` → `Run(pipe, HandleRequest, OpenDocumentCount)`（34-118）；**一请求一连接**（50-64）、4 字节长度前缀+UTF-8（120-129）、闲置 20 min 退出（36-40）、**关掉最后一份文档即退出**（107-116） | `WOC_DAEMON_IDLE_MS`/`WOC_DAEMON_ONCE` 可调 |
| helper 共享件 | `Tray.cs`(534)、`WocLocks.cs`(53) | 托盘图标常驻进程（`--tray`，状态文件即 IPC）；跨 app 互斥锁 `Local\woc-clipboard`/`Local\woc-state` | Tray.cs:40/65；WocLocks.cs:18 |
| helper 主体 | `helper\word\*.cs`(14 个, 9 752)、`helper\excel\EtHelper.cs`(2 784)、`helper\ppt\*.cs`(7 个, 6 932) | C#5、无第三方依赖、`[STAThread] Main`；`Json`（手写解析/序列化，WpsHelper.cs:35-228）、`Bag`（大小写不敏感的请求取值器，230-284）、`Dispatch` 命令表、每命令一个 `static object Fn(Bag a)` | word Dispatch **453-495**；excel **396-428**；ppt **445-489** |
| 多任务编排 | `scripts\tasks.ps1`(569) / `tasks-excel.ps1`(402) / `tasks-ppt.ps1`(309) / `tasks-session.ps1`(212) | 纯 PowerShell、**零 COM**：Word `split/merge/plan/ranges/selftest`（28、553-560）、Excel/PPT `split/merge/plan/selftest`、跨 app 波次 `plan/check/selftest` | 产出 `merged.json`/`items-*.json`/`batch-ops.json` 再喂 `<app>.ps1 batch` |
| 发布 | `publish.ps1`(74) | 只复制分发包到 `<D>`，反向删除目标里多余文件 | 见 §5 |

**三端的差异只在"app 常量 + 文档对象"**：`KWPS.Application`（Word，回退 `Word.Application`）/ `KET.Application`（Excel，**创建实例只认 KET**）/ `KWPP.Application`（PPT）。三端各自复制了一份 `Json`/`Bag`/`App`/`FindOpen`/`RequireDoc` 等 plumbing（**没有共享"文档层"，只有 shared 传输层**）——这是接新功能时最需要注意的结构事实。

### 1.2 `woc.ps1 word read -Doc X.docx -Count 0 -SkipTableCells` 完整调用链（含行号）

| # | 位置 | 发生什么 |
|---|---|---|
| 1 | `scripts\woc.ps1:19` | `$app = $args[0]` = `word` |
| 2 | `scripts\woc.ps1:24-28` | `$target = <scripts>\word.ps1`，存在性检查 |
| 3 | `scripts\woc.ps1:29` | `$rest = @($args \| Select-Object -Skip 1)` |
| 4 | `scripts\woc.ps1:35` | `& powershell -Mta -NoProfile -ExecutionPolicy Bypass -File $target @rest`（MTA 省 ~2 s STA 初始化） |
| 5 | `word.ps1:9-122` | `param()` 绑定：`$Command='read'`、`$Doc`、`$From=1`、`$Count=0`、`$SkipTableCells=$true`、`$Slow`、`$TimeoutSec=180` |
| 6 | `word.ps1:139-146` | 收集 helper 源码 = `helper\word\*.cs`(14) + `helper\shared\*.cs`(3)，按名排序 |
| 7 | `word.ps1:213-224` | dot-source `woc-common.ps1` → `Get-CachedSourceHash -App word -Sources $sources`（`woc-common.ps1:16-42`，命中缓存 0 ms，未命中用 `Get-FileHash`） |
| 8 | `word.ps1:225-245` | 缓存目录 `%LOCALAPPDATA%\wps-office-cooperate\word-helper-<sha12>`；exe 不存在则 `csc.exe` 编译（refs：System/System.Core/System.Drawing/System.Windows.Forms/Microsoft.CSharp，`/codepage:65001`） |
| 9 | `word.ps1:255` | `$Command -eq 'batch'` 不成立 → 跳过 |
| 10 | `word.ps1:273-275` | `$request['cmd']='read'`；`$Doc` 转绝对路径写 `doc` |
| 11 | `word.ps1:306-313` | read 分支：`from`/`count`（0=扫到底）/`skipTableCells:true`/`only`/`slow` |
| 12 | `word.ps1:629-646` | `Test-ParamConsumption`（`param-check.ps1:22-51`）——本轮参数全部进 request，无告警 |
| 13 | `word.ps1:647` | `ConvertTo-Json -Depth 12 -Compress` 成一行 JSON |
| 14 | `word.ps1:651-658` | `Invoke-Daemon $json 900`（定义 157-190）：连管道 `woc-word-<sha12>`（名字来自 155），发 4 字节长度+JSON，读回响应 → 原样 `Out.Write` 并 `exit 0`（**热路径到此结束，无进程启动**） |
| 15 | — | 若管道不在（首次/`WOC_NO_DAEMON=1`/daemon 挂了）：`word.ps1:660-681` 一次性回退——`%TEMP%` 建目录、写 `request.json`（664）、`Start-Process WpsHelper.exe <req> <resp>`（668）、`WaitForExit($TimeoutSec*1000)`（669，**唯一受 -TimeoutSec 约束的路径**）、读回响应（678）、清理临时目录 |
| 16 | `WpsHelper.cs:300-322` `Main` | 命中 `args[0]=="--serve"` → `PipeServer.Run(pipe, HandleRequest, OpenDocumentCount)`（308-309）；一次性模式则 `HandleRequest(File.ReadAllText(req))`（318）+ 写响应（319） |
| 17 | `PipeServer.cs:47-117` | `BeginWaitForConnection`（56）→ `ReadMessage`（120-129）→ `handler(request)`（86）→ 回写长度+JSON（89-96）→ 无文档即退出（107-116） |
| 18 | `WpsHelper.cs:340-419` `HandleRequest` | 清 `_reopened`（342）→ `Json.Parse`（346）→ `cmd` 小写（347）→ `doc`（349）→ `IsWriteCommand`（353，read **非**写命令）→ `Dispatch(cmd, bag)`（358）→ `Ok(cmd, dispatched)`（359, 431-438） |
| 19 | `WpsHelper.cs:453-460` `Dispatch` | `case "read": return ReadDoc(a);`（**460**） |
| 20 | `WpsHelper.cs:993-1156` `ReadDoc` | `RequireDoc(a, true)`（995）→ `Paragraphs.Count`（996）→ from/count（997-998）→ 过滤参数与正则（999-1006）→ `filtered = skipCells \|\| onlyRe != null`（1010）→ `TableSpans(doc)`（1014，表区间 947-961）→ **快路径门**（1033-1080：`total>=200 && filtered` 时先 `FastBlocksVerified`（`FastText.cs:344/353`，块模型+抽样校验），失败再 `XmlDocTexts`（`XmlText.cs:39`，一次 `Content.WordOpenXML`+本地解析），再失败整体回退慢路径）→ 主循环（1084-1107，快 `ReadOneParaFast` 1162-1177 / 慢 `ReadOnePara` 973-991）→ 表格清单（1110-1132）→ 结果字典（1134-1155：`name/readOnly/paragraphs/from/to/returned/scanned/filter/skippedTableCells/pages/revisions/comments/path/fastFail/items[]/tables[]`） |
| 21 | `WpsHelper.cs:418` | `Json.Write(response)`（序列化器 35-228，UTF-8 无 BOM） |
| 22 | 回到 `word.ps1:654` | 打印响应；`"ok":false` 时 `exit 2`，否则 `exit 0` |

> `RequireDoc`（`WpsHelper.cs:719-730`）：`FindOpen`（698-717，遍历 `Documents.Item(k).FullName` 小写比对）→ 找不到且允许创建则 `_reopened=true` + `OpenPath`（735-754，静默重开不弹窗），`HandleRequest` 事后写成 `reopened`/`unsavedAtClose`（368-392）。
> 写命令额外走 `State.cs` 的"人改过没有"指纹：`TakeSignature`（69-89）→ 执行 → `AfterWrite`（136-166）把 `changedOutside:true` 塞进结果。

---

## 2. 命令全景矩阵

图例：**包装参数**只列要点；**请求字段**是**发出去的 JSON 键名**（helper 侧大小写不敏感）；**实现**给 `文件:行`（helper 短名均在 `scripts\helper\<app>\`）；**返回**只列顶层主要字段。所有命令都额外支持 `-TimeoutSec`（**仅一次性回退路径生效**，见 §2.4-C）。

### 2.1 Word（39，`scripts\word.ps1` switch 295-627 / helper `WpsHelper.cs` Dispatch 453-495）

| 命令 | 包装参数 | 请求字段 | helper 实现 | 主要返回 | 语义 |
|---|---|---|---|---|---|
| `doctor` | — | `cmd` | WpsHelper.cs:797 | `apps{progId,installed,running}`、`processes`、`app`、`openDocuments`、`undoRecord`、`renderOk`、`encryptedAtRest` | 自检（套件/版本/渲染/进程） |
| `status` | — | `cmd` | WpsHelper.cs:887 | `running`、`processes`、`documents[{name,path,readOnly,saved,active,caption,hinted,paragraphs}]`、`app` | 列实例与文档 |
| `open` | `-Doc [-Show][-ReadOnly][-NoHint]` | `doc,show,readOnly,noHint` | WpsHelper.cs:898 | `name,path,readOnly,saved,paragraphs,tables,pages,revisions,comments`+`how`+`hint`+`multiInstanceWarning` | 接管/打开（默认显示窗口） |
| `read` | `-Doc [-From][-Count][-SkipTableCells][-Only][-Slow]` | `doc,from,count,skipTableCells,only,slow` | WpsHelper.cs:993 | `name,readOnly,paragraphs,from,to,returned,scanned,filter,skippedTableCells,skippedByOnly,pages,revisions,comments,path,fastFail,items[{i,text,style,list,inTable,table}],tables[{table,rows,columns,header}]` | 读段落（fast/xml/slow 三路径） |
| `search` | `-Doc -Text [-Style][-Max][-Case]` | `text,style,max,caseSensitive` | Search.cs:16（快路径 FastText.cs:164） | `hits[{para,style,inTable,at,context}]`、`totalHits`、`truncated`、`verified`、`atFallbacks`、`cappedParas` | 全文找词（Find+字符区间校验） |
| `selection` | `-Doc` | `doc` | Search.cs:124 | `text,length,paras,fromPara,toPara,style,inTable,table,row,col,images` | 读用户选区 |
| `undo` | `-Doc [-Count]` | `count` | Search.cs:323 | `requested,undone,revisions` | 撤自己最近 N 步 |
| `revise` | `-Doc -Items` | `items[]` | WpsHelper.cs:1185 | `applied`、`items[{para,op,at,hits,chars,deleted,textAfter,item}]`、`revisionsBefore/After/Added` | 修订模式批量改写（原子） |
| `comment` | `-Action list`｜`-Items`｜`-Delete -Index/-All`｜`-Para [-Anchor] -Text/-TextFile` | `action,items,delete,all,index,para,anchor,text` | WpsHelper.cs:1399 | `comments[{i,author,scope,text,para,anchor}]`、`items`、`added`、`deleted` | 批注增删列 |
| `render` | `-Doc -Page [-To] -Out [-Width]` | `page,to,out,width` | WpsHelper.cs:1556 | `page,path,width,pages,files,emfBytes` | 页 EMF→PNG（helper 写盘） |
| `font` | `-Doc -Name [-Track]` | `name,track` | WpsHelper.cs:1619 | `font,ranges,readBack,readBackLatin,paragraphs,tracked` | 统一字体（正文/页眉页脚/文本框） |
| `format` | `-Doc ([-Para]｜[-From -To]｜[-Style]｜[-Match]) +格式项 [-Slow]` | `para,from,to,expect,style,match`+`formatFields`(278-284) | Layout.cs:12 | `target,paragraphs,applied,samples,readBack,tracked` | 套字符/段落格式 |
| `style` | `-Doc -Style +格式项 [-NoNumbering]` | `style`+fmt、`numbering:false` | Layout.cs:232 | `style,applied,font,size` | 改样式本身 |
| `page` | `-Doc [+Paper][-Orientation][-Top…-Right][-HeaderText][-FooterText][-PageNumber][-PageNumberStart][-All]` | `paper,orientation,top,bottom,left,right,header,footer,pageNumber,pageNumberStart,all` | Layout.cs:330（查询模式 338-388） | `sections[{section,paperSizeId,orientation,marginsCm,headerText,footerText,pageNumbers{startNumberRaw,restartAtSectionRaw,firstDisplayedPage}}]`、`applied` | 纸张/边距/页眉页脚/页码（逐节） |
| `table` | 插入 `-Para -Data [-Borders][-Widths][-HeaderRow][-Caption]`；改表 `-Table [+Borders][-Autofit][-RowHeight][-KeepTogether]+fmt`；`-Action scan\|header-repeat\|caption\|rows` | `para,table,action,passes,row,force,data,borders,widths,header,caption,autofit,rowHeight,keepTogether`+fmt | Tables.cs:182（TabInsert 202 / TabFormat 397；scan Tables2.cs:30、header-repeat 122、caption-move 237、rows 366） | 插入:`table,rows,columns,widths,header,caption,applied,tables`；改表:`borderRead,applied,tables,count`；scan:每表行含 `para,split,headerRepeated,captionInFirstRow,thickBorders` | 建表/改表/体检/表头跨页/题名出框/删行 |
| `captions` | `-Doc [-Label][-Max]` | `label,max` | Captions.cs:235 | `tables[{table,para,rows,columns,caption,pos,label,number,format,gap,title}]`、`summary{duplicates,sequenceGaps,formats,spaces,positions,uncaptioned,unresolvedReferences}`、`issues[{rule,para,sample,suggest}]` | 题注清单+编号体检（只读，10521 段 6.4 s） |
| `caption` | `-Para -Label -Title [-Above][-Align]` | `para,label,title,above,align` | Tables.cs:989 | `para,label,title,text,field,number,anchorPara,fieldsRefreshed` | 生成 SEQ 域题注 |
| `picture` | 插 `-Path -Para [+-Width][-Caption][-Align]`；调 `-Image n [+-Width][-Height][-CropLeft…]` | `path,para,width,caption,align,action:"insert"`｜`image,height,cropLeft/Right/Top/Bottom,action:"adjust"` | Tables.cs:1183（PicInsert 1194 / PicAdjust 1348） | 插图:`action,image,width,height,naturalWidthPt,scale,clamped,warning,caption`；调整:`left,right,top,bottom,readBack,visible` | 插图/调尺寸裁剪 |
| `toc` | `-Doc [-Para][-Title][-Levels]` | `para,title,levels` | Toc.cs:82 | `toc,entries,tablesOfContents,atPara,title,levels,preview,tailRemoved` | 插/刷新目录域 |
| `reference` | `-Action list -Items [-AtPara][-Heading]`｜`-Action cite -Para -Numbers` | `action,para,items,numbers,heading` | Toc.cs:292（list 303 / cite 428） | 列表:`inserted,fields,heading,headingAt,first,last,hangingIndent`；引用:`atPos,superscript,paragraph` | 参考文献列表/上标引用 |
| `view` | `-Doc -Para` 或 `-Page` | `para`｜`page` | Toc.cs:549 | `scrolled,para,page,movedCaret,selected,focus,activated,viaScrollIntoView` | 滚到段/页（只动视图） |
| `fields` | `-Doc` | `doc` | Text.cs:13 | `toc,fieldsUpdated,tocUpdated,fieldsError,pages` | 刷新全部域 |
| `cell` | `-Doc (-Table -Row -Col` 或 `-Para) -Text +格式项` | `table,row,col`｜`para`、`text`+fmt | Text.cs:58 | `table,row,col,text,applied,readBack,tracked` | 改单个单元格 |
| `insert` | `-Doc [-Para][-Where after\|before\|end][-Items\|-Text][-Track]` | `items`｜`text`、`para,where,expect,track` | Text.cs:155 | `inserted,inTable,from,to,where,paragraphs,revisions,tracked` | 批量插段（默认不进修订） |
| `revisions` | `-Doc [-Action list\|accept\|reject][-Index\|-All]` | `action,all,index` | Review.cs:36（list 142） | list:`total,byType,authors,items[{i,type,kind,author,date,para,text}]`；裁决:`done,revisionsBefore/After` | 列/接受/拒绝修订 |
| `snapshot` | `-Doc [-Tag][-Out][-Slow]` | `tag,out,slow` | Review.cs:292 | `doc,tag,at,paragraphs,page,revisions,texts,path,digest,textsPath,chunkStats,fastFail` | 段落文本快照（默认 %LOCALAPPDATA%） |
| `diff` | `-Doc [-Since][-Slow]` | `since,slow` | Review.cs:495 | `changed,added,removed,same,paragraphs,from,summary,textsPath,degraded,fastFail` | 与快照比（人改的也算） |
| `textdiag` | `-Doc [-Slow]` | `slow` | FastText.cs:558 | `paragraphs,revisions,tables,tableRows,contentChars,blocks,blockDelta,blockSpotMatches,xmlFetchMs,xmlChars,xmlParagraphs,xmlPlusRows,xmlPlusRowsMatches,xmlParseMs,fastFail` | 读侧慢/回退诊断（只读） |
| `handoff` | `-Doc [-Note][-As]` | `note,out` | State.cs:199 | `journalEntries,snapshotPath,revisions,comments,name,path` | 生成 `.handoff.md` 交接单 |
| `audit` | `-Doc [-Max][-Action 规则名][-Slow]` | `max,rules,slow` | Audit.cs:95 | `findings[{para,rule,sample,suggest}]`、`count,byRule,truncated,path,ms,fastFail` | 机械检查 |
| `unit` | `-Doc [-Apply][-Normalize][-Units][-Track]` | `apply,normalize,units,track` | Units.cs:149 | `textHits,unicodeOk,units,fixed,fixedPerUnit,skipped,fragments,looseFragments,suspects,scriptPages,fixedOnScriptPage` | 单位上标批量修（三类只报不改） |
| `units` | 同 `unit` | 同 `unit` | 同一函数 | 同上 | 别名（ValidateSet 双列） |
| `hint` | `-Doc [-Text][-HintTag][-Bar][-Toast][-Sound][-Quiet][-Clear][-Tray]` | `tag,bar,text,clear,toast,sound,quiet,tray,wrapper` | WpsHelper.cs:1914 | `caption,captionSet,statusBar,statusBarText,statusBarNote,tag,cleared,toast,sound,tray,trayStopped` | 接管提示 |
| `finish` | `[-Note][-Toast][-Sound][-Quiet]` | `note,toast,sound,quiet` | WpsHelper.cs:1992 | `statusText,toast,sound,hintCleared,note` | 完成信号（不动文档） |
| `save` | `-Doc [-As]` | `as` | WpsHelper.cs:1701 | `name,path,saved,revisions` | 保存/另存（**清撤销栈**） |
| `close` | `-Doc [-Save]` | `save` | WpsHelper.cs:1742 | `closed,saved,documentsLeft,processes` | 关文档；无文档时退实例 |
| `quit` | — | `cmd` | WpsHelper.cs:2029 | `processesBefore,processesAfter` | 全关退实例（慎用） |
| `batch` | `-Ops ops.json` | `ops[]`（**helper 级请求**，每条形如 `{"cmd":"…",<请求字段>}`） | 包装层 word.ps1:255-271 → woc-common.ps1:86-138 | `count,totalMs,ops[{i,ms,ok,result\|error}]` | 一次进程跑一批（5 条 1982→891 ms） |
| `serve` | `-Action start\|status\|stop`（默认 start） | —（包装层本地处理，不进 helper） | word.ps1:602-622 + `Start-Daemon` 192-206 | `action,started` / `running,pipe` | 常驻 helper 控制（`open` 自动预热） |

### 2.2 Excel（28，`scripts\excel.ps1` switch 267-500 / helper `EtHelper.cs` Dispatch 396-428）

| 命令 | 包装参数 | 请求字段 | helper 实现 | 主要返回 | 语义 |
|---|---|---|---|---|---|
| `doctor` | — | `cmd` | EtHelper.cs:730 | `apps{progId,installed,running}`、`processes`、`render`、`app`、`openWorkbooks`、`undoNote` | 自检 |
| `status` | — | `cmd` | EtHelper.cs:816 | `running,processes,workbooks[{name,path,readOnly,saved,active,caption,hinted}],app{version,build,path}` | 列实例/工作簿 |
| `open` | `-Doc [-Show][-ReadOnly][-NoHint]` | `doc,show,readOnly,noHint` | EtHelper.cs:827 | `name,path,readOnly,saved,sheets[{index,name,visible,active,usedRange}]`+`how`+`hint`+`multiInstanceWarning` | 接管/打开 |
| `new` | `-Doc [-Show][-Overwrite][-NoHint]` | `doc,show,overwrite,noHint` | EtHelper.cs:883 | `how,note,hint`（`EXISTS`/`IN_USE` 为错误码） | 新建空白工作簿并落盘 |
| `read` | `-Doc [(-Sheet)\|(-Range "S!A1:F40")] [-Profile]` | `sheet,range,profile` | EtHelper.cs:920 | 概览:`book,sheets,meta,path,v,f`；表:`sheet,used,v,f`；`>60 行或 >15 列`（阈值 EtHelper.cs:954）自动改列概况 `columns[{col,numbers,min,median,max,topValues,distinct}],sampleHead,sampleTail` | 读工作簿/表/区域 |
| `search` | `-Doc -Text [-Sheet][-Max][-Case]` | `text,sheet,max,case` | EtHelper.cs:1071 | `hits[{sheet,cell,value,formula,text}]`、`count,truncated` | 找格（值+公式） |
| `selection` | `-Doc` | `doc` | EtHelper.cs:1130 | `selectionType,sheet,range,rows,cols,cells,values` | 读用户选中区域 |
| `write` | `-Doc -Items`（或单格 `-Cell/-Value/-Formula`） | `items[{sheet,cell\|range,value\|formula\|values,expect}]` | EtHelper.cs:1454 | `items[{sheet,cell,before,after,formula,applied}]`、`count,preSnapshot,undoNote` | 批量写（原子批/锚校验/回滚） |
| `fill` | `-Doc -Range -Formula [-Sheet]` | `range,formula,sheet` | EtHelper.cs:1631 | `sheet,range,rows,formula,preSnapshot,topAfter,bottomAfter` | 整列填公式（原生 FillDown） |
| `rows` | `-Doc [-Sheet] -Action insert\|delete -Index [-Count]` | `action,index,count,sheet` | EtHelper.cs:1189(true) | `sheet,action,what,at,count,note` | 插/删行 |
| `cols` | 同上 | 同上 | EtHelper.cs:1189(false) | 同上 | 插/删列 |
| `sheet` | `-Doc -Action list\|add\|rename\|delete\|activate [-Name][-NewName]` | `action,name,newName` | EtHelper.cs:1669 | `sheets[{index,name,visible,active,usedRange}]`（按 action） | 工作表增删改名/激活 |
| `comment` | `-Action list`｜`-Cell [-Text\|-TextFile]`｜`-Action delete (-All\|-Index)` | `action,cell,sheet,text,all,index` | EtHelper.cs:1710 | `comments[{sheet,cell,author,text}]` | 单元格批注 |
| `render` | `-Doc [-Sheet][-Range] -Out [-Width] [-IncludeCharts]` | `sheet,range,out,width,includeCharts` | EtHelper.cs:1807 | `files[{index,name,title,left,top}]`、`chartsIncluded,charts,width,sheet,range,requestedRange,png` | 区域 CopyPicture→PNG（可选扩到全部图表） |
| `chart` | `-Action add -Range [-Type][-Title][-XTitle][-YTitle][-Anchor][-Width][-Height][-Sheet]`；`list [-All]`；`delete\|move -Index` | `action,sheet,range,type,title,anchor,width,height,xtitle,ytitle,all,index` | EtHelper.cs:1268 | `index,name,type,title,left,top,width,height,seriesCount,seriesNames,sheets` | 建/列/删/移原生图表 |
| `autofit` | `-Doc [-Sheet][-Range]` | `sheet,range` | EtHelper.cs:1930 | —（未逐行核实） | 列宽行高自适应 |
| `format` | `-Doc -Range（`;` 多区域）[-Sheet] +格式项`；`-Items` 每块不同格式 | `range\|items,sheet`+`formatFields`+`bold/italic/underline/wrap:false` | EtHelper.cs:2028 | `sheet,ranges[{range,applied}]`、`count` | 批量排版（颜色已修 BGR） |
| `snapshot` | `-Doc [-Sheet 名1,名2][-Out]` | `sheets,out` | EtHelper.cs:2173 | `taken,path,sheets,cells` | 值+公式快照 |
| `diff` | `-Doc [-Since]` | `since` | EtHelper.cs:2228 | `since,changes[{sheet,cell,before,after,formula}],totalChanges,truncated` | 与快照比 |
| `restore` | `-Doc [-Since][-Cells][-KeepExtras]` | `since,cells,keepExtras` | EtHelper.cs:2293 | `restored,since,cleared,selective,skippedMissingSheet,extraSheets,note` | 一键/选择性回滚 |
| `calc` | `-Doc` | `doc` | EtHelper.cs:2454 | `calculated:true` | 全簿重算 |
| `save` | `-Doc [-As]` | `as` | EtHelper.cs:2462 | `name,path,saved`（`.pdf`=整簿导出） | 保存/另存 |
| `close` | `-Doc [-Save]` | `save` | EtHelper.cs:2512 | `closed,workbooksLeft,processes` | 关工作簿 |
| `quit` | — | `cmd` | EtHelper.cs:2777 | `processesBefore,processesAfter` | 全关退实例 |
| `batch` | `-Ops ops.json` | `ops[]` | excel.ps1:223-236 | 同 Word | 批量 |
| `finish` | `[-Note][-Cell][-Sheet][-Toast][-Sound][-Quiet][-Force]` | `note,cell,sheet,toast,sound,quiet,force` | EtHelper.cs:2726 | `statusLine,statusCell,toast,sound,hintCleared` | 完成信号（可选写一行状态格） |
| `hint` | `-Doc [-Text][-HintTag][-Bar][-Toast][-Sound][-Quiet][-Clear][-Tray]` | `tag,bar,text,clear,toast,sound,quiet,tray,wrapper` | EtHelper.cs:2674 | `caption,captionSet,statusBar,statusBarText,tray,trayStopped,cleared` | 接管提示（表格状态栏真的显示） |
| `serve` | `-Action start\|status\|stop` | —（本地） | excel.ps1:472-500 | 同 Word | 常驻 helper 控制 |

> Excel 无 `changedOutside` 机制；"文档被用户关掉"用**attachment marker**（`%LOCALAPPDATA%\wps-office-cooperate\excel-attached.json`）判 `DOC_CLOSED`——`RequireBook` 在 `EtHelper.cs:498-516`，marker 读写 522-579（带 `WocLocks.State` 原子写）。

### 2.3 PPT（41，`scripts\ppt.ps1` switch 300-684 / helper `WppHelper.cs` Dispatch 445-489 + 5 个分部文件）

| 命令 | 包装参数 | 请求字段 | helper 实现 | 主要返回 | 语义 |
|---|---|---|---|---|---|
| `doctor` | — | `cmd` | WppHelper.cs:963 | `wpsPpt,msPpt,apps,running,processes,openPresentations,renderNote` | 自检 |
| `status` | — | `cmd` | WppHelper.cs:992 | `running,processes,presentations[{name,path,saved,readOnly,slides,active}],app` | 列实例/文稿 |
| `open` | `-Doc [-Show][-ReadOnly][-NoHint]` | `doc,show,readOnly,noHint` | WppHelper.cs:1026 | `name,path,readOnly,saved,slides,slideSize,how,note`（路径不存在=新建，`how:"created"`） | 打开/新建并接管 |
| `read` | `-Doc [-From][-Count]` | `from,count` | WppHelper.cs:1065 | `i,layout,shapes[{i,name,type,ph,l,t,w,h,text}],takeoverBadge,notes,name,slides,from,to,items` | 逐页形状清单 |
| `search` | `-Doc -Text [-Max][-Case]` | `text,max,case` | WppHelper.cs:1116 | `hits[{slide,shape,text,context}]`、`count` | 跨页找文字 |
| `replace` | `-Doc -Find [-Replace][-Scope][-Slides][-Regex][-Case][-Max][-DryRun]`；`-Items` | `find,replace,items,scope,slides,regex,case,max,dryRun` | Replace.cs:116 | `scanned,changed,shapesTouched,hits[{slide,where,shape,path,start,before,after,context}],skipped[{reason}],preSnapshot,rules,capped` | 全稿查找替换（回读校验+快照） |
| `selection` | `-Doc` | `doc` | WppHelper.cs:1159 | `slide,selectedShapes,note` | 用户当前页/选中形状 |
| `slide` | `-Action add [-Preset][-Title][-Subtitle][-Number][-Items][-Layout][-At]`；`dup\|move\|delete -Index` | `action,preset,title,subtitle,number,items,layout,at,index,to` | WppHelper.cs:1188 | `added,slides,preset,presetNotes` | 预设版式建页/复制移动删除页 |
| `split` | `-Doc -Chunks "1-20,21-45" -Out [-Prefix][-DryRun]` | `chunks,out,prefix,dryRun` | Chapters.cs:88 | `source,slides,chunks,files,failed,dryRun` | 拆分册（整份复制再删页） |
| `merge` | `-Doc -Sources "a,b"` 或 `-Items` `[-After][-DryRun]` | `sources`｜`sourcesList`/`items`、`after,dryRun` | Chapters.cs:239 | `target,slidesBefore,slidesNow,appended,failed,dryRun` | 合并多稿（Copy+Paste） |
| `text` | 单条 `-Slide (-Shape\|-ShapeIdx) (-Text\|-TextFile) [-Expect]`；批量 `-Items` | `items[{slide,shape,expect,text}]` | WppHelper.cs:1401 | `slide,shape,expect,applied,count,note` | 改形状文字（锚校验，整批原子） |
| `textbox` | `-Slide (-Text\|-TextFile) [-Name][-L][-T][-W][-H] +格式项` | `slide,text,textFile,name,l,t,w,h`+fmt | WppHelper.cs:1519 | —（未逐行核实） | 加文本框 |
| `image` | `-Slide -Path [-L][-T][-W][-H]` | `slide,path,l,t,w,h` | WppHelper.cs:1541 | `added,slide,size,note` | 插图（中文路径先拷 ASCII 临时目录） |
| `picture` | `-Action list [-Slides]`；`replace -New [-Match\|-Ids][-DryRun][-FitInside][-KeepCrop]` | `action,new,match,ids,slides,dryRun,keepSize,resetCrop` | Deliver.cs:444 | list:`pictures[{slide,shape,before,linked,how,l,t,w,h}]`；replace:`changed,wouldChange,hits,skipped[{reason}]` | 全稿换图/换 logo |
| `comment` | `-Action add -Slide [-Shape] (-Text\|-TextFile) [-Author]`；`-Items`；`list`；`delete -Slide (-Index\|-All)` | `action,items,slide,shape,text,textFile,author,index,all,clear` | Deliver.cs:120 | add:`added,count,author,textFile`；list:`comments[{slide,index,shape,text,author}],total` | PPT 批注（锚点写进文字） |
| `export` | `-Format pdf\|png\|jpg\|md [-Slides][-Out][-Width][-Quality]` | `format,out,slides,width,quality` | Deliver.cs:308 | `format,out,slides,files,count,failed,encrypted,note` | 交付导出（PDF 落盘即密文） |
| `table` | `-Slide -Data [-L][-T][-W][-H][-Size]` | `slide,data,l,t,w,h,headerRow,size` | WppHelper.cs:1575 | —（未逐行核实） | 建表填内容 |
| `chart` | `-Action add -Slide -Data [-Type][-Title][-XTitle][-YTitle][-Legend][-L][-T][-W][-H]`；`list` | `action,slide,data,type,title,xtitle,ytitle,legend,l,t,w,h` | Chart.cs:107 | add:`slide,shape,chartType,series,categories,firstSeriesValues,rangeNote,warning`；list:`charts[{name,chartType,series,l,t,w,h,slide}]` | 原生可编辑图表（需窗口） |
| `shape` | `-Action add -Type rect\|roundrect\|ellipse [-L][-T][-W][-H][-Fill][-Line][-Text]`；`delete -Name` | `slide,action,type,name,l,t,w,h,line,text`+fmt | WppHelper.cs:1627 | —（未逐行核实） | 画色块/删形状 |
| `format` | `-Slide -Shape +格式项` | `slide,shape,shapeIdx`+fmt | WppHelper.cs:1678 | —（未逐行核实） | 改文字格式 |
| `font` | `-Doc [-Font][-From][-To]` | `font,from,to` | WppHelper.cs:1711 | `font,from,to,shapes,groups,tables,failed,samples` | 全稿统一字体（含组合/表格） |
| `theme` | `-Set 主题.json`/`-Apply`/`-Reset`/无=查看 | `set,apply,reset` | WppHelper.cs:2053（应用 2092） | `set,theme,applied`（`slides,restyledShapes,footers`） | 整稿风格打底/套用 |
| `master` | `-Action list`；`footer [-Text][-PageNum][-Date][-SkipCover][-Slides][-Clear]`；`layout -Layout 1..11 [-Slides][-DryRun]`；`rebase -Path -Out [-Layout][-KeepTemplatePages][-Slides]` | `action,slides,text,textFile,pageNum,date,skipCover,clear,layout,path,out,keepTemplatePages,dryRun` | Master.cs:36 | list:`designs,layouts,slides[{slide,layout,footer,pageNumber}]`；footer:`changed,footers,pageNumbers,dates,details`；layout:`layoutBefore/After,shapesBefore/After,warnings`；rebase:`pasted,templatePagesRemoved,sizeBefore/After,warning` | 版式/页脚/套外来模板 |
| `audit` | `-Doc [-From][-To][-MinFont]` | `from,to,minFont` | WppHelper.cs:2256 | `slides,checked,issues[{slide,shape,type,detail}],count` | 机械体检（老稿误报多） |
| `check` | `-Doc [-Profile][-Rules][-MinSeverity][-Config][-Out]`（包装层另有 `-ExitCode`） | `profile,rules,minSeverity,config,out` | Check.cs:82 | `summary{blocker,major,minor,info,total}`、`issues[{slide,shape,rule,severity,detail,fix}]`、`suppressed,suppressedBy,fonts,limits,verdict` | 交付前门禁（11 条规则分级） |
| `outline` | `-Doc [-Out]` | `out` | Check.cs:766 | `slides[{slide,title,bullets,notes}],count,out` | 导出大纲 |
| `align` | `-Slide -Mode … [-Shapes]` | `slide,mode,shapes` | WppHelper.cs:2374 | —（未逐行核实） | 对齐/分布/统一尺寸 |
| `zorder` | `-Slide (-Name\|-ShapeIdx) [-Where front\|back]` | `slide,name,shapeIdx,where` | WppHelper.cs:2489 | —（未逐行核实） | 置顶/置底 |
| `notes` | `-Slide (-Text\|-TextFile\|-Clear)` | `slide,text,textFile,clear` | WppHelper.cs:1815 | —（未逐行核实） | 写/清演讲备注 |
| `view` | `-Slide N` | `slide` | WppHelper.cs:1831 | —（未逐行核实） | 切到该页（只动视图） |
| `render` | `-Doc [-From][-To] -Out [-Width]`；`-Overview` | `from,to,out,width,overview` | WppHelper.cs:1849（Overview 1896） | `slides,files`／`overview,slides,note`（**每次最多 4 页**，WppHelper.cs:1864 强制） | 逐页 EMF→PNG / 6 页拼图 |
| `snapshot` | `-Doc [-Tag]` | `tag` | WppHelper.cs:2644 | `at,slides,path,tag,textShapes` | 文本级快照 |
| `diff` | `-Doc [-Since]` | `since` | WppHelper.cs:2669 | `since,summary,changed` | 与快照比 |
| `restore` | `-Doc [-Since][-Shapes]` | `since,shapes` | WppHelper.cs:2507 | `since,restored,failed,missing,details` | 文本级回滚 |
| `save` | `-Doc [-As]` | `as` | WppHelper.cs:2731 | —（未逐行核实） | 保存/另存（`.pdf`=导出） |
| `close` | `-Doc [-Save]` | `save` | WppHelper.cs:2779 | `closed,saved,presentationsLeft,processes` | 关文稿 |
| `quit` | — | `cmd` | WppHelper.cs:3352 | —（未逐行核实） | 全关退实例 |
| `batch` | `-Ops ops.json` | `ops[]` | ppt.ps1:257-270 | 同 Word | 批量 |
| `finish` | `[-Note][-Toast][-Sound][-Quiet]` | `note,toast,sound,quiet` | WppHelper.cs:3281 | `toast,sound,trayStopped,note` | 完成信号 |
| `hint` | `-Doc [-Text][-SlideMark\|-NoSlideMark][-Toast][-Sound][-Quiet][-Clear][-Tray]` | `tag,bar,text,clear,toast,sound,quiet,tray,wrapper,slideMark` | WppHelper.cs:3206 | `captionSupported,titleMarked,slideMarked,slideMarkCount,slideMarkWarning,inAppHint,limitation,tray,cleared` | 接管提示（Win32 标题+按需逐页徽章） |
| `serve` | `-Action start\|status\|stop` | —（本地） | ppt.ps1:636-658 | 同 Word | 常驻 helper 控制 |

### 2.4 速查表 / 脚本实现 / 文档之间的偏差（写新功能前必读）

| # | 偏差 | 证据 | 影响 |
|---|---|---|---|
| A | **`batch`、`serve` 只在 Word 速查表出现**，Excel/PPT 速查表无这两行，但三端都实现了 | Excel `excel.ps1:223`、`472`；PPT `ppt.ps1:257`、`636`；SKILL.md Excel 表 311-337、PPT 表 495-535 | 读文档会误以为 Excel/PPT 没有 batch/daemon |
| B | **Word `-Tables` 是死参数**：`param()` 声明（`word.ps1:19`）但没有任何 switch 分支消费 | SKILL.md:28 自认"至今没有任何命令消费它"；param-check 会告警 | 传了等于没传（只有 stderr 警告） |
| C | **`-TimeoutSec` 只在一次性回退路径生效**，daemon 热路径完全没有超时 | `word.ps1:669`、`excel.ps1:538`、`ppt.ps1:723`；daemon 分支 `word.ps1:651-658`/`excel.ps1:522`/`ppt.ps1:706` 无 timeout | SKILL.md:209/339/537"所有命令都可用 `-TimeoutSec`"**不成立**（WPS 弹对话框时 daemon 侧会一直挂着） |
| D | SKILL.md 指向**不存在的章节**：`serve` 行说"设计见 reference/workflows.md §15" | SKILL.md:185；`reference/workflows.md` 只到 `## 14. 接管提示` | 找不到常驻设计文档 |
| E | SKILL.md「参考」段最后一条 bullet **把 5 个文件说明挤成一行** | SKILL.md:725-729 | 仅观感问题 |
| F | **验收记录滞后**：`word-acceptance.md` 只到 v1.12、`excel-acceptance.md` 只到 v1.17，SKILL.md 已 v1.28 | 两份 acceptance 末尾小节 | 最后 6 个版本没有验收记录模板可抄 |
| G | **excel-quirks 与 SKILL.md 互相矛盾**（v1.16 条未按 v1.17 修正）：quirks 说 `render -Range` 不含浮动图表且 helper 逐张合成；SKILL.md v1.17 说 `Range.CopyPicture` **会**带上相交图表、`-IncludeCharts` 改为"扩视图"且不再二次合成 | `excel-quirks.md:39-41` vs SKILL.md:30、325；以 `EtHelper.cs:1807-1930` 为准 | 想复用 render 合成逻辑的人会被旧文档误导 |
| H | excel-quirks"边界"说 `comment -Action delete` 只支持 `-All`，实际已支持 `-Index` | `excel-quirks.md:32` vs `excel.ps1:390-393`、`EtHelper.cs:1710` | 文档滞后 |
| I | SKILL.md Excel `comment` 行只写 `-Action delete -All`，未写 `-Index` | SKILL.md:321 | 同上 |
| J | 三端 wrapper 有轻微格式瑕疵（case 与上一条块同行） | `excel.ps1:465` `}    'calc' {` | 不影响功能，别照抄排版 |
| K | 未发现"速查表有、脚本没有"的命令；**反向**（脚本有、速查表没写）只有 `batch`/`serve` 与 `units` 别名 | ValidateSet 计数实测 39/28/41 | — |

---

## 3. 新增一条命令：checklist 与最小可抄示例

### 3.1 checklist（按改动顺序，含行号锚点）

**A. 包装层 `scripts\<app>.ps1`**

1. `ValidateSet` 加命令名：word `11-13`、excel `11-12`、ppt `11-12`（不加则 PS 报参数校验失败）。
2. 需要新参数时在 `param()` 加：word `16-121`、excel `15-90`、ppt `15-125`。**命名约定：JSON 键 = 参数名首字母小写**（`$SkipTableCells`→`skipTableCells`），这样 `param-check` 的默认规则（`param-check.ps1:48`）自动认账。
3. `switch ($Command)` 里加 `'cmd' { … $request['k'] = … }`：word `295-627`、excel `267-500`、ppt `300-684`。**只把真正需要的参数写进 request**。
4. 参数名与键不机械对应时登记别名：word `$paramAliases` `632-636`、ppt `690-692`、excel `505-508`（跳过表 `$paramCommandSkips` excel `509-515`、默认值表 `516`）。
5. 故意"绑定了但不转发"的参数（命令行专用开关）加进 `$paramCommandSkips`：word `637-639`、ppt `693-696`、excel `509-515`。
6. **`batch` 不用改**：只要 helper 认这个 cmd，`{"cmd":"…"}` 就能经 batch 调用（`woc-common.ps1:86-138`）——也是"先跑通再补参数"的最快路径。

**B. helper `scripts\helper\<app>\*.cs`**

7. `Dispatch` 加 case：word `WpsHelper.cs:453-495`（默认分支 `494` 的报错串也顺手补上新名字）、excel `EtHelper.cs:396-428`、ppt `WppHelper.cs:445-489`。
8. 实现 `private static object Fn(Bag a)`：
   - 取参：`a.Need("doc")` / `a.Str(k, fb)` / `a.Int(k, fb)` / `a.Flag(k, fb)` / `a.List(k)`（`Bag` 定义 `WpsHelper.cs:230-284`，**键名大小写不敏感**，`Obj()` 用 OrdinalIgnoreCase，`WpsHelper.cs:76`）。
   - 取文档：word `RequireDoc`（`WpsHelper.cs:719-730`）、excel `RequireBook`（`EtHelper.cs:498-516`）、ppt `RequirePres`（`WppHelper.cs` 内，未逐行核实）。
   - 返回 `Dictionary<string,object>`（也可 List/标量），`Json.Write` 自动序列化；**不要**自己拼 JSON。
   - 错误抛 `new Fail("CODE","message")`（`WpsHelper.cs:25-31`）；COM 异常由 `HandleRequest` 统一映射成 `BUSY`/`COM_ERROR`（`399-412`）。
   - **C# 5 语法**（系统 `csc.exe` 4.x）：不能用字符串插值、`using var`、`?.`；源码**纯 ASCII**，中文写 `\uXXXX`（先例 `WpsHelper.cs:386-390`）。
9. 新文件**不需要注册**：wrapper 用 `Get-ChildItem helper\<app> -Filter *.cs` 全量收集（word `142`、excel `111`），共享件放 `helper\shared\`（word `145-146`）；编译参数见 word `239-240`（`/codepage:65001`、refs 含 System.Drawing/Windows.Forms）。
10. **写命令必做**：把命令名加进 word 的 `WRITE_COMMANDS`（`State.cs:18-21`），响应才会带 `changedOutside`（`AfterWrite` `136-166`）。Excel/PPT **没有**这套机制，靠 `preSnapshot` + `before/after`。
11. 涉及结构改动/批量时可复用：`WocLocks.State/Clipboard`（`shared\WocLocks.cs:18`）、attachment marker（excel `522-579`）、快照目录按**文档路径哈希**分目录（v1.20 起）。
12. 想进多任务编排：在对应 `tasks*.ps1` 里加映射（Word `tasks.ps1:553-560`；Excel `tasks-excel.ps1`；PPT `tasks-ppt.ps1`；跨 app `tasks-session.ps1`）。

**C. 文档（三处，缺一不可）**

13. `SKILL.md` 速查表加一行：Word `169-207`、Excel `311-337`、PPT `495-535`；复杂命令再加说明小节（可抄 `replace` 552-576 / `check` 657-675 的写法）。
14. `SKILL.md` 顶部版本段（`8-81`，最新在前，格式 `**版本 vX.YY（日期）**：…`）+ 故障排查表按需补行（Word `279-301` / Excel `456-469` / PPT `698-711`）。
15. `reference/<app>-acceptance.md` 追加验收记录；`README.md` 功能清单（可选，5-11 一带）。

**D. 测试**

16. 新断言脚本 `tests\<app>\check_<feature>.py`（**断言只放 python**，中文天然 UTF-8；`.ps1` 必须纯 ASCII，见 `word-quirks.md:24`）。
17. 需要真机驱动就加 `tests\<app>\run_<feature>.ps1`（抄 `run_replace.ps1` 结构：`-SkipRun`/`-Only`/`-PythonExe`，见 §4）。
18. 多任务/组合场景加到 `tests\multi-task\check_tasks.ps1`、`tests\excel\multi-task\`、`tests\ppt\multi-task\`。

**E. 发布**

19. 新 `.cs` 自动随包（`publish.ps1:45` 递归 glob）；**新增 `.ps1` 必须手工加进 `$includes`**（`publish.ps1:16-41`），否则部署目录没有它。
20. `.\publish.ps1` 后核对输出（当前 "published 52 files"）；首次调用自动重编（哈希变→新 exe+新管道名，旧 daemon 闲置自灭）。

### 3.2 最小可照抄示例：给 Excel 加一条只读命令 `usedrange`

**模板选择**：`excel calc` 是全仓库最短的端到端命令——包装层 3 行 + Dispatch 1 行 + 实现 7 行。照它就是最小改动量。

**改动 1/3：`scripts\excel.ps1`**

```powershell
# ① 第 11-12 行 ValidateSet 里加 'usedrange'
[ValidateSet('doctor', ..., 'calc', 'usedrange', 'save', ...)]

# ② switch 里（紧跟 'calc'，即 465-467 之后）加：
    'usedrange' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'usedrange needs -Doc' }
        if ($Sheet) { $request['sheet'] = $Sheet }
    }
```

**改动 2/3：`scripts\helper\excel\EtHelper.cs`**

```csharp
// ③ Dispatch（396-428）里，紧跟 case "calc"（420 行）加一行：
            case "usedrange": return UsedRange(a);

// ④ 文件末尾 "misc" 区、CalcBook（2454-2460）之后加实现：
    private static object UsedRange(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        string want = a.Str("sheet", null);
        dynamic ws = want != null ? wb.Worksheets.Item(want) : wb.Windows(1).ActiveSheet; // v1.17: ActiveSheet unreliable
        dynamic ur = ws.UsedRange;
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["sheet"] = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
        r["address"] = Convert.ToString(TryGet(() => (object)ur.Address(false, false), ""), CultureInfo.InvariantCulture);
        r["rows"] = TryGet(() => (object)(int)ur.Rows.Count, 0);
        r["cols"] = TryGet(() => (object)(int)ur.Columns.Count, 0);
        r["cells"] = TryGet(() => (object)(int)ur.Cells.Count, 0);
        return r;
    }
```

**改动 3/3：文档 + 测试**

```markdown
⑤ SKILL.md Excel 速查表（311-337 区）加一行：
| `usedrange` | `-Doc [-Sheet 名]` | 返回该表 usedRange 的地址/行列数/单元格数（只读） |
```
```python
# ⑥ tests/excel/check_usedrange.py（最小断言骨架，抄 check_hint.py 的调用方式）
import json, subprocess, os
EXCEL = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'scripts', 'excel.ps1')
DOC = os.path.join(os.environ['TEMP'], 'ur-probe.xlsx')
def run(*a):
    out = subprocess.run(['powershell', '-Mta', '-NoProfile', '-ExecutionPolicy', 'Bypass',
                          '-File', EXCEL] + list(a), capture_output=True, text=True, encoding='utf-8')
    return json.loads(out.stdout)
assert run('new', '-Doc', DOC, '-Overwrite', '-NoHint')['ok']
r = run('usedrange', '-Doc', DOC)['result']
assert set(('sheet', 'address', 'rows', 'cols', 'cells')) <= set(r), r
assert r['cells'] >= 1, r
run('close', '-Doc', DOC)
print('PASS usedrange', r)
```

**验证顺序（跑真机的正确姿势）**

```powershell
& <S>\scripts\woc.ps1 excel usedrange -Doc C:\tmp\t.xlsx      # 直呼
& <S>\scripts\woc.ps1 word  serve -Action status               # 加完新命令先确认 daemon 换了新管道（哈希变）
# batch 实验（不改包装层也能跑）：
#   ops.json = {"ops":[{"cmd":"usedrange","doc":"C:\\tmp\\t.xlsx"}]}
& <S>\scripts\woc.ps1 excel batch -Ops ops.json
python tests\excel\check_usedrange.py
.\publish.ps1
```

**若要加"写"命令的额外成本**：① `State.cs:18-21` 登记 → 才有 `changedOutside`；② 想清楚反悔手段（Word 修订 / Excel 快照 / PPT 文本快照）；③ 批量写自己挂起 `ScreenUpdating/Calculation`（先例 `WriteCells` `EtHelper.cs:1454`）；④ 返回里给 `before/after` 让人自证；⑤ 结构类命令（插删行/页）必须提示"先 snapshot"。

---

## 4. 测试体系地图

**共同约定**：`.ps1` 驱动只起 WPS、发命令、把每个响应 JSON 落盘；**所有断言都在 python**（独立第二实现，用 `python-pptx`/`openpyxl`/`win32com` 读结果）。`.ps1` 一律纯 ASCII（PS 5.1 按 ANSI 解析中文会炸）。测试**不随分发包**（`publish.ps1` 不含 `tests/`）。
**驱动开关（实测清单）**：`-SkipRun`（只重跑 python 断言）、`-Only full,restore`（PPT 场景子集）、`-SkipDecrypt`、`-DecryptEngine`、`-DecryptScript`、`-PythonExe`。**没有 `-Wps`/`-Office` 开关**。

| 测试 | 位置（行数） | 测什么 | 怎么跑 | 判据在哪 |
|---|---|---|---|---|
| Word 总验收 | `tests\word\run_acceptance.ps1`(100) + `validate.py`(500) | 语料生成→10 条 revise + 2 条批注→独立校验（accept-probe 副本上 `Revisions.AcceptAll` 比对） | `.\\run_acceptance.ps1 [-RevisedDoc …] [-SkipValidate]`；`python validate.py --doc … --targets … --edits …` | `validate.py`（`verdict` pass/fail，写 `validate-report.json`，exit 0/1/2）+ `reference\word-acceptance.md` |
| Word 题注回归 | `tests\word\captions\`（`make_fixture.py` 600 / `check_captions.py` 214 / `expected.json` / `README.md`） | 30 条断言：captions 清单与六类缺陷、`read` 的 `inTable` 边界与 `-SkipTableCells`、`scan.para`、`search` 段号校验与偏移 | `python check_captions.py [--verbose]`（跑 fixture 副本） | `check_captions.py` 断言 + `README.md` |
| 常驻 daemon 三端 | `tests\word\check_daemon.py`(228) | D1 serve 启停/管道命名；D2 **常驻 vs 一次性 JSON 逐字节相同**（剔 volatile 字段）；D3 真变快；D4 关最后文档自退；D5 daemon 停掉仍可用 | `python check_daemon.py [word\|excel\|ppt]`（40/40） | 脚本内 `check()` 计数 |
| 读侧快路径 | `tests\word\check_textpaths.py`(123) | P1 fast vs `-Slow` 的 digest/逐段文本一致（带表/带修订两份语料）；P2 快路径真生效；P3 `textdiag` 结构不变量 `w:p+表行==Paragraphs`；P4 `diff` 精确；P5 不可快时给 `fastFail` 原因 | `python check_textpaths.py [--keep]` | 脚本内断言 |
| 接管提示/托盘 | `tests\word\check_hint.py`(86)、`check_tray.py`(124)、`tests\excel\check_hint.py`、`check_tray.py`、`tests\ppt\check_hint.py`、`check_tray.py` | 标题标记/hinted/`-Clear`/`finish` 还原、托盘生命周期与状态文件、PPT 徽章 `-SlideMark` 策略与 Ctrl+S 自动修复（P6：裸 COM `Save()` 模拟用户存盘） | `python tests\<app>\check_hint.py` / `check_tray.py` | 脚本内断言 |
| PPT replace | `tests\ppt\run_replace.ps1`(285) + `check.py`(499) + `make_deck.py` | 12 个场景（dryrun/full/idem/restore/notes/master/slidesonly/slide1/regex/case/max/items/footer）：命中集合、逐字符格式不变、python-pptx 独立算期望 | `.\\run_replace.ps1 [-Only full,restore] [-SkipRun] [-SkipDecrypt]`（约 2–3 min） | `check.py`（唯一断言处） |
| PPT check/outline | `run_check.ps1`(106) + `check_check.py`(243) + `make_check_deck.py` | 每条规则在**指定页**触发、干净页**不误报**、`verdict`/`suppressed`/`-ExitCode`（`ppt.ps1:739` exit 3） | `.\\run_check.ps1 [-SkipRun]` | `check_check.py` |
| PPT comment/export/picture | `run_deliver.ps1`(148) + `check_deliver.py`(206) | 批注批量/回滚/锚点、PDF/PNG/MD 产物与密文标记、换图几何与 `skipped` 原因 | `.\\run_deliver.ps1` | `check_deliver.py` |
| PPT split/merge | `run_chapters.ps1`(133) + `check_chapters.py`(204) | 每页唯一标题 PAGE-n/A-1… ⇒ 丢页/多页/错序都查得出；母版标记验"设计是否跟拆" | `.\\run_chapters.ps1` | `check_chapters.py` |
| PPT master | `run_master.ps1`(126) + `check_master.py`(208) | `list`/`footer`/`layout`/`rebase`（模板带 `TPL-MARK` 母版标记） | `.\\run_master.ps1` | `check_master.py` |
| PPT chart | `run_chart.ps1`(107) + `check_chart.py`(147) | 图表数据是**真数据**（python-pptx 读图表缓存），不是图片 | `.\\run_chart.ps1` | `check_chart.py` |
| 多任务 Word | `tests\multi-task\check_tasks.ps1`(159) + `ranges.json` | split 覆盖无重叠/无缝隙；同段合并成一条；find 唯一；数字多重集 guard；plan 倒序 + batch-ops | `powershell -NoProfile -ExecutionPolicy Bypass -File tests\multi-task\check_tasks.ps1`（真 WPS，跑 `report-300.pristine.docx` 副本） | 脚本内 `Check()` |
| 多任务 Excel | `tests\excel\multi-task\check_tasks.ps1`(114) | 15 条：冲突仲裁、`expect` 缺省告警、批注每格一条、`diff` 对账 | 同上风格（`scripts\tasks-excel.ps1`） | 脚本内 `Check()` |
| 多任务 PPT | `tests\ppt\multi-task\check_tasks.ps1`(104) | 19 条：`(页,形状)` 合并、replace 重叠仲裁、全局→文本→替换→批注顺序 | 同上（`scripts\tasks-ppt.ps1`） | 脚本内 `Check()` |
| 语料/夹具 | `tests\word\corpus\{report-300.docx,.pristine,audit-fixture,edits.json,targets.json,meta.json}`、`tests\word\captions\captions-fixture.docx`、`tests\excel\work\*`、`tests\ppt\out*\`（多数 gitignore） | 固定语料让断言可复算 | `python make_corpus.py`、`make_audit_fixture.py`、`make_edits.py`、`make_deck.py` | — |

**跑测试的三条现实约束**：① 真机 WPS 必开，跑完查 wps/et/wpp 进程数（验收脚本自带 before/after 计数）；② 本机落盘即密文（`17DA5FA01633CD9A`），凡要 python 读 `save` 出来的文件都得先解密（`decrypt-office-pdf\decrypt-files.ps1` 或 FastDycrypt；PPT 驱动已封装，一行 ~9–45 s，故有 `-SkipDecrypt`）；③ `tests\ppt\local.config.json`（gitignore）存本机解密工具路径——中文路径不能写进 `.ps1`。

---

## 5. 部署与发布

| 项 | 事实（已用命令核实） |
|---|---|
| 发布脚本 | `<S>\publish.ps1`（74 行）：`param([string]$Target = "C:\Users\Administrator\.agents\skills\wps-office-cooperate")`（**字面量，不是 `$env:USERPROFILE`**，第 10 行） |
| 实际部署目录 | `C:\Users\Administrator\.agents\skills\wps-office-cooperate` —— **存在、真实目录（非 junction/symlink）**，52 个文件；顶层 `SKILL.md / README.md / reference / scripts / docs` |
| 是否同步 | **同步**：两侧 `SKILL.md` mtime 均 2026-11-02 19:58:14；部署侧 `scripts\helper` 下 25 个 `.cs` 与 dev 一致（word 14 / excel 1 / ppt 7 / shared 3） |
| 随包内容 | 白名单：`SKILL.md`、`README.md`、10 个 `reference\*.md`、12 个 `scripts\*.ps1`+`woc.cmd`、3 个 `docs\*.png`、`scripts\helper\**\*.cs`（25 个）——**2+10+12+3+25 = 52 个文件**，与部署目录实况吻合（`publish.ps1:16-45`） |
| 不随包 | `tests/`、`tools/`、`publish.ps1`、`DESIGN-word.md`、`DESIGN-parallel-tasks.md`、`.gitignore`、`tmp_axp/`、`docs/video-assets/`、**任何 `*.exe`**（helper 编译产物永不发布，首次调用现编 ~1 s） |
| 同步方式 | **逐文件 `Copy-Item -Force` + 反向清理**（`51-58` 复制、`63-69` 删除目标里不在白名单的文件并打印 `removed stale:`）——**不是 robocopy /MIR，没有哈希校验**；末尾打印 `published N files, K KB`（当前 52 files） |
| 安全阀 | target 为空或长度 <10 直接抛错（`48`）；只删 `$Target` 内文件 |
| git | `origin https://github.com/freeskychenjun/wps-office-cooperate.git`，分支 `master`，**0 个 tag**；版本号约定 = commit subject `vX.YY (app): 一句话`（如 `aac25a0 v1.28 (ppt): …`）+ `SKILL.md` 顶部版本段同步写 |
| `.gitignore` 关键项 | `tests/work/`、`tests/ppt/out*/`、`__pycache__/`、`*.exe`、`.tmp/`、`sandbox-*/`、`guide.html`、`tests/ppt/local.config.json`、`.dsh/`、`docs/video-assets/`、`tmp_axp/` |
| 运行时缓存（不在仓库） | `%LOCALAPPDATA%\wps-office-cooperate\`：`<app>-helper-<sha12>\`（编译产物，实测存量上百个）、`source-hash-<app>.json`、`word-state\`（按路径 16 位哈希的 journal）、`*-snapshots\`、`ppt-themes\`、`ppt-slide-mark\`、`excel-attached.json`/`ppt-attached.json`、`tray-*.json`、`ppt-title-hint.json` |

**发布后的生效路径**：部署目录里的 `<app>.ps1` 就地编译该目录的 `helper\*.cs` → 缓存目录按**源码哈希**区分，所以"改了源码没重编"不会串版本；反之**只改部署目录的 .ps1 不改 .cs**同样生效（脚本即逻辑）。helper 的 `--serve` 管道名同样派生自哈希 ⇒ 换版本后旧 daemon 自然失联自灭。

---

## 6. 已知限制台账（判断新功能可行性时直接查）

### 6.1 Word（WPS Writer 12.1.0.29161）

| # | 限制/实测坑 | 影响 | 出路（skill 内已有对策） |
|---|---|---|---|
| W1 | **每个表格行被算成一个空段落**：`Paragraphs.Count` = 正文 `w:p` + 表行数（7334 = 6779+555） | 段号≠正文段号，`Content.Text` 块模型在带表文档必然错位 | v1.13 起 `Content.WordOpenXML` 一次取全稿（`XmlText.cs:39`）；表内段用 `Table.Range.Paragraphs.Count` 权威标注（`WpsHelper.cs:1046-1080`） |
| W2 | 修订非零时 `Content.Text` 是"接受视图"，旧块模型必然回退 | 协作中读侧只剩慢路径 | XML 路径不取 `w:delText` 即与 `Range.Text` 对齐（`word-quirks.md:10`） |
| W3 | 内嵌图片在 `Range.Text` 里是 `/`；`w:sym`/单段多图段落易错 | 文本比对/段号会偏 | XML 侧 `w:drawing/w:pict/w:object`→`/`（`XmlText.cs:112-160`），风险段单独回读 |
| W4 | `Application.Name` 冒充 Word（`Name="Microsoft Word"`、`Version="12.0"`） | 用 Name 判套件会误判 | 看 `Path` 含 `Kingsoft` 或 `Build`（`word-quirks.md:12`） |
| W5 | 样式只认中文本地名/内建常量（`"Heading 1"` 报 E_FAIL） | 样式名写英文就失败 | 用 `Style.NameLocal` 读回的中文名；先 `read` 看 `style` |
| W6 | `Document.Range` 是**方法**不是属性 | `doc.Range.Information(...)` 必炸 | 一律 `doc.Range(a,b)` |
| W7 | 关掉最后一个文档 = **静默杀掉自动化实例** | 后续命令拿到空壳 | `close` 只在无文档时 `Quit` 并回报 `processes`；每次校验 `Documents.Count` |
| W8 | 每次 `New-Object`/`CreateInstance` 都是**新实例**，实例间互不可见 | 文档"开在另一个实例" | 先 `Marshal.GetActiveObject` 附着；`open` 给 `multiInstanceWarning`（`WpsHelper.cs:918-936`） |
| W9 | 没有 `Range.CopyAsPicture` | 无法截区域 | `render` 走 `Pages(n).EnhMetaFileBits`→PNG（`WpsHelper.cs:1556`） |
| W10 | 写入后视图跳回光标，无法"跟随滚动" | 用户看不到改动处 | 给段号+段首文字；核对用 `render` |
| W11 | 批注/修订作者固定为登录用户；`Comments.Add` 只 2 参 | AI 与用户改动分不开 | 署名写进批注正文 |
| W12 | 刚插入的文本会被下一次替换二次命中 | 链式替换出 `P--` | 一次算好写到位；同段一条 item |
| W13 | **PS 5.1 不能写 COM**（裸 `__ComObject`、`Range.Information` 调不到） | 只能在 C# 里做 COM | 架构本来就是"C# helper + PS 只转发"（`word-quirks.md:23`） |
| W14 | 含中文的 `.ps1` 会被按 ANSI 解析 | 乱码 + 离谱语法错 | `.ps1`/`.cs` 纯 ASCII，中文只走 JSON 文件（**本次勘察时亲测踩过一次**） |
| W15 | 批注锚点会让同段字符位置错 1 | `replace` 多删/少删一个字 | 已改读回校验 ±3 纠正（`word-quirks.md:25`；`Revise` `WpsHelper.cs:1185`） |
| W16 | `Table.Range.End` 是**排他端** | 紧接表格的正文段被误标 `inTable:true` | 判据 `start<=x<end`（`WpsHelper.cs:947-970`） |
| W17 | `Find` 的字符偏移→段号必须**按区间校验** | 长表格附近选错邻段且不触发兜底 | `Range.Start<=offset<Range.End` 裁决 + `verified:true` + `atFallbacks`（`FastText.cs:91-164`） |
| W18 | 题注判据：标签必须在**段首** | "如表3.3-2所示"被当题注；表内题名双报 | 共用题注解析器（`Captions.cs:235`） |
| W19 | Writer 的 `Application.StatusBar` **只写不读、界面不渲染** | 别承诺状态栏可见 | `hint` 返回 `statusBarNote`；进度用气泡/标题/托盘 |
| W20 | WPS 12 标签页**从右截断**（约 22 半角字符） | 后缀标记等于没打 | 标记放文件名**前面** |
| W21 | 用户明确否决**置顶悬浮窗** | 不要再提议/偷偷加回 | 常驻信息只靠"标签前缀+托盘"（SKILL.md:134） |
| W22 | `save` 会**清空撤销栈**；COM 格式改动的撤销不分作者 | 排版类改完很难反悔 | 铁律 6：用户看过/`render` 复核后才 save；文本优先走修订 |
| W23 | `-Tables` 参数没有任何命令消费 | 传了静默无效（仅告警） | 见 §2.4-B（要么实现要么删） |

### 6.2 Excel（WPS 表格）

| # | 限制/实测坑 | 影响 | 出路 |
|---|---|---|---|
| E1 | **一个 ET 实例同时开两个工作簿会挂起**（`Workbooks.Open` 不返回，随后幽灵簿/`E_FAIL`，只能杀 `et.exe`） | 多工作簿**不能并行** | 一个实例只开一个簿；多簿任务串行（SKILL.md:350） |
| E2 | `Application.Name` 冒充 Excel | 判套件误判 | 看 `Path`/`Build`；创建实例**只认 `KET.Application`**（`ET.Application` 未注册） |
| E3 | `SaveAs` 之后**工作簿换身份**（FullName 变新路径） | 按旧路径 `close` 报 `NOT_OPEN` | 另存后一律按新路径操作 |
| E4 | 写块数组必须 **0 基、尺寸精确**（按 1 基习惯造会静默不落位） | 数据"写了但没进去" | 走 `write -items.values`（`EtHelper.cs:1454`），别手写 COM |
| E5 | 多格 `Range.Text` 返回空（只有单格有显示文本） | before/after 拿不到 | 单格取；块写入后用 `read`/`render` 复核 |
| E6 | 渲染只能走剪贴板（`CopyPicture`→`Clipboard.GetImage`） | 并发渲染会串图 | `WocLocks.Clipboard`（`EtHelper.cs:1789`）跨 app 互斥 |
| E7 | **`Range.CopyPicture(xlScreen)` 会带上与区域相交的浮动图表**，区域外图表永远不在图里 | 想"只渲表格"会带上图 | v1.17 语义：`-IncludeCharts` 扩视图覆盖全部图表（`EtHelper.cs:1807-1930`）——**以代码为准，`excel-quirks.md:39-41` 已过时** |
| E8 | `chart` 数据**必须同表**；第 1 列=x、第 2..N 列各一系列、第 1 行系列名 | 跨表引用不可靠 | 辅助数据放同表右侧 |
| E9 | `format -Borders` 只认 `grid`/`thick`/`none`；对角线索引 5/6 v1.16 起恒清 | 其他值静默当 grid | 按枚举值传 |
| E10 | `-Color`/`-FillColor` 的 COM 低 24 位是 **BGR**（v0.1–v1.17 红蓝互换，v1.18 修） | 历史脚本颜色"反了" | 纯红=255、纯蓝=16711680 可读回核对 |
| E11 | 外部 COM 修改**不进撤销栈**，还会清掉用户的撤销历史 | 改坏只能靠快照 | `write/fill` 自动落 `preSnapshot`；`restore [-Cells]` 选择性回滚 |
| E12 | 关最后一份工作簿**静默杀实例** | 后续命令空壳 | `close` 回报 `workbooksLeft/processes` |
| E13 | `xlSheetVisible = -1`（不是 1） | 按 `==1` 判会把所有表读成隐藏 | `-1/1` 都算可见 |
| E14 | 单元素数组经 PS 函数返回会**被展开成标量** | helper 报 "items must be an array" | 函数里 `return ,@(...)` |
| E15 | `Workbook.ActiveWindow` 常为 null；`wb.ActiveSheet` 不可靠（同一分钟内会变） | 窗口标题/活动表定位错 | 统一 `wb.Windows(1).Caption` / `wb.Windows(1).ActiveSheet`（`excel-quirks.md:24,49`） |
| E16 | WPS 12 的窗口属于 **`wps.exe`**（`et.exe` 的 `MainWindowHandle=0`）；125% 缩放下 480 pt≈800 px | 截整窗会失败；图表被裁 | 别截屏，用 `render`；图表放内容下方或收小尺寸 |
| E17 | `SetSourceData` 对单列数据方向猜错（135 行=135 条系列） | 图全错 | 显式建系列；`seriesCount` 回读核对 |
| E18 | `Image.FromFile` 持文件锁 | 删/覆盖源 PNG 报 GDI+ 错误 | 先释放再动文件 |
| E19 | 落盘即密文（头部 `17DA5FA01633CD9A`） | python 读不回、发给同事打不开 | 见 `reference\encryption.md`；COM 路线不受影响 |
| E20 | `comment -Action delete` 现已支持 `-Index`（文档仍写只支持 `-All`） | 照旧文档会用错 | 以 `excel.ps1:390-393` 为准 |

### 6.3 PPT（WPS 演示）

| # | 限制/实测坑 | 影响 | 出路 |
|---|---|---|---|
| P1 | **无修订层、外部 COM 改动不进撤销栈** | 无 `undo` | 文本级 `snapshot`/`diff`/`restore`；`replace` 自动 `prereplace` 快照 |
| P2 | **WPS 演示不接受公式** | 表格/图表别想用公式 | 只写字面值 |
| P3 | `render` 走剪贴板（`Slide.Copy`→EMF→PNG，helper 写盘故**明文**）；`Slide.Export` 写的文件会被透明加密读不回来 | 想用 Export 会拿到读不回的密文 | 只用 `render`/`export -Format png` |
| P4 | **页码占位符渲染成重复 "1"** | 页脚页码全是 1 | 别用页码占位符；页脚写纯文字 |
| P5 | `open` 一个**不存在的路径 = 新建并立即保存** | 路径写错会留下垃圾稿 | v1.19 起返回 `how:"created"`+`note` |
| P6 | 手拼 JSON 的 Windows 路径里 `\U`/`\A` 是**非法转义**，旧解析器静默丢反斜杠 ⇒ 凭空多一份稿 | 数据污染 | v1.19 起非法转义直接 `BAD_ARGS`；请求一律走 JSON 序列化器 |
| P7 | 渲染 ≈355 ms/页、`check` 全规则 ≈970 ms/次、`read` 3 页 ≈632 ms | 逐页 render 会烧时间 | 设计期渲 `-Overview`，之后只渲改动页；`check` 收尾跑 `-MinSeverity blocker` |
| P8 | `chart` **必须有窗口**（`AddChart` 无窗口直接 E_FAIL/`NO_WINDOW`） | 后台建图必失败 | 先 `open -Show` |
| P9 | `chart` 只收**字面值**；`Series.Values` 写了不落盘；`SetSourceData` 被拒；`NewSeries` 不可用（>3 系列明确"未写入"）；点数靠增删工作表行对齐 | 复杂图表做不出来 | 数据经**内嵌工作簿**写；超 3 系列拆图 |
| P10 | `Slides.InsertFromFile` 在本机会**把目标稿清成 0 页** | 毁稿 | `merge` 用 `Slide.Copy()`+`Slides.Paste()` |
| P11 | 跨文稿粘贴**不带设计**（母版标记会丢） | 拆册外观走样 | `split` 用"整份复制再删区间外页" |
| P12 | 组合内的图片**只列不换**（子形状是组合局部坐标，COM 无 chOff/chExt） | 换 logo 会丢位置 | `picture replace` 进 `skipped`；要换先在 WPS 里移出组合 |
| P13 | 批注钉位置被 WPS 自己摆（传 100,100 读回 5,5） | 靠坐标对不上 | 锚点写进批注正文（`@形状名：…`） |
| P14 | `master -Action layout` **只认 `-Layout 1..11` 旧常量**；`CustomLayout` 对象静默无效；`ApplyTemplate` 返回成功但什么都不做 | 换版式"看起来成功"其实没变 | 用旧常量；换完 `render` 查掉没掉占位符（返 `warnings`） |
| P15 | 页脚/页码是**每页占位符**（开页脚会改变每页形状数） | `check -density` 跟着变 | 已知副作用，别当回归失败 |
| P16 | 母版上的标记会被**满版图片页**整片遮住（母版在页面形状之下） | 母版方案失效 | v1.27 起逐页贴、置顶（`woc-ai-takeover-mark`） |
| P17 | PPT 的 COM 标题**不可写**（`ActiveWindow.Caption` 只读等），只有共享窗口框 `PP12FrameClass` 接受 `SetWindowText`；**用户切标签页时 WPS 会重置标题** | 只能任务栏/Alt+Tab 看见，且会丢 | v1.26 起 Win32 标记 + 下一条命令自动补回；如实返回 `captionSupported:false` |
| P18 | **页面徽章是内容不是窗口状态**：用户 Ctrl+S 会存进文件 | 污染交付物 | v1.28 默认不贴；`-SlideMark` 显式开启 + 落盘前剥离 + 下一条命令自动修复（`badgeRepaired`） |
| P19 | 同页形状重名（旧稿"文本框 1"满天飞） | `-Shape` 命中的是第一个 | 用 `-ShapeIdx`，或 `textbox -Name` 命名 |
| P20 | 图片插入需**先拷到 ASCII 临时目录**（`AddPicture` 不认非 ASCII 路径） | 中文路径插图失败 | `image` 内部已处理（SKILL.md:509） |
| P21 | 换字体 → 字身变宽 → 多折行/挤爆框 | 排版跑版 | `font` 后必 `render -Overview`，个别框 `format -Size` 收一档 |

### 6.4 跨端 / 运维

| # | 事实 | 影响 / 做法 |
|---|---|---|
| X1 | **STA 公寓初始化 ~2.2 s**（本机安全过滤驱动所致，与 .NET/PS 无关；4 KB `[STAThread]` 探针 exe 也要 2.2 s） | 包装器一律 `-Mta`；helper 保持 `[STAThread]` |
| X2 | 常驻 helper 是**单实例串行**：并发命令排队，COM 往返总数不变 | **别用子代理并行刷同 app 命令**（数学上不省时间） |
| X3 | 每命令地板价 ≈0.5–1.6 s（含调用方 PowerShell 启动 ~0.2 s + 管道往返 1–15 ms） | 多条命令用 `batch`；一轮任务里别频繁 `close` 又重开 |
| X4 | **落盘即密文**：WPS 写的 `docx/xlsx/pptx/pdf` 头部 `17DA5FA01633CD9A` | COM 路线不受影响；交付前必须解密；helper **不产出** `ENCRYPTED` 错误码 |
| X5 | WPS 退出后会预拉起 1–2 个隐藏进程，约 1 分钟自灭 | 别当泄漏；只清理无窗口 `/Automation` 实例并**汇报进程数** |
| X6 | 用户中途关文档：Word 下条命令**静默重开**（`reopened`/`unsavedAtClose`）；Excel/PPT 报 `DOC_CLOSED` 必须显式 `open` | 铁律 3（SKILL.md:111-112） |
| X7 | 参数被静默丢弃的老毛病已有护栏（v1.18，stderr `WARN ignored parameter`） | 见到就说明命令不吃该参数 |
| X8 | `woc.cmd` 走 cmd.exe：`=A15^1.5`→`A151.5`、前导 `=` 被吃；PS 双引号里 `$B$5` 会被插值 | 用 `woc.ps1`；含 `$` 的公式用**单引号**或 `-Items` JSON |
| X9 | 三端 helper **各自复制**了一份 `Json/Bag/App/FindOpen/RequireDoc` | 新功能想"一次改三端"要分别落到三个 app（shared 只放传输/Tray/Locks） |
| X10 | `WocLocks` 只覆盖剪贴板与 state 文件，**不是**命令级并发锁 | 同一文档的并发写入仍靠"单写者"纪律（多任务模式即为此设计） |
| X11 | **本机 PS 5.1 对 UTF-8 文件的 `Get-Content`/`Set-Content` 就地改写会毁中文**（本次写这份文档时踩过一次，用 write/edit 工具才安全） | 改中文文本一律用 write/edit 工具；`-Encoding UTF8` 在 PS 5.1 也补不回已按 ANSI 解码的字节 |

---

## 7. 若要把 dsh-office 的能力接进来：最容易的两个接入点

1. **只读分析器类命令（Word 侧最省事，模板现成）**
   形态 = "一次取全文 + 本地算法 + 一条命令给结论"，不碰修订/撤销/快照，验收只需 python 独立复算。
   落点：包装层 `word.ps1` 加一个 case（照 `captions` `314-318` 或 `audit` `560-565`），helper 新文件放 `helper\word\` 并加 Dispatch 一行（`WpsHelper.cs:485` 一带）；读文本直接用现成两条通道——`FastBlocksVerified`（`FastText.cs:344/353`）与 `XmlDocTexts`（`XmlText.cs:39`），**不要**再写逐段 COM 循环。`textdiag`（`FastText.cs:558`）就是"新分析器怎么自证没走错路"的样板。

2. **Excel/PPT 的"一次性写 + 自证"命令（无修订层，机制最薄）**
   形态 = 读-改-写 + `before/after` + 自动 `preSnapshot`，反悔靠快照而非撤销栈。
   落点：`EtHelper.cs` 的 `WriteCells`（1454，原子批/expect/回滚/挂起重算）与 `FormatRange`（2028，items 每块不同格式）已把"批量写+校验+回滚"做全，新能力只需在 `Dispatch`（396-428）加 case + 复用这两个前置；PPT 侧同理复用 `Replace.cs:116` 的"写入前 `Characters()` 回读校验 + 快照 + restore"三件套。**最省事的实验路径**是先用 **`batch`**（`{"ops":[{"cmd":"…"}]}`，helper 级请求）跑通，再决定要不要加包装层参数（代价：少一层 `-Doc` 存在性校验与参数告警）。

---

## 8. 未核实清单（本文件明确没做到 100% 的地方）

1. PPT `format`/`textbox`/`table`/`shape`/`align`/`zorder`/`notes`/`view`/`save`/`quit` 的**返回字段**标注为"—（未逐行核实）"：实现行号准确，返回键名未逐个确认。
2. Excel `autofit` 的返回字段同上（`EtHelper.cs:1930` 未读内部）。
3. PPT `RequirePres`/`FindOpen` 的确切行号未定位（只确认存在同类 plumbing）。
4. `publish.ps1` 的"52 files"是**当前目录实况**；换机器/删过东西后会变，不要当常量。
5. 部署目录与 dev 的**内容级**一致只核了 `SKILL.md` mtime 与 `.cs` 数量（25/25）；逐文件哈希比对未做。
6. 三端 `helper` 的 `Json`/`Bag` 实现是否逐字节同构未验证（只确认同名 API 且键大小写不敏感）。
7. `docs\borrow\` 目录此前不存在，为本文件新建；`docs\` 下其他内容为 README 截图与 `video-assets/`（gitignore），与本任务无关。
