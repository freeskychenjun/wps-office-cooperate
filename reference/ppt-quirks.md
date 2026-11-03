# PPT 侧实测事实（WPS 演示）

> 一条一行，只记**实测过**的结论；每条带日期与 WPS 版本。与 `word-quirks.md` / `excel-quirks.md` 同构。
> 版本口径：WPS 演示 **12.1.0.29161**（本机），画布 960×540，全部在 `%TEMP%` 自建新稿上测，不动用户文稿。
> 用途：写/改 `scripts/helper/ppt/*.cs` 前先扫一遍，避免重踩。

## 2026-11-02 · v1.29 版式能力（crop / smartart / connect / bg / anim / chart 扩展）

- `Shape.PictureFormat.CropLeft/CropTop/CropRight/CropBottom` 可读写；`Shape.CropLeft` 属性**不存在**（写它就报错）。值**按点解释**，不是 0..1 比例。
- **写 0..1 的比例等于没裁**：`CropTop=0.14` 存得进、读得回 `0.14`，但渲染完全不变（图被拉伸铺满框）。必须写 `比例 × 缩放后的宽/高`（点）。
- 裁切生效后**不要再改 `Width/Height`**：一改，WPS 把裁切"洗掉"（比例还在、渲染变回拉伸）。顺序必须"先定框尺寸 → 写裁切 → 只调 Left/Top"。
- 写裁切时 WPS 会对框重排一次，`boxActual` 与请求框有偏差（实测请求 480×270 → 实际 480×224；请求 320×180 → 320×313）。要比例精确用 `-Fit contain`（实测 1200×600 → 300×150 居中，精确）。
- `crop -Fit contain` 走 `LockAspectRatio=-1` + 先设宽、超高再设高 + 居中，实测精确，是"图不能裁"时的安全路径。
- `Shapes.AddSmartArt(layout, l, t, w, h)` **能建图形**（返回形状名"图示 1"），`Application.SmartArtLayouts.Count = 134`，**全部中文名**（基本流程 / 循环关系 / 层次结构…），`Id` 仍是 PowerPoint 的 urn 尾段（`process1`、`hProcess9`…）→ 英文关键字必须先映射成中文关键字再匹配。
- **SmartArt 节点不可写**：`Shape.SmartArt.AllNodes` 在 WPS 12.1 上抛 `E_FAIL`（个别布局 `0x80010105 RPC_E_SERVERFAULT`）；`node.TextFrame2 / TextFrame / Shapes(1).TextFrame` 三条路都失败，`AddNode(5)/AddNode(2)` 同样失败。结论：**SmartArt 只能建，不能逐条写文字**；要可控框图用 `shape` + `connect`。
- `Shapes.AddConnector(kind, beginX, beginY, endX, endY)` + `ConnectorFormat.BeginConnect(shape, site)` / `EndConnect` 可行：`BeginConnected/EndConnected` 回读 -1（true）；站点序号 上=1 左=2 下=3 右=4；`RerouteConnections()` 不报错。
- 连接线 `Line.EndArrowheadStyle = 2`（三角）写入后回读 2；`BeginArrowheadStyle` 回读 1（无箭头）→ `-Arrow both` 需读回确认（本机回读 begin=1，说明"双向"未真正生效，见 v1.29 已知问题）。
- `Slide.Background.Fill.UserPicture(png)` **可行**，`Fill.Type` 回读 6（picture）；`Fill.Solid()` + `ForeColor.RGB` 回读 1；`Slide.FollowMasterBackground = 0/-1` 可写。
- `Slide.Background.Fill.OneColorGradient(1,2,0.35)` **写不报错但 `Fill.Type` 仍回读 1**（solid）→ 渐变背景在 WPS 上不可信，纹理背景一律"生成 PNG + UserPicture"。
- `SlideShowTransition.EntryEffect = 3849`（fade）写入后回读 3849（一致）；`Duration` 可写。**只能证明写进去了，不能证明放映真播**（需 F5 人看）。
- `Slide.TimeLine.MainSequence.AddEffect(shape, 10, 0, 1|2)` 可行：4 个形状加 4 条，`MainSequence.Count` 0→4。同样只证明入队，不证明真播。
- `Shapes.AddChart(-415, …)`（xlRadar）**不报错、`Chart.ChartType` 回读 -415，但渲染出来是柱形图** → 雷达图在 WPS 是"名义支持"，别承诺。
- `Series.Trendlines().Add(-4132,1,0,0)` 可调用并返回 Trendline；`DisplayEquation/DisplayRSquared` 写入不报错但回读 false（`equationShown:false`）→ 趋势线**画得出、方程/R² 标注不可靠**。
- `Chart.ApplyDataLabels(2)` 可行（`dataLabels:true`）；`Chart.Legend.Position` 写入后回读一致（bottom=-4107 / right=-4152）。
- 既有图表硬约束不变：必须有窗口、只吃字面值、种子 3 系列、`SetSourceData` 被拒；`rangeNote` 会在 `SeriesCollection(1).Values` 读回异常时报"plots 0 points"，**数据其实写进去了**（`firstSeriesValues` 正确），属已知误报。
- `Shapes.AddShape` 出来的形状**默认名是中文**（"圆角矩形 1"/"椭圆 3"），`-Name` 参数在 v1.29 才被 `shape` 命令消费（此前静默丢弃）→ 给形状命名后 `connect` 才能按名找到。
- 打开不存在的路径 = 新建空稿（`slides:0`），需先 `slide -Action add`；`open` 冷启动约 17 s，之后每命令 0.5–1.2 s；`chart` 每条 12–14 s（内嵌工作簿轮询）。
- 临时稿 `close` 后 `wpp` 进程数归 0（本机同机的 `wps.exe` 是别的任务的 Word 接管实例，别按映像名杀）。

## 2026-11-02 · v1.31 成品版式（cards/stats/timeline/process/compare/summary/gallery/split）

- **写 `TextFrame.AutoSize`（任何值）会把文本框压成 29 pt**：实测 `AddTextbox(1,l,t,w,h)` 之后写 `AutoSize = 0`，形状高度回读一律 29（标题框 52 → 29、卡片数字框 46 → 29、页码框 16 → 29 全都变 29），宽度和位置不受影响；后果是 `check` 的 `overflow` 规则拿 29 当框高，**满页报 blocker**（"text needs 33.6pt but the box is 29pt tall"）。做法：**不碰 `AutoSize`**，设完文字后把 `Left/Top/Width/Height` 再写一遍钉死。
- **`PictureFormat.Crop*` 的比例基数是"图片的自然尺寸"，不是当前框宽**：960×540 半页拦腰 480×540、900×900 方图按 `eL = overX/2` 写裁剪，实测框回读 492×540（少裁 12 pt），且 480+492 出页 → `check` 报 `outframe`。正确写法 `v = e × 自然宽 / 缩放后宽`（`自然宽` = `AddPicture` 之后立刻读到的 `pic.Width`），写完再 `check`：实测框**精确回到 480×540**、无越界。v1.29 的 `crop` 命令用的是"比例 × 缩放后尺寸"，因此它会**少裁几十点**（这正是 v1.29 记录的 480×270→480×224 偏差来源）。
- **页码必须用普通文本框**：`slide.HeadersFooters.SlideNumber.Visible` 在 WPS 演示上每页都显示 "1"（已知缺陷）。成品版式写 `n / 总数` 的普通文本框（命名 `pageNo`），代价是 `check -Rules pagenum` **每页都报 major "no visible slide number"**（判据就是那个占位符），交付时用 `-Config {"suppress":[{"rule":"pagenum"}]}` 抑制，别当成漏做。
- `check` 的 `overflow` 判据是**实测 `TextRange.BoundHeight` vs `Shape.Height`**（比值 > `overflowFactor` 1.05 + 2 pt 才算），所以框高一旦被 WPS 改错就会满屏误报；反过来，框高正确时它能抓出真溢出。
- **`audit` 的 `overlap` 会把"色块 + 落在块内的文字"逐对报出来**（8 页 72 条）：成品版式的卡片底色/面板按设计就压在文字下面，属已知误报，别按条数硬改；判据以 `check` 与 `render` 为准。
- `check` 的 `fontsize` 用 `profile` 的 `bodyMin`（汇报 14 pt）卡**所有**非标题非页脚文字，所以图注（10 pt）、小标签（13 pt）会报 major；要交就得 `-Config {"bodyMin":12}` 或按规则抑制。
- **失败版式必须整页回滚**：`BuildPreset` 半途抛异常时（例如 `ShapeD` 对 dynamic 的绑定失败）会把半成品页留在文稿里（下一档就错位一号，`slide 8` 变 `slide 9`）。v1.31 起 `BuildPreset` 捕获异常后 `slide.Delete()` 再抛。
- 版式几何在 **960×540 设计画布**上写死、按 `PageSetup.SlideWidth/Height` 双向缩放（`Sx = W/960`、`Sy = H/540`、字号取 `min(Sx,Sy)`），所以 4:3 稿不会把字挤出框；`slide -Preset` 的返回值已改为对象 `{preset,slide,shapes,overflowWarnings,fontSizeDrops,layout,...}`（v1.30 及以前是字符串）。

## 2026-11-03 · v1.32 canvas / grid / 形状图表 / reuse / 主题 Tags

- **WPS 12.1 拒绝 `Presentation.Tags`**：`pres.Tags` 能取到对象，`Tags.Add("wocTheme", json)` 不抛错，但 `Tags.Count` 回读仍是 0、`Tags.Item(k)` 取不到、`Tags("wocTheme")` 也读不回 → 主题**只能留在本机文件**（`%LOCALAPPDATA%\wps-office-cooperate\ppt-themes\<hash>.json`），换机就丢。判据用 `ppt theme -Probe`（写一个探针 tag 再读回、最后删掉），实测 `supported:false`；`theme -Set` 的返回里也带 `tags.supported/note`，**如实说"没写进去"**，不假装。读主题的顺序因此是 **文件 → 文稿 tag → 内置默认**（`themeSource` 字段说明命中的是哪一层）。
- **`Slide.Duplicate()` 返回的对象上 `MoveTo()` 被静默忽略**：`dup = src.Duplicate()` 之后 `dup.MoveTo(n)` 不报错、slide 也不动；同一个 `pres.Slides.Item(k).MoveTo(n)` 在**独立的一次命令**里是可用的（`slide -Action move` 实测能把第 7 页搬到第 2 页）。改用 `pres.Slides.Item(dup.SlideIndex).MoveTo(n)` 仍然无效 → 一个请求内**刚复制出来的页搬不动**。因此 `slide -Action reuse` 的 `-At` 是尽力而为：返回 `added`（实际落在哪）、`requestedAt`，搬不动时进 `warnings`；要精确位置就再来一条 `slide -Action move -Index <added> -To <目标>`。
- **`BuildFreeform` + `ConvertToShape` 在 WPS 可用**：饼图扇区/圆环用多边形近似时，`Shapes.BuildFreeform(0, x, y)` + `AddNodes(0, 0, x, y, Missing, Missing, Missing, Missing)` 逐点建，**最后要把第一个节点再加一次**把环闭合（否则填充缺一块）；`ConvertToShape()` 出来的形状 `Type` 回读是 **5**（圆角矩形那个值），别拿它判断"没建成"。真失败时降级用 `AddShape(142)`（饼图）+ `Adjustments[1]/[2]` 写起止角。
- `Shapes.AddLine` 在 WPS 上稳定（图表坐标轴、折线分段、拟合线全用它，一次几十条没有失败），`Line.EndArrowheadStyle=2` 画箭头、`DashStyle=4` 画虚线；连线的 `ConnectorFormat.BeginConnect/EndConnect` 之后 `BeginConnected/EndConnected` **回读为 true**（-1），这是"线真的连到框上"的唯一可信判据。
- **文本框能当标签底板**：`AddTextbox` 之后 `Fill.Solid()` + `Fill.ForeColor.RGB` + `Line.Visible=0` 生效（`connect -Label` 早就这么干），连线标签这么做才不会被线穿过字。
- **PS 包装器的 `Read-Items-Json` 会拆掉顶层 `items` 键**（有 `items` 就只返回那个数组），canvas/grid/reuse 的文档对象（title + items/rows + edges + texts/images）会被拆散 → 新增 `Read-Items-Doc`（整对象当一个数组元素送给 helper，由 helper 的 `KDoc` 认领）。**给这几个命令写 JSON 时，顶层键别再叫 `items` 之外的任意名**都行（`items`/`rows` 都可以，包装器已按命令分派）。
- 图表刻度：**别用"数据范围 ±6% 留白"**，出来的刻度是 `-1.7 / 62.2 / 1914` 这种难看且会在正值轴上出现负数；改成 1/2/5×10^k 取整（`KNiceTicks`）后是 `0 / 100 / 200 / 300`、`500 / 1000 / 1500`，并且首末刻度标签改成左/右对齐，否则第一个 x 标签会压到 y 标签上。
- 图片/图表标签框的**高度要给够**：`KindPage.Text` 的溢出判据是 `行数×字号×1.32` 对框高，8.5 pt 的刻度标签塞进 11 pt 高的框必然报 `possible overflow`（满页几十条噪声）。给 14 pt 高 + 单元素字号阶梯就没噪声了。
- **同一对框之间的两条连线会重叠**：`from A to B` 与 `from B to A` 的中点完全一样，标签叠在一起。做法：页内记 `LabelSpots`，第二个标签自动下移 17 pt（`KDrawEdges`），标签再加白底。
- **`render` 在 `slide -Action reuse`（`Slide.Duplicate`）之后必崩**：症状是 `{"code":"INTERNAL","message":"The helper exited (code ) without writing a response."}`——helper 进程直接死掉、没有任何响应（逐页重试、重启 daemon 都无效）。根因是 **WPS 把刚复制的幻灯片留在剪贴板上**，而 v1.32 的 `ClipboardKeeper.Keep()`（helper/shared）在"保存用户剪贴板内容"时被这个 slide 格式拖崩。**判据与解法**：先 `Set-Clipboard -Value x` 用纯文本覆盖剪贴板，再 `render`，一次就过（本次实测：覆盖前 3 次全崩，覆盖后连续 3 次正常）。**修法（v1.32 当天已落地）**：① `ClipboardKeeper.Snapshot()` 改成**只接管简单格式**（text/unicode/html/rtf/csv/bitmap/filedrop/png），OLE/slide/嵌入文档一律跳过不读，且剪贴板里没有文本/图片/文件时整段跳过；② `slide -Action reuse` / `dup` 内部用 `ClipboardKeeper.Keep()+Sync()` 包住 `Duplicate()`，顺带把用户剪贴板内容还回去。修后实测：`reuse` 之后直接 `render` **不再需要先清剪贴板**（Lead 复跑 2 页一次成功，用户文本剪贴板原样保留）；旧写法"先 `Set-Clipboard`"只留作未知前置状态下的救急手段。
- 形状绘制图表的坐标轴、刻度、图例全是独立形状，`check` 的 `density`（单页 >40 形状）会被它们撑破：一张散点图约 45–60 个形状。要交就抑制 `density`，或把数据点减到 10 个以内。
