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

