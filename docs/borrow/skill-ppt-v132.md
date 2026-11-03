# v1.32 · PPT 成品版式层 phase 2（canvas / grid / 形状图表 / reuse / 主题 Tags）

对标 `D:\DSH工作区\_research\dsh-office`（`helper\OfficeHelper.cs` 的 `Theme`/`Page` 类、`src\tools.ts:171-172` 的 grid/canvas 字段契约、`reuse_slide`），在 v1.31 的 8 档成品版式之上补第二层：
**自由版面（canvas）**、**结构化版面（grid）**、**形状绘制图表**、**整页复用（reuse）**、**主题写进 pptx**。

源码：`scripts/helper/ppt/Canvas.cs`（新文件，约 1300 行）、`scripts/helper/ppt/Kinds.cs`（`KindPage` 扩字段 + `BuildFinishedSlide` 分派）、`scripts/helper/ppt/WppHelper.cs`（`Bag.Items()`、`IsFinishedPreset`、`slide -Action reuse`、`DeckTheme.LoadWithSource` / `ThemeTagRead` / `ThemeTagWrite`、`theme -Probe`）、`scripts/ppt.ps1`（`Read-Items-Doc`、slide 的 `-From`、`theme -Probe`）。
踩坑与实测结论追加在 `reference/ppt-quirks.md` 的 2026-11-03 段。

---

## 1. `slide -Preset canvas`（自由版面 / 框图）

```json
{"title":"…",
 "items":[{"type":"node","id":"A","box":[48,170,210,74],"text":"数据采集层"}, …],
 "edges":[{"from":"A","to":"B","kind":"elbow","label":"原始数据","fromSide":"right","toSide":"left","dashed":false,"arrow":true}]}
```

- 坐标系是 **960×540 设计画布**，按 `PageSetup.SlideWidth/Height` 双向缩放（`KindCanvas`：`Sx=W/960`、`Sy=H/540`、字号 `min(Sx,Sy)`）。
- 正文区：有标题时 **x 48–912 / y 130–494**，无标题时 y 从 40 起（整页一张图）。**越界自动夹取并逐条进 `overflowWarnings`**（原文与被夹后的框都写进消息），绝不静默移动。
- item 类型：`text{text,size,bold,color,align,valign}`、`heading`、`card{head,text,value}`、`panel{fill,text,shape:rect\|round\|circle}`、`line`/`arrow`（box = 起点 + 增量）、`image{path,fit:cover\|contain}`、`stat{value,unit,label}`、`bullets{points}`、`steps{steps:[{head,text}]}`、`callout`、`quote{text,by}`、`node{id,box,text,fill}`、`chart{chart:{…}}`。未认识的类型**不丢**：画成纯文本并进 warnings（点名已知类型）。
- **edges 真的连到站点**：`Shapes.AddConnector(kind 1/2/3) → ConnectorFormat.BeginConnect/EndConnect → RerouteConnections → BeginConnected/EndConnected 回读`。边由几何自动选（沿用 `connect` 命令的 `SideAuto`：左右相邻右接左、上下相邻下接上），`fromSide/toSide` 可覆盖。**连接失败进 warnings**（`begin/end` 两个布尔都回读），返回里 `edges[]` 逐条给 `{from,to,connected:{begin,end},kind,fromSide,toSide,label}`。
- 连线标签：文本框 + 白底（`Fill.Solid()`）压在线中间；同一对框之间的两条边（A→B 与 B→A）中点相同，第二个标签自动下移 17 pt（`KindPage.LabelSpots`）。

## 2. `slide -Preset grid`（结构化版面）

```json
{"title":"…","rows":[[{"type":"image","span":1,"path":"…","fit":"cover"},
                      {"type":"stack","span":1,"items":[{"type":"heading","text":"…"},{"type":"text","text":"…"}]}],
                     [{"type":"card","head":"…","text":"…","value":"12"}, …],
                     [{"type":"callout","text":"结论一…"},{"type":"callout","text":"结论二…"}]]}
```

宽度与高度分开算：

1. **宽度**：每行内 `span`（默认 1）之和做分母、`gutter` 固定 18 pt 均分 → `cellW = (864 − 18×(n−1)) / Σspan`，一格宽 = `cellW × span`。
2. **行高**：先按内容估 `need`——`text`/`bullets`/`callout`/`quote`/`panel`/`card` 用 v1.31 那套 CJK/拉丁字宽估算器算行数（`KNeedHeight`），`chart` 150、`image` 120 当地板，`heading` 28、`stat` 58–78；
3. 所有 `need` 之和 > 可用高度 → **等比压缩并报 warning**（"drop a row or shorten the text"）；否则多出来的高度**只给能吸收的行**（含 `chart`/`image`/`panel`/`stack`，或显式 `"grow":true`），没有吸收行就按 need 比例摊。
4. `stack` 内部同一套算法，但**没有弹性项时不拉伸**（v1.31 教训：拉伸纯文字只会在行间留缝），内容顶对齐。

`card` 的字号/内边距由 `KCardMetrics()` 统一给（vs=24、hs=14、上 26 下 14），**测量与绘制共用同一组数**，所以 row 高估得准（早期版本按 `h×0.22` 动态算，量出的高和画出的高不一致，卡片正文被挤成 6 pt 框）。

## 3. 形状绘制图表（不建原生图表对象）

原生 `chart` 需要窗口、只吃字面值、种子 3 系列、`SetSourceData` 被拒，而且无法塞进指定框。所以 canvas/grid 里的 `{"type":"chart","chart":{…}}` 全部用**形状 + 线 + 文本**画：

- `scatter`：坐标轴 + 刻度 + 散点 + 可选最小二乘拟合线。拟合对**全部系列的数据点合并**做 OLS（`slope=(nΣxy−ΣxΣy)/(nΣx²−(Σx)²)`），`R²=1−SSE/SST`；拟合线用两个端点画、**按 y 轴带宽裁剪**（`(lo−b)/k`、`(hi−b)/k` 求交），方程与 `R²` 写在绘图区内右下角（`y = 2.264x + 12.725   R² = 0.999`）。
- `curve`：`series[].x[]/y[]`（也接受 `points:[[x,y]]`），分段直线 + 中点折线做柔化（`smooth` 默认 true）；点数 ≤12 时画数据点；**线末端标系列名**，同名位置冲突时下移 11 pt，越界时翻到左/下方；多系列再加顶部图例。
- `pie`/`doughnut`：扇区用 `BuildFreeform` 多边形近似（每 6° 一点，圆环再把内弧倒着走一遍），**最后补一个首点闭合**；`frac ≥ 7%` 的扇区在中间写百分比（白字），右侧图例给「类别 值」；圆环中间再叠一个白色圆形当前景洞（比依赖环形多边形稳）。`BuildFreeform` 被拒时降级成 `AddShape(142)` + `Adjustments` 起止角，并在 warnings 里说清楚（圆环会退化成实心饼）。
- **刻度用 1/2/5×10^k 取整**（`KNiceTicks`）：`raw = span/4` → `mag = 10^floor(log10 raw)` → `step = mag×{1,2,5,10}` → `lo/hi` 向外取整；刻度数 >9 时 step 再翻倍。首/末 x 标签左/右对齐（居中会压到 y 标签或出图）。数据全非负时 `lo` 收在 0（集水面积/流量不该出现负刻度）。
- 返回加 `chartTypes:["scatter","curve","pie","doughnut"]`，画不了的类型（`bar`/`line` 之类）进 warnings 并画一行说明文字。

## 4. `slide -Action reuse`（复用一页，换文字/图片）

```
ppt slide -Doc X -Action reuse -From 4 -At 7 -Items reuse.json
{"texts":{"title":"新标题","c6-t":"换掉这句话"},"images":{"g1-1":"C:\\…\\new.png"}}
```

- **`-From` 在这里是源页号**（`read/render/export` 里它是页码范围起点，返回的 `note` 会把这件事说清楚）；`-At` 是插入位置，缺省紧跟在源页之后。
- `Slide.Duplicate()` 复制整页（版式、母版、连线都在），再按**形状名**换文字 / 换图；换图沿用原几何（`AddPicture` 按旧 `l/t/w/h` 插入 → 删旧图 → 新图沿用旧名）。
- **替换不到的键绝不静默**：`missing:{texts:[],images:[]}` + 顶层 `warning`（还会告诉你去 `read -From <新页> -Count 1` 看真实形状名）；写入后回读校验，不符也算 missing。
- 返回 `{reused, added, requestedAt, slides, texts:[{shape,before,after,ok}], images:[{shape,path,l,t,w,h,ok}], applied, missing, warnings}`。
- **已知限制**：WPS 里 `Duplicate()` 返回的对象上 `MoveTo()` 被静默忽略（独立命令里同样的 `MoveTo` 又有效），所以一个请求内搬不动刚复制的页 → `added` 报**实际**落点、`requestedAt` 报要求的位置、搬不动进 `warnings`；要精确位置再补一条 `slide -Action move -Index <added> -To <目标>`（实测有效）。

## 5. 主题写进 pptx（探测结论：WPS 不支持）

- `DeckTheme.LoadWithSource(pres, out source)`：**本机主题文件 → 文稿 `Tags["wocTheme"]` → 内置默认**；`theme` 命令返回 `themeSource` 说明命中哪层。
- `theme -Set` 写文件后尝试 `Pres.Tags.Add("wocTheme", json)`，回读校验后给 `tags:{supported,readBack,chars,exact,note}`。
- **实测 WPS 12.1.0.29161：`Tags.Add` 不报错但 `Tags.Count` 仍为 0、读不回来** → 返回 `supported:false` + "WPS refused Presentation.Tags on this build - the theme stays in the local file only"。`theme -Probe` 单独做探针（写→读→删），不写任何主题数据。
- 结论：**保持文件方案**，跨机时把 `%LOCALAPPDATA%\wps-office-cooperate\ppt-themes\<hash>.json` 一起带走；换机后 `theme -Set` 一次即可。

---

## 验收（本次实测）

自造稿 `%TEMP%\woc-v132-ppt\v132.pptx`（960×540，7 页），全部 `render -Width 1400` 逐页目检：

页序（`-At` 搬不动，reuse 的副本落在源页之后，见下）：

| 页 | 版式 | 目检结论 |
|---|---|---|
| 1 | canvas 框图（4 节点 5 条边） | 边**连在框边**（右→左、下→上、虚曲线、回执反向），移动框线跟随；两条重叠边的标签**自动错开**且白底可读；越界项被夹取并报 warning |
| 2 | reuse（复制第 1 页，换标题与小标题） | 版式/连线/页脚全部保留，只换了 `title` 与 `c1` 两处文字（`applied:2`） |
| 3 | canvas 四张形状图表 | 散点轴 0/100/200/300、拟合线 `y=2.264x+12.725 R²=0.999` 落在轴内右下；双系列折线末端标 2023/2024 且互相错开；饼图 46/27/18/9 与圆环 72/28 是真扇区，百分比与图例对得上 |
| 4 | grid 左图右文 + 2×2 卡片 + 两条结论 | 四卡等高对齐、gutter 一致；图片 cover 裁切不出框；stack 内 heading/text/bullets 顶对齐不留缝 |
| 5 | reuse（复制第 4 页，换图） | 新图沿用旧几何 `423×153`、`ok:true`；不存在的文字键进 `missing`（`applied:1`） |
| 6 | grid 左数字右图 + 通栏散点 | stat 与图同高对齐；曲线轴 500…2500 / 2014…2024；拟合式在右下、图例在右上，不打架 |
| 7 | grid 上图 + 两结论（饼图居中） | 饼图与图例**作为一组居中**，不再一边空一半 |

- `ppt check -MinSeverity blocker` → `verdict=mechanically clean`、**blocker 0**。
- 返回值字段（canvas/grid）：`{preset, slide, shapes[], overflowWarnings[], fontSizeDrops[], layout[], designCanvas, bodyArea, chartTypes[], edges[]}`。

## 已知限制

- `chart` 只画 `scatter / curve / pie / doughnut`；`table` 类型不在这里画（用 `table` 命令补）。
- 形状图表不参与"读者改数"（要可编辑图表还是用原生 `chart`，代价是窗口 + 3 系列上限）。
- 形状图表一张 45–60 个形状，`check -Rules density` 会报；交付时按规则抑制。
- `reuse -At` 是尽力而为（见上），精确位置补一条 `slide -Action move`。
- **`reuse` 之后 `render` 会崩**（helper 无响应）：WPS 把复制的幻灯片留在剪贴板上，`ClipboardKeeper` 处理不了这个格式。先 `Set-Clipboard -Value x` 覆盖剪贴板再渲染即可（详见 `reference/ppt-quirks.md` 2026-11-03 段）。
- 主题无法写进 pptx（WPS 拒绝 `Presentation.Tags`），跨机要带主题文件。
