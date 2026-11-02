# w03 · dsh-office 的 PPT 域可借鉴清单

研究对象：`D:\DSH工作区\_research\dsh-office`（v1.2.0，MIT，copylee711/dsh-office）。
对照对象：本 skill `D:\DSH工作区\wps-office-cooperate`。

**路径基准**（全文引用均为「相对路径:行号」）：

- dsh-office 侧路径相对 `dsh-office/`：`helper/OfficeHelper.cs`（9893 行，全部 COM 实现单文件）、`src/tools.ts`（546 行，GUIDE + schema）、`docs/verification.md`（1075 行，逐版验收记录）、`README.md`。
- 本 skill 侧路径相对 `wps-office-cooperate/`：`SKILL.md`（733 行）、`scripts/ppt.ps1`、`scripts/helper/ppt/*.cs`、`scripts/tasks-ppt.ps1`。

**一句话结论**：dsh-office 的 PPT 能力不是一个「操作集」，而是**一层版式引擎**——12 套主题（token + 背景纹理 + 小部件风格 + 4 种章节页 + Beamer 边框）＋ 22 种成品版式（960×540 画布算法）＋ 自由画布/行列表格组合 ＋ 形状绘制图表 ＋ 原生 SmartArt。我们已有的 PPT 侧是「操作集 + 5 个简化预设版式」，真正的差距在**版式层**，不在单个 COM 调用。且 dsh-office 自己的验收记录里已经沉淀了**十几条 WPS 演示实测结论**，价值与我们自己的实测等价，可直接采信（见 §4.2）。

---

## §1 功能清单（op/字段 → 语义 → 实现位置 → 技巧）

### 1.1 主题系统（token 化 + 存进文稿 Tags）

| 项 | 语义 | 实现 | 技巧 |
|---|---|---|---|
| 主题表 | 12 套内置：ink / academic(别名 scholar) / paper / night / chalk / ocean / graphite / forest / sepia / crimson / slate / plum | `helper/OfficeHelper.cs:6630-6651` | 每套只给 7 个颜色 + 2 个字体 + `Dark`；`sepia` 带 `Texture="paper"`、`crimson` 带 `gradient`、`slate` 带 `grid`（`:6644-6646`）——主题 = 颜色 + 纹理 + 字体三合一，不是单纯配色 |
| token 全集 | Bg / Surface / Text / Muted / Primary / Accent / Line + TitleFont / BodyFont / Dark | `helper/OfficeHelper.cs:6603-6628` | 另有 5 组「小部件」token：`Corners`(soft/square/round/pill)、`Cards`(soft/outline/shadow)、`Marks`(square/dot/dash/triangle)、`Section`(solid/side/band/number)、`Chrome`(plain/beamer)；再加 `Transition`/`Animate`/`Base`/`Area`。**这三层 token 才让"同一套颜色能出不同气质"** |
| 落库 | 主题 JSON 存进 `deck.Tags["DSHTHEME"]` | `helper/OfficeHelper.cs:6653-6662`（读）、`:6716`（写） | 主题**随 pptx 文件走**，不用外部 sidecar；换机器/换会话不失忆。我们存在 `%LOCALAPPDATA%\wps-office-cooperate\ppt-themes\<sha16>.json`（`scripts/helper/ppt/WppHelper.cs:879-883`），文稿被移动/改名就丢 |
| 单色/字体覆盖 | `theme {bg,surface,text,muted,primary,accent,line,titleFont,bodyFont,dark}` | `helper/OfficeHelper.cs:6667-6678` | 覆盖后立刻 `Bgr()` 逐个验色（`:6715`），坏值当场报错而不是画到一半才炸 |
| 派生色 | 无固定色板，只有两条：角色名→颜色 `Tone()`；两色混合 `Mix()` | `helper/OfficeHelper.cs:8353-8370`、`:8373-8377` | 图表/多部件用 `Palette()`：primary、accent、各与 bg 混 45%、muted、混 70%、line（`:8380-8384`）——8 个互不撞的色，**不用手写调色板** |
| 章节页风格 | 4 式：solid（整页色底）/ side（左 320pt 色块装编号）/ band（横向色带）/ number（右侧 280pt 淡色大编号） | `helper/OfficeHelper.cs:7557-7630` | 未指定时**随机取一种**（`:6710-6711`），让不同文稿不千篇一律；`number` 式的字号上限 280 与被挤偏的教训见 `docs/verification.md:570` |
| Beamer 边框 | `chrome:"beamer"` 给全稿套 LaTeX Madrid 观感：顶细条 + 标题色带 + 三段式页脚 + 三角项目符号 | `helper/OfficeHelper.cs:7324-7339`（内容页头）、`:7366-7376`（页脚）、`:7379-7412`（封面/章节/结束）、`:8398-8411`（定理块） | 页脚的「报告人 / 短标题 / 页码」自动从封面 `meta` 与 `title` 取（`:7386-7391`），作者不用重复填 |

### 1.2 背景纹理（预生成一张图当背景，绕开 COM 渐变/纹理填充）

| 纹理 | 画法 | 位置 |
|---|---|---|
| `gradient` | 1600×900 位图：`LinearGradientBrush` 35° 从 Bg 到 `Mix(Bg,Primary,11%/22%)` + 左上大椭圆高光 | `helper/OfficeHelper.cs:6739-6746` |
| `grid` | 50pt 网格线，颜色 `Mix(Bg,Line,0.55)` | `:6747-6754` |
| `dots` | 36pt 间距 4pt 圆点 | `:6755-6759` |
| 图片铺底 | 按 cover 比例缩放居中 + 一层 Bg 色遮罩（alpha 226 浅色 / 205 深色）「感觉得到而不是看得见」 | `:6760-6771` |
| `paper` | ①`PathGradientBrush` 椭圆暗角（中心透明，四周 `Mix(Bg,#6B4E1E,0.5)`）②`LockBits`+`Marshal.Copy` 逐像素 ±6 噪点当纸纹 | `:6772-6786`、`:6788-6806` |
| 缓存 | key = 纹理+Bg+Primary+Line+Dark+版本号 → FNV 哈希 → `%TEMP%/dsh-office-ground-<hash>.jpg|png`；照片类存 JPEG q88 | `:6724-6730`、`:6807-6815` |
| 套用 | `Slide.Background.Fill.Solid(); .UserPicture(ground)` | `:7488-7496`（`try/catch` 失败只 Trace，不中断出页） |

> 关键设计：**背景是「一张静态图」**。这既避开了 PowerPoint/WPS 对 `Fill.OneColorGradient/TwoColorGradient/PresetTextured` 的对象模型差异，也让每页背景完全一致、渲染成本为零。代价：改主题要重生成图（已由哈希 key 覆盖），图片背景不能"随页面缩放"。

### 1.3 成品版式 `slide`（22 种 kind，960×540 画布）

**统一骨架**：`DesignedSlide()` `helper/OfficeHelper.cs:7461-7781`

- 画布：`p.Sx = SlideWidth/960; p.Sy = SlideHeight/540; p.S = min(Sx,Sy)`（`:7482-7486`）——所有版式按 960×540 写，自动适配 4:3/16:9/自定义尺寸；`area` 可把画布搬到模板页的某个自由区域（`:7481-7483`）。
- 页头：`Head()` 返回正文起点（`:7304-7363`）：48,40 一条 26×3 强调色短横 → kicker（10.5pt muted）→ 右上页码框 → 标题（27pt，>26 字降 23pt，主色粗体，带 `**强调**` 回填）→ 副标题 14.5pt → 0.75pt 分隔线 → 页脚 note 9pt（y=506）。
- 版式函数分发表：`:7760-7775`；未知 kind 报错时把全部合法值列出来（`:7775`）。
- 页码：名为 `PageNumber` 的文本框，批末统一 `Numbers()` 重写为 `01 / 12`（页脚区时为 `n / total`）（`:7823-7841`，`Renumber` 标志 `:7821`）。
- 切换：`Move()` `:7784-7796`，`EntryEffect`：fade=3849 / push=3852 / wipe=2817 / split=3848，`Duration=0.6`。
- 入场动画：`:7797-7817`，以 `Unit()`（`:6843-6847`）为「一条」，条内首个形状 `trigger=3`（接上一条），其余 `trigger=2`（同时进），`AddEffect(shape,10,0,trigger)`；总效果上限 160。**教训**：早期每页只有一个形状计入动画，导致"底板先出、文字后露"（`docs/verification.md:412`）。
- 失败回滚：`DesignSlide()` `:7448-7459` catch 里把半成品页删掉，不留脏页。
- 文字自适配：`Label()` `:6974-7016` 里 `Math.Max(6, (size<20 ? size*1.2 : size) * Math.Max(p.S,0.84))`——**小字号补偿 1.2 倍**、缩放不低于 0.84；`**词**` 用字符区间回填 bold + 强调色（`:6983-7012`）；含 `$` 转公式（`:7013`）。
- 垂直居中：`Settle(y,h,used)`——空余的 42% 放上方（`:7442-7446`），`BodyBottom=494`（`:7439`），治"内容只占上半页"。

| kind | 版面算法要点 | 位置 |
|---|---|---|
| cover / closing | 全文最复杂：左 56 起；标题 44/36/30 三档按字数；`stat` 大数字 40pt + 小单位 16pt；`style:"full"` 时满版图 + 左重右淡黑遮罩（`:7517-7519`）；无图时用 3 个圆/环做几何装饰（`:7527-7533`） | `:7509-7556` |
| section | 4 式见 §1.1；`style` 可覆盖主题的 `Section` | `:7557-7630` |
| agenda | 走 `Bullets(numbered:true)`，>4 条自动两栏 | `:7844`、`:7851-7852` |
| bullets | 8 条上限；有 `head` 时行高 108、否则 78；**按实测行宽回流**：先估 `Wide()` 折算行数，若仍溢出则按 1/0.93/0.87/0.8/0.74 逐档缩字并重排行高（`:7887-7913`），落笔后再用 `TextRange.BoundHeight` 实测，超了就 `drift` 下移后续行（`:7940-7948`）| `:7844-7953` |
| cards | 2-6 张，≤4 一排、>4 三列；单行时卡片高 250/230、两行 160；图标 34pt、value 36pt | `:7977-8007` |
| stats | ≤6 个；「只有数字」的单独分支：每个数字一张卡、值 38-68pt 按位数分档、单位 18pt 小字、下面 24×2 分隔线 | `:8009-8083` |
| chart | 图表占 100%/60%（有 `side` 时），右侧 ≤4 条说明按内容高度均分剩余空间（`air`）| `:9442-9492` |
| table | 默认用**原生 Table**；单元格含 `$` 时改 `DrawnTable()` 用形状逐格画（`:9366-9371`）| `:9354-9408` |
| formula | 3 排版：`focus`（单条大公式在色带上）/`side`（公式左、说明右，一一配对）/`stack`（公式叠说明）；卡片高度按公式"真实高度"`Tall()` 算（`\frac` 2.5 行、带上下限求和 3.1 行、`\begin{}` 按 `\\` 数）（`:8086-8100`）| `:8102-8238` |
| theorem/definition/lemma | 陈述卡片（标签胶囊 + 名称 + 陈述）+ 下方"证明思路"编号圆点；beamer 时改色带块头 | `:8387-8450` |
| gallery/figures | 2-6 图；竖图（平均宽高比<0.95）排一行竖格、否则每行 3 张；**末行不满时居中**；题注过长统一两行 | `:8453-8498` |
| summary/takeaways | 2-4 条大编号（`01`）；第 1 条反白/主色底当领句 | `:8501-8527` |
| timeline | 时间轴横线 + 圆点 + 日期（字号按字符宽度自适应 13-30pt）+ 标题 + 描述 | `:8240-8266` |
| process | ≤5 步：圆角框 + 左上序号圆盘 + 箭头（`AddShape(33)`）；框高按文字量动态算 | `:8268-8303` |
| compare | 两栏卡片，色带表头（左 primary / 右 accent），要点均分高度 | `:8305-8348` |
| image | 照片=满版 + 下 310pt 黑色渐变遮罩放白字；**图形/图表=完整显示**：宽高比<1.3 时图占左 380-540pt、说明在右（`:7645-7660`），否则满宽 | `:7631-7680` |
| split | 半页满版图（左或右），另半页 kicker+标题+正文+≤4 要点 | `:7681-7712` |
| quote | 110pt 引号 + 28pt 衬线正文 + 署名 | `:7713-7723` |
| code | 深色面板（`#1E2430`）+ 红黄绿三个窗口点 + 行号列 + 正则着色（注释/字符串/数字/关键字四色）+ 右侧 ≤4 条说明；行数上限 26、字号按最长行与行数反算 | `:8904-8979` |
| grid | 见 §1.4 | `:8884-8902` |
| canvas/custom/diagram | 见 §1.5 | `:8531-8759` |

**kind 别名**：`abstract` 类同义词做了归一，如 `closing`≈cover、`figures`≈gallery、`takeaways`≈summary、`compose|free`≈grid、`custom|diagram`≈canvas（`:7760-7774`）。

### 1.4 `grid`：不给坐标的组合版式

- 入口 `Grid()` `helper/OfficeHelper.cs:8884-8902`：`rows:[[part,part],[part],..]`，行内并排、行间上下。
- 布局核心 `GridRows()` `:8795-8881`：
  - 每格 `span`（相对宽度，默认 1）→ 宽 = `(w - gapX*(n-1)) * span / Σspan`；
  - 行高：`height` 显式 > 「固定高度部件」取最大值 > 留空（`0` = 吃掉剩余）；有 `weight` 的行按权重分剩余（`:8837-8843`）；
  - 总高超出时**所有行等比压缩**（`squeeze`），不足时整组垂直居中（`:8846-8848`）；
  - 部件默认高度表 `GridHigh()` `:8772-8792`：heading 34、node 72、callout 46、stat 118、formula 96/78、line 2、text 按 `行长*字号*0.98/宽度` 折算行数 `*size*1.5+10`；
  - `stack` 类型递归嵌一套行（`:8860-8865`）；
  - **有 `node` 时自动把 gap 从 20/16 放大到 46/40**，给连接线让路（`:8799-8804`）。
- 组装完把每格转成 canvas item（补 `box`）→ 交给 `Canvas()` 统一画（`:8898-8900`）。
- 可加 `edges:[{from,to,..}]`（`:8891-8897`）。

### 1.5 `canvas`：自由摆放 + 命名方框 + 连接线

- 入口 `Canvas()` `helper/OfficeHelper.cs:8531-8759`；`items:[{type, box:[x,y,w,h], ..}]`。
- 22 种 item type：text/title/heading/card/panel/node/circle/line/arrow/image/figure/formula/stat/chart/table/icon/bullets/steps/callout/quote/edge（`:8755` 给出全部合法值）。
- 连接线（**本任务的技术核心**）`:8558-8628`：
  1. `Shapes.AddConnector(kind, 0,0,10,10)`，kind：1=straight、2=elbow、3=curve（`:8566`）；
  2. **自动选边**：先判两框关系（左右分离 `apart` / 上下分离 `stacked`），左右相邻→右接左、上下相邻→下接上（`:8573-8600`）；
  3. curve 型先做**遮挡检测**：在两中心连线路径上采 20 点，若落在第三方框内则判定"被挡"，改为同侧出入（从上方绕过）（`:8578-8598`）；
  4. 接点用 `ConnectorFormat.BeginConnect(shape, index)` / `EndConnect`，index 由 `Site()` 算（`:8762-8769`）：4 站点按 上=1/左=2/下=3/右=4，8 站点（圆）为 `1+quarter*2`；
  5. 线色按角色色 `Tone()`、`weight`、`EndArrowheadStyle=2`、`dashed=→DashStyle=4`；
  6. **竖向连线的标签**：取 `link.Rotation` 判断是否被转 90°，交换外框 w/h 再算中点（`:8613-8616`）；曲线"拱形"标签压在弧顶（`:8617-8619`）；标签加页面底色填充块防止线穿过文字（`:8626`）。
- 能接连线的方框必须是**自带形状**：`NativeBox` 标志（`:6870`、`:8663`）——WPS 下自绘圆角自由曲线没有连接站点。
- 明暗自适应：填充色算亮度 `R*.299+G*.587+B*.114 < 150` 决定字色（`:8667-8670`）。
- 每个 item 单独一个动画步；`"with":true` 并入上一步（`:8635`）。

### 1.6 形状绘制图表（不是图表对象，是形状）

| 图表 | 算法 | 位置 |
|---|---|---|
| 柱/条/折线 | 刻度 `Ceiling()` 取 1/1.2/1.5/2/2.5/3/4/5/6/8/10×10^n，步数 3/4/5；折线在"数据都偏高"时**不起于 0** 而取低于最小值的整刻度（`:9530-9543`）；折线用 `BuildFreeform` 折线路径（**WPS 下必须显式 `Line.Visible=-1`**，`:9171-9174`、`:9605-9608`）；柱状支持 `highlight` 单色突出 | `:9494-9632` |
| 散点 + 拟合 | 最小二乘：`n*sxy-sx*sy` 斜率/截距，`R² = (n*sxy-sx*sy)²/((n*sxx-sx²)(n*syy-sy²))`；拟合线裁剪到绘图区内（`:9206-9212`）；标注 `y = 1.234x + 5.678，R² = 0.912`（`:9216`）| `:9196-9218` |
| 函数曲线/训练曲线 | 同 XY 框架；**线名标在线末端**而不是图例（`:9118-9123`）；末端标签自动防重叠：先按 y 排序后往下推、再反向回推（`Tips()` `:9275-9291`）；`endValues` 可带末值；`lines:[{y|text}]` 画参考线，`annotations:[{text,x,y}]` 画点标注 + 引线（`:9222-9271`）| `:9056-9272` |
| 坐标轴 | 两轴共用 `Axis()`/`Pinned()`/`Tick()`：右侧 5-6 格、`xMin/yMin` 可固定不回到 0（`:9094-9105`）；刻度小数位由步长反推（`:9047-9053`）| `:9024-9053` |
| 饼/圆环 | `AddShape(142)` 每片一个，`Adjustments[1]/[2]` 传起止角（-180..180），片间用页面底色 1.5pt 描边；单值≥359.9° 退化成圆；圆环用中心圆盘"挖空"+ 中心文字；右侧图例三列（名称/值/占比）| `:9294-9352` |
| 表格（形状版） | 列宽按内容长度权重（8-40 封顶）分配、隔行底色、表头主色反白、每行一条 0.5pt 线；**唯一能放真公式的表格**（`:8982-9021`）| `:9354-9408` |

> 为什么值得抄：`docs/verification.md:423` 明确记着「图表是形状画的，不是可编辑数据的 PowerPoint 图表对象」——这是**它的取舍**；反过来 `README.md:52` 的 Excel 侧才是对象图表。对我们更有意义的是：**形状图表在 WPS 演示里几乎必然可行**（全是 AddShape/AddLine/AddTextbox/BuildFreeform），而原生 `AddChart` 在 WPS 里有 5 条硬约束（见 §2）。

### 1.7 原生 SmartArt

- op：`add_smartart {slide, layout, items, left?, top?, width?, height?, colors?, look?, name?}`（`src/tools.ts:202`）。
- 版式表 20 条：`helper/OfficeHelper.cs:6290-6295`——`list/bullet_list/horizontal_list/process/chevron/arrows/timeline/steps/cycle/radial/hierarchy/org_chart/tree/pyramid/venn/matrix/funnel/target/balance/gear`，值是 PowerPoint 版式 id 的**尾段**（`process1`、`hierarchy1`、`venn1`…）。
- 真正调用：`target.Shapes.AddSmartArt(layout, l, t, w, h)`（`:6553`）；`layout` 对象的取法：遍历 `deck.Application.SmartArtLayouts`（1-based），用 `Id.EndsWith("/"+id)` 或 `Name` 全等匹配（`:6544-6551`），因此也接受「PowerPoint 里的版式名」。
- 清样例节点：新插入的图自带样例节点，先 `while (AllNodes.Count>1) AllNodes[Count].Delete()` 只留一个（`:6556`），再从这个节点长。
- 层级：`Branch()` `:6298-6335`——`{text, children:[..]}` 子节点用 `node.AddNode(5)`、同级用 `previous.AddNode(2)`；节点里 `$...$` 转成 Unicode 线性式（`Φ`、`x^2`），因为 SmartArt 节点只吃纯文本（`:6310-6321`）。
- `colors`/`look` 分别取 `Application.SmartArtColors[i]` / `SmartArtQuickStyles[i]`，包 `try/catch` 允许不支持（`:6565-6566`）。

### 1.8 图片：满版、focus 裁切、contain、缩放过小告警

| 能力 | 实现 | 位置 |
|---|---|---|
| 满版裁切 `Photo()` | `AddPicture` → `PictureFormat.Crop` 的**新式四参数**：`ShapeWidth/ShapeHeight/ShapeLeft/ShapeTop` + `PictureWidth/PictureHeight`（缩放到 cover 后的框），再算 `overX/overY` 与 `crop.PictureOffsetX/Y` | `helper/OfficeHelper.cs:7103-7121` |
| `focus` 语义 | 默认 `shareY=0.2`（**保留上五分之一——人像头部位置**）；`top/center/bottom/left/right` 可改；`crop.PictureOffsetY = overY*(0.5-shareY)` | `:7114-7118` |
| 裁切告警 | `lost = 1 - 显示面积/缩放后面积`，>0.4 记入 `Cut`（最多 8 条），结果里点名「哪张图只剩 x%、保留的是上部/中部」 | `:7119-7120`、`:7090-7096` |
| 完整显示 `Figure()` | 先画一张 `Backdrop()` 卡片（取四角像素做底色，四角不一致/半透明则白），留 10pt 内边距，等比 contain 居中 | `:7235-7250`、`:7213-7233` |
| 图形 vs 照片判定 | `IsFigure()`：读四角+两条边中点的像素，全亮（>232）且互相一致 ⇒ 图形（应当完整显示，不裁） | `:7189-7208` |
| 缩放过小告警 | `Lettering()`：`shown/drawn < 0.68`（整页）/`0.62`（半页）时记一条「10pt 字实际只有 x pt、图在页上宽几英寸」；同一批只给一次建议 | `:7159-7174`、`:7176-7187` |
| SVG | `Aspect()` 直接解析 SVG 的 `viewBox`/`width+height` 拿宽高比（`:7133-7148`）；图标本身就是 SVG（`:7066-7082`）；`add_image` 直接 `AddPicture` 传 .svg | `:7130`、`:6444` |
| 渐变遮罩 `Veil()` | 不做 24 条色条（接缝会露白线），而是**生成 256 级 alpha 的 2×256 PNG 拉伸**（`:7262-7281`）；`strong≈weak` 时退化为半透明色块（`:7256-7261`） | 同左；教训见 `docs/verification.md:431` |
| 图标 | 38 个 Feather 风格线性图标，写 SVG 到 `%TEMP%/dsh-office/icons/<name>-<color>.svg` 后当矢量图插入、命名为 `Icon <name>` | `:7024-7064`（图标表）、`:7066-7082` |

### 1.9 公式：两条完全不同的路（这正是我们最该抄的一段）

**A. PowerPoint：隐藏 Word + 剪贴板拿原生公式**（`helper/OfficeHelper.cs:5690-5692` 的注释、`:6171-6248`）

1. 起一个**不可见 Word 实例**，写 PID 到 `%LOCALAPPDATA%/dsh-office/math-word.pid`；崩溃残留下次启动 `Sweep()` 清掉（`:5708-5735`）。
2. `CommandBars.GetPressedMso("EquationLaTeXFormat")` 打开 LaTeX 输入模式（`:5759-5762`）——这样 Word 自己的 LaTeX 解析器读公式，**不用自己写 LaTeX→OMML**。
3. `OMaths.Add(range).BuildUp()` 建公式，`Unbuilt()` 校验（`:6193-6198`）。
4. `equation.Copy()` + `spot.Paste()`；**粘贴后比对文字是否变化**，没变就等 80+80×n ms 重贴，最多 8 次（`:6201-6218`）——PowerPoint 会"报成功但没贴"。
5. 贴进来是 Word 的黑色：回读实际区间，用原文字的颜色/字号覆盖；并处理「PPT 吃掉公式前空格导致整体前移一位」的边界（`:6223-6236`）。
6. 剪贴板：进来前 `Snapshot()` 存全部格式，出去后 `RestoreClipboard()`；用 `GetClipboardSequenceNumber()` 判断期间用户是否复制了新东西（是就不覆盖）；恢复时打 `CanIncludeInClipboardHistory=false`、`CanUploadToCloudClipboard=false`、`ExcludeClipboardContentFromMonitorProcessing`（`:5776-5830`）。
7. 表格单元格放不了公式 → `PptFlatMath()` 把 LaTeX 线性化后用 `Font.Subscript/Superscript` 造"假上下标"（`:5861-5923`）。

**B. WPS 演示：不接受公式 → 用 WPS 文字排版，取 EMF 当成矢量图叠上去**（`:5926-5929` 注释，`:6011-6169`）

1. 起一个不可见 `KWPS.Application`（WPS 文字），`PageSetup.LayoutMode=0` 关行网格，让行高等于公式高（`:5974-5986`）。
2. 公式以 OMML 形式 `InsertXML` 进文档（`FlatDocument()` 保留 Word 的公式结构）；行内公式在公式前**放一个黑体字母 H 当"基线尺"**，事后按它裁掉（`:6037-6041`）。
3. 尺寸：`Range(..).Information[5]` 取左右端点 x 算宽；超过 400pt 或折行就**先缩排再放回原尺寸**（最多 6 次），保证矢量不糊（`:6050-6068`）。
4. `Range.EnhMetaFileBits` 取 EMF；从 EMF 头（offset 24/28/32/36，单位 1/2540 cm）算出外框与高；`BaselineOf()` 把 EMF 渲成位图，从下往上扫找到"基线"（`:5943-5967`）。
5. 幻灯片上先占位：用 **U+2007（数字空格）** 填 n 个，字号调到 `height/1.75`（超出一行时才放大），再用 `BoundWidth` 实测回算个数（`:6080-6100`）。
6. 批末 `PlacePictures()`（`:6111-6169`）：按 U+2007 段定位 → `AddPicture` → `CropLeft/CropRight` 裁掉 EMF 两侧空白 → 按 `baseline` 对齐到文字基线（行内）或居中（独立行，额外下移 `0.3*size`）→ 命名为 `Formula of <shape> <i>` → **把图片插进该文本框的动画步之后**（`MainSequence.AddEffect` + `MoveTo(owner+1)`）（`:6151-6163`）。
7. 重写文本框时先 `DropPictures()` 删掉旧公式图（`:5998-6009`）。

### 1.10 其余 op（`PptOp` 全表）

| op | 关键点 | 位置 |
|---|---|---|
| `add_slide` | 版式表 7 项（title/title_content/two_content/title_only/blank/section/comparison），`at` 越界报错 | `:6337-6340`、`:6344-6356` |
| `move_slide` | **WPS 往后移少一位** → 循环把中间页逐张前移直到到位（`:6422-6428`）；移完 `Renumber` | `:6417-6431` |
| `reuse_slide` | 复制页→取形状→`TextFrame2.AutoSize=2`（缩字适配）→写文字；`images` 按旧几何换图并继承原名字（`:6381-6397`）；`delete` 删多余形状 | `:6359-6404` |
| `add_image` | `LockAspectRatio=0` 当 width+height 都给；返回占页面积百分比 + 缩放告警 | `:6440-6456` |
| `add_shape` | 15 种 `Forms` 常量表 + `line/arrow_line`（`AddLine`+`EndArrowheadStyle=2`） | `:6283-6287`、`:6481-6507` |
| `add_table` | 原生表；`header:false` 关首行样式；单元格含 `$` 走 `PptFlatMath` | `:6508-6536` |
| `format_text` | 支持 `find` 局部定位后再改字体/字号/粗斜下划线/颜色/对齐/项目符号/行距/`fit` | `:6571-6590` |
| `set_background` | `FollowMasterBackground=0` + `Fill.Solid` 单色（纹理走主题） | `:6473-6480` |
| `Shape()` 定位 | 支持 `title` / `body` / `#id` / 名称 / 序号，且**穿透组合** `Leaves()`；找不到时报"没有这个形状"并列出前 40 个（含 id） | `:5534-5574` |
| `PptWrite` | 行首 Tab（或两个空格）= 降一级项目符号 | `:5642-5680` |
| `add_slide` 逐字写入 | `Typing` 模式每段 `InsertAfter` + sleep 14ms，表格每行 sleep 25ms | `:5680-5688`、`:6531` |

### 1.11 套模板改稿链路（对我们场景最贴）

1. `office_open {path:新文件, template:用户的.pptx}` — 复制模板再打开。
2. `office_render {overview:true}` 看全貌；`office_read {slide:n}` 拿每页形状名与 `#id`（`:5592-5640`，装饰形状只计数不列）。
3. `reuse_slide {from:n, texts:{"#id":新文字}, images:{"#id":图}, delete:["#id"]}` — 换字换图，版式与动画保留（`:6359-6404`）。
4. 模板页放不下的内容（图表/公式/大数字）：`slide {kind, base:n, area:[l,t,w,h]}` — **在模板页副本上、指定自由区域内**画版式；区域内的示例内容按"中心点落区"删掉，但保留页标题与区外元素（logo/导航）（`:7416-7436`）。
5. `theme {base:n, area:[..]}` 把它设成默认，适合"模板只有背景图"（`:6713-6714`）。
6. `delete_slides {from:1,to:模板末页}` 收尾。
7. 缩放提示：`reuse_slide` 里 `AutoSize=2` 会让文字缩到原框内——模板字小成品也小（`docs/verification.md:421`）。

---

## §2 我方现状（对照 SKILL.md / ppt.ps1 / tasks-ppt.ps1）

### 2.1 已有（可直接复用，不必重造）

| 能力 | 我方位置 | 备注 |
|---|---|---|
| 5 种预设版式（title/section/bullets/two/end），几何随页尺寸自适应，文本框已命名供后续定位 | `scripts/helper/ppt/WppHelper.cs:1246-1376` | 注释写明「borrowed idea from dsh-office's ready-made slides, deliberately simple」（`:1244-1245`）——已吸收第一代思想，但只到"骨架"级 |
| 主题：5 色 + 2 字体 + 页脚/页码/封面跳过 + `-Apply` 全稿重刷（含深色块上的字自动转白） | `scripts/helper/ppt/WppHelper.cs:812-877`、`:2053-2188`；`SKILL.md:520` | 比 dsh-office 少 7 个颜色 token、纹理、卡片/圆角/符号三组小部件、章节页 4 式、Beamer |
| 原生图表（可编辑数据，写内嵌工作簿） | `scripts/helper/ppt/Chart.cs:1-21`（WPS 5 条硬约束的实测注记）、`:107+`；`SKILL.md:522`、`:591` | 与 dsh-office **路线相反**：它画形状，我们用对象。两者应并存（对象=读者可改；形状=能画散点拟合线/函数曲线/饼环/Box 只有对象能改数） |
| 插图（居中、尺寸、ASCII 临时副本绕中文路径、<20% 页宽告警） | `scripts/helper/ppt/WppHelper.cs:1541-1573`；`SKILL.md:509` | 缺满版 + focus + contain + 缩放比告警 |
| 原生表格 | `scripts/helper/ppt/WppHelper.cs:1575-1625`；`SKILL.md:513` | 单元格放不了公式（WPS） |
| 形状：rect / roundrect / ellipse | `scripts/helper/ppt/WppHelper.cs:1627-1674`；`SKILL.md:514` | dsh-office 15 种 + 直线/箭头线 |
| 对齐/层序/批量文字/批量替换/字体全稿/批注/备注/大纲 | `SKILL.md:502,507,516,511,517,525`；`scripts/helper/ppt/Replace.cs` | 我方在"改稿"侧的深度**超过** dsh-office（它没有跨页批量替换、没有 audit/check、没有 snapshot/restore） |
| 版式/页脚/页脚清空/换版式/套外部模板 rebase | `scripts/helper/ppt/Master.cs`；`SKILL.md:521` | dsh-office 的模板链路更细（reuse_slide 换图换字 + base/area 叠画） |
| 渲染/总览图（EMF→PNG，明文，6 页一图）/裁剪保留换图 | `SKILL.md:519`、`:510`；`scripts/helper/ppt/Deliver.cs:528-557` | 我方 `picture replace` 已读写 `PictureFormat.CropLeft/Top/Right/Bottom`，**证明 WPS 演示支持 4 个 Crop 属性**（这是 §4 的关键前提） |
| 拆册/合并/多任务合并成批（merge→plan→batch） | `SKILL.md:505-506`；`scripts/tasks-ppt.ps1:243-282` | dsh-office 无对应物（它是"一次生成"，我们是"多人多批改稿"） |

### 2.2 半有（有零件、缺策略）

| 项 | 我方 | dsh-office 多出来的部分 |
|---|---|---|
| 主题 | 5 色 2 字体，存在 `%LOCALAPPDATA%\...\ppt-themes\<sha16>.json`（跟路径绑定） | 12 套成品主题（含 3 套深色）、3 组小部件风格、4 种章节页、Beamer 边框、纹理背景、主题**存进 pptx Tags**（跟文件走） |
| 图片裁切 | 只会"保留/清掉旧裁剪"（`Deliver.cs:528-557`），不会主动裁 | focus 语义（默认上 1/5）、cover/contain 判定、`IsFigure` 图形判定、末行居中、裁掉>40% 报点名 |
| 文字适配 | `AutoFitSize()` 估算（CJK 1.0em / Latin 0.55em，`WppHelper.cs:1380-1396`）——**只估不回读** | 两段式：先估再落笔，落笔后用 `TextRange.BoundHeight` 实测回流（`helper/OfficeHelper.cs:7940-7948`）；`Settle()` 垂直居中（`:7442`）；小字号 1.2 倍补偿（`:7000`）。我们 `Check.cs:258-261` 已有 BoundHeight 判溢出，但只用于体检不用于排版 |
| 页码 | `master footer -PageNum` 走占位符，**本机 WPS 12 上每页显示 "1"**（`SKILL.md:483`） | 页码是普通文本框、名为 `PageNumber`，批末统一重写为 `01 / 12`（`:7823-7841`）——绕开占位符 bug |
| 换图 | `picture replace` 按 name/size/pos/ids 匹配，沿用旧几何 | `reuse_slide` 的 `images:{形状:图}` 语义更直接（按模板页命名形状整页换），且图片继承原形状名 |

### 2.3 没有（全新能力）

1. **成品版式层**：19+ 种 `slide kind`（cover/section×4/agenda/bullets/cards/stats/chart/table/formula/theorem/gallery/summary/code/timeline/process/compare/split/image/quote/closing）。
2. **`grid` 组合版式**（行列 + span/weight/height/stack，无坐标）。
3. **自由画布 `canvas`**（22 种 item、命名方框 + 连接线、自动选边、遮挡绕行）。
4. **连接线**（`AddConnector` + `BeginConnect/EndConnect` + `ConnectionSiteCount`）——我方零实现。
5. **形状绘制图表**：散点+最小二乘+R²、函数曲线（末端命名/参考线/标注）、饼/圆环、形状表格（能放真公式）。
6. **原生 SmartArt**（20 版式 + 层级）。
7. **切换与入场动画**（`SlideShowTransition.EntryEffect`、`TimeLine.MainSequence.AddEffect`）——我方零实现，SKILL.md 也未提。
8. **背景纹理**（paper/gradient/grid/dots/图片淡化）。
9. **公式**：我方 SKILL.md:477 直接写「WPS 演示不接受公式」就收工；dsh-office 用 WPS 文字排 EMF 再叠（§1.9B）**把这条路走通了**（`docs/verification.md:756-782`、`:914`）。
10. **SVG 图标**（38 个线性图标当矢量图插入）。
11. **图片铺底 + 遮罩**（`Veil`：256 级 alpha PNG）。
12. **模板页叠画**（`base` + `area` + 区域内示例内容清理 + 沿用模板标题位）。

---

## §3 差距与价值（面向「给领导/评审汇报的中文 PPT、套模板改稿」）

### P0 —— 直接决定"土不土"，且 WPS 演示几乎必然可做

| # | 能力 | 为什么值 | 依据 |
|---|---|---|---|
| P0-1 | **成品版式层（先做 8-10 种：cover / section×4 / bullets / cards / stats / summary / split / image / chart / table）** | 用户对 AI 做 PPT 的第一评价就是「能看但土、样式单一」，dsh-office 的结论是「差距在版面设计，不在单个操作」（`docs/verification.md:390`）。我们现在只有 5 个骨架预设，每一页长一个样 | 我方 `WppHelper.cs:1246` vs dsh-office `:7461-7781` |
| P0-2 | **主题 token 扩到 dsh-office 那 7 个 + 卡片/圆角/符号三组风格 + 存进 Tags** | 汇报稿要"像院里的稿"：主色/强调/标题/正文/次要/分隔/面 7 个角色色缺一不可；三组小部件是"同一配色不同气质"的开关 | `helper/OfficeHelper.cs:6603-6628`、`:6653-6662` |
| P0-3 | **图片满版 + focus 裁切 + contain + 裁切过多告警** | 领导汇报必有"满版大图封面""人物/成果对比图组"；WPS 里竖图默认居中裁会把头裁掉（dsh-office 专门把默认改为上 1/5） | `:7098-7127`、`:8453-8498`；我方 `Deliver.cs:530` 已证明 Crop 属性可用 |
| P0-4 | **页码改成普通文本框 + `01/12` 格式** | 本机 WPS 12 的页码占位符每页都渲染成 "1"（我们自己的实测坑，`SKILL.md:483`）——这直接是交付级 bug | dsh-office `:7823-7841` 的做法可原样移植 |
| P0-5 | **两段式文字适配（估 + BoundHeight 回流）** | 汇报稿最常见的硬伤就是"字撑破框"。我们已有估算器和体检器，只差把两者接成闭环 | `:7887-7948`；我方 `WppHelper.cs:1380`、`Check.cs:258` |

### P1 —— 显著提升专业度

| # | 能力 | 价值 | 成本 |
|---|---|---|---|
| P1-1 | **形状绘制图表**（散点+拟合+R²、函数曲线、饼/圆环） | 评审/科研场景：率定曲线、P-Ⅲ 适线、过程线，用原生图表对象画不了拟合线和 R² 标注；且形状图表在 WPS 里最稳 | 中（纯 AddShape/AddLine/AddTextbox/BuildFreeform，算法可照抄） |
| P1-2 | **grid 组合版式** | 「固定 kind 会重复自己、留半页空白」——组合版式是让一页"看起来是设计的"的最经济手段 | 中（`GridRows` 约 90 行） |
| P1-3 | **canvas + 连接线（框图）** | 汇报稿的技术路线图/系统架构图/流程图，靠 `add_shape` 摆不出连线；且 **dsh-office 已在 WPS 演示验证过连接线接在相对两边、弧线跨过并带标签**（`docs/verification.md:892`），风险比想象低 | 中 |
| P1-4 | **切换 + 入场动画** | "没有动画和背景"是用户原话（`docs/verification.md:386`）。dsh-office 用 `EntryEffect` 常量 + `AddEffect(shape,10,0,2|3)`，代码 30 行 | 低（但要实测 WPS 是否接受这些常量） |
| P1-5 | **背景纹理（paper/gradient/grid/dots）** | 「纯白背景不配革命历史题材」这类反馈（`:570`）；实现方式是"生成一张图当背景"，不碰 COM 渐变 | 低-中（GDI+ 画图约 120 行 + 缓存） |
| P1-6 | **公式矢量图（WPS 文字→EMF→叠图）** | 这是我方**唯一"不该由 dsh-office 独占"的领域**：我们本来就常驻 WPS，起一个不可见 WPS 文字比它更自然。水文/数学公式是评审稿刚需（本 skill Word 侧已有完整 WPS 文字公式能力可复用） | 高（EMF 布局、基线校准、占位空格、动画排序，dsh-office 用了约 250 行 + 多轮修复）；但**它已经把坑全踩完**，抄的是结论不是探路 |

### P2 —— 锦上添花

- P2-1 原生 SmartArt（20 版式）：好看但**WPS 支持度未验证**，且我们已有 `shape + 连接线` 的替代路径；如果 WPS 演示支持 `Application.SmartArtLayouts`，它比自画框图省一半代码。
- P2-2 SVG 图标集（38 个 Feather 图标 + 主题色重着色写临时 SVG）：`cards`/`stats`/`summary` 的点睛之笔，纯 AddPicture。
- P2-3 code 页（行号 + 正则着色）：院内汇报用得少，但"讲算法/模型"时很出彩；实现是纯文本着色，无 COM 风险。
- P2-4 Beamer chrome：只在用户明确要"数学味/LaTeX 风"时用；实现约 150 行。
- P2-5 模板页叠画（`base`+`area`）：我们已有 `master rebase`，缺"在一张模板页的空白区里画版式"这一层；套院里模板做汇报时价值高，但需要先把 P0-1 做出来才有东西可叠。

---

## §4 落地建议

### 4.1 架构与接法

**建议照 dsh-office 的形状做，但放进我们的 C# helper，而不是 PowerShell。**

- 新增 `scripts/helper/ppt/Design.cs`（部分类 `Program`，与现有 `WppHelper.cs`/`Chart.cs` 同构，纯 ASCII 源码 + 中文只走 JSON 字符串）：
  - `Theme` 类（24 个字段，照 `helper/OfficeHelper.cs:6603-6628`）+ `Named(name)` 12 套 + `Tone()/Mix()/Palette()`；
  - `Page` 类（`Slide/T/Sx/Sy/Ox/Oy/Motion/Units/Current`）+ `PX/PY/SX/SY/Block/Outline/Rule/Dot/Label/Icon/Photo/Figure/Veil`；
  - `Head()/Settle()/Numbers()/Move()` 骨架；
  - `DesignedSlide(deck, Bag)` 分发表 + 每种 kind 一个方法。
- 新命令：`ppt.ps1 design -Kind <kind> -Doc <pptx> -Items <json> [-At N] [-Theme <json>]`，或扩展 `slide -Preset` → `-Kind`（保持向后兼容：`-Preset title` 仍走旧 `BuildPreset`）。
- 主题存储迁移：`DeckTheme` 加字段 + **同时写 `pres.Tags`**（`Tags.Add("WOCDESIGN", json)`）。dsh-office 的经验是 Tags 随文件走，比我们的 sha16 文件路径健壮得多；保留旧文件路径做回退即可无缝升级。
- 一次调用画完一页（≈20-40 个形状）：dsh-office 的 `office_edit` 一批可下十几页；我们 `batch` 已有基础设施。**每页 COM 往返次数是主要成本**，务必每页只 `Save` 一次、`SlideAt` 只取一次。
- 页码/动画/切换一律放**批末**（`Numbers()`/`Move()` 的收尾思路，`:7778-7780`），避免边画边改触发 WPS 重排。

### 4.2 WPS 演示支持度评估（逐项）

| 能力 | 判断 | 依据 / 验证方法 |
|---|---|---|
| `AddShape` / `AddTextbox` / `AddLine` / `AddPicture` / `AddTable` | **能做**（我方已在用） | `WppHelper.cs:1261,1272,1561,1594,1662` 全为现役代码 |
| `BuildFreeform + AddNodes + ConvertToShape`（折线/曲线/自绘圆角） | **能做**，但**开放路径必须显式 `Line.Visible=-1`** | dsh-office 两处注释 + `docs/verification.md:811`（WPS 做的开放路径默认无轮廓，只设颜色线宽不出现） |
| `AddShape(5)` 圆角矩形 + `Adjustments[1]` 改圆角 | **做不了**（圆形圆角固定为短边 1/6） | `docs/verification.md:871`；dsh-office 因此在 WPS 下改用 Freeform 画圆角（`helper/OfficeHelper.cs:6887-6906`），我方可直接搬（`roundrect` 在 WPS 下退化成"小圆角"其实不难看，视场景） |
| `AddConnector` + `ConnectorFormat.BeginConnect/EndConnect` + `ConnectionSiteCount` | **能做** | `docs/verification.md:892`：11 页测试稿（含四层网络画布）在 **WPS 演示**渲染复核，连线接在相对两边、弧线跨过并带标签；`:878` 说明接边是插件自己算的（WPS 不重排） |
| `Shapes.AddSmartArt` + `Application.SmartArtLayouts/SmartArtColors/SmartArtQuickStyles` | **不确定，需实测** | dsh-office 只在 PowerPoint 验证（`docs/verification.md:371` 五页测试、`src/tools.ts:202` 参数说明），WPS 章节（`:688-726`）**未提 SmartArt**。WPS 12 有"智能图形"但 COM 是否暴露同名对象未知。探针：见下 §4.3-A |
| `PictureFormat.Crop` 新式对象（`ShapeWidth/PictureWidth/PictureOffsetX`） | **不确定，大概率做不了** | dsh-office 用它做满版裁切（`:7108-7118`），但**它的 WPS 章节没有"图片裁切"条目**；我方 `Deliver.cs:530-557` 用的是 4 个 `CropLeft/Top/Right/Bottom`，说明 WPS 走的是老式属性。**因此建议直接改用 4 属性 + 自己算比例**（dsh-office 的数学只用到 px 与比例，跟 API 无关） |
| `PictureFormat.CropLeft/Top/Right/Bottom` | **能做** | 我方现役代码 `scripts/helper/ppt/Deliver.cs:530-557`；Word 侧另有实测：WPS 存的是**100% 原图尺寸下的磅值**而非显示尺寸，需要 `PicCropFactor` 反算（`scripts/helper/word/Tables.cs:1365-1371`、`:1516-1543`）——PPT 侧很可能同构，**需实测确认** |
| `Slide.Background.Fill.Solid` + `ForeColor` | **能做** | 我方无现役调用，但 `master -Action rebase`/`theme -Apply` 改的是形状不是背景；`:6473-6480` 的写法在 WPS 演示应可用（同一 `Fill` 对象）；**建议探针确认 `Background.Fill.UserPicture`**（铺纹理图就靠它） |
| `Slide.Background.Fill.UserPicture(path)` | **不确定，需实测**（有回退方案） | dsh-office 包在 try/catch 里只 Trace（`:7494-7495`），说明作者也没在所有环境验证过。回退很便宜：**先插一张满版图片形状当"背景"，再插内容**（`AddPicture` 后 `ZOrder` 到底，我方已有 `zorder` 命令）。缺点：会被 `read` 当形状列出，需要在命名上过滤（如统一前缀 `BG `) |
| `Fill.OneColorGradient/PresetTextured/PresetGradient` | **建议不做** | 我们无法在 WPS 上验证且没必要——纹理用预生成图（§1.2）是更稳的等效方案 |
| `SlideShowTransition.EntryEffect`（3849/3852/2817/3848） | **不确定，需实测** | dsh-office 用了（`:7791-7793`）但它的"没有验证的"一节写明「动画和切换只核对了文件里的动画数量，没有实际放映看效果」（`docs/verification.md:417`）。判据：写入后 `read` 回读 `EntryEffect` 是否等于写入值 + `export pdf`/`render` 无异常 + **人工在 WPS 放映里看一眼**（这条必须人到场或用户确认） |
| `Slide.TimeLine.MainSequence.AddEffect(shape, 10, 0, trigger)` + `effect.Timing.Duration` | **不确定，需实测** | 同上。另外 WPS 的动画对象模型在 Word/Excel 侧已知残缺（本项目 memory：WPS `Row` 对象残缺），PPT 动画更大可能部分可用。**降级策略**：先只做切换，动画做成可选开关，失败静默跳过 |
| `TextFrame2.AutoSize = 2`（缩字适配） | **不确定，需实测** | dsh-office 在 `reuse_slide` 里用（`:6376`），但 WPS 章节未单列；我方替代方案更稳：自己按 `Font.Size` 逐档降（`BuildPreset` 的 `AutoFitSize` 就是我们自己的做法，`:1380-1396`），配合 `BoundHeight` 回读闭环即可不依赖 AutoSize |
| `TextRange.BoundWidth/BoundHeight/BoundLeft/BoundTop` | **大概率能做** | 我方 `Check.cs:490` 已在读 `BoundHeight` 做溢出体检；dsh-office 在 WPS 演示里用 `Bounds` 摆公式图（`:6132-6133`、`:6091`）——但注意它这条路径**也是靠 try/catch 兜的**，实际值需实测一次 |
| `Range.EnhMetaFileBits`（WPS 文字侧取 EMF） | **能做**（在 WPS 文字里） | 我方 Word 侧现役：`scripts/helper/word/WpsHelper.cs:834`、`:1576` 均读 `EnhMetaFileBits` 渲 PNG。公式 EMF 走同一条 |
| `KWPS.Application` 起隐藏实例排公式 | **能做** | dsh-office 全流程（`:5974-5986`）；我方 Word helper 本身就是这个实例，**可以复用现成 helper 而不是再起一个**（省 1 秒冷启动 + 不新增进程，正好契合用户"别留一堆 WPS 进程"的硬要求） |
| `Shapes.AddChart`（原生图表） | **能做但有 5 条硬约束** | 我方实测完整（`Chart.cs:6-21`）：需窗口、只收字面值、`Series.Values` 不落盘、`SetSourceData` 被拒、`NewSeries` 不可用 |
| SVG `AddPicture` | **不确定，需实测** | dsh-office 支持 SVG 插图与 SVG 图标（`:6444`、`:7077`），WPS 章节未验证；`docs/verification.md:446` 只测了微软 Office。探针：插入一张最小 SVG（1×1 的 `<svg viewBox>`）看是否报 `0x800C0005` 类错误。**降级**：图标改 PNG（40×40 也能看），或直接用 `AddShape` 拼（如圆点/短横） |
| WPS 演示"不接受公式" | **确认做不了原生**，但**图片公式能做** | 我方 `SKILL.md:477`；dsh-office 已验证可行并给了完整做法（`docs/verification.md:760-766`）。缺点：不可编辑、不能转回文本、改文本框要重排（同上 `:765`） |

### 4.3 探针验证思路（20 分钟，不碰用户文件）

统一原则：**新建临时 pptx**（`%TEMP%\woc-probe-<guid>.pptx`），每项一个 `try/catch`，把 `{item, ok, error.GetType().Name, message}` 写进 JSON；**逐项独立**（前一项失败不影响后一项），最后 `close` 并删文件。

- **A. 连接线/SmartArt/裁切/背景 四合一探针**（一个临时 3 页稿）：
  1. 第 1 页：`AddShape(1)` 两个框 → `AddConnector(2,…)` → `BeginConnect/EndConnect` → 回读 `ConnectorFormat.BeginConnectedShape.Name`（判据：能回读到、且 `Line.EndArrowheadStyle` 写入后回读一致）；
  2. 第 2 页：`Application.SmartArtLayouts.Count`（>0 即对象存在）→ 遍历找 id 尾段 `process1` → `Shapes.AddSmartArt` → 回读 `SmartArt.AllNodes.Count`（判据：layout 找到 + 形状创建 + 节点数≥1）；
  3. 第 3 页：`AddPicture` 一张竖图 → 试 `PictureFormat.Crop`（判 `E_FAIL`/`0x800A...`）→ 若失败改试 `CropLeft/Top/Right/Bottom` 并回读数值（判据：回读 = 写入）；
  4. 背景：`Slide.Background.Fill.UserPicture(<一张白底 png>)` → 回读 `Background.Fill.Type`（`msoFillPicture=6`）；失败则试 `Fill.Solid`。
- **B. 动画/切换探针**（同一临时稿）：写 `SlideShowTransition.EntryEffect = 3849` + `Duration=0.6` → 回读；`TimeLine.MainSequence.AddEffect(shape,10,0,2)` → 回读 `MainSequence.Count`。判据：写入回读一致且不抛 `E_FAIL`；**观感仍需人工放映一次**（这条要用户在 WPS 里看一眼，或 `export png` 只能验静态）。
- **C. 圆角/自绘探针**：`AddShape(5)` + `Adjustments[1]=0.5` 回读（预期 WPS 忽略）→ `BuildFreeform` 4 段圆弧小折线 → `ConvertToShape` → 回读名称与几何。判据：断言 WPS 下 `Adjustments` 回读 ≠ 0.5，且 Freeform 能成型（复现 `docs/verification.md:871`/`:884`）。
- **D. SVG 探针**：写 200 字节最小 SVG → `AddPicture` → 回读 `Width/Height`（>0 即成功）；同时插一张 PNG 作对照（若 PNG 也失败，说明是路径/加密问题而不是 SVG 问题——**用 ASCII 临时目录**，我方 `AsciiTempCopy` 已有现成函数，`WppHelper.cs:1559`）。
- **E. 文本度量探针**：写一段 40 字中文到一个 100pt 宽的框 → 读 `BoundWidth/BoundHeight/BoundLeft/BoundTop` 四个值，判 `BoundHeight>0` 且约等于 `行数*字号*1.2`（判据：数值合理，可用于回流）。

> 探针产物建议落到 `tests/ppt/probe-*.json`（现有 `tests/ppt/` 目录已有渲染与图表回归产物，结构一致）。**不要**在用户正开着的稿上跑任何一项。

### 4.4 风险与工作量

| 风险 | 影响 | 缓解 |
|---|---|---|
| 每页 20-40 个形状 → 文件变大、WPS 编辑卡 | 汇报稿 20 页 ≈ 600 形状；WPS 渲染/编辑可能明显变慢（我方实测 `render` 已 355ms/页，`SKILL.md:486`） | 装饰形状用最少数量；`check -density`（`SKILL.md:524`）加"形状数"阈值；提供 `-Lite` 开关（少图标、少装饰） |
| 外部 COM 改动不进撤销栈 | 用户 Ctrl+Z 撤不掉（`SKILL.md:477`） | 出页前 `snapshot`；`restore` 只能退文本，版式改动只能靠"重新生成那页"（可加 `slide -Kind ... -Replace N` 语义：先删页再画） |
| 动画/切换在 WPS 上不可用或不稳 | 白做工作量 | 做成"有则更好"：写好即 try/catch，回读校验失败则静默跳过并在结果里 `animations:"skipped (WPS refused)"` |
| 公式图片不可编辑 + 改文本框要重排 | 用户改一个字的公式要重来（`docs/verification.md:765`） | 在结果里明确 `formula:"picture, not editable"`；重写文本框命令里自动 `DropPictures` 再重排（照 `:5998-6009`） |
| 字体缺失导致排版全变 | 换机交付必踩（`SKILL.md:675`） | 主题里只允许白名单字体（微软雅黑/宋体/黑体/Consolas/Cambria Math），其余降级并警告 |
| 页尺寸非 960×540（4:3 / 16:9 / 自定义） | 版面错位 | 全程走 `Sx=W/960, Sy=H/540` 比例缩放（dsh-office 的 `Page` 就是这么做的，`:7482-7486`）；`-W/-H` 不写死 |

**工作量估计（C# 行数 / 人日）**

| 阶段 | 内容 | 规模 |
|---|---|---|
| 期 1（P0） | 主题 token 扩展 + Tags 存储 + `Page` 原语（Block/Outline/Rule/Dot/Label/Icon/Photo/Figure/Veil）+ `Head/Settle/Numbers/Move` + 10 种 kind | ≈1500-1800 行 / 4-6 人日 |
| 期 2（P1） | 形状图表（XY + 饼环 + 拟合）+ grid + canvas/连接线 + 切换动画 | ≈900-1100 行 / 3-4 人日 |
| 期 3（P1/P2） | 公式矢量图（复用 Word helper）+ 背景纹理 + SmartArt（视探针结果）+ code/gallery/theorem/timeline/process/compare | ≈500-800 行 / 3-5 人日 |

建议顺序：**先跑 §4.3 的 A/B/C/D/E 五个探针（半天），再定期 2/3 的具体范围**。探针结论会直接决定 SmartArt、背景 UserPicture、SVG、动画四项要不要写。

---

## §5 原文摘录（每条第 ≤30 行）

**摘录 1 · 主题的 token 结构（`helper/OfficeHelper.cs:6603-6628`）**

```csharp
    class Theme
    {
        public string Name = "ink";
        public string Bg = "#FFFFFF", Surface = "#F3F5F9", Text = "#1A2233", Muted = "#6B7280", Primary = "#0F3D75", Accent = "#C8102E", Line = "#D9DEE7";
        public string TitleFont = "微软雅黑", BodyFont = "微软雅黑";
        public bool Dark;
        public string Transition = "fade";
        public bool Animate = true;
        /// What lies behind the content: "" (the flat colour), paper, gradient, grid, dots, or "image:" and a path.
        public string Texture = "";
        /// How the section slides of the deck look: solid, side, band or number.
        public string Section = "";
        /// The small parts every layout is built from: the corners of cards and panels (soft, square, round, pill),
        /// the ground of a card (soft = a tinted fill, outline = a thin frame, shadow = white with a soft shadow)
        /// and the mark before a point (square, dot, dash).
        public string Corners = "soft", Cards = "soft", Marks = "square";
        /// The frame around every slide: "" (a dash, the title, a thin rule) or "beamer" (the look of a LaTeX
        /// Beamer talk: the title on a band across the top, a foot line with the speaker, the talk and the page).
        public string Chrome = "";
    }
```

**摘录 2 · 背景是"预生成一张图"（`helper/OfficeHelper.cs:6772-6786` + `:7494`）**

```csharp
                if (t.Texture == "paper")
                {
                    // Aged paper: darker towards the edges.
                    using (System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath())
                    {
                        path.AddEllipse(-W * 0.25f, -H * 0.35f, W * 1.5f, H * 1.7f);
                        using (System.Drawing.Drawing2D.PathGradientBrush brush = new System.Drawing.Drawing2D.PathGradientBrush(path))
                        {
                            brush.CenterColor = Color.FromArgb(0, bg);
                            brush.SurroundColors = new Color[] { Color.FromArgb(t.Dark ? 120 : 70, ColorTranslator.FromHtml(t.Dark ? "#000000" : Mix(t.Bg, "#6B4E1E", 0.5))) };
                            brush.FocusScales = new PointF(0.55f, 0.5f);
                            g.FillRectangle(brush, 0, 0, W, H);
                        }
                    }
                }
...
            try { string ground = Ground(t); if (ground != null) p.Slide.Background.Fill.UserPicture(ground); }
            catch (Exception error) { Trace("ground: " + error.Message); }
```

**摘录 3 · WPS 不给改圆角 → 自绘自由曲线（`helper/OfficeHelper.cs:6885-6906`）**

```csharp
        // WPS does not let the roundness of a rounded box be set: it stays a sixth of the short side, which on a
        // card is a bubble. There the outline is drawn: straight sides, each corner a short run of small steps.
        if (Suite == "wps" && rounded && !NativeBox && (corners == "soft" || corners == "round") && Math.Min(w, h) > 60)
        {
            try
            {
                float left = PX(p, x), top = PY(p, y), wide = SX(p, w), high = SY(p, h), r = (float)((corners == "round" ? 18 : 8) * p.S);
                float[] cx = { left + wide - r, left + wide - r, left + r, left + r }, cy = { top + r, top + high - r, top + high - r, top + r };
                dynamic builder = p.Slide.Shapes.BuildFreeform(0, left + r, top);
                for (int corner = 0; corner < 4; corner++)
                {
                    for (int step = 0; step <= 6; step++)
                    {
                        double angle = (-90 + corner * 90 + step * 15) * Math.PI / 180;
                        builder.AddNodes(0, 0, cx[corner] + (float)(r * Math.Cos(angle)), cy[corner] + (float)(r * Math.Sin(angle)));
                    }
                }
                builder.AddNodes(0, 0, left + r, top);
                shape = builder.ConvertToShape();
            }
            catch (Exception error) { Trace("drawn card: " + error.Message.Trim()); shape = null; }
        }
```

**摘录 4 · 满版图 + focus 裁切（`helper/OfficeHelper.cs:7103-7121`）**

```csharp
        dynamic picture = p.Slide.Shapes.AddPicture(path, 0, -1, bx, by, -1, -1);
        float pw = (float)picture.Width, ph = (float)picture.Height;
        float scale = Math.Max(bw / pw, bh / ph);
        try
        {
            dynamic crop = picture.PictureFormat.Crop;
            crop.ShapeWidth = bw; crop.ShapeHeight = bh; crop.ShapeLeft = bx; crop.ShapeTop = by;
            crop.PictureWidth = pw * scale; crop.PictureHeight = ph * scale;
            // What does not fit is cut away. Of a picture taller than its place the upper part is kept, a fifth down
            // from the top: that is where heads are. The author may say otherwise.
            float overX = pw * scale - bw, overY = ph * scale - bh;
            string focus = (Focus ?? "").ToLowerInvariant();
            float shareX = focus == "left" ? 0f : focus == "right" ? 1f : 0.5f;
            float shareY = focus == "top" ? 0f : focus == "bottom" ? 1f : focus == "center" || focus == "centre" || focus == "middle" ? 0.5f : 0.2f;
            crop.PictureOffsetX = overX * (0.5f - shareX);
            crop.PictureOffsetY = overY * (0.5f - shareY);
```

**摘录 5 · 连接线：选边 + 遮挡 + 接站点（`helper/OfficeHelper.cs:8564-8604` 节选）**

```csharp
                string way = it.Str("kind", it.Str("shape", "elbow")).ToLowerInvariant();
                bool bent = way == "curve" || way == "curved";
                dynamic link = p.Slide.Shapes.AddConnector(way == "straight" ? 1 : bent ? 3 : 2, 0, 0, 10, 10);
                // Which sides face each other: boxes side by side join right to left, boxes above one another bottom
                // to top. A curve is a jump over what lies between (a skip connection, a feedback): it leaves and
                // arrives on the same side, above a row and to the right of a column.
                ...
                bool across = apart && !stacked ? true : stacked && !apart ? false : Math.Abs(dx) >= Math.Abs(dy);
                ...
                if (bent && blocked) { leave = across ? "top" : "right"; arrive = leave; }
                else if (across) { leave = dx >= 0 ? "right" : "left"; arrive = dx >= 0 ? "left" : "right"; }
                else { leave = dy >= 0 ? "bottom" : "top"; arrive = dy >= 0 ? "top" : "bottom"; }
                leave = it.Str("fromSide", leave).ToLowerInvariant();
                arrive = it.Str("toSide", arrive).ToLowerInvariant();
                link.ConnectorFormat.BeginConnect(sa, Site(sa, leave));
                link.ConnectorFormat.EndConnect(sb, Site(sb, arrive));
```

**摘录 6 · WPS 演示公式 = WPS 文字排好的 EMF（`helper/OfficeHelper.cs:5925-5939`）**

```csharp
    /// $...$ in the text of a shape become native equations, in place.
    // WPS Presentation takes no equations: not built, not pasted. There a formula is set by WPS Writer, in a document
    // nobody sees, in the size and colour of the text it belongs to, and comes onto the slide as a vector picture.
    // In the text its place is kept by a run of blanks as wide as the picture — and as tall, so that the line makes
    // room — and once the batch is done and everything stands where it will stand, each picture is laid over its blanks.
    static dynamic PicWord, PicDoc;
    static readonly List<object[]> SlidePictures = new List<object[]>();
    const char Blank = '\u2007';
    /// How far down, in type sizes, a formula picture goes from the middle of its line to sit on the text's baseline.
    const double Lower = 0.3;
    /// Where the baseline of a line of slide text is, as a part of the line's height (ascent over ascent plus descent).
    const double TextBaseline = 0.8;
```

**摘录 7 · 切换与入场动画（`helper/OfficeHelper.cs:7786-7817`）**

```csharp
            string way = p.T.Transition;
            if (way != "none")
            {
                dynamic change = p.Slide.SlideShowTransition;
                change.EntryEffect = way == "push" ? 3852 : way == "wipe" ? 2817 : way == "split" ? 3848 : 3849;
                try { change.Duration = 0.6f; } catch (Exception) { }
            }
...
            dynamic sequence = p.Slide.TimeLine.MainSequence;
            int items = 0, effects = 0;
            foreach (List<object> unit in p.Units)
            {
                if (unit.Count == 0) continue;
                bool first = true;
                foreach (dynamic shape in unit)
                {
                    // The first shape of an item follows the item before it; the rest of the item comes with it.
                    dynamic effect = sequence.AddEffect(shape, 10, 0, first ? (items == 0 ? 2 : 3) : 2);
                    effect.Timing.Duration = 0.35f;
```

**摘录 8 · 最小二乘拟合线 + R²（`helper/OfficeHelper.cs:9196-9216` 节选）**

```csharp
            if (fit && !joined[s] && points.Count >= 3)
            {
                // The least-squares line through the points, and how well it holds.
                double n = points.Count, sx = 0, sy = 0, sxx = 0, sxy = 0, syy = 0;
                foreach (double[] point in points) { sx += point[0]; sy += point[1]; sxx += point[0] * point[0]; sxy += point[0] * point[1]; syy += point[1] * point[1]; }
                double den = n * sxx - sx * sx;
                if (Math.Abs(den) > 1e-12)
                {
                    double slope = (n * sxy - sx * sy) / den, cut = (sy - slope * sx) / n;
                    double varY = n * syy - sy * sy, r2 = varY > 1e-12 ? Math.Pow(n * sxy - sx * sy, 2) / (den * varY) : 1;
                    ...
                    Label(p, px + pw - 250, py + 4 + s * 16, 244, 14, "y = " + slope.ToString("0.###") + "x" + sign + Math.Abs(cut).ToString("0.###") + "，R² = " + r2.ToString("0.000"), 10, ...);
```

**摘录 9 · WPS 实测差异（`docs/verification.md:868-884`，节选 5 条）**

```
- 画布上的连线全部接在节点顶边：连线接哪一边原先交给应用自己重排，WPS 不重排。
- WPS 里卡片的圆角是短边的六分之一（几十磅）：WPS 不允许通过接口改圆角大小，PowerPoint 里是 8 磅。
- 画布连线：接哪一边由插件按两个节点的位置算（左右相邻接右到左，上下相邻接下到上），`curve` 是跨过中间节点的弧线，从同一侧出入；圆形节点的接点另算。竖向的连线在 WPS 里报的是转之前的外框，标签位置按转过的算。
- WPS 演示里的卡片轮廓改为自己画（直边加小段折线拼的圆角），圆角大小与 PowerPoint 一致；要接连线的节点仍用自带形状。
```

**摘录 10 · 曲线/折线在 WPS 里"只有点没有线"（`docs/verification.md:811` + 代码注释）**

```
- 曲线图、折线图在 WPS 里只有数据点、没有线：WPS 用 `BuildFreeform` 做出的开放路径默认没有轮廓，只设颜色和线宽不会让它出现。现在显式打开轮廓。模型当时因此改用 matplotlib 画图。
```
```csharp
                    dynamic path = Track(p, builder.ConvertToShape());
                    path.Fill.Visible = 0;
                    // WPS makes an open path without an outline, and a colour alone does not bring one.
                    try { path.Line.Visible = -1; } catch (Exception) { }
```

**摘录 11 · 动画的"一条"语义与实测教训（`docs/verification.md:412`）**

```
- 动画：用户反馈只有底板有入场动画、上面的文字先露出来。原因是每个条目只登记了一个形状。改为条目内的全部形状一起入场（首个形状接在上一条目之后，其余与它同时）。重新生成 14 页后统计：要点页 22 个形状中 14 个有动画，卡片页 24 中 18，图表页 41 中 34，没有动画的是页眉的 6–8 个固定形状。
```

**摘录 12 · 模板页叠画的接口（`src/tools.ts:192`）**

```ts
  3. Where no page fits (a chart, big figures, a formula, a table): slide {kind, .., base: n, area: [left, top, width, height]} draws that kind on a copy of template page n (a page with little on it: a content page with its sample deleted, or a plain background), inside the free area you give in points; set theme {primary, accent, text, bg, titleFont..} first to the template's colours (read them off the render) so the two match. theme {base, area} makes this the default for every slide op — the way to use a template that is only background pictures.
```

---

### 附：本次未核实/留待后续

- dsh-office `src/prompt.ts` / `src/index.ts` 里的系统提示如何引导模型选 kind（只看了 `src/tools.ts` 的 GUIDE，未读 prompt）。
- `docs/verification.md` 的 0.7~1.2 各节里关于 `grid`/`canvas` 的逐条验收细节只按 grep 命中行读了上下文，未逐节通读。
- 我方各命令的真实 COM 往返耗时（"每页 20-40 形状"的成本估算属推断，**需实测**）。
- WPS 演示对 `AddSmartArt`、`Background.Fill.UserPicture`、`PictureFormat.Crop`、`EntryEffect`/`AddEffect`、SVG `AddPicture` 的支持度：全部标 **需实测**，探针见 §4.3。
