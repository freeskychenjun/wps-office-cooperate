# skill-ppt-v131 · 成品版式层 phase 1（8 档 + 主题 token）

> 2026-11-02，skill `wps-office-cooperate` v1.30 → v1.31 的开发记录。
> 对标对象：dsh-office 的 `Theme` / `Page` 版式引擎（`helper/OfficeHelper.cs` 6600 行之后，
> 22 种 `slide` kind）；研究笔记见 `w03-ppt.md` §1.3 / §5。本文只记**落地了什么、怎么量、和上游的差别**。

## 1. 交付内容

`slide -Action add -Preset <kind>` 新增 8 档（数据一律走 `-Items <json>`，中文只走 JSON 文件）：

| kind | 数据字段 | 版式（960×540 设计画布） |
|---|---|---|
| `cards` | `[{head,text?,value?}]` 2–6 | ≤4 张一排、>4 三列两行；等高（单行 218 pt / 双行 172 pt）、等宽 18 pt 间隙；卡片底 = 主题 `cardFill`（缺省 = `surface`）、圆角 = `cardRadius`；左上 32×4 强调色短杠；value 32/26 pt 强调色 → head 17.5/15.5 pt → text 14/13 pt muted |
| `stats` | `{stats:[{value,unit?,label}]}` 2–4 | 每项一张 `surface` 卡（无 note 时 234 pt 高）；顶部 44×4 色杠（蓝/金交替）、数字 42–50 pt（**全页统一一档**，按最宽的一个定）强调色、单位 17 pt 小字跟在同一框里（`Characters()` 尾段改字号）、24×2 分隔线、label 14 pt |
| `timeline` | `{steps:[{when,head,text?}]}` 3–5 | 一条横线在正文中线（y≈305）+ 每步 14 pt 圆点；文字**上下交错**（偶数步贴线上方、奇数步贴线下方，各留 26 pt），块高按实测行数算，末步圆点与 when 用 accent；每块三框：when(accent bold)/head(16 bold)/text(14 muted) |
| `process` | `{steps:[{head,text?}]}` 3–5 | 3–5 个圆角描边块（描边 = 混色 hairline，1 pt），块高随最长文字（192 pt 起），左上 34 pt 圆形序号（蓝/金交替、白字），序号与 head(16)/text(14) 三段；块间 26 pt 间隙里放 `AddShape(33)` 右箭头（hairline 色） |
| `compare` | `{left:{head,points[]},right:{head,points[]}}` | 左右各 420 pt 卡片（`cardFill`），左蓝右金：40×4 色杠 + head 17 pt（同色）+ 6 pt 圆点要点（14 pt）；行高 = min(64, 可用高/最多条数) |
| `summary` | `{points:[{head,text?}]}` 2–4 | 纵向行（行高 min(122, 374/n)），左侧 40 pt 大编号（01/02/03，首条 accent 其余 primary），右侧 head 18 pt + 一句话 14 pt muted，行间 1 pt hairline |
| `gallery` | `{images:[{path,caption?,fit?}]}` 2–6 | ≤3 张一行、4 张 2×2、>4 张三列；**默认 `contain`：整图等比、不裁**，图底对齐单元格底部使同一行图注齐平；`-Fit cover` 才裁成统一比例并裁掉溢出（返回值 `fit` 说明用了哪种）；图注 10 pt muted 居中 |
| `split` | `{title,text?,points?,image,side}` | 半页满版图（`cover` 裁切、edge-to-edge 480×540）+ 另半页文字：56×4 色杠、标题 30 pt、正文 14 pt、≤4 条圆点要点 14 pt；页脚与页码只出现在文字半页（字号 10 pt） |

统一页脚：标题 28 pt（支持 `**强调**` → accent 加粗，`StripMarks()` 用 `Characters()` 回填）、
标题下 44×3.5 色杠 + 通栏 1 pt hairline、页脚（主题 `footer`）10 pt muted 左下、
页码 `n / 总数` 普通文本框（命名 `pageNo`）右下——**不用 WPS 页码占位符**（那玩意儿每页显示 "1"）。

返回值（8 档都一样）：

```json
{"preset":"cards","slide":1,
 "shapes":[{"name":"card1","type":"1","l":48,"t":184.4,"w":202.5,"h":218,"text":"..."}],
 "overflowWarnings":["possible overflow on shape \"cardText2\": ... needs about 58pt in a 44pt box even at 9.5pt"],
 "fontSizeDrops":[{"shape":"cardText1","from":14,"to":13}],
 "layout":[{"kind":"card","i":1,"l":48,"t":184.4,"w":202.5,"h":218}],
 "designCanvas":"960x540 scaled to 960x540","pageNumber":"1 / 8"}
```

`shapes` 是**从 WPS 回读**的真实几何（不是我们打算画什么）；`overflowWarnings` 是启发式
（CJK 1.0 em / 拉丁 0.53–0.64 em × 字号 × 1.32 行高，按框宽折行）在**阶梯降档到底仍放不下**时才报，
出页的形状也进这个数组；`fontSizeDrops` 记录每一次自动降档（14→13 这种）。

## 2. 主题 token 扩展（向后兼容旧主题文件）

`DeckTheme` 新增 4 个可选 token（`theme -Set` 的 JSON 直接接受，缺失取默认）：

| token | 默认 | 作用 |
|---|---|---|
| `accent2` | `#2E75B6` | 次强调色（双色对比/第二重点保留字；本期 compare 仍用 accent，留给后续） |
| `surface` | `#F2F5F9` | 卡片底色基色（`cardFill` 的缺省来源；旧主题文件没有它就取这个） |
| `cardFill` | 空 → 跟 `surface` | 卡片/面板底色；写 `null` 或不写 = 跟随 `surface`（`HasCardFill` 标记） |
| `cardRadius` | `12` pt | 卡片圆角（`Adjustments[1] = radius/min(w,h)`，0 = 直角、上限 0.5） |
| `sectionStyle` | 空 → `side`（原左色带） | 给已有 `section` 档位换外观：`solid`（整页色底）/`side`（左 30% 色带）/`band`（横贯色带）/`number`（右侧 220 pt 淡色大编号） |

## 3. 和上游 dsh-office 的差别（有意为之）

1. **只做 8 档**（上游 22 档）：`cover/closing` 由现有 `title/end` 承担，`agenda` 由 `bullets`，
   `chart/table/formula/theorem/code/grid/canvas` 各有独立命令或属 phase 2 的 grid/canvas。
2. **不抄它的 `Settle()` 魔法数字表**：改为"内容实测高度 → 剩余空间 42% 放上方"，且**字号阶梯回退**
   （18→16→14 这一类）由 `Text()` 的 `ladder` 参数统一执行，比上游的"先估行数再逐档缩字"少一半分支。
3. **不做入场动画/切换**（上游每条 kind 都挂 `Unit()` 动画步）：`anim` 已经是独立命令，版式层不管动画。
4. **不做 `ReuseSlide` / `grid` / `canvas`**（phase 2）。
5. **图片裁切改走 GDI+ 预裁**（2026-11-02 晚 Lead 复跑后修正）：第一版用「`PictureFormat.Crop*` 按自然尺寸换算」，子代理自己的 1200×600 素材碰巧通过，但 Lead 用 2:1 素材独立复跑时 `split` 图变成 **552×540、x=-36 出页**（`check` blocker `outframe`）——WPS 的 `Crop*` 按**点**解释，且之后写 `Width/Height` 会把裁切洗掉。现在 `KindPage.Picture(cover)` = **先按目标宽高比用 GDI+ 裁一份临时 PNG（竖图保上 1/5，横图居中），再按框精确插入**，完全不碰 COM crop；复跑回读 `splitImage l=0 t=0 w=480 h=540` 精确、8 页 `check -MinSeverity blocker` = 0。`gallery` 的 `-Fit cover` 走同一条路，默认 `contain` 不裁。

## 4. 验收（`%TEMP%\woc-ppt-v131\` 自建稿，8 页）

- `theme -Set` 自定义主题（surface `#EDF2F8`、cardRadius 10、sectionStyle `number`、中文页脚）。
- 8 档逐页加页：`overflowWarnings` 全 0（唯一一次非 0 是 split 图片框 492×540，修裁切换算后归零）。
- `check -MinSeverity major`：**blocker 0**；major 10 = `pagenum` 7（设计使然：页码是普通文本框）+ `fontsize` 3（图注 10 pt）。
- `check -MinSeverity blocker` → 0；`audit` 的 `overlap` 72 条全是"卡片底 + 其上的文字"，已知误报。
- 目检 `render -Overview` 两图 + 单页 1400 px 放大：卡片等高、时间线圆点与文字不错位、半图铺满、图注齐平。
