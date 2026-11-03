# PPT 预设版式（`slide -Preset`）：字段契约与示例

> 本文件 2026-11-04 从 `SKILL.md` 下沉（v1.34）：主文档只留「有哪些档位 + 怎么调用 + 一句话坑」，每档的字段契约、示例 JSON、返回值与限制看这里。
> 实现：`scripts/helper/ppt/Kinds.cs`（8 档成品版式）、`scripts/helper/ppt/Canvas.cs`（`canvas`/`grid`/形状图表）、`WppHelper.cs`（`slide -Action reuse`）。
> 实测坑（WPS 侧行为差异）见 [ppt-quirks.md](ppt-quirks.md)；主题 token 与 `theme` 命令见 `SKILL.md` 的 PPT 速查表。

---

## 0. 通用调用姿势

```powershell
# 一次定主题（含 v1.31 新 token：cardFill / cardRadius / sectionStyle）
ppt theme -Doc deck.pptx -Set theme.json
# 逐档加页：-Items 指向 JSON 文件（中文只在文件里，别塞命令行）
ppt slide -Doc deck.pptx -Action add -Preset cards -Items cards.json
ppt render -Doc deck.pptx -Overview -Out shots      # 60 秒看完全稿
```

- **数据放 JSON、标题也放 JSON**（`title` 字段）：七个内容档位都要求标题，缺了会报 `BAD_ARGS … needs -Title (or "title" in the -Items object)`；标题也可放命令行的 `-Title`。
- **统一外观**：标题 28pt（`**强调**` 会用 accent 加粗）、色杠 + 细线、页脚、页码「n / N」（普通文本框，**不用 WPS 页码占位符**——那玩意儿每页显示 "1"）；正文 14pt，放不下会**自动降一档字号**并记进 `fontSizeDrops`，降到底仍放不下才报 `overflowWarnings`；形状几何从 WPS 回读。
- **`overflowWarnings` 非空就别交**：要么删字、要么换档（例如 5 条要点拆成 `two`/`compare`）。
- 几何在 **960×540 设计画布**上写死、按 `PageSetup.SlideWidth/Height` 双向缩放（`Sx=W/960`、`Sy=H/540`、字号取 `min(Sx,Sy)`），所以 4:3 稿不会把字挤出框。
- 返回对象 `{preset,slide,shapes,overflowWarnings,fontSizeDrops,layout,…}`（v1.30 及以前是字符串）。

## 1. 旧档位（v1.1 起的骨架）

`title` / `section` / `bullets` / `two` / `end` —— 预设文本框已命名 `title`/`subtitle`/`number`/`body`/`left`/`right`，后续 `text`/`format` 可用 `-Shape` 直达。

## 2. 成品版式 8 档（v1.31）

| 档位 | `-Items` 契约 | 用途 |
|---|---|---|
| `cards` | `{"title":…,"cards":[{head,text?,value?}]}` | 2–6 张等高卡片 |
| `stats` | `{"title":…,"stats":[{value,unit?,label}]}` | 2–4 个大数字 |
| `timeline` | `{"title":…,"steps":[{when,head,text?}]}` | 3–5 步横向时间线，文字上下交错 |
| `process` | `{"title":…,"steps":[{head,text?}]}` | 3–5 步流程块 + 箭头 |
| `compare` | `{"title":…,"left":{"head","points"},"right":{…}}` | 左右对比 |
| `summary` | `{"title":…,"points":[{head,text?}]}` | 2–4 条编号总结 |
| `gallery` | `{"title":…,"images":[{path,caption?,fit?}]}` | 2–6 图带题注，默认 `contain` 整图不裁 |
| `split` | `{"title":…,"text"?,"points"?,"image","side"}` | 半页满版图 + 半页文字，`side` 选 `left`/`right` |

逐档命令示例（与上表一一对应）：

```powershell
ppt slide -Doc deck.pptx -Action add -Preset cards    -Items cards.json      # {"title":"…","cards":[{head,text,value}]}
ppt slide -Doc deck.pptx -Action add -Preset stats    -Items stats.json      # {"title":"…","stats":[{value,unit,label}]}
ppt slide -Doc deck.pptx -Action add -Preset timeline -Items timeline.json   # {"title":"…","steps":[{when,head,text}]}
ppt slide -Doc deck.pptx -Action add -Preset process  -Items process.json    # {"title":"…","steps":[{head,text}]}
ppt slide -Doc deck.pptx -Action add -Preset compare  -Items compare.json    # {"title":"…","left":{head,points},"right":{…}}
ppt slide -Doc deck.pptx -Action add -Preset summary  -Items summary.json    # {"title":"…","points":[{head,text}]}
ppt slide -Doc deck.pptx -Action add -Preset gallery  -Items gallery.json    # {"title":"…","images":[{path,caption}]}
ppt slide -Doc deck.pptx -Action add -Preset split    -Items split.json      # {"title":"…","image":"…jpg","side":"left"}
```

- 图片裁切走 **GDI+ 预裁再按框插入**（不碰 `PictureFormat.Crop*`——WPS 上它按点解释、且改宽高会把裁切洗掉，实测 1200×600 进 480×540 会变成 552×540 出页）；`gallery` 默认 `contain`，要统一比例裁切用 `-Fit cover`。
- 交付前 `check` 会对普通文本框页码报 `pagenum`、对 `gallery` 图注（10pt）报 `fontsize`，用 `-Config` 抑制，别为此改版式（见 [ppt-quirks.md](ppt-quirks.md)）。

## 3. `canvas`：自由版面 / 框图（v1.32）

```powershell
ppt slide -Doc deck.pptx -Action add -Preset canvas -Items canvas.json
# canvas.json = {"title":"…","items":[{"type":"node","id":"A","box":[60,150,220,80],"text":"采集层"},…],
#                "edges":[{"from":"A","to":"B","kind":"elbow","label":"原始数据"}]}
```

- **item 类型**：`text` / `heading` / `card` / `panel` / `line` / `arrow` / `image` / `stat` / `bullets` / `steps` / `callout` / `quote` / `node` / `chart`；每个 item 用 `box:[x,y,w,h]` 给坐标（960×540 设计画布）。
- **`edges` 用 `AddConnector` 真连到形状边**：`kind` 可为 `straight|elbow|curve`，边自动选（`fromSide`/`toSide` 可覆盖），`connected.begin/end` 回读为 `true` 才算连上，连不上必进 `warnings`；同一对框之间的双向标签自动错开 17pt。
- **越界自动夹取并逐条报**（`overflowWarnings`）：正文区 x48–912 / y130–494（无标题 y40 起）。
- **canvas 不会替你避让**：框重叠、文字压边不会自己挪（只报越界与连接失败），重叠要自己调 `box`。要"不用管坐标"就用 `grid`。

## 4. `grid`：结构化版面（v1.32）

```powershell
ppt slide -Doc deck.pptx -Action add -Preset grid -Items grid.json
# grid.json = {"title":"…","rows":[[{"type":"image","span":1,"path":"…","fit":"cover"},
#                                   {"type":"stack","span":1,"items":[{"type":"heading","text":"…"},{"type":"text","text":"…"}]}],
#                                  [{"type":"card","head":"…","value":"12"},{"type":"card","…":…}],
#                                  [{"type":"callout","text":"结论一"},{"type":"callout","text":"结论二"}]]}
```

- `rows` 是行数组、每行是 part 数组；`part` 有 `span`，**宽度按 span 均分**（gutter 18pt）、行高按内容估算（按 CJK/拉丁字宽算行数）、`chart`/`image`/`panel`/`stack` 吸收余量、`stack` 竖叠（无弹性项不拉伸）。
- 不用给坐标即可表达"左图右文 / 2×2 卡片 / 上表下结论"。

## 5. `reuse`：复用页（v1.32）

```powershell
ppt read -Doc deck.pptx -From 4 -Count 1                     # 先看形状名
ppt slide -Doc deck.pptx -Action reuse -From 4 -At 7 -Items reuse.json
# reuse.json = {"texts":{"title":"新标题"},"images":{"pic1":"C:\\…\\new.png"}}
```

- 复制整页后按**形状名**换文字 / 换图（沿用旧几何与旧名），保留版式与连线；替换不到的键进 `missing` 并给 warning；返回 `{reused,added,requestedAt,texts,images,applied,missing,pageNo}`。
- **复制页的页码会自动改成「新页号 / 总页数」**（页码是普通文本框，复制会留着源页的「1 / 1」）。
- **`-At` 是尽力而为**：WPS 对刚 `Duplicate()` 出来的页 `MoveTo` 被静默忽略（独立 `slide -Action move` 有效），所以返回里 `added` 是实际落点、`requestedAt` 是你要的位置；要精确位置再补一条 `slide -Action move -Index <added> -To <目标>`。
- `reuse` / `dup` 内部会**保存并恢复你的剪贴板**（`Duplicate` 借剪贴板；v1.32 修掉了"复用后再 `render` 必崩"的问题）。

## 6. 形状绘制图表（v1.32，不建原生图表对象）

- canvas/grid 的 item `{"type":"chart","chart":{…}}`：`scatter`（可 `fit:true` → 最小二乘拟合线 + 方程 + R²，按 y 带裁剪）、`curve`（分段柔化、末端标系列名防重叠、多系列图例）、`pie`/`doughnut`（`BuildFreeform` 多边形近似，≥7% 标百分比、饼+图例整组居中）。
- 刻度 1/2/5×10^k 取整、首末 x 标签左/右对齐、非负数据不出负刻度。
- **一张约 45–60 个形状**：`check` 的 `density` 规则会报（单页 >40 形状），交付时按规则抑制（`-Config`）或把点数控制在 10 个内。
- 要"读者能改数"的原生图表仍用 `chart` 命令（代价：必须有窗口、最多 3 系列）。

## 7. 已知限制（细节见 [ppt-quirks.md](ppt-quirks.md)）

- **主题写不进 pptx**：WPS 拒绝 `Presentation.Tags`（`Add` 不报错但 `Count` 恒 0）⇒ 主题存 `%LOCALAPPDATA%\wps-office-cooperate\ppt-themes\`，**换机要把这个 JSON 一起带走**再 `theme -Set`；`theme -Probe` 可复核。
- **失败版式必须整页回滚**：`BuildPreset` 半途抛异常会把半成品页留在文稿里（下一档就错位一号），v1.31 起捕获异常后 `slide.Delete()` 再抛。
- **`TextFrame.AutoSize` 别碰**：给文本框写一次 `AutoSize` 会把框压成约 29pt 高，随后 `check` 满屏 overflow。
