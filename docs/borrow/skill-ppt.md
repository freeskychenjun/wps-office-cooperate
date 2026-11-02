# skill-ppt.md · PPT 侧新增能力补丁（crop / smartart / connect / bg / anim / chart 扩展）

> 用途：交回 Lead 合并进 `SKILL.md`。本文件只写**新增内容的增量**，不重写既有段落。
> 实测环境：WPS 演示 **12.1.0.29161**，960×540 临时新稿（`%TEMP%` 自建，未动用户文稿），2026-11-02。
> 逐条实测与坑另见 `reference/ppt-quirks.md`。

## 1. 速查表（插到 SKILL.md 命令速查表 PPT 区）

| 命令 | 参数要点 |
|---|---|
| `crop` | `-Slide -Shape <图片名> [-Path <图>] [-Focus top\|center\|bottom\|left\|right] [-Fit cover\|contain] [-L -T -W -H] [-Name]`：把图片按目标框裁好/摆好；`cover`=铺满裁溢出，`contain`=完整显示居中；默认 `focus` = 水平居中、垂直保留上 1/5（人像头部），裁掉 >60% 时返回 `cut` 点名 |
| `smartart` | `-Action list [-Match <关键字>]` / `-Action add -Slide N -Preset <版式> -Items <json> [-L -T -W -H]`：`-Preset` 收英文关键字（process/cycle/hierarchy/pyramid/venn/list/matrix/funnel/target/gear/timeline/arrow/balance/picture）、中文关键字或版式全名；`-Items` 每项 `{text,children}` |
| `connect` | `-Slide N -FromShape <名> -ToShape <名> [-Kind straight\|elbow\|curve] [-Arrow none\|end\|both] [-Label <文字>] [-FromSide/-ToSide top\|left\|bottom\|right]`：**注意 `-From/-To` 是页码范围，连线两端用 `-FromShape/-ToShape`** |
| `bg` | `-Slide N\|-All [-Color #RRGGBB] [-Gradient [-Color]] [-Texture paper\|grid\|dot\|gradient] [-Image <图>] [-Reset]`：纹理是**生成一张 1600×900 PNG 再 `Background.Fill.UserPicture`**，不碰 COM 纹理填充 |
| `anim` | `-Slide N\|-All [-Transition none\|fade\|push\|wipe\|split] [-Effect fade\|fly-in\|wipe\|zoom\|appear] [-OnClick] [-Duration 0.6]`：返回每页 `readBack/matches`、`MainSequence` 前后条数，以及 `unsupported` |
| `chart`（扩展） | 新增 `-Type radar`；`-Trendline`（散点/折线加线性趋势线并显示方程+R²）、`-DataLabels`、`-LegendPos bottom\|top\|left\|right\|corner` |

## 2. 用法（插到 SKILL.md 的 PPT 操作示例区）

### 2.1 满版大图 / 人像套版

```
ppt crop -Doc deck.pptx -Slide 1 -Path portrait.jpg -L 0 -T 0 -W 960 -H 540 -Focus top
ppt crop -Doc deck.pptx -Slide 2 -Shape 图片 3 -L 40 -T 60 -W 480 -H 270 -Focus bottom
ppt crop -Doc deck.pptx -Slide 2 -Path chart.png -L 40 -T 360 -W 480 -H 150 -Fit contain
```

- `focus` 默认 = 水平居中 + **垂直保留上 1/5**（dsh-office 的经验：人像裁切保头）。竖图铺 16:9 只剩 ~28%，返回里的 `cut` 会点名，**必须 `render` 目检**再决定换 `-Focus` 还是 `-Fit contain`。
- 目标框比例与原图一致时不会裁（`crop` 全 0），此时等价于 `image -W -H`。
- 返回字段：`box`（请求框）、`boxActual`（WPS 实际框）、`crop`/`cropReadBack`（0..1 比例）、`showsPercent`、`cut`。

### 2.2 SmartArt（版式 → 文本）

```
ppt smartart -Doc deck.pptx -Action list -Match process          # 134 项里挑，中文名
ppt smartart -Doc deck.pptx -Slide 2 -Action add -Preset process \
    -Items items.json -L 60 -T 120 -W 840 -H 320
```

`items.json`：

```json
[{"text":"现状调查","children":[{"text":"资料收集"}]},{"text":"方案拟定"},{"text":"审查报批"}]
```

返回 `layout`（实际命中的中文版式名）、`layoutIndex/layoutsTotal`、`nodes`、`texts`。**关键字没命中会回退第一个版式并在 `note` 里说明**，读 `layout` 就知道是不是你要的那个。

> ⚠️ 本机 WPS 12.1：`Shapes.AddSmartArt` **能建图形**（"图示 1"），但 `Shape.SmartArt.AllNodes` 的读取/写入全部 `E_FAIL`/`RPC_E_SERVERFAULT`，节点文字写不进去（详见 quirks）。命令会如实返回 `nodes:0` + `note`。要可控的框图请用 `shape` + `connect`。

### 2.3 框图连线

```
ppt shape -Doc deck.pptx -Slide 3 -Type roundrect -L 80 -T 80 -W 200 -H 90 -Text 数据层 -Name BoxA -Fill #1F4E79
ppt shape -Doc deck.pptx -Slide 3 -Type roundrect -L 620 -T 320 -W 200 -H 90 -Text 应用层 -Name BoxB
ppt connect -Doc deck.pptx -Slide 3 -FromShape BoxA -ToShape BoxB -Kind elbow -Arrow end -Label 接口
```

- 边自动选：左右相邻→右接左，上下相邻→下接上；也可 `-FromSide/-ToSide` 强制。
- 返回 `connected.begin/end`：任一为 false 时**线画出来了但不会跟随框移动**（自绘 freeform / 组合没有连接站点），返回里会 `warning`。
- `-Label` 会加一个白底文本框压在连线中点（曲线/肘形的中点未必在线上，需目检）。

### 2.4 背景

```
ppt bg -Doc deck.pptx -All -Texture paper
ppt bg -Doc deck.pptx -Slide 1 -Texture grid -Color #F2F6FB
ppt bg -Doc deck.pptx -Slide 4 -Gradient -Color #EAF1F8
ppt bg -Doc deck.pptx -Slide 1 -Image bg.jpg
ppt bg -Doc deck.pptx -All -Reset
```

- 纹理 PNG 缓存在 `%TEMP%\woc-ppt\ground-<kind>-<color>.png`，改色/换纹理自动另存。
- 返回每页 `fillType`（1=solid / 6=picture）与 `foreColor`。`UserPicture` 在 WPS 上**真生效**（fillType 6）。
- `-Gradient` 走 `Fill.OneColorGradient(1,2,0.35)`：WPS 接受写入，但 `Fill.Type` 回读仍是 1（solid），靠不住——要"确定好看"的底色优先用 `-Texture gradient`（生成位图）。

### 2.5 切换与入场动画

```
ppt anim -Doc deck.pptx -All -Transition fade
ppt anim -Doc deck.pptx -Slide 2 -Effect fade -OnClick
ppt anim -Doc deck.pptx -Slide 3 -Effect fly-in -Duration 0.8
```

- 返回 `transition.matches`（写入值 = WPS 回读值）、`effect.added`（实际加进 `MainSequence` 的条数）。
- **只能证明"写进去了"，不能证明"放映时真播"**——这一步必须在 WPS 里 F5 人看。返回 `unsupported:true` 表示 WPS 据写或回读不符，别当成生效。

### 2.6 chart 扩展

```
ppt chart -Doc deck.pptx -Slide 1 -Data scatter.json -Type scatter -Trendline -DataLabels -LegendPos right
ppt chart -Doc deck.pptx -Slide 2 -Data dist.json -Type radar -Legend -LegendPos bottom
```

- `-Trendline`：对每个系列 `Series.Trendlines().Add(-4132,1,0,0)` 再开 `DisplayEquation/DisplayRSquared`；`decorations.trendlines[].equationShown` 是回读。
- **`-Type radar` 在本机 WPS 是"名义支持"**：`AddChart(-415)` 不报错、`ChartType` 回读 -415，但渲染出来是柱形图（见 quirks）。画雷达图请用形状自绘或换 `line`。
- 既有约束不变（窗口必需、只吃字面值、3 系列种子、`SetSourceData` 被拒、`rangeNote` 可能误报 0 点）。

## 3. 坑（进 SKILL.md 的 PPT 坑列表）

1. **`crop` 的单位是点、不是比例**：写 0.14 会被存下来、也读得回来，但渲染**完全不裁**（图被拉伸填满框）。必须写"点数 = 比例 × 缩放宽/高"。
2. **裁完之后不要再改 `Width/Height`**：WPS 会在写宽高时把已生效的裁切"洗掉"（比例仍在、渲染变回拉伸）。正确顺序：先定框尺寸 → 再写裁切 → 只调 `Left/Top`。
3. **裁切后 WPS 会对框做一次重排**，最终 `boxActual` 可能与请求框差几十点（实测 480×270 → 480×224）。所以永远 `render` 目检，`contain` 是比例精确的那条路。
4. **`Shape.CropLeft` 不存在**，只有 `Shape.PictureFormat.CropLeft/Top/Right/Bottom`。
5. **`-From/-To` 在 ppt 里是页码范围**（`read/search/export/render` 都在用），连线两端只能用 `-FromShape/-ToShape`；`-Layout` 是数字版式常量（slide/master 在用），SmartArt 版式走 `-Preset`。
6. **SmartArt 只能建不能改**（本机 12.1）：`AddSmartArt` 成功但 `SmartArt.AllNodes` E_FAIL；别承诺"SmartArt 里逐条写文字"。
7. **`Slide.FollowMasterBackground`**：设自定义背景前置 0（不跟随母版），`-Reset` 置 -1 + 白色纯填充。
8. **`bg` 与 `anim` 的 `-All`** 是"整稿"；不给 `-Slide` 时默认也是整稿（与其它命令默认第 1 页不同）。
9. 连接线**标签文本框**会进 `read` 的形状列表，名字形如 `文本框 5`；批量再处理时用 `-Label` 后记下名字或干脆不要标签。
10. 新增命令同样受"外部 COM 改动不进撤销栈"约束：改前 `snapshot`。

## 4. 命令一览变更（供 SKILL.md 首页 ValidateSet 列表同步）

`crop`、`smartart`、`connect`、`bg`、`anim` 五个；`chart` 增 `radar` 类型与 `-Trendline/-DataLabels/-LegendPos`。
