# dsh-office 借鉴清单：研究结论与落地状态

> 研究日期 2026-11-02 ｜ 对象 [`copylee711/dsh-office`](https://github.com/copylee711/dsh-office) v1.2.0（MIT，C# helper + TS 插件，同时支持微软 Office 与 WPS）
> 本机 skill：`wps-office-cooperate` v1.28 → v1.29 ｜ 环境：Windows + WPS 12.0/12.1.0.29161（`KWPS`/`KET`/`KWPP`）
> 本目录下五份域报告是本次研究的原始产出（约 31 万字符），本文是**结论层**：抄什么、不抄什么、WPS 实测结论、落地到哪一步。

---

## 1. 一句话结论

dsh-office 与我们做的是**同一件事的两条路线**：它是「DSH 插件 + 通用 Office 自动化工具集」（七个工具、按需下发操作清单、不碰用户界面），我们是「skill + 在 WPS 里当场协作」（修订、批注、快照回滚、逐条锚点校验）。**我们已有的读写操作、协作与安全机制普遍比它强**；差距集中在三处：**Word 的原生公式/脚注/版面收尾能力、Excel 的数据操作层（清洗/透视/成品感）、PPT 的成品版式层（主题 + 22 种设计版式）**。此外它是插件、有客户端 UI，所以「迷你卡片/逐字写入/静默模式」这类体验项按形态不适用于 skill。

| 领域 | 我方已有/更优 | 半有 | 没有（值得补的） | 本次落地 |
|---|---|---|---|---|
| Word | 修订/批注/题注/三线表/表头跨页/目录/引用/单位上标/审计/快照 diff | 图片尺寸、表格跨页 | **原生公式、原生脚注、fill_gaps 消空白、流程图、KeepTogether/不拆词两处实现缺陷** | 公式 + 脚注 + render 总览（见 §5） |
| Excel | 原子批写+锚点+回滚、快照/选择性还原、批注、查找/选区、多区域排版 | 大表概况（缺 LOOK/日期/错误值）、图表（缺组合图/次轴/标签） | **清洗套件 8 个 op、透视表、冻结/条件格式/数据验证/表对象、错误格回报** | 概况增强 + 数据操作套件 + 透视 + 成品感四件套 + 图表次轴（§5） |
| PPT | 查找替换、批注、快照/回滚、check 规则库、套模板、拆分合并、导出、接管标记 | 预设版式（仅 5 个骨架）、图表（3 系列限制） | **图片套版裁切、SmartArt、连接线框图、成品版式层、背景纹理、切换动画** | 裁切/套图 + SmartArt + 连接线 + 图表类型（§5） |
| 基础设施 | csc.exe 运行时编译 + 哈希缓存、常驻 helper/daemon | batch（失败继续 vs 首个失败即停）、参数告警只在 stderr | **Word/Excel 总览图、MessageFilter 忙重试、上下文成本（changelog 占 25%）** | changelog 拆出（已完成）+ 总览图 + 忙重试（§5） |

---

## 2. 研究方法与可信度

- 五路并行深读：Word / Excel / PPT / 基础设施 / 我方基线，逐条落到 `文件:行号`，产出本目录 `w0*.md`；对「WPS 到底支不支持」严格分三档：**已实测 / 据代码推断 / 需再测**。
- dsh-office 的 `docs/verification.md`（1075 行、28 版）是最有价值的工程资产：每一节记「环境 → 起因（用户原话 + 倒查）→ 改动 → 实测（定点复现 + 回归 + 真实模型对照组）→ **没有验证的**」。它的 WPS 适配结论可以采信（如「WPS 12.1.0.28505 上 4 张透视表/3 张图表与微软结果一致」）。
- 关键未知项由 Lead 用**临时文档**在 WPS 12 上实测（脚本见 `D:\DSH工作区\_research\probe\`，全部只碰 `%TEMP%`），结论见 §3。**不碰用户已打开的文档**。

---

## 3. WPS 实测台账（2026-11-02，WPS 12.0/12.1.0.29161）

### 3.1 Word（`KWPS.Application`）
| 能力 | 结论 | 备注 |
|---|---|---|
| `Footnotes.Add(range,'',text)` | ✅ 可行 | 计数与文本可回读；`Document.Footnotes` 与 `Range.Footnotes` 索引口径不同 |
| `FootnoteOptions.NumberingRule` | ❌ 属性不存在 | 「每节/每页重编号」控制不了，只能保留文档原设置 |
| `Document.OMaths.Add(range)` | ❌ 报「值不在预期范围内」 | 不能靠它把线性文本转公式 |
| `Range.InsertXML(<裸 m:oMathPara>)` | ❌ 调用成功但 `OMaths.Count` 仍 0 | 会被当普通内容 |
| **`Range.InsertXML(Flat OPC 整包)`** | ✅ **能生成原生公式** | `pkg:package` + `/_rels/.rels` + `/word/document.xml`，`OMaths` 0→1，可回读文本；**这是唯一可行路径** |
| `Fields.Add(...,"SEQ 表 \\* ARABIC")` | ✅ | 题注域可继续用 |
| `Fields.Add(...,"EQ ...")` | ❌ 结果为空 | 别用 EQ 域做公式 |
| `Shapes.AddShape` / `Shapes.AddConnector` | ✅ / ❌ | Writer 里连接线 `BeginConnect` 报 `E_FAIL` ⇒ 流程图只能用形状 + 直角折线 + 箭头 |
| `Application.UndoRecord.Start/EndCustomRecord` | ⚠️ 调用不报错 | 整批一次撤销「需再证」（我方已有 `undo -Count N` 兜底） |

### 3.2 Excel（`KET.Application`）
| 能力 | 结论 | 备注 |
|---|---|---|
| `PivotCaches().Create(1,range)` | ✅ **两参可用** | **三参/四参报「找不到 Create 的重载」**（dsh-office 用 4 参，PS/后期绑定下不行）；`CreatePivotTable(dest,name)` 两参可用；`AddDataField(f,caption,-4157)`、`TableRange1` 读值正确 |
| `ws.Sort`（SortFields+SetRange+Header+Apply） | ✅ | 降序实测生效 |
| `FormatConditions.Add(1,5,..)` / `AddColorScale(3)` / `AddDatabar()` / `AddTop10()` | ✅ 全部可用 | 条件格式不必降级 |
| `Validation.Add(3,1,1,'a,b,c')` | ✅ | 下拉列表可做 |
| `ListObjects.Add(1,range,Missing,1)` | ✅ | 表对象可做（默认名「表1」） |
| `Windows(1).FreezePanes` | ✅ | 需先 `Activate()` |
| `SpecialCells(12)` / `SpecialCells(-4123)` | ✅ 调用不报错 | 语义仍需逐项复核 |
| `Range.NumberFormat` | ⚠️ 单列返回 String，**整块多列返回 Null** | 日期识别必须按列取 |
| 组合图（`Series.ChartType` + `AxisGroup=2`） | ✅ | 系列建完再逐条改；`Axes(2,2)` 读取偶发 Null，需 try/catch |
| 错误格识别 | ✅ | `Value2` 为负整数错误码（如 `#DIV/0!` = `-2146826281`），可按码表映射 |

### 3.3 PPT（`KWPP.Application`，960×540）
| 能力 | 结论 | 备注 |
|---|---|---|
| `Shapes.AddConnector` + `BeginConnect/EndConnect` + 箭头 | ✅（两端都读回 -1） | **框图连线可做** |
| `SmartArtLayouts` / `Shapes.AddSmartArt` | ✅（134 种版式） | 组织结构图/流程/循环/维恩可做 |
| `Shape.PictureFormat.CropLeft/Top/Right/Bottom` | ✅ 可读写（比例 0–1） | `Shape.CropLeft` 直接访问不存在 |
| `Background.Fill`：`Solid` / `OneColorGradient` / `UserPicture` | ✅ | 满版图/渐变/纹理（先生成 PNG 再铺）都能做 |
| `SlideShowTransition.EntryEffect`、`TimeLine.MainSequence.AddEffect` | ⚠️ 属性可写、返回正常 | **是否真播未验证**（需要在放映里看），不能据属性成功就宣称生效 |
| `Shape.Adjustments` | ✅ 可写 | 但 dsh-office 记录 WPS 圆角实际被固定在短边 1/6 ⇒ 需要真目检 |
| `Application.Visible=$false` | ❌ `E_FAIL` | WPS 演示必须可见；`Presentations.Add()` 无参 |

---

## 4. 借鉴清单（按域，P0 = 立刻能省下手工活 / 堵住已知故障）

> 每条后面括号里是 dsh-office 的实现位置，便于回查。

### 4.1 Word
- **P0 原生公式**（Flat OPC + OMML；`helper/OfficeHelper.cs:1663-1769` + `src/omml.ts` 是 LaTeX→OMML 转换器）——水利报告公式（曼宁、P-Ⅲ、堰流）现在只能贴图或纯文本。
- **P0 原生脚注**（`OfficeHelper.cs` footnote 段、`README.md:104-118`）——校审/论文体例；`[fn:引注]` 内联写法可借鉴，编号规则 WPS 做不到。
- **P0 `fill_gaps` 消空白**（`:2967-3016`）：图表放不下整块移到下一页后，把后面的正文提上来填空白（不越过标题、冒号连带、空档阈值 `max(90pt,12% 页高)`）。WPS 卡 `Information(6)` 与 `Range.FormattedText` 两个原语，**降级方案：只报不改也有价值**。
- **P0 两处我方实现缺陷**（不是抄功能，是照它踩过的坑修）：① `helper/word/WpsHelper.cs:632` 的 `WordWrap = v?0:-1` **疑似极性取反**（微软侧结论 `-1` 才是整词），且只在 `format` 传参时生效；② `Tables.cs:975` 用 `Rows.Item().Range` 设 `KeepWithNext`，而 `Tables2.cs:5` 自己注明「WPS 会抛 `0x800A1767`」、异常被吞 ⇒ 报成功但没生效。
- **P0 脚注/图片段落的安全护栏**（`:3193`、`:1621-1649`）：含脚注的段落禁止整段 `set_text`；含公式/图片的段落自动给「最小值」行距，否则被固定行距裁切。
- **P1 `insert_diagram`**（`:3786-3793`、`:4047-4058`）：Word 原生形状流程图；WPS Writer 无连接线 ⇒ 用「形状 + 直角折线 + 箭头」。
- **P2 未知字段点名进返回值**（`:9669-9670`）：我方只有 stderr 的 `param-check` 告警，写进 JSON 模型才能自纠。

### 4.2 Excel
- **P0 大表概况的 LOOK 段 + 日期识别 + 错误格回报**（`:4511-4614`、`:4486-4501`、`:4153-4169`、`:4690-4711`）：写成文本的数字/日期、多余空格、类型混杂、空行、完全重复行、离群值（q3+3·IQR 且命中 ≤5% 才报）。**历年水文资料最脏的正是这几类**。
- **P1 清洗套件 8 个 op**（`clean`/`replace`/`fill_blanks`/`remove_duplicates`/`delete_rows{where}`/`sort`/`filter`/`copy_range`；`:5045-5244`、`:4928-4972`、`:4661-4680`、`:5298-5321`）：`¥1,299`/`89元`/`10%`→数字、`2025/3/1`/`20250301`/`3月1日`→日期；**`copy_range` 必须用 Excel 自己 Copy**（整块写回会重解析文本）；**`replace` 别用 `Range.Replace`**（WPS 会二次替换，`P`→`P-` 得 `P--`）。
- **P1 透视表**（`:5323-5412`）：频率统计/年月汇总；WPS 上用两参 Create，结果直读 `TableRange1`。
- **P1 图表组合图/次轴/标签/图例/占地区域与重叠提示**（`:4755-4875`）+ **WPS 标题被拒重试**（`:4736-4753`）：过程线+雨量柱是水文刚需。
- **P2 成品感四件套**（`:4973-5065`、`:5413-5500`）：冻结窗格、条件格式、数据验证、表对象（均已实测可行）。
- **不抄**：`Chart.Export` 单图落盘（本机落盘即密文）、「区域渲染自动含图表」的假设（WPS 不成立，我方 `-IncludeCharts` 才对）、「Excel 里数据活都让模型写代码」的定位（它是通用插件，我们的价值就是当场协作）。

### 4.3 PPT
- **P0 成品版式层**（`Theme`/`Page` 类 + `src/tools.ts:148-206`）：22 种设计版式（封面/目录/章节 4 式/要点/卡片/数据/图表/表格/公式卡/定理/图组/总结/对开/code/grid 行列组合/canvas 自由摆放）+ 主题 token 三层（7 色 + 卡片/圆角/符号 + 章节页样式）；**这是与我们差距最大的一块**——我们只有 5 个骨架预设（`WppHelper.cs:1246` 注释自认「借鉴但故意简化」）。
- **P0 图片套版裁切**（`:6011` 一带）：满版铺图按目标框比例裁，`focus: top` 默认保留上 1/5（人像头所在），裁掉过多要点名提醒。
- **P0 SmartArt / 连接线框图**：结构图、流程图、组织架构不用再拿图片顶。
- **P1 形状绘制图表**（`:9196`）：散点最小二乘 + 拟合线 + 方程 + R²、函数曲线（末端标线名、参考线、标注）、饼/圆环；表格单元格里的真公式。
- **P1 背景纹理**：先生成一张 1600×900 PNG 再 `UserPicture`（绕开 COM 渐变/纹理填充），旧纸张/网格/圆点/图片淡化铺底。
- **P1 主题存进 pptx `Tags`**（跟文件走，换机不丢；我方现在存 `%LOCALAPPDATA%`）、**页码用普通文本框写「01/12」**（绕开 WPS 占位符每页显示「1」的坑）。
- **P2 切换/入场动画**：属性可写但**实效未证**，落文档时必须标「未验证」。
- **不抄**：迷你卡片/逐字写入/静默模式（插件形态 + 用户已否决浮层）、`Slide.Export`（本机密文）。

### 4.4 基础设施
- **P0 Word/Excel `render -Overview` 总览图**（拼图算法 `OfficeHelper.cs:9756-9812`，宿主上限 `src/tools.ts:269-280`：6 页/张、一次 ≤18 页、cell 宽固定 640、浅灰底 + 「page N」+ 灰边框）。**不需要新取图路径**：复用我方 `WpsHelper.cs:1593-1616` 的 `SaveEmfAsPng`（已含 `factor=max(2,ceil(3000/w))` 超采样，正好保住细表格线）逐页出小图再拼；必须先切 `View.Type=3` 再信 `Pages.Count`；总览 cell 不能跟随 `-Width`（18 页会变十几 MP）。
- **P0 MessageFilter 忙重试**（`OfficeHelper.cs:115-135`，约 30 行）：`RetryRejectedCall → dwRejectType==2 && dwTickCount<6000 ? 150 : -1`，在 `[STAThread] Main` 里、任何 COM 调用前注册一次即可覆盖全部命令 ⇒ 用户正打字/开着模态框时从「180 s TIMEOUT」变成「≤6 s 自动通过或 `BUSY`」。
- **P0 上下文成本**：它常驻提示词只有 5,407 字符，三端操作清单按 app **首次使用时**随结果下发一次（`src/tools.ts:221-231`）；我方 `SKILL.md` 原 77,317 字符、其中版本 changelog 19,721 字符 = 25.5%（**已拆到 `reference/changelog.md`，本次完成**）。per-app 拆分**不做**：skill 没有「首次注入」钩子，拆成文件也一样进上下文，只在整场完全不用该 app 时才省。
- **P1 其它三条**：① 把「参数被忽略」写进 JSON 返回值（模型才能自纠，而不只是 stderr）；② 孤儿隐藏实例清扫（只杀 `MainWindowHandle==0` 的已知映像名 —— 正好避开我们「不按映像名批量杀 pythonw」的禁忌）；③ Excel `render` 后恢复用户剪贴板（我方现在直接 `Clear()`，会吞掉用户复制的内容）。
- **不抄**：在 `tools/execute` 阶段拦技能（无效，只有 `post-execute` 的 accept 有用 —— 它 `verification.md:1003-1011` 最贵的教训）。

---

## 5. 本次落地（v1.29，全部已实测）

三路并行实现 + Lead 亲自端到端验收（34 步，32 步一次通过，2 步"失败"是测试脚本自身问题，改正后复跑通过）。命令用法已并入 `SKILL.md`，逐条实测细节在 `reference/word-quirks.md`、`reference/excel-quirks.md`、`reference/ppt-quirks.md`。

| 域 | 落地内容 | 关键实测结论（WPS 12.1.0.29161） | 未做 / 边界 |
|---|---|---|---|
| Word | `formula`（LaTeX→原生 OMML）、`footnote`（list/insert/edit/delete/numbering）、`render -Overview`（6 页/张、≤18 页） | **WPS 只认 Flat OPC 整包**：裸 `<m:oMathPara>` 插完 `OMaths.Count=0`，`OMaths.Add(range)` 报「值不在预期范围内」；`Footnotes.Add` 可用；`Document.Footnotes.NumberingRule` 可读写（`FootnoteOptions` 上的不存在） | 脚注内嵌公式、公式字体字号参数、`\tag` 自动重编号（是手写数字非 SEQ 域）、`fill_gaps`、`insert_diagram` 未做 |
| Excel | 13 条新命令（`clean`/`fillblanks`/`dedupe`/`deletewhere`/`sort`/`filter`/`copy`/`addcol`/`pivot`/`freeze`/`cformat`/`validate`/`table`）+ `read -Profile` 的 LOOK/日期/错误码 + `fill` 错误格回报 + `chart` 组合图次轴/标签/图例/占地区域与重叠提示 + `render -Overview` | 透视表 `Create` **只能两参**；`Sort`/`FormatConditions`（含色阶/数据条/前N）/`Validation`/`ListObjects`/`FreezePanes` 全部可用；`Range.NumberFormat` **整块多列返回 Null**（必须按列取）；组合图标题被拒需重试 | `replace` 命令、跨簿 `copy`、`validate` 区域源、`pivot` 日期分组未实测；ListObject 内 `AutoFilter` 报 `E_ACCESSDENIED`（已转明确报错） |
| PPT | `crop`（focus 套版）、`connect`（框图连线）、`smartart`、`bg`（纹理/背景）、`anim`（切换/入场）、`chart` 趋势线/雷达/标签/图例 | `AddConnector`+`BeginConnect/EndConnect` 可用；`AddSmartArt` 可建**但节点写不进**（`AllNodes` E_FAIL）；裁切 `CropLeft` **单位是点**（比例会被存下但不渲染）、裁完改 W/H 会洗掉裁切；`OneColorGradient` 回读不可信 | SmartArt 节点文字、`radar` 真渲染（名义支持）、`-Arrow both` 起始箭头未生效——三项都是 WPS 硬限制，已如实标注而非假称成功 |
| 基础设施 | v1.1–v1.28 changelog 拆出（主文档少读 1.93 万字符）、`publish.ps1` 白名单补 `changelog.md`/`ppt-quirks.md`、新增"daemon 握死引用"故障排查条目 | 杀掉 WPS 进程后 daemon 会握着死 COM 引用 ⇒ 每条命令报 `RPC 服务器不可用`，`serve -Action stop` 后自愈 | MessageFilter 忙重试（~30 行，P0 未做）、参数忽略写进 JSON 返回值、孤儿隐藏实例清扫、Excel render 后恢复用户剪贴板、Word `fill_gaps`/`insert_diagram`、PPT 成品版式层（22 kind + 主题 token，最大差距，估 4–6 人日） |

**下一步（按价值/成本排序，均已在各域报告 §3 里给过依据）**：① ~~Word `fill_gaps` + 两处实现缺陷~~ → **v1.30 已完成**（`fillgaps` 新命令 + `noSplitWords` 极性修正 + `table -KeepTogether` 逐行回读；细节见 `docs/borrow/skill-word-v130.md`）；② Excel `replace` + `pivot` 日期分组实测；③ PPT 成品版式层（把 dsh-office 的 22 种版式 + 主题 token 移植成 `slide -Preset` 新档）；④ 基础设施 MessageFilter + render 剪贴板恢复。

---

## 6. 报告的用法

| 想看什么 | 看哪 |
|---|---|
| 某个 op 在 dsh-office 里的**字段与语义** | `w01-word.md` / `w02-excel.md` / `w03-ppt.md` 的 §1 |
| 我方**现在有没有**、差在哪 | 三份域报告的 §2 |
| **为什么**先做这几条（价值/工作量/风险） | 三份域报告 §3–§4 |
| WPS 到底支不支持某 API | §3 本表 + `w0*.md` 的 §4 支持度表 |
| 二次开发怎么接（改哪几处） | `w00-baseline.md` §3（新增命令 checklist，含行号） |
| 怎么验收（环境/起因/改动/实测/未验证） | `w04-infra.md` §4（从它 1075 行验收记录提炼的模板） |
| 原始实测脚本 | `D:\DSH工作区\_research\probe\probe-{excel,word,ppt}*.ps1` |
| 原始仓库 | `D:\DSH工作区\_research\dsh-office`（浅克隆，v1.2.0） |

> 维护纪律：本目录是**研究快照**，不是使用文档。落地后的用法写进 `SKILL.md` 与 `reference/`；研究结论若被后续实测推翻，就地加「⚠️ 已推翻：…」并注明日期，不要静默改写。
