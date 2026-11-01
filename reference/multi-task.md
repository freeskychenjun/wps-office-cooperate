# 多任务协作（Word 报告）：并行规划、合并写回

> 适用：一个会话里同时布置多个文档任务——多章节、同一章多个任务、章节任务 + 全文任务。
> 一句话：**并行的是"规划"，写回只有一个写者**；同段多任务必须合并成"一段一条 item"。
> 工具：`scripts/tasks.ps1`（split / merge / plan / selftest），配合 `word.ps1` 的 `read` / `revise` / `comment` / `table` / `unit` / `diff`。

---

## 1. 四步模型

```
① 冻结    word.ps1 status → open → snapshot -Tag pre → read -Count 0 -SkipTableCells
② 并行规划  每个任务一个子代理（零 COM）：产出 task-*.json 的 ops 清单
③ 合并    tasks.ps1 merge：按段合并、冲突裁决、数字/相似度 guard
④ 提交    tasks.ps1 plan：分章 items（降序）+ comments；父代理单写者按序应用
```

**关键纪律**

- 子代理**不要**调任何 skill 命令（切片和原文都在手上）；需要看版式才允许只读命令。
- 写回只由父代理做；同一时刻只有一个写者，因此**不需要锁**。
- 提交顺序固定：**表格结构（改段数）→ 文本批 → 段落格式 → unit → 批注 → diff/render**。
- 多章节时按**章节从后往前**提交；同一批 `revise` 内部由 skill 自动按 para 降序。

---

## 2. 任务卡（task-*.json）

```json
{
  "task": "T1",
  "title": "第3章润色",
  "priority": 30,
  "textOps": [
    {"para": 134, "expect": "区域构造位于", "op": "replace", "find": "抚州凹陷南缘", "replace": "抚州凹陷南缘（复核）"},
    {"para": 135, "expect": "根据钻探揭露", "op": "set_text", "text": "……整段新文本……", "allowRewrite": true}
  ],
  "commentOps": [{"para": 134, "anchor": "区域构造", "text": "建议补充构造单元名称"}],
  "tableOps": [{"para": 123, "action": "caption", "table": 2}],
  "formatOps": [{"para": 132, "style": "标题 1"}],
  "findings": ["第20段 1467 m³/s 与附表3不一致"]
}
```

字段说明

| 字段 | 说明 |
|---|---|
| `task` | 任务 id（汇报与合并报告里用） |
| `priority` | 数字（小=先应用）。建议：查证/事实 10，润色 30，格式 90 |
| `textOps[].para` | WPS 段号（`read` 返回的 `i`，表格行也算段号） |
| `textOps[].expect` | 段首文字，逐字照抄；不符 = 任务作废（防段号漂移） |
| `textOps[].op` | `replace`（优先，保留 run 级格式）或 `set_text`（整段重写） |
| `allowDigits` | 允许改数字（默认禁止；查证改数字必须显式开） |
| `allowRewrite` | 允许整段大改（相似度低于 `-MinSim` 时默认拦下） |
| `commentOps[].anchor` | 批注锚定文字，必须在该段原文中出现 |
| `findings` | 不改文档的结论（汇报用） |

**一段一条 textOp**：同一任务对同一段只能给一条；多条由规划侧先合并好。

---

## 3. 四条命令

### ranges —— 从标题样式自动切章（先跑这个）

```powershell
& <skill>\scripts\tasks.ps1 ranges -Read read.json -Out ranges.json -Styles "<一级标题样式>,<二级标题样式>"
```

- 样式名是 WPS 里的**中文名**（如 `标题 1`），从 `read` 返回的 `style` 字段原样取；不确定就先 `read -Count 0` 看几个标题段的 `style`。
- 只按**第一个样式**切章（后面的样式只用于识别层级，不切）；每章 = 本章标题段 → 下一个章标题前一段。
- 目录项用 `-SkipIfContains "……,.."` 过滤（逗号分隔多个子串）；文档没有标题样式时用 `-Pattern "<正则>"` 兜底。
- 输出 `ranges.json`：`[{"id":"ch1","title":"1 工程概况","from":18,"to":66}, …]`，并打印每章标题方便核对。
- 实测（306 段测试报告）：自动检出 6 章，与手工切分完全一致；目录中的"标题样式"条目会被样式判定自然排除。

### split —— 把一次 read 切成章节片（给子代理）

```powershell
& <skill>\scripts\tasks.ps1 split -Read read.json -Ranges ranges.json -Out <work>
```

- `read.json`：`word.ps1 read -Doc X -Count 0 -SkipTableCells` 的完整输出（含 `result.items`）。
- `ranges.json`：`tasks.ps1 ranges` 生成，或手写
  `[{"id":"ch3","title":"工程地质","from":132,"to":180}, …]`
  - `from/to` 用 **WPS 段号**；章节从 `标题 1/标题 2` 段落取。
  - **排除目录项**：目录段落常是"正文"样式且命中数字正则，切分前先剔除。
- 输出：`chapter-<id>.json`（`{id,title,from,to,count,paras:[{i,text,style}]}`）+ `index.json`（含 overlaps/problems 校验）。

### merge —— 合并多任务、做 guard

```powershell
& <skill>\scripts\tasks.ps1 merge -Read read.json -TaskFiles <dir 或 a.json,b.json> -Out <work> [-MinSim 0.55]
```

做四件事：

1. 校验 `expect`（与原文比对）、`anchor`、`para` 是否存在（表格段不在 read 里 → 报冲突）。
2. 按 `para` 把多个任务的 op 合成一条：按 `priority` 顺序在原文上模拟执行，得到最终文本。
3. 生成**单条** `replace`：取最小改动区间，并**自动扩展到在段内唯一**（`revise` 只替换首个命中，短 find 会打错位置）；规划给的 `find` 若本身段内歧义，也会自动向左右扩上下文（连同 `replace` 一起扩），扩不出来才报冲突。
4. Guard：数字 token 多重集（默认必须不变）、相似度下限（`set_text` 大改需 `allowRewrite`）。

输出 `merged.json`：`{stats, textOps(降序), tableOps, formatOps, comments, findings, conflicts}`。
**有冲突先解决再提交**——冲突项不会进 textOps。

### plan —— 拆成提交顺序，生成可执行的 items/批

```powershell
& <skill>\scripts\tasks.ps1 plan -Merged merged.json -Ranges ranges.json -Out <work> [-Doc <绝对路径>] [-BatchOps]
```

- 输出每章 `items-<id>.json`（已按 para 降序）、`comments.json`、`apply-order.json`（章节倒序）、
  以及 `-BatchOps` 时的 `batch-ops.json`（可直接 `word.ps1 batch -Ops`，**一次进程启动跑完提交**）。
- `-Doc` 建议给**绝对路径**（batch 模式下相对路径会按 helper 的工作目录解析）。

### selftest —— 结构自检

```powershell
& <skill>\scripts\tasks.ps1 selftest -Read read.json -Ranges ranges.json
```

---

## 4. 三种组合怎么排

### A. 多章节（第 3、5、7 章各润色）

1. 一次 read 全稿 → `split` 成 3 片 → 3 个子代理并行产出 task 文件。
2. `merge` → 每章一个 `items-<章>.json`。
3. 按 `apply-order`（**从后往前**）逐章 `revise`；或 `batch` 一次跑完。
4. `diff` / `render` 抽查，汇报，**等用户确认再 save**。

### B. 同一章多个任务（润色 + 单位 + 查证 + 表格 + 批注）

1. 任务卡各自出 ops（同一章、范围重叠）。
2. `merge` 按段合并：同一段被润色和查证同时改 → 合成一条（事实优先）。
3. 提交顺序：**表格结构 → 文本批 → 段落格式 → `unit -Apply` → 批注**。
   - 能折进 textOp 的单位修（`m3`→`m³` Unicode）直接折进去；run 级上标留给 `unit`。
4. 5 个任务通常只剩 3–4 次写调用。

### C. 章节任务 + 全文任务

| 全文任务 | 处理 |
|---|---|
| 只读侦察（audit/captions/法规核对/数字核对） | 并行跑，产出 findings + 需要改写的 textOps（与章节 op 合并） |
| 文本规范（术语/单位写法/地名统一） | **作为规则下发给章节规划**，或与章节 op 合并进同一批；不要两批各改一次 |
| 格式/结构（unit、字体、页码、表格题注） | 文本批之后；表格若改段数则最先跑并重读 |

顺序：**全文只读 → 章节文本 → 全文文本（合并）→ 全文格式 → 批注 → 复核**。

---

## 5. 汇报模板

```
收到 5 个任务：T1 第3章润色 / T2 全文单位上标 / T3 法规引用核对 / T4 表格题注 / T5 校审批注
排期：T2 在 T1 之后（避免上标被重写冲掉）；T4 最先（改段数）；T5 最后（按最终段号）
并行准备中：5 个规划子代理，零 COM……（约 2–4 分钟）
```

完成后两栏：**局部改动**（第 3 章 25 段）/ **全局改动**（全文 47 处单位、3 处引用），
再给 `diff` 摘要（changed/added/removed/revisions）与待裁决项。

---

## 6. 已知坑（都实测过）

1. **同一段不能放两条 item**：第 2 条报 `could not place the edit`，整批回滚（原子）。
2. **写回顺序必须从后往前**：插段会让后面段号 +1；升序会在下一批 `ANCHOR_MISSING`（被拦住，不会静默改错）。
3. **`set_text` 会丢段内 run 级格式**（上标、局部加粗）→ 优先 `replace`；整段重写后 `unit` 可补上标。
4. **`unit` 的脚本页守卫会误伤**：段落里出现 `P=1%` 这类内容可能被判成"模板/脚本示例"而跳过 → 跑完看 `skippedSamples`，误伤段落改走 textOp。
5. **表格行也是段号**：`read -SkipTableCells` 的 `i` 会跳号，直接用 `i` 切范围，不要重排。
6. **目录项会被当成章节**：切分前用样式 + "含 `……`" 过滤。
7. **`revise` 的 `replace` 只换首个命中**：合并器已把 find 扩到段内唯一；手写 items 时也要注意。
8. **批注在文本批之后**：文本改动会移动段内位置；批注 `anchor` 必须用改后文本核对。

---

## 7. 验收清单

- [ ] `tasks.ps1 selftest` PASS
- [ ] `merge` 的 `conflicts` 为空或已逐条说明
- [ ] 提交后 `diff.summary.changed` == 计划的段数（±0）
- [ ] `revisions` 数量与 `applied` 匹配（`set_text` ≈ 2×）
- [ ] `unit` 的 `skippedSamples` 已检查
- [ ] 未 `save`，等用户确认
- [ ] 汇报里给出局部/全局两栏与待裁决项
