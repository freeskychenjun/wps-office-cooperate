# 跨 app 多任务编排：一个会话里同时处理 Word / Excel / PPT

> 适用：一个会话里同时有 Word 报告、Excel 台账、PPT 汇报稿的任务。目标不是"把 COM 并行起来"（做不到），而是**让用户一次布置、任务自动排队、写者永远只有一个**。

---

## 1. 物理约束（决定调度规则）

| 约束 | 表现 | 调度含义 |
|---|---|---|
| 同一文档只能有一个写者 | Word 段号漂移、Excel 区间互相覆盖、PPT 页号漂移 | 同一文档的任务**必须串行**（按提交顺序） |
| WPS COM 是全局串行通道 | 轻命令可重叠（3 路 status 1.9×），重命令并发反而更慢（0.57×） | 跨文档可以"并行规划"，但**真并行度上限 2–3**，别指望线性加速 |
| 剪贴板共享 | Excel `render` 与 PPT `render` 并发会串图 | 两端 render 由 helper 排队（`Local\woc-clipboard`），不需要调度器处理 |
| 一个 ET 实例只能开一个工作簿 | 第二个 `open` 会挂起 | 多工作簿任务**必须串行**（close → open） |
| 每端一个常驻 helper | 同一 app 的请求按到达顺序排队 | 不需要额外锁；调度器只管"别让两个写者同时进同一文档" |

---

## 2. 任务板（board.json）

```json
{"tasks": [
  {"id": "T1", "app": "word",  "doc": "D:\\报告.docx", "kind": "write", "priority": 10},
  {"id": "T2", "app": "excel", "doc": "D:\\台账.xlsx",  "kind": "write", "priority": 10},
  {"id": "T3", "app": "word",  "doc": "D:\\报告.docx",  "kind": "write", "priority": 20, "dependsOn": ["T1"]},
  {"id": "T4", "app": "ppt",   "doc": "D:\\汇报.pptx",  "kind": "write", "dependsOn": ["T1"]}
]}
```

| 字段 | 说明 |
|---|---|
| `app` | `word` / `excel` / `ppt` |
| `doc` | 文档绝对路径（资源键 = app + 路径） |
| `kind` | `write`（默认）/ `read`（只读任务；仍与同文档写任务串行，避免读到半成品） |
| `priority` | 数字（小=先），同文档多任务按板内顺序 + 该值 |
| `dependsOn` | 显式依赖（跨文档的数据依赖，比如"Excel 出图后再贴进 PPT"） |

**隐式规则**：同一 `doc` 的任务自动串成顺序链，不需要手写 `dependsOn`。

---

## 3. 调度命令

```powershell
& <skill>\scripts\tasks-session.ps1 check -Board board.json
& <skill>\scripts\tasks-session.ps1 plan  -Board board.json -Out schedule.json [-MaxParallel 3]
```

`plan` 输出**波次（wave）**：

```
wave 1 (parallel): T1:word, T2:excel, T5:excel
wave 2 (parallel): T3:word, T4:ppt
```

- 同一波内**不同文档**：可以同时开工；
- 波与波之间：上一波全部完成，或该任务的 `dependsOn` 完成；
- 同文档的第二个任务一定在后面的波；
- 波内任务数超过 `-MaxParallel`（默认 3）会给 warning——建议拆子批。

---

## 4. 每个任务内部仍是"并行规划、单写者提交"

调度器只决定**波次**；每个任务内部按各 app 的工具做：

| app | 工具 | 参考 |
|---|---|---|
| Word | `tasks.ps1`（ranges/split/merge/plan） | `reference/multi-task.md` |
| Excel | `tasks-excel.ps1`（split/merge/plan） | `reference/multi-task-excel.md` |
| PPT | `tasks-ppt.ps1`（split/merge/plan） | `reference/multi-task-ppt.md` |

即：一个波里的 Word 任务和 Excel 任务可以各自派子代理并行规划，然后**串行**提交（各写各的文档，互不影响）。

---

## 5. 例：一份报告 + 一份台账 + 一份汇报

| 任务 | 内容 | 波次 |
|---|---|---|
| T1 | 报告第 3 章润色（Word，修订） | 1 |
| T2 | 台账补 F 列公式（Excel） | 1 |
| T3 | 报告全文单位上标（Word，同文档） | 2 |
| T4 | 汇报稿换 logo + 统一字体（PPT） | 2（依赖 T1 的结论页？按需） |
| T5 | 汇总报告/台账结果写进汇报（跨文档依赖） | 3 |

执行：波 1 同时跑 T1（Word）与 T2（Excel）；波 2 跑 T3、T4；波 3 收尾。每波结束汇报一次，全部结束给总表，**不 save**，等用户裁决。

---

## 6. 汇报与失败处理

- 每波开始前报"这一波会动哪几份文档"；每波结束报"改了什么 / 待裁决"。
- 某任务失败：**不影响其它波里已完成的任务**；该任务标记 failed，后续依赖它的任务跳过并说明。
- 用户中途追加任务：进板后重新 `plan`（已有任务的波次尽量不变），把新任务排进对应波或新增波。
- 中断/换会话：板与波次结果都在 `board.json`/`schedule.json` 里，可恢复。

---

## 7. 验收清单

- [ ] `tasks-session.ps1 check` 无环、无未知依赖
- [ ] 同一文档的任务不在同一波
- [ ] 波内任务数 ≤ 3（或已拆子批）
- [ ] 每波提交前对目标文档 `snapshot`
- [ ] 波次之间核对 `diff`；结构类任务单独跑并重读
- [ ] 全部结束未 `save`
