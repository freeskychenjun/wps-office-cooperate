# PPT 多任务：并行规划、合并写回

> 适用：一份演示文稿里同时改多页（文字/替换/格式/图片/图表/批注），或"全局风格 + 逐页微调"这类组合。
> 一句话：**并行规划、单写者提交**；PPT 无修订层、`restore` 只还原文字，所以结构类改动要先做并单独快照。

---

## 1. 与 Word / Excel 多任务的关键差异

| | Word | Excel | PPT |
|---|---|---|---|
| 身份锚点 | 段落号 | 单元格地址 | **页号 + 形状名**（页号会因加/删/移页漂移） |
| 回滚 | 修订逐条 | `snapshot`/`restore`（含 `-Cells` 选择回滚） | `snapshot`/`restore` **只还原文字**（版式/新增形状不还原） |
| 全局任务 | `unit`/`font` | 整簿 `format` | `theme` / `master` / `font`：**一改波及全稿**，必须最先做 |
| 结构任务 | 插删段 | 插删行/列、改表名 | 加/删/移页、`split`/`merge`：**页号整体偏移** |

---

## 2. 四步模型

```
① split   ppt.ps1 read 按页区间取形状/文字 → slice-<id>.json（给规划子代理）
② 规划    子代理产出 task-*.json（不碰 WPS）
③ merge   按 (页, 形状) 合并、replace 规则重叠仲裁、批注/全局/结构分列 → merged.json
④ plan    生成 text/replace/comment-items + global/structure-ops + batch-ops.json
```

---

## 3. 任务卡（task-*.json）

```json
{
  "task": "P1",
  "priority": 10,
  "texts":    [{"slide": 3, "shape": "title", "expect": "原开头", "text": "新文字"}],
  "replaces": [{"find": "旧词", "replace": "新词", "scope": "all", "slides": "1-5"}],
  "formats":  [{"slide": 3, "shape": "title", "font": "黑体", "size": 28}],
  "comments": [{"slide": 3, "shape": "title", "text": "校审意见"}],
  "globals":  [{"kind": "font", "font": "宋体"}, {"kind": "master", "action": "footer", "text": "页脚"}],
  "structures": [{"kind": "slide", "action": "add", "preset": "title", "title": "新页"}],
  "pictures": [], "charts": []
}
```

| 字段 | 说明 |
|---|---|
| `priority` | 数字（小=先应用）。冲突时保留优先级更小者 |
| `texts[]` | 一页一个形状只能有一个最终文本；`expect` 是当前文字开头（可选但强烈建议） |
| `replaces[]` | 全稿查找替换规则；`slides` 支持 `1-5,8`，空/`all`=全稿 |
| `comments[]` | 锚到形状（批注文字里会自动带 `@形状名`） |
| `globals[]` | `font`（全稿字体）、`theme`（套主题）、`master`（页脚/版式） |
| `structures[]` | 加/删/移页等；**单独文件，最先跑，跑完重读再规划其余任务** |

---

## 4. 三条命令

```powershell
& <skill>\scripts\tasks-ppt.ps1 split -Doc <pptx> -Ranges ranges.json -Out <work>
& <skill>\scripts\tasks-ppt.ps1 merge -TaskFiles <work> -Out <work>
& <skill>\scripts\tasks-ppt.ps1 plan  -Merged <work>\merged.json -Out <work> -Doc <绝对路径> -BatchOps
& <skill>\scripts\ppt.ps1 batch -Ops <work>\batch-ops.json
```

**merge 的判定**

1. 同一 `(slide, shape)` 两条不同文本 → 冲突，保留优先级小者。
2. 两条 `replace` 规则 `find` 相同且页区间重叠 → 冲突，保留第一条（避免同一处替换两次）。
3. 同一 `(slide, shape)` 两条不同批注 → 冲突，保留第一条。
4. `globals` 按优先级排序；多个 `theme` 只最后生效（给 warning）。
5. `structures`/`pictures`/`charts` 原样传递，`plan` 单独输出文件；出现结构操作时警告"页号会漂移，先跑结构再重读"。

**plan 输出**：`text-items.json`、`replace-items.json`、`comment-items.json`、`structure-ops.json`/`picture-ops.json`/`chart-ops.json`、`apply-order.json`；`-BatchOps` 时 `batch-ops.json`（globals → text → replace → comment）。

---

## 5. 提交顺序与验证

```
snapshot → （结构：单独跑 + 重读）→ 全局(theme/master/font) → text → replace → 图片/图表 → 批注 → check → render -Overview
```

- **结构先去**：加/删/移页后所有页号变化，后置任务必须重读重规划。
- **全局先于局部**：`theme -Apply`、`font`、`master` 会把整稿刷一遍；先逐页改再套主题等于白改。
- `replace` 自带 `prereplace` 快照，可 `restore -Since prereplace` 一键退回；`text` 是整批原子（一条不过整批不改）。
- 验证：`diff -Since pre` 看页/形状文字变化；`check` 只读门禁；`render -Overview` 目检（每页 ~355 ms，不要边改边渲）。

---

## 6. 例：三页改文字 + 全稿字体 + 换 logo

| 任务 | 内容 | 产出 |
|---|---|---|
| P1 | 第 2、3 页标题润色 | `texts` 2 条 |
| P2 | 全稿字体统一宋体 | `globals` 1 条（font） |
| P3 | 全稿换 logo | `pictures` 1 条 |

合并提交：`font`（全局）→ `text`（逐页）→ `picture`（换图）→ `check` → `render -Overview`。

---

## 7. 坑（实测）

1. **页号漂移**：`slide -Action add/delete/move`、`split`/`merge` 之后，旧页号全部可能失效 → 结构任务单独跑 + 重读。
2. **`restore` 只还原文字**：版式、新增形状、图片不还原；结构/版式改动前必须 `snapshot` + 另存副本。
3. **全局任务顺序**：`theme`/`master`/`font` 先做，否则局部格式被重刷。
4. **`replace` 与 `text` 的先后**：同一处既替换又改文本会互相覆盖——合并时按 `(slide, shape)` 判冲突；无法判断时先 `text` 后 `replace`，并复读。
5. **图表必须有窗口**：`open -Show` 后再画；`render` 走剪贴板，与 Excel 渲染共享互斥锁（v1.21 起）。
6. 快照按**文档路径哈希**分目录（v1.21 起），同名不同目录不再串基线。

---

## 8. 验收清单

- [ ] `merge` conflicts 为空或逐条说明
- [ ] 结构操作单独跑过并重读
- [ ] 全局任务在逐页任务之前
- [ ] 提交前 `snapshot` 成功
- [ ] `diff` 与计划一致；`check` 无新增 blocker；`render -Overview` 目检过
