# Excel 多任务：并行规划、合并写回

> 适用：一个会话里同时布置多个 Excel 任务——同一工作簿的多区域补公式、统一格式、加批注、做图表；或多个工作簿串行处理。
> 一句话：**并行的是"规划"，写回只有一个写者，且 Excel 没有修订层**——所以合并层比 Word 更保守（宁可报冲突，不猜）。

---

## 1. 与 Word 多任务的关键差异

| | Word | Excel |
|---|---|---|
| 回滚语义 | 修订可逐条接受/拒绝 | **无修订层**：COM 写入不进撤销栈；只能 `snapshot` + `restore`（还支持 `-Cells` 选择性回滚） |
| 锚点 | 段落号 + `expect` 段首 | 单元格/区域地址 + `expect` 当前显示文本 |
| 会漂移的 | 插删段 → 段号 +1 | 插删行/列、改表名/删表 → 后续**所有地址**变化 |
| 合并粒度 | 一段一条 item | 一个单元格只能有一个最终值；重叠区域必须仲裁 |

**结论**：结构类任务（`rows/cols/sheet`）最先做，做完**必须重新 read**，再提交值/公式批。

---

## 2. 四步模型

```
① split    Excel read -Range 逐块取数 → slice-<id>.json（给规划子代理，零 COM）
② 规划     子代理只产出 ops 清单（task-*.json），不碰 WPS
③ merge    按单元格/区域合并、冲突仲裁、expect 锚检查 → merged.json
④ plan     生成 writes/formats/comments + batch-ops.json；一次 batch 进程跑完
```

---

## 3. 任务卡（task-*.json）

```json
{
  "task": "E1",
  "title": "补 F 列公式",
  "priority": 10,
  "writes": [
    {"sheet": "数据", "cell": "F2", "formula": "=E2*2", "expect": "0"},
    {"sheet": "数据", "cell": "F3:F100", "values": [[1],[2]], "expect": ""}
  ],
  "formats": [
    {"sheet": "数据", "range": "A1:F1", "bold": true, "fillColor": "#D9E1F2"}
  ],
  "comments": [
    {"sheet": "数据", "cell": "F2", "text": "公式口径待确认"}
  ],
  "structures": [],
  "charts": []
}
```

| 字段 | 说明 |
|---|---|
| `priority` | 数字（小=先应用）。冲突时**保留优先级更小**的一条 |
| `writes[]` | 每条必须**恰好一个** `value` / `formula` / `values`；`cell` 可为单格或区域；`expect` 强烈建议（缺了会进 warnings） |
| `formats[]` | 与 `format -Items` 同名字段；一条可含多个区域（分号） |
| `comments[]` | 每格最多一条批注（Excel 限制） |
| `structures[]` | `rows/cols/sheet` 结构操作——**单独文件**，先应用再重读 |
| `charts[]` | 图表操作，排在 format 之后 |

---

## 4. 三条命令

```powershell
# ① 取数切片（每块一个 slice，交给一个规划子代理）
& <skill>\scripts\tasks-excel.ps1 split -Doc <xlsx> -Ranges ranges.json -Out <work>
#   ranges.json = [{"id":"r1","sheet":"Sheet1","range":"A1:F200"}, ...]

# ② 合并（校验 + 冲突仲裁 + 生成 merged.json）
& <skill>\scripts\tasks-excel.ps1 merge -TaskFiles <work> -Out <work>

# ③ 排序并生成可执行批（-Doc 给绝对路径；-BatchOps 生成 batch-ops.json）
& <skill>\scripts\tasks-excel.ps1 plan -Merged <work>\merged.json -Out <work> -Doc <绝对路径> -BatchOps

# ④ 提交（一次进程启动）
& <skill>\scripts\excel.ps1 batch -Ops <work>\batch-ops.json
```

**merge 的判定**

1. 同一 `(sheet, cell/rect)` 被两个任务写：值/公式相同 → 去重；不同 → **冲突**，保留优先级更小者并在 `merged.conflicts` 里写明。
2. 区域用矩形相交判定（`A1:B2` 与 `B2:C3` 算重叠），不只是比地址字符串。
3. `expect` 缺失 → warnings（盲写，建议补）。
4. 同一格两条不同批注 → 冲突，保留第一条。
5. `structures`/`charts` 原样传递，`plan` 单独输出 `structure-ops.json`/`chart-ops.json`。
6. 出现结构操作时 merge 会警告：**本计划里的地址是结构变更前的**，必须先跑结构、重读、再出值/公式计划。

**plan 的输出**：`writes.json`、`formats.json`、`comments.json`、`apply-order.json`（阶段与计数），`-BatchOps` 时再加 `batch-ops.json`（`write` → `format` → 每条 `comment` 一个 op）。

---

## 5. 提交与验证

```
snapshot → batch(write, format, comment) → calc → diff -Since last → render（可选）
```

- **写前 `snapshot`**：Excel 没有修订层，这是唯一可靠的后悔药；`write` 自己也会落 pre-batch 快照，但跨阶段的整批回滚要靠手动的这一份。
- `diff` 会列出每个变化格（before→after、公式联动），用它核对"计划内的格变了、计划外没变"。
- 写/格式失败时 `write` 批是原子的（整批一条不改）；`batch` **不跨 op 回滚**——上一 op 成功、下一 op 失败时，前面的已生效，务必用 `diff` 对账。

---

## 6. 例：同一工作簿三个任务

| 任务 | 内容 | 产出 |
|---|---|---|
| E1 | 补 F2:F136 公式 | `writes` 一条 `values`/`formula` 块 + `expect` |
| E2 | 表头格式 + 数字格式 | `formats` 两条 |
| E3 | 校审批注 5 条 | `comments` 5 条 |

合并后一次 `batch` 跑完：`write`（原子）→ `format` → 5 条 `comment`；随后 `calc` + `diff` 对账。

---

## 7. 坑（实测）

1. **一个 ET 实例只能可靠地开一个工作簿**：打开第二个工作簿会挂起（`Workbooks.Open` 不返回，实例变幽灵，需结束 `et.exe`；部署版同样复现）⇒ 多工作簿任务串行（`close` 再 `open`）。
2. **结构操作改地址**：`rows/cols insert/delete`、`sheet rename/delete` 之后，所有旧地址可能失效；必须重读。
3. **没有修订层**：不要向用户承诺 Ctrl+Z；写前快照、写后 diff。
4. 快照现在按**文档路径哈希**分目录（v1.21 起），同名不同目录不会再串基线。
5. `render` 走系统剪贴板，与 PPT 渲染共享 `Local\woc-clipboard` 互斥锁（v1.21 起）——并发渲染不会再串图。
6. `format -Borders` 只认 `grid/thick/none`；`chart` 数据必须同表；这些既有坑不变。

---

## 8. 验收清单

- [ ] `merge` 的 conflicts 为空或逐条说明
- [ ] 每条写都有 `expect`（或明确接受盲写告警）
- [ ] 提交前 `snapshot` 成功
- [ ] `diff` 的 cells 与计划一致（计划外无变化）
- [ ] 结构操作单独跑过并重读过
- [ ] 未 `save` 前先 `render`/`read` 核对
