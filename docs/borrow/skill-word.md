# SKILL.md 补丁（Word）：`formula` / `footnote` / `render -Overview`

> 本文件是 **SKILL.md 的补丁片段**（v1.29 计划）。按下面的位置贴进 `SKILL.md` 即可；
> 实测环境 WPS Office **12.1.0.29161**（`KWPS.Application`，COM 由 WPS 接管）。
> 实现细节与踩坑见 `reference/word-quirks.md` 末尾「2026-11-02 公式/脚注/总览实测」一节。

---

## 补丁 1：Word 命令速查表（`## Word 命令速查` 的表格）

**① 替换现有 `render` 行**（原第 106 行）：

```markdown
| `render` | `-Doc <路径> -Page 1 [-To 4] -Out <目录> [-Width 1100]`；**总览：`-Overview [-Page 1] [-To 18]`** | 页面 EMF → PNG（宽 320–2400），返回每页路径与尺寸。**`-Overview`（v1.29）= 联系表总览**：每张图 **6 页、3 列**、格宽 **640**、浅灰底 + 每格上方「第 N 页」，**一次最多 18 页**（超出返回 `capped:true` + `more`）；一次调用看清排版错位/半空页/图表出框，可疑页再用 `-Page n` 放大到 1100 复核 |
```

**② 在 `render` 行之后新增两行**：

```markdown
| `formula` | 单条：`-Doc -Para N [-Expect "<该段开头>"] -Latex "<LaTeX>" [-Display] [-Number "1"] [-Anchor "文字" -Where after\|before]`；**批量：`-Items <json>`**（`[{"para":N,"expect":"...","latex":"...","display":true,"number":"1"}]`） | **LaTeX → WPS 原生公式**（真 OMML，不是图片、不是线性文字）。行内式插到段末（或 `-Anchor` 指定处），`-Display` 独占一段并居中，`-Number`/`\tag{n}` 生成右对齐编号 `(n)`。**插入后校验 `OMaths.Count` 必须增加**，否则回滚并返回 `MATH_FAILED`；LaTeX 超出子集返回 `MATH_UNSUPPORTED`（点名不支持的命令），两者都**不留残留**。一次 ≤12 条 |
| `footnote` | `-Action list`；`insert -Para N [-Expect] [-Find "词" -Where after\|before] (-Text "…" \| -TextFile <txt>)`；`edit -Index N [-Expect "…"] (-Text\|-TextFile)`；`delete -Index N`；`numbering -Number continuous\|section\|page` | **原生脚注 CRUD**。脚注是真脚注（WPS 自己编号，引用标记 U+0002），`list` 给 `i/label/para/page/text`（`i` **永远是 `Document.Footnotes` 的序号**，不是 WPS 的 `Range.Footnotes` 序号）。`edit/delete` 建议带 `-Expect`（该脚注正文开头），不匹配报 `ANCHOR_MISMATCH` 而不是改错条 |
```

## 补丁 2：Word 典型工作流（`## Word 典型工作流` 末尾追加）

```markdown
13. **写公式（LaTeX，v1.29）**：`read` 定位段号 → `formula -Para N -Expect "<段首>" -Latex "..."`（行内式）；
    要编号的独立公式用 `-Display -Number "3"`。**中文/含 `^` `{}` `\` 的 LaTeX 一定用单引号**，或整批走 `-Items <json>`。
    插完 `read` 一眼（公式在 `text` 里显示为数学字符，另有 `omaths` 计数）→ `render` 抽一页看分式/根号/上下标是否成形。
14. **加脚注（v1.29）**：`footnote -Action insert -Para N -Expect "<段首>" [-Find "<锚词>" -Where after] -TextFile <txt>`（中文正文走文件）
    → `footnote -Action list` 拿 `i` → 需要改就 `edit -Index i -Expect "<原文开头>"`，要删 `delete -Index i`。
    **不要**用 `revise`/`insert` 伪造脚注（上标数字 + 页脚手写），那不是原生脚注，编号会随增删而错。
15. **交付前过全稿（v1.29）**：`render -Overview`（≤18 页/次）→ 看 3 张总览图 → 可疑页 `render -Page n -Width 1100` 放大 → 再 `save`。
```

## 补丁 3：Word 故障排查表（`## Word 故障排查` 追加）

```markdown
| `MATH_UNSUPPORTED` + 命令名 | 该 LaTeX 用了转换器子集之外的命令（如 `\substack`、`\mathop`、`\begin{array}{cc}` 的列格式串） | 换写法：矩阵用 `\begin{matrix}` / `\begin{pmatrix}` / `\begin{cases}`（不写列格式）；`\text{}` 放中文除外——**中文放 `\text{中文}` 会进方程但字体不保证**，报告类公式优先 ASCII |
| `MATH_FAILED`（"WPS inserted the XML but built no equation"） | WPS 没把包认成公式：包结构不完整 / 段落上下文特殊（如插在域代码里） | 已回滚，文档没脏。换段落下手，或先把该段从域/文本框里挪出来；仍失败就先插一个空段再 `formula -Para <空段>` |
| 脚注 `ANCHOR_MISMATCH` | `-Index` 指向的脚注正文与你给的 `-Expect` 不同（多半是前面插过/删过脚注，序号已移位） | 先 `footnote -Action list` 重取 `i` 再改；**不要**跳过后重试 |
| 脚注里要写公式 | 目前**不支持**（返回 `BAD_ARGS` 明说） | 把公式放正文用 `formula`；脚注里用 Unicode 字符（`x²`、`α ≤ β`） |
| `render -Overview` 只给 1 张图且 `capped:true` | 一次最多 18 页（3 张 × 6 页） | `-Page 19 -Overview` 接着看；**不要**一次要 60 页 |
```

## 补丁 4：铁律区的"公式/脚注"提醒（可选，放在 `## 铁律` 第 5 条后）

```markdown
- **公式与脚注是"原生对象"，不是文字（v1.29）**：`formula` 写的是真 OMML 公式、`footnote` 写的是真脚注——
  它们**不随 `revise` 的 `replace` 改字**（公式内部文字不能用 `find/replace` 改，脚注正文也不在正文段落里）。
  要改公式就删了重插（`formula` 无删除命令，用 `undo` 或用户手工删）；要改脚注一律 `footnote -Action edit`。
```

---

## 用法示例（实测可复制）

```powershell
# 行内公式：第 12 段末尾（先 read 拿 expect，防改错段）
& <skill>\scripts\woc.ps1 word formula -Doc X.docx -Para 12 -Expect "式中" -Latex 'E = mc^2'

# 独立成段 + 右对齐编号 (3)
& <skill>\scripts\woc.ps1 word formula -Doc X.docx -Para 12 -Expect "式中" `
    -Latex '\frac{Q}{A\sqrt{R}} = \frac{1}{n}' -Display -Number "3"

# 带上下限与矩阵（一次成功，返回 omaths/readBack）
& <skill>\scripts\woc.ps1 word formula -Doc X.docx -Para 20 -Latex '\sum_{i=1}^{n} x_i^2 = \begin{pmatrix} a & b \\ c & d \end{pmatrix}' -Display

# 批量（中文/特殊字符多时首选）：formula.json = {"items":[...]}
& <skill>\scripts\woc.ps1 word formula -Doc X.docx -Items formula.json

# 脚注：插在"泥沙"二字之后
& <skill>\scripts\woc.ps1 word footnote -Doc X.docx -Action insert -Para 33 -Expect "泥沙" `
    -Find "泥沙" -Where after -TextFile footnote-1.txt
& <skill>\scripts\woc.ps1 word footnote -Doc X.docx -Action list
& <skill>\scripts\woc.ps1 word footnote -Doc X.docx -Action edit -Index 1 -Expect "含沙量" -TextFile footnote-1b.txt
& <skill>\scripts\woc.ps1 word footnote -Doc X.docx -Action delete -Index 2

# 总览：≤18 页一次看完，1 张图 = 6 页
& <skill>\scripts\woc.ps1 word render -Doc X.docx -Overview -Page 1 -To 18 -Out .\shots
```

## 支持的 LaTeX 子集（`scripts/helper/word/Omml.cs`）

| 类别 | 支持 |
|---|---|
| 上下标 / 素数 | `x^2` `x_i` `x_i^2` `x^{n+1}` `f'` |
| 分式 / 根式 | `\frac{}{}` `\dfrac` `\tfrac` `\binom{}{}` `\sqrt{}` `\sqrt[n]{}` |
| 定界符 | `\left( \right)` `\left[ \right]` `\left\{ \right\}` `\left\| \right.`（`.` = 空） |
| 大型运算符 | `\sum` `\prod` `\int` `\iint` `\iiint` `\oint` `\bigcup` `\bigcap`（`_`/`^` 成上下限，自动吸收操作数） |
| 极限 / 函数名 | `\lim_{x \to 0}` `\limsup` `\max` `\min` `\sup` `\inf`；`\sin \cos \tan \log \ln \exp \det \arg` 等直立 |
| 希腊字母 / 关系符号 | `\alpha…\omega` `\Gamma…\Omega`；`\le \ge \ne \approx \equiv \pm \times \cdot \div \infty \partial \nabla \to \Rightarrow \in \subset \cup \cap \forall \exists` 等 |
| 重音 / 上下划线 | `\hat \bar \vec \tilde \dot \ddot \check \acute \grave \overline \underline` |
| 环境 | `\begin{matrix\|pmatrix\|bmatrix\|Bmatrix\|vmatrix\|Vmatrix}`、`\begin{cases}`、`\begin{aligned\|align\|gather\|split}`（`&` 分列、`\\` 换行） |
| 字体 / 文本 | `\mathrm` `\mathbf` `\mathit` `\mathcal` `\mathbb` `\text` `\operatorname` |
| 其他 | `\quad \qquad \, \; \:`、`\tag{n}`（编号，等价 `-Number`，转换前被剥离）、`\label`（忽略） |

**不支持（会明确报 `MATH_UNSUPPORTED` + 命令名）**：`\substack`、`\mathop`、`\stackrel`、`\overset/\underset`、
`\begin{array}{cc}`（列格式串）、`\left.` 之外的 `\genfrac`、`\boxed`、`\cancel`、`\color`、`\newcommand` 等。
**遇到不支持先降级写法**（分段函数用 `cases`、矩阵不写列格式），不要反复试。

## 未做 / 边界（写进 SKILL 时如实保留）

1. **脚注里的公式**不支持（`footnote` 见到 `$` 直接 `BAD_ARGS`）。dsh-office 走"脚注正文里再 `Mathify`"，
   我方未实现。
2. **公式的字体/字号/颜色**没有专门参数：要改就走 `format -Match` 命不中公式（公式不是普通文字），
   实际做法是整段 `format` 后再 `formula` 重插，或让用户手工调。
3. `\tag` 编号是**手写数字**（`(1)` 是普通文字），不是 SEQ 域：增删公式后不会自动重排。
   真实报告需要自动编号时，仍建议 `caption` 式的域方案（未实现）。
4. `render -Overview` 不做"跨节横向页混排"的特殊处理：页序用 `Panes.Item(1).Pages` 的全局页序，
   含横向节的稿子已实测页号连续（7 页样张），但**分节符很多的老稿**建议先 `render -Page n` 抽两页对号。
5. 多实例：本机连开若干命令会看到隐藏 `wps.exe /Automation` 实例变多（`GetActiveObject` 偶发失败→新建实例），
   收尾时 `close` 并在 `status`/`doctor` 里看 `processes.wps`；只清**无窗口**的自动化实例。
