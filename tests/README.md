# tests —— wps-office-cooperate 的测试与验收

本目录是仓库的唯一测试入口。测试**不是** pytest：沿用仓库原有约定——"断言脚本（Python / PowerShell）+
`run_*.ps1` 驱动器 + 编号断言输出"，一行命令汇总 PASS / FAIL / SKIP 与耗时。

## 两层：fast（离线）与 wps（需要 WPS）

| 层 | 需要什么 | 内容 | 用途 |
|---|---|---|---|
| `fast` | Python 3.12+、Windows PowerShell 5.1，**不需要 Office** | 夹具生成（python-docx / python-pptx）、docx 静态事实断言、PPT oracle 计算、`tasks*.ps1 selftest`、`publish -WhatIf` 冒烟、编码/语法门 | 本地秒级回归；GitHub Actions 只跑这一层 |
| `wps` | **WPS Office 12.1+**（Writer / ET / WPP），pywin32、Pillow；本机若透明加密还需解密工具 | `tests\**\run_*.ps1` 驱动器、三端 `check_tasks.ps1` 多任务回归 | 端到端 COM 验收，会真的启动 WPS 并改文档副本 |

## 快速开始

```powershell
pip install -r requirements-dev.txt

# 离线层（CI 模式，约 20 秒）
powershell -NoProfile -ExecutionPolicy Bypass -File tests\run_all.ps1 -Layer fast

# 全部（fast 先跑，再跑 wps 层；WPS 层需要 WPS 空闲、可能要十几分钟）
powershell -NoProfile -ExecutionPolicy Bypass -File tests\run_all.ps1 -Layer all

# 只跑名字里含 keyword 的测试（逗号分隔，大小写不敏感的子串匹配）
powershell -NoProfile -ExecutionPolicy Bypass -File tests\run_all.ps1 -Layer fast -Only encoding,normalize

# 只看有哪些测试，不执行
powershell -NoProfile -ExecutionPolicy Bypass -File tests\run_all.ps1 -List
```

输出示例（`-Layer fast`）：

```
PASS encoding                              3.1s  encoding: PASS - 27 .ps1 file(s) are pure ASCII
PASS ps1-syntax                            0.1s  27 .ps1 parsed
...
== summary: 10 passed, 0 failed, 0 skipped in 20.1s (layer fast)
```

退出码：全部通过（含 SKIP）为 `0`，任一 FAIL 为 `1`，没有匹配的测试为 `2`。CI/脚本可直接据此判定。

## fast 层包含什么

| 测试名 | 断言 |
|---|---|
| `encoding` | `publish.ps1` + `scripts\*.ps1` + `tests\**\*.ps1` 全部**纯 ASCII**（允许 UTF-8 BOM，满足 PS 5.1 无 BOM 会按 ANSI 解析的约束） |
| `ps1-syntax` | 上面这批 `.ps1` 在 PS 5.1 语法下都能 `ParseFile`（防止改坏包装器到调用时才爆） |
| `word-normalize-static` | 现场生成 `word table -Action normalize` 的夹具，再跑 `check_normalize.py --static-only`：14 张表 / 4 个"题注写在第 1 行" / 2 张缺 `w:tblBorders` / 0 修订 / 0 批注 |
| `ppt-fixtures` | `tests\ppt\make_*.py` 六个夹具生成器全部能跑（python-pptx 是 WPS 层的独立 oracle） |
| `ppt-oracle` | `check.py --list` 对 replace 夹具算出的期望命中数 = 16（oracle 与夹具没漂移） |
| `tasks-selftest` | `scripts\tasks.ps1 selftest`：章节区间覆盖、最小替换回环、数字守卫、相似度 |
| `tasks-excel-selftest` / `tasks-ppt-selftest` / `tasks-session-selftest` | 三端规划器的纯逻辑自测 |
| `publish-whatif` | `publish.ps1 -WhatIf` 走完整解析/校验并列出复制清单，**不落盘**；若 `publish.ps1` 还没有 `-WhatIf` 则 SKIP |

生成物统一落在 `tests\out-fast\`（gitignore），`tests\clean.ps1` 可清理。

## wps 层与自动发现

`run_all.ps1 -Layer wps` 会自动递归发现：

- `tests\**\run_*.ps1` —— 各端驱动器（`tests\word\run_acceptance.ps1`、`tests\ppt\run_*.ps1` 等）；
- `tests\**\check_tasks.ps1` —— 多任务协作（split / merge / plan / 仲裁）回归。

新增一个 WPS 端到端测试**不需要改 runner**：把驱动器命名成 `run_<topic>.ps1` 放进对应的
`tests\<app>\` 即可。约定：

- 驱动器只负责"造副本 → 通过 `scripts\woc.ps1`（或对应 app 包装器）调 skill → 把每个 JSON 响应落盘"；
- 断言集中在断言脚本（Python 用 python-pptx / python-docx / openpyxl 独立算期望值，见 `tests\ppt\check.py`）；
- 退出码 `0` = 通过，非 `0` = 失败（runner 只看退出码）；
- 绝不修改被跟踪的语料，只操作临时副本。

单独跑某个驱动器也可以，例如：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\ppt\run_replace.ps1
python tests\word\check_normalize.py --fixture "D:\path\to\real-report.docx"
```

### Word normalize 夹具（自给自足）

`tests\word\check_normalize.py` 以前默认指向仓库外的乐平报告副本，新克隆第一条 Word 测试就挂。
现在：

- 默认夹具是 `tests\word\out\normalize-fixture.docx`，由 `tests\word\make_normalize_fixture.py`
  现场生成（14 张表、4 个题注写在第 1 行、其中 12 张显式写 `w:tblBorders`、0 修订 0 批注），
  所以**不需要任何外部文件**就能跑静态断言；
- 合成夹具没有分页符 / 分节符 / 正文表引用，因此默认只跑静态部分（等价于 `--static-only`）；
  端到端 A1-A9 断言需要真实报告：`python tests\word\check_normalize.py --fixture <report.docx>`；
- `--static-only`、`--skip-edit`、`--verbose` 等开关保持可用。

## 环境依赖

`requirements-dev.txt` 是 Python 依赖清单（python-docx / python-pptx / openpyxl / Pillow / pywin32），
版本是下限不是锁版本。另外：

- **WPS Office 12.1**（本 skill 的目标宿主，不是 Microsoft Office）：只有 wps 层需要；版本差异会真实影响
  COM 行为，报告 bug 时请带上 WPS 版本。
- 本机若开了透明加密（落盘即密文，文件头 `17DA5FA01633CD9A`），PPT 的"文件级"断言需要一个解密工具：
  在 `tests\ppt\local.config.json`（不提交、gitignore）里写 `{"decryptEngine": "<FastDycrypt.exe 路径>"}`，
  中文路径不能写进 `.ps1`。
- CI 只跑 fast 层：`.github/workflows/tests.yml`（windows-latest + `requirements-dev.txt`）。

## 目录

| 目录 | 内容 |
|---|---|
| `tests\word\` | 语料生成（`make_corpus.py` / `make_edits.py` / `make_audit_fixture.py`）、`validate.py`、表格题注（`captions\`）、`table -Action normalize`（`check_normalize.py` + `make_normalize_fixture.py`） |
| `tests\excel\` | `check_hint.py` / `check_tray.py`（WPS 层）、`run_data.ps1` 数据层回归（由数据侧维护） |
| `tests\ppt\` | `make_*.py` 夹具生成器 + `check_*.py` 独立 oracle + `run_*.ps1` 驱动器（replace / check / deliver / master / chapters / chart） |
| `tests\multi-task\` | `ranges.json` 章节图 + `check_tasks.ps1`（split / merge / plan 回归） |
| `tests\out-fast\` | `run_all.ps1` 的生成物（gitignore） |

## 清理

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\clean.ps1 -DryRun   # 只打印
powershell -NoProfile -ExecutionPolicy Bypass -File tests\clean.ps1           # 真删
```

只清 `tests\**\out*`、`tests\**\__pycache__`、`.tmp\impl`；任何越界路径（不在 `tests\` 或 `.tmp\` 下）
直接中止，不会误删源码或部署目录。

## 编码规则（所有 `.ps1`）

Windows PowerShell 5.1 把无 BOM 的 `.ps1` 按 ANSI 读，中文字面量会变乱码、甚至报语法错。所以：

- **`.ps1` 必须纯 ASCII**（需要中文就放到 UTF-8 数据文件里，或通过运行时参数传）；
- 允许 UTF-8 BOM（`scripts\woc.ps1` 就是 BOM + ASCII），但 BOM 之后仍必须全部是 ASCII；
- `tests\check_encoding.ps1` 是这条规则的守门员，已接进 `run_all.ps1` 的 `encoding` 测试。
