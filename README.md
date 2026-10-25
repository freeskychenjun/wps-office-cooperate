# wps-office-cooperate

在 WPS 里**实时**改 Office 文档的 skill：用户开着窗口，agent 通过 COM 直接改同一份活文档，所有改动当场可见。

- **Word（WPS Writer，`KWPS.Application`）**：修订模式逐段改写（可逐条接受/拒绝）、批注、页面渲染核对、排版、表格/题注/图片/目录/参考文献、快照对比、undo、交接单。
- **Excel（WPS 表格，`KET.Application`）**：批量写值/公式（锚点校验、原子批、失败回滚）、整列 fill、单元格批注、区域渲染、快照/对比/一键回滚、图表、行列表结构编辑。
- PPT（WPS 演示）规划中。

由 `wps-word-cooperate`（v0.8）与 `wps-excel-cooperate`（v0.6）于 2026-10-25 合并而成，命令与行为不变。

## 入口

```
powershell -NoProfile -ExecutionPolicy Bypass -File <skill>\scripts\office.ps1 word|excel <命令> [参数]
```

`word.ps1` / `excel.ps1` 也可直接调用。输出一段 JSON：`{"ok":true,"result":{...}}` 或 `{"ok":false,"error":{"code":"...","message":"..."}}`。

环境：Windows + WPS（没装微软 Office 也行）。首次调用用系统 `csc.exe` 编译 helper（约 1 秒，按源码哈希缓存）。

## 目录

```
SKILL.md               唯一权威说明（速查 + 铁律）
scripts/office.ps1     app 分发入口（word|excel|ppt）
scripts/word.ps1       Word 入口（原 wps.ps1）
scripts/excel.ps1      Excel 入口（原 et.ps1）
scripts/helper/word/   Word helper C# 源码（编译缓存按哈希）
scripts/helper/excel/  Excel helper C# 源码
reference/             workflows / quirks / encryption / acceptance
publish.ps1            发布到部署目录（不带 tests/docs/tools）
```

## 发布

```
.\publish.ps1            # 部署到 C:\Users\<user>\.agents\skills\wps-office-cooperate
.\publish.ps1 -Target d  # 部署到别处
```

## 开发

- helper 是纯 C#5 单/多文件，改完直接跑 `word.ps1 status` / `excel.ps1 doctor`，源码哈希变了自动重编。
- 测试基线在 `tests/word`（验收脚本 + 语料）、`tests/excel/work`。
- 版本历史：[reference/word-acceptance.md](reference/word-acceptance.md)、[reference/excel-acceptance.md](reference/excel-acceptance.md)。
