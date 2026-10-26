# 落盘即密文：本机的透明加密与交付

本机（以及任何装了单位透明加密的机器）有一个会让"AI 改完的文件别人打不开"的环境事实：**WPS 写出的文件在磁盘上是密文**。本文只讲怎么认、怎么绕、怎么交付；加解密本身交给专用工具。

## 1. 现象：怎么一眼认出密文

| 检查 | 明文 | 本机密文 |
|---|---|---|
| 文件头 8 字节 | docx/pptx/xlsx 是 `50 4B 03 04`（`PK`） | `17 DA 5F A0 16 33 CD 9A` |
| `python zipfile` 打开 | 正常 | `BadZipFile: File is not a zip file` |
| `PyMuPDF` 打开导出的 PDF | 正常 | `code=7: no objects found` / `Failed to open file` |

自查一条命令（PowerShell 5.1 可用）：

```powershell
(Get-Content -Path "D:\报告\某报告.docx" -Encoding Byte -TotalCount 8 | ForEach-Object { '{0:X2}' -f $_ }) -join ' '
```

- 输出 `50 4B 03 04 …` → 明文，可直接发人。
- 输出 `17 DA 5F A0 16 33 CD 9A` → 密文，见第 3 节。

**实测凭证**：

- **2026-10-23 本机**：本 skill 的 `save` 保存出的 `.tmp\smoke.docx`（44,833 字节）头部即 `17 DA 5F A0 16 33 CD 9A`，`python zipfile` 报 `BadZipFile: File is not a zip file`；同一份文件在 WPS 里能正常打开、读写、渲染。
- **2026-10-21 本机**（[研] §2.3 ① 同期数据）：`ExportAsFixedFormat` 导出的 PDF 同样是密文——`isPlainPDF=False`、`/Type/Page` 命中 0，且密文比明文大 1024 字节。

**关键区分**：被加密的是**落盘的文件字节**，不是**屏幕上的文档**。WPS 打开、编辑、渲染、导出全都正常——只有"别的进程直接读这个文件"才失败。

## 2. 为什么 COM 路线不受影响

透明加密驱动（本机是 Agile DGS 一类）工作在文件系统层，按"**哪个进程在读/写**"决定是否解密/加密：

- `wps.exe` 在驱动白名单里：它**读**文件时驱动给明文，它**写**文件时驱动自动加密。
- 所以本 skill 的 `read` / `revise` / `comment` / `render` 全部在 COM 内部完成——数据从 WPS 进程里出来，根本不经过 `python`、`zipfile`、LibreOffice、PyMuPDF，也就碰不到"密文"这一层。
- 唯一落盘的动作是 `save`，而那是 WPS 自己写的 → 结果是密文（这是驱动的行为，不是 bug）。

反向推论（很重要）：**任何离线脚本路线**（python-docx / openpyxl / LibreOffice headless / fitz 渲 PDF）读这份文件都会失败，因为它不在白名单里，驱动不给它解密。这既解释了"为什么 COM 是这里的正确选择"，也解释了"为什么 AI 改完的文件同事打不开"。

另一条相关事实：`python-docx` 自己写出的 .docx **不受影响**（头是 `PK`），因为写盘的不是白名单进程——所以"有的文件能读、有的不能"不是随机故障，看是谁写的（既有记忆《本机文档处理环境坑》实测）。

## 3. 交付与转发：发之前必须解密

顺序固定，别跳步：

1. `save` 之后跑第 1 节的自查命令，确认头部。
2. 要发给同事 / 上传 / 归档前，解密。**首选本机自研的 `FastDycrypt`**（离线算法实现，不依赖白名单进程，~9 s/份，任意格式）：

   ```powershell
   # 就地解密（原密文在解密成功后立即删除，不留 .dgsbak）
   powershell -NoProfile -ExecutionPolicy Bypass -File `
     "D:\DSH工作区\FastDycrypter\FastDycrypt\engine\dec.ps1" `
     -Paths "<文件或目录>" -InPlace -DeleteBackupAfterSuccess -NoPause -NoLog
   ```
   - 资源管理器里右键 `FastDycrypt` 走的是同一引擎；批量/脚本化就用上面这条。
   - 默认只处理 Office/PDF/图片/CAD 扩展名；其它类型加 `-AllFiles`；目录递归加 `-Recurse`。
   - 日志默认追加到 `%LOCALAPPDATA%\FastDycrypt\decrypt.log`（`-NoLog` 关闭、`-LogFile` 改路径）。
   - 备选：**`decrypt-office-pdf`** 技能（`decrypt-files.ps1`，走 `D:\应用软件\右键解密\` 的伪装白名单程序，~43 s/份；伪装 `powerpnt.exe` 是单实例，残留实例会把下一次调用卡死）。
   - 两者的用法与参数以各自工具/技能的说明为准，本文不复述其内部实现。
   - **没有装这些工具的机器**（比如拿到本 skill 的同事电脑）：用单位统一的解密途径——本单位的透明加密通常配套右键解密菜单或解密工具（如 `D:\应用软件\右键解密\`）；再没有就联系单位 IT。**不要**因此尝试用脚本改写文件。
3. 解密后再跑一次自查：头部应恢复成 `PK`（PDF 恢复 `%PDF`）。**没恢复就是没成功**，别发。
4. 发给同事的文件用解密后的那份；源文件保持原状即可。

两条禁忌：

- **别手工"修"密文**：改文件头、改扩展名、用文本编辑器另存，只会把文件彻底弄坏。加密是整文件级的，只有专用解密工具能还原。
- **别把已解密的文件再丢回 WPS 存一遍**：会重新加密。要改就在 WPS 里改、改完再解密一次。

## 4. 什么时候完全不用管

同事机器上**没有**这套加密驱动时，一切正常：WPS 写出的就是明文，`python zipfile` 能读，直接发就行。判断方法就是第 1 节那三条——头部是 `PK`、`zipfile` 能打开、同事能打开，三者一致就无需处理。

同一个团队里"我的文件同事打不开、同事的文件我能打开"，基本可以断定是**写入方进程不同**（白名单 vs 非白名单）而不是文件损坏。

## 5. 已知边界

- 本节所有结论基于本机的 Agile DGS 类驱动；其它厂商的透明加密产品头部标记与白名单行为**未验证**。
- 本 skill 不做加解密：加密是驱动行为，解密必须用专用工具。`ENCRYPTED` 这类错误码的出现，通常意味着"这份文件当前没有任何白名单进程能读它"。
- 加密对 `.doc`/`.docx`/`.pdf`/图片/CAD 都生效（[研] §2.3 ① 与既有记忆），但各格式的头部标记只有 `17DA5FA01633CD9A` 这一种被本机确认过。
