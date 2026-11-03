# 发版 checklist（wps-office-cooperate）

> v1.34 起启用。起因：v1.32→v1.33 发版时 `SKILL.md` 的版本头条漏更，主文档声明"版本 v1.32"而 HEAD 与 `reference/changelog.md` 已是 v1.33——**同一份版本历史存在四份活副本**，漏更一次就是一次漂移。
>
> 口径：**版本号只有一个权威源 = `SKILL.md` frontmatter 的 `version`**；版本历史只有一个出处 = `reference/changelog.md`。其余位置只做"当前版本要点"，不再各留全量记录。

## 发版必改的五处

| # | 位置 | 改什么 | 校验 |
|---|---|---|---|
| 1 | `SKILL.md` frontmatter | `version: X.YY` | 与 #3 #4 一致 |
| 2 | `SKILL.md`「## 当前版本」小节 | 3–5 行**当前版要点**（不写历史；旧版内容只进 changelog） | 版本号一致；无 v(N-1) 叙述残留 |
| 3 | `reference/changelog.md` | **新条目放最上面**，格式 `**版本 vX.YY（日期）**：` | 首条 = 当前版 |
| 4 | `README.md` | 能力清单补本轮新能力；`vN` 相关表述同步 | 搜 `v1.` 看有没有旧版残留 |
| 5 | commit message | `vX.YY: <一句话>` | 与 #1 一致 |

## 每次都要确认

- [ ] `reference/**` 有新增文件？→ 登记到 `publish.ps1` 的 `$includes`，否则部署目录缺文件。
- [ ] 新增相对 markdown 链接？→ 全部能解析到真实文件（`reference/`、`docs/` 里被引用的路径都在）。
- [ ] 改了 `.ps1`？→ 纯 ASCII + UTF-8 BOM，跑 `tests\check_encoding.ps1`。
- [ ] 跑了 `tests\run_all.ps1 -SkipWps`（fast 层）；动到 COM 侧再跑 `-Layer wps`。
- [ ] 纯文档改动（本类）也要过一遍：`git diff --stat` 只看文档；不 `git commit`（由人决定）。

## 反模式（已发生过的）

- **在主文档里再写一份版本历史**——v1.29 前主文档前 74 行是版本记录，占整份 skill 25.5%，且与 changelog 双写，于是漏更。现在历史只许在 `reference/changelog.md`。
- **把实测细节塞进速查表单元格**——v1.34 之前 `table`/`slide` 两行各 1.9K / 1.5K 字符，成了"changelog 第二现场"。速查表只放参数 + 一句话坑位 + 指针，细节进 `reference/*-quirks.md`。
- **靠人记忆同步版本号**——现在有这五处清单；漏了就按上表回查。
