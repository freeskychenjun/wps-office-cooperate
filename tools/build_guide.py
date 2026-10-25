# -*- coding: utf-8 -*-
"""Build a single-file HTML guide for the wps-word-cooperate skill.

Images are downscaled and embedded as base64 so the file can be copied on its own
(someone pastes it into an e-mail or a shared folder and it still works offline).
"""
import base64
import io
import os

from PIL import Image

BASE = r"C:\Users\Administrator\.agents\skills\wps-word-cooperate"
OUT = os.path.join(BASE, "guide.html")


def embed(path, width=1000, quality=82):
    img = Image.open(path)
    # 裁掉页面图四周的大块空白，说明书里别浪费半屏白纸
    gray = img.convert("L")
    box = gray.point(lambda v: 255 if v < 200 else 0).getbbox()
    if box:
        pad = 12
        left = max(0, box[0] - pad); top = max(0, box[1] - pad)
        right = min(img.width, box[2] + pad); bottom = min(img.height, box[3] + pad)
        if (right - left) > 200 and (bottom - top) > 200:
            img = img.crop((left, top, right, bottom))
    if img.width > width:
        h = int(round(img.height * width / img.width))
        img = img.convert("RGB").resize((width, h), Image.LANCZOS)
    else:
        img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=quality, optimize=True, progressive=True)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


img_rev = embed(os.path.join(BASE, "docs", "example-revisions.png"))
img_tbl = embed(os.path.join(BASE, "docs", "example-table.png"), width=900)

CSS = """
:root{
  --ink:#1b1f24; --muted:#5c6773; --line:#e3e7ec; --bg:#fbfcfd; --card:#ffffff;
  --accent:#1f4e79; --accent2:#c2452d; --code-bg:#f4f6f8; --tip:#eef6ef; --warn:#fdf3ee;
  --shadow:0 1px 2px rgba(16,24,40,.04), 0 8px 24px rgba(16,24,40,.06);
}
@media (prefers-color-scheme: dark){
  :root{ --ink:#e8eaed; --muted:#9aa4b0; --line:#2b3138; --bg:#15181c; --card:#1b1f24;
         --accent:#8ab4f8; --accent2:#f28b82; --code-bg:#20252b; --tip:#1c2a20; --warn:#2b211c;
         --shadow:none; }
}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{
  margin:0; background:var(--bg); color:var(--ink);
  font:16px/1.75 -apple-system,"Segoe UI","PingFang SC","Microsoft YaHei","Noto Sans SC",sans-serif;
  -webkit-font-smoothing:antialiased;
}
.wrap{display:grid; grid-template-columns:264px minmax(0,1fr); gap:40px; max-width:1240px; margin:0 auto; padding:0 28px}
nav.toc{position:sticky; top:0; align-self:start; max-height:100vh; overflow:auto; padding:34px 0 40px; font-size:14px}
nav.toc h2{font-size:12px; letter-spacing:.12em; text-transform:uppercase; color:var(--muted); margin:0 0 10px}
nav.toc a{display:block; padding:5px 10px; border-left:2px solid transparent; color:var(--muted); text-decoration:none; border-radius:0 6px 6px 0}
nav.toc a:hover{color:var(--ink); background:var(--code-bg)}
nav.toc a.sub{padding-left:22px; font-size:13px}
main{padding:34px 0 80px; min-width:0}
header.hero{border-bottom:1px solid var(--line); padding-bottom:26px; margin-bottom:34px}
.badge{display:inline-block; font-size:12px; letter-spacing:.06em; padding:3px 9px; border-radius:999px;
       background:var(--code-bg); color:var(--muted); margin-right:6px}
h1{font-size:34px; line-height:1.25; margin:14px 0 8px; letter-spacing:-.01em}
h1 code{font-size:.8em}
.lede{font-size:17px; color:var(--muted); margin:0; max-width:46em}
h2{font-size:23px; margin:52px 0 14px; padding-top:10px; letter-spacing:-.01em}
h3{font-size:17px; margin:30px 0 8px}
p{margin:10px 0}
ul,ol{margin:10px 0; padding-left:1.35em}
li{margin:5px 0}
code{font-family:"Cascadia Mono",Consolas,"SF Mono",Menlo,monospace; font-size:.9em;
     background:var(--code-bg); padding:.12em .38em; border-radius:4px}
pre{background:var(--code-bg); border:1px solid var(--line); border-radius:10px; padding:14px 16px;
    overflow:auto; font-size:13.5px; line-height:1.6}
pre code{background:none; padding:0; font-size:inherit}
table{width:100%; border-collapse:collapse; margin:14px 0; font-size:14.5px}
th,td{text-align:left; padding:9px 12px; border-bottom:1px solid var(--line); vertical-align:top}
th{font-size:12.5px; letter-spacing:.04em; text-transform:uppercase; color:var(--muted); font-weight:600}
tbody tr:hover{background:var(--code-bg)}
.card{background:var(--card); border:1px solid var(--line); border-radius:14px; padding:20px 22px; box-shadow:var(--shadow); margin:18px 0}
.grid2{display:grid; grid-template-columns:1fr 1fr; gap:18px}
.grid3{display:grid; grid-template-columns:repeat(3,1fr); gap:16px}
.tip{background:var(--tip); border-left:3px solid #3f8b52; padding:12px 16px; border-radius:0 8px 8px 0; margin:16px 0}
.warn{background:var(--warn); border-left:3px solid var(--accent2); padding:12px 16px; border-radius:0 8px 8px 0; margin:16px 0}
figure{margin:20px 0}
figure img{width:100%; border:1px solid var(--line); border-radius:10px; display:block}
figcaption{font-size:13px; color:var(--muted); margin-top:8px}
.steps{counter-reset:s; list-style:none; padding:0}
.steps li{counter-increment:s; position:relative; padding-left:44px; margin:14px 0}
.steps li::before{content:counter(s); position:absolute; left:0; top:1px; width:28px; height:28px;
  border-radius:50%; background:var(--accent); color:#fff; font-size:14px; font-weight:600;
  display:grid; place-items:center}
.copy{font-size:12.5px; color:var(--muted)}
footer{border-top:1px solid var(--line); margin-top:60px; padding-top:18px; font-size:13px; color:var(--muted)}
@media (max-width:900px){
  .wrap{grid-template-columns:1fr; gap:0}
  nav.toc{position:static; max-height:none; padding:18px 0; border-bottom:1px solid var(--line)}
  .grid2,.grid3{grid-template-columns:1fr}
  h1{font-size:27px}
}
@media print{
  nav.toc{display:none}
  .wrap{grid-template-columns:1fr; max-width:none}
  body{background:#fff}
  .card,pre,table{box-shadow:none}
  h2{page-break-after:avoid}
}
"""

HTML = """<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>wps-word-cooperate · 使用说明</title>
<style>__CSS__</style>
</head>
<body>
<div class="wrap">

<nav class="toc">
  <h2>目录</h2>
  <a href="#what">这是什么</a>
  <a href="#install">装它</a>
  <a href="#start">五分钟上手</a>
  <a href="#commands">命令全景</a>
  <a href="#format">格式项速查</a>
  <a href="#collab">和人接力</a>
  <a href="#scenario">两套剧本</a>
  <a href="#pitfalls">实测坑</a>
  <a href="#faq">常见问题</a>
  <a href="#notfor">什么时候别用</a>
  <a href="#dev">给想改它的人</a>
  <a href="#limits">边界与未验证</a>
</nav>

<main>

<header class="hero">
  <div>
    <span class="badge">27 条命令</span>
    <span class="badge">Windows + WPS</span>
    <span class="badge">零第三方依赖</span>
    <span class="badge">v0.4</span>
  </div>
  <h1><code>wps-word-cooperate</code><br>在 WPS 里和 AI 一起改 Word</h1>
  <p class="lede">AI 直接在你已经打开的 WPS 文字里改文档——改动实时出现在窗口里，文字修改是<b>修订</b>，你能逐条接受或拒绝；排版、表格、题注、目录、批注、版式核对也能一并做。不重写文件、不改变版式、不悄悄改掉原文。</p>
</header>

<h2 id="what">这是什么</h2>

<p>它调的是 <b>WPS 自己的自动化接口（COM）</b>，不是"读文件—改文件—写文件"。这带来三个直接好处：</p>

<div class="grid3">
  <div class="card"><h3>文件开着也能改</h3><p class="copy">不会因为"文件被占用"失败。已经打开的直接接着用，不会再开一份。</p></div>
  <div class="card"><h3>格式不会走样</h3><p class="copy">排版由 WPS 自己做，存下来还是 WPS 的样子，不会被第三方库重写得面目全非。</p></div>
  <div class="card"><h3>你能边看边改</h3><p class="copy">AI 每次动手前读到的都是文档当下的样子；你随时可以插手，它下次读到的是你改过的版本。</p></div>
</div>

<figure>
  <img src="__IMG_REV__" alt="AI 改完后的样子：修订红字与右侧批注气泡">
  <figcaption>AI 改完的样子：正文里是修订（红色下划线/删除线），右侧是批注气泡。你可以在 WPS 的「审阅」里逐条接受或拒绝。</figcaption>
</figure>

<div class="tip">
  <b>和"让 AI 写个脚本改 docx"的区别：</b>脚本路线你看不到过程、文件开着还会写不进去、公式与版式会走样；这条路是<b>让 WPS 自己动手</b>，你在旁边看着。
</div>

<h2 id="install">装它</h2>

<ol class="steps">
  <li><b>拷文件夹</b><br>把整个 <code>wps-word-cooperate</code> 文件夹放到你的技能目录，通常是
    <code>C:\\Users\\&lt;你的用户名&gt;\\.agents\\skills\\</code>。<br>
    也可以装进院里的 skill 管理平台。</li>
  <li><b>确认有 WPS</b><br>这台机器要装 <b>WPS Office</b>（文字组件即可）。微软 Office 也能用，会自动回退。</li>
  <li><b>自检</b>
<pre><code>powershell -NoProfile -ExecutionPolicy Bypass -File ^
  "C:\\Users\\&lt;你&gt;\\.agents\\skills\\wps-word-cooperate\\scripts\\wps.ps1" doctor</code></pre>
    看到 <code>wps.installed: true</code> 就装好了。第一次跑会用系统自带的 <code>csc.exe</code> 编译一个约 1 秒的小程序，之后一直用缓存。</li>
</ol>

<div class="warn">
  <b>唯一依赖：</b>.NET Framework 4.8（Windows 10/11 自带）+ WPS。<b>不需要 Python</b>，<b>不需要装插件</b>，<b>不需要联网</b>。
</div>

<h2 id="start">五分钟上手</h2>

<p>最省事：直接跟 AI 说人话。</p>

<div class="card">
  <h3>例子 1 · 润色并逐条看</h3>
  <p>「打开 <code>D:\\报告\\某报告.docx</code>，读一遍，把第 3 章里"大概 1234 m3/s"改成"约 1234 m³/s"，我能看着改。」</p>
  <p class="copy">→ AI 以修订形式改动，你在 WPS 里接受/拒绝。</p>
</div>

<div class="card">
  <h3>例子 2 · 校审批注</h3>
  <p>「按检查要点过一遍第 2 章，问题直接在原文对应段落加批注。」</p>
  <p class="copy">→ 批注挂在具体字句上，右侧气泡可见。</p>
</div>

<div class="card">
  <h3>例子 3 · 一键排版</h3>
  <p>「全文正文改成仿宋小四、1.5 倍行距、首行缩进 2 字符，标题 1 黑体三号居中，页脚加页码。」</p>
  <p class="copy">→ 格式类改动不进修订，<b>一次 Ctrl+Z 可整批撤回</b>。</p>
</div>

<p>想自己调命令也简单（路径按你机器改）：</p>
<pre><code>$s = "C:\\Users\\&lt;你&gt;\\.agents\\skills\\wps-word-cooperate\\scripts\\wps.ps1"

powershell -NoProfile -ExecutionPolicy Bypass -File $s status
powershell -NoProfile -ExecutionPolicy Bypass -File $s open   -Doc "D:\\报告\\某报告.docx"
powershell -NoProfile -ExecutionPolicy Bypass -File $s read   -Doc "D:\\报告\\某报告.docx" -From 1 -Count 60
powershell -NoProfile -ExecutionPolicy Bypass -File $s revise -Doc "D:\\报告\\某报告.docx" -Items edits.json
powershell -NoProfile -ExecutionPolicy Bypass -File $s render -Doc "D:\\报告\\某报告.docx" -Page 1 -To 3 -Out "D:\\报告\\_pages"
powershell -NoProfile -ExecutionPolicy Bypass -File $s save   -Doc "D:\\报告\\某报告.docx"</code></pre>

<p><code>edits.json</code> 长这样（<code>expect</code> 是"这段开头的几个字"，用来保证改的位置没错）：</p>
<pre><code>{"items":[
  {"para":12,"expect":"3.2 工程地质","op":"replace","find":"孔口高程","replace":"孔口地面高程"},
  {"para":15,"expect":"根据地质勘察","op":"set_text","text":"根据地质勘察成果，坝址区基岩为……"},
  {"para":31,"expect":"本阶段不再","op":"delete"}
]}</code></pre>

<h2 id="commands">命令全景（27 条）</h2>

<h3>会话与自检</h3>
<table>
<thead><tr><th style="width:15%">命令</th><th>作用</th></tr></thead>
<tbody>
<tr><td><code>doctor</code></td><td>环境自检：哪个套件可用、版本、是否在跑、开着的文档、渲染能力、进程数</td></tr>
<tr><td><code>status</code></td><td>在跑的实例与文档清单（路径 / 是否只读 / 是否已保存 / 段数）</td></tr>
<tr><td><code>open</code></td><td>打开或复用；返回 readOnly、段落数、页数</td></tr>
<tr><td><code>close</code> / <code>quit</code></td><td>关一份 / 关全部并退出（<code>quit</code> 会波及你开的其它文档，慎用）</td></tr>
</tbody>
</table>

<h3>读与改</h3>
<table>
<thead><tr><th style="width:15%">命令</th><th>作用</th></tr></thead>
<tbody>
<tr><td><code>read</code></td><td>段落（编号 · 样式 · 自动编号 · 文本 · 是否在表格里）+ 表格清单 + 页数/修订数/批注数</td></tr>
<tr><td><code>revise</code></td><td><b>修订模式</b>批量改写：<code>replace</code>（字符级）· <code>set_text</code> · <code>insert_after</code> · <code>delete</code>；先全量校验，有一条不符就一条不改</td></tr>
<tr><td><code>comment</code></td><td>加批注（可锚到某几个字）；<code>-Delete -Index N</code> 删批注</td></tr>
<tr><td><code>insert</code></td><td><b>批量插入段落</b>（默认干净写入、不进修订；<code>-Track</code> 才留痕），可带样式与格式</td></tr>
<tr><td><code>cell</code></td><td>改一个表格单元格的文字（可带字体字号对齐）</td></tr>
</tbody>
</table>

<h3>排版</h3>
<table>
<thead><tr><th style="width:15%">命令</th><th>作用</th></tr></thead>
<tbody>
<tr><td><code>font</code></td><td>整篇统一字体（正文、表格、页眉页脚、文本框一起换，中英文都换）</td></tr>
<tr><td><code>format</code></td><td>整篇 / 一段 / 区间 / 某样式的所有段：字符格式 + 段落格式</td></tr>
<tr><td><code>style</code></td><td>改<b>样式本身</b>（正文、标题 1…），所有用它的段落一起变；可顺带去掉自动编号</td></tr>
<tr><td><code>page</code></td><td>纸张 / 方向 / 页边距 / 页眉页脚文字 / 页码（逐节）</td></tr>
</tbody>
</table>

<h3>表格、图表、目录</h3>
<table>
<thead><tr><th style="width:15%">命令</th><th>作用</th></tr></thead>
<tbody>
<tr><td><code>table</code></td><td>建表（<code>grid/three-line/none</code> 框线、表头跨页重复、列宽、行不跨页、表题）/ 改表</td></tr>
<tr><td><code>caption</code></td><td><b>域</b>式题注「表 1 / 图 1」，自动编号、题不落单</td></tr>
<tr><td><code>picture</code></td><td>插图（可带图题）/ 调宽高 / 按百分比裁剪；缩太小会告警</td></tr>
<tr><td><code>toc</code></td><td>插入并刷新目录域（依赖标题样式）</td></tr>
<tr><td><code>reference</code></td><td>参考文献列表（悬挂缩进）/ 段末上标引用 <code>[1,2,5]</code></td></tr>
<tr><td><code>audit</code></td><td><b>机械检查</b>：单位未用上标、重复词、括号混用、连续空格、全角数字、标题跳级、图表缺题注——返回段落号与建议，供判断后落批注</td></tr>
<tr><td><code>fields</code></td><td><b>一键刷新全部域</b>：页码、目录、题注、引用（改完正文、交付前跑一次）</td></tr>
<tr><td><code>view</code></td><td>把窗口滚到某段/某页（不动光标）</td></tr>
</tbody>
</table>

<h3>协作交接</h3>
<table>
<thead><tr><th style="width:15%">命令</th><th>作用</th></tr></thead>
<tbody>
<tr><td><code>revisions</code></td><td>列出<b>未处理</b>的修订（类型/作者/所在段/文本），或按人要求接受/拒绝</td></tr>
<tr><td><code>snapshot</code></td><td>存一份段落文本快照（给 <code>diff</code> 用，不写进文档目录）</td></tr>
<tr><td><code>diff</code></td><td>和快照比：哪些段被改了（before→after）、哪些增删——<b>人改的和 AI 改的都算</b></td></tr>
<tr><td><code>handoff</code></td><td>在文档旁生成 <code>.handoff.md</code> 交接单：状态 / 待办 / 操作记录 / 变化表</td></tr>
</tbody>
</table>

<h3>看与存</h3>
<table>
<thead><tr><th style="width:15%">命令</th><th>作用</th></tr></thead>
<tbody>
<tr><td><code>render</code></td><td>把页面渲成 PNG 给 AI 核对版式（走打印视图的页面图元，不截图、不落 PDF）</td></tr>
<tr><td><code>save</code></td><td>保存 / 另存（覆盖写同一个文件时保持 WPS 自己的格式）</td></tr>
</tbody>
</table>

<h2 id="format">格式项速查</h2>

<p><code>format</code> / <code>style</code> / <code>table</code> / <code>caption</code> / <code>insert</code> 共用这套参数：</p>

<table>
<thead><tr><th style="width:22%">参数</th><th style="width:30%">取值</th><th>说明</th></tr></thead>
<tbody>
<tr><td><code>-Font</code></td><td><code>仿宋</code></td><td>中英文一起换</td></tr>
<tr><td><code>-LatinFont</code></td><td><code>"Times New Roman"</code></td><td>只换西文，不动中文字体</td></tr>
<tr><td><code>-Size</code></td><td><code>12</code> / <code>"10.5pt"</code></td><td>字号（磅）</td></tr>
<tr><td><code>-Bold</code> <code>-Italic</code> <code>-Underline</code></td><td>开关</td><td>关掉用 <code>-Bold:$false</code></td></tr>
<tr><td><code>-Super</code> <code>-Sub</code></td><td>开关</td><td>上标 / 下标</td></tr>
<tr><td><code>-Color</code></td><td><code>"#C00000"</code></td><td>字体颜色</td></tr>
<tr><td><code>-Align</code></td><td><code>left|center|right|justify</code></td><td>中文也认 <code>居左/居中/居右/两端对齐</code></td></tr>
<tr><td><code>-FirstLineIndent</code></td><td><code>2</code> / <code>"0.74cm"</code> / <code>0</code></td><td><b>首行缩进</b>：纯数字＝字符数，带单位＝长度</td></tr>
<tr><td><code>-IndentChars</code></td><td><code>2</code></td><td>左缩进（按字符）</td></tr>
<tr><td><code>-LineSpacing</code></td><td><code>1.5</code> / <code>"20pt"</code> / <code>"at least 20pt"</code> / <code>single</code></td><td>倍数 / 固定值 / 最小值</td></tr>
<tr><td><code>-SpaceBefore</code> <code>-SpaceAfter</code></td><td><code>"6pt"</code> / <code>0</code></td><td>段前 / 段后</td></tr>
<tr><td><code>-PageBreakBefore</code></td><td>开关</td><td>该段另起一页</td></tr>
<tr><td><code>-NoSplitWords</code></td><td>开关</td><td>整词换行（治"37000 被拆成 37+000"）</td></tr>
<tr><td><code>-Track</code></td><td>开关</td><td>让格式改动也进修订（默认不进）</td></tr>
</tbody>
</table>

<h2 id="collab">和人接力：这是它最该被用的方式</h2>

<p>一套完整的交接只需要四个动作。下面是<b>真实跑过</b>的一轮（592 段、31 页的报告）：</p>

<ol class="steps">
  <li><b>AI 开工前拍快照</b><br><code>snapshot -Tag round1</code> —— 记下"我接手时文档长什么样"。</li>
  <li><b>AI 干活</b><br>两处文字修订 + 一条批注，全部以修订形式留在文档里。</li>
  <li><b>人裁决 / 手改</b><br>人在 WPS 里接受了 1 处修订，并手改了另一段（"4 种"→"5种"）。</li>
  <li><b>AI 读结果继续</b><br><code>revisions</code> 看到待裁决 4→2 条；<code>diff -Since round1</code> 精确报出：
    第 61 段（我的修订）、第 68 段（人被采纳的）、第 85 段（人手改的 4→5 种）。<br>
    然后 <code>handoff -Note "…"</code> 在文档旁生成交接单。</li>
</ol>

<div class="tip">
  <b>关键在于 AI 不需要你把改了什么复述一遍</b>，也不用把 592 段重读一遍——<code>revisions</code> 告诉你还剩多少待裁决，<code>diff</code> 告诉你文本变成了什么样。
</div>

<div class="warn">
  <b>一个必须记住的语义：</b>WPS 里的文本是<b>"接受视图"</b>——插入的字临时算进去、删掉的字临时不算。所以<br>
  · <b>「还有多少待裁决」看 <code>revisions</code></b><br>
  · <b>「文本最后长什么样」看 <code>diff</code></b>
</div>

<div class="card" style="border-left:3px solid var(--accent)">
  <h3>一条通道（先守这条，再谈别的）</h3>
  <p>一旦这份文档在本 skill 手上（WPS 里开着），<b>这次协作里就不要把它交给别的文档 skill</b>：校审用 <code>audit</code> + <code>comment -Items</code>，润色用 <code>revise</code>，排版用 <code>format</code> / <code>style</code> / <code>page</code>。</p>
  <p class="copy">原因不是体验，是<b>会坏</b>：那几个 skill 改的是磁盘文件，而这份文件正被 WPS 独占（实测写盘 PermissionError）；它们还产出副本，一改就分成两份内容不同的稿子。</p>
  <p class="copy">例外：用户要「不动原件、出一份新稿」或「无人值守批量跑」——那就先 <code>save</code>、<code>close</code>，把文件交出去。</p>
</div>

<h2 id="scenario">两套协作剧本</h2>

<div class="grid2">
  <div class="card">
    <h3>A · 批次式（推荐）</h3>
    <p>AI 一口气干完一段 → 人审 → AI 再接手第二轮。</p>
    <p><code>snapshot</code> → 批量改（修订）→ 请人处理 → <code>diff</code> + <code>revisions</code> 读结果 → 第二轮 → <code>save</code> → <code>handoff</code></p>
    <p class="copy">原则：一轮只做一类事（润色 / 排版 / 校审分开），人处理完再开下一轮；每轮开头都拍快照。</p>
  </div>
  <div class="card">
    <h3>B · 并发式</h3>
    <p>你和 AI 同时在看同一份文档、来回交替。</p>
    <p>① 写之前先看返回值里的 <code>changedOutside</code>；② <code>expect</code> 报 <code>ANCHOR_MISSING</code> 时先 <code>read</code> 那段——很可能是你改了它；③ 不替你裁决修订；④ 不动你的光标。</p>
    <p class="copy">注意：感知是"轮次级"的——AI 只在下一条命令时才被告知"中间有人动过"，做不到实时跟随你打字。</p>
  </div>
</div>

<h2 id="pitfalls">实测坑（照做少返工）</h2>

<table>
<thead><tr><th style="width:30%">现象</th><th>原因与对策</th></tr></thead>
<tbody>
<tr><td>表格 <code>-Para</code> 插表后位置很怪</td><td>锚点段落若在表格里，会变成"往单元格里塞表"。<code>read</code> 里带 <code>inTable: true</code> 的段落会报错；要嵌表加 <code>nested: true</code>，否则换一个正文段。</td></tr>
<tr><td>插入的段落跑进了表格单元格</td><td>锚点紧挨表格时可能发生；返回里 <code>inTable: true</code> 就是提示。要插在表格前，改用 <code>-Where before</code> 指向表格后那一段。</td></tr>
<tr><td>改了批注旁边的字，多出/少了 1 个字</td><td>批注锚点占一个字符位置却不在 <code>Range.Text</code> 里。程序会先读回校验并自动纠正，纠正不了就拒绝执行；必要时先 <code>comment -Delete</code> 摘掉批注再改。</td></tr>
<tr><td>页码出现两套</td><td>中文报告常在<b>页眉</b>编页码；<code>page -PageNumber</code> 是往<b>页脚</b>插。动手前先 <code>render</code> 看一眼，或干脆不传。</td></tr>
<tr><td>目录是空的 / 显示"错误！未找到目录项"</td><td>文档里没有「标题 1/2/3」样式的段落。先用 <code>style</code> 或人工套标题样式。</td></tr>
<tr><td>截图里加粗看不出来</td><td>WPS 的 EMF 渲图对 <code>Bold</code> 呈现不可靠，docx 里是对的——别据此判断没加粗。</td></tr>
<tr><td>图片里的字小得看不清</td><td>插图时被缩得太小；<code>picture</code> 会返回 <code>scale</code> 和 <code>warning</code>，按提示重画或加宽。</td></tr>
<tr><td>标题文字本身带编号，样式又自动编号</td><td>出现双重编号；用 <code>style -NoNumbering</code> 去掉自动的那个。</td></tr>
</tbody>
</table>

<h2 id="faq">常见问题</h2>

<table>
<thead><tr><th style="width:26%">错误码 / 现象</th><th>怎么办</th></tr></thead>
<tbody>
<tr><td><code>NO_APP</code></td><td>没装 WPS 文字，或 COM 没注册。手动打开一次 WPS 再试。</td></tr>
<tr><td><code>NOT_OPEN</code></td><td>路径和实际打开的不是同一个（另存过、短路径、相对路径）。先 <code>status</code>，把 <code>path</code> 原样复制。</td></tr>
<tr><td><code>READONLY</code></td><td>文件被别处占用或只读。关掉占用的程序，或另存一份再改——不要强行覆盖。</td></tr>
<tr><td><code>ANCHOR_MISSING</code></td><td>段落编号漂移 / 原文对不上 / 可能是人改了那段。<b>整批一条都没改</b>，重新 <code>read</code> 核对再来。</td></tr>
<tr><td><code>BUSY</code> / <code>TIMEOUT</code></td><td>你正在打字，或有对话框挡着。关掉对话框，等几秒重试，别连发。</td></tr>
<tr><td>中文乱码</td><td>用 <code>powershell -File</code> 调脚本（内部固定 UTF-8），别自己把 JSON 拼到命令行里。</td></tr>
<tr><td><code>wps.exe</code> 变多</td><td>中途异常没 <code>close</code>。先 <code>status</code> 看实例数，只结束<b>无窗口</b>的自动化实例，别按名字批量杀。</td></tr>
<tr><td>保存后文件发出去打不开</td><td>单位透明加密：保存后 docx/PDF 落盘即密文（文件头 <code>17DA5FA01633CD9A</code>）。发之前先解密。</td></tr>
</tbody>
</table>

<h2 id="notfor">什么时候<b>不要</b>用它</h2>
<ul>
  <li><b>无人值守、上千段一次跑完</b>的批量润色 → 用离线的 <code>docx-paragraph-polish</code>（不依赖 WPS，也不受加密影响）</li>
  <li><b>纯排版、不需要人看过程</b> → <code>design-report-format</code></li>
  <li><b>只要一份检查报告、不改原文</b> → <code>design-report-review</code></li>
  <li><b>Excel / PPT</b> → 这个 skill 只做 Word</li>
</ul>

<figure>
  <img src="__IMG_TBL__" alt="三线表与题注的验收页">
  <figcaption>验收页：三线表（上下粗线、表头下细线）+ 域式题注「表 1 …」。这些都是一条命令生成的，且都在 WPS 里可直接编辑。</figcaption>
</figure>

<h2 id="dev">给想改它的人</h2>

<pre><code>scripts/wps.ps1          薄包装：编译缓存 → 写 UTF-8 请求 → 跑 helper → 回 JSON
scripts/helper/          C# 5，4,691 行，partial class Program
  WpsHelper.cs  1263 行  会话 / 读 / 修订 / 批注 / 渲染 / 字体 / JSON / COM / 格式原语
  Layout.cs      293 行  format / style / page
  Tables.cs     1401 行  table / caption / picture
  Toc.cs         617 行  toc / reference / view
  Text.cs        230 行  fields / cell / insert
  Review.cs      598 行  revisions / snapshot / diff
  State.cs       289 行  并发检测 / 写入流水 / handoff
tests/                   306 段合成报告 + 独立 COM 校验器（可复跑验收）
reference/               WPS 差异清单、加密与交付说明</code></pre>

<h3>加一条命令</h3>
<ol>
  <li>在对应 <code>.cs</code> 里写 <code>private static object XxxDoc(Bag a)</code>；</li>
  <li><code>WpsHelper.cs</code> 的 <code>Dispatch</code> 里加一行 <code>case "xxx":</code>；</li>
  <li><code>wps.ps1</code> 的 <code>ValidateSet</code> 和 <code>switch</code> 里各加一段。</li>
</ol>
<p>改完直接跑 <code>wps.ps1 status</code>：源码哈希变了会自动重编（约 1 秒）。</p>

<h2 id="limits">边界与未验证</h2>

<div class="grid2">
  <div class="card">
    <h3>做不到的</h3>
    <ul>
      <li>逐字写入动画、跟随滚动、右下角迷你卡片、无窗口静默实例——这些是宿主插件的能力，skill 没有窗口句柄</li>
      <li>并发的"实时感知"：只能在下一条命令时告诉你"中间有人动过"</li>
      <li>公式（LaTeX→原生）、流程图（形状+连线）、版面回填——参考实现有，这里没做</li>
    </ul>
  </div>
  <div class="card">
    <h3>实测环境</h3>
    <ul>
      <li>WPS Office <b>12.1.0.29161</b>（其它版本未验证）</li>
      <li>合成报告 306 段 / 29 页；真实报告 592 段 / 31 页，全流程跑过</li>
      <li>页眉页脚里的修订、旧 <code>.doc</code> 格式、SVG 插图未验证</li>
      <li>只核对过渲染图，没实际打印</li>
    </ul>
  </div>
</div>

<footer>
  wps-word-cooperate · v0.4 · 27 条命令 · 最后更新 2026-10-23<br>
  配套文档：<code>README.md</code>（给人）· <code>SKILL.md</code>（给 AI）· <code>reference/</code>（WPS 差异与加密）<br>
  做法参考 <a href="https://github.com/copylee711/dsh-office">copylee711/dsh-office</a>（MIT），代码为本项目自写 · 本项目 MIT License
</footer>

</main>
</div>
</body>
</html>
"""

html = HTML.replace("__CSS__", CSS).replace("__IMG_REV__", img_rev).replace("__IMG_TBL__", img_tbl)
with open(OUT, "w", encoding="utf-8", newline="\n") as fh:
    fh.write(html)
print("wrote", OUT, round(len(html.encode("utf-8")) / 1024), "KB")
