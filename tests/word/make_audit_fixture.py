# Build audit-fixture.docx: copy report-300.docx and plant one known finding per rule.
# Pure ASCII source. Prints ground truth + existing styles for verification.
import shutil
from docx import Document

SRC = r'D:\DSH工作区\wps-office-cooperate\tests\word\corpus\report-300.docx'
DST = r'D:\DSH工作区\wps-office-cooperate\tests\word\corpus\audit-fixture.docx'
shutil.copyfile(SRC, DST)

doc = Document(DST)
styles = sorted(set(p.style.name for p in doc.paragraphs))
print('existing styles sample:', [s for s in styles if 'ead' in s or '标题' in s][:10])

def para(text, style=None):
    p = doc.add_paragraph()
    if style is not None:
        p.style = doc.styles[style]
    p.add_run(text)
    return p

# unit x2 (one paragraph, two hits -> single Match -> 1 finding; plus a second paragraph)
para('设计流量为1000 m3/s，集水面积254 km2，校核流量3500 m3/s。')
para('多年平均径流深为 812.5 mm，最大一日暴雨量为 225 mm。')
para('水库总库容为 1.28 亿m3，调洪库容 0.56 亿m3。')
# repeat-word
para('本工程工程规模为大（2）型。')
# bracket-mix in ONE paragraph
para('（见表3）与(table 4)的数据一致。')
# double space
para('连续  两个空格的段落。')
# space before fullwidth punct
para('句末多了空格 ，逗号前。')
# fullwidth digit
para('实测洪峰流量为３５００每秒立方米量级。')
# trailing space
para('这一段的末尾有一个空格 ')
# 4 consecutive empty paragraphs -> empty-run at the 4th
for _ in range(4):
    para('')
# heading jump: Heading 1 followed directly by Heading 3
h_styles = [s for s in styles if s.lower().startswith('heading') or s.startswith('标题')]
print('heading styles present:', h_styles)
need = []
for name in ('Heading 1', 'Heading 3'):
    try:
        doc.styles[name]
    except KeyError:
        need.append(name)
if need:
    from docx.enum.style import WD_STYLE_TYPE
    for name in need:
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = doc.styles('Heading 2') if False else None
    print('added styles:', need)
para('第九章 跳级测试章', 'Heading 1')
para('9.0.1 直接跳到三级标题', 'Heading 3')
doc.save(DST)
print('saved. paragraphs now:', len(Document(DST).paragraphs))
