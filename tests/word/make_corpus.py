#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Generate the acceptance corpus for the `wps-word-cooperate` skill.

Outputs (all under tests/corpus/):
  report-300.docx            ~300-paragraph Chinese water-conservancy report
  report-300.pristine.docx   byte copy of the freshly generated file (baseline)
  targets.json               5 polish targets (3 required + 2 extra): para/original/find/replace
  meta.json                  paragraph index map python-docx -> WPS + tag registry

The docx is written by python-docx and is NEVER saved by WPS here, so it stays
plaintext (zip) and can be re-read by python-docx later.  WPS re-saves of the
*revised* copy will be encrypted (17DA5FA01633CD9A) - that is the point of the
acceptance test.
"""

import json
import os
import random
import shutil
import sys
import zipfile

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = os.path.dirname(os.path.abspath(__file__))
CORPUS = os.path.join(HERE, "corpus")
OUT_DOCX = os.path.join(CORPUS, "report-300.docx")
PRISTINE = os.path.join(CORPUS, "report-300.pristine.docx")
TARGETS_JSON = os.path.join(CORPUS, "targets.json")
META_JSON = os.path.join(CORPUS, "meta.json")

MIN_LEN, MAX_LEN = 40, 150


# --------------------------------------------------------------------------
# vocabulary / number filler
# --------------------------------------------------------------------------
def nums(rng):
    return dict(
        A=rng.choice([820, 960, 1140, 1280, 1560, 1830, 2140, 2470, 2860, 3240]),
        L=rng.choice([62, 78, 95, 112, 128, 146, 163, 178]),
        L2=rng.choice([8.6, 10.2, 12.4, 14.8, 16.5, 18.2]),
        J=rng.choice([1.2, 1.8, 2.4, 3.1, 3.6, 4.2, 4.8, 5.5]),
        veg=rng.randint(58, 86),
        d=round(rng.uniform(1.2, 4.6), 1),
        mu=round(rng.uniform(0.8, 8.5), 1),
        pop=round(rng.uniform(2.5, 12.0), 1),
        f=rng.choice([10, 20, 50]),
        pct=rng.choice([1, 2, 5]),
        rain=rng.randint(1500, 1900),
        temp=round(rng.uniform(16.8, 18.6), 1),
        b=rng.randint(45, 95),
        year=rng.choice([1958, 1972, 1985, 1996, 2003, 2011, 2016, 2019]),
        len3=round(rng.uniform(2.4, 9.8), 1),
        invest=rng.randint(8600, 46000),
        invest2=rng.randint(5200, 32000),
        Q=rng.randint(820, 2500),
        Q5=rng.randint(560, 1600),
        W=rng.randint(1200, 5600),
        n=rng.randint(38, 66),
        Qm=rng.randint(600, 1200),
        cv=round(rng.uniform(0.38, 0.62), 2),
        cs=round(rng.uniform(2.0, 3.5), 1),
        m1=round(rng.uniform(3.2, 6.8), 1),
        K=round(rng.uniform(1.2, 3.6), 1),
        K2=round(rng.uniform(1.08, 1.42), 2),
        th=rng.randint(18, 36),
        nz=round(rng.uniform(0.025, 0.040), 3),
        Hmax=round(rng.uniform(66.5, 72.5), 2),
        Qmax=rng.randint(1200, 2400),
        month=rng.choice([5, 6, 7]),
        mod=round(rng.uniform(0.6, 1.4), 2),
        sed=rng.randint(8, 46),
        sedge=rng.randint(68, 88),
        qh=round(rng.uniform(3.2, 8.6), 1),
        ct=round(rng.uniform(2.2, 6.4), 1),
        fak=rng.randint(90, 160),
        dw=rng.randint(1, 5),
        wf=round(rng.uniform(3.5, 9.2), 1),
        pg=round(rng.uniform(0.05, 0.10), 2),
        dis=round(rng.uniform(3.5, 18.0), 1),
        dis2=round(rng.uniform(2.5, 12.0), 1),
        bw=rng.choice([6.0, 7.0, 8.0, 9.0]),
        sup=round(rng.uniform(1.0, 2.0), 1),
        lv=rng.choice([3, 4, 5]),
        s1=rng.choice([2.0, 2.5, 3.0]),
        s2=rng.choice([1.5, 1.8, 2.0]),
        thk=round(rng.uniform(0.12, 0.30), 2),
        cnt=rng.randint(2, 6),
        q=rng.randint(15, 120),
        bk=rng.choice([2.0, 2.5, 3.0, 4.0]),
        kong=rng.randint(1, 3),
        comp=rng.choice([96, 97, 98]),
        road=rng.choice([6.0, 6.5, 7.0]),
        mg=rng.randint(5, 20),
        zhuang=rng.randint(20, 60),
        gr=rng.randint(2, 6),
        dp=rng.choice([0.5, 0.6, 0.8]),
        lp=rng.choice([6.0, 8.0, 10.0, 12.0]),
        gt=rng.randint(12, 24),
        gb=rng.randint(1, 3),
        zm=rng.randint(8, 18),
        ph=round(rng.uniform(3.2, 6.8), 1),
        pw=rng.choice([4.0, 5.0, 6.0]),
        tr=rng.randint(1, 3),
        tw=rng.uniform(8.5, 46.2),
        tf=rng.uniform(6.2, 32.5),
        conc=rng.uniform(1.2, 8.6),
        pp=rng.randint(120, 460),
        jj=rng.randint(18, 76),
        dg=rng.randint(600, 3200),
        cc=rng.randint(2, 8),
        duan=rng.randint(3, 6),
        rate=rng.choice([10, 15, 20]),
        irr=rng.uniform(6.8, 12.4),
        bc=rng.uniform(1.12, 1.68),
        loss=rng.randint(320, 2400),
    )


POOLS = {
    "overview": [
        "白沙河流域位于江西省中部，属鄱阳湖水系抚河支流，集水面积{A} km²，主河道全长{L} km。",
        "流域地势南高北低，上游为低山丘陵，中下游为河谷平原，植被覆盖率达{veg}%。",
        "本次治理河段自白沙镇上游{d} km 处起至河口止，治理长度{L2} km。",
        "工程任务以防洪为主，兼顾排涝与生态修复，保护耕地{mu} 万亩、人口{pop} 万人。",
        "根据《防洪标准》（GB 50201—2014），防护区等级为Ⅳ等，防洪标准采用{f} 年一遇。",
        "工程区属亚热带湿润季风气候，多年平均降水量{rain} mm，多年平均气温{temp} ℃。",
        "河道现状行洪断面不足，局部卡口段宽度仅{b} m，汛期水位壅高明显。",
        "上一轮治理于{year} 年实施，完成堤防加固{len3} km，本次为续建配套。",
        "工程总投资{invest} 万元，其中建筑工程投资{invest2} 万元。",
        "设计洪水采用 P={pct}%，相应洪峰流量{Q} m³/s，洪水总量{W} 万 m³。",
    ],
    "hydrology": [
        "白沙水文站设立于{year} 年，控制集水面积{A} km²，实测系列长度{n} 年。",
        "采用 P-Ⅲ 型曲线适线，均值{Qm} m³/s，Cv={cv}，Cs/Cv={cs}。",
        "设计洪水计算采用瞬时单位线法，汇流参数 m1={m1}，K={K} h。",
        "经计算，P=1% 设计洪峰流量为{Q} m³/s，P=5% 为{Q5} m³/s。",
        "设计洪水过程按同倍比放大，放大倍比{K2}，主峰历时{th} h。",
        "水位流量关系采用曼宁公式推求，糙率 n={nz}，比降 J={J}‰。",
        "实测最高水位{Hmax} m，相应流量{Qmax} m³/s，发生于{year} 年{month} 月。",
        "受下游顶托影响，高水段水位流量关系呈绳套状，需分涨落分别定线。",
        "设计洪水成果与上下游相邻河段成果对比，洪峰模数{mod} m³/s/km²，合理性较好。",
        "泥沙多年平均输沙量{sed} 万 t，汛期占全年{sedge}%。",
    ],
    "geology": [
        "工程区出露地层主要为白垩系上统南雄组紫红色砂岩、泥质粉砂岩。",
        "区域构造位于抚州凹陷南缘，断裂以北东向为主，规模较小。",
        "河床覆盖层为第四系冲积层，厚度{qh} m，以砂卵石为主，结构稍密。",
        "堤基上部为粉质黏土，厚{ct} m，承载力特征值{fak} kPa。",
        "地下水位埋深{dw} m，对混凝土无腐蚀性，对钢结构具弱腐蚀性。",
        "根据钻探揭露，岩体风化带厚度{wf} m，弱风化岩体完整性较好。",
        "场地类别为Ⅱ类，地震动峰值加速度{pg} g，抗震设防烈度Ⅵ度。",
        "天然建筑材料储量满足要求，砂砾料运距{dis} km，土料运距{dis2} km。",
        "主要工程地质问题为堤基渗漏与砂性土液化，需采取相应处理措施。",
    ],
    "layout": [
        "工程总体布置沿现有堤线加固，堤顶宽度{bw} m，堤顶高程按设计洪水位加超高{sup} m 确定。",
        "堤防级别为{lv} 级，迎水侧边坡 1:{s1}，背水侧边坡 1:{s2}。",
        "护岸工程采用生态联锁砖护坡，厚度{thk} m，下设砂砾石垫层。",
        "穿堤建筑物共{cnt} 座，均为涵闸，设计流量{q} m³/s。",
        "进水闸闸室为开敞式钢筋混凝土结构，闸孔净宽{bk} m，共{kong} 孔。",
        "堤身填筑料采用黏性土，压实度不小于{comp}%，渗透系数按规范要求控制。",
        "堤防与公路交叉段设置防汛道路，路面宽{road} m，采用沥青混凝土。",
        "工程管理范围线按堤脚外{mg} m 控制，界桩{zhuang} 个。",
        "为满足生态要求，堤顶设置绿化带，宽度{gr} m，种植乡土树种。",
        "建筑物地基采用水泥土搅拌桩处理，桩径{dp} m，桩长{lp} m。",
    ],
    "construction": [
        "施工总工期{gt} 个月，其中准备期{gb} 个月，主体工程施工期{zm} 个月。",
        "施工导流采用分期围堰，一期围堰高{ph} m，堰顶宽{pw} m。",
        "河床砂卵石料丰富，混凝土骨料可就近开采，运距{dis} km。",
        "施工用水从河道抽取，用电由附近 10 kV 线路接入，设变压器{tr} 台。",
        "主体工程量：土方开挖{tw:.1f} 万 m³，土方填筑{tf:.1f} 万 m³，混凝土{conc:.1f} 万 m³。",
        "高峰期施工人数{pp} 人，施工机械{jj} 台（套）。",
        "施工期需做好水土保持，设置临时排水沟{dg} m、沉沙池{cc} 座。",
        "围堰拆除安排在枯水期，拆除料用于堤后压渗填筑。",
        "施工组织采用分段流水作业，分为{duan} 个作业面同步推进。",
        "质量检测按《水利水电工程施工质量检验与评定规程》执行，抽检频率{rate}%。",
    ],
    "conclusion": [
        "本工程实施后，治理河段防洪标准提高到{f} 年一遇，保护人口{pop} 万人。",
        "工程经济内部收益率{irr:.1f}%，效益费用比{bc:.2f}，经济上合理。",
        "工程建成后可减少洪灾损失{loss} 万元/年，社会效益显著。",
        "建议加快初步设计批复，落实建设资金，尽早开工实施。",
        "建议在施工期加强水位观测，及时调整围堰高程，确保度汛安全。",
        "下阶段应结合施工详勘，进一步优化堤线布置与建筑物型式。",
        "本阶段成果满足可行性研究深度要求，可作为下阶段设计依据。",
        "建议建立工程运行管理机制，明确管护责任主体与经费来源。",
        "工程实施需与河道采砂管理、岸线保护规划相衔接。",
        "结论：工程建设必要性充分，技术方案可行，建议尽快实施。",
    ],
}


def gen_body(pool_name, rng, used):
    """Return a unique paragraph of MIN_LEN..MAX_LEN Chinese chars."""
    pool = POOLS[pool_name]
    for _ in range(120):
        k = rng.choice([2, 3, 3, 3, 4])
        clauses = rng.sample(pool, k)
        n = nums(rng)
        text = "".join(c.format(**n) for c in clauses)
        if len(text) > MAX_LEN:
            n = nums(rng)
            text = "".join(c.format(**n) for c in clauses[:2])
        if len(text) < MIN_LEN:
            continue
        if len(text) > MAX_LEN:
            continue
        if text in used:
            continue
        used.add(text)
        return text
    raise RuntimeError("cannot generate body paragraph for pool %s" % pool_name)


# --------------------------------------------------------------------------
# document styling
# --------------------------------------------------------------------------
def set_style_font(style, ascii_font, ea_font, size, bold=None):
    style.font.name = ascii_font
    style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    style.font.color.rgb = RGBColor(0, 0, 0)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), ascii_font)
    rfonts.set(qn("w:hAnsi"), ascii_font)
    rfonts.set(qn("w:eastAsia"), ea_font)
    rfonts.set(qn("w:cs"), ascii_font)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]


def setup_styles(doc):
    normal = doc.styles["Normal"]
    set_style_font(normal, "Times New Roman", "宋体", 12)
    pf = normal.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    h1 = doc.styles["Heading 1"]
    set_style_font(h1, "Times New Roman", "黑体", 16, bold=True)
    h1.paragraph_format.space_before = Pt(18)
    h1.paragraph_format.space_after = Pt(10)
    h1.paragraph_format.keep_with_next = True
    h1.paragraph_format.page_break_before = True
    h1.paragraph_format.line_spacing = 1.5

    h2 = doc.styles["Heading 2"]
    set_style_font(h2, "Times New Roman", "黑体", 14, bold=True)
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after = Pt(6)
    h2.paragraph_format.keep_with_next = True
    h2.paragraph_format.line_spacing = 1.5

    title = doc.styles["Title"]
    set_style_font(title, "Times New Roman", "黑体", 22, bold=True)
    title.paragraph_format.space_before = Pt(110)
    title.paragraph_format.space_after = Pt(24)
    title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.line_spacing = 1.5


def setup_page(doc):
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.54)
    sec.bottom_margin = Cm(2.54)
    sec.left_margin = Cm(3.0)
    sec.right_margin = Cm(2.5)


def setup_run_font(run, ea="宋体", size=None, bold=None):
    run.font.name = "Times New Roman"
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), ea)


# --------------------------------------------------------------------------
# content plan
# --------------------------------------------------------------------------
TARGET_A = "根据本次实测断面资料，进水闸孔口高程为 52.36 m，与原设计图纸相差 0.12 m，需在施工图阶段复核。"
TARGET_B = "根据水文计算成果，P=1% 设计洪峰流量大概 1234 m3/s，相应水位为 68.42 m，与 2019 年实测洪痕基本吻合。"
TARGET_C = "施工期项目部组织对上下游围堰进行了进行了三次检查，未发现集中渗漏，堰体稳定满足要求。"
EXTRA_R4 = "根据现场踏勘，坝址河段比较顺直，河床宽约 120 m，两岸为第四系冲积层，堤线布置条件较好。"
EXTRA_R5 = "根据本次审查意见，建议下阶段补充地质勘探工作量，重点查明断层破碎带与软弱夹层的分布范围。"

SPECIAL_SLOTS = {
    "2.2 设计洪水": {1: ("target_B", TARGET_B)},
    "3.2 工程地质条件": {1: ("target_A", TARGET_A)},
    "4.2 堤防工程": {1: ("r4", EXTRA_R4)},
    "5.2 施工导流": {1: ("target_C", TARGET_C)},
    "6.2 建议": {1: ("r5", EXTRA_R5)},
}
SET_TEXT_SLOTS = {
    "1.3 主要建设内容": (3, "set_1"),
    "4.1 工程总体布置": (2, "set_2"),
    "5.3 施工进度": (4, "set_3"),
}
INSERT_AFTER = {
    # section title -> [(tag, offset from the end of the section's paragraphs), ...]
    # Both anchors sit at the very end of the document so that no other edited
    # paragraph index is shifted by the insertions.
    "6.2 建议": [("ins_1", 2), ("ins_2", 1)],
}

CHAPTERS = [
    ("1 工程概况", [
        ("1.1 流域概况", "overview"),
        ("1.2 工程任务与规模", "overview"),
        ("1.3 主要建设内容", "overview"),
    ]),
    ("2 水文", [
        ("2.1 测站与资料", "hydrology"),
        ("2.2 设计洪水", "hydrology"),
        ("2.3 水位流量关系", "hydrology"),
    ]),
    ("3 工程地质", [
        ("3.1 区域地质", "geology"),
        ("3.2 工程地质条件", "geology"),
        ("3.3 天然建筑材料", "geology"),
    ]),
    ("4 工程布置及建筑物", [
        ("4.1 工程总体布置", "layout"),
        ("4.2 堤防工程", "layout"),
        ("4.3 建筑物设计", "layout"),
    ]),
    ("5 施工组织设计", [
        ("5.1 施工条件", "construction"),
        ("5.2 施工导流", "construction"),
        ("5.3 施工进度", "construction"),
    ]),
    ("6 结论与建议", [
        ("6.1 结论", "conclusion"),
        ("6.2 建议", "conclusion"),
    ]),
]

BODY_TARGET = 250  # 292 python paragraphs = 6 cover + 1 toc head + 10 toc + 24 headings + 250 body + 1 caption
TOC_ENTRIES = [
    "1 工程概况", "2 水文", "3 工程地质", "4 工程布置及建筑物",
    "5 施工组织设计", "6 结论与建议",
    "1.1 流域概况", "2.2 设计洪水", "4.2 堤防工程", "6.2 建议",
]


def build_plan(rng):
    plan = []

    def P(role, text, style=None, tag=None):
        plan.append({"kind": "p", "role": role, "text": text, "style": style, "tag": tag})

    # cover -----------------------------------------------------------------
    P("cover_title", "江西省白沙河流域防洪治理工程可行性研究报告", style="Title")
    P("cover_sub", "（送审稿）")
    P("blank", "")
    P("cover_unit", "编制单位：江西省水利规划设计研究院")
    P("cover_date", "编制日期：二〇二五年十二月")
    P("blank", "")

    # table of contents -----------------------------------------------------
    P("toc_head", "目　　录")
    for i, entry in enumerate(TOC_ENTRIES):
        dots = "…" * max(6, 26 - len(entry) // 2)
        P("toc", "%s %s %d" % (entry, dots, i + 1))

    # sections --------------------------------------------------------------
    sections = [(ch, sec, pool) for ch, secs in CHAPTERS for sec, pool in secs_norm(secs)]
    quotas = {}
    total = 0
    for idx, (_, title, _) in enumerate(sections):
        q = 12 + (1 if idx < 18 else 0) + (1 if idx < 16 else 0)
        quotas[title] = q
        total += q
    # exact fit to BODY_TARGET
    diff = BODY_TARGET - total
    keys = [t for _, t, _ in sections]
    i = 0
    while diff != 0:
        k = keys[i % len(keys)]
        quotas[k] += 1 if diff > 0 else -1
        diff += -1 if diff > 0 else 1
        i += 1

    used = set()
    sec_body_tags = {}
    for chapter_title, secs in CHAPTERS:
        P("h1", chapter_title, style="Heading 1")
        for sec_title, pool in secs_norm(secs):
            P("h2", sec_title, style="Heading 2")
            specials = SPECIAL_SLOTS.get(sec_title, {})
            set_slot = SET_TEXT_SLOTS.get(sec_title)
            body_tags = []
            for pos in range(quotas[sec_title]):
                tag = None
                text = None
                if pos in specials:
                    tag, text = specials[pos]
                elif set_slot and pos == set_slot[0]:
                    tag = set_slot[1]
                if text is None:
                    text = gen_body(pool, rng, used)
                P("body", text, style=None, tag=tag)
                if tag:
                    body_tags.append((pos, tag))
            if sec_title in INSERT_AFTER:
                for tag, back in INSERT_AFTER[sec_title]:
                    plan[-back]["tag"] = tag
            sec_body_tags[sec_title] = body_tags
            # table sits at the end of 2.2
            if sec_title == "2.2 设计洪水":
                P("caption", "表 1 主要水文特征值表")
                plan.append({"kind": "table"})
    return plan


def secs_norm(secs):
    return secs


# --------------------------------------------------------------------------
# rendering
# --------------------------------------------------------------------------
def add_table(doc):
    data = [
        ["断面名称", "集水面积（km²）", "P=1% 洪峰流量（m³/s）", "相应水位（m）"],
        ["坝址断面", "1280", "1234", "68.42"],
        ["河口断面", "2140", "1980", "61.75"],
    ]
    table = doc.add_table(rows=3, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.line_spacing = 1.0
            run = p.add_run(val)
            setup_run_font(run, ea="宋体", size=10.5, bold=(r == 0))
    return table


def render(doc, plan):
    tags = {}
    py_index = 0
    for item in plan:
        if item["kind"] == "table":
            add_table(doc)
            continue
        role = item["role"]
        text = item["text"]
        style = item.get("style")
        p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
        run = p.add_run(text)
        pf = p.paragraph_format
        if role == "cover_title":
            pass  # Title style handles it
        elif role == "cover_sub":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.space_before = Pt(6)
            setup_run_font(run, ea="宋体", size=16)
        elif role == "cover_unit" or role == "cover_date":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.space_before = Pt(6)
            setup_run_font(run, ea="宋体", size=14)
        elif role == "toc_head":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.space_before = Pt(12)
            pf.space_after = Pt(18)
            setup_run_font(run, ea="黑体", size=18, bold=True)
        elif role == "toc":
            pf.line_spacing = 1.5
            setup_run_font(run, ea="宋体", size=12)
        elif role == "h1" or role == "h2":
            pass  # heading styles handle it
        elif role == "caption":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.space_before = Pt(6)
            pf.space_after = Pt(3)
            pf.line_spacing = 1.0
            setup_run_font(run, ea="黑体", size=10.5)
        elif role == "body":
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            pf.first_line_indent = Pt(24)
            pf.line_spacing = 1.5
        tag = item.get("tag")
        if tag:
            tags[tag] = py_index
        py_index += 1
    return tags


def build_wps_map(doc):
    """Map 0-based python-paragraph index -> 1-based WPS Paragraphs index.

    WPS counts the paragraphs inside table cells inline *plus one empty
    "row end" paragraph per table row* (measured: a 3x4 table occupies 15
    paragraphs, not 12).  python-docx does not expose either.
    """
    mapping = {}
    py = -1
    wps = 0
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            py += 1
            wps += 1
            mapping[py] = wps
        elif child.tag == qn("w:tbl"):
            for tr in child.findall(qn("w:tr")):
                for tc in tr.findall(qn("w:tc")):
                    for p in tc.findall(qn("w:p")):
                        wps += 1
                wps += 1  # row-end paragraph mark, counted by WPS
    return mapping, wps


def main():
    os.makedirs(CORPUS, exist_ok=True)
    rng = random.Random(20251023)

    plan = build_plan(rng)

    doc = Document()
    setup_styles(doc)
    setup_page(doc)
    tags = render(doc, plan)

    cp = doc.core_properties
    cp.title = "江西省白沙河流域防洪治理工程可行性研究报告"
    cp.author = "江西省水利规划设计研究院"
    cp.subject = "wps-word-cooperate skill acceptance corpus"
    cp.comments = "synthetic test document"

    doc.save(OUT_DOCX)

    # plaintext sanity check (must NOT be encrypted yet)
    with open(OUT_DOCX, "rb") as fh:
        head = fh.read(16)
    if not zipfile.is_zipfile(OUT_DOCX):
        raise SystemExit("generated docx is not a zip (head=%s) - did WPS save it?" % head.hex())

    # reload to get real paragraph texts (what WPS will see)
    doc2 = Document(OUT_DOCX)
    py_texts = [p.text for p in doc2.paragraphs]
    mapping, wps_total = build_wps_map(doc2)

    # map tag -> py -> wps
    tag_info = {}
    for tag, py in tags.items():
        if py not in mapping:
            raise SystemExit("tag %s not found in body" % tag)
        tag_info[tag] = {"py": py, "wps": mapping[py], "text": py_texts[py]}

    # ---- assertions ------------------------------------------------------
    assert 290 <= len(py_texts) <= 320, "python paragraph count %d out of range" % len(py_texts)
    assert 290 <= wps_total <= 320, "wps paragraph count %d out of range" % wps_total
    assert len(doc2.tables) == 1, "expected exactly 1 table"
    body_like = [t for t in py_texts if t]
    for t in body_like:
        if len(t) > 200:  # headings/cover/toc excluded loosely
            continue
    # body length check only for body-role paragraphs recorded in plan
    body_texts = [it["text"] for it in plan if it.get("role") == "body"]
    assert len(body_texts) == BODY_TARGET, "body count %d != %d" % (len(body_texts), BODY_TARGET)
    bad = [t for t in body_texts if not (MIN_LEN <= len(t) <= MAX_LEN)]
    assert not bad, "body length out of range: %r" % bad[:3]

    # find strings must be unique in the whole document
    find_map = {
        "target_A": "孔口高程",
        "target_B": "大概 1234 m3/s",
        "target_C": "进行了进行",
        "r4": "河段比较顺直",
        "r5": "建议下阶段补充",
    }
    all_texts = list(py_texts)
    for tb in doc2.tables:
        for row in tb.rows:
            for cell in row.cells:
                all_texts.append(cell.text)
    for tag, needle in find_map.items():
        hits = sum(t.count(needle) for t in all_texts)
        assert hits == 1, "find %r appears %d times (expected 1)" % (needle, hits)

    # ---- targets.json ----------------------------------------------------
    rep = {
        "target_A": ("孔口高程", "孔口地面高程"),
        "target_B": ("大概 1234 m3/s", "约 1234 m³/s"),
        "target_C": ("进行了进行", "进行了"),
        "r4": ("河段比较顺直", "河段较为顺直"),
        "r5": ("建议下阶段补充", "建议下阶段进一步补充"),
    }
    targets = []
    for tag, (find, repl) in rep.items():
        info = tag_info[tag]
        original = info["text"]
        assert find in original, "%s: find not in original" % tag
        assert original.count(find) == 1
        targets.append({
            "para": info["wps"],
            "original": original,
            "find": find,
            "replace": repl,
            "kind": "target" if tag.startswith("target_") else "extra",
        })
    targets.sort(key=lambda x: x["para"])
    with open(TARGETS_JSON, "w", encoding="utf-8") as fh:
        json.dump(targets, fh, ensure_ascii=False, indent=2)

    meta = {
        "doc": OUT_DOCX,
        "counts": {
            "python_paragraphs": len(py_texts),
            "wps_paragraphs": wps_total,
            "body_paragraphs": len(body_texts),
            "tables": len(doc2.tables),
            "table_paragraphs": wps_total - len(py_texts),
        },
        "tags": tag_info,
        "python_to_wps": {str(k): v for k, v in mapping.items()},
    }
    with open(META_JSON, "w", encoding="utf-8") as fh:
        json.dump(meta, fh, ensure_ascii=False, indent=2)

    # ---- pristine baseline copy -----------------------------------------
    shutil.copy2(OUT_DOCX, PRISTINE)

    print(json.dumps({
        "docx": OUT_DOCX,
        "pristine": PRISTINE,
        "paragraphs_python": len(py_texts),
        "paragraphs_wps": wps_total,
        "body_paragraphs": len(body_texts),
        "tables": len(doc2.tables),
        "targets": [{"para": t["para"], "kind": t["kind"], "find": t["find"]} for t in targets],
        "head_hex": head.hex(),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
