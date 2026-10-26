# -*- coding: utf-8 -*-
"""生成 Word 表格题注/编号清单回归语料（幂等脚本）。

用法（本机 Python）：
    & C:\\Users\\Administrator\\.dsh\\dsh-runtimes\\dsh-primary-runtime\\dependencies\\python\\python.exe make_fixture.py

产物（均在本脚本所在目录）：
    captions-fixture.docx   语料本体：约 20 张表、8 张 35x13 大表、3 章正文
    expected.json           期望值：19 条题注 + 刻意埋入的缺陷清单

特性：
  * 每次运行都整体重建 docx 与 expected.json，可反复覆盖；
  * 不含随机数、时间戳等非确定内容，两次运行的 docx 字节数一致；
  * 生成后自动重新打开 docx 做一遍自检并打印报告（见 self_check）。

文档结构（题注位置按规格交替）：
  above    题注为表格上方紧邻的独立段落（居中、无首行缩进），题注与表格之间无空段；
  in-table 题注写在表格第 1 行（该行合并为 1 格），数据从第 2 行开始。
"""

from __future__ import annotations

import datetime
import io
import json
import sys
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

HERE = Path(__file__).resolve().parent
DOCX = HERE / "captions-fixture.docx"
EXPECTED = HERE / "expected.json"

# ---- 规格常量 ---------------------------------------------------------------
FW_SPACE = "\u3000"        # 全角空格
HW_SPACE2 = "  "           # 两个半角空格
BIG_ROWS, BIG_COLS = 35, 13
BIG_TABLES = (3, 4, 5, 8, 9, 10, 11, 19)   # 8 张 35 行 x 13 列大表
SMALL_ROWS = 6

# 表序, 表号, 标题, 题注位置, 表号与标题间的空白, 小表列数（4~6）
SPECS = [
    (1,  "2.1.1", "峡江站天然径流成果表",           "above",    HW_SPACE2, 5),
    (2,  "2.1.2", "峡江站天然径流月分配表",         "in-table", FW_SPACE,  4),
    (3,  "2.2.1", "峡江站逐日平均流量表（一）",     "above",    "",        6),
    (4,  "2.2.2", "峡江站逐日平均流量表（二）",     "in-table", HW_SPACE2, 6),
    (5,  "2.2.3", "峡江站逐日平均流量表（三）",     "above",    HW_SPACE2, 6),
    (6,  "2.4.1", "坝址水位流量关系",               "above",    HW_SPACE2, 4),
    (7,  "2.4.3", "坝前水位流量关系",               "above",    HW_SPACE2, 4),
    (8,  "3.1-1", "设计洪水成果表",                 "in-table", HW_SPACE2, 6),
    (9,  "3.1-2", "分期洪水成果表",                 "in-table", HW_SPACE2, 6),
    (10, "3.2-1", "库容曲线",                       "above",    HW_SPACE2, 6),
    (11, "3.2-2", "泄流能力曲线",                   "above",    HW_SPACE2, 6),
    (12, "3.3-1", "调度规则表",                     "above",    HW_SPACE2, 6),
    (13, "3.3-2", "水库运行调度规则",               "above",    HW_SPACE2, 5),
    (14, "3.3-3", "Q泄i-1与△Q关系",                 "above",    HW_SPACE2, 4),
    (15, "3.3-3", "泄流能力复核表",                 "in-table", HW_SPACE2, 5),
    (16, "3.3-4", "水库特征水位表",                 "above",    HW_SPACE2, 6),
    (17, "4.1-1", "工程量汇总表",                   "in-table", HW_SPACE2, 5),
    (18, "4.1-2", "投资估算表",                     "above",    HW_SPACE2, 6),
    (20, "4.2-1", "结论对照表",                     "above",    HW_SPACE2, 5),
]

# 表 19 刻意无题注：上方是一段含「如表中数据所示」的正常正文，表内第 1 行是普通表头。
NO_CAPTION_TABLES = (19,)
# 表 19 表头（13 列，前三列必须为 项目 / 单位 / 数量）
T19_HEADER = [
    "项目", "单位", "数量", "建筑工程", "机电设备", "金属结构",
    "临时工程", "合计", "备注", "复核", "审核", "批准", "日期",
]

# 小表表头（列数须与 SPECS 中一致）
SMALL_HEADERS = {
    1:  ["项目", "单位", "均值", "Cv", "Cs"],
    2:  ["月份", "径流量(亿m³)", "占比(%)", "备注"],
    6:  ["水位(m)", "流量(m³/s)", "面积(m²)", "流速(m/s)"],
    7:  ["水位(m)", "流量(m³/s)", "面积(m²)", "流速(m/s)"],
    12: ["时段", "起调水位(m)", "控制水位(m)", "下泄流量(m³/s)", "调度方式", "备注"],
    13: ["编号", "适用时段", "控制条件", "下泄流量(m³/s)", "备注"],
    14: ["水位(m)", "Q泄(m³/s)", "△Q(m³/s)", "备注"],
    15: ["水位(m)", "原泄流量(m³/s)", "复核泄流量(m³/s)", "差值(m³/s)", "备注"],
    16: ["序号", "特征水位", "单位", "数值", "对应库容(亿m³)", "备注"],
    17: ["序号", "项目", "单位", "数量", "备注"],
    18: ["序号", "项目", "单位", "数量", "单价(元)", "合价(万元)"],
    20: ["序号", "项目", "结论", "依据", "备注"],
}

TITLE_TEXT = "新干枢纽扩容工程可行性研究报告（节选）"

# 正文：("h2", 标题) / ("p", 段落) / ("t", 表序)
SECTIONS = [
    ("2 水文", [
        ("p", "赣江中游河段的水文情势直接决定新干枢纽的规模与调度方式。本章依据峡江站等测站的实测系列，分析径流、洪水与水位流量关系。"),
        ("p", "测站资料系列截至 2023 年，按规范要求进行了一致性、可靠性与代表性分析。"),
        ("p", "本章同时说明资料的引用与换算基准，未作特殊说明时水位均采用 1985 国家高程基准。"),
        ("p", "流量成果按瞬时流量与日平均流量两种口径给出，使用时应注意区分。"),
        ("h2", "2.1 基本资料"),
        ("p", "峡江站位于赣江中游，集水面积约 3.7 万 km²，为国家基本水文站，观测项目包括水位、流量、含沙量与降水量。"),
        ("p", "本次共收集到该站 1953—2023 年共 71 年实测径流资料，系列长度满足规范要求。"),
        ("p", "测站高程系统已统一换算至 1985 国家高程基准。"),
        ("p", "流量资料采用流速仪法施测，测次分布基本均匀，高水期测次偏少。"),
        ("p", "资料系列按一致性、可靠性、代表性三项要求进行了审查，未发现明显跳变或系统偏差。"),
        ("p", "含沙量资料自 1965 年起施测，本次仅用于泥沙分析，未参与径流计算。"),
        ("p", "降水量资料用于径流还原与产流参数率定。"),
        ("p", "径流成果按年、月两个尺度统计，月分配采用多年平均月径流过程。"),
        ("p", "年径流统计采用水文年，水文年起始日为 4 月 1 日。"),
        ("p", "径流还原计算采用分项还原法，还原水量占实测径流的比例小于 15%。"),
        ("p", "天然径流系列经还原后长度仍为 71 年，可满足频率计算要求。"),
        ("t", 1),
        ("p", "如表2.1.1所示，峡江站多年平均径流量为 682 亿 m³，年径流变差系数 Cv 为 0.22。"),
        ("p", "径流的年内分配不均，汛期 4—9 月径流量约占全年的七成。"),
        ("p", "月分配成果按各月多年平均值统计，各月合计与年径流量一致。"),
        ("t", 2),
        ("p", "受上游水库调度影响，枯水期月平均流量较天然情况有所增大。"),
        ("p", "本次径流成果已与上下游相邻站的成果进行对照，沿程变化合理。"),
        ("p", "月分配成果同时用于兴利调节计算与发电量估算。"),
        ("p", "月分配成果的不确定度主要来自上游水库的调度方式。"),
        ("p", "成果的精度满足可行性研究阶段的要求。"),
        ("h2", "2.2 径流与流量"),
        ("p", "为反映逐日流量过程，按丰、平、枯三个代表年整理逐日平均流量成果。"),
        ("p", "代表年的选择以年径流量接近设计频率为原则，兼顾系列的代表性。"),
        ("p", "逐日流量成果由实测水位经水位流量关系换算得到。"),
        ("p", "日平均流量按 8 时、20 时观测值算术平均后换算。"),
        ("p", "三个代表年的年径流量分别为 905 亿 m³、668 亿 m³、452 亿 m³。"),
        ("t", 3),
        ("p", "丰水年的流量过程陡涨陡落，汛期最大日平均流量超过 8000 m³/s。"),
        ("p", "逐日流量成果按日历日排列，缺测日按前后日插补。"),
        ("p", "平水年的流量过程相对平缓，枯水期流量维持在 200 m³/s 左右。"),
        ("t", 4),
        ("p", "平水年成果主要用于发电量估算与出力过程分析。"),
        ("p", "代表年的洪水过程与实测过程基本一致。"),
        ("p", "逐日成果的统计口径为水文年，跨年时段按起始年归属。"),
        ("p", "枯水年的枯季流量明显偏小，对下游供水与生态流量保障提出更高要求。"),
        ("t", 5),
        ("p", "枯水年成果用于检验生态流量控制指标的满足程度。"),
        ("p", "上述三组逐日流量成果作为调节计算与动能经济分析的输入。"),
        ("p", "三组成果的年内总量与相应频率的年径流量偏差在 2% 以内。"),
        ("p", "成果已用于调节计算，调节流量过程与实测过程趋势一致。"),
        ("p", "施工期导流的分月设计流量按分期洪水成果推算。"),
        ("p", "逐日流量过程同时用于施工期导流的分月设计流量估算。"),
        ("p", "成果的时段划分与水文年一致。"),
        ("h2", "2.4 水位流量关系"),
        ("p", "坝址河段河道顺直、断面规整，具备稳定的水位流量关系。"),
        ("p", "关系曲线采用实测大断面与实测流量点据综合拟定，高水部分按水力学方法延长。"),
        ("p", "大断面测量采用全站仪，测点间距不大于 5 m。"),
        ("p", "断面成果已换算至统一基面，左右岸高程采用同一系统。"),
        ("p", "定线时按水位分级检验，各级水位下的流量偏差控制在 3% 以内。"),
        ("p", "水位流量关系的定线精度按规范要求检验，定线误差满足要求。"),
        ("p", "高水部分的延长采用水力学公式与上下游站对照两种方法互校。"),
        ("t", 6),
        ("p", "坝址断面的水位流量关系在高水部分受下游回水影响略有抬升。"),
        ("p", "关系曲线以单值形式给出，回水影响通过分段修正。"),
        ("t", 7),
        ("p", "坝前水位流量关系用于泄流能力与调洪演算的衔接。"),
        ("p", "上、下游关系曲线的差值反映了坝前壅水的高度。"),
        ("p", "各项关系的适用范围为坝前水位 24.0～38.0 m。"),
        ("p", "关系曲线在汛期与枯期分别检验，未发现明显偏离。"),
        ("p", "关系曲线每年汛前复核一次，遇较大洪水后加测校核。"),
        ("p", "曲线的控制断面与坝址断面位置一致，未作距离改正。"),
        ("p", "如表2.9.9所示，各频率设计洪水成果见附表。"),
        ("p", "上述关系在枢纽运行期应定期复核，遇河道采砂等人类活动影响时及时修正。"),
    ]),
    ("3 工程任务和规模", [
        ("p", "新干枢纽以发电、航运为主，兼顾防洪与水资源配置。"),
        ("p", "工程规模论证的核心是设计洪水、库容与泄流能力、调度规则三项。"),
        ("p", "本章各节成果经专业间协调后确定。"),
        ("p", "各节成果的编制口径与第 2 章一致。"),
        ("h2", "3.1 设计洪水"),
        ("p", "设计洪水采用实测洪水系列排频计算，历史洪水按文献考证成果加入系列。"),
        ("p", "洪水系列采用 1953—2023 年共 71 年实测资料，另加入 2 场历史洪水。"),
        ("p", "历史洪水的洪峰流量按洪痕调查与水力学方法推算。"),
        ("p", "系列中的特大洪水按重现期作经验处理，并按规范要求进行连续性改正。"),
        ("p", "洪峰与各时段洪量采用同一典型年放大，成果相互匹配。"),
        ("p", "频率计算采用 P-Ⅲ 型曲线适线，参数以矩法与适线法结合确定。"),
        ("p", "频率曲线的统计参数采用目估适线，兼顾上下段点据。"),
        ("p", "设计断面的洪水成果按面积比指数法移用，指数取 0.7。"),
        ("t", 8),
        ("p", "设计洪水成果经与上下游梯级对比协调后确定，详见表3.1-1。"),
        ("p", "各频率设计洪水过程按典型年同倍比放大。"),
        ("p", "成果与邻站同期成果对比，峰量关系协调。"),
        ("p", "分期洪水按汛前、汛后两个分期分别计算。"),
        ("p", "分期的划分依据流域暴雨的时程分布特征。"),
        ("t", 9),
        ("p", "分期洪水成果用于施工期导流与初期蓄水安排。"),
        ("p", "汛前分期的洪水明显小于主汛期，可提高施工期度汛标准。"),
        ("p", "汛后分期的洪水过程对蓄水时机选择影响较大。"),
        ("p", "各频率成果均已按规范要求计及历史洪水的影响。"),
        ("p", "分期洪水的分期界限按暴雨季节划分，汛前为 4—6 月。"),
        ("p", "分期成果已考虑分期内的洪水遭遇情况。"),
        ("h2", "3.2 库容与泄流能力"),
        ("p", "库容曲线依据 1:2000 库区地形图量算，分层间距 0.5 m。"),
        ("p", "量算范围自坝址至库尾回水末端，支流库容单独计列。"),
        ("p", "量算成果已与上一阶段成果对比，偏差小于 3%。"),
        ("p", "库容曲线分高程列出面积与库容，便于插值使用。"),
        ("p", "库区地形图的比例尺与精度满足可研阶段要求。"),
        ("p", "库容成果已扣除库区已有的侵占部分。"),
        ("t", 10),
        ("p", "泄流能力曲线按闸孔出流与堰流两种流态分别计算。"),
        ("p", "闸门全开时按堰流控制，局部开启时按闸孔出流控制。"),
        ("p", "泄流能力计算采用堰流与闸孔出流公式，系数取自规范。"),
        ("p", "闸门全开与局部开启两种工况分别给出泄流量。"),
        ("t", 11),
        ("p", "泄流能力规模按校核洪水位下泄量满足要求确定。"),
        ("p", "泄流建筑物的孔口尺寸与闸门型式已与金结专业核实。"),
        ("p", "曲线计算中糙率取 0.014，与模型试验成果基本一致。"),
        ("p", "泄流能力复核时采用设计洪水与校核洪水两种工况。"),
        ("p", "泄流能力曲线与库容曲线共同用于调洪演算。"),
        ("h2", "3.3 调度规则"),
        ("p", "水库调度遵循防洪优先、兴利服从防洪的原则。"),
        ("p", "调度规则按汛期、非汛期与特殊工况分别拟定。"),
        ("p", "调度规则的编制依据为现行调度规程与工程任务要求。"),
        ("p", "调度图按保证出力与防洪安全双目标编制。"),
        ("p", "汛期限制水位按防洪要求与兴利蓄水要求综合确定。"),
        ("p", "汛期限制水位以下预留防洪库容，汛末按计划蓄水。"),
        ("t", 12),
        ("p", "各项规则的适用条件与下泄控制指标见表内数值，具体条文见表3.3-2。"),
        ("p", "特殊工况包括闸门故障、超标准洪水与生态应急调度。"),
        ("p", "各工况下的下泄流量按泄流能力曲线反算。"),
        ("p", "规则条文的编号与下泄流量分级一一对应。"),
        ("t", 13),
        # 表 13 之后不留空段，紧接本段（inTable 误判复现）
        ("p", "为满足表3.3-2的控制要求，泄流能力按上述曲线确定，相应坝前水位按水位流量关系内插。"),
        ("p", "内插计算采用分段线性方法，控制精度为 0.01 m。"),
        ("p", "泄流量的增量关系按水位分级统计，用于控制下泄流量。"),
        ("t", 14),
        ("p", "增量关系用于调度图的下泄流量修正。"),
        ("p", "泄流能力复核按最新闸门尺寸与糙率取值进行。"),
        ("t", 15),
        ("p", "复核成果与原设计成果基本一致，个别水位级差在 2% 以内。"),
        ("p", "水库特征水位按规范规定的荷载组合确定。"),
        ("t", 16),
        ("p", "特征水位成果作为大坝安全与库区淹没处理的依据。"),
        ("p", "调度规则在运行期应根据实测资料滚动修订。"),
        ("p", "规则条文的表述与运行单位的操作规程保持一致。"),
        ("p", "调度规则的修改须经主管部门批准后执行。"),
        ("p", "特征水位包括正常蓄水位、死水位、防洪高水位与校核洪水位。"),
        ("p", "各特征水位对应的库容已与库容曲线核对。"),
        ("p", "调度规则与特征水位成果共同构成运行调度的依据。"),
    ]),
    ("4 结论", [
        ("p", "本章汇总工程量、投资与主要结论。"),
        ("p", "各项成果均按可行性研究阶段的深度要求编制。"),
        ("p", "投资估算的价格水平年与工程量计算口径一致。"),
        ("h2", "4.1 工程量与投资"),
        ("p", "工程量按建筑工程、机电设备及安装工程、金属结构设备及安装工程、临时工程四部分计列。"),
        ("p", "工程量计算依据为水工专业提供的单体工程量与施工组织设计。"),
        ("p", "建筑工程量按分部分项计列，土石方按自然方与压实方分别统计。"),
        ("p", "机电设备及安装工程量按台套计列，安装费按设备原价的百分比计取。"),
        ("p", "各分部工程量按单位与数量分别统计，计量单位与概算定额一致。"),
        ("p", "金属结构设备及安装工程量按吨位计列。"),
        ("p", "临时工程量包括导流工程、施工交通与施工房屋建筑。"),
        ("t", 17),
        ("p", "汇总工程量与各专业提供的成果一致，未计列重复项目。"),
        ("p", "汇总表中的合计数量与各分部成果逐项核对。"),
        ("p", "投资估算按 2024 年第二季度价格水平编制。"),
        ("p", "投资估算由建筑工程费、机电设备及安装工程费、金属结构设备及安装工程费、临时工程费与其他费用构成。"),
        ("p", "其他费用包括建设管理费、生产准备费与科研勘测设计费。"),
        ("p", "建筑工程按扩大指标法估算，机电与金结设备按厂家询价计取。"),
        ("t", 18),
        ("p", "投资估算成果与工程量汇总成果相互对应，如表4.1-2所列。"),
        ("p", "基本预备费按工程费用的 5% 计取，价差预备费按现行规定不计。"),
        ("p", "建设期利息与生产准备费在下一阶段投资中计列。"),
        ("p", "估算成果已与同类工程指标进行类比分析。"),
        ("h2", "4.2 结论"),
        ("p", "经水文、规模与投资三方面论证，新干枢纽扩容工程在技术经济上可行。"),
        ("p", "工程建成后可增加年发电量约 3.6 亿 kW·h。"),
        ("p", "枢纽建成后可改善上游航道条件约 28 km。"),
        ("p", "防洪方面可提高下游保护区的防洪标准。"),
        ("p", "各项工程量与投资汇总如表中数据所示，其中投资估算与表4.1-2所列成果一致，各分部工程的单位与数量已逐项复核。"),
        ("t", 19),
        ("p", "上述汇总成果不含建设期利息与生产准备费。"),
        ("p", "主要结论与相关章节成果的对照关系见表内数值。"),
        ("t", 20),
        ("p", "结论对照表中各项指标均满足可研阶段的深度要求。"),
        ("p", "结论对照表中各项指标的口径与相关章节一致。"),
        ("p", "下一阶段应结合施工组织设计进一步复核工程量。"),
        ("p", "工程的防洪、发电、航运效益均较显著。"),
        ("p", "建议下阶段补充库区淹没调查与移民安置规划工作。"),
        ("p", "建议下阶段开展水工模型试验，验证泄流能力。"),
        ("p", "建议同步开展环境影响评价与水土保持方案编制。"),
    ]),
]


# ---- 排版辅助 ---------------------------------------------------------------
def _set_east_asian(style_or_run, font_name: str) -> None:
    """给样式或 run 同时设置西文与中文字体。"""
    element = style_or_run.element if hasattr(style_or_run, "element") else style_or_run._element
    rpr = element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)
    rfonts.set(qn("w:eastAsia"), font_name)


def setup_styles(doc: Document) -> None:
    """设置 A4 页面与中文字体（正文宋体五号）。"""
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.17)
    section.right_margin = Cm(3.17)

    normal = doc.styles["Normal"]
    normal.font.name = "宋体"
    normal.font.size = Pt(10.5)          # 五号
    _set_east_asian(normal, "宋体")

    for name, size in (("Heading 1", 16), ("Heading 2", 14)):
        style = doc.styles[name]
        style.font.name = "黑体"
        style.font.size = Pt(size)
        _set_east_asian(style, "黑体")


def add_body_paragraph(doc: Document, text: str):
    """正文段落：首行缩进两字符、两端对齐。"""
    p = doc.add_paragraph(text)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.first_line_indent = Pt(21)        # 2 字符 x 10.5pt
    pf.line_spacing = 1.5
    return p


def add_caption_paragraph(doc: Document, text: str):
    """above 型题注：居中、无首行缩进的独立段落。"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.first_line_indent = Pt(0)
    pf.line_spacing = 1.5
    p.add_run(text)
    return p


# ---- 表格构造 ---------------------------------------------------------------
def caption_text(number: str, gap: str, title: str) -> str:
    return f"表{number}{gap}{title}"


def para_text(element) -> str:
    """取一个 w:p 元素的可见文字。

    注意：不能直接用 lxml 的 ``itertext()``——python-docx 给 ``w:r`` / ``w:p``
    加了 ``text`` 属性，``itertext()`` 会把同一段文字重复计 3 次。
    """
    return "".join(t.text or "" for t in element.iter(qn("w:t")))


def _fill_numeric(table, start_row: int, start_value: int) -> None:
    """从 start_row 起逐行填连续数字，保证每个单元格非空。"""
    value = start_value
    for row in table.rows[start_row:]:
        for cell in row.cells:
            cell.text = str(value)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.font.size = Pt(9)       # 小五，避免 13 列大表撑得过宽
            value += 1


def _merge_caption_row(table, text: str) -> None:
    """把第 1 行合并为 1 格，只放题注文字。"""
    row = table.rows[0]
    merged = row.cells[0].merge(row.cells[-1])
    # 合并后原各格的空段落要清掉，只留一个段落放题注
    for extra in merged.paragraphs[1:]:
        extra._element.getparent().remove(extra._element)
    para = merged.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.text = text


def build_table(doc: Document, spec, index: int):
    """按规格插入一张表；返回 python-docx 的 Table 对象。"""
    has_caption = spec is not None
    is_big = (spec[0] if has_caption else index) in BIG_TABLES
    rows = BIG_ROWS if is_big else SMALL_ROWS
    cols = BIG_COLS if is_big else spec[5]

    table = doc.add_table(rows=rows, cols=cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # 表格内字号小五，避免 13 列大表撑得离谱
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)

    if not has_caption:                       # 表 19：普通表头 + 数字
        for c, head in enumerate(T19_HEADER[:cols]):
            cell = table.cell(0, c)
            cell.text = head
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _fill_numeric(table, 1, 1)
        return table

    _, number, title, pos, gap, _cols = spec
    text = caption_text(number, gap, title)

    if pos == "in-table":
        _merge_caption_row(table, text)
        _fill_numeric(table, 1, 1)
    else:
        header = SMALL_HEADERS.get(spec[0]) or []
        if is_big:
            _fill_numeric(table, 0, 1)        # 大表直接填数字（规格允许 1..455）
        else:
            if len(header) != cols:
                raise ValueError(f"表 {spec[0]} 表头列数 {len(header)} 与 {cols} 不符")
            for c, head in enumerate(header):
                cell = table.cell(0, c)
                cell.text = head
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            _fill_numeric(table, 1, 1)
    return table


# ---- 主流程 -----------------------------------------------------------------
def build() -> None:
    doc = Document()
    setup_styles(doc)

    cp = doc.core_properties
    cp.title = "captions 回归语料"
    cp.author = "fixture-dev"
    cp.last_modified_by = "fixture-dev"
    cp.revision = 1
    cp.created = datetime.datetime(2026, 1, 1, 0, 0, 0)
    cp.modified = datetime.datetime(2026, 1, 1, 0, 0, 0)

    title_para = doc.add_paragraph(TITLE_TEXT)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_para.paragraph_format.space_before = Pt(12)
    title_para.paragraph_format.space_after = Pt(12)
    for run in title_para.runs:
        run.font.name = "宋体"
        run.font.size = Pt(18)
        run.bold = True
        _set_east_asian(run, "宋体")

    specs_by_index = {spec[0]: spec for spec in SPECS}
    placed = []

    for h1, items in SECTIONS:
        doc.add_heading(h1, level=1)
        for kind, payload in items:
            if kind == "h2":
                doc.add_heading(payload, level=2)
            elif kind == "p":
                add_body_paragraph(doc, payload)
            elif kind == "t":
                spec = specs_by_index.get(payload)
                if spec is not None and spec[3] == "above":
                    add_caption_paragraph(doc, caption_text(spec[1], spec[4], spec[2]))
                build_table(doc, spec, payload)
                placed.append(payload)
            else:  # pragma: no cover
                raise ValueError(kind)

    if sorted(placed) != sorted([s[0] for s in SPECS] + list(NO_CAPTION_TABLES)):
        raise RuntimeError(f"表格插入顺序/数量不符：{placed}")

    DOCX.parent.mkdir(parents=True, exist_ok=True)
    save_deterministic(doc, DOCX)


def save_deterministic(doc: Document, path: Path) -> None:
    """保存 docx 并把 zip 条目时间戳固定，使两次运行产出字节一致。

    python-docx 保存时用当前时间做 zip 条目时间，虽不影响内容与大小，
    但会让每次运行的字节流不同。这里先在内存里生成，再按固定时间写出。
    """
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    with zipfile.ZipFile(buf) as src:
        entries = [(info.filename, src.read(info.filename)) for info in src.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst:
        for name, data in entries:
            zi = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            dst.writestr(zi, data)


def build_expected() -> dict:
    captions = [
        {
            "table": spec[0],
            "label": "表",
            "number": spec[1],
            "title": spec[2],
            "pos": spec[3],
            "gap": spec[4],
            "style": None,
        }
        for spec in SPECS
    ]
    dot = [spec[1] for spec in SPECS if "." in spec[1] and "-" not in spec[1]]
    hyphen = [spec[1] for spec in SPECS if "-" in spec[1]]
    return {
        "doc": "captions-fixture.docx",
        "tables": 20,
        "bigTables": 8,
        "bigTableSize": [BIG_ROWS, BIG_COLS],
        "captions": captions,
        "defects": {
            "duplicates": [{"number": "3.3-3", "tables": [14, 15]}],
            "sequenceGaps": [
                {"chapter": "2", "after": "2.4.1", "before": "2.4.3", "missing": ["2.4.2"]}
            ],
            "formatMixed": {"dot": dot, "hyphen": hyphen},
            "missingCaptionTables": list(NO_CAPTION_TABLES),
            "spaceMixed": True,
            "danglingReferences": ["2.9.9"],
            "resolvedReferences": ["2.1.1", "3.1-1", "3.3-2", "4.1-2"],
            "proseAfterTable": {"afterTable": 13, "text": "为满足表3.3-2的控制要求"},
        },
        "notes": (
            "题注位置交替：above 为表格上方紧邻的独立段落（居中、无缩进、题注后无空段），"
            "in-table 为表格第 1 行合并单元格。表号与标题之间的空白逐表不同（两个半角空格 / "
            "一个全角空格 / 无空格），用于验证空格不一场景。第 2 章用点分表号 2.1.1，"
            "第 3、4 章用短横表号 3.1-1；表 14 与表 15 重号 3.3-3；2.4 组缺 2.4.2；"
            "正文含悬空引用 2.9.9 与 4 条正常引用；表 13 之后不留空段直接接正文；"
            "表 19 无题注，上方正文含「如表中数据所示」，表内第 1 行为普通表头。"
        ),
    }


def write_expected() -> None:
    data = build_expected()
    EXPECTED.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


# ---- 自检 -------------------------------------------------------------------
def self_check() -> int:
    """重新打开生成的 docx，逐条核对规格并打印报告。返回失败项数量。"""
    doc = Document(str(DOCX))
    body = doc.element.body
    children = list(body.iterchildren())
    p_tag, tbl_tag = qn("w:p"), qn("w:tbl")

    problems: list[str] = []
    total_p = sum(1 for _ in body.iter(p_tag))
    body_p = sum(1 for c in children if c.tag == p_tag)
    tbl_elements = [c for c in children if c.tag == tbl_tag]
    chars = sum(len(t.text or "") for t in body.iter(qn("w:t")))

    print("=" * 68)
    print("自检报告：captions-fixture.docx")
    print("=" * 68)
    print(f"文件大小            : {DOCX.stat().st_size:,} 字节")
    print(f"表格数              : {len(doc.tables)}（含表格元素 {len(tbl_elements)}）")
    print(f"全部 w:p 段落数     : {total_p}（含表格单元格）")
    print(f"正文级段落数        : {body_p}")
    print(f"字符数（w:t 文本）  : {chars}")

    if len(doc.tables) != 20:
        problems.append(f"表格数 {len(doc.tables)} != 20")
    if not (4000 <= total_p <= 6000):
        problems.append(f"全部 w:p 段落数 {total_p} 不在 4000–6000")

    big = [
        i + 1
        for i, t in enumerate(doc.tables)
        if len(t.rows) == BIG_ROWS and len(t.columns) == BIG_COLS
    ]
    print(f"35x13 大表序号      : {big}")
    if big != sorted(BIG_TABLES):
        problems.append(f"大表序号 {big} != {sorted(BIG_TABLES)}")

    print("-" * 68)
    print("逐表行列数 / 题注：")
    for i, t in enumerate(doc.tables, start=1):
        # 表 19 之外，取题注：above 取上一段，in-table 取第 1 行合并格
        spec = next((s for s in SPECS if s[0] == i), None)
        if spec is None:
            prev = tbl_elements[i - 1].getprevious()
            prev_text = para_text(prev) if prev is not None else ""
            print(f"  #{i:>2} {len(t.rows):>2}x{len(t.columns):<2} 无题注"
                  f"  上方段落含『如表中数据所示』={'如表中数据所示' in prev_text}"
                  f"  表头={'|'.join(c.text for c in t.rows[0].cells[:3])}")
            if "如表中数据所示" not in prev_text:
                problems.append("表 19 上方段落缺『如表中数据所示』")
            continue
        if spec[3] == "in-table":
            got = t.rows[0].cells[0].text
        else:
            prev = tbl_elements[i - 1].getprevious()
            got = para_text(prev) if prev is not None else ""
        want = caption_text(spec[1], spec[4], spec[2])
        flag = "OK " if got == want else "BAD"
        print(f"  #{i:>2} {len(t.rows):>2}x{len(t.columns):<2} {flag} {spec[3]:<8} [{got}]")
        if got != want:
            problems.append(f"表 {i} 题注不符：{got!r} != {want!r}")

    # 表 13 之后紧邻段落
    after13 = tbl_elements[12].getnext()
    after13_text = para_text(after13) if after13 is not None else ""
    print("-" * 68)
    print(f"表 13 之后紧邻元素     : {after13.tag.split('}')[-1] if after13 is not None else None}")
    print(f"表 13 之后段落正文     : {after13_text}")
    if after13 is None or after13.tag != p_tag or not after13_text.startswith(
        "为满足表3.3-2的控制要求"
    ):
        problems.append("表 13 之后未紧跟指定正文段落")

    # 文本统计
    all_text = "\n".join(para_text(p) for p in body.iter(p_tag))
    dup = all_text.count("表3.3-3")
    dang = all_text.count("表2.9.9")
    print(f"『表3.3-3』出现次数   : {dup}")
    print(f"『表2.9.9』出现次数   : {dang}")
    for ref in ("如表2.1.1所示", "详见表3.1-1", "见表3.3-2", "如表4.1-2所列"):
        hit = all_text.count(ref)
        print(f"  resolved 引用 {ref:<12}: {hit} 次")
        if hit < 1:
            problems.append(f"缺正常引用：{ref}")
    if dup != 2:
        problems.append(f"『表3.3-3』出现 {dup} 次 != 2")
    if dang != 1:
        problems.append(f"『表2.9.9』出现 {dang} 次 != 1")

    # 题注位置核对（above 必须紧邻表格、上方无空段；in-table 第 1 行只有题注）
    for i, t in enumerate(doc.tables, start=1):
        spec = next((s for s in SPECS if s[0] == i), None)
        if spec is None:
            continue
        want = caption_text(spec[1], spec[4], spec[2])
        if spec[3] == "above":
            prev = tbl_elements[i - 1].getprevious()
            if prev is None or prev.tag != p_tag or para_text(prev) != want:
                problems.append(f"表 {i} 的 above 题注未紧邻表格")
            if spec[0] in (1, 3, 5, 6, 7, 10, 11, 12, 13, 14, 16, 18, 20):
                # 题注与表格之间不应有空段：上一元素必须是题注段落本身（已核对）
                pass
        else:
            if t.rows[0].cells[0].text != want:
                problems.append(f"表 {i} 的 in-table 题注不符")

    # 段间不应有连续空段（题注后不得有空段）
    empty_in_row = 0
    for prev, cur in zip(children, children[1:]):
        if prev.tag == p_tag and cur.tag == p_tag:
            if not para_text(prev).strip() and not para_text(cur).strip():
                empty_in_row += 1
    print(f"连续空段对数         : {empty_in_row}")
    if empty_in_row:
        problems.append(f"存在 {empty_in_row} 处连续空段")

    print("-" * 68)
    if problems:
        print(f"自检失败（{len(problems)} 项）：")
        for item in problems:
            print("  - " + item)
    else:
        print("自检全部通过。")
    print("=" * 68)
    return len(problems)


def main(argv: list[str]) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    build()
    write_expected()
    print(f"已生成：{DOCX}")
    print(f"已生成：{EXPECTED}")
    return 0 if self_check() == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
