# -*- coding: utf-8 -*-
"""
废旧轮胎的回收与资源化利用 — Chinese academic deck (review / evidence-map arc).
Built with python-pptx per the nature-paper2ppt skill toolchain policy.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

# ---------------- palette ----------------
INK     = "2B2B2B"   # tire-black text
FOREST  = "1F4E3D"   # primary dark green
GREEN   = "2C5F2D"   # green
MOSS    = "8FB996"   # light moss
MOSSBG  = "E4EEE7"   # very light green panel
AMBER   = "E8A33D"   # accent energy
TERRA   = "BE6B3A"   # pyrolysis oil
STEEL   = "6E7B72"   # muted grey-green (captions)
LIGHT   = "F2F5F2"   # light panel
WHITE   = "FFFFFF"
LINEGR  = "D2DCD5"   # light divider

EA   = "Microsoft YaHei"
EAB  = "Microsoft YaHei"   # use same family; weight via bold
LAT  = "Arial"

EMU_IN = 914400
SW, SH = 13.333, 7.5

prs = Presentation()
prs.slide_width  = Inches(SW)
prs.slide_height = Inches(SH)
BLANK = prs.slide_layouts[6]

# ---------------- helpers ----------------
def slide():
    return prs.slides.add_slide(BLANK)

def _set_run(r, txt, sz, b, c, ea=EA, lat=LAT, italic=False):
    r.text = txt
    f = r.font
    f.size = Pt(sz)
    f.bold = b
    f.italic = italic
    f.color.rgb = RGBColor.from_string(c)
    f.name = lat
    rPr = r._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set('typeface', ea)

def rect(s, l, t, w, h, fill=None, line=None, line_w=1.0, shape=MSO_SHAPE.RECTANGLE,
         shadow=False):
    sp = s.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = RGBColor.from_string(line); sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    if shadow:
        el = sp._element.spPr
        ef = el.makeelement(qn('a:effectLst'), {})
        sh = el.makeelement(qn('a:outerShdw'),
            {'blurRad':'90000','dist':'38100','dir':'5400000','rotWithShape':'0'})
        clr = el.makeelement(qn('a:srgbClr'), {'val':'9AA89E'})
        alpha = el.makeelement(qn('a:alpha'), {'val':'42000'})
        clr.append(alpha); sh.append(clr); ef.append(sh); el.append(ef)
    return sp

def text(s, l, t, w, h, paras, anchor=MSO_ANCHOR.TOP, wrap=True,
         ml=0.0, mr=0.0, mt=0.0, mb=0.0):
    """paras: list of dicts: {runs:[(txt,sz,bold,color, [ea],[lat],[italic])],
       align, space_before, space_after, line}"""
    tb = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(ml); tf.margin_right = Inches(mr)
    tf.margin_top = Inches(mt); tf.margin_bottom = Inches(mb)
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get('align', PP_ALIGN.LEFT)
        if 'space_after' in p:  para.space_after = Pt(p['space_after'])
        if 'space_before' in p: para.space_before = Pt(p['space_before'])
        if 'line' in p:
            para.line_spacing = p['line']
        for rdef in p['runs']:
            r = para.add_run()
            txt_, sz, b, c = rdef[0], rdef[1], rdef[2], rdef[3]
            ea = rdef[4] if len(rdef) > 4 and rdef[4] else EA
            lat = rdef[5] if len(rdef) > 5 and rdef[5] else LAT
            it = rdef[6] if len(rdef) > 6 else False
            _set_run(r, txt_, sz, b, c, ea, lat, it)
    return tb

def circle(s, cx, cy, d, fill, line=None, line_w=1.0):
    return rect(s, cx-d/2, cy-d/2, d, d, fill=fill, line=line, line_w=line_w,
                shape=MSO_SHAPE.OVAL)

def notes(s, txt):
    s.notes_slide.notes_text_frame.text = txt

def kicker(s, label, x=0.7, y=0.55):
    """small green dot + uppercase kicker label above title"""
    circle(s, x+0.07, y+0.09, 0.16, AMBER)
    text(s, x+0.22, y-0.07, 6.0, 0.34,
         [{'runs':[(label, 12, True, GREEN)]}], anchor=MSO_ANCHOR.MIDDLE)

def title(s, t, y=0.86, x=0.7, w=12.0, color=INK, sz=29):
    text(s, x, y, w, 0.95,
         [{'runs':[(t, sz, True, color)], 'line':1.02}], anchor=MSO_ANCHOR.TOP)

def bottomstrip(s, txt, color=STEEL):
    text(s, 0.7, 7.02, 12.0, 0.34,
         [{'runs':[(txt, 9, False, color)]}], anchor=MSO_ANCHOR.MIDDLE)

# =====================================================================
# Slide 1 — Cover
# =====================================================================
s = slide()
rect(s, 0, 0, SW, SH, fill=FOREST)
rect(s, 0, 0, SW, SH*0.5, fill=FOREST)  # keep solid
# motif: stacked tire rings (concentric ovals) bottom-right, subtle
for i, d in enumerate([3.6, 2.7, 1.8]):
    circle(s, 11.6, 6.2, d, fill=None, line=MOSS, line_w=1.4)
circle(s, 11.6, 6.2, 0.9, fill=GREEN, line=MOSS, line_w=1.4)
# tag
rect(s, 0.9, 1.35, 2.5, 0.42, fill=AMBER)
text(s, 0.9, 1.35, 2.5, 0.42, [{'runs':[('环境工程 · 资源循环', 11.5, True, INK)],
     'align':PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)
# title
text(s, 0.85, 2.15, 10.5, 2.2, [
    {'runs':[('废旧轮胎的回收', 50, True, WHITE)], 'line':1.04},
    {'runs':[('与', 50, True, MOSS),('资源化利用', 50, True, AMBER)], 'line':1.04},
], anchor=MSO_ANCHOR.TOP)
# subtitle
text(s, 0.9, 4.55, 10.0, 0.7,
     [{'runs':[('从“黑色污染”到循环经济：技术路线、价值分级与发展趋势', 16, False, MOSSBG)]}])
# divider dot row
circle(s, 1.0, 5.45, 0.12, AMBER)
text(s, 1.2, 5.27, 10.0, 0.4,
     [{'runs':[('专题汇报', 13, True, WHITE)]}], anchor=MSO_ANCHOR.MIDDLE)
# footer meta band
text(s, 0.9, 6.55, 8.0, 0.5, [
    {'runs':[('汇报人：（姓名）', 12, False, MOSSBG)]},
    {'runs':[('环境工程专业 · 2026', 11, False, MOSS)]},
])
notes(s, "开场白：大家好。今天我汇报的主题是废旧轮胎的回收与资源化利用。"
          "轮胎是现代交通不可缺少的部件，但报废之后却是典型的固体废物——俗称'黑色污染'。"
          "我会从它为什么是个环境难题讲起，再依次介绍主流的资源化技术路线、它们的对比、"
          "目前的争议与瓶颈，最后落到循环经济和未来发展方向。整个汇报大约十三分钟。")

# =====================================================================
# Slide 2 — Why it matters: scale + hazards (big stats)
# =====================================================================
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
kicker(s, "为什么重要 · WHY NOW")
title(s, "废旧轮胎：增长最快的固体废物之一，俗称“黑色污染”")
# three big stat callouts
stats = [
    ("约 3 亿条", "我国每年新增报废轮胎（行业公开估算）", AMBER),
    ("逾 1000 万吨", "对应年报废重量，且仍逐年增长", GREEN),
    (">100 年", "自然环境中难以降解的时间尺度", TERRA),
]
x0, w, gap = 0.7, 3.86, 0.27
for i,(big,small,col) in enumerate(stats):
    x = x0 + i*(w+gap)
    rect(s, x, 1.95, w, 1.75, fill=LIGHT, shadow=True)
    rect(s, x, 1.95, 0.12, 1.75, fill=col)
    text(s, x+0.32, 2.12, w-0.5, 0.85,
         [{'runs':[(big, 33, True, col)]}], anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+0.32, 2.98, w-0.55, 0.6,
         [{'runs':[(small, 12, False, INK)], 'line':1.12}], anchor=MSO_ANCHOR.TOP)
# hazard band
rect(s, 0.7, 4.05, 11.93, 2.55, fill=MOSSBG)
text(s, 1.05, 4.25, 11.3, 0.5,
     [{'runs':[('堆存与不当处置带来的环境与安全风险', 16, True, FOREST)]}])
haz = [
    ("占地与蚊媒", "整胎中空、不可压缩，长期露天堆存，积水滋生蚊虫、传播疫病"),
    ("火灾隐患", "胶料热值高、一旦起火极难扑灭，释放浓烟与有毒气体"),
    ("土壤与水体", "添加剂、重金属及微塑料随风化淋溶，污染土壤和地下水"),
    ("资源浪费", "含优质橡胶、炭黑与钢丝，简单填埋即丢弃了可回收资源"),
]
cw = (11.3-0.3)/2
for i,(h,d) in enumerate(haz):
    cx = 1.05 + (i%2)*(cw+0.3)
    cy = 4.85 + (i//2)*0.85
    circle(s, cx+0.12, cy+0.13, 0.22, FOREST)
    text(s, cx+0.34, cy-0.05, cw-0.4, 0.8, [
        {'runs':[(h+'  ', 12.5, True, INK),(d, 11, False, INK)], 'line':1.1}])
bottomstrip(s, "数据为行业公开统计/估算，用于说明问题规模；不同来源口径略有差异。")
notes(s, "首先说明问题的规模。我国是轮胎生产和消费大国，每年新增报废轮胎大约三亿条，"
          "对应重量超过一千万吨，而且随着汽车保有量上升还在增长。轮胎以橡胶为主、加上炭黑和钢丝，"
          "在自然环境里上百年都难以降解。如果只是露天堆放或填埋，会带来四类风险："
          "一是整胎中空积水、滋生蚊虫；二是极易起火且很难扑灭；三是添加剂和重金属淋溶污染土壤水体；"
          "四是把本可回收的优质资源白白浪费。所以废旧轮胎既是环境负担，也是放错位置的资源。"
          "强调：这里的数字是行业公开估算，用来说明量级。")

# =====================================================================
# Slide 3 — Tire composition (doughnut chart): a misplaced resource
# =====================================================================
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
kicker(s, "资源属性 · COMPOSITION")
title(s, "轮胎是“放错位置的资源”：可回收组分占绝大多数")
# left text rail
text(s, 0.7, 2.1, 4.7, 3.6, [
    {'runs':[('一条轮胎并非单一材料，而是橡胶、炭黑、钢丝与纤维的复合体。',13.5,False,INK)],'line':1.28,'space_after':10},
    {'runs':[('• ',13.5,True,GREEN),('天然/合成橡胶',13.5,True,INK),
             ('——可再生胶、可热解回收',12,False,STEEL)],'space_after':7,'line':1.2},
    {'runs':[('• ',13.5,True,GREEN),('炭黑',13.5,True,INK),
             ('——补强填料，热解可回收再炭黑',12,False,STEEL)],'space_after':7,'line':1.2},
    {'runs':[('• ',13.5,True,GREEN),('钢丝',13.5,True,INK),
             ('——可磁选回收，回炉炼钢',12,False,STEEL)],'space_after':7,'line':1.2},
    {'runs':[('• ',13.5,True,GREEN),('纺织纤维及添加剂',13.5,True,INK)],'space_after':12,'line':1.2},
])
rect(s, 0.7, 5.55, 4.7, 1.05, fill=MOSSBG)
text(s, 0.95, 5.62, 4.3, 0.95,
     [{'runs':[('结论：',12.5,True,FOREST),('约 9 成质量是可循环材料——回收的核心在于“高值化”而非简单处置。',12.5,False,INK)],'line':1.18}],
     anchor=MSO_ANCHOR.MIDDLE)
# doughnut chart on right
cd = CategoryChartData()
cd.categories = ['橡胶（天然+合成）', '炭黑', '钢丝', '纤维/织物', '其他添加剂']
cd.add_series('组成', (45, 25, 15, 5, 10))
gframe = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(5.9), Inches(1.95),
                            Inches(6.7), Inches(4.9), cd)
chart = gframe.chart
chart.has_title = True
chart.chart_title.text_frame.text = "典型乘用车轮胎的质量组成（近似）"
tfont = chart.chart_title.text_frame.paragraphs[0].runs[0].font
tfont.size = Pt(13); tfont.bold = True; tfont.name = EA
tfont.color.rgb = RGBColor.from_string(INK)
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.RIGHT
chart.legend.include_in_layout = False
chart.legend.font.size = Pt(11); chart.legend.font.name = EA
chart.legend.font.color.rgb = RGBColor.from_string(INK)
plot = chart.plots[0]
plot.has_data_labels = True
dl = plot.data_labels
dl.number_format = '0"%"'; dl.number_format_is_linked = False
dl.font.size = Pt(11); dl.font.bold = True; dl.font.name = EA
dl.font.color.rgb = RGBColor.from_string(WHITE)
seg_colors = [GREEN, INK, STEEL, MOSS, AMBER]
for i, pt in enumerate(plot.series[0].points):
    pt.format.fill.solid(); pt.format.fill.fore_color.rgb = RGBColor.from_string(seg_colors[i])
bottomstrip(s, "整理自橡胶工业公开资料；不同胎型/规格组分比例有差异。")
notes(s, "这一页回答'为什么值得回收'。一条轮胎不是单一材料，而是橡胶、炭黑、钢丝和纤维的复合体。"
          "右边的环形图给出典型乘用车胎的质量组成：橡胶约百分之四十五，炭黑约四分之一，钢丝约百分之十五，"
          "其余是纤维和各种添加剂。关键结论是——大约九成质量都是可循环利用的材料。"
          "所以问题的核心不是'怎么扔掉'，而是'怎么把这些资源高值化地拿回来'，这就引出后面的技术路线。")

# =====================================================================
# Slide 4 — Framework: 4R hierarchy (full-width process)
# =====================================================================
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
kicker(s, "全景框架 · FRAMEWORK")
title(s, "资源化利用的总体框架：按“价值优先”分级的 4R 路线")
# four stage cards with descending value indication
stages = [
    ("减量化", "Reduce", "延长寿命·耐磨设计·合理充气\n从源头减少报废量", FOREST),
    ("再利用", "Reuse", "翻新轮胎·整胎工程利用\n保留制品形态，能耗最低", GREEN),
    ("再生利用", "Recycle", "胶粉、再生胶、热解回收\n材料重新进入产业链", AMBER),
    ("能量回收", "Recover", "作替代燃料（TDF）\n回收热值，价值最低", TERRA),
]
n=4; gap=0.3; w=(11.93-(n-1)*gap)/n; x0=0.7; y=2.25; h=2.5
for i,(zh,en,desc,col) in enumerate(stages):
    x=x0+i*(w+gap)
    rect(s, x, y, w, h, fill=LIGHT, line=col, line_w=1.5, shadow=True)
    rect(s, x, y, w, 0.78, fill=col)
    text(s, x, y+0.02, w, 0.78, [
        {'runs':[(zh, 17, True, WHITE)],'align':PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)
    text(s, x, y+0.82, w, 0.32, [
        {'runs':[(en, 11, True, col)],'align':PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+0.18, y+1.2, w-0.36, h-1.3, [
        {'runs':[(ln, 11.5, False, INK)],'align':PP_ALIGN.CENTER,'line':1.18}
        for ln in desc.split('\n')], anchor=MSO_ANCHOR.TOP)
    if i < n-1:
        ax = x+w+0.02
        text(s, ax, y+0.7, gap, 0.6, [{'runs':[('▶',14,True,STEEL)],'align':PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)
# value-priority arrow band
rect(s, 0.7, 5.05, 11.93, 0.55, fill=FOREST)
text(s, 1.0, 5.05, 6.0, 0.55, [{'runs':[('资源/环境价值：高',12,True,WHITE)]}], anchor=MSO_ANCHOR.MIDDLE)
text(s, 6.6, 5.05, 5.93, 0.55, [{'runs':[('低（应尽量减少占比）',12,True,MOSSBG)],'align':PP_ALIGN.RIGHT}],
     anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.7, 5.85, 11.93, 0.95,
     [{'runs':[('核心思想：',13,True,FOREST),
               ('优先延长寿命与再利用，其次让材料循环，最后才是能量回收与填埋——后续每条路线都对应其中一个层级。',13,False,INK)],'line':1.25}])
bottomstrip(s, "框架依据固体废物管理“减量—再利用—再生—回收”通行优先级整理。")
notes(s, "这页是全篇的'地图'。废物管理有一个通行的优先级，常被概括为4R：减量化、再利用、再生利用、能量回收，"
          "价值和环境效益从左到右递减。减量化是从源头减少报废，比如耐磨设计、合理充气延长寿命；"
          "再利用是保留制品形态，比如翻新和整胎工程利用，能耗最低；再生利用是把材料重新拿回来，"
          "包括胶粉、再生胶和热解；能量回收是当替代燃料、只回收热值，价值最低。"
          "核心思想就一句话：能高值利用就不要低值处置。后面四条路线，正好对应这张图里的不同层级。")

# =====================================================================
# Slide 5 — Route 1: reuse & retreading
# =====================================================================
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
kicker(s, "路径一 · REUSE")
title(s, "原形利用与翻新：能耗最低、价值最高的“第一选择”")
# left: two sub-blocks
def routeblock(x,y,w,h,head,col,items):
    rect(s, x, y, w, h, fill=LIGHT, shadow=True)
    rect(s, x, y, w, 0.62, fill=col)
    text(s, x+0.25, y, w-0.4, 0.62, [{'runs':[(head,15,True,WHITE)]}], anchor=MSO_ANCHOR.MIDDLE)
    paras=[]
    for it in items:
        paras.append({'runs':[('• ',12.5,True,col),(it[0]+'  ',12.5,True,INK),(it[1],11.5,False,STEEL)],
                      'line':1.16,'space_after':8})
    text(s, x+0.28, y+0.78, w-0.5, h-0.9, paras, anchor=MSO_ANCHOR.TOP)

routeblock(0.7, 2.05, 5.85, 2.25, "轮胎翻新（Retreading）", GREEN, [
    ("旧胎体 + 新胎面","对仍完好的胎体重新硫化贴面"),
    ("载重/航空胎为主","可翻新 1–3 次，按里程更划算"),
    ("节能减排显著","较新胎大幅降低原材料与能耗"),
])
routeblock(0.7, 4.45, 5.85, 2.2, "整胎直接利用", AMBER, [
    ("土木工程","护坡、挡土墙、河岸防护"),
    ("港口/船舶","码头护舷、缓冲防撞"),
    ("生态与民用","人工鱼礁、游乐设施、隔音墙"),
])
# right: highlight panel — value-first message + simple bar of energy
rect(s, 6.85, 2.05, 5.78, 4.6, fill=FOREST)
text(s, 7.2, 2.3, 5.2, 0.8, [{'runs':[('为什么把它放在第一位？',17,True,WHITE)],'line':1.1}])
text(s, 7.2, 3.05, 5.15, 2.0, [
    {'runs':[('保留了轮胎的',13,False,MOSSBG),('制品形态与结构',13,True,WHITE),
             ('，无需破碎、再加工或高温处理。',13,False,MOSSBG)],'line':1.35,'space_after':10},
    {'runs':[('因此',13,False,MOSSBG),('单位能耗与碳排放最低',13,True,AMBER),
             ('，资源价值保留最完整。',13,False,MOSSBG)],'line':1.35},
])
# mini comparison chips
chips=[("翻新一条载重胎","≈ 节省可观原料与能耗"),("整胎工程利用","近零再加工能耗")]
for i,(a,b) in enumerate(chips):
    cy=5.25+i*0.62
    circle(s, 7.42, cy+0.2, 0.2, AMBER)
    text(s, 7.7, cy, 4.7, 0.5, [{'runs':[(a+'：',12,True,WHITE),(b,12,False,MOSSBG)]}],
         anchor=MSO_ANCHOR.MIDDLE)
bottomstrip(s, "翻新与整胎利用对应 4R 中的“再利用”层级，是优先推荐的处置方式。")
notes(s, "从这页开始进入具体的技术路线，顺序大致按价值从高到低。"
          "第一条是原形利用和翻新，也是最该优先做的。翻新是把还完好的胎体重新贴一层新胎面再硫化，"
          "在载重胎和航空胎上很常见，一条胎体可以翻新一到三次，按行驶里程算很划算。"
          "整胎直接利用则是把废胎用在土木工程的护坡挡墙、港口码头的护舷防撞、以及人工鱼礁、隔音墙等。"
          "它放第一位的原因在右边：保留了轮胎的形态和结构，不需要破碎和高温处理，"
          "所以单位能耗和碳排放最低，资源价值保留得最完整。")

# =====================================================================
# Slide 6 — Route 2: crumb rubber
# =====================================================================
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
kicker(s, "路径二 · CRUMB RUBBER")
title(s, "胶粉利用：把废胎磨成“原料”，规模化进入材料市场")
# top: mini process flow
flow = ["废旧轮胎","破碎/除杂","常温或低温研磨","分级胶粉(目数)","下游应用"]
fx=0.7; fw=2.18; fy=1.95; fh=0.72; fg=0.28
for i,st in enumerate(flow):
    x=fx+i*(fw+fg)
    col = AMBER if i==len(flow)-1 else GREEN
    rect(s, x, fy, fw, fh, fill=(col if i in (0,len(flow)-1) else LIGHT),
         line=GREEN, line_w=1.2, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    tcol = WHITE if i in (0,len(flow)-1) else INK
    text(s, x, fy, fw, fh, [{'runs':[(st,11.5,True,tcol)],'align':PP_ALIGN.CENTER,'line':1.05}],
         anchor=MSO_ANCHOR.MIDDLE)
    if i < len(flow)-1:
        text(s, x+fw, fy, fg, fh, [{'runs':[('→',15,True,STEEL)],'align':PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)
# four application cards
apps = [
    ("橡胶改性沥青","掺入胶粉改性道路沥青，提升抗裂、抗车辙与降噪，是吸纳量最大的出路之一", AMBER),
    ("弹性铺装","塑胶跑道、运动场地、儿童活动场的弹性面层与垫层", GREEN),
    ("橡胶制品","防水卷材、密封件、橡胶地砖、汽车配件等再制造", TERRA),
    ("建材与填料","改性混凝土、轻质砌块、铁路道砟垫层等土木应用", FOREST),
]
cw=(11.93-3*0.3)/4; cy=3.15; ch=2.55
for i,(h,d,col) in enumerate(apps):
    x=0.7+i*(cw+0.3)
    rect(s, x, cy, cw, ch, fill=LIGHT, shadow=True)
    circle(s, x+cw/2, cy+0.62, 0.66, col)
    text(s, x+cw/2-0.33, cy+0.3, 0.66, 0.66, [{'runs':[(str(i+1),20,True,WHITE)],'align':PP_ALIGN.CENTER}],
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+0.1, cy+1.08, cw-0.2, 0.45, [{'runs':[(h,12.5,True,INK)],'align':PP_ALIGN.CENTER,'line':1.05}],
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+0.18, cy+1.55, cw-0.36, ch-1.6, [{'runs':[(d,10.5,False,STEEL)],'align':PP_ALIGN.CENTER,'line':1.18}],
         anchor=MSO_ANCHOR.TOP)
rect(s, 0.7, 5.95, 11.93, 0.72, fill=MOSSBG)
text(s, 1.0, 5.95, 11.4, 0.72,
     [{'runs':[('要点：',12.5,True,FOREST),('胶粉技术成熟、规模大，但多为“等值或降级”利用；其中橡胶改性沥青兼顾用量与性能，被视为重点方向。',12.5,False,INK)],'line':1.18}],
     anchor=MSO_ANCHOR.MIDDLE)
notes(s, "第二条路线是胶粉利用，也是目前规模最大的方向之一。流程是把废胎破碎除掉钢丝和纤维，"
          "再常温或低温研磨成不同目数的胶粉，然后进入下游。最重要的四个出路："
          "一是橡胶改性沥青，把胶粉掺进道路沥青，能提升抗裂、抗车辙和降噪，吸纳量很大；"
          "二是弹性铺装，比如塑胶跑道和运动场地；三是各类再生橡胶制品；四是建材和土木填料。"
          "要点是：胶粉技术成熟、规模大，但大多属于等值或降级利用；其中改性沥青因为用量大、性能好，被当作重点推广方向。")

# =====================================================================
# Slide 7 — Route 3: reclaimed rubber / devulcanization
# =====================================================================
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
kicker(s, "路径三 · RECLAIM")
title(s, "再生胶与脱硫：让交联橡胶重新“可塑化”")
# left: concept — why hard + how
text(s, 0.7, 2.0, 6.0, 1.5, [
    {'runs':[('硫化使橡胶分子形成三维交联网络，',13.5,False,INK),
             ('既不熔融也不溶解',13.5,True,TERRA),
             ('，这正是橡胶难以回收的根源。',13.5,False,INK)],'line':1.3,'space_after':8},
    {'runs':[('再生（脱硫）的目标：',13.5,True,FOREST),
             ('选择性切断硫交联键，恢复部分塑性与加工性能。',13.5,False,INK)],'line':1.3},
])
# devulcanization methods as labelled rows
methods=[
    ("物理法","机械、热、超声、微波等输入能量打断交联","6E7B72"),
    ("化学法","加入再生剂/脱硫剂促进交联键断裂","BE6B3A"),
    ("生物法","微生物或酶温和脱硫，绿色但较慢","2C5F2D"),
]
my=3.55
for i,(m,d,col) in enumerate(methods):
    y=my+i*0.72
    rect(s, 0.7, y, 0.16, 0.56, fill=col)
    text(s, 1.0, y, 5.6, 0.56, [{'runs':[(m+'：',12.5,True,INK),(d,11.5,False,STEEL)],'line':1.12}],
         anchor=MSO_ANCHOR.MIDDLE)
rect(s, 0.7, 5.95, 6.0, 0.72, fill=MOSSBG)
text(s, 0.95, 5.95, 5.6, 0.72,
     [{'runs':[('再生胶可部分替代生胶，',12,True,FOREST),('降低成本、减少新胶消耗。',12,False,INK)],'line':1.18}],
     anchor=MSO_ANCHOR.MIDDLE)
# right: pros/cons two stacked panels
rect(s, 7.0, 2.0, 5.63, 2.25, fill=MOSSBG)
text(s, 7.3, 2.18, 5.1, 0.4, [{'runs':[('优势',14,True,GREEN)]}])
for i,t in enumerate(["替代部分天然/合成生胶，节约资源与成本",
                       "工艺相对成熟，已形成产业规模",
                       "可循环再加工，链条比较灵活"]):
    text(s, 7.5, 2.6+i*0.5, 4.9, 0.45, [{'runs':[('✓ ',12.5,True,GREEN),(t,11.5,False,INK)],'line':1.12}],
         anchor=MSO_ANCHOR.MIDDLE)
rect(s, 7.0, 4.42, 5.63, 2.25, fill=LIGHT)
text(s, 7.3, 4.6, 5.1, 0.4, [{'runs':[('局限与挑战',14,True,TERRA)]}])
for i,t in enumerate(["力学性能较原胶下降，掺用比例受限",
                       "传统工艺可能产生异味与废气",
                       "向“绿色脱硫、高性能再生”升级中"]):
    text(s, 7.5, 5.02+i*0.5, 4.9, 0.45, [{'runs':[('! ',12.5,True,TERRA),(t,11.5,False,INK)],'line':1.12}],
         anchor=MSO_ANCHOR.MIDDLE)
bottomstrip(s, "再生胶属“再生利用”层级，关键在提升性能与降低脱硫过程的二次污染。")
notes(s, "第三条路线是再生胶，要先理解橡胶为什么难回收：硫化让橡胶分子形成三维交联网络，"
          "既不熔化也不溶解。再生、也叫脱硫，目标就是有选择地切断硫交联键，让它恢复一部分塑性，能重新加工。"
          "方法分三类：物理法靠机械、热、超声、微波等输入能量；化学法加再生剂促进断键；生物法用微生物或酶，"
          "更绿色但速度慢。再生胶的好处是能部分替代生胶、节约成本，工艺也成熟；"
          "局限是力学性能比原胶差、掺用比例有限，传统工艺还可能有异味废气，所以现在都在往绿色脱硫和高性能再生方向升级。")

# =====================================================================
# Slide 8 — Route 4: pyrolysis (hero) — flow + product bar chart
# =====================================================================
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
kicker(s, "路径四 · PYROLYSIS")
title(s, "热解资源化：缺氧高温下“一胎四宝”全组分回收")
# left column: vertical process flow
fx=0.7; fw=4.0
steps=[("废旧轮胎/胶块","预处理：破碎、除杂",GREEN),
       ("缺氧/无氧高温热解","约 400–600 ℃ 热裂解",AMBER),
       ("冷凝 + 气固分离","分离液/固/气三相",FOREST)]
sy=2.05; sh=1.18; sg=0.28
for i,(a,b,col) in enumerate(steps):
    y=sy+i*(sh+sg)
    rect(s, fx, y, fw, sh, fill=LIGHT, line=col, line_w=1.5, shape=MSO_SHAPE.ROUNDED_RECTANGLE, shadow=True)
    rect(s, fx, y, 0.13, sh, fill=col)
    text(s, fx+0.32, y+0.12, fw-0.5, sh-0.2, [
        {'runs':[(a,13.5,True,INK)],'line':1.1,'space_after':3},
        {'runs':[(b,11,False,STEEL)],'line':1.1}], anchor=MSO_ANCHOR.MIDDLE)
    if i<len(steps)-1:
        text(s, fx, y+sh-0.05, fw, sg+0.1, [{'runs':[('▼',14,True,STEEL)],'align':PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)
# right: product yield bar chart
cd = CategoryChartData()
cd.categories = ['热解油', '炭黑(回收炭)', '钢丝', '可燃气']
cd.add_series('产率', (45, 33, 12, 10))
gframe = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(5.1), Inches(1.95),
                            Inches(7.5), Inches(3.55), cd)
chart = gframe.chart
chart.has_title = True
chart.chart_title.text_frame.text = "典型废轮胎热解产物产率（质量%，近似范围）"
tf0=chart.chart_title.text_frame.paragraphs[0].runs[0].font
tf0.size=Pt(12.5); tf0.bold=True; tf0.name=EA; tf0.color.rgb=RGBColor.from_string(INK)
chart.has_legend=False
cat_ax=chart.category_axis; val_ax=chart.value_axis
cat_ax.tick_labels.font.size=Pt(11); cat_ax.tick_labels.font.name=EA
cat_ax.tick_labels.font.color.rgb=RGBColor.from_string(INK)
val_ax.tick_labels.font.size=Pt(9); val_ax.tick_labels.font.name=LAT
val_ax.has_major_gridlines=True
val_ax.maximum_scale=60; val_ax.minimum_scale=0
ser=chart.plots[0].series[0]
chart.plots[0].has_data_labels=True
ser.data_labels.number_format='0"%"'; ser.data_labels.number_format_is_linked=False
ser.data_labels.font.size=Pt(11); ser.data_labels.font.bold=True; ser.data_labels.font.name=LAT
ser.data_labels.font.color.rgb=RGBColor.from_string(INK)
ser.data_labels.position=XL_LABEL_POSITION.OUTSIDE_END
bar_colors=[TERRA, INK, STEEL, AMBER]
for i,pt in enumerate(ser.points):
    pt.format.fill.solid(); pt.format.fill.fore_color.rgb=RGBColor.from_string(bar_colors[i])
# product value strip (bottom right)
rect(s, 5.1, 5.7, 7.53, 0.98, fill=MOSSBG)
text(s, 5.35, 5.78, 7.1, 0.95, [
    {'runs':[('产物去向：',12,True,FOREST)],'line':1.15,'space_after':2},
    {'runs':[('热解油',11.5,True,TERRA),('→燃料/化工原料  ',11,False,INK),
             ('回收炭黑',11.5,True,INK),('→补强填料/活性炭  ',11,False,INK),
             ('钢丝',11.5,True,STEEL),('→回炉炼钢  ',11,False,INK),
             ('可燃气',11.5,True,AMBER),('→供热自给',11,False,INK)],'line':1.25}])
bottomstrip(s, "产率为典型范围，随原料、温度与工艺差异较大；属“再生利用”中的化学回收路线。")
notes(s, "第四条路线是热解，也是近年最受关注的化学回收方式，我重点讲。"
          "做法是把破碎除杂后的胶块放进缺氧或无氧的反应器，加热到大约四百到六百摄氏度，"
          "让橡胶大分子热裂解，再经过冷凝和气固分离，得到液、固、气三相产物，可以说是'一胎四宝'："
          "热解油大约占四成多，可作燃料或化工原料；回收炭黑约三分之一，可做补强填料或进一步做活性炭；"
          "钢丝约一成多，直接回炉炼钢；还有约一成可燃气，通常回用给反应器供热，实现能量自给。"
          "它的优势是几乎全组分回收、理论上不留残渣；但产率受原料和工艺影响很大，"
          "而且产品品质和尾气控制是关键——这正好引出下一页的争议。")

# =====================================================================
# Slide 9 — Comparison table
# =====================================================================
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
kicker(s, "横向对比 · COMPARE")
title(s, "四条路线对比：没有“万能解”，按资源价值分级组合")
rows = [
    ["技术路线","4R 层级","主要产物/用途","相对价值","主要局限"],
    ["翻新 / 整胎利用","再利用","翻新胎、工程构件","★★★★","适用胎体有限、市场分散"],
    ["胶粉利用","再生利用","改性沥青、弹性场地","★★★","多为降级利用、需求受政策影响"],
    ["再生胶 / 脱硫","再生利用","再生胶替代生胶","★★★","性能下降、易二次污染"],
    ["热解","再生(化学)","油、炭黑、钢丝、气","★★★☆","品质波动、尾气与经济性"],
    ["能量回收(TDF)","能量回收","替代燃料热值","★★","仅回收热值、排放需严控"],
]
nrows=len(rows); ncols=5
tl, tt = 0.7, 2.0
tw, th = 11.93, 4.35
gtbl = s.shapes.add_table(nrows, ncols, Inches(tl), Inches(tt), Inches(tw), Inches(th)).table
# disable default banding style
gtbl.first_row=False; gtbl.horz_banding=False
colw=[2.55, 1.7, 3.1, 1.55, 3.03]
for j,wv in enumerate(colw):
    gtbl.columns[j].width=Inches(wv)
for i in range(nrows):
    gtbl.rows[i].height=Inches(th/nrows)
    for j in range(ncols):
        cell=gtbl.cell(i,j)
        cell.margin_left=Inches(0.12); cell.margin_right=Inches(0.08)
        cell.margin_top=Inches(0.04); cell.margin_bottom=Inches(0.04)
        cell.vertical_anchor=MSO_ANCHOR.MIDDLE
        if i==0:
            cell.fill.solid(); cell.fill.fore_color.rgb=RGBColor.from_string(FOREST)
        else:
            cell.fill.solid(); cell.fill.fore_color.rgb=RGBColor.from_string(LIGHT if i%2 else WHITE)
        tf=cell.text_frame; tf.word_wrap=True
        p=tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT if j in (0,2,4) else PP_ALIGN.CENTER
        r=p.add_run()
        is_head = (i==0)
        col = WHITE if is_head else INK
        sz = 12 if is_head else 11.5
        if j==3 and not is_head: col=AMBER; sz=13
        _set_run(r, rows[i][j], sz, is_head, col)
bottomstrip(s, "★为相对资源价值的定性示意，非定量评分；实际选择取决于胎源、规模与政策。")
notes(s, "把四条路线放在一起对比。表格按4R层级排列，右边用星号做一个定性的'相对价值'示意，注意这不是定量打分。"
          "可以看到：翻新和整胎利用价值最高，但能用的胎体有限、市场也比较分散；"
          "胶粉和再生胶规模大、技术成熟，但多是降级利用，还受政策和需求影响；"
          "热解能全组分回收、价值较高，难点在产品品质波动、尾气治理和经济性;"
          "能量回收只拿热值，价值最低，排放必须严格控制。"
          "结论是没有万能解，实际应当按胎源、规模和政策,把几条路线组合起来、分级利用。")

# =====================================================================
# Slide 10 — Controversy & bottlenecks
# =====================================================================
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
kicker(s, "争议与瓶颈 · CHALLENGES")
title(s, "现实瓶颈：回收“看起来美好”，落地仍有四道坎")
issues=[
    ("回收体系散、乱","个体回收为主、来源分散，正规企业“吃不饱”，“小作坊”难监管",FOREST),
    ("二次污染风险","土法炼油、落后热解尾气与废水处理不达标，造成新的污染",TERRA),
    ("低值化与同质化","大量产能集中在低端产品，高值化、差异化利用不足",AMBER),
    ("标准与经济性","产品标准、激励政策不完善，部分路线盈利依赖补贴",GREEN),
]
cw=(11.93-0.4)/2; ch=2.05; x0=0.7; y0=2.05; gx=0.4; gy=0.35
for i,(h,d,col) in enumerate(issues):
    x=x0+(i%2)*(cw+gx); y=y0+(i//2)*(ch+gy)
    rect(s, x, y, cw, ch, fill=LIGHT, shadow=True)
    circle(s, x+0.55, y+0.62, 0.62, col)
    text(s, x+0.25, y+0.32, 0.62, 0.62, [{'runs':[(f"{i+1:02d}",17,True,WHITE)],'align':PP_ALIGN.CENTER}],
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+1.05, y+0.22, cw-1.25, 0.5, [{'runs':[(h,15,True,INK)],'line':1.05}], anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+1.05, y+0.82, cw-1.3, ch-1.0, [{'runs':[(d,12,False,STEEL)],'line':1.25}], anchor=MSO_ANCHOR.TOP)
notes(s, "前面讲的技术听起来都很好，但落地还有四道坎。"
          "第一,回收体系散而乱:以个体回收为主、胎源分散,正规企业经常吃不饱,而小作坊又难监管。"
          "第二,二次污染:一些土法炼油和落后的热解装置,尾气废水处理不达标,反而制造了新的污染。"
          "第三,低值化和同质化:很多产能挤在低端产品,真正高值、差异化的利用还不够。"
          "第四,标准和经济性:产品标准和激励政策还不完善,有些路线的盈利甚至依赖补贴。"
          "这四点决定了行业还不能完全靠市场自发跑通,需要制度来托底,这就引出循环经济的思路。")

# =====================================================================
# Slide 11 — Circular economy & EPR (synthesis)
# =====================================================================
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
kicker(s, "综合观点 · SYNTHESIS")
title(s, "出路：以“分级利用 + 生产者责任”构建闭环")
# left: priority ladder (4 tiers)
text(s, 0.7, 1.95, 5.6, 0.45, [{'runs':[('① 分级利用：价值优先的处置顺序',14,True,FOREST)]}])
ladder=[("再利用(翻新/整胎)","最优先",GREEN),
        ("再生利用(胶粉/再生胶)","次之",MOSS),
        ("化学回收(热解)","补充",AMBER),
        ("能量回收 / 安全处置","兜底",STEEL)]
ly=2.5; lh=0.78; lg=0.16
widths=[5.6,4.9,4.2,3.5]
for i,(t,tag,col) in enumerate(ladder):
    w=widths[i]
    y=ly+i*(lh+lg)
    rect(s, 0.7, y, w, lh, fill=col)
    text(s, 0.95, y, w-1.4, lh, [{'runs':[(t,12.5,True,WHITE)]}], anchor=MSO_ANCHOR.MIDDLE)
    text(s, 0.7+w-1.25, y, 1.2, lh, [{'runs':[(tag,11,True,WHITE)],'align':PP_ALIGN.RIGHT}], anchor=MSO_ANCHOR.MIDDLE)
# right: EPR closed loop
text(s, 6.7, 1.95, 6.0, 0.45, [{'runs':[('② 生产者责任延伸（EPR）+ 循环闭环',14,True,FOREST)]}])
loop=["生产","销售/使用","回收","资源化","再入产业链"]
cx_, cy_, R = 9.6, 4.5, 1.3
import math
for i,node in enumerate(loop):
    ang = -math.pi/2 + i*(2*math.pi/len(loop))
    nx = cx_ + R*math.cos(ang)
    ny = cy_ + R*math.sin(ang)
    col = AMBER if i>=2 else GREEN
    circle(s, nx, ny, 1.02, col, line=WHITE, line_w=1.5)
    text(s, nx-0.5, ny-0.5, 1.0, 1.0, [{'runs':[(node,10.5,True,WHITE)],'align':PP_ALIGN.CENTER,'line':1.0}],
         anchor=MSO_ANCHOR.MIDDLE)
circle(s, cx_, cy_, 1.55, None, line=MOSS, line_w=2.0)
text(s, cx_-0.7, cy_-0.4, 1.4, 0.8, [{'runs':[('闭环',13,True,FOREST)],'align':PP_ALIGN.CENTER},
     {'runs':[('Circular',9,False,STEEL)],'align':PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)
rect(s, 0.7, 6.35, 11.93, 0.55, fill=FOREST)
text(s, 1.0, 6.35, 11.4, 0.55,
     [{'runs':[('一句话：',12,True,AMBER),('让“谁生产、谁负责回收”，叠加价值优先的分级利用，把废旧轮胎真正纳入循环经济。',12,True,WHITE)]}],
     anchor=MSO_ANCHOR.MIDDLE)
notes(s, "针对这些瓶颈,综合的出路可以概括成两条。"
          "第一是分级利用:严格按价值优先的顺序处置——能翻新、整胎利用就别破碎;"
          "其次是胶粉、再生胶;再次是热解这类化学回收;能量回收和安全处置只作兜底。"
          "第二是制度,核心是生产者责任延伸,也就是EPR:谁生产、谁就要对回收负责,"
          "用押金、回收目标、生态设计等手段,把回收的责任和成本内部化。"
          "两者结合,就能把右边这个'生产—使用—回收—资源化—再入产业链'的闭环真正转起来,"
          "这就是循环经济在废旧轮胎上的落地。")

# =====================================================================
# Slide 12 — Future directions
# =====================================================================
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
kicker(s, "未来方向 · OUTLOOK")
title(s, "发展趋势：向“绿色、高值、智能”升级")
futs=[
    ("绿色低碳工艺","连续化、清洁化热解；绿色脱硫；过程能量自给与尾气深度治理"),
    ("产品高值化","热解油提质、回收炭黑改性、制活性炭等高附加值产品"),
    ("智能回收体系","数字化溯源、智能分拣与区域化回收网络，提升正规回收率"),
    ("绿色设计与政策","可回收/易翻新的轮胎设计 + EPR、标准与激励协同发力"),
]
cw=(11.93-3*0.32)/4; ch=3.5; y0=2.1
icons=["♻","¥","◎","✦"]
cols=[GREEN, AMBER, FOREST, TERRA]
for i,(h,d) in enumerate(futs):
    x=0.7+i*(cw+0.32)
    rect(s, x, y0, cw, ch, fill=LIGHT, shadow=True)
    rect(s, x, y0, cw, 0.16, fill=cols[i])
    circle(s, x+cw/2, y0+0.95, 0.92, cols[i])
    text(s, x+cw/2-0.46, y0+0.49, 0.92, 0.92, [{'runs':[(icons[i],26,True,WHITE)],'align':PP_ALIGN.CENTER}],
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+0.12, y0+1.55, cw-0.24, 0.6, [{'runs':[(h,13.5,True,INK)],'align':PP_ALIGN.CENTER,'line':1.08}],
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+0.2, y0+2.2, cw-0.4, ch-2.3, [{'runs':[(d,11,False,STEEL)],'align':PP_ALIGN.CENTER,'line':1.22}],
         anchor=MSO_ANCHOR.TOP)
rect(s, 0.7, 5.95, 11.93, 0.72, fill=MOSSBG)
text(s, 1.0, 5.95, 11.4, 0.72,
     [{'runs':[('趋势主线：',12.5,True,FOREST),('从“处理废物”转向“经营资源”——技术、产品、体系与政策四轮驱动。',12.5,False,INK)],'line':1.18}],
     anchor=MSO_ANCHOR.MIDDLE)
notes(s, "面向未来,行业升级有四个方向。"
          "一是绿色低碳工艺:把热解做成连续化、清洁化,推广绿色脱硫,让过程能量自给、尾气深度治理;"
          "二是产品高值化,比如热解油提质、回收炭黑改性、做活性炭,提升附加值;"
          "三是智能回收体系,用数字化溯源、智能分拣和区域化网络,把正规回收率提上来;"
          "四是绿色设计加政策协同,从源头设计就考虑可回收、易翻新,再用EPR、标准和激励一起推动。"
          "一句话概括这条主线:从'处理废物'转向'经营资源',靠技术、产品、体系和政策四轮驱动。")

# =====================================================================
# Slide 13 — Summary / takeaways
# =====================================================================
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
kicker(s, "总结 · TAKEAWAYS")
title(s, "小结：把“黑色污染”变成“城市矿产”")
points=[
    ("是负担，更是资源","约九成质量可循环——橡胶、炭黑与钢丝都值得拿回来"),
    ("路线分级、各有取舍","翻新→胶粉/再生胶→热解→能量回收，价值递减、互为补充"),
    ("瓶颈在体系与品质","回收散乱、二次污染、低值化与标准不足是主要短板"),
    ("出路是循环经济","分级利用 + 生产者责任，推动绿色、高值、智能化升级"),
]
y0=2.05
for i,(h,d) in enumerate(points):
    y=y0+i*1.12
    circle(s, 1.0, y+0.42, 0.62, FOREST if i%2==0 else AMBER)
    text(s, 0.69, y+0.12, 0.62, 0.62, [{'runs':[(str(i+1),19,True,WHITE)],'align':PP_ALIGN.CENTER}],
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, 1.55, y+0.04, 4.2, 0.85, [{'runs':[(h,15,True,INK)],'line':1.12}], anchor=MSO_ANCHOR.MIDDLE)
    rect(s, 5.85, y+0.18, 0.04, 0.62, fill=LINEGR)
    text(s, 6.1, y+0.04, 6.5, 0.9, [{'runs':[(d,12.5,False,STEEL)],'line':1.22}], anchor=MSO_ANCHOR.MIDDLE)
bottomstrip(s, "作为环境工程的研究与实践方向，废旧轮胎的高值化资源化兼具环境与经济价值。")
notes(s, "最后做个小结,四句话。"
          "第一,废旧轮胎是负担、更是资源,约九成质量都能循环;"
          "第二,技术路线要分级、各有取舍,从翻新到胶粉、再生胶、热解,再到能量回收,价值递减、相互补充;"
          "第三,真正的瓶颈不在单项技术,而在回收体系、二次污染、低值化和标准;"
          "第四,出路是循环经济——靠分级利用加生产者责任,推动绿色、高值和智能化升级。"
          "对我们环境工程专业来说,这正是一个兼具环境效益和经济价值的实践方向。")

# =====================================================================
# Slide 14 — Thanks (closing, dark)
# =====================================================================
s = slide()
rect(s, 0, 0, SW, SH, fill=FOREST)
for d in [3.6, 2.6, 1.6]:
    circle(s, 11.7, 6.3, d, None, line=MOSS, line_w=1.3)
circle(s, 11.7, 6.3, 0.85, GREEN, line=MOSS, line_w=1.3)
text(s, 0.9, 2.7, 11.5, 1.4, [{'runs':[('谢谢聆听!',46,True,WHITE)]}])
text(s, 0.92, 4.1, 11.0, 0.6, [{'runs':[('欢迎批评指正与交流讨论',16,False,MOSSBG)]}])
circle(s, 1.02, 5.05, 0.12, AMBER)
text(s, 1.22, 4.87, 10.0, 0.4, [{'runs':[('废旧轮胎的回收与资源化利用 · 环境工程',12,True,MOSSBG)]}],
     anchor=MSO_ANCHOR.MIDDLE)
notes(s, "我的汇报到这里,谢谢大家,欢迎老师和同学批评指正。")

# ---------------- save ----------------
out = "output/final_presentation_cn.pptx"
prs.save(out)
print("Saved", out, "with", len(prs.slides._sldIdLst), "slides")
