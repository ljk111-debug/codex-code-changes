from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE

OUT = r"G:\\codex test\\周文豪_软件测试工程师_优化版.docx"

doc = Document()
sec = doc.sections[0]
sec.top_margin = Cm(1.25)
sec.bottom_margin = Cm(1.15)
sec.left_margin = Cm(1.45)
sec.right_margin = Cm(1.45)
sec.header_distance = Cm(0.5)
sec.footer_distance = Cm(0.5)

BLACK = RGBColor(0, 0, 0)
NAVY = RGBColor(24, 55, 86)
GRAY = RGBColor(90, 98, 108)
LIGHT = "E9EEF3"

def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tcPr.append(shd)
    shd.set(qn('w:fill'), fill)

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = tcPr.first_child_found_in('w:tcBorders')
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tcPr.append(borders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        if edge in kwargs:
            tag = 'w:{}'.format(edge)
            element = borders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                borders.append(element)
            for key in ['val', 'sz', 'space', 'color']:
                if key in kwargs[edge]:
                    element.set(qn('w:{}'.format(key)), str(kwargs[edge][key]))

def set_run_font(run, name='Microsoft YaHei', size=None, bold=None, color=BLACK):
    run.font.name = name
    run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), name)
    run._element.get_or_add_rPr().rFonts.set(qn('w:ascii'), name)
    run._element.get_or_add_rPr().rFonts.set(qn('w:hAnsi'), name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    run.font.color.rgb = color

def set_para(p, before=0, after=0, line=1.0, align=None, keep=False):
    fmt = p.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing = line
    if align is not None:
        p.alignment = align
    if keep:
        pPr = p._p.get_or_add_pPr()
        k = OxmlElement('w:keepNext')
        pPr.append(k)

styles = doc.styles
normal = styles['Normal']
normal.font.name = 'Microsoft YaHei'
normal._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
normal.font.size = Pt(9.2)
normal.font.color.rgb = BLACK

for style_name in ['Title', 'Heading 1', 'Heading 2']:
    st = styles[style_name]
    st.font.name = 'Microsoft YaHei'
    st._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
    st.font.color.rgb = BLACK

def add_text(p, text, size=9.2, bold=False, color=BLACK):
    r = p.add_run(text)
    set_run_font(r, size=size, bold=bold, color=color)
    return r

def add_section(title):
    p = doc.add_paragraph()
    set_para(p, before=8, after=3, line=1.0, keep=True)
    r = p.add_run(title.upper())
    set_run_font(r, size=10.6, bold=True, color=BLACK)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '5')
    bottom.set(qn('w:space'), '5')
    bottom.set(qn('w:color'), 'D6DEE7')
    pbdr.append(bottom)
    pPr.append(pbdr)
    return p

def add_bullet(label, text, size=8.9, after=1.5):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(0.38)
    p.paragraph_format.first_line_indent = Cm(-0.18)
    set_para(p, after=after, line=1.08)
    add_text(p, label, size=size, bold=True, color=NAVY)
    add_text(p, text, size=size)
    return p

# Header block
p = doc.add_paragraph()
set_para(p, after=1, line=1.0)
add_text(p, '周文豪', size=23, bold=True, color=BLACK)
add_text(p, '   软件测试工程师 | 自动化测试', size=12, bold=True, color=NAVY)
p = doc.add_paragraph()
set_para(p, after=5, line=1.0)
add_text(p, '19330110072  |  llljingkang00@gmail.com  |  湖南涉外经济学院 · 软件工程本科在读', size=9.2, color=GRAY)

add_section('个人概述')
p = doc.add_paragraph()
set_para(p, after=2.5, line=1.15)
add_text(p, '具备软件测试理论、接口测试、UI/接口自动化和日志分析能力，', size=9.2)
add_text(p, '有耳机固件及 App 测试实习经历。', size=9.2, bold=True, color=NAVY)
add_text(p, '能够从需求分析、用例设计、缺陷跟踪到回归验证参与版本质量保障；基于 Python + Pytest 完成 OTA 自动化验证，使单轮回归由约 2 小时缩短至 20 分钟。', size=9.2)

add_section('核心能力')
skills = [
    ('测试方法', '等价类、边界值、场景法、判定表、正交实验法；熟悉冒烟、回归、稳定性和兼容性测试。'),
    ('自动化测试', 'Python + Pytest + Playwright/Requests；能够编写 UI、接口及 OTA 自动化用例，使用 Allure 生成测试报告。'),
    ('接口与数据', 'Postman RESTful 接口测试；MySQL 增删改查、多表关联、业务数据一致性校验。'),
    ('工程化工具', 'Jira 缺陷管理、Git 版本控制、Jenkins 持续集成；掌握 Linux 命令、串口/ADB 日志抓取与分析。'),
    ('语言与学习', '可流畅阅读英文技术文档与需求文档；善用 AI 辅助日志聚类、关键字段提取和用例设计。'),
]
for label, text in skills:
    add_bullet(label + '：', text, size=8.8, after=1.2)

add_section('实习经历')
p = doc.add_paragraph()
set_para(p, after=1.5, line=1.0, keep=True)
add_text(p, 'Anker 创新科技（Soundcore 声阔）', size=10.5, bold=True)
add_text(p, '  |  软件测试实习生', size=9.5, bold=True, color=NAVY)
add_text(p, '                                      2026.06 - 2026.09', size=8.8, color=GRAY)
p = doc.add_paragraph()
set_para(p, after=3, line=1.08)
add_text(p, '负责耳机固件及 App 功能测试、OTA 自动化验证与日志分析，参与版本质量保障全流程。', size=8.9, color=GRAY)
add_bullet('用例设计与执行：', '独立完成测试分析与用例设计，累计编写并执行 300+ 条用例，覆盖音频播放、蓝牙连接与多设备切换、降噪/通透模式、触控、EQ、充电低电及 OTA 升级等核心场景。')
add_bullet('缺陷管理：', '通过 Jira 提交并跟踪 40+ 个缺陷，其中严重及以上 8 个；规范记录复现步骤、日志和预期结果，协同开发、固件及产品团队推动修复，缺陷验证闭环率 100%。')
add_bullet('OTA 自动化：', '基于 Python + Pytest 编写 OTA 升级自动化脚本，覆盖升级包校验、断点续传和失败回滚等 20+ 条用例；单轮回归由约 2 小时降至 20 分钟，效率提升约 80%。')
add_bullet('稳定性与兼容性：', '执行长时播放、反复配对及 Android/iOS 跨机型测试 100+ 轮，推动 5 类高频问题在量产前闭环。')
add_bullet('日志定位：', '使用串口/ADB 抓取耳机与 App 日志，结合 AI 工具完成日志聚类和关键字段提取，定位蓝牙断连、音频卡顿等问题；平均定位时间缩短约 50%，输出分析报告 10+ 份。')

add_section('项目经历')
p = doc.add_paragraph()
set_para(p, after=1.5, line=1.0, keep=True)
add_text(p, '电商商城系统测试', size=10.5, bold=True)
add_text(p, '  |  测试工程师', size=9.5, bold=True, color=NAVY)
add_text(p, '                                      2025.09 - 2025.12', size=8.8, color=GRAY)
p = doc.add_paragraph()
set_para(p, after=3, line=1.08)
add_text(p, 'B2C 电商平台，覆盖注册登录、商品管理、购物车、订单、支付结算和个人中心；技术栈：SpringBoot + Vue + MySQL。', size=8.9, color=GRAY)
add_bullet('测试设计：', '独立完成登录、商品、购物车和订单模块的测试分析，输出测试用例 120+ 条。')
add_bullet('接口测试：', '使用 Postman 验证核心接口参数、状态码及异常场景，通过参数化完成 60+ 条用例回归。')
add_bullet('数据校验：', '使用 MySQL 校验库存、订单状态和用户数据一致性，推动修复 6 处业务逻辑缺陷。')
add_bullet('UI 自动化：', '使用 Pytest + Playwright 编写下单、支付主流程用例 30+ 条，将冒烟测试由 40 分钟缩短至 8 分钟。')

add_section('教育背景')
p = doc.add_paragraph()
set_para(p, after=1.5, line=1.0)
add_text(p, '湖南涉外经济学院', size=9.8, bold=True)
add_text(p, '  |  软件工程（本科）', size=9.2, color=NAVY)
add_text(p, '                                      2023.09 - 2027.06', size=8.8, color=GRAY)
p = doc.add_paragraph()
set_para(p, after=1, line=1.08)
add_text(p, '主修课程：', size=8.8, bold=True)
add_text(p, '软件测试、Java 程序设计、数据结构与算法、数据库原理、计算机网络、操作系统', size=8.8)

# Footer page number field
footer = sec.footer
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_para(fp, before=2, line=1.0)
run = fp.add_run('周文豪 · 软件测试工程师')
set_run_font(run, size=7.5, color=GRAY)

doc.save(OUT)
print(OUT)

