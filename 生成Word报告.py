# -*- coding: utf-8 -*-
"""
将 销售数据分析报告.md 转成 Word 文档，并在对应小节嵌入图表
"""
import os, re
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

BASE = r'D:\Python\Python_item\PythonProject\销售数据分析'
MD = os.path.join(BASE, '销售数据分析报告.md')
CHART_DIR = os.path.join(BASE, '图表')
OUT = os.path.join(BASE, '销售数据分析报告.docx')

# 图表文件名 -> 插入位置（在哪个小节标题后）
charts = [
    ('1_每日销售趋势.png',       '2.1'),
    ('2_分周汇总.png',           '2.2'),
    ('10_星期几成交规律.png',     '2.3'),
    ('12_门店成交与线索质量.png', '3.1'),
    ('3_各城市成交对比.png',     '3.1'),
    ('11_城市车型热力图.png',     '3.2'),
    ('5_各车型销量占比.png',     '4.1'),
    ('6_各车型转化率.png',     '4.2'),
    ('7_销售顾问业绩排行.png',     '5.1'),
    ('8_销售顾问产能转化散点.png', '5.2'),
    ('9_整体转化漏斗.png',       '六'),
]

# 中文映射：小节号 -> 对应图表
section_charts = {
    '### 2.1': ['1_每日销售趋势.png'],
    '### 2.2': ['2_分周汇总.png'],
    '### 2.3': ['10_星期几成交规律.png'],
    '### 3.1': ['3_各城市成交对比.png', '12_门店成交与线索质量.png'],
    '### 3.2': ['11_城市车型热力图.png'],
    '### 4.1': ['5_各车型销量占比.png'],
    '### 4.2': ['6_各车型转化率.png'],
    '### 5.1': ['7_销售顾问业绩排行.png'],
    '### 5.2': ['8_销售顾问产能转化散点.png'],
    '## 六':   ['9_整体转化漏斗.png'],
}

def set_cn_font(run, name='微软雅黑', size=None, bold=None, color=None):
    run.font.name = name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    if size: run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if color: run.font.color.rgb = color

def parse_md_table(lines):
    """解析markdown表格 -> list of rows(list of str)"""
    rows = []
    for ln in lines:
        ln = ln.strip()
        if not ln.startswith('|'):
            continue
        cells = [c.strip() for c in ln.strip('|').split('|')]
        if all(re.fullmatch(r':?-{2,}:?', c) for c in cells):  # 分隔行
            continue
        rows.append(cells)
    return rows

doc = Document()
# 全局默认中文字体
style = doc.styles['Normal']
style.font.name = '微软雅黑'
style.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
style.font.size = Pt(10.5)

with open(MD, encoding='utf-8') as f:
    lines = f.read().splitlines()

i = 0
n = len(lines)
while i < n:
    line = lines[i].rstrip()

    # 跳过开头的 fenced code block 标记 ``` 
    if line.strip().startswith('```'):
        i += 1
        continue

    # 表格：连续读入
    if line.startswith('|'):
        tbl_start = i
        while i < n and lines[i].strip().startswith('|'):
            i += 1
        rows = parse_md_table(lines[tbl_start:i])
        if rows:
            t = doc.add_table(rows=len(rows), cols=len(rows[0]))
            t.style = 'Light Grid Accent 1'
            t.alignment = WD_TABLE_ALIGNMENT.CENTER
            for r_idx, row in enumerate(rows):
                for c_idx, cell in enumerate(row):
                    cell_text = cell.replace('**','').replace('`','')
                    run = t.cell(r_idx, c_idx).paragraphs[0].add_run(cell_text)
                    set_cn_font(run, size=9, bold=(r_idx==0))
        continue

    # 标题
    clean_h = lambda s: s.replace('`','').replace('**','')
    if line.startswith('# '):
        doc.add_heading(clean_h(line[2:].strip()), level=0)
    elif line.startswith('## '):
        h = doc.add_heading(clean_h(line[3:].strip()), level=1)
    elif line.startswith('### '):
        h = doc.add_heading(clean_h(line[4:].strip()), level=2)
        # 判断该小节是否有配图
        for key, imgs in section_charts.items():
            if line.strip().startswith(key) or line.strip()==key:
                for img in imgs:
                    p = os.path.join(CHART_DIR, img)
                    if os.path.exists(p):
                        doc.add_picture(p, width=Inches(6.0))
                        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                break
    elif line.startswith('#### '):
        doc.add_heading(line[5:].strip(), level=3)
    # 分隔线
    elif re.fullmatch(r'-{3,}', line.strip()):
        pass
    # 引用块
    elif line.startswith('>'):
        p = doc.add_paragraph()
        run = p.add_run(line.lstrip('> '))
        set_cn_font(run, size=9.5, color=RGBColor(0x66,0x66,0x66))
    # 无序列表
    elif line.startswith('- '):
        p = doc.add_paragraph(style='List Bullet')
        run = p.add_run(line[2:].strip())
        set_cn_font(run)
    # 有序列表(数字.) 
    elif re.match(r'^\d+\.\s', line):
        p = doc.add_paragraph(style='List Number')
        run = p.add_run(re.sub(r'^\d+\.\s','',line))
        set_cn_font(run)
    # 代码块(line内)
    elif line.strip():
        p = doc.add_paragraph()
        run = p.add_run(line)
        set_cn_font(run)

    i += 1

doc.save(OUT)
print('Word 已生成:', OUT)
