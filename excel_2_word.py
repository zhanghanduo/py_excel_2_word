#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
import xlrd
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
from docx.oxml.ns import qn

from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE

# Open the file
path_read = os.path.join(os.getcwd(), 'test.xlsx')
wb = xlrd.open_workbook(path_read)

# Get the list of the sheets name
sheet_list = wb.sheet_names()
# Select one sheet and get its size
s = wb.sheet_by_name(sheet_list[0])  # or s = wb.sheet_by_index(1)
print(s.nrows, s.ncols)

for i in range(s.nrows - 1):

    document = Document()
    document.styles['Normal'].font.name = u'宋体'
    document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')

    style_head = document.styles.add_style('r_head', WD_STYLE_TYPE.PARAGRAPH)
    style_head.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    style_head.font.name = u'宋体'
    style_head.font.size = Pt(9)
    style_head._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    foot_head = document.add_paragraph('JL(JLBZ-F-08)-02JD(01)',style = 'r_head')

    style = document.styles.add_style('rtl', WD_STYLE_TYPE.PARAGRAPH)
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style.font.name = u'宋体'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')

    head = document.add_paragraph('压力表检定记录',style = 'rtl')

    tab1 = document.add_table(rows=6, cols=4)
    tab1.style = 'Table Grid'
    cells0 = tab1.rows[0].cells
    cells0[0].text = '证书编号'
    cells0[2].text = '流    水    号'
    cells0[2].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cells0[3].text = str(s.cell(i+1, 1).value)
    cells0[0].width = Inches(0.95)
    cells0[1].width = Inches(2.13)
    cells0[2].width = Inches(1.13)
    cells0[3].width = Inches(1.80)

    cells1 = tab1.rows[1].cells
    cells1[0].text = '送检单位'
    cells1[1].text = str(s.cell(i + 1, 0).value)
    cells1[2].text = '计量器具名称'
    cells1[3].text = '压力表'
    cells1[3].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cells1[0].width = Inches(0.95)
    cells1[1].width = Inches(2.13)
    cells1[2].width = Inches(1.13)
    cells1[3].width = Inches(1.80)

    cells2 = tab1.rows[2].cells
    cells2[0].text = '制造单位'
    cells2[1].text = str(s.cell(i + 1, 2).value)
    cells2[2].text = '检 定 地 点'
    cells2[3].text = '本院热工所(压力)'
    cells2[0].width = Inches(0.95)
    cells2[1].width = Inches(2.13)
    cells2[2].width = Inches(1.13)
    cells2[3].width = Inches(1.80)

    cells3 = tab1.rows[3].cells
    cells3[0].text = '型号/规格'
    cells3[1].text = str(s.cell(i + 1, 3).value)
    cells3[2].text = '检 定 依 据'
    cells3[3].text = 'JJG52-2013'
    cells3[0].width = Inches(0.95)
    cells3[1].width = Inches(2.13)
    cells3[2].width = Inches(1.13)
    cells3[3].width = Inches(1.80)

    cells4 = tab1.rows[4].cells
    cells4[0].text = '测量范围'
    cells4[1].text = str(s.cell(i + 1, 3).value)
    cells4[2].text = '最大允许误差'
    cells4[3].text = '± 0.016 MPa'
    cells4[0].width = Inches(0.95)
    cells4[1].width = Inches(2.13)
    cells4[2].width = Inches(1.13)
    cells4[3].width = Inches(1.80)

    cells5 = tab1.rows[5].cells
    cells5[0].text = '准确度等级'
    cells5[1].text = str(s.cell(i + 1, 5).value) + '   级'
    cells5[2].text = '分   度   值'
    cells5[3].text = '0.05 MPa'
    cells5[0].width = Inches(0.95)
    cells5[1].width = Inches(2.13)
    cells5[2].width = Inches(1.13)
    cells5[3].width = Inches(1.80)

    tab2 = document.add_table(rows=1, cols=6)
    tab2.style = 'Table Grid'
    cell20 = tab2.rows[0].cells
    cell20[0].text = '出厂编号'
    no_ = str(s.cell(i + 1, 4).value)
    cell20[1].text = no_.rstrip('0').rstrip('.') if '.0' in no_ else no_ # incase change the number by program
    cell20[2].text = '环境温度'
    cell20[3].text = str(s.cell(i + 1, 6).value) + '  ℃'
    cell20[4].text = '相对湿度'
    cell20[5].text = str(s.cell(i + 1, 7).value) + '  %'
    cell20[0].width = Inches(0.95)
    cell20[1].width = Inches(1.8)
    cell20[2].width = Inches(0.95)
    cell20[3].width = Inches(0.59)

    tab3 = document.add_table(rows=1, cols=1)
    tab3.style = 'Table Grid'
    cell30 = tab3.rows[0].cells
    cell30[0].text = '检定所使用的计量标准'

    tab4 = document.add_table(rows=2, cols=3)
    tab4.style = 'Table Grid'
    cell40 = tab4.rows[0].cells
    cell40[0].text = '名称'
    cell40[1].text = '准确度等级'
    cell40[2].text = '证书编号'
    cell41 = tab4.rows[1].cells
    cell41[0].text = '0.05级活塞式压力计标准装置'
    cell41[1].text = '0.05'
    cell41[2].text = '[1987]辽量标法证字第1038号'

    tab5 = document.add_table(rows=3, cols=1)
    tab5.style = 'Table Grid'
    cell50 = tab5.rows[0].cells
    cell50[0].text = '(1)开机检查是否符合要求　　　　　　是　　　'



    path_write = os.path.join(os.getcwd(), 'result', str(i + 1) + '.docx')
    document.save(path_write)