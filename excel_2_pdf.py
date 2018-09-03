#!/usr/bin/python
# -*- coding: utf-8 -*-

import xlrd
from PyPDF2 import PdfFileWriter

# Open the file
wb = xlrd.open_workbook('C:\\workspace\\test.xlsx')

# Get the list of the sheets name
sheet_list = wb.sheet_names()
# Select one sheet and get its size
s = wb.sheet_by_name(sheet_list[0])  # or s = wb.sheet_by_index(1)
print("excel size of sheet 1: ", s.nrows, s.ncols)





