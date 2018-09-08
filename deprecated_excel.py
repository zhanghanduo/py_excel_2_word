#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
import sys
import importlib
import logging
import time
import datetime

from xlutils.copy import copy
import xlrd
import xlwt
from xlwt import easyxf

def convert(config):

    # 1. 读取数据源excel的数据，读入内存
    path_source = os.path.join(os.getcwd(), config['read_params']['source_name'])
    doc_source = xlrd.open_workbook(path_source)

    # 1.1 获取数据源所有工作簿(sheet)
    sheet_list_s = doc_source.sheet_names()
    # 1.2 默认第一个工作簿(sheet_list_s[0])，输出表格行列数
    s = doc_source.sheet_by_name(sheet_list_s[0])  # or s = doc_source.sheet_by_index(1)
    print("row and column number of data source: {} | {}".format(s.nrows, s.ncols))

    # 2. 读取模板excel的数据，读入内存
    path_template = os.path.join(os.getcwd(), config['read_params']['template_name'])
    doc_template = xlrd.open_workbook(path_template, formatting_info=True)

    # 2.1 获取模板excel所有工作簿(sheet)
    sheet_list_t = doc_template.sheet_names()
    print("names of sheets: {}".format(sheet_list_t))
    print("num of sheets: {}".format(len(sheet_list_t)))

    # 2.2 尝试读取第一个工作簿(sheet_list[0])，输出表格行列数
    temp0 = doc_template.sheet_by_name(sheet_list_t[0])  # or s = doc_source.sheet_by_index(1)
    print("row and column number of template 0: {} | {}".format(temp0.nrows, temp0.ncols))

    # 3. 建立子目录，用于生成excel文档
    sub_working_dir = '{}/{}/{}'.format(
        os.getcwd(), config['output_dir'],
        time.strftime("%d%H%M%S", time.localtime()))
    if not os.path.exists(sub_working_dir):
        os.makedirs(sub_working_dir)
    logging.info("sub working dir: %s" % sub_working_dir)

    # 4. 源数据每一行都进行处理
    for i in range(s.nrows - 2):
        row_index = i + 2
        workbook = copy(doc_template)                              # 建立模板副本用于修改添加信息

        # # 4.1 读取源数据中的测量范围，提取数字(上限，下限)
        # up_lim = s.cell(row_index, 6).value
        # down_lim = s.cell(row_index, 5).value

        # # 4.2 根据上限判断使用哪个模板，建立工作簿副本w_sheet
        # if up_lim < 1:
        #     w_sheet = workbook.get_sheet(0)
        # elif up_lim < 10:
        #     w_sheet = workbook.get_sheet(1)
        # elif up_lim < 50:
        #     w_sheet = workbook.get_sheet(2)
        # else:
        #     print("error!")

        # # 5.1 送检单位
        # w_sheet.write(3, 18, str(s.cell(row_index, 0).value))

        # # 5.2 流水号
        # w_sheet.write(4, 18, str(s.cell(row_index, 1).value))

        # path_write = os.path.join(sub_working_dir, config['output_dir'] + str(i + 1) + '.xls')

        workbook.save(path_write)


def main():
    logging.basicConfig(level=logging.DEBUG,
                        format="[%(asctime)s %(filename)s] %(message)s")

    if len(sys.argv) != 2:
        logging.error("Usage: python training.py params.py")
        sys.exit()
    params_path = sys.argv[1]
    if not os.path.isfile(params_path):
        logging.error("no params file found! path: {}".format(params_path))
        sys.exit()
    config = importlib.import_module(params_path[:-3]).PARAMS
    
    convert(config)

if __name__ == "__main__":
    main()