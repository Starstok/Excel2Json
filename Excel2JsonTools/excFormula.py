#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import xlrd
import json

'''
pid\modelId\poiCode
生成Excel匹配公式
'''
def createFormula(sheet, dataNcols, baseData):
    body = ''
    for rown in range(1, sheet.nrows):
        name = conve(sheet.cell_value(rown, 0))
        value = conve(sheet.cell_value(rown, dataNcols))
        body += 'IF(${}="{}",{},'.format(baseData, name, value)
        # 最后一个去掉,
        if sheet.nrows - rown == 1:
            name = conve(sheet.cell_value(rown, 0))
            pid = conve(sheet.cell_value(rown, 1))
            body += 'IF(${}="{}",{}'.format(baseData, name, value)
    start = "="
    end = ")" * rown
    formula = start + body + end
    # print(formula)
    return formula

def conFormula(inPutFile, dataType, baseData):
    workbook = xlrd.open_workbook(inPutFile)    # 打开文件
    sheetCount = len(workbook.sheets())         #sheet数量
    sheet = workbook.sheet_by_index(0)          # 获取第一个sheet
    # sheet的名称，行数，列数
    print("表格名:", sheet.name, "行数:", sheet.nrows, "列数:", sheet.ncols)

    if sheet.ncols > 5: # 如果表格少于5列，返回1
        return 1
    
    if dataType == "pid":
        dataNcols = 1   # 数据所在的列数
        return createFormula(sheet, dataNcols, baseData)
    elif dataType == "modelId":
        dataNcols = 2
        return createFormula(sheet, dataNcols, baseData)
    elif dataType == "poiCode":
        dataNcols = 3
        return createFormula(sheet, dataNcols, baseData)

def conve(data):
    if type(data) == float:
        intData = int(data)
    else:
        intData = data
    return intData
