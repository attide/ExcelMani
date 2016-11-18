#!/usr/bin/env python
# -*- coding: utf-8 -*-

u'生成线下质量分析统计的相关表格'

__author__ = 'Fengyue Zhu'

import xlrd
import xlwt
import sys
import re
from datetime import date,datetime,time
from xlrd import open_workbook,xldate_as_tuple

STAGE = u'可前置在哪个阶段发现'
SEVERITY = u'Severity'
# PRIORITY = u'Priority'
RESOLUTION = u'Resolution'
OWNER = u'负责人'
START_TIME = u'创建时间'
SOLVED_TIME = u'解决时间'
TITLE = u'标题'
BUG_TYPE = u'Bug类型'

DISTRIBUTION = u'每日分布'
MODULE_DIS = u'模块分布'

NOT_A_BUG = u'Not a Bug'

OUTPUT_XLS = 'v3_res.xls'

def getIdx(first_row,key):
    if key in first_row:
        target_index = first_row.index(key)
    else:
        print '\'', key ,'\' not in subtitle'
        return
    return target_index

def getFirstName(name_data):
    name_list = re.split(r'\,+', name_data)
    first_name = re.split(r'\(', name_list[0])[0]
    return first_name

def getModule(title_data):
    if re.match(u'【(.*?)】', title_data):
        module_key = re.match(u'【(.*?)】', title_data).groups()[0]
        if u"直播" in module_key:
            module_key = u'直播'
        if u"组件" in module_key:
            module_key = u'组件'
    else:
        module_key = title_data
    if any(crash_key in module_key for crash_key in (u'崩溃',u'Crash',u'crash',u'闪退')):
        module_key = u'崩溃'
    return module_key

def getDateStr(date_val,bk):
    date_value = xldate_as_tuple(date_val,bk.datemode)
    date_key = date(*date_value[:3])
    return str(date_key)

def countCol(target_col):
    if target_col == []:
        return
    result = {}
    for vals in target_col:
        if vals not in result:
            result[vals] = 0
        result[vals] += 1
    return result

def mergeDate(start_list,solved_list):
    if start_list == [] or solved_list == []:
        return
    result = {}
    for vals in start_list:
        if vals not in result:
            result[vals] = [0,0]
        result[vals][0] += 1
    for vals in solved_list:
        if vals not in result:
            result[vals] = [0,0]
        result[vals][1] += 1
    return result

def comXlwt(book, dic, key):
    if dic is None:
        return
    sheet1 = book.add_sheet(key,cell_overwrite_ok=True)
    sheet1.write(0,0,key)
    sheet1.write(0,1,u'数量')
    row_num = 1
    val_sum = 0
    for i in dic:
        if i == '':
            sheet1.write(row_num,0,u'未知')
        else:
            sheet1.write(row_num,0,i)
        sheet1.write(row_num,1,dic[i])
        row_num+=1
        val_sum+=dic[i]
    sheet1.write(row_num,0,u'合计')
    sheet1.write(row_num,1,val_sum)
    print 'Sheet \'', key,' \' has bee added'

def dateXlwt(book, dic, key):
    if(dic is None):
        return
    sheet1 = book.add_sheet(key,cell_overwrite_ok=True)
    sheet1.write(0,0,key)
    sheet1.write(0,1,u'数量')
    row_num = 1
    val_sum1 = 0
    val_sum2 = 0
    for i in dic:
        if i == '':
            sheet1.write(row_num,0,u'未知')
        else:
            sheet1.write(row_num,0,i)
        sheet1.write(row_num,1,dic[i][0])
        sheet1.write(row_num,2,dic[i][1])
        row_num+=1
        val_sum1+=dic[i][0]
        val_sum2+=dic[i][1]
    sheet1.write(row_num,0,u'合计')
    sheet1.write(row_num,1,val_sum1)
    sheet1.write(row_num,2,val_sum2)
    print 'Sheet \'', key,' \' has bee added'

def main(file_name):
    bk = xlrd.open_workbook(file_name)
    try:
        sh = bk.sheet_by_index(0)
    except:
        print 'no sheet in %s' % filename
        return
    first_row = sh.row_values(0)
    nrows = sh.nrows
    stage_idx = getIdx(first_row,STAGE)
    severity_idx = getIdx(first_row,SEVERITY)
    resolution_idx = getIdx(first_row,RESOLUTION)
    owner_idx = getIdx(first_row,OWNER)
    start_time_idx = getIdx(first_row, START_TIME)
    solved_time_idx = getIdx(first_row, SOLVED_TIME)
    title_idx = getIdx(first_row,TITLE)
    bug_type_idx = getIdx(first_row,BUG_TYPE)
    stage_col = []
    severity_col = []
    resolution_col = []
    owner_col = []
    start_time_col = []
    solved_time_col = []
    module_col = []
    bug_type_col = []
    for i in range(1,nrows):
        row_data = sh.row_values(i)
        if row_data[resolution_idx] == NOT_A_BUG:
            continue
        else:
            if stage_idx:
                stage_col.append(row_data[stage_idx])
            if severity_idx:
                severity_col.append(row_data[severity_idx])
            if resolution_idx:
                resolution_col.append(row_data[resolution_idx])
            if owner_idx:
                owner_col.append(getFirstName(row_data[owner_idx]))
            if title_idx:
                module_col.append(getModule(row_data[title_idx]))
            if bug_type_idx:
                bug_type_col.append(row_data[bug_type_idx])
            if start_time_idx and solved_time_idx:
                start_time_col.append(getDateStr(row_data[start_time_idx],bk))
                solved_time_col.append(getDateStr(row_data[solved_time_idx],bk))
    book = xlwt.Workbook()
    comXlwt(book,countCol(stage_col),STAGE)
    comXlwt(book,countCol(severity_col),SEVERITY)
    comXlwt(book,countCol(resolution_col),RESOLUTION)
    comXlwt(book,countCol(owner_col),OWNER)
    comXlwt(book,countCol(module_col),MODULE_DIS)
    comXlwt(book,countCol(bug_type_col),BUG_TYPE)
    dateXlwt(book,mergeDate(start_time_col,solved_time_col),DISTRIBUTION)
    book.save(OUTPUT_XLS)

if __name__ == '__main__':
    args = sys.argv
    if len(args) == 2:
        # print args[1]
        file_name = args[1].decode('utf-8')
        main(file_name)
    else:
        print 'Invalid arguments'

