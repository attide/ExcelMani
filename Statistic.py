#!/usr/bin/env python
# -*- coding: utf-8 -*-

u'统计Priority为0和1，统计创建时间和解决时间相差超过8小时的'

u'Priority为2和3，统计创建时间和解决时间相差超过24小时的'

u'输入表格：编号、标题、负责人、创建时间、解决时间、时间差（单位小时）'

__author__ = 'Fengyue Zhu'

import xlrd
import xlwt
import sys
import re
from datetime import date,datetime,time
from xlrd import open_workbook,xldate_as_tuple

PRIORITY = u'Priority'
START_TIME = u'创建时间'
SOLVED_TIME = u'解决时间'
BUG_ID = u'编号'
TITLE = u'标题'
OWNER = u'负责人'
TIME_DIFF = u'时间差（单位小时）'
RESOLUTION = u'Resolution'
NOT_A_BUG = u'Not a Bug'
P0_PRIORITY = u'P0-Highest'
P1_PRIORITY = u'P1-High'
P2_PRIORITY = u'P2-Middle'
P3_PRIORITY = u'P3-Lowest'

OUTPUT_XLS = 'sta_res.xls'

def getIdx(filename,key):
    bk = xlrd.open_workbook(filename)
    try:
        sh = bk.sheet_by_index(0)
    except:
        print 'no sheet in %s' % filename
        return
    first_row = sh.row_values(0)
    if key in first_row:
        target_index = first_row.index(key)
    else:
        print '\'', key ,'\' not in subtitle'
        return
    return target_index

def getContent(filename):
    bk = xlrd.open_workbook(filename)
    try:
        sh = bk.sheet_by_index(0)
    except:
        print 'no sheet in %s' % filename
        return
    nrows = sh.nrows
    row_list = []
    for i in range(1,nrows):
        row_data = sh.row_values(i)
        row_list.append(row_data)
    return row_list

def getDatetimeVal(datetime_num,filename):
    bk = xlrd.open_workbook(filename)
    try:
        sh = bk.sheet_by_index(0)
    except:
        print 'no sheet in %s' % filename
        return
    datetime_value = xldate_as_tuple(datetime_num,bk.datemode)
    return datetime(*datetime_value)

def resXlwt(result_list_p01,result_list_p23):
    book = xlwt.Workbook()
    if(result_list_p01 is None):
        pass
    else:
        sheet1 = book.add_sheet('Result P0,P1',cell_overwrite_ok=True)
        row_num = 0
        cell_num = 0
        for row in result_list_p01:
            for cell in row:
                sheet1.write(row_num,cell_num,cell)
                cell_num+=1
            row_num+=1
            cell_num=0
    if(result_list_p23 is None):
        pass
    else:
        sheet2 = book.add_sheet('Result P2,P3',cell_overwrite_ok=True)
        row_num = 0
        cell_num = 0
        for row in result_list_p23:
            for cell in row:
                sheet2.write(row_num,cell_num,cell)
                cell_num+=1
            row_num+=1
            cell_num=0
    book.save(OUTPUT_XLS)
    print 'Sheet Result has bee added'

def main(file_name):
    priority_idx = getIdx(file_name, PRIORITY)
    start_time_idx = getIdx(file_name, START_TIME)
    solved_time_idx = getIdx(file_name, SOLVED_TIME)
    bug_id_idx = getIdx(file_name, BUG_ID)
    title_idx = getIdx(file_name,TITLE)
    owner_idx = getIdx(file_name,OWNER)
    resolution_idx = getIdx(file_name,RESOLUTION)
    content = getContent(file_name)
    # 删除'Not a bug'的行
    bug_list = []
    for row_i in content:
        if row_i[resolution_idx] == NOT_A_BUG:
            pass
        else:
            bug_list.append(row_i)
    # print len(bug_list)
    subtitle = [BUG_ID,TITLE,OWNER,PRIORITY,START_TIME,SOLVED_TIME,TIME_DIFF]
    result_list_p01 = []
    result_list_p01.append(subtitle)
    result_list_p23 = []
    result_list_p23.append(subtitle)
    for row_i in bug_list:
        start_time_val = getDatetimeVal(row_i[start_time_idx],file_name)
        end_time_val = getDatetimeVal(row_i[solved_time_idx],file_name)
        time_diff_val = (end_time_val-start_time_val).total_seconds()/(60*60)
        if (row_i[priority_idx] in (P2_PRIORITY,P3_PRIORITY) and time_diff_val >= 24):
        # if True:
            res_row = [row_i[bug_id_idx],row_i[title_idx],row_i[owner_idx],row_i[priority_idx],
                str(start_time_val),
                str(end_time_val),
                int(time_diff_val)]
            result_list_p23.append(res_row)
        if (row_i[priority_idx] in (P0_PRIORITY,P1_PRIORITY) and time_diff_val >= 8):
            res_row = [row_i[bug_id_idx],row_i[title_idx],row_i[owner_idx],row_i[priority_idx],
                str(start_time_val),
                str(end_time_val),
                int(time_diff_val)]
            result_list_p01.append(res_row)
    resXlwt(result_list_p01,result_list_p23)

if __name__ == '__main__':
    args = sys.argv
    if len(args) == 2:
        file_name = args[1].decode('utf-8')
        main(file_name)
    else:
        print 'Invalid arguments'
