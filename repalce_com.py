import xlrd
from xlrd import open_workbook
import os
import pandas as pd
from xlutils.copy import copy

str_num = {
        '力荐':'5',
        '推荐':'4',
        '还行':'3',
        '较差':'2',
        '很差':'1'
    }

def replace_evaluation():
    replace_path = "X:\pythonCode\getComment\comments\深夜食堂.xls"
    xlsx = xlrd.open_workbook(replace_path,formatting_info=True)
    shell = xlsx.sheet_by_index(0)
    rows = shell.nrows
    wb  = copy(xlsx)
    new_shell = wb.get_shell(0)
    #读取行数_
    rowsNum = new_shell.nrows
    print(rowsNum)
    for i in range(rowsNum):
        cell_2 = new_shell.cell_value(i, 2)
        num = str_num.get(cell_2)
        new_shell.write(i,2,num)
    pd.read_excel(wb).to_csv('X:\pythonCode\getComment\comments\syst.csv',encoding='utf-8')


if __name__=='__main__':
    replace_evaluation()
