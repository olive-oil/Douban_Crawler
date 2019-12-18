import os
import pandas as pd
import numpy as py
import xlrd

save_path = 'X:\pythonCode\getComment\comments'
def mergeToCsv(film_name,com_range):
    data = []
    for i in range(0, 25, 1):
        print(i)
        path = str(float(i))+'.'+film_name+com_range+'豆瓣影评.xlsx'
        print(path)
        xlsx = xlrd.open_workbook(path)
        sh0 = xlsx.sheet_by_index(0)
        rowsNum = sh0.nrows
        if i == 0:
            data.append(sh0.row_values(0))
        for j in range(1, rowsNum):
            data.append(sh0.row_values(j))
    content = pd.DataFrame(data)
    save_path_h = save_path + '\\'+com_range+'_'+film_name
    print(save_path_h)
    if os.path.exists(save_path_h):
        os.remove(save_path_h)
    content.to_excel(save_path_h+'.xlsx',header=False,index=False)

if __name__ == '__main__':
    film_name = input("电影名：")
    com_range = input("评论等级：")
    mergeToCsv(film_name,com_range)
