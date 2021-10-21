import os
import sys

import xlrd
from xlutils.copy import copy
import datetime


def main(rootdir='第四章作业', homework_number=4):
    # 文件位置
    # rootdir = '第四章作业'
    # # 第几次作业
    # homework_number = 4

    execl_file = xlrd.open_workbook(r'作业提交情况点名册.xls')
    sheet = execl_file.sheet_by_index(0)
    nrows = sheet.nrows
    ncols = sheet.ncols
    excel = copy(execl_file)
    table = excel.get_sheet(0)
    
    # 未查找到的提交文件
    findless_list = []
    # 查找错误的提交文件
    error_list = []

    name_list = []
    name_str = ""
    for i in range(8, nrows):
        name_list.append(sheet.cell(i, 2).value)
        name_str += sheet.cell(i, 2).value

    list = os.listdir(rootdir)
    for i in range(0, len(list)):
        file_name = list[i].split(".")[0]
        flag = -1
        name = ""
        for ch in file_name:
            if '\u4e00' <= ch <= '\u9fff':
                if (name_str.find(ch) != -1) and flag == -1:
                    flag = 1
                    name += ch
                elif name_str.find(ch) != -1 and flag != -1:
                    if name_str.find(name + ch) != -1:
                        name += ch
                    elif len(name) <= 1:
                        flag = 1
                        name = ch
                        # flag = name_str.find(ch)
        if name != "":
            print(name)
            try:
                loc = name_list.index(name)
            except ValueError:
                error_list.append(list[i])
            else:
                table.write(loc + 8, homework_number + 5, 10)
        else:
            findless_list.append(list[i])

    today = datetime.date.today()
    formatted_today = today.strftime('%y%m%d')
    excel.save('作业提交情况点名册.xls')
    print('未找到的文件：', findless_list)
    print('发生错误的文件：', error_list)


if __name__ == '__main__':
    main(sys.argv[1], int(sys.argv[2]))
    print("Finish")
