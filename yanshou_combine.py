import xlrd
import xlsxwriter
import glob2 as glob
import os

wei_zhi = "NULL"

#获取要合并的表格
def get_exce():
    global wei_zhi
    wei_zhi = input("请输入Excel所在目录：")
    all_exce = glob.glob(wei_zhi + "*.xls") + glob.glob(wei_zhi + "*.xlsx")
    print("该目录下有" + str(len(all_exce)) + "个Excel文件：")
    if(len(all_exce) == 0):
        return 0
    else:
        for i in range(len(all_exce)):
            print(all_exce[i])
        return all_exce

#打开Excel文件
def open_exce(name):
    fh = xlrd.open_workbook(name)
    return fh

#获取Excel文件下所有sheet
def get_sheet(fh):
    sheets = fh.sheets()
    return sheets

#获取sheet下有多少行数据
def get_sheetrow_num(sheet):
    return sheet.nrows

#获取sheet下有多少列数据
def get_sheetcol_num(sheet):
    return sheet.ncols

#获取sheet下的数据
def get_sheet_data(sheet,row):
    for i in range(row):
        values = sheet.row_values(i)
        all_data1.append(values)
    return all_data1



