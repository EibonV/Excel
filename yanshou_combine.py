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
    nrow = sheet.nrows
    return nrow

#获取sheet下有多少列数据
def get_sheetcol_num(sheet):
    ncol = sheet.ncols
    return ncol

#获取sheet下所需列的数据
def get_sheet_data(sheet,row,col):
    all_data0 = []
    all_data0 = sheet.col_values(col)
    for i in range(row):
        del(all_data0[0])
    return all_data0

#获取“项目名称”行列
def get_xiangmu(sheet,row,col):
    for i in range(row):
        for j in range(col):
            xiangmu_value = sheet.cell(i,j).value
            if xiangmu_value == "项目名称":
                xiangmu = [i,j]
                return xiangmu
                break

#获取“工程编号”行列
def get_bianhao(sheet,row,col):
    for i in range(row):
        for j in range(col):
            bianhao_value = sheet.cell(i,j).value
            if bianhao_value == "工程编号":
                bianhao = [i,j]
                return bianhao
                break

#获取“验收日期”行列
def get_date(sheet,row,col):
    for i in range(row):
        for j in range(col):
            date_value = sheet.cell(i,j).value
            if date_value == "验收日期":
                date = [i,j]
                return date
                break

if __name__=='__main__':
    all_exce = get_exce()
    #得到要合并的所有exce表格数据
    if(all_exce == 0):
        print("该目录下无.xlsx文件！请检查您输入的目录是否有误！")
        os.system('pause')
        exit()
    all_data1 = [] #用于保存项目名称列的数据
    all_data2 = [] #用于保存工程编号列的数据
    all_data3 = [] #用于保存所有验收日期列的数据
    
    #下面开始文件数据获取
    for exce in all_exce:
        fh = open_exce(exce) #打开文件
        sheets = get_sheet(fh) #获取所有sheet

    #获取项目名称数据
        for sheet in range(len(sheets)):
            sheetrow = get_sheetrow_num(sheets[sheet]) #获取列表总行数
            sheetcol = get_sheetcol_num(sheets[sheet]) #获取列表总列数
            xiangmu = get_xiangmu(sheets[sheet],sheetrow,sheetcol) #获取项目名称所在行列
            all_data1 = get_sheet_data(sheets[sheet],xiangmu[0],xiangmu[1]) #获取项目名称所有数据

    #获取工程编号数据
        for sheet in range(len(sheets)):
            sheetrow = get_sheetrow_num(sheets[sheet]) #获取列表总行数
            sheetcol = get_sheetcol_num(sheets[sheet]) #获取列表总列数
            bianhao = get_bianhao(sheets[sheet],sheetrow,sheetcol)
            all_data2 = get_sheet_data(sheets[sheet],bianhao[0],bianhao[1])

    #获取验收时间数据
        for sheet in range(len(sheets)):
            sheetrow = get_sheetrow_num(sheets[sheet]) #获取列表总行数
            sheetcol = get_sheetcol_num(sheets[sheet]) #获取列表总列数
            date = get_date(sheets[sheet],sheetrow,sheetcol)
            all_data3 = get_sheet_data(sheets[sheet],date[0],date[1])

    #写入新表
    #新建表格
    new_exce = wei_zhi + "test.xlsx"
    fh1 = xlsxwriter.Workbook(new_exce)
    new_sheet = fh1.add_worksheet()

    #写入第一列
    for i in range(len(all_data1)):
        c = all_data1[i]
        new_sheet.write(i,0,c)

    #写入第二列
    for i in range(len(all_data2)):
        c = all_data2[i]
        new_sheet.write(i,1,c)

    #写入第三列
    for i in range(len(all_data3)):
        c = all_data3[i]
        new_sheet.write(i,2,c)

    #关闭exce表    
    fh1.close
    print("文件合并成功,请查看“" + wei_zhi + "”目录下的test.xlsx文件！")

    os.system('pause')
    os.system('pause')

        



    