import xlrd
import xlsxwriter
import glob2 as glob
import os

biao_tou = "NULL"
wei_zhi = "NULL"

#获取要合并的所有exce表格
def get_exce():
    global wei_zhi
    wei_zhi = input("请输入Excel文件所在的目录：")
    all_exce = glob.glob(wei_zhi + "*.xlsx")
    print("该目录下有" + str(len(all_exce)) + "个exce文件：")
    if(len(all_exce) == 0):
        return 0
    else:
         for i in range(len(all_exce)):
             print(all_exce[i])
         return all_exce					
        


#打开Exce文件
def open_exce(name):
    fh = xlrd.open_workbook(name)
    return fh

#获取exce文件下的所有sheet
def get_sheet(fh):
    sheets = fh.sheets()
    return sheets


#获取sheet下有多少行数据
def get_sheetrow_num(sheet):
    return sheet.nrows
    


#获取sheet下的数据
def get_sheet_data(sheet,row):
    for i in range(row):
        if (i == 0):
            global biao_tou
            biao_tou = sheet.row_values(i)
            continue
        values = sheet.row_values(i)
        all_data1.append(values)
        
    return all_data1
    

if __name__=='__main__':
    all_exce = get_exce()
    #得到要合并的所有exce表格数据
    if(all_exce == 0):
        print("该目录下无.xlsx文件！请检查您输入的目录是否有误！")
        os.system('pause')
        exit()

    all_data1 = []
    #用于保存合并的所有行的数据


    #下面开始文件数据的获取
    for exce in all_exce:
        fh = open_exce(exce)
        #打开文件
        sheets = get_sheet(fh)
        #获取文件下的sheet数量


        for sheet in range(len(sheets)):
            row = get_sheetrow_num(sheets[sheet])
            #获取一个sheet下的所有的数据的行数

            all_data2 = get_sheet_data(sheets[sheet],row)
            #获取一个sheet下的所有行的数据

    all_data2.insert(0,biao_tou)
    #表头写入

    


    #下面开始文件数据的写入
    new_exce = wei_zhi + "test.xlsx"
    #新建的exce文件名字

    
    fh1 = xlsxwriter.Workbook(new_exce)
    #新建一个exce表

    new_sheet = fh1.add_worksheet()
    #新建一个sheet表

    for i in range(len(all_data2)):
        for j in range(len(all_data2[i])):
            c = all_data2[i][j]
            new_sheet.write(i,j,c)
            
    fh1.close()
    #关闭该exce表
    
    print("文件合并成功,请查看“" + wei_zhi + "”目录下的test.xlsx文件！")
            
    os.system('pause')
    os.system('pause')