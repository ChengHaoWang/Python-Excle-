# -*- coding: utf-8 -*-
"""
Created on Sat Nov 24 17:36:11 2018

@author: Wang
"""

import xlrd
import xlwt
import xlwings as xw
from datetime import date,datetime

def read_excel():
    # 打开文件
    #workbook = xlrd.open_workbook(r'E:\第四题.xlsx')
    # 获取所有sheet


    # 根据sheet索引或者名称获取sheet内容
    wb = xw.Book(r'E:\第四题1.xlsx') # sheet索引从0开始
    sheet=wb.sheets['第四题']
    #print (sheet.sheet_names())
    # sheet的名称，行数，列数
    print ("ghgvjhbjh")
    #print (sheet.name,sheet.nrows,sheet.ncols)

    # 获取整行和整列的值（数组）
    #rows = sheet.row_values(3) # 获取第四行内容
    #cols = sheet.col_values(11) # 获取第十二列内容
    cols = sheet.range('L2:L35582').value
    #print rows
    #print cols

    # 获取单元格内容
    #print ("这是cols")
    #print (cols)
    #sheet.range('A1').api.EntireRow.Delete()
    #删除不足六年的数据
    count=0
    #while not cols[count] is None:
    total=35581
    while count<total:
        conm=cols[count]
        if conm==cols[count+5]:
            count+=6
        else:
            temp=count
            temp2=count
            while temp<(temp2 + 5):
                temp=temp+1
                if cols[count]!=cols[count+1]:
                    sheet.range('A'+str(count+2)).api.EntireRow.Delete()
                    total=total-1
                    cols = sheet.range('L2:L'+str(total)).value
                    print("不等") 
                    break
                else:
                    #执行删除
                    sheet.range('A'+str(count+2)).api.EntireRow.Delete()
                    print(count+2)  
                    total=total-1
                    cols = sheet.range('L2:L'+str(total)).value
                    #print(cols[count])
                    #print(cols[count+1])
            print("完成一个公司")
            
    wb.save()    
if __name__ == '__main__':
    read_excel()