# -*- coding: utf-8 -*-
"""
Created on Sat Nov 24 19:32:29 2018

@author: Wang
"""

import xlrd
import xlwt
import xlwings as xw

def read_excel():
    wb = xw.Book(r'E:\第四题1.xlsx')
    sheet=wb.sheets['第四题']
    cols = sheet.range('P2:P23426').value
    count=0
    
    total=23426
    while count<total-1:
        if not cols[count] is None and not cols[count+1] is None and not cols[count+2] is None and not cols[count+3] is None and not cols[count+4] is None:
            count+=5
        else:
            sheet.range('A'+str(count+2)).api.EntireRow.Delete()
            sheet.range('A'+str(count+2)).api.EntireRow.Delete()
            sheet.range('A'+str(count+2)).api.EntireRow.Delete()
            sheet.range('A'+str(count+2)).api.EntireRow.Delete()
            sheet.range('A'+str(count+2)).api.EntireRow.Delete()
            total-=5
            cols = sheet.range('P2:P'+str(total)).value
    
    '''
    while count<28110:
        sheet.range('A'+str(count+2)).api.EntireRow.Delete()
        count+=5
    '''    
    '''
    while count<28109:
        if not cols[count] is None and not cols[count+1] is None and cols[count+1]!=0 and cols[count]!=0:
            if (count+1) % 6 !=0:
                sheet.range('P'+str(count+3)).value=cols[count+1]/cols[count]
            count+=1
    
        else:
            count+=1
    '''
            
    #wb.save()    
if __name__ == '__main__':
    read_excel()