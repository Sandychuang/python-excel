# -*- coding: utf-8 -*-
"""
Created on Fri Mar 30 14:58:29 2018

@author: zhangshuangxi
"""

import openpyxl
import os
import win32com
from win32com.client import Dispatch, constants

wb= openpyxl.load_workbook('MTO.xlsm',data_only=False,keep_vba=True)
#print wb.get_sheet_names()
wb1= wb[u'\u8f93\u5165\u6570\u636e']
wb2= wb[u'\u6307\u6807']
print(wb1['G35'].value)
print(wb2['E26'].value)
wb1['G35']=260
wb.save('mto1.xlsm')
wb.close()


#os.startfile('deal.xlsm')
xlsApp= win32com.client.Dispatch('Excel.Application')
xlsApp.Workbooks.Open('D:\Documents\Desktop\mto1.xlsm')
xlsApp.Run('de')
xlsApp.DisplayAlerts=False
xlsApp.Application.save
xlsApp.Application.quit()

wbnew= openpyxl.load_workbook('mto1.xlsm',data_only=True,keep_vba=True)

wb1= wbnew[u'\u8f93\u5165\u6570\u636e']
wb2= wbnew[u'\u6307\u6807']
wbnew.close()
#os.remove('mto1.xlsm')
#os.remove('mto6.xlsx')
print(wb1['G35'].value)
print(wb2['E26'].value)

