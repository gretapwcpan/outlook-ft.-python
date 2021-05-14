# -*- coding: utf-8 -*-
"""
Created on Wed Jul  1 17:59:48 2020

"""

import pandas as pd
import win32com.client as win32
import openpyxl
import xlwings as xw
from openpyxl import load_workbook
import os
import shutil


#clear old reports
try:
    os.remove('output.xlsx')
except:
    pass


#Screen fbl5n and create attachments
data = pd.read_csv('export.csv', encoding = 'utf-8')
ID = pd.read_csv('ID.csv', encoding = 'utf-8')
#adjust data and save it as output in excel
data = pd.merge(data, ID, left_on= 'Last changed by', right_on = 'WBI', how = 'left').drop('WBI', axis =1)
#re-arrange order of columns
col = list(data.columns)[:2] + [list(data.columns)[-1]]+ list(data.columns)[2: -1]
data = data[col].drop(['New value.1','Old value.1'], axis = 1)
data['Comment if something you think is not right'] = ""
data.to_excel(r'C:\Users\nxf33342\Desktop\monthly auto sending report\output.xlsx', encoding = 'utf-8', index = False)





#visualization
wb = load_workbook('./output.xlsx')
ws = wb['Sheet1']
ws.row_dimensions[1].height = 140



#fit column cells with lehgth of data 

excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(r'C:\Users\nxf33342\Desktop\monthly auto sending report\output.xlsx')
ws = wb.Worksheets("Sheet1")
ws.Columns.AutoFit()
ws.Range("A1", "I1").Interior.ColorIndex = 15 #grey
ws.Range("J1", "J1").Interior.ColorIndex = 36 #green
wb.Save()
excel.Application.Quit()

Edate = input('input the end of the period ')
Bdate = input('input the begin of the period ')

try:
    os.rename(os.path.join(r'C:\Users\nxf33342\Desktop\monthly auto sending report\output.xlsx'), 'Credit limits check {}.xlsx'.format(Edate))
except:
    pass
#also save in sharedrive
hyperlink = 'K:\Regular Reports\Credit limits check\Credit limits check {}.xlsx'.format(Edate)
shutil.move(r'C:\Users\nxf33342\Desktop\monthly auto sending report\Credit limits check {}.xlsx'.format(Edate), hyperlink)

#send data
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
receivers = ['Global_Credit@nxp.com']
mail.To = receivers[0]
mail.Subject = 'Action required - Credit limits check {}'.format(Edate)

mail.Body = 'Dear Team Members,\r\nPlease take a close look again at all credit limit modifications from {} to {}\r\n'.format(Bdate, Edate) + r'<\\twgtpetc02ms014\crm\Regular Reports\Credit limits check\Credit limits check {}.xlsx>'.format(Edate)+'\n*No action required if the information is believed right.\n*Do revise the figures in FD32 and let Charlene, your direct supervisor know if something is wrong.\n\nThanks a lot for your assistance.\n\nBest Regards,\nGreta'
#mail.Attachments.Add(r'C:\Users\nxf33342\Desktop\monthly auto sending report\Credit limits check {}.xlsx'.format(Edate))

mail.Send()

#set a fixed time to send e-mail
