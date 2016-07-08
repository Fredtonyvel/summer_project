import openpyxl
import os
from openpyxl.workbook import Workbook
from openpyxl.styles import Font, Fill
from datetime import datetime

now = datetime.now()
#print now.strftime("%m/%d/%Y %H:%M:%S")

#loading workbook => summer_projxl.xlsx
wb = openpyxl.load_workbook('summer_projxl.xlsx')
type(wb)
ws = wb.active	#Allows you to write into speadsheet
c = ws['A1']		
d = ws['B1']
e = ws['C1']
f = ws['D1']

#Applying style changes to each object
c.font = Font(bold=True)
d.font = Font(bold=True)	
e.font = Font(bold=True)
f.font = Font(bold=True)

#Writing values to cells (Title of each column)
sheet = wb.get_sheet_by_name('Sheet1')	#Picked Sheet1
sheet['A1'] = 'Name'
sheet['B1'] = 'Id_Num'
sheet['C1'] = 'Date'
sheet['D1'] = 'Time'

#Adding values to each cell under appropiate column
sheet['A2'] = 'Freddy'
sheet['B2'] = '515'
sheet['C2'] = now.strftime("%m/%d/%Y")
sheet['D2'] = now.strftime("%H:%M:%S")

sheet['A3'] = 'Gary'
sheet['B3'] = '428'
sheet['C3'] = now.strftime("%m/%d/%Y")
sheet['D3'] = now.strftime("%H:%M:%S")

sheet['A4'] = 'Ashok'
sheet['B4'] = '596'
sheet['C4'] = now.strftime("%m/%d/%Y")
sheet['D4'] = now.strftime("%H:%M:%S")

sheet['A5'] = 'Timothy'
sheet['B5'] = '354'
sheet['C5'] = now.strftime("%m/%d/%Y")
sheet['D5'] = now.strftime("%H:%M:%S")

sheet['A6'] = 'Paul'
sheet['B6'] = '896'
sheet['C6'] = now.strftime("%m/%d/%Y")
sheet['D6'] = now.strftime("%H:%M:%S")

sheet['A7'] = 'Jasmine'
sheet['B7'] = '596'
sheet['C7'] = now.strftime("%m/%d/%Y")
sheet['D7'] = now.strftime("%H:%M:%S")

sheet['A8'] = 'Xavier'
sheet['B8'] = '381'
sheet['C8'] = now.strftime("%m/%d/%Y")
sheet['D8'] = now.strftime("%H:%M:%S")

sheet['A9'] = 'Jane'
sheet['B9'] = '786'
sheet['C9'] = now.strftime("%m/%d/%Y")
sheet['D9'] = now.strftime("%H:%M:%S")

sheet['A10'] = 'Bob'
sheet['B10'] = '685'
sheet['C10'] = now.strftime("%m/%d/%Y")
sheet['D10'] = now.strftime("%H:%M:%S")

sheet['A11'] = 'Zack'
sheet['B11'] = '262'
sheet['C11'] = now.strftime("%m/%d/%Y")
sheet['D11'] = now.strftime("%H:%M:%S")

#Changing column dimensions
#ws.column_dimensions['D'].width = 10.71

wb.save('summer_projxl.xlsx')	#Saves updates on excel file

#Using Python to enter commands through CMD prompt
os.system('summer_projxl.xlsx')