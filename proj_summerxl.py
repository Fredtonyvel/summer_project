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

#Changing column dimensions
ws.column_dimensions['A'].width = 20
ws.column_dimensions['B'].width = 20
ws.column_dimensions['C'].width = 15
ws.column_dimensions['D'].width = 15


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
sheet['A1'] = 'ID Data'
sheet['B1'] = 'Student\'s Name'
sheet['C1'] = 'Date'
sheet['D1'] = 'Time'

#Adding values to each cell under appropiate column
sheet['A2'] = '515'
sheet['B2'] = 'Freddy'
sheet['C2'] = now.strftime("%m/%d/%Y")  #format mm/dd/yyyy
sheet['D2'] = now.strftime("%H:%M:%S")  #format hh:mm

sheet['A3'] = '428'
sheet['B3'] = 'Gary'
sheet['C3'] = now.strftime("%m/%d/%Y")
sheet['D3'] = now.strftime("%H:%M:%S")

sheet['A4'] = '596'
sheet['B4'] = 'Ashok'
sheet['C4'] = now.strftime("%m/%d/%Y")
sheet['D4'] = now.strftime("%H:%M:%S")

sheet['A5'] = '354'
sheet['B5'] = 'Timothy'
sheet['C5'] = now.strftime("%m/%d/%Y")
sheet['D5'] = now.strftime("%H:%M:%S")

sheet['A6'] = '896'
sheet['B6'] = 'Paul'
sheet['C6'] = now.strftime("%m/%d/%Y")
sheet['D6'] = now.strftime("%H:%M:%S")

sheet['A7'] = '596'
sheet['B7'] = 'Jasmine'
sheet['C7'] = now.strftime("%m/%d/%Y")
sheet['D7'] = now.strftime("%H:%M:%S")

sheet['A8'] = '381'
sheet['B8'] = 'Xavier'
sheet['C8'] = now.strftime("%m/%d/%Y")
sheet['D8'] = now.strftime("%H:%M:%S")

sheet['A9'] = '786'
sheet['B9'] = 'Jane'
sheet['C9'] = now.strftime("%m/%d/%Y")
sheet['D9'] = now.strftime("%H:%M:%S")

sheet['A10'] = '685'
sheet['B10'] = 'Bob'
sheet['C10'] = now.strftime("%m/%d/%Y")
sheet['D10'] = now.strftime("%H:%M:%S")

sheet['A11'] = '262'
sheet['B11'] = 'Zack'
sheet['C11'] = now.strftime("%m/%d/%Y")
sheet['D11'] = now.strftime("%H:%M:%S")

for rowOfCellObjects in sheet['A1': 'D11']:
	for cellObj in rowOfCellObjects:
		print(cellObj.coordinate, cellObj.value)
	print('--- END OF ROW ---\n')

wb.save('summer_projxl.xlsx')	#Saves updates on excel file

#Using Python to enter commands through CMD prompt
os.system('summer_projxl.xlsx')