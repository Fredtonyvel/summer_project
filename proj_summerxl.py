import openpyxl
import os
from openpyxl.workbook import Workbook
from openpyxl.styles import Font, Fill

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

#Writing values to cells
sheet = wb.get_sheet_by_name('Sheet1')	#Picked Sheet1
sheet['A1'] = 'Name'
sheet['B1'] = 'Id_Num'
sheet['C1'] = 'Date'
sheet['D1'] = 'Time'

#Changing column dimensions
#ws.column_dimensions['D'].width = 10.71

wb.save('summer_projxl.xlsx')	#Saves updates on excel file

#Using Python to enter commands through CMD prompt
os.system('summer_projxl.xlsx')