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

#Writing values to cells (Title of each column)
sheet = wb.get_sheet_by_name('Sheet1')	#Picked Sheet1
sheet['A1'] = 'Name'
sheet['B1'] = 'Id_Num'
sheet['C1'] = 'Date'
sheet['D1'] = 'Time'

#Adding values to each cell under appropiate column
sheet['A2'] = 'Freddy'
sheet['B2'] = '515'
sheet['C2'] = '6-30-2016'
sheet['D2'] = '3:53'

sheet['A3'] = 'Gary'
sheet['B3'] = '428'
sheet['C3'] = '6-30-2016'
sheet['D3'] = '3:54'

sheet['A4'] = 'Ashok'
sheet['B4'] = '596'
sheet['C4'] = '6-30-2016'
sheet['D4'] = '3:55'

#Changing column dimensions
#ws.column_dimensions['D'].width = 10.71

wb.save('summer_projxl.xlsx')	#Saves updates on excel file

#Using Python to enter commands through CMD prompt
os.system('summer_projxl.xlsx')