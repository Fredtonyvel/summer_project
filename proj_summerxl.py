import openpyxl
import os

wb = openpyxl.load_workbook('summer_projxl.xlsx')
type(wb)

#Writing values to cells
sheet = wb.get_sheet_by_name('Sheet1')
sheet['A1'] = 'Name'
sheet['B1'] = 'Date'
sheet['C1'] = 'Time'
print sheet['A1'].value
print sheet['B1'].value
print sheet['C1'].value
wb.save('summer_projxl.xlsx')

#Using Python to enter commands through CMD prompt
os.system('summer_projxl.xlsx')