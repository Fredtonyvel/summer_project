import openpyxl
import calendar
import os
from datetime import datetime
from openpyxl.cell import get_column_letter

#loading workbook => summer_projxl.xlsx
wb = openpyxl.load_workbook('month_pyxl.xlsx')
type(wb)
ws = wb.active	#Allows you to write into speadsheet

#Writing values to cells (Title of each column)
sheet = wb.get_sheet_by_name('Sheet1')	#Picked Sheet1

sheet['A1'] = 'Date'
sheet['B1'] = 'Time'

print sheet['A1'].value
print sheet['B1'].value





now = datetime.now()

for i in range (2, 11):
	date = now.strftime("%m/%d/%Y")
	sheet['A' + str(i)] = date
	print i, ' = ', date

print "\n"

for l in range (2,11):
	#time = now.strftime("%H:%M:%S")
	time = now.strftime("%H:%M")
	sheet['B' + str(l)] = time
	print l, ' = ', time

print "\nsheet.max_row =", sheet.max_row
print "sheet.max_column =", get_column_letter(sheet.max_column)

print "\n"




sheet_value = sheet[get_column_letter(sheet.max_column) + str(sheet.max_row)].value
'''
#for i in range (1, sheet.max_row):
if (sheet_value == '12:29'):
	print "Match!"

else:
	print "No Match!"
'''

if (sheet_value == '12:41'):
	sheet_count = sheet.max_column
	sheet_count += 1
	print "New sheet.max_column =", get_column_letter(sheet_count)
else:
	print "Didn't work!"






wb.save('month_pyxl.xlsx')	#Saves updates on excel file

#Using Python to enter commands through CMD prompt
os.system('month_pyxl.xlsx')