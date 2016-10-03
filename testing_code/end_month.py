import calendar
import datetime


year = int(raw_input("Enter year as a number: "))
#print calendar.prcal(year)
month = int(raw_input("Enter month as a number: "))

#if month == 2:
#	month2 = 'February'

#print "\nThe month of", month2, "=", calendar.monthrange(year, month)
print "\nThe month of", month, "=", calendar.monthrange(year, month)

#print calendar.monthrange(2016, 7)

print '\n', calendar.month(2016, 7, w = 0, l = 0)


print "\n\n", "The last day of July is = ", calendar.mdays[datetime.date.today().month]
date = datetime.date.today()
print date, "\n"

'''
count = 0;
last_day = calendar.mdays[datetime.date.today().month]
day = 31

if( day == last_day):
	for i in range (0, 31):
		count += 1
		print "count = ", count 
		#print "day = last_day"
'''