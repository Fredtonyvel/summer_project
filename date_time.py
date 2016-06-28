'''
from datetime import datetime #imports datetime library to be used 

print datetime.now()
'''

'''
from datetime import datetime
now = datetime.now()

current_year = now.year
current_month = now.month
current_day = now.day

print now
print "The year is: " current_year
print "The month is: " current_month
print "The day is: " current_day
'''

from datetime import datetime
now = datetime.now()

print '%s-%s-%s' % (now.year, now.month, now.day)
print '%s/%s/%s' % (now.month, now.day, now.year) #in (mm/dd/yyyy) form
print '%s:%s:%s' % (now.hour, now.minute, now.second) #in (hh:mm:ss) form
