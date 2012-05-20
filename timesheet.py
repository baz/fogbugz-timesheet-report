import fbSettings
from fogbugz import FogBugz
from datetime import datetime, timedelta
from pytz import timezone
from decimal import *
import pytz
import csv

fb = FogBugz(fbSettings.URL, fbSettings.TOKEN)

# Use correct timezone
userTimezone = timezone('Australia/Sydney')

# Localized dates according to user's timezones
# Midnight 14th
startDateLocal = userTimezone.localize(datetime(2012, 05, 14, 0, 0))
# 11.59pm 19th
endDateLocal = userTimezone.localize(datetime(2012, 05, 19, 23, 59))

# API expects UTC dates so convert
utc = pytz.utc
startDateUTC = utc.normalize(startDateLocal.astimezone(utc))
endDateUTC = utc.normalize(endDateLocal.astimezone(utc))

# Expected format by API
formatString = '%Y-%m-%dT%H:%M:%SZ'

# Logged-in user's timesheet
response = fb.listIntervals(dtStart = startDateUTC.strftime(formatString), dtEnd = endDateUTC.strftime(formatString))

# Reference of cases and their project
allCases = fb.search(q='orderby:\'lastupdated\'',cols="ixBug,sProject")
caseProjects = {}
for case in allCases.findAll('case'):
	caseProjects[case.ixbug.string] = case.sproject;

f = open('report.csv', 'wt')
try:
	writer = csv.writer(f, quoting=csv.QUOTE_NONNUMERIC)
	writer.writerow(('ID', 'Start Date', 'End Date', 'Duration', 'Duration (seconds)', 'Description', 'Project'))
	outputFormatString = '%d %b %Y, %H:%M'
	totalDeltaSeconds = 0
	for interval in response.findAll('interval'):
		# Convert dates from API (UTC) to user's timezone
		startDateConvertedUTC = utc.localize(datetime.strptime(interval.dtstart.string, formatString))
		startDateConvertedLocal = startDateConvertedUTC.astimezone(userTimezone)
		startDateString = startDateConvertedLocal.strftime(outputFormatString)

		endDateConvertedUTC = utc.localize(datetime.strptime(interval.dtend.string, formatString))
		endDateConvertedLocal = endDateConvertedUTC.astimezone(userTimezone)
		endDateString = endDateConvertedLocal.strftime(outputFormatString)

		delta = endDateConvertedUTC - startDateConvertedUTC
		totalDeltaSeconds = totalDeltaSeconds + delta.seconds

		project = caseProjects[interval.ixbug.string].string.encode('UTF-8')
		writer.writerow((interval.ixbug.string, startDateString, endDateString, str(delta), delta.seconds, interval.stitle.string.encode('UTF-8'), project))

	# Write total
	totalHours = totalDeltaSeconds / float(60 * 60)
	writer.writerow(('TOTAL (hours)', '', '', totalHours))
finally:
	f.close()

print open('report.csv', 'rt').read()
