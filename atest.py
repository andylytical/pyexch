import pyexch
import datetime
import pprint

ptr_regex = { 'NOTWORK': '(sick|doctor|dr. appt|vacation|OOTO|OOO|out of the office|out of office)' }
pyex = pyexch.PyExch( regex_map=ptr_regex )
#pyex = pyexch.PyExch()
start = datetime.datetime( 2017, 1, 1 )
#events = pyex.get_events_filtered( start )
#pprint.pprint( events )
report = pyex.per_day_report( start )
pprint.pprint( report )
