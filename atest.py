import pyexch
import datetime
import pprint

pyex = pyexch.PyExch()

start = datetime.datetime( 2018, 1, 1 )
#events = pyex.get_events_filtered( start )
#pprint.pprint( events )
report = pyex.per_day_report( start )
pprint.pprint( report )
