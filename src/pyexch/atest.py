import pyexch
import datetime
import pprint

import logging
#fmt = '%(levelname)s %(message)s'
log_lvl = logging.DEBUG
fmt = '%(levelname)s [%(filename)s:%(funcName)s:%(lineno)s] %(message)s'
logging.basicConfig( level=log_lvl, format=fmt )
# no_debug = [ 'exchangelib' ]
# for key in no_debug:
#     logging.getLogger(key).setLevel(logging.CRITICAL)

pyex = pyexch.PyExch()

start = datetime.datetime( 2022, 8, 1 )
events = pyex.get_events_filtered( start )
for e in events:
    logging.debug(f'Start:{e.start.date()} End:{e.end.date()} Subject:{e.subject}')
# pprint.pprint( events )
# report = pyex.per_day_report( start )
# pprint.pprint( report )
