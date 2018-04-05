import pyexch
import datetime
import pprint
import netrc
import os

#netrcfn = os.getenv( 'NETRC' )
#nrc = netrc.netrc( netrcfn )
#( u, a, p ) = nrc.authenticators( 'EXCH' )
#nrc_parts = { 'login': u, 'account': a, 'pwd': p }
#pyex = pyexch.PyExch( **nrc_parts )

#ptr_regex = { 'NOTWORK': '(sick|doctor|dr. appt|vacation|OOTO|OOO|out of the office|out of office)' }
#pyex = pyexch.PyExch( regex_map=ptr_regex )

pyex = pyexch.PyExch()

start = datetime.datetime( 2018, 1, 1 )
#events = pyex.get_events_filtered( start )
#pprint.pprint( events )
report = pyex.per_day_report( start )
pprint.pprint( report )
