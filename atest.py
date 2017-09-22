import pyexch
import datetime
import pprint
import netrc
import os

netrcfn = os.getenv( 'NETRC' )
nrc = netrc.netrc( netrcfn )
( u, a, p ) = nrc.authenticators( 'PYEXCH' )
nrc_parts = { 'login': u, 'account': a, 'pwd': p }

#ptr_regex = { 'NOTWORK': '(sick|doctor|dr. appt|vacation|OOTO|OOO|out of the office|out of office)' }
#pyex = pyexch.PyExch( regex_map=ptr_regex )
pyex = pyexch.PyExch( **nrc_parts )
start = datetime.datetime( 2017, 1, 1 )
#events = pyex.get_events_filtered( start )
#pprint.pprint( events )
report = pyex.per_day_report( start )
pprint.pprint( report )
