# pyexch
Python wrapper to Exchange calendar

# Usage Summary
* Oauth Config (`/home/.ssh/exchange_oauth.yaml`)
```yaml
---
tenant_id: '...'
client_id: '...'
client_secret: '...'
scope:
  - 'https://outlook.office365.com/Calendars.ReadWrite.Shared'
  - 'https://outlook.office365.com/EWS.AccessAsUser.All'
```
(see also Azure App notes in: https://github.com/ncsa/asd-triage-scheduler)
* Netrc setup (`/home/.ssh/netrc`)
```
machine EXCH
account primary_SMTP@illinois.edu
```
* Environment variables
  * `OAUTH_CONFIG_FILE=/home/.ssh/exchange_oauth.yaml`
  * `OAUTH_TOKEN_FILE=/home/.ssh/exchange_token`
  * `NETRC=/home/.ssh/netrc`
  * `PYEXCH_REGEX_JSON='{"SICK":"(sick|doctor|dr.appt)","VACATION":"(vacation|PTO|paid time off|personal day)"}'`
* Install pyexch
  * `pip install git+https://github.com/andylytical/pyexch@v3.1.0`
* Use pyexch
```
import pyexch

# Get parameter values from enviroment variables
px = pyexch.pyexch.PyExch()

# Pass parameter values directly
px = pyexch.pyexch.PyExch(
  oauth_conf={'tenant_id':'...',
  client_id='...',
  client_secret='...',
  scope=( ... , ... ),
  account='primary_SMTP@illinois.edu',
  regex_map={'SICK':'(sick|doctor|dr.appt)', ... }
)

# Get all events since 1 Aug 2022 till now,
# with subject matching a regular expression in PYEXCH_REGEX_JSON
start = datetime.datetime(2022, 8, 1)
events = px.get_events_filtered( start )
for e in events:
  print( f'Subj:{e.subject}, Type:{e.type}' )
  actual_exch_event = e.raw_event
  # do something with the raw exchange event

# Create a new all day event
day = datetime.date( 2022, 11, 24)
px.new_all_day_event( day, subject='Holiday')

# Invite others to a Zoom meeting
px.new_event(
  start=datetime.datetime( 2022, 8, 15, hour=9, minute=0 )
  end=start + datetime.timedelta( seconds=900 ) #50 minute meeting
  subject='Morning standup',
  attendees=( 'myboss@illinois.edu','mystaff@illinois.edu' ),
  location='https://illinois.zoom.us/j/12345678?pwd=xxxyyy',
)
```
# Notes
* Creating an instance will attempt to login,
  which will in turn generate a URL for authorization,
  wait for authcode_response,
  retrieve access & refresh tokens,
  and store them in OAUTH_TOKEN_FILE.
* Will use existing tokens in OAUTH_TOKEN_FILE.
* If existing tokens are expired, a new authorization sequence will be
  initiated.
