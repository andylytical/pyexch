import datetime
import pprint
import re
import os
import getpass
import collections
import logging

# From ENV
import tzlocal
import exchangelib

# TODO - Allow env var PYEXCH_REGEX as a filename to a YAML file
# TODO - (better?) Allow env var PYEXCH_CONFIG as a filename to a YAML file

###
# Work-around for missing TZ's in exchangelib
#
# Hopefully will get fixed in https://github.com/ncsa/exchangelib/pull/1
missing_timezones = { 'America/Chicago': 'Central Standard Time' }
exchangelib.EWSTimeZone.PYTZ_TO_MS_MAP.update( missing_timezones )
#
###

simple_event = collections.namedtuple( 'SimpleEvent', [ "start",
                                                        "end",
                                                        "elapsed",
                                                        "is_all_day",
                                                        "type" ] )


class PyExch( object ):
    ''' Get calendar events from Exchange, convert to hours (worked, sick, 
        vacation, etc.) per day.
    '''

    DEFAULT_REGEX_CLASSES = {
        'SICK'     : '(sick|doctor|dr. appt)',
        'VACATION' : '(vacation|OOTO|OOO|out of the office|out of office)',
    }

    def __init__( self, user=None, ad_domain=None, email_domain=None, pwd=None, regex_map=None ):
        ''' User 
        should be in DOMAIN\\USERNAME format.
        '''
        self.user = user
        self.ad_domain = ad_domain
        self.email_domain = email_domain
        self.pwd = pwd
        self.regex_map = regex_map
        self._try_load_from_env()
        self.full_user = '{}\\{}'.format( self.ad_domain, self.user )
        self.full_email = '{}@{}'.format( self.user, self.email_domain )
        if not regex_map:
            self.regex_map = self.DEFAULT_REGEX_CLASSES
        self.re_map = { k: re.compile( v, re.IGNORECASE ) for k,v in self.regex_map.items() }
        self.tz = None
        self._set_timezone()
        self.credentials = exchangelib.Credentials( username=self.full_user, 
                                                    password=self.pwd )
        self.account = exchangelib.Account( primary_smtp_address=self.full_email, 
                                            credentials=self.credentials,
                                            autodiscover=True, 
                                            access_type=exchangelib.DELEGATE )
        

    def _try_load_from_env( self ):
        # USER
        if not self.user:
            self.user = os.getenv( 'PYEXCH_USER' )
        if not self.user:
            self.user = getpass.getuser()
        # AD_DOMAIN
        if not self.ad_domain:
            self.ad_domain = os.environ[ 'PYEXCH_AD_DOMAIN' ]
        # EMAIL_DOMAIN
        if not self.email_domain:
            self.email_domain = os.environ[ 'PYEXCH_EMAIL_DOMAIN' ]
        # PASSWORD
        if not self.pwd:
            pfile = None
            pfile = os.getenv( 'PYEXCH_PWD_FILE' )
            if pfile:
                with open( pfile ) as f:
                    self.pwd = f.read().strip()
            else:
                self.pwd = getpass.getpass()

    def _set_timezone( self ):
        tz_str = tzlocal.get_localzone()
        self.tz = exchangelib.EWSTimeZone.from_pytz( tz_str )
        pprint.pprint( [ 'LOCALTIMEZONE', self.tz ] )


    def get_events_filtered( self, start ):
        logging.debug( 'Enter get_events_filtered' )
        calendar_events = []
        cal_start = exchangelib.EWSDateTime.from_datetime( start )
        if not start.tzinfo:
            cal_start = self.tz.localize( exchangelib.EWSDateTime.from_datetime( start ) )
        cal_end = self.tz.localize( exchangelib.EWSDateTime.now() )
        items = self.account.calendar.view( start=cal_start, end=cal_end )
        for item in items:
            for typ,regx in self.re_map.items():
                if regx.search( item.subject ):
                    calendar_events.append( self.as_simple_event( item, typ ) )
        return calendar_events


    def as_simple_event( self, event, typ ):
        start = event.start.astimezone( self.tz )
        end = event.end.astimezone( self.tz )
        elapsed = end - start
        is_all_day = event.is_all_day
        return simple_event( start, end, elapsed, is_all_day, typ )


    def event_to_daily_data( self, e ):
        elapsed = int( e.elapsed.total_seconds() )
        daily_secs = []
        dayone_secs = elapsed
        if elapsed > 86400:
            # num seconds from start time of event to end of day
            dayone_secs = 86400 - ( e.start.hour * 3600 ) - ( e.start.minute * 60 ) - e.start.second
        daily_secs.append( dayone_secs )
        elapsed = elapsed - dayone_secs
        while elapsed > 0:
            daily_secs.append( min( 86400, elapsed ) )
            elapsed = elapsed - 86400
        daily_data = {}
        for i in range( len( daily_secs ) ):
            diff = datetime.timedelta( days=i )
            thedate = ( e.start + diff ).date()
            daily_data[ thedate ] = { e.type: daily_secs[i] }
        return daily_data


    def per_day_report( self, start ):
        ''' For each day, return a dict mapping 
            regex CLASS to seconds spent in that class
            where "CLASS" is a key from regex_map (passed to init)
        '''
        raw_events = self.get_events_filtered( start )
        dates = {}
        for e in raw_events:
            daily_data = self.event_to_daily_data( e )
            #pprint.pprint( daily_data )
            for thedate, data in daily_data.items():
                if thedate not in dates:
                    dates[ thedate ] = {}
                for ev_type, secs in data.items():
                    if ev_type not in dates[ thedate ]:
                        dates[ thedate ][ ev_type ] = 0
                    dates[ thedate ][ ev_type ] += secs
        return dates


if __name__ == '__main__':
    raise UserWarning( "Command line invocation unsupported" )
