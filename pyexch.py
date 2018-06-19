import datetime
import re
import os
import getpass
import collections
import json
import logging
import netrc

# From LOCAL
import tzlocal
import exchangelib
import exchangelib.errors

LOGR = logging.getLogger(__name__)

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
                                                        "type",
                                                        "location",
                                                        "subject" ] )


class PyExch( object ):
    ''' Get calendar events from Exchange, convert to hours (worked, sick, 
        vacation, etc.) per day.
    '''

    def __init__( self, login=None, pwd=None, account=None, regex_map=None ):
        ''' + login is the exchange credential login name
              NOTE: Might be in the form "user@domain" or "domain\\\\user"
            + pwd is the exchange credential password
            + account is the primary SMTP address of the exchange account associated with "login"
            + regex_map is a map of KEY to REGEX used by get_events_filtered()
              Filtering works as follows:
              If subject (of exchange event) matches <REGEX>, then a new
              simple_event is created with type=KEY.  
              If subject does not match any <REGEX>, then exchange event is ignored.
              Resolution priority: 
              1. <regex_map> parameter
              2. PYEXCH_REGEX_JSON environment variable
        '''
        self.login = login
        self.pwd = pwd
        self.account = account
        self.regex_map = regex_map
        self._try_load_from_env()
        if not self.regex_map:
            raise UserWarning( 'Cannot proceed with null regex_map' )
        self.re_map = { k: re.compile( v, re.IGNORECASE ) for k,v in self.regex_map.items() }
        self.tz = None
        self._set_timezone()
        self._validate_auth()
        self.credentials = exchangelib.Credentials( username=self.login, 
                                                    password=self.pwd )
        acct_parms_campus = { 'primary_smtp_address': self.account, 
                              'access_type': exchangelib.DELEGATE,
                              'credentials': self.credentials,
                              'autodiscover': True,
                            }
        # manually specify server for EWS cloud hosted calendar
        ews_config = exchangelib.Configuration( 
            server='outlook.office365.com',
            credentials=self.credentials
        )
        acct_parms_ews = { 'primary_smtp_address': self.account, 
                           'access_type': exchangelib.DELEGATE,
                           'config': ews_config,
                           'autodiscover': False,
                         }

        try:
            # autodiscovery works for campus hosted calendar
            self.exch_account = exchangelib.Account( **acct_parms_campus )
        except ( exchangelib.errors.AutoDiscoverFailed ) as e:
            # manually specify server for EWS cloud hosted calendar
            self.exch_account = exchangelib.Account( **acct_parms_ews )


    def _try_load_from_env( self ):
        # attempt to load from NETRC
        netrc_file = os.getenv( 'NETRC' )
        nrc = netrc.netrc( netrc_file )
        nrc_parts = nrc.authenticators( 'EXCH' )
        if nrc_parts:
            if not self.login:
                self.login = nrc_parts[0]
            if not self.account:
                self.account = nrc_parts[1]
            if not self.pwd:
                self.pwd = nrc_parts[2]
        # REGEX
        if not self.regex_map:
            json_str = os.getenv( 'PYEXCH_REGEX_JSON' )
            if json_str:
                self.regex_map = json.loads( json_str )
            

    def _validate_auth( self ):
#        logging.debug( f"exch login: {self.login}" )
#        logging.debug( f"exch account: {self.account}" )
#        logging.debug( f"exch pwd: {self.pwd}" )
        if not self.login:
            raise UserWarning( 'Cannot proceed with empty exchange login' )
        if not self.account:
            raise UserWarning( 'Cannot proceed with empty exchange account' )
        if not self.pwd:
            raise UserWarning( 'Cannot proceed with empty exchange pwd' )


    def _set_timezone( self ):
        tz_str = tzlocal.get_localzone()
        self.tz = exchangelib.EWSTimeZone.from_pytz( tz_str )
        LOGR.debug( [ 'LOCALTIMEZONE', self.tz ] )


    def get_events_filtered( self, start, end=None ):
        LOGR.debug( 'Enter get_events_filtered' )
        calendar_events = []
        cal_start = exchangelib.EWSDateTime.from_datetime( start )
        if not start.tzinfo:
            cal_start = self.tz.localize( exchangelib.EWSDateTime.from_datetime( start ) )
        cal_end = self.tz.localize( exchangelib.EWSDateTime.now() ) #default
        if end:
            #override default
            cal_end = exchangelib.EWSDateTime.from_datetime( end )
            if not end.tzinfo:
                cal_end = self.tz.localize( exchangelib.EWSDateTime.from_datetime( end ) )
        items = self.exch_account.calendar.view( start=cal_start, end=cal_end )
        for item in items:
            for typ,regx in self.re_map.items():
                if regx.search( item.subject ):
                    calendar_events.append( self.as_simple_event( item, typ ) )
        return calendar_events


    def as_simple_event( self, event, typ ):
        start = event.start.astimezone( self.tz )
        end = event.end.astimezone( self.tz )
        elapsed = end - start
        return simple_event( start=start, 
                             end=end, 
                             elapsed=elapsed, 
                             is_all_day=event.is_all_day, 
                             type=typ, 
                             location=event.location, 
                             subject=event.subject )


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
            where "CLASS" is a KEY from regex_map (passed to init)
        '''
        raw_events = self.get_events_filtered( start )
        dates = {}
        for e in raw_events:
            daily_data = self.event_to_daily_data( e )
            LOGR.debug( daily_data )
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
