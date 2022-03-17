import collections
import datetime
import getpass
import json
import logging
import netrc
import os
import pprint
import re

# From LOCAL
import exchangelib
import exchangelib.errors

LOGR = logging.getLogger(__name__)

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
        self.exch_account = None
        self._try_load_from_env()
        if not self.regex_map:
            raise UserWarning( 'Cannot proceed with null regex_map' )
        self.re_map = { k: re.compile( v, re.IGNORECASE ) for k,v in self.regex_map.items() }
        self.tz = exchangelib.EWSTimeZone.localzone()
        self._validate_auth()
        self.credentials = exchangelib.Credentials( username=self.login, 
                                                    password=self.pwd )
########
#        # Auto-discovery works with local campus server,
#        #  but local server is long gone, so no longer using this method.
#        #  Leaving this here as an example of auto-discovery.
#        acct_parms_campus = { 'primary_smtp_address': self.account, 
#                              'access_type': exchangelib.DELEGATE,
#                              'credentials': self.credentials,
#                              'autodiscover': True,
#                            }
#        try:
#            # autodiscovery works for campus hosted calendar
#            self.exch_account = exchangelib.Account( **acct_parms_campus )
#        except ( exchangelib.errors.AutoDiscoverFailed ) as e:
#            pass
########

        # manually specify server for Outlook 365 (cloud hosted)
        ews_config = exchangelib.Configuration( 
            server='outlook.office365.com',
            credentials=self.credentials
        )
        acct_parms_ews = { 'primary_smtp_address': self.account, 
                           'access_type': exchangelib.DELEGATE,
                           'config': ews_config,
                           'autodiscover': False,
                         }
        self.exch_account = exchangelib.Account( **acct_parms_ews )
        if not self.exch_account:
            raise UserWarning( 'Error while logging into Exchange Cloud Service' )


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


    def get_events_filtered( self, start, end=None ):
        LOGR.debug( 'Enter get_events_filtered' )
        calendar_events = []
        cal_start = exchangelib.EWSDateTime.from_datetime( start )
        if not start.tzinfo:
            cal_start = exchangelib.EWSDateTime.from_datetime( start ).astimezone( self.tz )
        LOGR.debug( pprint.pformat( f'Cal_Start: {cal_start}' ) )
        cal_end = exchangelib.EWSDateTime.now().astimezone( self.tz ) #default
        if end:
            #override default
            cal_end = exchangelib.EWSDateTime.from_datetime( end )
            if not end.tzinfo:
                cal_end = exchangelib.EWSDateTime.from_datetime( end ).astimezone( self.tz )
        LOGR.debug( pprint.pformat( f'Cal_End: {cal_end}' ) )
        items = self.exch_account.calendar.view( start=cal_start, end=cal_end )
        for item in items:
            for typ,regx in self.re_map.items():
                if regx.search( item.subject ):
                    calendar_events.append( self.as_simple_event( item, typ ) )
        return calendar_events


    def as_simple_event( self, event, typ ):
        try:
            start = event.start.astimezone( self.tz )
            end = event.end.astimezone( self.tz )
        except AttributeError as e:
            # convert EWSDate to EWSDateTime (common for all-day events)
            sy, sm, sd = [ getattr( event.start, a ) for a in ( 'year', 'month', 'day' ) ]
            ey, em, ed = [ getattr( event.end, a ) for a in ( 'year', 'month', 'day' ) ]
            start = exchangelib.EWSDateTime( sy, sm, sd, tzinfo=self.tz )
            end = exchangelib.EWSDateTime( ey, em, ed, 23, 59, 59, tzinfo=self.tz )
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


    def per_day_report( self, start, end=None ):
        ''' For each day, return a dict mapping 
            regex CLASS to seconds spent in that class
            where "CLASS" is a KEY from regex_map (passed to init)
        '''
        raw_events = self.get_events_filtered( start, end )
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
