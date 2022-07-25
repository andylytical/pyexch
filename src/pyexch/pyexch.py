import collections
import datetime
import exchangelib
import exchangelib.errors
import getpass
import json
import logging
import netrc
import oauthlib
import os
import pathlib
import pprint
import re
import requests_oauthlib
import yaml

LOGR = logging.getLogger(__name__)

simple_event = collections.namedtuple( 'SimpleEvent', [
    "start",
    "end",
    "elapsed",
    "is_all_day",
    "type",
    "location",
    "subject",
    "raw_event" ] )


class PyExch( object ):
    ''' Get calendar events from Exchange, convert to hours (worked, sick,
        vacation, etc.) per day.
    '''

    netrc_file_default = f"{os.environ['HOME']}/.ssh/netrc"
    oauth_config_file_default = f"{os.environ['HOME']}/.ssh/exchange_oauth.yaml"
    token_file_default = f"{os.environ['HOME']}/.ssh/exchange_token"


    def __init__( self, login=None, account=None, regex_map=None ):
        ''' + login is the exchange credential login name
              NOTE: Might be in the form "user@domain" or "domain\\\\user"
            + account is the primary SMTP address of the exchange account to be accessed
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
        self.account = account
        self.regex_map = regex_map
        self.exch_account = None
        self.oauth_data = None
        self.token_file = None
        self._try_load_from_env()
        if not self.regex_map:
            raise UserWarning( 'Cannot proceed with null regex_map' )
        self.re_map = { k: re.compile( v, re.IGNORECASE ) for k,v in self.regex_map.items() }
        self.tz = exchangelib.EWSTimeZone.localzone()
        self.credentials = exchangelib.OAuth2AuthorizationCodeCredentials(
            client_id=self.oauth_data['client_id'],
            client_secret=self.oauth_data['client_secret'],
            tenant_id=self.oauth_data['tenant_id'],
            access_token=self._get_token(),
        )
        ews_config = exchangelib.Configuration(
            server='outlook.office365.com',
            credentials=self.credentials,
            auth_type=exchangelib.OAUTH2,
        )
        acct_parms_ews = {
            'primary_smtp_address': self.account,
            'access_type': exchangelib.DELEGATE,
            'config': ews_config,
            'autodiscover': False,
        }
        self.exch_account = exchangelib.Account( **acct_parms_ews )
        if not self.exch_account:
            raise UserWarning( 'Error while logging into Exchange Cloud Service' )


    def _try_load_from_env( self ):
        # attempt to load from NETRC
        netrc_file = os.getenv( 'NETRC', self.netrc_file_default )
        nrc = netrc.netrc( netrc_file )
        nrc_parts = nrc.authenticators( 'EXCH' )
        if nrc_parts:
            if not self.login:
                self.login = nrc_parts[0]
            if not self.account:
                self.account = nrc_parts[1]
        # REGEX
        if not self.regex_map:
            json_str = os.getenv( 'PYEXCH_REGEX_JSON' )
            if json_str:
                self.regex_map = json.loads( json_str )
        # OAUTH CONFIG
        if not self.oauth_data:
            oauth_config_file = os.getenv( 'PYEXCH_OAUTH_CONFIG', self.oauth_config_file_default )
            p = pathlib.Path( oauth_config_file )
            self.oauth_data = yaml.safe_load( p.read_text() )
        # OAUTH TOKEN FILE
        if not self.token_file:
            self.token_file = os.getenv( 'PYEXCH_TOKEN_FILE', self.token_file_default )


    def _get_token( self ):
        p = pathlib.Path( self.token_file )
        if p.is_file():
            with p.open() as f:
                token = json.load( f )
            if self._is_token_expired( token ):
                token = self._new_token_from_auth_code()
        else:
            token = self._new_token_from_auth_code()
        LOGR.debug( f"Token:\n{token}" )
        return token


    def _new_token_from_auth_code( self ):
        base_url = ''.join((
            'https://login.microsoftonline.com/',
            self.oauth_data['tenant_id'],
            '/oauth2/v2.0',
            ))
        auth_url = f'{base_url}/authorize'
        token_url = f'{base_url}/token'
        redirect_url = 'https://login.microsoftonline.com/common/oauth2/nativeclient'
        oa2session = requests_oauthlib.OAuth2Session(
            self.oauth_data['client_id'],
            scope=self.oauth_data['scope'],
            redirect_uri=redirect_url
            )
        auth_uri, state = oa2session.authorization_url( auth_url )
        print( f"AUTH URI:\n{auth_uri}" )
        response = input( 'Paste AUTH URI HTTP response here:' )
        token = oa2session.fetch_token(
            token_url,
            authorization_response=response,
            include_client_id=True
            )
        p = pathlib.Path( self.token_file )
        with p.open( mode='w' ) as f:
            json.dump( token, f )


    def _is_token_expired( self, token ):
        expiration = datetime.datetime.fromtimestamp( token['expires_at'] )
        now = datetime.datetime.now()
        return now >= expiration


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
        return simple_event(
            start=start,
            end=end,
            elapsed=elapsed,
            is_all_day=event.is_all_day,
            type=typ,
            location=event.location,
            subject=event.subject,
            raw_event=event )


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


    def new_event( self, start, end, subject, attendees, location, is_all_day=False, categories=None, free=False ):
        ''' start = Python datetime.datetime
            end = Python datetime.datetime
            subject = String
            attendees = list of emails
            location = String (usually the URL to a Zoom meeting)
            is_all_day = Boolean, set to True to make the meeting all day
            categories = list of categories
            free = Boolean, if true, event will not block the calendar
        '''
        # convert start and end to timezone aware EWSDateTime's
        cal_start = exchangelib.EWSDateTime.from_datetime( start )
        if not start.tzinfo:
            cal_start = exchangelib.EWSDateTime.from_datetime( start ).astimezone( self.tz )
        cal_end = exchangelib.EWSDateTime.from_datetime( end )
        if not end.tzinfo:
            cal_end = exchangelib.EWSDateTime.from_datetime( end ).astimezone( self.tz )
        params = {
            'account': self.exch_account,
            'folder': self.exch_account.calendar,
            'start': cal_start,
            'end': cal_end,
            'is_all_day': is_all_day,
            'subject': subject,
            'required_attendees': attendees,
            'location': location,
        }
        if len(categories) > 0:
            params[ 'categories' ] = categories
        if free:
            params[ 'legacy_free_busy_status' ] = 'Free'
        item = exchangelib.CalendarItem( **params )
        item.save(send_meeting_invitations=exchangelib.items.SEND_ONLY_TO_ALL)
        return item


    def new_all_day_event( self, date, subject, attendees, location, categories=None, free=False ):
        ''' convenience method for an all day event
            date = Python datetime.date
            subject = String
            attendees = list of emails
            location = String (usually the URL to a Zoom meeting)
            categories = list of categories
            free = Boolean, if true, event will not block the calendar
        '''
        params = {
            'start' : datetime.datetime.combine( date, datetime.time.min ),
            'end' : datetime.datetime.combine( date, datetime.time.max ),
            'subject' : subject,
            'attendees' : attendees,
            'location' : location,
            'is_all_day' : True,
            'categories' : categories,
            'free' : free,
        }
        return self.new_event( **params )


if __name__ == '__main__':
    raise UserWarning( "Command line invocation unsupported" )
