from datetime import datetime
import logging

import pytz

log = logging.getLogger(__name__)


class EWSDateTime:
    """
    Extends the normal datetime implementation to satisfy EWS
    """

    def __init__(self, year, month, day, hour=0, minute=0, second=0, microsecond=0, tzinfo=None):
        """
        EWS expects dates in UTC. To be on the safe side, we require dates to be created with a timezone. Due to
        datetime() weirdness on DST, we can't inherit from the datetime.datetime class. Instead, emulate it using
        self.dt and __getattr__(). If a non-UTC timezone is used, the local timezone must be passed to Exchange using
        the MeetingTimezone, StartTimezone, EndTimezone or TimeZoneContext elements when creating items.
        """
        if tzinfo is None:
            raise ValueError('Must specify a timezone on EWSDateTime objects')
        self.dt = tzinfo.localize(datetime(year, month, day, hour, minute, second, microsecond))

    def ewsformat(self, tzinfo=None):
        """
        ISO 8601 format to satisfy xs:datetime as interpreted by EWS. Example: 2009-01-15T13:45:56Z
        """
        if tzinfo:
            if tzinfo == pytz.utc:
                return self.dt.astimezone(tzinfo).strftime('%Y-%m-%dT%H:%M:%SZ')
            return self.dt.astimezone(tzinfo).strftime('%Y-%m-%dT%H:%M:%S')
        if self.dt.tzinfo == pytz.utc:
            return self.dt.strftime('%Y-%m-%dT%H:%M:%SZ')
        return self.dt.strftime('%Y-%m-%dT%H:%M:%S')

    def __getattr__(self, attr):
        return getattr(self.dt, attr)

    def to_datetime(self):
        return self.dt

    @classmethod
    def from_datetime(cls, d, tzinfo=None):
        if not (d.tzinfo or tzinfo):
            raise ValueError('Must specify a timezone')
        return cls(d.year, d.month, d.day, d.hour, d.minute, d.second, d.microsecond, d.tzinfo or tzinfo)

    def __add__(self, other):
        # Can add timespans to get a new EWSDateTime
        t = self.dt.__add__(other)
        return self.__class__.from_datetime(t)

    def __sub__(self, other):
        # Can subtract datetimes to get a timespan, or timespan to get a new EWSDateTime
        t = self.dt.__sub__(other)
        if type(t) == datetime:
            return self.__class__.from_datetime(t)
        return t

    def __eq__(self, other):
        return self.dt == (other.dt if isinstance(other, self.__class__) else other)

    def __ne__(self, other):
        return self.dt != (other.dt if isinstance(other, self.__class__) else other)

    def __lt__(self, other):
        return self.__cmp__(other) < 0

    def __le__(self, other):
        return self.__cmp__(other) <= 0

    def __gt__(self, other):
        return self.__cmp__(other) > 0

    def __ge__(self, other):
        return self.__cmp__(other) >= 0

    def __cmp__(self, other):
        a, b = self.dt, other.dt if isinstance(other, self.__class__) else other
        return (a > b) - (a < b)

    def get_ewstimezone(self):
        return EWSTimeZone.from_pytz(self.dt.tzinfo)

    @classmethod
    def from_string(cls, date_string):
        dt = datetime.strptime(date_string, '%Y-%m-%dT%H:%M:%SZ')
        return cls.from_datetime(dt, tzinfo=pytz.utc)

    @classmethod
    def now(cls, tzinfo=pytz.utc):
        t = datetime.now(tzinfo)
        return cls.from_datetime(t)

    def __repr__(self):
        # Mimic datetime repr behavior
        attrs = [self.dt.year, self.dt.month, self.dt.day, self.dt.hour, self.dt.minute, self.dt.second,
                 self.dt.microsecond]
        for i in (6, 5):
            if attrs[i]:
                break
            del attrs[i]
        if self.dt.tzinfo:
            return self.__class__.__name__ + '(%s, tzinfo=%s)' % (', '.join([str(i) for i in attrs]),
                                                                  repr(self.dt.tzinfo))
        else:
            return self.__class__.__name__ + repr(tuple(attrs))
    
    def __str__(self):
        return str(self.dt)


class EWSTimeZone:
    """
    Represents a timezone as expected by the EWS TimezoneContext / TimezoneDefinition XML element, and returned by
    services.GetServerTimeZones.
    """
    def __init__(self, tz, ms_id, name):
        self.tz = tz  # pytz timezone
        self.ms_id = ms_id
        self.name = name

    @classmethod
    def from_pytz(cls, tz):
        try:
            ms_id = cls.PYTZ_TO_MS_MAP[str(tz)]
        except KeyError as e:
            raise ValueError('Please add a mapping from "%s" to MS timezone in PYTZ_TO_MS_TZMAP' % tz) from e
        name = cls.MS_TIMEZONE_DEFINITIONS[ms_id]
        return cls(tz, ms_id, name)

    @classmethod
    def from_location(cls, location):
        tz = pytz.timezone(location)
        return cls.from_pytz(tz)

    def localize(self, *args, **kwargs):
        return self.tz.localize(*args, **kwargs)

    def __str__(self):
        return '%s (%s)' % (self.ms_id, self.name)
    
    def __repr__(self):
        return self.__class__.__name__ + repr((self.tz, self.ms_id, self.name))

    # Manually maintained translation between pytz timezone name and MS timezone IDs
    PYTZ_TO_MS_MAP = {
        'Europe/Copenhagen': 'Romance Standard Time',
        'UTC': 'UTC',
        'GMT': 'GMT Standard Time',
    }

    # This is a somewhat authoritative list of the timezones available on an Exchange server. Format is (id, name).
    MS_TIMEZONE_DEFINITIONS = dict([
        ('Dateline Standard Time', '(UTC-12:00) International Date Line West'),
        ('UTC-11', '(UTC-11:00) Coordinated Universal Time-11'),
        ('Samoa Standard Time', '(UTC-11:00) Midway Island, Samoa'),
        ('Hawaiian Standard Time', '(UTC-10:00) Hawaii'),
        ('Alaskan Standard Time', '(UTC-09:00) Alaska'),
        ('Pacific Standard Time', '(UTC-08:00) Pacific Time (US & Canada)'),
        ('Pacific Standard Time (Mexico)', '(UTC-08:00) Tijuana, Baja California'),
        ('US Mountain Standard Time', '(UTC-07:00) Arizona'),
        ('Mountain Standard Time (Mexico)', '(UTC-07:00) Chihuahua, La Paz, Mazatlan'),
        ('Mountain Standard Time', '(UTC-07:00) Mountain Time (US & Canada)'),
        ('Central America Standard Time', '(UTC-06:00) Central America'),
        ('Central Standard Time', '(UTC-06:00) Central Time (US & Canada)'),
        ('Central Standard Time (Mexico)', '(UTC-06:00) Guadalajara, Mexico City, Monterrey'),
        ('Canada Central Standard Time', '(UTC-06:00) Saskatchewan'),
        ('SA Pacific Standard Time', '(UTC-05:00) Bogota, Lima, Quito'),
        ('Eastern Standard Time', '(UTC-05:00) Eastern Time (US & Canada)'),
        ('US Eastern Standard Time', '(UTC-05:00) Indiana (East)'),
        ('Venezuela Standard Time', '(UTC-04:30) Caracas'),
        ('Paraguay Standard Time', '(UTC-04:00) Asuncion'),
        ('Atlantic Standard Time', '(UTC-04:00) Atlantic Time (Canada)'),
        ('SA Western Standard Time', '(UTC-04:00) Georgetown, La Paz, San Juan'),
        ('Central Brazilian Standard Time', '(UTC-04:00) Manaus'),
        ('Pacific SA Standard Time', '(UTC-04:00) Santiago'),
        ('Newfoundland Standard Time', '(UTC-03:30) Newfoundland'),
        ('E. South America Standard Time', '(UTC-03:00) Brasilia'),
        ('Argentina Standard Time', '(UTC-03:00) Buenos Aires'),
        ('SA Eastern Standard Time', '(UTC-03:00) Cayenne'),
        ('Greenland Standard Time', '(UTC-03:00) Greenland'),
        ('Montevideo Standard Time', '(UTC-03:00) Montevideo'),
        ('UTC-02', '(UTC-02:00) Coordinated Universal Time-02'),
        ('Mid-Atlantic Standard Time', '(UTC-02:00) Mid-Atlantic'),
        ('Azores Standard Time', '(UTC-01:00) Azores'),
        ('Cape Verde Standard Time', '(UTC-01:00) Cape Verde Is.'),
        ('Morocco Standard Time', '(UTC) Casablanca'),
        ('UTC', '(UTC) Coordinated Universal Time'),
        ('GMT Standard Time', '(UTC) Greenwich Mean Time : Dublin, Edinburgh, Lisbon, London'),
        ('Greenwich Standard Time', '(UTC) Monrovia, Reykjavik'),
        ('W. Europe Standard Time', '(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna'),
        ('Central Europe Standard Time', '(UTC+01:00) Belgrade, Bratislava, Budapest, Ljubljana, Prague'),
        ('Romance Standard Time', '(UTC+01:00) Brussels, Copenhagen, Madrid, Paris'),
        ('Central European Standard Time', '(UTC+01:00) Sarajevo, Skopje, Warsaw, Zagreb'),
        ('W. Central Africa Standard Time', '(UTC+01:00) West Central Africa'),
        ('Namibia Standard Time', '(UTC+02:00) Windhoek'),
        ('Jordan Standard Time', '(UTC+02:00) Amman'),
        ('GTB Standard Time', '(UTC+02:00) Athens, Bucharest, Istanbul'),
        ('Middle East Standard Time', '(UTC+02:00) Beirut'),
        ('Egypt Standard Time', '(UTC+02:00) Cairo'),
        ('Syria Standard Time', '(UTC+02:00) Damascus'),
        ('South Africa Standard Time', '(UTC+02:00) Harare, Pretoria'),
        ('FLE Standard Time', '(UTC+02:00) Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius'),
        ('Israel Standard Time', '(UTC+02:00) Jerusalem'),
        ('E. Europe Standard Time', '(UTC+02:00) Minsk'),
        ('Arabic Standard Time', '(UTC+03:00) Baghdad'),
        ('Arab Standard Time', '(UTC+03:00) Kuwait, Riyadh'),
        ('Russian Standard Time', '(UTC+03:00) Moscow, St. Petersburg, Volgograd'),
        ('E. Africa Standard Time', '(UTC+03:00) Nairobi'),
        ('Iran Standard Time', '(UTC+03:30) Tehran'),
        ('Georgian Standard Time', '(UTC+03:00) Tbilisi'),
        ('Arabian Standard Time', '(UTC+04:00) Abu Dhabi, Muscat'),
        ('Azerbaijan Standard Time', '(UTC+04:00) Baku'),
        ('Mauritius Standard Time', '(UTC+04:00) Port Louis'),
        ('Caucasus Standard Time', '(UTC+04:00) Yerevan'),
        ('Afghanistan Standard Time', '(UTC+04:30) Kabul'),
        ('Ekaterinburg Standard Time', '(UTC+05:00) Ekaterinburg'),
        ('Pakistan Standard Time', '(UTC+05:00) Islamabad, Karachi'),
        ('West Asia Standard Time', '(UTC+05:00) Tashkent'),
        ('India Standard Time', '(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi'),
        ('Sri Lanka Standard Time', '(UTC+05:30) Sri Jayawardenepura'),
        ('Nepal Standard Time', '(UTC+05:45) Kathmandu'),
        ('N. Central Asia Standard Time', '(UTC+06:00) Almaty, Novosibirsk'),
        ('Central Asia Standard Time', '(UTC+06:00) Astana'),
        ('Bangladesh Standard Time', '(UTC+06:00) Dhaka'),
        ('Myanmar Standard Time', '(UTC+06:30) Yangon (Rangoon)'),
        ('SE Asia Standard Time', '(UTC+07:00) Bangkok, Hanoi, Jakarta'),
        ('North Asia Standard Time', '(UTC+07:00) Krasnoyarsk'),
        ('China Standard Time', '(UTC+08:00) Beijing, Chongqing, Hong Kong, Urumqi'),
        ('North Asia East Standard Time', '(UTC+08:00) Irkutsk, Ulaan Bataar'),
        ('Singapore Standard Time', '(UTC+08:00) Kuala Lumpur, Singapore'),
        ('W. Australia Standard Time', '(UTC+08:00) Perth'),
        ('Taipei Standard Time', '(UTC+08:00) Taipei'),
        ('Ulaanbaatar Standard Time', '(UTC+08:00) Ulaanbaatar'),
        ('Tokyo Standard Time', '(UTC+09:00) Osaka, Sapporo, Tokyo'),
        ('Korea Standard Time', '(UTC+09:00) Seoul'),
        ('Yakutsk Standard Time', '(UTC+09:00) Yakutsk'),
        ('Cen. Australia Standard Time', '(UTC+09:30) Adelaide'),
        ('AUS Central Standard Time', '(UTC+09:30) Darwin'),
        ('E. Australia Standard Time', '(UTC+10:00) Brisbane'),
        ('AUS Eastern Standard Time', '(UTC+10:00) Canberra, Melbourne, Sydney'),
        ('West Pacific Standard Time', '(UTC+10:00) Guam, Port Moresby'),
        ('Tasmania Standard Time', '(UTC+10:00) Hobart'),
        ('Vladivostok Standard Time', '(UTC+10:00) Vladivostok'),
        ('Magadan Standard Time', '(UTC+11:00) Magadan'),
        ('Central Pacific Standard Time', '(UTC+11:00) Magadan, Solomon Is., New Caledonia'),
        ('New Zealand Standard Time', '(UTC+12:00) Auckland, Wellington'),
        ('UTC+12', '(UTC+12:00) Coordinated Universal Time+12'),
        ('Fiji Standard Time', '(UTC+12:00) Fiji, Marshall Is.'),
        ('Kamchatka Standard Time', '(UTC+12:00) Petropavlovsk-Kamchatsky'),
        ('Tonga Standard Time', "(UTC+13:00) Nuku'alofa"),
    ])
