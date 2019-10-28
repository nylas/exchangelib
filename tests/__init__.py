# coding=utf-8
from collections import namedtuple
import datetime
from decimal import Decimal
import logging
import os
import random
import string
import time
import unittest
import unittest.util

import flake8.defaults
import flake8.main.application
import pytz
from six import PY2
from yaml import safe_load

from exchangelib.account import Account
from exchangelib.attachments import FileAttachment
from exchangelib.autodiscover import AutodiscoverProtocol
from exchangelib.configuration import Configuration
from exchangelib.credentials import DELEGATE, IMPERSONATION, Credentials
from exchangelib.errors import ErrorItemNotFound, ErrorInvalidOperation, UnknownTimeZone, ErrorAccessDenied, \
    ErrorFolderNotFound, AmbiguousTimeError, NonExistentTimeError, ErrorNoPublicFolderReplicaAvailable
from exchangelib.ewsdatetime import EWSDateTime, EWSDate, EWSTimeZone, UTC, UTC_NOW
from exchangelib.fields import BooleanField, IntegerField, DecimalField, TextField, EmailAddressField, URIField, \
    ChoiceField, BodyField, DateTimeField, Base64Field, PhoneNumberField, EmailAddressesField, TimeZoneField, \
    PhysicalAddressField, ExtendedPropertyField, MailboxField, AttendeesField, AttachmentField, CharListField, \
    MailboxListField, EWSElementField, CultureField, CharField, TextListField, PermissionSetField, MimeContentField
from exchangelib.indexed_properties import EmailAddress, PhysicalAddress, PhoneNumber
from exchangelib.items import CalendarItem, Task
from exchangelib.properties import Attendee, Mailbox, EWSElement, PermissionSet, Permission, UserId
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
from exchangelib.recurrence import Recurrence, DailyPattern
from exchangelib.util import  PrettyXmlHandler

if PY2:
    FileNotFoundError = IOError

mock_account = namedtuple('mock_account', ('protocol', 'version'))
mock_protocol = namedtuple('mock_protocol', ('version', 'service_endpoint'))
mock_version = namedtuple('mock_version', ('build',))

# Show full repr() output for object instances in unittest error messages
unittest.util._MAX_LENGTH = 2000


def mock_post(url, status_code, headers, text=''):
    req = namedtuple('request', ['headers'])(headers={})
    c = text.encode('utf-8')
    return lambda **kwargs: namedtuple(
        'response', ['status_code', 'headers', 'text', 'content', 'request', 'history', 'url']
    )(status_code=status_code, headers=headers, text=text, content=c, request=req, history=None, url=url)


def mock_session_exception(exc_cls):
    def raise_exc(**kwargs):
        raise exc_cls()

    return raise_exc


class MockResponse(object):
    def __init__(self, c):
        self.c = c

    def iter_content(self):
        return self.c


class TimedTestCase(unittest.TestCase):
    SLOW_TEST_DURATION = 5  # Log tests that are slower than this value (in seconds)

    def setUp(self):
        self.maxDiff = None
        self.t1 = time.time()

    def tearDown(self):
        t2 = time.time() - self.t1
        if t2 > self.SLOW_TEST_DURATION:
            print("{:07.3f} : {}".format(t2, self.id()))


class StyleTest(TimedTestCase):
    def test_flake8(self):
        import exchangelib
        flake8.defaults.MAX_LINE_LENGTH = 120
        app = flake8.main.application.Application()
        app.run(exchangelib.__path__)
        # If this fails, look at stdout for actual error messages
        self.assertEqual(app.result_count, 0)


class ItemTest(TimedTestCase):
    def test_task_validation(self):
        tz = EWSTimeZone.timezone('Europe/Copenhagen')
        task = Task(due_date=tz.localize(EWSDateTime(2017, 1, 1)), start_date=tz.localize(EWSDateTime(2017, 2, 1)))
        task.clean()
        # We reset due date if it's before start date
        self.assertEqual(task.due_date, tz.localize(EWSDateTime(2017, 2, 1)))
        self.assertEqual(task.due_date, task.start_date)

        task = Task(complete_date=tz.localize(EWSDateTime(2099, 1, 1)), status=Task.NOT_STARTED)
        task.clean()
        # We reset status if complete_date is set
        self.assertEqual(task.status, Task.COMPLETED)
        # We also reset complete date to now() if it's in the future
        self.assertEqual(task.complete_date.date(), UTC_NOW().date())

        task = Task(complete_date=tz.localize(EWSDateTime(2017, 1, 1)), start_date=tz.localize(EWSDateTime(2017, 2, 1)))
        task.clean()
        # We also reset complete date to start_date if it's before start_date
        self.assertEqual(task.complete_date, task.start_date)

        task = Task(percent_complete=Decimal('50.0'), status=Task.COMPLETED)
        task.clean()
        # We reset percent_complete to 100.0 if state is completed
        self.assertEqual(task.percent_complete, Decimal(100))

        task = Task(percent_complete=Decimal('50.0'), status=Task.NOT_STARTED)
        task.clean()
        # We reset percent_complete to 0.0 if state is not_started
        self.assertEqual(task.percent_complete, Decimal(0))


class EWSTest(TimedTestCase):
    @classmethod
    def setUpClass(cls):
        # There's no official Exchange server we can test against, and we can't really provide credentials for our
        # own test server to everyone on the Internet. Travis-CI uses the encrypted settings.yml.enc for testing.
        #
        # If you want to test against your own server and account, create your own settings.yml with credentials for
        # that server. 'settings.yml.sample' is provided as a template.
        try:
            with open(os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings.yml')) as f:
                settings = safe_load(f)
        except FileNotFoundError:
            print('Skipping %s - no settings.yml file found' % cls.__name__)
            print('Copy settings.yml.sample to settings.yml and enter values for your test server')
            raise unittest.SkipTest('Skipping %s - no settings.yml file found' % cls.__name__)

        cls.verify_ssl = settings.get('verify_ssl', True)
        if not cls.verify_ssl:
            # Allow unverified TLS if requested in settings file
            BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

        # Speed up tests a bit. We don't need to wait 10 seconds for every nonexisting server in the discover dance
        AutodiscoverProtocol.TIMEOUT = 2

        # Create an account shared by all tests
        tz = EWSTimeZone.timezone('Europe/Copenhagen')
        config = Configuration(
            server=settings['server'],
            credentials=Credentials(settings['username'], settings['password'])
        )
        cls.account = Account(primary_smtp_address=settings['account'], access_type=DELEGATE, config=config,
                              locale='da_DK', default_timezone=tz)

    def setUp(self):
        super(EWSTest, self).setUp()
        # Create a random category for each test to avoid crosstalk
        self.categories = [get_random_string(length=16, spaces=False, special=False)]

    def wipe_test_account(self):
        # Deletes up all deleteable items in the test account. Not run in a normal test run
        self.account.root.wipe(page_size=100)

    def bulk_delete(self, ids):
        # Clean up items and check return values
        for res in self.account.bulk_delete(ids):
            self.assertEqual(res, True)

    def random_val(self, field):
        if isinstance(field, ExtendedPropertyField):
            if field.value_cls.property_type == 'StringArray':
                return [get_random_string(255) for _ in range(random.randint(1, 4))]
            if field.value_cls.property_type == 'IntegerArray':
                return [get_random_int(0, 256) for _ in range(random.randint(1, 4))]
            if field.value_cls.property_type == 'BinaryArray':
                return [get_random_string(255).encode() for _ in range(random.randint(1, 4))]
            if field.value_cls.property_type == 'String':
                return get_random_string(255)
            if field.value_cls.property_type == 'Integer':
                return get_random_int(0, 256)
            if field.value_cls.property_type == 'Binary':
                # In the test_extended_distinguished_property test, EWS rull return 4 NULL bytes after char 16 if we
                # send a longer bytes sequence.
                return get_random_string(16).encode()
            raise ValueError('Unsupported field %s' % field)
        if isinstance(field, URIField):
            return get_random_url()
        if isinstance(field, EmailAddressField):
            return get_random_email()
        if isinstance(field, ChoiceField):
            return get_random_choice(field.supported_choices(version=self.account.version))
        if isinstance(field, CultureField):
            return get_random_choice(['da-DK', 'de-DE', 'en-US', 'es-ES', 'fr-CA', 'nl-NL', 'ru-RU', 'sv-SE'])
        if isinstance(field, BodyField):
            return get_random_string(400)
        if isinstance(field, CharListField):
            return [get_random_string(16) for _ in range(random.randint(1, 4))]
        if isinstance(field, TextListField):
            return [get_random_string(400) for _ in range(random.randint(1, 4))]
        if isinstance(field, CharField):
            return get_random_string(field.max_length)
        if isinstance(field, TextField):
            return get_random_string(400)
        if isinstance(field, MimeContentField):
            return get_random_string(400)
        if isinstance(field, Base64Field):
            return get_random_bytes(400)
        if isinstance(field, BooleanField):
            return get_random_bool()
        if isinstance(field, DecimalField):
            return get_random_decimal(field.min or 1, field.max or 99)
        if isinstance(field, IntegerField):
            return get_random_int(field.min or 0, field.max or 256)
        if isinstance(field, DateTimeField):
            return get_random_datetime(tz=self.account.default_timezone)
        if isinstance(field, AttachmentField):
            return [FileAttachment(name='my_file.txt', content=get_random_bytes(400))]
        if isinstance(field, MailboxListField):
            # email_address must be a real account on the server(?)
            # TODO: Mailbox has multiple optional args but vals must match server account, so we can't easily test
            if get_random_bool():
                return [Mailbox(email_address=self.account.primary_smtp_address)]
            else:
                return [self.account.primary_smtp_address]
        if isinstance(field, MailboxField):
            # email_address must be a real account on the server(?)
            # TODO: Mailbox has multiple optional args but vals must match server account, so we can't easily test
            if get_random_bool():
                return Mailbox(email_address=self.account.primary_smtp_address)
            else:
                return self.account.primary_smtp_address
        if isinstance(field, AttendeesField):
            # Attendee must refer to a real mailbox on the server(?). We're only sure to have one
            if get_random_bool():
                mbx = Mailbox(email_address=self.account.primary_smtp_address)
            else:
                mbx = self.account.primary_smtp_address
            with_last_response_time = get_random_bool()
            if with_last_response_time:
                return [
                    Attendee(mailbox=mbx, response_type='Accept',
                             last_response_time=get_random_datetime(tz=self.account.default_timezone))
                ]
            else:
                if get_random_bool():
                    return [Attendee(mailbox=mbx, response_type='Accept')]
                else:
                    return [self.account.primary_smtp_address]
        if isinstance(field, EmailAddressesField):
            addrs = []
            for label in EmailAddress.get_field_by_fieldname('label').supported_choices(version=self.account.version):
                addr = EmailAddress(email=get_random_email())
                addr.label = label
                addrs.append(addr)
            return addrs
        if isinstance(field, PhysicalAddressField):
            addrs = []
            for label in PhysicalAddress.get_field_by_fieldname('label')\
                    .supported_choices(version=self.account.version):
                addr = PhysicalAddress(street=get_random_string(32), city=get_random_string(32),
                                       state=get_random_string(32), country=get_random_string(32),
                                       zipcode=get_random_string(8))
                addr.label = label
                addrs.append(addr)
            return addrs
        if isinstance(field, PhoneNumberField):
            pns = []
            for label in PhoneNumber.get_field_by_fieldname('label').supported_choices(version=self.account.version):
                pn = PhoneNumber(phone_number=get_random_string(16))
                pn.label = label
                pns.append(pn)
            return pns
        if isinstance(field, EWSElementField):
            if field.value_cls == Recurrence:
                return Recurrence(pattern=DailyPattern(interval=5), start=get_random_date(), number=7)
        if isinstance(field, TimeZoneField):
            while True:
                try:
                    return EWSTimeZone.timezone(random.choice(pytz.all_timezones))
                except UnknownTimeZone:
                    pass
        if isinstance(field, PermissionSetField):
            return PermissionSet(
                permissions=[
                    Permission(
                        user_id=UserId(primary_smtp_address=self.account.primary_smtp_address),
                    )
                ]
            )
        raise ValueError('Unknown field %s' % field)


class CommonTest(EWSTest):
    def test_magic(self):
        self.assertIn(self.account.protocol.version.api_version, str(self.account.protocol))
        self.assertIn(self.account.protocol.credentials.username, str(self.account.protocol.credentials))
        self.assertIn(self.account.primary_smtp_address, str(self.account))
        self.assertIn(str(self.account.version.build.major_version), repr(self.account.version))
        for item in (
                self.account.protocol,
                self.account.version,
        ):
            # Just test that these at least don't throw errors
            repr(item)
            str(item)
        for attr in (
                'admin_audit_logs',
                'archive_deleted_items',
                'archive_inbox',
                'archive_msg_folder_root',
                'archive_recoverable_items_deletions',
                'archive_recoverable_items_purges',
                'archive_recoverable_items_root',
                'archive_recoverable_items_versions',
                'archive_root',
                'calendar',
                'conflicts',
                'contacts',
                'conversation_history',
                'directory',
                'drafts',
                'favorites',
                'im_contact_list',
                'inbox',
                'journal',
                'junk',
                'local_failures',
                'msg_folder_root',
                'my_contacts',
                'notes',
                'outbox',
                'people_connect',
                'public_folders_root',
                'quick_contacts',
                'recipient_cache',
                'recoverable_items_deletions',
                'recoverable_items_purges',
                'recoverable_items_root',
                'recoverable_items_versions',
                'search_folders',
                'sent',
                'server_failures',
                'sync_issues',
                'tasks',
                'todo_search',
                'trash',
                'voice_mail',
        ):
            # Test distinguished folder shortcuts. Some may raise ErrorAccessDenied
            try:
                item = getattr(self.account, attr)
            except (ErrorAccessDenied, ErrorFolderNotFound, ErrorItemNotFound, ErrorInvalidOperation,
                    ErrorNoPublicFolderReplicaAvailable):
                continue
            else:
                repr(item)
                str(item)
                self.assertTrue(item.is_distinguished)

    def test_from_xml(self):
        # Test for all EWSElement classes that they handle None as input to from_xml()
        import exchangelib
        for mod in (exchangelib.attachments, exchangelib.extended_properties, exchangelib.indexed_properties,
                    exchangelib.folders, exchangelib.items, exchangelib.properties):
            for k, v in vars(mod).items():
                if type(v) != type:
                    continue
                if not issubclass(v, EWSElement):
                    continue
                # from_xml() does not support None input
                with self.assertRaises(Exception):
                    v.from_xml(elem=None, account=None)


class BaseItemTest(EWSTest):
    TEST_FOLDER = None
    FOLDER_CLASS = None
    ITEM_CLASS = None

    @classmethod
    def setUpClass(cls):
        if cls is BaseItemTest:
            raise unittest.SkipTest("Skip BaseItemTest, it's only for inheritance")
        super(BaseItemTest, cls).setUpClass()

    def setUp(self):
        super(BaseItemTest, self).setUp()
        self.test_folder = getattr(self.account, self.TEST_FOLDER)
        self.assertEqual(type(self.test_folder), self.FOLDER_CLASS)
        self.assertEqual(self.test_folder.DISTINGUISHED_FOLDER_ID, self.TEST_FOLDER)
        self.test_folder.filter(categories__contains=self.categories).delete()

    def tearDown(self):
        self.test_folder.filter(categories__contains=self.categories).delete()
        # Delete all delivery receipts
        self.test_folder.filter(subject__startswith='Delivered: Subject: ').delete()
        super(BaseItemTest, self).tearDown()

    def get_random_insert_kwargs(self):
        insert_kwargs = {}
        for f in self.ITEM_CLASS.FIELDS:
            if not f.supports_version(self.account.version):
                # Cannot be used with this EWS version
                continue
            if self.ITEM_CLASS == CalendarItem and f in CalendarItem.timezone_fields():
                # Timezone fields will (and must) be populated automatically from the timestamp
                continue
            if f.is_read_only:
                # These cannot be created
                continue
            if f.name == 'mime_content':
                # This needs special formatting. See separate test_mime_content() test
                continue
            if f.name == 'attachments':
                # Testing attachments is heavy. Leave this to specific tests
                insert_kwargs[f.name] = []
                continue
            if f.name == 'resources':
                # The test server doesn't have any resources
                insert_kwargs[f.name] = []
                continue
            if f.name == 'optional_attendees':
                # 'optional_attendees' and 'required_attendees' are mutually exclusive
                insert_kwargs[f.name] = None
                continue
            if f.name == 'start':
                start = get_random_date()
                insert_kwargs[f.name], insert_kwargs['end'] = \
                    get_random_datetime_range(start_date=start, end_date=start, tz=self.account.default_timezone)
                insert_kwargs['recurrence'] = self.random_val(self.ITEM_CLASS.get_field_by_fieldname('recurrence'))
                insert_kwargs['recurrence'].boundary.start = insert_kwargs[f.name].date()
                continue
            if f.name == 'end':
                continue
            if f.name == 'recurrence':
                continue
            if f.name == 'due_date':
                # start_date must be before due_date
                insert_kwargs['start_date'], insert_kwargs[f.name] = \
                    get_random_datetime_range(tz=self.account.default_timezone)
                continue
            if f.name == 'start_date':
                continue
            if f.name == 'status':
                # Start with an incomplete task
                status = get_random_choice(set(f.supported_choices(version=self.account.version)) - {Task.COMPLETED})
                insert_kwargs[f.name] = status
                if status == Task.NOT_STARTED:
                    insert_kwargs['percent_complete'] = Decimal(0)
                else:
                    insert_kwargs['percent_complete'] = get_random_decimal(1, 99)
                continue
            if f.name == 'percent_complete':
                continue
            insert_kwargs[f.name] = self.random_val(f)
        return insert_kwargs

    def get_random_update_kwargs(self, item, insert_kwargs):
        update_kwargs = {}
        now = UTC_NOW()
        for f in self.ITEM_CLASS.FIELDS:
            if not f.supports_version(self.account.version):
                # Cannot be used with this EWS version
                continue
            if self.ITEM_CLASS == CalendarItem and f in CalendarItem.timezone_fields():
                # Timezone fields will (and must) be populated automatically from the timestamp
                continue
            if f.is_read_only:
                # These cannot be changed
                continue
            if not item.is_draft and f.is_read_only_after_send:
                # These cannot be changed when the item is no longer a draft
                continue
            if f.name == 'message_id' and f.is_read_only_after_send:
                # Cannot be updated, regardless of draft status
                continue
            if f.name == 'attachments':
                # Testing attachments is heavy. Leave this to specific tests
                update_kwargs[f.name] = []
                continue
            if f.name == 'resources':
                # The test server doesn't have any resources
                update_kwargs[f.name] = []
                continue
            if isinstance(f, AttachmentField):
                # Attachments are handled separately
                continue
            if f.name == 'start':
                start = get_random_date(start_date=insert_kwargs['end'].date())
                update_kwargs[f.name], update_kwargs['end'] = \
                    get_random_datetime_range(start_date=start, end_date=start, tz=self.account.default_timezone)
                update_kwargs['recurrence'] = self.random_val(self.ITEM_CLASS.get_field_by_fieldname('recurrence'))
                update_kwargs['recurrence'].boundary.start = update_kwargs[f.name].date()
                continue
            if f.name == 'end':
                continue
            if f.name == 'recurrence':
                continue
            if f.name == 'due_date':
                # start_date must be before due_date, and before complete_date which must be in the past
                update_kwargs['start_date'], update_kwargs[f.name] = \
                    get_random_datetime_range(end_date=now.date(), tz=self.account.default_timezone)
                continue
            if f.name == 'start_date':
                continue
            if f.name == 'status':
                # Update task to a completed state. complete_date must be a date in the past, and < than start_date
                update_kwargs[f.name] = Task.COMPLETED
                update_kwargs['percent_complete'] = Decimal(100)
                continue
            if f.name == 'percent_complete':
                continue
            if f.name == 'reminder_is_set':
                if self.ITEM_CLASS == Task:
                    # Task type doesn't allow updating 'reminder_is_set' to True
                    update_kwargs[f.name] = False
                else:
                    update_kwargs[f.name] = not insert_kwargs[f.name]
                continue
            if isinstance(f, BooleanField):
                update_kwargs[f.name] = not insert_kwargs[f.name]
                continue
            if f.value_cls in (Mailbox, Attendee):
                if insert_kwargs[f.name] is None:
                    update_kwargs[f.name] = self.random_val(f)
                else:
                    update_kwargs[f.name] = None
                continue
            update_kwargs[f.name] = self.random_val(f)
        if update_kwargs.get('is_all_day', False):
            # For is_all_day items, EWS will remove the time part of start and end values
            update_kwargs['start'] = update_kwargs['start'].replace(hour=0, minute=0, second=0, microsecond=0)
            update_kwargs['end'] = \
                update_kwargs['end'].replace(hour=0, minute=0, second=0, microsecond=0) + datetime.timedelta(days=1)
        if self.ITEM_CLASS == CalendarItem:
            # EWS always sets due date to 'start'
            update_kwargs['reminder_due_by'] = update_kwargs['start']
        return update_kwargs

    def get_test_item(self, folder=None, categories=None):
        item_kwargs = self.get_random_insert_kwargs()
        item_kwargs['categories'] = categories or self.categories
        return self.ITEM_CLASS(folder=folder or self.test_folder, **item_kwargs)


def get_random_bool():
    return bool(random.randint(0, 1))


def get_random_int(min_val=0, max_val=2147483647):
    return random.randint(min_val, max_val)


def get_random_decimal(min_val=0, max_val=100):
    precision = 2
    val = get_random_int(min_val, max_val * 10**precision) / 10.0**precision
    return Decimal('{:.2f}'.format(val))


def get_random_choice(choices):
    return random.sample(choices, 1)[0]


def get_random_string(length, spaces=True, special=True):
    chars = string.ascii_letters + string.digits
    if special:
        chars += ':.-_'
    if spaces:
        chars += ' '
    # We want random strings that don't end in spaces - Exchange strips these
    res = ''.join(map(lambda i: random.choice(chars), range(length))).strip()
    if len(res) < length:
        # If strip() made the string shorter, make sure to fill it up
        res += get_random_string(length - len(res), spaces=False)
    return res


def get_random_bytes(*args, **kwargs):
    return get_random_string(*args, **kwargs).encode('utf-8')


def get_random_url():
    path_len = random.randint(1, 16)
    domain_len = random.randint(1, 30)
    tld_len = random.randint(2, 4)
    return 'http://%s.%s/%s.html' % tuple(map(
        lambda i: get_random_string(i, spaces=False, special=False).lower(),
        (domain_len, tld_len, path_len)
    ))


def get_random_email():
    account_len = random.randint(1, 6)
    domain_len = random.randint(1, 30)
    tld_len = random.randint(2, 4)
    return '%s@%s.%s' % tuple(map(
        lambda i: get_random_string(i, spaces=False, special=False).lower(),
        (account_len, domain_len, tld_len)
    ))


# The timezone we're testing (CET/CEST) had a DST date change in 1996 (see
# https://en.wikipedia.org/wiki/Summer_Time_in_Europe). The Microsoft timezone definition on the server
# does not observe that, but pytz does. So random datetimes before 1996 will fail tests randomly.

def get_random_date(start_date=EWSDate(1996, 1, 1), end_date=EWSDate(2030, 1, 1)):
    # Keep with a reasonable date range. A wider date range is unstable WRT timezones
    return EWSDate.fromordinal(random.randint(start_date.toordinal(), end_date.toordinal()))


def get_random_datetime(start_date=EWSDate(1996, 1, 1), end_date=EWSDate(2030, 1, 1), tz=UTC):
    # Create a random datetime with minute precision. Both dates are inclusive.
    # Keep with a reasonable date range. A wider date range than the default values is unstable WRT timezones.
    while True:
        try:
            random_date = get_random_date(start_date=start_date, end_date=end_date)
            random_datetime = datetime.datetime.combine(random_date, datetime.time.min) \
                + datetime.timedelta(minutes=random.randint(0, 60 * 24))
            return tz.localize(EWSDateTime.from_datetime(random_datetime), is_dst=None)
        except (AmbiguousTimeError, NonExistentTimeError):
            pass


def get_random_datetime_range(start_date=EWSDate(1996, 1, 1), end_date=EWSDate(2030, 1, 1), tz=UTC):
    # Create two random datetimes.  Both dates are inclusive.
    # Keep with a reasonable date range. A wider date range than the default values is unstable WRT timezones.
    # Calendar items raise ErrorCalendarDurationIsTooLong if duration is > 5 years.
    return sorted([
        get_random_datetime(start_date=start_date, end_date=end_date, tz=tz),
        get_random_datetime(start_date=start_date, end_date=end_date, tz=tz),
    ])


if __name__ == '__main__':
    import sys

    if '-q' in sys.argv:
        sys.argv.remove('-q')
        logging.basicConfig(level=logging.CRITICAL)
        verbosity = 0
    else:
        logging.basicConfig(level=logging.DEBUG, handlers=[PrettyXmlHandler()])
        verbosity = 1

    unittest.main(verbosity=verbosity)
else:
    # Don't print warnings and stack traces mixed with test progress. We'll get the debug info for test failures later.
    logging.basicConfig(level=logging.CRITICAL)
