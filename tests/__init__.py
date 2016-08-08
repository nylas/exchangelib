import os
import unittest
from datetime import timedelta, datetime, date, time
import random
import string
from decimal import Decimal

from yaml import load

from exchangelib.account import Account
from exchangelib.autodiscover import discover
from exchangelib.configuration import Configuration
from exchangelib.credentials import DELEGATE, Credentials
from exchangelib.ewsdatetime import EWSDateTime, EWSDate, EWSTimeZone, UTC, UTC_NOW
from exchangelib.folders import CalendarItem, Attendee, Mailbox, Message, ExternId, Choice, Email, Contact, Task, \
    EmailAddress, PhysicalAddress, PhoneNumber, IndexedField, RoomList
from exchangelib.restriction import Restriction
from exchangelib.services import GetServerTimeZones, GetRoomLists, GetRooms, AllProperties, IdOnly
from exchangelib.util import xml_to_str, chunkify, peek
from exchangelib.version import Build


class BuildTest(unittest.TestCase):
    def test_magic(self):
        with self.assertRaises(ValueError):
            Build(7, 0)
        self.assertEqual(str(Build(9, 8, 7, 6)), '9.8.7.6')

    def test_compare(self):
        self.assertEqual(Build(15, 0, 1, 2), Build(15, 0, 1, 2))
        self.assertLess(Build(15, 0, 1, 2), Build(15, 0, 1, 3))
        self.assertLess(Build(15, 0, 1, 2), Build(15, 0, 2, 2))
        self.assertLess(Build(15, 0, 1, 2), Build(15, 1, 1, 2))
        self.assertLess(Build(15, 0, 1, 2), Build(16, 0, 1, 2))
        self.assertLessEqual(Build(15, 0, 1, 2), Build(15, 0, 1, 2))
        self.assertGreater(Build(15, 0, 1, 2), Build(15, 0, 1, 1))
        self.assertGreater(Build(15, 0, 1, 2), Build(15, 0, 0, 2))
        self.assertGreater(Build(15, 1, 1, 2), Build(15, 0, 1, 2))
        self.assertGreater(Build(15, 0, 1, 2), Build(14, 0, 1, 2))
        self.assertGreaterEqual(Build(15, 0, 1, 2), Build(15, 0, 1, 2))

    def test_api_version(self):
        self.assertEqual(Build(8, 0).api_version(), 'Exchange2007')
        self.assertEqual(Build(8, 1).api_version(), 'Exchange2007_SP1')
        self.assertEqual(Build(8, 2).api_version(), 'Exchange2007_SP1')
        self.assertEqual(Build(8, 3).api_version(), 'Exchange2007_SP1')
        self.assertEqual(Build(15, 0, 1, 1).api_version(), 'Exchange2013')
        self.assertEqual(Build(15, 0, 1, 1).api_version(), 'Exchange2013')
        self.assertEqual(Build(15, 0, 847, 0).api_version(), 'Exchange2013_SP1')
        with self.assertRaises(KeyError):
            Build(16, 0).api_version()
        with self.assertRaises(KeyError):
            Build(15, 4).api_version()


class CredentialsTest(unittest.TestCase):
    def test_hash(self):
        # Test that we can use credentials as a dict key
        self.assertEqual(hash(Credentials('a', 'b')), hash(Credentials('a', 'b')))
        self.assertNotEqual(hash(Credentials('a', 'b')), hash(Credentials('a', 'a')))
        self.assertNotEqual(hash(Credentials('a', 'b')), hash(Credentials('b', 'b')))

    def test_equality(self):
        self.assertEqual(Credentials('a', 'b'), Credentials('a', 'b'))
        self.assertNotEqual(Credentials('a', 'b'), Credentials('a', 'a'))
        self.assertNotEqual(Credentials('a', 'b'), Credentials('b', 'b'))


class EWSDateTest(unittest.TestCase):
    def test_ewsdatetime(self):
        tz = EWSTimeZone.timezone('Europe/Copenhagen')
        self.assertIsInstance(tz, EWSTimeZone)
        self.assertEqual(tz.ms_id, 'Romance Standard Time')
        self.assertEqual(tz.ms_name, '(UTC+01:00) Brussels, Copenhagen, Madrid, Paris')

        dt = tz.localize(EWSDateTime(2000, 1, 2, 3, 4, 5))
        self.assertIsInstance(dt, EWSDateTime)
        self.assertIsInstance(dt.tzinfo, EWSTimeZone)
        self.assertEqual(dt.tzinfo.ms_id, tz.ms_id)
        self.assertEqual(dt.tzinfo.ms_name, tz.ms_name)
        self.assertEqual(str(dt), '2000-01-02 03:04:05+01:00')
        self.assertEqual(
            repr(dt),
            "EWSDateTime(2000, 1, 2, 3, 4, 5, tzinfo=<DstTzInfo 'Europe/Copenhagen' CET+1:00:00 STD>)"
        )
        self.assertIsInstance(dt + timedelta(days=1), EWSDateTime)
        self.assertIsInstance(dt - timedelta(days=1), EWSDateTime)
        self.assertIsInstance(dt - EWSDateTime.now(tz=tz), timedelta)
        self.assertIsInstance(EWSDateTime.now(tz=tz), EWSDateTime)
        self.assertEqual(dt, EWSDateTime.from_datetime(tz.localize(datetime(2000, 1, 2, 3, 4, 5))))
        self.assertEqual(dt.ewsformat(), '2000-01-02T03:04:05')
        utc_tz = EWSTimeZone.timezone('UTC')
        self.assertEqual(dt.astimezone(utc_tz).ewsformat(), '2000-01-02T02:04:05Z')
        # Test summertime
        dt = tz.localize(EWSDateTime(2000, 8, 2, 3, 4, 5))
        self.assertEqual(dt.astimezone(utc_tz).ewsformat(), '2000-08-02T01:04:05Z')


class RestrictionTest(unittest.TestCase):
    def test_parse(self):
        xml = Restriction.parse_source(
            "calendar:Start > '2016-01-15T13:45:56Z' and (not calendar:Subject == 'EWS Test')"
        )
        result = '''\
<m:Restriction>
    <t:And>
        <t:IsGreaterThan>
            <t:FieldURI FieldURI="calendar:Start" />
            <t:FieldURIOrConstant>
                <t:Constant Value="2016-01-15T13:45:56Z" />
            </t:FieldURIOrConstant>
        </t:IsGreaterThan>
        <t:Not>
            <t:IsEqualTo>
                <t:FieldURI FieldURI="calendar:Subject" />
                <t:FieldURIOrConstant>
                    <t:Constant Value="EWS Test" />
                </t:FieldURIOrConstant>
            </t:IsEqualTo>
        </t:Not>
    </t:And>
</m:Restriction>'''
        self.assertEqual(xml_to_str(xml), ''.join(l.lstrip() for l in result.split('\n')))

    def test_from_params(self):
        tz = EWSTimeZone.timezone('Europe/Copenhagen')
        start = tz.localize(EWSDateTime(1900, 9, 26, 8, 0, 0))
        end = tz.localize(EWSDateTime(2200, 9, 26, 11, 0, 0))
        xml = Restriction.from_params(start=start, end=end, categories=['FOO', 'BAR'])
        result = '''\
<m:Restriction>
    <t:And>
        <t:IsGreaterThan>
            <t:FieldURI FieldURI="calendar:End" />
            <t:FieldURIOrConstant>
                <t:Constant Value="1900-09-26T07:10:00Z" />
            </t:FieldURIOrConstant>
        </t:IsGreaterThan>
        <t:IsLessThan>
            <t:FieldURI FieldURI="calendar:Start" />
            <t:FieldURIOrConstant>
                <t:Constant Value="2200-09-26T10:00:00Z" />
            </t:FieldURIOrConstant>
        </t:IsLessThan>
        <t:Or>
            <t:Contains ContainmentComparison="Exact" ContainmentMode="Substring">
                <t:FieldURI FieldURI="item:Categories" />
                <t:Constant Value="FOO" />
            </t:Contains>
            <t:Contains ContainmentComparison="Exact" ContainmentMode="Substring">
                <t:FieldURI FieldURI="item:Categories" />
                <t:Constant Value="BAR" />
            </t:Contains>
        </t:Or>
    </t:And>
</m:Restriction>'''
        self.assertEqual(str(xml), ''.join(l.lstrip() for l in result.split('\n')))


class UtilTest(unittest.TestCase):
    def test_chunkify(self):
        # Test list, tuple, set, range, map and generator
        seq = [1, 2, 3, 4, 5]
        self.assertEqual(list(chunkify(seq, chunksize=2)), [[1, 2], [3, 4], [5]])

        seq = (1, 2, 3, 4, 6, 7, 9)
        self.assertEqual(list(chunkify(seq, chunksize=3)), [(1, 2, 3), (4, 6, 7), (9,)])

        seq = {1, 2, 3, 4, 5}
        self.assertEqual(list(chunkify(seq, chunksize=2)), [[1, 2], [3, 4], [5,]])

        seq = range(5)
        self.assertEqual(list(chunkify(seq, chunksize=2)), [range(0, 2), range(2, 4), range(4, 5)])

        seq = map(int, range(5))
        self.assertEqual(list(chunkify(seq, chunksize=2)), [[0, 1], [2, 3], [4]])

        seq = (i for i in range(5))
        self.assertEqual(list(chunkify(seq, chunksize=2)), [[0, 1], [2, 3], [4]])

    def test_peek(self):
        # Test peeking into various sequence types

        # tuple
        is_empty, seq = peek(tuple())
        self.assertEqual((is_empty, list(seq)), (True, []))
        is_empty, seq = peek((1, 2, 3))
        self.assertEqual((is_empty, list(seq)), (False, [1, 2, 3]))

        # list
        is_empty, seq = peek([])
        self.assertEqual((is_empty, list(seq)), (True, []))
        is_empty, seq = peek([1, 2, 3])
        self.assertEqual((is_empty, list(seq)), (False, [1, 2, 3]))

        # set
        is_empty, seq = peek(set())
        self.assertEqual((is_empty, list(seq)), (True, []))
        is_empty, seq = peek({1, 2, 3})
        self.assertEqual((is_empty, list(seq)), (False, [1, 2, 3]))

        # range
        is_empty, seq = peek(range(0))
        self.assertEqual((is_empty, list(seq)), (True, []))
        is_empty, seq = peek(range(1, 4))
        self.assertEqual((is_empty, list(seq)), (False, [1, 2, 3]))

        # map
        is_empty, seq = peek(map(int, []))
        self.assertEqual((is_empty, list(seq)), (True, []))
        is_empty, seq = peek(map(int, [1, 2, 3]))
        self.assertEqual((is_empty, list(seq)), (False, [1, 2, 3]))

        # generator
        is_empty, seq = peek((i for i in []))
        self.assertEqual((is_empty, list(seq)), (True, []))
        is_empty, seq = peek((i for i in [1, 2, 3]))
        self.assertEqual((is_empty, list(seq)), (False, [1, 2, 3]))


class EWSTest(unittest.TestCase):
    def setUp(self):
        # There's no official Exchange server we can test against, and we can't really provide credentials for our
        # own test server to anyone on the Internet. You need to create your own settings.yml with credentials for
        # your own test server. 'settings.yml.sample' is provided as a template.
        try:
            with open(os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings.yml')) as f:
                settings = load(f)
        except FileNotFoundError:
            print('Copy settings.yml.sample to settings.yml and enter values for your test server')
            raise
        self.tz = EWSTimeZone.timezone('Europe/Copenhagen')
        self.categories = ['Test']
        self.config = Configuration(server=settings['server'], username=settings['username'],
                                    password=settings['password'], verify_ssl=settings['verify_ssl'])
        self.account = Account(primary_smtp_address=settings['account'], access_type=DELEGATE, config=self.config)
        self.maxDiff = None

    def random_val(self, field_type):
        if field_type == ExternId:
            return get_random_string(255)
        if field_type == str:
            return get_random_string(255)
        if field_type == [str]:
            return [get_random_string(16) for _ in range(random.randint(1, 4))]
        if field_type == int:
            return get_random_int(0, 256)
        if field_type == Decimal:
            return get_random_decimal(0, 100)
        if field_type == bool:
            return get_random_bool()
        if field_type == EWSDateTime:
            return get_random_datetime()
        if field_type == Email:
            return get_random_email()
        if field_type == Mailbox:
            # email_address must be a real account on the server(?)
            # TODO: Mailbox has multiple optional args, but they must match the server account, so we can't easily test.
            return Mailbox(email_address=self.account.primary_smtp_address)
        if field_type == [Mailbox]:
            # Mailbox must be a real mailbox on the server(?). We're only sure to have one
            return [self.random_val(Mailbox)]
        if field_type == Attendee:
            with_last_response_time = get_random_bool()
            if with_last_response_time:
                return Attendee(mailbox=self.random_val(Mailbox), response_type='Accept',
                                last_response_time=self.random_val(EWSDateTime))
            else:
                return Attendee(mailbox=self.random_val(Mailbox), response_type='Accept')
        if field_type == [Attendee]:
            # Attendee must refer to a real mailbox on the server(?). We're only sure to have one
            return [self.random_val(Attendee)]
        if field_type == EmailAddress:
            return EmailAddress(email=get_random_email())
        if field_type == [EmailAddress]:
            addrs = []
            for label in EmailAddress.LABELS:
                addr = self.random_val(EmailAddress)
                addr.label = label
                addrs.append(addr)
            return addrs
        if field_type == PhysicalAddress:
            return PhysicalAddress(
                street=get_random_string(32), city=get_random_string(32), state=get_random_string(32),
                country=get_random_string(32), zipcode=get_random_string(8))
        if field_type == [PhysicalAddress]:
            addrs = []
            for label in PhysicalAddress.LABELS:
                addr = self.random_val(PhysicalAddress)
                addr.label = label
                addrs.append(addr)
            return addrs
        if field_type == PhoneNumber:
            return PhoneNumber(phone_number=get_random_string(16))
        if field_type == [PhoneNumber]:
            pns = []
            for label in PhoneNumber.LABELS:
                pn = self.random_val(PhoneNumber)
                pn.label = label
                pns.append(pn)
            return pns
        assert False, 'Unknown field type %s' % field_type


class CommonTest(EWSTest):
    def test_credentials(self):
        self.assertEqual(self.account.access_type, DELEGATE)
        self.assertTrue(self.config.protocol.test())

    def test_get_timezones(self):
        ws = GetServerTimeZones(self.config.protocol)
        data = ws.call()
        self.assertAlmostEqual(len(data), 100, delta=10, msg=data)

    def test_get_roomlists(self):
        # The test server is not guaranteed to have any room lists which makes this test less useful
        ws = GetRoomLists(self.config.protocol)
        roomlists = ws.call()
        self.assertEqual(roomlists, [])

    def test_get_rooms(self):
        # The test server is not guaranteed to have any rooms or room lists which makes this test less useful
        roomlist = RoomList(email_address='my.roomlist@example.com')
        ws = GetRooms(self.config.protocol)
        roomlists = ws.call(roomlist=roomlist)
        self.assertEqual(roomlists, [])

    def test_getfolders(self):
        folders = self.account.root.get_folders()
        self.assertEqual(len(folders), 61, sorted(f.name for f in folders))

    def test_sessionpool(self):
        start = self.tz.localize(EWSDateTime(2011, 10, 12, 8))
        end = self.tz.localize(EWSDateTime(2011, 10, 12, 10))
        items = []
        for i in range(150):
            subject = 'Test Subject %s' % i
            item = CalendarItem(item_id='', changekey='', start=start, end=end, subject=subject, categories=self.categories)
            items.append(item)
        return_ids = self.account.calendar.add_items(items=items)
        self.assertEqual(len(return_ids), len(items))
        ids = self.account.calendar.find_items(start=start, end=end, categories=self.categories, shape=IdOnly)
        self.assertEqual(len(ids), len(items))
        items = self.account.calendar.get_items(return_ids)
        for i, item in enumerate(items):
            subject = 'Test Subject %s' % i
            self.assertEqual(item.start, start)
            self.assertEqual(item.end, end)
            self.assertEqual(item.subject, subject)
            self.assertEqual(item.categories, self.categories)
        status = self.account.calendar.delete_items(ids)
        self.assertEqual(set(status), {(True, None)})

    def test_magic(self):
        self.assertIn(self.config.protocol.version.api_version, str(self.config.protocol))
        self.assertIn(self.config.credentials.username, str(self.config.credentials))
        self.assertIn(self.account.primary_smtp_address, str(self.account))

    def test_autodiscover(self):
        primary_smtp_address, protocol = discover(email=self.account.primary_smtp_address,
                                                  credentials=self.config.credentials)
        self.assertEqual(primary_smtp_address, self.account.primary_smtp_address)
        self.assertEqual(protocol.service_endpoint.lower(), self.config.protocol.service_endpoint.lower())
        self.assertEqual(protocol.version.build, self.config.protocol.version.build)
        self.assertEqual(id(protocol), id(self.config.protocol), (protocol, self.config.protocol))


class BaseItemTest(EWSTest):
    TEST_FOLDER = None
    ITEM_CLASS = None

    @classmethod
    def setUpClass(cls):
        if cls is BaseItemTest:
            raise unittest.SkipTest("Skip BaseItemTest, it's only for inheritance")
        super().setUpClass()

    def setUp(self):
        super().setUp()
        self.test_folder = getattr(self.account, self.TEST_FOLDER)

    def tearDown(self):
        ids = self.test_folder.find_items(categories=self.categories, shape=IdOnly)
        self.test_folder.delete_items(ids, all_occurrences=True)

    def get_test_item(self):
        item_kwargs = {}
        for f in self.ITEM_CLASS.required_fields():
            if f == 'start':
                item_kwargs['start'], item_kwargs['end'] = get_random_datetime_range()
                continue
            if f == 'end':
                continue
            field_type = self.ITEM_CLASS.type_for_field(f)
            if field_type == Choice:
                # ITEM_CLASS.__init__ should select a default choice for us
                continue
            item_kwargs[f] = self.random_val(field_type)
        return self.ITEM_CLASS(item_id='', changekey='', categories=self.categories, **item_kwargs)

    def test_magic(self):
        item = self.get_test_item()
        if self.ITEM_CLASS == CalendarItem:
            self.assertIn('ItemId', str(item))
        else:
            self.assertIn(item.__class__.__name__, str(item))
        self.assertIn(item.__class__.__name__, repr(item))

    def test_empty_args(self):
        # We allow empty sequences for these methods
        self.assertEqual(self.test_folder.add_items(items=[]), [])
        self.assertEqual(self.test_folder.get_items(ids=[]), [])
        self.assertEqual(self.test_folder.update_items(items=[]), [])
        self.assertEqual(self.test_folder.delete_items(ids=[]), [])

    def test_finditems(self):
        item = self.get_test_item()
        self.test_folder.add_items(items=[item, item])
        items = self.test_folder.find_items(categories=self.categories, shape=AllProperties)
        for item in items:
            assert isinstance(item, self.ITEM_CLASS)
        self.assertEqual(len(items), 2)
        self.test_folder.delete_items(items, all_occurrences=True)

    def test_getitems(self):
        item = self.get_test_item()
        self.test_folder.add_items(items=[item, item])
        ids = self.test_folder.find_items(categories=self.categories, shape=IdOnly)
        items = self.test_folder.get_items(ids=ids)
        for item in items:
            assert isinstance(item, self.ITEM_CLASS)
        self.assertEqual(len(items), 2)
        self.test_folder.delete_items(items, all_occurrences=True)

    def test_extra_fields(self):
        item = self.get_test_item()
        self.test_folder.add_items(items=[item, item])
        ids = self.test_folder.find_items(categories=self.categories, shape=IdOnly)
        self.test_folder.with_extra_fields = True
        items = self.test_folder.get_items(ids=ids)
        self.test_folder.with_extra_fields = False
        for item in items:
            assert isinstance(item, self.ITEM_CLASS)
            for f in self.ITEM_CLASS.fieldnames(with_extra=True):
                self.assertTrue(hasattr(item, f))
        self.assertEqual(len(items), 2)
        self.test_folder.delete_items(items, all_occurrences=True)

    def test_item(self):
        # Test insert
        insert_kwargs = {}
        for f in self.ITEM_CLASS.fieldnames():
            if f in self.ITEM_CLASS.readonly_fields():
                # These cannot be created
                continue
            if f == 'resources':
                # The test server doesn't have any resources
                continue
            if f == 'optional_attendees':
                # 'optional_attendees' and 'required_attendees' are mutually exclusive
                insert_kwargs[f] = None
                continue
            if f == 'start':
                insert_kwargs['start'], insert_kwargs['end'] = get_random_datetime_range()
                continue
            if f == 'end':
                continue
            if f == 'due_date':
                # start_date must be before due_date
                insert_kwargs['start_date'], insert_kwargs['due_date'] = get_random_datetime_range()
                continue
            if f == 'start_date':
                continue
            if f == 'status':
                # Start with an incomplete task
                status = get_random_choice(Task.choices_for_field(f) - {Task.COMPLETED})
                insert_kwargs[f] = status
                insert_kwargs['percent_complete'] = Decimal(0) if status == Task.NOT_STARTED else get_random_decimal(0, 100)
                continue
            if f == 'percent_complete':
                continue
            field_type = self.ITEM_CLASS.type_for_field(f)
            if field_type == Choice:
                insert_kwargs[f] = get_random_choice(self.ITEM_CLASS.choices_for_field(f))
                continue
            insert_kwargs[f] = self.random_val(field_type)
        item = self.ITEM_CLASS(item_id='', changekey='', **insert_kwargs)
        # Test with generator as argument
        insert_ids = self.test_folder.add_items(items=(i for i in [item]))
        self.assertEqual(len(insert_ids), 1)
        assert isinstance(insert_ids[0], tuple)
        find_ids = self.test_folder.find_items(categories=insert_kwargs['categories'], shape=IdOnly)
        self.assertEqual(len(find_ids), 1)
        self.assertEqual(len(find_ids[0]), 2)
        self.assertEqual(insert_ids, find_ids)
        # Test with generator as argument
        item = self.test_folder.get_items(ids=(i for i in find_ids))[0]
        for f in self.ITEM_CLASS.fieldnames():
            if f in self.ITEM_CLASS.readonly_fields():
                continue
            if f == 'resources':
                continue
            if isinstance(self.ITEM_CLASS.type_for_field(f), list):
                if not (getattr(item, f) is None and insert_kwargs[f] is None):
                    self.assertSetEqual(set(getattr(item, f)), set(insert_kwargs[f]), (f, repr(item), insert_kwargs))
            else:
                self.assertEqual(getattr(item, f), insert_kwargs[f], (f, repr(item), insert_kwargs))

        # Test update
        update_kwargs = {}
        now = UTC_NOW()
        for f in self.ITEM_CLASS.fieldnames():
            if f in self.ITEM_CLASS.readonly_fields():
                # These cannot be changed
                continue
            if f == 'resources':
                # The test server doesn't have any resources
                continue
            if f == 'start':
                update_kwargs['start'], update_kwargs['end'] = get_random_datetime_range()
                continue
            if f == 'end':
                continue
            if f == 'due_date':
                # start_date must be before due_date, and before complete_date which must be in the past
                d1, d2 = get_random_datetime(end_date=now), get_random_datetime(end_date=now)
                update_kwargs['start_date'], update_kwargs['due_date'] = sorted([d1, d2])
                continue
            if f == 'start_date':
                continue
            if f == 'status':
                # Update task to a completed state. complete_date must be a date in the past, and < than start_date
                update_kwargs[f] = Task.COMPLETED
                update_kwargs['percent_complete'] = Decimal(100)
                continue
            if f == 'percent_complete':
                continue
            if f == 'reminder_is_set' and self.ITEM_CLASS == Task:
                # Task type doesn't allow updating 'reminder_is_set' to True. TODO: Really?
                update_kwargs[f] = False
                continue
            field_type = self.ITEM_CLASS.type_for_field(f)
            if field_type == bool:
                update_kwargs[f] = not(insert_kwargs[f])
                continue
            if field_type == Choice:
                update_kwargs[f] = get_random_choice(self.ITEM_CLASS.choices_for_field(f))
                continue
            if field_type in (Mailbox, [Mailbox], Attendee, [Attendee]):
                if insert_kwargs[f] is None:
                    update_kwargs[f] = self.random_val(field_type)
                else:
                    update_kwargs[f] = None
                continue
            update_kwargs[f] = self.random_val(field_type)
        # Test with generator as argument
        update_ids = self.test_folder.update_items(items=(i for i in [(item, update_kwargs), ]))
        self.assertEqual(len(update_ids), 1)
        self.assertEqual(len(update_ids[0]), 2, update_ids)
        self.assertEqual(insert_ids[0][0], update_ids[0][0])  # ID should be the same
        self.assertNotEqual(insert_ids[0][1], update_ids[0][1])  # Changekey should not be the same when item is updated
        item = self.test_folder.get_items(update_ids)[0]
        for f in self.ITEM_CLASS.fieldnames():
            if f in self.ITEM_CLASS.readonly_fields():
                continue
            if f == 'resources':
                # The test server doesn't have any resources
                continue
            field_type = self.ITEM_CLASS.type_for_field(f)
            if isinstance(field_type, list):
                if issubclass(field_type[0], IndexedField):
                    # TODO: We don't know how to update IndexedField types yet
                    continue
                if not (getattr(item, f) is None and update_kwargs[f] is None):
                    self.assertSetEqual(set(getattr(item, f)), set(update_kwargs[f]), (f, repr(item), update_kwargs))
            else:
                self.assertEqual(getattr(item, f), update_kwargs[f], (f, repr(item), update_kwargs))

        # Test wiping or removing string, int, Choice and bool fields
        wipe_kwargs = {}
        for f in self.ITEM_CLASS.fieldnames():
            if f in self.ITEM_CLASS.required_fields():
                # These cannot be deleted
                continue
            if f in self.ITEM_CLASS.readonly_fields():
                # These cannot be changed
                continue
            field_type = self.ITEM_CLASS.type_for_field(f)
            if field_type == ExternId:
                wipe_kwargs[f] = ''
            elif field_type in (bool, str, int, Choice, Email):
                wipe_kwargs[f] = None
        wipe_ids = self.test_folder.update_items([(item, wipe_kwargs), ])
        self.assertEqual(len(wipe_ids), 1)
        self.assertEqual(len(wipe_ids[0]), 2, wipe_ids)
        self.assertEqual(insert_ids[0][0], wipe_ids[0][0])  # ID should be the same
        self.assertNotEqual(insert_ids[0][1], wipe_ids[0][1])  # Changekey should not be the same when item is updated
        item = self.test_folder.get_items(wipe_ids)[0]
        for f in self.ITEM_CLASS.fieldnames():
            if f in self.ITEM_CLASS.required_fields():
                continue
            if f in self.ITEM_CLASS.readonly_fields():
                continue
            field_type = self.ITEM_CLASS.type_for_field(f)
            if field_type in (str, ExternId, bool, int, Choice, Email):
                self.assertEqual(getattr(item, f), wipe_kwargs[f], (f, repr(item), insert_kwargs))

        # Test extern_id = None, which deletes the extended property entirely
        extern_id = None
        wipe2_ids = self.test_folder.update_items([(item, {'extern_id': extern_id}), ])
        self.assertEqual(len(wipe2_ids), 1)
        self.assertEqual(len(wipe2_ids[0]), 2, wipe2_ids)
        self.assertEqual(insert_ids[0][0], wipe2_ids[0][0])  # ID should be the same
        self.assertNotEqual(insert_ids[0][1], wipe2_ids[0][1])  # Changekey should not be the same when item is updated
        item = self.test_folder.get_items(wipe2_ids)[0]
        self.assertEqual(item.extern_id, extern_id)

        # Remove test item. Test with generator as argument
        status = self.test_folder.delete_items(ids=(i for i in wipe2_ids), all_occurrences=True)
        self.assertEqual(status, [(True, None)])


class CalendarTest(BaseItemTest):
    TEST_FOLDER = 'calendar'
    ITEM_CLASS = CalendarItem


class InboxTest(BaseItemTest):
    TEST_FOLDER = 'inbox'
    ITEM_CLASS = Message


class TasksTest(BaseItemTest):
    TEST_FOLDER = 'tasks'
    ITEM_CLASS = Task


class ContactsTest(BaseItemTest):
    TEST_FOLDER = 'contacts'
    ITEM_CLASS = Contact


def get_random_bool():
    return bool(random.randint(0, 1))


def get_random_int(min=0, max=2147483647):
    return random.randint(min, max)


def get_random_decimal(min=0, max=100):
    # Return a random decimal with 6-digit precision
    major = get_random_int(min, max)
    minor = 0 if major == max else get_random_int(0, 999999)
    return Decimal('%s.%s' % (major, minor))


def get_random_choice(choices):
    return random.sample(choices, 1)[0]


def get_random_string(length, spaces=True):
    chars = string.ascii_letters + string.digits + ':.-_'
    if spaces:
        chars += ' '
    # We want random strings that don't end in spaces - Exchange strips these
    res = ''.join(map(lambda i: random.choice(chars), range(length))).strip()
    if len(res) < length:
        # If strip() made the string shorter, make sure to fill it up
        res += get_random_string(length - len(res), spaces=False)
    return res


def get_random_email():
    account_len = random.randint(1, 6)
    domain_len = random.randint(1, 30)
    tld_len = random.randint(2, 4)
    return '%s@%s.%s' % tuple(map(
        lambda i: get_random_string(i, spaces=False).lower(),
        (account_len, domain_len, tld_len)
    ))


def get_random_date(start_date=date(1900, 1, 1), end_date=date(2100, 1, 1)):
    return EWSDate.fromordinal(random.randint(start_date.toordinal(), end_date.toordinal()))


def get_random_datetime(start_date=date(1900, 1, 1), end_date=date(2100, 1, 1)):
    # Create a random datetime with minute precision
    random_date = get_random_date(start_date=start_date, end_date=end_date)
    random_datetime = datetime.combine(random_date, time.min) + timedelta(minutes=random.randint(0, 60*24))
    return UTC.localize(EWSDateTime.from_datetime(random_datetime))


def get_random_datetime_range():
    # Create two random datetimes. Calendar items raise ErrorCalendarDurationIsTooLong if duration is > 5 years.
    dt1 = get_random_datetime()
    dt2 = dt1 + timedelta(minutes=random.randint(0, 60*24*365*5))
    return dt1, dt2


if __name__ == '__main__':
    import logging
    # loglevel = logging.DEBUG
    loglevel = logging.INFO
    logging.basicConfig(level=loglevel)
    logging.getLogger('exchangelib').setLevel(loglevel)
    unittest.main()
