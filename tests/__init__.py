import os
import unittest
from datetime import timedelta, datetime

from yaml import load

from exchangelib.account import Account
from exchangelib.configuration import Configuration
from exchangelib.credentials import DELEGATE
from exchangelib.ewsdatetime import EWSDateTime, EWSTimeZone
from exchangelib.folders import CalendarItem, MailBox
from exchangelib.services import GetServerTimeZones, AllProperties, IdOnly


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


class EWSTest(unittest.TestCase):
    def setUp(self):
        self.tz = EWSTimeZone.timezone('Europe/Copenhagen')
        try:
            with open(os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings.yml')) as f:
                settings = load(f)
        except FileNotFoundError:
            print('Copy settings.yml.sample to settings.yml and enter values for your test server')
            raise
        self.categories = ['Test']
        self.config = Configuration(server=settings['server'], username=settings['username'],
                                    password=settings['password'])
        self.account = Account(primary_smtp_address=settings['account'], access_type=DELEGATE, config=self.config)

    def tearDown(self):
        start = self.tz.localize(EWSDateTime(1900, 9, 26, 8, 0, 0))
        end = self.tz.localize(EWSDateTime(2200, 9, 26, 11, 0, 0))
        ids = self.account.calendar.find_items(start=start, end=end, categories=self.categories, shape=IdOnly)
        if ids:
            self.account.calendar.delete_items(ids)

    def test_credentials(self):
        self.assertEqual(self.account.access_type, DELEGATE)
        self.assertTrue(self.config.protocol.test())

    def test_get_timezones(self):
        ws = GetServerTimeZones(self.config.protocol)
        data = ws.call()
        self.assertAlmostEqual(len(data), 100, delta=10, msg=data)

    def test_getfolders(self):
        folders = self.account.root.get_folders()
        self.assertEqual(len(folders), 61, sorted(f.name for f in folders))

    def test_finditems(self):
        start = self.tz.localize(EWSDateTime(2009, 9, 26, 8, 0, 0))
        end = self.tz.localize(EWSDateTime(2009, 9, 26, 11, 0, 0))
        subject = 'Test Subject'
        body = 'Test Body'
        location = 'Test Location'
        item = CalendarItem(item_id='', changekey='', start=start, end=end, subject=subject, body=body,
                            location=location, reminder_is_set=False, categories=self.categories)
        self.account.calendar.add_items(items=[item, item])
        items = self.account.calendar.find_items(start=start, end=end, categories=self.categories, shape=AllProperties)
        for item in items:
            assert isinstance(item, CalendarItem)
        self.assertEqual(len(items), 2)
        self.account.calendar.delete_items(items)

    def test_getitems(self):
        start = self.tz.localize(EWSDateTime(2009, 9, 26, 8, 0, 0))
        end = self.tz.localize(EWSDateTime(2009, 9, 26, 11, 0, 0))
        subject = 'Test Subject'
        body = 'Test Body'
        location = 'Test Location'
        item = CalendarItem(item_id='', changekey='', start=start, end=end, subject=subject, body=body,
                            location=location, reminder_is_set=False, categories=self.categories)
        self.account.calendar.add_items(items=[item, item])
        ids = self.account.calendar.find_items(start=start, end=end, categories=self.categories, shape=IdOnly)
        items = self.account.calendar.get_items(ids=ids)
        for item in items:
            assert isinstance(item, CalendarItem)
        self.assertEqual(len(items), 2)
        self.account.calendar.delete_items(items)

    def test_extra_fields(self):
        start = self.tz.localize(EWSDateTime(2009, 9, 26, 8, 0, 0))
        end = self.tz.localize(EWSDateTime(2009, 9, 26, 11, 0, 0))
        subject = 'Test Subject'
        body = 'Test Body'
        location = 'Test Location'
        item = CalendarItem(item_id='', changekey='', start=start, end=end, subject=subject, body=body,
                            location=location, reminder_is_set=False, categories=self.categories)
        self.account.calendar.add_items(items=[item, item])
        ids = self.account.calendar.find_items(start=start, end=end, categories=self.categories, shape=IdOnly)
        self.account.calendar.with_extra_fields = True
        items = self.account.calendar.get_items(ids=ids)
        self.account.calendar.with_extra_fields = False
        for item in items:
            assert isinstance(item, CalendarItem)
            for f in CalendarItem.fieldnames(with_extra=True):
                self.assertTrue(hasattr(item, f))
        self.assertEqual(len(items), 2)
        self.account.calendar.delete_items(items)

    def test_item(self):
        # Test insert
        start = self.tz.localize(EWSDateTime(2011, 10, 12, 8))
        end = self.tz.localize(EWSDateTime(2011, 10, 12, 10))
        subject = 'Test Subject'
        body = 'Test Body'
        location = 'Test Location'
        extern_id = '123'
        reminder_is_set = False
        item = CalendarItem(item_id='', changekey='', start=start, end=end, subject=subject, body=body,
                            location=location, reminder_is_set=reminder_is_set, categories=self.categories,
                            extern_id=extern_id)
        return_ids = self.account.calendar.add_items(items=[item])
        self.assertEqual(len(return_ids), 1)
        for item_id in return_ids:
            assert isinstance(item_id, tuple)
        ids = self.account.calendar.find_items(start=start, end=end, categories=self.categories, shape=IdOnly)
        self.assertEqual(len(ids[0]), 2)
        self.assertEqual(len(ids), 1)
        self.assertEqual(return_ids, ids)
        item = self.account.calendar.get_items(ids)[0]
        self.assertEqual(item.start, start)
        self.assertEqual(item.end, end)
        self.assertEqual(item.subject, subject)
        self.assertEqual(item.location, location)
        self.assertEqual(item.body, body)
        self.assertEqual(item.categories, self.categories)
        self.assertEqual(item.extern_id, extern_id)
        self.assertEqual(item.organizer.email_address, self.account.primary_smtp_address)
        self.assertEqual(item.organizer.mailbox_type, 'Mailbox')
        self.assertEqual(item.organizer.item_id, None)
        self.assertEqual(item.reminder_is_set, reminder_is_set)

        # Test update
        start = self.tz.localize(EWSDateTime(2012, 9, 12, 16))
        end = self.tz.localize(EWSDateTime(2012, 9, 12, 17))
        subject = 'New Subject'
        body = 'New Body'
        location = 'New Location'
        categories = ['a', 'b']
        extern_id = '456'
        reminder_is_set = True
        ids = self.account.calendar.update_items(
            [
                (item, {'start': start, 'end': end, 'subject': subject, 'body': body, 'location': location,
                        'categories': categories, 'extern_id': extern_id, 'reminder_is_set': reminder_is_set}),
            ]
        )
        self.assertEqual(len(ids[0]), 2, ids)
        self.assertEqual(len(ids), 1)
        self.assertEqual(return_ids[0][0], ids[0][0])  # ID should be the same
        self.assertNotEqual(return_ids[0][1], ids[0][1])  # Changekey should not be the same when item is updated
        item = self.account.calendar.get_items(ids)[0]
        self.assertEqual(item.start, start)
        self.assertEqual(item.end, end)
        self.assertEqual(item.subject, subject)
        self.assertEqual(item.location, location)
        self.assertEqual(item.body, body)
        self.assertEqual(item.categories, categories)
        self.assertEqual(item.extern_id, extern_id)
        self.assertEqual(item.organizer.email_address, self.account.primary_smtp_address)
        self.assertEqual(item.organizer.mailbox_type, 'Mailbox')
        self.assertEqual(item.organizer.item_id, None)
        self.assertEqual(item.reminder_is_set, reminder_is_set)

        # Test wiping fields
        subject = ''
        body = ''
        location = ''
        extern_id = None
        # reminder_is_set = None  # reminder_is_set cannot be deleted
        ids = self.account.calendar.update_items(
            [
                (item, {'subject': subject, 'body': body, 'location': location, 'extern_id': extern_id}),
            ]
        )
        self.assertEqual(len(ids[0]), 2, ids)
        self.assertEqual(len(ids), 1)
        self.assertEqual(return_ids[0][0], ids[0][0])  # ID should be the same
        self.assertNotEqual(return_ids[0][1], ids[0][1])  # Changekey should not be the same when item is updated
        item = self.account.calendar.get_items(ids)[0]
        self.assertEqual(item.subject, subject)
        self.assertEqual(item.location, location)
        self.assertEqual(item.body, body)
        self.assertEqual(item.extern_id, extern_id)

        # Test extern_id = None vs extern_id = ''
        extern_id = ''
        # reminder_is_set = None  # reminder_is_set cannot be deleted
        ids = self.account.calendar.update_items(
            [
                (item, {'extern_id': extern_id}),
            ]
        )
        self.assertEqual(len(ids[0]), 2, ids)
        self.assertEqual(len(ids), 1)
        self.assertEqual(return_ids[0][0], ids[0][0])  # ID should be the same
        self.assertNotEqual(return_ids[0][1], ids[0][1])  # Changekey should not be the same when item is updated
        item = self.account.calendar.get_items(ids)[0]
        self.assertEqual(item.extern_id, extern_id)

        # Remove test item
        status = self.account.calendar.delete_items(ids)
        self.assertEqual(status, [(True, None)])

    def test_sessionpool(self):
        start = self.tz.localize(EWSDateTime(2011, 10, 12, 8))
        end = self.tz.localize(EWSDateTime(2011, 10, 12, 10))
        items = []
        for i in range(150):
            subject = 'Test Subject %s' % i
            body = 'Test Body %s' % i
            location = 'Test Location %s' % i
            item = CalendarItem(item_id='', changekey='', start=start, end=end, subject=subject, body=body,
                                location=location, reminder_is_set=False, categories=self.categories)
            items.append(item)
        return_ids = self.account.calendar.add_items(items=items)
        self.assertEqual(len(return_ids), len(items))
        ids = self.account.calendar.find_items(start=start, end=end, categories=self.categories, shape=IdOnly)
        self.assertEqual(len(ids), len(items))
        items = self.account.calendar.get_items(return_ids)
        for i, item in enumerate(items):
            subject = 'Test Subject %s' % i
            body = 'Test Body %s' % i
            location = 'Test Location %s' % i
            self.assertEqual(item.start, start)
            self.assertEqual(item.end, end)
            self.assertEqual(item.subject, subject)
            self.assertEqual(item.location, location)
            self.assertEqual(item.body, body)
            self.assertEqual(item.categories, self.categories)
        status = self.account.calendar.delete_items(ids)
        self.assertEqual(set(status), {(True, None)})


if __name__ == '__main__':
    import logging
    # loglevel = logging.DEBUG
    loglevel = logging.INFO
    logging.basicConfig(level=loglevel)
    logging.getLogger('exchangelib').setLevel(loglevel)
    unittest.main()
