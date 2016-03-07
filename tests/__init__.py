import unittest

from pytz import timezone
from yaml import load

from exchangelib.account import Account
from exchangelib.configuration import Configuration
from exchangelib.credentials import DELEGATE
from exchangelib.ewsdatetime import EWSDateTime
from exchangelib.folders import CalendarItem
from exchangelib.services import GetServerTimeZones, AllProperties, IdOnly


class EWSTest(unittest.TestCase):
    def setUp(self):
        self.tzname = 'Europe/Copenhagen'
        try:
            with open('settings.yml') as f:
                settings = load(f)
        except FileNotFoundError:
            print('Copy settings.yml.sample to settings.yml and enter values for your test server')
            raise
        self.categories = ['Test']
        self.config = Configuration(server=settings['server'], username=settings['username'],
                                    password=settings['password'], timezone=self.tzname)
        self.account = Account(primary_smtp_address=settings['account'], access_type=DELEGATE, config=self.config)

    def tearDown(self):
        tz = timezone(self.tzname)
        start = EWSDateTime(1900, 9, 26, 8, 0, 0, tzinfo=tz)
        end = EWSDateTime(2200, 9, 26, 11, 0, 0, tzinfo=tz)
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
        tz = timezone(self.tzname)
        start = EWSDateTime(2009, 9, 26, 8, 0, 0, tzinfo=tz)
        end = EWSDateTime(2009, 9, 26, 11, 0, 0, tzinfo=tz)
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
        tz = timezone(self.tzname)
        start = EWSDateTime(2009, 9, 26, 8, 0, 0, tzinfo=tz)
        end = EWSDateTime(2009, 9, 26, 11, 0, 0, tzinfo=tz)
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
        tz = timezone(self.tzname)
        start = EWSDateTime(2009, 9, 26, 8, 0, 0, tzinfo=tz)
        end = EWSDateTime(2009, 9, 26, 11, 0, 0, tzinfo=tz)
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
        tz = timezone(self.tzname)

        # Test insert
        start = EWSDateTime(2011, 10, 12, 8, tzinfo=tz)
        end = EWSDateTime(2011, 10, 12, 10, tzinfo=tz)
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
        self.assertEqual(item.reminder_is_set, reminder_is_set)

        # Test update
        start = EWSDateTime(2012, 9, 12, 16, tzinfo=tz)
        end = EWSDateTime(2012, 9, 12, 17, tzinfo=tz)
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
        tz = timezone(self.tzname)
        start = EWSDateTime(2011, 10, 12, 8, tzinfo=tz)
        end = EWSDateTime(2011, 10, 12, 10, tzinfo=tz)
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
