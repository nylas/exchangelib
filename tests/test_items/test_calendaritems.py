import datetime

from exchangelib.errors import ErrorInvalidOperation
from exchangelib.ewsdatetime import EWSDateTime, UTC
from exchangelib.folders import Calendar
from exchangelib.items import CalendarItem

from ..common import get_random_string, get_random_datetime_range, get_random_date
from .test_basics import CommonItemTest


class CalendarTest(CommonItemTest):
    TEST_FOLDER = 'calendar'
    FOLDER_CLASS = Calendar
    ITEM_CLASS = CalendarItem

    def test_updating_timestamps(self):
        # Test that we can update an item without changing anything, and maintain the hidden timezone fields as local
        # timezones, and that returned timestamps are in UTC.
        item = self.get_test_item()
        item.reminder_is_set = True
        item.is_all_day = False
        item.save()
        for i in self.account.calendar.filter(categories__contains=self.categories).only('start', 'end', 'categories'):
            self.assertEqual(i.start, item.start)
            self.assertEqual(i.start.tzinfo, UTC)
            self.assertEqual(i.end, item.end)
            self.assertEqual(i.end.tzinfo, UTC)
            self.assertEqual(i._start_timezone, self.account.default_timezone)
            self.assertEqual(i._end_timezone, self.account.default_timezone)
            i.save(update_fields=['start', 'end'])
            self.assertEqual(i.start, item.start)
            self.assertEqual(i.start.tzinfo, UTC)
            self.assertEqual(i.end, item.end)
            self.assertEqual(i.end.tzinfo, UTC)
            self.assertEqual(i._start_timezone, self.account.default_timezone)
            self.assertEqual(i._end_timezone, self.account.default_timezone)
        for i in self.account.calendar.filter(categories__contains=self.categories).only('start', 'end', 'categories'):
            self.assertEqual(i.start, item.start)
            self.assertEqual(i.start.tzinfo, UTC)
            self.assertEqual(i.end, item.end)
            self.assertEqual(i.end.tzinfo, UTC)
            self.assertEqual(i._start_timezone, self.account.default_timezone)
            self.assertEqual(i._end_timezone, self.account.default_timezone)
            i.delete()

    def test_update_to_non_utc_datetime(self):
        # Test updating with non-UTC datetime values. This is a separate code path in UpdateItem code
        item = self.get_test_item()
        item.reminder_is_set = True
        item.is_all_day = False
        item.save()
        # Update start, end and recurrence with timezoned datetimes. For some reason, EWS throws
        # 'ErrorOccurrenceTimeSpanTooBig' is we go back in time.
        start = get_random_date(start_date=item.start.date() + datetime.timedelta(days=1))
        dt_start, dt_end = [
            dt.astimezone(self.account.default_timezone) for dt in
            get_random_datetime_range(start_date=start, end_date=start, tz=self.account.default_timezone)
        ]
        item.start, item.end = dt_start, dt_end
        item.recurrence.boundary.start = dt_start.date()
        item.save()
        item.refresh()
        self.assertEqual(item.start, dt_start)
        self.assertEqual(item.end, dt_end)

    def test_all_day_datetimes(self):
        # Test that we can use plain dates for start and end values for all-day items
        start = get_random_date()
        start_dt, end_dt = get_random_datetime_range(
            start_date=start,
            end_date=start + datetime.timedelta(days=365),
            tz=self.account.default_timezone
        )
        # Assign datetimes for start and end
        item = self.ITEM_CLASS(folder=self.test_folder, start=start_dt, end=end_dt, is_all_day=True,
                               categories=self.categories).save()

        # Returned item start and end values should be EWSDate instances
        item = self.test_folder.all().only('is_all_day', 'start', 'end').get(id=item.id, changekey=item.changekey)
        self.assertEqual(item.is_all_day, True)
        self.assertEqual(item.start, start_dt.date())
        self.assertEqual(item.end, end_dt.date())
        item.save()  # Make sure we can update
        item.delete()

        # We are also allowed to assign plain dates as values for all-day items
        item = self.ITEM_CLASS(folder=self.test_folder, start=start_dt.date(), end=end_dt.date(), is_all_day=True,
                               categories=self.categories).save()

        # Returned item start and end values should be EWSDate instances
        item = self.test_folder.all().only('is_all_day', 'start', 'end').get(id=item.id, changekey=item.changekey)
        self.assertEqual(item.is_all_day, True)
        self.assertEqual(item.start, start_dt.date())
        self.assertEqual(item.end, end_dt.date())
        item.save()  # Make sure we can update

    def test_view(self):
        item1 = self.ITEM_CLASS(
            account=self.account,
            folder=self.test_folder,
            subject=get_random_string(16),
            start=self.account.default_timezone.localize(EWSDateTime(2016, 1, 1, 8)),
            end=self.account.default_timezone.localize(EWSDateTime(2016, 1, 1, 10)),
            categories=self.categories,
        )
        item2 = self.ITEM_CLASS(
            account=self.account,
            folder=self.test_folder,
            subject=get_random_string(16),
            start=self.account.default_timezone.localize(EWSDateTime(2016, 2, 1, 8)),
            end=self.account.default_timezone.localize(EWSDateTime(2016, 2, 1, 10)),
            categories=self.categories,
        )
        self.test_folder.bulk_create(items=[item1, item2])

        # Test missing args
        with self.assertRaises(TypeError):
            self.test_folder.view()
        # Test bad args
        with self.assertRaises(ValueError):
            list(self.test_folder.view(start=item1.end, end=item1.start))
        with self.assertRaises(TypeError):
            list(self.test_folder.view(start='xxx', end=item1.end))
        with self.assertRaises(ValueError):
            list(self.test_folder.view(start=item1.start, end=item1.end, max_items=0))

        def match_cat(i):
            return set(i.categories or []) == set(self.categories)

        # Test dates
        self.assertEqual(
            len([i for i in self.test_folder.view(start=item1.start, end=item1.end) if match_cat(i)]),
            1
        )
        self.assertEqual(
            len([i for i in self.test_folder.view(start=item1.start, end=item2.end) if match_cat(i)]),
            2
        )
        # Edge cases. Get view from end of item1 to start of item2. Should logically return 0 items, but Exchange wants
        # it differently and returns item1 even though there is no overlap.
        self.assertEqual(
            len([i for i in self.test_folder.view(start=item1.end, end=item2.start) if match_cat(i)]),
            1
        )
        self.assertEqual(
            len([i for i in self.test_folder.view(start=item1.start, end=item2.start) if match_cat(i)]),
            1
        )

        # Test max_items
        self.assertEqual(
            len([i for i in self.test_folder.view(start=item1.start, end=item2.end, max_items=9999) if match_cat(i)]),
            2
        )
        self.assertEqual(
            len(self.test_folder.view(start=item1.start, end=item2.end, max_items=1)),
            1
        )

        # Test chaining
        qs = self.test_folder.view(start=item1.start, end=item2.end)
        self.assertTrue(qs.count() >= 2)
        with self.assertRaises(ErrorInvalidOperation):
            qs.filter(subject=item1.subject).count()  # EWS does not allow restrictions
        self.assertListEqual(
            [i for i in qs.order_by('subject').values('subject') if i['subject'] in (item1.subject, item2.subject)],
            [{'subject': s} for s in sorted([item1.subject, item2.subject])]
        )
