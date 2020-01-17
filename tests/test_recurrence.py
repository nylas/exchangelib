from exchangelib import EWSDate
from exchangelib.fields import MONDAY, FEBRUARY, AUGUST, SECOND, LAST, WEEKEND_DAY
from exchangelib.recurrence import Recurrence, AbsoluteYearlyPattern, RelativeYearlyPattern, AbsoluteMonthlyPattern, \
    RelativeMonthlyPattern, WeeklyPattern, DailyPattern, NoEndPattern, EndDatePattern, NumberedPattern

from .common import TimedTestCase


class RecurrenceTest(TimedTestCase):
    def test_magic(self):
        pattern = AbsoluteYearlyPattern(month=FEBRUARY, day_of_month=28)
        self.assertEqual(str(pattern), 'Occurs on day 28 of February')
        pattern = RelativeYearlyPattern(month=AUGUST, week_number=SECOND, weekday=MONDAY)
        self.assertEqual(str(pattern), 'Occurs on weekday Monday in the Second week of August')
        pattern = AbsoluteMonthlyPattern(interval=3, day_of_month=31)
        self.assertEqual(str(pattern), 'Occurs on day 31 of every 3 month(s)')
        pattern = RelativeMonthlyPattern(interval=2, week_number=LAST, weekday=5)
        self.assertEqual(str(pattern), 'Occurs on weekday Friday in the Last week of every 2 month(s)')
        pattern = WeeklyPattern(interval=4, weekdays=WEEKEND_DAY, first_day_of_week=7)
        self.assertEqual(str(pattern),
                         'Occurs on weekdays WeekendDay of every 4 week(s) where the first day of the week is Sunday')
        pattern = DailyPattern(interval=6)
        self.assertEqual(str(pattern), 'Occurs every 6 day(s)')

    def test_validation(self):
        p = DailyPattern(interval=3)
        d_start = EWSDate(2017, 9, 1)
        d_end = EWSDate(2017, 9, 7)
        with self.assertRaises(ValueError):
            Recurrence(pattern=p, boundary='foo', start='bar')  # Specify *either* boundary *or* start, end and number
        with self.assertRaises(ValueError):
            Recurrence(pattern=p, start='foo', end='bar', number='baz')  # number is invalid when end is present
        with self.assertRaises(ValueError):
            Recurrence(pattern=p, end='bar', number='baz')  # Must have start
        r = Recurrence(pattern=p, start=d_start)
        self.assertEqual(r.boundary, NoEndPattern(start=d_start))
        r = Recurrence(pattern=p, start=d_start, end=d_end)
        self.assertEqual(r.boundary, EndDatePattern(start=d_start, end=d_end))
        r = Recurrence(pattern=p, start=d_start, number=1)
        self.assertEqual(r.boundary, NumberedPattern(start=d_start, number=1))
