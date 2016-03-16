Exchange Web Services client library
====================================
This module provides an well-performing interface for communicating with a Microsoft Exchange 2007-2016 Server or
Office365 using Exchange Web Services (EWS). It currently only implements autodiscover and functions for manipulating
calendar items.

Usage
~~~~~

Here is a simple example that inserts, retrieves and deletes calendar items in an Exchange calendar::

    from exchangelib import DELEGATE, services
    from exchangelib.configuration import Configuration
    from exchangelib.account import Account
    from exchangelib.ewsdatetime import EWSDateTime, EWSTimeZone
    from exchangelib.folders import CalendarItem

    year, month, day = 2016, 3, 20
    tz = EWSTimeZone.timezone('Europe/Copenhagen')

    calendar_items = []
    for hour in range(7, 17):
        calendar_items.append(CalendarItem(
            start=tz.localize(EWSDateTime(year, month, day, hour, 30)),
            end=tz.localize(EWSDateTime(year, month, day, hour+1, 15)),
            subject='Test item',
            body='Hello from Python',
            location='devnull',
            categories=['foo', 'bar'],
        ))

    config = Configuration(username='MYWINDOMAIN\myusername', password='topsecret')
    account = Account(primary_smtp_address='john@example.com', config=config, autodiscover=True, access_type=DELEGATE)

    res = account.calendar.add_items(calendar_items)
    print(res)

    ids = account.calendar.find_items(
        start=tz.localize(EWSDateTime(year, month, day)),
        end=tz.localize(EWSDateTime(year, month, day+1)),
        categories=['foo', 'bar'],
        shape=services.IdOnly,
    )
    print(ids)

    res = account.calendar.delete_items(ids)
    print(res)
