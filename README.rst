Exchange Web Services client library
====================================
This module provides an well-performing, well-behaving and simple interface for communicating with a Microsoft Exchange
2007-2016 Server or Office365 using Exchange Web Services (EWS). It currently implements autodiscover, and functions for
searching, creating, updating and deleting calendar, mailbox, task and contact items.


.. image:: https://badge.fury.io/py/exchangelib.svg
    :target: https://badge.fury.io/py/exchangelib

.. image:: https://landscape.io/github/ecederstrand/exchangelib/master/landscape.png
   :target: https://landscape.io/github/ecederstrand/exchangelib/master

.. image:: https://secure.travis-ci.org/ecederstrand/exchangelib.png
    :target: http://travis-ci.org/ecederstrand/exchangelib

.. image:: https://coveralls.io/repos/github/ecederstrand/exchangelib/badge.svg?branch=
    :target: https://coveralls.io/github/ecederstrand/exchangelib?branch=


Usage
~~~~~

Here is a simple example that inserts, retrieves and deletes calendar items in an Exchange calendar:

.. code-block:: python

    from exchangelib import DELEGATE, IMPERSONATION, IdOnly, Account, Credentials, \
        EWSDateTime, EWSTimeZone, Configuration, NTLM
    from exchangelib.folders import Calendar, CalendarItem

    year, month, day = 2016, 3, 20
    tz = EWSTimeZone.timezone('Europe/Copenhagen')

    # Build a list of calendar items
    calendar_items = []
    for hour in range(7, 17):
        calendar_items.append(CalendarItem(
            start=tz.localize(EWSDateTime(year, month, day, hour, 30)),
            end=tz.localize(EWSDateTime(year, month, day, hour + 1, 15)),
            subject='Test item',
            body='Hello from Python',
            location='devnull',
            categories=['foo', 'bar'],
        ))

    # Username in WINDOMAIN\username format. Office365 wants usernames in PrimarySMTPAddress
    # ('myusername@example.com') format. UPN format is also supported.
    credentials = Credentials(username='MYWINDOMAIN\\myusername', password='topsecret')

    # If your credentials have been given impersonation access to the target account, use
    # access_type=IMPERSONATION
    account = Account(primary_smtp_address='john@example.com', credentials=credentials,
                      autodiscover=True, access_type=DELEGATE)

    # If the server doesn't support autodiscover, use a Configuration object to set the
    # server location:
    # config = Configuration(server='mail.example.com', username='MYWINDOMAIN\\myusername',
    #                        password='topsecret', auth_type=NTLM)
    # account = Account(primary_smtp_address='john@example.com', config=config,
    #                   access_type=DELEGATE)


    # Create the calendar items in the user's standard calendar.  If you want to access a
    # non-standard calendar, choose a different one from account.folders[Calendar]
    res = account.calendar.add_items(calendar_items)
    print(res)

    # Get Exchange ID and changekey of the calendar items we just created. We search by
    # categories so we only get the items created by us.
    ids = account.calendar.find_items(
        start=tz.localize(EWSDateTime(year, month, day)),
        end=tz.localize(EWSDateTime(year, month, day + 1)),
        categories=['foo', 'bar'],
        shape=IdOnly,
    )
    print(ids)

    # Get the rest of the attributes on the calendar items we just created
    items = account.calendar.get_items(ids)
    print(items)

    # Delete the calendar items again
    res = account.calendar.delete_items(ids)
    print(res)
