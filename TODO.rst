=====
TO DO
=====

Cool things to work on:
-----------------------
* Mail attachments: https://github.com/ecederstrand/exchangelib/issues/15
* Notification subscriptions: https://msdn.microsoft.com/en-us/library/office/dn458790(v=exchg.150).aspx and https://msdn.microsoft.com/en-us/library/office/dn458791(v=exchg.150).aspx
* Password change for accounts: https://support.office.com/en-us/article/Change-password-in-Outlook-Web-App-50bb1309-6f53-4c24-8bfd-ed24ca9e872c
* SendItem service to send draft emails: https://msdn.microsoft.com/en-us/library/office/aa580238(v=exchg.150).aspx
* Make it possible to configure the returned item attributes with ``my_folder.filter(foo=bar).only('subject', 'body')``
* Make ``my_folder.find_items()`` lazy to support the above chaining.
* Make it possible to save or update an item with ``CalendarItem.save()`` or ``Message.send()``
* Support all well-known Messages folders as attributesto account.
