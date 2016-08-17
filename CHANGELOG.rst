==========
Change Log
==========

1.6.0
-----
* Complete rewrite of ``Folder.find_items()``. The old ``start``, ``end``, ``subject`` and
  ``categories`` args are deprecated in favor of a Django QuerySet filter() syntax. The
  supported lookup types are ``__gt``, ``__lt``, ``__gte``, ``__lte``, ``__range``, ``__in``,
  ``__exact``,``__iexact``, ``__contains``,``__icontains``, ``__contains``, ``__icontains``,
  ``__startswith``, ``__istartswith``, plus an additional ``__not`` which translates to ``!=``.
  Additionally, *all* fields on the item are now supported in ``Folder.find_items()``.

  **WARNING**: This change is backwards-incompatible! Old uses of ``Folder.find_items()`` like this:

  .. code-block:: python

      ids = account.calendar.find_items(
          start=tz.localize(EWSDateTime(year, month, day)),
          end=tz.localize(EWSDateTime(year, month, day + 1)),
          categories=['foo', 'bar'],
      )

  must be rewritten like this:

  .. code-block:: python

      ids = account.calendar.find_items(
          start__lt=tz.localize(EWSDateTime(year, month, day + 1)),
          end__gt=tz.localize(EWSDateTime(year, month, day)),
          categories__contains=['foo', 'bar'],
      )
      
  failing to do so will most likely result in empty or wrong results.

* Added a ``exchangelib.restrictions.Q`` class much like Django Q objects that can be used to
  create even more complex filtering. Q objects must be passed directly to ``exchangelib.services.FindItem``.

1.3.6
-----
* Don't require sequence arguments to ``Folder.*_items()`` methods to support ``len()``
  (e.g. generators and ``map`` instances are now supported)
* Allow empty sequences as argument to ``Folder.*_items()`` methods


1.3.4
-----
* Add support for ``required_attendees``, ``optional_attendees`` and ``resources``
  attribute on ``folders.CalendarItem``. These are implemented with a new ``folders.Attendee``
  class.


1.3.3
-----
* Add support for ``organizer`` attribute on ``CalendarItem``.  Implemented with a
  new ``folders.Mailbox`` class.


1.2
---
* Initial import
