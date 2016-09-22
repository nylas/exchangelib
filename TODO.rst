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
* Let filter() return full objects, unless explicitly told otherwise. Use FindItem if .only() is non-empty and doesn't
  contain 'dangerous' fields (like 'body'), and fallback to GetItems otherwise.
* Make ``my_folder.filter()`` lazy and support chaining.
* Support chaining with ``filter()``, ``exclude()``, ``only()``, ``order_by()`` and ``reverse()``
  and endpoints ``values()``, ``values_list()``, ``iterator()``, ``get()``, ``count()``, ``exists()`` and ``delete()``
* Move ``Folder.fetch()`` to ``Account.fetch()`` where it logically belongs, since it gets items anywhere in
  the mailbox, not just in the folder. To do this, ``.from_xml()`` must be prepared to handle all item classes (see
  _get_elements_in_container() in services.py). Requires implementing ``.only()`` first since ``.GetItem`` payload XML
  depends on ``self.item_model.additional_property_elems``.
* Enforce SUBJECT_MAXLENGTH and LOCATION_MAXLENGTH (and other validations on item fields)
