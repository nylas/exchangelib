=====
TO DO
=====

Cool things to work on:
-----------------------
* Mail attachments: https://github.com/ecederstrand/exchangelib/issues/15
* Notification subscriptions: https://msdn.microsoft.com/en-us/library/office/dn458790(v=exchg.150).aspx and https://msdn.microsoft.com/en-us/library/office/dn458791(v=exchg.150).aspx
* Password change for accounts: https://support.office.com/en-us/article/Change-password-in-Outlook-Web-App-50bb1309-6f53-4c24-8bfd-ed24ca9e872c
* SendItem service to send draft emails: https://msdn.microsoft.com/en-us/library/office/aa580238(v=exchg.150).aspx
* Make it possible to configure the returned item attributes with ``my_folder.filter(foo=bar).only('subject', 'body')``.
  Deprecate 'with_extra' and EXTRA_ITEM_FIELDS
* Make ``my_folder.filter()`` lazy to support the above chaining.
* Move ``Folder.get_items()`` to ``Account.get_items()`` where it logically belongs, since it gets items anywhere in
  the mailbox, not just in the folder. To do this, ``from_xml()`` must guess the correct item class. Requires
  implementing ``.only()`` first since get_xml() depends on ``self.item_model.additional_property_elems``.
* Enforce SUBJECT_MAXLENGTH and LOCATION_MAXLENGTH (and other string fields)
