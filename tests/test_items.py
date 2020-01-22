import datetime
from decimal import Decimal
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from keyword import kwlist
import time
import unittest
import unittest.util

from dateutil.relativedelta import relativedelta
from exchangelib.account import SAVE_ONLY, SEND_ONLY, SEND_AND_SAVE_COPY
from exchangelib.attachments import ItemAttachment
from exchangelib.errors import ErrorItemNotFound, ErrorInvalidOperation, ErrorInvalidChangeKey, \
    ErrorUnsupportedPathForQuery, ErrorInvalidValueForProperty, ErrorPropertyUpdate, ErrorInvalidPropertySet, \
    ErrorInvalidIdMalformed
from exchangelib.ewsdatetime import EWSDateTime, EWSTimeZone, UTC, UTC_NOW
from exchangelib.extended_properties import ExtendedProperty, ExternId
from exchangelib.fields import TextField, BodyField, ExtendedPropertyField, FieldPath, CultureField, IdField, \
    CharField, ChoiceField, AttachmentField, BooleanField
from exchangelib.folders import Calendar, Inbox, Tasks, Contacts, Folder, FolderCollection
from exchangelib.indexed_properties import EmailAddress, PhysicalAddress, SingleFieldIndexedElement, \
    MultiFieldIndexedElement
from exchangelib.items import Item, CalendarItem, Message, Contact, Task, DistributionList, Persona, BaseItem, \
    SHALLOW, ASSOCIATED
from exchangelib.properties import Mailbox, Member, Attendee
from exchangelib.queryset import QuerySet, DoesNotExist, MultipleObjectsReturned
from exchangelib.restriction import Restriction, Q
from exchangelib.services import GetPersona
from exchangelib.util import value_to_xml_text
from exchangelib.version import Build, EXCHANGE_2007, EXCHANGE_2013

from .common import EWSTest, get_random_string, get_random_datetime_range, get_random_date, \
    get_random_email, get_random_decimal, get_random_choice, get_random_int, mock_version


class BaseItemTest(EWSTest):
    TEST_FOLDER = None
    FOLDER_CLASS = None
    ITEM_CLASS = None

    @classmethod
    def setUpClass(cls):
        if cls is BaseItemTest:
            raise unittest.SkipTest("Skip BaseItemTest, it's only for inheritance")
        super().setUpClass()

    def setUp(self):
        super().setUp()
        self.test_folder = getattr(self.account, self.TEST_FOLDER)
        self.assertEqual(type(self.test_folder), self.FOLDER_CLASS)
        self.assertEqual(self.test_folder.DISTINGUISHED_FOLDER_ID, self.TEST_FOLDER)
        self.test_folder.filter(categories__contains=self.categories).delete()

    def tearDown(self):
        self.test_folder.filter(categories__contains=self.categories).delete()
        # Delete all delivery receipts
        self.test_folder.filter(subject__startswith='Delivered: Subject: ').delete()
        super().tearDown()

    def get_random_insert_kwargs(self):
        insert_kwargs = {}
        for f in self.ITEM_CLASS.FIELDS:
            if not f.supports_version(self.account.version):
                # Cannot be used with this EWS version
                continue
            if self.ITEM_CLASS == CalendarItem and f in CalendarItem.timezone_fields():
                # Timezone fields will (and must) be populated automatically from the timestamp
                continue
            if f.is_read_only:
                # These cannot be created
                continue
            if f.name == 'mime_content':
                # This needs special formatting. See separate test_mime_content() test
                continue
            if f.name == 'attachments':
                # Testing attachments is heavy. Leave this to specific tests
                insert_kwargs[f.name] = []
                continue
            if f.name == 'resources':
                # The test server doesn't have any resources
                insert_kwargs[f.name] = []
                continue
            if f.name == 'optional_attendees':
                # 'optional_attendees' and 'required_attendees' are mutually exclusive
                insert_kwargs[f.name] = None
                continue
            if f.name == 'start':
                start = get_random_date()
                insert_kwargs[f.name], insert_kwargs['end'] = \
                    get_random_datetime_range(start_date=start, end_date=start, tz=self.account.default_timezone)
                insert_kwargs['recurrence'] = self.random_val(self.ITEM_CLASS.get_field_by_fieldname('recurrence'))
                insert_kwargs['recurrence'].boundary.start = insert_kwargs[f.name].date()
                continue
            if f.name == 'end':
                continue
            if f.name == 'is_all_day':
                # For CalendarItem instances, the 'is_all_day' attribute affects the 'start' and 'end' values. Changing
                # from 'false' to 'true' removes the time part of these datetimes.
                insert_kwargs['is_all_day'] = False
                continue
            if f.name == 'recurrence':
                continue
            if f.name == 'due_date':
                # start_date must be before due_date
                insert_kwargs['start_date'], insert_kwargs[f.name] = \
                    get_random_datetime_range(tz=self.account.default_timezone)
                continue
            if f.name == 'start_date':
                continue
            if f.name == 'status':
                # Start with an incomplete task
                status = get_random_choice(set(f.supported_choices(version=self.account.version)) - {Task.COMPLETED})
                insert_kwargs[f.name] = status
                if status == Task.NOT_STARTED:
                    insert_kwargs['percent_complete'] = Decimal(0)
                else:
                    insert_kwargs['percent_complete'] = get_random_decimal(1, 99)
                continue
            if f.name == 'percent_complete':
                continue
            insert_kwargs[f.name] = self.random_val(f)
        return insert_kwargs

    def get_random_update_kwargs(self, item, insert_kwargs):
        update_kwargs = {}
        now = UTC_NOW()
        for f in self.ITEM_CLASS.FIELDS:
            if not f.supports_version(self.account.version):
                # Cannot be used with this EWS version
                continue
            if self.ITEM_CLASS == CalendarItem and f in CalendarItem.timezone_fields():
                # Timezone fields will (and must) be populated automatically from the timestamp
                continue
            if f.is_read_only:
                # These cannot be changed
                continue
            if not item.is_draft and f.is_read_only_after_send:
                # These cannot be changed when the item is no longer a draft
                continue
            if f.name == 'message_id' and f.is_read_only_after_send:
                # Cannot be updated, regardless of draft status
                continue
            if f.name == 'attachments':
                # Testing attachments is heavy. Leave this to specific tests
                update_kwargs[f.name] = []
                continue
            if f.name == 'resources':
                # The test server doesn't have any resources
                update_kwargs[f.name] = []
                continue
            if isinstance(f, AttachmentField):
                # Attachments are handled separately
                continue
            if f.name == 'start':
                start = get_random_date(start_date=insert_kwargs['end'].date())
                update_kwargs[f.name], update_kwargs['end'] = \
                    get_random_datetime_range(start_date=start, end_date=start, tz=self.account.default_timezone)
                update_kwargs['recurrence'] = self.random_val(self.ITEM_CLASS.get_field_by_fieldname('recurrence'))
                update_kwargs['recurrence'].boundary.start = update_kwargs[f.name].date()
                continue
            if f.name == 'end':
                continue
            if f.name == 'recurrence':
                continue
            if f.name == 'due_date':
                # start_date must be before due_date, and before complete_date which must be in the past
                update_kwargs['start_date'], update_kwargs[f.name] = \
                    get_random_datetime_range(end_date=now.date(), tz=self.account.default_timezone)
                continue
            if f.name == 'start_date':
                continue
            if f.name == 'status':
                # Update task to a completed state. complete_date must be a date in the past, and < than start_date
                update_kwargs[f.name] = Task.COMPLETED
                update_kwargs['percent_complete'] = Decimal(100)
                continue
            if f.name == 'percent_complete':
                continue
            if f.name == 'reminder_is_set':
                if self.ITEM_CLASS == Task:
                    # Task type doesn't allow updating 'reminder_is_set' to True
                    update_kwargs[f.name] = False
                else:
                    update_kwargs[f.name] = not insert_kwargs[f.name]
                continue
            if isinstance(f, BooleanField):
                update_kwargs[f.name] = not insert_kwargs[f.name]
                continue
            if f.value_cls in (Mailbox, Attendee):
                if insert_kwargs[f.name] is None:
                    update_kwargs[f.name] = self.random_val(f)
                else:
                    update_kwargs[f.name] = None
                continue
            update_kwargs[f.name] = self.random_val(f)
        if update_kwargs.get('is_all_day', False):
            # For is_all_day items, EWS will remove the time part of start and end values
            update_kwargs['start'] = update_kwargs['start'].replace(hour=0, minute=0, second=0, microsecond=0)
            update_kwargs['end'] = \
                update_kwargs['end'].replace(hour=0, minute=0, second=0, microsecond=0) + datetime.timedelta(days=1)
        if self.ITEM_CLASS == CalendarItem:
            # EWS always sets due date to 'start'
            update_kwargs['reminder_due_by'] = update_kwargs['start']
        return update_kwargs

    def get_test_item(self, folder=None, categories=None):
        item_kwargs = self.get_random_insert_kwargs()
        item_kwargs['categories'] = categories or self.categories
        return self.ITEM_CLASS(folder=folder or self.test_folder, **item_kwargs)


class ItemQuerySetTest(BaseItemTest):
    TEST_FOLDER = 'inbox'
    FOLDER_CLASS = Inbox
    ITEM_CLASS = Message

    def test_querysets(self):
        test_items = []
        for i in range(4):
            item = self.get_test_item()
            item.subject = 'Item %s' % i
            item.save()
            test_items.append(item)
        qs = QuerySet(
            folder_collection=FolderCollection(account=self.account, folders=[self.test_folder])
        ).filter(categories__contains=self.categories)
        test_cat = self.categories[0]
        self.assertEqual(
            set((i.subject, i.categories[0]) for i in qs),
            {('Item 0', test_cat), ('Item 1', test_cat), ('Item 2', test_cat), ('Item 3', test_cat)}
        )
        self.assertEqual(
            [(i.subject, i.categories[0]) for i in qs.none()],
            []
        )
        self.assertEqual(
            [(i.subject, i.categories[0]) for i in qs.filter(subject__startswith='Item 2')],
            [('Item 2', test_cat)]
        )
        self.assertEqual(
            set((i.subject, i.categories[0]) for i in qs.exclude(subject__startswith='Item 2')),
            {('Item 0', test_cat), ('Item 1', test_cat), ('Item 3', test_cat)}
        )
        self.assertEqual(
            set((i.subject, i.categories) for i in qs.only('subject')),
            {('Item 0', None), ('Item 1', None), ('Item 2', None), ('Item 3', None)}
        )
        self.assertEqual(
            [(i.subject, i.categories[0]) for i in qs.order_by('subject')],
            [('Item 0', test_cat), ('Item 1', test_cat), ('Item 2', test_cat), ('Item 3', test_cat)]
        )
        self.assertEqual(  # Test '-some_field' syntax for reverse sorting
            [(i.subject, i.categories[0]) for i in qs.order_by('-subject')],
            [('Item 3', test_cat), ('Item 2', test_cat), ('Item 1', test_cat), ('Item 0', test_cat)]
        )
        self.assertEqual(  # Test ordering on a field that we don't need to fetch
            [(i.subject, i.categories[0]) for i in qs.order_by('-subject').only('categories')],
            [(None, test_cat), (None, test_cat), (None, test_cat), (None, test_cat)]
        )
        self.assertEqual(
            [(i.subject, i.categories[0]) for i in qs.order_by('subject').reverse()],
            [('Item 3', test_cat), ('Item 2', test_cat), ('Item 1', test_cat), ('Item 0', test_cat)]
        )
        with self.assertRaises(ValueError):
            list(qs.values([]))
        self.assertEqual(
            [i for i in qs.order_by('subject').values('subject')],
            [{'subject': 'Item 0'}, {'subject': 'Item 1'}, {'subject': 'Item 2'}, {'subject': 'Item 3'}]
        )

        # Test .values() in combinations of 'id' and 'changekey', which are handled specially
        self.assertEqual(
            list(qs.order_by('subject').values('id')),
            [{'id': i.id} for i in test_items]
        )
        self.assertEqual(
            list(qs.order_by('subject').values('changekey')),
            [{'changekey': i.changekey} for i in test_items]
        )
        self.assertEqual(
            list(qs.order_by('subject').values('id', 'changekey')),
            [{k: getattr(i, k) for k in ('id', 'changekey')} for i in test_items]
        )

        self.assertEqual(
            set(i for i in qs.values_list('subject')),
            {('Item 0',), ('Item 1',), ('Item 2',), ('Item 3',)}
        )

        # Test .values_list() in combinations of 'id' and 'changekey', which are handled specially
        self.assertEqual(
            list(qs.order_by('subject').values_list('id')),
            [(i.id,) for i in test_items]
        )
        self.assertEqual(
            list(qs.order_by('subject').values_list('changekey')),
            [(i.changekey,) for i in test_items]
        )
        self.assertEqual(
            list(qs.order_by('subject').values_list('id', 'changekey')),
            [(i.id, i.changekey) for i in test_items]
        )

        self.assertEqual(
            set(i.subject for i in qs.only('subject')),
            {'Item 0', 'Item 1', 'Item 2', 'Item 3'}
        )

        # Test .only() in combinations of 'id' and 'changekey', which are handled specially
        self.assertEqual(
            list((i.id,) for i in qs.order_by('subject').only('id')),
            [(i.id,) for i in test_items]
        )
        self.assertEqual(
            list((i.changekey,) for i in qs.order_by('subject').only('changekey')),
            [(i.changekey,) for i in test_items]
        )
        self.assertEqual(
            list((i.id, i.changekey) for i in qs.order_by('subject').only('id', 'changekey')),
            [(i.id, i.changekey) for i in test_items]
        )

        with self.assertRaises(ValueError):
            list(qs.values_list('id', 'changekey', flat=True))
        with self.assertRaises(AttributeError):
            list(qs.values_list('id', xxx=True))
        self.assertEqual(
            list(qs.order_by('subject').values_list('id', flat=True)),
            [i.id for i in test_items]
        )
        self.assertEqual(
            list(qs.order_by('subject').values_list('changekey', flat=True)),
            [i.changekey for i in test_items]
        )
        self.assertEqual(
            set(i for i in qs.values_list('subject', flat=True)),
            {'Item 0', 'Item 1', 'Item 2', 'Item 3'}
        )
        self.assertEqual(
            qs.values_list('subject', flat=True).get(subject='Item 2'),
            'Item 2'
        )
        self.assertEqual(
            set((i.subject, i.categories[0]) for i in qs.exclude(subject__startswith='Item 2')),
            {('Item 0', test_cat), ('Item 1', test_cat), ('Item 3', test_cat)}
        )
        # Test that we can sort on a field that we don't want
        self.assertEqual(
            [i.categories[0] for i in qs.only('categories').order_by('subject')],
            [test_cat, test_cat, test_cat, test_cat]
        )
        # Test iterator
        self.assertEqual(
            set((i.subject, i.categories[0]) for i in qs.iterator()),
            {('Item 0', test_cat), ('Item 1', test_cat), ('Item 2', test_cat), ('Item 3', test_cat)}
        )
        # Test that iterator() preserves the result format
        self.assertEqual(
            set((i[0], i[1][0]) for i in qs.values_list('subject', 'categories').iterator()),
            {('Item 0', test_cat), ('Item 1', test_cat), ('Item 2', test_cat), ('Item 3', test_cat)}
        )
        self.assertEqual(qs.get(subject='Item 3').subject, 'Item 3')
        with self.assertRaises(DoesNotExist):
            qs.get(subject='Item XXX')
        with self.assertRaises(MultipleObjectsReturned):
            qs.get(subject__startswith='Item')
        # len() and count()
        self.assertEqual(len(qs), 4)
        self.assertEqual(qs.count(), 4)
        # Indexing and slicing
        self.assertTrue(isinstance(qs[0], self.ITEM_CLASS))
        self.assertEqual(len(list(qs[1:3])), 2)
        self.assertEqual(len(qs), 4)
        with self.assertRaises(IndexError):
            print(qs[99999])
        # Exists
        self.assertEqual(qs.exists(), True)
        self.assertEqual(qs.filter(subject='Test XXX').exists(), False)
        self.assertEqual(
            qs.filter(subject__startswith='Item').delete(),
            [True, True, True, True]
        )

    def test_queryset_failure(self):
        qs = QuerySet(
            folder_collection=FolderCollection(account=self.account, folders=[self.test_folder])
        ).filter(categories__contains=self.categories)
        with self.assertRaises(ValueError):
            qs.order_by('XXX')
        with self.assertRaises(ValueError):
            qs.values('XXX')
        with self.assertRaises(ValueError):
            qs.values_list('XXX')
        with self.assertRaises(ValueError):
            qs.only('XXX')
        with self.assertRaises(ValueError):
            qs.reverse()  # We can't reverse when we haven't defined an order yet

    def test_cached_queryset_corner_cases(self):
        test_items = []
        for i in range(4):
            item = self.get_test_item()
            item.subject = 'Item %s' % i
            item.save()
            test_items.append(item)
        qs = QuerySet(
            folder_collection=FolderCollection(account=self.account, folders=[self.test_folder])
        ).filter(categories__contains=self.categories).order_by('subject')
        for _ in qs:
            # Build up the cache
            pass
        self.assertEqual(len(qs._cache), 4)
        with self.assertRaises(MultipleObjectsReturned):
            qs.get()  # Get with a full cache
        self.assertEqual(qs[2].subject, 'Item 2')  # Index with a full cache
        self.assertEqual(qs[-2].subject, 'Item 2')  # Negative index with a full cache
        qs.delete()  # Delete with a full cache
        self.assertEqual(qs.count(), 0)  # QuerySet is empty after delete
        self.assertEqual(list(qs.none()), [])

    def test_queryset_get_by_id(self):
        item = self.get_test_item().save()
        with self.assertRaises(ValueError):
            list(self.test_folder.filter(id__in=[item.id]))
        with self.assertRaises(ValueError):
            list(self.test_folder.get(id=item.id, changekey=item.changekey, subject='XXX'))
        with self.assertRaises(ValueError):
            list(self.test_folder.get(id=None, changekey=item.changekey))

        # Test a simple get()
        get_item = self.test_folder.get(id=item.id, changekey=item.changekey)
        self.assertEqual(item.id, get_item.id)
        self.assertEqual(item.changekey, get_item.changekey)
        self.assertEqual(item.subject, get_item.subject)
        self.assertEqual(item.body, get_item.body)

        # Test get() with ID only
        get_item = self.test_folder.get(id=item.id)
        self.assertEqual(item.id, get_item.id)
        self.assertEqual(item.changekey, get_item.changekey)
        self.assertEqual(item.subject, get_item.subject)
        self.assertEqual(item.body, get_item.body)
        get_item = self.test_folder.get(id=item.id, changekey=None)
        self.assertEqual(item.id, get_item.id)
        self.assertEqual(item.changekey, get_item.changekey)
        self.assertEqual(item.subject, get_item.subject)
        self.assertEqual(item.body, get_item.body)

        # Test a get() from queryset
        get_item = self.test_folder.all().get(id=item.id, changekey=item.changekey)
        self.assertEqual(item.id, get_item.id)
        self.assertEqual(item.changekey, get_item.changekey)
        self.assertEqual(item.subject, get_item.subject)
        self.assertEqual(item.body, get_item.body)

        # Test a get() with only()
        get_item = self.test_folder.all().only('subject').get(id=item.id, changekey=item.changekey)
        self.assertEqual(item.id, get_item.id)
        self.assertEqual(item.changekey, get_item.changekey)
        self.assertEqual(item.subject, get_item.subject)
        self.assertIsNone(get_item.body)

    def test_paging(self):
        # Test that paging services work correctly. Default EWS paging size is 1000 items. Our default is 100 items.
        items = []
        for _ in range(11):
            i = self.get_test_item()
            del i.attachments[:]
            items.append(i)
        self.test_folder.bulk_create(items=items)
        ids = self.test_folder.filter(categories__contains=self.categories).values_list('id', 'changekey')
        ids.page_size = 10
        self.bulk_delete(ids.iterator())

    def test_slicing(self):
        # Test that slicing works correctly
        items = []
        for i in range(4):
            item = self.get_test_item()
            item.subject = 'Subj %s' % i
            del item.attachments[:]
            items.append(item)
        ids = self.test_folder.bulk_create(items=items)
        qs = self.test_folder.filter(categories__contains=self.categories).only('subject').order_by('subject')

        # Test positive index
        self.assertEqual(
            qs._copy_self()[0].subject,
            'Subj 0'
        )
        # Test positive index
        self.assertEqual(
            qs._copy_self()[3].subject,
            'Subj 3'
        )
        # Test negative index
        self.assertEqual(
            qs._copy_self()[-2].subject,
            'Subj 2'
        )
        # Test positive slice
        self.assertEqual(
            [i.subject for i in qs._copy_self()[0:2]],
            ['Subj 0', 'Subj 1']
        )
        # Test positive slice
        self.assertEqual(
            [i.subject for i in qs._copy_self()[2:4]],
            ['Subj 2', 'Subj 3']
        )
        # Test positive open slice
        self.assertEqual(
            [i.subject for i in qs._copy_self()[:2]],
            ['Subj 0', 'Subj 1']
        )
        # Test positive open slice
        self.assertEqual(
            [i.subject for i in qs._copy_self()[2:]],
            ['Subj 2', 'Subj 3']
        )
        # Test negative slice
        self.assertEqual(
            [i.subject for i in qs._copy_self()[-3:-1]],
            ['Subj 1', 'Subj 2']
        )
        # Test negative slice
        self.assertEqual(
            [i.subject for i in qs._copy_self()[1:-1]],
            ['Subj 1', 'Subj 2']
        )
        # Test negative open slice
        self.assertEqual(
            [i.subject for i in qs._copy_self()[:-2]],
            ['Subj 0', 'Subj 1']
        )
        # Test negative open slice
        self.assertEqual(
            [i.subject for i in qs._copy_self()[-2:]],
            ['Subj 2', 'Subj 3']
        )
        # Test positive slice with step
        self.assertEqual(
            [i.subject for i in qs._copy_self()[0:4:2]],
            ['Subj 0', 'Subj 2']
        )
        # Test negative slice with step
        self.assertEqual(
            [i.subject for i in qs._copy_self()[4:0:-2]],
            ['Subj 3', 'Subj 1']
        )

    def test_delete_via_queryset(self):
        self.get_test_item().save()
        qs = self.test_folder.filter(categories__contains=self.categories)
        self.assertEqual(qs.count(), 1)
        qs.delete()
        self.assertEqual(qs.count(), 0)

    def test_send_via_queryset(self):
        self.get_test_item().save()
        qs = self.test_folder.filter(categories__contains=self.categories)
        to_folder = self.account.sent
        to_folder_qs = to_folder.filter(categories__contains=self.categories)
        self.assertEqual(qs.count(), 1)
        self.assertEqual(to_folder_qs.count(), 0)
        qs.send(copy_to_folder=to_folder)
        time.sleep(5)  # Requests are supposed to be transactional, but apparently not...
        self.assertEqual(qs.count(), 0)
        self.assertEqual(to_folder_qs.count(), 1)

    def test_send_with_no_copy_via_queryset(self):
        self.get_test_item().save()
        qs = self.test_folder.filter(categories__contains=self.categories)
        to_folder = self.account.sent
        to_folder_qs = to_folder.filter(categories__contains=self.categories)
        self.assertEqual(qs.count(), 1)
        self.assertEqual(to_folder_qs.count(), 0)
        qs.send(save_copy=False)
        time.sleep(5)  # Requests are supposed to be transactional, but apparently not...
        self.assertEqual(qs.count(), 0)
        self.assertEqual(to_folder_qs.count(), 0)

    def test_copy_via_queryset(self):
        self.get_test_item().save()
        qs = self.test_folder.filter(categories__contains=self.categories)
        to_folder = self.account.trash
        to_folder_qs = to_folder.filter(categories__contains=self.categories)
        self.assertEqual(qs.count(), 1)
        self.assertEqual(to_folder_qs.count(), 0)
        qs.copy(to_folder=to_folder)
        self.assertEqual(qs.count(), 1)
        self.assertEqual(to_folder_qs.count(), 1)

    def test_move_via_queryset(self):
        self.get_test_item().save()
        qs = self.test_folder.filter(categories__contains=self.categories)
        to_folder = self.account.trash
        to_folder_qs = to_folder.filter(categories__contains=self.categories)
        self.assertEqual(qs.count(), 1)
        self.assertEqual(to_folder_qs.count(), 0)
        qs.move(to_folder=to_folder)
        self.assertEqual(qs.count(), 0)
        self.assertEqual(to_folder_qs.count(), 1)

    def test_depth(self):
        self.assertGreaterEqual(self.test_folder.all().depth(ASSOCIATED).count(), 0)
        self.assertGreaterEqual(self.test_folder.all().depth(SHALLOW).count(), 0)


class ItemHelperTest(BaseItemTest):
    TEST_FOLDER = 'inbox'
    FOLDER_CLASS = Inbox
    ITEM_CLASS = Message

    def test_save_with_update_fields(self):
        item = self.get_test_item()
        with self.assertRaises(ValueError):
            item.save(update_fields=['subject'])  # update_fields does not work on item creation
        item.save()
        item.subject = 'XXX'
        item.body = 'YYY'
        item.save(update_fields=['subject'])
        item.refresh()
        self.assertEqual(item.subject, 'XXX')
        self.assertNotEqual(item.body, 'YYY')

        # Test invalid 'update_fields' input
        with self.assertRaises(ValueError) as e:
            item.save(update_fields=['xxx'])
        self.assertEqual(
            e.exception.args[0],
            "Field name(s) 'xxx' are not valid for a '%s' item" % self.ITEM_CLASS.__name__
        )
        with self.assertRaises(ValueError) as e:
            item.save(update_fields='subject')
        self.assertEqual(
            e.exception.args[0],
            "Field name(s) 's', 'u', 'b', 'j', 'e', 'c', 't' are not valid for a '%s' item" % self.ITEM_CLASS.__name__
        )

    def test_soft_delete(self):
        # First, empty trash bin
        self.account.trash.filter(categories__contains=self.categories).delete()
        self.account.recoverable_items_deletions.filter(categories__contains=self.categories).delete()
        item = self.get_test_item().save()
        item_id = (item.id, item.changekey)
        # Soft delete
        item.soft_delete()
        for e in self.account.fetch(ids=[item_id]):
            # It's gone from the test folder
            self.assertIsInstance(e, ErrorItemNotFound)
        # Really gone, not just changed ItemId
        self.assertEqual(len(self.test_folder.filter(categories__contains=item.categories)), 0)
        self.assertEqual(len(self.account.trash.filter(categories__contains=item.categories)), 0)
        # But we can find it in the recoverable items folder
        self.assertEqual(len(self.account.recoverable_items_deletions.filter(categories__contains=item.categories)), 1)

    def test_move_to_trash(self):
        # First, empty trash bin
        self.account.trash.filter(categories__contains=self.categories).delete()
        item = self.get_test_item().save()
        item_id = (item.id, item.changekey)
        # Move to trash
        item.move_to_trash()
        for e in self.account.fetch(ids=[item_id]):
            # Not in the test folder anymore
            self.assertIsInstance(e, ErrorItemNotFound)
        # Really gone, not just changed ItemId
        self.assertEqual(len(self.test_folder.filter(categories__contains=item.categories)), 0)
        # Test that the item moved to trash
        item = self.account.trash.get(categories__contains=item.categories)
        moved_item = list(self.account.fetch(ids=[item]))[0]
        # The item was copied, so the ItemId has changed. Let's compare the subject instead
        self.assertEqual(item.subject, moved_item.subject)

    def test_copy(self):
        # First, empty trash bin
        self.account.trash.filter(categories__contains=self.categories).delete()
        item = self.get_test_item().save()
        # Copy to trash. We use trash because it can contain all item types.
        copy_item_id, copy_changekey = item.copy(to_folder=self.account.trash)
        # Test that the item still exists in the folder
        self.assertEqual(len(self.test_folder.filter(categories__contains=item.categories)), 1)
        # Test that the copied item exists in trash
        copied_item = self.account.trash.get(categories__contains=item.categories)
        self.assertNotEqual(item.id, copied_item.id)
        self.assertNotEqual(item.changekey, copied_item.changekey)
        self.assertEqual(copy_item_id, copied_item.id)
        self.assertEqual(copy_changekey, copied_item.changekey)

    def test_move(self):
        # First, empty trash bin
        self.account.trash.filter(categories__contains=self.categories).delete()
        item = self.get_test_item().save()
        item_id = (item.id, item.changekey)
        # Move to trash. We use trash because it can contain all item types. This changes the ItemId
        item.move(to_folder=self.account.trash)
        for e in self.account.fetch(ids=[item_id]):
            # original item ID no longer exists
            self.assertIsInstance(e, ErrorItemNotFound)
        # Test that the item moved to trash
        self.assertEqual(len(self.test_folder.filter(categories__contains=item.categories)), 0)
        moved_item = self.account.trash.get(categories__contains=item.categories)
        self.assertEqual(item.id, moved_item.id)
        self.assertEqual(item.changekey, moved_item.changekey)

    def test_refresh(self):
        # Test that we can refresh items, and that refresh fails if the item no longer exists on the server
        item = self.get_test_item().save()
        orig_subject = item.subject
        item.subject = 'XXX'
        item.refresh()
        self.assertEqual(item.subject, orig_subject)
        item.delete()
        with self.assertRaises(ValueError):
            # Item no longer has an ID
            item.refresh()


class BulkMethodTest(BaseItemTest):
    TEST_FOLDER = 'inbox'
    FOLDER_CLASS = Inbox
    ITEM_CLASS = Message

    def test_fetch(self):
        item = self.get_test_item()
        self.test_folder.bulk_create(items=[item, item])
        ids = self.test_folder.filter(categories__contains=item.categories)
        items = list(self.account.fetch(ids=ids))
        for item in items:
            self.assertIsInstance(item, self.ITEM_CLASS)
        self.assertEqual(len(items), 2)

        items = list(self.account.fetch(ids=ids, only_fields=['subject']))
        self.assertEqual(len(items), 2)

        items = list(self.account.fetch(ids=ids, only_fields=[FieldPath.from_string('subject', self.test_folder)]))
        self.assertEqual(len(items), 2)

    def test_empty_args(self):
        # We allow empty sequences for these methods
        self.assertEqual(self.test_folder.bulk_create(items=[]), [])
        self.assertEqual(list(self.account.fetch(ids=[])), [])
        self.assertEqual(self.account.bulk_create(folder=self.test_folder, items=[]), [])
        self.assertEqual(self.account.bulk_update(items=[]), [])
        self.assertEqual(self.account.bulk_delete(ids=[]), [])
        self.assertEqual(self.account.bulk_send(ids=[]), [])
        self.assertEqual(self.account.bulk_copy(ids=[], to_folder=self.account.trash), [])
        self.assertEqual(self.account.bulk_move(ids=[], to_folder=self.account.trash), [])
        self.assertEqual(self.account.upload(data=[]), [])
        self.assertEqual(self.account.export(items=[]), [])

    def test_qs_args(self):
        # We allow querysets for these methods
        qs = self.test_folder.none()
        self.assertEqual(list(self.account.fetch(ids=qs)), [])
        with self.assertRaises(ValueError):
            # bulk_update() does not allow queryset input
            self.assertEqual(self.account.bulk_update(items=qs), [])
        self.assertEqual(self.account.bulk_delete(ids=qs), [])
        self.assertEqual(self.account.bulk_send(ids=qs), [])
        self.assertEqual(self.account.bulk_copy(ids=qs, to_folder=self.account.trash), [])
        self.assertEqual(self.account.bulk_move(ids=qs, to_folder=self.account.trash), [])
        with self.assertRaises(ValueError):
            # upload() does not allow queryset input
            self.assertEqual(self.account.upload(data=qs), [])
        self.assertEqual(self.account.export(items=qs), [])

    def test_no_kwargs(self):
        self.assertEqual(self.test_folder.bulk_create([]), [])
        self.assertEqual(list(self.account.fetch([])), [])
        self.assertEqual(self.account.bulk_create(self.test_folder, []), [])
        self.assertEqual(self.account.bulk_update([]), [])
        self.assertEqual(self.account.bulk_delete([]), [])
        self.assertEqual(self.account.bulk_send([]), [])
        self.assertEqual(self.account.bulk_copy([], to_folder=self.account.trash), [])
        self.assertEqual(self.account.bulk_move([], to_folder=self.account.trash), [])
        self.assertEqual(self.account.upload([]), [])
        self.assertEqual(self.account.export([]), [])

    def test_invalid_bulk_args(self):
        # Test bulk_create
        with self.assertRaises(ValueError):
            # Folder must belong to account
            self.account.bulk_create(folder=Folder(root=None), items=[])
        with self.assertRaises(AttributeError):
            # Must have folder on save
            self.account.bulk_create(folder=None, items=[], message_disposition=SAVE_ONLY)
        # Test that we can send_and_save with a default folder
        self.account.bulk_create(folder=None, items=[], message_disposition=SEND_AND_SAVE_COPY)
        with self.assertRaises(AttributeError):
            # Must not have folder on send-only
            self.account.bulk_create(folder=self.test_folder, items=[], message_disposition=SEND_ONLY)

        # Test bulk_update
        with self.assertRaises(ValueError):
            # Cannot update in send-only mode
            self.account.bulk_update(items=[], message_disposition=SEND_ONLY)

    def test_bulk_failure(self):
        # Test that bulk_* can handle EWS errors and return the errors in order without losing non-failure results
        items1 = [self.get_test_item().save() for _ in range(3)]
        items1[1].changekey = 'XXX'
        for i, res in enumerate(self.account.bulk_delete(items1)):
            if i == 1:
                self.assertIsInstance(res, ErrorInvalidChangeKey)
            else:
                self.assertEqual(res, True)
        items2 = [self.get_test_item().save() for _ in range(3)]
        items2[1].id = 'AAAA=='
        for i, res in enumerate(self.account.bulk_delete(items2)):
            if i == 1:
                self.assertIsInstance(res, ErrorInvalidIdMalformed)
            else:
                self.assertEqual(res, True)
        items3 = [self.get_test_item().save() for _ in range(3)]
        items3[1].id = items1[0].id
        for i, res in enumerate(self.account.fetch(items3)):
            if i == 1:
                self.assertIsInstance(res, ErrorItemNotFound)
            else:
                self.assertIsInstance(res, Item)


class CommonItemTest(BaseItemTest):
    @classmethod
    def setUpClass(cls):
        if cls is CommonItemTest:
            raise unittest.SkipTest("Skip CommonItemTest, it's only for inheritance")
        super().setUpClass()

    def test_field_names(self):
        # Test that fieldnames don't clash with Python keywords
        for f in self.ITEM_CLASS.FIELDS:
            self.assertNotIn(f.name, kwlist)

    def test_magic(self):
        item = self.get_test_item()
        self.assertIn('subject=', str(item))
        self.assertIn(item.__class__.__name__, repr(item))

    def test_queryset_nonsearchable_fields(self):
        for f in self.ITEM_CLASS.FIELDS:
            with self.subTest(f=f):
                if f.is_searchable or isinstance(f, IdField) or not f.supports_version(self.account.version):
                    continue
                if f.name in ('percent_complete', 'allow_new_time_proposal'):
                    # These fields don't raise an error when used in a filter, but also don't match anything in a filter
                    continue
                try:
                    filter_val = f.clean(self.random_val(f))
                    filter_kwargs = {'%s__in' % f.name: filter_val} if f.is_list else {f.name: filter_val}

                    # We raise ValueError when searching on an is_searchable=False field
                    with self.assertRaises(ValueError):
                        list(self.test_folder.filter(**filter_kwargs))

                    # Make sure the is_searchable=False setting is correct by searching anyway and testing that this
                    # fails server-side. This only works for values that we are actually able to convert to a search
                    # string.
                    try:
                        value_to_xml_text(filter_val)
                    except NotImplementedError:
                        continue

                    f.is_searchable = True
                    if f.name in ('reminder_due_by',):
                        # Filtering is accepted but doesn't work
                        self.assertEqual(
                            len(self.test_folder.filter(**filter_kwargs)),
                            0
                        )
                    else:
                        with self.assertRaises((ErrorUnsupportedPathForQuery, ErrorInvalidValueForProperty)):
                            list(self.test_folder.filter(**filter_kwargs))
                finally:
                    f.is_searchable = False

    def test_filter_on_all_fields(self):
        # Test that we can filter on all field names
        # TODO: Test filtering on subfields of IndexedField
        item = self.get_test_item().save()
        common_qs = self.test_folder.filter(categories__contains=self.categories)
        for f in self.ITEM_CLASS.FIELDS:
            if not f.supports_version(self.account.version):
                # Cannot be used with this EWS version
                continue
            if not f.is_searchable:
                # Cannot be used in a QuerySet
                continue
            val = getattr(item, f.name)
            if val is None:
                # We cannot filter on None values
                continue
            if self.ITEM_CLASS == Contact and f.name in ('body', 'display_name'):
                # filtering 'body' or 'display_name' on Contact items doesn't work at all. Error in EWS?
                continue
            if f.is_list:
                # Filter multi-value fields with =, __in and __contains
                if issubclass(f.value_cls, MultiFieldIndexedElement):
                    # For these, we need to filter on the subfield
                    filter_kwargs = []
                    for v in val:
                        for subfield in f.value_cls.supported_fields(version=self.account.version):
                            field_path = FieldPath(field=f, label=v.label, subfield=subfield)
                            path, subval = field_path.path, field_path.get_value(item)
                            if subval is None:
                                continue
                            filter_kwargs.extend([
                                {path: subval}, {'%s__in' % path: [subval]}, {'%s__contains' % path: [subval]}
                            ])
                elif issubclass(f.value_cls, SingleFieldIndexedElement):
                    # For these, we may filter by item or subfield value
                    filter_kwargs = []
                    for v in val:
                        for subfield in f.value_cls.supported_fields(version=self.account.version):
                            field_path = FieldPath(field=f, label=v.label, subfield=subfield)
                            path, subval = field_path.path, field_path.get_value(item)
                            if subval is None:
                                continue
                            filter_kwargs.extend([
                                {f.name: v}, {path: subval},
                                {'%s__in' % path: [subval]}, {'%s__contains' % path: [subval]}
                            ])
                else:
                    filter_kwargs = [{'%s__in' % f.name: val}, {'%s__contains' % f.name: val}]
            else:
                # Filter all others with =, __in and __contains. We could have more filters here, but these should
                # always match.
                filter_kwargs = [{f.name: val}, {'%s__in' % f.name: [val]}]
                if isinstance(f, TextField) and not isinstance(f, ChoiceField):
                    # Choice fields cannot be filtered using __contains. Sort of makes sense.
                    random_start = get_random_int(min_val=0, max_val=len(val)//2)
                    random_end = get_random_int(min_val=len(val)//2+1, max_val=len(val))
                    filter_kwargs.append({'%s__contains' % f.name: val[random_start:random_end]})
            for kw in filter_kwargs:
                with self.subTest(f=f, kw=kw):
                    matches = len(common_qs.filter(**kw))
                    if isinstance(f, TextField) and f.is_complex:
                        # Complex text fields sometimes fail a search using generated data. In production,
                        # they almost always work anyway. Give it one more try after 10 seconds; it seems EWS does
                        # some sort of indexing that needs to catch up.
                        if not matches:
                            time.sleep(10)
                            matches = len(common_qs.filter(**kw))
                            if not matches and isinstance(f, BodyField):
                                # The body field is particularly nasty in this area. Give up
                                continue
                    self.assertEqual(matches, 1, (f.name, val, kw))

    def test_text_field_settings(self):
        # Test that the max_length and is_complex field settings are correctly set for text fields
        item = self.get_test_item().save()
        for f in self.ITEM_CLASS.FIELDS:
            with self.subTest(f=f):
                if not f.supports_version(self.account.version):
                    # Cannot be used with this EWS version
                    continue
                if not isinstance(f, TextField):
                    continue
                if isinstance(f, ChoiceField):
                    # This one can't contain random values
                    continue
                if isinstance(f, CultureField):
                    # This one can't contain random values
                    continue
                if f.is_read_only:
                    continue
                if f.name == 'categories':
                    # We're filtering on this one, so leave it alone
                    continue
                old_max_length = getattr(f, 'max_length', None)
                old_is_complex = f.is_complex
                try:
                    # Set a string long enough to not be handled by FindItems
                    f.max_length = 4000
                    if f.is_list:
                        setattr(item, f.name, [get_random_string(f.max_length) for _ in range(len(getattr(item, f.name)))])
                    else:
                        setattr(item, f.name, get_random_string(f.max_length))
                    try:
                        item.save(update_fields=[f.name])
                    except ErrorPropertyUpdate:
                        # Some fields throw this error when updated to a huge value
                        self.assertIn(f.name, ['given_name', 'middle_name', 'surname'])
                        continue
                    except ErrorInvalidPropertySet:
                        # Some fields can not be updated after save
                        self.assertTrue(f.is_read_only_after_send)
                        continue
                    # is_complex=True forces the query to use GetItems which will always get the full value
                    f.is_complex = True
                    new_full_item = self.test_folder.all().only(f.name).get(categories__contains=self.categories)
                    new_full = getattr(new_full_item, f.name)
                    if old_max_length:
                        if f.is_list:
                            for s in new_full:
                                self.assertLessEqual(len(s), old_max_length, (f.name, len(s), old_max_length))
                        else:
                            self.assertLessEqual(len(new_full), old_max_length, (f.name, len(new_full), old_max_length))

                    # is_complex=False forces the query to use FindItems which will only get the short value
                    f.is_complex = False
                    new_short_item = self.test_folder.all().only(f.name).get(categories__contains=self.categories)
                    new_short = getattr(new_short_item, f.name)

                    if not old_is_complex:
                        self.assertEqual(new_short, new_full, (f.name, new_short, new_full))
                finally:
                    if old_max_length:
                        f.max_length = old_max_length
                    else:
                        delattr(f, 'max_length')
                    f.is_complex = old_is_complex

    def test_save_and_delete(self):
        # Test that we can create, update and delete single items using methods directly on the item.
        insert_kwargs = self.get_random_insert_kwargs()
        insert_kwargs['categories'] = self.categories
        item = self.ITEM_CLASS(folder=self.test_folder, **insert_kwargs)
        self.assertIsNone(item.id)
        self.assertIsNone(item.changekey)

        # Create
        item.save()
        self.assertIsNotNone(item.id)
        self.assertIsNotNone(item.changekey)
        for k, v in insert_kwargs.items():
            self.assertEqual(getattr(item, k), v, (k, getattr(item, k), v))
        # Test that whatever we have locally also matches whatever is in the DB
        fresh_item = list(self.account.fetch(ids=[item]))[0]
        for f in item.FIELDS:
            with self.subTest(f=f):
                old, new = getattr(item, f.name), getattr(fresh_item, f.name)
                if f.is_read_only and old is None:
                    # Some fields are automatically set server-side
                    continue
                if f.name == 'reminder_due_by':
                    # EWS sets a default value if it is not set on insert. Ignore
                    continue
                if f.name == 'mime_content':
                    # This will change depending on other contents fields
                    continue
                if f.is_list:
                    old, new = set(old or ()), set(new or ())
                self.assertEqual(old, new, (f.name, old, new))

        # Update
        update_kwargs = self.get_random_update_kwargs(item=item, insert_kwargs=insert_kwargs)
        for k, v in update_kwargs.items():
            setattr(item, k, v)
        item.save()
        for k, v in update_kwargs.items():
            self.assertEqual(getattr(item, k), v, (k, getattr(item, k), v))
        # Test that whatever we have locally also matches whatever is in the DB
        fresh_item = list(self.account.fetch(ids=[item]))[0]
        for f in item.FIELDS:
            with self.subTest(f=f):
                old, new = getattr(item, f.name), getattr(fresh_item, f.name)
                if f.is_read_only and old is None:
                    # Some fields are automatically updated server-side
                    continue
                if f.name == 'mime_content':
                    # This will change depending on other contents fields
                    continue
                if f.name == 'reminder_due_by':
                    if new is None:
                        # EWS does not always return a value if reminder_is_set is False.
                        continue
                    if old is not None:
                        # EWS sometimes randomly sets the new reminder due date to one month before or after we
                        # wanted it, and sometimes 30 days before or after. But only sometimes...
                        old_date = old.astimezone(self.account.default_timezone).date()
                        new_date = new.astimezone(self.account.default_timezone).date()
                        if relativedelta(month=1) + new_date == old_date:
                            item.reminder_due_by = new
                            continue
                        if relativedelta(month=1) + old_date == new_date:
                            item.reminder_due_by = new
                            continue
                        elif abs(old_date - new_date) == datetime.timedelta(days=30):
                            item.reminder_due_by = new
                            continue
                if f.is_list:
                    old, new = set(old or ()), set(new or ())
                self.assertEqual(old, new, (f.name, old, new))

        # Hard delete
        item_id = (item.id, item.changekey)
        item.delete()
        for e in self.account.fetch(ids=[item_id]):
            # It's gone from the account
            self.assertIsInstance(e, ErrorItemNotFound)
        # Really gone, not just changed ItemId
        items = self.test_folder.filter(categories__contains=item.categories)
        self.assertEqual(len(items), 0)

    def test_item(self):
        # Test insert
        insert_kwargs = self.get_random_insert_kwargs()
        insert_kwargs['categories'] = self.categories
        item = self.ITEM_CLASS(folder=self.test_folder, **insert_kwargs)
        # Test with generator as argument
        insert_ids = self.test_folder.bulk_create(items=(i for i in [item]))
        self.assertEqual(len(insert_ids), 1)
        self.assertIsInstance(insert_ids[0], BaseItem)
        find_ids = self.test_folder.filter(categories__contains=item.categories).values_list('id', 'changekey')
        self.assertEqual(len(find_ids), 1)
        self.assertEqual(len(find_ids[0]), 2, find_ids[0])
        self.assertEqual(insert_ids, list(find_ids))
        # Test with generator as argument
        item = list(self.account.fetch(ids=(i for i in find_ids)))[0]
        for f in self.ITEM_CLASS.FIELDS:
            with self.subTest(f=f):
                if not f.supports_version(self.account.version):
                    # Cannot be used with this EWS version
                    continue
                if self.ITEM_CLASS == CalendarItem and f in CalendarItem.timezone_fields():
                    # Timezone fields will (and must) be populated automatically from the timestamp
                    continue
                if f.is_read_only:
                    continue
                if f.name == 'reminder_due_by':
                    # EWS sets a default value if it is not set on insert. Ignore
                    continue
                if f.name == 'mime_content':
                    # This will change depending on other contents fields
                    continue
                old, new = getattr(item, f.name), insert_kwargs[f.name]
                if f.is_list:
                    old, new = set(old or ()), set(new or ())
                self.assertEqual(old, new, (f.name, old, new))

        # Test update
        update_kwargs = self.get_random_update_kwargs(item=item, insert_kwargs=insert_kwargs)
        if self.ITEM_CLASS in (Contact, DistributionList):
            # Contact and DistributionList don't support mime_type updates at all
            update_kwargs.pop('mime_content', None)
        update_fieldnames = [f for f in update_kwargs.keys() if f != 'attachments']
        for k, v in update_kwargs.items():
            setattr(item, k, v)
        # Test with generator as argument
        update_ids = self.account.bulk_update(items=(i for i in [(item, update_fieldnames)]))
        self.assertEqual(len(update_ids), 1)
        self.assertEqual(len(update_ids[0]), 2, update_ids)
        self.assertEqual(insert_ids[0].id, update_ids[0][0])  # ID should be the same
        self.assertNotEqual(insert_ids[0].changekey, update_ids[0][1])  # Changekey should change when item is updated
        item = list(self.account.fetch(update_ids))[0]
        for f in self.ITEM_CLASS.FIELDS:
            with self.subTest(f=f):
                if not f.supports_version(self.account.version):
                    # Cannot be used with this EWS version
                    continue
                if self.ITEM_CLASS == CalendarItem and f in CalendarItem.timezone_fields():
                    # Timezone fields will (and must) be populated automatically from the timestamp
                    continue
                if f.is_read_only or f.is_read_only_after_send:
                    # These cannot be changed
                    continue
                if f.name == 'mime_content':
                    # This will change depending on other contents fields
                    continue
                old, new = getattr(item, f.name), update_kwargs[f.name]
                if f.name == 'reminder_due_by':
                    if old is None:
                        # EWS does not always return a value if reminder_is_set is False. Set one now
                        item.reminder_due_by = new
                        continue
                    elif old is not None and new is not None:
                        # EWS sometimes randomly sets the new reminder due date to one month before or after we
                        # wanted it, and sometimes 30 days before or after. But only sometimes...
                        old_date = old.astimezone(self.account.default_timezone).date()
                        new_date = new.astimezone(self.account.default_timezone).date()
                        if relativedelta(month=1) + new_date == old_date:
                            item.reminder_due_by = new
                            continue
                        if relativedelta(month=1) + old_date == new_date:
                            item.reminder_due_by = new
                            continue
                        elif abs(old_date - new_date) == datetime.timedelta(days=30):
                            item.reminder_due_by = new
                            continue
                if f.is_list:
                    old, new = set(old or ()), set(new or ())
                self.assertEqual(old, new, (f.name, old, new))

        # Test wiping or removing fields
        wipe_kwargs = {}
        for f in self.ITEM_CLASS.FIELDS:
            if not f.supports_version(self.account.version):
                # Cannot be used with this EWS version
                continue
            if self.ITEM_CLASS == CalendarItem and f in CalendarItem.timezone_fields():
                # Timezone fields will (and must) be populated automatically from the timestamp
                continue
            if f.is_required or f.is_required_after_save:
                # These cannot be deleted
                continue
            if f.is_read_only or f.is_read_only_after_send:
                # These cannot be changed
                continue
            wipe_kwargs[f.name] = None
        for k, v in wipe_kwargs.items():
            setattr(item, k, v)
        wipe_ids = self.account.bulk_update([(item, update_fieldnames), ])
        self.assertEqual(len(wipe_ids), 1)
        self.assertEqual(len(wipe_ids[0]), 2, wipe_ids)
        self.assertEqual(insert_ids[0].id, wipe_ids[0][0])  # ID should be the same
        self.assertNotEqual(insert_ids[0].changekey,
                            wipe_ids[0][1])  # Changekey should not be the same when item is updated
        item = list(self.account.fetch(wipe_ids))[0]
        for f in self.ITEM_CLASS.FIELDS:
            with self.subTest(f=f):
                if not f.supports_version(self.account.version):
                    # Cannot be used with this EWS version
                    continue
                if self.ITEM_CLASS == CalendarItem and f in CalendarItem.timezone_fields():
                    # Timezone fields will (and must) be populated automatically from the timestamp
                    continue
                if f.is_required or f.is_required_after_save:
                    continue
                if f.is_read_only or f.is_read_only_after_send:
                    continue
                old, new = getattr(item, f.name), wipe_kwargs[f.name]
                if f.is_list:
                    old, new = set(old or ()), set(new or ())
                self.assertEqual(old, new, (f.name, old, new))

        try:
            self.ITEM_CLASS.register('extern_id', ExternId)
            # Test extern_id = None, which deletes the extended property entirely
            extern_id = None
            item.extern_id = extern_id
            wipe2_ids = self.account.bulk_update([(item, ['extern_id']), ])
            self.assertEqual(len(wipe2_ids), 1)
            self.assertEqual(len(wipe2_ids[0]), 2, wipe2_ids)
            self.assertEqual(insert_ids[0].id, wipe2_ids[0][0])  # ID must be the same
            self.assertNotEqual(insert_ids[0].changekey, wipe2_ids[0][1])  # Changekey must change when item is updated
            item = list(self.account.fetch(wipe2_ids))[0]
            self.assertEqual(item.extern_id, extern_id)
        finally:
            self.ITEM_CLASS.deregister('extern_id')


class GenericItemTest(CommonItemTest):
    # Tests that don't need to be run for every single folder type
    TEST_FOLDER = 'inbox'
    FOLDER_CLASS = Inbox
    ITEM_CLASS = Message

    def test_validation(self):
        item = self.get_test_item()
        item.clean()
        for f in self.ITEM_CLASS.FIELDS:
            with self.subTest(f=f):
                # Test field max_length
                if isinstance(f, CharField) and f.max_length:
                    with self.assertRaises(ValueError):
                        setattr(item, f.name, 'a' * (f.max_length + 1))
                        item.clean()
                        setattr(item, f.name, 'a')

    def test_invalid_direct_args(self):
        with self.assertRaises(ValueError):
            item = self.get_test_item()
            item.account = None
            item.save()  # Must have account on save
        with self.assertRaises(ValueError):
            item = self.get_test_item()
            item.id = 'XXX'  # Fake a saved item
            item.account = None
            item.save()  # Must have account on update
        with self.assertRaises(ValueError):
            item = self.get_test_item()
            item.save(update_fields=['foo', 'bar'])  # update_fields is only valid on update

        with self.assertRaises(ValueError):
            item = self.get_test_item()
            item.account = None
            item.refresh()  # Must have account on refresh
        with self.assertRaises(ValueError):
            item = self.get_test_item()
            item.refresh()  # Refresh an item that has not been saved
        with self.assertRaises(ErrorItemNotFound):
            item = self.get_test_item()
            item.save()
            item_id, changekey = item.id, item.changekey
            item.delete()
            item.id, item.changekey = item_id, changekey
            item.refresh()  # Refresh an item that doesn't exist

        with self.assertRaises(ValueError):
            item = self.get_test_item()
            item.account = None
            item.copy(to_folder=self.test_folder)  # Must have an account on copy
        with self.assertRaises(ValueError):
            item = self.get_test_item()
            item.copy(to_folder=self.test_folder)  # Must be an existing item
        with self.assertRaises(ErrorItemNotFound):
            item = self.get_test_item()
            item.save()
            item_id, changekey = item.id, item.changekey
            item.delete()
            item.id, item.changekey = item_id, changekey
            item.copy(to_folder=self.test_folder)  # Item disappeared

        with self.assertRaises(ValueError):
            item = self.get_test_item()
            item.account = None
            item.move(to_folder=self.test_folder)  # Must have an account on move
        with self.assertRaises(ValueError):
            item = self.get_test_item()
            item.move(to_folder=self.test_folder)  # Must be an existing item
        with self.assertRaises(ErrorItemNotFound):
            item = self.get_test_item()
            item.save()
            item_id, changekey = item.id, item.changekey
            item.delete()
            item.id, item.changekey = item_id, changekey
            item.move(to_folder=self.test_folder)  # Item disappeared

        with self.assertRaises(ValueError):
            item = self.get_test_item()
            item.account = None
            item.delete()  # Must have an account
        with self.assertRaises(ValueError):
            item = self.get_test_item()
            item.delete()  # Must be an existing item
        with self.assertRaises(ErrorItemNotFound):
            item = self.get_test_item()
            item.save()
            item_id, changekey = item.id, item.changekey
            item.delete()
            item.id, item.changekey = item_id, changekey
            item.delete()  # Item disappeared

    def test_invalid_kwargs_on_send(self):
        # Only Message class has the send() method
        with self.assertRaises(ValueError):
            item = self.get_test_item()
            item.account = None
            item.send()  # Must have account on send
        with self.assertRaises(ErrorItemNotFound):
            item = self.get_test_item()
            item.save()
            item_id, changekey = item.id, item.changekey
            item.delete()
            item.id, item.changekey = item_id, changekey
            item.send()  # Item disappeared
        with self.assertRaises(AttributeError):
            item = self.get_test_item()
            item.send(copy_to_folder=self.account.trash, save_copy=False)  # Inconsistent args

    def test_unsupported_fields(self):
        # Create a field that is not supported by any current versions. Test that we fail when using this field
        class UnsupportedProp(ExtendedProperty):
            property_set_id = 'deadcafe-beef-beef-beef-deadcafebeef'
            property_name = 'Unsupported Property'
            property_type = 'String'

        attr_name = 'unsupported_property'
        self.ITEM_CLASS.register(attr_name=attr_name, attr_cls=UnsupportedProp)
        try:
            for f in self.ITEM_CLASS.FIELDS:
                if f.name == attr_name:
                    f.supported_from = Build(99, 99, 99, 99)

            with self.assertRaises(ValueError):
                self.test_folder.get(**{attr_name: 'XXX'})
            with self.assertRaises(ValueError):
                list(self.test_folder.filter(**{attr_name: 'XXX'}))
            with self.assertRaises(ValueError):
                list(self.test_folder.all().only(attr_name))
            with self.assertRaises(ValueError):
                list(self.test_folder.all().values(attr_name))
            with self.assertRaises(ValueError):
                list(self.test_folder.all().values_list(attr_name))
        finally:
            self.ITEM_CLASS.deregister(attr_name=attr_name)

    def test_order_by(self):
        # Test order_by() on normal field
        test_items = []
        for i in range(4):
            item = self.get_test_item()
            item.subject = 'Subj %s' % i
            test_items.append(item)
        self.test_folder.bulk_create(items=test_items)
        qs = QuerySet(
            folder_collection=FolderCollection(account=self.account, folders=[self.test_folder])
        ).filter(categories__contains=self.categories)
        self.assertEqual(
            [i for i in qs.order_by('subject').values_list('subject', flat=True)],
            ['Subj 0', 'Subj 1', 'Subj 2', 'Subj 3']
        )
        self.assertEqual(
            [i for i in qs.order_by('-subject').values_list('subject', flat=True)],
            ['Subj 3', 'Subj 2', 'Subj 1', 'Subj 0']
        )
        self.bulk_delete(qs)

        try:
            self.ITEM_CLASS.register('extern_id', ExternId)
            # Test order_by() on ExtendedProperty
            test_items = []
            for i in range(4):
                item = self.get_test_item()
                item.extern_id = 'ID %s' % i
                test_items.append(item)
            self.test_folder.bulk_create(items=test_items)
            qs = QuerySet(
                folder_collection=FolderCollection(account=self.account, folders=[self.test_folder])
            ).filter(categories__contains=self.categories)
            self.assertEqual(
                [i for i in qs.order_by('extern_id').values_list('extern_id', flat=True)],
                ['ID 0', 'ID 1', 'ID 2', 'ID 3']
            )
            self.assertEqual(
                [i for i in qs.order_by('-extern_id').values_list('extern_id', flat=True)],
                ['ID 3', 'ID 2', 'ID 1', 'ID 0']
            )
        finally:
            self.ITEM_CLASS.deregister('extern_id')
        self.bulk_delete(qs)

        # Test sorting on multiple fields
        try:
            self.ITEM_CLASS.register('extern_id', ExternId)
            test_items = []
            for i in range(2):
                for j in range(2):
                    item = self.get_test_item()
                    item.subject = 'Subj %s' % i
                    item.extern_id = 'ID %s' % j
                    test_items.append(item)
            self.test_folder.bulk_create(items=test_items)
            qs = QuerySet(
                folder_collection=FolderCollection(account=self.account, folders=[self.test_folder])
            ).filter(categories__contains=self.categories)
            self.assertEqual(
                [i for i in qs.order_by('subject', 'extern_id').values('subject', 'extern_id')],
                [{'subject': 'Subj 0', 'extern_id': 'ID 0'},
                 {'subject': 'Subj 0', 'extern_id': 'ID 1'},
                 {'subject': 'Subj 1', 'extern_id': 'ID 0'},
                 {'subject': 'Subj 1', 'extern_id': 'ID 1'}]
            )
            self.assertEqual(
                [i for i in qs.order_by('-subject', 'extern_id').values('subject', 'extern_id')],
                [{'subject': 'Subj 1', 'extern_id': 'ID 0'},
                 {'subject': 'Subj 1', 'extern_id': 'ID 1'},
                 {'subject': 'Subj 0', 'extern_id': 'ID 0'},
                 {'subject': 'Subj 0', 'extern_id': 'ID 1'}]
            )
            self.assertEqual(
                [i for i in qs.order_by('subject', '-extern_id').values('subject', 'extern_id')],
                [{'subject': 'Subj 0', 'extern_id': 'ID 1'},
                 {'subject': 'Subj 0', 'extern_id': 'ID 0'},
                 {'subject': 'Subj 1', 'extern_id': 'ID 1'},
                 {'subject': 'Subj 1', 'extern_id': 'ID 0'}]
            )
            self.assertEqual(
                [i for i in qs.order_by('-subject', '-extern_id').values('subject', 'extern_id')],
                [{'subject': 'Subj 1', 'extern_id': 'ID 1'},
                 {'subject': 'Subj 1', 'extern_id': 'ID 0'},
                 {'subject': 'Subj 0', 'extern_id': 'ID 1'},
                 {'subject': 'Subj 0', 'extern_id': 'ID 0'}]
            )
        finally:
            self.ITEM_CLASS.deregister('extern_id')

    def test_finditems(self):
        now = UTC_NOW()

        # Test argument types
        item = self.get_test_item()
        ids = self.test_folder.bulk_create(items=[item])
        # No arguments. There may be leftover items in the folder, so just make sure there's at least one.
        self.assertGreaterEqual(
            len(self.test_folder.filter()),
            1
        )
        # Q object
        self.assertEqual(
            len(self.test_folder.filter(Q(subject=item.subject))),
            1
        )
        # Multiple Q objects
        self.assertEqual(
            len(self.test_folder.filter(Q(subject=item.subject), ~Q(subject=item.subject[:-3] + 'XXX'))),
            1
        )
        # Multiple Q object and kwargs
        self.assertEqual(
            len(self.test_folder.filter(Q(subject=item.subject), categories__contains=item.categories)),
            1
        )
        self.bulk_delete(ids)

        # Test categories which are handled specially - only '__contains' and '__in' lookups are supported
        item = self.get_test_item(categories=['TestA', 'TestB'])
        ids = self.test_folder.bulk_create(items=[item])
        common_qs = self.test_folder.filter(subject=item.subject)  # Guard against other simultaneous runs
        self.assertEqual(
            len(common_qs.filter(categories__contains='ci6xahH1')),  # Plain string
            0
        )
        self.assertEqual(
            len(common_qs.filter(categories__contains=['ci6xahH1'])),  # Same, but as list
            0
        )
        self.assertEqual(
            len(common_qs.filter(categories__contains=['TestA', 'TestC'])),  # One wrong category
            0
        )
        self.assertEqual(
            len(common_qs.filter(categories__contains=['TESTA'])),  # Test case insensitivity
            1
        )
        self.assertEqual(
            len(common_qs.filter(categories__contains=['testa'])),  # Test case insensitivity
            1
        )
        self.assertEqual(
            len(common_qs.filter(categories__contains=['TestA'])),  # Partial
            1
        )
        self.assertEqual(
            len(common_qs.filter(categories__contains=item.categories)),  # Exact match
            1
        )
        with self.assertRaises(ValueError):
            len(common_qs.filter(categories__in='ci6xahH1'))  # Plain string is not supported
        self.assertEqual(
            len(common_qs.filter(categories__in=['ci6xahH1'])),  # Same, but as list
            0
        )
        self.assertEqual(
            len(common_qs.filter(categories__in=['TestA', 'TestC'])),  # One wrong category
            1
        )
        self.assertEqual(
            len(common_qs.filter(categories__in=['TestA'])),  # Partial
            1
        )
        self.assertEqual(
            len(common_qs.filter(categories__in=item.categories)),  # Exact match
            1
        )
        self.bulk_delete(ids)

        common_qs = self.test_folder.filter(categories__contains=self.categories)
        one_hour = datetime.timedelta(hours=1)
        two_hours = datetime.timedelta(hours=2)
        # Test 'exists'
        ids = self.test_folder.bulk_create(items=[self.get_test_item()])
        self.assertEqual(
            len(common_qs.filter(datetime_created__exists=True)),
            1
        )
        self.assertEqual(
            len(common_qs.filter(datetime_created__exists=False)),
            0
        )
        self.bulk_delete(ids)

        # Test 'range'
        ids = self.test_folder.bulk_create(items=[self.get_test_item()])
        self.assertEqual(
            len(common_qs.filter(datetime_created__range=(now + one_hour, now + two_hours))),
            0
        )
        self.assertEqual(
            len(common_qs.filter(datetime_created__range=(now - one_hour, now + one_hour))),
            1
        )
        self.bulk_delete(ids)

        # Test '>'
        ids = self.test_folder.bulk_create(items=[self.get_test_item()])
        self.assertEqual(
            len(common_qs.filter(datetime_created__gt=now + one_hour)),
            0
        )
        self.assertEqual(
            len(common_qs.filter(datetime_created__gt=now - one_hour)),
            1
        )
        self.bulk_delete(ids)

        # Test '>='
        ids = self.test_folder.bulk_create(items=[self.get_test_item()])
        self.assertEqual(
            len(common_qs.filter(datetime_created__gte=now + one_hour)),
            0
        )
        self.assertEqual(
            len(common_qs.filter(datetime_created__gte=now - one_hour)),
            1
        )
        self.bulk_delete(ids)

        # Test '<'
        ids = self.test_folder.bulk_create(items=[self.get_test_item()])
        self.assertEqual(
            len(common_qs.filter(datetime_created__lt=now - one_hour)),
            0
        )
        self.assertEqual(
            len(common_qs.filter(datetime_created__lt=now + one_hour)),
            1
        )
        self.bulk_delete(ids)

        # Test '<='
        ids = self.test_folder.bulk_create(items=[self.get_test_item()])
        self.assertEqual(
            len(common_qs.filter(datetime_created__lte=now - one_hour)),
            0
        )
        self.assertEqual(
            len(common_qs.filter(datetime_created__lte=now + one_hour)),
            1
        )
        self.bulk_delete(ids)

        # Test '='
        item = self.get_test_item()
        ids = self.test_folder.bulk_create(items=[item])
        self.assertEqual(
            len(common_qs.filter(subject=item.subject[:-3] + 'XXX')),
            0
        )
        self.assertEqual(
            len(common_qs.filter(subject=item.subject)),
            1
        )
        self.bulk_delete(ids)

        # Test '!='
        item = self.get_test_item()
        ids = self.test_folder.bulk_create(items=[item])
        self.assertEqual(
            len(common_qs.filter(subject__not=item.subject)),
            0
        )
        self.assertEqual(
            len(common_qs.filter(subject__not=item.subject[:-3] + 'XXX')),
            1
        )
        self.bulk_delete(ids)

        # Test 'exact'
        item = self.get_test_item()
        item.subject = 'aA' + item.subject[2:]
        ids = self.test_folder.bulk_create(items=[item])
        self.assertEqual(
            len(common_qs.filter(subject__exact=item.subject[:-3] + 'XXX')),
            0
        )
        self.assertEqual(
            len(common_qs.filter(subject__exact=item.subject.lower())),
            0
        )
        self.assertEqual(
            len(common_qs.filter(subject__exact=item.subject.upper())),
            0
        )
        self.assertEqual(
            len(common_qs.filter(subject__exact=item.subject)),
            1
        )
        self.bulk_delete(ids)

        # Test 'iexact'
        item = self.get_test_item()
        item.subject = 'aA' + item.subject[2:]
        ids = self.test_folder.bulk_create(items=[item])
        self.assertEqual(
            len(common_qs.filter(subject__iexact=item.subject[:-3] + 'XXX')),
            0
        )
        self.assertIn(
            len(common_qs.filter(subject__iexact=item.subject.lower())),
            (0, 1)  # iexact search is broken on some EWS versions
        )
        self.assertIn(
            len(common_qs.filter(subject__iexact=item.subject.upper())),
            (0, 1)  # iexact search is broken on some EWS versions
        )
        self.assertEqual(
            len(common_qs.filter(subject__iexact=item.subject)),
            1
        )
        self.bulk_delete(ids)

        # Test 'contains'
        item = self.get_test_item()
        item.subject = item.subject[2:8] + 'aA' + item.subject[8:]
        ids = self.test_folder.bulk_create(items=[item])
        self.assertEqual(
            len(common_qs.filter(subject__contains=item.subject[2:14] + 'XXX')),
            0
        )
        self.assertEqual(
            len(common_qs.filter(subject__contains=item.subject[2:14].lower())),
            0
        )
        self.assertEqual(
            len(common_qs.filter(subject__contains=item.subject[2:14].upper())),
            0
        )
        self.assertEqual(
            len(common_qs.filter(subject__contains=item.subject[2:14])),
            1
        )
        self.bulk_delete(ids)

        # Test 'icontains'
        item = self.get_test_item()
        item.subject = item.subject[2:8] + 'aA' + item.subject[8:]
        ids = self.test_folder.bulk_create(items=[item])
        self.assertEqual(
            len(common_qs.filter(subject__icontains=item.subject[2:14] + 'XXX')),
            0
        )
        self.assertIn(
            len(common_qs.filter(subject__icontains=item.subject[2:14].lower())),
            (0, 1)  # icontains search is broken on some EWS versions
        )
        self.assertIn(
            len(common_qs.filter(subject__icontains=item.subject[2:14].upper())),
            (0, 1)  # icontains search is broken on some EWS versions
        )
        self.assertEqual(
            len(common_qs.filter(subject__icontains=item.subject[2:14])),
            1
        )
        self.bulk_delete(ids)

        # Test 'startswith'
        item = self.get_test_item()
        item.subject = 'aA' + item.subject[2:]
        ids = self.test_folder.bulk_create(items=[item])
        self.assertEqual(
            len(common_qs.filter(subject__startswith='XXX' + item.subject[:12])),
            0
        )
        self.assertEqual(
            len(common_qs.filter(subject__startswith=item.subject[:12].lower())),
            0
        )
        self.assertEqual(
            len(common_qs.filter(subject__startswith=item.subject[:12].upper())),
            0
        )
        self.assertEqual(
            len(common_qs.filter(subject__startswith=item.subject[:12])),
            1
        )
        self.bulk_delete(ids)

        # Test 'istartswith'
        item = self.get_test_item()
        item.subject = 'aA' + item.subject[2:]
        ids = self.test_folder.bulk_create(items=[item])
        self.assertEqual(
            len(common_qs.filter(subject__istartswith='XXX' + item.subject[:12])),
            0
        )
        self.assertIn(
            len(common_qs.filter(subject__istartswith=item.subject[:12].lower())),
            (0, 1)  # istartswith search is broken on some EWS versions
        )
        self.assertIn(
            len(common_qs.filter(subject__istartswith=item.subject[:12].upper())),
            (0, 1)  # istartswith search is broken on some EWS versions
        )
        self.assertEqual(
            len(common_qs.filter(subject__istartswith=item.subject[:12])),
            1
        )
        self.bulk_delete(ids)

    def test_filter_with_querystring(self):
        # QueryString is only supported from Exchange 2010
        with self.assertRaises(NotImplementedError):
            Q('Subject:XXX').to_xml(self.test_folder, version=mock_version(build=EXCHANGE_2007),
                                    applies_to=Restriction.ITEMS)

        # We don't allow QueryString in combination with other restrictions
        with self.assertRaises(ValueError):
            self.test_folder.filter('Subject:XXX', foo='bar')
        with self.assertRaises(ValueError):
            self.test_folder.filter('Subject:XXX').filter(foo='bar')
        with self.assertRaises(ValueError):
            self.test_folder.filter(foo='bar').filter('Subject:XXX')

        item = self.get_test_item()
        item.subject = get_random_string(length=8, spaces=False, special=False)
        item.save()
        # For some reason, the querystring search doesn't work instantly. We may have to wait for up to 60 seconds.
        # I'm too impatient for that, so also allow empty results. This makes the test almost worthless but I blame EWS.
        self.assertIn(
            len(self.test_folder.filter('Subject:%s' % item.subject)),
            (0, 1)
        )

    def test_complex_fields(self):
        # Test that complex fields can be fetched using only(). This is a test for #141.
        item = self.get_test_item().save()
        for f in self.ITEM_CLASS.FIELDS:
            with self.subTest(f=f):
                if not f.supports_version(self.account.version):
                    # Cannot be used with this EWS version
                    continue
                if f.name in ('optional_attendees', 'required_attendees', 'resources'):
                    continue
                if f.is_read_only:
                    continue
                if f.name == 'reminder_due_by':
                    # EWS sets a default value if it is not set on insert. Ignore
                    continue
                if f.name == 'mime_content':
                    # This will change depending on other contents fields
                    continue
                old = getattr(item, f.name)
                # Test field as single element in only()
                fresh_item = self.test_folder.all().only(f.name).get(categories__contains=item.categories)
                new = getattr(fresh_item, f.name)
                if f.is_list:
                    old, new = set(old or ()), set(new or ())
                self.assertEqual(old, new, (f.name, old, new))
                # Test field as one of the elements in only()
                fresh_item = self.test_folder.all().only('subject', f.name).get(categories__contains=item.categories)
                new = getattr(fresh_item, f.name)
                if f.is_list:
                    old, new = set(old or ()), set(new or ())
                self.assertEqual(old, new, (f.name, old, new))

    def test_text_body(self):
        if self.account.version.build < EXCHANGE_2013:
            raise self.skipTest('Exchange version too old')
        item = self.get_test_item()
        item.body = 'X' * 500  # Make body longer than the normal 256 char text field limit
        item.save()
        fresh_item = self.test_folder.filter(categories__contains=item.categories).only('text_body')[0]
        self.assertEqual(fresh_item.text_body, item.body)

    def test_only_fields(self):
        item = self.get_test_item().save()
        item = self.test_folder.get(categories__contains=item.categories)
        self.assertIsInstance(item, self.ITEM_CLASS)
        for f in self.ITEM_CLASS.FIELDS:
            with self.subTest(f=f):
                self.assertTrue(hasattr(item, f.name))
                if not f.supports_version(self.account.version):
                    # Cannot be used with this EWS version
                    continue
                if f.name in ('optional_attendees', 'required_attendees', 'resources'):
                    continue
                if f.name == 'reminder_due_by' and not item.reminder_is_set:
                    # We delete the due date if reminder is not set
                    continue
                elif f.is_read_only:
                    continue
                self.assertIsNotNone(getattr(item, f.name), (f, getattr(item, f.name)))
        only_fields = ('subject', 'body', 'categories')
        item = self.test_folder.all().only(*only_fields).get(categories__contains=item.categories)
        self.assertIsInstance(item, self.ITEM_CLASS)
        for f in self.ITEM_CLASS.FIELDS:
            with self.subTest(f=f):
                self.assertTrue(hasattr(item, f.name))
                if not f.supports_version(self.account.version):
                    # Cannot be used with this EWS version
                    continue
                if f.name in only_fields:
                    self.assertIsNotNone(getattr(item, f.name), (f.name, getattr(item, f.name)))
                elif f.is_required:
                    v = getattr(item, f.name)
                    if f.name == 'attachments':
                        self.assertEqual(v, [], (f.name, v))
                    elif f.default is None:
                        self.assertIsNone(v, (f.name, v))
                    else:
                        self.assertEqual(v, f.default, (f.name, v))

    def test_export_and_upload(self):
        # 15 new items which we will attempt to export and re-upload
        items = [self.get_test_item().save() for _ in range(15)]
        ids = [(i.id, i.changekey) for i in items]
        # re-fetch items because there will be some extra fields added by the server
        items = list(self.account.fetch(items))

        # Try exporting and making sure we get the right response
        export_results = self.account.export(items)
        self.assertEqual(len(items), len(export_results))
        for result in export_results:
            self.assertIsInstance(result, str)

        # Try reuploading our results
        upload_results = self.account.upload([(self.test_folder, data) for data in export_results])
        self.assertEqual(len(items), len(upload_results), (items, upload_results))
        for result in upload_results:
            # Must be a completely new ItemId
            self.assertIsInstance(result, tuple)
            self.assertNotIn(result, ids)

        # Check the items uploaded are the same as the original items
        def to_dict(item):
            dict_item = {}
            # fieldnames is everything except the ID so we'll use it to compare
            for f in item.FIELDS:
                # datetime_created and last_modified_time aren't copied, but instead are added to the new item after
                # uploading. This means mime_content and size can also change. Items also get new IDs on upload. And
                # meeting_count values are dependent on contents of current calendar. Form query strings contain the
                # item ID and will also change.
                if f.name in {'id', 'changekey', 'first_occurrence', 'last_occurrence', 'datetime_created',
                              'last_modified_time', 'mime_content', 'size', 'conversation_id',
                              'adjacent_meeting_count', 'conflicting_meeting_count',
                              'web_client_read_form_query_string', 'web_client_edit_form_query_string'}:
                    continue
                dict_item[f.name] = getattr(item, f.name)
                if f.name == 'attachments':
                    # Attachments get new IDs on upload. Wipe them here so we can compare the other fields
                    for a in dict_item[f.name]:
                        a.attachment_id = None
            return dict_item

        uploaded_items = sorted([to_dict(item) for item in self.account.fetch(upload_results)],
                                key=lambda i: i['subject'])
        original_items = sorted([to_dict(item) for item in items], key=lambda i: i['subject'])
        self.assertListEqual(original_items, uploaded_items)

    def test_export_with_error(self):
        # 15 new items which we will attempt to export and re-upload
        items = [self.get_test_item().save() for _ in range(15)]
        # Use id tuples for export here because deleting an item clears it's
        #  id.
        ids = [(item.id, item.changekey) for item in items]
        # Delete one of the items, this will cause an error
        items[3].delete()

        export_results = self.account.export(ids)
        self.assertEqual(len(items), len(export_results))
        for idx, result in enumerate(export_results):
            if idx == 3:
                # If it is the one returning the error
                self.assertIsInstance(result, ErrorItemNotFound)
            else:
                self.assertIsInstance(result, str)

        # Clean up after yourself
        del ids[3]  # Sending the deleted one through will cause an error

    def test_item_attachments(self):
        item = self.get_test_item(folder=self.test_folder)
        item.attachments = []

        attached_item1 = self.get_test_item(folder=self.test_folder)
        attached_item1.attachments = []
        attached_item1.save()
        attachment1 = ItemAttachment(name='attachment1', item=attached_item1)
        item.attach(attachment1)

        self.assertEqual(len(item.attachments), 1)
        item.save()
        fresh_item = list(self.account.fetch(ids=[item]))[0]
        self.assertEqual(len(fresh_item.attachments), 1)
        fresh_attachments = sorted(fresh_item.attachments, key=lambda a: a.name)
        self.assertEqual(fresh_attachments[0].name, 'attachment1')
        self.assertIsInstance(fresh_attachments[0].item, self.ITEM_CLASS)

        for f in self.ITEM_CLASS.FIELDS:
            with self.subTest(f=f):
                # Normalize some values we don't control
                if f.is_read_only:
                    continue
                if self.ITEM_CLASS == CalendarItem and f in CalendarItem.timezone_fields():
                    # Timezone fields will (and must) be populated automatically from the timestamp
                    continue
                if isinstance(f, ExtendedPropertyField):
                    # Attachments don't have these values. It may be possible to request it if we can find the FieldURI
                    continue
                if f.name == 'is_read':
                    # This is always true for item attachments?
                    continue
                if f.name == 'reminder_due_by':
                    # EWS sets a default value if it is not set on insert. Ignore
                    continue
                if f.name == 'mime_content':
                    # This will change depending on other contents fields
                    continue
                old_val = getattr(attached_item1, f.name)
                new_val = getattr(fresh_attachments[0].item, f.name)
                if f.is_list:
                    old_val, new_val = set(old_val or ()), set(new_val or ())
                self.assertEqual(old_val, new_val, (f.name, old_val, new_val))

        # Test attach on saved object
        attached_item2 = self.get_test_item(folder=self.test_folder)
        attached_item2.attachments = []
        attached_item2.save()
        attachment2 = ItemAttachment(name='attachment2', item=attached_item2)
        item.attach(attachment2)

        self.assertEqual(len(item.attachments), 2)
        fresh_item = list(self.account.fetch(ids=[item]))[0]
        self.assertEqual(len(fresh_item.attachments), 2)
        fresh_attachments = sorted(fresh_item.attachments, key=lambda a: a.name)
        self.assertEqual(fresh_attachments[0].name, 'attachment1')
        self.assertIsInstance(fresh_attachments[0].item, self.ITEM_CLASS)

        for f in self.ITEM_CLASS.FIELDS:
            with self.subTest(f=f):
                # Normalize some values we don't control
                if f.is_read_only:
                    continue
                if self.ITEM_CLASS == CalendarItem and f in CalendarItem.timezone_fields():
                    # Timezone fields will (and must) be populated automatically from the timestamp
                    continue
                if isinstance(f, ExtendedPropertyField):
                    # Attachments don't have these values. It may be possible to request it if we can find the FieldURI
                    continue
                if f.name == 'reminder_due_by':
                    # EWS sets a default value if it is not set on insert. Ignore
                    continue
                if f.name == 'is_read':
                    # This is always true for item attachments?
                    continue
                if f.name == 'mime_content':
                    # This will change depending on other contents fields
                    continue
                old_val = getattr(attached_item1, f.name)
                new_val = getattr(fresh_attachments[0].item, f.name)
                if f.is_list:
                    old_val, new_val = set(old_val or ()), set(new_val or ())
                self.assertEqual(old_val, new_val, (f.name, old_val, new_val))

        self.assertEqual(fresh_attachments[1].name, 'attachment2')
        self.assertIsInstance(fresh_attachments[1].item, self.ITEM_CLASS)

        for f in self.ITEM_CLASS.FIELDS:
            # Normalize some values we don't control
            if f.is_read_only:
                continue
            if self.ITEM_CLASS == CalendarItem and f in CalendarItem.timezone_fields():
                # Timezone fields will (and must) be populated automatically from the timestamp
                continue
            if isinstance(f, ExtendedPropertyField):
                # Attachments don't have these values. It may be possible to request it if we can find the FieldURI
                continue
            if f.name == 'reminder_due_by':
                # EWS sets a default value if it is not set on insert. Ignore
                continue
            if f.name == 'is_read':
                # This is always true for item attachments?
                continue
            if f.name == 'mime_content':
                # This will change depending on other contents fields
                continue
            old_val = getattr(attached_item2, f.name)
            new_val = getattr(fresh_attachments[1].item, f.name)
            if f.is_list:
                old_val, new_val = set(old_val or ()), set(new_val or ())
            self.assertEqual(old_val, new_val, (f.name, old_val, new_val))

        # Test detach
        item.detach(attachment2)
        self.assertTrue(attachment2.attachment_id is None)
        self.assertTrue(attachment2.parent_item is None)
        fresh_item = list(self.account.fetch(ids=[item]))[0]
        self.assertEqual(len(fresh_item.attachments), 1)
        fresh_attachments = sorted(fresh_item.attachments, key=lambda a: a.name)

        for f in self.ITEM_CLASS.FIELDS:
            with self.subTest(f=f):
                # Normalize some values we don't control
                if f.is_read_only:
                    continue
                if self.ITEM_CLASS == CalendarItem and f in CalendarItem.timezone_fields():
                    # Timezone fields will (and must) be populated automatically from the timestamp
                    continue
                if isinstance(f, ExtendedPropertyField):
                    # Attachments don't have these values. It may be possible to request it if we can find the FieldURI
                    continue
                if f.name == 'reminder_due_by':
                    # EWS sets a default value if it is not set on insert. Ignore
                    continue
                if f.name == 'is_read':
                    # This is always true for item attachments?
                    continue
                if f.name == 'mime_content':
                    # This will change depending on other contents fields
                    continue
                old_val = getattr(attached_item1, f.name)
                new_val = getattr(fresh_attachments[0].item, f.name)
                if f.is_list:
                    old_val, new_val = set(old_val or ()), set(new_val or ())
                self.assertEqual(old_val, new_val, (f.name, old_val, new_val))

        # Test attach with non-saved item
        attached_item3 = self.get_test_item(folder=self.test_folder)
        attached_item3.attachments = []
        attachment3 = ItemAttachment(name='attachment2', item=attached_item3)
        item.attach(attachment3)
        item.detach(attachment3)


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
        # Test that start and end datetimes for all-day items are returned in the datetime of the account.
        start = get_random_date()
        start_dt, end_dt = get_random_datetime_range(
            start_date=start,
            end_date=start + datetime.timedelta(days=365),
            tz=self.account.default_timezone
        )
        item = self.ITEM_CLASS(folder=self.test_folder, start=start_dt, end=end_dt, is_all_day=True,
                               categories=self.categories)
        item.save()

        item = self.test_folder.all().only('start', 'end').get(id=item.id, changekey=item.changekey)
        self.assertEqual(item.start.astimezone(self.account.default_timezone).time(), datetime.time(0, 0))
        self.assertEqual(item.end.astimezone(self.account.default_timezone).time(), datetime.time(0, 0))

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
            return set(i.categories) == set(self.categories)

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


class MessagesTest(CommonItemTest):
    # Just test one of the Message-type folders
    TEST_FOLDER = 'inbox'
    FOLDER_CLASS = Inbox
    ITEM_CLASS = Message
    INCOMING_MESSAGE_TIMEOUT = 20

    def get_incoming_message(self, subject):
        t1 = time.monotonic()
        while True:
            t2 = time.monotonic()
            if t2 - t1 > self.INCOMING_MESSAGE_TIMEOUT:
                raise self.skipTest('Too bad. Gave up in %s waiting for the incoming message to show up' % self.id())
            try:
                return self.account.inbox.get(subject=subject)
            except DoesNotExist:
                time.sleep(5)

    def test_send(self):
        # Test that we can send (only) Message items
        item = self.get_test_item()
        item.folder = None
        item.send()
        self.assertIsNone(item.id)
        self.assertIsNone(item.changekey)
        self.assertEqual(len(self.test_folder.filter(categories__contains=item.categories)), 0)

    def test_send_and_save(self):
        # Test that we can send_and_save Message items
        item = self.get_test_item()
        item.send_and_save()
        self.assertIsNone(item.id)
        self.assertIsNone(item.changekey)
        time.sleep(5)  # Requests are supposed to be transactional, but apparently not...
        # Also, the sent item may be followed by an automatic message with the same category
        self.assertGreaterEqual(len(self.test_folder.filter(categories__contains=item.categories)), 1)

        # Test update, although it makes little sense
        item = self.get_test_item()
        item.save()
        item.send_and_save()
        time.sleep(5)  # Requests are supposed to be transactional, but apparently not...
        # Also, the sent item may be followed by an automatic message with the same category
        self.assertGreaterEqual(len(self.test_folder.filter(categories__contains=item.categories)), 1)

    def test_send_draft(self):
        item = self.get_test_item()
        item.folder = self.account.drafts
        item.is_draft = True
        item.save()  # Save a draft
        item.send()  # Send the draft
        self.assertIsNone(item.id)
        self.assertIsNone(item.changekey)
        self.assertIsNone(item.folder)
        self.assertEqual(len(self.test_folder.filter(categories__contains=item.categories)), 0)

    def test_send_and_copy_to_folder(self):
        item = self.get_test_item()
        item.send(save_copy=True, copy_to_folder=self.account.sent)  # Send the draft and save to the sent folder
        self.assertIsNone(item.id)
        self.assertIsNone(item.changekey)
        self.assertEqual(item.folder, self.account.sent)
        time.sleep(5)  # Requests are supposed to be transactional, but apparently not...
        self.assertEqual(len(self.account.sent.filter(categories__contains=item.categories)), 1)

    def test_bulk_send(self):
        with self.assertRaises(AttributeError):
            self.account.bulk_send(ids=[], save_copy=False, copy_to_folder=self.account.trash)
        item = self.get_test_item()
        item.save()
        for res in self.account.bulk_send(ids=[item]):
            self.assertEqual(res, True)
        time.sleep(10)  # Requests are supposed to be transactional, but apparently not...
        # By default, sent items are placed in the sent folder
        ids = self.account.sent.filter(categories__contains=item.categories).values_list('id', 'changekey')
        self.assertEqual(len(ids), 1)

    def test_reply(self):
        # Test that we can reply to a Message item. EWS only allows items that have been sent to receive a reply
        item = self.get_test_item()
        item.folder = None
        item.send()  # get_test_item() sets the to_recipients to the test account
        sent_item = self.get_incoming_message(item.subject)
        new_subject = ('Re: %s' % sent_item.subject)[:255]
        sent_item.reply(subject=new_subject, body='Hello reply', to_recipients=[item.author])
        reply = self.get_incoming_message(new_subject)
        self.account.bulk_delete([sent_item, reply])

    def test_reply_all(self):
        # Test that we can reply-all a Message item. EWS only allows items that have been sent to receive a reply
        item = self.get_test_item(folder=None)
        item.folder = None
        item.send()
        sent_item = self.get_incoming_message(item.subject)
        new_subject = ('Re: %s' % sent_item.subject)[:255]
        sent_item.reply_all(subject=new_subject, body='Hello reply')
        reply = self.get_incoming_message(new_subject)
        self.account.bulk_delete([sent_item, reply])

    def test_forward(self):
        # Test that we can forward a Message item. EWS only allows items that have been sent to receive a reply
        item = self.get_test_item(folder=None)
        item.folder = None
        item.send()
        sent_item = self.get_incoming_message(item.subject)
        new_subject = ('Re: %s' % sent_item.subject)[:255]
        sent_item.forward(subject=new_subject, body='Hello reply', to_recipients=[item.author])
        reply = self.get_incoming_message(new_subject)
        reply2 = sent_item.create_forward(subject=new_subject, body='Hello reply', to_recipients=[item.author])
        reply2 = reply2.save(self.account.drafts)
        self.assertIsInstance(reply2, Message)

        self.account.bulk_delete([sent_item, reply, reply2])

    def test_mime_content(self):
        # Tests the 'mime_content' field
        subject = get_random_string(16)
        msg = MIMEMultipart()
        msg['From'] = self.account.primary_smtp_address
        msg['To'] = self.account.primary_smtp_address
        msg['Subject'] = subject
        body = 'MIME test mail'
        msg.attach(MIMEText(body, 'plain', _charset='utf-8'))
        mime_content = msg.as_bytes()
        item = self.ITEM_CLASS(
            folder=self.test_folder,
            to_recipients=[self.account.primary_smtp_address],
            mime_content=mime_content,
            categories=self.categories,
        ).save()
        self.assertEqual(self.test_folder.get(subject=subject).body, body)


class TasksTest(CommonItemTest):
    TEST_FOLDER = 'tasks'
    FOLDER_CLASS = Tasks
    ITEM_CLASS = Task

    def test_task_validation(self):
        tz = EWSTimeZone.timezone('Europe/Copenhagen')
        task = Task(due_date=tz.localize(EWSDateTime(2017, 1, 1)), start_date=tz.localize(EWSDateTime(2017, 2, 1)))
        task.clean()
        # We reset due date if it's before start date
        self.assertEqual(task.due_date, tz.localize(EWSDateTime(2017, 2, 1)))
        self.assertEqual(task.due_date, task.start_date)

        task = Task(complete_date=tz.localize(EWSDateTime(2099, 1, 1)), status=Task.NOT_STARTED)
        task.clean()
        # We reset status if complete_date is set
        self.assertEqual(task.status, Task.COMPLETED)
        # We also reset complete date to now() if it's in the future
        self.assertEqual(task.complete_date.date(), UTC_NOW().date())

        task = Task(complete_date=tz.localize(EWSDateTime(2017, 1, 1)), start_date=tz.localize(EWSDateTime(2017, 2, 1)))
        task.clean()
        # We also reset complete date to start_date if it's before start_date
        self.assertEqual(task.complete_date, task.start_date)

        task = Task(percent_complete=Decimal('50.0'), status=Task.COMPLETED)
        task.clean()
        # We reset percent_complete to 100.0 if state is completed
        self.assertEqual(task.percent_complete, Decimal(100))

        task = Task(percent_complete=Decimal('50.0'), status=Task.NOT_STARTED)
        task.clean()
        # We reset percent_complete to 0.0 if state is not_started
        self.assertEqual(task.percent_complete, Decimal(0))

    def test_complete(self):
        item = self.get_test_item().save()
        item.refresh()
        self.assertNotEqual(item.status, Task.COMPLETED)
        self.assertNotEqual(item.percent_complete, Decimal(100))
        item.complete()
        item.refresh()
        self.assertEqual(item.status, Task.COMPLETED)
        self.assertEqual(item.percent_complete, Decimal(100))


class ContactsTest(CommonItemTest):
    TEST_FOLDER = 'contacts'
    FOLDER_CLASS = Contacts
    ITEM_CLASS = Contact

    def test_order_by_on_indexed_field(self):
        # Test order_by() on IndexedField (simple and multi-subfield). Only Contact items have these
        test_items = []
        label = self.random_val(EmailAddress.get_field_by_fieldname('label'))
        for i in range(4):
            item = self.get_test_item()
            item.email_addresses = [EmailAddress(email='%s@foo.com' % i, label=label)]
            test_items.append(item)
        self.test_folder.bulk_create(items=test_items)
        qs = QuerySet(
            folder_collection=FolderCollection(account=self.account, folders=[self.test_folder])
        ).filter(categories__contains=self.categories)
        self.assertEqual(
            [i[0].email for i in qs.order_by('email_addresses__%s' % label)
                .values_list('email_addresses', flat=True)],
            ['0@foo.com', '1@foo.com', '2@foo.com', '3@foo.com']
        )
        self.assertEqual(
            [i[0].email for i in qs.order_by('-email_addresses__%s' % label)
                .values_list('email_addresses', flat=True)],
            ['3@foo.com', '2@foo.com', '1@foo.com', '0@foo.com']
        )
        self.bulk_delete(qs)

        test_items = []
        label = self.random_val(PhysicalAddress.get_field_by_fieldname('label'))
        for i in range(4):
            item = self.get_test_item()
            item.physical_addresses = [PhysicalAddress(street='Elm St %s' % i, label=label)]
            test_items.append(item)
        self.test_folder.bulk_create(items=test_items)
        qs = QuerySet(
            folder_collection=FolderCollection(account=self.account, folders=[self.test_folder])
        ).filter(categories__contains=self.categories)
        self.assertEqual(
            [i[0].street for i in qs.order_by('physical_addresses__%s__street' % label)
                .values_list('physical_addresses', flat=True)],
            ['Elm St 0', 'Elm St 1', 'Elm St 2', 'Elm St 3']
        )
        self.assertEqual(
            [i[0].street for i in qs.order_by('-physical_addresses__%s__street' % label)
                .values_list('physical_addresses', flat=True)],
            ['Elm St 3', 'Elm St 2', 'Elm St 1', 'Elm St 0']
        )
        self.bulk_delete(qs)

    def test_order_by_failure(self):
        # Test error handling on indexed properties with labels and subfields
        qs = QuerySet(
            folder_collection=FolderCollection(account=self.account, folders=[self.test_folder])
        ).filter(categories__contains=self.categories)
        with self.assertRaises(ValueError):
            qs.order_by('email_addresses')  # Must have label
        with self.assertRaises(ValueError):
            qs.order_by('email_addresses__FOO')  # Must have a valid label
        with self.assertRaises(ValueError):
            qs.order_by('email_addresses__EmailAddress1__FOO')  # Must not have a subfield
        with self.assertRaises(ValueError):
            qs.order_by('physical_addresses__Business')  # Must have a subfield
        with self.assertRaises(ValueError):
            qs.order_by('physical_addresses__Business__FOO')  # Must have a valid subfield

    def test_distribution_lists(self):
        dl = DistributionList(folder=self.test_folder, display_name=get_random_string(255), categories=self.categories)
        dl.save()
        new_dl = self.test_folder.get(categories__contains=dl.categories)
        self.assertEqual(new_dl.display_name, dl.display_name)
        self.assertEqual(new_dl.members, None)
        dl.refresh()

        dl.members = set(
            # We set mailbox_type to OneOff because otherwise the email address must be an actual account
            Member(mailbox=Mailbox(email_address=get_random_email(), mailbox_type='OneOff')) for _ in range(4)
        )
        dl.save()
        new_dl = self.test_folder.get(categories__contains=dl.categories)
        self.assertEqual({m.mailbox.email_address for m in new_dl.members}, dl.members)

        dl.delete()

    def test_find_people(self):
        # The test server may not have any contacts. Just test that the FindPeople service and helpers work
        self.assertGreaterEqual(len(list(self.test_folder.people())), 0)
        self.assertGreaterEqual(
            len(list(
                self.test_folder.people().only('display_name').filter(display_name='john').order_by('display_name')
            )),
            0
        )

    def test_get_persona(self):
        # The test server may not have any personas. Just test that the service response with something we can parse
        persona = Persona(id='AAA=', changekey='xxx')
        try:
            GetPersona(protocol=self.account.protocol).call(persona=persona)
        except ErrorInvalidIdMalformed:
            pass
