from exchangelib.errors import ErrorInvalidIdMalformed
from exchangelib.folders import Contacts, FolderCollection
from exchangelib.indexed_properties import EmailAddress, PhysicalAddress
from exchangelib.items import Contact, DistributionList, Persona
from exchangelib.properties import Mailbox, Member
from exchangelib.queryset import QuerySet
from exchangelib.services import GetPersona

from ..common import get_random_string, get_random_email
from .test_basics import CommonItemTest


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
