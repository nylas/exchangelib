# coding=utf-8
from collections import namedtuple

from exchangelib import FolderCollection, Q
from exchangelib.folders import Inbox
from exchangelib.queryset import QuerySet

from .common import TimedTestCase


class QuerySetTest(TimedTestCase):
    def test_magic(self):
        self.assertEqual(
            str(
                QuerySet(folder_collection=FolderCollection(account=None, folders=[Inbox(root='XXX', name='FooBox')]))
            ),
            'QuerySet(q=Q(), folders=[Inbox (FooBox)])'
        )

    def test_from_folder(self):
        MockRoot = namedtuple('Root', ['account'])
        folder = Inbox(root=MockRoot(account='XXX'))
        self.assertIsInstance(folder.all(), QuerySet)
        self.assertIsInstance(folder.none(), QuerySet)
        self.assertIsInstance(folder.filter(subject='foo'), QuerySet)
        self.assertIsInstance(folder.exclude(subject='foo'), QuerySet)

    def test_queryset_copy(self):
        qs = QuerySet(folder_collection=FolderCollection(account=None, folders=[Inbox(root='XXX')]))
        qs.q = Q()
        qs.only_fields = ('a', 'b')
        qs.order_fields = ('c', 'd')
        qs.return_format = QuerySet.NONE

        # Initially, immutable items have the same id()
        new_qs = qs._copy_self()
        self.assertNotEqual(id(qs), id(new_qs))
        self.assertEqual(id(qs.folder_collection), id(new_qs.folder_collection))
        self.assertEqual(id(qs._cache), id(new_qs._cache))
        self.assertEqual(qs._cache, new_qs._cache)
        self.assertNotEqual(id(qs.q), id(new_qs.q))
        self.assertEqual(qs.q, new_qs.q)
        self.assertEqual(id(qs.only_fields), id(new_qs.only_fields))
        self.assertEqual(qs.only_fields, new_qs.only_fields)
        self.assertEqual(id(qs.order_fields), id(new_qs.order_fields))
        self.assertEqual(qs.order_fields, new_qs.order_fields)
        self.assertEqual(id(qs.return_format), id(new_qs.return_format))
        self.assertEqual(qs.return_format, new_qs.return_format)

        # Set the same values, forcing a new id()
        new_qs.q = Q()
        new_qs.only_fields = ('a', 'b')
        new_qs.order_fields = ('c', 'd')
        new_qs.return_format = QuerySet.NONE

        self.assertNotEqual(id(qs), id(new_qs))
        self.assertEqual(id(qs.folder_collection), id(new_qs.folder_collection))
        self.assertEqual(id(qs._cache), id(new_qs._cache))
        self.assertEqual(qs._cache, new_qs._cache)
        self.assertNotEqual(id(qs.q), id(new_qs.q))
        self.assertEqual(qs.q, new_qs.q)
        self.assertEqual(qs.only_fields, new_qs.only_fields)
        self.assertEqual(qs.order_fields, new_qs.order_fields)
        self.assertEqual(qs.return_format, new_qs.return_format)

        # Set the new values, forcing a new id()
        new_qs.q = Q(foo=5)
        new_qs.only_fields = ('c', 'd')
        new_qs.order_fields = ('e', 'f')
        new_qs.return_format = QuerySet.VALUES

        self.assertNotEqual(id(qs), id(new_qs))
        self.assertEqual(id(qs.folder_collection), id(new_qs.folder_collection))
        self.assertEqual(id(qs._cache), id(new_qs._cache))
        self.assertEqual(qs._cache, new_qs._cache)
        self.assertNotEqual(id(qs.q), id(new_qs.q))
        self.assertNotEqual(qs.q, new_qs.q)
        self.assertNotEqual(id(qs.only_fields), id(new_qs.only_fields))
        self.assertNotEqual(qs.only_fields, new_qs.only_fields)
        self.assertNotEqual(id(qs.order_fields), id(new_qs.order_fields))
        self.assertNotEqual(qs.order_fields, new_qs.order_fields)
        self.assertNotEqual(id(qs.return_format), id(new_qs.return_format))
        self.assertNotEqual(qs.return_format, new_qs.return_format)
