# coding=utf-8
from __future__ import unicode_literals

import logging

from future.utils import python_2_unicode_compatible
from six import string_types

from .fields import IntegerField, TextField, DateTimeField, FieldPath, EffectiveRightsField, MailboxField
from .items import Item, CalendarItem, Contact, Message, Task, MeetingRequest, MeetingResponse, MeetingCancellation, \
    DistributionList, ITEM_CLASSES, ITEM_TRAVERSAL_CHOICES, SHAPE_CHOICES, IdOnly
from .properties import ItemId, Mailbox, EWSElement
from .queryset import QuerySet
from .restriction import Restriction
from .services import FindFolder, GetFolder, FindItem
from .transport import MNS
from .util import create_element, value_to_xml_text

string_type = string_types[0]
log = logging.getLogger(__name__)

# Traversal enums
SHALLOW = 'Shallow'
SOFT_DELETED = 'SoftDeleted'
DEEP = 'Deep'
FOLDER_TRAVERSAL_CHOICES = (SHALLOW, DEEP, SOFT_DELETED)


class FolderId(ItemId):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa579461(v=exchg.150).aspx
    ELEMENT_NAME = 'FolderId'

    __slots__ = ItemId.__slots__


class DistinguishedFolderId(ItemId):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa580808(v=exchg.150).aspx
    ELEMENT_NAME = 'DistinguishedFolderId'

    FIELDS = [
        TextField('id', field_uri=ItemId.ID_ATTR, is_required=True),
        TextField('changekey', field_uri=ItemId.CHANGEKEY_ATTR, is_required=False),
        MailboxField('mailbox', is_required=False)
    ]

    __slots__ = 'id', 'changekey', 'mailbox'

    def to_xml(self, version):
        self.clean(version=version)
        elem = create_element(self.request_tag())
        # Use .set() to not fill up the create_element() cache with unique values
        elem.set(self.ID_ATTR, self.id)
        if self.changekey:
            elem.set(self.CHANGEKEY_ATTR, self.changekey)
        if self.mailbox:
            elem.append(self.get_field_by_fieldname('mailbox').to_xml(self.mailbox, version=version))
        return elem


class CalendarView(EWSElement):
    """
    MSDN: https://msdn.microsoft.com/en-US/library/office/aa564515%28v=exchg.150%29.aspx
    """
    ELEMENT_NAME = 'CalendarView'
    NAMESPACE = MNS

    FIELDS = [
        DateTimeField('start', field_uri='StartDate', is_required=True),
        DateTimeField('end', field_uri='EndDate', is_required=True),
        IntegerField('max_items', field_uri='MaxEntriesReturned', min=1),
    ]

    __slots__ = ('start', 'end', 'max_items')

    def clean(self, version=None):
        super(CalendarView, self).clean(version=version)
        if self.end < self.start:
            raise ValueError("'start' must be before 'end'")

    def to_xml(self, version):
        self.clean(version=version)
        i = create_element(self.request_tag())
        for f in self.supported_fields(version=version):
            value = getattr(self, f.name)
            if value is None:
                continue
            i.set(f.field_uri, value_to_xml_text(value))
        return i


@python_2_unicode_compatible
class Folder(EWSElement):
    DISTINGUISHED_FOLDER_ID = None  # See https://msdn.microsoft.com/en-us/library/office/aa580808(v=exchg.150).aspx
    # Default item type for this folder. See http://msdn.microsoft.com/en-us/library/hh354773(v=exchg.80).aspx
    CONTAINER_CLASS = None
    supported_item_models = ITEM_CLASSES  # The Item types that this folder can contain. Default is all
    LOCALIZED_NAMES = dict()  # A map of (str)locale: (tuple)localized_folder_names
    ITEM_MODEL_MAP = {cls.response_tag(): cls for cls in ITEM_CLASSES}
    FIELDS = [
        TextField('folder_id', field_uri='folder:FolderId', is_searchable=False),
        TextField('changekey', field_uri='folder:Changekey', is_searchable=False),
        TextField('folder_class', field_uri='folder:FolderClass'),
        TextField('name', field_uri='folder:DisplayName'),
        IntegerField('total_count', field_uri='folder:TotalCount', is_read_only=True),
        IntegerField('child_folder_count', field_uri='folder:ChildFolderCount', is_read_only=True),
        IntegerField('unread_count', field_uri='folder:UnreadCount', is_read_only=True),
        EffectiveRightsField('effective_rights', field_uri='folder:EffectiveRights', is_read_only=True),
    ]

    __slots__ = ('account', 'folder_id', 'changekey', 'folder_class', 'name', 'total_count', 'child_folder_count',
                 'unread_count', 'effective_rights')

    def __init__(self, **kwargs):
        self.account = kwargs.pop('account', None)
        super(Folder, self).__init__(**kwargs)
        # pylint: disable=access-member-before-definition
        if self.name is None:
            self.name = self.DISTINGUISHED_FOLDER_ID
        log.debug('%s created for %s', self, self.account)

    def clean(self, version=None):
        super(Folder, self).clean(version=version)
        if self.account is not None:
            from .account import Account
            assert isinstance(self.account, Account)
        if not self.is_distinguished:
            assert self.folder_id
        if self.folder_id:
            assert self.changekey

    @property
    def is_distinguished(self):
        return self.name and self.DISTINGUISHED_FOLDER_ID and self.name.lower() == self.DISTINGUISHED_FOLDER_ID.lower()

    @staticmethod
    def folder_cls_from_container_class(container_class):
        """Returns a reasonable folder class given a container class, e.g. 'IPF.Note'
        """
        try:
            return {cls.CONTAINER_CLASS: cls for cls in (Calendar, Contacts, Messages, Tasks)}[container_class]
        except KeyError:
            return Folder

    @staticmethod
    def folder_cls_from_folder_name(folder_name, locale):
        """Returns the folder class that matches a localized folder name.

        locale is a string, e.g. 'da_DK'
        """
        folder_classes = set(WELLKNOWN_FOLDERS.values())
        for folder_cls in folder_classes:
            for localized_name in folder_cls.LOCALIZED_NAMES.get(locale, []):
                if folder_name.lower() == localized_name.lower():
                    return folder_cls
        raise KeyError()

    @classmethod
    def item_model_from_tag(cls, tag):
        try:
            return cls.ITEM_MODEL_MAP[tag]
        except KeyError:
            item_model = Folder.ITEM_MODEL_MAP[tag]
            raise ValueError('Item type %s was unexpected in a %s folder' % (item_model.__name__, cls.__name__))

    def allowed_fields(self):
        # Return non-ID fields of all item classes allowed in this folder type
        fields = set()
        for item_model in self.supported_item_models:
            fields.update(set(item_model.supported_fields(version=self.account.version if self.account else None)))
        return fields

    def complex_fields(self):
        return {f for f in self.allowed_fields() if f.is_complex}

    def validate_fields(self, fields):
        # Takes a list of fieldnames or FieldPath objects meant for fetching, and checks that they are valid for this
        # folder. Turns them into FieldPath objects and adds internal timezone fields if necessary.
        from .version import EXCHANGE_2010
        allowed_fields = self.allowed_fields()
        fields = list(fields)
        has_start, has_end = False, False
        for i, field_path in enumerate(fields):
            # Allow both FieldPath instances and string field paths as input
            if isinstance(field_path, string_types):
                field_path = FieldPath.from_string(field_path, folder=self)
                fields[i] = field_path
            if not isinstance(field_path, FieldPath):
                raise ValueError("Field '%s' must be a string or FieldPath object" % field_path)
            if field_path.field not in allowed_fields:
                raise ValueError("'%s' is not a valid field on %s" % (field_path.field, self.supported_item_models))
            if field_path.field.name == 'start':
                has_start = True
            elif field_path.field.name == 'end':
                has_end = True

        # For CalendarItem items, we want to inject internal timezone fields. See also CalendarItem.clean()
        if CalendarItem in self.supported_item_models:
            meeting_tz_field, start_tz_field, end_tz_field = CalendarItem.timezone_fields()
            if self.account.version.build < EXCHANGE_2010:
                if has_start or has_end:
                    fields.append(FieldPath(field=meeting_tz_field))
            else:
                if has_start:
                    fields.append(FieldPath(field=start_tz_field))
                if has_end:
                    fields.append(FieldPath(field=end_tz_field))
        return fields

    @classmethod
    def get_item_field_by_fieldname(cls, fieldname):
        for item_model in cls.supported_item_models:
            try:
                return item_model.get_field_by_fieldname(fieldname)
            except ValueError:
                pass
        raise ValueError("Unknown fieldname '%s' on class '%s'" % (fieldname, cls.__name__))

    def all(self):
        return QuerySet(self).all()

    def none(self):
        return QuerySet(self).none()

    def filter(self, *args, **kwargs):
        """
        Finds items in the folder.

        Non-keyword args may be a list of Q instances.

        Optional extra keyword arguments follow a Django-like QuerySet filter syntax (see
           https://docs.djangoproject.com/en/1.10/ref/models/querysets/#field-lookups).

        We don't support '__year' and other date-related lookups. We also don't support '__endswith' or '__iendswith'.

        We support the additional '__not' lookup in place of Django's exclude() for simple cases. For more complicated
        cases you need to create a Q object and use ~Q().

        Examples:

            my_account.inbox.filter(datetime_received__gt=EWSDateTime(2016, 1, 1))
            my_account.calendar.filter(start__range=(EWSDateTime(2016, 1, 1), EWSDateTime(2017, 1, 1)))
            my_account.tasks.filter(subject='Hi mom')
            my_account.tasks.filter(subject__not='Hi mom')
            my_account.tasks.filter(subject__contains='Foo')
            my_account.tasks.filter(subject__icontains='foo')

        'endswith' and 'iendswith' could be emulated by searching with 'contains' or 'icontains' and then
        post-processing items. Fetch the field in question with additional_fields and remove items where the search
        string is not a postfix.
        """
        return QuerySet(self).filter(*args, **kwargs)

    def exclude(self, *args, **kwargs):
        return QuerySet(self).exclude(*args, **kwargs)

    def get(self, *args, **kwargs):
        return QuerySet(self).get(*args, **kwargs)

    def find_items(self, q, shape=IdOnly, depth=SHALLOW, additional_fields=tuple(), order_fields=None,
                   calendar_view=None, page_size=None, max_items=None):
        """
        Private method to call the FindItem service

        :param q: a Q instance containing any restrictions
        :param shape: controls the exact fields returned are governed by. Be aware that complex elements can only be
                      fetched with fetch().
        :param depth: controls the whether to return soft-deleted items or not.
        :param additional_fields: the extra properties we want on the return objects. If None, we'll return
                                  (item_id, changekey) tuples instead of Item obejcts.
        :param order_fields: the SortOrder fields, if any
        :param calendar_view: a CalendarView instance, if any
        :param page_size: the requested number of items per page
        :param max_items: the max number of items to return
        :return: a generator for the returned item IDs or items
        """
        assert shape in SHAPE_CHOICES
        assert depth in ITEM_TRAVERSAL_CHOICES
        if additional_fields:
            allowed_fields = self.allowed_fields()
            complex_fields = self.complex_fields()
            for f in additional_fields:
                if f.field not in allowed_fields:
                    raise ValueError("'%s' is not a valid field on %s" % (f.field.name, self.supported_item_models))
                if f.field in complex_fields:
                    raise ValueError("find_items() does not support field '%s'. Use fetch() instead" % f.field.name)
        if calendar_view is not None:
            assert isinstance(calendar_view, CalendarView)
        if page_size is None:
            # Set a sane default
            page_size = FindItem.CHUNKSIZE
        assert isinstance(page_size, int)

        # Build up any restrictions
        if q.is_empty():
            restriction = None
            query_string = None
        elif q.query_string:
            restriction = None
            query_string = Restriction(q, folder=self)
        else:
            restriction = Restriction(q, folder=self)
            query_string = None
        log.debug(
            'Finding %s items for %s (shape: %s, depth: %s, additional_fields: %s, restriction: %s)',
            self.DISTINGUISHED_FOLDER_ID,
            self.account,
            shape,
            depth,
            additional_fields,
            restriction.q if restriction else None,
        )
        items = FindItem(account=self.account, folders=[self]).call(
            additional_fields=additional_fields,
            restriction=restriction,
            order_fields=order_fields,
            shape=shape,
            query_string=query_string,
            depth=depth,
            calendar_view=calendar_view,
            page_size=page_size,
            max_items=calendar_view.max_items if calendar_view else max_items,
        )
        if shape == IdOnly and additional_fields is None:
            for i in items:
                yield i if isinstance(i, Exception) else Item.id_from_xml(i)
        else:
            for i in items:
                if isinstance(i, Exception):
                    yield i
                else:
                    item = self.item_model_from_tag(i.tag).from_xml(elem=i, account=self.account)
                    item.folder = self
                    yield item

    def bulk_create(self, items, *args, **kwargs):
        return self.account.bulk_create(folder=self, items=items, *args, **kwargs)

    def fetch(self, *args, **kwargs):
        return self.account.fetch(folder=self, *args, **kwargs)

    def wipe(self):
        # Recursively deletes all items in this folder and all subfolders. Use with caution!
        for f in self.get_folders():
            f.wipe()
            # TODO: Also delete non-distinguished folders here when we support folder deletion
        log.debug('Wiping folder %s', self)
        for i in self.all().delete():
            if isinstance(i, Exception):
                raise i

    def test_access(self):
        """
        Does a simple FindItem to test (read) access to the folder. Maybe the account doesn't exist, maybe the
        service user doesn't have access to the calendar. This will throw the most common errors.
        """
        list(self.filter(subject='DUMMY').values_list('subject'))
        return True

    @classmethod
    def from_xml(cls, elem, account):
        # fld_type = re.sub('{.*}', '', elem.tag)
        fld_id_elem = elem.find(FolderId.response_tag())
        fld_id = fld_id_elem.get(FolderId.ID_ATTR)
        changekey = fld_id_elem.get(FolderId.CHANGEKEY_ATTR)
        kwargs = {f.name: f.from_xml(elem=elem, account=account) for f in cls.supported_fields()}
        elem.clear()
        return cls(account=account, folder_id=fld_id, changekey=changekey, **kwargs)

    def to_xml(self, version):
        self.clean(version=version)
        if self.is_distinguished:
            return DistinguishedFolderId(
                id=self.DISTINGUISHED_FOLDER_ID,
                changekey=self.changekey,
                mailbox=Mailbox(email_address=self.account.primary_smtp_address)
            ).to_xml(version=version)
        return FolderId(id=self.folder_id, changekey=self.changekey).to_xml(version=version)

    @classmethod
    def supported_fields(cls, version=None):
        return tuple(f for f in cls.FIELDS if f.name not in ('folder_id', 'changekey') and f.supports_version(version))

    def get_folders(self, shape=IdOnly, depth=DEEP):
        # 'depth' controls whether to return direct children or recurse into sub-folders
        if not self.account:
            raise ValueError('Folder must have an account')
        assert shape in SHAPE_CHOICES
        assert depth in FOLDER_TRAVERSAL_CHOICES
        additional_fields = [FieldPath(field=f) for f in self.supported_fields(version=self.account.version)]
        for elem in FindFolder(account=self.account, folders=[self]).call(
                additional_fields=additional_fields,
                shape=shape,
                depth=depth,
                page_size=100,
                max_items=None,
        ):
            # TODO: Support the Restriction class for folders, too
            # The "FolderClass" element value is the only indication we have in the FindFolder response of which
            # folder class we should create the folder with.
            #
            # We should be able to just use the name, but apparently default folder names can be renamed to a set of
            # localized names using a PowerShell command:
            #     https://technet.microsoft.com/da-dk/library/dd351103(v=exchg.160).aspx
            #
            # Instead, search for a folder class using the localized name. If none are found, fall back to getting the
            # folder class by the "FolderClass" value.
            #
            # TODO: fld_class.LOCALIZED_NAMES is most definitely neither complete nor authoritative
            if isinstance(elem, Exception):
                yield elem
                continue
            dummy_fld = Folder.from_xml(elem=elem, account=self.account)  # We use from_xml() only to parse elem
            try:
                folder_cls = self.folder_cls_from_folder_name(folder_name=dummy_fld.name, locale=self.account.locale)
                log.debug('Folder class %s matches localized folder name %s', folder_cls, dummy_fld.name)
            except KeyError:
                folder_cls = self.folder_cls_from_container_class(dummy_fld.folder_class)
                log.debug('Folder class %s matches container class %s (%s)', folder_cls, dummy_fld.folder_class,
                          dummy_fld.name)
            yield folder_cls(account=self.account, **{f.name: getattr(dummy_fld, f.name) for f in folder_cls.FIELDS})

    def get_folder_by_name(self, name):
        """Takes a case-sensitive folder name and returns an instance of that folder, if a folder with that name exists
        as a direct or indirect subfolder of this folder.
        """
        assert isinstance(name, string_types)
        matching_folders = []
        for f in self.get_folders(depth=DEEP):
            if f.name == name:
                matching_folders.append(f)
        if not matching_folders:
            raise ValueError('No subfolders found with name %s' % name)
        if len(matching_folders) > 1:
            raise ValueError('Multiple subfolders found with name %s' % name)
        return matching_folders[0]

    @classmethod
    def get_distinguished(cls, account, shape=IdOnly):
        assert shape in SHAPE_CHOICES
        additional_fields = [FieldPath(field=f) for f in cls.supported_fields(version=account.version)]
        folders = []
        for elem in GetFolder(account=account).call(
                folders=[cls(account=account)],
                additional_fields=additional_fields,
                shape=shape
        ):
            if isinstance(elem, Exception):
                raise elem
            folder = cls.from_xml(elem=elem, account=account)
            folder.account = account
            folders.append(folder)
        assert len(folders) == 1
        return folders[0]

    def refresh(self):
        if not self.account:
            raise ValueError('Folder must have an account')
        if not self.folder_id:
            raise ValueError('Folder must have an ID')
        additional_fields = [FieldPath(field=f) for f in self.supported_fields(version=self.account.version)]
        folders = []
        for elem in GetFolder(account=self.account).call(
                folders=[self],
                additional_fields=additional_fields,
                shape=IdOnly
        ):
            if isinstance(elem, Exception):
                raise elem
            folder = self.from_xml(elem=elem, account=self.account)
            folders.append(folder)
        assert len(folders) == 1
        fresh_folder = folders[0]
        assert self.folder_id == fresh_folder.folder_id
        # Apparently, the changekey may get updated
        for f in self.FIELDS:
            setattr(self, f.name, getattr(fresh_folder, f.name))

    def __repr__(self):
        return self.__class__.__name__ + \
               repr((self.account, self.name, self.total_count, self.unread_count, self.child_folder_count,
                     self.folder_class, self.folder_id, self.changekey))

    def __str__(self):
        return '%s (%s)' % (self.__class__.__name__, self.name)


class Root(Folder):
    DISTINGUISHED_FOLDER_ID = 'root'

    __slots__ = Folder.__slots__


class Calendar(Folder):
    """
    An interface for the Exchange calendar
    """
    DISTINGUISHED_FOLDER_ID = 'calendar'
    CONTAINER_CLASS = 'IPF.Appointment'
    supported_item_models = (CalendarItem,)

    LOCALIZED_NAMES = {
        'da_DK': (u'Kalender',),
        'de_DE': (u'Kalender',),
        'en_US': (u'Calendar',),
        'es_ES': (u'Calendario',),
        'fr_CA': (u'Calendrier',),
        'nl_NL': (u'Agenda',),
        'ru_RU': (u'Календарь',),
        'sv_SE': (u'Kalender',),
    }

    __slots__ = Folder.__slots__

    def view(self, start, end, max_items=None, *args, **kwargs):
        """ Implements the CalendarView option to FindItem. The difference between filter() and view() is that filter()
        only returns the master CalendarItem for recurring items, while view() unfolds recurring items and returns all
        CalendarItem occurrences as one would normally expect when presenting a calendar.

        Supports the same semantics as filter, except for 'start' and 'end' keyword attributes which are both required
        and behave differently than filter. Here, they denote the start and end of the timespan of the view. All items
        the overlap the timespan are returned (items that end exactly on 'start' are also returned, for some reason).

        EWS does not allow combining CalendarView with search restrictions (filter and exclude).

        'max_items' defines the maximum number of items returned in this view. Optional.
        """
        qs = QuerySet(self).filter(*args, **kwargs)
        qs.calendar_view = CalendarView(start=start, end=end, max_items=max_items)
        return qs


class DeletedItems(Folder):
    DISTINGUISHED_FOLDER_ID = 'deleteditems'
    CONTAINER_CLASS = 'IPF.Note'
    supported_item_models = ITEM_CLASSES

    LOCALIZED_NAMES = {
        'da_DK': (u'Slettet post',),
        'de_DE': (u'Gelöschte Elemente',),
        'en_US': (u'Deleted Items',),
        'es_ES': (u'Elementos eliminados',),
        'fr_CA': (u'Éléments supprimés',),
        'nl_NL': (u'Verwijderde items',),
        'ru_RU': (u'Удаленные',),
        'sv_SE': (u'Borttaget',),
    }

    __slots__ = Folder.__slots__


class Messages(Folder):
    CONTAINER_CLASS = 'IPF.Note'
    supported_item_models = (Message, MeetingRequest, MeetingResponse, MeetingCancellation)

    __slots__ = Folder.__slots__


class Drafts(Messages):
    DISTINGUISHED_FOLDER_ID = 'drafts'

    LOCALIZED_NAMES = {
        'da_DK': (u'Kladder',),
        'de_DE': (u'Entwürfe',),
        'en_US': (u'Drafts',),
        'es_ES': (u'Borradores',),
        'fr_CA': (u'Brouillons',),
        'nl_NL': (u'Concepten',),
        'ru_RU': (u'Черновики',),
        'sv_SE': (u'Utkast',),
    }

    __slots__ = Folder.__slots__


class Inbox(Messages):
    DISTINGUISHED_FOLDER_ID = 'inbox'

    LOCALIZED_NAMES = {
        'da_DK': (u'Indbakke',),
        'de_DE': (u'Posteingang',),
        'en_US': (u'Inbox',),
        'es_ES': (u'Bandeja de entrada',),
        'fr_CA': (u'Boîte de réception',),
        'nl_NL': (u'Postvak IN',),
        'ru_RU': (u'Входящие',),
        'sv_SE': (u'Inkorgen',),
    }

    __slots__ = Folder.__slots__


class Outbox(Messages):
    DISTINGUISHED_FOLDER_ID = 'outbox'

    LOCALIZED_NAMES = {
        'da_DK': (u'Udbakke',),
        'de_DE': (u'Kalender',),
        'en_US': (u'Outbox',),
        'es_ES': (u'Bandeja de salida',),
        'fr_CA': (u"Boîte d'envoi",),
        'nl_NL': (u'Postvak UIT',),
        'ru_RU': (u'Исходящие',),
        'sv_SE': (u'Utkorgen',),
    }

    __slots__ = Folder.__slots__


class SentItems(Messages):
    DISTINGUISHED_FOLDER_ID = 'sentitems'

    LOCALIZED_NAMES = {
        'da_DK': (u'Sendt post',),
        'de_DE': (u'Gesendete Elemente',),
        'en_US': (u'Sent Items',),
        'es_ES': (u'Elementos enviados',),
        'fr_CA': (u'Éléments envoyés',),
        'nl_NL': (u'Verzonden items',),
        'ru_RU': (u'Отправленные',),
        'sv_SE': (u'Skickat',),
    }

    __slots__ = Folder.__slots__


class JunkEmail(Messages):
    DISTINGUISHED_FOLDER_ID = 'junkemail'

    LOCALIZED_NAMES = {
        'da_DK': (u'Uønsket e-mail',),
        'de_DE': (u'Junk-E-Mail',),
        'en_US': (u'Junk E-mail',),
        'es_ES': (u'Correo no deseado',),
        'fr_CA': (u'Courrier indésirables',),
        'nl_NL': (u'Ongewenste e-mail',),
        'ru_RU': (u'Нежелательная почта',),
        'sv_SE': (u'Skräppost',),
    }


class RecoverableItemsDeletions(Folder):
    DISTINGUISHED_FOLDER_ID = 'recoverableitemsdeletions'
    supported_item_models = ITEM_CLASSES

    LOCALIZED_NAMES = {
    }

    __slots__ = Folder.__slots__


class RecoverableItemsRoot(Folder):
    DISTINGUISHED_FOLDER_ID = 'recoverableitemsroot'
    supported_item_models = ITEM_CLASSES

    LOCALIZED_NAMES = {
    }

    __slots__ = Folder.__slots__


class Tasks(Folder):
    DISTINGUISHED_FOLDER_ID = 'tasks'
    CONTAINER_CLASS = 'IPF.Task'
    supported_item_models = (Task,)

    LOCALIZED_NAMES = {
        'da_DK': (u'Opgaver',),
        'de_DE': (u'Aufgaben',),
        'en_US': (u'Tasks',),
        'es_ES': (u'Tareas',),
        'fr_CA': (u'Tâches',),
        'nl_NL': (u'Taken',),
        'ru_RU': (u'Задачи',),
        'sv_SE': (u'Uppgifter',),
    }

    __slots__ = Folder.__slots__


class Contacts(Folder):
    DISTINGUISHED_FOLDER_ID = 'contacts'
    CONTAINER_CLASS = 'IPF.Contact'
    supported_item_models = (Contact, DistributionList)

    LOCALIZED_NAMES = {
        'da_DK': (u'Kontaktpersoner',),
        'de_DE': (u'Kontakte',),
        'en_US': (u'Contacts',),
        'es_ES': (u'Contactos',),
        'fr_CA': (u'Contacts',),
        'nl_NL': (u'Contactpersonen',),
        'ru_RU': (u'Контакты',),
        'sv_SE': (u'Kontakter',),
    }

    __slots__ = Folder.__slots__


class GenericFolder(Folder):
    __slots__ = Folder.__slots__


class WellknownFolder(Folder):
    # Use this class until we have specific folder implementations
    __slots__ = Folder.__slots__


# See http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.wellknownfoldername(v=exchg.80).aspx
WELLKNOWN_FOLDERS = dict([
    ('Calendar', Calendar),
    ('Contacts', Contacts),
    ('DeletedItems', DeletedItems),
    ('Drafts', Drafts),
    ('Inbox', Inbox),
    ('Journal', WellknownFolder),
    ('Notes', WellknownFolder),
    ('Outbox', Outbox),
    ('SentItems', SentItems),
    ('Tasks', Tasks),
    ('MsgFolderRoot', WellknownFolder),
    ('PublicFoldersRoot', WellknownFolder),
    ('Root', Root),
    ('JunkEmail', JunkEmail),
    ('Search', WellknownFolder),
    ('VoiceMail', WellknownFolder),
    ('RecoverableItemsRoot', RecoverableItemsRoot),
    ('RecoverableItemsDeletions', RecoverableItemsDeletions),
    ('RecoverableItemsVersions', WellknownFolder),
    ('RecoverableItemsPurges', WellknownFolder),
    ('ArchiveRoot', WellknownFolder),
    ('ArchiveMsgFolderRoot', WellknownFolder),
    ('ArchiveDeletedItems', WellknownFolder),
    ('ArchiveRecoverableItemsRoot', Folder),
    ('ArchiveRecoverableItemsDeletions', WellknownFolder),
    ('ArchiveRecoverableItemsVersions', WellknownFolder),
    ('ArchiveRecoverableItemsPurges', WellknownFolder),
    ('SyncIssues', WellknownFolder),
    ('Conflicts', WellknownFolder),
    ('LocalFailures', WellknownFolder),
    ('ServerFailures', WellknownFolder),
    ('RecipientCache', WellknownFolder),
    ('QuickContacts', WellknownFolder),
    ('ConversationHistory', WellknownFolder),
    ('ToDoSearch', WellknownFolder),
    ('', GenericFolder),
])
