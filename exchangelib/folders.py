# coding=utf-8
import logging

from future.utils import python_2_unicode_compatible
from six import string_types

from .ewsdatetime import EWSDateTime, UTC
from .fields import SimpleField
from .items import Item, CalendarItem, Contact, Message, Task, MeetingRequest, MeetingResponse, MeetingCancellation, \
    ITEM_CLASSES
from .properties import ItemId, EWSElement
from .queryset import QuerySet
from .restriction import Restriction
from .services import IdOnly, FindFolder, GetFolder, FindItem, SHALLOW, DEEP, ITEM_TRAVERSAL_CHOICES, \
    FOLDER_TRAVERSAL_CHOICES, SHAPE_CHOICES
from .util import create_element, value_to_xml_text

string_type = string_types[0]
log = logging.getLogger(__name__)


class FolderId(ItemId):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa579461(v=exchg.150).aspx
    ELEMENT_NAME = 'FolderId'

    __slots__ = ('id', 'changekey')


class CalendarView(EWSElement):
    """
    MSDN: https://msdn.microsoft.com/en-US/library/office/aa564515%28v=exchg.150%29.aspx
    """
    ELEMENT_NAME = 'CalendarView'

    __slots__ = ('start', 'end', 'max_items')

    def __init__(self, start, end, max_items=None):
        self.start = start
        self.end = end
        self.max_items = max_items
        self.clean()

    def clean(self):
        if not isinstance(self.start, EWSDateTime):
            raise ValueError("'start' must be an EWSDateTime")
        if not isinstance(self.end, EWSDateTime):
            raise ValueError("'end' must be an EWSDateTime")
        if not getattr(self.start, 'tzinfo'):
            raise ValueError("'start' must be timezone aware")
        if not getattr(self.end, 'tzinfo'):
            raise ValueError("'end' must be timezone aware")
        if self.end < self.start:
            raise AttributeError("'start' must be before 'end'")
        if self.max_items is not None:
            if not isinstance(self.max_items, int):
                raise ValueError("'max_items' must be an int")
            if self.max_items < 1:
                raise ValueError("'max_items' must be a positive integer")

    @classmethod
    def request_tag(cls):
        return 'm:%s' % cls.ELEMENT_NAME

    def to_xml(self, version):
        self.clean()
        elem = create_element(self.request_tag())
        # Use .set() to not fill up the create_element() cache with unique values
        elem.set('StartDate', value_to_xml_text(self.start.astimezone(UTC)))
        elem.set('EndDate', value_to_xml_text(self.end.astimezone(UTC)))
        if self.max_items is not None:
            elem.set('MaxEntriesReturned', value_to_xml_text(self.max_items))
        return elem


@python_2_unicode_compatible
class Folder(EWSElement):
    DISTINGUISHED_FOLDER_ID = None  # See https://msdn.microsoft.com/en-us/library/office/aa580808(v=exchg.150).aspx
    # Default item type for this folder. See http://msdn.microsoft.com/en-us/library/hh354773(v=exchg.80).aspx
    CONTAINER_CLASS = None
    supported_item_models = ITEM_CLASSES  # The Item types that this folder can contain. Default is all
    LOCALIZED_NAMES = dict()  # A map of (str)locale: (tuple)localized_folder_names
    ITEM_MODEL_MAP = {cls.response_tag(): cls for cls in ITEM_CLASSES}
    FOLDER_FIELDS = (
        SimpleField('folder_id', field_uri='folder:FolderId', value_cls=string_type),
        SimpleField('changekey', field_uri='folder:Changekey', value_cls=string_type),
        SimpleField('name', field_uri='folder:DisplayName', value_cls=string_type),
        SimpleField('folder_class', field_uri='folder:FolderClass', value_cls=string_type),
        SimpleField('total_count', field_uri='folder:TotalCount', value_cls=int),
        SimpleField('unread_count', field_uri='folder:UnreadCount', value_cls=int),
        SimpleField('child_folder_count', field_uri='folder:ChildFolderCount', value_cls=int),
    )
    FOLDER_FIELDS_MAP = {f.name: f for f in FOLDER_FIELDS}

    __slots__ = ('account',) + tuple(f.name for f in FOLDER_FIELDS)

    def __init__(self, account, **kwargs):
        from .account import Account
        assert isinstance(account, Account)
        self.account = account
        for f in self.FOLDER_FIELDS:
            setattr(self, f.name, kwargs.pop(f.name, None))
        if kwargs:
            raise TypeError("%s are invalid keyword arguments for this function" %
                            ', '.join("'%s'" % k for k in kwargs.keys()))
        self.clean()
        log.debug('%s created for %s', self, self.account)

    def clean(self):
        if self.name is None:
            self.name = self.DISTINGUISHED_FOLDER_ID
        if not self.is_distinguished:
            assert self.folder_id
        if self.folder_id:
            assert self.changekey

    @property
    def is_distinguished(self):
        if not self.name or not self.DISTINGUISHED_FOLDER_ID:
            return False
        return self.name.lower() == self.DISTINGUISHED_FOLDER_ID.lower()

    @staticmethod
    def folder_cls_from_container_class(container_class):
        """Returns a reasonable folder class given a container class, e.g. 'IPF.Note'
        """
        return {cls.CONTAINER_CLASS: cls for cls in (Calendar, Contacts, Messages, Tasks)}.get(container_class, Folder)

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
        return cls.ITEM_MODEL_MAP[tag]

    @classmethod
    def allowed_fields(cls):
        fields = set()
        for item_model in cls.supported_item_models:
            fields.update(item_model.ITEM_FIELDS)
        return fields

    @classmethod
    def complex_fields(cls):
        fields = set()
        for item_model in cls.supported_item_models:
            for f in item_model.ITEM_FIELDS:
                if f.is_complex:
                    fields.add(f)
        return fields

    @classmethod
    def get_item_field_by_fieldname(cls, fieldname):
        for item_model in cls.supported_item_models:
            try:
                return item_model.ITEM_FIELDS_MAP[fieldname]
            except KeyError:
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

    def find_items(self, q, shape=IdOnly, depth=SHALLOW, additional_fields=tuple(), order=None, calendar_view=None,
                   page_size=None):
        """
        Private method to call the FindItem service

        :param q: a Q instance containing any restrictions
        :param shape: controls the exact fields returned are governed by. Be aware that complex elements can only be
                      fetched with fetch().
        :param depth: controls the whether to return soft-deleted items or not.
        :param additional_fields: the extra properties we want on the return objects
        :param order: the SortOrder field, if any
        :param calendar_view: a CalendarView instance, if any
        :param page_size: the requested number of items per page
        :return: a generator for the returned item IDs or items
        """
        assert shape in SHAPE_CHOICES
        assert depth in ITEM_TRAVERSAL_CHOICES
        if additional_fields:
            allowed_fields = self.allowed_fields()
            complex_fields = self.complex_fields()
            for f in additional_fields:
                if f not in allowed_fields:
                    raise ValueError("'%s' is not a field on %s" % (f, self.supported_item_models))
                if f in complex_fields:
                    raise ValueError("find_items() does not support field '%s'. Use fetch() instead" % f)
        if calendar_view is not None:
            assert isinstance(calendar_view, CalendarView)
        if page_size is None:
            # Set a sane default
            page_size = FindItem.CHUNKSIZE
        assert isinstance(page_size, int)

        # Build up any restrictions
        if q.is_empty():
            restriction = None
        else:
            restriction = Restriction(q.translate_fields(folder_class=self.__class__))
        log.debug(
            'Finding %s items for %s (shape: %s, depth: %s, additional_fields: %s, restriction: %s)',
            self.DISTINGUISHED_FOLDER_ID,
            self.account,
            shape,
            depth,
            additional_fields,
            restriction.q if restriction else None,
        )
        items = FindItem(folder=self).call(
            additional_fields=additional_fields,
            restriction=restriction,
            order=order,
            shape=shape,
            depth=depth,
            calendar_view=calendar_view,
            page_size=page_size,
        )
        if shape == IdOnly and additional_fields is None:
            for i in items:
                yield i if isinstance(i, Exception) else Item.id_from_xml(i)
        else:
            for i in items:
                if isinstance(i, Exception):
                    yield i
                else:
                    yield self.item_model_from_tag(i.tag).from_xml(elem=i, account=self.account, folder=self)

    def bulk_create(self, items, *args, **kwargs):
        return self.account.bulk_create(folder=self, items=items, *args, **kwargs)

    def fetch(self, *args, **kwargs):
        return self.account.fetch(folder=self, *args, **kwargs)

    def test_access(self):
        """
        Does a simple FindItem to test (read) access to the folder. Maybe the account doesn't exist, maybe the
        service user doesn't have access to the calendar. This will throw the most common errors.
        """
        list(self.filter(subject='DUMMY').values_list('subject'))
        return True

    @classmethod
    def from_xml(cls, elem, account=None):
        # fld_type = re.sub('{.*}', '', elem.tag)
        fld_id_elem = elem.find(FolderId.response_tag())
        fld_id = fld_id_elem.get(FolderId.ID_ATTR)
        changekey = fld_id_elem.get(FolderId.CHANGEKEY_ATTR)
        kwargs = {f.name: f.from_xml(elem) for f in cls.FOLDER_FIELDS if f.name not in ('folder_id', 'changekey')}
        elem.clear()
        return cls(account=account, folder_id=fld_id, changekey=changekey, **kwargs)

    def to_xml(self, version):
        return FolderId(id=self.folder_id, changekey=self.changekey).to_xml(version=version)

    def get_folders(self, shape=IdOnly, depth=DEEP):
        # 'depth' controls whether to return direct children or recurse into sub-folders
        assert shape in SHAPE_CHOICES
        assert depth in FOLDER_TRAVERSAL_CHOICES
        folders = []
        for elem in FindFolder(folder=self).call(
                additional_fields=[f for f in self.FOLDER_FIELDS if f.name not in ('folder_id', 'changekey')],
                shape=shape,
                depth=depth,
                page_size=100,
        ):
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
                folders.append(elem)
                continue
            dummy_fld = Folder.from_xml(elem=elem, account=self.account)  # We use from_xml() only to parse elem
            try:
                folder_cls = self.folder_cls_from_folder_name(folder_name=dummy_fld.name, locale=self.account.locale)
                log.debug('Folder class %s matches localized folder name %s', folder_cls, dummy_fld.name)
            except KeyError:
                folder_cls = self.folder_cls_from_container_class(dummy_fld.folder_class)
                log.debug('Folder class %s matches container class %s (%s)', folder_cls, dummy_fld.folder_class,
                          dummy_fld.name)
            folders.append(folder_cls(account=self.account,
                                      **{f.name: getattr(dummy_fld, f.name) for f in folder_cls.FOLDER_FIELDS}))
        return folders

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
        folders = []
        for elem in GetFolder(account=account).call(
                folder=None,
                distinguished_folder_id=cls.DISTINGUISHED_FOLDER_ID,
                additional_fields=[f for f in cls.FOLDER_FIELDS if f.name not in ('folder_id', 'changekey')],
                shape=shape
        ):
            if isinstance(elem, Exception):
                folders.append(elem)
                continue
            folders.append(cls.from_xml(elem=elem, account=account))
        assert len(folders) == 1
        return folders[0]

    def refresh(self):
        if not self.account:
            raise ValueError('Folder must have an account')
        if not self.folder_id:
            raise ValueError('Folder must have an ID')
        folders = []
        for elem in GetFolder(account=self.account).call(
                folder=self,
                distinguished_folder_id=None,
                additional_fields=[f for f in self.FOLDER_FIELDS if f.name not in ('folder_id', 'changekey')],
                shape=IdOnly
        ):
            if isinstance(elem, Exception):
                folders.append(elem)
                continue
            folders.append(self.from_xml(elem=elem, account=self.account))
        assert len(folders) == 1
        fresh_folder = folders[0]
        assert self.folder_id == fresh_folder.folder_id
        # Apparently, the changekey may get updated
        for f in self.FOLDER_FIELDS:
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
        'da_DK': ('Kalender',),
        'de_DE': ('Kalender',),
        'en_US': ('Calendar',),
        'es_ES': ('Calendario',),
        'fr_CA': ('Calendrier',),
        'nl_NL': ('Agenda',),
        'ru_RU': ('Календарь',),
        'sv_SE': ('Kalender',),
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
        'da_DK': ('Slettet post',),
        'de_DE': ('Gelöschte Elemente',),
        'en_US': ('Deleted Items',),
        'es_ES': ('Elementos eliminados',),
        'fr_CA': ('Éléments supprimés',),
        'nl_NL': ('Verwijderde items',),
        'ru_RU': ('Удаленные',),
        'sv_SE': ('Borttaget',),
    }

    __slots__ = Folder.__slots__


class Messages(Folder):
    CONTAINER_CLASS = 'IPF.Note'
    supported_item_models = (Message, MeetingRequest, MeetingResponse, MeetingCancellation)

    __slots__ = Folder.__slots__


class Drafts(Messages):
    DISTINGUISHED_FOLDER_ID = 'drafts'

    LOCALIZED_NAMES = {
        'da_DK': ('Kladder',),
        'de_DE': ('Entwürfe',),
        'en_US': ('Drafts',),
        'es_ES': ('Borradores',),
        'fr_CA': ('Brouillons',),
        'nl_NL': ('Concepten',),
        'ru_RU': ('Черновики',),
        'sv_SE': ('Utkast',),
    }

    __slots__ = Folder.__slots__


class Inbox(Messages):
    DISTINGUISHED_FOLDER_ID = 'inbox'

    LOCALIZED_NAMES = {
        'da_DK': ('Indbakke',),
        'de_DE': ('Posteingang',),
        'en_US': ('Inbox',),
        'es_ES': ('Bandeja de entrada',),
        'fr_CA': ('Boîte de réception',),
        'nl_NL': ('Postvak IN',),
        'ru_RU': ('Входящие',),
        'sv_SE': ('Inkorgen',),
    }

    __slots__ = Folder.__slots__


class Outbox(Messages):
    DISTINGUISHED_FOLDER_ID = 'outbox'

    LOCALIZED_NAMES = {
        'da_DK': ('Udbakke',),
        'de_DE': ('Kalender',),
        'en_US': ('Outbox',),
        'es_ES': ('Bandeja de salida',),
        'fr_CA': ("Boîte d'envoi",),
        'nl_NL': ('Postvak UIT',),
        'ru_RU': ('Исходящие',),
        'sv_SE': ('Utkorgen',),
    }

    __slots__ = Folder.__slots__


class SentItems(Messages):
    DISTINGUISHED_FOLDER_ID = 'sentitems'

    LOCALIZED_NAMES = {
        'da_DK': ('Sendt post',),
        'de_DE': ('Gesendete Elemente',),
        'en_US': ('Sent Items',),
        'es_ES': ('Elementos enviados',),
        'fr_CA': ('Éléments envoyés',),
        'nl_NL': ('Verzonden items',),
        'ru_RU': ('Отправленные',),
        'sv_SE': ('Skickat',),
    }

    __slots__ = Folder.__slots__


class JunkEmail(Messages):
    DISTINGUISHED_FOLDER_ID = 'junkemail'

    LOCALIZED_NAMES = {
        'da_DK': ('Uønsket e-mail',),
        'de_DE': ('Junk-E-Mail',),
        'en_US': ('Junk E-mail',),
        'es_ES': ('Correo no deseado',),
        'fr_CA': ('Courrier indésirables',),
        'nl_NL': ('Ongewenste e-mail',),
        'ru_RU': ('Нежелательная почта',),
        'sv_SE': ('Skräppost',),
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
        'da_DK': ('Opgaver',),
        'de_DE': ('Aufgaben',),
        'en_US': ('Tasks',),
        'es_ES': ('Tareas',),
        'fr_CA': ('Tâches',),
        'nl_NL': ('Taken',),
        'ru_RU': ('Задачи',),
        'sv_SE': ('Uppgifter',),
    }

    __slots__ = Folder.__slots__


class Contacts(Folder):
    DISTINGUISHED_FOLDER_ID = 'contacts'
    CONTAINER_CLASS = 'IPF.Contact'
    supported_item_models = (Contact,)

    LOCALIZED_NAMES = {
        'da_DK': ('Kontaktpersoner',),
        'de_DE': ('Kontakte',),
        'en_US': ('Contacts',),
        'es_ES': ('Contactos',),
        'fr_CA': ('Contacts',),
        'nl_NL': ('Contactpersonen',),
        'ru_RU': ('Контакты',),
        'sv_SE': ('Kontakter',),
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
