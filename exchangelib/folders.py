# coding=utf-8
from __future__ import unicode_literals

from fnmatch import fnmatch
import logging
from operator import attrgetter

from future.utils import python_2_unicode_compatible
from six import text_type, string_types

from .errors import ErrorAccessDenied, ErrorCannotDeleteObject, ErrorFolderNotFound
from .fields import IntegerField, TextField, DateTimeField, FieldPath, EffectiveRightsField, MailboxField, IdField, \
    EWSElementField
from .items import Item, CalendarItem, Contact, Message, Task, MeetingRequest, MeetingResponse, MeetingCancellation, \
    DistributionList, ITEM_CLASSES, ITEM_TRAVERSAL_CHOICES, SHAPE_CHOICES, IdOnly
from .properties import ItemId, Mailbox, EWSElement, ParentFolderId
from .queryset import QuerySet
from .restriction import Restriction
from .services import FindFolder, GetFolder, FindItem
from .transport import MNS

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
        IdField('id', field_uri=ItemId.ID_ATTR, is_required=True),
        IdField('changekey', field_uri=ItemId.CHANGEKEY_ATTR),
        MailboxField('mailbox'),
    ]

    __slots__ = ItemId.__slots__ + ('mailbox',)


class CalendarView(EWSElement):
    """
    MSDN: https://msdn.microsoft.com/en-US/library/office/aa564515%28v=exchg.150%29.aspx
    """
    ELEMENT_NAME = 'CalendarView'
    NAMESPACE = MNS

    FIELDS = [
        DateTimeField('start', field_uri='StartDate', is_required=True, is_attribute=True),
        DateTimeField('end', field_uri='EndDate', is_required=True, is_attribute=True),
        IntegerField('max_items', field_uri='MaxEntriesReturned', min=1, is_attribute=True),
    ]

    __slots__ = ('start', 'end', 'max_items')

    def clean(self, version=None):
        super(CalendarView, self).clean(version=version)
        if self.end < self.start:
            raise ValueError("'start' must be before 'end'")


@python_2_unicode_compatible
class Folder(EWSElement):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa581334(v=exchg.150).aspx
    """
    DISTINGUISHED_FOLDER_ID = None  # See https://msdn.microsoft.com/en-us/library/office/aa580808(v=exchg.150).aspx
    # Default item type for this folder. See http://msdn.microsoft.com/en-us/library/hh354773(v=exchg.80).aspx
    CONTAINER_CLASS = None
    supported_item_models = ITEM_CLASSES  # The Item types that this folder can contain. Default is all
    LOCALIZED_NAMES = dict()  # A map of (str)locale: (tuple)localized_folder_names
    ITEM_MODEL_MAP = {cls.response_tag(): cls for cls in ITEM_CLASSES}
    FIELDS = [
        IdField('folder_id', field_uri=FolderId.ID_ATTR),
        IdField('changekey', field_uri=FolderId.CHANGEKEY_ATTR),
        EWSElementField('parent_folder_id', field_uri='folder:ParentFolderId', value_cls=ParentFolderId),
        TextField('folder_class', field_uri='folder:FolderClass'),
        TextField('name', field_uri='folder:DisplayName'),
        IntegerField('total_count', field_uri='folder:TotalCount', is_read_only=True),
        IntegerField('child_folder_count', field_uri='folder:ChildFolderCount', is_read_only=True),
        IntegerField('unread_count', field_uri='folder:UnreadCount', is_read_only=True),
        EffectiveRightsField('effective_rights', field_uri='folder:EffectiveRights', is_read_only=True),
    ]

    def __init__(self, **kwargs):
        self.account = kwargs.pop('account', None)
        super(Folder, self).__init__(**kwargs)

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
    def parent(self):
        if self.parent_folder_id.id == self.folder_id:
            # Some folders have a parent that references itself. Avoid circulare references here
            return None
        return self.account.root.get_folder(self.parent_folder_id.id)

    @property
    def children(self):
        for c in self.account.root.get_children(self):
            yield c

    def get_folder_by_name(self, name):
        """Takes a case-sensitive folder name and returns an instance of that folder, if a folder with that name exists
        as a direct or indirect subfolder of this folder.
        """
        import warnings
        warnings.warn('The get_folder_by_name() method is deprecated. Use "[f for f in self.walk() if f.name == name]" '
                      'or "some_folder / \'Sub Folder\'" instead, to find folders by name.')
        matching_folders = [f for f in self.walk() if f.name == name]
        if not matching_folders:
            raise ValueError('No subfolders found with name %s' % name)
        if len(matching_folders) > 1:
            raise ValueError('Multiple subfolders found with name %s' % name)
        return matching_folders[0]

    @property
    def parts(self):
        parts = [self]
        f = self.parent
        while f:
            parts.insert(0, f)
            f = f.parent
        return parts

    @property
    def root(self):
        return self.parts[0]

    @property
    def absolute(self):
        return ''.join('/%s' % p.name for p in self.parts)

    def walk(self):
        for c in self.children:
            yield c
            for f in c.walk():
                yield f

    def glob(self, pattern):
        split_pattern = pattern.rsplit('/', 1)
        head, tail = (split_pattern[0], None) if len(split_pattern) == 1 else split_pattern
        if head == '':
            # We got an absolute path. Restart globbing at root
            for f in self.root.glob(tail or '*'):
                yield f
        elif head == '..':
            # Relative path with reference to parent. Restart globbing at parent
            if not self.parent:
                raise ValueError('Already at top')
            for f in self.parent.glob(tail or '*'):
                yield f
        elif head == '**':
            # Match anything here or in any subfolder at arbitrary depth
            for c in self.walk():
                if fnmatch(c.name, tail or '*'):
                    yield c
        else:
            # Regular pattern
            for c in self.children:
                if not fnmatch(c.name, head):
                    continue
                if tail is None:
                    yield c
                    continue
                for f in c.glob(tail):
                    yield f

    def tree(self):
        """
        Returns a string representation of the folder structure of this folder. Example:

        root
        ├── inbox
        │   └── todos
        └── archive
            ├── Last Job
            ├── exchangelib issues
            └── Mom
        """
        tree = '%s\n' % self.name
        children = list(self.children)
        for i, c in enumerate(sorted(children, key=attrgetter('name')), start=1):
            nodes = c.tree().split('\n')
            for j, node in enumerate(nodes, start=1):
                if i != len(children) and j == 1:
                    # Not the last child, but the first node, which is the name of the child
                    tree += '├── %s\n' % node
                elif i != len(children) and j > 1:
                    # Not the last child, and not name of child
                    tree += '│   %s\n' % node
                elif i == len(children) and j == 1:
                    # Not the last child, but the first node, which is the name of the child
                    tree += '└── %s\n' % node
                else:  # Last child, and not name of child
                    tree += '    %s\n' % node
        return tree.strip()

    @property
    def is_distinguished(self):
        return self.name and self.DISTINGUISHED_FOLDER_ID and self.name.lower() == self.DISTINGUISHED_FOLDER_ID.lower()

    @staticmethod
    def folder_cls_from_container_class(container_class):
        """Returns a reasonable folder class given a container class, e.g. 'IPF.Note'. Don't iterate WELLKNOWN_FOLDERS
        because many folder classes have the same CONTAINER_CLASS.
        """
        for folder_cls in (Messages, Tasks, Calendar, Contacts, GALContacts, RecipientCache):
            if folder_cls.CONTAINER_CLASS == container_class:
                return folder_cls
        return Folder

    @staticmethod
    def folder_cls_from_folder_name(folder_name, locale):
        """Returns the folder class that matches a localized folder name.

        locale is a string, e.g. 'da_DK'
        """
        for folder_cls in WELLKNOWN_FOLDERS:
            localized_names = {s.lower() for s in folder_cls.LOCALIZED_NAMES.get(locale, [])}
            if folder_name.lower() in localized_names:
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

    def find_items(self, q, shape=IdOnly, depth=SHALLOW, additional_fields=None, order_fields=None,
                   calendar_view=None, page_size=None, max_items=None):
        """
        Private method to call the FindItem service

        :param q: a Q instance containing any restrictions
        :param shape: controls whether to return (item_id, chanegkey) tuples or Item objects. If additional_fields is
               non-null, we always return Item objects.
        :param depth: controls the whether to return soft-deleted items or not.
        :param additional_fields: the extra properties we want on the return objects. If None, we'll fetch all fields.
               Be aware that complex elements can only be fetched with fetch().
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
        for f in self.children:
            # TODO: Also delete non-distinguished folders here when we support folder deletion
            try:
                f.wipe()
            except ErrorAccessDenied:
                log.warning('Not allowed to wipe %s', f)
        if isinstance(self, Root):
            log.debug('Skipping root - wiping is not supported')
            return
        log.debug('Wiping folder %s', self)
        for i in self.all().delete():
            if isinstance(i, ErrorCannotDeleteObject):
                log.warning('Not allowed to delete item (%s)', i)
            elif isinstance(i, Exception):
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
        if not kwargs['name']:
            # Some folders are returned with an empty 'DisplayName' element. Assign a default name to them.
            # TODO: Only do this if we actually requested the 'name' field.
            kwargs['name'] = cls.DISTINGUISHED_FOLDER_ID
        elem.clear()
        return cls(account=account, folder_id=fld_id, changekey=changekey, **kwargs)

    def to_xml(self, version):
        self.clean(version=version)
        if self.is_distinguished:
            # Don't add the changekey here. When modifying folder content, we usually don't care if others have changed
            # the folder content since we fetched the changekey.
            return DistinguishedFolderId(
                id=self.DISTINGUISHED_FOLDER_ID,
                mailbox=Mailbox(email_address=self.account.primary_smtp_address)
            ).to_xml(version=version)
        return FolderId(id=self.folder_id, changekey=self.changekey).to_xml(version=version)

    @classmethod
    def supported_fields(cls, version=None):
        return tuple(f for f in cls.FIELDS if f.name not in ('folder_id', 'changekey') and f.supports_version(version))

    def find_folders(self, shape=IdOnly, depth=DEEP):
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
            # folder class we should create the folder with. And many folders share the same 'FolderClass' value, e.g.
            # Inbox and DeletedItems. We want to distinguish between these because otherwise we can't locate the right
            # folders for e.g. Account.inbox and Account.trash.
            #
            # We should be able to just use the name, but apparently default folder names can be renamed to a set of
            # localized names using a PowerShell command:
            #     https://technet.microsoft.com/da-dk/library/dd351103(v=exchg.160).aspx
            #
            # Instead, search for a folder class using the localized name. If none are found, fall back to getting the
            # folder class by the "FolderClass" value.
            if isinstance(elem, Exception):
                yield elem
                continue
            dummy_fld = Folder.from_xml(elem=elem, account=self.account)  # We use from_xml() only to parse elem
            try:
                # TODO: fld_class.LOCALIZED_NAMES is most definitely neither complete nor authoritative
                folder_cls = self.folder_cls_from_folder_name(folder_name=dummy_fld.name, locale=self.account.locale)
                log.debug('Folder class %s matches localized folder name %s', folder_cls, dummy_fld.name)
            except KeyError:
                folder_cls = self.folder_cls_from_container_class(dummy_fld.folder_class)
                log.debug('Folder class %s matches container class %s (%s)', folder_cls, dummy_fld.folder_class,
                          dummy_fld.name)
            yield folder_cls(account=self.account, **{f.name: getattr(dummy_fld, f.name) for f in folder_cls.FIELDS})

    @classmethod
    def get_folders(cls, account, ids, additional_fields=None):
        if additional_fields is None:
            additional_fields = [FieldPath(field=f) for f in cls.supported_fields(version=account.version)]
        for elem in GetFolder(account=account).call(
                folders=ids,
                additional_fields=additional_fields,
                shape=IdOnly
        ):
            if isinstance(elem, Exception):
                raise elem
            yield cls.from_xml(elem=elem, account=account)

    @classmethod
    def get_distinguished(cls, account):
        assert cls.DISTINGUISHED_FOLDER_ID
        folders = list(cls.get_folders(account=account, ids=[cls(account=account, name=cls.DISTINGUISHED_FOLDER_ID)]))
        if not folders:
            raise ErrorFolderNotFound('Could not find distinguished folder %s' % cls.DISTINGUISHED_FOLDER_ID)
        assert len(folders) == 1
        assert isinstance(folders[0], cls)
        return folders[0]

    def refresh(self):
        if not self.account:
            raise ValueError('Folder must have an account')
        if not self.folder_id:
            raise ValueError('Folder must have an ID')
        folders = list(self.get_folders(account=self.account, ids=[self]))
        if not folders:
            raise ErrorFolderNotFound('Folder %s disappeared' % self)
        assert len(folders) == 1
        fresh_folder = folders[0]
        assert self.folder_id == fresh_folder.folder_id
        # Apparently, the changekey may get updated
        for f in self.FIELDS:
            setattr(self, f.name, getattr(fresh_folder, f.name))

    def __truediv__(self, other):
        if other == '..':
            if not self.parent:
                raise ValueError('Already at top')
            return self.parent
        if other == '.':
            return self
        for c in self.children:
            if c.name == other:
                return c
        raise ErrorFolderNotFound("No subfolder with name '%s'" % other)

    # Python 2 requires __div__
    __div__ = __truediv__

    def __repr__(self):
        return self.__class__.__name__ + \
               repr((self.account, self.name, self.total_count, self.unread_count, self.child_folder_count,
                     self.folder_class, self.folder_id, self.changekey))

    def __str__(self):
        return '%s (%s)' % (self.__class__.__name__, self.name)


class Root(Folder):
    DISTINGUISHED_FOLDER_ID = 'root'

    def __init__(self, **kwargs):
        super(Root, self).__init__(**kwargs)
        self._subfolders = None  # See self._folders_map()

    def refresh(self):
        self._subfolders = None
        super(Root, self).refresh()

    @property
    def tois(self):
        # 'Top of Information Store' is a folder available in some Exchange accounts. It usually contains the
        # distinguished folders belonging to the account (inbox, calendar, trash etc.).
        return self / 'Top of Information Store'

    def get_folder(self, folder_id):
        return self._folders_map.get(folder_id, None)

    def get_children(self, folder):
        for f in self._folders_map.values():
            if not f.parent:
                continue
            if f.parent.folder_id == folder.folder_id:
                yield f

    @property
    def _folders_map(self):
        if self._subfolders is not None:
            return self._subfolders

        # Map root, and all subfolders of root, at arbitrary depth by folder ID
        folders_map = {self.folder_id: self}
        try:
            for f in self.find_folders(depth=DEEP):
                if isinstance(f, Exception):
                    raise f
                folders_map[f.folder_id] = f
        except ErrorAccessDenied:
            # We may not have GetFolder or FindFolder access
            pass
        self._subfolders = folders_map
        return folders_map

    def get_default_folder(self, folder_cls):
        # Returns the distinguished folder instance of type folder_cls belonging to this account. If no distinguished
        # folder was found, try as best we can to return the default folder of type 'folder_cls'
        assert folder_cls.DISTINGUISHED_FOLDER_ID
        try:
            # Get the default folder
            log.debug('Testing default %s folder with GetFolder', folder_cls)
            f = folder_cls.get_distinguished(account=self.account)
            return self._folders_map.get(f.folder_id, f)  # Use cached instance if available
        except ErrorAccessDenied:
            # Maybe we just don't have GetFolder access? Try FindItems instead
            log.debug('Testing default %s folder with FindItem', folder_cls)
            f = folder_cls(account=self.account, name=folder_cls.DISTINGUISHED_FOLDER_ID)
            f.test_access()
            return self._folders_map.get(f.folder_id, f)  # Use cached instance if available
        except ErrorFolderNotFound:
            # There's no folder named fld_class.DISTINGUISHED_FOLDER_ID. Try to guess which folder is the default.
            # Exchange makes this unnecessarily difficult.
            log.debug('Searching default %s folder in full folder list', folder_cls)

        candidates = []
        # Try direct children of TOIS first. TOIS might not exist.
        try:
            same_type = [f for f in self.tois.children if type(f) == folder_cls]
            are_distinguished = [f for f in same_type if f.is_distinguished]
            if are_distinguished:
                candidates = are_distinguished
            else:
                localized_names = {s.lower() for s in folder_cls.LOCALIZED_NAMES.get(self.account.locale, [])}
                candidates = [f for f in same_type if f.name.lower() in localized_names]
        except ErrorFolderNotFound:
            pass

        if candidates:
            if len(candidates) > 1:
                raise ValueError(
                    'Multiple possible default %s folders in TOIS: %s'
                    % (folder_cls, [text_type(f.name) for f in candidates])
                )
            return candidates[0]

        # No candidates in TOIS. Try direct children of root.
        same_type = [f for f in self.children if type(f) == folder_cls]
        are_distinguished = [f for f in same_type if f.is_distinguished]
        if are_distinguished:
            candidates = are_distinguished
        else:
            localized_names = {s.lower() for s in folder_cls.LOCALIZED_NAMES.get(self.account.locale, [])}
            candidates = [f for f in same_type if f.name.lower() in localized_names]

        if candidates:
            if len(candidates) > 1:
                raise ValueError('Multiple possible default %s folders in root: %s'
                                 % (folder_cls, [text_type(f.name) for f in candidates]))
            return candidates[0]

        raise ErrorFolderNotFound('No useable default %s folders' % folder_cls)


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


class Messages(Folder):
    CONTAINER_CLASS = 'IPF.Note'
    supported_item_models = (Message, MeetingRequest, MeetingResponse, MeetingCancellation)


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


class Outbox(Messages):
    DISTINGUISHED_FOLDER_ID = 'outbox'

    LOCALIZED_NAMES = {
        'da_DK': (u'Udbakke',),
        'de_DE': (u'Postausgang',),
        'en_US': (u'Outbox',),
        'es_ES': (u'Bandeja de salida',),
        'fr_CA': (u"Boîte d'envoi",),
        'nl_NL': (u'Postvak UIT',),
        'ru_RU': (u'Исходящие',),
        'sv_SE': (u'Utkorgen',),
    }


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


class GALContacts(Contacts):
    DISTINGUISHED_FOLDER_ID = None
    CONTAINER_CLASS = 'IPF.Contact.GalContacts'

    LOCALIZED_NAMES = {}


class RecipientCache(Contacts):
    DISTINGUISHED_FOLDER_ID = 'recipientcache'
    CONTAINER_CLASS = 'IPF.Contact.RecipientCache'

    LOCALIZED_NAMES = {}


class WellknownFolder(Folder):
    # Use this class until we have specific folder implementations
    supported_item_models = ITEM_CLASSES


class AdminAuditLogs(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'adminauditlogs'


class ArchiveDeletedItems(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'archivedeleteditems'


class ArchiveInbox(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'archiveinbox'


class ArchiveMsgFolderRoot(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'archivemsgfolderroot'


class ArchiveRecoverableItemsDeletions(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'archiverecoverableitemsdeletions'


class ArchiveRecoverableItemsPurges(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'archiverecoverableitemspurges'


class ArchiveRecoverableItemsRoot(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'archiverecoverableitemsroot'


class ArchiveRecoverableItemsVersions(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'archiverecoverableitemsversions'


class ArchiveRoot(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'archiveroot'


class Conflicts(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'conflicts'


class ConversationHistory(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'conversationhistory'


class Directory(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'directory'


class Favorites(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'favorites'


class IMContactList(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'imcontactlist'


class Journal(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'journal'


class LocalFailures(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'localfailures'


class MsgFolderRoot(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'msgfolderroot'


class MyContacts(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'mycontacts'


class Notes(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'notes'


class PeopleConnect(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'peopleconnect'


class PublicFoldersRoot(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'publicfoldersroot'


class QuickContacts(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'quickcontacts'


class RecoverableItemsDeletions(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'recoverableitemsdeletions'


class RecoverableItemsPurges(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'recoverableitemspurges'


class RecoverableItemsRoot(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'recoverableitemsroot'


class RecoverableItemsVersions(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'recoverableitemsversions'


class SearchFolders(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'searchfolders'


class ServerFailures(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'serverfailures'


class SyncIssues(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'syncissues'


class ToDoSearch(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'todosearch'


class VoiceMail(WellknownFolder):
    DISTINGUISHED_FOLDER_ID = 'voicemail'


# See http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.wellknownfoldername(v=exchg.80).aspx
# and https://msdn.microsoft.com/en-us/library/office/aa580808(v=exchg.150).aspx
WELLKNOWN_FOLDERS = [
    AdminAuditLogs,
    ArchiveDeletedItems,
    ArchiveInbox,
    ArchiveMsgFolderRoot,
    ArchiveRecoverableItemsDeletions,
    ArchiveRecoverableItemsPurges,
    ArchiveRecoverableItemsRoot,
    ArchiveRecoverableItemsVersions,
    ArchiveRoot,
    Calendar,
    Conflicts,
    Contacts,
    ConversationHistory,
    DeletedItems,
    Directory,
    Drafts,
    Favorites,
    IMContactList,
    Inbox,
    Journal,
    JunkEmail,
    LocalFailures,
    MsgFolderRoot,
    MyContacts,
    Notes,
    Outbox,
    PeopleConnect,
    PublicFoldersRoot,
    QuickContacts,
    RecipientCache,
    RecoverableItemsDeletions,
    RecoverableItemsPurges,
    RecoverableItemsRoot,
    RecoverableItemsVersions,
    Root,
    SearchFolders,
    SentItems,
    ServerFailures,
    SyncIssues,
    Tasks,
    ToDoSearch,
    VoiceMail,
]
