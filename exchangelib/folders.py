"""
Implements a selection of the folders and folder items found in an Exchange account.

Exchange is very picky about things like the order of XML elements in SOAP requests, so we need to generate XML
automatically instead of taking advantage of Python SOAP libraries and the WSDL file.
"""

from copy import deepcopy
from logging import getLogger
from xml.etree.cElementTree import Element
import re

from .credentials import DELEGATE
from .ewsdatetime import EWSDateTime, EWSTimeZone
from .restriction import Restriction
from .services import TNS, FindItem, IdOnly, SHALLOW, DEEP, DeleteItem, CreateItem, UpdateItem, FindFolder, GetFolder, \
    GetItem
from .util import set_xml_attr, get_xml_attrs, get_xml_attr, set_xml_value

log = getLogger(__name__)

ElementType = type(Element('x'))  # Type is auto-generated inside cElementTree


class Item:
    ITEM_TYPE = 'Item'
    FIELDURI_PREFIX = 'item'

    # 'extern_id' is not a native EWS Item field. We use it for identification when item originates in an external
    # system. The field is implemented as an extended property on the Item.
    ITEM_FIELDS = (
        'item_id',
        'changekey',
        'subject',
        'body',
        'reminder_is_set',
        'categories',
        'extern_id',
    )
    # These are optional
    EXTRA_ITEM_FIELDS = (
        'datetime_created',
        'datetime_sent',
        'datetime_recieved',
        'last_modified_name',
        'last_modified_time',
    )
    ATTR_FIELDURI_MAP = {
        'subject': 'Subject',
        'body': 'Body',
        'reminder_is_set': 'ReminderIsSet',
        'categories': 'Categories',
        'datetime_created': 'DateTimeCreated',
        'datetime_sent': 'DateTimeSent',
        'datetime_recieved': 'DateTimeReceived',
        'last_modified_name': 'LastModifiedName',
        'last_modified_time': 'LastModifiedTime',
        'extern_id': dict(PropertySetId='c11ff724-aa03-4555-9952-8fa248a11c3e', PropertyName='External ID',
                          PropertyType='String'),
    }
    FIELD_TYPE_MAP = {
        'subject': str,
        'body': str,
        'reminder_is_set': bool,
        'categories': list,
        'datetime_created': EWSDateTime,
        'datetime_sent': EWSDateTime,
        'datetime_recieved': EWSDateTime,
        'last_modified_name': str,
        'last_modified_time': EWSDateTime,
        'extern_id': dict,
    }

    def __init__(self, **kwargs):
        for k in Item.ITEM_FIELDS + Item.EXTRA_ITEM_FIELDS:
            default = False if k == 'reminder_is_set' else None
            setattr(self, k, kwargs.pop(k, default))
        for k, v in kwargs.items():
            raise TypeError("'%s' is an invalid keyword argument for this function" % k)

    @classmethod
    def fieldnames(cls, with_extra=False):
        # Return non-ID field names
        if with_extra:
            return cls.ITEM_FIELDS[2:] + cls.EXTRA_ITEM_FIELDS
        return cls.ITEM_FIELDS[2:]

    @classmethod
    def fielduri_for_field(cls, fieldname):
        try:
            field_uri = cls.ATTR_FIELDURI_MAP[fieldname]
            if isinstance(field_uri, str):
                return '%s:%s' % (cls.FIELDURI_PREFIX, field_uri)
            return field_uri
        except KeyError:
            raise ValueError("No fielduri defined for fieldname '%s'" % fieldname)

    @classmethod
    def request_xml_elem_for_field(cls, fieldname):
        try:
            field_uri = cls.ATTR_FIELDURI_MAP[fieldname]
            if isinstance(field_uri, dict):
                extended_property = Element('t:ExtendedProperty')
                extended_property_field_uri = Element('t:ExtendedFieldURI', **field_uri)
                extended_property.append(extended_property_field_uri)
                return extended_property
            else:
                elem = Element('t:%s' % field_uri)
                if fieldname == 'body':
                    elem.set('BodyType', 'Text')
                return elem
        except KeyError:
            raise ValueError("No fielduri defined for fieldname '%s'" % fieldname)

    @classmethod
    def response_xml_elem_for_field(cls, fieldname):
        try:
            return '{%s}%s' % (TNS, cls.ATTR_FIELDURI_MAP[fieldname])
        except KeyError:
            raise ValueError("No fielduri defined for fieldname '%s'" % fieldname)

    @classmethod
    def type_for_field(cls, fieldname):
        try:
            return cls.FIELD_TYPE_MAP[fieldname]
        except KeyError:
            raise ValueError("No type defined for fieldname '%s'" % fieldname)

    @classmethod
    def additional_property_fields(cls, with_extra=False):
        fields = []
        fielduri = Element('t:FieldURI')
        extended_fielduri = Element('t:ExtendedFieldURI')
        for f in cls.fieldnames(with_extra=with_extra):
            field_uri = cls.fielduri_for_field(f)
            if isinstance(field_uri, str):
                f = deepcopy(fielduri)
                f.set('FieldURI', field_uri)
                fields.append(f)
            else:
                f = deepcopy(extended_fielduri)
                for k, v in field_uri.items():
                    f.set(k, v)
                fields.append(f)
        return fields

    @classmethod
    def id_from_xml(cls, elem):
        id_elem = elem.find('{%s}ItemId' % TNS)
        return id_elem.get('Id'), id_elem.get('ChangeKey')

    @classmethod
    def from_xml(cls, elem, with_extra=False):
        item_id, changekey = cls.id_from_xml(elem)
        kwargs = {}
        extended_properties = elem.findall('{%s}ExtendedProperty' % TNS)
        for fieldname in cls.fieldnames(with_extra=with_extra):
            t = cls.type_for_field(fieldname)
            if t == EWSDateTime:
                str_val = get_xml_attr(elem, cls.response_xml_elem_for_field(fieldname))
                kwargs[fieldname] = EWSDateTime.from_string(str_val)
            elif t == bool:
                val = get_xml_attr(elem, cls.response_xml_elem_for_field(fieldname))
                kwargs[fieldname] = True if val == 'true' else False
            elif t == str:
                kwargs[fieldname] = get_xml_attr(elem, cls.response_xml_elem_for_field(fieldname)) or ''
            elif t == list:
                iter_elem = elem.find(cls.response_xml_elem_for_field(fieldname))
                kwargs[fieldname] = get_xml_attrs(iter_elem, '{%s}String' % TNS) if iter_elem is not None else None
            elif t == dict:
                field_uri = cls.fielduri_for_field(fieldname)
                extended_field_value = None
                for e in extended_properties:
                    extended_field_uri = e.find('{%s}ExtendedFieldURI' % TNS)
                    match = True
                    for k, v in field_uri.items():
                        if extended_field_uri.get(k) != v:
                            match = False
                            break
                    if match:
                        extended_field_value = get_xml_attr(e, '{%s}Value' % TNS) or ''
                        break
                kwargs[fieldname] = extended_field_value
            else:
                assert False, 'Field %s type %s not supported' % (fieldname, t)
        return cls(item_id=item_id, changekey=changekey, **kwargs)


class Folder:
    DISTINGUISHED_FOLDER_ID = None
    CONTAINER_CLASS = None  # See http://msdn.microsoft.com/en-us/library/hh354773(v=exchg.80).aspx
    item_model = Item

    def __init__(self, account, name=None, folder_class=None, folder_id=None, changekey=None):
        self.account = account
        self.name = name or self.DISTINGUISHED_FOLDER_ID
        self.folder_class = folder_class
        self.folder_id = folder_id
        self.changekey = changekey
        if not self.is_distinguished:
            assert self.folder_id
        if self.folder_id:
            assert self.changekey
        self.with_extra_fields = False
        log.debug('%s created for %s', self.__class__.__name__, account)

    @property
    def is_distinguished(self):
        return self.name.lower() == self.DISTINGUISHED_FOLDER_ID

    @classmethod
    def attr_to_fielduri(cls, fieldname):
        return cls.item_model.fielduri_for_field(fieldname)

    @classmethod
    def attr_to_request_xml_elem(cls, fieldname):
        return cls.item_model.request_xml_elem_for_field(fieldname)

    @classmethod
    def attr_to_response_xml_elem(cls, fieldname):
        return cls.item_model.response_xml_elem_for_field(fieldname)

    def find_items(self, start=None, end=None, categories=None, shape=IdOnly, depth=SHALLOW):
        """
        Finds all items in the folder, optionally restricted by start- and enddates and a list of categories
        """
        log.debug(
            'Finding %s items for %s from %s to %s with cats %s shape %s',
            self.DISTINGUISHED_FOLDER_ID,
            self.account,
            start,
            end,
            categories,
            shape
        )
        xml_func = self.item_model.id_from_xml if shape == IdOnly else self.item_model.from_xml
        # Define the extra properties we want on the return objects. 'body' field can only be fetched with GetItem.
        additional_fields = ['item:Categories'] if categories else None
        # Define any search restrictions we want to set on the search.
        # TODO Filtering by category doesn't work on Exchange 2010, returning "ErrorContainsFilterWrongType:
        # The Contains filter can only be used for string properties." Fall back to filtering after getting all items
        # instead. This may be a legal problem because we get ALL items, including private appointments.
        restriction = Restriction.from_params(self.DISTINGUISHED_FOLDER_ID, start=start, end=end)
        items = FindItem(self.account.protocol).call(folder=self, additional_fields=additional_fields,
                                                     restriction=restriction, shape=shape, depth=depth)
        if not categories:
            log.debug('Found %s items', len(items))
            return list(map(xml_func, items))

        # Filter for category. Searching for categories only works with 'Or' operator on Exchange 2007, so we need to
        # ignore items with only some but not all categories present.
        filtered_items = []
        categoryset = set(categories)
        for item in items:
            if not isinstance(item, ElementType):
                status, error = item
                assert not status
                log.warning('Error fetching items: %s', error)
                continue
            cats = item.find(self.attr_to_response_xml_elem('categories'))
            if cats is not None:
                item_cats = get_xml_attrs(cats, '{%s}String' % TNS)
                if categoryset.issubset(set(item_cats)):
                    filtered_items.append(item)
        log.debug('%s of %s items match category', len(filtered_items), len(items))
        return list(map(xml_func, filtered_items))

    def add_items(self, items):
        """
        Creates new items in the folder. 'items' is an iterable of Item objects. Returns a list of (id, changekey)
        tuples in the same order as the input.
        """
        assert len(items)
        log.debug('Adding %s calendar items', len(items))
        return list(map(self.item_model.id_from_xml, CreateItem(self.account.protocol).call(folder=self, items=items)))

    def delete_items(self, ids):
        """
        Deletes items in the folder. 'ids' is an iterable of either (item_id, changekey) tuples or Item objects.
        """
        assert len(ids)
        return DeleteItem(self.account.protocol).call(folder=self, ids=ids)

    def update_items(self, items):
        """
        Updates items in the folder. 'items' is an iterable of tuples containing two elements:

            1. either an (item_id, changekey) tuple or an Item object
            2. a dict containing the Item attributes to change

        """
        assert len(items)
        return list(map(self.item_model.id_from_xml, UpdateItem(self.account.protocol).call(folder=self, items=items)))

    def get_items(self, ids):
        # get_xml() uses self.with_extra_fields. Pass this to from_xml()
        assert len(ids)
        return list(map(
            lambda i: self.item_model.from_xml(i, self.with_extra_fields),
            GetItem(self.account.protocol).call(folder=self, ids=ids))
        )

    def test_access(self):
        """
        Does a simple FindItem to test (read) access to the mailbox. Maybe the account doesn't exist, maybe the
        service user doesn't have access to the calendar.
        """
        now = EWSDateTime.now(tz=EWSTimeZone.timezone('UTC'))
        restriction = Restriction.from_params(self.DISTINGUISHED_FOLDER_ID, start=now, end=now)
        FindItem(self.account.protocol).call(folder=self, restriction=restriction, shape=IdOnly)
        return True

    def folderid_xml(self):
        if self.folder_id:
            assert self.changekey
            return Element('t:FolderId', Id=self.folder_id, ChangeKey=self.changekey)
        else:
            # Only use distinguished ID if we don't have the folder ID
            distinguishedfolderid = Element('t:DistinguishedFolderId', Id=self.DISTINGUISHED_FOLDER_ID)
            if self.account.access_type == DELEGATE:
                mailbox = Element('t:Mailbox')
                set_xml_attr(mailbox, 't:EmailAddress', self.account.primary_smtp_address)
                distinguishedfolderid.append(mailbox)
            return distinguishedfolderid

    def get_xml(self, ids):
        assert len(ids)
        assert isinstance(ids[0], (tuple, Item))
        # This list should be configurable. 'body' element can only be fetched with GetItem.
        # CalendarItem.from_xml() specifies the items we currently expect. For full list, see
        # https://msdn.microsoft.com/en-us/library/office/aa494315(v=exchg.150).aspx
        log.debug(
            'Getting %s %s items for %s',
            len(ids),
            self.DISTINGUISHED_FOLDER_ID,
            self.account
        )
        getitem = Element('m:%s' % GetItem.SERVICE_NAME)
        itemshape = Element('m:ItemShape')
        set_xml_attr(itemshape, 't:BaseShape', IdOnly)
        additional_properties = self.item_model.additional_property_fields(with_extra=self.with_extra_fields)
        if additional_properties:
            additionalproperties = Element('t:AdditionalProperties')
            for p in additional_properties:
                additionalproperties.append(p)
            itemshape.append(additionalproperties)
        getitem.append(itemshape)
        itemids = Element('m:ItemIds')
        itemid = Element('t:ItemId')
        # copy.deepcopy() is an order of magnitude faster than having
        # Element() inside the loop. It matters because 'ids' may
        # be a very large list.
        for item in ids:
            item_id, changekey = item if isinstance(item, tuple) else (item.item_id, item.changekey)
            i = deepcopy(itemid)
            i.set('Id', item_id)
            i.set('ChangeKey', changekey)
            itemids.append(i)
        getitem.append(itemids)
        return getitem

    def create_xml(self, items):
        assert len(items)
        # Takes an account name, a folder name, a list of Calendar.Item obejcts and a function to convert items to XML
        # Elements
        if isinstance(self, Calendar):
            createitem = Element('m:%s' % CreateItem.SERVICE_NAME, SendMeetingInvitations='SendToNone')
        elif isinstance(self, Inbox):
            createitem = Element('m:%s' % CreateItem.SERVICE_NAME, MessageDisposition='SaveOnly')
        else:
            createitem = Element('m:%s' % CreateItem.SERVICE_NAME)

        saveditemfolderid = Element('m:SavedItemFolderId')
        saveditemfolderid.append(self.folderid_xml())
        createitem.append(saveditemfolderid)

        myitems = Element('m:Items')
        for item in self._create_item_xml(items):
            myitems.append(item)
        createitem.append(myitems)
        return createitem

    def _create_item_xml(self, items):
        raise NotImplementedError()

    def delete_xml(self, ids):
        assert len(ids)
        assert isinstance(ids[0], (tuple, Item))
        # Prepare reuseable Element objects
        if isinstance(self, Calendar):
            deleteitem = Element('m:%s' % DeleteItem.SERVICE_NAME, DeleteType='HardDelete',
                                 SendMeetingCancellations='SendToNone')
        elif isinstance(self, Inbox):
            deleteitem = Element('m:%s' % DeleteItem.SERVICE_NAME, DeleteType='HardDelete')
        elif isinstance(self, Tasks):
            deleteitem = Element('m:%s' % DeleteItem.SERVICE_NAME, DeleteType='HardDelete',
                                 AffectedTaskOccurrences='SpecifiedOccurrenceOnly')
        else:
            deleteitem = Element('m:%s' % DeleteItem.SERVICE_NAME, DeleteType='HardDelete')
        if self.account.version.major_version >= 15:
            deleteitem.set('SuppressReadReceipts', 'true')

        itemids = Element('m:ItemIds')
        itemid = Element('t:ItemId')

        # copy.deepcopy() is an order of magnitude faster than having
        # Element() inside the loop. It matters because 'ids' may
        # be a very large list.
        for item in ids:
            item_id, changekey = item if isinstance(item, tuple) else (item.item_id, item.changekey)
            i = deepcopy(itemid)
            i.set('Id', item_id)
            i.set('ChangeKey', changekey)
            itemids.append(i)
        deleteitem.append(itemids)
        return deleteitem

    def update_xml(self, items):
        assert len(items)
        assert isinstance(items[0], tuple)
        assert isinstance(items[0][0], (tuple, Item))
        assert isinstance(items[0][1], dict)
        # Prepare reuseable Element objects
        if isinstance(self, Calendar):
            updateitem = Element('m:%s' % UpdateItem.SERVICE_NAME, ConflictResolution='AutoResolve',
                                 SendMeetingInvitationsOrCancellations='SendToNone')
        elif isinstance(self, Inbox):
            updateitem = Element('m:%s' % UpdateItem.SERVICE_NAME, ConflictResolution='AutoResolve',
                                 MessageDisposition='SaveOnly')
        else:
            updateitem = Element('m:%s' % UpdateItem.SERVICE_NAME, ConflictResolution='AutoResolve')
        if self.account.version.major_version >= 15:
            updateitem.set('SuppressReadReceipts', 'true')

        itemchanges = Element('m:ItemChanges')
        itemchange = Element('t:ItemChange')
        itemid = Element('t:ItemId')
        updates = Element('t:Updates')
        setitemfield = Element('t:SetItemField')
        deleteitemfield = Element('t:DeleteItemField')
        fielduri = Element('t:FieldURI')
        extended_fielduri = Element('t:ExtendedFieldURI')
        folderitem = Element('t:%s' % self.item_model.ITEM_TYPE)

        # copy.deepcopy() is an order of magnitude faster than having
        # Element() inside the loop. It matters because 'ids' may
        # be a very large list.
        for item, update_dict in items:
            assert len(update_dict)
            i_itemchange = deepcopy(itemchange)
            item_id, changekey = item if isinstance(item, tuple) else (item.item_id, item.changekey)
            i_itemid = deepcopy(itemid)
            i_itemid.set('Id', item_id)
            i_itemid.set('ChangeKey', changekey)
            i_itemchange.append(i_itemid)
            i_updates = deepcopy(updates)
            meeting_timezone_added = False
            for fieldname, val in update_dict.items():
                field_uri = self.attr_to_fielduri(fieldname)
                if isinstance(field_uri, str):
                    i_fielduri = deepcopy(fielduri)
                    i_fielduri.set('FieldURI', field_uri)
                else:
                    i_fielduri = deepcopy(extended_fielduri)
                    for k, v in field_uri.items():
                        i_fielduri.set(k, v)
                if val is None:
                    # A value of None means we want to remove this field from the item
                    i_deleteitemfield = deepcopy(deleteitemfield)
                    i_deleteitemfield.append(i_fielduri)
                    i_updates.append(i_deleteitemfield)
                    continue
                i_setitemfield = deepcopy(setitemfield)
                i_setitemfield.append(i_fielduri)
                i_folderitem = deepcopy(folderitem)
                i_value = self.attr_to_request_xml_elem(fieldname)
                if isinstance(val, EWSDateTime):
                    i_folderitem.append(set_xml_value(i_value, val))
                    i_setitemfield.append(i_folderitem)
                    i_updates.append(i_setitemfield)

                    # Always set timezone explicitly when updating date fields. Exchange 2007 wants "MeetingTimeZone"
                    # instead of explicit timezone on each datetime field.
                    if self.account.version.major_version < 14:
                        if meeting_timezone_added:
                            # Let's hope that we're not changing timezone, or that both 'start' and 'end' are supplied.
                            # Exchange 2007 doesn't support different timezone on start and end.
                            continue
                        i_setitemfield_tz = deepcopy(setitemfield)
                        i_fielduri = deepcopy(fielduri)
                        i_fielduri.set('FieldURI', 'calendar:MeetingTimeZone')
                        i_setitemfield_tz.append(i_fielduri)
                        i_folderitem = deepcopy(folderitem)
                        meeting_timezone = Element('t:MeetingTimeZone', TimeZoneName=val.tzinfo.ms_id)
                        i_folderitem.append(meeting_timezone)
                        i_setitemfield_tz.append(i_folderitem)
                        i_updates.append(i_setitemfield_tz)
                        meeting_timezone_added = True
                    else:
                        i_setitemfield_tz = deepcopy(setitemfield)
                        i_fielduri = deepcopy(fielduri)
                        if fieldname == 'start':
                            field_uri = 'calendar:StartTimeZone'
                            timezone_element = Element('t:StartTimeZone', Id=val.tzinfo.ms_id, Name=val.tzinfo.ms_name)
                        elif fieldname == 'end':
                            field_uri = 'calendar:EndTimeZone'
                            timezone_element = Element('t:EndTimeZone', Id=val.tzinfo.ms_id, Name=val.tzinfo.ms_name)
                        else:
                            assert False, 'Cannot set timezone for field %s' % fieldname
                        i_fielduri.set('FieldURI', field_uri)
                        i_setitemfield_tz.append(i_fielduri)
                        i_folderitem = deepcopy(folderitem)
                        i_folderitem.append(timezone_element)
                        i_setitemfield_tz.append(i_folderitem)
                        i_updates.append(i_setitemfield_tz)
                else:
                    if isinstance(field_uri, str):
                        i_folderitem.append(set_xml_value(i_value, val))
                    else:
                        i_folderitem.append(set_xml_attr(i_value, 't:Value', val))
                    i_setitemfield.append(i_folderitem)
                    i_updates.append(i_setitemfield)
            i_itemchange.append(i_updates)
            itemchanges.append(i_itemchange)
        updateitem.append(itemchanges)
        return updateitem

    @classmethod
    def from_xml(cls, account, elem):
        fld_type = re.sub('{.*}', '', elem.tag)
        fld_class = FOLDER_TYPE_MAP[fld_type]
        fld_id_elem = elem.find('{%s}FolderId' % TNS)
        fld_id = fld_id_elem.get('Id')
        changekey = fld_id_elem.get('ChangeKey')
        display_name = get_xml_attr(elem, '{%s}DisplayName' % TNS)
        folder_class = get_xml_attr(elem, '{%s}FolderClass' % TNS)
        return fld_class(account=account, name=display_name, folder_class=folder_class, folder_id=fld_id,
                         changekey=changekey)

    def get_folders(self, shape=IdOnly, depth=DEEP):
        folders = []
        for elem in FindFolder(self.account.protocol).call(
                folder=self,
                additional_fields=['folder:DisplayName', 'folder:FolderClass'],
                shape=shape,
                depth=depth
        ):
            folders.append(self.from_xml(self.account, elem))
        return folders

    def get_folder(self, shape=IdOnly):
        folders = []
        for elem in GetFolder(self.account.protocol).call(
                folder=self,
                additional_fields=['folder:DisplayName', 'folder:FolderClass'],
                shape=shape
        ):
            folders.append(self.from_xml(self.account, elem))
        assert len(folders) == 1
        return folders[0]

    def __repr__(self):
        return self.__class__.__name__ + \
               repr((self.account, self.name, self.folder_class, self.folder_id, self.changekey))

    def __str__(self):
        return '%s (%s)' % (self.__class__.__name__, self.name)


class Root(Folder):
    DISTINGUISHED_FOLDER_ID = 'root'


class CalendarItem(Item):
    """
    Models a calendar item. Not all attributes are supported. See full list at
    https://msdn.microsoft.com/en-us/library/office/aa564765(v=exchg.150).aspx
    """
    ITEM_TYPE = 'CalendarItem'
    SUBJECT_MAXLENGTH = 255
    LOCATION_MAXLENGTH = 255
    FIELDURI_PREFIX = 'calendar'
    ITEM_FIELDS = (
        'start',
        'end',
        'location',
        'organizer',
        'legacy_free_busy_status',
    )
    ATTR_FIELDURI_MAP = {
        'start': 'Start',
        'end': 'End',
        'location': 'Location',
        'organizer': 'Organizer',
        'legacy_free_busy_status': 'LegacyFreeBusyStatus',
    }
    FIELD_TYPE_MAP = {
        'start': EWSDateTime,
        'end': EWSDateTime,
        'location': str,
        'organizer': str,
        'legacy_free_busy_status': str,
    }

    @classmethod
    def fieldnames(cls, with_extra=False):
        return cls.ITEM_FIELDS + Item.fieldnames(with_extra=with_extra)

    @classmethod
    def fielduri_for_field(cls, fieldname):
        try:
            field_uri = cls.ATTR_FIELDURI_MAP[fieldname]
            if isinstance(field_uri, str):
                return '%s:%s' % (cls.FIELDURI_PREFIX, field_uri)
            return field_uri
        except KeyError:
            return Item.fielduri_for_field(fieldname)

    @classmethod
    def request_xml_elem_for_field(cls, fieldname):
        try:
            field_uri = cls.ATTR_FIELDURI_MAP[fieldname]
            if isinstance(field_uri, dict):
                extended_property = Element('t:ExtendedProperty')
                extended_property_field_uri = Element('t:ExtendedFieldURI', **field_uri)
                extended_property.append(extended_property_field_uri)
                return extended_property
            else:
                return Element('t:%s' % field_uri)
        except KeyError:
            return Item.request_xml_elem_for_field(fieldname)

    @classmethod
    def response_xml_elem_for_field(cls, fieldname):
        try:
            return '{%s}%s' % (TNS, cls.ATTR_FIELDURI_MAP[fieldname])
        except KeyError:
            return Item.response_xml_elem_for_field(fieldname)

    @classmethod
    def type_for_field(cls, fieldname):
        try:
            return cls.FIELD_TYPE_MAP[fieldname]
        except KeyError:
            return Item.type_for_field(fieldname)

    def __init__(self, **kwargs):
        for k in self.ITEM_FIELDS:
            default = 'Busy' if k == 'legacy_free_busy_status' else None
            v = kwargs.pop(k, default)
            if k in ('start', 'end') and v and not getattr(v, 'tzinfo'):
                raise ValueError("'%s' must be timezone aware")
            setattr(self, k, v)
        super().__init__(**kwargs)

    def __repr__(self):
        return self.__class__.__name__ + repr(getattr(self, k) for k in (Item.ITEM_FIELDS + self.ITEM_FIELDS))

    def __str__(self):
        return '''\
ItemId: %(item_id)s
Changekey: %(changekey)s
Subject: %(subject)s
Start: %(start)s
End: %(end)s
Location: %(location)s
Body: %(body)s
Has reminder: %(reminder_is_set)s
Categories: %(categories)s
Extern ID: %(extern_id)s''' % self.__dict__


class Calendar(Folder):
    """
    An interface for the Exchange calendar
    """
    DISTINGUISHED_FOLDER_ID = 'calendar'
    CONTAINER_CLASS = 'IPF.Appointment'
    item_model = CalendarItem

    # These must be capitalized
    # TODO: This is most definitely not either complete or authoritative
    LOCALIZED_NAMES = (
        'Kalender',
    )

    def _create_item_xml(self, items):
        """
        Takes an array of Calendar.Item objects and generates the XML.
        """
        calendar_items = []
        # Prepare Element objects for CalendarItem creation
        calendar_item = Element('t:%s' % self.item_model.ITEM_TYPE)
        subject = self.attr_to_request_xml_elem('subject')
        body = self.attr_to_request_xml_elem('body')
        reminder = self.attr_to_request_xml_elem('reminder_is_set')
        start = self.attr_to_request_xml_elem('start')
        end = self.attr_to_request_xml_elem('end')
        busystatus = self.attr_to_request_xml_elem('legacy_free_busy_status')
        location = self.attr_to_request_xml_elem('location')
        organizer = self.attr_to_request_xml_elem('organizer')
        categories = self.attr_to_request_xml_elem('categories')
        extern_id = self.attr_to_request_xml_elem('extern_id')

        # Using copy.deepcopy() increases performance drastically on large arrays
        for item in items:
            # WARNING: The order of addition of XML elements is VERY important. Exchange expects XML elements in a
            # specific, non-documented order and will fail with meaningless errors if the order is wrong.
            i = deepcopy(calendar_item)
            i.append(set_xml_value(deepcopy(subject), item.subject))
            if item.body:
                i.append(set_xml_value(deepcopy(body), item.body))
            if item.categories:
                i.append(set_xml_value(deepcopy(categories), item.categories))
            i.append(set_xml_value(deepcopy(reminder), item.reminder_is_set))
            if item.extern_id is not None:
                i.append(set_xml_attr(deepcopy(extern_id), 't:Value', item.extern_id))
            i.append(set_xml_value(deepcopy(start), item.start))
            i.append(set_xml_value(deepcopy(end), item.end))
            i.append(set_xml_value(deepcopy(busystatus), item.legacy_free_busy_status))
            if item.location:
                i.append(set_xml_value(deepcopy(location), item.location))
            if item.organizer:
                i.append(set_xml_value(deepcopy(organizer), item.organizer))
            if self.account.version.major_version < 14:
                meeting_timezone = Element('t:MeetingTimeZone', TimeZoneName=item.start.tzinfo.ms_id)
                i.append(meeting_timezone)
            else:
                start_timezone = Element('t:StartTimeZone', Id=item.start.tzinfo.ms_id, Name=item.start.tzinfo.ms_name)
                i.append(start_timezone)
                end_timezone = Element('t:EndTimeZone', Id=item.end.tzinfo.ms_id, Name=item.end.tzinfo.ms_name)
                i.append(end_timezone)
            calendar_items.append(i)
        return calendar_items


class Message(Item):
    ITEM_TYPE = 'Message'


class Inbox(Folder):
    ITEM_TYPE = 'Message'
    DISTINGUISHED_FOLDER_ID = 'inbox'
    CONTAINER_CLASS = 'IPF.Note'


class Task(Item):
    ITEM_TYPE = 'Task'


class Tasks(Folder):
    DISTINGUISHED_FOLDER_ID = 'tasks'
    CONTAINER_CLASS = 'IPF.Task'


class Contact(Item):
    ITEM_TYPE = 'Contact'


class Contacts(Folder):
    DISTINGUISHED_FOLDER_ID = 'contacts'
    CONTAINER_CLASS = 'IPF.Contact'


class GenericFolder(Folder):
    pass


class WellknownFolder(Folder):
    # Use this class until we have specific folder implementations
    pass


# See http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.wellknownfoldername(v=exchg.80).aspx
WELLKNOWN_FOLDERS = dict([
    ('Calendar', Calendar),
    ('Contacts', Contacts),
    ('DeletedItems', WellknownFolder),
    ('Drafts', WellknownFolder),
    ('Inbox', Inbox),
    ('Journal', WellknownFolder),
    ('Notes', WellknownFolder),
    ('Outbox', WellknownFolder),
    ('SentItems', WellknownFolder),
    ('Tasks', Tasks),
    ('MsgFolderRoot', WellknownFolder),
    ('PublicFoldersRoot', WellknownFolder),
    ('Root', Root),
    ('JunkEmail', WellknownFolder),
    ('Search', WellknownFolder),
    ('VoiceMail', WellknownFolder),
    ('RecoverableItemsRoot', WellknownFolder),
    ('RecoverableItemsDeletions', WellknownFolder),
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

FOLDER_TYPE_MAP = dict()
for folder_name, folder_model in WELLKNOWN_FOLDERS.items():
    folder_type = '%sFolder' % folder_name
    FOLDER_TYPE_MAP[folder_type] = folder_model
