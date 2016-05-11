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


class XMLElement:
    root = None

    def to_xml(self):
        raise NotImplementedError()

    @classmethod
    def from_xml(cls, elem):
        raise NotImplementedError()


class ItemId(XMLElement):
    root = Element('t:ItemId')

    def __init__(self, id, changekey):
        self.id = id
        self.changekey = changekey

    def to_xml(self):
        # copy.deepcopy() is an order of magnitude faster than creating a new Element()
        item_id = deepcopy(self.root)
        item_id.set('Id', self.id)
        item_id.set('ChangeKey', self.changekey)
        return item_id

    @classmethod
    def from_xml(cls, elem):
        if not elem:
            return None
        return cls(id=elem.get('Id'), changekey=elem.get('ChangeKey'))

    def __repr__(self):
        return self.__class__.__name__ + repr((self.id, self.changekey))


class MailBox(XMLElement):
    root = Element('t:Mailbox')
    MAILBOX_TYPES = {'Mailbox', 'PublicDL', 'PrivateDL', 'Contact', 'PublicFolder', 'Unknown', 'OneOff'}

    def __init__(self, name=None, email_address=None, mailbox_type=None, item_id=None):
        # There's also the 'RoutingType' element, but it's optional and must have value "SMTP"
        if mailbox_type:
            assert mailbox_type in self.MAILBOX_TYPES
        self.name = name
        self.email_address = email_address
        self.mailbox_type = mailbox_type
        self.item_id = item_id

    def to_xml(self):
        if not self.email_address and not self.item_id:
            # See "Remarks" section of https://msdn.microsoft.com/en-us/library/office/aa565036(v=exchg.150).aspx
            raise AttributeError('Mailbox must have either email_address or item_id')
        mailbox = deepcopy(self.root)
        if self.name:
            set_xml_attr(mailbox, 't:Name', self.name)
        if self.email_address:
            set_xml_attr(mailbox, 't:EmailAddress', self.email_address)
        if self.mailbox_type:
            set_xml_attr(mailbox, 't:MailboxType', self.mailbox_type)
        if self.item_id:
            mailbox.append(self.item_id.to_xml())
        return mailbox

    @classmethod
    def from_xml(cls, elem):
        if not elem:
            return None
        mailbox = elem.find('{%s}Mailbox' % TNS)
        if not mailbox:
            return None
        return cls(
            name=get_xml_attr(mailbox, '{%s}Name' % TNS),
            email_address=get_xml_attr(mailbox, '{%s}EmailAddress' % TNS),
            mailbox_type=get_xml_attr(mailbox, '{%s}MailboxType' % TNS),
            item_id=ItemId.from_xml(mailbox.find('{%s}ItemId' % TNS)),
        )

    def __repr__(self):
        return self.__class__.__name__ + repr((self.name, self.email_address, self.mailbox_type, self.item_id))


class ExtendedProperty(XMLElement):
    root = Element('t:ExtendedProperty')
    property_id = None
    property_name = None
    property_type = None

    def __init__(self, value):
        self.value = value

    @classmethod
    def field_uri_xml(cls):
        return Element(
            't:ExtendedFieldURI',
            PropertySetId=cls.property_id,
            PropertyName=cls.property_name,
            PropertyType=cls.property_type,
        )

    def to_xml(self):
        extended_property = deepcopy(self.root)
        extended_field_uri = self.field_uri_xml()
        extended_property.append(extended_field_uri)
        set_xml_attr(extended_property, 't:Value', self.value)
        return extended_property

    @classmethod
    def get_value(cls, elem):
        # Gets value of this specific ExtendedProperty from a list of 'ExtendedProperty' XML elements
        extended_field_value = None
        for e in elem:
            extended_field_uri = e.find('{%s}ExtendedFieldURI' % TNS)
            match = True

            for k, v in (
                    ('PropertySetId', cls.property_id),
                    ('PropertyName', cls.property_name),
                    ('PropertyType', cls.property_type),
            ):
                if extended_field_uri.get(k) != v:
                    match = False
                    break
            if match:
                extended_field_value = get_xml_attr(e, '{%s}Value' % TNS) or ''
                break
        return extended_field_value

    def __str__(self):
        return self.__class__.__name__ + repr((self.value,))


class ExternId(ExtendedProperty):
    # 'c11ff724-aa03-4555-9952-8fa248a11c3e' is arbirtary. We just want a unique UUID.
    property_id = 'c11ff724-aa03-4555-9952-8fa248a11c3e'
    property_name = 'External ID'
    property_type = 'String'

    def __init__(self, extern_id):
        super().__init__(value=extern_id)


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
        'extern_id': ExternId,
    }
    FIELD_TYPE_MAP = {
        'item_id': str,
        'changekey': str,
        'subject': str,
        'body': str,
        'reminder_is_set': bool,
        'categories': list,
        'datetime_created': EWSDateTime,
        'datetime_sent': EWSDateTime,
        'datetime_recieved': EWSDateTime,
        'last_modified_name': str,
        'last_modified_time': EWSDateTime,
        'extern_id': ExternId,
    }

    def __init__(self, **kwargs):
        for k in Item.ITEM_FIELDS + Item.EXTRA_ITEM_FIELDS:
            default = False if k == 'reminder_is_set' else None
            v = kwargs.pop(k, default)
            if v is not None:
                field_type = self.type_for_field(k)
                if k != 'extern_id' and not isinstance(v, field_type):
                    raise TypeError('Field %s value "%s" must be of type %s' % (k, v, field_type))
            setattr(self, k, v)
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
        assert isinstance(fieldname, str)
        try:
            elem = Element('t:%s' % cls.ATTR_FIELDURI_MAP[fieldname])
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
        for f in cls.fieldnames(with_extra=with_extra):
            field_uri = cls.fielduri_for_field(f)
            if isinstance(field_uri, str):
                f = deepcopy(fielduri)
                f.set('FieldURI', field_uri)
                fields.append(f)
            else:
                # ExtendedProperty
                fields.append(field_uri.field_uri_xml())
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
            elif issubclass(t, ExtendedProperty):
                kwargs[fieldname] = t.get_value(extended_properties)
            elif issubclass(t, XMLElement):
                kwargs[fieldname] = t.from_xml(elem.find(cls.response_xml_elem_for_field(fieldname)))
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
                mailbox = MailBox(email_address=self.account.primary_smtp_address)
                distinguishedfolderid.append(mailbox.to_xml())
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
        item_ids = Element('m:ItemIds')
        for item in ids:
            item_id = ItemId(*item) if isinstance(item, tuple) else ItemId(item.item_id, item.changekey)
            item_ids.append(item_id.to_xml())
        getitem.append(item_ids)
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
        deleteitem = Element('m:%s' % DeleteItem.SERVICE_NAME, DeleteType='HardDelete')
        if isinstance(self, Calendar):
            deleteitem.set('SendMeetingCancellations', 'SendToNone')
        elif isinstance(self, Tasks):
            deleteitem.set('AffectedTaskOccurrences', 'SpecifiedOccurrenceOnly')
        if self.account.version.major_version >= 15:
            deleteitem.set('SuppressReadReceipts', 'true')

        item_ids = Element('m:ItemIds')
        for item in ids:
            item_id = ItemId(*item) if isinstance(item, tuple) else ItemId(item.item_id, item.changekey)
            item_ids.append(item_id.to_xml())
        deleteitem.append(item_ids)
        return deleteitem

    def update_xml(self, items):
        assert len(items)
        assert isinstance(items[0], tuple)
        assert isinstance(items[0][0], (tuple, Item))
        assert isinstance(items[0][1], dict)
        # Prepare reuseable Element objects
        updateitem = Element('m:%s' % UpdateItem.SERVICE_NAME, ConflictResolution='AutoResolve')
        if isinstance(self, Calendar):
            updateitem.set('SendMeetingInvitationsOrCancellations', 'SendToNone')
        elif isinstance(self, Inbox):
            updateitem.set('MessageDisposition', 'SaveOnly')
        if self.account.version.major_version >= 15:
            updateitem.set('SuppressReadReceipts', 'true')

        itemchanges = Element('m:ItemChanges')
        itemchange = Element('t:ItemChange')
        updates = Element('t:Updates')
        setitemfield = Element('t:SetItemField')
        deleteitemfield = Element('t:DeleteItemField')
        fielduri = Element('t:FieldURI')
        folderitem = Element('t:%s' % self.item_model.ITEM_TYPE)

        # copy.deepcopy() is an order of magnitude faster than having
        # Element() inside the loop. It matters because 'ids' may
        # be a very large list.
        for item, update_dict in items:
            assert len(update_dict)
            i_itemchange = deepcopy(itemchange)
            item_id = ItemId(*item) if isinstance(item, tuple) else ItemId(item.item_id, item.changekey)
            i_itemchange.append(item_id.to_xml())
            i_updates = deepcopy(updates)
            meeting_timezone_added = False
            for fieldname, val in update_dict.items():
                if val is not None and fieldname == 'extern_id':
                    val = ExternId(val)
                field_uri = self.attr_to_fielduri(fieldname)
                if isinstance(field_uri, str):
                    i_fielduri = deepcopy(fielduri)
                    i_fielduri.set('FieldURI', field_uri)
                else:
                    # ExtendedProperty
                    i_fielduri = field_uri.field_uri_xml()
                if val is None:
                    # A value of None means we want to remove this field from the item
                    i_deleteitemfield = deepcopy(deleteitemfield)
                    i_deleteitemfield.append(i_fielduri)
                    i_updates.append(i_deleteitemfield)
                    continue
                i_setitemfield = deepcopy(setitemfield)
                i_setitemfield.append(i_fielduri)
                i_folderitem = deepcopy(folderitem)
                if isinstance(val, EWSDateTime):
                    i_value = self.attr_to_request_xml_elem(fieldname)
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
                    if isinstance(val, XMLElement):
                        i_folderitem.append(val.to_xml())
                    else:
                        i_value = self.attr_to_request_xml_elem(fieldname)
                        i_folderitem.append(set_xml_value(i_value, val))
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
        'organizer': MailBox,
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
        assert isinstance(fieldname, str)
        try:
            return Element('t:%s' % cls.ATTR_FIELDURI_MAP[fieldname])
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
                i.append(ExternId(item.extern_id).to_xml())
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
