"""
Implements a selection of the folders and folder items found in an Exchange account.

Exchange is very picky about things like the order of XML elements in SOAP requests, so we need to generate XML
automatically instead of taking advantage of Python SOAP libraries and the WSDL file.
"""

from logging import getLogger
import re

from .credentials import DELEGATE
from .ewsdatetime import EWSDateTime, EWSTimeZone
from .restriction import Restriction
from .services import TNS, FindItem, IdOnly, SHALLOW, DEEP, DeleteItem, CreateItem, UpdateItem, FindFolder, GetFolder, \
    GetItem
from .util import create_element, add_xml_child, get_xml_attrs, get_xml_attr, set_xml_value, ElementType, peek

log = getLogger(__name__)


class EWSElement:
    ELEMENT_NAME = None

    __slots__ = tuple()

    def to_xml(self, version):
        raise NotImplementedError()

    @classmethod
    def from_xml(cls, elem):
        raise NotImplementedError()

    @classmethod
    def request_tag(cls):
        return 't:%s' % cls.ELEMENT_NAME

    @classmethod
    def response_tag(cls):
        return '{%s}%s' % (TNS, cls.ELEMENT_NAME)


class ItemId(EWSElement):
    ELEMENT_NAME = 'ItemId'

    __slots__ = ('id', 'changekey')

    def __init__(self, id, changekey):
        assert isinstance(id, str)
        assert isinstance(changekey, str)
        self.id = id
        self.changekey = changekey

    def to_xml(self, version):
        # Don't use create_element with extra args. It caches results and Id is always unique.
        elem = create_element(self.request_tag())
        elem.set('Id', self.id)
        elem.set('ChangeKey', self.changekey)
        return elem

    @classmethod
    def from_xml(cls, elem):
        if not elem:
            return None
        assert elem.tag == cls.response_tag()
        return cls(id=elem.get('Id'), changekey=elem.get('ChangeKey'))


class Mailbox(EWSElement):
    ELEMENT_NAME = 'Mailbox'
    MAILBOX_TYPES = {'Mailbox', 'PublicDL', 'PrivateDL', 'Contact', 'PublicFolder', 'Unknown', 'OneOff'}

    __slots__ = ('name', 'email_address', 'mailbox_type', 'item_id')

    def __init__(self, name=None, email_address=None, mailbox_type=None, item_id=None):
        # There's also the 'RoutingType' element, but it's optional and must have value "SMTP"
        if name is not None:
            assert isinstance(name, str)
        if email_address is not None:
            assert isinstance(email_address, str)
        if mailbox_type is not None:
            assert mailbox_type in self.MAILBOX_TYPES
        if item_id is not None:
            assert isinstance(item_id, ItemId)
        self.name = name
        self.email_address = email_address
        self.mailbox_type = mailbox_type
        self.item_id = item_id

    def to_xml(self, version):
        if not self.email_address and not self.item_id:
            # See "Remarks" section of https://msdn.microsoft.com/en-us/library/office/aa565036(v=exchg.150).aspx
            raise AttributeError('Mailbox must have either email_address or item_id')
        mailbox = create_element(self.request_tag())
        if self.name:
            add_xml_child(mailbox, 't:Name', self.name)
        if self.email_address:
            add_xml_child(mailbox, 't:EmailAddress', self.email_address)
        if self.mailbox_type:
            add_xml_child(mailbox, 't:MailboxType', self.mailbox_type)
        if self.item_id:
            set_xml_value(mailbox, self.item_id, version)
        return mailbox

    @classmethod
    def from_xml(cls, elem):
        if not elem:
            return None
        assert elem.tag == cls.response_tag()
        return cls(
            name=get_xml_attr(elem, '{%s}Name' % TNS),
            email_address=get_xml_attr(elem, '{%s}EmailAddress' % TNS),
            mailbox_type=get_xml_attr(elem, '{%s}MailboxType' % TNS),
            item_id=ItemId.from_xml(elem.find(ItemId.response_tag())),
        )

    def __repr__(self):
        return self.__class__.__name__ + repr((self.name, self.email_address, self.mailbox_type, self.item_id))


class ExtendedProperty(EWSElement):
    ELEMENT_NAME = 'ExtendedProperty'

    property_id = None
    property_name = None
    property_type = None

    __slots__ = ('value',)

    def __init__(self, value):
        assert isinstance(value, str)
        self.value = value

    @classmethod
    def field_uri_xml(cls):
        return create_element(
            't:ExtendedFieldURI',
            PropertySetId=cls.property_id,
            PropertyName=cls.property_name,
            PropertyType=cls.property_type
        )

    def to_xml(self, version):
        extended_property = create_element(self.request_tag())
        set_xml_value(extended_property, self.field_uri_xml(), version)
        add_xml_child(extended_property, 't:Value', self.value)
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

    def __repr__(self):
        return self.__class__.__name__ + repr((self.value,))


class ExternId(ExtendedProperty):
    # 'c11ff724-aa03-4555-9952-8fa248a11c3e' is arbirtary. We just want a unique UUID.
    property_id = 'c11ff724-aa03-4555-9952-8fa248a11c3e'
    property_name = 'External ID'
    property_type = 'String'

    __slots__ = ('value',)

    def __init__(self, extern_id):
        super().__init__(value=extern_id)


class Attendee(EWSElement):
    ELEMENT_NAME = 'Attendee'
    RESPONSE_TYPES = {'Unknown', 'Organizer', 'Tentative', 'Accept', 'Decline', 'NoResponseReceived'}

    __slots__ = ('mailbox', 'response_type', 'last_response_time')

    def __init__(self, mailbox, response_type, last_response_time=None):
        assert isinstance(mailbox, Mailbox)
        assert response_type in self.RESPONSE_TYPES
        if last_response_time is not None:
            assert isinstance(last_response_time, EWSDateTime)
        self.mailbox = mailbox
        self.response_type = response_type
        self.last_response_time = last_response_time

    def to_xml(self, version):
        attendee = create_element(self.request_tag())
        set_xml_value(attendee, self.mailbox, version)
        add_xml_child(attendee, 't:ResponseType', self.response_type)
        add_xml_child(attendee, 't:LastResponseTime', self.last_response_time)
        return attendee

    @classmethod
    def from_xml(cls, elem):
        if not elem:
            return None
        assert elem.tag == cls.response_tag()
        last_response_time = get_xml_attr(elem, '{%s}LastResponseTime' % TNS)
        return cls(
            mailbox=Mailbox.from_xml(elem.find(Mailbox.response_tag())),
            response_type=get_xml_attr(elem, '{%s}ResponseType' % TNS),
            last_response_time=EWSDateTime.from_string(last_response_time) if last_response_time else None,
        )

    def __repr__(self):
        return self.__class__.__name__ + repr((self.mailbox, self.response_type, self.last_response_time))


class Item(EWSElement):
    ELEMENT_NAME = 'Item'
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
        'categories': [str],
        'datetime_created': EWSDateTime,
        'datetime_sent': EWSDateTime,
        'datetime_recieved': EWSDateTime,
        'last_modified_name': str,
        'last_modified_time': EWSDateTime,
        'extern_id': ExternId,
    }

    __slots__ = ITEM_FIELDS + EXTRA_ITEM_FIELDS

    def __init__(self, **kwargs):
        for k in Item.ITEM_FIELDS + Item.EXTRA_ITEM_FIELDS:
            default = False if k == 'reminder_is_set' else None
            v = kwargs.pop(k, default)
            if v is not None:
                # Test if arguments have same type as specified in FIELD_TYPE_MAP. 'extern_id' is psecial because we
                # implement it internally as the ExternId class but want to keep the attribute as a simple str.
                # 'field_type' may be a list with a single type. In that case we want to check all list members
                field_type = self.type_for_field(k)
                if isinstance(field_type, list):
                    elem_type = field_type[0]
                    assert isinstance(v, list)
                    for item in v:
                        if not isinstance(item, elem_type):
                            raise TypeError('Field %s value "%s" must be of type %s' % (k, v, field_type))
                elif k != 'extern_id' and not isinstance(v, field_type):
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
    def elem_for_field(cls, fieldname):
        assert isinstance(fieldname, str)
        try:
            if fieldname == 'body':
                return create_element('t:%s' % cls.ATTR_FIELDURI_MAP[fieldname], BodyType='Text')
            return create_element('t:%s' % cls.ATTR_FIELDURI_MAP[fieldname])
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
    def additional_property_elems(cls, with_extra=False):
        fields = []
        for f in cls.fieldnames(with_extra=with_extra):
            field_uri = cls.fielduri_for_field(f)
            if isinstance(field_uri, str):
                fields.append(create_element('t:FieldURI', FieldURI=field_uri))
            else:
                # ExtendedProperty
                fields.append(field_uri.field_uri_xml())
        return fields

    @classmethod
    def id_from_xml(cls, elem):
        id_elem = elem.find(ItemId.response_tag())
        return id_elem.get('Id'), id_elem.get('ChangeKey')

    @classmethod
    def from_xml(cls, elem, with_extra=False):
        assert elem.tag == cls.response_tag()
        item_id, changekey = cls.id_from_xml(elem)
        kwargs = {}
        extended_properties = elem.findall(ExtendedProperty.response_tag())
        for fieldname in cls.fieldnames(with_extra=with_extra):
            field_type = cls.type_for_field(fieldname)
            if field_type == EWSDateTime:
                str_val = get_xml_attr(elem, cls.response_xml_elem_for_field(fieldname))
                kwargs[fieldname] = EWSDateTime.from_string(str_val)
            elif field_type == bool:
                val = get_xml_attr(elem, cls.response_xml_elem_for_field(fieldname))
                kwargs[fieldname] = True if val == 'true' else False
            elif field_type == str:
                kwargs[fieldname] = get_xml_attr(elem, cls.response_xml_elem_for_field(fieldname)) or ''
            elif isinstance(field_type, list):
                list_type = field_type[0]
                iter_elem = elem.find(cls.response_xml_elem_for_field(fieldname))
                if iter_elem is None:
                    kwargs[fieldname] = None
                else:
                    if list_type == str:
                        kwargs[fieldname] = get_xml_attrs(iter_elem, '{%s}String' % TNS)
                    else:
                        kwargs[fieldname] = [list_type.from_xml(elem)
                                             for elem in iter_elem.findall(list_type.response_tag())]
            elif issubclass(field_type, ExtendedProperty):
                kwargs[fieldname] = field_type.get_value(extended_properties)
            elif issubclass(field_type, EWSElement):
                if fieldname == 'organizer':
                    # We want the nested Mailbox, not the Organizer element itself
                    organizer = elem.find(cls.response_xml_elem_for_field(fieldname))
                    if organizer is None:
                        kwargs[fieldname] = None
                    else:
                        kwargs[fieldname] = field_type.from_xml(organizer.find(Mailbox.response_tag()))
                else:
                    kwargs[fieldname] = field_type.from_xml(elem.find(cls.response_xml_elem_for_field(fieldname)))
            else:
                assert False, 'Field %s type %s not supported' % (fieldname, field_type)
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
        is_empty, items = peek(items)
        if is_empty:
            # We accept generators, so it's not always convenient for caller to know up-front if 'items' is empty. Allow
            # empty 'items' and return early.
            return []
        return list(map(self.item_model.id_from_xml, CreateItem(self.account.protocol).call(folder=self, items=items)))

    def delete_items(self, ids):
        """
        Deletes items in the folder. 'ids' is an iterable of either (item_id, changekey) tuples or Item objects.
        """
        is_empty, ids = peek(ids)
        if is_empty:
            # We accept generators, so it's not always convenient for caller to know up-front if 'items' is empty. Allow
            # empty 'items' and return early.
            return []
        return DeleteItem(self.account.protocol).call(folder=self, ids=ids)

    def update_items(self, items):
        """
        Updates items in the folder. 'items' is an iterable of tuples containing two elements:

            1. either an (item_id, changekey) tuple or an Item object
            2. a dict containing the Item attributes to change

        """
        is_empty, items = peek(items)
        if is_empty:
            # We accept generators, so it's not always convenient for caller to know up-front if 'items' is empty. Allow
            # empty 'items' and return early.
            return []
        return list(map(self.item_model.id_from_xml, UpdateItem(self.account.protocol).call(folder=self, items=items)))

    def get_items(self, ids):
        # get_xml() uses self.with_extra_fields. Pass this to from_xml()
        is_empty, ids = peek(ids)
        if is_empty:
            # We accept generators, so it's not always convenient for caller to know up-front if 'items' is empty. Allow
            # empty 'items' and return early.
            return []
        return list(map(
            lambda i: self.item_model.from_xml(i, self.with_extra_fields),
            GetItem(self.account.protocol).call(folder=self, ids=ids)
        ))

    def test_access(self):
        """
        Does a simple FindItem to test (read) access to the mailbox. Maybe the account doesn't exist, maybe the
        service user doesn't have access to the calendar. This will throw the most common errors.
        """
        now = EWSDateTime.now(tz=EWSTimeZone.timezone('UTC'))
        restriction = Restriction.from_params(self.DISTINGUISHED_FOLDER_ID, start=now, end=now)
        FindItem(self.account.protocol).call(folder=self, restriction=restriction, shape=IdOnly)
        return True

    def folderid_xml(self):
        if self.folder_id:
            assert self.changekey
            return create_element('t:FolderId', Id=self.folder_id, ChangeKey=self.changekey)
        else:
            # Only use distinguished ID if we don't have the folder ID
            distinguishedfolderid = create_element('t:DistinguishedFolderId', Id=self.DISTINGUISHED_FOLDER_ID)
            if self.account.access_type == DELEGATE:
                mailbox = Mailbox(email_address=self.account.primary_smtp_address)
                set_xml_value(distinguishedfolderid, mailbox, self.account.version)
            return distinguishedfolderid

    def get_xml(self, ids):
        # This list should be configurable. 'body' element can only be fetched with GetItem.
        # CalendarItem.from_xml() specifies the items we currently expect. For full list, see
        # https://msdn.microsoft.com/en-us/library/office/aa494315(v=exchg.150).aspx
        log.debug(
            'Getting %s items for %s',
            self.DISTINGUISHED_FOLDER_ID,
            self.account
        )
        getitem = create_element('m:%s' % GetItem.SERVICE_NAME)
        itemshape = create_element('m:ItemShape')
        add_xml_child(itemshape, 't:BaseShape', IdOnly)
        additional_properties = self.item_model.additional_property_elems(with_extra=self.with_extra_fields)
        if additional_properties:
            add_xml_child(itemshape, 't:AdditionalProperties', additional_properties)
        getitem.append(itemshape)
        item_ids = create_element('m:ItemIds')
        n = 0
        for item in ids:
            n += 1
            item_id = ItemId(*item) if isinstance(item, tuple) else ItemId(item.item_id, item.changekey)
            set_xml_value(item_ids, item_id, self.account.version)
        if not n:
            raise AttributeError('"ids" must not be empty')
        getitem.append(item_ids)
        return getitem

    def create_xml(self, items):
        # Takes an account name, a folder name, a list of Calendar.Item obejcts and a function to convert items to XML
        # Elements
        if isinstance(self, Calendar):
            createitem = create_element('m:%s' % CreateItem.SERVICE_NAME, SendMeetingInvitations='SendToNone')
        elif isinstance(self, Inbox):
            createitem = create_element('m:%s' % CreateItem.SERVICE_NAME, MessageDisposition='SaveOnly')
        else:
            createitem = create_element('m:%s' % CreateItem.SERVICE_NAME)
        add_xml_child(createitem, 'm:SavedItemFolderId', self.folderid_xml())
        item_elems = [i.to_xml(self.account.version) for i in items]
        if not item_elems:
            raise AttributeError('"items" must not be empty')
        add_xml_child(createitem, 'm:Items', item_elems)
        return createitem

    def delete_xml(self, ids):
        # Prepare reuseable Element objects
        if isinstance(self, Calendar):
            deleteitem = create_element('m:%s' % DeleteItem.SERVICE_NAME, DeleteType='HardDelete',
                                        SendMeetingCancellations='SendToNone')
        elif isinstance(self, Tasks):
            deleteitem = create_element('m:%s' % DeleteItem.SERVICE_NAME, DeleteType='HardDelete',
                                        AffectedTaskOccurrences='SpecifiedOccurrenceOnly')
        else:
            deleteitem = create_element('m:%s' % DeleteItem.SERVICE_NAME, DeleteType='HardDelete')
        if self.account.version.major_version >= 15:
            deleteitem.set('SuppressReadReceipts', 'true')

        item_ids = create_element('m:ItemIds')
        n = 0
        for item in ids:
            n += 1
            item_id = ItemId(*item) if isinstance(item, tuple) else ItemId(item.item_id, item.changekey)
            set_xml_value(item_ids, item_id, self.account.version)
        if not n:
            raise AttributeError('"ids" must not be empty')
        deleteitem.append(item_ids)
        return deleteitem

    def update_xml(self, items):
        # Prepare reuseable Element objects
        if isinstance(self, Calendar):
            updateitem = create_element('m:%s' % UpdateItem.SERVICE_NAME, ConflictResolution='AutoResolve',
                                        SendMeetingInvitationsOrCancellations='SendToNone')
        elif isinstance(self, Inbox):
            updateitem = create_element('m:%s' % UpdateItem.SERVICE_NAME, ConflictResolution='AutoResolve',
                                        MessageDisposition='SaveOnly')
        else:
            updateitem = create_element('m:%s' % UpdateItem.SERVICE_NAME, ConflictResolution='AutoResolve')
        if self.account.version.major_version >= 15:
            updateitem.set('SuppressReadReceipts', 'true')

        itemchanges = create_element('m:ItemChanges')
        n = 0
        for item, update_dict in items:
            n += 1
            if not update_dict:
                raise AttributeError('"update_dict" must not be empty')
            itemchange = create_element('t:ItemChange')
            item_id = ItemId(*item) if isinstance(item, tuple) else ItemId(item.item_id, item.changekey)
            set_xml_value(itemchange, item_id, self.account.version)
            updates = create_element('t:Updates')
            meeting_timezone_added = False
            for fieldname, val in update_dict.items():
                # Skip fields that are read-only in Exchange
                if fieldname in ('organizer',):
                    log.warning('%s is a read-only field. Skipping', fieldname)
                    continue
                if fieldname == 'extern_id' and val is not None:
                    val = ExternId(val)
                field_uri = self.attr_to_fielduri(fieldname)
                if isinstance(field_uri, str):
                    fielduri = create_element('t:FieldURI', FieldURI=field_uri)
                else:
                    # ExtendedProperty
                    fielduri = field_uri.field_uri_xml()
                if val is None:
                    # A value of None means we want to remove this field from the item
                    add_xml_child(updates, 't:DeleteItemField', fielduri)
                    continue
                setitemfield = create_element('t:SetItemField')
                setitemfield.append(fielduri)
                folderitem = create_element(self.item_model.request_tag())

                if isinstance(val, EWSElement):
                    set_xml_value(folderitem, val, self.account.version)
                else:
                    folderitem.append(
                        set_xml_value(self.item_model.elem_for_field(fieldname), val, self.account.version)
                    )
                setitemfield.append(folderitem)
                updates.append(setitemfield)

                if isinstance(val, EWSDateTime):
                    # Always set timezone explicitly when updating date fields. Exchange 2007 wants "MeetingTimeZone"
                    # instead of explicit timezone on each datetime field.
                    setitemfield_tz = create_element('t:SetItemField')
                    folderitem_tz = create_element(self.item_model.request_tag())
                    if self.account.version.major_version < 14:
                        if meeting_timezone_added:
                            # Let's hope that we're not changing timezone, or that both 'start' and 'end' are supplied.
                            # Exchange 2007 doesn't support different timezone on start and end.
                            continue
                        fielduri_tz = create_element('t:FieldURI', FieldURI='calendar:MeetingTimeZone')
                        timezone = create_element('t:MeetingTimeZone', TimeZoneName=val.tzinfo.ms_id)
                        meeting_timezone_added = True
                    else:
                        if fieldname == 'start':
                            fielduri_tz = create_element('t:FieldURI', FieldURI='calendar:StartTimeZone')
                            timezone = create_element('t:StartTimeZone', Id=val.tzinfo.ms_id, Name=val.tzinfo.ms_name)
                        elif fieldname == 'end':
                            fielduri_tz = create_element('t:FieldURI', FieldURI='calendar:EndTimeZone')
                            timezone = create_element('t:EndTimeZone', Id=val.tzinfo.ms_id, Name=val.tzinfo.ms_name)
                        else:
                            assert False, 'Cannot set timezone for field %s' % fieldname
                    setitemfield_tz.append(fielduri_tz)
                    folderitem_tz.append(timezone)
                    setitemfield_tz.append(folderitem_tz)
                    updates.append(setitemfield_tz)
            itemchange.append(updates)
            itemchanges.append(itemchange)
        if not n:
            raise AttributeError('"items" must not be empty')
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
    ELEMENT_NAME = 'CalendarItem'
    SUBJECT_MAXLENGTH = 255
    LOCATION_MAXLENGTH = 255
    FIELDURI_PREFIX = 'calendar'
    ITEM_FIELDS = (
        'start',
        'end',
        'location',
        'organizer',  # Read-only in Exchange
        'legacy_free_busy_status',
        'required_attendees',
        'optional_attendees',
        'resources',
    )
    ATTR_FIELDURI_MAP = {
        'start': 'Start',
        'end': 'End',
        'location': 'Location',
        'organizer': 'Organizer',
        'legacy_free_busy_status': 'LegacyFreeBusyStatus',
        'required_attendees': 'RequiredAttendees',
        'optional_attendees': 'OptionalAttendees',
        'resources': 'Resources',
    }
    FIELD_TYPE_MAP = {
        'start': EWSDateTime,
        'end': EWSDateTime,
        'location': str,
        'organizer': Mailbox,
        'legacy_free_busy_status': str,
        'required_attendees': [Attendee],
        'optional_attendees': [Attendee],
        'resources': [Attendee],
    }

    __slots__ = ITEM_FIELDS + Item.ITEM_FIELDS + Item.EXTRA_ITEM_FIELDS

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
    def elem_for_field(cls, fieldname):
        assert isinstance(fieldname, str)
        try:
            return create_element('t:%s' % cls.ATTR_FIELDURI_MAP[fieldname])
        except KeyError:
            return Item.elem_for_field(fieldname)

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

    def to_xml(self, version):
        # WARNING: The order of addition of XML elements is VERY important. Exchange expects XML elements in a
        # specific, non-documented order and will fail with meaningless errors if the order is wrong.
        i = create_element(self.request_tag())
        i.append(set_xml_value(self.elem_for_field('subject'), self.subject, version))
        if self.body:
            i.append(set_xml_value(self.elem_for_field('body'), self.body, version))
        if self.categories:
            i.append(set_xml_value(self.elem_for_field('categories'), self.categories, version))
        i.append(set_xml_value(self.elem_for_field('reminder_is_set'), self.reminder_is_set, version))
        if self.extern_id is not None:
            set_xml_value(i, ExternId(self.extern_id), version)
        i.append(set_xml_value(self.elem_for_field('start'), self.start, version))
        i.append(set_xml_value(self.elem_for_field('end'), self.end, version))
        i.append(set_xml_value(self.elem_for_field('legacy_free_busy_status'), self.legacy_free_busy_status, version))
        if self.location:
            i.append(set_xml_value(self.elem_for_field('location'), self.location, version))
        if self.organizer:
            log.warning('organizer is a read-only field. Skipping')
        if self.required_attendees:
            i.append(set_xml_value(self.elem_for_field('required_attendees'), self.required_attendees, version))
        if self.optional_attendees:
            i.append(set_xml_value(self.elem_for_field('optional_attendees'), self.optional_attendees, version))
        if self.resources:
            i.append(set_xml_value(self.elem_for_field('resources'), self.resources, version))
        if version.major_version < 14:
            i.append(create_element('t:MeetingTimeZone', TimeZoneName=self.start.tzinfo.ms_id))
        else:
            i.append(create_element('t:StartTimeZone', Id=self.start.tzinfo.ms_id, Name=self.start.tzinfo.ms_name))
            i.append(create_element('t:EndTimeZone', Id=self.end.tzinfo.ms_id, Name=self.end.tzinfo.ms_name))
        return i

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


class Message(Item):
    ELEMENT_NAME = 'Message'


class Inbox(Folder):
    ELEMENT_NAME = 'Message'
    DISTINGUISHED_FOLDER_ID = 'inbox'
    CONTAINER_CLASS = 'IPF.Note'


class Task(Item):
    ELEMENT_NAME = 'Task'


class Tasks(Folder):
    DISTINGUISHED_FOLDER_ID = 'tasks'
    CONTAINER_CLASS = 'IPF.Task'


class Contact(Item):
    ELEMENT_NAME = 'Contact'


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
