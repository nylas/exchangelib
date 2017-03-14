# coding=utf-8
"""
Implements a selection of the folders and folder items found in an Exchange account.

Exchange is very picky about things like the order of XML elements in SOAP requests, so we need to generate XML
automatically instead of taking advantage of Python SOAP libraries and the WSDL file.
"""

from __future__ import unicode_literals

import base64
import mimetypes
import warnings
from decimal import Decimal
from logging import getLogger

from future.utils import python_2_unicode_compatible
from six import text_type, string_types

from .ewsdatetime import EWSDateTime, UTC, UTC_NOW
from .queryset import QuerySet
from .restriction import Restriction, Q
from .services import TNS, IdOnly, SHALLOW, DEEP, FindFolder, GetFolder, FindItem, GetAttachment, CreateAttachment, \
    DeleteAttachment, MNS, ITEM_TRAVERSAL_CHOICES, FOLDER_TRAVERSAL_CHOICES, SHAPE_CHOICES
from .util import create_element, add_xml_child, get_xml_attrs, get_xml_attr, set_xml_value, value_to_xml_text, \
    xml_text_to_value
from .version import EXCHANGE_2010

string_type = string_types[0]
log = getLogger(__name__)

# MessageDisposition values. See https://msdn.microsoft.com/en-us/library/office/aa565209(v=exchg.150).aspx
SAVE_ONLY = 'SaveOnly'
SEND_ONLY = 'SendOnly'
SEND_AND_SAVE_COPY = 'SendAndSaveCopy'
MESSAGE_DISPOSITION_CHOICES = (SAVE_ONLY, SEND_ONLY, SEND_AND_SAVE_COPY)

# SendMeetingInvitations values: see https://msdn.microsoft.com/en-us/library/office/aa565209(v=exchg.150).aspx
# SendMeetingInvitationsOrCancellations: see https://msdn.microsoft.com/en-us/library/office/aa580254(v=exchg.150).aspx
# SendMeetingCancellations values: see https://msdn.microsoft.com/en-us/library/office/aa562961(v=exchg.150).aspx
SEND_TO_NONE = 'SendToNone'
SEND_ONLY_TO_ALL = 'SendOnlyToAll'
SEND_ONLY_TO_CHANGED = 'SendOnlyToChanged'
SEND_TO_ALL_AND_SAVE_COPY = 'SendToAllAndSaveCopy'
SEND_TO_CHANGED_AND_SAVE_COPY = 'SendToChangedAndSaveCopy'
SEND_MEETING_INVITATIONS_CHOICES = (SEND_TO_NONE, SEND_ONLY_TO_ALL, SEND_TO_ALL_AND_SAVE_COPY)
SEND_MEETING_INVITATIONS_AND_CANCELLATIONS_CHOICES = (SEND_TO_NONE, SEND_ONLY_TO_ALL, SEND_ONLY_TO_CHANGED,
                                                      SEND_TO_ALL_AND_SAVE_COPY, SEND_TO_CHANGED_AND_SAVE_COPY)
SEND_MEETING_CANCELLATIONS_CHOICES = (SEND_TO_NONE, SEND_ONLY_TO_ALL, SEND_TO_ALL_AND_SAVE_COPY)

# AffectedTaskOccurrences values. See https://msdn.microsoft.com/en-us/library/office/aa562961(v=exchg.150).aspx
ALL_OCCURRENCIES = 'AllOccurrences'
SPECIFIED_OCCURRENCE_ONLY = 'SpecifiedOccurrenceOnly'
AFFECTED_TASK_OCCURRENCES_CHOICES = (ALL_OCCURRENCIES, SPECIFIED_OCCURRENCE_ONLY)

# ConflictResolution values. See https://msdn.microsoft.com/en-us/library/office/aa580254(v=exchg.150).aspx
NEVER_OVERWRITE = 'NeverOverwrite'
AUTO_RESOLVE = 'AutoResolve'
ALWAYS_OVERWRITE = 'AlwaysOverwrite'
CONFLICT_RESOLUTION_CHOICES = (NEVER_OVERWRITE, AUTO_RESOLVE, ALWAYS_OVERWRITE)

# DeleteType values. See https://msdn.microsoft.com/en-us/library/office/aa562961(v=exchg.150).aspx
HARD_DELETE = 'HardDelete'
SOFT_DELETE = 'SoftDelete'
MOVE_TO_DELETED_ITEMS = 'MoveToDeletedItems'
DELETE_TYPE_CHOICES = (HARD_DELETE, SOFT_DELETE, MOVE_TO_DELETED_ITEMS)


class Choice(text_type):
    # A helper class used for string enums
    pass


class Email(text_type):
    # A helper class used for email address string
    pass


class AnyURI(text_type):
    # Helper to mark strings that must conform to xsd:anyURI
    # If we want an URI validator, see http://stackoverflow.com/questions/14466585/is-this-regex-correct-for-xsdanyuri
    pass


class Body(text_type):
    # Helper to mark the 'body' field as a complex attribute.
    # MSDN: https://msdn.microsoft.com/en-us/library/office/jj219983(v=exchg.150).aspx
    body_type = 'Text'


class HTMLBody(Body):
    # Helper to mark the 'body' field as a complex attribute.
    # MSDN: https://msdn.microsoft.com/en-us/library/office/jj219983(v=exchg.150).aspx
    body_type = 'HTML'


class Subject(text_type):
    # A helper class used for subject string
    MAXLENGTH = 255

    def clean(self):
        if len(self) > self.MAXLENGTH:
            raise ValueError("'%s' value '%s' exceeds length %s" % (self.__class__.__name__, self, self.MAXLENGTH))


class Location(text_type):
    # A helper class used for location string
    MAXLENGTH = 255

    def clean(self):
        if len(self) > self.MAXLENGTH:
            raise ValueError("'%s' value '%s' exceeds length %s" % (self.__class__.__name__, self, self.MAXLENGTH))


class MimeContent(text_type):
    # Helper to work with the base64 encoded MimeContent Message field
    def b64encode(self):
        return base64.b64encode(self).decode('ascii')

    def b64decode(self):
        return base64.b64decode(self)


class EWSElement(object):
    ELEMENT_NAME = None

    __slots__ = tuple()

    def clean(self):
        # Perform any attribute validation here
        pass

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

    def __eq__(self, other):
        return hash(self) == hash(other)

    def __hash__(self):
        return hash(tuple(getattr(self, f) for f in self.__slots__))

    def __repr__(self):
        return self.__class__.__name__ + repr(tuple(getattr(self, f) for f in self.__slots__))


class MessageHeader(EWSElement):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa565307(v=exchg.150).aspx
    ELEMENT_NAME = 'InternetMessageHeader'
    NAME_ATTR = 'HeaderName'

    __slots__ = ('name', 'value')

    def __init__(self, name, value):
        self.name = name
        self.value = value

    def to_xml(self, version):
        self.clean()
        elem = create_element(self.request_tag())
        # Use .set() to not fill up the create_element() cache with unique values
        elem.set(self.NAME_ATTR, self.name)
        set_xml_value(elem, self.value, version)
        return elem

    @classmethod
    def from_xml(cls, elem):
        if elem is None:
            return None
        assert elem.tag == cls.response_tag(), (cls, elem.tag, cls.response_tag())
        res = cls(name=elem.get(cls.NAME_ATTR), value=elem.text)
        elem.clear()
        return res


class ItemId(EWSElement):
    # 'id' and 'changekey' are UUIDs generated by Exchange
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa580234(v=exchg.150).aspx
    ELEMENT_NAME = 'ItemId'

    ID_ATTR = 'Id'
    CHANGEKEY_ATTR = 'ChangeKey'

    __slots__ = ('id', 'changekey')

    def __init__(self, id, changekey):
        self.id = id
        self.changekey = changekey
        self.clean()

    def clean(self):
        if not isinstance(self.id, string_types) or not self.id:
            raise ValueError("id '%s' must be a non-empty string" % id)
        if not isinstance(self.changekey, string_types) or not self.changekey:
            raise ValueError("changekey '%s' must be a non-empty string" % self.changekey)

    def to_xml(self, version):
        self.clean()
        elem = create_element(self.request_tag())
        # Use .set() to not fill up the create_element() cache with unique values
        elem.set(self.ID_ATTR, self.id)
        elem.set(self.CHANGEKEY_ATTR, self.changekey)
        return elem

    @classmethod
    def from_xml(cls, elem):
        if elem is None:
            return None
        assert elem.tag == cls.response_tag(), (cls, elem.tag, cls.response_tag())
        res = cls(id=elem.get(cls.ID_ATTR), changekey=elem.get(cls.CHANGEKEY_ATTR))
        elem.clear()
        return res

    def __eq__(self, other):
        # A more efficient version of super().__eq__
        if other is None:
            return False
        return self.id == other.id and self.changekey == other.changekey


class FolderId(ItemId):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa579461(v=exchg.150).aspx
    ELEMENT_NAME = 'FolderId'

    __slots__ = ('id', 'changekey')


class ParentItemId(ItemId):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa563720(v=exchg.150).aspx
    ELEMENT_NAME = 'ParentItemId'

    __slots__ = ('id', 'changekey')

    @classmethod
    def request_tag(cls):
        return 'm:%s' % cls.ELEMENT_NAME


class RootItemId(ItemId):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/bb204277(v=exchg.150).aspx
    ELEMENT_NAME = 'RootItemId'
    ID_ATTR = 'RootItemId'
    CHANGEKEY_ATTR = 'RootItemChangeKey'

    __slots__ = ('id', 'changekey')

    @classmethod
    def request_tag(cls):
        return 'm:%s' % cls.ELEMENT_NAME

    @classmethod
    def response_tag(cls):
        return '{%s}%s' % (MNS, cls.ELEMENT_NAME)


class AttachmentId(EWSElement):
    # 'id' and 'changekey' are UUIDs generated by Exchange
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa580987(v=exchg.150).aspx
    ELEMENT_NAME = 'AttachmentId'

    ID_ATTR = 'Id'
    ROOT_ID_ATTR = 'RootItemId'
    ROOT_CHANGEKEY_ATTR = 'RootItemChangeKey'

    __slots__ = ('id', 'root_id', 'root_changekey')

    def __init__(self, id, root_id=None, root_changekey=None):
        self.id = id
        self.root_id = root_id
        self.root_changekey = root_changekey
        self.clean()

    def clean(self):
        if not isinstance(self.id, string_types) or not id:
            raise ValueError("id '%s' must be a non-empty string" % id)
        if self.root_id is not None or self.root_changekey is not None:
            if self.root_id is not None and (not isinstance(self.root_id, string_types) or not self.root_id):
                raise ValueError("root_id '%s' must be a non-empty string" % self.root_id)
            if self.root_changekey is not None and \
                    (not isinstance(self.root_changekey, string_types) or not self.root_changekey):
                raise ValueError("root_changekey '%s' must be a non-empty string" % self.root_changekey)

    def to_xml(self, version):
        self.clean()
        elem = create_element(self.request_tag())
        # Use .set() to not fill up the create_element() cache with unique values
        elem.set(self.ID_ATTR, self.id)
        if self.root_id:
            elem.set(self.ROOT_ID_ATTR, self.root_id)
        if self.root_changekey:
            elem.set(self.ROOT_CHANGEKEY_ATTR, self.root_changekey)
        return elem

    @classmethod
    def from_xml(cls, elem):
        if elem is None:
            return None
        assert elem.tag == cls.response_tag(), (cls, elem.tag, cls.response_tag())
        res = cls(
            id=elem.get(cls.ID_ATTR),
            root_id=elem.get(cls.ROOT_ID_ATTR),
            root_changekey=elem.get(cls.ROOT_CHANGEKEY_ATTR)
        )
        elem.clear()
        return res


class Attachment(EWSElement):
    """
    Parent class for FileAttachment and ItemAttachment
    """
    # TODO: Rewrite these as tuple of Field() elements
    ATTACHMENT_FIELDS = {
        'name': ('Name', string_type),
        'content_type': ('ContentType', string_type),
        'attachment_id': (AttachmentId, AttachmentId),
        'content_id': ('ContentId', string_type),
        'content_location': ('ContentLocation', AnyURI),
        'size': ('Size', int),
        'last_modified_time': ('LastModifiedTime', EWSDateTime),
        'is_inline': ('IsInline', bool),
    }
    ORDERED_FIELDS = (
        'attachment_id', 'name', 'content_type', 'content_id', 'content_location', 'size', 'last_modified_time',
        'is_inline',
    )

    __slots__ = ('parent_item',) + ORDERED_FIELDS

    def __init__(self, parent_item=None, attachment_id=None, name=None, content_type=None, content_id=None,
                 content_location=None, size=None, last_modified_time=None, is_inline=None):
        self.parent_item = parent_item
        self.name = name
        self.content_type = content_type
        self.attachment_id = attachment_id
        self.content_id = content_id
        self.content_location = content_location
        self.size = size  # Size is attachment size in bytes
        self.last_modified_time = last_modified_time
        self.is_inline = is_inline
        self.clean()

    def clean(self):
        if self.content_type is None and self.name is not None:
            self.content_type = mimetypes.guess_type(self.name)[0] or 'application/octet-stream'

    def to_xml(self, version):
        self.clean()
        entry = create_element(self.request_tag())
        for field_name in self.ORDERED_FIELDS:
            if field_name == 'size':
                # 'Size' is read-only
                continue
            val = getattr(self, field_name)
            if field_name == 'content':
                # EWS wants file content base64-encoded
                if val is None:
                    raise ValueError("File attachment must contain data")
                val = base64.b64encode(val).decode('ascii')
            if val is None:
                continue
            field_uri = self.ATTACHMENT_FIELDS[field_name][0]
            if isinstance(field_uri, string_types):
                add_xml_child(entry, 't:%s' % field_uri, val)
            elif issubclass(field_uri, EWSElement):
                set_xml_value(entry, val, version)
            else:
                assert False, 'field_uri %s not supported' % field_uri
        return entry

    @classmethod
    def from_xml(cls, elem, parent_item=None):
        if elem is None:
            return None
        assert elem.tag == cls.response_tag(), (cls, elem.tag, cls.response_tag())
        kwargs = {}
        for field_name, (field_uri, field_type) in cls.ATTACHMENT_FIELDS.items():
            if field_name == 'item':
                kwargs[field_name] = None
                for item_cls in ITEM_CLASSES:
                    item_elem = elem.find(item_cls.response_tag())
                    if item_elem is not None:
                        account = parent_item.account if parent_item else None
                        kwargs[field_name] = item_cls.from_xml(elem=item_elem, account=account)
                        break
                continue
            if issubclass(field_type, EWSElement):
                kwargs[field_name] = field_type.from_xml(elem=elem.find(field_type.response_tag()))
                continue
            response_tag = '{%s}%s' % (TNS, field_uri)
            val = get_xml_attr(elem, response_tag)
            if field_name == 'content':
                kwargs[field_name] = None if val is None else base64.b64decode(val)
                continue
            if field_name == 'last_modified_time' and val is not None and not val.endswith('Z'):
                # Sometimes, EWS will send timestamps without the 'Z' for UTC. It seems like the values are still
                # UTC, so mark them as such so EWSDateTime can still interpret the timestamps.
                val += 'Z'
            kwargs[field_name] = xml_text_to_value(value=val, value_type=field_type)
        elem.clear()
        return cls(parent_item=parent_item, **kwargs)

    def attach(self):
        # Adds this attachment to an item and updates the item_id and updated changekey on the parent item
        if self.attachment_id:
            raise ValueError('This attachment has already been created')
        if not self.parent_item.account:
            raise ValueError('Parent item %s must have an account' % self.parent_item)
        items = list(
            i if isinstance(i, Exception) else self.from_xml(elem=i)
            for i in CreateAttachment(account=self.parent_item.account).call(parent_item=self.parent_item, items=[self])
        )
        assert len(items) == 1
        root_item_id = items[0]
        if isinstance(root_item_id, Exception):
            raise root_item_id
        attachment_id = root_item_id.attachment_id
        assert attachment_id.root_id == self.parent_item.item_id
        assert attachment_id.root_changekey != self.parent_item.changekey
        self.parent_item.changekey = attachment_id.root_changekey
        # EWS does not like receiving root_id and root_changekey on subsequent requests
        attachment_id.root_id = None
        attachment_id.root_changekey = None
        self.attachment_id = attachment_id

    def detach(self):
        # Deletes an attachment remotely and returns the item_id and updated changekey of the parent item
        if not self.attachment_id:
            raise ValueError('This attachment has not been created')
        if not self.parent_item:
            raise ValueError('This attachment is not attached to an item')
        items = list(
            i if isinstance(i, Exception) else RootItemId.from_xml(elem=i)
            for i in DeleteAttachment(account=self.parent_item.account).call(items=[self.attachment_id])
        )
        assert len(items) == 1
        root_item_id = items[0]
        if isinstance(root_item_id, Exception):
            raise root_item_id
        assert root_item_id.id == self.parent_item.item_id
        assert root_item_id.changekey != self.parent_item.changekey
        self.parent_item.changekey = root_item_id.changekey
        self.parent_item = None
        self.attachment_id = None

    def __hash__(self):
        if self.attachment_id is None:
            return hash(tuple(getattr(self, f) for f in self.__slots__[1:]))
        return hash(self.attachment_id)

    def __repr__(self):
        return self.__class__.__name__ + '(%s)' % ', '.join(
            '%s=%s' % (k, repr(getattr(self, k))) for k in self.ORDERED_FIELDS if k not in ('item', 'content')
        )


class IndexedElement(EWSElement):
    LABELS = set()
    SUB_FIELD_ELEMENT_NAMES = dict()
    __slots__ = tuple()


class EmailAddress(IndexedElement):
    # MSDN:  https://msdn.microsoft.com/en-us/library/office/aa564757(v=exchg.150).aspx
    ELEMENT_NAME = 'Entry'
    LABELS = {'EmailAddress1', 'EmailAddress2', 'EmailAddress3'}

    __slots__ = ('label', 'email')

    def __init__(self, email, label='EmailAddress1'):
        self.label = label
        self.email = email
        self.clean()

    def clean(self):
        assert self.label in self.LABELS, self.label
        assert isinstance(self.email, string_types), self.email

    def to_xml(self, version):
        self.clean()
        entry = create_element(self.request_tag(), Key=self.label)
        set_xml_value(entry, self.email, version)
        return entry

    @classmethod
    def from_xml(cls, elem):
        if elem is None:
            return None
        assert elem.tag == cls.response_tag(), (cls, elem.tag, cls.response_tag())
        res = cls(
            label=elem.get('Key'),
            email=elem.text or elem.get('Name'),  # Sometimes elem.text is empty. Exchange saves the same in 'Name' attr
        )
        elem.clear()
        return res


class PhoneNumber(IndexedElement):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa565941(v=exchg.150).aspx
    ELEMENT_NAME = 'Entry'
    LABELS = {
        'AssistantPhone', 'BusinessFax', 'BusinessPhone', 'BusinessPhone2', 'Callback', 'CarPhone', 'CompanyMainPhone',
        'HomeFax', 'HomePhone', 'HomePhone2', 'Isdn', 'MobilePhone', 'OtherFax', 'OtherTelephone', 'Pager',
        'PrimaryPhone', 'RadioPhone', 'Telex', 'TtyTddPhone',
    }

    __slots__ = ('label', 'phone_number')

    def __init__(self, phone_number, label='PrimaryPhone'):
        self.label = label
        self.phone_number = phone_number
        self.clean()

    def clean(self):
        assert self.label in self.LABELS, self.label
        assert isinstance(self.phone_number, (int, string_type)), self.phone_number

    def to_xml(self, version):
        self.clean()
        entry = create_element(self.request_tag(), Key=self.label)
        set_xml_value(entry, text_type(self.phone_number), version)
        return entry

    @classmethod
    def from_xml(cls, elem):
        if elem is None:
            return None
        assert elem.tag == cls.response_tag(), (cls, elem.tag, cls.response_tag())
        res = cls(
            label=elem.get('Key'),
            phone_number=elem.text,
        )
        elem.clear()
        return res


class PhysicalAddress(IndexedElement):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa564323(v=exchg.150).aspx
    ELEMENT_NAME = 'Entry'
    LABELS = {'Business', 'Home', 'Other'}
    SUB_FIELD_ELEMENT_NAMES = {
        'street': 'Street',
        'city': 'City',
        'state': 'State',
        'country': 'CountryOrRegion',
        'zipcode': 'PostalCode',
    }

    __slots__ = ('label', 'street', 'city', 'state', 'country', 'zipcode')

    def __init__(self, street=None, city=None, state=None, country=None, zipcode=None, label='Business'):
        self.label = label
        self.street = street  # Street *and* house number (and similar info)
        self.city = city
        self.state = state
        self.country = country
        self.zipcode = zipcode
        self.clean()

    def clean(self):
        assert self.label in self.LABELS, self.label
        if self.street is not None:
            assert isinstance(self.street, string_types), self.street
        if self.city is not None:
            assert isinstance(self.city, string_types), self.city
        if self.state is not None:
            assert isinstance(self.state, string_types), self.state
        if self.country is not None:
            assert isinstance(self.country, string_types), self.country
        if self.zipcode is not None:
            assert isinstance(self.zipcode, (string_type, int)), self.zipcode

    def to_xml(self, version):
        self.clean()
        entry = create_element(self.request_tag(), Key=self.label)
        for attr in self.__slots__:
            if attr == 'label':
                continue
            val = getattr(self, attr)
            if val is not None:
                add_xml_child(entry, 't:%s' % self.SUB_FIELD_ELEMENT_NAMES[attr], val)
        return entry

    @classmethod
    def from_xml(cls, elem):
        if elem is None:
            return None
        assert elem.tag == cls.response_tag(), (cls, elem.tag, cls.response_tag())
        kwargs = dict(label=elem.get('Key'))
        for k, v in cls.SUB_FIELD_ELEMENT_NAMES.items():
            kwargs[k] = get_xml_attr(elem, '{%s}%s' % (TNS, v))
        elem.clear()
        return cls(**kwargs)


class Mailbox(EWSElement):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa565036(v=exchg.150).aspx
    ELEMENT_NAME = 'Mailbox'
    MAILBOX_TYPES = {'Mailbox', 'PublicDL', 'PrivateDL', 'Contact', 'PublicFolder', 'Unknown', 'OneOff'}

    __slots__ = ('name', 'email_address', 'mailbox_type', 'item_id')

    def __init__(self, name=None, email_address=None, mailbox_type=None, item_id=None):
        # There's also the 'RoutingType' element, but it's optional and must have value "SMTP"
        self.name = name
        self.email_address = email_address
        self.mailbox_type = mailbox_type
        self.item_id = item_id
        self.clean()

    def clean(self):
        if self.name is not None:
            assert isinstance(self.name, string_types)
        if self.email_address is not None:
            assert isinstance(self.email_address, string_types)
        if self.mailbox_type is not None:
            assert self.mailbox_type in self.MAILBOX_TYPES
        if self.item_id is not None:
            assert isinstance(self.item_id, ItemId)
        if not self.email_address and not self.item_id:
            # See "Remarks" section of https://msdn.microsoft.com/en-us/library/office/aa565036(v=exchg.150).aspx
            raise AttributeError('Mailbox must have either email_address or item_id')

    def to_xml(self, version):
        self.clean()
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
        if elem is None:
            return None
        assert elem.tag == cls.response_tag(), (elem.tag, cls.response_tag())
        res = cls(
            name=get_xml_attr(elem, '{%s}Name' % TNS),
            email_address=get_xml_attr(elem, '{%s}EmailAddress' % TNS),
            mailbox_type=get_xml_attr(elem, '{%s}MailboxType' % TNS),
            item_id=ItemId.from_xml(elem=elem.find(ItemId.response_tag())),
        )
        elem.clear()
        return res

    def __hash__(self):
        # Exchange may add 'mailbox_type' and 'name' on insert. We're satisfied if the item_id or email address matches.
        if self.item_id:
            return hash(self.item_id)
        return hash(self.email_address.lower())


class RoomList(Mailbox):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/dd899514(v=exchg.150).aspx
    ELEMENT_NAME = 'RoomList'
    # In a GetRoomLists response, room lists are delivered as Address elements
    # MSDN: https://msdn.microsoft.com/en-us/library/office/dd899404(v=exchg.150).aspx
    RESPONSE_ELEMENT_NAME = 'Address'

    @classmethod
    def request_tag(cls):
        return 'm:%s' % cls.ELEMENT_NAME

    @classmethod
    def response_tag(cls):
        return '{%s}%s' % (TNS, cls.RESPONSE_ELEMENT_NAME)


class Room(Mailbox):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/dd899479(v=exchg.150).aspx
    ELEMENT_NAME = 'Room'

    @classmethod
    def from_xml(cls, elem):
        if elem is None:
            return None
        assert elem.tag == cls.response_tag(), (elem.tag, cls.response_tag())
        id_elem = elem.find('{%s}Id' % TNS)
        res = cls(
            name=get_xml_attr(id_elem, '{%s}Name' % TNS),
            email_address=get_xml_attr(id_elem, '{%s}EmailAddress' % TNS),
            mailbox_type=get_xml_attr(id_elem, '{%s}MailboxType' % TNS),
            item_id=ItemId.from_xml(elem=id_elem.find(ItemId.response_tag())),
        )
        elem.clear()
        return res


class ExtendedProperty(EWSElement):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa566405(v=exchg.150).aspx

    Property_* values: https://msdn.microsoft.com/en-us/library/office/aa564843(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'ExtendedProperty'

    DISTINGUISHED_SETS = {
        'Meeting',
        'Appointment',
        'Common',
        'PublicStrings',
        'Address',
        'InternetHeaders',
        'CalendarAssistant',
        'UnifiedMessaging',
    }
    PROPERTY_TYPES = {
        'ApplicationTime',
        'Binary',
        'BinaryArray',
        'Boolean',
        'CLSID',
        'CLSIDArray',
        'Currency',
        'CurrencyArray',
        'Double',
        'DoubleArray',
        # 'Error',
        'Float',
        'FloatArray',
        'Integer',
        'IntegerArray',
        'Long',
        'LongArray',
        # 'Null',
        # 'Object',
        # 'ObjectArray',
        'Short',
        'ShortArray',
        # 'SystemTime',  # Not implemented yet
        # 'SystemTimeArray',  # Not implemented yet
        'String',
        'StringArray',
    }  # The commented-out types cannot be used for setting or getting (see docs) and are thus not very useful here

    distinguished_property_set_id = None
    property_set_id = None
    property_tag = None  # hex integer (e.g. 0x8000) or string ('0x8000')
    property_name = None
    property_id = None  # integer as hex-formatted int (e.g. 0x8000) or normal int (32768)
    property_type = None

    __slots__ = ('value',)

    def __init__(self, value):
        self.value = value

    def clean(self):
        if self.distinguished_property_set_id:
            assert not any([self.property_set_id, self.property_tag])
            assert any([self.property_id, self.property_name])
            assert self.distinguished_property_set_id in self.DISTINGUISHED_SETS
        if self.property_set_id:
            assert not any([self.distinguished_property_set_id, self.property_tag])
            assert any([self.property_id, self.property_name])
        if self.property_tag:
            assert not any([
                self.distinguished_property_set_id, self.property_set_id, self.property_name, self.property_id
            ])
            if 0x8000 <= self.property_tag_as_int() <= 0xFFFE:
                raise ValueError(
                    "'property_tag' value '%s' is reserved for custom properties" % self.property_tag_as_hex()
                )
        if self.property_name:
            assert not any([self.property_id, self.property_tag])
            assert any([self.distinguished_property_set_id, self.property_set_id])
        if self.property_id:
            assert not any([self.property_name, self.property_tag])
            assert any([self.distinguished_property_set_id, self.property_set_id])
        assert self.property_type in self.PROPERTY_TYPES

        python_type = self.python_type()
        if self.is_array_type():
            for v in self.value:
                assert isinstance(v, python_type)
        else:
            assert isinstance(self.value, python_type)

    @classmethod
    def is_array_type(cls):
        return cls.property_type.endswith('Array')

    @classmethod
    def property_tag_as_int(cls):
        if isinstance(cls.property_tag, string_types):
            return int(cls.property_tag, base=16)
        return cls.property_tag

    @classmethod
    def property_tag_as_hex(cls):
        return hex(cls.property_tag) if isinstance(cls.property_tag, int) else cls.property_tag

    @classmethod
    def python_type(cls):
        # Return the best equivalent for a Python type for the property type of this class
        base_type = cls.property_type[:-5] if cls.is_array_type() else cls.property_type
        return {
            'ApplicationTime': Decimal,
            'Binary': bytes,
            'Boolean': bool,
            'CLSID': string_type,
            'Currency': int,
            'Double': Decimal,
            'Float': Decimal,
            'Integer': int,
            'Long': int,
            'Short': int,
            # 'SystemTime': int,
            'String': string_type,
        }[base_type]

    def to_xml(self, version):
        if self.is_array_type():
            values = create_element('t:Values')
            for v in self.value:
                add_xml_child(values, 't:Value', v)
            return values
        else:
            value = create_element('t:Value')
            set_xml_value(value, self.value, version=version)
            return value

    @classmethod
    def from_xml(cls, elems):
        # Gets value of this specific ExtendedProperty from a list of 'ExtendedProperty' XML elements
        python_type = cls.python_type()
        extended_field_value = None
        for e in elems:
            extended_field_uri = e.find('{%s}ExtendedFieldURI' % TNS)
            match = True

            for k, v in (
                    ('DistinguishedPropertySetId', cls.distinguished_property_set_id),
                    ('PropertySetId', cls.property_set_id),
                    ('PropertyTag', cls.property_tag_as_hex()),
                    ('PropertyName', cls.property_name),
                    ('PropertyId', value_to_xml_text(cls.property_id) if cls.property_id else None),
                    ('PropertyType', cls.property_type),
            ):
                if extended_field_uri.get(k) != v:
                    match = False
                    break
            if match:
                if cls.is_array_type():
                    extended_field_value = [
                        xml_text_to_value(value=val, value_type=python_type)
                        for val in get_xml_attrs(e, '{%s}Value' % TNS)
                    ]
                else:
                    extended_field_value = xml_text_to_value(
                        value=get_xml_attr(e, '{%s}Value' % TNS), value_type=python_type)
                    if python_type == string_type and not extended_field_value:
                        # For string types, we want to return the empty string instead of None if the element was
                        # actually found, but there was no XML value. For other types, it would be more problematic
                        # to make that distinction, e.g. return False for bool, 0 for int, etc.
                        extended_field_value = ''
                break
        return extended_field_value


class ExternId(ExtendedProperty):
    # This is a custom extended property defined by us. It's useful for synchronization purposes, to attach a unique ID
    # from an external system. Strictly, this is an field that should probably not be registered by default since it's
    # not part of EWS, but it's been around since the beginning of this library and would be a pain for consumers to
    # register manually.

    property_set_id = 'c11ff724-aa03-4555-9952-8fa248a11c3e'  # This is arbitrary. We just want a unique UUID.
    property_name = 'External ID'
    property_type = 'String'

    __slots__ = ExtendedProperty.__slots__


class Attendee(EWSElement):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa580339(v=exchg.150).aspx
    ELEMENT_NAME = 'Attendee'
    RESPONSE_TYPES = {'Unknown', 'Organizer', 'Tentative', 'Accept', 'Decline', 'NoResponseReceived'}

    __slots__ = ('mailbox', 'response_type', 'last_response_time')

    def __init__(self, mailbox, response_type, last_response_time=None):
        self.mailbox = mailbox
        self.response_type = response_type
        self.last_response_time = last_response_time
        self.clean()

    def clean(self):
        if isinstance(self.mailbox, string_types):
            self.mailbox = Mailbox(email_address=self.mailbox)
        assert isinstance(self.mailbox, Mailbox)
        assert self.response_type in self.RESPONSE_TYPES
        if self.last_response_time is not None:
            assert isinstance(self.last_response_time, EWSDateTime)

    def to_xml(self, version):
        self.clean()
        attendee = create_element(self.request_tag())
        set_xml_value(attendee, self.mailbox, version)
        add_xml_child(attendee, 't:ResponseType', self.response_type)
        if self.last_response_time:
            add_xml_child(attendee, 't:LastResponseTime', self.last_response_time)
        return attendee

    @classmethod
    def from_xml(cls, elem):
        if elem is None:
            return None
        assert elem.tag == cls.response_tag(), (cls, elem.tag, cls.response_tag())
        last_response_time = get_xml_attr(elem, '{%s}LastResponseTime' % TNS)
        res = cls(
            mailbox=Mailbox.from_xml(elem=elem.find(Mailbox.response_tag())),
            response_type=get_xml_attr(elem, '{%s}ResponseType' % TNS) or 'Unknown',
            last_response_time=EWSDateTime.from_string(last_response_time) if last_response_time else None,
        )
        elem.clear()
        return res

    @classmethod
    def set_field_xml(cls, field_elem, items, version):
        # Builds the XML for a SetItemField element
        for item in items:
            attendee = create_element(cls.request_tag())
            set_xml_value(attendee, item.mailbox, version)
            field_elem.append(attendee)
        return field_elem

    def __hash__(self):
        # TODO: maybe take 'response_type' and 'last_response_time' into account?
        return hash(self.mailbox)


class Field(object):
    """
    Holds information related to an item field
    """
    def __init__(self, name, value_cls, from_version=None, choices=None, default=None, is_list=False,
                 is_complex=False, is_required=False, is_read_only=False, is_read_only_after_send=False):
        self.name = name
        self.value_cls = value_cls
        self.from_version = from_version
        self.choices = choices
        self.default = default  # Default value if none is given
        self.is_list = is_list
        # Is the field a complex EWS type? Quoting the EWS FindItem docs:
        #
        #   The FindItem operation returns only the first 512 bytes of any streamable property. For Unicode, it returns
        #   the first 255 characters by using a null-terminated Unicode string. It does not return any of the message
        #   body formats or the recipient lists.
        self.is_complex = is_complex
        self.is_required = is_required
        self.is_read_only = is_read_only
        # Set this for fields that raise ErrorInvalidPropertyUpdateSentMessage on update after send
        self.is_read_only_after_send = is_read_only_after_send

    def clean(self, value):
        if value is None:
            if self.is_required and self.default is None:
                raise ValueError("'%s' is a required field with no default" % self.name)
            if self.is_list and self.value_cls == Attachment:
                return []
            return self.default
        if self.value_cls == EWSDateTime and not getattr(value, 'tzinfo'):
            raise ValueError("Field '%s' must be timezone aware" % self.name)
        if self.value_cls == Choice and value not in self.choices:
            raise ValueError("Field '%s' value '%s' is not a valid choice (%s)" % (self.name, value, self.choices))

        # For value_cls that are subclasses of string types, convert simple string values to their subclass equivalent
        # (e.g. str to Body and str to Subject) so we can call value.clean()
        if issubclass(self.value_cls, string_types) and self.value_cls != string_type \
                and not isinstance(value, self.value_cls):
            value = self.value_cls(value)
        elif self.value_cls == Mailbox:
            if self.is_list:
                value = [Mailbox(email_address=s) if isinstance(s, string_types) else s for s in value]
            elif isinstance(value, string_types):
                value = Mailbox(email_address=value)
        elif self.value_cls == Attendee:
            if self.is_list:
                value = [Attendee(mailbox=Mailbox(email_address=s), response_type='Accept')
                         if isinstance(s, string_types) else s for s in value]
            elif isinstance(value, string_types):
                value = Attendee(mailbox=Mailbox(email_address=value), response_type='Accept')

        if self.is_list:
            if not isinstance(value, (tuple, list)):
                raise ValueError("Field '%s' value '%s' must be a list" % (self.name, value))
            for v in value:
                if not isinstance(v, self.value_cls):
                    raise TypeError('Field %s value "%s" must be of type %s' % (self.name, v, self.value_cls))
                if hasattr(v, 'clean'):
                    v.clean()
        else:
            if not isinstance(value, self.value_cls):
                raise ValueError("Field '%s' value '%s' must be of type %s" % (self.name, value, self.value_cls))
            if hasattr(value, 'clean'):
                value.clean()
        return value

    def from_xml(self, elem):
        if self.is_list:
            iter_elem = elem.find(self.response_tag())
            if self.value_cls == string_type:
                if iter_elem is not None:
                    return get_xml_attrs(iter_elem, '{%s}String' % TNS)
            elif self.value_cls == Attachment:
                # Look for both FileAttachment and ItemAttachment
                if iter_elem is not None:
                    attachments = []
                    for att_type in (FileAttachment, ItemAttachment):
                        attachments.extend(
                            [att_type.from_xml(e) for e in iter_elem.findall(att_type.response_tag())]
                        )
                    return attachments
            elif issubclass(self.value_cls, EWSElement):
                if iter_elem is not None:
                    return [self.value_cls.from_xml(e) for e in iter_elem.findall(self.value_cls.response_tag())]
            else:
                assert False, 'Field %s type %s not supported' % (self.name, self.value_cls)
        else:
            field_elem = elem.find(self.response_tag())
            if issubclass(self.value_cls, (bool, int, Decimal, string_type, EWSDateTime)):
                val = None if field_elem is None else field_elem.text or None
                if val is not None:
                    try:
                        val = xml_text_to_value(value=val, value_type=self.value_cls)
                    except ValueError:
                        pass
                    except KeyError:
                        assert False, 'Field %s type %s not supported' % (self.name, self.value_cls)
                    if self.name == 'body':
                        body_type = field_elem.get('BodyType')
                        try:
                            val = {
                                Body.body_type: lambda v: Body(v),
                                HTMLBody.body_type: lambda v: HTMLBody(v),
                            }[body_type](val)
                        except KeyError:
                            assert False, "Unknown BodyType '%s'" % body_type
                    return val
            elif issubclass(self.value_cls, EWSElement):
                sub_elem = elem.find(self.response_tag())
                if sub_elem is not None:
                    if self.value_cls == Mailbox:
                        # We want the nested Mailbox, not the wrapper element
                        return self.value_cls.from_xml(sub_elem.find(Mailbox.response_tag()))
                    else:
                        return self.value_cls.from_xml(sub_elem)
            else:
                assert False, 'Field %s type %s not supported' % (self.name, self.value_cls)
        return self.default

    def to_xml(self, value, version):
        raise NotImplementedError()

    def field_uri_xml(self):
        raise NotImplementedError()

    def set_field_xml(self):
        raise NotImplementedError()

    def request_tag(self):
        raise NotImplementedError()

    def response_tag(self):
        raise NotImplementedError()

    def __eq__(self, other):
        return hash(self) == hash(other)

    def __hash__(self):
        raise NotImplementedError()

    def __repr__(self):
        return self.__class__.__name__ + repr((self.name, self.value_cls))


class SimpleField(Field):
    def __init__(self, *args, **kwargs):
        field_uri = kwargs.pop('field_uri')
        super(SimpleField, self).__init__(*args, **kwargs)
        # See all valid FieldURI values at https://msdn.microsoft.com/en-us/library/office/aa494315(v=exchg.150).aspx
        # field_uri_prefix is the prefix part of the FieldURI.
        self.field_uri = field_uri
        self.field_uri_prefix, self.field_uri_postfix = field_uri.split(':')

    def to_xml(self, value, version):
        field_elem = create_element(self.request_tag())
        if self.name == 'body':
            body_type = HTMLBody.body_type if isinstance(value, HTMLBody) else Body.body_type
            field_elem.set('BodyType', body_type)
        return set_xml_value(field_elem, value, version=version)

    def field_uri_xml(self):
        return create_element('t:FieldURI', FieldURI=self.field_uri)

    def request_tag(self):
        return 't:%s' % self.field_uri_postfix

    def response_tag(self):
        return '{%s}%s' % (TNS, self.field_uri_postfix)

    def __hash__(self):
        return hash(self.field_uri)


class IndexedField(SimpleField):
    PARENT_ELEMENT_NAME = None
    VALUE_CLS = None

    def __init__(self, *args, **kwargs):
        super(IndexedField, self).__init__(*args, **kwargs)
        assert issubclass(self.value_cls, IndexedElement)

    def field_uri_xml(self, label=None, subfield=None):
        if not label:
            # Return elements for all labels
            elems = []
            for l in self.value_cls.LABELS:
                elem = self.field_uri_xml(label=l)
                if isinstance(elem, list):
                    elems.extend(elem)
                else:
                    elems.append(elem)
            return elems
        if self.value_cls.SUB_FIELD_ELEMENT_NAMES:
            if not subfield:
                # Return elements for all sub-fields
                return [self.field_uri_xml(label=label, subfield=s)
                        for s in self.value_cls.SUB_FIELD_ELEMENT_NAMES.keys()]
            assert subfield in self.value_cls.SUB_FIELD_ELEMENT_NAMES, (subfield, self.value_cls.SUB_FIELD_ELEMENT_NAMES)
            field_uri = '%s:%s' % (self.field_uri, self.value_cls.SUB_FIELD_ELEMENT_NAMES[subfield])
        else:
            field_uri = self.field_uri
        assert label in self.value_cls.LABELS, (label, self.value_cls.LABELS)
        return create_element('t:IndexedFieldURI', FieldURI=field_uri, FieldIndex=label)

    def to_xml(self, value, version):
        return set_xml_value(create_element('t:%s' % self.PARENT_ELEMENT_NAME), value, version)

    @classmethod
    def response_tag(cls):
        return '{%s}%s' % (TNS, cls.PARENT_ELEMENT_NAME)

    def __hash__(self):
        return hash(self.field_uri)


class EmailAddressField(IndexedField):
    PARENT_ELEMENT_NAME = 'EmailAddresses'


class PhoneNumberField(IndexedField):
    PARENT_ELEMENT_NAME = 'PhoneNumbers'


class PhysicalAddressField(IndexedField):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa564323(v=exchg.150).aspx
    PARENT_ELEMENT_NAME = 'PhysicalAddresses'


class ExtendedPropertyField(Field):
    def __init__(self, *args, **kwargs):
        super(ExtendedPropertyField, self).__init__(*args, **kwargs)
        assert issubclass(self.value_cls, ExtendedProperty)

    def clean(self, value):
        if value is None:
            if self.is_required:
                raise ValueError("'%s' is a required field" % self.name)
            return self.default
        if not isinstance(value, self.value_cls):
            # Allow keeping ExtendedProperty field values as their simple Python type, but run clean() anyway
            tmp = self.value_cls(value)
            tmp.clean()
            return value
        value.clean()
        return value

    def field_uri_xml(self):
        elem = create_element('t:ExtendedFieldURI')
        cls = self.value_cls
        if cls.distinguished_property_set_id:
            elem.set('DistinguishedPropertySetId', cls.distinguished_property_set_id)
        if cls.property_set_id:
            elem.set('PropertySetId', cls.property_set_id)
        if cls.property_tag:
            hex_val = int(cls.property_tag, base=16) if isinstance(cls.property_tag, string_types) else cls.property_tag
            elem.set('PropertyTag', hex(hex_val))
        if cls.property_name:
            elem.set('PropertyName', cls.property_name)
        if cls.property_id:
            elem.set('PropertyId', value_to_xml_text(cls.property_id))
        elem.set('PropertyType', cls.property_type)
        return elem

    def from_xml(self, elem):
        extended_properties = elem.findall(self.response_tag())
        return self.value_cls.from_xml(extended_properties)

    def to_xml(self, value, version):
        extended_property = create_element(self.request_tag())
        set_xml_value(extended_property, self.field_uri_xml(), version=version)
        if isinstance(value, self.value_cls):
            set_xml_value(extended_property, value, version=version)
        else:
            # Allow keeping ExtendedProperty field values as their simple Python type
            set_xml_value(extended_property, self.value_cls(value), version=version)
        return extended_property

    def request_tag(self):
        return 't:%s' % ExtendedProperty.ELEMENT_NAME

    @classmethod
    def response_tag(cls):
        return '{%s}%s' % (TNS, ExtendedProperty.ELEMENT_NAME)

    def __hash__(self):
        return hash(self.name)


class Item(EWSElement):
    ELEMENT_NAME = 'Item'

    # ITEM_FIELDS is an ordered list of attributes supported by this item class. Not all possible attributes are
    # supported. See full list at
    # https://msdn.microsoft.com/en-us/library/office/aa580790(v=exchg.150).aspx

    # 'extern_id' is not a native EWS Item field. We use it for identification when item originates in an external
    # system. The field is implemented as an extended property on the Item.
    ITEM_FIELDS = (
        SimpleField('item_id', field_uri='item:ItemId', value_cls=string_type, is_read_only=True),
        SimpleField('changekey', field_uri='item:ChangeKey', value_cls=string_type, is_read_only=True),
        # TODO: MimeContent actually supports writing, but is still untested
        SimpleField('mime_content', field_uri='item:MimeContent', value_cls=MimeContent, is_read_only=True),
        SimpleField('subject', field_uri='item:Subject', value_cls=Subject),
        SimpleField('sensitivity', field_uri='item:Sensitivity', value_cls=Choice,
                    choices={'Normal', 'Personal', 'Private', 'Confidential'}, is_required=True, default='Normal'),
        SimpleField('body', field_uri='item:Body', value_cls=Body, is_complex=True),  # Body or HTMLBody
        SimpleField('attachments', field_uri='item:Attachments', value_cls=Attachment, default=None, is_list=True,
                    is_complex=True),  # ItemAttachment or FileAttachment
        SimpleField('datetime_received', field_uri='item:DateTimeReceived', value_cls=EWSDateTime, is_read_only=True),
        SimpleField('categories', field_uri='item:Categories', value_cls=string_type, is_list=True),
        SimpleField('importance', field_uri='item:Importance', value_cls=Choice,
                    choices={'Low', 'Normal', 'High'}, is_required=True, default='Normal'),
        SimpleField('is_draft', field_uri='item:IsDraft', value_cls=bool, is_read_only=True),
        SimpleField('headers', field_uri='item:InternetMessageHeaders', value_cls=MessageHeader, is_list=True,
                    is_read_only=True),
        SimpleField('datetime_sent', field_uri='item:DateTimeSent', value_cls=EWSDateTime, is_read_only=True),
        SimpleField('datetime_created', field_uri='item:DateTimeCreated', value_cls=EWSDateTime, is_read_only=True),
        SimpleField('reminder_is_set', field_uri='item:ReminderIsSet', value_cls=bool, is_required=True, default=False),
        # ExtendedProperty fields go here
        SimpleField('last_modified_name', field_uri='item:LastModifiedName', value_cls=string_type, is_read_only=True),
        SimpleField('last_modified_time', field_uri='item:LastModifiedTime', value_cls=EWSDateTime, is_read_only=True),
    )
    ITEM_FIELDS_MAP = {f.name: f for f in ITEM_FIELDS}

    def __init__(self, **kwargs):
        # 'account' is optional but allows calling 'send()' and 'delete()'
        # 'folder' is optional but allows calling 'save()'
        from .account import Account
        self.account = kwargs.pop('account', None)
        if self.account is not None:
            assert isinstance(self.account, Account)
        self.folder = kwargs.pop('folder', None)
        if self.folder is not None:
            assert isinstance(self.folder, Folder)

        for f in self.ITEM_FIELDS:
            setattr(self, f.name, kwargs.pop(f.name, None))
        for k, v in kwargs.items():
            raise TypeError("'%s' is an invalid keyword argument for this function" % k)
        # self.clean()
        if self.attachments is None:
            self.attachments = []
        for a in self.attachments:
            if a.parent_item:
                assert a.parent_item is self  # An attachment cannot refer to 'self' in __init__
            else:
                a.parent_item = self
            self.attach(self.attachments)

    def clean(self):
        for f in self.ITEM_FIELDS:
            val = getattr(self, f.name)
            setattr(self, f.name, f.clean(val))

    def save(self, conflict_resolution=AUTO_RESOLVE, send_meeting_invitations=SEND_TO_NONE):
        item = self._save(message_disposition=SAVE_ONLY, conflict_resolution=conflict_resolution,
                          send_meeting_invitations=send_meeting_invitations)
        if self.item_id:
            # _save() returns tuple()
            item_id, changekey = item
            assert self.item_id == item_id
            assert self.changekey != changekey
            self.changekey = changekey
        else:
            # _save() returns Item
            self.item_id, self.changekey = item.item_id, item.changekey
            for old_att, new_att in zip(self.attachments, item.attachments):
                assert old_att.attachment_id is None
                assert new_att.attachment_id is not None
                old_att.attachment_id = new_att.attachment_id
        return self

    def _save(self, message_disposition, conflict_resolution, send_meeting_invitations):
        if not self.account:
            raise ValueError('Item must have an account')
        if self.item_id:
            assert self.changekey
            update_fieldnames = []
            for f in self.ITEM_FIELDS:
                if f.name == 'attachments':
                    # Attachments are handled separately after item creation
                    continue
                if f.is_read_only:
                    # These cannot be changed
                    continue
                if not self.is_draft and f.is_read_only_after_send:
                    # These cannot be changed when the item is no longer a draft
                    continue
                if f.is_required and getattr(self, f.name) is None:
                    continue
                update_fieldnames.append(f.name)
            # bulk_update() returns a tuple
            res = self.account.bulk_update(
                items=[(self, update_fieldnames)], message_disposition=message_disposition,
                conflict_resolution=conflict_resolution,
                send_meeting_invitations_or_cancellations=send_meeting_invitations)
            if message_disposition == SEND_AND_SAVE_COPY:
                assert len(res) == 0
                return None
            else:
                if not res:
                    raise ValueError('Item disappeared')
                assert len(res) == 1, res
                if isinstance(res[0], Exception):
                    raise res[0]
                return res[0]
        else:
            # bulk_create() returns an Item because we want to return item_id on both main item *and* attachments
            res = self.account.bulk_create(
                items=[self], folder=self.folder, message_disposition=message_disposition,
                send_meeting_invitations=send_meeting_invitations)
            if message_disposition in (SEND_ONLY, SEND_AND_SAVE_COPY):
                assert len(res) == 0
                return None
            else:
                assert len(res) == 1, res
                if isinstance(res[0], Exception):
                    raise res[0]
                return res[0]

    def refresh(self):
        # Updates the item based on fresh data from EWS
        if not self.account:
            raise ValueError('Item must have an account')
        if not self.item_id:
            raise ValueError('Item must have an ID')
        res = list(self.account.fetch(ids=[self]))
        if not res:
            raise ValueError('Item disappeared')
        assert len(res) == 1, res
        if isinstance(res[0], Exception):
            raise res[0]
        fresh_item = res[0]
        assert self.item_id == fresh_item.item_id
        assert self.changekey == fresh_item.changekey
        for f in self.ITEM_FIELDS:
            setattr(self, f.name, getattr(fresh_item, f.name))

    def move(self, to_folder):
        if not self.account:
            raise ValueError('Item must have an account')
        if not self.item_id:
            raise ValueError('Item must have an ID')
        res = self.account.bulk_move(ids=[self], to_folder=to_folder)
        if not res:
            raise ValueError('Item disappeared')
        assert len(res) == 1, res
        if isinstance(res[0], Exception):
            raise res[0]
        self.item_id, self.changekey = res[0]
        self.folder = to_folder

    def move_to_trash(self, send_meeting_cancellations=SEND_TO_NONE,
                      affected_task_occurrences=SPECIFIED_OCCURRENCE_ONLY, suppress_read_receipts=True):
        # Delete and move to the trash folder.
        self._delete(delete_type=MOVE_TO_DELETED_ITEMS, send_meeting_cancellations=send_meeting_cancellations,
                     affected_task_occurrences=affected_task_occurrences, suppress_read_receipts=suppress_read_receipts)
        self.item_id, self.changekey = None, None
        self.folder = self.folder.account.trash

    def soft_delete(self, send_meeting_cancellations=SEND_TO_NONE, affected_task_occurrences=SPECIFIED_OCCURRENCE_ONLY,
                    suppress_read_receipts=True):
        # Delete and move to the dumpster, if it is enabled.
        self._delete(delete_type=SOFT_DELETE, send_meeting_cancellations=send_meeting_cancellations,
                     affected_task_occurrences=affected_task_occurrences, suppress_read_receipts=suppress_read_receipts)
        self.item_id, self.changekey = None, None
        self.folder = self.folder.account.recoverable_deleted_items

    def delete(self, send_meeting_cancellations=SEND_TO_NONE, affected_task_occurrences=SPECIFIED_OCCURRENCE_ONLY,
               suppress_read_receipts=True):
        # Remove the item permanently. No copies are stored anywhere.
        self._delete(delete_type=HARD_DELETE, send_meeting_cancellations=send_meeting_cancellations,
                     affected_task_occurrences=affected_task_occurrences, suppress_read_receipts=suppress_read_receipts)
        self.item_id, self.changekey, self.folder = None, None, None

    def _delete(self, delete_type, send_meeting_cancellations, affected_task_occurrences, suppress_read_receipts):
        if not self.account:
            raise ValueError('Item must have an account')
        if not self.item_id:
            raise ValueError('Item must have an ID')
        res = self.account.bulk_delete(
            ids=[self], delete_type=delete_type, send_meeting_cancellations=send_meeting_cancellations,
            affected_task_occurrences=affected_task_occurrences, suppress_read_receipts=suppress_read_receipts)
        if not res:
            raise ValueError('Item disappeared')
        assert len(res) == 1, res
        if isinstance(res[0], Exception):
            raise res[0]

    def attach(self, attachments):
        """Add an attachment, or a list of attachments, to this item. If the item has already been saved, the
        attachments will be created on the server immediately. If the item has not yet been saved, the attachments will
        be created on the server the item is saved.

        Adding attachments to an existing item will update the changekey of the item.
        """
        if isinstance(attachments, Attachment):
            attachments = [attachments]
        for a in attachments:
            assert isinstance(a, Attachment)
            if not a.parent_item:
                a.parent_item = self
            if self.item_id and not a.attachment_id:
                # Already saved object. Attach the attachment server-side now
                a.attach()
            if a not in self.attachments:
                self.attachments.append(a)

    def detach(self, attachments):
        """Remove an attachment, or a list of attachments, from this item. If the item has already been saved, the
        attachments will be deleted on the server immediately. If the item has not yet been saved, the attachments will
        simply not be created on the server the item is saved.

        Removing attachments from an existing item will update the changekey of the item.
        """
        if isinstance(attachments, Attachment):
            attachments = [attachments]
        for a in attachments:
            assert isinstance(a, Attachment)
            assert a.parent_item is self
            if self.item_id:
                # Item is already created. Detach  the attachment server-side now
                a.detach()
            if a in self.attachments:
                self.attachments.remove(a)

    @classmethod
    def fieldnames(cls):
        # Return non-ID field names
        return set(f.name for f in cls.ITEM_FIELDS if f.name not in ('item_id', 'changekey'))

    @classmethod
    def required_fields(cls):
        return set(f for f in cls.ITEM_FIELDS if f.is_required)

    @classmethod
    def readonly_fields(cls):
        return set(f for f in cls.ITEM_FIELDS if f.is_read_only)

    @classmethod
    def readonly_after_send_fields(cls):
        return set(f for f in cls.ITEM_FIELDS if f.is_read_only_after_send)

    @classmethod
    def complex_fields(cls):
        return set(f for f in cls.ITEM_FIELDS if f.is_complex)

    @classmethod
    def id_from_xml(cls, elem):
        id_elem = elem.find(ItemId.response_tag())
        if id_elem is None:
            return None, None
        return id_elem.get(ItemId.ID_ATTR), id_elem.get(ItemId.CHANGEKEY_ATTR)

    @classmethod
    def from_xml(cls, elem, account=None, folder=None):
        assert elem.tag == cls.response_tag(), (cls, elem.tag, cls.response_tag())
        item_id, changekey = cls.id_from_xml(elem)
        kwargs = {f.name: f.from_xml(elem) for f in cls.ITEM_FIELDS if f.name not in ('item_id', 'changekey')}
        elem.clear()
        return cls(item_id=item_id, changekey=changekey, account=account, folder=folder, **kwargs)

    def to_xml(self, version):
        self.clean()
        # WARNING: The order of addition of XML elements is VERY important. Exchange expects XML elements in a
        # specific, non-documented order and will fail with meaningless errors if the order is wrong.
        i = create_element(self.request_tag())
        for f in self.ITEM_FIELDS:
            if f.is_read_only:
                continue
            value = getattr(self, f.name)
            if value is None:
                continue
            if f.is_list and not value:
                continue
            i.append(f.to_xml(value, version=version))
        return i

    @classmethod
    def register(cls, attr_name, attr_cls):
        """
        Register a custom extended property in this item class so they can be accessed just like any other attribute
        """
        if attr_name in cls.ITEM_FIELDS_MAP:
            raise AttributeError("%s' is already registered" % attr_name)
        if not issubclass(attr_cls, ExtendedProperty):
            raise ValueError("'%s' must be a subclass of ExtendedProperty" % attr_cls)
        # Find the correct index for the extended property and insert the new field. See Item.ITEM_FIELDS for comment
        updated_item_fields = []
        for f in cls.ITEM_FIELDS:
            updated_item_fields.append(f)
            if f.name == 'reminder_is_set':
                # This is a bit hacky and will need to change if we add new item fields after 'reminder_is_set'
                updated_item_fields.append(ExtendedPropertyField(attr_name, value_cls=attr_cls))
        cls.ITEM_FIELDS = tuple(updated_item_fields)
        # Rebuild map
        cls.ITEM_FIELDS_MAP = {f.name: f for f in cls.ITEM_FIELDS}

    @classmethod
    def deregister(cls, attr_name):
        """
        De-register an extended property that has been registered with register()
        """
        # TODO: ExtendedProperty goes in between <HasAttachments/><ExtendedProperty/><Culture/>
        # TODO: See https://msdn.microsoft.com/en-us/library/office/aa580790(v=exchg.150).aspx
        if attr_name not in cls.ITEM_FIELDS_MAP:
            raise AttributeError("%s' is not registered" % attr_name)
        if not isinstance(cls.ITEM_FIELDS_MAP[attr_name], ExtendedPropertyField):
            raise AttributeError("'%s' is not registered as an ExtendedProperty" % attr_name)
        cls.ITEM_FIELDS = tuple(f for f in cls.ITEM_FIELDS if f.name != attr_name)
        # Rebuild map
        cls.ITEM_FIELDS_MAP = {f.name: f for f in cls.ITEM_FIELDS}

    def __eq__(self, other):
        if isinstance(other, tuple):
            item_id, changekey = other
            return self.item_id == item_id and self.changekey == changekey
        return self.item_id == other.item_id and self.changekey == other.changekey

    def __hash__(self):
        # If we have an item_id and changekey, use that as key. Else return a hash of all attributes
        if self.item_id:
            return hash((self.item_id, self.changekey))
        return hash(tuple(
            tuple(tuple(getattr(self, f.name) or ()) if f.is_list else getattr(self, f.name) for f in self.ITEM_FIELDS)
        ))

    def __str__(self):
        return '\n'.join('%s: %s' % (f, getattr(self, f.name)) for f in self.ITEM_FIELDS)

    def __repr__(self):
        return self.__class__.__name__ + '(%s)' % ', '.join(
            '%s=%s' % (f.name, repr(getattr(self, f.name))) for f in self.ITEM_FIELDS
        )


@python_2_unicode_compatible
class BulkCreateResult(Item):
    ITEM_FIELDS = (
        SimpleField('item_id', field_uri='item:ItemId', value_cls=string_type, is_read_only=True, is_required=True),
        SimpleField('changekey', field_uri='item:ChangeKey', value_cls=string_type, is_read_only=True, is_required=True),
        SimpleField('attachments', field_uri='item:Attachments', value_cls=Attachment, default=(), is_list=True,
                    is_complex=True),  # ItemAttachment or FileAttachment
    )
    ITEM_FIELDS_MAP = {f.name: f for f in ITEM_FIELDS}

    @classmethod
    def from_xml(cls, elem, account=None, folder=None):
        item_id, changekey = cls.id_from_xml(elem)
        kwargs = {f.name: f.from_xml(elem) for f in cls.ITEM_FIELDS if f.name not in ('item_id', 'changekey')}
        elem.clear()
        return cls(item_id=item_id, changekey=changekey, account=account, folder=folder, **kwargs)


@python_2_unicode_compatible
class CalendarItem(Item):
    """
    Models a calendar item. Not all attributes are supported. See full list at
    https://msdn.microsoft.com/en-us/library/office/aa564765(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'CalendarItem'
    ITEM_FIELDS = Item.ITEM_FIELDS + (
        SimpleField('start', field_uri='calendar:Start', value_cls=EWSDateTime, is_required=True),
        SimpleField('end', field_uri='calendar:End', value_cls=EWSDateTime, is_required=True),
        SimpleField('is_all_day', field_uri='calendar:IsAllDayEvent', value_cls=bool, is_required=True, default=False),
        # TODO: The 'WorkingElsewhere' status was added in Exchange2015 but we don't support versioned choices yet
        SimpleField('legacy_free_busy_status', field_uri='calendar:LegacyFreeBusyStatus', value_cls=Choice,
                    choices={'Free', 'Tentative', 'Busy', 'OOF', 'NoData'}, is_required=True, default='Busy'),
        SimpleField('location', field_uri='calendar:Location', value_cls=Location),
        SimpleField('organizer', field_uri='calendar:Organizer', value_cls=Mailbox, is_read_only=True),
        SimpleField('required_attendees', field_uri='calendar:RequiredAttendees', value_cls=Attendee, is_list=True),
        SimpleField('optional_attendees', field_uri='calendar:OptionalAttendees', value_cls=Attendee, is_list=True),
        SimpleField('resources', field_uri='calendar:Resources', value_cls=Attendee, is_list=True),
    )
    ITEM_FIELDS_MAP = {f.name: f for f in ITEM_FIELDS}

    def to_xml(self, version):
        # WARNING: The order of addition of XML elements is VERY important. Exchange expects XML elements in a
        # specific, non-documented order and will fail with meaningless errors if the order is wrong.
        i = super(CalendarItem, self).to_xml(version=version)
        if version.build < EXCHANGE_2010:
            i.append(create_element('t:MeetingTimeZone', TimeZoneName=self.start.tzinfo.ms_id))
        else:
            i.append(create_element('t:StartTimeZone', Id=self.start.tzinfo.ms_id, Name=self.start.tzinfo.ms_name))
            i.append(create_element('t:EndTimeZone', Id=self.end.tzinfo.ms_id, Name=self.end.tzinfo.ms_name))
        return i


class Message(Item):
    ELEMENT_NAME = 'Message'
    # Supported attrs: see https://msdn.microsoft.com/en-us/library/office/aa494306(v=exchg.150).aspx
    # TODO: This list is incomplete
    ITEM_FIELDS = Item.ITEM_FIELDS + (
        SimpleField('to_recipients', field_uri='message:ToRecipients', value_cls=Mailbox, is_list=True,
                    is_read_only_after_send=True),
        SimpleField('cc_recipients', field_uri='message:CcRecipients', value_cls=Mailbox, is_list=True,
                    is_read_only_after_send=True),
        SimpleField('bcc_recipients', field_uri='message:BccRecipients', value_cls=Mailbox, is_list=True,
                    is_read_only_after_send=True),
        SimpleField('is_read_receipt_requested', field_uri='message:IsReadReceiptRequested', value_cls=bool,
                    is_required=True, default=False, is_read_only_after_send=True),
        SimpleField('is_delivery_receipt_requested', field_uri='message:IsDeliveryReceiptRequested', value_cls=bool,
                    is_required=True, default=False, is_read_only_after_send=True),
        SimpleField('sender', field_uri='message:Sender', value_cls=Mailbox, is_read_only=True,
                    is_read_only_after_send=True),
        # We can't use fieldname 'from' since it's a Python keyword
        SimpleField('author', field_uri='message:From', value_cls=Mailbox, is_read_only_after_send=True),
        SimpleField('is_read', field_uri='message:IsRead', value_cls=bool, is_required=True, default=False),
        SimpleField('is_response_requested', field_uri='message:IsResponseRequested', value_cls=bool, default=False,
                    is_required=True),
        SimpleField('reply_to', field_uri='message:ReplyTo', value_cls=Mailbox, is_list=True,
                    is_read_only_after_send=True),
        SimpleField('message_id', field_uri='message:InternetMessageId', value_cls=string_type, is_read_only=True,
                    is_read_only_after_send=True),
    )
    ITEM_FIELDS_MAP = {f.name: f for f in ITEM_FIELDS}

    def send(self, save_copy=True, copy_to_folder=None, conflict_resolution=AUTO_RESOLVE,
             send_meeting_invitations=SEND_TO_NONE):
        # Only sends a message. The message can either be an existing draft stored in EWS or a new message that does
        # not yet exist in EWS.
        if not self.account:
            raise ValueError('Item must have an account')
        if self.item_id:
            res = self.account.bulk_send(ids=[self], save_copy=save_copy, copy_to_folder=copy_to_folder)
            if not res:
                raise ValueError('Item disappeared')
            assert len(res) == 1, res
            if isinstance(res[0], Exception):
                raise res[0]
            # The item will be deleted from the original folder
            self.item_id, self.changekey = None, None
            self.folder = copy_to_folder
        else:
            # New message
            if copy_to_folder:
                if not save_copy:
                    raise AttributeError("'save_copy' must be True when 'copy_to_folder' is set")
                # This would better be done via send_and_save() but lets just support it here
                self.folder = copy_to_folder
                return self.send_and_save(conflict_resolution=conflict_resolution,
                                          send_meeting_invitations=send_meeting_invitations)
            assert copy_to_folder is None
            res = self._save(message_disposition=SEND_ONLY, conflict_resolution=conflict_resolution,
                             send_meeting_invitations=send_meeting_invitations)
            assert res is None

    def send_and_save(self, conflict_resolution=AUTO_RESOLVE, send_meeting_invitations=SEND_TO_NONE):
        # Sends Message and saves a copy in the parent folder. Does not return an ItemId.
        res = self._save(message_disposition=SEND_AND_SAVE_COPY, conflict_resolution=conflict_resolution,
                         send_meeting_invitations=send_meeting_invitations)
        assert res is None


class Task(Item):
    ELEMENT_NAME = 'Task'
    NOT_STARTED = 'NotStarted'
    COMPLETED = 'Completed'
    # Supported attrs: see https://msdn.microsoft.com/en-us/library/office/aa563930(v=exchg.150).aspx
    # TODO: This list is incomplete
    ITEM_FIELDS = Item.ITEM_FIELDS + (
        SimpleField('actual_work', field_uri='task:ActualWork', value_cls=int),
        SimpleField('assigned_time', field_uri='task:AssignedTime', value_cls=EWSDateTime, is_read_only=True),
        SimpleField('billing_information', field_uri='task:BillingInformation', value_cls=string_type),
        SimpleField('change_count', field_uri='task:ChangeCount', value_cls=int, is_read_only=True),
        SimpleField('companies', field_uri='task:Companies', value_cls=string_type, is_list=True),
        SimpleField('contacts', field_uri='task:Contacts', value_cls=string_type, is_list=True),
        SimpleField('delegation_state', field_uri='task:DelegationState', value_cls=Choice,
                    choices={'NoMatch', 'OwnNew', 'Owned', 'Accepted', 'Declined', 'Max'}, is_read_only=True),
        SimpleField('delegator', field_uri='task:Delegator', value_cls=string_type, is_read_only=True),
        # 'complete_date' can be set, but is ignored by the server, which sets it to now()
        SimpleField('complete_date', field_uri='task:CompleteDate', value_cls=EWSDateTime, is_read_only=True),
        SimpleField('due_date', field_uri='task:DueDate', value_cls=EWSDateTime),
        SimpleField('is_complete', field_uri='task:IsComplete', value_cls=bool, is_read_only=True),
        SimpleField('is_recurring', field_uri='task:IsRecurring', value_cls=bool, is_read_only=True),
        SimpleField('is_team_task', field_uri='task:IsTeamTask', value_cls=bool, is_read_only=True),
        SimpleField('mileage', field_uri='task:Mileage', value_cls=string_type),
        SimpleField('owner', field_uri='task:Owner', value_cls=string_type, is_read_only=True),
        SimpleField('percent_complete', field_uri='task:PercentComplete', value_cls=Decimal),
        SimpleField('start_date', field_uri='task:StartDate', value_cls=EWSDateTime),
        SimpleField('status', field_uri='task:Status', value_cls=Choice, choices={
            NOT_STARTED, 'InProgress', COMPLETED, 'WaitingOnOthers', 'Deferred'
        }, is_required=True, default=NOT_STARTED),
        SimpleField('status_description', field_uri='task:StatusDescription', value_cls=string_type, is_read_only=True),
        SimpleField('total_work', field_uri='task:TotalWork', value_cls=int),
    )
    ITEM_FIELDS_MAP = {f.name: f for f in ITEM_FIELDS}

    def clean(self):
        super(Task, self).clean()
        if self.due_date and self.start_date and self.due_date < self.start_date:
            log.warning("'due_date' must be greater than 'start_date' (%s vs %s). Resetting 'due_date'",
                        self.due_date, self.start_date)
            self.due_date = self.start_date
        if self.complete_date:
            if self.status != self.COMPLETED:
                log.warning("'status' must be '%s' when 'complete_date' is set (%s). Resetting",
                            self.COMPLETED, self.status)
                self.status = self.COMPLETED
            now = UTC_NOW()
            if (self.complete_date - now).total_seconds() > 120:
                # 'complete_date' can be set automatically by the server. Allow some grace between local and server time
                log.warning("'complete_date' must be in the past (%s vs %s). Resetting", self.complete_date, now)
                self.complete_date = now
            if self.start_date and self.complete_date < self.start_date:
                log.warning("'complete_date' must be greater than 'start_date' (%s vs %s). Resetting",
                            self.complete_date, self.start_date)
                self.complete_date = self.start_date
        if self.percent_complete is not None:
            assert isinstance(self.percent_complete, Decimal)
            assert Decimal(0) <= self.percent_complete <= Decimal(100), self.percent_complete
            if self.status == self.COMPLETED and self.percent_complete != Decimal(100):
                # percent_complete must be 100% if task is complete
                log.warning("'percent_complete' must be 100 when 'status' is '%s' (%s). Resetting",
                            self.COMPLETED, self.percent_complete)
                self.percent_complete = Decimal(100)
            elif self.status == self.NOT_STARTED and self.percent_complete != Decimal(0):
                # percent_complete must be 0% if task is not started
                log.warning("'percent_complete' must be 0 when 'status' is '%s' (%s). Resetting",
                            self.NOT_STARTED, self.percent_complete)
                self.percent_complete = Decimal(0)


class Contact(Item):
    ELEMENT_NAME = 'Contact'
    # Supported attrs: see https://msdn.microsoft.com/en-us/library/office/aa581315(v=exchg.150).aspx
    # TODO: This list is incomplete
    ITEM_FIELDS = Item.ITEM_FIELDS + (
        SimpleField('file_as', field_uri='contacts:FileAs', value_cls=string_type),
        SimpleField('file_as_mapping', field_uri='contacts:FileAsMapping', value_cls=Choice, choices={
            'None', 'LastCommaFirst', 'FirstSpaceLast', 'Company', 'LastCommaFirstCompany', 'CompanyLastFirst',
            'LastFirst', 'LastFirstCompany', 'CompanyLastCommaFirst', 'LastFirstSuffix', 'LastSpaceFirstCompany',
            'CompanyLastSpaceFirst', 'LastSpaceFirst', 'DisplayName', 'FirstName', 'LastFirstMiddleSuffix', 'LastName',
            'Empty',
        }),
        SimpleField('display_name', field_uri='contacts:DisplayName', value_cls=string_type, is_required=True, default=''),
        SimpleField('given_name', field_uri='contacts:GivenName', value_cls=string_type),
        SimpleField('initials', field_uri='contacts:Initials', value_cls=string_type),
        SimpleField('middle_name', field_uri='contacts:MiddleName', value_cls=string_type),
        SimpleField('nickname', field_uri='contacts:Nickname', value_cls=string_type),
        SimpleField('company_name', field_uri='contacts:CompanyName', value_cls=string_type),
        EmailAddressField('email_addresses', field_uri='contacts:EmailAddress', value_cls=EmailAddress, is_list=True),
        PhysicalAddressField('physical_addresses', field_uri='contacts:PhysicalAddress', value_cls=PhysicalAddress,
                             is_list=True),
        PhoneNumberField('phone_numbers', field_uri='contacts:PhoneNumber', value_cls=PhoneNumber, is_list=True),
        SimpleField('assistant_name', field_uri='contacts:AssistantName', value_cls=string_type),
        SimpleField('birthday', field_uri='contacts:Birthday', value_cls=EWSDateTime),
        SimpleField('business_homepage', field_uri='contacts:BusinessHomePage', value_cls=AnyURI),
        SimpleField('companies', field_uri='contacts:Companies', value_cls=string_type, is_list=True),
        SimpleField('department', field_uri='contacts:Department', value_cls=string_type),
        SimpleField('generation', field_uri='contacts:Generation', value_cls=string_type),
        # SimpleField('im_addresses', field_uri='contacts:ImAddresses', value_cls=ImAddress, is_list=True),
        SimpleField('job_title', field_uri='contacts:JobTitle', value_cls=string_type),
        SimpleField('manager', field_uri='contacts:Manager', value_cls=string_type),
        SimpleField('mileage', field_uri='contacts:Mileage', value_cls=string_type),
        SimpleField('office', field_uri='contacts:OfficeLocation', value_cls=string_type),
        SimpleField('profession', field_uri='contacts:Profession', value_cls=string_type),
        SimpleField('surname', field_uri='contacts:Surname', value_cls=string_type),
        # SimpleField('email_alias', field_uri='contacts:Alias', , value_cls=Email),
        # SimpleField('notes', field_uri='contacts:Notes', value_cls=string_type),  # Only available from Exchange 2010 SP2
    )
    ITEM_FIELDS_MAP = {f.name: f for f in ITEM_FIELDS}


class MeetingRequest(Item):
    # Supported attrs: https://msdn.microsoft.com/en-us/library/office/aa565229(v=exchg.150).aspx
    # TODO: Untested and unfinished. Only the bare minimum supported to allow reading a folder that contains meeting
    # requests.
    ELEMENT_NAME = 'MeetingRequest'
    ITEM_FIELDS = (
        SimpleField('subject', field_uri='item:Subject', value_cls=Subject, is_required=True, is_read_only=True),
        SimpleField('author', field_uri='message:From', value_cls=Mailbox, is_read_only=True),
        SimpleField('is_read', field_uri='message:IsRead', value_cls=bool, is_read_only=True),
        SimpleField('start', field_uri='calendar:Start', value_cls=EWSDateTime, is_read_only=True),
        SimpleField('end', field_uri='calendar:End', value_cls=EWSDateTime, is_read_only=True),
    )
    ITEM_FIELDS_MAP = {f.name: f for f in ITEM_FIELDS}

    #__slots__ = ('account', 'folder') + tuple(f.name for f in ITEM_FIELDS)


class MeetingResponse(Item):
    # Supported attrs: https://msdn.microsoft.com/en-us/library/office/aa564337(v=exchg.150).aspx
    # TODO: Untested and unfinished. Only the bare minimum supported to allow reading a folder that contains meeting
    # responses.
    ELEMENT_NAME = 'MeetingResponse'
    ITEM_FIELDS = (
        SimpleField('subject', field_uri='item:Subject', value_cls=Subject, is_required=True, is_read_only=True),
        SimpleField('author', field_uri='message:From', value_cls=Mailbox, is_read_only=True),
        SimpleField('is_read', field_uri='message:IsRead', value_cls=bool, is_read_only=True),
        SimpleField('start', field_uri='calendar:Start', value_cls=EWSDateTime, is_read_only=True),
        SimpleField('end', field_uri='calendar:End', value_cls=EWSDateTime, is_read_only=True),
    )
    ITEM_FIELDS_MAP = {f.name: f for f in ITEM_FIELDS}

    #__slots__ = ('account', 'folder') + tuple(f.name for f in ITEM_FIELDS)


class MeetingCancellation(Item):
    # Supported attrs: https://msdn.microsoft.com/en-us/library/office/aa564685(v=exchg.150).aspx
    # TODO: Untested and unfinished. Only the bare minimum supported to allow reading a folder that contains meeting
    # cancellations.
    ELEMENT_NAME = 'MeetingCancellation'
    ITEM_FIELDS = (
        SimpleField('subject', field_uri='item:Subject', value_cls=Subject, is_required=True, is_read_only=True),
        SimpleField('author', field_uri='message:From', value_cls=Mailbox, is_read_only=True),
        SimpleField('is_read', field_uri='message:IsRead', value_cls=bool, is_read_only=True),
        SimpleField('start', field_uri='calendar:Start', value_cls=EWSDateTime, is_read_only=True),
        SimpleField('end', field_uri='calendar:End', value_cls=EWSDateTime, is_read_only=True),
    )
    ITEM_FIELDS_MAP = {f.name: f for f in ITEM_FIELDS}

    #__slots__ = ('account', 'folder') + tuple(f.name for f in ITEM_FIELDS)


class FileAttachment(Attachment):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa580492(v=exchg.150).aspx
    """
    # TODO: This class is most likely inefficient for large data. Investigate methods to reduce copying
    ELEMENT_NAME = 'FileAttachment'
    ATTACHMENT_FIELDS = {
        'content': ('Content', bytes),
        'is_contact_photo': ('IsContactPhoto', bool),
    }
    ATTACHMENT_FIELDS.update(Attachment.ATTACHMENT_FIELDS)
    ORDERED_FIELDS = (
        'attachment_id', 'name', 'content_type', 'content_id', 'content_location', 'size', 'last_modified_time',
        'is_inline', 'is_contact_photo', 'content',
    )

    __slots__ = ('parent_item',) + tuple(ORDERED_FIELDS[:-1]) + ('_content',)

    def __init__(self, *args, **kwargs):
        self._content = kwargs.pop('content', None)
        self.is_contact_photo = kwargs.pop('is_contact_photo', None)
        super(FileAttachment, self).__init__(*args, **kwargs)
        for field_name, (_, field_type) in self.ATTACHMENT_FIELDS.items():
            if field_name == 'content':
                field_name = '_content'
            val = getattr(self, field_name)
            if val is not None and not isinstance(val, field_type):
                raise ValueError("Field '%s' must be of type '%s'" % (field_name, field_type))

    @property
    def content(self):
        if self.attachment_id is None:
            return self._content
        if self._content is not None:
            return self._content
        # We have an ID to the data but still haven't called GetAttachment to get the actual data. Do that now.
        if not self.parent_item.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        elems = list(GetAttachment(account=self.parent_item.account).call(
            items=[self.attachment_id], include_mime_content=False))
        assert len(elems) == 1
        elem = elems[0]
        if isinstance(elem, Exception):
            raise elem
        assert not isinstance(elem, tuple), elem
        # Don't use get_xml_attr() here because we want to handle empty file content as '', not None
        val = elem.find('{%s}%s' % (TNS, self.ATTACHMENT_FIELDS['content'][0]))
        if val is None:
            self._content = None
        else:
            self._content = base64.b64decode(val.text)
            elem.clear()
        return self._content

    @content.setter
    def content(self, value):
        assert isinstance(value, bytes)
        self._content = value


class ItemAttachment(Attachment):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa562997(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'ItemAttachment'
    ATTACHMENT_FIELDS = {
        'item': (Item, Item),
    }
    ATTACHMENT_FIELDS.update(Attachment.ATTACHMENT_FIELDS)
    ORDERED_FIELDS = (
        'attachment_id', 'name', 'content_type', 'content_id', 'content_location', 'size', 'last_modified_time',
        'is_inline', 'item',
    )

    __slots__ = ('parent_item',) + tuple(ORDERED_FIELDS[:-1]) + ('_item',)

    def __init__(self, *args, **kwargs):
        self._item = kwargs.pop('item', None)
        super(ItemAttachment, self).__init__(*args, **kwargs)
        for field_name, (_, field_type) in self.ATTACHMENT_FIELDS.items():
            if field_name == 'item':
                field_name = '_item'
            val = getattr(self, field_name)
            if val is not None and not isinstance(val, field_type):
                raise ValueError("Field '%s' must be of type '%s'" % (field_name, field_type))

    @property
    def item(self):
        if self.attachment_id is None:
            return self._item
        if self._item is not None:
            return self._item
        # We have an ID to the data but still haven't called GetAttachment to get the actual data. Do that now.
        if not self.parent_item.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        items = list(
            i if isinstance(i, Exception) else self.__class__.from_xml(elem=i)
            for i in GetAttachment(account=self.parent_item.account).call(
                items=[self.attachment_id], include_mime_content=True)
        )
        assert len(items) == 1
        _attachment = items[0]
        if isinstance(_attachment, Exception):
            raise _attachment
        self._item = _attachment._item
        return self._item

    @item.setter
    def item(self, value):
        assert isinstance(value, Item)
        self._item = value


ITEM_CLASSES = (CalendarItem, Contact, Message, Task, MeetingRequest, MeetingResponse, MeetingCancellation)


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
        SimpleField('child_folder_count', field_uri='folder:ChildFolderCount',value_cls=int),
    )
    FOLDER_FIELDS_MAP = {f.name: f for f in FOLDER_FIELDS}

    __slots__ = ('account',) + tuple(f.name for f in FOLDER_FIELDS)

    def __init__(self, account, **kwargs):
        self.account = account

        for f in self.FOLDER_FIELDS:
            setattr(self, f.name, kwargs.pop(f.name, None))
        for k, v in kwargs.items():
            raise TypeError("'%s' is an invalid keyword argument for this function" % k)
        self.clean()
        log.debug('%s created for %s', self, account)

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
            fields.update(item_model.complex_fields())
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
        return QuerySet(self).filter(*args, **kwargs)

    def exclude(self, *args, **kwargs):
        return QuerySet(self).exclude(*args, **kwargs)

    def get(self, *args, **kwargs):
        return QuerySet(self).get(*args, **kwargs)

    def find_items(self, *args, **kwargs):
        """
        Finds items in the folder.

        'shape' controls the exact fields returned are governed by. Be aware that complex elements can only be fetched
        with fetch().

        'depth' controls the whether to return soft-deleted items or not.

        Non-keyword args may be a search expression as supported by Restriction.from_source(), or a list of Q instances.

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

        """
        # 'endswith' and 'iendswith' could be implemented by searching with 'contains' or 'icontains' and then
        # post-processing items. Fetch the field in question with additional_fields and remove items where the search
        # string is not a postfix.

        shape = kwargs.pop('shape', IdOnly)
        depth = kwargs.pop('depth', SHALLOW)
        assert shape in SHAPE_CHOICES
        assert depth in ITEM_TRAVERSAL_CHOICES

        # Define the extra properties we want on the return objects
        additional_fields = kwargs.pop('additional_fields', tuple())
        if additional_fields:
            allowed_fields = self.allowed_fields()
            complex_fields = self.complex_fields()
            for f in additional_fields:
                if f not in allowed_fields:
                    raise ValueError("'%s' is not a field on %s" % (f, self.supported_item_models))
                if f in complex_fields:
                    raise ValueError("find_items() does not support field '%s'. Use fetch() instead" % f)

        # Get the SortOrder field, if any
        order = kwargs.pop('order', None)

        # Get the CalendarView, if any
        calendar_view = kwargs.pop('calendar_view', None)

        # Get the requested number of items per page. Set a sane default and disallow None
        page_size = kwargs.pop('page_size', None) or FindItem.CHUNKSIZE

        # Build up any restrictions
        q = Q.from_filter_args(self.__class__, *args, **kwargs)
        if q and not q.is_empty():
            restriction = Restriction(q.translate_fields(folder_class=self.__class__))
        else:
            restriction = None
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

    def add_items(self, *args, **kwargs):
        warnings.warn('add_items() is deprecated. Use bulk_create() instead', PendingDeprecationWarning)
        return self.bulk_create(*args, **kwargs)

    def bulk_create(self, items, *args, **kwargs):
        return self.account.bulk_create(folder=self, items=items, *args, **kwargs)

    def delete_items(self, ids, *args, **kwargs):
        warnings.warn('delete_items() is deprecated. Use bulk_delete() instead', PendingDeprecationWarning)
        return self.bulk_delete(ids, *args, **kwargs)

    def bulk_delete(self, ids, *args, **kwargs):
        warnings.warn('Folder.bulk_delete() is deprecated. Use Account.bulk_delete() instead', PendingDeprecationWarning)
        return self.account.bulk_delete(ids, *args, **kwargs)

    def update_items(self, items, *args, **kwargs):
        warnings.warn('update_items() is deprecated. Use bulk_update() instead', PendingDeprecationWarning)
        return self.bulk_update(items, *args, **kwargs)

    def bulk_update(self, items, *args, **kwargs):
        warnings.warn('Folder.bulk_update() is deprecated. Use Account.bulk_update() instead', PendingDeprecationWarning)
        return self.account.bulk_update(items, *args, **kwargs)

    def get_items(self, *args, **kwargs):
        warnings.warn('get_items() is deprecated. Use fetch() instead', PendingDeprecationWarning)
        return self.fetch(*args, **kwargs)

    def fetch(self, *args, **kwargs):
        if hasattr(self, 'with_extra_fields'):
            raise DeprecationWarning(
                "'%(cls)s.with_extra_fields' is deprecated. Use 'fetch(ids, only_fields=[...])' instead"
                % dict(cls=self.__class__.__name__))
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
        assert account
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


# Pre-register these extended properties
CalendarItem.register('extern_id', ExternId)
Message.register('extern_id', ExternId)
Contact.register('extern_id', ExternId)
Task.register('extern_id', ExternId)
