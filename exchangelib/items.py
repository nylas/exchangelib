import logging
from decimal import Decimal

from future.utils import python_2_unicode_compatible
from six import string_types

from .attachments import Attachment
from .ewsdatetime import EWSDateTime, UTC_NOW
from .extended_properties import ExtendedProperty
from .fields import SimpleField, PhoneNumberField, EmailAddressField, PhysicalAddressField, ExtendedPropertyField
from .indexed_properties import PhoneNumber, EmailAddress, PhysicalAddress
from .properties import EWSElement, Body, Subject, Location, AnyURI, MimeContent, MessageHeader, ItemId, Choice, \
    Mailbox, Attendee
from .util import create_element
from .version import EXCHANGE_2010

string_type = string_types[0]
log = logging.getLogger(__name__)


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
        # Reminder related fields
        SimpleField('reminder_is_set', field_uri='item:ReminderIsSet', value_cls=bool, is_required=True, default=False),
        SimpleField('reminder_due_by', field_uri='item:ReminderDueBy', value_cls=EWSDateTime, is_required=False,
                    is_required_after_save=True),
        SimpleField('reminder_minutes_before_start', field_uri='item:ReminderMinutesBeforeStart', value_cls=int,
                    is_required_after_save=True, default=0),
        # ExtendedProperty fields go here
        SimpleField('last_modified_name', field_uri='item:LastModifiedName', value_cls=string_type, is_read_only=True),
        SimpleField('last_modified_time', field_uri='item:LastModifiedTime', value_cls=EWSDateTime, is_read_only=True),
    )
    ITEM_FIELDS_MAP = {f.name: f for f in ITEM_FIELDS}

    def __init__(self, **kwargs):
        # 'account' is optional but allows calling 'send()' and 'delete()'
        # 'folder' is optional but allows calling 'save()'
        from .account import Account
        from .folders import Folder
        self.account = kwargs.pop('account', None)
        if self.account is not None:
            assert isinstance(self.account, Account)
        self.folder = kwargs.pop('folder', None)
        if self.folder is not None:
            assert isinstance(self.folder, Folder)

        for f in self.ITEM_FIELDS:
            setattr(self, f.name, kwargs.pop(f.name, None))
        if kwargs:
            raise TypeError("%s are invalid keyword arguments for this function" %
                            ', '.join("'%s'" % k for k in kwargs.keys()))
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

    def save(self, update_fields=None, conflict_resolution=AUTO_RESOLVE, send_meeting_invitations=SEND_TO_NONE):
        item = self._save(update_fieldnames=update_fields, message_disposition=SAVE_ONLY,
                          conflict_resolution=conflict_resolution, send_meeting_invitations=send_meeting_invitations)
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

    def _save(self, update_fieldnames, message_disposition, conflict_resolution, send_meeting_invitations):
        if not self.account:
            raise ValueError('Item must have an account')
        if self.item_id:
            assert self.changekey
            if not update_fieldnames:
                # The fields to update was not specified explicitly. Update all fields where update is possible
                update_fieldnames = []
                for f in self.ITEM_FIELDS:
                    if f.name == 'attachments':
                        # Attachments are handled separately after item creation
                        continue
                    if f.is_read_only:
                        # These cannot be changed
                        continue
                    if f.is_required or f.is_required_after_save:
                        if getattr(self, f.name) is None or (f.is_list and not getattr(self, f.name)):
                            # These are required and cannot be deleted
                            continue
                    if not self.is_draft and f.is_read_only_after_send:
                        # These cannot be changed when the item is no longer a draft
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
            if update_fieldnames:
                raise ValueError("'update_fields' can only be specified when updating an item")
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
            if value is None or (f.is_list and not value):
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
        return '\n'.join('%s: %s' % (f.name, getattr(self, f.name)) for f in self.ITEM_FIELDS)

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

    __slots__ = ('item_id', 'changekey', 'attachments')

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
            res = self._save(update_fieldnames=None, message_disposition=SEND_ONLY,
                             conflict_resolution=conflict_resolution, send_meeting_invitations=send_meeting_invitations)
            assert res is None

    def send_and_save(self, update_fields=None, conflict_resolution=AUTO_RESOLVE,
                      send_meeting_invitations=SEND_TO_NONE):
        # Sends Message and saves a copy in the parent folder. Does not return an ItemId.
        res = self._save(update_fieldnames=update_fields, message_disposition=SEND_AND_SAVE_COPY,
                         conflict_resolution=conflict_resolution, send_meeting_invitations=send_meeting_invitations)
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
        SimpleField('percent_complete', field_uri='task:PercentComplete', value_cls=Decimal, is_required=True,
                    default=Decimal(0.0)),
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
        SimpleField('display_name', field_uri='contacts:DisplayName', value_cls=string_type, is_required=True,
                    default=''),
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


ITEM_CLASSES = (CalendarItem, Contact, Message, Task, MeetingRequest, MeetingResponse, MeetingCancellation)
