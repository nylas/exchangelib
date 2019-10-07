from __future__ import unicode_literals

from decimal import Decimal
import logging
import warnings

from future.utils import python_2_unicode_compatible

from .ewsdatetime import UTC_NOW
from .extended_properties import ExtendedProperty
from .fields import BooleanField, IntegerField, DecimalField, Base64Field, TextField, CharListField, ChoiceField, \
    URIField, BodyField, DateTimeField, MessageHeaderField, PhoneNumberField, EmailAddressesField, \
    PhysicalAddressField, ExtendedPropertyField, AttachmentField, RecurrenceField, MailboxField, MailboxListField, \
    AttendeesField, Choice, OccurrenceField, OccurrenceListField, MemberListField, EWSElementField, \
    EffectiveRightsField, TimeZoneField, CultureField, IdField, CharField, TextListField, EnumAsIntField, \
    EmailAddressField, FreeBusyStatusField, ReferenceItemIdField, AssociatedCalendarItemIdField, OccurrenceItemIdField
from .properties import EWSElement, ItemId, ConversationId, ParentFolderId, Attendee, ReferenceItemId, \
    AssociatedCalendarItemId, PersonaId
from .recurrence import FirstOccurrence, LastOccurrence, Occurrence, DeletedOccurrence
from .util import is_iterable
from .version import EXCHANGE_2007_SP1, EXCHANGE_2010, EXCHANGE_2013

log = logging.getLogger(__name__)

# Overall Types Schema: https://msdn.microsoft.com/en-us/library/hh354700(v=exchg.150).aspx

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

# Conference Type values. See https://msdn.microsoft.com/en-US/library/office/aa563529(v=exchg.150).aspx
CONFERENCE_TYPES = ('NetMeeting', 'NetShow', 'Chat')

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

# Traversal enums
SHALLOW = 'Shallow'
SOFT_DELETED = 'SoftDeleted'
ASSOCIATED = 'Associated'
ITEM_TRAVERSAL_CHOICES = (SHALLOW, SOFT_DELETED, ASSOCIATED)

# Shape enums
ID_ONLY = 'IdOnly'
DEFAULT = 'Default'
# AllProperties doesn't actually get all properties in FindItem, just the "first-class" ones. See
#    http://msdn.microsoft.com/en-us/library/office/dn600367(v=exchg.150).aspx
ALL_PROPERTIES = 'AllProperties'
SHAPE_CHOICES = (ID_ONLY, DEFAULT, ALL_PROPERTIES)

# Contacts search (ResolveNames) scope enums
ACTIVE_DIRECTORY = 'ActiveDirectory'
ACTIVE_DIRECTORY_CONTACTS = 'ActiveDirectoryContacts'
CONTACTS = 'Contacts'
CONTACTS_ACTIVE_DIRECTORY = 'ContactsActiveDirectory'
SEARCH_SCOPE_CHOICES = (ACTIVE_DIRECTORY, ACTIVE_DIRECTORY_CONTACTS, CONTACTS, CONTACTS_ACTIVE_DIRECTORY)


class RegisterMixIn(EWSElement):
    INSERT_AFTER_FIELD = None

    @classmethod
    def register(cls, attr_name, attr_cls):
        """
        Register a custom extended property in this item class so they can be accessed just like any other attribute
        """
        if not cls.INSERT_AFTER_FIELD:
            raise ValueError('Class %s is missing INSERT_AFTER_FIELD value' % cls)
        try:
            cls.get_field_by_fieldname(attr_name)
        except ValueError:
            pass
        else:
            raise ValueError("'%s' is already registered" % attr_name)
        if not issubclass(attr_cls, ExtendedProperty):
            raise ValueError("%r must be a subclass of ExtendedProperty" % attr_cls)
        # Check if class attributes are properly defined
        attr_cls.validate_cls()
        # ExtendedProperty is not a real field, but a placeholder in the fields list. See
        #   https://msdn.microsoft.com/en-us/library/office/aa580790(v=exchg.150).aspx
        #
        # Find the correct index for the new extended property, and insert.
        idx = tuple(f.name for f in cls.FIELDS).index(cls.INSERT_AFTER_FIELD) + 1
        field = ExtendedPropertyField(attr_name, value_cls=attr_cls)
        cls.add_field(field, idx=idx)

    @classmethod
    def deregister(cls, attr_name):
        """
        De-register an extended property that has been registered with register()
        """
        try:
            field = cls.get_field_by_fieldname(attr_name)
        except ValueError:
            raise ValueError("'%s' is not registered" % attr_name)
        if not isinstance(field, ExtendedPropertyField):
            raise ValueError("'%s' is not registered as an ExtendedProperty" % attr_name)
        cls.remove_field(field)


class Item(RegisterMixIn):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa580790(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'Item'

    # FIELDS is an ordered list of attributes supported by this item class
    FIELDS = [
        Base64Field('mime_content', field_uri='item:MimeContent', is_read_only_after_send=True),
        IdField('id', field_uri=ItemId.ID_ATTR, is_read_only=True),
        IdField('changekey', field_uri=ItemId.CHANGEKEY_ATTR, is_read_only=True),
        EWSElementField('parent_folder_id', field_uri='item:ParentFolderId', value_cls=ParentFolderId,
                        is_read_only=True),
        CharField('item_class', field_uri='item:ItemClass', is_read_only=True),
        CharField('subject', field_uri='item:Subject'),
        ChoiceField('sensitivity', field_uri='item:Sensitivity', choices={
            Choice('Normal'), Choice('Personal'), Choice('Private'), Choice('Confidential')
        }, is_required=True, default='Normal'),
        TextField('text_body', field_uri='item:TextBody', is_read_only=True, supported_from=EXCHANGE_2013),
        BodyField('body', field_uri='item:Body'),  # Accepts and returns Body or HTMLBody instances
        AttachmentField('attachments', field_uri='item:Attachments'),  # ItemAttachment or FileAttachment
        DateTimeField('datetime_received', field_uri='item:DateTimeReceived', is_read_only=True),
        IntegerField('size', field_uri='item:Size', is_read_only=True),  # Item size in bytes
        CharListField('categories', field_uri='item:Categories'),
        ChoiceField('importance', field_uri='item:Importance', choices={
            Choice('Low'), Choice('Normal'), Choice('High')
        }, is_required=True, default='Normal'),
        TextField('in_reply_to', field_uri='item:InReplyTo'),
        BooleanField('is_submitted', field_uri='item:IsSubmitted', is_read_only=True),
        BooleanField('is_draft', field_uri='item:IsDraft', is_read_only=True),
        BooleanField('is_from_me', field_uri='item:IsFromMe', is_read_only=True),
        BooleanField('is_resend', field_uri='item:IsResend', is_read_only=True),
        BooleanField('is_unmodified', field_uri='item:IsUnmodified', is_read_only=True),
        MessageHeaderField('headers', field_uri='item:InternetMessageHeaders', is_read_only=True),
        DateTimeField('datetime_sent', field_uri='item:DateTimeSent', is_read_only=True),
        DateTimeField('datetime_created', field_uri='item:DateTimeCreated', is_read_only=True),
        # Placeholder for ResponseObjects
        DateTimeField('reminder_due_by', field_uri='item:ReminderDueBy', is_required_after_save=True,
                      is_searchable=False),
        BooleanField('reminder_is_set', field_uri='item:ReminderIsSet', is_required=True, default=False),
        IntegerField('reminder_minutes_before_start', field_uri='item:ReminderMinutesBeforeStart',
                     is_required_after_save=True, min=0, default=0),
        TextField('display_cc', field_uri='item:DisplayCc', is_read_only=True),
        TextField('display_to', field_uri='item:DisplayTo', is_read_only=True),
        BooleanField('has_attachments', field_uri='item:HasAttachments', is_read_only=True),
        # ExtendedProperty fields go here
        CultureField('culture', field_uri='item:Culture', is_required_after_save=True, is_searchable=False),
        EffectiveRightsField('effective_rights', field_uri='item:EffectiveRights', is_read_only=True),
        CharField('last_modified_name', field_uri='item:LastModifiedName', is_read_only=True),
        DateTimeField('last_modified_time', field_uri='item:LastModifiedTime', is_read_only=True),
        BooleanField('is_associated', field_uri='item:IsAssociated', is_read_only=True, supported_from=EXCHANGE_2010),
        URIField('web_client_read_form_query_string', field_uri='item:WebClientReadFormQueryString',
                 is_read_only=True, supported_from=EXCHANGE_2010),
        URIField('web_client_edit_form_query_string', field_uri='item:WebClientEditFormQueryString',
                 is_read_only=True, supported_from=EXCHANGE_2010),
        EWSElementField('conversation_id', field_uri='item:ConversationId', value_cls=ConversationId,
                        is_read_only=True, supported_from=EXCHANGE_2010),
        BodyField('unique_body', field_uri='item:UniqueBody', is_read_only=True, supported_from=EXCHANGE_2010),
    ]

    # Used to register extended properties
    INSERT_AFTER_FIELD = 'has_attachments'

    # We can't use __slots__ because we need to add extended properties dynamically

    def __init__(self, **kwargs):
        # 'account' is optional but allows calling 'send()' and 'delete()'
        # 'folder' is optional but allows calling 'save()'. If 'folder' has an account, and 'account' is not set,
        # we use folder.account.
        from .folders import Folder
        from .account import Account
        self.account = kwargs.pop('account', None)
        if self.account is not None and not isinstance(self.account, Account):
            raise ValueError("'account' %r must be an Account instance" % self.account)
        self.folder = kwargs.pop('folder', None)
        if self.folder is not None:
            if not isinstance(self.folder, Folder):
                raise ValueError("'folder' %r must be a Folder instance" % self.folder)
            if self.folder.account is not None:
                if self.account is not None:
                    # Make sure the account from kwargs matches the folder account
                    if self.account != self.folder.account:
                        raise ValueError("'account' does not match 'folder.account'")
                self.account = self.folder.account
        if 'item_id' in kwargs:
            warnings.warn("The 'item_id' attribute is deprecated. Use 'id' instead.", PendingDeprecationWarning)
            kwargs['id'] = kwargs.pop('item_id')
        super(Item, self).__init__(**kwargs)
        # pylint: disable=access-member-before-definition
        if self.attachments:
            for a in self.attachments:
                if a.parent_item and a.parent_item is not self:
                    raise ValueError("'parent_item' of attachment %s must point to this item" % a)
                else:
                    a.parent_item = self
                self.attach(self.attachments)
        else:
            self.attachments = []

    @property
    def item_id(self):
        warnings.warn("The 'item_id' attribute is deprecated. Use 'id' instead.", PendingDeprecationWarning)
        return self.id

    @item_id.setter
    def item_id(self, value):
        warnings.warn("The 'item_id' attribute is deprecated. Use 'id' instead.", PendingDeprecationWarning)
        self.id = value

    @classmethod
    def get_field_by_fieldname(cls, fieldname):
        if fieldname == 'item_id':
            warnings.warn("The 'item_id' attribute is deprecated. Use 'id' instead.", PendingDeprecationWarning)
            fieldname = 'id'
        return super(Item, cls).get_field_by_fieldname(fieldname)

    def save(self, update_fields=None, conflict_resolution=AUTO_RESOLVE, send_meeting_invitations=SEND_TO_NONE):
        occurrence_item_id = getattr(self, 'occurrence_item_id', None)
        if self.id or occurrence_item_id:
            item_id, changekey = self._update(
                update_fieldnames=update_fields,
                message_disposition=SAVE_ONLY,
                conflict_resolution=conflict_resolution,
                send_meeting_invitations=send_meeting_invitations
            )
            if not occurrence_item_id and self.id != item_id:
                # We can't match ids when using OccurrenceItemId, since that field
                # doesn't include the item id of the individual occurrence. But
                # otherwise, the item id should be the same.
                raise ValueError("'id' mismatch in returned update response")
            # Don't check that changekeys are different. No-op saves will sometimes leave the changekey intact
            self.changekey = changekey
        else:
            if update_fields:
                raise ValueError("'update_fields' is only valid for updates")
            tmp_attachments = None
            if self.account and self.account.version.build < EXCHANGE_2010 and self.attachments:
                # Exchange 2007 can't save attachments immediately. You need to first save, then attach. Store
                # the attachment of this item temporarily and attach later.
                tmp_attachments, self.attachments = self.attachments, []
            item = self._create(message_disposition=SAVE_ONLY, send_meeting_invitations=send_meeting_invitations)
            self.id, self.changekey = item.id, item.changekey
            for old_att, new_att in zip(self.attachments, item.attachments):
                if old_att.attachment_id is not None:
                    raise ValueError("Old 'attachment_id' is not empty")
                if new_att.attachment_id is None:
                    raise ValueError("New 'attachment_id' is empty")
                old_att.attachment_id = new_att.attachment_id
            if tmp_attachments:
                # Exchange 2007 workaround. See above
                self.attach(tmp_attachments)
        return self

    def _create(self, message_disposition, send_meeting_invitations):
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        # bulk_create() returns an Item because we want to return the ID of both the main item *and* attachments
        res = self.account.bulk_create(
            items=[self], folder=self.folder, message_disposition=message_disposition,
            send_meeting_invitations=send_meeting_invitations)
        if message_disposition in (SEND_ONLY, SEND_AND_SAVE_COPY):
            if res:
                raise ValueError('Got a response in non-save mode')
            return None
        if len(res) != 1:
            raise ValueError('Expected result length 1, but got %s' % res)
        if isinstance(res[0], Exception):
            raise res[0]
        return res[0]

    def _update_fieldnames(self):
        # Return the list of fields we are allowed to update
        update_fieldnames = []
        for f in self.supported_fields(version=self.account.version):
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
            if f.name == 'message_id' and f.is_read_only_after_send:
                # 'message_id' doesn't support updating, no matter the draft status
                continue
            if f.name == 'mime_content' and isinstance(self, (Contact, DistributionList)):
                # Contact and DistributionList don't support updating mime_content, no matter the draft status
                continue
            update_fieldnames.append(f.name)
        return update_fieldnames

    def _update(self, update_fieldnames, message_disposition, conflict_resolution, send_meeting_invitations):
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        occurrence_item_id = getattr(self, 'occurrence_item_id', None)
        if not self.changekey and not getattr(occurrence_item_id, 'changekey', None):
            raise ValueError('%s must have changekey' % self.__class__.__name__)
        if not update_fieldnames:
            # The fields to update was not specified explicitly. Update all fields where update is possible
            update_fieldnames = self._update_fieldnames()
        # bulk_update() returns a tuple
        res = self.account.bulk_update(
            items=[(self, update_fieldnames)], message_disposition=message_disposition,
            conflict_resolution=conflict_resolution,
            send_meeting_invitations_or_cancellations=send_meeting_invitations)
        if message_disposition == SEND_AND_SAVE_COPY:
            if res:
                raise ValueError('Got a response in non-save mode')
            return None
        if len(res) != 1:
            raise ValueError('Expected result length 1, but got %s' % res)
        if isinstance(res[0], Exception):
            raise res[0]
        return res[0]

    def refresh(self):
        # Updates the item based on fresh data from EWS
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        if not self.id:
            raise ValueError('%s must have an ID' % self.__class__.__name__)
        res = list(self.account.fetch(ids=[self]))
        if len(res) != 1:
            raise ValueError('Expected result length 1, but got %s' % res)
        if isinstance(res[0], Exception):
            raise res[0]
        fresh_item = res[0]
        if self.id != fresh_item.id:
            raise ValueError('Unexpected ID of fresh item')
        for f in self.FIELDS:
            setattr(self, f.name, getattr(fresh_item, f.name))

    def copy(self, to_folder):
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        if not self.id:
            raise ValueError('%s must have an ID' % self.__class__.__name__)
        res = self.account.bulk_copy(ids=[self], to_folder=to_folder)
        if len(res) != 1:
            raise ValueError('Expected result length 1, but got %s' % res)
        if isinstance(res[0], Exception):
            raise res[0]
        return res[0]

    def move(self, to_folder):
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        if not self.id:
            raise ValueError('%s must have an ID' % self.__class__.__name__)
        res = self.account.bulk_move(ids=[self], to_folder=to_folder)
        if not res:
            # Assume 'to_folder' is a public folder or a folder in a different mailbox
            self.id, self.changekey = None, None
            return
        if len(res) != 1:
            raise ValueError('Expected result length 1, but got %s' % res)
        if isinstance(res[0], Exception):
            raise res[0]
        self.id, self.changekey = res[0]
        self.folder = to_folder

    def move_to_trash(self, send_meeting_cancellations=SEND_TO_NONE, affected_task_occurrences=ALL_OCCURRENCIES,
                      suppress_read_receipts=True):
        # Delete and move to the trash folder.
        self._delete(delete_type=MOVE_TO_DELETED_ITEMS, send_meeting_cancellations=send_meeting_cancellations,
                     affected_task_occurrences=affected_task_occurrences, suppress_read_receipts=suppress_read_receipts)
        self.id, self.changekey = None, None
        self.folder = self.account.trash

    def soft_delete(self, send_meeting_cancellations=SEND_TO_NONE, affected_task_occurrences=ALL_OCCURRENCIES,
                    suppress_read_receipts=True):
        # Delete and move to the dumpster, if it is enabled.
        self._delete(delete_type=SOFT_DELETE, send_meeting_cancellations=send_meeting_cancellations,
                     affected_task_occurrences=affected_task_occurrences, suppress_read_receipts=suppress_read_receipts)
        self.id, self.changekey = None, None
        self.folder = self.account.recoverable_items_deletions

    def delete(self, send_meeting_cancellations=SEND_TO_NONE, affected_task_occurrences=ALL_OCCURRENCIES,
               suppress_read_receipts=True):
        # Remove the item permanently. No copies are stored anywhere.
        self._delete(delete_type=HARD_DELETE, send_meeting_cancellations=send_meeting_cancellations,
                     affected_task_occurrences=affected_task_occurrences, suppress_read_receipts=suppress_read_receipts)
        self.id, self.changekey, self.folder = None, None, None

    def _delete(self, delete_type, send_meeting_cancellations, affected_task_occurrences, suppress_read_receipts):
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        if not self.id:
            raise ValueError('%s must have an ID' % self.__class__.__name__)
        res = self.account.bulk_delete(
            ids=[self], delete_type=delete_type, send_meeting_cancellations=send_meeting_cancellations,
            affected_task_occurrences=affected_task_occurrences, suppress_read_receipts=suppress_read_receipts)
        if len(res) != 1:
            raise ValueError('Expected result length 1, but got %s' % res)
        if isinstance(res[0], Exception):
            raise res[0]

    def attach(self, attachments):
        """Add an attachment, or a list of attachments, to this item. If the item has already been saved, the
        attachments will be created on the server immediately. If the item has not yet been saved, the attachments will
        be created on the server the item is saved.

        Adding attachments to an existing item will update the changekey of the item.
        """
        if not is_iterable(attachments, generators_allowed=True):
            attachments = [attachments]
        for a in attachments:
            if not a.parent_item:
                a.parent_item = self
            if self.id and not a.attachment_id:
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
        if not is_iterable(attachments, generators_allowed=True):
            attachments = [attachments]
        for a in attachments:
            if a.parent_item is not self:
                raise ValueError('Attachment does not belong to this item')
            if self.id:
                # Item is already created. Detach  the attachment server-side now
                a.detach()
            if a in self.attachments:
                self.attachments.remove(a)

    @classmethod
    def id_from_xml(cls, elem):
        id_elem = elem.find(ItemId.response_tag())
        if id_elem is None:
            return None, None
        return id_elem.get(ItemId.ID_ATTR), id_elem.get(ItemId.CHANGEKEY_ATTR)

    @classmethod
    def from_xml(cls, elem, account):
        item_id, changekey = cls.id_from_xml(elem=elem)
        kwargs = {f.name: f.from_xml(elem=elem, account=account) for f in cls.supported_fields()}
        elem.clear()
        return cls(account=account, id=item_id, changekey=changekey, **kwargs)

    def __eq__(self, other):
        if isinstance(other, tuple):
            return hash((self.id, self.changekey)) == hash(other)
        return super(Item, self).__eq__(other)

    def __hash__(self):
        # If we have an ID and changekey, use that as key. Else return a hash of all attributes
        if self.id:
            return hash((self.id, self.changekey))
        return super(Item, self).__hash__()


@python_2_unicode_compatible
class BulkCreateResult(Item):
    """
    A dummy class to store return values from a CreateItem service call
    """
    FIELDS = [
        IdField('id', field_uri=ItemId.ID_ATTR, is_required=True, is_read_only=True),
        IdField('changekey', field_uri=ItemId.CHANGEKEY_ATTR, is_required=True, is_read_only=True),
        AttachmentField('attachments', field_uri='item:Attachments'),  # ItemAttachment or FileAttachment
    ]


# CalendarItemType enums
SINGLE = 'Single'
OCCURRENCE = 'Occurrence'
EXCEPTION = 'Exception'
RECURRING_MASTER = 'RecurringMaster'
CALENDAR_ITEM_CHOICES = (SINGLE, OCCURRENCE, EXCEPTION, RECURRING_MASTER)


@python_2_unicode_compatible
class CalendarItem(Item):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa564765(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'CalendarItem'
    FIELDS = Item.FIELDS + [
        OccurrenceItemIdField('occurrence_item_id', field_uri='calendar:OccurrenceItemId', is_syncback_only=True),
        TextField('uid', field_uri='calendar:UID', is_required_after_save=True, is_searchable=False),
        DateTimeField('start', field_uri='calendar:Start', is_required=True),
        DateTimeField('end', field_uri='calendar:End', is_required=True),
        DateTimeField('original_start', field_uri='calendar:OriginalStart', is_read_only=True),
        BooleanField('is_all_day', field_uri='calendar:IsAllDayEvent', is_required=True, default=False),
        FreeBusyStatusField('legacy_free_busy_status', field_uri='calendar:LegacyFreeBusyStatus', is_required=True,
                            default='Busy'),
        TextField('location', field_uri='calendar:Location'),
        TextField('when', field_uri='calendar:When'),
        BooleanField('is_meeting', field_uri='calendar:IsMeeting', is_read_only=True),
        BooleanField('is_cancelled', field_uri='calendar:IsCancelled', is_read_only=True),
        BooleanField('is_recurring', field_uri='calendar:IsRecurring', is_read_only=True),
        BooleanField('meeting_request_was_sent', field_uri='calendar:MeetingRequestWasSent', is_read_only=True),
        BooleanField('is_response_requested', field_uri='calendar:IsResponseRequested', default=None,
                     is_required_after_save=True, is_searchable=False),
        ChoiceField('type', field_uri='calendar:CalendarItemType', choices={Choice(c) for c in CALENDAR_ITEM_CHOICES},
                    is_read_only=True),
        ChoiceField('my_response_type', field_uri='calendar:MyResponseType', choices={
            Choice(c) for c in Attendee.RESPONSE_TYPES
        }, is_read_only=True),
        MailboxField('organizer', field_uri='calendar:Organizer', is_read_only=True),
        AttendeesField('required_attendees', field_uri='calendar:RequiredAttendees', is_searchable=False),
        AttendeesField('optional_attendees', field_uri='calendar:OptionalAttendees', is_searchable=False),
        AttendeesField('resources', field_uri='calendar:Resources', is_searchable=False),
        IntegerField('conflicting_meeting_count', field_uri='calendar:ConflictingMeetingCount', is_read_only=True),
        IntegerField('adjacent_meeting_count', field_uri='calendar:AdjacentMeetingCount', is_read_only=True),
        # Placeholder for ConflictingMeetings
        # Placeholder for AdjacentMeetings
        CharField('duration', field_uri='calendar:Duration', is_read_only=True),
        DateTimeField('appointment_reply_time', field_uri='calendar:AppointmentReplyTime', is_read_only=True),
        IntegerField('appointment_sequence_number', field_uri='calendar:AppointmentSequenceNumber', is_read_only=True),
        # Placeholder for AppointmentState
        # AppointmentState is an EnumListField-like field, but with bitmask values:
        #    https://msdn.microsoft.com/en-us/library/office/aa564700(v=exchg.150).aspx
        # We could probably subclass EnumListField to implement this field.
        RecurrenceField('recurrence', field_uri='calendar:Recurrence', is_searchable=False),
        OccurrenceField('first_occurrence', field_uri='calendar:FirstOccurrence', value_cls=FirstOccurrence,
                        is_read_only=True),
        OccurrenceField('last_occurrence', field_uri='calendar:LastOccurrence', value_cls=LastOccurrence,
                        is_read_only=True),
        OccurrenceListField('modified_occurrences', field_uri='calendar:ModifiedOccurrences', value_cls=Occurrence,
                            is_read_only=True),
        OccurrenceListField('deleted_occurrences', field_uri='calendar:DeletedOccurrences', value_cls=DeletedOccurrence,
                            is_read_only=True),
        TimeZoneField('_meeting_timezone', field_uri='calendar:MeetingTimeZone', deprecated_from=EXCHANGE_2010,
                      is_searchable=False),
        TimeZoneField('_start_timezone', field_uri='calendar:StartTimeZone', supported_from=EXCHANGE_2010,
                      is_searchable=False),
        TimeZoneField('_end_timezone', field_uri='calendar:EndTimeZone', supported_from=EXCHANGE_2010,
                      is_searchable=False),
        EnumAsIntField('conference_type', field_uri='calendar:ConferenceType', enum=CONFERENCE_TYPES, min=0,
                       default=None, is_required_after_save=True),
        BooleanField('allow_new_time_proposal', field_uri='calendar:AllowNewTimeProposal', default=None,
                     is_required_after_save=True, is_searchable=False),
        BooleanField('is_online_meeting', field_uri='calendar:IsOnlineMeeting', default=None,
                     is_read_only=True),
        URIField('meeting_workspace_url', field_uri='calendar:MeetingWorkspaceUrl'),
        URIField('net_show_url', field_uri='calendar:NetShowUrl'),
    ]

    @classmethod
    def timezone_fields(cls):
        return [f for f in cls.FIELDS if isinstance(f, TimeZoneField)]

    def clean_timezone_fields(self, version):
        # pylint: disable=access-member-before-definition
        # Sets proper values on the timezone fields if they are not already set
        if version.build < EXCHANGE_2010:
            if self._meeting_timezone is None and self.start is not None:
                self._meeting_timezone = self.start.tzinfo
            self._start_timezone = None
            self._end_timezone = None
        else:
            self._meeting_timezone = None
            if self._start_timezone is None and self.start is not None:
                self._start_timezone = self.start.tzinfo
            if self._end_timezone is None and self.end is not None:
                self._end_timezone = self.end.tzinfo

    def clean(self, version=None):
        # pylint: disable=access-member-before-definition
        super(CalendarItem, self).clean(version=version)
        if self.start and self.end and self.end < self.start:
            raise ValueError("'end' must be greater than 'start' (%s -> %s)" % (self.start, self.end))
        if version:
            self.clean_timezone_fields(version=version)

    def accept(self, **kwargs):
        return AcceptItem(
            account=self.account,
            reference_item_id=ReferenceItemId(id=self.id, changekey=self.changekey),
            **kwargs
        ).send()

    def cancel(self, **kwargs):
        return CancelCalendarItem(
            account=self.account,
            reference_item_id=ReferenceItemId(id=self.id, changekey=self.changekey),
            **kwargs
        ).send()

    def decline(self, **kwargs):
        return DeclineItem(
            account=self.account,
            reference_item_id=ReferenceItemId(id=self.id, changekey=self.changekey),
            **kwargs
        ).send()

    def tentatively_accept(self, **kwargs):
        return TentativelyAcceptItem(
            account=self.account,
            reference_item_id=ReferenceItemId(id=self.id, changekey=self.changekey),
            **kwargs
        ).send()

    def _update_fieldnames(self):
        update_fields = super(CalendarItem, self)._update_fieldnames()
        if self.type == OCCURRENCE:
            # Some CalendarItem fields cannot be updated when the item is an occurrence. The values are empty when we
            # receive them so would have been updated because they are set to None.
            update_fields.remove('recurrence')
            update_fields.remove('uid')
        return update_fields


class Message(Item):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa494306(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'Message'
    FIELDS = Item.FIELDS + [
        MailboxField('sender', field_uri='message:Sender', is_read_only=True, is_read_only_after_send=True),
        MailboxListField('to_recipients', field_uri='message:ToRecipients', is_read_only_after_send=True,
                         is_searchable=False),
        MailboxListField('cc_recipients', field_uri='message:CcRecipients', is_read_only_after_send=True,
                         is_searchable=False),
        MailboxListField('bcc_recipients', field_uri='message:BccRecipients', is_read_only_after_send=True,
                         is_searchable=False),
        BooleanField('is_read_receipt_requested', field_uri='message:IsReadReceiptRequested',
                     is_required=True, default=False, is_read_only_after_send=True),
        BooleanField('is_delivery_receipt_requested', field_uri='message:IsDeliveryReceiptRequested',
                     is_required=True, default=False, is_read_only_after_send=True),
        Base64Field('conversation_index', field_uri='message:ConversationIndex', is_read_only=True),
        CharField('conversation_topic', field_uri='message:ConversationTopic', is_read_only=True),
        # Rename 'From' to 'author'. We can't use fieldname 'from' since it's a Python keyword.
        MailboxField('author', field_uri='message:From', is_read_only_after_send=True),
        CharField('message_id', field_uri='message:InternetMessageId', is_read_only_after_send=True),
        BooleanField('is_read', field_uri='message:IsRead', is_required=True, default=False),
        BooleanField('is_response_requested', field_uri='message:IsResponseRequested', default=False, is_required=True),
        TextField('references', field_uri='message:References'),
        MailboxField('reply_to', field_uri='message:ReplyTo', is_read_only_after_send=True, is_searchable=False),
        MailboxField('received_by', field_uri='message:ReceivedBy', is_read_only=True),
        MailboxField('received_representing', field_uri='message:ReceivedRepresenting', is_read_only=True),
        # Placeholder for ReminderMessageData
    ]

    def send(self, save_copy=True, copy_to_folder=None, conflict_resolution=AUTO_RESOLVE,
             send_meeting_invitations=SEND_TO_NONE):
        # Only sends a message. The message can either be an existing draft stored in EWS or a new message that does
        # not yet exist in EWS.
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        if self.id:
            res = self.account.bulk_send(ids=[self], save_copy=save_copy, copy_to_folder=copy_to_folder)
            if len(res) != 1:
                raise ValueError('Expected result length 1, but got %s' % res)
            if isinstance(res[0], Exception):
                raise res[0]
            # The item will be deleted from the original folder
            self.id, self.changekey = None, None
            self.folder = copy_to_folder
            return None

        # New message
        if copy_to_folder:
            if not save_copy:
                raise AttributeError("'save_copy' must be True when 'copy_to_folder' is set")
            # This would better be done via send_and_save() but lets just support it here
            self.folder = copy_to_folder
            return self.send_and_save(conflict_resolution=conflict_resolution,
                                      send_meeting_invitations=send_meeting_invitations)

        if self.account.version.build < EXCHANGE_2010 and self.attachments:
            # Exchange 2007 can't send attachments immediately. You need to first save, then attach, then send.
            # This is done in send_and_save(). send() will delete the item again.
            self.send_and_save(conflict_resolution=conflict_resolution,
                               send_meeting_invitations=send_meeting_invitations)
            return None

        res = self._create(message_disposition=SEND_ONLY, send_meeting_invitations=send_meeting_invitations)
        if res:
            raise ValueError('Unexpected response in send-only mode')
        return None

    def send_and_save(self, update_fields=None, conflict_resolution=AUTO_RESOLVE,
                      send_meeting_invitations=SEND_TO_NONE):
        # Sends Message and saves a copy in the parent folder. Does not return an ItemId.
        if self.id:
            self._update(
                update_fieldnames=update_fields,
                message_disposition=SEND_AND_SAVE_COPY,
                conflict_resolution=conflict_resolution,
                send_meeting_invitations=send_meeting_invitations
            )
        else:
            if self.account.version.build < EXCHANGE_2010 and self.attachments:
                # Exchange 2007 can't send-and-save attachments immediately. You need to first save, then attach, then
                # send. This is done in save().
                self.save(update_fields=update_fields, conflict_resolution=conflict_resolution,
                          send_meeting_invitations=send_meeting_invitations)
                self.send(save_copy=False, conflict_resolution=conflict_resolution,
                          send_meeting_invitations=send_meeting_invitations)
            else:
                res = self._create(
                    message_disposition=SEND_AND_SAVE_COPY,
                    send_meeting_invitations=send_meeting_invitations
                )
                if res:
                    raise ValueError('Unexpected response in send-only mode')

    def reply(self, subject, body, to_recipients=None, cc_recipients=None, bcc_recipients=None):
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        if not self.id:
            raise ValueError('%s must have an ID' % self.__class__.__name__)
        if not to_recipients and not self.author:
            raise ValueError("'to_recipients' must be set when message has no 'author'")
        ReplyToItem(
            account=self.account,
            reference_item_id=ReferenceItemId(id=self.id, changekey=self.changekey),
            subject=subject,
            new_body=body,
            to_recipients=to_recipients or [self.author],
            cc_recipients=cc_recipients,
            bcc_recipients=bcc_recipients,
        ).send()

    def reply_all(self, subject, body):
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        if not self.id:
            raise ValueError('%s must have an ID' % self.__class__.__name__)
        ReplyAllToItem(
            account=self.account,
            reference_item_id=ReferenceItemId(id=self.id, changekey=self.changekey),
            subject=subject,
            new_body=body,
            to_recipients=self.to_recipients or [self.author],
            cc_recipients=self.cc_recipients,
            bcc_recipients=self.bcc_recipients,
        ).send()

    def forward(self, subject, body, to_recipients, cc_recipients=None, bcc_recipients=None):
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        if not self.id:
            raise ValueError('%s must have an ID' % self.__class__.__name__)
        ForwardItem(
            account=self.account,
            reference_item_id=ReferenceItemId(id=self.id, changekey=self.changekey),
            subject=subject,
            new_body=body,
            to_recipients=to_recipients,
            cc_recipients=cc_recipients,
            bcc_recipients=bcc_recipients,
        ).send()


class Task(Item):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa563930(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'Task'
    NOT_STARTED = 'NotStarted'
    COMPLETED = 'Completed'
    FIELDS = Item.FIELDS + [
        IntegerField('actual_work', field_uri='task:ActualWork', min=0),
        DateTimeField('assigned_time', field_uri='task:AssignedTime', is_read_only=True),
        TextField('billing_information', field_uri='task:BillingInformation'),
        IntegerField('change_count', field_uri='task:ChangeCount', is_read_only=True, min=0),
        TextListField('companies', field_uri='task:Companies'),
        # 'complete_date' can be set, but is ignored by the server, which sets it to now()
        DateTimeField('complete_date', field_uri='task:CompleteDate', is_read_only=True),
        TextListField('contacts', field_uri='task:Contacts'),
        ChoiceField('delegation_state', field_uri='task:DelegationState', choices={
            Choice('NoMatch'), Choice('OwnNew'), Choice('Owned'), Choice('Accepted'), Choice('Declined'), Choice('Max')
        }, is_read_only=True),
        CharField('delegator', field_uri='task:Delegator', is_read_only=True),
        DateTimeField('due_date', field_uri='task:DueDate'),
        BooleanField('is_editable', field_uri='task:IsAssignmentEditable', is_read_only=True),
        BooleanField('is_complete', field_uri='task:IsComplete', is_read_only=True),
        BooleanField('is_recurring', field_uri='task:IsRecurring', is_read_only=True),
        BooleanField('is_team_task', field_uri='task:IsTeamTask', is_read_only=True),
        TextField('mileage', field_uri='task:Mileage'),
        CharField('owner', field_uri='task:Owner', is_read_only=True),
        DecimalField('percent_complete', field_uri='task:PercentComplete', is_required=True, default=Decimal(0.0),
                     is_searchable=False),
        # Placeholder for Recurrence
        DateTimeField('start_date', field_uri='task:StartDate'),
        ChoiceField('status', field_uri='task:Status', choices={
            Choice(NOT_STARTED), Choice('InProgress'), Choice(COMPLETED), Choice('WaitingOnOthers'), Choice('Deferred')
        }, is_required=True, is_searchable=False, default=NOT_STARTED),
        CharField('status_description', field_uri='task:StatusDescription', is_read_only=True),
        IntegerField('total_work', field_uri='task:TotalWork', min=0),
    ]

    def clean(self, version=None):
        # pylint: disable=access-member-before-definition
        super(Task, self).clean(version=version)
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
                # Reset complete_date values that are in the future
                # 'complete_date' can be set automatically by the server. Allow some grace between local and server time
                log.warning("'complete_date' must be in the past (%s vs %s). Resetting", self.complete_date, now)
                self.complete_date = now
            if self.start_date and self.complete_date < self.start_date:
                log.warning("'complete_date' must be greater than 'start_date' (%s vs %s). Resetting",
                            self.complete_date, self.start_date)
                self.complete_date = self.start_date
        if self.percent_complete is not None:
            if not Decimal(0) <= self.percent_complete <= Decimal(100):
                raise ValueError("'percent_complete' must be in range 0.0 - 100.0")
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

    def complete(self):
        # pylint: disable=access-member-before-definition
        # A helper method to mark a task as complete on the server
        self.status = Task.COMPLETED
        self.percent_complete = Decimal(100)
        self.save()


class Contact(Item):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa581315(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'Contact'
    FIELDS = Item.FIELDS + [
        TextField('file_as', field_uri='contacts:FileAs'),
        ChoiceField('file_as_mapping', field_uri='contacts:FileAsMapping', choices={
            Choice('None'), Choice('LastCommaFirst'), Choice('FirstSpaceLast'), Choice('Company'),
            Choice('LastCommaFirstCompany'), Choice('CompanyLastFirst'), Choice('LastFirst'),
            Choice('LastFirstCompany'), Choice('CompanyLastCommaFirst'), Choice('LastFirstSuffix'),
            Choice('LastSpaceFirstCompany'), Choice('CompanyLastSpaceFirst'), Choice('LastSpaceFirst'),
            Choice('DisplayName'), Choice('FirstName'), Choice('LastFirstMiddleSuffix'), Choice('LastName'),
            Choice('Empty'),
        }),
        TextField('display_name', field_uri='contacts:DisplayName', is_required=True),
        CharField('given_name', field_uri='contacts:GivenName'),
        TextField('initials', field_uri='contacts:Initials'),
        CharField('middle_name', field_uri='contacts:MiddleName'),
        TextField('nickname', field_uri='contacts:Nickname'),
        # Placeholder for CompleteName
        TextField('company_name', field_uri='contacts:CompanyName'),
        EmailAddressesField('email_addresses', field_uri='contacts:EmailAddress'),
        PhysicalAddressField('physical_addresses', field_uri='contacts:PhysicalAddress'),
        PhoneNumberField('phone_numbers', field_uri='contacts:PhoneNumber'),
        TextField('assistant_name', field_uri='contacts:AssistantName'),
        DateTimeField('birthday', field_uri='contacts:Birthday'),
        URIField('business_homepage', field_uri='contacts:BusinessHomePage'),
        TextListField('children', field_uri='contacts:Children'),
        TextListField('companies', field_uri='contacts:Companies', is_searchable=False),
        ChoiceField('contact_source', field_uri='contacts:ContactSource', choices={
            Choice('Store'), Choice('ActiveDirectory')
        }, is_read_only=True),
        TextField('department', field_uri='contacts:Department'),
        TextField('generation', field_uri='contacts:Generation'),
        CharField('im_addresses', field_uri='contacts:ImAddresses', is_read_only=True),
        TextField('job_title', field_uri='contacts:JobTitle'),
        TextField('manager', field_uri='contacts:Manager'),
        TextField('mileage', field_uri='contacts:Mileage'),
        TextField('office', field_uri='contacts:OfficeLocation'),
        ChoiceField('postal_address_index', field_uri='contacts:PostalAddressIndex', choices={
            Choice('Business'), Choice('Home'), Choice('Other'), Choice('None')
        }, default='None', is_required_after_save=True),
        TextField('profession', field_uri='contacts:Profession'),
        TextField('spouse_name', field_uri='contacts:SpouseName'),
        CharField('surname', field_uri='contacts:Surname'),
        DateTimeField('wedding_anniversary', field_uri='contacts:WeddingAnniversary'),
        BooleanField('has_picture', field_uri='contacts:HasPicture', supported_from=EXCHANGE_2010, is_read_only=True),
        TextField('phonetic_full_name', field_uri='contacts:PhoneticFullName', supported_from=EXCHANGE_2013,
                  is_read_only=True),
        TextField('phonetic_first_name', field_uri='contacts:PhoneticFirstName', supported_from=EXCHANGE_2013,
                  is_read_only=True),
        TextField('phonetic_last_name', field_uri='contacts:PhoneticLastName', supported_from=EXCHANGE_2013,
                  is_read_only=True),
        EmailAddressField('email_alias', field_uri='contacts:Alias', is_read_only=True),
        CharField('notes', field_uri='contacts:Notes', supported_from=EXCHANGE_2013, is_read_only=True),
        # Placeholder for Photo
        # Placeholder for UserSMIMECertificate
        # Placeholder for MSExchangeCertificate
        TextField('directory_id', field_uri='contacts:DirectoryId', supported_from=EXCHANGE_2013, is_read_only=True),
        # Placeholder for ManagerMailbox
        # Placeholder for DirectReports
    ]


class DistributionList(Item):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa566353(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'DistributionList'
    FIELDS = Item.FIELDS + [
        CharField('display_name', field_uri='contacts:DisplayName', is_required=True),
        CharField('file_as', field_uri='contacts:FileAs', is_read_only=True),
        ChoiceField('contact_source', field_uri='contacts:ContactSource', choices={
            Choice('Store'), Choice('ActiveDirectory')
        }, is_read_only=True),
        MemberListField('members', field_uri='distributionlist:Members'),
    ]


class PostItem(Item):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/bb891851(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'PostItem'
    FIELDS = Item.FIELDS + [
        Base64Field('conversation_index', field_uri='message:ConversationIndex', is_read_only=True),
        CharField('conversation_topic', field_uri='message:ConversationTopic', is_read_only=True),
        MailboxField('author', field_uri='message:From', is_read_only_after_send=True),
        CharField('message_id', field_uri='message:InternetMessageId', is_read_only_after_send=True),
        BooleanField('is_read', field_uri='message:IsRead', is_required=True, default=False),
        DateTimeField('posted_time', field_uri='postitem:PostedTime', is_read_only=True),
        MailboxField('sender', field_uri='message:Sender', is_read_only=True, is_read_only_after_send=True),
    ]


class PostReplyItem(Item):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/bb891896(v=exchg.150).aspx
    """
    # TODO: Untested and unfinished.
    ELEMENT_NAME = 'PostReplyItem'


class BaseMeetingItem(Item):
    """
    A base class for meeting requests that share the same fields (Message, Request, Response, Cancellation)

    MSDN: https://msdn.microsoft.com/en-us/library/aa580757(v=exchg.150).aspx
        Certain types are created as a side effect of doing something else. Meeting messages, for example, are created
        when you send a calendar item to attendees; they are not explicitly created.

    Therefore BaseMeetingItem inherits from  EWSElement has no save() or send() method

    """
    FIELDS = Item.FIELDS + [
        MailboxField('sender', field_uri='message:Sender', is_read_only=True, is_read_only_after_send=True),
        MailboxListField('to_recipients', field_uri='message:ToRecipients', is_read_only_after_send=True,
                         is_searchable=False),
        MailboxListField('cc_recipients', field_uri='message:CcRecipients', is_read_only_after_send=True,
                         is_searchable=False),
        MailboxListField('bcc_recipients', field_uri='message:BccRecipients', is_read_only_after_send=True,
                         is_searchable=False),
        BooleanField('is_read_receipt_requested', field_uri='message:IsReadReceiptRequested',
                     is_required=True, default=False, is_read_only_after_send=True),
        BooleanField('is_delivery_receipt_requested', field_uri='message:IsDeliveryReceiptRequested',
                     is_required=True, default=False, is_read_only_after_send=True),
        Base64Field('conversation_index', field_uri='message:ConversationIndex', is_read_only=True),
        CharField('conversation_topic', field_uri='message:ConversationTopic', is_read_only=True),
        # Rename 'From' to 'author'. We can't use fieldname 'from' since it's a Python keyword.
        MailboxField('author', field_uri='message:From', is_read_only_after_send=True),
        CharField('message_id', field_uri='message:InternetMessageId', is_read_only_after_send=True),
        BooleanField('is_read', field_uri='message:IsRead', is_required=True, default=False),
        BooleanField('is_response_requested', field_uri='message:IsResponseRequested', default=False, is_required=True),
        TextField('references', field_uri='message:References'),
        MailboxField('reply_to', field_uri='message:ReplyTo', is_read_only_after_send=True, is_searchable=False),
        AssociatedCalendarItemIdField('associated_calendar_item_id', field_uri='meeting:AssociatedCalendarItemId',
                                      value_cls=AssociatedCalendarItemId),
        BooleanField('is_delegated', field_uri='meeting:IsDelegated', is_read_only=True, default=False),
        BooleanField('is_out_of_date', field_uri='meeting:IsOutOfDate', is_read_only=True, default=False),
        BooleanField('has_been_processed', field_uri='meeting:HasBeenProcessed', is_read_only=True, default=False),
        ChoiceField('response_type', field_uri='meeting:ResponseType',
                    choices={Choice('Unknown'), Choice('Organizer'), Choice('Tentative'),
                             Choice('Accept'), Choice('Decline'), Choice('NoResponseReceived')},
                    is_required=True, default='Unknown'),
    ]

    # Used to register extended properties
    INSERT_AFTER_FIELD = 'has_attachments'


class MeetingRequest(BaseMeetingItem):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa565229(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'MeetingRequest'
    FIELDS = Message.FIELDS + [
        ChoiceField('meeting_request_type', field_uri='meetingRequest:MeetingRequestType',
                    choices={Choice('FullUpdate'), Choice('InformationalUpdate'), Choice('NewMeetingRequest'),
                             Choice('None'), Choice('Outdated'), Choice('PrincipalWantsCopy'),
                             Choice('SilentUpdate')},
                    default='None'),
        ChoiceField('intended_free_busy_status', field_uri='meetingRequest:IntendedFreeBusyStatus', choices={
                    Choice('Free'), Choice('Tentative'), Choice('Busy'), Choice('OOF'), Choice('NoData')},
                    is_required=True, default='Busy'),
        DateTimeField('start', field_uri='calendar:Start', is_read_only=True, supported_from=EXCHANGE_2010),
        DateTimeField('end', field_uri='calendar:End', is_read_only=True, supported_from=EXCHANGE_2010),
        DateTimeField('original_start', field_uri='calendar:OriginalStart', is_read_only=True),
        BooleanField('is_all_day', field_uri='calendar:IsAllDayEvent', is_required=True, default=False),
        ChoiceField('legacy_free_busy_status', field_uri='calendar:LegacyFreeBusyStatus', choices={
            Choice('Free'), Choice('Tentative'), Choice('Busy'), Choice('OOF'), Choice('NoData'),
            Choice('WorkingElsewhere', supported_from=EXCHANGE_2013)
        }, is_required=True, default='Busy'),
        TextField('location', field_uri='calendar:Location'),
        TextField('when', field_uri='calendar:When'),
        BooleanField('is_meeting', field_uri='calendar:IsMeeting', is_read_only=True),
        BooleanField('is_cancelled', field_uri='calendar:IsCancelled', is_read_only=True),
        BooleanField('is_recurring', field_uri='calendar:IsRecurring', is_read_only=True),
        BooleanField('meeting_request_was_sent', field_uri='calendar:MeetingRequestWasSent', is_read_only=True),
        ChoiceField('type', field_uri='calendar:CalendarItemType', choices={
            Choice('Single'), Choice('Occurrence'), Choice('Exception'), Choice('RecurringMaster'),
        }, is_read_only=True),
        ChoiceField('my_response_type', field_uri='calendar:MyResponseType', choices={
            Choice(c) for c in Attendee.RESPONSE_TYPES
        }, is_read_only=True),
        MailboxField('organizer', field_uri='calendar:Organizer', is_read_only=True),
        AttendeesField('required_attendees', field_uri='calendar:RequiredAttendees', is_searchable=False),
        AttendeesField('optional_attendees', field_uri='calendar:OptionalAttendees', is_searchable=False),
        AttendeesField('resources', field_uri='calendar:Resources', is_searchable=False),
        IntegerField('conflicting_meeting_count', field_uri='calendar:ConflictingMeetingCount', is_read_only=True),
        IntegerField('adjacent_meeting_count', field_uri='calendar:AdjacentMeetingCount', is_read_only=True),
        # Placeholder for calendar:ConflictingMeetings
        # Placeholder for calendar:AdjacentMeetings
        CharField('duration', field_uri='calendar:Duration', is_read_only=True),
        # Placeholder for calendar:TimeZone
        DateTimeField('appointment_reply_time', field_uri='calendar:AppointmentReplyTime', is_read_only=True),
        IntegerField('appointment_sequence_number', field_uri='calendar:AppointmentSequenceNumber', is_read_only=True),
        # Placeholder for calendar:AppointmentState
        # AppointmentState is an EnumListField-like field, but with bitmask values:
        #    https://msdn.microsoft.com/en-us/library/office/aa564700(v=exchg.150).aspx
        # We could probably subclass EnumListField to implement this field.
        RecurrenceField('recurrence', field_uri='calendar:Recurrence', is_searchable=False),
        OccurrenceField('first_occurrence', field_uri='calendar:FirstOccurrence', value_cls=FirstOccurrence,
                        is_read_only=True),
        OccurrenceField('last_occurrence', field_uri='calendar:LastOccurrence', value_cls=LastOccurrence,
                        is_read_only=True),
        OccurrenceListField('modified_occurrences', field_uri='calendar:ModifiedOccurrences', value_cls=Occurrence,
                            is_read_only=True),
        OccurrenceListField('deleted_occurrences', field_uri='calendar:DeletedOccurrences', value_cls=DeletedOccurrence,
                            is_read_only=True),
        TimeZoneField('_meeting_timezone', field_uri='calendar:MeetingTimeZone', deprecated_from=EXCHANGE_2010,
                      is_read_only=True, is_searchable=False),
        TimeZoneField('_start_timezone', field_uri='calendar:StartTimeZone', supported_from=EXCHANGE_2010,
                      is_read_only=True, is_searchable=False),
        TimeZoneField('_end_timezone', field_uri='calendar:EndTimeZone', supported_from=EXCHANGE_2010,
                      is_read_only=True, is_searchable=False),
        EnumAsIntField('conference_type', field_uri='calendar:ConferenceType', enum=CONFERENCE_TYPES, min=0,
                       default=None, is_required_after_save=True),
        BooleanField('allow_new_time_proposal', field_uri='calendar:AllowNewTimeProposal', default=None,
                     is_required_after_save=True, is_searchable=False),
        BooleanField('is_online_meeting', field_uri='calendar:IsOnlineMeeting', default=None,
                     is_required_after_save=True),
        URIField('meeting_workspace_url', field_uri='calendar:MeetingWorkspaceUrl'),
        URIField('net_show_url', field_uri='calendar:NetShowUrl'),
    ]

    def accept(self, **kwargs):
        return AcceptItem(
            account=self.account,
            reference_item_id=ReferenceItemId(id=self.id, changekey=self.changekey),
            **kwargs
        ).send()

    def decline(self, **kwargs):
        return DeclineItem(
            account=self.account,
            reference_item_id=ReferenceItemId(id=self.id, changekey=self.changekey),
            **kwargs
        ).send()

    def tentatively_accept(self, **kwargs):
        return TentativelyAcceptItem(
            account=self.account,
            reference_item_id=ReferenceItemId(id=self.id, changekey=self.changekey),
            **kwargs
        ).send()


class MeetingMessage(BaseMeetingItem):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa565359(v=exchg.150).aspx
    """
    # TODO: Untested - not sure if this is ever used
    ELEMENT_NAME = 'MeetingMessage'


class MeetingResponse(BaseMeetingItem):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa564337(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'MeetingResponse'
    FIELDS = Message.FIELDS + [
        MailboxField('received_by', field_uri='message:ReceivedBy', is_read_only=True),
        MailboxField('received_representing', field_uri='message:ReceivedRepresenting', is_read_only=True),
    ]


class MeetingCancellation(BaseMeetingItem):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa564685(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'MeetingCancellation'


class BaseMeetingReplyItem(Item):
    # A base class for meeting request reply items that share the same fields (Accept, TentativelyAccept, Decline)
    FIELDS = [
        CharField('item_class', field_uri='item:ItemClass', is_read_only=True),
        ChoiceField('sensitivity', field_uri='item:Sensitivity', choices={
            Choice('Normal'), Choice('Personal'), Choice('Private'), Choice('Confidential')
        }, is_required=True, default='Normal'),
        BodyField('body', field_uri='item:Body'),  # Accepts and returns Body or HTMLBody instances
        AttachmentField('attachments', field_uri='item:Attachments'),  # ItemAttachment or FileAttachment
        MessageHeaderField('headers', field_uri='item:InternetMessageHeaders', is_read_only=True),
        MailboxField('sender', field_uri='message:Sender', is_read_only=True, is_read_only_after_send=True),
        MailboxListField('to_recipients', field_uri='message:ToRecipients', is_read_only_after_send=True,
                         is_searchable=False),
        MailboxListField('cc_recipients', field_uri='message:CcRecipients', is_read_only_after_send=True,
                         is_searchable=False),
        MailboxListField('bcc_recipients', field_uri='message:BccRecipients', is_read_only_after_send=True,
                         is_searchable=False),
        BooleanField('is_read_receipt_requested', field_uri='message:IsReadReceiptRequested',
                     is_required=True, default=False, is_read_only_after_send=True),
        BooleanField('is_delivery_receipt_requested', field_uri='message:IsDeliveryReceiptRequested',
                     is_required=True, default=False, is_read_only_after_send=True),
        ReferenceItemIdField('reference_item_id', field_uri='item:ReferenceItemId', value_cls=ReferenceItemId),
        MailboxField('received_by', field_uri='message:ReceivedBy', is_read_only=True),
        MailboxField('received_representing', field_uri='message:ReceivedRepresenting', is_read_only=True),
        DateTimeField('proposed_start', field_uri='meeting:ProposedStart', supported_from=EXCHANGE_2013),
        DateTimeField('proposed_end', field_uri='meeting:ProposedEnd', supported_from=EXCHANGE_2013),
    ]

    def send(self, message_disposition=SEND_AND_SAVE_COPY):
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)

        res = self.account.bulk_create(items=[self], folder=self.folder, message_disposition=message_disposition)

        for r_item in res:
            if isinstance(r_item, Exception):
                raise r_item
        return res


class AcceptItem(BaseMeetingReplyItem):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/aa562964(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'AcceptItem'


class TentativelyAcceptItem(BaseMeetingReplyItem):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/aa565438(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'TentativelyAcceptItem'


class DeclineItem(BaseMeetingReplyItem):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/aa579729(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'DeclineItem'


class BaseReplyItem(EWSElement):
    # A base class for reply/forward elements that share the same fields
    FIELDS = [
        CharField('subject', field_uri='Subject'),
        BodyField('body', field_uri='Body'),  # Accepts and returns Body or HTMLBody instances
        MailboxListField('to_recipients', field_uri='ToRecipients'),
        MailboxListField('cc_recipients', field_uri='CcRecipients'),
        MailboxListField('bcc_recipients', field_uri='BccRecipients'),
        BooleanField('is_read_receipt_requested', field_uri='IsReadReceiptRequested'),
        BooleanField('is_delivery_receipt_requested', field_uri='IsDeliveryReceiptRequested'),
        MailboxField('author', field_uri='From'),
        EWSElementField('reference_item_id', value_cls=ReferenceItemId),
        BodyField('new_body', field_uri='NewBodyContent'),  # Accepts and returns Body or HTMLBody instances
        MailboxField('received_by', field_uri='ReceivedBy', supported_from=EXCHANGE_2007_SP1),
        MailboxField('received_by_representing', field_uri='ReceivedRepresenting', supported_from=EXCHANGE_2007_SP1),
    ]

    def __init__(self, **kwargs):
        # 'account' is optional but allows calling 'send()'
        from .account import Account
        self.account = kwargs.pop('account', None)
        if self.account is not None and not isinstance(self.account, Account):
            raise ValueError("'account' %r must be an Account instance" % self.account)
        super(BaseReplyItem, self).__init__(**kwargs)

    def send(self, save_copy=True, copy_to_folder=None):
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        if copy_to_folder:
            if not save_copy:
                raise AttributeError("'save_copy' must be True when 'copy_to_folder' is set")
        message_disposition = SEND_AND_SAVE_COPY if save_copy else SEND_ONLY
        res = self.account.bulk_create(items=[self], folder=copy_to_folder, message_disposition=message_disposition)
        if res and isinstance(res[0], Exception):
            raise res[0]


class ReplyToItem(BaseReplyItem):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa580287(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'ReplyToItem'


class ReplyAllToItem(BaseReplyItem):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa563988(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'ReplyAllToItem'


class ForwardItem(BaseReplyItem):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa564250(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'ForwardItem'


class CancelCalendarItem(BaseReplyItem):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/aa564482(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'CancelCalendarItem'


class Persona(EWSElement):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/jj191299(v=exchg.150).aspx
    ELEMENT_NAME = 'Persona'
    FIELDS = [
        EWSElementField('persona_id', field_uri='persona:PersonaId', value_cls=PersonaId),
        CharField('file_as', field_uri='persona:FileAs'),
        CharField('display_name', field_uri='persona:DisplayName'),
        CharField('given_name', field_uri='persona:GivenName'),
        TextField('middle_name', field_uri='persona:MiddleName'),
        CharField('surname', field_uri='persona:Surname'),
        TextField('generation', field_uri='persona:Generation'),
        TextField('nickname', field_uri='persona:Nickname'),
        CharField('title', field_uri='persona:Title'),
        TextField('department', field_uri='persona:Department'),
        CharField('company_name', field_uri='persona:CompanyName'),
        CharField('im_address', field_uri='persona:ImAddress'),
        TextField('initials', field_uri='persona:Initials'),
    ]

    @classmethod
    def id_from_xml(cls, elem):
        id_elem = elem.find(PersonaId.response_tag())
        if id_elem is None:
            return None, None
        return id_elem.get(PersonaId.ID_ATTR), id_elem.get(PersonaId.CHANGEKEY_ATTR)

    def __eq__(self, other):
        if isinstance(other, tuple):
            return hash(self.persona_id) == hash(other)
        return super(Persona, self).__eq__(other)

    def __hash__(self):
        # If we have a persona_id, use that as key. Else return a hash of all attributes
        if self.persona_id:
            return hash(self.persona_id)
        return super(Persona, self).__hash__()


ITEM_CLASSES = (CalendarItem, Contact, DistributionList, Message, PostItem, Task, MeetingRequest, MeetingResponse,
                MeetingCancellation)
