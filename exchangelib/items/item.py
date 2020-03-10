import logging

from ..fields import BooleanField, IntegerField, TextField, CharListField, ChoiceField, URIField, BodyField, \
    DateTimeField, MessageHeaderField, AttachmentField, Choice, EWSElementField, EffectiveRightsField, CultureField, \
    CharField, MimeContentField
from ..properties import ConversationId, ParentFolderId, ReferenceItemId, OccurrenceItemId, RecurringMasterItemId
from ..util import is_iterable
from ..version import EXCHANGE_2010, EXCHANGE_2013
from .base import BaseItem, SAVE_ONLY, SEND_ONLY, SEND_AND_SAVE_COPY

log = logging.getLogger(__name__)

# SendMeetingInvitations values. See
# https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/createitem
# SendMeetingInvitationsOrCancellations. See
# https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/updateitem
# SendMeetingCancellations values. See
# https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/deleteitem
SEND_TO_NONE = 'SendToNone'
SEND_ONLY_TO_ALL = 'SendOnlyToAll'
SEND_ONLY_TO_CHANGED = 'SendOnlyToChanged'
SEND_TO_ALL_AND_SAVE_COPY = 'SendToAllAndSaveCopy'
SEND_TO_CHANGED_AND_SAVE_COPY = 'SendToChangedAndSaveCopy'
SEND_MEETING_INVITATIONS_CHOICES = (SEND_TO_NONE, SEND_ONLY_TO_ALL, SEND_TO_ALL_AND_SAVE_COPY)
SEND_MEETING_INVITATIONS_AND_CANCELLATIONS_CHOICES = (SEND_TO_NONE, SEND_ONLY_TO_ALL, SEND_ONLY_TO_CHANGED,
                                                      SEND_TO_ALL_AND_SAVE_COPY, SEND_TO_CHANGED_AND_SAVE_COPY)
SEND_MEETING_CANCELLATIONS_CHOICES = (SEND_TO_NONE, SEND_ONLY_TO_ALL, SEND_TO_ALL_AND_SAVE_COPY)

# AffectedTaskOccurrences values. See
# https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/deleteitem
ALL_OCCURRENCIES = 'AllOccurrences'
SPECIFIED_OCCURRENCE_ONLY = 'SpecifiedOccurrenceOnly'
AFFECTED_TASK_OCCURRENCES_CHOICES = (ALL_OCCURRENCIES, SPECIFIED_OCCURRENCE_ONLY)

# ConflictResolution values. See
# https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/updateitem
NEVER_OVERWRITE = 'NeverOverwrite'
AUTO_RESOLVE = 'AutoResolve'
ALWAYS_OVERWRITE = 'AlwaysOverwrite'
CONFLICT_RESOLUTION_CHOICES = (NEVER_OVERWRITE, AUTO_RESOLVE, ALWAYS_OVERWRITE)

# DeleteType values. See
# https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/deleteitem
HARD_DELETE = 'HardDelete'
SOFT_DELETE = 'SoftDelete'
MOVE_TO_DELETED_ITEMS = 'MoveToDeletedItems'
DELETE_TYPE_CHOICES = (HARD_DELETE, SOFT_DELETE, MOVE_TO_DELETED_ITEMS)


class Item(BaseItem):
    """
    MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/item
    """
    ELEMENT_NAME = 'Item'

    LOCAL_FIELDS = [
        MimeContentField('mime_content', field_uri='item:MimeContent', is_read_only_after_send=True),
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

    FIELDS = LOCAL_FIELDS[0:1] + BaseItem.FIELDS + LOCAL_FIELDS[1:]

    __slots__ = tuple(f.name for f in LOCAL_FIELDS)

    # Used to register extended properties
    INSERT_AFTER_FIELD = 'has_attachments'

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        # pylint: disable=access-member-before-definition
        if self.attachments:
            for a in self.attachments:
                if a.parent_item:
                    if a.parent_item is not self:
                        raise ValueError("'parent_item' of attachment %s must point to this item" % a)
                else:
                    a.parent_item = self
                self.attach(self.attachments)
        else:
            self.attachments = []

    def save(self, update_fields=None, conflict_resolution=AUTO_RESOLVE, send_meeting_invitations=SEND_TO_NONE):
        if self.id:
            item_id, changekey = self._update(
                update_fieldnames=update_fields,
                message_disposition=SAVE_ONLY,
                conflict_resolution=conflict_resolution,
                send_meeting_invitations=send_meeting_invitations
            )
            if self.id != item_id and not isinstance(self._id, (OccurrenceItemId, RecurringMasterItemId)):
                # When we update an item with an OccurrenceItemId as ID, EWS returns the ID of the occurrence, so
                # the ID of this item changes.
                raise ValueError("'id' mismatch in returned update response")
            # Don't check that changekeys are different. No-op saves will sometimes leave the changekey intact
            self._id = self.ID_ELEMENT_CLS(item_id, changekey)
        else:
            if update_fields:
                raise ValueError("'update_fields' is only valid for updates")
            tmp_attachments = None
            if self.account and self.account.version.build < EXCHANGE_2010 and self.attachments:
                # Exchange 2007 can't save attachments immediately. You need to first save, then attach. Store
                # the attachment of this item temporarily and attach later.
                tmp_attachments, self.attachments = self.attachments, []
            item = self._create(message_disposition=SAVE_ONLY, send_meeting_invitations=send_meeting_invitations)
            self._id = self.ID_ELEMENT_CLS(item.id, item.changekey)
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
        from .contact import Contact, DistributionList
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
        if not self.changekey:
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
        if self.id != fresh_item.id and not isinstance(self._id, (OccurrenceItemId, RecurringMasterItemId)):
            # When we refresh an item with an OccurrenceItemId as ID, EWS returns the ID of the occurrence, so
            # the ID of this item changes.
            raise ValueError("'id' mismatch in returned update response")
        for f in self.FIELDS:
            setattr(self, f.name, getattr(fresh_item, f.name))
        # 'parent_item' should point to 'self', not 'fresh_item'. That way, 'fresh_item' can be garbage collected.
        for a in self.attachments:
            a.parent_item = self
        del fresh_item

    def copy(self, to_folder):
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        if not self.id:
            raise ValueError('%s must have an ID' % self.__class__.__name__)
        res = self.account.bulk_copy(ids=[self], to_folder=to_folder)
        if not res:
            # Assume 'to_folder' is a public folder or a folder in a different mailbox
            return
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
            self._id = None
            return
        if len(res) != 1:
            raise ValueError('Expected result length 1, but got %s' % res)
        if isinstance(res[0], Exception):
            raise res[0]
        self._id = self.ID_ELEMENT_CLS(*res[0])
        self.folder = to_folder

    def move_to_trash(self, send_meeting_cancellations=SEND_TO_NONE, affected_task_occurrences=ALL_OCCURRENCIES,
                      suppress_read_receipts=True):
        # Delete and move to the trash folder.
        self._delete(delete_type=MOVE_TO_DELETED_ITEMS, send_meeting_cancellations=send_meeting_cancellations,
                     affected_task_occurrences=affected_task_occurrences, suppress_read_receipts=suppress_read_receipts)
        self._id = None
        self.folder = self.account.trash

    def soft_delete(self, send_meeting_cancellations=SEND_TO_NONE, affected_task_occurrences=ALL_OCCURRENCIES,
                    suppress_read_receipts=True):
        # Delete and move to the dumpster, if it is enabled.
        self._delete(delete_type=SOFT_DELETE, send_meeting_cancellations=send_meeting_cancellations,
                     affected_task_occurrences=affected_task_occurrences, suppress_read_receipts=suppress_read_receipts)
        self._id = None
        self.folder = self.account.recoverable_items_deletions

    def delete(self, send_meeting_cancellations=SEND_TO_NONE, affected_task_occurrences=ALL_OCCURRENCIES,
               suppress_read_receipts=True):
        # Remove the item permanently. No copies are stored anywhere.
        self._delete(delete_type=HARD_DELETE, send_meeting_cancellations=send_meeting_cancellations,
                     affected_task_occurrences=affected_task_occurrences, suppress_read_receipts=suppress_read_receipts)
        self._id, self.folder = None, None

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

    def archive(self, to_folder):
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        if not self.id:
            raise ValueError('%s must have an ID' % self.__class__.__name__)
        res = self.account.bulk_archive(ids=[self], to_folder=to_folder)
        if len(res) != 1:
            raise ValueError('Expected result length 1, but got %s' % res)
        if isinstance(res[0], Exception):
            raise res[0]
        return res[0]

    def attach(self, attachments):
        """Add an attachment, or a list of attachments, to this item. If the item has already been saved, the
        attachments will be created on the server immediately. If the item has not yet been saved, the attachments will
        be created on the server when the item is saved.

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
        if attachments is self.attachments:
            # Don't remove from the same list we are iterating
            attachments = list(attachments)
        for a in attachments:
            if a.parent_item is not self:
                raise ValueError('Attachment does not belong to this item')
            if self.id:
                # Item is already created. Detach  the attachment server-side now
                a.detach()
            if a in self.attachments:
                self.attachments.remove(a)

    def create_forward(self, subject, body, to_recipients, cc_recipients=None, bcc_recipients=None):
        from .message import ForwardItem
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        if not self.id:
            raise ValueError('%s must have an ID' % self.__class__.__name__)
        return ForwardItem(
            account=self.account,
            reference_item_id=ReferenceItemId(id=self.id, changekey=self.changekey),
            subject=subject,
            new_body=body,
            to_recipients=to_recipients,
            cc_recipients=cc_recipients,
            bcc_recipients=bcc_recipients,
        )

    def forward(self, subject, body, to_recipients, cc_recipients=None, bcc_recipients=None):
        self.create_forward(
            subject,
            body,
            to_recipients,
            cc_recipients,
            bcc_recipients,
        ).send()


class BulkCreateResult(BaseItem):
    """A dummy class to store return values from a CreateItem service call"""
    LOCAL_FIELDS = [
        AttachmentField('attachments', field_uri='item:Attachments'),  # ItemAttachment or FileAttachment
    ]
    FIELDS = BaseItem.FIELDS + LOCAL_FIELDS

    __slots__ = tuple(f.name for f in LOCAL_FIELDS)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        # pylint: disable=access-member-before-definition
        if self.attachments is None:
            self.attachments = []
