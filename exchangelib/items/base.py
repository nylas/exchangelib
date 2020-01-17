import logging

from ..extended_properties import ExtendedProperty
from ..fields import BooleanField, ExtendedPropertyField, BodyField, MailboxField, MailboxListField, EWSElementField, \
    CharField
from ..properties import InvalidField, IdChangeKeyMixIn, EWSElement, ReferenceItemId
from ..version import EXCHANGE_2007_SP1

log = logging.getLogger(__name__)

# MessageDisposition values. See
# https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/createitem
SAVE_ONLY = 'SaveOnly'
SEND_ONLY = 'SendOnly'
SEND_AND_SAVE_COPY = 'SendAndSaveCopy'
MESSAGE_DISPOSITION_CHOICES = (SAVE_ONLY, SEND_ONLY, SEND_AND_SAVE_COPY)


class RegisterMixIn(IdChangeKeyMixIn):
    """Base class for classes that can change their list of supported fields dynamically"""

    # This class implements dynamic fields on an element class, so we need to include __dict__ in __slots__
    __slots__ = ('__dict__',)

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
        except InvalidField:
            pass
        else:
            raise ValueError("'%s' is already registered" % attr_name)
        if not issubclass(attr_cls, ExtendedProperty):
            raise ValueError("%r must be a subclass of ExtendedProperty" % attr_cls)
        # Check if class attributes are properly defined
        attr_cls.validate_cls()
        # ExtendedProperty is not a real field, but a placeholder in the fields list. See
        #   https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/item
        #
        # Find the correct index for the new extended property, and insert.
        field = ExtendedPropertyField(attr_name, value_cls=attr_cls)
        cls.add_field(field, insert_after=cls.INSERT_AFTER_FIELD)

    @classmethod
    def deregister(cls, attr_name):
        """
        De-register an extended property that has been registered with register()
        """
        try:
            field = cls.get_field_by_fieldname(attr_name)
        except InvalidField:
            raise ValueError("'%s' is not registered" % attr_name)
        if not isinstance(field, ExtendedPropertyField):
            raise ValueError("'%s' is not registered as an ExtendedProperty" % attr_name)
        cls.remove_field(field)


class BaseItem(RegisterMixIn):
    """Base class for all other classes that implement EWS items"""
    __slots__ = ('account', 'folder')

    def __init__(self, **kwargs):
        # 'account' is optional but allows calling 'send()' and 'delete()'
        # 'folder' is optional but allows calling 'save()'. If 'folder' has an account, and 'account' is not set,
        # we use folder.account.
        from ..folders import BaseFolder
        from ..account import Account
        self.account = kwargs.pop('account', None)
        if self.account is not None and not isinstance(self.account, Account):
            raise ValueError("'account' %r must be an Account instance" % self.account)
        self.folder = kwargs.pop('folder', None)
        if self.folder is not None:
            if not isinstance(self.folder, BaseFolder):
                raise ValueError("'folder' %r must be a Folder instance" % self.folder)
            if self.folder.account is not None:
                if self.account is not None:
                    # Make sure the account from kwargs matches the folder account
                    if self.account != self.folder.account:
                        raise ValueError("'account' does not match 'folder.account'")
                self.account = self.folder.account
        super().__init__(**kwargs)

    @classmethod
    def from_xml(cls, elem, account):
        item = super().from_xml(elem=elem, account=account)
        item.account = account
        return item


class BaseReplyItem(EWSElement):
    """Base class for reply/forward elements that share the same fields"""
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

    __slots__ = tuple(f.name for f in FIELDS) + ('account',)

    def __init__(self, **kwargs):
        # 'account' is optional but allows calling 'send()' and 'save()'
        from ..account import Account
        self.account = kwargs.pop('account', None)
        if self.account is not None and not isinstance(self.account, Account):
            raise ValueError("'account' %r must be an Account instance" % self.account)
        super().__init__(**kwargs)

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

    def save(self, folder):
        """
        save reply/forward and retrieve the item result for further modification,
        you may want to use account.drafts as the folder.
        """
        if not self.account:
            raise ValueError('%s must have an account' % self.__class__.__name__)
        res = self.account.bulk_create(items=[self], folder=folder, message_disposition=SAVE_ONLY)
        if res and isinstance(res[0], Exception):
            raise res[0]
        res = list(self.account.fetch(res))  # retrieve result
        if res and isinstance(res[0], Exception):
            raise res[0]
        return res[0]
