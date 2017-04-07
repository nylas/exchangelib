# Add noqa on top-level convenience imports
from __future__ import unicode_literals

from .account import Account
from .attachments import FileAttachment, ItemAttachment
from .autodiscover import discover
from .configuration import Configuration
from .credentials import DELEGATE, IMPERSONATION, Credentials, ServiceAccount
from .ewsdatetime import EWSDateTime, EWSTimeZone, UTC, UTC_NOW
from .extended_properties import ExtendedProperty, ExternId
from .folders import SHALLOW, DEEP
from .items import CalendarItem, Contact, Message, Task
from .properties import Body, HTMLBody, ItemId, Mailbox, Attendee, Room, RoomList
from .restriction import Q
from .transport import BASIC, DIGEST, NTLM
from .version import Build, Version

__all__ = [
    'Account',
    'FileAttachment', 'ItemAttachment',
    'discover',
    'Configuration',
    'DELEGATE', 'IMPERSONATION', 'Credentials', 'ServiceAccount',
    'EWSDateTime', 'EWSTimeZone', 'UTC', 'UTC_NOW',
    'ExtendedProperty',
    'CalendarItem', 'Contact', 'Message', 'Task',
    'ItemId', 'Mailbox', 'Attendee', 'Room', 'RoomList', 'Body', 'HTMLBody',
    'Q',
    'SHALLOW', 'DEEP',
    'BASIC', 'DIGEST', 'NTLM',
    'Build', 'Version',
]


def close_connections():
    from .autodiscover import close_connections as close_autodiscover_connections
    from .protocol import close_connections as close_protocol_connections
    close_autodiscover_connections()
    close_protocol_connections()


# Pre-register these extended properties
CalendarItem.register('extern_id', ExternId)
Message.register('extern_id', ExternId)
Contact.register('extern_id', ExternId)
Task.register('extern_id', ExternId)
