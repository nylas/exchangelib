# Add noqa on top-level convenience imports
from __future__ import unicode_literals

from .account import Account
from .attachments import FileAttachment, ItemAttachment
from .autodiscover import discover
from .configuration import Configuration
from .credentials import DELEGATE, IMPERSONATION, Credentials, OAuthCredentials, ServiceAccount
from .ewsdatetime import EWSDate, EWSDateTime, EWSTimeZone, UTC, UTC_NOW
from .extended_properties import ExtendedProperty, ExternId, Flag, CalendarColor
from .folders import Folder, FolderCollection, SHALLOW, DEEP
from .items import AcceptItem, TentativelyAcceptItem, DeclineItem, CalendarItem, CancelCalendarItem, Contact, \
    DistributionList, Message, PostItem, Task
from .properties import Body, HTMLBody, ItemId, Mailbox, Attendee, Room, RoomList, UID
from .protocol import BaseProtocol
from .restriction import Q
from .transport import BASIC, DIGEST, NTLM, GSSAPI
from .version import Build, Version
from .settings import OofSettings

__version__ = '1.11.5'

__all__ = [
    '__version__',
    'Account',
    'FileAttachment', 'ItemAttachment',
    'discover',
    'Configuration',
    'DELEGATE', 'IMPERSONATION', 'Credentials', 'ServiceAccount',
    'EWSDate', 'EWSDateTime', 'EWSTimeZone', 'UTC', 'UTC_NOW',
    'ExtendedProperty',
    'AcceptItem', 'TentativelyAcceptItem', 'DeclineItem',
    'CalendarItem', 'CancelCalendarItem', 'Contact', 'DistributionList', 'Message', 'PostItem', 'Task',
    'ItemId', 'Mailbox', 'Attendee', 'Room', 'RoomList', 'Body', 'HTMLBody', 'UID',
    'OofSettings',
    'Q',
    'Folder', 'FolderCollection', 'SHALLOW', 'DEEP',
    'BASIC', 'DIGEST', 'NTLM', 'GSSAPI',
    'Build', 'Version',
]

# Set a default user agent, e.g. "exchangelib/3.1.1"
BaseProtocol.USERAGENT = "%s/%s" % (__name__, __version__)


def close_connections():
    from .autodiscover import close_connections as close_autodiscover_connections
    from .protocol import close_connections as close_protocol_connections
    close_autodiscover_connections()
    close_protocol_connections()


# Pre-register these extended properties. They are not part of the standard EWS fields but are useful for identification
# when item originates in an external system.

CalendarItem.register('extern_id', ExternId)
Message.register('extern_id', ExternId)
Contact.register('extern_id', ExternId)
Task.register('extern_id', ExternId)

############# Nylas Registered Extended Properties ###########
Message.register('flag', Flag)
Folder.register('color', CalendarColor)
