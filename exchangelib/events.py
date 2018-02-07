from __future__ import unicode_literals

from xml.etree.ElementTree import Element

from typing import List

from exchangelib.fields import DateTimeField, IdField, Field, IntegerField, TextField, IdAndChangekeyField
from exchangelib.properties import EWSElement
from exchangelib.transport import TNS


class Event(EWSElement):
    ELEMENT_NAME = 'Event'
    NAMESPACE = TNS

    @classmethod
    def has_all_required_xml_fields(cls, elem):
        # type: (Element) -> bool
        for f in cls.FIELDS:
            if elem.find(f.response_tag()) is None and f.is_required:
                return False
        return True


class CopiedEvent(Event):
    ELEMENT_NAME = 'CopiedEvent'
    FIELDS = [  # type: List[Field]
        TextField('watermark', field_uri='Watermark', is_required=False),
        DateTimeField('timestamp', field_uri='TimeStamp'),
        IdAndChangekeyField('parent_folder_id', field_uri='ParentFolderId', is_attribute=False),
        IdAndChangekeyField('old_parent_folder_id', field_uri='OldParentFolderId', is_attribute=False)
    ]
    __slots__ = ('watermark', 'timestamp', 'parent_folder_id', 'old_folder_id')


class FolderCopiedEvent(CopiedEvent):
    FIELDS = CopiedEvent.FIELDS + [
        IdAndChangekeyField('folder_id', field_uri='FolderId', is_required=True, is_attribute=False),
        IdAndChangekeyField('old_folder_id', field_uri='OldFolderId', is_required=True, is_attribute=False),
    ]
    __slots__ = CopiedEvent.__slots__ + ('folder_id', 'old_folder_id')


class ItemCopiedEvent(CopiedEvent):
    FIELDS = CopiedEvent.FIELDS + [
        IdAndChangekeyField('item_id', field_uri='ItemId', is_required=True, is_attribute=False),
        IdAndChangekeyField('old_item_id', field_uri='OldItemId', is_required=True, is_attribute=False),
    ]
    __slots__ = CopiedEvent.__slots__ + ('item_id', 'old_item_id')


class CreatedEvent(Event):
    ELEMENT_NAME = 'CreatedEvent'
    FIELDS = [  # type: List[Field]
        TextField('watermark', field_uri='Watermark', is_required=False),
        DateTimeField('timestamp', field_uri='TimeStamp'),
        IdAndChangekeyField('parent_folder_id', field_uri='ParentFolderId', is_required=True, is_attribute=False),
    ]
    __slots__ = ('watermark', 'timestamp', 'parent_folder_id')


class ItemCreatedEvent(CreatedEvent):
    FIELDS = CreatedEvent.FIELDS + [
        IdAndChangekeyField('item_id', field_uri='ItemId', is_required=True, is_attribute=False),
    ]
    __slots__ = CreatedEvent.__slots__ + ('item_id',)


class FolderCreatedEvent(CreatedEvent):
    FIELDS = CreatedEvent.FIELDS + [
        IdAndChangekeyField('folder_id', field_uri='FolderId', is_required=True, is_attribute=False),
    ]
    __slots__ = CreatedEvent.__slots__ + ('folder_id',)


class DeletedEvent(Event):
    ELEMENT_NAME = 'DeletedEvent'
    FIELDS = [  # type: List[Field]
        TextField('watermark', field_uri='Watermark', is_required=False),
        DateTimeField('timestamp', field_uri='TimeStamp'),
        IdAndChangekeyField('parent_folder_id', field_uri='ParentFolderId', is_required=True, is_attribute=False),
    ]
    __slots__ = ('watermark', 'timestamp', 'parent_folder_id')


class ItemDeletedEvent(DeletedEvent):
    FIELDS = DeletedEvent.FIELDS + [
        IdAndChangekeyField('item_id', field_uri='ItemId', is_required=True, is_attribute=False),
    ]
    __slots__ = DeletedEvent.__slots__ + ('item_id',)


class FolderDeletedEvent(DeletedEvent):
    FIELDS = DeletedEvent.FIELDS + [
        IdAndChangekeyField('folder_id', field_uri='FolderId', is_required=True, is_attribute=False),
    ]
    __slots__ = DeletedEvent.__slots__ + ('folder_id',)


class ModifiedEvent(Event):
    ELEMENT_NAME = 'ModifiedEvent'
    FIELDS = [  # type: List[Field]
        TextField('watermark', field_uri='Watermark', is_required=False),
        DateTimeField('timestamp', field_uri='TimeStamp'),
        IdAndChangekeyField('parent_folder_id', field_uri='ParentFolderId', is_required=True, is_attribute=False),
        IntegerField('unread_count', field_uri='UnreadCount', is_read_only=True),
    ]
    __slots__ = ('watermark', 'timestamp', 'parent_folder_id', 'unread_count')


class ItemModifiedEvent(ModifiedEvent):
    FIELDS = ModifiedEvent.FIELDS + [
        IdAndChangekeyField('item_id', field_uri='ItemId', is_required=True, is_attribute=False),
    ]
    __slots__ = ModifiedEvent.__slots__ + ('item_id',)


class FolderModifiedEvent(ModifiedEvent):
    FIELDS = ModifiedEvent.FIELDS + [
        IdAndChangekeyField('folder_id', field_uri='FolderId', is_required=True, is_attribute=False),
    ]
    __slots__ = ModifiedEvent.__slots__ + ('folder_id',)


class MovedEvent(Event):
    ELEMENT_NAME = 'MovedEvent'
    FIELDS = [  # type: List[Field]
        TextField('watermark', field_uri='Watermark', is_required=False),
        DateTimeField('timestamp', field_uri='TimeStamp'),
        IdAndChangekeyField('parent_folder_id', field_uri='ParentFolderId', is_required=True, is_attribute=False),
        IdAndChangekeyField('old_parent_folder_id', field_uri='OldParentFolderId', is_required=True, is_attribute=False),
    ]
    __slots__ = ('watermark', 'timestamp', 'parent_folder_id', 'old_parent_folder_id')


class ItemMovedEvent(MovedEvent):
    FIELDS = MovedEvent.FIELDS + [
        IdAndChangekeyField('item_id', field_uri='ItemId', is_required=True, is_attribute=False),
        IdAndChangekeyField('old_item_id', field_uri='OldItemId', is_required=True, is_attribute=False),
    ]
    __slots__ = MovedEvent.__slots__ + ('item_id', 'old_item_id')


class FolderMovedEvent(MovedEvent):
    FIELDS = MovedEvent.FIELDS + [
        IdAndChangekeyField('folder_id', field_uri='FolderId', is_required=True, is_attribute=False),
        IdAndChangekeyField('old_folder_id', field_uri='OldFolderId', is_required=True, is_attribute=False),
    ]
    __slots__ = MovedEvent.__slots__ + ('folder_id', 'old_folder_id')


class NewMailEvent(Event):
    ELEMENT_NAME = 'NewMailEvent'
    FIELDS = [
        TextField('watermark', field_uri='Watermark', is_required=False),
        DateTimeField('timestamp', field_uri='TimeStamp'),
        IdAndChangekeyField('item_id', field_uri='ItemId', is_required=True, is_attribute=False),
        IdAndChangekeyField('parent_folder_id', field_uri='ParentFolderId', is_required=True, is_attribute=False),
    ]
    __slots__ = ('watermark', 'timestamp', 'item_id', 'parent_folder_id')


class StatusEvent(Event):
    ELEMENT_NAME = 'StatusEvent'
    FIELDS = [
        TextField('watermark', field_uri='Watermark', is_required=False),
    ]
    __slots__ = ('watermark',)


class FreeBusyChangedEvent(Event):
    ELEMENT_NAME = 'FreeBusyChangedEvent'
    FIELDS = [
        TextField('watermark', field_uri='Watermark', is_required=False),
        DateTimeField('timestamp', field_uri='TimeStamp'),
        IdAndChangekeyField('item_id', field_uri='ItemId', is_required=True, is_attribute=False),
        IdAndChangekeyField('parent_folder_id', field_uri='ParentFolderId', is_required=True, is_attribute=False),
    ]
    __slots__ = ('watermark', 'timestamp', 'item_id', 'parent_folder_id')


EVENT_TYPES = [
    CopiedEvent,
    FolderCopiedEvent,
    ItemCopiedEvent,
    CreatedEvent,
    ItemCreatedEvent,
    FolderCreatedEvent,
    DeletedEvent,
    ItemDeletedEvent,
    FolderDeletedEvent,
    ModifiedEvent,
    ItemModifiedEvent,
    FolderModifiedEvent,
    MovedEvent,
    ItemMovedEvent,
    FolderMovedEvent,
    NewMailEvent,
    StatusEvent,
    FreeBusyChangedEvent,
]

CONCRETE_EVENT_CLASSES = [
    FolderCopiedEvent,
    ItemCopiedEvent,
    ItemCreatedEvent,
    FolderCreatedEvent,
    ItemDeletedEvent,
    FolderDeletedEvent,
    ItemModifiedEvent,
    FolderModifiedEvent,
    ItemMovedEvent,
    FolderMovedEvent,
    NewMailEvent,
    StatusEvent,
    FreeBusyChangedEvent,
]
