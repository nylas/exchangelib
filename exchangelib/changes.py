from __future__ import unicode_literals

from exchangelib.fields import ItemField, FolderField, IdAndChangekeyField
from exchangelib.properties import EWSElement
from exchangelib.transport import TNS


class Change(EWSElement):
    NAMESPACE = TNS


class ItemChange(Change):
    FIELDS = [
        ItemField('item'),
    ]
    __slots__ = ('item',)


class CreateItemChange(ItemChange):
    ELEMENT_NAME = 'Create'


class UpdateItemChange(ItemChange):
    ELEMENT_NAME = 'Update'


class DeleteItemChange(ItemChange):
    ELEMENT_NAME = 'Delete'
    FIELDS = [
        IdAndChangekeyField('item_id', field_uri='ItemId'),
    ]
    __slots__ = ('item_id',)


class ReadFlagChange(ItemChange):
    ELEMENT_NAME = 'ReadFlagChange'


class FolderChange(Change):
    FIELDS = [
        FolderField('folder'),
    ]
    __slots__ = ('folder',)


class CreateFolderChange(FolderChange):
    ELEMENT_NAME = 'Create'


class UpdateFolderChange(FolderChange):
    ELEMENT_NAME = 'Update'


class DeleteFolderChange(FolderChange):
    FIELDS = [
        IdAndChangekeyField('item_id', field_uri='FolderId'),
    ]
    __slots__ = ('item_id',)
    ELEMENT_NAME = 'Delete'
