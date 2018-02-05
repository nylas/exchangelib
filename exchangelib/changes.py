from __future__ import unicode_literals

from exchangelib import Folder
from exchangelib.fields import EWSElementField, ItemField, FolderField
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
    ELEMENT_NAME = 'Delete'
