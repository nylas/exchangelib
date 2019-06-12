from . import move_item


class CopyItem(move_item.MoveItem):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa565012(v=exchg.150).aspx
    """
    SERVICE_NAME = 'CopyItem'
