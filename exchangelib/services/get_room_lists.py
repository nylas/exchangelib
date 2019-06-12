from ..util import create_element, MNS
from ..version import EXCHANGE_2010
from .common import EWSService


class GetRoomLists(EWSService):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/dd899486(v=exchg.150).aspx
    """
    SERVICE_NAME = 'GetRoomLists'
    element_container_name = '{%s}RoomLists' % MNS

    def call(self):
        from ..properties import RoomList

        if self.protocol.version.build < EXCHANGE_2010:
            raise NotImplementedError('%s is only supported for Exchange 2010 servers and later' % self.SERVICE_NAME)
        elements = self._get_elements(payload=self.get_payload())
        return [RoomList.from_xml(elem=elem, account=None) for elem in elements]

    def get_payload(self):
        return create_element('m:%s' % self.SERVICE_NAME)
