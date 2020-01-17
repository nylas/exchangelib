from ..util import create_element, set_xml_value, MNS
from .common import EWSAccountService


class SetUserOofSettings(EWSAccountService):
    """
    Set automatic replies for the specified mailbox.
    MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/setuseroofsettings-operation
    """
    SERVICE_NAME = 'SetUserOofSettings'

    def call(self, oof_settings, mailbox):
        res = list(self._get_elements(payload=self.get_payload(oof_settings=oof_settings, mailbox=mailbox)))
        if len(res) != 1:
            raise ValueError("Expected 'res' length 1, got %s" % res)
        return res[0]

    def get_payload(self, oof_settings, mailbox):
        from ..properties import AvailabilityMailbox
        payload = create_element('m:%sRequest' % self.SERVICE_NAME)
        set_xml_value(payload, AvailabilityMailbox.from_mailbox(mailbox), version=self.account.version)
        set_xml_value(payload, oof_settings, version=self.account.version)
        return payload

    def _get_element_container(self, message, response_message=None, name=None):
        response_message = message.find('{%s}ResponseMessage' % MNS)
        return super()._get_element_container(
            message=message, response_message=response_message, name=name
        )
