from exchangelib.fields import CharField, DateTimeField, TextField, ChoiceField, Choice
from exchangelib.properties import EWSElement
from exchangelib.transport import TNS
from exchangelib.util import create_element

class DurationDatetimeField(DateTimeField):

    _CONTAINER = '{%s}Duration' % (TNS)
    def from_xml(self, elem, account):
        duration = elem.find(self._CONTAINER)
        if duration is not None:
            return super().from_xml(duration, account)

    def to_xml(self, value, version):
        return super().to_xml(value, version)

class ReplyField(TextField):
    def from_xml(self, elem, account):
        reply = field_elem = elem.find(self.response_tag())
        if reply is not None:
            message = reply.find('{%s}Message' % TNS)
            if message is not None:
                return message.text

class OofSettings(EWSElement):
    NAMESPACE = TNS
    ELEMENT_NAME = 'OofSettings'
    FIELDS = [
        ChoiceField('state', field_uri='OofState', is_required=True,
                    choices={Choice('Enabled'), Choice('Scheduled'), Choice('Disabled')}),
        ChoiceField('external_audience', field_uri='ExternalAudience',
                    choices={Choice('None'), Choice('Known'), Choice('All')}, default='All'),
        DurationDatetimeField('time_start', field_uri='StartTime'),
        DurationDatetimeField('time_end', field_uri='EndTime'),
        ReplyField('reply_internal', field_uri='InternalReply'),
        ReplyField('reply_external', field_uri='ExternalReply'),
    ]

    def to_xml(self, version):
        result = super().to_xml(version)
        result.tag = 't:UserOofSettings'
        duration = create_element('t:Duration')
        for name in 'StartTime', 'EndTime':
            elem = result.find('t:%s' %  name)
            result.remove(elem)
            duration.append(elem)
        result.append(duration)
        for name in 'InternalReply', 'ExternalReply':
            elem = result.find('t:%s' %  name)
            message = create_element('t:Message')
            message.text = elem.text
            elem.text = None
            elem.append(message)
        return result