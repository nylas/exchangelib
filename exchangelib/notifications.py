from __future__ import unicode_literals

from exchangelib.fields import TextField, BooleanField, SubField, EventListField
from exchangelib.properties import EWSElement
from exchangelib.util import TNS, MNS


class Notification(EWSElement):
    ELEMENT_NAME = 'Notification'
    FIELDS = [
        TextField('subscription_id', field_uri='SubscriptionId'),
        TextField('previous_watermark', field_uri='PreviousWatermark'),
        BooleanField('more_events', field_uri='MoreEvents'),
        EventListField('events'),
    ]
    NAMESPACE = MNS
    __slots__ = ('subscription_id', 'previous_watermark', 'more_events', 'events')


class ConnectionStatus(EWSElement):
    ELEMENT_NAME = 'ConnectionStatus'
    NAMESPACE = MNS
    FIELDS = [
        SubField('status'),
    ]
    __slots__ = ('status',)
