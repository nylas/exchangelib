from collections import OrderedDict

from ..util import create_element
from ..version import EXCHANGE_2013_SP1
from .common import EWSAccountService, EWSPooledMixIn, create_item_ids_element


class DeleteItem(EWSAccountService, EWSPooledMixIn):
    """
    Takes a folder and a list of (id, changekey) tuples. Returns result of deletion as a list of tuples
    (success[True|False], errormessage), in the same order as the input list.

    MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/deleteitem

    """
    SERVICE_NAME = 'DeleteItem'
    element_container_name = None  # DeleteItem doesn't return a response object, just status in XML attrs

    def call(self, items, delete_type, send_meeting_cancellations, affected_task_occurrences, suppress_read_receipts):
        return self._pool_requests(payload_func=self.get_payload, **dict(
            items=items,
            delete_type=delete_type,
            send_meeting_cancellations=send_meeting_cancellations,
            affected_task_occurrences=affected_task_occurrences,
            suppress_read_receipts=suppress_read_receipts,
        ))

    def get_payload(self, items, delete_type, send_meeting_cancellations, affected_task_occurrences,
                    suppress_read_receipts):
        # Takes a list of (id, changekey) tuples or Item objects and returns the XML for a DeleteItem request.
        if self.account.version.build >= EXCHANGE_2013_SP1:
            deleteitem = create_element(
                'm:%s' % self.SERVICE_NAME,
                attrs=OrderedDict([
                    ('DeleteType', delete_type),
                    ('SendMeetingCancellations', send_meeting_cancellations),
                    ('AffectedTaskOccurrences', affected_task_occurrences),
                    ('SuppressReadReceipts', 'true' if suppress_read_receipts else 'false'),
                ])
            )
        else:
            deleteitem = create_element(
                'm:%s' % self.SERVICE_NAME,
                attrs=OrderedDict([
                    ('DeleteType', delete_type),
                    ('SendMeetingCancellations', send_meeting_cancellations),
                    ('AffectedTaskOccurrences', affected_task_occurrences),
                 ])
            )

        item_ids = create_item_ids_element(items=items, version=self.account.version)
        deleteitem.append(item_ids)
        return deleteitem
