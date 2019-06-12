from ..util import create_element, add_xml_child, DummyResponse, StreamingBase64Parser, StreamingContentHandler, \
    ElementNotFound, MNS
from .common import EWSAccountService, create_attachment_ids_element


class GetAttachment(EWSAccountService):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa494316(v=exchg.150).aspx
    """
    SERVICE_NAME = 'GetAttachment'
    element_container_name = '{%s}Attachments' % MNS
    streaming = True

    def call(self, items, include_mime_content):
        return self._get_elements(payload=self.get_payload(
            items=items,
            include_mime_content=include_mime_content,
        ))

    def get_payload(self, items, include_mime_content):
        payload = create_element('m:%s' % self.SERVICE_NAME)
        # TODO: Support additional properties of AttachmentShape. See
        # https://msdn.microsoft.com/en-us/library/office/aa563727(v=exchg.150).aspx
        if include_mime_content:
            attachment_shape = create_element('m:AttachmentShape')
            add_xml_child(attachment_shape, 't:IncludeMimeContent', 'true')
            payload.append(attachment_shape)
        attachment_ids = create_attachment_ids_element(items=items, version=self.account.version)
        payload.append(attachment_ids)
        return payload

    @classmethod
    def _get_soap_payload(cls, response, **parse_opts):
        if not parse_opts.get('stream_file_content', False):
            return super(GetAttachment, cls)._get_soap_payload(response, **parse_opts)

        from ..attachments import FileAttachment
        parser = StreamingBase64Parser()
        field = FileAttachment.get_field_by_fieldname('_content')
        handler = StreamingContentHandler(parser=parser, ns=field.namespace, element_name=field.field_uri)
        parser.setContentHandler(handler)
        return parser.parse(response)

    def stream_file_content(self, attachment_id):
        # The streaming XML parser can only stream content of one attachment
        payload = self.get_payload(items=[attachment_id], include_mime_content=False)
        try:
            for chunk in self._get_response_xml(payload=payload, stream_file_content=True):
                yield chunk
        except ElementNotFound as enf:
            # When the returned XML does not contain a Content element, ElementNotFound is thrown by parser.parse().
            # Let the non-streaming SOAP parser parse the response and hook into the normal exception handling.
            # Wrap in DummyResponse because _get_soap_payload() expects an iter_content() method.
            response = DummyResponse(url=None, headers=None, request_headers=None, content=enf.data)
            res = super(GetAttachment, self)._get_soap_payload(response=response)
            for e in self._get_elements_in_response(response=res):
                if isinstance(e, Exception):
                    raise e
            # The returned content did not contain any EWS exceptions. Give up and re-raise the original exception.
            raise enf
