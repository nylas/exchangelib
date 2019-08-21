from __future__ import unicode_literals

import logging

from .protocol import Protocol, FailFast

log = logging.getLogger(__name__)


class Configuration(object):
    """
    Assembles a connection protocol when autodiscover is not used.

    If the server is not configured with autodiscover, the following should be sufficient:

        config = Configuration(server='example.com', credentials=Credentials('MYWINDOMAIN\\myusername', 'topsecret'))
        account = Account(primary_smtp_address='john@example.com', config=config)

    You can also set the EWS service endpoint directly:

        config = Configuration(service_endpoint='https://mail.example.com/EWS/Exchange.asmx', credentials=...)

    If you know which authentication type the server uses, you add that as a hint:

        config = Configuration(service_endpoint='https://example.com/EWS/Exchange.asmx', auth_type=NTLM, credentials=..)

    If you want to use autodiscover, don't use a Configuration object. Instead, set up an account like this:

        credentials = Credentials(username='MYWINDOMAIN\\myusername', password='topsecret')
        account = Account(primary_smtp_address='john@example.com', credentials=credentials, autodiscover=True)

    """
    def __init__(self, credentials=None, server=None, service_endpoint=None, auth_type=None, version=None,
                 retry_policy=None):
        if not (server or service_endpoint):
            raise AttributeError('Either server or service_endpoint must be provided')
        # Set up a default protocol that non-autodiscover accounts can use
        if not service_endpoint:
            service_endpoint = 'https://%s/EWS/Exchange.asmx' % server
        if not retry_policy:
            retry_policy = FailFast()
        self.protocol = Protocol(
            service_endpoint=service_endpoint,
            auth_type=auth_type,
            credentials=credentials,
            version=version,
            retry_policy=retry_policy,
        )

    def __repr__(self):
        return self.__class__.__name__ + repr((self.protocol,))
