import logging

from .credentials import Credentials
from .protocol import Protocol

log = logging.getLogger(__name__)


class Configuration:
    """
    Stores default credentials when connecting via a system account, and default connection protocol when autodiscover
    is not activated for an Account.
    """
    def __init__(self, server=None, username=None, password=None, has_ssl=True, ews_auth_type=None, ews_url=None):
        if username:
            if not password:
                raise AttributeError('Password must be provided when username is provided')
            self.credentials = Credentials(username, password)
        else:
            self.credentials = None
        if server or ews_url:
            if not self.credentials:
                raise AttributeError('Credentials must be provided when server is provided')
            # Set up a default protocol that non-autodiscover accounts can use
            if not ews_url:
                ews_url = '%s://%s/EWS/Exchange.asmx' % ('https' if has_ssl else 'http', server)
            self.protocol = Protocol(
                ews_url=ews_url,
                ews_auth_type=ews_auth_type,
                has_ssl=has_ssl,
                credentials=self.credentials,
            )
        else:
            self.protocol = None
