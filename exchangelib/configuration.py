import logging

from cached_property import threaded_cached_property

from .credentials import BaseCredentials
from .protocol import RetryPolicy, FailFast
from .transport import AUTH_TYPE_MAP
from .util import split_url
from .version import Version

log = logging.getLogger(__name__)


class Configuration:
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
        if not isinstance(credentials, (BaseCredentials, type(None))):
            raise ValueError("'credentials' %r must be a Credentials instance" % credentials)
        if server and service_endpoint:
            raise AttributeError("Only one of 'server' or 'service_endpoint' must be provided")
        if auth_type is not None and auth_type not in AUTH_TYPE_MAP:
            raise ValueError("'auth_type' %r must be one of %s"
                             % (auth_type, ', '.join("'%s'" % k for k in sorted(AUTH_TYPE_MAP.keys()))))
        if not retry_policy:
            retry_policy = FailFast()
        if not isinstance(version, (Version, type(None))):
            raise ValueError("'version' %r must be a Version instance" % version)
        if not isinstance(retry_policy, RetryPolicy):
            raise ValueError("'retry_policy' %r must be a RetryPolicy instance" % retry_policy)
        self._credentials = credentials
        if server:
            self.service_endpoint = 'https://%s/EWS/Exchange.asmx' % server
        else:
            self.service_endpoint = service_endpoint
        self.auth_type = auth_type
        self.version = version
        self.retry_policy = retry_policy

    @property
    def credentials(self):
        # Do not update credentials from this class. Instead, do it from Protocol
        return self._credentials

    @threaded_cached_property
    def server(self):
        return split_url(self.service_endpoint)[1]

    def __repr__(self):
        return self.__class__.__name__ + '(%s)' % ', '.join('%s=%r' % (k, getattr(self, k)) for k in (
            'credentials', 'service_endpoint', 'auth_type', 'version', 'retry_policy'
        ))
