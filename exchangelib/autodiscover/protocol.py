from future.utils import python_2_unicode_compatible

from ..protocol import BaseProtocol, FailFast


@python_2_unicode_compatible
class AutodiscoverProtocol(BaseProtocol):
    """Protocol which implements the bare essentials for autodiscover"""
    TIMEOUT = 10  # Seconds
    # When connecting to servers that may not be serving the correct endpoint, we should use a retry policy that does
    # not leave us hanging for a long time on each step in the protocol.
    INITIAL_RETRY_POLICY = FailFast()

    def __str__(self):
        return '''\
Autodiscover endpoint: %s
Auth type: %s''' % (
            self.service_endpoint,
            self.auth_type,
        )
