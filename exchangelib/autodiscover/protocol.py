from future.utils import python_2_unicode_compatible

from ..protocol import BaseProtocol


@python_2_unicode_compatible
class AutodiscoverProtocol(BaseProtocol):
    """Protocol which implements the bare essentials for autodiscover"""
    TIMEOUT = 10  # Seconds

    def __str__(self):
        return '''\
Autodiscover endpoint: %s
Auth type: %s''' % (
            self.service_endpoint,
            self.auth_type,
        )
