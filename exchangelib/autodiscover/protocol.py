from ..protocol import BaseProtocol


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
