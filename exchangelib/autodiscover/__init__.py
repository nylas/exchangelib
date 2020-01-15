# coding=utf-8
"""
Autodiscover is a Microsoft method for automatically getting the endpoint of the Exchange server and other
connection-related settings holding the email address using only the email address, and username and password of the
user.

The protocol for autodiscovering an email address is described in detail in
https://docs.microsoft.com/en-us/previous-versions/office/developer/exchange-server-interoperability-guidance. Handling
error messages is described here:
https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/handling-autodiscover-error-messages.

WARNING: The autodiscover protocol is very complicated. If you have problems autodiscovering using this implementation,
start by doing an official test at https://testconnectivity.microsoft.com
"""
import os

from .cache import AutodiscoverCache, autodiscover_cache
from .discovery import Autodiscovery
from .protocol import AutodiscoverProtocol

if os.environ.get('EXCHANGELIB_AUTODISCOVER_VERSION', 'legacy') == 'legacy':
    # Default to the legacy implementation
    from .legacy import discover
else:
    from .discovery import discover


def close_connections():
    with autodiscover_cache:
        autodiscover_cache.close()


def clear_cache():
    with autodiscover_cache:
        autodiscover_cache.clear()


__all__ = [
    'AutodiscoverCache', 'AutodiscoverProtocol', 'Autodiscovery', 'discover', 'autodiscover_cache',
    'close_connections', 'clear_cache'
]
