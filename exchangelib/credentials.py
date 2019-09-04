# coding=utf-8
"""
Implements an Exchange user object and access types. Exchange provides two different ways of granting access for a
login to a specific account. Impersonation is used mainly for service accounts that connect via EWS. Delegate is used
for ad-hoc access e.g. granted manually by the user.
See http://blogs.msdn.com/b/exchangedev/archive/2009/06/15/exchange-impersonation-vs-delegate-access.aspx
"""
from __future__ import unicode_literals

import logging

from future.utils import python_2_unicode_compatible

from .util import PickleMixIn

log = logging.getLogger(__name__)

IMPERSONATION = 'impersonation'
DELEGATE = 'delegate'
ACCESS_TYPES = (IMPERSONATION, DELEGATE)


class BaseCredentials(object):
    pass


@python_2_unicode_compatible
class Credentials(BaseCredentials, PickleMixIn):
    """
    Keeps login info the way Exchange likes it.

    :param username: Usernames for authentication are of one of these forms:
    * PrimarySMTPAddress
    * WINDOMAIN\\username
    * User Principal Name (UPN)

    :param password: Clear-text password
    """
    EMAIL = 'email'
    DOMAIN = 'domain'
    UPN = 'upn'

    __slots__ = ('username', 'password', 'type')

    def __init__(self, username, password):
        if username.count('@') == 1:
            self.type = self.EMAIL
        elif username.count('\\') == 1:
            self.type = self.DOMAIN
        else:
            self.type = self.UPN
        self.username = username
        self.password = password

    def __eq__(self, other):
        for k in self.__slots__:
            if getattr(self, k) != getattr(other, k):
                return False
        return True

    def __hash__(self):
        return hash((self.username, self.password))

    def __repr__(self):
        return self.__class__.__name__ + repr((self.username, '********'))

    def __str__(self):
        return self.username


@python_2_unicode_compatible
class OAuth2Credentials(BaseCredentials, PickleMixIn):
    """Login info for OAuth 2.0 authentication
    """
    __slots__ = ('client_id', 'client_secret', 'tenant_id')

    def __init__(self, client_id, client_secret, tenant_id):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id

    def __eq__(self, other):
        for k in self.__slots__:
            if getattr(self, k) != getattr(other, k):
                return False
        return True

    def __hash__(self):
        return hash((self.client_id, self.client_secret, self.tenant_id))

    def __repr__(self):
        return self.__class__.__name__ + repr((self.client_id, '********'))

    def __str__(self):
        return self.client_id
