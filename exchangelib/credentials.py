"""
Implements an Exchange user object and access types. Exchange provides two different ways of granting access for a
login to a specific account. Impersonation is used mainly for service accounts that connect via EWS. Delegate is used
for ad-hoc access e.g. granted manually by the user.
See http://blogs.msdn.com/b/exchangedev/archive/2009/06/15/exchange-impersonation-vs-delegate-access.aspx
"""
import logging

log = logging.getLogger(__name__)

IMPERSONATION = 'impersonation'
DELEGATE = 'delegate'

EMAIL = 'email'
DOMAIN = 'domain'
UPN = 'upn'


class Credentials:
    # Keeps login info the way Exchange likes it. Usernames for authentication are of one of these forms:
    #
    #    * PrimarySMTPAddress
    #    * WINDOMAIN\username
    #    * User Principal Name (UPN)
    def __init__(self, username, password):
        if username.count('@') == 1:
            self.type = EMAIL
        elif username.count('\\') == 1:
            self.type = DOMAIN
        else:
            self.type = DOMAIN
        self.username = username
        self.password = password

    def __str__(self):
        return self.username
