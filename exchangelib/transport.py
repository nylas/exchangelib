import logging
import time

import requests.auth
import requests_ntlm
import requests_oauthlib

from .credentials import IMPERSONATION
from .errors import UnauthorizedError, TransportError
from .util import create_element, add_xml_child, xml_to_str, ns_translation, _may_retry_on_error, _back_off_if_needed, \
    DummyResponse, CONNECTION_ERRORS

log = logging.getLogger(__name__)

# Authentication method enums
NOAUTH = 'no authentication'
NTLM = 'NTLM'
BASIC = 'basic'
DIGEST = 'digest'
GSSAPI = 'gssapi'
SSPI = 'sspi'
OAUTH2 = 'OAuth 2.0'

AUTH_TYPE_MAP = {
    NTLM: requests_ntlm.HttpNtlmAuth,
    BASIC: requests.auth.HTTPBasicAuth,
    DIGEST: requests.auth.HTTPDigestAuth,
    OAUTH2: requests_oauthlib.OAuth2,
    NOAUTH: None,
}
try:
    import requests_kerberos
    AUTH_TYPE_MAP[GSSAPI] = requests_kerberos.HTTPKerberosAuth
except ImportError:
    # Kerberos auth is optional
    pass
try:
    import requests_negotiate_sspi
    AUTH_TYPE_MAP[SSPI] = requests_negotiate_sspi.HttpNegotiateAuth
except ImportError:
    # SSPI auth is optional
    pass

DEFAULT_ENCODING = 'utf-8'
DEFAULT_HEADERS = {'Content-Type': 'text/xml; charset=%s' % DEFAULT_ENCODING, 'Accept-Encoding': 'gzip, deflate'}


def extra_headers(account):
    """Generate extra HTTP headers
    """
    if account:
        # See
        # https://blogs.msdn.microsoft.com/webdav_101/2015/05/11/best-practices-ews-authentication-and-access-issues/
        return {'X-AnchorMailbox': account.primary_smtp_address}
    return None


def wrap(content, api_version, account=None):
    """
    Generate the necessary boilerplate XML for a raw SOAP request. The XML is specific to the server version.
    ExchangeImpersonation allows to act as the user we want to impersonate.
    """
    envelope = create_element('s:Envelope', nsmap=ns_translation)
    header = create_element('s:Header')
    requestserverversion = create_element('t:RequestServerVersion', attrs=dict(Version=api_version))
    header.append(requestserverversion)
    if account:
        if account.access_type == IMPERSONATION:
            exchangeimpersonation = create_element('t:ExchangeImpersonation')
            connectingsid = create_element('t:ConnectingSID')
            add_xml_child(connectingsid, 't:PrimarySmtpAddress', account.primary_smtp_address)
            exchangeimpersonation.append(connectingsid)
            header.append(exchangeimpersonation)
        timezonecontext = create_element('t:TimeZoneContext')
        timezonedefinition = create_element('t:TimeZoneDefinition', attrs=dict(Id=account.default_timezone.ms_id))
        timezonecontext.append(timezonedefinition)
        header.append(timezonecontext)
    envelope.append(header)
    body = create_element('s:Body')
    body.append(content)
    envelope.append(body)
    return xml_to_str(envelope, encoding=DEFAULT_ENCODING, xml_declaration=True)


def get_auth_instance(auth_type, **kwargs):
    """
    Returns an *Auth instance suitable for the requests package
    """
    model = AUTH_TYPE_MAP[auth_type]
    if model is None:
        return None
    if auth_type == GSSAPI:
        # Kerberos auth relies on credentials supplied via a ticket available externally to this library
        return model()
    if auth_type == SSPI:
        # SSPI auth does not require credentials, but can have it
        return model(**kwargs)
    return model(**kwargs)


def get_service_authtype(service_endpoint, retry_policy, api_versions, name):
    # Get auth type by tasting headers from the server. Only do POST requests. HEAD is too error prone, and some servers
    # are set up to redirect to OWA on all requests except POST to /EWS/Exchange.asmx
    #
    # We don't know the API version yet, but we need it to create a valid request because some Exchange servers only
    # respond when given a valid request. Try all known versions. Gross.
    from .protocol import BaseProtocol
    retry = 0
    wait = 10  # seconds
    t_start = time.monotonic()
    headers = DEFAULT_HEADERS.copy()
    for api_version in api_versions:
        data = dummy_xml(api_version=api_version, name=name)
        log.debug('Requesting %s from %s', data, service_endpoint)
        while True:
            _back_off_if_needed(retry_policy.back_off_until)
            log.debug('Trying to get service auth type for %s', service_endpoint)
            with BaseProtocol.raw_session() as s:
                try:
                    r = s.post(url=service_endpoint, headers=headers, data=data, allow_redirects=False,
                               timeout=BaseProtocol.TIMEOUT)
                    break
                except CONNECTION_ERRORS as e:
                    # Don't retry on TLS errors. They will most likely be persistent.
                    total_wait = time.monotonic() - t_start
                    r = DummyResponse(url=service_endpoint, headers={}, request_headers=headers)
                    if _may_retry_on_error(response=r, retry_policy=retry_policy, wait=total_wait):
                        log.info("Connection error on URL %s (retry %s, error: %s). Cool down %s secs",
                                 service_endpoint, retry, e, wait)
                        retry_policy.back_off(wait)
                        retry += 1
                        continue
                    else:
                        raise TransportError(str(e)) from e
        if r.status_code not in (200, 401):
            log.debug('Unexpected response: %s %s', r.status_code, r.reason)
            continue
        try:
            auth_type = get_auth_method_from_response(response=r)
            log.debug('Auth type is %s', auth_type)
            return auth_type, api_version
        except UnauthorizedError:
            continue
    raise TransportError('Failed to get auth type from service')


def get_auth_method_from_response(response):
    # First, get the auth method from headers. Then, test credentials. Don't handle redirects - burden is on caller.
    log.debug('Request headers: %s', response.request.headers)
    log.debug('Response headers: %s', response.headers)
    if response.status_code == 200:
        return NOAUTH
    # Get auth type from headers
    for key, val in response.headers.items():
        if key.lower() == 'www-authenticate':
            # Requests will combine multiple HTTP headers into one in 'request.headers'
            vals = _tokenize(val.lower())
            for v in vals:
                if v.startswith('realm'):
                    realm = v.split('=')[1].strip('"')
                    log.debug('realm: %s', realm)
            # Prefer most secure auth method if more than one is offered. See discussion at
            # http://docs.oracle.com/javase/7/docs/technotes/guides/net/http-auth.html
            if 'digest' in vals:
                return DIGEST
            if 'ntlm' in vals:
                return NTLM
            if 'basic' in vals:
                return BASIC
    raise UnauthorizedError('No compatible auth type was reported by server')


def _tokenize(val):
    # Splits cookie auth values
    auth_methods = []
    auth_method = ''
    quote = False
    for c in val:
        if c in (' ', ',') and not quote:
            if auth_method not in ('', ','):
                auth_methods.append(auth_method)
            auth_method = ''
            continue
        elif c == '"':
            auth_method += c
            if quote:
                auth_methods.append(auth_method)
                auth_method = ''
            quote = not quote
            continue
        auth_method += c
    if auth_method:
        auth_methods.append(auth_method)
    return auth_methods


def dummy_xml(api_version, name):
    # Generate a minimal, valid EWS request
    from .services import ResolveNames  # Avoid circular import
    return wrap(content=ResolveNames(protocol=None).get_payload(
        unresolved_entries=[name],
        parent_folders=None,
        return_full_contact_data=False,
        search_scope=None,
        contact_data_shape=None,
    ), api_version=api_version)
