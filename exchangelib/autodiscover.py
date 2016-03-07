"""
Autodiscover is a Microsoft method for automatically getting the hostname of the Exchange server and the server
version of the server holding the email address using only the email address and password of the user (and possibly
User Principal Name). The protocol for autodiscovering an email address is described in detail in
http://msdn.microsoft.com/en-us/library/hh352638(v=exchg.140).aspx. Handling error messages is described here:
http://msdn.microsoft.com/en-us/library/office/dn467392(v=exchg.150).aspx. This is not fully implemented.

WARNING: We are taking many shortcuts here, like assuming SSL and following 302 Redirects automatically.
If you have problems autodiscovering, start by doing an official test at https://testconnectivity.microsoft.com
"""

# TODO: According to Microsoft, we may cache the URL of the autodiscover service forever, or until it stops responding.
# My previous experience with Exchange products in mind, I'm not sure if I should trust that advice. But it could save
# some valuable seconds every time we start a new connection to a known server. In any case, this info would require
# persistent storage.

from xml.etree.cElementTree import Element, tostring
import logging
from threading import Lock
import queue
from urllib import parse

import requests.exceptions
import dns.resolver

from .credentials import Credentials
from .version import API_VERSIONS
from .errors import AutoDiscoverFailed, AutoDiscoverRedirect, TransportError, RedirectError, ErrorNonExistentMailbox
from .protocol import Protocol
from . import transport
from .util import get_xml_attr, set_xml_attr, to_xml, is_xml, post_ratelimited, get_redirect_url

log = logging.getLogger(__name__)

REQUEST_NS = 'http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006'
AUTODISCOVER_NS = 'http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006'
ERROR_NS = 'http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006'
RESPONSE_NS = 'http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a'

REQUEST_TIMEOUT = 10  # Seconds

# Used to cache the autoconfigure URL for a specific hostname
_autodiscover_cache = {}
_autodiscover_cache_lock = Lock()


def close_connections():
    for domain, cached_values in _autodiscover_cache.items():
        log.debug('Domain %s: Closing sessions', domain)
        autodiscover_protocol = cached_values[2]
        autodiscover_protocol.close()


def discover(email, credentials):
    """
    Performs the autodiscover dance and returns a Protocol on success. The autodiscover and EWs server might not be
    the same, so we use a different Protocol to do the autodiscover request, and return a hopefully-cached Protocol
    to the callee.
    """
    log.debug('Attempting autodiscover on email %s', email)
    assert isinstance(credentials, Credentials)
    domain = get_domain(email)
    # Use lock to guard against multiple threads competing to cache information
    if domain in _autodiscover_cache:
        # Python dict() is thread safe, so accessing _autodiscover_cache without a lock should be OK
        hostname, has_ssl, protocol = _autodiscover_cache[domain]
        assert isinstance(protocol, AutodiscoverProtocol)
        log.debug('Cache hit for domain %s: %s', domain, hostname)
        try:
            # This is the main path when the cache is primed
            primary_smtp_address, protocol = _autodiscover_quick(hostname=hostname, credentials=credentials,
                                                                 email=email, has_ssl=has_ssl,
                                                                 autodiscover_protocol=protocol)
            assert primary_smtp_address
            assert isinstance(protocol, Protocol)
            return primary_smtp_address, protocol
        except AutoDiscoverRedirect as e:
            log.debug('%s redirects to %s', email, e.redirect_email)
            if email.lower() == e.redirect_email.lower():
                raise AutoDiscoverFailed('Redirect to same email address: %s' % email) from e
            # Start over with the new email address
            return discover(e.redirect_email, credentials)
        # This is unreachable

    log.debug('Waiting for _autodiscover_cache_lock')
    with _autodiscover_cache_lock:
        log.debug('_autodiscover_cache_lock acquired')
        # Don't recurse while holding the lock!
        if domain in _autodiscover_cache:
            # Cache was primed by some other thread while we were waiting for the lock.
            log.debug('Cache filled for domain %s while we were waiting', domain)
        else:
            log.debug('Cache miss for domain %s', domain)
            log.debug('Cache contents: %s', _autodiscover_cache)
            try:
                # This eventually fills the cache in _autodiscover_hostname
                primary_smtp_address, protocol = _try_autodiscover(hostname=domain, credentials=credentials,
                                                                   email=email)
                assert primary_smtp_address
                assert isinstance(protocol, Protocol)
                return primary_smtp_address, protocol
            except AutoDiscoverRedirect as e:
                if email.lower() == e.redirect_email.lower():
                    raise AutoDiscoverFailed('Redirect to same email address: %s' % email) from e
                log.debug('%s redirects to %s', email, e.redirect_email)
                email = e.redirect_email
            finally:
                log.debug('Releasing_autodiscover_cache_lock')
    # We fell out of the with statement, so either cache was filled by someone else, or autodiscover redirected us to
    # another email address. Start over.
    return discover(email, credentials)


def _try_autodiscover(hostname, credentials, email):
    # Implements the full chain of autodiscover server discovery attempts. Tries to return autodiscover data from the
    # final host.
    try:
        return _autodiscover_hostname(hostname=hostname, credentials=credentials, email=email)
    except RedirectError as e:
        return _try_autodiscover(e.server, credentials, email)
    except AutoDiscoverFailed:
        try:
            return _autodiscover_hostname(hostname='autodiscover.%s' % hostname, credentials=credentials, email=email)
        except RedirectError as e:
            return _try_autodiscover(e.server, credentials, email)
        except AutoDiscoverFailed:
            try:
                return _autodiscover_hostname(hostname='autodiscover.%s' % hostname, credentials=credentials,
                                              email=email, has_ssl=False)
            except RedirectError as e:
                return _try_autodiscover(e.server, credentials, email)
            except AutoDiscoverFailed:
                try:
                    hostname_from_dns = _get_canonical_name(hostname='autodiscover.%s' % hostname)
                    if not hostname_from_dns:
                        hostname_from_dns = _get_hostname_from_srv(hostname='autodiscover.%s' % hostname)
                    # Start over with new hostname
                    return _try_autodiscover(hostname=hostname_from_dns, credentials=credentials, email=email)
                except AutoDiscoverFailed:
                    hostname_from_dns = _get_hostname_from_srv(hostname='_autodiscover._tcp.%s' % hostname)
                    # Start over with new hostname
                    return _try_autodiscover(hostname=hostname_from_dns, credentials=credentials, email=email)


def _autodiscover_hostname(hostname, credentials, email, has_ssl=True, auth_type=None):
    # Tries to get autodiscover data on a specific host. If we are HTTP redirected, we restart the autodiscover dance on
    # the new host.
    scheme = 'https' if has_ssl else 'http'
    url = '%s://%s/Autodiscover/Autodiscover.xml' % (scheme, hostname)
    log.debug('Trying autodiscover on %s', url)
    if not auth_type:
        try:
            auth_type = _get_autodiscover_auth_type(hostname=hostname, url=url, has_ssl=has_ssl, email=email)
        except RedirectError as e:
            log.debug(e)
            redirect_url, redirect_hostname, redirect_has_ssl = e.url, e.server, e.has_ssl
            log.debug('We were redirected to %s', redirect_url)
            canonical_hostname = _get_canonical_name(redirect_hostname)
            if canonical_hostname:
                log.debug('Canonical hostname is %s', canonical_hostname)
                redirect_hostname = canonical_hostname
            # Try the process on the new host, without 'www'. This is beyond the autodiscover protocol
            if redirect_hostname.startswith('www.'):
                redirect_hostname = redirect_hostname[4:]
            if redirect_hostname == hostname:
                log.debug('We were redirected to the same host')
                raise AutoDiscoverFailed('We were redirected to the same host') from e
            raise RedirectError(url=None, server=redirect_hostname, has_ssl=redirect_has_ssl) from e

    autodiscover_protocol = AutodiscoverProtocol(url=url, has_ssl=has_ssl, credentials=credentials, auth_type=auth_type)
    r = _get_autodiscover_response(protocol=autodiscover_protocol, url=url, email=email, verify=True)
    if r.status_code == 302:
        redirect_url, server, has_ssl = get_redirect_url(r, hostname, has_ssl)
        log.debug('We were redirected to %s', redirect_url)
        # Don't raise RedirectError here because we need to pass the ssl and auth_type data
        return _autodiscover_hostname(server, credentials, email, has_ssl=has_ssl, auth_type=None)
    domain = get_domain(email)
    try:
        server, has_ssl, ews_url, ews_auth_type, primary_smtp_address = _parse_response(r.text)
        if not primary_smtp_address:
            primary_smtp_address = email
    except (ErrorNonExistentMailbox, AutoDiscoverRedirect):
        # These are both valid responses from an autodiscover server, showing that we have found the correct
        # server for the original domain. Fill cache before re-raising
        log.debug('Adding cache entry for %s (hostname %s, has_ssl %s)' % (domain, hostname, has_ssl))
        _autodiscover_cache[domain] = hostname, has_ssl, autodiscover_protocol
        raise

    real_ews_auth_type = transport.get_service_authtype(server=server, has_ssl=has_ssl, ews_url=ews_url,
                                                        versions=API_VERSIONS)
    if ews_auth_type != real_ews_auth_type:
        log.debug('Autodiscover and real server disagree on auth method for %s (%s vs %s). Using server version',
                  email, ews_auth_type, real_ews_auth_type)
        ews_auth_type = real_ews_auth_type

    # Cache the final hostname of the autodiscover service so we don't need to autodiscover the same domain again
    log.debug('Adding cache entry for %s (hostname %s, has_ssl %s)' % (domain, hostname, has_ssl))
    _autodiscover_cache[domain] = hostname, has_ssl, autodiscover_protocol
    protocol = Protocol(has_ssl=has_ssl, ews_url=ews_url, credentials=credentials, ews_auth_type=ews_auth_type)
    return primary_smtp_address, protocol


def _autodiscover_quick(hostname, credentials, email, has_ssl, autodiscover_protocol):
    scheme = 'https' if has_ssl else 'http'
    url = '%s://%s/autodiscover/autodiscover.xml' % (scheme, hostname)
    r = _get_autodiscover_response(protocol=autodiscover_protocol, url=url, email=email, verify=True)
    server, has_ssl, ews_url, ews_auth_type, primary_smtp_address = _parse_response(r.text)
    if not primary_smtp_address:
        primary_smtp_address = email
    log.debug('Autodiscover success: %s may connect to %s as primary email %s', email, ews_url, primary_smtp_address)
    protocol = Protocol(has_ssl=has_ssl, ews_url=ews_url, credentials=credentials, ews_auth_type=ews_auth_type)
    return primary_smtp_address, protocol


def _get_autodiscover_auth_type(hostname, url, has_ssl, email, encoding='utf-8'):
    try:
        data = _get_autodiscover_payload(email=email, encoding=encoding)
        return transport.get_autodiscover_authtype(server=hostname, has_ssl=has_ssl, url=url, data=data,
                                                   timeout=REQUEST_TIMEOUT)
    except (TransportError, requests.exceptions.ConnectionError, requests.exceptions.Timeout,
            requests.exceptions.SSLError) as e:
        if isinstance(e, RedirectError):
            raise
        log.debug('Error guessing auth type: %s', e)
        raise AutoDiscoverFailed('Error guessing auth type: %s' % e) from e


def _get_autodiscover_payload(email, encoding='utf-8'):
    # Builds a full Autodiscover XML request
    payload = Element('Autodiscover')
    payload.set('xmlns', REQUEST_NS)
    request = Element('Request')
    set_xml_attr(request, 'EMailAddress', email)
    set_xml_attr(request, 'AcceptableResponseSchema', RESPONSE_NS)
    payload.append(request)
    return ('<?xml version="1.0" encoding="%s"?>' % encoding).encode(encoding) + tostring(payload, encoding=encoding)


def _get_autodiscover_response(protocol, url, email, encoding='utf-8', verify=True):
    data = _get_autodiscover_payload(email=email, encoding=encoding)
    headers = {'Content-Type': 'text/xml; charset=%s' % encoding}
    try:
        # Rate-limiting is an issue with autodiscover if the same setup is hosting EWS and autodiscover and we just
        # hammered the server with requests. We allow redirects since some autodiscover servers will issue different
        # redirects depending on the POST data content.
        session = protocol.get_session()
        r, session = post_ratelimited(protocol=protocol, session=session, url=url, headers=headers, data=data,
                                      timeout=protocol.timeout, verify=verify, allow_redirects=True)
        protocol.release_session(session)
        log.debug('Response headers: %s', r.headers)
    except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as e:
        log.debug('Connection error on %s: %s', url, e)
        # Don't raise AutoDiscoverFailed here. Connection errors could just as well be a valid but misbehaving server.
        raise
    except RedirectError:
        raise
    except TransportError:
        log.debug('No access to %s using %s', url, protocol.ews_auth_type)
        raise AutoDiscoverFailed('No access to %s using %s' % (url, protocol.ews_auth_type))
    if r.status_code == 302:
        # Give caller a chance to re-do the request
        return r
    if r.status_code != 200:
        log.debug('%s returned HTTP %s', url, r.status_code)
        # raise an uncatched error for now, until we understand this failure case
        raise TransportError('%s returned HTTP %s' % (url, r.status_code))
    if not is_xml(r.text):
        # This is normal - e.g. a greedy webserver serving custom HTTP 404's as 200 OK
        log.debug('URL %s: This is not XML: %s', url, r.text[:1000])
        raise AutoDiscoverFailed('URL %s: This is not XML: %s' % (url, r.text[:1000]))
    return r


def _parse_response(response, encoding='utf-8'):
    # We could return lots more interesting things here
    # log.debug('Autodiscover response: %s', response)
    autodiscover = to_xml(response, encoding=encoding)
    resp = autodiscover.find('{%s}Response' % RESPONSE_NS)
    if resp is None:
        resp = autodiscover.find('{%s}Response' % ERROR_NS)
        error = resp.find('{%s}Error' % ERROR_NS)
        errorcode = get_xml_attr(error, '{%s}ErrorCode' % ERROR_NS)
        message = get_xml_attr(error, '{%s}Message' % ERROR_NS)
        if message in ('The e-mail address cannot be found.', "The email address can't be found."):
            raise ErrorNonExistentMailbox('The SMTP address has no mailbox associated with it')
        raise AutoDiscoverFailed('Unknown error %s: %s' % (errorcode, message))
    account = resp.find('{%s}Account' % RESPONSE_NS)
    action = get_xml_attr(account, '{%s}Action' % RESPONSE_NS)
    redirect_email = get_xml_attr(account, '{%s}RedirectAddr' % RESPONSE_NS)
    if action == 'redirectAddr' and redirect_email:
        # This is redirection to e.g. Office365
        raise AutoDiscoverRedirect(redirect_email)
    user = resp.find('{%s}User' % RESPONSE_NS)
    # AutoDiscoverSMTPAddress might not be present in the XML, so primary_smtp_address might be None. In this
    # case, the original email address IS the primary address
    primary_smtp_address = get_xml_attr(user, '{%s}AutoDiscoverSMTPAddress' % RESPONSE_NS)
    account_type = get_xml_attr(account, '{%s}AccountType' % RESPONSE_NS)
    assert account_type == 'email'
    protocols = account.findall('{%s}Protocol' % RESPONSE_NS)
    # There are three possible protocol types: EXCH, EXPR and WEB. EXPR is for EWS. See
    # http://blogs.technet.com/b/exchange/archive/2008/09/26/3406344.aspx
    for protocol in protocols:
        if get_xml_attr(protocol, '{%s}Type' % RESPONSE_NS) != 'EXPR':
            continue
        server = get_xml_attr(protocol, '{%s}Server' % RESPONSE_NS)
        has_ssl = True if get_xml_attr(protocol, '{%s}SSL' % RESPONSE_NS) == 'On' else False
        ews_url = get_xml_attr(protocol, '{%s}EwsUrl' % RESPONSE_NS)
        auth_package = get_xml_attr(protocol, '{%s}AuthPackage' % RESPONSE_NS)
        try:
            ews_auth_type = {
                'ntlm': transport.NTLM,
                'basic': transport.BASIC,
                'digest': transport.DIGEST,
                None: transport.NOAUTH,
            }[auth_package.lower() if auth_package else None]
        except KeyError:
            log.warning("Unknown auth package '%s'")
            ews_auth_type = transport.UNKNOWN
        log.debug('Primary SMTP:%s, EWS endpoint:%s, auth type:%s', primary_smtp_address, ews_url, ews_auth_type)
        assert server
        assert has_ssl in (True, False)
        assert ews_url
        assert ews_auth_type
        return server, has_ssl, ews_url, ews_auth_type, primary_smtp_address
    raise AutoDiscoverFailed('Invalid AutoDiscover response: %s' % response)


def _get_canonical_name(hostname):
    log.debug('Attempting to get canonical name for %s', hostname)
    resolver = dns.resolver.Resolver()
    resolver.timeout = REQUEST_TIMEOUT
    try:
        canonical_name = resolver.query(hostname).canonical_name.to_unicode().rstrip('.')
    except dns.resolver.NXDOMAIN:
        log.debug('Nonexistent domain %s', hostname)
        return None
    if canonical_name != hostname:
        log.debug('%s has canonical name %s', hostname, canonical_name)
        return canonical_name
    return None


def _get_hostname_from_srv(hostname):
    # May return e.g.:
    #   canonical name = mail.ucl.dk.
    #   service = 8 100 443 webmail.ucn.dk.
    # or throw dns.resolver.NoAnswer
    log.debug('Attempting to get SRV record on %s', hostname)
    resolver = dns.resolver.Resolver()
    resolver.timeout = REQUEST_TIMEOUT
    try:
        answers = resolver.query(hostname, 'SRV')
        for rdata in answers:
            try:
                vals = rdata.to_text().strip().rstrip('.').split(' ')
                priority, weight, port, svr = int(vals[0]), int(vals[1]), int(vals[2]), vals[3]
            except (ValueError, KeyError) as e:
                raise AutoDiscoverFailed('Incompatible SRV record for %s (%s)' % (hostname, rdata.to_text())) from e
            else:
                return svr
    except dns.resolver.NoNameservers as e:
        raise AutoDiscoverFailed('No name servers for %s' % hostname) from e
    except dns.resolver.NoAnswer as e:
        raise AutoDiscoverFailed('No SRV record for %s' % hostname) from e
    except dns.resolver.NXDOMAIN as e:
        raise AutoDiscoverFailed('Nonexistent domain %s' % hostname) from e


def get_domain(email):
    try:
        return email.split('@')[1].lower().strip()
    except (IndexError, AttributeError) as e:
        raise ValueError("'%s' is not a valid email" % email) from e

POOLSIZE = 4


class AutodiscoverProtocol(Protocol):
    # Dummy class for post_ratelimited which implements the bare essentials
    SESSION_POOLSIZE = 1

    def __init__(self, url, has_ssl, credentials, auth_type):
        assert isinstance(credentials, Credentials)
        self.server = parse.urlparse(url).hostname.lower()
        self.credentials = credentials
        self.ews_url = url
        self.ews_auth_type = auth_type
        self.has_ssl = has_ssl
        self.timeout = REQUEST_TIMEOUT
        self._session_pool = queue.LifoQueue(maxsize=POOLSIZE)
        for i in range(POOLSIZE):
            self._session_pool.put(self.create_session(), block=False)
