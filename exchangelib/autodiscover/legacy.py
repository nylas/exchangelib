import logging
import time

import dns.resolver

from ..configuration import Configuration
from ..credentials import BaseCredentials
from ..errors import AutoDiscoverFailed, AutoDiscoverRedirect, AutoDiscoverCircularRedirect, TransportError, \
    RedirectError, ErrorNonExistentMailbox, UnauthorizedError, RelativeRedirect
from ..protocol import Protocol, RetryPolicy, FailFast
from ..transport import DEFAULT_HEADERS, get_auth_method_from_response
from ..util import is_xml, post_ratelimited, get_domain, is_valid_hostname, _back_off_if_needed, _may_retry_on_error, \
    get_redirect_url, DummyResponse, CONNECTION_ERRORS, TLS_ERRORS
from .cache import autodiscover_cache
from .properties import Autodiscover
from .protocol import AutodiscoverProtocol

log = logging.getLogger(__name__)

# When connecting to servers that may not be serving the correct endpoint, we should use a retry policy that does
# not leave us hanging for a long time on each step in the protocol.
INITIAL_RETRY_POLICY = FailFast()


def discover(email, credentials=None, auth_type=None, retry_policy=None):
    """
    Performs the autodiscover dance and returns the primary SMTP address of the account and a Protocol on success. The
    autodiscover and EWS server might not be the same, so we use a different Protocol to do the autodiscover request,
    and return a hopefully-cached Protocol to the callee.
    """
    log.debug('Attempting autodiscover on email %s', email)
    if not isinstance(credentials, (BaseCredentials, type(None))):
        raise ValueError("'credentials' %r must be a Credentials instance" % credentials)
    if not isinstance(retry_policy, (RetryPolicy, type(None))):
        raise ValueError("'retry_policy' %r must be a RetryPolicy instance" % retry_policy)
    domain = get_domain(email)
    # We may be using multiple different credentials and changing our minds on TLS verification. This key combination
    # should be safe.
    autodiscover_key = (domain, credentials)
    # Use lock to guard against multiple threads competing to cache information
    log.debug('Waiting for autodiscover_cache lock')
    with autodiscover_cache:
        # Don't recurse while holding the lock!
        log.debug('autodiscover_cache lock acquired')
        if autodiscover_key in autodiscover_cache:
            protocol = autodiscover_cache[autodiscover_key]
            if not isinstance(protocol, AutodiscoverProtocol):
                raise ValueError('Unexpected autodiscover cache contents: %s' % protocol)
            # Reset auth type and retry policy if we requested non-default values
            if auth_type:
                protocol.config.auth_type = auth_type
            if retry_policy:
                protocol.config.retry_policy = retry_policy
            log.debug('Cache hit for domain %s credentials %s: %s', domain, credentials, protocol.service_endpoint)
            try:
                # This is the main path when the cache is primed
                return _autodiscover_quick(credentials=credentials, email=email, protocol=protocol)
            except AutoDiscoverFailed:
                # Autodiscover no longer works with this domain. Clear cache and try again after releasing the lock
                del autodiscover_cache[autodiscover_key]
            except AutoDiscoverRedirect as e:
                log.debug('%s redirects to %s', email, e.redirect_email)
                if email.lower() == e.redirect_email.lower():
                    raise AutoDiscoverCircularRedirect('Redirect to same email address: %s' % email) from None
                # Start over with the new email address after releasing the lock
                email = e.redirect_email
        else:
            log.debug('Cache miss for domain %s credentials %s', domain, credentials)
            log.debug('Cache contents: %s', autodiscover_cache)
            try:
                # This eventually fills the cache in _autodiscover_hostname
                return _try_autodiscover(hostname=domain, credentials=credentials, email=email,
                                         auth_type=auth_type, retry_policy=retry_policy)
            except AutoDiscoverRedirect as e:
                if email.lower() == e.redirect_email.lower():
                    raise AutoDiscoverCircularRedirect('Redirect to same email address: %s' % email) from None
                log.debug('%s redirects to %s', email, e.redirect_email)
                # Start over with the new email address after releasing the lock
                email = e.redirect_email
    log.debug('Released autodiscover_cache_lock')
    # We fell out of the with statement, so either cache was filled by someone else, or autodiscover redirected us to
    # another email address. Start over after releasing the lock.
    return discover(email=email, credentials=credentials, auth_type=auth_type, retry_policy=retry_policy)


def _try_autodiscover(hostname, credentials, email, auth_type, retry_policy):
    # Implements the full chain of autodiscover server discovery attempts. Tries to return autodiscover data from the
    # final host.
    try:
        return _autodiscover_hostname(hostname=hostname, credentials=credentials, email=email, has_ssl=True,
                                      auth_type=auth_type, retry_policy=retry_policy)
    except RedirectError as e:
        if not e.has_ssl:
            raise AutoDiscoverFailed(
                '%s redirected us to %s but only HTTPS redirects allowed' % (hostname, e.url)
            ) from None
        log.info('%s redirected us to %s', hostname, e.server)
        return _try_autodiscover(hostname=e.server, credentials=credentials, email=email, auth_type=auth_type,
                                 retry_policy=retry_policy)
    except AutoDiscoverFailed as e:
        log.info('Autodiscover on %s failed (%s). Trying autodiscover.%s', hostname, e, hostname)
        try:
            return _autodiscover_hostname(hostname='autodiscover.%s' % hostname, credentials=credentials, email=email,
                                          has_ssl=True, auth_type=auth_type, retry_policy=retry_policy)
        except RedirectError as e:
            if not e.has_ssl:
                raise AutoDiscoverFailed(
                    'autodiscover.%s redirected us to %s but only HTTPS redirects allowed' % (hostname, e.url)
                ) from None
            log.info('%s redirected us to %s', hostname, e.server)
            return _try_autodiscover(hostname=e.server, credentials=credentials, email=email,
                                     auth_type=auth_type, retry_policy=retry_policy)
        except AutoDiscoverFailed:
            log.info('Autodiscover on %s failed (%s). Trying autodiscover.%s (plain HTTP)', hostname, e, hostname)
            try:
                return _autodiscover_hostname(hostname='autodiscover.%s' % hostname, credentials=credentials,
                                              email=email, has_ssl=False, auth_type=auth_type,
                                              retry_policy=retry_policy)
            except RedirectError as e:
                if not e.has_ssl:
                    raise AutoDiscoverFailed(
                        'autodiscover.%s redirected us to %s but only HTTPS redirects allowed' % (hostname, e.url)
                    ) from None
                log.info('autodiscover.%s redirected us to %s', hostname, e.server)
                return _try_autodiscover(hostname=e.server, credentials=credentials, email=email,
                                         auth_type=auth_type, retry_policy=retry_policy)
            except AutoDiscoverFailed as e:
                log.info('Autodiscover on autodiscover.%s (no TLS) failed (%s). Trying DNS records', hostname, e)
                hostname_from_dns = _get_canonical_name(hostname='autodiscover.%s' % hostname)
                try:
                    if not hostname_from_dns:
                        log.info('No canonical name on autodiscover.%s Trying SRV record', hostname)
                        hostname_from_dns = _get_hostname_from_srv(hostname='autodiscover.%s' % hostname)
                    # Start over with new hostname
                    return _try_autodiscover(hostname=hostname_from_dns, credentials=credentials, email=email,
                                             auth_type=auth_type, retry_policy=retry_policy)
                except AutoDiscoverFailed as e:
                    log.info('Autodiscover on %s failed (%s). Trying _autodiscover._tcp.%s', hostname_from_dns, e,
                             hostname)
                    # Start over with new hostname
                    try:
                        hostname_from_dns = _get_hostname_from_srv(hostname='_autodiscover._tcp.%s' % hostname)
                        return _try_autodiscover(hostname=hostname_from_dns, credentials=credentials, email=email,
                                                 auth_type=auth_type, retry_policy=retry_policy)
                    except AutoDiscoverFailed:
                        raise AutoDiscoverFailed('All steps in the autodiscover protocol failed') from None


def _get_auth_type_or_raise(url, email, hostname, retry_policy):
    # Returns the auth type of the URL. Raises any redirection errors. This tests host DNS, port availability, and TLS
    # validation (if applicable).
    try:
        return _get_auth_type(url=url, email=email, retry_policy=retry_policy)
    except RedirectError as e:
        redirect_url, redirect_hostname, redirect_has_ssl = e.url, e.server, e.has_ssl
        log.debug('We were redirected to %s', redirect_url)
        if redirect_hostname.startswith('www.'):
            # Try the process on the new host, without 'www'. This is beyond the autodiscover protocol and an attempt to
            # work around seriously misconfigured Exchange servers. It's probably better to just show the Exchange
            # admins the report from https://testconnectivity.microsoft.com
            redirect_hostname = redirect_hostname[4:]
        canonical_hostname = _get_canonical_name(redirect_hostname)
        if canonical_hostname:
            log.debug('Canonical hostname is %s', canonical_hostname)
            redirect_hostname = canonical_hostname
        if redirect_hostname == hostname:
            log.debug('We were redirected to the same host')
            raise AutoDiscoverFailed('We were redirected to the same host') from None
        raise RedirectError(url='%s://%s' % ('https' if redirect_has_ssl else 'http', redirect_hostname)) from None


def _autodiscover_hostname(hostname, credentials, email, has_ssl, auth_type, retry_policy):
    # Tries to get autodiscover data on a specific host. If we are HTTP redirected, we restart the autodiscover dance on
    # the new host.
    url = '%s://%s/Autodiscover/Autodiscover.xml' % ('https' if has_ssl else 'http', hostname)
    log.info('Trying autodiscover on %s', url)
    if not is_valid_hostname(hostname, timeout=AutodiscoverProtocol.TIMEOUT):
        # 'requests' is really bad at reporting that a hostname cannot be resolved. Let's check this separately.
        raise AutoDiscoverFailed('%r has no DNS entry' % hostname) from None
    # We are connecting to an unknown server here. It's probable that servers in the autodiscover sequence are
    # unresponsive or send any kind of ill-formed response. We shouldn't use a retry policy meant for a trusted
    # endpoint here.
    if auth_type is None:
        auth_type = _get_auth_type_or_raise(url=url, email=email, hostname=hostname, retry_policy=INITIAL_RETRY_POLICY)
    autodiscover_protocol = AutodiscoverProtocol(config=Configuration(
        service_endpoint=url, credentials=credentials, auth_type=auth_type, retry_policy=INITIAL_RETRY_POLICY
    ))
    r = _get_response(protocol=autodiscover_protocol, email=email)
    domain = get_domain(email)
    try:
        ad_response = _parse_response(r.content)
    except (ErrorNonExistentMailbox, AutoDiscoverRedirect):
        # These are both valid responses from an autodiscover server, showing that we have found the correct
        # server for the original domain. Fill cache before re-raising
        log.debug('Adding cache entry for %s (hostname %s)', domain, hostname)
        # We have already acquired the cache lock at this point
        autodiscover_cache[(domain, credentials)] = autodiscover_protocol
        raise

    # Cache the final hostname of the autodiscover service so we don't need to autodiscover the same domain again
    log.debug('Adding cache entry for %s (hostname %s, has_ssl %s)', domain, hostname, has_ssl)
    # We have already acquired the cache lock at this point
    autodiscover_cache[(domain, credentials)] = autodiscover_protocol
    # Autodiscover response contains an auth type, but we don't want to spend time here testing if it actually works.
    # Instead of forcing a possibly-wrong auth type, just let Protocol auto-detect the auth type.
    ews_url = ad_response.protocol.ews_url
    if not ad_response.autodiscover_smtp_address:
        # Autodiscover does not always return an email address. In that case, the requesting email should be used
        ad_response.user.autodiscover_smtp_address = email

    return ad_response, Protocol(config=Configuration(
        service_endpoint=ews_url, credentials=credentials, auth_type=None, retry_policy=retry_policy
    ))


def _autodiscover_quick(credentials, email, protocol):
    r = _get_response(protocol=protocol, email=email)
    ad_response = _parse_response(r.content)
    ews_url = ad_response.protocol.ews_url
    log.debug('Autodiscover success: %s may connect to %s', email, ews_url)
    # Autodiscover response contains an auth type, but we don't want to spend time here testing if it actually works.
    # Instead of forcing a possibly-wrong auth type, just let Protocol auto-detect the auth type.
    if not ad_response.autodiscover_smtp_address:
        # Autodiscover does not always return an email address. In that case, the requesting email should be used
        ad_response.user.autodiscover_smtp_address = email
    return ad_response, Protocol(config=Configuration(
        service_endpoint=ews_url, credentials=credentials, auth_type=None, retry_policy=protocol.retry_policy
    ))


def _get_auth_type(url, email, retry_policy):
    try:
        data = Autodiscover.payload(email=email)
        return get_autodiscover_authtype(service_endpoint=url, retry_policy=retry_policy, data=data)
    except TransportError as e:
        if isinstance(e, RedirectError):
            raise
        raise AutoDiscoverFailed('Error guessing auth type: %s' % e) from None


def _get_response(protocol, email):
    data = Autodiscover.payload(email=email)
    try:
        # Rate-limiting is an issue with autodiscover if the same setup is hosting EWS and autodiscover and we just
        # hammered the server with requests. We allow redirects since some autodiscover servers will issue different
        # redirects depending on the POST data content.
        session = protocol.get_session()
        r, session = post_ratelimited(protocol=protocol, session=session, url=protocol.service_endpoint,
                                      headers=DEFAULT_HEADERS.copy(), data=data, allow_redirects=True)
        protocol.release_session(session)
        log.debug('Response headers: %s', r.headers)
    except RedirectError:
        raise
    except (TransportError, UnauthorizedError):
        log.debug('No access to %s using %s', protocol.service_endpoint, protocol.auth_type)
        raise AutoDiscoverFailed('No access to %s using %s' % (protocol.service_endpoint, protocol.auth_type)) from None
    if not is_xml(r.content):
        # This is normal - e.g. a greedy webserver serving custom HTTP 404's as 200 OK
        log.debug('URL %s: This is not XML: %r', protocol.service_endpoint, r.content[:1000])
        raise AutoDiscoverFailed('URL %s: This is not XML: %r' % (protocol.service_endpoint, r.content[:1000]))
    return r


def _parse_response(bytes_content):
    try:
        ad = Autodiscover.from_bytes(bytes_content=bytes_content)
    except ValueError as e:
        raise AutoDiscoverFailed(str(e))
    if ad.response is None:
        ad.raise_errors()
    ad_response = ad.response
    if ad_response.redirect_address:
        # This is redirection to e.g. Office365
        raise AutoDiscoverRedirect(ad_response.redirect_address)
    try:
        ews_url = ad_response.protocol.ews_url
    except ValueError:
        raise AutoDiscoverFailed('No valid protocols in response: %s' % bytes_content)
    if not ews_url:
        raise ValueError("Required element 'EwsUrl' not found in response")
    log.debug('Primary SMTP: %s, EWS endpoint: %s', ad_response.autodiscover_smtp_address, ews_url)
    return ad_response


def _get_canonical_name(hostname):
    log.debug('Attempting to get canonical name for %s', hostname)
    resolver = dns.resolver.Resolver()
    resolver.timeout = AutodiscoverProtocol.TIMEOUT
    try:
        canonical_name = resolver.query(hostname).canonical_name.to_unicode().rstrip('.')
    except (dns.resolver.NXDOMAIN, dns.resolver.NoAnswer):
        log.debug('Nonexistent domain %s', hostname)
        return None
    if canonical_name != hostname:
        log.debug('%s has canonical name %s', hostname, canonical_name)
        return canonical_name
    return None


def _get_hostname_from_srv(hostname):
    # An SRV entry may contain e.g.:
    #   canonical name = mail.ucl.dk.
    #   service = 8 100 443 webmail.ucn.dk.
    # or throw dns.resolver.NoAnswer
    # The first three numbers in the service line are priority, weight, port
    log.debug('Attempting to get SRV record on %s', hostname)
    resolver = dns.resolver.Resolver()
    resolver.timeout = AutodiscoverProtocol.TIMEOUT
    try:
        answers = resolver.query(hostname, 'SRV')
        for rdata in answers:
            try:
                vals = rdata.to_text().strip().rstrip('.').split(' ')
                # pylint: disable=expression-not-assigned
                int(vals[0]), int(vals[1]), int(vals[2])  # Just to raise errors if these are not ints
                svr = vals[3]
                return svr
            except (ValueError, IndexError):
                raise AutoDiscoverFailed('Incompatible SRV record for %s (%s)' % (hostname, rdata.to_text())) from None
    except dns.resolver.NoNameservers:
        raise AutoDiscoverFailed('No name servers for %s' % hostname) from None
    except dns.resolver.NoAnswer:
        raise AutoDiscoverFailed('No SRV record for %s' % hostname) from None
    except dns.resolver.NXDOMAIN:
        raise AutoDiscoverFailed('Nonexistent domain %s' % hostname) from None


def get_autodiscover_authtype(service_endpoint, retry_policy, data):
    # Get auth type by tasting headers from the server. Only do POST requests. HEAD is too error prone, and some servers
    # are set up to redirect to OWA on all requests except POST to the autodiscover endpoint.
    #
    # 'service_endpoint' could be any random server at this point, so we need to take adequate precautions.
    log.debug('Requesting %s from %s', data, service_endpoint)
    retry = 0
    wait = 10  # seconds
    t_start = time.time()
    headers = DEFAULT_HEADERS.copy()
    while True:
        _back_off_if_needed(retry_policy.back_off_until)
        log.debug('Trying to get autodiscover auth type for %s', service_endpoint)
        with AutodiscoverProtocol.raw_session() as s:
            try:
                r = s.post(url=service_endpoint, headers=headers, data=data, allow_redirects=False,
                           timeout=AutodiscoverProtocol.TIMEOUT)
                break
            except TLS_ERRORS as e:
                # Don't retry on TLS errors. They will most likely be persistent.
                raise TransportError(str(e)) from e
            except CONNECTION_ERRORS as e:
                total_wait = time.time() - t_start
                r = DummyResponse(url=service_endpoint, headers={}, request_headers=headers)
                if _may_retry_on_error(response=r, retry_policy=retry_policy, wait=total_wait):
                    log.info("Connection error on URL %s (retry %s, error: %s). Cool down %s secs",
                             service_endpoint, retry, e, wait)
                    retry_policy.back_off(wait)
                    retry += 1
                    continue
                else:
                    raise TransportError(str(e)) from e
    if r.status_code in (301, 302):
        try:
            redirect_url = get_redirect_url(r, allow_relative=False)
        except RelativeRedirect:
            raise TransportError('Redirect to same host when trying to get auth method')
        raise RedirectError(url=redirect_url)
    if r.status_code not in (200, 401):
        raise TransportError('Unexpected response: %s %s' % (r.status_code, r.reason))
    return get_auth_method_from_response(response=r)
