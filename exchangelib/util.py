import re
from xml.etree.ElementTree import Element
import logging
import time
from threading import get_ident
from datetime import datetime
from copy import deepcopy
import itertools
from types import GeneratorType

from .errors import TransportError, RateLimitError, RedirectError

log = logging.getLogger(__name__)
ElementType = type(Element('x'))  # Type is auto-generated inside cElementTree


# Some control characters are illegal in XML 1.0 (and XML 1.1). Some Exchange serveres will emit XML 1.0
# containing characters only allowed in XML 1.1
_illegal_xml_chars_RE = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1F\uD800-\uDFFF\uFFFE\uFFFF]')
BOM = '\xef\xbb\xbf'  # UTF-8 byte order mark which tails some Microsoft XML strings


def chunkify(iterable, chunksize):
    """
    Splits an iterable into chunks of size ``chunksize``. The last chunk may be smaller than ``chunksize``.
    """
    if hasattr(iterable, '__getitem__'):
        # lists, tuples
        for i in range(0, len(iterable), chunksize):
            yield iterable[i:i + chunksize]
    else:
        # generators, sets
        chunk = []
        for i in iterable:
            chunk.append(i)
            if len(chunk) == chunksize:
                yield chunk
                chunk = []
        if chunk:
            yield chunk


def peek(iterable):
    """
    Checks if an iterable is empty and returns status and the rewinded generator
    """
    if isinstance(iterable, GeneratorType):
        try:
            first = next(iterable)
        except StopIteration:
            return True, iterable
        return False, itertools.chain([first], iterable)
    else:
        return len(iterable) == 0, iterable



def xml_to_str(tree, encoding='utf-8'):
    from xml.etree.ElementTree import tostring
    # tostring returns bytecode unless encoding is 'unicode'. We ALWAYS want bytecode so we can convert to unicode
    if encoding == 'unicode':
        encoding = 'utf-8'
    return tostring(tree, encoding=encoding).decode(encoding)


def get_xml_attr(tree, name):
    elem = tree.find(name)
    if elem is not None and elem.text:  # Must compare with None, see XML docs
        return elem.text.strip() or None
    return None


def get_xml_attrs(tree, name):
    return [elem.text.strip() for elem in tree.findall(name) if elem.text is not None]


def set_xml_value(elem, value, version):
    from .ewsdatetime import EWSDateTime
    from .folders import EWSElement
    if isinstance(value, str):
        elem.text = safe_xml_value(value)
    elif isinstance(value, (tuple, list)):
        for v in value:
            if isinstance(v, EWSElement):
                assert version
                elem.append(v.to_xml(version))
            elif isinstance(v, ElementType):
                elem.append(v)
            elif isinstance(v, str):
                add_xml_child(elem, 't:String', v)
            else:
                raise AttributeError('Unsupported type %s for list value %s on elem %s' % (type(v), v, elem))
    elif isinstance(value, bool):
        elem.text = '1' if value else '0'
    elif isinstance(value, EWSDateTime):
        elem.text = value.ewsformat()
    elif isinstance(value, EWSElement):
        assert version
        elem.append(value.to_xml(version))
    elif isinstance(value, ElementType):
        elem.append(value)
    else:
        raise AttributeError('Unsupported type %s for value %s on elem %s' % (type(value), value, elem))
    return elem


def safe_xml_value(value, replacement='?'):
    return str(_illegal_xml_chars_RE.sub(replacement, value))


# Keeps a cache of Element objects to deepcopy
_deepcopy_cache = dict()


def create_element(name, **attrs):
    # copy.deepcopy() is an order of magnitude faster than creating a new Element() every time
    key = (name, tuple(attrs.items()))  # dict requires key to be immutable
    if name not in _deepcopy_cache:
        _deepcopy_cache[key] = Element(name, **attrs)
    return deepcopy(_deepcopy_cache[key])


def add_xml_child(tree, name, value):
    # We're calling add_xml_child many places where we don't have the version handy. Don't pass EWSElement or list of
    # EWSElement to this function!
    tree.append(set_xml_value(elem=create_element(name), value=value, version=None))


def to_xml(text, encoding):
    from xml.etree.ElementTree import fromstring, ParseError
    processed = text.lstrip(BOM).encode(encoding or 'utf-8')
    try:
        return fromstring(processed)
    except ParseError:
        from io import StringIO
        from lxml.etree import XMLParser, parse, tostring
        # Exchange servers may spit out the weirdest XML. lxml is pretty good at recovering from errors
        log.warning('Fallback to lxml processing of faulty XML')
        magical_parser = XMLParser(encoding=encoding or 'utf-8', recover=True)
        root = parse(StringIO(processed), magical_parser)
        try:
            return fromstring(tostring(root))
        except ParseError as e:
            line_no, col_no = e.lineno, e.offset
            try:
                offending_line = processed.splitlines()[line_no - 1]
            except IndexError:
                offending_line = ''
            offending_excerpt = offending_line[max(0, col_no - 20):col_no + 20].decode('ascii', 'ignore')
            raise ParseError('%s\nOffending text: [...]%s[...]' % (str(e), offending_excerpt)) from e


def is_xml(text):
    """
    Helper function. Lightweight test if response is an XML doc
    """
    return text.lstrip(BOM)[0:5] == '<?xml'


class DummyRequest:
    headers = {}


class DummyResponse:
    status_code = 401
    headers = {}
    text = ''
    request = DummyRequest()


def get_redirect_url(response, server=None, has_ssl=None):
    from urllib.parse import urlparse
    redirect_url = response.headers.get('location', None)
    if not redirect_url:
        raise TransportError('302 redirect but no location header')
    # At least some are kind enough to supply a new location
    url = urlparse(redirect_url)
    response_url = urlparse(response.url)
    if server is None:
        server = response_url.netloc
    if has_ssl is None:
        has_ssl = response_url.scheme == 'https'
    scheme = url.scheme or ('https' if has_ssl else 'http')
    has_ssl = scheme == 'https'
    if url.netloc:
        server = url.netloc
    server = server.lower()
    redirect_url = '%s://%s%s' % (scheme, server, url.path)
    if redirect_url == response.url:
        # And some are mean enough to redirect to the same location
        raise TransportError('Redirect to same location: %s' % redirect_url)
    return redirect_url, server, has_ssl


def post_ratelimited(protocol, session, url, headers, data, timeout=None, verify=True, allow_redirects=False):
    from socket import timeout as SocketTimeout
    from requests.exceptions import ConnectionError, ReadTimeout, ChunkedEncodingError
    # The contract on sessions here is to return the session that ends up being used, or retiring the session if we
    # intend to raise an exception. We give up on max_wait timeout, not number of retries
    r = None
    wait = 10  # seconds
    max_wait = 3600  # seconds
    redirects = 0
    max_redirects = 5  # We don't want to be sent into an endless redirect loop
    log_msg = '''\
Retry: %(i)s
Waited: %(wait)s
Timeout: %(timeout)s
Session: %(session_id)s
Thread: %(thread_id)s
Auth type: %(auth)s
URL: %(url)s
Verify: %(verify)s
Allow redirects: %(allow_redirects)s
Response time: %(response_time)s
Status code: %(status_code)s
Request headers: %(request_headers)s
Response headers: %(response_headers)s'''
    log_vals = dict(i=0, wait=0, timeout=timeout, session_id=session.session_id, thread_id=get_ident(),
                    auth=session.auth, url=url, response_time=None, status_code=None, request_headers=headers,
                    response_headers=None, verify=verify, allow_redirects=allow_redirects)
    try:
        while True:
            log.debug('Session %(session_id)s thread %(thread_id)s: retry %(i)s timeout %(timeout)s POST\'ing to '
                      '%(url)s after %(wait)s s wait', log_vals)
            d1 = datetime.now()
            try:
                r = session.post(url=url, headers=headers, data=data, allow_redirects=False, timeout=timeout,
                                 verify=verify)
            except (ChunkedEncodingError, ConnectionError, ConnectionResetError, ReadTimeout, SocketTimeout):
                log.debug(
                    'Session %(session_id)s thread %(thread_id)s: timeout or connection error POST\'ing to %(url)s',
                    log_vals)
                r = DummyResponse()
                r.request.headers = headers
                r.headers = {'DummyResponseHeader': None}
            d2 = datetime.now()
            log_vals['response_time'] = str(d2 - d1)
            log_vals['status_code'] = r.status_code
            log_vals['request_headers'] = r.request.headers
            log_vals['response_headers'] = r.headers
            log.debug(log_msg, log_vals)
            # log.debug('Request data: %s', data)
            # log.debug('Response data: %s', getattr(r, 'text'))
            # The genericerrorpage.htm/internalerror.asp is ridiculous behaviour for random outages. Redirect to
            # '/internalsite/internalerror.asp' or '/internalsite/initparams.aspx' is caused by e.g. SSL certificate
            # f*ckups on the Exchange server.
            if (r.status_code == 401) \
                    or (r.headers.get('connection') == 'close') \
                    or (r.status_code == 302 and r.headers.get('location').lower() ==
                        '/ews/genericerrorpage.htm?aspxerrorpath=/ews/exchange.asmx')\
                    or (r.status_code == 503):
                # Maybe stale session. Get brand new one. But wait a bit, since the server may be rate-limiting us.
                # This can be 302 redirect to error page, 401 authentication error or 503 service unavailable
                if r.status_code not in (302, 401, 503):
                    # Only retry if we didn't get a useful response
                    break
                log_vals['i'] += 1
                log_vals['wait'] = wait  # We set it to 0 initially
                if wait > max_wait:
                    # We lost patience. Session is cleaned up in outer loop
                    raise RateLimitError(
                        'Session %(session_id)s URL %(url)s: Max timeout reached' % log_vals)
                log.info("Session %(session_id)s thread %(thread_id)s: Connection error on URL %(url)s "
                         "(code %(status_code)s). Cool down %(wait)s secs", log_vals)
                time.sleep(wait)  # Increase delay for every retry
                wait *= 2
                session = protocol.renew_session(session)
                log_vals['wait'] = wait
                log_vals['session_id'] = session.session_id
                continue
            if r.status_code == 302:
                # If we get a normal 302 redirect, requests will issue a GET to that URL. We still want to POST
                redirect_url, server, has_ssl = get_redirect_url(response=r, server=protocol.server,
                                                                 has_ssl=protocol.has_ssl)
                if not allow_redirects:
                    raise TransportError('Redirect not allowed but we were redirected (%s -> %s)' % (url, redirect_url))
                if has_ssl != protocol.has_ssl or server != protocol.server:
                    log.debug("'allow_redirects' only supports relative redirects (%s -> %s)", url, redirect_url)
                    raise RedirectError(url=redirect_url, server=server, has_ssl=has_ssl)
                url = redirect_url
                log_vals['url'] = url
                log.debug('302 Redirected to %s', url)
                redirects += 1
                if redirects > max_redirects:
                    raise TransportError('Max redirect count exceeded')
                continue
            break
    except (RateLimitError, RedirectError) as e:
        log.warning(e.value)
        protocol.retire_session(session)
        raise
    except Exception as e:
        # Let higher layers handle this. Add data for better debugging.
        log_msg = '%(exc_cls)s: %(exc_msg)s\n' + log_msg
        log_vals['exc_cls'] = e.__class__.__name__
        log_vals['exc_msg'] = str(e)
        log_msg += '\nRequest data: %(data)s'
        log_vals['data'] = data
        log_msg += '\nResponse data: %(text)s'
        try:
            log_vals['text'] = r.text
        except (NameError, AttributeError):
            log_vals['text'] = ''
        log.error(log_msg, log_vals)
        protocol.retire_session(session)
        raise
    if r.status_code != 200:
        if r.text and is_xml(r.text):
            # Some genius at Microsoft thinks it's OK to send 500 error messages with valid SOAP response
            log.debug('Got status code %s but trying to parse content anyway', r.status_code)
        else:
            # This could be anything. Let higher layers handle this
            protocol.retire_session(session)
            log_msg += '\nRequest data: %(data)s'
            log_vals['data'] = data
            try:
                log_msg += '\nResponse data: %(text)s'
                log_vals['text'] = r.text
            except (NameError, AttributeError):
                pass
            raise TransportError('Unknown failure\n' + log_msg % log_vals)
    log.debug('Session %(session_id)s thread %(thread_id)s: Useful response from %(url)s', log_vals)
    return r, session
