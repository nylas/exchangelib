import io
from itertools import chain
import logging

import requests
import requests_mock

from exchangelib import FailFast, FaultTolerance
from exchangelib.errors import RelativeRedirect, TransportError, RateLimitError, RedirectError, UnauthorizedError,\
    CASError
import exchangelib.util
from exchangelib.util import chunkify, peek, get_redirect_url, get_domain, PrettyXmlHandler, to_xml, BOM_UTF8, \
    ParseError, post_ratelimited, safe_b64decode, CONNECTION_ERRORS

from .common import EWSTest, mock_post, mock_session_exception


class UtilTest(EWSTest):
    def test_chunkify(self):
        # Test tuple, list, set, range, map, chain and generator
        seq = [1, 2, 3, 4, 5]
        self.assertEqual(list(chunkify(seq, chunksize=2)), [[1, 2], [3, 4], [5]])

        seq = (1, 2, 3, 4, 6, 7, 9)
        self.assertEqual(list(chunkify(seq, chunksize=3)), [(1, 2, 3), (4, 6, 7), (9,)])

        seq = {1, 2, 3, 4, 5}
        self.assertEqual(list(chunkify(seq, chunksize=2)), [[1, 2], [3, 4], [5, ]])

        seq = range(5)
        self.assertEqual(list(chunkify(seq, chunksize=2)), [range(0, 2), range(2, 4), range(4, 5)])

        seq = map(int, range(5))
        self.assertEqual(list(chunkify(seq, chunksize=2)), [[0, 1], [2, 3], [4]])

        seq = chain(*[[i] for i in range(5)])
        self.assertEqual(list(chunkify(seq, chunksize=2)), [[0, 1], [2, 3], [4]])

        seq = (i for i in range(5))
        self.assertEqual(list(chunkify(seq, chunksize=2)), [[0, 1], [2, 3], [4]])

    def test_peek(self):
        # Test peeking into various sequence types

        # tuple
        is_empty, seq = peek(tuple())
        self.assertEqual((is_empty, list(seq)), (True, []))
        is_empty, seq = peek((1, 2, 3))
        self.assertEqual((is_empty, list(seq)), (False, [1, 2, 3]))

        # list
        is_empty, seq = peek([])
        self.assertEqual((is_empty, list(seq)), (True, []))
        is_empty, seq = peek([1, 2, 3])
        self.assertEqual((is_empty, list(seq)), (False, [1, 2, 3]))

        # set
        is_empty, seq = peek(set())
        self.assertEqual((is_empty, list(seq)), (True, []))
        is_empty, seq = peek({1, 2, 3})
        self.assertEqual((is_empty, list(seq)), (False, [1, 2, 3]))

        # range
        is_empty, seq = peek(range(0))
        self.assertEqual((is_empty, list(seq)), (True, []))
        is_empty, seq = peek(range(1, 4))
        self.assertEqual((is_empty, list(seq)), (False, [1, 2, 3]))

        # map
        is_empty, seq = peek(map(int, []))
        self.assertEqual((is_empty, list(seq)), (True, []))
        is_empty, seq = peek(map(int, [1, 2, 3]))
        self.assertEqual((is_empty, list(seq)), (False, [1, 2, 3]))

        # generator
        is_empty, seq = peek((i for i in []))
        self.assertEqual((is_empty, list(seq)), (True, []))
        is_empty, seq = peek((i for i in [1, 2, 3]))
        self.assertEqual((is_empty, list(seq)), (False, [1, 2, 3]))

    @requests_mock.mock()
    def test_get_redirect_url(self, m):
        m.get('https://httpbin.org/redirect-to', status_code=302, headers={'location': 'https://example.com/'})
        r = requests.get('https://httpbin.org/redirect-to?url=https://example.com/', allow_redirects=False)
        self.assertEqual(get_redirect_url(r), 'https://example.com/')

        m.get('https://httpbin.org/redirect-to', status_code=302, headers={'location': 'http://example.com/'})
        r = requests.get('https://httpbin.org/redirect-to?url=http://example.com/', allow_redirects=False)
        self.assertEqual(get_redirect_url(r), 'http://example.com/')

        m.get('https://httpbin.org/redirect-to', status_code=302, headers={'location': '/example'})
        r = requests.get('https://httpbin.org/redirect-to?url=/example', allow_redirects=False)
        self.assertEqual(get_redirect_url(r), 'https://httpbin.org/example')

        m.get('https://httpbin.org/redirect-to', status_code=302, headers={'location': 'https://example.com'})
        with self.assertRaises(RelativeRedirect):
            r = requests.get('https://httpbin.org/redirect-to?url=https://example.com', allow_redirects=False)
            get_redirect_url(r, require_relative=True)

        m.get('https://httpbin.org/redirect-to', status_code=302, headers={'location': '/example'})
        with self.assertRaises(RelativeRedirect):
            r = requests.get('https://httpbin.org/redirect-to?url=/example', allow_redirects=False)
            get_redirect_url(r, allow_relative=False)

    def test_to_xml(self):
        to_xml(b'<?xml version="1.0" encoding="UTF-8"?><foo></foo>')
        to_xml(BOM_UTF8+b'<?xml version="1.0" encoding="UTF-8"?><foo></foo>')
        to_xml(BOM_UTF8+b'<?xml version="1.0" encoding="UTF-8"?><foo>&broken</foo>')
        with self.assertRaises(ParseError):
            to_xml(b'foo')
        try:
            to_xml(b'<t:Foo><t:Bar>Baz</t:Bar></t:Foo>')
        except ParseError as e:
            # Not all lxml versions throw an error here, so we can't use assertRaises
            self.assertIn('Offending text: [...]<t:Foo><t:Bar>Baz</t[...]', e.args[0])

    def test_get_domain(self):
        self.assertEqual(get_domain('foo@example.com'), 'example.com')
        with self.assertRaises(ValueError):
            get_domain('blah')

    def test_pretty_xml_handler(self):
        # Test that a normal, non-XML log record is passed through unchanged
        stream = io.StringIO()
        stream.isatty = lambda: True
        h = PrettyXmlHandler(stream=stream)
        self.assertTrue(h.is_tty())
        r = logging.LogRecord(
            name='baz', level=logging.INFO, pathname='/foo/bar', lineno=1, msg='hello', args=(), exc_info=None
        )
        h.emit(r)
        h.stream.seek(0)
        self.assertEqual(h.stream.read(), 'hello\n')

        # Test formatting of an XML record. It should contain newlines and color codes.
        stream = io.StringIO()
        stream.isatty = lambda: True
        h = PrettyXmlHandler(stream=stream)
        r = logging.LogRecord(
            name='baz', level=logging.DEBUG, pathname='/foo/bar', lineno=1, msg='hello %(xml_foo)s',
            args=({'xml_foo': b'<?xml version="1.0" encoding="UTF-8"?><foo>bar</foo>'},), exc_info=None)
        h.emit(r)
        h.stream.seek(0)
        self.assertEqual(
            h.stream.read(),
            "hello \x1b[36m<?xml version='1.0' encoding='utf-8'?>\x1b[39;49;00m\n\x1b[94m"
            "<foo\x1b[39;49;00m\x1b[94m>\x1b[39;49;00mbar\x1b[94m</foo>\x1b[39;49;00m\n\n"
        )

    def test_post_ratelimited(self):
        url = 'https://example.com'

        protocol = self.account.protocol
        retry_policy = protocol.config.retry_policy
        RETRY_WAIT = exchangelib.util.RETRY_WAIT
        MAX_REDIRECTS = exchangelib.util.MAX_REDIRECTS

        session = protocol.get_session()
        try:
            # Make sure we fail fast in error cases
            protocol.config.retry_policy = FailFast()

            # Test the straight, HTTP 200 path
            session.post = mock_post(url, 200, {}, 'foo')
            r, session = post_ratelimited(protocol=protocol, session=session, url='http://', headers=None, data='')
            self.assertEqual(r.content, b'foo')

            # Test exceptions raises by the POST request
            for err_cls in CONNECTION_ERRORS:
                session.post = mock_session_exception(err_cls)
                with self.assertRaises(err_cls):
                    r, session = post_ratelimited(
                        protocol=protocol, session=session, url='http://', headers=None, data='')

            # Test bad exit codes and headers
            session.post = mock_post(url, 401, {})
            with self.assertRaises(UnauthorizedError):
                r, session = post_ratelimited(protocol=protocol, session=session, url='http://', headers=None, data='')
            session.post = mock_post(url, 999, {'connection': 'close'})
            with self.assertRaises(TransportError):
                r, session = post_ratelimited(protocol=protocol, session=session, url='http://', headers=None, data='')
            session.post = mock_post(url, 302,
                                     {'location': '/ews/genericerrorpage.htm?aspxerrorpath=/ews/exchange.asmx'})
            with self.assertRaises(TransportError):
                r, session = post_ratelimited(protocol=protocol, session=session, url='http://', headers=None, data='')
            session.post = mock_post(url, 503, {})
            with self.assertRaises(TransportError):
                r, session = post_ratelimited(protocol=protocol, session=session, url='http://', headers=None, data='')

            # No redirect header
            session.post = mock_post(url, 302, {})
            with self.assertRaises(TransportError):
                r, session = post_ratelimited(protocol=protocol, session=session, url=url, headers=None, data='')
            # Redirect header to same location
            session.post = mock_post(url, 302, {'location': url})
            with self.assertRaises(TransportError):
                r, session = post_ratelimited(protocol=protocol, session=session, url=url, headers=None, data='')
            # Redirect header to relative location
            session.post = mock_post(url, 302, {'location': url + '/foo'})
            with self.assertRaises(RedirectError):
                r, session = post_ratelimited(protocol=protocol, session=session, url=url, headers=None, data='')
            # Redirect header to other location and allow_redirects=False
            session.post = mock_post(url, 302, {'location': 'https://contoso.com'})
            with self.assertRaises(TransportError):
                r, session = post_ratelimited(protocol=protocol, session=session, url=url, headers=None, data='')
            # Redirect header to other location and allow_redirects=True
            exchangelib.util.MAX_REDIRECTS = 0
            session.post = mock_post(url, 302, {'location': 'https://contoso.com'})
            with self.assertRaises(TransportError):
                r, session = post_ratelimited(protocol=protocol, session=session, url=url, headers=None, data='',
                                              allow_redirects=True)

            # CAS error
            session.post = mock_post(url, 999, {'X-CasErrorCode': 'AAARGH!'})
            with self.assertRaises(CASError):
                r, session = post_ratelimited(protocol=protocol, session=session, url=url, headers=None, data='')

            # Allow XML data in a non-HTTP 200 response
            session.post = mock_post(url, 500, {}, '<?xml version="1.0" ?><foo></foo>')
            r, session = post_ratelimited(protocol=protocol, session=session, url=url, headers=None, data='')
            self.assertEqual(r.content, b'<?xml version="1.0" ?><foo></foo>')

            # Bad status_code and bad text
            session.post = mock_post(url, 999, {})
            with self.assertRaises(TransportError):
                r, session = post_ratelimited(protocol=protocol, session=session, url=url, headers=None, data='')

            # Test rate limit exceeded
            exchangelib.util.RETRY_WAIT = 1
            protocol.config.retry_policy = FaultTolerance(max_wait=0.5)  # Fail after first RETRY_WAIT
            session.post = mock_post(url, 503, {'connection': 'close'})
            # Mock renew_session to return the same session so the session object's 'post' method is still mocked
            protocol.renew_session = lambda s: s
            with self.assertRaises(RateLimitError) as rle:
                r, session = post_ratelimited(protocol=protocol, session=session, url='http://', headers=None, data='')
            self.assertEqual(rle.exception.status_code, 503)
            self.assertEqual(rle.exception.url, url)
            self.assertTrue(1 <= rle.exception.total_wait < 2)  # One RETRY_WAIT plus some overhead

            # Test something larger than the default wait, so we retry at least once
            protocol.retry_policy.max_wait = 3  # Fail after second RETRY_WAIT
            session.post = mock_post(url, 503, {'connection': 'close'})
            with self.assertRaises(RateLimitError) as rle:
                r, session = post_ratelimited(protocol=protocol, session=session, url='http://', headers=None, data='')
            self.assertEqual(rle.exception.status_code, 503)
            self.assertEqual(rle.exception.url, url)
            # We double the wait for each retry, so this is RETRY_WAIT + 2*RETRY_WAIT plus some overhead
            self.assertTrue(3 <= rle.exception.total_wait < 4, rle.exception.total_wait)
        finally:
            protocol.retire_session(session)  # We have patched the session, so discard it
            # Restore patched attributes and functions
            protocol.config.retry_policy = retry_policy
            exchangelib.util.RETRY_WAIT = RETRY_WAIT
            exchangelib.util.MAX_REDIRECTS = MAX_REDIRECTS

            try:
                delattr(protocol, 'renew_session')
            except AttributeError:
                pass

    def test_safe_b64decode(self):
        # Test correctly padded string
        self.assertEqual(safe_b64decode('SGVsbG8gd29ybGQ='), b'Hello world')
        # Test incorrectly padded string
        self.assertEqual(safe_b64decode('SGVsbG8gd29ybGQ'), b'Hello world')
        # Test binary data
        self.assertEqual(safe_b64decode(b'SGVsbG8gd29ybGQ='), b'Hello world')
        # Test incorrectly padded binary data
        self.assertEqual(safe_b64decode(b'SGVsbG8gd29ybGQ'), b'Hello world')
