from collections import namedtuple
import glob
from types import MethodType

import dns
import requests_mock

from exchangelib import DELEGATE
import exchangelib.autodiscover.discovery
from exchangelib import Credentials, NTLM, FailFast, Configuration, Account
from exchangelib.autodiscover import close_connections, clear_cache, autodiscover_cache, AutodiscoverProtocol, \
    Autodiscovery
from exchangelib.autodiscover.properties import Autodiscover
from exchangelib.errors import ErrorNonExistentMailbox, AutoDiscoverCircularRedirect, AutoDiscoverFailed
from exchangelib.protocol import FaultTolerance
from exchangelib.util import get_domain
from .common import EWSTest


class AutodiscoverTest(EWSTest):
    def setUp(self):
        super().setUp()

        # Enable retries, to make tests more robust
        Autodiscovery.INITIAL_RETRY_POLICY = FaultTolerance(max_wait=30)
        Autodiscovery.RETRY_WAIT = 5

        # Each test should start with a clean autodiscover cache
        clear_cache()

        # Some mocking helpers
        self.domain = get_domain(self.account.primary_smtp_address)
        self.dummy_ad_endpoint = 'https://%s/Autodiscover/Autodiscover.xml' % self.domain
        self.dummy_ews_endpoint = 'https://expr.example.com/EWS/Exchange.asmx'
        self.dummy_ad_response = b'''\
<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006">
    <Response xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a">
        <User>
            <AutoDiscoverSMTPAddress>%s</AutoDiscoverSMTPAddress>
        </User>
        <Account>
            <AccountType>email</AccountType>
            <Action>settings</Action>
            <Protocol>
                <Type>EXPR</Type>
                <EwsUrl>%s</EwsUrl>
            </Protocol>
        </Account>
    </Response>
</Autodiscover>''' % (self.account.primary_smtp_address.encode(), self.dummy_ews_endpoint.encode())
        self.dummy_ews_response = b'''\
<?xml version='1.0' encoding='utf-8'?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <h:ServerVersionInfo xmlns:h="http://schemas.microsoft.com/exchange/services/2006/types"
    MajorVersion="15" MinorVersion="1" MajorBuildNumber="1847" MinorBuildNumber="5" Version="V2017_07_11"/>
  </soap:Header>
  <soap:Body>
    <m:ResolveNamesResponse
    xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <m:ResponseMessages>
        <m:ResolveNamesResponseMessage ResponseClass="Success">
          <m:ResponseCode>NoError</m:ResponseCode>
          <m:ResolutionSet TotalItemsInView="0" IncludesLastItemInRange="true">
          </m:ResolutionSet>
        </m:ResolveNamesResponseMessage>
      </m:ResponseMessages>
    </m:ResolveNamesResponse>
  </soap:Body>
</soap:Envelope>
'''

    @requests_mock.mock(real_http=False)  # Just make sure we don't issue any real HTTP here
    def test_magic(self, m):
        # Just test we don't fail when calling repr() and str(). Insert a dummy cache entry for testing
        clear_cache()
        c = Credentials('leet_user', 'cannaguess')
        autodiscover_cache[('example.com', c)] = AutodiscoverProtocol(config=Configuration(
            service_endpoint='https://example.com/Autodiscover/Autodiscover.xml',
            credentials=c,
            auth_type=NTLM,
            retry_policy=FailFast(),
        ))
        self.assertEqual(len(autodiscover_cache), 1)
        str(autodiscover_cache)
        repr(autodiscover_cache)
        for protocol in autodiscover_cache._protocols.values():
            str(protocol)
            repr(protocol)

    def test_autodiscover_empty_cache(self):
        # A live test of the entire process with an empty cache
        clear_cache()
        ad_response, protocol = exchangelib.autodiscover.discovery.discover(
            email=self.account.primary_smtp_address,
            credentials=self.account.protocol.credentials,
            retry_policy=self.retry_policy,
        )
        self.assertEqual(ad_response.autodiscover_smtp_address, self.account.primary_smtp_address)
        self.assertEqual(protocol.service_endpoint.lower(), self.account.protocol.service_endpoint.lower())
        self.assertEqual(protocol.version.build, self.account.protocol.version.build)

    def test_autodiscover_failure(self):
        # A live test that errors can be raised. Here, we try to aútodiscover a non-existing email address
        if not self.settings.get('autodiscover_server'):
            self.skipTest("Skipping %s - no 'autodiscover_server' entry in settings.yml" % self.__class__.__name__)
        # Autodiscovery may take a long time. Prime the cache with the autodiscover server from the config file
        ad_endpoint = 'https://%s/Autodiscover/Autodiscover.xml' % self.settings['autodiscover_server']
        cache_key = (self.domain, self.account.protocol.credentials)
        autodiscover_cache[cache_key] = AutodiscoverProtocol(config=Configuration(
            service_endpoint=ad_endpoint,
            credentials=self.account.protocol.credentials,
            auth_type=NTLM,
            retry_policy=self.retry_policy,
        ))
        with self.assertRaises(ErrorNonExistentMailbox):
            exchangelib.autodiscover.discovery.discover(
                email='XXX.' + self.account.primary_smtp_address,
                credentials=self.account.protocol.credentials,
                retry_policy=self.retry_policy,
            )

    def test_failed_login_via_account(self):
        Autodiscovery.INITIAL_RETRY_POLICY = FaultTolerance(max_wait=10)
        clear_cache()
        with self.assertRaises(AutoDiscoverFailed):
            Account(
                primary_smtp_address=self.account.primary_smtp_address,
                access_type=DELEGATE,
                credentials=Credentials(self.account.protocol.credentials.username, 'WRONG_PASSWORD'),
                autodiscover=True,
                locale='da_DK',
            )

    @requests_mock.mock(real_http=False)  # Just make sure we don't issue any real HTTP here
    def test_close_autodiscover_connections(self, m):
        # A live test that we can close TCP connections
        clear_cache()
        c = Credentials('leet_user', 'cannaguess')
        autodiscover_cache[('example.com', c)] = AutodiscoverProtocol(config=Configuration(
            service_endpoint='https://example.com/Autodiscover/Autodiscover.xml',
            credentials=c,
            auth_type=NTLM,
            retry_policy=FailFast(),
        ))
        self.assertEqual(len(autodiscover_cache), 1)
        close_connections()

    @requests_mock.mock(real_http=False)  # Just make sure we don't issue any real HTTP here
    def test_autodiscover_direct_gc(self, m):
        # Test garbage collection of the autodiscover cache
        clear_cache()
        c = Credentials('leet_user', 'cannaguess')
        autodiscover_cache[('example.com', c)] = AutodiscoverProtocol(config=Configuration(
            service_endpoint='https://example.com/Autodiscover/Autodiscover.xml',
            credentials=c,
            auth_type=NTLM,
            retry_policy=FailFast(),
        ))
        self.assertEqual(len(autodiscover_cache), 1)
        autodiscover_cache.__del__()

    @requests_mock.mock(real_http=False)
    def test_autodiscover_cache(self, m):
        # Mock the default endpoint that we test in step 1 of autodiscovery
        m.post(self.dummy_ad_endpoint, status_code=200, content=self.dummy_ad_response)
        # Also mock the EWS URL. We try to guess its auth method as part of autodiscovery
        m.post(self.dummy_ews_endpoint, status_code=200)
        discovery = Autodiscovery(
            email=self.account.primary_smtp_address,
            credentials=self.account.protocol.credentials,
            retry_policy=self.retry_policy,
        )
        # Not cached
        self.assertNotIn(discovery._cache_key, autodiscover_cache)
        discovery.discover()
        # Now it's cached
        self.assertIn(discovery._cache_key, autodiscover_cache)
        # Make sure the cache can be looked by value, not by id(). This is important for multi-threading/processing
        self.assertIn((
            self.account.primary_smtp_address.split('@')[1],
            Credentials(self.account.protocol.credentials.username, self.account.protocol.credentials.password),
            True
        ), autodiscover_cache)
        # Poison the cache with a failing autodiscover endpoint. discover() must handle this and rebuild the cache
        autodiscover_cache[discovery._cache_key] = AutodiscoverProtocol(config=Configuration(
            service_endpoint='https://example.com/Autodiscover/Autodiscover.xml',
            credentials=Credentials('leet_user', 'cannaguess'),
            auth_type=NTLM,
            retry_policy=FailFast(),
        ))
        m.post('https://example.com/Autodiscover/Autodiscover.xml', status_code=404)
        discovery.discover()
        self.assertIn(discovery._cache_key, autodiscover_cache)

        # Make sure that the cache is actually used on the second call to discover()
        _orig = discovery._step_1

        def _mock(slf, *args, **kwargs):
            raise NotImplementedError()

        discovery._step_1 = MethodType(_mock, discovery)
        discovery.discover()

        # Fake that another thread added the cache entry into the persistent storage but we don't have it in our
        # in-memory cache. The cache should work anyway.
        autodiscover_cache._protocols.clear()
        discovery.discover()
        discovery._step_1 = _orig

        # Make sure we can delete cache entries even though we don't have it in our in-memory cache
        autodiscover_cache._protocols.clear()
        del autodiscover_cache[discovery._cache_key]
        # This should also work if the cache does not contain the entry anymore
        del autodiscover_cache[discovery._cache_key]

    @requests_mock.mock(real_http=False)  # Just make sure we don't issue any real HTTP here
    def test_corrupt_autodiscover_cache(self, m):
        # Insert a fake Protocol instance into the cache and test that we can recover
        key = (2, 'foo', 4)
        autodiscover_cache[key] = namedtuple('P', ['service_endpoint', 'auth_type', 'retry_policy'])(1, 'bar', 'baz')
        # Check that it exists. 'in' goes directly to the file
        self.assertTrue(key in autodiscover_cache)
        # Destroy the backing cache file(s)
        for db_file in glob.glob(autodiscover_cache._storage_file + '*'):
            with open(db_file, 'w') as f:
                f.write('XXX')
        # Check that we can recover from a destroyed file and that the entry no longer exists
        self.assertFalse(key in autodiscover_cache)

    @requests_mock.mock(real_http=False)  # Just make sure we don't issue any real HTTP here
    def test_autodiscover_from_account(self, m):
        # Test that autodiscovery via account creation works
        clear_cache()
        # Mock the default endpoint that we test in step 1 of autodiscovery
        m.post(self.dummy_ad_endpoint, status_code=200, content=self.dummy_ad_response)
        # Also mock the EWS URL. We try to guess its auth method as part of autodiscovery
        m.post(self.dummy_ews_endpoint, status_code=200, content=self.dummy_ews_response)
        self.assertEqual(len(autodiscover_cache), 0)
        account = Account(
            primary_smtp_address=self.account.primary_smtp_address,
            config=Configuration(
                credentials=self.account.protocol.credentials,
                retry_policy=self.retry_policy,
            ),
            autodiscover=True,
            locale='da_DK',
        )
        self.assertEqual(account.primary_smtp_address, self.account.primary_smtp_address)
        self.assertEqual(account.protocol.service_endpoint.lower(), self.dummy_ews_endpoint.lower())
        # Make sure cache is full
        self.assertEqual(len(autodiscover_cache), 1)
        self.assertTrue((account.domain, self.account.protocol.credentials, True) in autodiscover_cache)
        # Test that autodiscover works with a full cache
        account = Account(
            primary_smtp_address=self.account.primary_smtp_address,
            config=Configuration(
                credentials=self.account.protocol.credentials,
                retry_policy=self.retry_policy,
            ),
            autodiscover=True,
            locale='da_DK',
        )
        self.assertEqual(account.primary_smtp_address, self.account.primary_smtp_address)
        # Test cache manipulation
        key = (account.domain, self.account.protocol.credentials, True)
        self.assertTrue(key in autodiscover_cache)
        del autodiscover_cache[key]
        self.assertFalse(key in autodiscover_cache)

    @requests_mock.mock(real_http=False)
    def test_autodiscover_redirect(self, m):
        # Test various aspects of autodiscover redirection. Mock all HTTP responses because we can't force a live server
        # to send us into the correct code paths.
        clear_cache()
        # Mock the default endpoint that we test in step 1 of autodiscovery
        m.post(self.dummy_ad_endpoint, status_code=200, content=self.dummy_ad_response)
        # Also mock the EWS URL. We try to guess its auth method as part of autodiscovery
        m.post(self.dummy_ews_endpoint, status_code=200)
        discovery = Autodiscovery(
            email=self.account.primary_smtp_address,
            credentials=self.account.protocol.credentials,
            retry_policy=self.retry_policy,
        )
        discovery.discover()

        # Make sure we discover a different return address
        m.post(self.dummy_ad_endpoint, status_code=200, content=b'''\
<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006">
    <Response xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a">
        <User>
            <AutoDiscoverSMTPAddress>john@example.com</AutoDiscoverSMTPAddress>
        </User>
        <Account>
            <AccountType>email</AccountType>
            <Action>settings</Action>
            <Protocol>
                <Type>EXPR</Type>
                <EwsUrl>https://expr.example.com/EWS/Exchange.asmx</EwsUrl>
            </Protocol>
        </Account>
    </Response>
</Autodiscover>''')
        # Also mock the EWS URL. We try to guess its auth method as part of autodiscovery
        m.post('https://expr.example.com/EWS/Exchange.asmx', status_code=200)
        ad_response, p = discovery.discover()
        self.assertEqual(ad_response.autodiscover_smtp_address, 'john@example.com')

        # Make sure we discover an address redirect to the same domain. We have to mock the same URL with two different
        # responses. We do that with a response list.
        redirect_addr_content = b'''\
<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006">
    <Response xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a">
        <Account>
            <Action>redirectAddr</Action>
            <RedirectAddr>redirect_me@%s</RedirectAddr>
        </Account>
    </Response>
</Autodiscover>''' % self.domain.encode()
        settings_content = b'''\
<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006">
    <Response xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a">
        <User>
            <AutoDiscoverSMTPAddress>redirected@%s</AutoDiscoverSMTPAddress>
        </User>
        <Account>
            <AccountType>email</AccountType>
            <Action>settings</Action>
            <Protocol>
                <Type>EXPR</Type>
                <EwsUrl>https://redirected.%s/EWS/Exchange.asmx</EwsUrl>
            </Protocol>
        </Account>
    </Response>
</Autodiscover>''' % (self.domain.encode(), self.domain.encode())
        # Also mock the EWS URL. We try to guess its auth method as part of autodiscovery
        m.post('https://redirected.%s/EWS/Exchange.asmx' % self.domain, status_code=200)

        m.post(self.dummy_ad_endpoint, [
            dict(status_code=200, content=redirect_addr_content),
            dict(status_code=200, content=settings_content),
        ])
        ad_response, p = discovery.discover()
        self.assertEqual(ad_response.autodiscover_smtp_address, 'redirected@%s' % self.domain)
        self.assertEqual(ad_response.protocol.ews_url, 'https://redirected.%s/EWS/Exchange.asmx' % self.domain)

        # Test that we catch circular redirects on the same domain with a primed cache. Just mock the endpoint to
        # return the same redirect response on every request.
        self.assertEqual(len(autodiscover_cache), 1)
        m.post(self.dummy_ad_endpoint, status_code=200, content=b'''\
<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006">
    <Response xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a">
        <Account>
            <Action>redirectAddr</Action>
            <RedirectAddr>foo@%s</RedirectAddr>
        </Account>
    </Response>
</Autodiscover>''' % self.domain.encode())
        self.assertEqual(len(autodiscover_cache), 1)
        with self.assertRaises(AutoDiscoverCircularRedirect):
            discovery.discover()

        # Test that we also catch circular redirects when cache is empty
        clear_cache()
        self.assertEqual(len(autodiscover_cache), 0)
        with self.assertRaises(AutoDiscoverCircularRedirect):
            discovery.discover()

        # Test that we can handle being asked to redirect to an address on a different domain
        m.post(self.dummy_ad_endpoint, status_code=200, content=b'''\
<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006">
    <Response xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a">
        <Account>
            <Action>redirectAddr</Action>
            <RedirectAddr>john@example.com</RedirectAddr>
        </Account>
    </Response>
</Autodiscover>''')
        m.post('https://example.com/Autodiscover/Autodiscover.xml', status_code=200, content=b'''\
<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006">
    <Response xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a">
        <User>
            <AutoDiscoverSMTPAddress>john@redirected.example.com</AutoDiscoverSMTPAddress>
        </User>
        <Account>
            <AccountType>email</AccountType>
            <Action>settings</Action>
            <Protocol>
                <Type>EXPR</Type>
                <EwsUrl>https://redirected.example.com/EWS/Exchange.asmx</EwsUrl>
            </Protocol>
        </Account>
    </Response>
</Autodiscover>''')
        # Also mock the EWS URL. We try to guess its auth method as part of autodiscovery
        m.post('https://redirected.example.com/EWS/Exchange.asmx', status_code=200)
        ad_response, p = discovery.discover()
        self.assertEqual(ad_response.autodiscover_smtp_address, 'john@redirected.example.com')
        self.assertEqual(ad_response.protocol.ews_url, 'https://redirected.example.com/EWS/Exchange.asmx')

    def test_get_srv_records(self):
        from exchangelib.autodiscover.discovery import _get_srv_records, SrvRecord
        # Unknown domain
        self.assertEqual(_get_srv_records('example.XXXXX'), [])
        # No SRV record
        self.assertEqual(_get_srv_records('example.com'), [])
        # Finding a real server that has a correct SRV record is not easy. Mock it
        _orig = dns.resolver.Resolver

        class _Mock1:
            def query(self, hostname, cat):
                class A:
                    def to_text(self):
                        # Return a valid record
                        return '1 2 3 example.com.'
                return [A()]

        dns.resolver.Resolver = _Mock1
        # Test a valid record
        self.assertEqual(_get_srv_records('example.com.'), [SrvRecord(priority=1, weight=2, port=3, srv='example.com')])

        class _Mock2:
            def query(self, hostname, cat):
                class A:
                    def to_text(self):
                        # Return malformed data
                        return 'XXXXXXX'
                return [A()]

        dns.resolver.Resolver = _Mock2
        # Test an invalid record
        self.assertEqual(_get_srv_records('example.com'), [])
        dns.resolver.Resolver = _orig

    def test_select_srv_host(self):
        from exchangelib.autodiscover.discovery import _select_srv_host, SrvRecord
        with self.assertRaises(ValueError):
            # Empty list
            _select_srv_host([])
        with self.assertRaises(ValueError):
            # No records with TLS port
            _select_srv_host([SrvRecord(priority=1, weight=2, port=3, srv='example.com')])
        # One record
        self.assertEqual(
            _select_srv_host([SrvRecord(priority=1, weight=2, port=443, srv='example.com')]),
            'example.com'
        )
        # Highest priority record
        self.assertEqual(
            _select_srv_host([
                SrvRecord(priority=10, weight=2, port=443, srv='10.example.com'),
                SrvRecord(priority=1, weight=2, port=443, srv='1.example.com'),
            ]),
            '10.example.com'
        )
        # Highest priority record no matter how it's sorted
        self.assertEqual(
            _select_srv_host([
                SrvRecord(priority=1, weight=2, port=443, srv='1.example.com'),
                SrvRecord(priority=10, weight=2, port=443, srv='10.example.com'),
            ]),
            '10.example.com'
        )

    def test_parse_response(self):
        # Test parsing of various XML responses
        with self.assertRaises(ValueError):
            Autodiscover.from_bytes(b'XXX')  # Invalid response

        xml = b'''<?xml version="1.0" encoding="utf-8"?><foo>bar</foo>'''
        with self.assertRaises(ValueError):
            Autodiscover.from_bytes(xml)  # Invalid XML response

        # Redirect to different email address
        xml = b'''\
<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006">
    <Response xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a">
        <User>
            <AutoDiscoverSMTPAddress>john@demo.affect-it.dk</AutoDiscoverSMTPAddress>
        </User>
        <Account>
            <Action>redirectAddr</Action>
            <RedirectAddr>foo@example.com</RedirectAddr>
        </Account>
    </Response>
</Autodiscover>'''
        self.assertEqual(Autodiscover.from_bytes(xml).response.redirect_address, 'foo@example.com')

        # Redirect to different URL
        xml = b'''\
<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006">
    <Response xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a">
        <User>
            <AutoDiscoverSMTPAddress>john@demo.affect-it.dk</AutoDiscoverSMTPAddress>
        </User>
        <Account>
            <Action>redirectUrl</Action>
            <RedirectURL>https://example.com/foo.asmx</RedirectURL>
        </Account>
    </Response>
</Autodiscover>'''
        self.assertEqual(Autodiscover.from_bytes(xml).response.redirect_url, 'https://example.com/foo.asmx')

        # Select EXPR if it's there, and there are multiple available
        xml = b'''\
<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006">
    <Response xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a">
        <User>
            <AutoDiscoverSMTPAddress>john@demo.affect-it.dk</AutoDiscoverSMTPAddress>
        </User>
        <Account>
            <AccountType>email</AccountType>
            <Action>settings</Action>
            <Protocol>
                <Type>EXCH</Type>
                <EwsUrl>https://exch.example.com/EWS/Exchange.asmx</EwsUrl>
            </Protocol>
            <Protocol>
                <Type>EXPR</Type>
                <EwsUrl>https://expr.example.com/EWS/Exchange.asmx</EwsUrl>
            </Protocol>
        </Account>
    </Response>
</Autodiscover>'''
        self.assertEqual(
            Autodiscover.from_bytes(xml).response.protocol.ews_url,
            'https://expr.example.com/EWS/Exchange.asmx'
        )

        # Select EXPR if EXPR is unavailable
        xml = b'''\
<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006">
    <Response xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a">
        <User>
            <AutoDiscoverSMTPAddress>john@demo.affect-it.dk</AutoDiscoverSMTPAddress>
        </User>
        <Account>
            <AccountType>email</AccountType>
            <Action>settings</Action>
            <Protocol>
                <Type>EXCH</Type>
                <EwsUrl>https://exch.example.com/EWS/Exchange.asmx</EwsUrl>
            </Protocol>
        </Account>
    </Response>
</Autodiscover>'''
        self.assertEqual(
            Autodiscover.from_bytes(xml).response.protocol.ews_url,
            'https://exch.example.com/EWS/Exchange.asmx'
        )

        # Fail if neither EXPR nor EXPR are unavailable
        xml = b'''\
<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006">
    <Response xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a">
        <User>
            <AutoDiscoverSMTPAddress>john@demo.affect-it.dk</AutoDiscoverSMTPAddress>
        </User>
        <Account>
            <AccountType>email</AccountType>
            <Action>settings</Action>
            <Protocol>
                <Type>XXX</Type>
                <EwsUrl>https://xxx.example.com/EWS/Exchange.asmx</EwsUrl>
            </Protocol>
        </Account>
    </Response>
</Autodiscover>'''
        with self.assertRaises(ValueError):
            Autodiscover.from_bytes(xml).response.protocol
