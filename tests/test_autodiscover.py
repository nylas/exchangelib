# coding=utf-8
from collections import namedtuple
import glob
import os
import tempfile
import warnings

import dns
import requests_mock

import exchangelib.autodiscover
from exchangelib import discover, Credentials, NTLM, FailFast, Configuration, Account
from exchangelib.autodiscover import close_connections, AutodiscoverProtocol
from exchangelib.errors import ErrorNonExistentMailbox, AutoDiscoverRedirect, AutoDiscoverCircularRedirect, \
    AutoDiscoverFailed
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter

from .common import EWSTest


class AutodiscoverTest(EWSTest):
    def test_magic(self):
        # Just test we don't fail
        from exchangelib.autodiscover import _autodiscover_cache
        discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)
        str(_autodiscover_cache)
        repr(_autodiscover_cache)
        for protocol in _autodiscover_cache._protocols.values():
            str(protocol)
            repr(protocol)

    def test_autodiscover(self):
        primary_smtp_address, protocol = discover(email=self.account.primary_smtp_address,
                                                  credentials=self.account.protocol.credentials)
        self.assertEqual(primary_smtp_address, self.account.primary_smtp_address)
        self.assertEqual(protocol.service_endpoint.lower(), self.account.protocol.service_endpoint.lower())
        self.assertEqual(protocol.version.build, self.account.protocol.version.build)

    def test_autodiscover_failure(self):
        # Empty the cache
        from exchangelib.autodiscover import _autodiscover_cache
        _autodiscover_cache.clear()
        with self.assertRaises(ErrorNonExistentMailbox):
            # Test that error is raised with an empty cache
            discover(email='XXX.' + self.account.primary_smtp_address, credentials=self.account.protocol.credentials)
        with self.assertRaises(ErrorNonExistentMailbox):
            # Test that error is raised with a full cache
            discover(email='XXX.' + self.account.primary_smtp_address, credentials=self.account.protocol.credentials)

    def test_close_autodiscover_connections(self):
        discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)
        close_connections()

    def test_autodiscover_gc(self):
        # This is what Python garbage collection does
        from exchangelib.autodiscover import _autodiscover_cache
        discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)
        del _autodiscover_cache

    def test_autodiscover_direct_gc(self):
        # This is what Python garbage collection does
        from exchangelib.autodiscover import _autodiscover_cache
        discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)
        _autodiscover_cache.__del__()

    @requests_mock.mock(real_http=True)
    def test_autodiscover_cache(self, m):
        # Empty the cache
        from exchangelib.autodiscover import _autodiscover_cache
        _autodiscover_cache.clear()
        cache_key = (self.account.domain, self.account.protocol.credentials)
        # Not cached
        self.assertNotIn(cache_key, _autodiscover_cache)
        discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)
        # Now it's cached
        self.assertIn(cache_key, _autodiscover_cache)
        # Make sure the cache can be looked by value, not by id(). This is important for multi-threading/processing
        self.assertIn((
            self.account.primary_smtp_address.split('@')[1],
            Credentials(self.account.protocol.credentials.username, self.account.protocol.credentials.password),
            True
        ), _autodiscover_cache)
        # Poison the cache. discover() must survive and rebuild the cache
        _autodiscover_cache[cache_key] = AutodiscoverProtocol(config=Configuration(
            service_endpoint='https://example.com/blackhole.asmx',
            credentials=Credentials('leet_user', 'cannaguess'),
            auth_type=NTLM,
            retry_policy=FailFast(),
        ))
        m.post('https://example.com/blackhole.asmx', status_code=404)
        discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)
        self.assertIn(cache_key, _autodiscover_cache)

        # Make sure that the cache is actually used on the second call to discover()
        _orig = exchangelib.autodiscover._try_autodiscover

        def _mock(*args, **kwargs):
            raise NotImplementedError()
        exchangelib.autodiscover._try_autodiscover = _mock
        discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)
        # Fake that another thread added the cache entry into the persistent storage but we don't have it in our
        # in-memory cache. The cache should work anyway.
        _autodiscover_cache._protocols.clear()
        discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)
        exchangelib.autodiscover._try_autodiscover = _orig
        # Make sure we can delete cache entries even though we don't have it in our in-memory cache
        _autodiscover_cache._protocols.clear()
        del _autodiscover_cache[cache_key]
        # This should also work if the cache does not contain the entry anymore
        del _autodiscover_cache[cache_key]

    def test_corrupt_autodiscover_cache(self):
        # Insert a fake Protocol instance into the cache
        from exchangelib.autodiscover import _autodiscover_cache
        key = (2, 'foo', 4)
        _autodiscover_cache[key] = namedtuple('P', ['service_endpoint', 'auth_type', 'retry_policy'])(1, 'bar', 'baz')
        # Check that it exists. 'in' goes directly to the file
        self.assertTrue(key in _autodiscover_cache)
        # Destroy the backing cache file(s)
        for db_file in glob.glob(_autodiscover_cache._storage_file + '*'):
            with open(db_file, 'w') as f:
                f.write('XXX')
        # Check that we can recover from a destroyed file and that the entry no longer exists
        self.assertFalse(key in _autodiscover_cache)

    def test_autodiscover_from_account(self):
        from exchangelib.autodiscover import _autodiscover_cache
        _autodiscover_cache.clear()
        account = Account(primary_smtp_address=self.account.primary_smtp_address,
                          credentials=self.account.protocol.credentials,
                          autodiscover=True, locale='da_DK')
        self.assertEqual(account.primary_smtp_address, self.account.primary_smtp_address)
        self.assertEqual(account.protocol.service_endpoint.lower(), self.account.protocol.service_endpoint.lower())
        self.assertEqual(account.protocol.version.build, self.account.protocol.version.build)
        # Make sure cache is full
        self.assertTrue((account.domain, self.account.protocol.credentials, True) in _autodiscover_cache)
        # Test that autodiscover works with a full cache
        account = Account(primary_smtp_address=self.account.primary_smtp_address,
                          credentials=self.account.protocol.credentials,
                          autodiscover=True, locale='da_DK')
        self.assertEqual(account.primary_smtp_address, self.account.primary_smtp_address)
        # Test cache manipulation
        key = (account.domain, self.account.protocol.credentials, True)
        self.assertTrue(key in _autodiscover_cache)
        del _autodiscover_cache[key]
        self.assertFalse(key in _autodiscover_cache)
        del _autodiscover_cache

    def test_autodiscover_redirect(self):
        # Prime the cache
        email, p = discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)
        _orig = exchangelib.autodiscover._autodiscover_quick

        # Test that we can get another address back than the address we're looking up
        def _mock1(*args, **kwargs):
            return 'john@example.com', p
        exchangelib.autodiscover._autodiscover_quick = _mock1
        test_email, p = discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)
        self.assertEqual(test_email, 'john@example.com')

        # Test that we can survive being asked to lookup with another address
        def _mock2(*args, **kwargs):
            email = kwargs['email']
            if email == 'xxxxxx@%s' % self.account.domain:
                raise ErrorNonExistentMailbox(email)
            raise AutoDiscoverRedirect(redirect_email='xxxxxx@'+self.account.domain)
        exchangelib.autodiscover._autodiscover_quick = _mock2
        with self.assertRaises(ErrorNonExistentMailbox):
            discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)

        # Test that we catch circular redirects
        def _mock3(*args, **kwargs):
            raise AutoDiscoverRedirect(redirect_email=self.account.primary_smtp_address)
        exchangelib.autodiscover._autodiscover_quick = _mock3
        with self.assertRaises(AutoDiscoverCircularRedirect):
            discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)
        exchangelib.autodiscover._autodiscover_quick = _orig

        # Test that we catch circular redirects when cache is empty. This is a different code path
        _orig = exchangelib.autodiscover._try_autodiscover
        def _mock4(*args, **kwargs):
            raise AutoDiscoverRedirect(redirect_email=self.account.primary_smtp_address)
        exchangelib.autodiscover._try_autodiscover = _mock4
        exchangelib.autodiscover._autodiscover_cache.clear()
        with self.assertRaises(AutoDiscoverCircularRedirect):
            discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)
        exchangelib.autodiscover._try_autodiscover = _orig

        # Test that we can survive being asked to lookup with another address, when cache is empty
        def _mock5(*args, **kwargs):
            email = kwargs['email']
            if email == 'xxxxxx@%s' % self.account.domain:
                raise ErrorNonExistentMailbox(email)
            raise AutoDiscoverRedirect(redirect_email='xxxxxx@'+self.account.domain)
        exchangelib.autodiscover._try_autodiscover = _mock5
        exchangelib.autodiscover._autodiscover_cache.clear()
        with self.assertRaises(ErrorNonExistentMailbox):
            discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)
        exchangelib.autodiscover._try_autodiscover = _orig

    def test_canonical_lookup(self):
        from exchangelib.autodiscover import _get_canonical_name
        self.assertEqual(_get_canonical_name('example.com'), None)
        self.assertEqual(_get_canonical_name('example.com.'), 'example.com')
        self.assertEqual(_get_canonical_name('example.XXXXX.'), None)

    def test_srv(self):
        from exchangelib.autodiscover import _get_hostname_from_srv
        with self.assertRaises(AutoDiscoverFailed):
            # Unknown doomain
            _get_hostname_from_srv('example.XXXXX.')
        with self.assertRaises(AutoDiscoverFailed):
            # No SRV record
            _get_hostname_from_srv('example.com.')
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
        self.assertEqual(_get_hostname_from_srv('example.com.'), 'example.com')

        class _Mock2:
            def query(self, hostname, cat):
                class A:
                    def to_text(self):
                        # Return malformed data
                        return 'XXXXXXX'
                return [A()]
        dns.resolver.Resolver = _Mock2
        # Test an invalid record
        with self.assertRaises(AutoDiscoverFailed):
            _get_hostname_from_srv('example.com.')
        dns.resolver.Resolver = _orig

    def test_parse_response(self):
        from exchangelib.autodiscover import _parse_response
        with self.assertRaises(AutoDiscoverFailed):
            _parse_response(b'XXX')  # Invalid response

        xml = b'''<?xml version="1.0" encoding="utf-8"?><foo>bar</foo>'''
        with self.assertRaises(AutoDiscoverFailed):
            _parse_response(xml)  # Invalid XML response

        # Redirection
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
        with self.assertRaises(AutoDiscoverRedirect) as e:
            _parse_response(xml)  # Redirect to primary email
        self.assertEqual(e.exception.redirect_email, 'foo@example.com')

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
        self.assertEqual(_parse_response(xml)[0], 'https://expr.example.com/EWS/Exchange.asmx')

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
        self.assertEqual(_parse_response(xml)[0], 'https://exch.example.com/EWS/Exchange.asmx')

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
        with self.assertRaises(AutoDiscoverFailed):
            _parse_response(xml)

    def test_disable_ssl_verification(self):
        if not self.verify_ssl:
            # We can only run this test if we haven't already disabled TLS
            raise self.skipTest('TLS verification already disabled')
        import exchangelib.autodiscover

        default_adapter_cls = BaseProtocol.HTTP_ADAPTER_CLS

        # A normal discover should succeed
        exchangelib.autodiscover._autodiscover_cache.clear()
        discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)

        # Smash TLS verification using an untrusted certificate
        with tempfile.NamedTemporaryFile() as f:
            f.write(b'''\
 -----BEGIN CERTIFICATE-----
MIIENzCCAx+gAwIBAgIJAOYfYfw7NCOcMA0GCSqGSIb3DQEBBQUAMIGxMQswCQYD
VQQGEwJVUzERMA8GA1UECAwITWFyeWxhbmQxFDASBgNVBAcMC0ZvcmVzdCBIaWxs
MScwJQYDVQQKDB5UaGUgQXBhY2hlIFNvZnR3YXJlIEZvdW5kYXRpb24xFjAUBgNV
BAsMDUFwYWNoZSBUaHJpZnQxEjAQBgNVBAMMCWxvY2FsaG9zdDEkMCIGCSqGSIb3
DQEJARYVZGV2QHRocmlmdC5hcGFjaGUub3JnMB4XDTE0MDQwNzE4NTgwMFoXDTIy
MDYyNDE4NTgwMFowgbExCzAJBgNVBAYTAlVTMREwDwYDVQQIDAhNYXJ5bGFuZDEU
MBIGA1UEBwwLRm9yZXN0IEhpbGwxJzAlBgNVBAoMHlRoZSBBcGFjaGUgU29mdHdh
cmUgRm91bmRhdGlvbjEWMBQGA1UECwwNQXBhY2hlIFRocmlmdDESMBAGA1UEAwwJ
bG9jYWxob3N0MSQwIgYJKoZIhvcNAQkBFhVkZXZAdGhyaWZ0LmFwYWNoZS5vcmcw
ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCqE9TE9wEXp5LRtLQVDSGQ
GV78+7ZtP/I/ZaJ6Q6ZGlfxDFvZjFF73seNhAvlKlYm/jflIHYLnNOCySN8I2Xw6
L9MbC+jvwkEKfQo4eDoxZnOZjNF5J1/lZtBeOowMkhhzBMH1Rds351/HjKNg6ZKg
2Cldd0j7HbDtEixOLgLbPRpBcaYrLrNMasf3Hal+x8/b8ue28x93HSQBGmZmMIUw
AinEu/fNP4lLGl/0kZb76TnyRpYSPYojtS6CnkH+QLYnsRREXJYwD1Xku62LipkX
wCkRTnZ5nUsDMX6FPKgjQFQCWDXG/N096+PRUQAChhrXsJ+gF3NqWtDmtrhVQF4n
AgMBAAGjUDBOMB0GA1UdDgQWBBQo8v0wzQPx3EEexJPGlxPK1PpgKjAfBgNVHSME
GDAWgBQo8v0wzQPx3EEexJPGlxPK1PpgKjAMBgNVHRMEBTADAQH/MA0GCSqGSIb3
DQEBBQUAA4IBAQBGFRiJslcX0aJkwZpzTwSUdgcfKbpvNEbCNtVohfQVTI4a/oN5
U+yqDZJg3vOaOuiAZqyHcIlZ8qyesCgRN314Tl4/JQ++CW8mKj1meTgo5YFxcZYm
T9vsI3C+Nzn84DINgI9mx6yktIt3QOKZRDpzyPkUzxsyJ8J427DaimDrjTR+fTwD
1Dh09xeeMnSa5zeV1HEDyJTqCXutLetwQ/IyfmMBhIx+nvB5f67pz/m+Dv6V0r3I
p4HCcdnDUDGJbfqtoqsAATQQWO+WWuswB6mOhDbvPTxhRpZq6AkgWqv4S+u3M2GO
r5p9FrBgavAw5bKO54C0oQKpN/5fta5l6Ws0
-----END CERTIFICATE-----''')
            try:
                os.environ['REQUESTS_CA_BUNDLE'] = f.name

                # Now discover should fail. TLS errors mean we exhaust all autodiscover attempts
                with self.assertRaises(AutoDiscoverFailed):
                    exchangelib.autodiscover._autodiscover_cache.clear()
                    discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)

                # Disable insecure TLS warnings
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    # Make sure we can survive TLS validation errors when using the custom adapter
                    exchangelib.autodiscover._autodiscover_cache.clear()
                    BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
                    discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)

                    # Test that the custom adapter also works when validation is OK again
                    del os.environ['REQUESTS_CA_BUNDLE']
                    exchangelib.autodiscover._autodiscover_cache.clear()
                    discover(email=self.account.primary_smtp_address, credentials=self.account.protocol.credentials)
            finally:
                # Reset environment
                os.environ.pop('REQUESTS_CA_BUNDLE', None)  # May already have been deleted
                exchangelib.autodiscover._autodiscover_cache.clear()
                BaseProtocol.HTTP_ADAPTER_CLS = default_adapter_cls
