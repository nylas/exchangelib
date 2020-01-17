import datetime
import math
import time

import requests_mock

from exchangelib import Configuration, Credentials, NTLM, FailFast, FaultTolerance, Version, Build
from exchangelib.transport import AUTH_TYPE_MAP

from .common import TimedTestCase


class ConfigurationTest(TimedTestCase):
    def test_init(self):
        with self.assertRaises(ValueError) as e:
            Configuration(credentials='foo')
        self.assertEqual(e.exception.args[0], "'credentials' 'foo' must be a Credentials instance")
        with self.assertRaises(AttributeError) as e:
            Configuration(server='foo', service_endpoint='bar')
        self.assertEqual(e.exception.args[0], "Only one of 'server' or 'service_endpoint' must be provided")
        with self.assertRaises(ValueError) as e:
            Configuration(auth_type='foo')
        self.assertEqual(
            e.exception.args[0],
            "'auth_type' 'foo' must be one of %s" % ', '.join("'%s'" % k for k in sorted(AUTH_TYPE_MAP.keys()))
        )
        with self.assertRaises(ValueError) as e:
            Configuration(version='foo')
        self.assertEqual(e.exception.args[0], "'version' 'foo' must be a Version instance")
        with self.assertRaises(ValueError) as e:
            Configuration(retry_policy='foo')
        self.assertEqual(e.exception.args[0], "'retry_policy' 'foo' must be a RetryPolicy instance")

    def test_magic(self):
        config = Configuration(
            server='example.com',
            credentials=Credentials('foo', 'bar'),
            auth_type=NTLM,
            version=Version(build=Build(15, 1, 2, 3), api_version='foo'),
        )
        # Just test that these work
        str(config)
        repr(config)

    @requests_mock.mock()  # Just to make sure we don't make any requests
    def test_hardcode_all(self, m):
        # Test that we can hardcode everything without having a working server. This is useful if neither tasting or
        # guessing missing values works.
        Configuration(
            server='example.com',
            credentials=Credentials('foo', 'bar'),
            auth_type=NTLM,
            version=Version(build=Build(15, 1, 2, 3), api_version='foo'),
        )

    def test_fail_fast_back_off(self):
        # Test that FailFast does not support back-off logic
        c = FailFast()
        self.assertIsNone(c.back_off_until)
        with self.assertRaises(AttributeError):
            c.back_off_until = 1

    def test_service_account_back_off(self):
        # Test back-off logic in FaultTolerance
        sa = FaultTolerance()

        # Initially, the value is None
        self.assertIsNone(sa.back_off_until)

        # Test a non-expired back off value
        in_a_while = datetime.datetime.now() + datetime.timedelta(seconds=10)
        sa.back_off_until = in_a_while
        self.assertEqual(sa.back_off_until, in_a_while)

        # Test an expired back off value
        sa.back_off_until = datetime.datetime.now()
        time.sleep(0.001)
        self.assertIsNone(sa.back_off_until)

        # Test the back_off() helper
        sa.back_off(10)
        # This is not a precise test. Assuming fast computers, there should be less than 1 second between the two lines.
        self.assertEqual(int(math.ceil((sa.back_off_until - datetime.datetime.now()).total_seconds())), 10)

        # Test expiry
        sa.back_off(0)
        time.sleep(0.001)
        self.assertIsNone(sa.back_off_until)

        # Test default value
        sa.back_off(None)
        self.assertEqual(int(math.ceil((sa.back_off_until - datetime.datetime.now()).total_seconds())), 60)
