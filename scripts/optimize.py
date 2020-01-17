#!/usr/bin/env python

# Measures bulk create and delete performance for different session pool sizes and payload chunksizes
import copy
import logging
import os
import time

from yaml import safe_load

from exchangelib import DELEGATE, Configuration, Account, EWSDateTime, EWSTimeZone, CalendarItem, Credentials, \
    FaultTolerance

logging.basicConfig(level=logging.WARNING)

try:
    with open(os.path.join(os.path.dirname(__file__), '../settings.yml')) as f:
        settings = safe_load(f)
except FileNotFoundError:
    print('Copy settings.yml.sample to settings.yml and enter values for your test server')
    raise

categories = ['perftest']
tz = EWSTimeZone.timezone('America/New_York')

verify_ssl = settings.get('verify_ssl', True)
if not verify_ssl:
    from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
    BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

config = Configuration(
    server=settings['server'],
    credentials=Credentials(settings['username'], settings['password']),
    retry_policy=FaultTolerance(),
)
print('Exchange server: %s' % config.service_endpoint)

account = Account(config=config, primary_smtp_address=settings['account'], access_type=DELEGATE)

# Remove leftovers from earlier tests
account.calendar.filter(categories__contains=categories).delete()


# Calendar item generator
def generate_items(count):
    start = tz.localize(EWSDateTime(2000, 3, 1, 8, 30, 0))
    end = tz.localize(EWSDateTime(2000, 3, 1, 9, 15, 0))
    tpl_item = CalendarItem(
        start=start,
        end=end,
        body='This is a performance optimization test of server %s intended to find the optimal batch size and '
             'concurrent connection pool size of this server.' % account.protocol.server,
        location="It's safe to delete this",
        categories=categories,
    )
    for j in range(count):
        item = copy.copy(tpl_item)
        item.subject = 'Performance optimization test %s by exchangelib' % j,
        yield item


# Worker
def test(items, chunk_size):
    t1 = time.monotonic()
    ids = account.calendar.bulk_create(items=items, chunk_size=chunk_size)
    t2 = time.monotonic()
    account.bulk_delete(ids=ids, chunk_size=chunk_size)
    t3 = time.monotonic()

    delta1 = t2 - t1
    rate1 = len(ids) / delta1
    delta2 = t3 - t2
    rate2 = len(ids) / delta2
    print(('Time to process %s items (batchsize %s, poolsize %s): %s / %s (%s / %s per sec)' % (
        len(ids), chunk_size, account.protocol.poolsize, delta1, delta2, rate1, rate2)))


# Generate items
calitems = list(generate_items(500))

print('\nTesting batch size')
for i in range(1, 11):
    chunk_size = 25 * i
    account.protocol.poolsize = 5
    test(calitems, chunk_size)
    time.sleep(60)  # Sleep 1 minute. Performance will deteriorate over time if we give the server tie to recover

print('\nTesting pool size')
for i in range(1, 11):
    chunk_size = 10
    account.protocol.poolsize = i
    test(calitems, chunk_size)
    time.sleep(60)
