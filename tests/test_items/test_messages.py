from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time

from exchangelib.folders import Inbox
from exchangelib.items import Message
from exchangelib.queryset import DoesNotExist

from ..common import get_random_string
from .test_basics import CommonItemTest


class MessagesTest(CommonItemTest):
    # Just test one of the Message-type folders
    TEST_FOLDER = 'inbox'
    FOLDER_CLASS = Inbox
    ITEM_CLASS = Message
    INCOMING_MESSAGE_TIMEOUT = 20

    def get_incoming_message(self, subject):
        t1 = time.monotonic()
        while True:
            t2 = time.monotonic()
            if t2 - t1 > self.INCOMING_MESSAGE_TIMEOUT:
                raise self.skipTest('Too bad. Gave up in %s waiting for the incoming message to show up' % self.id())
            try:
                return self.account.inbox.get(subject=subject)
            except DoesNotExist:
                time.sleep(5)

    def test_send(self):
        # Test that we can send (only) Message items
        item = self.get_test_item()
        item.folder = None
        item.send()
        self.assertIsNone(item.id)
        self.assertIsNone(item.changekey)
        self.assertEqual(self.test_folder.filter(categories__contains=item.categories).count(), 0)

    def test_send_and_save(self):
        # Test that we can send_and_save Message items
        item = self.get_test_item()
        item.send_and_save()
        self.assertIsNone(item.id)
        self.assertIsNone(item.changekey)
        time.sleep(5)  # Requests are supposed to be transactional, but apparently not...
        # Also, the sent item may be followed by an automatic message with the same category
        self.assertGreaterEqual(self.test_folder.filter(categories__contains=item.categories).count(), 1)

        # Test update, although it makes little sense
        item = self.get_test_item()
        item.save()
        item.send_and_save()
        time.sleep(5)  # Requests are supposed to be transactional, but apparently not...
        # Also, the sent item may be followed by an automatic message with the same category
        self.assertGreaterEqual(self.test_folder.filter(categories__contains=item.categories).count(), 1)

    def test_send_draft(self):
        item = self.get_test_item()
        item.folder = self.account.drafts
        item.is_draft = True
        item.save()  # Save a draft
        item.send()  # Send the draft
        self.assertIsNone(item.id)
        self.assertIsNone(item.changekey)
        self.assertEqual(item.folder, self.account.sent)
        self.assertEqual(self.test_folder.filter(categories__contains=item.categories).count(), 0)

    def test_send_and_copy_to_folder(self):
        item = self.get_test_item()
        item.send(save_copy=True, copy_to_folder=self.account.sent)  # Send the draft and save to the sent folder
        self.assertIsNone(item.id)
        self.assertIsNone(item.changekey)
        self.assertEqual(item.folder, self.account.sent)
        time.sleep(5)  # Requests are supposed to be transactional, but apparently not...
        self.assertEqual(self.account.sent.filter(categories__contains=item.categories).count(), 1)

    def test_bulk_send(self):
        with self.assertRaises(AttributeError):
            self.account.bulk_send(ids=[], save_copy=False, copy_to_folder=self.account.trash)
        item = self.get_test_item()
        item.save()
        for res in self.account.bulk_send(ids=[item]):
            self.assertEqual(res, True)
        time.sleep(10)  # Requests are supposed to be transactional, but apparently not...
        # By default, sent items are placed in the sent folder
        self.assertEqual(self.account.sent.filter(categories__contains=item.categories).count(), 1)

    def test_reply(self):
        # Test that we can reply to a Message item. EWS only allows items that have been sent to receive a reply
        item = self.get_test_item()
        item.folder = None
        item.send()  # get_test_item() sets the to_recipients to the test account
        sent_item = self.get_incoming_message(item.subject)
        new_subject = ('Re: %s' % sent_item.subject)[:255]
        sent_item.reply(subject=new_subject, body='Hello reply', to_recipients=[item.author])
        reply = self.get_incoming_message(new_subject)
        self.account.bulk_delete([sent_item, reply])

    def test_reply_all(self):
        # Test that we can reply-all a Message item. EWS only allows items that have been sent to receive a reply
        item = self.get_test_item(folder=None)
        item.folder = None
        item.send()
        sent_item = self.get_incoming_message(item.subject)
        new_subject = ('Re: %s' % sent_item.subject)[:255]
        sent_item.reply_all(subject=new_subject, body='Hello reply')
        reply = self.get_incoming_message(new_subject)
        self.account.bulk_delete([sent_item, reply])

    def test_forward(self):
        # Test that we can forward a Message item. EWS only allows items that have been sent to receive a reply
        item = self.get_test_item(folder=None)
        item.folder = None
        item.send()
        sent_item = self.get_incoming_message(item.subject)
        new_subject = ('Re: %s' % sent_item.subject)[:255]
        sent_item.forward(subject=new_subject, body='Hello reply', to_recipients=[item.author])
        reply = self.get_incoming_message(new_subject)
        forward = sent_item.create_forward(subject=new_subject, body='Hello reply', to_recipients=[item.author])
        res = forward.save(self.account.drafts)
        self.account.bulk_delete([sent_item, reply, res])

    def test_mime_content(self):
        # Tests the 'mime_content' field
        subject = get_random_string(16)
        msg = MIMEMultipart()
        msg['From'] = self.account.primary_smtp_address
        msg['To'] = self.account.primary_smtp_address
        msg['Subject'] = subject
        body = 'MIME test mail'
        msg.attach(MIMEText(body, 'plain', _charset='utf-8'))
        mime_content = msg.as_bytes()
        self.ITEM_CLASS(
            folder=self.test_folder,
            to_recipients=[self.account.primary_smtp_address],
            mime_content=mime_content,
            categories=self.categories,
        ).save()
        self.assertEqual(self.test_folder.get(subject=subject).body, body)
