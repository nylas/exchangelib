import flake8.defaults
import flake8.main.application

from exchangelib.errors import ErrorAccessDenied, ErrorFolderNotFound, ErrorItemNotFound, ErrorInvalidOperation, \
    ErrorNoPublicFolderReplicaAvailable
from exchangelib.properties import EWSElement

from .common import EWSTest, TimedTestCase


class StyleTest(TimedTestCase):
    def test_flake8(self):
        import exchangelib
        flake8.defaults.MAX_LINE_LENGTH = 120
        app = flake8.main.application.Application()
        app.run(exchangelib.__path__)
        # If this fails, look at stdout for actual error messages
        self.assertEqual(app.result_count, 0)


class CommonTest(EWSTest):
    def test_magic(self):
        self.assertIn(self.account.protocol.version.api_version, str(self.account.protocol))
        self.assertIn(self.account.protocol.credentials.username, str(self.account.protocol.credentials))
        self.assertIn(self.account.primary_smtp_address, str(self.account))
        self.assertIn(str(self.account.version.build.major_version), repr(self.account.version))
        for item in (
                self.account.protocol,
                self.account.version,
        ):
            with self.subTest(item=item):
                # Just test that these at least don't throw errors
                repr(item)
                str(item)
        for attr in (
                'admin_audit_logs',
                'archive_deleted_items',
                'archive_inbox',
                'archive_msg_folder_root',
                'archive_recoverable_items_deletions',
                'archive_recoverable_items_purges',
                'archive_recoverable_items_root',
                'archive_recoverable_items_versions',
                'archive_root',
                'calendar',
                'conflicts',
                'contacts',
                'conversation_history',
                'directory',
                'drafts',
                'favorites',
                'im_contact_list',
                'inbox',
                'journal',
                'junk',
                'local_failures',
                'msg_folder_root',
                'my_contacts',
                'notes',
                'outbox',
                'people_connect',
                'public_folders_root',
                'quick_contacts',
                'recipient_cache',
                'recoverable_items_deletions',
                'recoverable_items_purges',
                'recoverable_items_root',
                'recoverable_items_versions',
                'search_folders',
                'sent',
                'server_failures',
                'sync_issues',
                'tasks',
                'todo_search',
                'trash',
                'voice_mail',
        ):
            with self.subTest(attr=attr):
                # Test distinguished folder shortcuts. Some may raise ErrorAccessDenied
                try:
                    item = getattr(self.account, attr)
                except (ErrorAccessDenied, ErrorFolderNotFound, ErrorItemNotFound, ErrorInvalidOperation,
                        ErrorNoPublicFolderReplicaAvailable):
                    continue
                else:
                    repr(item)
                    str(item)
                    self.assertTrue(item.is_distinguished)

    def test_from_xml(self):
        # Test for all EWSElement classes that they handle None as input to from_xml()
        import exchangelib
        for mod in (exchangelib.attachments, exchangelib.extended_properties, exchangelib.indexed_properties,
                    exchangelib.folders, exchangelib.items, exchangelib.properties):
            for k, v in vars(mod).items():
                with self.subTest(k=k, v=v):
                    if type(v) != type:
                        continue
                    if not issubclass(v, EWSElement):
                        continue
                    # from_xml() does not support None input
                    with self.assertRaises(Exception):
                        v.from_xml(elem=None, account=None)
