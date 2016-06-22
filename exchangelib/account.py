from logging import getLogger

from .autodiscover import discover
from .configuration import Configuration
from .credentials import Credentials, DELEGATE, IMPERSONATION
from .errors import ErrorFolderNotFound, ErrorAccessDenied
from .folders import Root, Calendar, Inbox, Tasks, Contacts, SHALLOW, DEEP
from .protocol import Protocol
from .ewsdatetime import EWSDateTime, UTC

log = getLogger(__name__)


class Account:
    """
    Models an Exchange server user account. The primary key for accounts is their PrimarySMTPAddress
    """
    def __init__(self, primary_smtp_address, fullname=None, credentials=None, access_type=None, autodiscover=False,
                 config=None):
        if '@' not in primary_smtp_address:
            raise ValueError("primary_smtp_address '%s' is not an email address" % primary_smtp_address)
        self.primary_smtp_address = primary_smtp_address
        self.fullname = fullname
        if not (credentials or config):
            raise AttributeError('Either config or credentials must be supplied')
        self.credentials = credentials or config.credentials
        # Assume delegate access if individual credentials are provided. Else, assume service user with impersonation
        self.access_type = access_type or (DELEGATE if credentials else IMPERSONATION)
        assert self.access_type in (DELEGATE, IMPERSONATION)
        if autodiscover:
            self.primary_smtp_address, self.protocol = discover(email=self.primary_smtp_address,
                                                                credentials=self.credentials)
            if config:
                assert isinstance(config, Configuration)
        else:
            self.protocol = config.protocol
        # We may need to override the default server version on a per-account basis because Microsoft may report one
        # server version up-front but delegate account requests to an older backend server.
        self.version = self.protocol.version
        self.root = Root(self)

        assert isinstance(self.credentials, Credentials)
        assert isinstance(self.protocol, Protocol)
        log.debug('Added account: %s', self)

    @property
    def folders(self):
        if hasattr(self, '_folders'):
            return self._folders
        # 'Top of Information Store' is a folder available in some Exchange accounts. It only contains folders
        # owned by the account.
        self._folders = self.root.get_folders(depth=SHALLOW)  # Start by searching top-level folders.
        has_tois = False
        for folder in self._folders:
            if folder.name == 'Top of Information Store':
                has_tois = True
                self._folders = folder.get_folders(depth=SHALLOW)
                break
        if not has_tois:
            # We need to dig deeper. Get everything.
            self._folders = self.root.get_folders(depth=DEEP)
        return self._folders

    @property
    def calendar(self):
        if hasattr(self, '_calendar'):
            return self._calendar
        # If the account contains a shared calendar from a different user, that calendar will be in the folder list.
        # Attempt not to return one of those. An account may not always have a calendar called "Calendar", but a
        # Calendar folder with a localized name instead. Return that, if it's available.
        try:
            # Get the default calendar
            self._calendar = Calendar(self).get_folder()
        except ErrorAccessDenied:
            # Maybe we just don't have GetFolder access? Try FindItems instead
            self._calendar = Calendar(self)
            dt = UTC.localize(EWSDateTime(2000, 1, 1))
            self._calendar.find_items(start=dt, end=dt, categories=['DUMMY'])
        except ErrorFolderNotFound as e:
            # There's no folder called 'calendar'. Try to guess which calendar folder is the default. Exchange makes
            # this unnecessarily difficult.
            calendars = []
            for folder in self.folders:
                # Search for a calendar wth a localized name
                # TODO: Calendar.LOCALIZED_NAMES is most definitely neither complete nor authoritative
                if folder.folder_class != folder.CONTAINER_CLASS:
                    # This is a pseudo-folder
                    continue
                if folder.name.title() in Calendar.LOCALIZED_NAMES:
                    calendars.append(folder)
            if not calendars:
                # There was no calendar folder with a localized name. Use the distinguished folder instead.
                for folder in self.folders:
                    if folder.folder_class != folder.CONTAINER_CLASS and folder.is_distinguished:
                        calendars.append(folder)
            if not calendars:
                raise ErrorFolderNotFound('No useable calendar folders') from e
            assert len(calendars) == 1, 'Multiple calendars could be default: %s' % [str(f) for f in calendars]
            self._calendar = calendars[0]
        return self._calendar

    @property
    def inbox(self):
        if hasattr(self, '_inbox'):
            return self._inbox
        self._inbox = None
        for folder in self.folders[Inbox]:
            if folder.is_distinguished:
                self._inbox = folder
                break
        return self._inbox

    @property
    def tasks(self):
        if hasattr(self, '_tasks'):
            return self._tasks
        self._tasks = None
        for folder in self.folders[Tasks]:
            if folder.is_distinguished:
                self._tasks = folder
                break
        return self._tasks

    @property
    def contacts(self):
        if hasattr(self, '_contacts'):
            return self._contacts
        self._contacts = None
        for folder in self.folders[Contacts]:
            if folder.is_distinguished:
                self._contacts = folder
                break
        return self._contacts

    def get_domain(self):
        return self.primary_smtp_address.split('@')[1].lower().strip()

    def __str__(self):
        txt = '%s' % self.primary_smtp_address
        if self.fullname:
            txt += ' (%s)' % self.fullname
        return txt
