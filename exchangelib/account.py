from logging import getLogger

from .autodiscover import discover
from .credentials import DELEGATE, IMPERSONATION
from .errors import ErrorFolderNotFound, ErrorAccessDenied
from .folders import Root, Calendar, Messages, Tasks, Contacts, SHALLOW, DEEP, WELLKNOWN_FOLDERS
from .protocol import Protocol
from .util import get_domain

log = getLogger(__name__)


class Account:
    """
    Models an Exchange server user account. The primary key for an account is its PrimarySMTPAddress
    """
    def __init__(self, primary_smtp_address, fullname=None, access_type=None, autodiscover=False, credentials=None,
                 config=None, verify_ssl=True, locale='da_DK'):
        if '@' not in primary_smtp_address:
            raise ValueError("primary_smtp_address '%s' is not an email address" % primary_smtp_address)
        self.primary_smtp_address = primary_smtp_address
        self.fullname = fullname
        self.locale = locale
        # Assume delegate access if individual credentials are provided. Else, assume service user with impersonation
        self.access_type = access_type or (DELEGATE if credentials else IMPERSONATION)
        assert self.access_type in (DELEGATE, IMPERSONATION)
        if autodiscover:
            if not credentials:
                raise AttributeError('autodiscover requires credentials')
            self.primary_smtp_address, self.protocol = discover(email=self.primary_smtp_address,
                                                                credentials=credentials, verify_ssl=verify_ssl)
            if config:
                raise AttributeError('config is ignored when autodiscover is active')
        else:
            if not config:
                raise AttributeError('non-autodiscover requires a config')
            self.protocol = config.protocol
        # We may need to override the default server version on a per-account basis because Microsoft may report one
        # server version up-front but delegate account requests to an older backend server.
        self.version = self.protocol.version
        self.root = Root(self)

        assert isinstance(self.protocol, Protocol)
        log.debug('Added account: %s', self)

    @property
    def folders(self):
        if hasattr(self, '_folders'):
            return self._folders
        # 'Top of Information Store' is a folder available in some Exchange accounts. It only contains folders
        # owned by the account.
        folders = self.root.get_folders(depth=SHALLOW)  # Start by searching top-level folders.
        has_tois = False
        for folder in folders:
            if folder.name == 'Top of Information Store':
                has_tois = True
                folders = folder.get_folders(depth=SHALLOW)
                break
        if not has_tois:
            # We need to dig deeper. Get everything.
            folders = self.root.get_folders(depth=DEEP)
        self._folders = dict((m, []) for m in WELLKNOWN_FOLDERS.values())
        for f in folders:
            self._folders[f.__class__].append(f)
        return self._folders

    def _get_default_folder(self, fld_class):
        try:
            # Get the default folder
            log.debug('Testing default %s folder with GetFolder', fld_class.__name__)
            return fld_class(self).get_folder()
        except ErrorAccessDenied:
            # Maybe we just don't have GetFolder access? Try FindItems instead
            log.debug('Testing default %s folder with FindItem', fld_class.__name__)
            fld = fld_class(self)
            fld.find_items(subject='DUMMY')
            return fld
        except ErrorFolderNotFound as e:
            # There's no folder named fld_class.DISTINGUISHED_FOLDER_ID. Try to guess which folder is the default.
            # Exchange makes this unnecessarily difficult.
            log.debug('Searching default %s folder in full folder list', fld_class.__name__)
            flds = []
            for folder in self.folders[fld_class]:
                # Search for a folder wth a localized name. This is a hack because I can't find a way to get the
                # default Calendar, Inbox, etc. folders without looking at the folder name which could be localized.
                # TODO: fld_class.LOCALIZED_NAMES is most definitely neither complete nor authoritative
                if folder.name.title() in fld_class.LOCALIZED_NAMES.get(self.locale, []):
                    flds.append(folder)
            if not flds:
                # There was no folder with a localized name. Use the distinguished folder instead.
                for folder in self.folders[fld_class]:
                    if folder.is_distinguished:
                        flds.append(folder)
            if not flds:
                raise ErrorFolderNotFound('No useable default %s folders' % fld_class.__name__) from e
            assert len(flds) == 1, 'Multiple possible default %s folders: %s' % (
                fld_class.__name__, [str(f) for f in flds])
            return flds[0]

    @property
    def calendar(self):
        if hasattr(self, '_calendar'):
            return self._calendar
        # If the account contains a shared calendar from a different user, that calendar will be in the folder list.
        # Attempt not to return one of those. An account may not always have a calendar called "Calendar", but a
        # Calendar folder with a localized name instead. Return that, if it's available.
        self._calendar = self._get_default_folder(Calendar)
        return self._calendar

    @property
    def inbox(self):
        if hasattr(self, '_inbox'):
            return self._inbox
        self._inbox = self._get_default_folder(Messages)
        return self._inbox

    @property
    def tasks(self):
        if hasattr(self, '_tasks'):
            return self._tasks
        self._tasks = self._get_default_folder(Tasks)
        return self._tasks

    @property
    def contacts(self):
        if hasattr(self, '_contacts'):
            return self._contacts
        self._contacts = self._get_default_folder(Contacts)
        return self._contacts

    @property
    def domain(self):
        return get_domain(self.primary_smtp_address)

    def __str__(self):
        txt = '%s' % self.primary_smtp_address
        if self.fullname:
            txt += ' (%s)' % self.fullname
        return txt
