import logging
from xml.etree.ElementTree import ParseError

import requests.sessions
import requests.adapters

from .errors import UnauthorizedError, TransportError, EWSWarning
from .transport import TNS, SOAPNS, dummy_xml, get_auth_instance
from .util import is_xml, to_xml, post_ratelimited

log = logging.getLogger(__name__)

# Legend for dict:
#   Key: shortname
#   Values: (EWS API version ID, full name)

# 'shortname' comes from types.xsd and is the official version of the server, corresponding to the version numbers
# supplied in SOAP headers. 'API version' is the version name supplied in the RequestServerVersion element in SOAP
# headers and describes the EWS API version the server accepts. Valid values for this element are described here:
#    http://msdn.microsoft.com/en-us/library/bb891876(v=exchg.150).aspx

VERSIONS = {
    'Exchange2007': ('Exchange2007', 'Microsoft Exchange Server 2007'),
    'Exchange2007_SP1': ('Exchange2007_SP1', 'Microsoft Exchange Server 2007 SP1'),
    'Exchange2007_SP2': ('Exchange2007_SP1', 'Microsoft Exchange Server 2007 SP2'),
    'Exchange2007_SP3': ('Exchange2007_SP1', 'Microsoft Exchange Server 2007 SP3'),
    'Exchange2010': ('Exchange2010', 'Microsoft Exchange Server 2010'),
    'Exchange2010_SP1': ('Exchange2010_SP1', 'Microsoft Exchange Server 2010 SP1'),
    'Exchange2010_SP2': ('Exchange2010_SP2', 'Microsoft Exchange Server 2010 SP2'),
    'Exchange2010_SP3': ('Exchange2010_SP2', 'Microsoft Exchange Server 2010 SP3'),
    'Exchange2013': ('Exchange2013', 'Microsoft Exchange Server 2013'),
    'Exchange2013_SP1': ('Exchange2013_SP1', 'Microsoft Exchange Server 2013 SP1'),
    'Exchange2015': ('Exchange2015', 'Microsoft Exchange Server 2015'),
    'Exchange2015_SP1': ('Exchange2015_SP1', 'Microsoft Exchange Server 2015 SP1'),
    'Exchange2016': ('Exchange2016', 'Microsoft Exchange Server 2016'),
}

# List of build numbers here: https://technet.microsoft.com/en-gb/library/hh135098(v=exchg.150).aspx
API_VERSION_FROM_BUILD_NUMBER = {
    8: {
        0: 'Exchange2007',
        1: 'Exchange2007_SP1',
        2: 'Exchange2007_SP1',
        3: 'Exchange2007_SP1',
    },
    14: {
        0: 'Exchange2010',
        1: 'Exchange2010_SP1',
        2: 'Exchange2010_SP2',
        3: 'Exchange2010_SP2',
    },
    15: {
        0: 'Exchange2013',  # Minor builds above 847 are Exchange2013_SP1
        1: 'Exchange2016',
    },
}

# Build a list of unique API versions, used when guessing API version supported by the server.  Use reverse order so we
# get the newest API version supported by the server.
API_VERSIONS = sorted({v[0] for v in VERSIONS.values()}, reverse=True)


class Version:
    """
    Holds information about the server version
    """

    def __init__(self, major_version, minor_version, major_build, minor_build, api_version):
        self.major_version = major_version
        self.minor_version = minor_version
        self.major_build = major_build
        self.minor_build = minor_build
        self.api_version = api_version
        if major_version < 8:
            raise ValueError("Exchange major versions below 8 don't support EWS (%s)", str(self))

    @property
    def build(self):
        return '%s.%s.%s.%s' % (self.major_version, self.minor_version, self.major_build, self.minor_build)

    @property
    def fullname(self):
        return VERSIONS[self.api_version][1]

    @classmethod
    def guess(cls, protocol):
        """
        Tries to ask the server which version it has. We haven't set up an Account object yet, so generate a request
        by hand. We only need a response header containing a ServerVersionInfo element. Apparently, EWS has no problem
        supplying one version in its types.xsd and reporting another in its SOAP headers. Trust the SOAP version.
        """
        log.debug('Asking server for version info')
        # Can't use a session object from the protocol pool for docs because sessions are created with service auth.
        try:
            auth = get_auth_instance(credentials=protocol.credentials, auth_type=protocol.docs_auth_type)
            shortname = cls._get_shortname_from_docs(auth=auth, types_url=protocol.types_url)
            log.debug('Shortname according to %s: %s', protocol.types_url, shortname)
        except (TransportError, UnauthorizedError) as e:
            log.warning(str(e))
            shortname = None
        api_version = VERSIONS[shortname][0] if shortname else None
        return cls._guess_version_from_service(protocol=protocol, ews_url=protocol.ews_url, hint=api_version)

    @classmethod
    def _get_shortname_from_docs(cls, auth, types_url):
        # Get the server version from types.xsd. A server response provides the build numbers. We can't necessarily use
        # the service auth type since it may not be the same as the auth type for docs.
        log.debug('Getting %s with auth type %s', types_url, auth.__class__.__name__)
        # Some servers send an empty response if we send 'Connection': 'close' header
        with requests.sessions.Session() as s:
            r = s.get(url=types_url, auth=auth, allow_redirects=False, stream=False)
        log.debug('Request headers: %s', r.request.headers)
        log.debug('Response code: %s', r.status_code)
        log.debug('Response headers: %s', r.headers)
        if r.status_code == 401:
            raise UnauthorizedError('Wrong username or password for %s' % types_url)
        if r.status_code == 302:
            log.debug('We were redirected. Cant get version info from docs')
            return None
        if r.status_code == 503:
            log.debug('Service is unavailable. Cant get version info from docs')
            return None
        if r.status_code != 200:
            if 'The referenced account is currently locked out' in r.text:
                raise TransportError('The service account is currently locked out')
            raise TransportError('Unexpected HTTP status %s when getting %s (%s)' % (r.status_code, types_url, r.text))
        if not is_xml(r.text):
            raise TransportError('Unexpected result when getting %s. Maybe this is not an EWS server?%s' % (
                types_url,
                '\n\n%s[...]' % r.text[:200] if len(r.text) > 200 else '\n\n%s' % r.text if r.text else '',
            ))
        return to_xml(r.text, encoding=r.encoding).get('version')

    @classmethod
    def _guess_version_from_service(cls, protocol, ews_url, hint=None):
        # We need to guess the version. If we got a shortname from docs, start guessing that
        if hint:
            api_versions = [hint] + [v for v in API_VERSIONS if v != hint]
        else:
            api_versions = API_VERSIONS
        for api_version in api_versions:
            try:
                return cls._get_version_from_service(protocol=protocol, ews_url=ews_url, api_version=api_version)
            except EWSWarning:
                continue
        raise TransportError('Unable to guess version')

    @classmethod
    def _get_version_from_service(cls, protocol, ews_url, api_version):
        assert api_version
        xml = dummy_xml(version=api_version)
        # Create a minimal, valid EWS request to force Exchange into accepting the request and returning EWS xml
        # containing server version info. Some servers will only reply with their version if a valid POST is sent.
        session = protocol.get_session()
        log.debug('Test if service API version is %s using auth %s', api_version, session.auth.__class__.__name__)
        r, session = post_ratelimited(protocol=protocol, session=session, url=ews_url, headers=None, data=xml,
                                      timeout=protocol.timeout, verify=True, allow_redirects=False)
        protocol.release_session(session)

        if r.status_code == 401:
            raise UnauthorizedError('Wrong username or password for %s' % ews_url)
        elif r.status_code == 302:
            log.debug('We were redirected. Cant get version info from docs')
            return None
        elif r.status_code == 503:
            log.debug('Service is unavailable. Cant get version info from docs')
            return None
        if r.status_code == 400:
            raise EWSWarning('Bad request')
        if r.status_code == 500 and ('The specified server version is invalid' in r.text or
                                     'ErrorInvalidSchemaVersionForMailboxVersion' in r.text):
            raise EWSWarning('Invalid server version')
        if r.status_code != 200:
            if 'The referenced account is currently locked out' in r.text:
                raise TransportError('The service account is currently locked out')
            raise TransportError('Unexpected HTTP status %s when getting %s (%s)' % (r.status_code, ews_url, r.text))
        log.debug('Response data: %s', r.text)
        try:
            header = to_xml(r.text, encoding=r.encoding).find('{%s}Header' % SOAPNS)
            if not header:
                raise ParseError()
        except ParseError as e:
            raise EWSWarning('Unknown XML response from %s (response: %s)' % (ews_url, r.text)) from e
        info = header.find('{%s}ServerVersionInfo' % TNS)
        if info is None:
            raise TransportError('No ServerVersionInfo in response: %s' % r.text)

        version = cls.from_response(requested_api_version=api_version, response=r)
        log.debug('Service version is: %s', version)
        return version

    @classmethod
    def from_response(cls, requested_api_version, response):
        try:
            header = to_xml(response.text, encoding=response.encoding).find('{%s}Header' % SOAPNS)
            if not header:
                raise ParseError()
        except ParseError as e:
            raise EWSWarning('Unknown XML response from %s (response: %s)' % (response, response.text)) from e
        info = header.find('{%s}ServerVersionInfo' % TNS)
        if info is None:
            raise TransportError('No ServerVersionInfo in response: %s' % response.text)

        major_version, minor_version, major_build, minor_build = \
            [int(info.get(k)) for k in ['MajorVersion', 'MinorVersion', 'MajorBuildNumber', 'MinorBuildNumber']]
        for k, v in dict(MajorVersion=major_version, MinorVersion=minor_version, MajorBuildNumber=major_build,
                         MinorBuildNumber=minor_build).items():
            if v is None:
                raise TransportError('No %s in response: %s' % (k, response.text))
        api_version_from_server = info.get('Version')
        if api_version_from_server is None:
            # Not all Exchange servers send the Version element
            api_version_from_server = cls.api_version_from_build_number(major_version, minor_version, major_build)
        if api_version_from_server != requested_api_version:
            if api_version_from_server.startswith('V2_') \
                    or api_version_from_server.startswith('V2015_') \
                    or api_version_from_server.startswith('V2016_'):
                # Office 365 is an expert in sending invalid server versions...
                log.info('API version "%s" worked but server reports version "%s". Using "%s"', requested_api_version,
                         api_version_from_server, requested_api_version)
                api_version_from_server = requested_api_version
            else:
                # Work around a bug in Exchange that reports a bogus API version in the XML response. Trust server
                # response except 'V2_nn' or 'V201[5,6]_nn_mm' which is bogus
                log.info('API version "%s" worked but server reports version "%s". Using "%s"', requested_api_version,
                         api_version_from_server, api_version_from_server)
        return cls(major_version, minor_version, major_build, minor_build, api_version_from_server)

    @staticmethod
    def api_version_from_build_number(major_version, minor_version, major_build):
        api_version = API_VERSION_FROM_BUILD_NUMBER[major_version][minor_version]
        if major_version == 15 and major_version == 0 and major_build >= 847:
            api_version = 'Exchange2013_SP1'
        return api_version

    def __str__(self):
        return 'Build=%s, API=%s, Fullname=%s' % (self.build, self.api_version, self.fullname)
