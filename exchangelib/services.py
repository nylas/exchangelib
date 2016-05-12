"""
Implement a selection of EWS services.

Exchange is very picky about things like the order of XML elements in SOAP requests, so we need to generate XML
automatically instead of taking advantage of Python SOAP libraries and the WSDL file.

Exchange EWS references:
    - 2007: http://msdn.microsoft.com/en-us/library/bb409286(v=exchg.80).aspx
    - 2010: http://msdn.microsoft.com/en-us/library/bb409286(v=exchg.140).aspx
    - 2013: http://msdn.microsoft.com/en-us/library/bb409286(v=exchg.150).aspx
"""

import logging
import itertools
from xml.parsers.expat import ExpatError
import traceback

from . import errors
from .errors import EWSWarning, TransportError, SOAPError, ErrorTimeoutExpired, ErrorBatchProcessingStopped, \
    ErrorQuotaExceeded, ErrorCannotDeleteObject, ErrorCreateItemAccessDenied, ErrorFolderNotFound, \
    ErrorNonExistentMailbox, ErrorMailboxStoreUnavailable, ErrorImpersonateUserDenied, ErrorInternalServerError, \
    ErrorInternalServerTransientError, ErrorNoRespondingCASInDestinationSite, ErrorImpersonationFailed, \
    ErrorMailboxMoveInProgress, ErrorAccessDenied, ErrorConnectionFailed, RateLimitError, ErrorServerBusy, \
    ErrorTooManyObjectsOpened, ErrorInvalidLicense, ErrorInvalidSchemaVersionForMailboxVersion, \
    ErrorInvalidServerVersion
from .transport import wrap, SOAPNS, TNS, MNS, ENS
from .util import chunkify, create_element, add_xml_child, get_xml_attr, to_xml, post_ratelimited, ElementType, \
    xml_to_str

log = logging.getLogger(__name__)


# Shape enums
IdOnly = 'IdOnly'
# This doesn't actually get all properties in FindItem, just the "first-class" ones. See
#    http://msdn.microsoft.com/en-us/library/office/dn600367(v=exchg.150).aspx
AllProperties = 'AllProperties'

# Traversal enums
SHALLOW = 'Shallow'
DEEP = 'Deep'
SOFTDELETED = 'SoftDeleted'


class EWSService:
    SERVICE_NAME = None
    element_container_name = None
    extra_element_names = []

    def __init__(self, protocol, element_name=None):
        self.protocol = protocol
        self.element_name = element_name

    def payload(self, version, account, *args, **kwargs):
        return wrap(content=self._get_payload(*args, **kwargs), version=version, account=account)

    def _get_payload(self, *args, **kwargs):
        raise NotImplementedError()

    def _get_elements(self, payload, account=None):
        assert isinstance(payload, ElementType)
        try:
            response = self._get_response_xml(payload=payload, account=account)
            return self._get_elements_in_response(response=response)
        except (ErrorQuotaExceeded, ErrorCannotDeleteObject, ErrorCreateItemAccessDenied, ErrorTimeoutExpired,
                ErrorFolderNotFound, ErrorNonExistentMailbox, ErrorMailboxStoreUnavailable, ErrorImpersonateUserDenied,
                ErrorInternalServerError, ErrorInternalServerTransientError, ErrorNoRespondingCASInDestinationSite,
                ErrorImpersonationFailed, ErrorMailboxMoveInProgress, ErrorAccessDenied, ErrorConnectionFailed,
                RateLimitError, ErrorServerBusy, ErrorTooManyObjectsOpened, ErrorInvalidLicense):
            # These are known and understood, and don't require a backtrace
            # TODO: ErrorTooManyObjectsOpened means there are too many connections to the database. We should be able to
            # act on this by lowering the self.protocol connection pool size.
            raise
        except Exception:
            # This may run from a thread pool, which obfuscates the stack trace. Print trace immediately.
            log.warning('EWS %s, account %s: Exception in _get_elements: %s', self.protocol.ews_url, account,
                        traceback.format_exc(20))
            raise

    def _get_response_xml(self, payload, account=None):
        # Takes an XML tree and returns SOAP payload as an XML tree
        assert isinstance(payload, ElementType)
        # Microsoft really doesn't want to make our lives easy. The server may report one version in our initial version
        # guessing tango, but then the server may decide that any arbitrary legacy backend server may actually process
        # the request for an account. Prepare to handle ErrorInvalidSchemaVersionForMailboxVersion errors and set the
        # server version per-account.
        from .version import API_VERSIONS, Version
        hint = account.version.api_version if account else self.protocol.version.api_version
        api_versions = [hint] + [v for v in API_VERSIONS if v != hint]
        for api_version in api_versions:
            session = self.protocol.get_session()
            soap_payload = wrap(content=payload, version=api_version, account=account)
            r, session = post_ratelimited(
                protocol=self.protocol,
                session=session,
                url=self.protocol.ews_url,
                headers=None,
                data=soap_payload,
                timeout=self.protocol.timeout,
                verify=True,
                allow_redirects=False)
            self.protocol.release_session(session)
            log.debug('Trying API version %s for account %s', api_version, account)
            try:
                soap_response_payload = to_xml(r.text, encoding=r.encoding or 'utf-8')
            except ExpatError as e:
                raise SOAPError('SOAP response is not XML: %s' % e) from e
            try:
                res = self._get_soap_payload(soap_response=soap_response_payload)
            except (ErrorInvalidSchemaVersionForMailboxVersion, ErrorInvalidServerVersion):
                assert account # This should never happen for non-account services
                # The guessed server version is wrong for this account. Try the next version
                log.debug('API version %s was invalid for account %s', api_version, account)
                continue
            if account and account.version.api_version != api_version:
                # The api_version that worked was different than our hint. Set new version for account
                log.info('New API version for account %s (%s -> %s)', account, account.version.api_version, api_version)
                account.version = Version.from_response(requested_api_version=api_version, response=r)
            return res
        raise ErrorInvalidSchemaVersionForMailboxVersion('Tried versions %s but all were invalid for account %s' %
                                                         (api_versions, account))

    def _get_soap_payload(self, soap_response):
        log_prefix = 'EWS %s, service %s' % (self.protocol.ews_url, self.SERVICE_NAME)
        assert isinstance(soap_response, ElementType)
        body = soap_response.find('{%s}Body' % SOAPNS)
        if body is None:
            raise TransportError('No Body element in SOAP response')
        response = body.find('{%s}%sResponse' % (MNS, self.SERVICE_NAME))
        if response is None:
            fault = body.find('{%s}Fault' % SOAPNS)
            if fault is None:
                raise SOAPError('Unknown SOAP response: %s' % xml_to_str(body))
            self._raise_soap_errors(fault=fault)  # Will throw SOAPError
        response_messages = response.find('{%s}ResponseMessages' % MNS)
        if response_messages is None:
            raise TransportError('%s: No ResponseMessages element in response' % log_prefix)
        return response_messages.findall('{%s}%sResponseMessage' % (MNS, self.SERVICE_NAME))

    def _raise_soap_errors(self, fault):
        assert isinstance(fault, ElementType)
        log_prefix = 'EWS %s, service %s' % (self.protocol.ews_url, self.SERVICE_NAME)
        # Fault: See http://www.w3.org/TR/2000/NOTE-SOAP-20000508/#_Toc478383507
        faultcode = get_xml_attr(fault, 'faultcode')
        faultstring = get_xml_attr(fault, 'faultstring')
        faultactor = get_xml_attr(fault, 'faultactor')
        detail = fault.find('detail')
        if detail is not None:
            code, msg = None, None
            if detail.find('{%s}ResponseCode' % ENS) is not None:
                code = get_xml_attr(detail, '{%s}ResponseCode' % ENS)
            if detail.find('{%s}Message' % ENS) is not None:
                msg = get_xml_attr(detail, '{%s}Message' % ENS)
            try:
                raise vars(errors)[code](msg)
            except KeyError:
                detail = '%s: code: %s msg: %s (%s)' % (log_prefix, code, msg, xml_to_str(detail))
        try:
            raise vars(errors)[faultcode](faultstring)
        except KeyError:
            pass
        raise SOAPError('SOAP error code: %s string: %s actor: %s detail: %s' % (
            faultcode, faultstring, faultactor, detail))

    def _get_element_container(self, message, name=None):
        assert isinstance(message, ElementType)
        # ResponseClass: See http://msdn.microsoft.com/en-us/library/aa566424(v=EXCHG.140).aspx
        response_class = message.get('ResponseClass')
        # ResponseCode, MessageText: See http://msdn.microsoft.com/en-us/library/aa580757(v=EXCHG.140).aspx
        response_code = get_xml_attr(message, '{%s}ResponseCode' % MNS)
        msg_text = get_xml_attr(message, '{%s}MessageText' % MNS)
        msg_xml = get_xml_attr(message, '{%s}MessageXml' % MNS)
        if response_class == 'Success' and response_code == 'NoError':
            if not name:
                return True
            container = message.find(name)
            if container is None:
                raise TransportError('No %s elements in ResponseMessage (%s)' % (name, xml_to_str(message)))
            return container
        if response_class == 'Warning':
            return self._raise_warnings(code=response_code, text=msg_text, xml=msg_xml)
        # rspclass == 'Error', or 'Success' and not 'NoError'
        return self._raise_errors(code=response_code, text=msg_text, xml=msg_xml)

    def _raise_warnings(self, code, text, xml):
        try:
            return self._raise_errors(code=code, text=text, xml=xml)
        except ErrorBatchProcessingStopped as e:
            raise EWSWarning(e.value) from e

    @staticmethod
    def _raise_errors(code, text, xml):
        if code == 'NoError':
            return True
        if not code:
            raise TransportError('Empty ResponseCode in ResponseMessage (MessageText: %s, MessageXml: %s)' % (
                text, xml))
        try:
            # Raise the error corresponding to the ResponseCode
            raise vars(errors)[code](text)
        except KeyError as e:
            # Should not happen
            raise TransportError('Unknown ResponseCode in ResponseMessage: %s (MessageText: %s, MessageXml: %s)' % (
                code, text, xml)) from e

    def _get_elements_in_response(self, response):
        assert isinstance(response, list)
        elements = []
        for msg in response:
            assert isinstance(msg, ElementType)
            try:
                container = self._get_element_container(message=msg, name=self.element_container_name)
                if isinstance(container, ElementType):
                    elements.extend(self._get_elements_in_container(container=container))
                else:
                    elements.append((container, None))
            except (ErrorTimeoutExpired, ErrorBatchProcessingStopped):
                raise
            except EWSWarning as e:
                elements.append((False, '%s' % e.value))
                continue
        return elements

    def _get_elements_in_container(self, container):
        elems = container.findall(self.element_name)
        for element_name in self.extra_element_names:
            elems.extend(container.findall(element_name))
        return elems


class EWSAccountService(EWSService):
    def call(self, account, **kwargs):
        raise NotImplementedError()


class EWSFolderService(EWSService):
    def call(self, folder, **kwargs):
        raise NotImplementedError()


class PagingEWSService(EWSService):
    def _paged_call(self, **kwargs):
        # TODO This is awkward. The function must work with _get_payload() of both folder- and account-based services
        account = kwargs['folder'].account if 'folder' in kwargs else kwargs['account']
        log_prefix = 'EWS %s, account %s, service %s' % (self.protocol.ews_url, account, self.SERVICE_NAME)
        elements = []
        offset = 0
        while True:
            log.debug('%s: Getting %s at offset %s', log_prefix, self.element_name, offset)
            kwargs['offset'] = offset
            payload = self._get_payload(**kwargs)
            response = self._get_response_xml(payload=payload, account=account)
            page, offset = self._get_page(response)
            if isinstance(page, ElementType):
                container = page.find(self.element_container_name)
                if container is None:
                    raise TransportError('No %s elements in ResponseMessage (%s)' % (self.element_container_name,
                                                                                     xml_to_str(page)))
                elements.extend(self._get_elements_in_container(container=container))
            if not offset:
                break
        return elements

    def _get_page(self, response):
        assert len(response) == 1
        log_prefix = 'EWS %s, service %s' % (self.protocol.ews_url, self.SERVICE_NAME)
        rootfolder = self._get_element_container(message=response[0], name='{%s}RootFolder' % MNS)
        is_last_page = rootfolder.get('IncludesLastItemInRange').lower() in ('true', '0')
        offset = rootfolder.get('IndexedPagingOffset')
        if offset is None and not is_last_page:
            log.warning("Not last page in range, but Exchange didn't send a page offset. Assuming first page")
            offset = '1'
        next_offset = 0 if is_last_page else int(offset)
        if not int(rootfolder.get('TotalItemsInView')):
            assert next_offset == 0
            rootfolder = None
        log.debug('%s: Got page with next offset %s (last_page %s)', log_prefix, next_offset, is_last_page)
        return rootfolder, next_offset


class GetServerTimeZones(EWSService):
    SERVICE_NAME = 'GetServerTimeZones'
    element_container_name = '{%s}TimeZoneDefinitions' % MNS

    def __init__(self, *args, **kwargs):
        kwargs['element_name'] = '{%s}TimeZoneDefinition' % TNS
        super().__init__(*args, **kwargs)

    def call(self, **kwargs):
        if self.protocol.version.major_version < 14:
            raise NotImplementedError('%s is only supported for Exchange 2010 servers and later' % self.SERVICE_NAME)
        return self._get_elements(payload=self._get_payload(**kwargs))

    def _get_payload(self, returnfulltimezonedata=False):
        return create_element('m:%s' % self.SERVICE_NAME, ReturnFullTimeZoneData=(
            'true' if returnfulltimezonedata else 'false'))

    def _get_elements_in_container(self, container):
        timezones = []
        timezonedefs = container.findall(self.element_name)
        for timezonedef in timezonedefs:
            tz_id = timezonedef.get('Id')
            name = timezonedef.get('Name')
            timezones.append((tz_id, name))
        return timezones


class EWSPooledService(EWSService):
    CHUNKSIZE = None

    def _pool_requests(self, account, payload_func, items):
        log.debug('Processing %s items in chunks of %s', len(items), self.CHUNKSIZE)
        # Chop items list into suitable pieces and let worker threads chew on the work. The order of the output result
        # list must be the same as the input id list, so the caller knows which status message belongs to which ID.
        func = lambda n: self._get_elements(account=account, payload=payload_func(n))
        return list(itertools.chain(*account.protocol.thread_pool.map(func, chunkify(items, self.CHUNKSIZE))))


class GetItem(EWSPooledService):
    """
    Take a list of (id, changekey) tuples and returns a list of items in stable order
    """
    CHUNKSIZE = 100
    SERVICE_NAME = 'GetItem'
    element_container_name = '{%s}Items' % MNS

    def call(self, folder, **kwargs):
        self.element_name = folder.item_model.response_tag()
        return self._pool_requests(account=folder.account, payload_func=folder.get_xml, items=kwargs['ids'])


class CreateItem(EWSPooledService):
    """
    Takes folder and a list of items. Returns result of creation as a list of tuples (success[True|False],
    errormessage), in the same order as the input list.

    """
    CHUNKSIZE = 25
    SERVICE_NAME = 'CreateItem'
    element_container_name = '{%s}Items' % MNS

    def call(self, folder, **kwargs):
        self.element_name = folder.item_model.response_tag()
        return self._pool_requests(account=folder.account, payload_func=folder.create_xml, items=kwargs['items'])


class DeleteItem(EWSPooledService):
    """
    Takes a folder and a list of (id, changekey) tuples. Returns result of deletion as a list of tuples
    (success[True|False], errormessage), in the same order as the input list.

    """
    CHUNKSIZE = 25
    SERVICE_NAME = 'DeleteItem'
    element_container_name = None  # DeleteItem doesn't return a response object, just status in XML attrs

    def call(self, folder, **kwargs):
        return self._pool_requests(account=folder.account, payload_func=folder.delete_xml, items=kwargs['ids'])


class UpdateItem(EWSPooledService):
    CHUNKSIZE = 25
    SERVICE_NAME = 'UpdateItem'
    element_container_name = '{%s}Items' % MNS

    def call(self, folder, **kwargs):
        self.element_name = folder.item_model.response_tag()
        return self._pool_requests(account=folder.account, payload_func=folder.update_xml, items=kwargs['items'])


class FindItem(PagingEWSService, EWSFolderService):
    """
    Gets all items for 'account' in folder 'folder_id', optionally expanded with 'additional_fields' Element,
    optionally restricted by a Restriction definition.
    """
    SERVICE_NAME = 'FindItem'
    element_container_name = '{%s}Items' % TNS

    def call(self, folder, **kwargs):
        self.element_name = folder.item_model.response_tag()
        return self._paged_call(folder=folder, **kwargs)

    def _get_payload(self, folder, additional_fields=None, restriction=None, shape=IdOnly, depth=SHALLOW, offset=0):
        log.debug(
            'Finding %s items for %s extra fields %s restriction %s shape %s offset %s',
            folder.DISTINGUISHED_FOLDER_ID,
            folder.account,
            additional_fields,
            restriction,
            shape,
            offset,
        )
        finditem = create_element('m:%s' % self.SERVICE_NAME, Traversal=SHALLOW)
        itemshape = create_element('m:ItemShape')
        add_xml_child(itemshape, 't:BaseShape', shape)
        if additional_fields:
            additionalproperties = create_element('t:AdditionalProperties')
            for field_uri in additional_fields:
                additionalproperties.append(create_element('t:FieldURI', FieldURI=field_uri))
            itemshape.append(additionalproperties)
        finditem.append(itemshape)
        indexedpageviewitem = create_element('m:IndexedPageItemView', Offset=str(offset), BasePoint='Beginning')
        finditem.append(indexedpageviewitem)
        if restriction:
            finditem.append(restriction.xml)
        parentfolderids = create_element('m:ParentFolderIds')
        parentfolderids.append(folder.folderid_xml())
        finditem.append(parentfolderids)
        return finditem


class FindFolder(PagingEWSService, EWSFolderService):
    """
    Gets a list of folders belonging to an account.
    """
    SERVICE_NAME = 'FindFolder'
    element_container_name = '{%s}Folders' % TNS
    # See http://msdn.microsoft.com/en-us/library/aa564009(v=exchg.150).aspx
    extra_element_names = [
        '{%s}CalendarFolder' % TNS,
        '{%s}ContactsFolder' % TNS,
        '{%s}SearchFolder' % TNS,
        '{%s}TasksFolder' % TNS,
    ]

    def __init__(self, *args, **kwargs):
        kwargs['element_name'] = '{%s}Folder' % TNS
        super().__init__(*args, **kwargs)

    def call(self, folder, **kwargs):
        return self._paged_call(folder=folder, **kwargs)

    def _get_payload(self, folder, additional_fields=None, shape=IdOnly, depth=DEEP, offset=0):
        log.debug(
            'Getting folders for %s, root:%s, extra fields:%s, shape:%s',
            folder.account,
            folder.name,
            additional_fields,
            shape
        )
        findfolder = create_element('m:%s' % self.SERVICE_NAME, Traversal=depth)
        foldershape = create_element('m:FolderShape')
        add_xml_child(foldershape, 't:BaseShape', shape)
        if additional_fields:
            additionalproperties = create_element('t:AdditionalProperties')
            for field_uri in additional_fields:
                additionalproperties.append(create_element('t:FieldURI', FieldURI=field_uri))
            foldershape.append(additionalproperties)
        findfolder.append(foldershape)
        if folder.account.protocol.version.major_version >= 14:
            indexedpageviewitem = create_element('m:IndexedPageFolderView', Offset=str(offset), BasePoint='Beginning')
            findfolder.append(indexedpageviewitem)
        else:
            assert offset == 0, 'Offset is %s' % offset
        parentfolderids = create_element('m:ParentFolderIds')
        parentfolderids.append(folder.folderid_xml())
        findfolder.append(parentfolderids)
        return findfolder


class GetFolder(EWSFolderService):
    SERVICE_NAME = 'GetFolder'
    element_container_name = '{%s}Folders' % MNS
    # See http://msdn.microsoft.com/en-us/library/aa564009(v=exchg.150).aspx
    extra_element_names = [
        '{%s}CalendarFolder' % TNS,
        '{%s}ContactsFolder' % TNS,
        '{%s}SearchFolder' % TNS,
        '{%s}TasksFolder' % TNS,
    ]

    def __init__(self, *args, **kwargs):
        kwargs['element_name'] = '{%s}Folder' % TNS
        super().__init__(*args, **kwargs)

    def call(self, folder, **kwargs):
        return self._get_elements(payload=self._get_payload(folder, **kwargs), account=folder.account)

    def _get_payload(self, folder, additional_fields=None, shape=IdOnly):
        log.debug(
            'Getting folders for %s, folder:%s, extra fields:%s, shape:%s',
            folder.account,
            folder.name,
            additional_fields,
            shape
        )
        getfolder = create_element('m:%s' % self.SERVICE_NAME)
        foldershape = create_element('m:FolderShape')
        add_xml_child(foldershape, 't:BaseShape', shape)
        if additional_fields:
            additionalproperties = create_element('t:AdditionalProperties')
            for field_uri in additional_fields:
                additionalproperties.append(create_element('t:FieldURI', FieldURI=field_uri))
            foldershape.append(additionalproperties)
        getfolder.append(foldershape)
        folderids = create_element('m:FolderIds')
        folderids.append(folder.folderid_xml())
        getfolder.append(folderids)
        return getfolder


class ResolveNames(EWSAccountService):
    SERVICE_NAME = 'ResolveNames'
    element_container_name = '{%s}ResolutionSet' % MNS

    def __init__(self, *args, **kwargs):
        kwargs['element_name'] = '{%s}Resolution' % TNS
        super().__init__(*args, **kwargs)

    def call(self, **kwargs):
        return self._get_elements(payload=self._get_payload(**kwargs))

    def _get_payload(self, unresolvedentries, returnfullcontactdata=False):
        payload = create_element('m:%s' % self.SERVICE_NAME, ReturnFullContactData=(
            'true' if returnfullcontactdata else 'false'))
        assert len(unresolvedentries)
        for entry in unresolvedentries:
            add_xml_child(payload, 'm:UnresolvedEntry', entry)
        return payload
