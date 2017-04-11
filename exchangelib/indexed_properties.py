from __future__ import unicode_literals

import logging

from six import string_types, text_type

from .fields import EmailSubField, LabelField, SubField, TextField, Choice
from .properties import EWSElement
from .util import create_element, set_xml_value, add_xml_child

string_type = string_types[0]
log = logging.getLogger(__name__)


class IndexedElement(EWSElement):
    LABELS = set()
    LABEL_FIELD = None

    __slots__ = ('label',)

    def __init__(self, **kwargs):
        self.label = kwargs.pop('label', None)
        super(IndexedElement, self).__init__(**kwargs)

    def clean(self, version=None):
        self.LABEL_FIELD.clean(self.label, version=version)
        super(IndexedElement, self).clean(version=version)


class SingleFieldIndexedElement(IndexedElement):
    __slots__ = ('label',)

    @classmethod
    def from_xml(cls, elem):
        if elem is None:
            return None
        assert elem.tag == cls.response_tag(), (cls, elem.tag, cls.response_tag())
        kwargs = {f.name: f.from_xml(elem=elem) for f in cls.FIELDS}
        kwargs[cls.LABEL_FIELD.name] = elem.get(cls.LABEL_FIELD.field_uri)
        elem.clear()
        return cls(**kwargs)

    def to_xml(self, version):
        self.clean(version=version)
        entry = create_element(self.request_tag(), Key=self.label)
        for f in self.supported_fields(version=version):
            set_xml_value(entry, f.to_xml(getattr(self, f.name), version=version), version)
        return entry


class EmailAddress(SingleFieldIndexedElement):
    # MSDN:  https://msdn.microsoft.com/en-us/library/office/aa564757(v=exchg.150).aspx
    ELEMENT_NAME = 'Entry'
    LABEL_FIELD = LabelField('label', field_uri='Key', choices={
        Choice('EmailAddress1'), Choice('EmailAddress2'), Choice('EmailAddress3')
    }, default='EmailAddress1')
    FIELDS = [
        EmailSubField('email'),
    ]

    __slots__ = ('label', 'email')


class PhoneNumber(SingleFieldIndexedElement):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa565941(v=exchg.150).aspx
    ELEMENT_NAME = 'Entry'
    LABEL_FIELD = LabelField('label', field_uri='Key', choices={
        Choice('AssistantPhone'), Choice('BusinessFax'), Choice('BusinessPhone'), Choice('BusinessPhone2'),
        Choice('Callback'), Choice('CarPhone'), Choice('CompanyMainPhone'), Choice('HomeFax'), Choice('HomePhone'),
        Choice('HomePhone2'), Choice('Isdn'), Choice('MobilePhone'), Choice('OtherFax'), Choice('OtherTelephone'),
        Choice('Pager'), Choice('PrimaryPhone'), Choice('RadioPhone'), Choice('Telex'), Choice('TtyTddPhone'),
    }, default='PrimaryPhone')
    FIELDS = [
        SubField('phone_number'),
    ]

    __slots__ = ('label', 'phone_number')


class MultiFieldIndexedElement(IndexedElement):
    __slots__ = ('label',)

    @classmethod
    def from_xml(cls, elem):
        if elem is None:
            return None
        assert elem.tag == cls.response_tag(), (cls, elem.tag, cls.response_tag())
        kwargs = {f.name: f.from_xml(elem=elem) for f in cls.FIELDS}
        kwargs['label'] = cls.LABEL_FIELD.from_xml(elem=elem)
        elem.clear()
        return cls(**kwargs)

    def to_xml(self, version):
        self.clean(version=version)
        entry = create_element(self.request_tag(), Key=self.label)
        for f in self.supported_fields(version=version):
            value = getattr(self, f.name)
            if value is not None:
                add_xml_child(entry, f.request_tag(), value)
        return entry


class PhysicalAddress(MultiFieldIndexedElement):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa564323(v=exchg.150).aspx
    ELEMENT_NAME = 'Entry'
    LABEL_FIELD = LabelField('label', field_uri='Key', choices={
        Choice('Business'), Choice('Home'), Choice('Other')
    }, default='Business')
    FIELDS = [
        TextField('street', field_uri='Street'),  # Street, house number, etc.
        TextField('city', field_uri='City'),
        TextField('state', field_uri='State'),
        TextField('country', field_uri='CountryOrRegion'),
        TextField('zipcode', field_uri='PostalCode'),
    ]

    __slots__ = ('label', 'street', 'city', 'state', 'country', 'zipcode')

    def clean(self, version=None):
        # pylint: disable=access-member-before-definition
        if isinstance(self.zipcode, int):
            self.zipcode = text_type(self.zipcode)
        super(PhysicalAddress, self).clean(version=version)
