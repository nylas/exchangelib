from __future__ import unicode_literals

import logging

from six import text_type

from .fields import EmailSubField, LabelField, SubField, NamedSubField, Choice
from .properties import EWSElement

log = logging.getLogger(__name__)


class IndexedElement(EWSElement):
    LABELS = set()

    __slots__ = ('label',)


class SingleFieldIndexedElement(IndexedElement):
    __slots__ = ('label',)

    @classmethod
    def value_field(cls, version=None):
        fields = cls.supported_fields(version=version)
        if len(fields) != 1:
            raise ValueError('This class must have only one field (found %s)' % fields)
        return fields[0]


class EmailAddress(SingleFieldIndexedElement):
    # MSDN:  https://msdn.microsoft.com/en-us/library/office/aa564757(v=exchg.150).aspx
    ELEMENT_NAME = 'Entry'
    FIELDS = [
        LabelField('label', field_uri='Key', choices={
            Choice('EmailAddress1'), Choice('EmailAddress2'), Choice('EmailAddress3')
        }, default='EmailAddress1'),
        EmailSubField('email'),
    ]

    __slots__ = ('label', 'email')


class PhoneNumber(SingleFieldIndexedElement):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa565941(v=exchg.150).aspx
    ELEMENT_NAME = 'Entry'
    FIELDS = [
        LabelField('label', field_uri='Key', choices={
            Choice('AssistantPhone'), Choice('BusinessFax'), Choice('BusinessPhone'), Choice('BusinessPhone2'),
            Choice('Callback'), Choice('CarPhone'), Choice('CompanyMainPhone'), Choice('HomeFax'), Choice('HomePhone'),
            Choice('HomePhone2'), Choice('Isdn'), Choice('MobilePhone'), Choice('OtherFax'), Choice('OtherTelephone'),
            Choice('Pager'), Choice('PrimaryPhone'), Choice('RadioPhone'), Choice('Telex'), Choice('TtyTddPhone'),
        }, default='PrimaryPhone'),
        SubField('phone_number'),
    ]

    __slots__ = ('label', 'phone_number')


class MultiFieldIndexedElement(IndexedElement):
    __slots__ = ('label',)


class PhysicalAddress(MultiFieldIndexedElement):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa564323(v=exchg.150).aspx
    ELEMENT_NAME = 'Entry'
    FIELDS = [
        LabelField('label', field_uri='Key', choices={
            Choice('Business'), Choice('Home'), Choice('Other')
        }, default='Business'),
        NamedSubField('street', field_uri='Street'),  # Street, house number, etc.
        NamedSubField('city', field_uri='City'),
        NamedSubField('state', field_uri='State'),
        NamedSubField('country', field_uri='CountryOrRegion'),
        NamedSubField('zipcode', field_uri='PostalCode'),
    ]

    __slots__ = ('label', 'street', 'city', 'state', 'country', 'zipcode')

    def clean(self, version=None):
        # pylint: disable=access-member-before-definition
        if isinstance(self.zipcode, int):
            self.zipcode = text_type(self.zipcode)
        super(PhysicalAddress, self).clean(version=version)
