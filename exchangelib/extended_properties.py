from __future__ import unicode_literals

import base64
import logging
from decimal import Decimal

from six import string_types

from .ewsdatetime import EWSDateTime
from .properties import EWSElement
from .util import create_element, add_xml_child, get_xml_attrs, get_xml_attr, set_xml_value, value_to_xml_text, \
    xml_text_to_value, is_iterable, TNS

log = logging.getLogger(__name__)


class ExtendedProperty(EWSElement):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa566405(v=exchg.150).aspx

    Property_* values: https://msdn.microsoft.com/en-us/library/office/aa564843(v=exchg.150).aspx
    """
    ELEMENT_NAME = 'ExtendedProperty'

    DISTINGUISHED_SETS = {
        'Meeting',
        'Appointment',
        'Common',
        'PublicStrings',
        'Address',
        'InternetHeaders',
        'CalendarAssistant',
        'UnifiedMessaging',
    }
    PROPERTY_TYPES = {
        'ApplicationTime',
        'Binary',
        'BinaryArray',
        'Boolean',
        'CLSID',
        'CLSIDArray',
        'Currency',
        'CurrencyArray',
        'Double',
        'DoubleArray',
        # 'Error',
        'Float',
        'FloatArray',
        'Integer',
        'IntegerArray',
        'Long',
        'LongArray',
        # 'Null',
        # 'Object',
        # 'ObjectArray',
        'Short',
        'ShortArray',
        'SystemTime',
        'SystemTimeArray',
        'String',
        'StringArray',
    }  # The commented-out types cannot be used for setting or getting (see docs) and are thus not very useful here

    distinguished_property_set_id = None
    property_set_id = None
    property_tag = None  # hex integer (e.g. 0x8000) or string ('0x8000')
    property_name = None
    property_id = None  # integer as hex-formatted int (e.g. 0x8000) or normal int (32768)
    property_type = ''

    __slots__ = ('value',)

    def __init__(self, *args, **kwargs):
        if not kwargs:
            # Allow to set attributes without keyword
            kwargs = dict(zip(self.__slots__, args))
        self.value = kwargs.pop('value')
        super(ExtendedProperty, self).__init__(**kwargs)

    @classmethod
    def validate_cls(cls):
        # Validate values of class attributes and their inter-dependencies
        cls._validate_distinguished_property_set_id()
        cls._validate_property_set_id()
        cls._validate_property_tag()
        cls._validate_property_name()
        cls._validate_property_id()
        cls._validate_property_type()

    @classmethod
    def _validate_distinguished_property_set_id(cls):
        if cls.distinguished_property_set_id:
            if any([cls.property_set_id, cls.property_tag]):
                raise ValueError(
                    "When 'distinguished_property_set_id' is set, 'property_set_id' and 'property_tag' must be None"
                )
            if not any([cls.property_id, cls.property_name]):
                raise ValueError(
                    "When 'distinguished_property_set_id' is set, 'property_id' or 'property_name' must also be set"
                )
            if cls.distinguished_property_set_id not in cls.DISTINGUISHED_SETS:
                raise ValueError(
                    "'distinguished_property_set_id' value '%s' must be one of %s"
                    % (cls.distinguished_property_set_id, sorted(cls.DISTINGUISHED_SETS))
                )

    @classmethod
    def _validate_property_set_id(cls):
        if cls.property_set_id:
            if any([cls.distinguished_property_set_id, cls.property_tag]):
                raise ValueError(
                    "When 'property_set_id' is set, 'distinguished_property_set_id' and 'property_tag' must be None"
                )
            if not any([cls.property_id, cls.property_name]):
                raise ValueError(
                    "When 'property_set_id' is set, 'property_id' or 'property_name' must also be set"
                )

    @classmethod
    def _validate_property_tag(cls):
        if cls.property_tag:
            if any([
                cls.distinguished_property_set_id, cls.property_set_id, cls.property_name, cls.property_id
            ]):
                raise ValueError("When 'property_tag' is set, only 'property_type' must be set")
            if 0x8000 <= cls.property_tag_as_int() <= 0xFFFE:
                raise ValueError(
                    "'property_tag' value '%s' is reserved for custom properties" % cls.property_tag_as_hex()
                )

    @classmethod
    def _validate_property_name(cls):
        if cls.property_name:
            if any([cls.property_id, cls.property_tag]):
                raise ValueError("When 'property_name' is set, 'property_id' and 'property_tag' must be None")
            if not any([cls.distinguished_property_set_id, cls.property_set_id]):
                raise ValueError(
                    "When 'property_name' is set, 'distinguished_property_set_id' or 'property_set_id' must also be set"
                )

    @classmethod
    def _validate_property_id(cls):
        if cls.property_id:
            if any([cls.property_name, cls.property_tag]):
                raise ValueError("When 'property_id' is set, 'property_name' and 'property_tag' must be None")
            if not any([cls.distinguished_property_set_id, cls.property_set_id]):
                raise ValueError(
                    "When 'property_id' is set, 'distinguished_property_set_id' or 'property_set_id' must also be set"
                )

    @classmethod
    def _validate_property_type(cls):
        if cls.property_type not in cls.PROPERTY_TYPES:
            raise ValueError(
                "'property_type' value '%s' must be one of %s" % (cls.property_type, sorted(cls.PROPERTY_TYPES))
            )

    def clean(self, version=None):
        self.validate_cls()
        python_type = self.python_type()
        if self.is_array_type():
            if not is_iterable(self.value):
                raise ValueError("'%s' value %r must be a list" % (self.__class__.__name__, self.value))
            for v in self.value:
                if not isinstance(v, python_type):
                    raise ValueError("'value' element %r must be an instance of %s" % (v, python_type))
        else:
            if not isinstance(self.value, python_type):
                raise ValueError("'value' %r must be an instance of %s" % (self.value, python_type))

    @classmethod
    def from_xml(cls, elem, account):
        # Gets value of this specific ExtendedProperty from a list of 'ExtendedProperty' XML elements
        python_type = cls.python_type()
        if cls.is_array_type():
            values = elem.find('{%s}Values' % TNS)
            if cls.is_binary_type():
                return [base64.b64decode(val) for val in get_xml_attrs(values, '{%s}Value' % TNS)]
            return [
                xml_text_to_value(value=val, value_type=python_type)
                for val in get_xml_attrs(values, '{%s}Value' % TNS)
            ]
        if cls.is_binary_type():
            return base64.b64decode(get_xml_attr(elem, '{%s}Value' % TNS))
        extended_field_value = xml_text_to_value(value=get_xml_attr(elem, '{%s}Value' % TNS), value_type=python_type)
        if python_type == string_types[0] and not extended_field_value:
            # For string types, we want to return the empty string instead of None if the element was
            # actually found, but there was no XML value. For other types, it would be more problematic
            # to make that distinction, e.g. return False for bool, 0 for int, etc.
            return ''
        return extended_field_value

    def to_xml(self, version):
        if self.is_array_type():
            values = create_element('t:Values')
            for v in self.value:
                if self.is_binary_type():
                    add_xml_child(values, 't:Value', base64.b64encode(v).decode('ascii'))
                else:
                    add_xml_child(values, 't:Value', v)
            return values
        val = base64.b64encode(self.value).decode('ascii') if self.is_binary_type() else self.value
        return set_xml_value(create_element('t:Value'), val, version=version)

    @classmethod
    def is_array_type(cls):
        return cls.property_type.endswith('Array')

    @classmethod
    def is_binary_type(cls):
        # We can't just test python_type() == bytes, because str == bytes in Python2
        return 'Binary' in cls.property_type

    @classmethod
    def property_tag_as_int(cls):
        if isinstance(cls.property_tag, string_types):
            return int(cls.property_tag, base=16)
        return cls.property_tag

    @classmethod
    def property_tag_as_hex(cls):
        return hex(cls.property_tag) if isinstance(cls.property_tag, int) else cls.property_tag

    @classmethod
    def python_type(cls):
        # Return the best equivalent for a Python type for the property type of this class
        base_type = cls.property_type[:-5] if cls.is_array_type() else cls.property_type
        return {
            'ApplicationTime': Decimal,
            'Binary': bytes,
            'Boolean': bool,
            'CLSID': string_types[0],
            'Currency': int,
            'Double': Decimal,
            'Float': Decimal,
            'Integer': int,
            'Long': int,
            'Short': int,
            'SystemTime': EWSDateTime,
            'String': string_types[0],
        }[base_type]

    @classmethod
    def properties_map(cls):
        # EWS returns PropertySetId values in lowercase in XML
        return {
            'DistinguishedPropertySetId': cls.distinguished_property_set_id,
            'PropertySetId': cls.property_set_id.lower() if cls.property_set_id else None,
            'PropertyTag': cls.property_tag_as_hex(),
            'PropertyName': cls.property_name,
            'PropertyId': value_to_xml_text(cls.property_id) if cls.property_id else None,
            'PropertyType': cls.property_type,
        }


class ExternId(ExtendedProperty):
    # This is a custom extended property defined by us. It's useful for synchronization purposes, to attach a unique ID
    # from an external system. Strictly, this is an field that should probably not be registered by default since it's
    # not part of EWS, but it's been around since the beginning of this library and would be a pain for consumers to
    # register manually.

    property_set_id = 'c11ff724-aa03-4555-9952-8fa248a11c3e'  # This is arbitrary. We just want a unique UUID.
    property_name = 'External ID'
    property_type = 'String'

    __slots__ = ExtendedProperty.__slots__
