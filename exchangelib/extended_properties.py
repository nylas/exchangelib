import logging
from decimal import Decimal

from six import string_types

from .properties import EWSElement
from .services import TNS
from .util import create_element, add_xml_child, get_xml_attrs, get_xml_attr, set_xml_value, value_to_xml_text, \
    xml_text_to_value

string_type = string_types[0]
log = logging.getLogger(__name__)


class ExtendedProperty(EWSElement):
    """
    MSDN: https://msdn.microsoft.com/en-us/library/office/aa566405(v=exchg.150).aspx

    Property_* values: https://msdn.microsoft.com/en-us/library/office/aa564843(v=exchg.150).aspx
    """
    __metaclass__ = EWSElement

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
        # 'SystemTime',  # Not implemented yet
        # 'SystemTimeArray',  # Not implemented yet
        'String',
        'StringArray',
    }  # The commented-out types cannot be used for setting or getting (see docs) and are thus not very useful here

    distinguished_property_set_id = None
    property_set_id = None
    property_tag = None  # hex integer (e.g. 0x8000) or string ('0x8000')
    property_name = None
    property_id = None  # integer as hex-formatted int (e.g. 0x8000) or normal int (32768)
    property_type = None

    __slots__ = ('value',)

    def __init__(self, value):
        self.value = value

    def clean(self):
        if self.distinguished_property_set_id:
            assert not any([self.property_set_id, self.property_tag])
            assert any([self.property_id, self.property_name])
            assert self.distinguished_property_set_id in self.DISTINGUISHED_SETS
        if self.property_set_id:
            assert not any([self.distinguished_property_set_id, self.property_tag])
            assert any([self.property_id, self.property_name])
        if self.property_tag:
            assert not any([
                self.distinguished_property_set_id, self.property_set_id, self.property_name, self.property_id
            ])
            if 0x8000 <= self.property_tag_as_int() <= 0xFFFE:
                raise ValueError(
                    "'property_tag' value '%s' is reserved for custom properties" % self.property_tag_as_hex()
                )
        if self.property_name:
            assert not any([self.property_id, self.property_tag])
            assert any([self.distinguished_property_set_id, self.property_set_id])
        if self.property_id:
            assert not any([self.property_name, self.property_tag])
            assert any([self.distinguished_property_set_id, self.property_set_id])
        assert self.property_type in self.PROPERTY_TYPES

        python_type = self.python_type()
        if self.is_array_type():
            for v in self.value:
                assert isinstance(v, python_type)
        else:
            assert isinstance(self.value, python_type)

    @classmethod
    def is_array_type(cls):
        return cls.property_type.endswith('Array')

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
            'CLSID': string_type,
            'Currency': int,
            'Double': Decimal,
            'Float': Decimal,
            'Integer': int,
            'Long': int,
            'Short': int,
            # 'SystemTime': int,
            'String': string_type,
        }[base_type]

    def to_xml(self, version):
        if self.is_array_type():
            values = create_element('t:Values')
            for v in self.value:
                add_xml_child(values, 't:Value', v)
            return values
        else:
            value = create_element('t:Value')
            set_xml_value(value, self.value, version=version)
            return value

    @classmethod
    def from_xml(cls, elems):
        # Gets value of this specific ExtendedProperty from a list of 'ExtendedProperty' XML elements
        python_type = cls.python_type()
        extended_field_value = None
        for e in elems:
            extended_field_uri = e.find('{%s}ExtendedFieldURI' % TNS)
            match = True

            for k, v in (
                    ('DistinguishedPropertySetId', cls.distinguished_property_set_id),
                    ('PropertySetId', cls.property_set_id),
                    ('PropertyTag', cls.property_tag_as_hex()),
                    ('PropertyName', cls.property_name),
                    ('PropertyId', value_to_xml_text(cls.property_id) if cls.property_id else None),
                    ('PropertyType', cls.property_type),
            ):
                if extended_field_uri.get(k) != v:
                    match = False
                    break
            if match:
                if cls.is_array_type():
                    extended_field_value = [
                        xml_text_to_value(value=val, value_type=python_type)
                        for val in get_xml_attrs(e, '{%s}Value' % TNS)
                    ]
                else:
                    extended_field_value = xml_text_to_value(
                        value=get_xml_attr(e, '{%s}Value' % TNS), value_type=python_type)
                    if python_type == string_type and not extended_field_value:
                        # For string types, we want to return the empty string instead of None if the element was
                        # actually found, but there was no XML value. For other types, it would be more problematic
                        # to make that distinction, e.g. return False for bool, 0 for int, etc.
                        extended_field_value = ''
                break
        return extended_field_value


class ExternId(ExtendedProperty):
    # This is a custom extended property defined by us. It's useful for synchronization purposes, to attach a unique ID
    # from an external system. Strictly, this is an field that should probably not be registered by default since it's
    # not part of EWS, but it's been around since the beginning of this library and would be a pain for consumers to
    # register manually.

    property_set_id = 'c11ff724-aa03-4555-9952-8fa248a11c3e'  # This is arbitrary. We just want a unique UUID.
    property_name = 'External ID'
    property_type = 'String'

    __slots__ = ExtendedProperty.__slots__