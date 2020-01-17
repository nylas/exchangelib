from collections import namedtuple
from decimal import Decimal

from exchangelib import Version, EWSDateTime, EWSTimeZone, UTC
from exchangelib.errors import ErrorInvalidServerVersion
from exchangelib.extended_properties import ExternId
from exchangelib.fields import BooleanField, IntegerField, DecimalField, TextField, ChoiceField, DateTimeField, \
    Base64Field, TimeZoneField, ExtendedPropertyField, CharListField, Choice, DateField, EnumField, EnumListField, \
    CharField
from exchangelib.indexed_properties import SingleFieldIndexedElement
from exchangelib.version import EXCHANGE_2007, EXCHANGE_2010, EXCHANGE_2013
from exchangelib.util import to_xml, TNS

from .common import TimedTestCase


class FieldTest(TimedTestCase):
    def test_value_validation(self):
        field = TextField('foo', field_uri='bar', is_required=True, default=None)
        with self.assertRaises(ValueError) as e:
            field.clean(None)  # Must have a default value on None input
        self.assertEqual(str(e.exception), "'foo' is a required field with no default")

        field = TextField('foo', field_uri='bar', is_required=True, default='XXX')
        self.assertEqual(field.clean(None), 'XXX')

        field = CharListField('foo', field_uri='bar')
        with self.assertRaises(ValueError) as e:
            field.clean('XXX')  # Must be a list type
        self.assertEqual(str(e.exception), "Field 'foo' value 'XXX' must be a list")

        field = CharListField('foo', field_uri='bar')
        with self.assertRaises(TypeError) as e:
            field.clean([1, 2, 3])  # List items must be correct type
        self.assertEqual(str(e.exception), "Field 'foo' value 1 must be of type <class 'str'>")

        field = CharField('foo', field_uri='bar')
        with self.assertRaises(TypeError) as e:
            field.clean(1)  # Value must be correct type
        self.assertEqual(str(e.exception), "Field 'foo' value 1 must be of type <class 'str'>")
        with self.assertRaises(ValueError) as e:
            field.clean('X' * 256)  # Value length must be within max_length
        self.assertEqual(
            str(e.exception),
            "'foo' value 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
            "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
            "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' exceeds length 255"
        )

        field = DateTimeField('foo', field_uri='bar')
        with self.assertRaises(ValueError) as e:
            field.clean(EWSDateTime(2017, 1, 1))  # Datetime values must be timezone aware
        self.assertEqual(str(e.exception), "Value '2017-01-01 00:00:00' on field 'foo' must be timezone aware")

        field = ChoiceField('foo', field_uri='bar', choices=[Choice('foo'), Choice('bar')])
        with self.assertRaises(ValueError) as e:
            field.clean('XXX')  # Value must be a valid choice
        self.assertEqual(str(e.exception), "Invalid choice 'XXX' for field 'foo'. Valid choices are: foo, bar")

        # A few tests on extended properties that override base methods
        field = ExtendedPropertyField('foo', value_cls=ExternId, is_required=True)
        with self.assertRaises(ValueError) as e:
            field.clean(None)  # Value is required
        self.assertEqual(str(e.exception), "'foo' is a required field")
        with self.assertRaises(TypeError) as e:
            field.clean(123)  # Correct type is required
        self.assertEqual(str(e.exception), "'ExternId' value 123 must be an instance of <class 'str'>")
        self.assertEqual(field.clean('XXX'), 'XXX')  # We can clean a simple value and keep it as a simple value
        self.assertEqual(field.clean(ExternId('XXX')), ExternId('XXX'))  # We can clean an ExternId instance as well

        class ExternIdArray(ExternId):
            property_type = 'StringArray'

        field = ExtendedPropertyField('foo', value_cls=ExternIdArray, is_required=True)
        with self.assertRaises(ValueError)as e:
            field.clean(None)  # Value is required
        self.assertEqual(str(e.exception), "'foo' is a required field")
        with self.assertRaises(ValueError)as e:
            field.clean(123)  # Must be an iterable
        self.assertEqual(str(e.exception), "'ExternIdArray' value 123 must be a list")
        with self.assertRaises(TypeError) as e:
            field.clean([123])  # Correct type is required
        self.assertEqual(str(e.exception), "'ExternIdArray' value element 123 must be an instance of <class 'str'>")

        # Test min/max on IntegerField
        field = IntegerField('foo', field_uri='bar', min=5, max=10)
        with self.assertRaises(ValueError) as e:
            field.clean(2)
        self.assertEqual(str(e.exception), "Value 2 on field 'foo' must be greater than 5")
        with self.assertRaises(ValueError)as e:
            field.clean(12)
        self.assertEqual(str(e.exception), "Value 12 on field 'foo' must be less than 10")

        # Test min/max on DecimalField
        field = DecimalField('foo', field_uri='bar', min=5, max=10)
        with self.assertRaises(ValueError) as e:
            field.clean(Decimal(2))
        self.assertEqual(str(e.exception), "Value Decimal('2') on field 'foo' must be greater than 5")
        with self.assertRaises(ValueError)as e:
            field.clean(Decimal(12))
        self.assertEqual(str(e.exception), "Value Decimal('12') on field 'foo' must be less than 10")

        # Test enum validation
        field = EnumField('foo', field_uri='bar', enum=['a', 'b', 'c'])
        with self.assertRaises(ValueError)as e:
            field.clean(0)  # Enums start at 1
        self.assertEqual(str(e.exception), "Value 0 on field 'foo' must be greater than 1")
        with self.assertRaises(ValueError) as e:
            field.clean(4)  # Spills over list
        self.assertEqual(str(e.exception), "Value 4 on field 'foo' must be less than 3")
        with self.assertRaises(ValueError) as e:
            field.clean('d')  # Value not in enum
        self.assertEqual(str(e.exception), "Value 'd' on field 'foo' must be one of ['a', 'b', 'c']")

        # Test enum list validation
        field = EnumListField('foo', field_uri='bar', enum=['a', 'b', 'c'])
        with self.assertRaises(ValueError)as e:
            field.clean([])
        self.assertEqual(str(e.exception), "Value '[]' on field 'foo' must not be empty")
        with self.assertRaises(ValueError) as e:
            field.clean([0])
        self.assertEqual(str(e.exception), "Value 0 on field 'foo' must be greater than 1")
        with self.assertRaises(ValueError) as e:
            field.clean([1, 1])  # Values must be unique
        self.assertEqual(str(e.exception), "List entries '[1, 1]' on field 'foo' must be unique")
        with self.assertRaises(ValueError) as e:
            field.clean(['d'])
        self.assertEqual(str(e.exception), "List value 'd' on field 'foo' must be one of ['a', 'b', 'c']")

    def test_garbage_input(self):
        # Test that we can survive garbage input for common field types
        tz = EWSTimeZone.timezone('Europe/Copenhagen')
        account = namedtuple('Account', ['default_timezone'])(default_timezone=tz)
        payload = b'''\
<?xml version="1.0" encoding="utf-8"?>
<Envelope xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    <t:Item>
        <t:Foo>THIS_IS_GARBAGE</t:Foo>
    </t:Item>
</Envelope>'''
        elem = to_xml(payload).find('{%s}Item' % TNS)
        for field_cls in (Base64Field, BooleanField, IntegerField, DateField, DateTimeField, DecimalField):
            field = field_cls('foo', field_uri='item:Foo', is_required=True, default='DUMMY')
            self.assertEqual(field.from_xml(elem=elem, account=account), None)

        # Test MS timezones
        payload = b'''\
<?xml version="1.0" encoding="utf-8"?>
<Envelope xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    <t:Item>
        <t:Foo Id="THIS_IS_GARBAGE"></t:Foo>
    </t:Item>
</Envelope>'''
        elem = to_xml(payload).find('{%s}Item' % TNS)
        field = TimeZoneField('foo', field_uri='item:Foo', default='DUMMY')
        self.assertEqual(field.from_xml(elem=elem, account=account), None)

    def test_versioned_field(self):
        field = TextField('foo', field_uri='bar', supported_from=EXCHANGE_2010)
        with self.assertRaises(ErrorInvalidServerVersion):
            field.clean('baz', version=Version(EXCHANGE_2007))
        field.clean('baz', version=Version(EXCHANGE_2010))
        field.clean('baz', version=Version(EXCHANGE_2013))

    def test_versioned_choice(self):
        field = ChoiceField('foo', field_uri='bar', choices={
            Choice('c1'), Choice('c2', supported_from=EXCHANGE_2010)
        })
        with self.assertRaises(ValueError):
            field.clean('XXX')  # Value must be a valid choice
        field.clean('c2', version=None)
        with self.assertRaises(ErrorInvalidServerVersion):
            field.clean('c2', version=Version(EXCHANGE_2007))
        field.clean('c2', version=Version(EXCHANGE_2010))
        field.clean('c2', version=Version(EXCHANGE_2013))

    def test_naive_datetime(self):
        # Test that we can survive naive datetimes on a datetime field
        tz = EWSTimeZone.timezone('Europe/Copenhagen')
        account = namedtuple('Account', ['default_timezone'])(default_timezone=tz)
        default_value = tz.localize(EWSDateTime(2017, 1, 2, 3, 4))
        field = DateTimeField('foo', field_uri='item:DateTimeSent', default=default_value)

        # TZ-aware datetime string
        payload = b'''\
<?xml version="1.0" encoding="utf-8"?>
<Envelope xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    <t:Item>
        <t:DateTimeSent>2017-06-21T18:40:02Z</t:DateTimeSent>
    </t:Item>
</Envelope>'''
        elem = to_xml(payload).find('{%s}Item' % TNS)
        self.assertEqual(field.from_xml(elem=elem, account=account), UTC.localize(EWSDateTime(2017, 6, 21, 18, 40, 2)))

        # Naive datetime string is localized to tz of the account
        payload = b'''\
<?xml version="1.0" encoding="utf-8"?>
<Envelope xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    <t:Item>
        <t:DateTimeSent>2017-06-21T18:40:02</t:DateTimeSent>
    </t:Item>
</Envelope>'''
        elem = to_xml(payload).find('{%s}Item' % TNS)
        self.assertEqual(field.from_xml(elem=elem, account=account), tz.localize(EWSDateTime(2017, 6, 21, 18, 40, 2)))

        # Garbage string returns None
        payload = b'''\
<?xml version="1.0" encoding="utf-8"?>
<Envelope xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    <t:Item>
        <t:DateTimeSent>THIS_IS_GARBAGE</t:DateTimeSent>
    </t:Item>
</Envelope>'''
        elem = to_xml(payload).find('{%s}Item' % TNS)
        self.assertEqual(field.from_xml(elem=elem, account=account), None)

        # Element not found returns default value
        payload = b'''\
<?xml version="1.0" encoding="utf-8"?>
<Envelope xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    <t:Item>
    </t:Item>
</Envelope>'''
        elem = to_xml(payload).find('{%s}Item' % TNS)
        self.assertEqual(field.from_xml(elem=elem, account=account), default_value)

    def test_single_field_indexed_element(self):
        # A SingleFieldIndexedElement must have only one field defined
        class TestField(SingleFieldIndexedElement):
            FIELDS = [CharField('a'), CharField('b')]

        with self.assertRaises(ValueError):
            TestField.value_field()
