import abc
from decimal import Decimal
import logging

from six import string_types

from .ewsdatetime import EWSDateTime
from .services import TNS
from .util import create_element, get_xml_attrs, set_xml_value, value_to_xml_text, xml_text_to_value

string_type = string_types[0]
log = logging.getLogger(__name__)


class Field(object):
    """
    Holds information related to an item field
    """
    __metaclass__ = abc.ABCMeta
    value_cls = None

    def __init__(self, name, is_list=False, is_complex=False, is_required=False, is_required_after_save=False,
                 is_read_only=False, is_read_only_after_send=False, default=None):
        self.name = name
        self.default = default  # Default value if none is given
        self.is_list = is_list
        # Is the field a complex EWS type? Quoting the EWS FindItem docs:
        #
        #   The FindItem operation returns only the first 512 bytes of any streamable property. For Unicode, it returns
        #   the first 255 characters by using a null-terminated Unicode string. It does not return any of the message
        #   body formats or the recipient lists.
        #
        self.is_complex = is_complex
        self.is_required = is_required
        # Some fields cannot be deleted on update. Default to True if 'is_required' is set
        self.is_required_after_save = is_required or is_required_after_save
        self.is_read_only = is_read_only
        # Set this for fields that raise ErrorInvalidPropertyUpdateSentMessage on update after send. Default to True
        # if 'is_read_only' is set
        self.is_read_only_after_send = is_read_only or is_read_only_after_send

    @abc.abstractmethod
    def clean(self, value):
        if value is None:
            if self.is_required and self.default is None:
                raise ValueError("'%s' is a required field with no default" % self.name)
            return self.default
        if self.is_list:
            if not isinstance(value, (tuple, list, set)):
                raise ValueError("Field '%s' value '%s' must be a list" % (self.name, value))
            for v in value:
                if not isinstance(v, self.value_cls):
                    raise TypeError('Field %s value "%s" must be of type %s' % (self.name, v, self.value_cls))
                if hasattr(v, 'clean'):
                    v.clean()
        else:
            if not isinstance(value, self.value_cls):
                raise ValueError("Field '%s' value '%s' must be of type %s" % (self.name, value, self.value_cls))
            if hasattr(value, 'clean'):
                value.clean()
        return value

    @abc.abstractmethod
    def from_xml(self, elem):
        raise NotImplementedError()

    @abc.abstractmethod
    def to_xml(self, value, version):
        raise NotImplementedError()

    def __eq__(self, other):
        return hash(self) == hash(other)

    def __hash__(self):
        raise NotImplementedError()

    def __repr__(self):
        return self.__class__.__name__ + repr((self.name, self.value_cls))


class FieldURIField(Field):
    def __init__(self, *args, **kwargs):
        self.field_uri = kwargs.pop('field_uri', None)
        super(FieldURIField, self).__init__(*args, **kwargs)
        # See all valid FieldURI values at https://msdn.microsoft.com/en-us/library/office/aa494315(v=exchg.150).aspx
        # The field_uri has a prefix when the FieldURI points to an Item field.
        if self.field_uri is None:
            self.field_uri_postfix = None
        elif ':' in self.field_uri:
            self.field_uri_postfix = self.field_uri.split(':')[1]
        else:
            self.field_uri_postfix = self.field_uri

    def to_xml(self, value, version):
        field_elem = create_element(self.request_tag())
        return set_xml_value(field_elem, value, version=version)

    def field_uri_xml(self):
        return create_element('t:FieldURI', FieldURI=self.field_uri)

    def request_tag(self):
        assert self.field_uri_postfix
        return 't:%s' % self.field_uri_postfix

    def response_tag(self):
        assert self.field_uri_postfix
        return '{%s}%s' % (TNS, self.field_uri_postfix)

    def __hash__(self):
        return hash(self.field_uri)


class BooleanField(FieldURIField):
    value_cls = bool

    def clean(self, value):
        assert not self.is_list
        return super(BooleanField, self).clean(value)

    def from_xml(self, elem):
        field_elem = elem.find(self.response_tag())
        val = None if field_elem is None else field_elem.text or None
        if val is not None:
            try:
                val = xml_text_to_value(value=val, value_type=self.value_cls)
            except ValueError:
                pass
            return val
        return self.default


class IntegerField(FieldURIField):
    value_cls = int

    def clean(self, value):
        assert not self.is_list
        return super(IntegerField, self).clean(value)

    def from_xml(self, elem):
        field_elem = elem.find(self.response_tag())
        val = None if field_elem is None else field_elem.text or None
        if val is not None:
            try:
                val = xml_text_to_value(value=val, value_type=self.value_cls)
            except ValueError:
                pass
            return val
        return self.default


class DecimalField(FieldURIField):
    value_cls = Decimal

    def clean(self, value):
        assert not self.is_list
        return super(DecimalField, self).clean(value)

    def from_xml(self, elem):
        field_elem = elem.find(self.response_tag())
        val = None if field_elem is None else field_elem.text or None
        if val is not None:
            try:
                val = xml_text_to_value(value=val, value_type=self.value_cls)
            except ValueError:
                pass
            return val
        return self.default


class Base64Field(FieldURIField):
    value_cls = bytes

    def clean(self, value):
        assert not self.is_list
        return super(Base64Field, self).clean(value)

    def from_xml(self, elem):
        field_elem = elem.find(self.response_tag())
        val = None if field_elem is None else field_elem.text or None
        if val is not None:
            try:
                val = xml_text_to_value(value=val, value_type=self.value_cls)
            except ValueError:
                pass
            return val
        return self.default

    def to_xml(self, value, version):
        field_elem = create_element(self.request_tag())
        return set_xml_value(field_elem, value, version=version)


class DateTimeField(FieldURIField):
    value_cls = EWSDateTime

    def clean(self, value):
        assert not self.is_list
        if value is not None and isinstance(value, EWSDateTime) and not getattr(value, 'tzinfo'):
            raise ValueError("Field '%s' must be timezone aware" % self.name)
        return super(DateTimeField, self).clean(value)

    def from_xml(self, elem):
        field_elem = elem.find(self.response_tag())
        val = None if field_elem is None else field_elem.text or None
        if val is not None:
            if not val.endswith('Z'):
                # Sometimes, EWS will send timestamps without the 'Z' for UTC. It seems like the values are
                # still UTC, so mark them as such so EWSDateTime can still interpret the timestamps.
                val += 'Z'
            try:
                return xml_text_to_value(value=val, value_type=self.value_cls)
            except ValueError:
                pass
        return self.default


class TextField(FieldURIField):
    value_cls = string_type

    def __init__(self, *args, **kwargs):
        self.max_length = kwargs.pop('max_length', None)
        super(TextField, self).__init__(*args, **kwargs)

    def clean(self, value):
        value = super(TextField, self).clean(value)
        if value is not None and self.max_length and len(value) > self.max_length:
            raise ValueError("'%s' value '%s' exceeds length %s" % (self.name, value, self.max_length))
        return value

    def from_xml(self, elem):
        if self.is_list:
            iter_elem = elem.find(self.response_tag())
            if iter_elem is not None:
                return get_xml_attrs(iter_elem, '{%s}String' % TNS)
        else:
            field_elem = elem.find(self.response_tag())
            val = None if field_elem is None else field_elem.text or None
            if val is not None:
                return val
        return self.default


class URIField(TextField):
    # Helper to mark strings that must conform to xsd:anyURI
    # If we want an URI validator, see http://stackoverflow.com/questions/14466585/is-this-regex-correct-for-xsdanyuri
    def clean(self, value):
        assert not self.is_list
        return super(URIField, self).clean(value)


class EmailField(TextField):
    # A helper class used for email address string
    def clean(self, value):
        assert not self.is_list
        return super(EmailField, self).clean(value)


class ChoiceField(TextField):
    def __init__(self, *args, **kwargs):
        self.choices = kwargs.pop('choices')
        super(ChoiceField, self).__init__(*args, **kwargs)

    def clean(self, value):
        assert not self.is_list
        value = super(ChoiceField, self).clean(value)
        if value is not None and value not in self.choices:
            raise ValueError("Field '%s' value '%s' is not a valid choice (%s)" % (self.name, value, self.choices))
        return value


class BodyField(TextField):
    def __init__(self, *args, **kwargs):
        from .properties import Body
        self.value_cls = Body
        kwargs['is_complex'] = True
        super(BodyField, self).__init__(*args, **kwargs)

    def clean(self, value):
        assert not self.is_list
        if value is not None and not isinstance(value, self.value_cls):
            value = self.value_cls(value)
        return super(BodyField, self).clean(value)

    def from_xml(self, elem):
        from .properties import Body, HTMLBody
        field_elem = elem.find(self.response_tag())
        val = None if field_elem is None else field_elem.text or None
        if val is not None:
            body_type = field_elem.get('BodyType')
            return {
                Body.body_type: Body,
                HTMLBody.body_type: HTMLBody,
            }[body_type](val)
        return self.default

    def to_xml(self, value, version):
        from .properties import Body, HTMLBody
        field_elem = create_element(self.request_tag())
        body_type = {
            Body: Body.body_type,
            HTMLBody: HTMLBody.body_type,
        }[type(value)]
        field_elem.set('BodyType', body_type)
        return set_xml_value(field_elem, value, version=version)


class EWSElementField(FieldURIField):
    def __init__(self, *args, **kwargs):
        self.value_cls = kwargs.pop('value_cls')
        super(EWSElementField, self).__init__(*args, **kwargs)

    def clean(self, value):
        return super(EWSElementField, self).clean(value)

    def from_xml(self, elem):
        if self.is_list:
            iter_elem = elem.find(self.response_tag())
            if iter_elem is not None:
                return [self.value_cls.from_xml(elem=e) for e in iter_elem.findall(self.value_cls.response_tag())]
        else:
            if self.field_uri is None:
                sub_elem = elem.find(self.value_cls.response_tag())
            else:
                sub_elem = elem.find(self.response_tag())
            if sub_elem is not None:
                return self.value_cls.from_xml(elem=sub_elem)
        return self.default

    def to_xml(self, value, version):
        if self.field_uri is None:
            return value.to_xml(version=version)
        field_elem = create_element(self.request_tag())
        return set_xml_value(field_elem, value, version=version)


class MailboxField(EWSElementField):
    def __init__(self, *args, **kwargs):
        from .properties import Mailbox
        kwargs['value_cls'] = Mailbox
        super(MailboxField, self).__init__(*args, **kwargs)

    def clean(self, value):
        if value is not None:
            if self.is_list:
                value = [self.value_cls(email_address=s) if isinstance(s, string_types) else s for s in value]
            elif isinstance(value, string_types):
                value = self.value_cls(email_address=value)
        return super(MailboxField, self).clean(value)

    def from_xml(self, elem):
        if self.is_list:
            iter_elem = elem.find(self.response_tag())
            if iter_elem is not None:
                return [self.value_cls.from_xml(elem=e) for e in iter_elem.findall(self.value_cls.response_tag())]
        else:
            if self.field_uri is None:
                sub_elem = elem.find(self.value_cls.response_tag())
            else:
                sub_elem = elem.find(self.response_tag())
            if sub_elem is not None:
                if self.field_uri is not None:
                    # We want the nested Mailbox, not the wrapper element
                    return self.value_cls.from_xml(elem=sub_elem.find(self.value_cls.response_tag()))
                else:
                    return self.value_cls.from_xml(elem=sub_elem)
        return self.default


class AttendeesField(EWSElementField):
    def __init__(self, *args, **kwargs):
        from .properties import Attendee
        kwargs['value_cls'] = Attendee
        kwargs['is_list'] = True
        super(AttendeesField, self).__init__(*args, **kwargs)

    def clean(self, value):
        from .properties import Mailbox
        if value is not None:
            value = [self.value_cls(mailbox=Mailbox(email_address=s), response_type='Accept')
                     if isinstance(s, string_types) else s for s in value]
        return super(AttendeesField, self).clean(value)

    def from_xml(self, elem):
        iter_elem = elem.find(self.response_tag())
        if iter_elem is not None:
            return [self.value_cls.from_xml(elem=e) for e in iter_elem.findall(self.value_cls.response_tag())]
        return self.default


class AttachmentField(EWSElementField):
    def __init__(self, *args, **kwargs):
        from .attachments import Attachment
        kwargs['value_cls'] = Attachment
        kwargs['is_list'] = True
        kwargs['is_complex'] = True
        super(AttachmentField, self).__init__(*args, **kwargs)

    def from_xml(self, elem):
        from .attachments import FileAttachment, ItemAttachment
        iter_elem = elem.find(self.response_tag())
        # Look for both FileAttachment and ItemAttachment
        if iter_elem is not None:
            attachments = []
            for att_type in (FileAttachment, ItemAttachment):
                attachments.extend(
                    [att_type.from_xml(elem=e) for e in iter_elem.findall(att_type.response_tag())]
                )
            return attachments
        return self.default


class IndexedField(FieldURIField):
    PARENT_ELEMENT_NAME = None

    def __init__(self, *args, **kwargs):
        self.value_cls = kwargs.pop('value_cls')
        super(IndexedField, self).__init__(*args, **kwargs)

    def clean(self, value):
        return super(IndexedField, self).clean(value)

    def from_xml(self, elem):
        if self.is_list:
            iter_elem = elem.find(self.response_tag())
            if iter_elem is not None:
                return [self.value_cls.from_xml(elem=e) for e in iter_elem.findall(self.value_cls.response_tag())]
        else:
            sub_elem = elem.find(self.response_tag())
            if sub_elem is not None:
                return self.value_cls.from_xml(elem=sub_elem)
        return self.default

    def to_xml(self, value, version):
        return set_xml_value(create_element('t:%s' % self.PARENT_ELEMENT_NAME), value, version)

    def field_uri_xml(self, label=None, subfield=None):
        from .indexed_properties import MultiFieldIndexedElement
        if not label:
            # Return elements for all labels
            elems = []
            for l in self.value_cls.LABELS:
                elem = self.field_uri_xml(label=l)
                if isinstance(elem, list):
                    elems.extend(elem)
                else:
                    elems.append(elem)
            return elems
        if issubclass(self.value_cls, MultiFieldIndexedElement):
            if not subfield:
                # Return elements for all sub-fields
                return [self.field_uri_xml(label=label, subfield=f) for f in self.value_cls.FIELDS]
            assert subfield in self.value_cls.FIELDS
            field_uri = '%s:%s' % (self.field_uri, subfield.field_uri)
        else:
            field_uri = self.field_uri
        assert label in self.value_cls.LABELS, (label, self.value_cls.LABELS)
        return create_element('t:IndexedFieldURI', FieldURI=field_uri, FieldIndex=label)

    @classmethod
    def response_tag(cls):
        return '{%s}%s' % (TNS, cls.PARENT_ELEMENT_NAME)

    def __hash__(self):
        return hash(self.field_uri)


class EmailAddressField(IndexedField):
    PARENT_ELEMENT_NAME = 'EmailAddresses'

    def __init__(self, *args, **kwargs):
        from .indexed_properties import EmailAddress
        kwargs['value_cls'] = EmailAddress
        kwargs['is_list'] = True
        super(EmailAddressField, self).__init__(*args, **kwargs)


class PhoneNumberField(IndexedField):
    PARENT_ELEMENT_NAME = 'PhoneNumbers'

    def __init__(self, *args, **kwargs):
        from .indexed_properties import PhoneNumber
        kwargs['value_cls'] = PhoneNumber
        kwargs['is_list'] = True
        super(PhoneNumberField, self).__init__(*args, **kwargs)


class PhysicalAddressField(IndexedField):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa564323(v=exchg.150).aspx
    PARENT_ELEMENT_NAME = 'PhysicalAddresses'

    def __init__(self, *args, **kwargs):
        from .indexed_properties import PhysicalAddress
        kwargs['value_cls'] = PhysicalAddress
        kwargs['is_list'] = True
        super(PhysicalAddressField, self).__init__(*args, **kwargs)


class ExtendedPropertyField(Field):
    def __init__(self, *args, **kwargs):
        self.value_cls = kwargs.pop('value_cls')
        super(ExtendedPropertyField, self).__init__(*args, **kwargs)

    def clean(self, value):
        assert not self.is_list
        if value is None:
            if self.is_required:
                raise ValueError("'%s' is a required field" % self.name)
            return self.default
        elif not isinstance(value, self.value_cls):
            # Allow keeping ExtendedProperty field values as their simple Python type, but run clean() anyway
            tmp = self.value_cls(value)
            tmp.clean()
            return value
        value.clean()
        return value

    def field_uri_xml(self):
        elem = create_element('t:ExtendedFieldURI')
        cls = self.value_cls
        if cls.distinguished_property_set_id:
            elem.set('DistinguishedPropertySetId', cls.distinguished_property_set_id)
        if cls.property_set_id:
            elem.set('PropertySetId', cls.property_set_id)
        if cls.property_tag:
            elem.set('PropertyTag', cls.property_tag_as_hex())
        if cls.property_name:
            elem.set('PropertyName', cls.property_name)
        if cls.property_id:
            elem.set('PropertyId', value_to_xml_text(cls.property_id))
        elem.set('PropertyType', cls.property_type)
        return elem

    def from_xml(self, elem):
        extended_properties = elem.findall(self.value_cls.response_tag())
        for extended_property in extended_properties:
            extended_field_uri = extended_property.find('{%s}ExtendedFieldURI' % TNS)
            match = True
            for k, v in self.value_cls.properties_map().items():
                if extended_field_uri.get(k) != v:
                    match = False
                    break
            if match:
                return self.value_cls.from_xml(elem=extended_property)
        return self.default

    def to_xml(self, value, version):
        extended_property = create_element(self.value_cls.request_tag())
        set_xml_value(extended_property, self.field_uri_xml(), version=version)
        if isinstance(value, self.value_cls):
            set_xml_value(extended_property, value, version=version)
        else:
            # Allow keeping ExtendedProperty field values as their simple Python type
            set_xml_value(extended_property, self.value_cls(value), version=version)
        return extended_property

    def __hash__(self):
        return hash(self.name)


class LabelField(ChoiceField):
    def from_xml(self, elem):
        return elem.get(self.field_uri)


class SubField(Field):
    value_cls = string_type

    def clean(self, value):
        assert not self.is_list
        return super(SubField, self).clean(value)

    def from_xml(self, elem):
        return elem.text

    def to_xml(self, value, version):
        return value

    def __hash__(self):
        return hash(self.name)


class EmailSubField(SubField):
    value_cls = string_type

    def from_xml(self, elem):
        return elem.text or elem.get('Name')  # Sometimes elem.text is empty. Exchange saves the same in 'Name' attr


class ItemField(FieldURIField):
    def __init__(self, *args, **kwargs):
        super(ItemField, self).__init__(*args, **kwargs)

    @property
    def value_cls(self):
        # This is a workaround for circular imports. Item
        from .items import Item
        return Item

    def clean(self, value):
        assert not self.is_list
        return super(ItemField, self).clean(value)

    def from_xml(self, elem):
        from .items import ITEM_CLASSES
        for item_cls in ITEM_CLASSES:
            item_elem = elem.find(item_cls.response_tag())
            if item_elem is not None:
                return item_cls.from_xml(elem=item_elem)

    def to_xml(self, value, version):
        # We don't want to wrap in an Item element
        return value.to_xml(version=version)
