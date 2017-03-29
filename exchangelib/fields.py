import logging
from decimal import Decimal

from six import string_types

from .ewsdatetime import EWSDateTime
from .extended_properties import ExtendedProperty
from .services import TNS
from .util import create_element, get_xml_attrs, set_xml_value, value_to_xml_text, xml_text_to_value

string_type = string_types[0]
log = logging.getLogger(__name__)


class Field(object):
    """
    Holds information related to an item field
    """
    def __init__(self, name, value_cls, from_version=None, choices=None, default=None, is_list=False,
                 is_complex=False, is_required=False, is_required_after_save=False, is_read_only=False,
                 is_read_only_after_send=False):
        self.name = name
        self.value_cls = value_cls
        self.from_version = from_version
        self.choices = choices
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

    def clean(self, value):
        from .attachments import Attachment
        from .properties import Attendee, Choice, Mailbox
        if value is None:
            if self.is_required and self.default is None:
                raise ValueError("'%s' is a required field with no default" % self.name)
            if self.is_list and self.value_cls == Attachment:
                return []
            return self.default
        if self.value_cls == EWSDateTime and not getattr(value, 'tzinfo'):
            raise ValueError("Field '%s' must be timezone aware" % self.name)
        if self.value_cls == Choice and value not in self.choices:
            raise ValueError("Field '%s' value '%s' is not a valid choice (%s)" % (self.name, value, self.choices))

        # For value_cls that are subclasses of string types, convert simple string values to their subclass equivalent
        # (e.g. str to Body and str to Subject) so we can call value.clean()
        if issubclass(self.value_cls, string_types) and self.value_cls != string_type \
                and not isinstance(value, self.value_cls):
            value = self.value_cls(value)
        elif issubclass(self.value_cls, bytes) and self.value_cls != bytes \
                and not isinstance(value, self.value_cls):
            value = self.value_cls(value)
        elif self.value_cls == Mailbox:
            if self.is_list:
                value = [Mailbox(email_address=s) if isinstance(s, string_types) else s for s in value]
            elif isinstance(value, string_types):
                value = Mailbox(email_address=value)
        elif self.value_cls == Attendee:
            if self.is_list:
                value = [Attendee(mailbox=Mailbox(email_address=s), response_type='Accept')
                         if isinstance(s, string_types) else s for s in value]
            elif isinstance(value, string_types):
                value = Attendee(mailbox=Mailbox(email_address=value), response_type='Accept')

        if self.is_list:
            if not isinstance(value, (tuple, list)):
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

    def from_xml(self, elem):
        from .attachments import Attachment, FileAttachment, ItemAttachment
        from .properties import Body, HTMLBody, Content, Mailbox, EWSElement
        if self.is_list:
            iter_elem = elem.find(self.response_tag())
            if self.value_cls == string_type:
                if iter_elem is not None:
                    return get_xml_attrs(iter_elem, '{%s}String' % TNS)
            elif self.value_cls == Attachment:
                # Look for both FileAttachment and ItemAttachment
                if iter_elem is not None:
                    attachments = []
                    for att_type in (FileAttachment, ItemAttachment):
                        attachments.extend(
                            [att_type.from_xml(elem=e) for e in iter_elem.findall(att_type.response_tag())]
                        )
                    return attachments
            elif issubclass(self.value_cls, EWSElement):
                if iter_elem is not None:
                    return [self.value_cls.from_xml(elem=e) for e in iter_elem.findall(self.value_cls.response_tag())]
            else:
                assert False, 'Field %s type %s not supported' % (self.name, self.value_cls)
        else:
            field_elem = elem.find(self.response_tag())
            if issubclass(self.value_cls, (bool, int, Decimal, bytes, string_type, EWSDateTime)):
                val = None if field_elem is None else field_elem.text or None
                if val is not None:
                    if issubclass(self.value_cls, EWSDateTime) and not val.endswith('Z'):
                        # Sometimes, EWS will send timestamps without the 'Z' for UTC. It seems like the values are
                        # still UTC, so mark them as such so EWSDateTime can still interpret the timestamps.
                        val += 'Z'
                    try:
                        val = xml_text_to_value(value=val, value_type=self.value_cls)
                    except ValueError:
                        pass
                    except KeyError:
                        assert False, 'Field %s type %s not supported' % (self.name, self.value_cls)
                    if issubclass(self.value_cls, Body):
                        body_type = field_elem.get('BodyType')
                        try:
                            return {
                                Body.body_type: Body,
                                HTMLBody.body_type: HTMLBody,
                            }[body_type](val)
                        except KeyError:
                            assert False, "Unknown BodyType '%s'" % body_type
                    if issubclass(self.value_cls, Content):
                        return val.b64decode()
                    return val
            elif issubclass(self.value_cls, EWSElement):
                sub_elem = elem.find(self.response_tag())
                if sub_elem is not None:
                    if self.value_cls == Mailbox:
                        # We want the nested Mailbox, not the wrapper element
                        return self.value_cls.from_xml(elem=sub_elem.find(Mailbox.response_tag()))
                    else:
                        return self.value_cls.from_xml(elem=sub_elem)
            else:
                assert False, 'Field %s type %s not supported' % (self.name, self.value_cls)
        return self.default

    def to_xml(self, value, version):
        raise NotImplementedError()

    def field_uri_xml(self):
        raise NotImplementedError()

    def request_tag(self):
        raise NotImplementedError()

    def response_tag(self):
        raise NotImplementedError()

    def __eq__(self, other):
        return hash(self) == hash(other)

    def __hash__(self):
        raise NotImplementedError()

    def __repr__(self):
        return self.__class__.__name__ + repr((self.name, self.value_cls))


class SimpleField(Field):
    def __init__(self, *args, **kwargs):
        field_uri = kwargs.pop('field_uri')
        super(SimpleField, self).__init__(*args, **kwargs)
        # See all valid FieldURI values at https://msdn.microsoft.com/en-us/library/office/aa494315(v=exchg.150).aspx
        # field_uri_prefix is the prefix part of the FieldURI.
        self.field_uri = field_uri
        if ':' in field_uri:
            self.field_uri_prefix, self.field_uri_postfix = field_uri.split(':')
        else:
            self.field_uri_prefix, self.field_uri_postfix = None, field_uri

    def to_xml(self, value, version):
        from .properties import Body, HTMLBody, Content
        field_elem = create_element(self.request_tag())
        if issubclass(self.value_cls, Body):
            body_type = {
                Body: Body.body_type,
                HTMLBody: HTMLBody.body_type,
            }[type(value)]
            field_elem.set('BodyType', body_type)
        if issubclass(self.value_cls, Content):
            value = value.b64encode()
        return set_xml_value(field_elem, value, version=version)

    def field_uri_xml(self):
        return create_element('t:FieldURI', FieldURI=self.field_uri)

    def request_tag(self):
        return 't:%s' % self.field_uri_postfix

    def response_tag(self):
        return '{%s}%s' % (TNS, self.field_uri_postfix)

    def __hash__(self):
        return hash(self.field_uri)


class IndexedField(SimpleField):
    PARENT_ELEMENT_NAME = None

    def __init__(self, *args, **kwargs):
        from .indexed_properties import IndexedElement
        super(IndexedField, self).__init__(*args, **kwargs)
        assert issubclass(self.value_cls, IndexedElement)

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

    def to_xml(self, value, version):
        return set_xml_value(create_element('t:%s' % self.PARENT_ELEMENT_NAME), value, version)

    @classmethod
    def response_tag(cls):
        return '{%s}%s' % (TNS, cls.PARENT_ELEMENT_NAME)

    def __hash__(self):
        return hash(self.field_uri)


class EmailAddressField(IndexedField):
    PARENT_ELEMENT_NAME = 'EmailAddresses'


class PhoneNumberField(IndexedField):
    PARENT_ELEMENT_NAME = 'PhoneNumbers'


class PhysicalAddressField(IndexedField):
    # MSDN: https://msdn.microsoft.com/en-us/library/office/aa564323(v=exchg.150).aspx
    PARENT_ELEMENT_NAME = 'PhysicalAddresses'


class ExtendedPropertyField(Field):
    def __init__(self, *args, **kwargs):
        super(ExtendedPropertyField, self).__init__(*args, **kwargs)
        assert issubclass(self.value_cls, ExtendedProperty)

    def clean(self, value):
        if value is None:
            if self.is_required:
                raise ValueError("'%s' is a required field" % self.name)
            return self.default
        if not isinstance(value, self.value_cls):
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
            hex_val = int(cls.property_tag, base=16) if isinstance(cls.property_tag, string_types) else cls.property_tag
            elem.set('PropertyTag', hex(hex_val))
        if cls.property_name:
            elem.set('PropertyName', cls.property_name)
        if cls.property_id:
            elem.set('PropertyId', value_to_xml_text(cls.property_id))
        elem.set('PropertyType', cls.property_type)
        return elem

    def from_xml(self, elem):
        extended_properties = elem.findall(self.response_tag())
        return self.value_cls.from_xml(elems=extended_properties)

    def to_xml(self, value, version):
        extended_property = create_element(self.request_tag())
        set_xml_value(extended_property, self.field_uri_xml(), version=version)
        if isinstance(value, self.value_cls):
            set_xml_value(extended_property, value, version=version)
        else:
            # Allow keeping ExtendedProperty field values as their simple Python type
            set_xml_value(extended_property, self.value_cls(value), version=version)
        return extended_property

    def request_tag(self):
        return 't:%s' % ExtendedProperty.ELEMENT_NAME

    @classmethod
    def response_tag(cls):
        return '{%s}%s' % (TNS, ExtendedProperty.ELEMENT_NAME)

    def __hash__(self):
        return hash(self.name)


class LabelField(SimpleField):
    def from_xml(self, elem):
        return elem.get(self.field_uri)


class SubField(SimpleField):
    def __init__(self, *args, **kwargs):
        kwargs['field_uri'] = ''
        super(SubField, self).__init__(*args, **kwargs)

    def from_xml(self, elem):
        return elem.text

    def to_xml(self, value, version):
        return value


class EmailSubField(SubField):
    def from_xml(self, elem):
        return elem.text or elem.get('Name')  # Sometimes elem.text is empty. Exchange saves the same in 'Name' attr


class ItemField(SimpleField):
    def __init__(self, *args, **kwargs):
        kwargs['value_cls'] = None
        super(ItemField, self).__init__(*args, **kwargs)

    @property
    def value_cls(self):
        from .items import Item
        return Item

    @value_cls.setter
    def value_cls(self, value):
        pass

    def from_xml(self, elem):
        from .items import ITEM_CLASSES
        for item_cls in ITEM_CLASSES:
            item_elem = elem.find(item_cls.response_tag())
            if item_elem is not None:
                return item_cls.from_xml(elem=item_elem)

    def to_xml(self, value, version):
        # We don't want to wrap in an Item element
        return value.to_xml(version=version)
