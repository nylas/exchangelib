# coding=utf-8
import base64
import logging

from future.utils import python_2_unicode_compatible
from six import string_types

from .util import create_element, xml_to_str, value_to_xml_text, is_iterable
from .version import EXCHANGE_2010

log = logging.getLogger(__name__)


@python_2_unicode_compatible
class Q(object):
    # Connection types
    AND = 'AND'
    OR = 'OR'
    NOT = 'NOT'
    CONN_TYPES = {AND, OR, NOT}

    # EWS Operators
    EQ = '=='
    NE = '!='
    GT = '>'
    GTE = '>='
    LT = '<'
    LTE = '<='
    EXACT = 'exact'
    IEXACT = 'iexact'
    CONTAINS = 'contains'
    ICONTAINS = 'icontains'
    STARTSWITH = 'startswith'
    ISTARTSWITH = 'istartswith'
    EXISTS = 'exists'
    OP_TYPES = {EQ, NE, GT, GTE, LT, LTE, EXACT, IEXACT, CONTAINS, ICONTAINS, STARTSWITH, ISTARTSWITH, EXISTS}
    CONTAINS_OPS = {EXACT, IEXACT, CONTAINS, ICONTAINS, STARTSWITH, ISTARTSWITH}

    # Valid lookups
    LOOKUP_RANGE = 'range'
    LOOKUP_IN = 'in'
    LOOKUP_NOT = 'not'
    LOOKUP_GT = 'gt'
    LOOKUP_GTE = 'gte'
    LOOKUP_LT = 'lt'
    LOOKUP_LTE = 'lte'
    LOOKUP_EXACT = 'exact'
    LOOKUP_IEXACT = 'iexact'
    LOOKUP_CONTAINS = 'contains'
    LOOKUP_ICONTAINS = 'icontains'
    LOOKUP_STARTSWITH = 'startswith'
    LOOKUP_ISTARTSWITH = 'istartswith'
    LOOKUP_EXISTS = 'exists'
    LOOKUP_TYPES = {LOOKUP_RANGE, LOOKUP_IN, LOOKUP_NOT, LOOKUP_GT, LOOKUP_GTE, LOOKUP_LT, LOOKUP_LTE, LOOKUP_EXACT,
                    LOOKUP_IEXACT, LOOKUP_CONTAINS, LOOKUP_ICONTAINS, LOOKUP_STARTSWITH, LOOKUP_ISTARTSWITH,
                    LOOKUP_EXISTS}

    __slots__ = ('conn_type', 'field_path', 'op', 'value', 'children', 'query_string')

    def __init__(self, *args, **kwargs):
        self.conn_type = kwargs.pop('conn_type', self.AND)

        self.field_path = None  # Name of the field we want to filter on
        self.op = None
        self.value = None
        self.query_string = None

        # Parsing of args and kwargs may require child elements
        self.children = []

        # Remove any empty Q elements in args before proceeding
        args = tuple(a for a in args if not (isinstance(a, self.__class__) and a.is_empty()))

        # Check for query string, or Q object containing query string, as the only argument
        if len(args) == 1 and not kwargs:
            if isinstance(args[0], string_types):
                self.query_string = args[0]
                return
            if isinstance(args[0], self.__class__) and args[0].query_string:
                self.query_string = args[0].query_string
                return

        # Parse args which must be Q objects
        for q in args:
            if not isinstance(q, self.__class__):
                raise ValueError("Non-keyword arg %r must be a Q object" % q)
            if q.query_string:
                raise ValueError(
                    'A query string cannot be combined with other restrictions (args: %r, kwargs: %r)' % (args, kwargs)
                )
            self.children.append(q)

        # Parse keyword args and extract the filter
        is_single_kwarg = len(args) == 0 and len(kwargs) == 1
        for key, value in kwargs.items():
            children = self._get_children_from_kwarg(key=key, value=value, is_single_kwarg=is_single_kwarg)
            self.children.extend(children)

        if len(self.children) == 1 and self.field_path is None and self.conn_type != self.NOT:
            # We only have one child and no expression on ourselves, so we are a no-op. Flatten by taking over the child
            self._promote()

    def _get_children_from_kwarg(self, key, value, is_single_kwarg=False):
        # Generates Q objects corresponding to a single keyword argument. Makes this a leaf if there are no children to
        # generate.
        key_parts = key.rsplit('__', 1)
        if len(key_parts) == 2 and key_parts[1] in self.LOOKUP_TYPES:
            # This is a kwarg with a lookup at the end
            field_path, lookup = key_parts
            if lookup == self.LOOKUP_EXISTS:
                # value=True will fall through to further processing
                if not value:
                    return [~self.__class__(**{key: True})]

            if lookup == self.LOOKUP_RANGE:
                # EWS doesn't have a 'range' operator. Emulate 'foo__range=(1, 2)' as 'foo__gte=1 and foo__lte=2'
                # (both values inclusive).
                if len(value) != 2:
                    raise ValueError("Value of lookup '%s' must have exactly 2 elements" % key)
                return [
                    self.__class__(**{'%s__gte' % field_path: value[0]}),
                    self.__class__(**{'%s__lte' % field_path: value[1]}),
                ]

            if lookup == self.LOOKUP_IN:
                # EWS doesn't have an '__in' operator. Allow '__in' lookups on list and non-list field types,
                # specifying a list value. We'll emulate it as a set of OR'ed exact matches.
                if not is_iterable(value, generators_allowed=True):
                    raise ValueError("Value for lookup %r must be a list" % key)
                children = [self.__class__(**{field_path: v}) for v in value]
                return [self.__class__(*children, conn_type=self.OR)]

            # Filtering on list types is a bit quirky. The only lookup type I have found to work is:
            #
            #     item:Categories == 'foo' AND item:Categories == 'bar' AND ...
            #
            #     item:Categories == 'foo' OR item:Categories == 'bar' OR ...
            #
            # The former returns items that have all these categories, but maybe also others. The latter returns
            # items that have at least one of these categories. This translates to the 'contains' and 'in' lookups.
            # Both versions are case-insensitive.
            #
            # Exact matching and case-sensitive or partial-string matching is not possible since that requires the
            # 'Contains' element which only supports matching on string elements, not arrays.
            #
            # Exact matching of categories (i.e. match ['a', 'b'] but not ['a', 'b', 'c']) could be implemented by
            # post-processing items by fetch the categories field unconditionally and removing the items that don't
            # have an exact match.
            if lookup == self.LOOKUP_CONTAINS and is_iterable(value, generators_allowed=True):
                # '__contains' lookups on list field types
                children = [self.__class__(**{field_path: v}) for v in value]
                return [self.__class__(*children, conn_type=self.AND)]

            try:
                op = self._lookup_to_op(lookup)
            except KeyError:
                raise ValueError("Lookup '%s' is not supported (called as '%s=%r')" % (lookup, key, value))
        else:
            field_path, op = key, self.EQ

        if not is_single_kwarg:
            return [self.__class__(**{key: value})]

        # This is a single-kwarg Q object with a lookup that requires a single value. Make this a leaf
        self.field_path = field_path
        self.op = op
        self.value = value
        return []

    def _promote(self):
        # Flatten by taking over the only child
        if len(self.children) != 1:
            raise ValueError('Can only flatten when child count is 1')
        if self.field_path is not None:
            raise ValueError("Can only flatten when 'field_path' is not set")
        q = self.children[0]
        self.conn_type = q.conn_type
        self.field_path = q.field_path
        self.op = q.op
        self.value = q.value
        self.query_string = q.query_string
        self.children = q.children

    def clean(self):
        # Do some basic checks on the attributes, using a generic folder and no Exchange version restrictions. to_xml()
        # does a really good job of validating. There's no reason to replicate much of that here.
        from .folders import Folder
        self.to_xml(folders=[Folder()], version=None)

    @classmethod
    def _lookup_to_op(cls, lookup):
        return {
            cls.LOOKUP_NOT: cls.NE,
            cls.LOOKUP_GT: cls.GT,
            cls.LOOKUP_GTE: cls.GTE,
            cls.LOOKUP_LT: cls.LT,
            cls.LOOKUP_LTE: cls.LTE,
            cls.LOOKUP_EXACT: cls.EXACT,
            cls.LOOKUP_IEXACT: cls.IEXACT,
            cls.LOOKUP_CONTAINS: cls.CONTAINS,
            cls.LOOKUP_ICONTAINS: cls.ICONTAINS,
            cls.LOOKUP_STARTSWITH: cls.STARTSWITH,
            cls.LOOKUP_ISTARTSWITH: cls.ISTARTSWITH,
            cls.LOOKUP_EXISTS: cls.EXISTS,
        }[lookup]

    @classmethod
    def _conn_to_xml(cls, conn_type):
        xml_tag_map = {
            cls.AND: 't:And',
            cls.OR: 't:Or',
            cls.NOT: 't:Not',
        }
        return create_element(xml_tag_map[conn_type])

    @classmethod
    def _op_to_xml(cls, op):
        xml_tag_map = {
            cls.EQ: 't:IsEqualTo',
            cls.NE: 't:IsNotEqualTo',
            cls.GTE: 't:IsGreaterThanOrEqualTo',
            cls.LTE: 't:IsLessThanOrEqualTo',
            cls.LT: 't:IsLessThan',
            cls.GT: 't:IsGreaterThan',
            cls.EXISTS: 't:Exists',
        }
        if op in xml_tag_map:
            return create_element(xml_tag_map[op])
        valid_ops = cls.EXACT, cls.IEXACT, cls.CONTAINS, cls.ICONTAINS, cls.STARTSWITH, cls.ISTARTSWITH
        if op not in valid_ops:
            raise ValueError("'op' %s must be one of %s" % (op, valid_ops))

        # For description of Contains attribute values, see
        #     https://msdn.microsoft.com/en-us/library/office/aa580702(v=exchg.150).aspx
        #
        # Possible ContainmentMode values:
        #     FullString, Prefixed, Substring, PrefixOnWords, ExactPhrase
        # Django lookups have no equivalent of PrefixOnWords and ExactPhrase (and I'm unsure how they actually
        # work).
        #
        # EWS has no equivalent of '__endswith' or '__iendswith'. That could be emulated using '__contains' and
        # '__icontains' and filtering results afterwards in Python. But it could be inefficient because we might be
        # fetching and discarding a lot of non-matching items, plus we would need to always fetch the field we're
        # matching on, to be able to do the filtering. I think it's better to leave this to the consumer, i.e.:
        #
        # items = [i for i in fld.filter(subject__contains=suffix) if i.subject.endswith(suffix)]
        # items = [i for i in fld.filter(subject__icontains=suffix) if i.subject.lower().endswith(suffix.lower())]
        #
        # Possible ContainmentComparison values (there are more, but the rest are "To be removed"):
        #     Exact, IgnoreCase, IgnoreNonSpacingCharacters, IgnoreCaseAndNonSpacingCharacters
        # I'm unsure about non-spacing characters, but as I read
        #    https://en.wikipedia.org/wiki/Graphic_character#Spacing_and_non-spacing_characters
        # we shouldn't ignore them ('a' would match both 'a' and 'å', the latter having a non-spacing character).
        if op in {cls.EXACT, cls.IEXACT}:
            match_mode = 'FullString'
        elif op in (cls.CONTAINS, cls.ICONTAINS):
            match_mode = 'Substring'
        elif op in (cls.STARTSWITH, cls.ISTARTSWITH):
            match_mode = 'Prefixed'
        else:
            raise ValueError('Unsupported op: %s' % op)
        if op in (cls.IEXACT, cls.ICONTAINS, cls.ISTARTSWITH):
            compare_mode = 'IgnoreCase'
        else:
            compare_mode = 'Exact'
        return create_element('t:Contains', ContainmentMode=match_mode, ContainmentComparison=compare_mode)

    def is_leaf(self):
        return not self.children

    def is_empty(self):
        return self.is_leaf() and self.field_path is None and self.query_string is None

    def expr(self):
        if self.is_empty():
            return None
        if self.query_string:
            return self.query_string
        if self.is_leaf():
            expr = '%s %s %s' % (self.field_path, self.op, repr(self.value))
        else:
            # Sort children by field name so we get stable output (for easier testing). Children should never be empty.
            expr = (' %s ' % (self.AND if self.conn_type == self.NOT else self.conn_type)).join(
                (c.expr() if c.is_leaf() or c.conn_type == self.NOT else '(%s)' % c.expr())
                for c in sorted(self.children, key=lambda i: i.field_path or '')
            )
        if self.conn_type == self.NOT:
            # Add the NOT operator. Put children in parens if there is more than one child.
            if self.is_leaf() or len(self.children) == 1:
                return self.conn_type + ' %s' % expr
            return self.conn_type + ' (%s)' % expr
        return expr

    def to_xml(self, folders, version):
        if self.query_string:
            if version.build < EXCHANGE_2010:
                raise NotImplementedError('QueryString filtering is only supported for Exchange 2010 servers and later')
            elem = create_element('m:QueryString')
            elem.text = self.query_string
            return elem
        # Translate this Q object to a valid Restriction XML tree
        elem = self.xml_elem(folders=folders, version=version)
        if elem is None:
            return None
        restriction = create_element('m:Restriction')
        restriction.append(elem)
        return restriction

    def _check_integrity(self):
        if self.is_empty():
            return
        if self.query_string:
            if any([self.field_path, self.op, self.value, self.children]):
                raise ValueError('Query strings cannot be combined with other settings')
            return
        if self.conn_type not in self.CONN_TYPES:
            raise ValueError("'conn_type' %s must be one of %s" % (self.conn_type, self.CONN_TYPES))
        if not self.is_leaf():
            return
        if not self.field_path:
            raise ValueError("'field_path' must be set")
        if self.op not in self.OP_TYPES:
            raise ValueError("'op' %s must be one of %s" % (self.op, self.OP_TYPES))
        if self.op == self.EXISTS:
            if self.value is not True:
                raise ValueError("'value' must be True when operator is EXISTS")
        if self.value is None:
            raise ValueError('Value for filter on field path "%s" cannot be None' % self.field_path)
        if is_iterable(self.value, generators_allowed=True):
            raise ValueError(
                'Value %r for filter on field path "%s" must be a single value' % (self.value, self.field_path)
            )

    def _validate_field_path(self, field_path, folder):
        from .indexed_properties import MultiFieldIndexedElement
        if field_path.field not in folder.allowed_fields():
            raise ValueError(
                "'%s' is not a valid field when filtering on %s" % (field_path.field.name, folder.__class__.__name__))
        if not field_path.field.is_searchable:
            raise ValueError("EWS does not support filtering on field '%s'" % field_path.field.name)
        if field_path.subfield and not field_path.subfield.is_searchable:
            raise ValueError("EWS does not support filtering on subfield '%s'" % field_path.subfield.name)
        if field_path.field.value_cls is not None and issubclass(field_path.field.value_cls, MultiFieldIndexedElement) \
                and not field_path.subfield:
            raise ValueError("Field path '%s' must contain a subfield" % self.field_path)

    def _get_field_path(self, folders):
        # Convert the string field path to a real FieldPath object. The path is validated using the given folders.
        from .fields import FieldPath
        for folder in folders:
            try:
                field_path = FieldPath.from_string(field_path=self.field_path, folder=folder)
                self._validate_field_path(field_path=field_path, folder=folder)
                break
            except ValueError:
                continue
        else:
            raise ValueError("Unknown fieldname '%s' on folders '%s'" % (self.field_path, folders))
        return field_path

    def _get_clean_value(self, field_path, version):
        if self.op == self.EXISTS:
            return None
        clean_field = field_path.subfield if (field_path.subfield and field_path.label) else field_path.field
        if clean_field.is_list:
            # With __contains, we allow filtering by only one value even though the field is a list type
            return clean_field.clean(value=[self.value], version=version)[0]
        else:
            return clean_field.clean(value=self.value, version=version)

    def xml_elem(self, folders, version):
        # Recursively build an XML tree structure of this Q object. If this is an empty leaf (the equivalent of Q()),
        # return None.
        from .indexed_properties import SingleFieldIndexedElement
        from .extended_properties import ExtendedProperty
        from .fields import GenericDateField
        # Don't check self.value just yet. We want to return error messages on the field path first, and then the value.
        # This is done in _get_field_path() and _get_clean_value(), respectively.
        self._check_integrity()
        if self.is_empty():
            return None
        if self.is_leaf():
            elem = self._op_to_xml(self.op)
            field_path = self._get_field_path(folders)
            clean_value = self._get_clean_value(field_path=field_path, version=version)
            if isinstance(field_path.field, GenericDateField):
                pass
            elif issubclass(field_path.field.value_cls, ExtendedProperty) and field_path.field.value_cls.is_binary_type():
                # We need to base64-encode binary data
                clean_value = base64.b64encode(clean_value.value).decode('ascii')
            elif issubclass(field_path.field.value_cls, SingleFieldIndexedElement) and not field_path.label:
                # We allow a filter shortcut of e.g. email_addresses__contains=EmailAddress(label='Foo', ...) instead of
                # email_addresses__Foo_email_address=.... Set FieldPath label now so we can generate the field_uri.
                field_path.label = clean_value.label
            elem.append(field_path.to_xml())
            constant = create_element('t:Constant')
            if self.op != self.EXISTS:
                # Use .set() to not fill up the create_element() cache with unique values
                constant.set('Value', value_to_xml_text(clean_value))
                if self.op in self.CONTAINS_OPS:
                    elem.append(constant)
                else:
                    uriorconst = create_element('t:FieldURIOrConstant')
                    uriorconst.append(constant)
                    elem.append(uriorconst)
        elif len(self.children) == 1:
            # We have only one child
            elem = self.children[0].xml_elem(folders=folders, version=version)
        else:
            # We have multiple children. If conn_type is NOT, then group children with AND. We'll add the NOT later
            elem = self._conn_to_xml(self.AND if self.conn_type == self.NOT else self.conn_type)
            # Sort children by field name so we get stable output (for easier testing). Children should never be empty
            for c in sorted(self.children, key=lambda i: i.field_path or ''):
                elem.append(c.xml_elem(folders=folders, version=version))
        if elem is None:
            return None  # Should not be necessary, but play safe
        if self.conn_type == self.NOT:
            # Encapsulate everything in the NOT element
            not_elem = self._conn_to_xml(self.conn_type)
            not_elem.append(elem)
            return not_elem
        return elem

    def __and__(self, other):
        # & operator. Return a new Q with two children and conn_type AND
        return self.__class__(self, other, conn_type=self.AND)

    def __or__(self, other):
        # | operator. Return a new Q with two children and conn_type OR
        return self.__class__(self, other, conn_type=self.OR)

    def __invert__(self):
        # ~ operator. If op has an inverse, change op. Else return a new Q with conn_type NOT
        if self.conn_type == self.NOT:
            # This is NOT NOT. Change to AND
            self.conn_type = self.AND
            if len(self.children) == 1 and self.field_path is None:
                self._promote()
            return self
        if self.is_leaf():
            if self.op == self.EQ:
                self.op = self.NE
                return self
            if self.op == self.NE:
                self.op = self.EQ
                return self
            if self.op == self.GT:
                self.op = self.LTE
                return self
            if self.op == self.GTE:
                self.op = self.LT
                return self
            if self.op == self.LT:
                self.op = self.GTE
                return self
            if self.op == self.LTE:
                self.op = self.GT
                return self
        return self.__class__(self, conn_type=self.NOT)

    def __eq__(self, other):
        return repr(self) == repr(other)

    def __hash__(self):
        return hash(repr(self))

    def __str__(self):
        return self.expr() or 'Q()'

    def __repr__(self):
        if self.is_leaf():
            if self.query_string:
                return self.__class__.__name__ + '(%s)' % repr(self.query_string)
            return self.__class__.__name__ + '(%s %s %s)' % (self.field_path, self.op, repr(self.value))
        sorted_children = tuple(sorted(self.children, key=lambda i: i.field_path or ''))
        if self.conn_type == self.NOT or len(self.children) > 1:
            return self.__class__.__name__ + repr((self.conn_type,) + sorted_children)
        return self.__class__.__name__ + repr(sorted_children)


@python_2_unicode_compatible
class Restriction(object):
    """
    Implements an EWS Restriction type.

    """

    def __init__(self, q, folders):
        if not isinstance(q, Q):
            raise ValueError("'q' value %r must be a Q instance" % q)
        if q.is_empty():
            raise ValueError("Q object must not be empty")
        from .folders import Folder
        for folder in folders:
            if not isinstance(folder, Folder):
                raise ValueError("'folder' value %r must be a Folder instance" % folder)
        self.q = q
        self.folders = folders

    def to_xml(self, version):
        return self.q.to_xml(folders=self.folders, version=version)

    def __str__(self):
        """
        Prints the XML syntax tree
        """
        return xml_to_str(self.to_xml(version=None))
