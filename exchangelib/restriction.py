# coding=utf-8
import logging
from threading import Lock

from future.utils import python_2_unicode_compatible

from .ewsdatetime import EWSDateTime, UTC
from .util import create_element, xml_to_str, value_to_xml_text

log = logging.getLogger(__name__)

_source_cache = dict()
_source_cache_lock = Lock()


@python_2_unicode_compatible
class Q(object):
    # Connection types
    AND = 'AND'
    OR = 'OR'
    NOT = 'NOT'
    CONN_TYPES = (AND, OR, NOT)

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
    OP_TYPES = (EQ, NE, GT, GTE, LT, LTE, EXACT, IEXACT, CONTAINS, ICONTAINS, STARTSWITH, ISTARTSWITH, EXISTS)
    CONTAINS_OPS = (EXACT, IEXACT, CONTAINS, ICONTAINS, STARTSWITH, ISTARTSWITH)

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

    __slots__ = 'conn_type', 'fieldname', 'op', 'value', 'children'

    def __init__(self, *args, **kwargs):
        self.conn_type = kwargs.pop('conn_type', self.AND)

        self.fieldname = None  # Name of the field we want to filter on
        self.op = None
        self.value = None

        # Build children of Q objects from *args and **kwargs
        self.children = []

        for q in args:
            if not isinstance(q, self.__class__):
                raise ValueError("Non-keyword arg '%s' must be a Q object" % q)
            if not q.is_empty():
                self.children.append(q)

        for key, value in kwargs.items():
            if '__' in key:
                fieldname, lookup = key.rsplit('__')
                if lookup == self.LOOKUP_EXISTS:
                    # value=True will fall through to further processing
                    if not value:
                        self.children.append(~self.__class__(**{key: True}))
                        continue

                if lookup == self.LOOKUP_RANGE:
                    # EWS doesn't have a 'range' operator. Emulate 'foo__range=(1, 2)' as 'foo__gte=1 and foo__lte=2'
                    # (both values inclusive).
                    if len(value) != 2:
                        raise ValueError("Value of lookup '%s' must have exactly 2 elements" % key)
                    self.children.append(self.__class__(**{'%s__gte' % fieldname: value[0]}))
                    self.children.append(self.__class__(**{'%s__lte' % fieldname: value[1]}))
                    continue

                if lookup == self.LOOKUP_IN:
                    # Allow '__in' lookups on list and non-list field types, specifying a list
                    if not isinstance(value, (tuple, list, set)):
                        raise ValueError("Value for lookup '%s' must be a list" % key)
                    children = [self.__class__(**{fieldname: v}) for v in value]
                    self.children.append(self.__class__(*children, conn_type=self.OR))
                    continue

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
                # post-processing items. Fetch 'item:Categories' with additional_fields and remove the items that don't
                # have an exact match, after the call to FindItems.
                if lookup == self.LOOKUP_CONTAINS and isinstance(value, (tuple, list, set)):
                    # '__contains' lookups on list field types
                    children = [self.__class__(**{fieldname: v}) for v in value]
                    self.children.append(self.__class__(*children, conn_type=self.AND))
                    continue
                try:
                    op = self._lookup_to_op(lookup)
                except KeyError:
                    raise ValueError("Lookup '%s' is not supported (called as '%s=%r')" % (lookup, key, value))
            else:
                fieldname, op = key, self.EQ

            if len(args) == 0 and len(kwargs) == 1:
                # This is a single-kwarg Q object with a lookup that requires a single value. Make this a leaf
                self.fieldname = fieldname
                self.op = op
                self.value = value
                break

            self.children.append(self.__class__(**{key: value}))

        if len(self.children) == 1 and self.fieldname is None and self.conn_type != self.NOT:
            # We only have one child and no expression on ourselves, so we are a no-op. Flatten by taking over the child
            self._promote()

        self.clean()

    def _promote(self):
        # Flatten by taking over the only child
        assert len(self.children) == 1 and self.fieldname is None
        q = self.children[0]
        self.conn_type = q.conn_type
        self.fieldname = q.fieldname
        self.op = q.op
        self.value = q.value
        self.children = q.children

    def clean(self):
        if self.is_empty():
            return
        assert self.conn_type in self.CONN_TYPES
        if not self.is_leaf():
            return
        assert self.fieldname
        assert self.op in self.OP_TYPES
        if self.op == self.EXISTS:
            assert self.value is True
        if self.value is None:
            raise ValueError('Value for filter on field "%s" cannot be None' % self.fieldname)
        if isinstance(self.value, (tuple, list, set)):
            raise ValueError('Value for filter on field "%s" must be a single value' % self.fieldname)
        try:
            value_to_xml_text(self.value)
        except ValueError:
            raise ValueError('Value "%s" for filter in field "%s" is unsupported' % (self.value, self.fieldname))
        if isinstance(self.value, EWSDateTime):
            # We want to convert all values to UTC
            self.value = self.value.astimezone(UTC)

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
        assert op in (cls.EXACT, cls.IEXACT, cls.CONTAINS, cls.ICONTAINS, cls.STARTSWITH, cls.ISTARTSWITH)

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
            assert False
        if op in (cls.IEXACT, cls.ICONTAINS, cls.ISTARTSWITH):
            compare_mode = 'IgnoreCase'
        else:
            compare_mode = 'Exact'
        return create_element('t:Contains', ContainmentMode=match_mode, ContainmentComparison=compare_mode)

    def is_leaf(self):
        return not self.children

    def is_empty(self):
        return self.is_leaf() and self.fieldname is None

    def expr(self):
        if self.is_empty():
            return None
        if self.is_leaf():
            expr = '%s %s %s' % (self.fieldname, self.op, repr(self.value))
        else:
            # Sort children by field name so we get stable output (for easier testing). Children should never be empty.
            expr = (' %s ' % (self.AND if self.conn_type == self.NOT else self.conn_type)).join(
                (c.expr() if c.is_leaf() or c.conn_type == self.NOT else '(%s)' % c.expr())
                for c in sorted(self.children, key=lambda i: i.fieldname or '')
            )
        if self.conn_type == self.NOT:
            # Add the NOT operator. Put children in parens if there is more than one child.
            if self.is_leaf() or len(self.children) == 1:
                return self.conn_type + ' %s' % expr
            return self.conn_type + ' (%s)' % expr
        return expr

    @staticmethod
    def _validate_field(field, folder):
        if field not in folder.allowed_fields():
            raise ValueError("'%s' is not a valid field when filtering on %s" % (field.name, folder.__class__.__name__))
        if not field.is_searchable:
            raise ValueError("EWS does not support filtering on field '%s'" % field.name)

    def to_xml(self, folder, version):
        # Translate this Q object to a valid Restriction XML tree
        elem = self.xml_elem(folder=folder, version=version)
        if elem is None:
            return None
        from xml.etree.ElementTree import ElementTree
        restriction = create_element('m:Restriction')
        restriction.append(elem)
        return ElementTree(restriction).getroot()

    def xml_elem(self, folder, version):
        # Recursively build an XML tree structure of this Q object. If this is an empty leaf (the equivalent of Q()),
        # return None.
        from .fields import IndexedField
        if self.is_empty():
            return None
        if self.is_leaf():
            elem = self._op_to_xml(self.op)
            field = folder.get_item_field_by_fieldname(self.fieldname)
            self._validate_field(field=field, folder=folder)
            if isinstance(field, IndexedField):
                field_uri = field.field_uri_xml(version=version, label=self.value.label)
            else:
                field_uri = field.field_uri_xml(version=version)
            elem.append(field_uri)
            constant = create_element('t:Constant')
            if self.op != self.EXISTS:
                # Use .set() to not fill up the create_element() cache with unique values
                if field.is_list and not isinstance(self.value, (tuple, list, set)):
                    # With __contains, we allow filtering by only one value even though the field is a liste type
                    value = field.clean(value=[self.value], version=version)[0]
                else:
                    value = field.clean(value=self.value, version=version)
                constant.set('Value', value_to_xml_text(value))
                if self.op in self.CONTAINS_OPS:
                    elem.append(constant)
                else:
                    uriorconst = create_element('t:FieldURIOrConstant')
                    uriorconst.append(constant)
                    elem.append(uriorconst)
        else:
            # We have multiple children. If conn_type is NOT, then group children with AND. We'll add the NOT later
            elem = self._conn_to_xml(self.AND if self.conn_type == self.NOT else self.conn_type)
            # Sort children by field name so we get stable output (for easier testing). Children should never be empty
            for c in sorted(self.children, key=lambda i: i.fieldname or ''):
                elem.append(c.xml_elem(folder=folder, version=version))
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
            if len(self.children) == 1 and self.fieldname is None:
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

    def __str__(self):
        return self.expr()

    def __repr__(self):
        if self.is_leaf():
            return self.__class__.__name__ + '(%s %s %s)' % (self.fieldname, self.op, repr(self.value))
        sorted_children = tuple(sorted(self.children, key=lambda i: i.fieldname or ''))
        if self.conn_type == self.NOT or len(self.children) > 1:
            return self.__class__.__name__ + repr((self.conn_type,) + sorted_children)
        return self.__class__.__name__ + repr(sorted_children)


@python_2_unicode_compatible
class Restriction(object):
    """
    Implements an EWS Restriction type.

    """

    def __init__(self, q, folder):
        assert isinstance(q, Q)
        if q.is_empty():
            raise ValueError("Q object must not be empty")
        from .folders import Folder
        assert isinstance(folder, Folder)
        self.q = q
        self.folder = folder

    def to_xml(self, version):
        return self.q.to_xml(folder=self.folder, version=version)

    def __str__(self):
        """
        Prints the XML syntax tree
        """
        return xml_to_str(self.to_xml(version=None))
