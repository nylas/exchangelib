# coding=utf-8
from __future__ import unicode_literals

import functools
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

    def __init__(self, *args, **kwargs):
        if 'conn_type' in kwargs:
            self.conn_type = kwargs.pop('conn_type')
        else:
            self.conn_type = self.AND
        assert self.conn_type in self.CONN_TYPES

        self.fieldname = None  # Name of the field we want to filter on
        self.op = None
        self.value = None

        # Build children of Q objects from *args and **kwargs
        self.children = []
        for q in args:
            if not isinstance(q, self.__class__):
                if isinstance(q, Restriction):
                    q = q.q
                else:
                    raise AttributeError("'%s' must be a Q or Restriction instance" % q)
            if not q.is_empty():
                self.children.append(q)

        for key, value in kwargs.items():
            if value is None:
                raise ValueError('Value for Q kwarg "%s" cannot be None' % key)
            if '__' in key:
                fieldname, lookup = key.rsplit('__')
                if lookup == self.LOOKUP_RANGE:
                    # EWS doesn't have a 'range' operator. Emulate 'foo__range=(1, 2)' as 'foo__gte=1 and foo__lte=2'
                    # (both values inclusive).
                    if len(value) != 2:
                        raise ValueError("Value of kwarg '%s' must have exactly 2 elements" % key)
                    self.children.append(self.__class__(**{'%s__gte' % fieldname: value[0]}))
                    self.children.append(self.__class__(**{'%s__lte' % fieldname: value[1]}))
                    continue
                if lookup == self.LOOKUP_IN:
                    # EWS doesn't have an 'in' operator. Emulate 'foo in (1, 2, ...)' as 'foo==1 or foo==2 or ...'
                    or_args = []
                    for val in value:
                        or_args.append(self.__class__(**{fieldname: val}))
                    self.children.append(Q(*or_args, conn_type=self.OR))
                    continue
                else:
                    op = self._lookup_to_op(lookup)
            else:
                fieldname, op = key, self.EQ

            assert op in self.OP_TYPES
            try:
                value_to_xml_text(value)
            except ValueError:
                raise ValueError('Value "%s" for filter kwarg "%s" is unsupported' % (value, key))
            if len(args) == 0 and len(kwargs) == 1:
                self.fieldname = fieldname
                self.op = op
                if isinstance(value, EWSDateTime):
                    # We want to convert all values to UTC
                    if not getattr(value, 'tzinfo'):
                        raise ValueError("'%s' must be timezone aware" % fieldname)
                    self.value = value.astimezone(UTC)
                else:
                    self.value = value
            else:
                self.children.append(Q(**{key: value}))

    @classmethod
    def from_filter_args(cls, folder_class, *args, **kwargs):
        # args and kwargs are Django-style q args and field lookups
        q = None
        if args:
            q_args = []
            for arg in args:
                if not isinstance(arg, Q):
                    raise ValueError("Non-keyword arg '%s' must be a Q object" % arg)
                q_args.append(arg)
            # AND all the given Q objects together
            q = functools.reduce(lambda a, b: a & b, q_args)
        if kwargs:
            kwargs_q = q or Q()
            for key, value in kwargs.items():
                if value is None:
                    raise ValueError('Value for filter kwarg "%s" cannot be None' % key)
                if '__' in key:
                    fieldname, lookup = key.rsplit('__')
                else:
                    fieldname, lookup = key, None
                field = folder_class.get_item_field_by_fieldname(fieldname)
                cls._validate_field(field, folder_class)
                if lookup == Q.LOOKUP_EXISTS:
                    if value:
                        kwargs_q &= Q(**{key: True})
                    else:
                        kwargs_q &= ~Q(**{key: True})
                    continue
                # Filtering by category is a bit quirky. The only lookup type I have found to work is:
                #
                #     item:Categories == 'foo' AND item:Categories == 'bar' AND ...
                #
                #     item:Categories == 'foo' OR item:Categories == 'bar' OR ...
                #
                # The former returns items that have these categories, but maybe also others. The latter returns
                # items that have at least one of these categories. This translates to the 'contains' and 'in' lookups.
                # Both versions are case-insensitive.
                #
                # Exact matching and case-sensitive or partial-string matching is not possible since that requires the
                # 'Contains' element which only supports matching on string elements, not arrays.
                #
                # Exact matching of categories (i.e. match ['a', 'b'] but not ['a', 'b', 'c']) could be implemented by
                # post-processing items. Fetch 'item:Categories' with additional_fields and remove the items that don't
                # have an exact match, after the call to FindItems.
                if field.is_list:
                    if lookup not in (Q.LOOKUP_CONTAINS, Q.LOOKUP_IN):
                        raise ValueError(
                            "Field '%(field)s' can only be filtered using '%(field)s__contains=['a', 'b', ...]' and "
                            "'%(field)s__in=['a', 'b', ...]' and '%(field)s__exists=True|False'" % dict(field=fieldname))
                    if isinstance(value, (list, tuple, set)):
                        children = [Q(**{fieldname: v}) for v in value]
                        if lookup == Q.LOOKUP_CONTAINS:
                            kwargs_q &= Q(*children, conn_type=Q.AND)
                        elif lookup == Q.LOOKUP_IN:
                            kwargs_q &= Q(*children, conn_type=Q.OR)
                        else:
                            assert False
                    else:
                        kwargs_q &= Q(**{fieldname: value})
                    continue
                if lookup == Q.LOOKUP_IN and isinstance(value, (list, tuple, set)):
                    # Allow '__in' lookup on non-list field types, specifying a list or a simple value
                    children = [Q(**{fieldname: v}) for v in value]
                    kwargs_q &= Q(*children, conn_type=Q.OR)
                    continue
                if isinstance(value, (list, tuple, set)) and lookup != Q.LOOKUP_RANGE:
                    raise ValueError('Value for filter kwarg "%s" must be a single value' % key)
                kwargs_q &= Q(**{key: value})
            q = kwargs_q
        return q

    @classmethod
    def _lookup_to_op(cls, lookup):
        try:
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
        except KeyError:
            raise ValueError("Lookup '%s' is not supported" % lookup)

    @classmethod
    def _conn_to_xml(cls, conn_type):
        if conn_type == cls.AND:
            return create_element('t:And')
        if conn_type == cls.OR:
            return create_element('t:Or')
        if conn_type == cls.NOT:
            return create_element('t:Not')
        raise ValueError("Unknown conn_type: '%s'" % conn_type)

    @classmethod
    def _op_to_xml(cls, op):
        if op == cls.EQ:
            return create_element('t:IsEqualTo')
        if op == cls.NE:
            return create_element('t:IsNotEqualTo')
        if op == cls.GTE:
            return create_element('t:IsGreaterThanOrEqualTo')
        if op == cls.LTE:
            return create_element('t:IsLessThanOrEqualTo')
        if op == cls.LT:
            return create_element('t:IsLessThan')
        if op == cls.GT:
            return create_element('t:IsGreaterThan')
        if op == cls.EXISTS:
            return create_element('t:Exists')
        if op in (cls.EXACT, cls.IEXACT, cls.CONTAINS, cls.ICONTAINS, cls.STARTSWITH, cls.ISTARTSWITH):
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
        raise ValueError("Unknown op: '%s'" % op)

    def is_leaf(self):
        return not self.children

    def is_empty(self):
        return self.is_leaf() and self.fieldname is None

    def expr(self):
        if self.is_empty():
            return None
        if self.is_leaf():
            assert self.fieldname and self.op and self.value is not None
            expr = '%s %s %s' % (self.fieldname, self.op, repr(self.value))
        elif len(self.children) == 1:
            # Flatten the tree a bit
            expr = self.children[0].expr()
        else:
            # Sort children by field name so we get stable output (for easier testing). Children should never be empty.
            expr = (' %s ' % (self.AND if self.conn_type == self.NOT else self.conn_type)).join(
                (c.expr() if c.is_leaf() or c.conn_type == self.NOT else '(%s)' % c.expr())
                for c in sorted(self.children, key=lambda i: i.fieldname or '')
            )
        if not expr:
            return None  # Should not be necessary, but play safe
        if self.conn_type == self.NOT:
            # Add the NOT operator. Put children in parens if there is more than one child.
            if self.is_leaf() or (len(self.children) == 1 and self.children[0].is_leaf()):
                expr = self.conn_type + ' %s' % expr
            else:
                expr = self.conn_type + ' (%s)' % expr
        return expr

    @staticmethod
    def _validate_field(field, folder_class):
        from .attachments import Attachment
        from .properties import Mailbox, Attendee
        from .indexed_properties import PhysicalAddress
        if field not in folder_class.allowed_fields():
            raise ValueError("'%s' is not a valid field when filtering on %s" % (field.name, folder_class.__name__))
        if field.name in ('status', 'companies', 'reminder_due_by'):
            raise ValueError("EWS does not support filtering on field '%s'" % field.name)
        if field.is_list and field.value_cls in (Attachment, Mailbox, Attendee, PhysicalAddress):
            raise ValueError("EWS does not support filtering on %s (non-searchable field type [%s])" % (
                field.name, field.value_cls.__name__))

    def to_xml(self, folder_class):
        # Translate this Q object to a valid Restriction XML tree
        elem = self.xml_elem(folder_class=folder_class)
        if elem is None:
            return None
        from xml.etree.ElementTree import ElementTree
        restriction = create_element('m:Restriction')
        restriction.append(elem)
        return ElementTree(restriction).getroot()

    def xml_elem(self, folder_class):
        # Return an XML tree structure of this Q object. First, remove any empty children. If conn_type is AND or OR and
        # there is exactly one child, ignore the AND/OR and treat this node as a leaf. If this is an empty leaf
        # (equivalent of Q()), return None.
        from .fields import IndexedField
        if self.is_empty():
            return None
        if self.is_leaf():
            assert self.fieldname and self.op and self.value is not None
            elem = self._op_to_xml(self.op)
            field = folder_class.get_item_field_by_fieldname(self.fieldname)
            self._validate_field(field=field, folder_class=folder_class)
            if isinstance(field, IndexedField):
                field_uri = field.field_uri_xml(label=self.value.label)
            else:
                field_uri = field.field_uri_xml()
            elem.append(field_uri)
            constant = create_element('t:Constant')
            if self.op != self.EXISTS:
                # Use .set() to not fill up the create_element() cache with unique values
                constant.set('Value', value_to_xml_text(self.value))
                if self.op in self.CONTAINS_OPS:
                    elem.append(constant)
                else:
                    uriorconst = create_element('t:FieldURIOrConstant')
                    uriorconst.append(constant)
                    elem.append(uriorconst)
        elif len(self.children) == 1:
            # Flatten the tree a bit
            elem = self.children[0].xml_elem(folder_class=folder_class)
        else:
            # We have multiple children. If conn_type is NOT, then group children with AND. We'll add the NOT later
            elem = self._conn_to_xml(self.AND if self.conn_type == self.NOT else self.conn_type)
            # Sort children by field name so we get stable output (for easier testing). Children should never be empty
            for c in sorted(self.children, key=lambda i: i.fieldname or ''):
                elem.append(c.xml_elem(folder_class=folder_class))
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
        # ~ operator. If this is a leaf and op has an inverse, change op. Else return a new Q with conn_type NOT
        if self.conn_type == self.NOT:
            # This is NOT NOT. Change to AND
            self.conn_type = self.AND
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
        if self.conn_type == self.NOT or len(self.children) > 1:
            return self.__class__.__name__ + repr((self.conn_type,) + tuple(self.children))
        return self.__class__.__name__ + repr(tuple(self.children))


@python_2_unicode_compatible
class Restriction(object):
    """
    Implements an EWS Restriction type.

    """

    def __init__(self, q, folder_class):
        if not isinstance(q, Q):
            raise ValueError("'q' must be a Q object (%s)", type(q))
        if q.is_empty():
            raise ValueError("Q object must not be empty")
        from .folders import Folder
        assert issubclass(folder_class, Folder)
        self.q = q
        self.folder_class = folder_class

    def to_xml(self, version):
        return self.q.to_xml(folder_class=self.folder_class)

    def __str__(self):
        """
        Prints the XML syntax tree
        """
        return xml_to_str(self.to_xml(version=None))
