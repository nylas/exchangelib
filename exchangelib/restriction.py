import logging
from threading import Lock

from .ewsdatetime import EWSDateTime, UTC
from .util import create_element, xml_to_str, value_to_xml_text

log = logging.getLogger(__name__)

_source_cache = dict()
_source_cache_lock = Lock()


class Q:
    # Connection types
    AND = 'AND'
    OR = 'OR'
    NOT = 'NOT'
    CONN_TYPES = (AND, OR, NOT)

    # Operators
    EQ = '=='
    NE = '!='
    GT = '>'
    GTE = '>='
    LT = '<'
    LTE = '<='
    IN = 'in'
    EXACT = 'exact'
    IEXACT = 'iexact'
    CONTAINS = 'contains'
    ICONTAINS = 'icontains'
    STARTSWITH = 'startswith'
    ISTARTSWITH = 'istartswith'
    RANGE = 'range'
    OP_TYPES = (EQ, NE, GT, GTE, LT, LTE, IN, EXACT, IEXACT, CONTAINS, ICONTAINS, STARTSWITH, ISTARTSWITH, RANGE)
    CONTAINS_OPS = (EXACT, IEXACT, CONTAINS, ICONTAINS, STARTSWITH, ISTARTSWITH)

    def __init__(self, *args, **kwargs):
        if 'conn_type' in kwargs:
            self.conn_type = kwargs.pop('conn_type')
        else:
            self.conn_type = self.AND
        assert self.conn_type in self.CONN_TYPES

        self.field = None
        self.op = None
        self.value = None

        # Build children of Q objects from *args and **kwargs
        self.children = []
        for q in args:
            if not isinstance(q, self.__class__):
                raise AttributeError("'%s' must be a Q instance")
            if not q.is_empty():
                self.children.append(q)

        for key, value in kwargs.items():
            if '__' in key:
                field, lookup = key.rsplit('__')
                if lookup == self.RANGE:
                    # EWS doesn't have a 'range' operator. Emulate 'foo__range=(1, 2)' as 'foo__gte=1 and foo__lte=2'
                    # (both values inclusive).
                    self.children.append(self.__class__(**{'%s__gte' % field: value[0]}))
                    self.children.append(self.__class__(**{'%s__lte' % field: value[1]}))
                    continue
                if lookup == self.IN:
                    # EWS doesn't have an 'in' operator. Emulate 'foo in (1, 2, ...)' as 'foo==1 or foo==2 or ...'
                    or_args = []
                    for val in value:
                        or_args.append(self.__class__(**{field: val}))
                    self.children.append(Q(*or_args, conn_type=self.OR))
                    continue
                else:
                    op = self._lookup_to_op(lookup)
            else:
                field, op = key, self.EQ
            assert op in self.OP_TYPES
            if len(args) == 0 and len(kwargs) == 1:
                self.field = field
                self.op = op
                if isinstance(value, EWSDateTime):
                    # We want to convert all values to UTC
                    self.value = value.astimezone(UTC)
                else:
                    self.value = value
            else:
                self.children.append(Q(**{key: value}))

    @classmethod
    def _lookup_to_op(cls, lookup):
        try:
            return {
                'not': cls.NE,
                'gt': cls.GT,
                'gte': cls.GTE,
                'lt': cls.LT,
                'lte': cls.LTE,
                'exact': cls.EXACT,
                'iexact': cls.IEXACT,
                'contains': cls.CONTAINS,
                'icontains': cls.ICONTAINS,
                'startswith': cls.STARTSWITH,
                'istartswith': cls.ISTARTSWITH,
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
        if op in (cls.EXACT, cls.IEXACT, cls.CONTAINS, cls.ICONTAINS, cls.STARTSWITH, cls.ISTARTSWITH):
            # For description of Contains attribute values, see
            #     https://msdn.microsoft.com/en-us/library/office/aa580702(v=exchg.150).aspx
            #
            # Possible ContainmentMode values:
            #     FullString, Prefixed, Substring, PrefixOnWords, ExactPhrase
            # Django lookups have no equivalent of PrefixOnWords and ExactPhrase (and I'm unsure how they actually
            # work).
            # TODO EWS has no equivalent of '__endswith' or '__iendswith' so that would need to be emulated using
            # Substring and post-processing in Python.
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
        return self.is_leaf() and self.field is None

    def expr(self):
        if self.is_empty():
            return None
        if self.is_leaf():
            assert self.field and self.op and self.value is not None
            expr = '%s %s %s' % (self.field, self.op, repr(self.value))
        elif len(self.children) == 1:
            # Flatten the tree a bit
            expr = self.children[0].expr()
        else:
            # Sort children by field name so we get stable output (for easier testing). Children should never be empty.
            expr = (' %s ' % (self.AND if self.conn_type == self.NOT else self.conn_type)).join(
                (c.expr() if c.is_leaf() or c.conn_type == self.NOT else '(%s)' % c.expr())
                for c in sorted(self.children, key=lambda i: i.field or '')
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

    def to_xml(self):
        elem = self._to_xml_elem()
        if elem is None:
            return None
        from xml.etree.ElementTree import ElementTree
        restriction = create_element('m:Restriction')
        restriction.append(self._to_xml_elem())
        return ElementTree(restriction).getroot()

    def _to_xml_elem(self):
        # Return an XML tree structure of this Q object. First, remove any empty children. If conn_type is AND or OR and
        # there is exactly one child, ignore the AND/OR and treat this node as a leaf. If this is an empty leaf
        # (equivalent of Q()), return None.
        if self.is_empty():
            return None
        if self.is_leaf():
            assert self.field and self.op and self.value is not None
            elem = self._op_to_xml(self.op)
            field = create_element('t:FieldURI', FieldURI=self.field)
            elem.append(field)
            constant = create_element('t:Constant', Value=value_to_xml_text(self.value))
            if self.op in self.CONTAINS_OPS:
                elem.append(constant)
            else:
                uriorconst = create_element('t:FieldURIOrConstant')
                uriorconst.append(constant)
                elem.append(uriorconst)
        elif len(self.children) == 1:
            # Flatten the tree a bit
            elem = self.children[0]._to_xml_elem()
        else:
            # We have multiple children. If conn_type is NOT, then group children with AND. We'll add the NOT later
            elem = self._conn_to_xml(self.AND if self.conn_type == self.NOT else self.conn_type)
            # Sort children by field name so we get stable output (for easier testing). Children should never be empty
            for c in sorted(self.children, key=lambda i: i.field or ''):
                elem.append(c._to_xml_elem())
        if elem is None:
            return None  # Should not be necessary, but play safe
        if self.conn_type == self.NOT:
            # Encapsulate everything in the NOT element
            not_elem = self._conn_to_xml(self.conn_type)
            not_elem.append(elem)
            return not_elem
        return elem

    def __and__(self, other):
        # Return a new Q with two children and conn_type AND
        return self.__class__(self, other, conn_type=self.AND)

    def __or__(self, other):
        # Return a new Q with two children and conn_type OR
        return self.__class__(self, other, conn_type=self.OR)

    def __invert__(self):
        # If this is a leaf and op has an inverse, change op. Else return a new Q with conn_type NOT
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

    def __repr__(self):
        if self.is_leaf():
            return self.__class__.__name__ + '(%s %s %s)' % (self.field, self.op, repr(self.value))
        if self.conn_type == self.NOT or len(self.children) > 1:
            return self.__class__.__name__ + repr((self.conn_type,) + tuple(self.children))
        return self.__class__.__name__ + repr(tuple(self.children))


class Restriction:
    """
    Implements an EWS Restriction type.

    """
    def __init__(self, xml):
        from xml.etree.ElementTree import Element
        if not isinstance(xml, Element):
            raise ValueError("'xml' must be an ElementTree (%s)", type(xml))
        self.xml = xml

    @classmethod
    def from_source(cls, source):
        """
        source is a search expression in Python syntax. EWS Item fieldnames may be spelled with a colon (:). They will
        be escaped as underscores (_) since colons are not allowed in Python identifiers. Example:

            calendar:Start > '2009-01-15T13:45:56Z' and not (item:Subject == 'EWS Test' or item:Subject == 'Foo')
        """
        with _source_cache_lock:
            # Something within the parser module seems to be deadlocking. Wrap in lock
            if source not in _source_cache:
                from parser import expr
                log.debug('Parsing source: %s', source)
                st = expr(cls._escape(source)).tolist()
                from xml.etree.ElementTree import ElementTree
                etree = ElementTree(cls._parse_syntaxtree(st))
                # etree.register_namespace('t', 'http://schemas.microsoft.com/exchange/services/2006/messages')
                # etree.register_namespace('m', 'http://schemas.microsoft.com/exchange/services/2006/types')
                log.debug('Source parsed')
                _source_cache[source] = etree.getroot()
        return cls(_source_cache[source])

    @classmethod
    def _parse_syntaxtree(cls, slist):
        """
        Takes a Python syntax tree containing a search restriction expression and returns the tree as EWS-formatted XML
        """
        from token import NAME, EQEQUAL, NOTEQUAL, GREATEREQUAL, LESSEQUAL, LESS, GREATER, STRING, LPAR, RPAR, \
            NEWLINE, ENDMARKER
        from symbol import and_test, or_test, not_test, comparison, eval_input, sym_name, atom
        key = slist[0]
        if isinstance(slist[1], list):
            if len(slist) == 2:
                # Let nested 2-element lists pass transparently
                return cls._parse_syntaxtree(slist[1])
            if key == atom:
                # This is a parens with contents. Continue without the parens since they are unnecessary when building
                # a tree - a node *is* the parens.
                assert slist[1][0] == LPAR
                assert slist[3][0] == RPAR
                return cls._parse_syntaxtree(slist[2])
            if key == or_test:
                e = create_element('t:Or')
                for item in [cls._parse_syntaxtree(l) for l in slist[1:]]:
                    if item is not None:
                        e.append(item)
                return e
            if key == and_test:
                e = create_element('t:And')
                for item in [cls._parse_syntaxtree(l) for l in slist[1:]]:
                    if item is not None:
                        e.append(item)
                return e
            if key == not_test:
                e = create_element('t:Not')
                for item in [cls._parse_syntaxtree(l) for l in slist[1:]]:
                    if item is not None:
                        e.append(item)
                return e
            if key == comparison:
                op = cls._parse_syntaxtree(slist[2])
                field = cls._parse_syntaxtree(slist[1])
                constant = cls._parse_syntaxtree(slist[3])
                op.append(field)
                if op.tag == 't:Contains':
                    op.append(constant)
                else:
                    uriorconst = create_element('t:FieldURIOrConstant')
                    uriorconst.append(constant)
                    op.append(uriorconst)
                return op
            if key == eval_input:
                e = create_element('m:Restriction')
                for item in [cls._parse_syntaxtree(l) for l in slist[1:]]:
                    if item is not None:
                        e.append(item)
                return e
            raise ValueError('Unknown element type: %s %s (slist %s len %s)' % (key, sym_name[key], slist, len(slist)))
        else:
            val = slist[1]
            if key == NAME:
                if val in ('and', 'or', 'not'):
                    return None
                if val == 'in':
                    return create_element('t:Contains', ContainmentMode='Substring', ContainmentComparison='Exact')
                return create_element('t:FieldURI', FieldURI=cls._unescape(val))
            if key == EQEQUAL:
                return create_element('t:IsEqualTo')
            if key == NOTEQUAL:
                return create_element('t:IsNotEqualTo')
            if key == GREATEREQUAL:
                return create_element('t:IsGreaterThanOrEqualTo')
            if key == LESSEQUAL:
                return create_element('t:IsLessThanOrEqualTo')
            if key == LESS:
                return create_element('t:IsLessThan')
            if key == GREATER:
                return create_element('t:IsGreaterThan')
            if key == STRING:
                # This is a string, so strip single/double quotes
                return create_element('t:Constant', Value=val.strip('"\''))
            if key in (LPAR, RPAR, NEWLINE, ENDMARKER):
                return None
            raise ValueError('Unknown token type: %s %s' % (key, val))

    @staticmethod
    def _escape(source):
        # Make the syntax of the expression legal Python syntax by replacing ':' in identifiers with '_'. Play safe and
        # only do this for known property prefixes. See Table 1 and 5 in
        # https://msdn.microsoft.com/en-us/library/office/dn467898(v=exchg.150).aspx
        for prefix in ('message:', 'calendar:', 'contacts:', 'conversation:', 'distributionlist:', 'folder:', 'item:',
                       'meeting:', 'meetingRequest:', 'postitem:', 'task:'):
            new = prefix[:-1] + '_'
            source = source.replace(prefix, new)
        return source

    @staticmethod
    def _unescape(fieldname):
        # Switch back to correct fieldname spelling. Inverse of _unescape()
        if '_' not in fieldname:
            return fieldname
        prefix, field = fieldname.split('_', maxsplit=1)
        # There aren't any valid FieldURI values with an underscore
        assert prefix in ('message', 'calendar', 'contacts', 'conversation', 'distributionlist', 'folder', 'item',
                          'meeting', 'meetingRequest', 'postitem', 'task')
        return '%s:%s' % (prefix, field)

    def __str__(self):
        """
        Prints the XML syntax tree
        """
        return xml_to_str(self.xml)
