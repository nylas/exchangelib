import logging
from threading import Lock

from .ewsdatetime import UTC
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
    CONTAINS = 'contains'
    RANGE = 'range'
    OP_TYPES = (EQ, NE, GT, GTE, LT, LTE, IN, CONTAINS, RANGE)

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
                    # Interpret 'foo_range=(1, 2)' as 'foo__gte=1 and foo__lte=2'
                    self.children.append(self.__class__(**{'%s__%s' % (field, self.GTE): value[0]}))
                    self.children.append(self.__class__(**{'%s__%s' % (field, self.LTE): value[1]}))
                    continue
                else:
                    op = self._lookup_to_op(lookup)
            else:
                field, op = key, self.EQ
            assert op in self.OP_TYPES
            if len(args) == 0 and len(kwargs) == 1:
                self.field = field
                self.op = op
                self.value = value
            else:
                self.children.append(Q(**{key: value}))

    @classmethod
    def _lookup_to_op(cls, lookup):
        try:
            return {
                'gt': cls.GT,
                'gte': cls.GTE,
                'lt': cls.LT,
                'lte': cls.LTE,
                'contains': cls.CONTAINS,
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
        if op == cls.CONTAINS:
            return create_element('t:Contains', ContainmentMode='Substring', ContainmentComparison='Exact')
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
        from xml.etree.ElementTree import ElementTree
        return ElementTree(self._to_xml_elem()).getroot()

    def _to_xml_elem(self):
        # there are exactly one child, ignore the AND/OR. If this is an empty leaf (equivalent of Q()), return None
        pass

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

    @classmethod
    def from_params(cls, start=None, end=None, categories=None, subject=None):
        # Builds a search expression string using the most common criteria and returns a Restriction object
        if not (start or end or categories):
            return None
        search_expr = []
        if start:
            search_expr.append('calendar:End > "%s"' % start.astimezone(UTC).ewsformat())
        if end:
            search_expr.append('calendar:Start < "%s"' % end.astimezone(UTC).ewsformat())
        if subject:
            search_expr.append('item:Subject = "%s"' % subject)
        if categories:
            if len(categories) == 1:
                search_expr.append('item:Categories in "%s"' % categories[0])
            else:
                expr2 = []
                for cat in categories:
                    # TODO Searching for items with multiple categories seems to be broken in EWS. 'And' operator
                    # returns no items, and searching for a list of categories doesn't work either.
                    expr2.append('item:Categories in "%s"' % cat)
                search_expr.append('( ' + ' or '.join(expr2) + ' )')
        expr_str = ' and '.join(search_expr)
        return cls.from_source(expr_str)

    def __str__(self):
        """
        Prints the XML syntax tree
        """
        return xml_to_str(self.xml)
