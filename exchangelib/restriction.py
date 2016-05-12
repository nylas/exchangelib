import logging
from threading import Lock

from .ewsdatetime import EWSTimeZone
from .util import create_element, xml_to_str

log = logging.getLogger(__name__)

_source_cache = dict()
_source_cache_lock = Lock()


class Restriction:
    """
    Implements an EWS Restriction type.

    """
    def __init__(self, source):
        """
        source is a search expression in Python syntax. EWS Item fieldnames may be spelled with a colon (:). They will
        be escaped as underscores (_) since colons are not allowed in Python identifiers. Example:

              calendar:Start > '2009-01-15T13:45:56Z' and ( not calendar:Subject == 'EWS Test' )
        """
        with _source_cache_lock:
            # Something within the parser module seems to be deadlocking. Wrap in lock
            if source not in _source_cache:
                _source_cache[source] = self.parse_source(source)
        self.xml = _source_cache[source]

    @classmethod
    def parse_source(cls, source):
        """
        Takes a string and returns an XML tree.

        """
        from parser import expr
        from xml.etree.ElementTree import ElementTree
        log.debug('Parsing source: %s', source)

        # Make the syntax of the expression legal Python syntax by replacing ':' in identifiers with '_'. Play safe and
        # only do this for known field URI prefixes.
        for prefix in ('conversation:', 'postitem:', 'distributionlist:', 'contacts:', 'task:', 'calendar:',
                       'meetingRequest:', 'meeting:', 'message:', 'item:', 'folder:'):
            new = prefix[:-1] + '_'
            source = source.replace(prefix, new)

        st = expr(source).tolist()
        etree = ElementTree(cls._parse_syntaxtree(st))
        # etree.register_namespace('t', 'http://schemas.microsoft.com/exchange/services/2006/messages')
        # etree.register_namespace('m', 'http://schemas.microsoft.com/exchange/services/2006/types')
        log.debug('Source parsed')
        return etree.getroot()

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
                return create_element('t:FieldURI', FieldURI=val.replace('_', ':'))  # Switch back to correct spelling
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

    @classmethod
    def from_params(cls, folder_id, start=None, end=None, categories=None):
        # Builds a search expression string using the most common criteria and returns a Restriction object
        if not (start or end or categories):
            return None
        search_expr = []
        tz = EWSTimeZone.timezone('UTC')
        if start:
            search_expr.append('%s:End > "%s"' % (folder_id, start.astimezone(tz).ewsformat()))
        if end:
            search_expr.append('%s:Start < "%s"' % (folder_id, end.astimezone(tz).ewsformat()))
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
        return cls(expr_str)

    def __str__(self):
        """
        Prints the XML syntax tree
        """
        return xml_to_str(self.xml)
