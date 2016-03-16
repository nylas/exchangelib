from parser import expr
from token import NAME, EQEQUAL, NOTEQUAL, GREATEREQUAL, LESSEQUAL, LESS, GREATER, STRING, LPAR, RPAR, NEWLINE, \
    ENDMARKER
from symbol import and_test, or_test, not_test, comparison, eval_input, sym_name
from xml.etree.cElementTree import Element, ElementTree, tostring
import logging
from threading import Lock

from .ewsdatetime import EWSTimeZone

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
                _source_cache[source] = self.parse_source(source).getroot()
        self.xml = _source_cache[source]

    def parse_source(self, source):
        """
        Takes a string and returns an XML tree.

        """
        log.debug('Parsing source: %s', source)

        # Make the syntax of the expression legal Python syntax. Play safe and only do this for known field URI prefixes
        for prefix in ('conversation:', 'postitem:', 'distributionlist:', 'contacts:', 'task:', 'calendar:',
                       'meetingRequest:', 'meeting:', 'message:', 'item:', 'folder:'):
            new = prefix.replace(':', '_')
            source = source.replace(prefix, new)

        st = expr(source).tolist()
        etree = ElementTree(self.parse_syntaxtree(st))
        # etree.register_namespace('t', 'http://schemas.microsoft.com/exchange/services/2006/messages')
        # etree.register_namespace('m', 'http://schemas.microsoft.com/exchange/services/2006/types')
        log.debug('Source parsed')
        return etree

    def parse_syntaxtree(self, slist):
        """
        Takes a Python syntax tree containing a search restriction expression and returns the tree as EWS-formatted XML
        """
        key = slist[0]
        if isinstance(slist[1], list):
            if len(slist) == 2:
                # Let nested 2-element lists pass transparently
                return self.parse_syntaxtree(slist[1])
            else:
                if key == or_test:
                    e = Element('t:Or')
                    for item in [self.parse_syntaxtree(l) for l in slist[1:]]:
                        if item is not None:
                            e.append(item)
                    return e
                if key == and_test:
                    e = Element('t:And')
                    for item in [self.parse_syntaxtree(l) for l in slist[1:]]:
                        if item is not None:
                            e.append(item)
                    return e
                if key == not_test:
                    e = Element('t:Not')
                    for item in [self.parse_syntaxtree(l) for l in slist[1:]]:
                        if item is not None:
                            e.append(item)
                    return e
                if key == comparison:
                    op = self.parse_syntaxtree(slist[2])
                    field = self.parse_syntaxtree(slist[1])
                    constant = self.parse_syntaxtree(slist[3])
                    op.append(field)
                    if op.tag == 't:Contains':
                        op.append(constant)
                    else:
                        uriorconst = Element('t:FieldURIOrConstant')
                        uriorconst.append(constant)
                        op.append(uriorconst)
                    return op
                if key == eval_input:
                    e = Element('m:Restriction')
                    for item in [self.parse_syntaxtree(l) for l in slist[1:]]:
                        if item is not None:
                            e.append(item)
                    return e
                raise ValueError('Unknown element type: %s %s' % (key, sym_name[key]))
        else:
            val = slist[1]
            if key == NAME:
                if val in ('and', 'or', 'not'):
                    return None
                if val == 'in':
                    return Element('t:Contains', ContainmentMode='Substring', ContainmentComparison='Exact')
                return Element('t:FieldURI', FieldURI=val.replace('_', ':'))  # Switch back to correct spelling
            if key == EQEQUAL:
                return Element('t:IsEqualTo')
            if key == NOTEQUAL:
                return Element('t:IsNotEqualTo')
            if key == GREATEREQUAL:
                return Element('t:IsGreaterThanOrEqualTo')
            if key == LESSEQUAL:
                return Element('t:IsLessThanOrEqualTo')
            if key == LESS:
                return Element('t:IsLessThan')
            if key == GREATER:
                return Element('t:IsGreaterThan')
            if key == STRING:
                return Element('t:Constant', Value=val.strip('"\''))  # This is a string, so strip single/double quotes
            if key in (LPAR, RPAR, NEWLINE, ENDMARKER):
                return None
            raise ValueError('Unknown token type: %s %s' % (key, val))

    @classmethod
    def from_params(cls, folder_id, start=None, end=None, categories=None):
        # Builds a search expression string and returns the equivalent XML tree ready for inclusion in an EWS request
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
        return tostring(self.xml).decode('utf-8')
