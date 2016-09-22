import logging

from .restriction import Q
from .services import IdOnly, AllProperties

log = logging.getLogger(__name__)


class MultipleObjectsReturned(Exception):
    pass


class DoesNotExist(Exception):
    pass


class QuerySet:
    """
    A Django QuerySet-like class for querying items. Defers queries until the QuerySet is consumed. Supports chaining to
    build up complex queries.

    Django QuerySet documentation: https://docs.djangoproject.com/en/dev/ref/models/querysets/
    """
    def __init__(self, folder):
        self.folder = folder
        self.q = Q()
        self.only_fields = None
        self.order_fields = None
        self.reversed = False

        self._cache = None

    def copy(self):
        new_qs = self.__class__(self.folder)
        new_qs.q = self.q
        new_qs.only_fields = self.only_fields
        new_qs.order_fields = self.order_fields
        new_qs.reversed = self.reversed
        new_qs._cache = self._cache
        return new_qs

    def _check_fields(self, field_names):
        allowed_field_names = self.folder.item_model.fieldnames()
        for f in field_names:
            if f not in allowed_field_names:
                raise ValueError("Unknown fieldname '%s'" % f)

    def _query(self):
        shape = IdOnly if self.only_fields else AllProperties
        items = self.folder.find_items(self.q, additional_fields=self.only_fields, shape=shape)
        if self.order_fields:
            items = sorted(items, key=lambda i: tuple(getattr(i, f) for f in self.order_fields))
        if self.reversed:
            items = reversed(items)
        return items

    def __iter__(self):
        # Fill cache if this is the first iteration. Return an iterator over the cached results.
        if self._cache is None:
            log.debug('Filling cache')
            if self.q is None:
                self._cache = []
            else:
                self._cache = list(self._query())
        return iter(self._cache)

    ###############################
    #
    # Methods that support chaining
    #
    ###############################
    # Return copies of self, so foo_qs.filter(...) doesn't surprise a following call to e.g. foo_qs.all()
    def all(self):
        # Invalidate cache and return all objects
        new_qs = self.copy()
        new_qs._cache = None
        return new_qs

    def none(self):
        new_qs = self.copy()
        new_qs._cache = None
        new_qs.q = None
        return new_qs

    def filter(self, *args, **kwargs):
        new_qs = self.copy()
        q = Q.from_filter_args(self.folder.item_model, *args, **kwargs)
        new_qs.q = q if new_qs.q is None else new_qs.q & q
        return new_qs

    def exclude(self, *args, **kwargs):
        new_qs = self.copy()
        q = ~Q.from_filter_args(self.folder.item_model, *args, **kwargs)
        new_qs.q = q if new_qs.q is None else new_qs.q & q
        return new_qs

    def only(self, *args):
        try:
            self._check_fields(args)
        except ValueError as e:
            raise ValueError("%s in only()" % e.args[0])
        new_qs = self.copy()
        new_qs.only_fields = args
        return new_qs

    def order_by(self, *args):
        # TODO: support '-my_field' for reverse sort
        try:
            self._check_fields(args)
        except ValueError as e:
            raise ValueError("%s in order_by()" % e.args[0])
        new_qs = self.copy()
        new_qs.order_fields = args
        return new_qs

    def reverse(self):
        new_qs = self.copy()
        new_qs.reversed = not self.reversed
        return new_qs

    ###########################
    #
    # Methods that end chaining
    #
    ###########################
    def values(self, *args):
        try:
            self._check_fields(args)
        except ValueError as e:
            raise ValueError("%s in values()" % e.args[0])
        new_qs = self.copy()
        new_qs.only_fields = args
        if len(args) == 0:
            raise ValueError('values_list() requires at least one field name')
        for i in new_qs:
            yield {k: getattr(i, k) for k in args}

    def values_list(self, *args, flat=False):
        try:
            self._check_fields(args)
        except ValueError as e:
            raise ValueError("%s in values_list()" % e.args[0])
        new_qs = self.copy()
        new_qs.only_fields = args
        if len(args) == 0:
            raise ValueError('values_list() requires at least one field name')
        if flat and len(args) != 1:
            raise ValueError('flat=True requires exactly one field name')
        if flat:
            for i in new_qs:
                yield getattr(i, args[0])
        else:
            for i in new_qs:
                yield tuple(getattr(i, f) for f in args)

    def iterator(self):
        # Return an iterator that doesn't bother with caching
        return self._query()

    def get(self, *args, **kwargs):
        new_qs = self.filter(*args, **kwargs)
        items = list(new_qs)
        if len(items) == 0:
            raise DoesNotExist()
        if len(items) != 1:
            raise MultipleObjectsReturned()
        return items[0]

    def count(self):
        # Get the item count with as little effort as possible
        new_qs = self.copy()
        new_qs.only_fields = []
        new_qs.order_fields = None
        new_qs.reverse = False
        return len(list(new_qs))

    def exists(self):
        return self.count() > 0

    def delete(self):
        from .folders import ALL_OCCURRENCIES
        return self.folder.bulk_delete(ids=self, affected_task_occurrences=ALL_OCCURRENCIES)
