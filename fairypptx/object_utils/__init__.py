"""Utility functions of Objects.
 
Utility functions for handling Objects.

"""
import builtins
from _ctypes import COMError
from contextlib import contextmanager
from collections.abc import Sequence
from collections import UserDict, UserString
from collections.abc import Mapping

import numpy as np

from fairypptx import registory_utils


def get_type(instance):
    """Return the Capitalized Object Type Name."""
    return getattr(type(instance), "__com_interface__").__name__.strip("_").capitalize()


def is_object(instance, name=None):
    """Return whether ``instance`` is regarded as Powerpoint Object.

    Args:
        instance: the checked instance.
        name(str): the Object Type name.

    Note
    --------------
    Return ``True``, if ``instance`` has ``__com_interface__``.
    Hence, this check is very weak so ``instance`` may be not related to
    PowerPoint, even if ``True`` is returned.

    """
    try:
        c_name = get_type(instance)
    except AttributeError:
        return False
    if name:
        return name.capitalize() == c_name
    else:
        return True


def upstream(instance, name):
    """Return `name` Object as long as `instance.parent..parent` is achievable.

    Args:
        instance: instance.
        name(str): the Target ObjectType Name.

    Raises: ValueError:
    """
    target = instance
    while True:
        try:
            type_name = get_type(target)
        except AttributeError:
            raise ValueError(f"`{instance}` is not regarded as Object.")
        else:
            if type_name == name.capitalize():
                return target
        try:
            target = target.parent
        except AttributeError as e:
            raise ValueError(f"`{name}` is not an ancestor of `{instance}`.")


def setattr(instance, attr, value):
    """Extension of `setattr` so that `.` specifier is valid.
    Args:
        instance: Object.
        attr (str or Sequence of str): specifier.
        value: value

    Raises:
        ValueError: When `value` is not invalid.
        AttributeError: When attribute is not existent.

    """
    elems = _listify(attr)
    target = instance
    for elem in elems[:-1]:
        target = builtins.getattr(target, elem)
    try:
        builtins.setattr(target, elems[-1], value)
    except AttributeError as e:
        raise AttributeError(f"`{value}` cannot be set to `{attr}`") from e
    except COMError as e:
        raise ValueError(f"`{value}` cannot be set to `{attr}`") from e


_NOT_SPECIFIED = object()


def getattr(instance, attr, default=_NOT_SPECIFIED):
    """Extension of `setattr` so that `.` specifier is valid.
    Args:
        instance: Object.
        attr (str or Sequence of str): specifier.
        default(Any): the default value
                      when attribute is not exist.
    Raises:
        AttributeError: Attribute is not existent and default is not set.
    """
    if default is not _NOT_SPECIFIED:
        try:
            return getattr(instance, attr)
        except AttributeError:
            return default
    else:
        elems = _listify(attr)
        target = instance
        for elem in elems:
            target = builtins.getattr(target, elem)
        return target


def hasattr(instance, attr):
    """Extension of `builtins.hasattr` so that `.` specifier is valid."""
    try:
        getattr(instance, attr)
    except AttributeError:
        return False
    return True


@contextmanager
def stored(instance, attrs):
    """Store the values of `instance`'s `attrs`.
    Args:
        instance: Object
        attrs (str, Sequence of str or Sequence of Sequence):
    """

    stock = dict()
    if isinstance(attrs, str):
        attrs = [attrs]
    # Stocked Phase.
    for attr in attrs:
        elems = _listify(attr)
        stock[tuple(elems)] = getattr(instance, elems)

    # Rollback Phase
    def _rollback():
        for key, value in stock.items():
            setattr(instance, key, value)

    try:
        yield
    except Exception as e:
        _rollback()
        raise e
    else:
        _rollback()


def _listify(attr):
    if isinstance(attr, str):
        attr = attr.split(".")
    if not isinstance(attr, Sequence):
        raise ValueError(f"`{attr}` is not valid specifier.")
    return attr


class ObjectDictMixin(UserDict):
    """Mixin for Object Dict.

    Args:
        arg: the target of Object or Dict.
        name: the Name of Object.

    Protocol of Derivative classes.
    --------------------------------

    * `data`: class attribute is required as a template.
    * `name`: Object Name.
              If `None`, then the classname is the same as the Object name.
    * `readonly`: `data` is stored,
                   however, when `setattr` is called at `apply`,
                   the keys of `readonly` is skipped.
                   Typically, read-only properties of Object
                   is set here.

    Provided Service
    ------------------------------
    * conversion between `Object` and UserDict.
    * fetch / register of Object / UserDict.
    * key access to Object API.
        - [TODO] `key` of `dict` must be case incensitive.
            - It leads to th duplicated (key, value).

    Note
    ----------------------------------
    `fetch` <-> `register`.

    When you implement `to_dict`,
    typically you also implement`apply`.
    """

    data = dict()
    name = None
    readonly = []

    def __init__(self, arg=None, **kwargs):
        self.cls = type(self)
        if self.cls.name:
            self.name = self.cls.name
        else:
            self.name = self.cls.__name__
        self.data, self._api = self._construct(arg, **kwargs)

    @property
    def api(self):
        return self._api

    def detached(self):
        """Return this class without `_api`."""
        return self.cls(self)

    def __getstate__(self):
        # This is for serialization of `pickle`.
        return {"data": self.data, "name": self.name, "readonly": self.readonly}

    def __setitem__(self, key, item):
        super().__setitem__(key, item)
        if self._api:
            if not hasattr(self._api, key):
                api_type = get_type(self._api)
                raise KeyError(f"`{key}` does not exist in `{api_type}`.")
            setattr(self._api, key, item)
        else:
            # When `api` is not given, this behaves as normal `UserDict`.
            pass

    def __setattr__(self, name, value):

        # Without `_api`, the behavior is the same as `UserDict`.
        if "_api" not in self.__dict__:
            object.__setattr__(self, name, value)
            return
        if self._api is None:
            object.__setattr__(self, name, value)
            return

        if name in self.__dict__ or name in type(self).__dict__:
            object.__setattr__(self, name, value)
        elif hasattr(self.api, name):
            setattr(self.api, name, value)
            self[name] = value
        else:
            # TODO: Maybe require modification.
            object.__setattr__(self, name, value)

    def to_dict(self, api_object):
        """Convert `Object` to `dict`.

        As a default,  values of `keys` of `class.data`
        is copied.
        """
        keys = self.cls.data.keys()
        data = {key: getattr(api_object, key) for key in keys}
        return data

    def apply(self, api_object):
        """Apply `self` to `api_object`."""
        readonly_props = set(self.readonly)
        for key, value in self.data.items():
            if key not in readonly_props:
                setattr(api_object, key, value)
        return api_object

    def register(self, key, disk=False):
        """Register to the storage."""
        name = self._get_name()
        registory_utils.register(name, key, self.data, extension=".json", disk=disk)

    @classmethod
    def fetch(cls, key, disk=True):
        """Construct the instance with `key` object."""
        name = cls._get_name()
        data = registory_utils.fetch(name, key, disk=True)
        return cls(data)

    def _construct(self, arg, **kwargs):
        name = self._get_name()
        api = None
        if arg is None:
            data = dict(self.cls.data)
        elif is_object(arg, name):
            data = self.to_dict(arg)
            api = arg
        elif isinstance(arg, Mapping):
            data = dict(arg)
        else:
            raise ValueError(f"Cannot interpret `{arg}`.")
        # If specified, update is performed.
        data.update(kwargs)
        return data, api

    @classmethod
    def _get_name(cls):
        if cls.name:
            name = cls.name
        else:
            name = cls.__name__
        return name


class ObjectClassMixin:
    """Provide functions useful for classes which corresponds to Object Class.


    Provided Service
    -----------------
    * Fetch / Setter of `api` attribute.
    * __getattr__ / __setattr__ are revised so that when attribute is not found for the class,

    Extension Cases
    ----------------
    * Implement `fetch_api` for the case when `arg` is None or other types.
    * Set `name` attribute if the name of the class is not equal to Object Type Name.
    """

    name = None

    def __init__(self, arg=None):
        self.cls = type(self)
        if self.cls.name:
            self.name = name
        else:
            self.name = self.cls.__name__

        self._api = self._fetch_api(arg)

    def fetch_api(self, arg):
        """Fetch `api` from `arg`."""
        raise ValueError("Cannot interpret `arg`; `{arg}`.")

    @property
    def api(self):
        return self._api

    @property
    def app(self):
        from fairypptx import Application

        return Application(self._api.Application)

    def _fetch_api(self, arg):
        if isinstance(arg, self.cls):
            return arg.api
        elif is_object(arg, self.name):
            return arg
        return self.fetch_api(arg)

    def __getattr__(self, name):
        if "_api" not in self.__dict__:
            raise AttributeError
        return getattr(self.__dict__["_api"], name)

    def __setattr__(self, name, value):
        if "_api" not in self.__dict__:
            object.__setattr__(self, name, value)

        if name in self.__dict__ or name in type(self).__dict__:
            object.__setattr__(self, name, value)
        elif hasattr(self.api, name):
            setattr(self.api, name, value)
        else:
            # TODO: Maybe require modification.
            object.__setattr__(self, name, value)


class ObjectItems:
    """Utility class for handing `api.Item(index)` function.
    Args
        api_object: `.api` object for the parent.
                    This must have the attribute `Item`.
        child_cls: The wrapper class used for each element of Items.
    """

    def __init__(self, api_object, child_cls):
        self.api = api_object
        self.cls = child_cls

    def __len__(self):
        return self.api.Count

    def __getitem__(self, key):
        if isinstance(key, (int, np.integer)):
            if key < 0:
                key = key + len(self) - 1
            if not (0 <= key < len(self)):
                raise IndexError(
                    f"Size is {len(self)}, index is {key} is out of range."
                )
            return self.cls(self.api.Item(key + 1))
        elif isinstance(key, slice):
            indices = range(*key.indices(len(self)))
            return [self[index] for index in indices]
        elif isinstance(key, Sequence):
            return [self[index] for index in key]
        raise TypeError(
            f"Key's type is invalid; key = `{key}`, type(key) = `{type(key)}`"
        )

    def normalize(self, key):
        """Return int or Sequence of int, which represents indices.

        If key is int, the return is in [0, len(sel) - 1].
        If key specifies, slice or Sequence,  each element of the return is in [0, len(self) - 1].

        Raise:
            IndexError.
            TypeError.
        """
        if isinstance(key, int):
            if key < 0:
                key = key + len(self)
            if not (0 <= key < len(self)):
                raise IndexError(
                    f"Size is {len(self)}, index is {key} is out of range."
                )
            return key

        elif isinstance(key, slice):
            indices = list(range(*key.indices(len(self))))
            return indices

        elif isinstance(key, Sequence):
            for elem in key:
                if not (0 <= elem < len(self)):
                    raise IndexError(
                        f"index out of range; len=`{len(self)}`, but index is `{elem}`"
                    )
            return key

        raise TypeError(f"Invalid Argument; `{key}`")


if __name__ == "__main__":
    pass
