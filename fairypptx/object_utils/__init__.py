"""Utility functions of Objects.
 
Utility functions for handling Objects.

"""
from typing import Self, Sequence, Any, cast
import builtins
from win32com.client import CDispatch, DispatchBaseClass, CoClassBaseClass
from pywintypes import com_error
from contextlib import contextmanager
from collections.abc import Sequence
from collections import UserDict, UserString
from collections.abc import Mapping

import numpy as np

from fairypptx import registry_utils


def get_type(instance):
    """Return the Capitalized Object Type Name."""
    if instance is None:
        return None
    return instance.__class__.__name__.strip("_").capitalize()
    #except AttributeError: 
    #    return None
    # return getattr(type(instance), "__com_interface__").__name__.strip("_").capitalize()


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
    flag = isinstance(instance, (DispatchBaseClass, CoClassBaseClass))
    c_name = get_type(instance)
    if name:
        return flag and (name.capitalize() == c_name)
    else:
        return flag


def upstream(instance, name):
    """Return `name` Object as long as `instance.parent..parent` is achievable.

    Args:
        instance: instance.
        name(str): the Target ObjectType Name.

    Raises: ValueError:
    """
    target = instance
    while True:
        type_name = get_type(target)
        # print("type_name", type_name, name.capitalize())
        if type_name.capitalize() == name.capitalize():
            return target
        try:
            target = target.Parent
        except AttributeError as e:
            raise ValueError(f"`{name}` is not an ancestor of `{instance.__class__}`.")


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
    except com_error as e:
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


def to_api2(api_object):
    """Convert to the second version Object.

    In Powerpoint Object model, 
    some classes have 2 apis. (e.g. `TextRange` and `TextRange2`.) 

    This function returns 2 version of `api`.  

    TODO: This list is not complete. 
    """
    if is_object(api_object) and get_type(api_object).endswith("2"):
        return api_object

    def _to_textrange2_api(textrange_api):
        shape_obj = upstream(api_object, "Shape")
        start = textrange_api.Start
        length = textrange_api.Length
        return shape_obj.TextFrame2.TextRange.GetCharacters(start, length)

    if is_object(api_object, "TextFrame"):
        shape_api = upstream(api_object, "Shape")
        return shape_api.TextFrame2
    if is_object(api_object, "TextRange"):
        return _to_textrange2_api(api_object)
    if is_object(api_object, "Font"):
        tr_api = upstream(api_object, "TextRange")
        tr_api2 = _to_textrange2_api(tr_api)
        return tr_api2.Font
    if is_object(api_object, "ParagraphFormat"):
        tr_api = upstream(api_object, "TextRange")
        tr_api2 = _to_textrange2_api(tr_api)
        return tr_api2.ParagraphFormat
    if hasattr(api_object, "api"):
        return to_api2(api_object)  # Maybe high-level class is given? 
    raise ValueError("Cannot interpret the given `api_object`. ", api_object.__class__)


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


    Usage
    ----------------------------------

    * `apply`: When you want to apply the content of `this class`
               to `Object`, using the stored dict.
    * `to_dict`: You you want to store the data from `Object` as `dict`.

    - When you implement `to_dict`, typically you also implement`apply`.
    - Some properties are read-only, so you may have to call the function other than `setattr`.
      For those cases, please implement your customized `apply`.

    The expected protocol of these classes as follows:
    what this class intend to realize this protocol.

    - __init__(self, arg):
        - If `arg` is Object, then `self.api` points to `arg` and `self.data` represents 
        - If `arg` is `Mapping`, then `self.api` is None  and `self.data` is equivalent to `arg`.  

    - def apply(self, api_object): 
        - `self.data`'s contents are reflected onto `api_object`.   

    - `fetch` / `register` / `api`:  (You can easily guess the contents.)

    I feel if these interfaces are realized, the feel of usage does not change from perspective of users.  
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
        name = self.name
        registry_utils.register(name, key, self.data, extension=".json", disk=disk)

    @classmethod
    def fetch(cls, key, disk=True):
        """Construct the instance with `key` object."""
        name = self.name
        data = registry_utils.fetch(name, key, disk=True)
        return cls(data)

    def _construct(self, arg, **kwargs):
        name = self.name
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



class ObjectDictMixin2(Mapping):
    """Represents `api` / `api2`'s extension for

    For some `Objects`, `api` and `api2` exist, 
    (c.f. ParagraphFormat, TextRange)

    data2 is added to `ObjectDictMixin`
    """

    data = dict()
    data2 = dict()

    readonly = []
    name = None

    def __init__(self, arg=None):
        cls = type(self)
        if cls.name is None:
            cls.name = cls.__name__
        self._construct(arg)
        assert bool(self.data) or bool(self.data2)


    def to_dict(self, api):
        cls = type(self)
        return {key: getattr(api, key) for key in self.data}

    def to_dict2(self, api2):
        cls = type(self)
        return {key: getattr(api2, key) for key in self.data2}

    def apply(self, api):
        api2 = to_api2(api)
        readonly_props = set(self.readonly)

        for key, value in self.data.items():
            if key not in readonly_props:
                if value is not None:
                    setattr(api, key, value)

        for key, value in self.data2.items():
            if key not in readonly_props:
                if value is not None:
                    setattr(api2, key, value)
        return api

    @property
    def api(self):
        return self._api

    @property
    def api2(self):
        if self.api:
            return to_api2(self.api)
        else:
            return None

    def to_api2(self, api):
        return to_api2(api)


    def _construct(self, arg):
        cls = type(self)
        self._api = None
        if arg is None:
            self.data = cls.data.copy()
            self.data2 = cls.data2.copy()
        elif is_object(arg, cls.name):
            self._api = arg
            self.data = self.to_dict(self.api)
            self.data2 = self.to_dict2(self.api2)
        elif isinstance(arg, Mapping):
            for key, value in arg.items(): 
                self[key] = value
        else:
            raise ValueError("Given `arg` is not appropriate.", arg)


    def __repr__(self):
        return repr(self.data) + "\n" + repr(self.data2)

    def items(self):
        import itertools 
        return itertools.chain(self.data.items(), self.data2.items())

    def __len__(self):
        return len(self.data) + len(self.data2)

    def __getitem__(self, key):
        if key in self.data:
            return self.data[key]
        if key in self.data2:
            return self.data2[key]
        raise KeyError(key)

    def __setitem__(self, key, item):
        is_api_set =  self.api and (key not in self.readonly)

        if key in self.data:
            self.data[key] = item
            if is_api_set:
                if item is not None:
                    setattr(self.api, key, item)
            return 
        elif key in self.data2:
            self.data2[key] = item
            if is_api_set:
                if item is not None:
                    setattr(self.api2, key, item)
            return
        elif is_api_set: 
            # This is a fallback, 
            try:
                setattr(self.api, key, item)
            except AttributeError as e:
                pass
            else:
                self.data[key] = item
                return 
            try:
                setattr(self.api2, key, item)
            except AttributeError as e:
                pass
            else:
                self.data2[key] = item
                return 
        raise KeyError("Cannot set key", key)

    def __delitem__(self, key):
        del self.data[key]
        del self.data2[key]

    def __iter__(self):
        import itertools
        return itertools.chain(iter(self.data), iter(self.data2))

    def __contains__(self, key):
        return key in self.data or key in self.data2

    def __getstate__(self):
        """For `pickle` serialization
        """
        return {"name": self.name,
                "data": self.data,
                "data2": self.data2,
                "readonly": self.readonly}

    def register(self, key, disk=False):
        """Register to the storage."""
        name = type(self).name
        registry_utils.register(name, key, self, extension=".pkl", disk=disk)


    @classmethod
    def fetch(cls, key, disk=True):
        """Construct the instance with `key` object."""
        name = cls.name
        return registry_utils.fetch(name, key, disk=True)



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


class ObjectItems[T]:
    """Utility class for handing `api.Item(index)` function.
    Args
        api_object: `.api` object for the parent.
                    This must have the attribute `Item`.
        child_cls: The wrapper class used for each element of Items.
    """

    def __init__(self, api_object: Any, child_cls: type[T]):
        self._api = api_object
        self.cls = child_cls
        
    @property
    def api(self):
        return self._api

    def __len__(self) -> int:
        return self.api.Count

    def __getitem__(self, key: int | slice | Sequence[int]) -> T | list[T]:
        if isinstance(key, (int, np.number)):
            key = int(key)
            if key < 0:
                key = key + len(self)
            if not (0 <= key < len(self)):
                raise IndexError(
                    f"Size is {len(self)}, index is {key} is out of range."
                )
            return cast(T, self.cls(self.api.Item(key + 1)))

        elif isinstance(key, slice):
            indices = range(*key.indices(len(self)))
            return cast(list[T], [self[index] for index in indices])
        elif isinstance(key, Sequence):
            return cast(list[T], [self[index] for index in key])
        raise TypeError(
            f"Key's type is invalid; key = `{key}`, type(key) = `{type(key)}`"
        )

    def normalize(self, key: int | slice | Sequence[int]) -> int | Sequence[int]:
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
