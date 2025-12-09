"""Utility functions of Objects.
 
Utility functions for handling Objects.

"""
from typing import Sequence, Any, cast
import builtins
from win32com.client import DispatchBaseClass, CoClassBaseClass
from pywintypes import com_error
from contextlib import contextmanager
from collections.abc import Sequence

import numpy as np



def get_type(instance):
    """Return the Capitalized Object Type Name."""
    if instance is None:
        return None
    return instance.__class__.__name__.strip("_").capitalize()


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
