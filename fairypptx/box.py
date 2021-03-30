"""Class for 2D - Rectangle 

When len of Sequence is 4... (Left, Top, Width, Height).
When len of Sequence is 2 (y-coordination, x-coordination).

"""

# TODO: Box should be Immutable. 
from collections import UserDict
from collections.abc import Mapping, Sequence 
from types import GeneratorType

from fairypptx import object_utils

_attributes = ("Left", "Top", "Width", "Height")

class hybridmethod:
    """Hyrbird method of `classmethod` and `instancemethod`. 
    If called as `classmethod`, then `*args` is directly used.
    If called as `instancemethod`, then `self` is added to `*args`. 

    Example:
    @hybridmethod
    def func(cls, *args, **kwargs):
        pass
    """
    def __init__(self, f):
        self.f = f

    def __get__(self, obj, klass=None):
        if klass is None:
            klass = type(obj)
        def newfunc(*args, **kwargs):
            if len(args) == 1:
                if isinstance(args[0], (Sequence, GeneratorType)):
                    args = tuple(args[0])
            if obj:
                args = tuple(args) + (obj, )
            return self.f(klass, *args, **kwargs)
        return newfunc

class EmptySet(Exception):
    """Exception for Empty set.  
    """


class Box(UserDict):
    """ This class keeps (2D-coordination.)
    (``Left``, ``Top``, ``Width``, and ``Height``)
    """

    def __init__(self, *args, **kwargs):
        if args and kwargs:
            raise TypeError("Empty arguments fo `Box`.")
        if bool(args) is False and bool(kwargs) is False:
            raise TypeError("Only either of `*args` or `**kwargs` is accepted.")  

        if not args:
            data = self._construct(kwargs)
        elif len(args) == 1:
            data = self._construct(args[0])
        else: 
            data = self._construct(args)

        self.data = data

    def _construct(self, arg): 
        if isinstance(arg, Mapping):
            return {key: arg[key] for key in _attributes}
        if isinstance(arg, Sequence):
            if len(arg) == 2:
                assert all(isinstance(elem, Interval) for elem in arg)
                return {"Left": arg[1].start,
                        "Top": arg[0].start, 
                        "Width": arg[1].length,
                        "Height": arg[0].length}
            elif len(arg) == 4:
                return dict(zip(_attributes, arg))
            else:
                raise ValueError("Invalid Sequence Argument. `{arg}`.")
        if object_utils.is_object(arg, "Shape"):
            return {key: getattr(arg, key) for key in _attributes}
        raise ValueError(f"Cannot convert to Box with `{arg}`.")

    @property
    def right(self):
        return self.Left + self.Width

    @property
    def bottom(self):
        return self.Top + self.Height

    def __getattr__(self, key):
        n_key = key.capitalize()
        if n_key in self.data:
            return self.data[n_key]
        raise AttributeError(f"`key` {key} is not Found.")

    @property
    def x_interval(self):
        return Interval(self.left, self.right)
    h_interval = x_interval

    @property
    def y_interval(self):
        return Interval(self.top, self.bottom)
    v_interval = y_interval

    @property
    def x_length(self):
        return self.Width
    h_length = x_length

    @property
    def y_length(self):
        return self.Height
    v_length = y_length

    @property
    def center(self):
        # When len of `Sequence` is 2, the order is y - x.
        return ((self.Top + self.Height) / 2, (self.Left + self.Width) / 2, )

    @property
    def area(self):
        return self.Width * self.Height

    def __eq__(self, other):
        return all((getattr(self, name) == getattr(other, name) for name in _attributes))

    @hybridmethod
    def cover(cls, *args):
        y_interval = Interval.cover((arg.y_interval for arg in args))
        x_interval = Interval.cover((arg.x_interval for arg in args))
        return Box(y_interval, x_interval)

    @hybridmethod
    def intersection(cls, *args):
        y_interval = Interval.intersection((arg.y_interval for arg in args))
        x_interval = Interval.intersection((arg.x_interval for arg in args))
        return Box(y_interval, x_interval)

    @classmethod
    def from_vertices(cls, vertices):
        """Return the curcumscribed rectangle of vertices.
        """
        xmin, xmax = min(vertices[:, 0]), max(vertices[:, 0])
        ymin, ymax = min(vertices[:, 1]), max(vertices[:, 1])
        left, top, width, height = xmin, ymin, xmax - xmin, ymax - ymin
        return Box(left, top, width, height)


class Interval:
    """Closed Interval of Real Value.
    [self.start, self.end]
    """
    def __init__(self, *args):
        if len(args) == 1:
            args = args[0]
        if len(args) != 2:
            raise TypeError("Invalid `args`.")
        start, end = args
        start, end = min(start, end), max(start, end)
        self._start = start
        self._end = end

    @property
    def start(self):
        return self._start

    @property
    def end(self):
        return self._end

    @property
    def length(self):
        return self.end - self.start

    @property
    def center(self):
        return (self.start + self.end) / 2

    @hybridmethod
    def cover(cls, *args):
        """Return the minimum Interval which covers
        """
        if not args:
            raise EmptySet
        start = min((arg.start for arg in args))
        end = max((arg.end for arg in args))
        return Interval(start, end)

    @hybridmethod
    def intersection(cls, *args):
        if not args:
            raise EmptySet
        assert isinstance(args[0], Interval)
        start = max((arg.start for arg in args))
        end = min((arg.end for arg in args))
        if end < start:
            raise EmptySet
        return Interval(start, end)

    def issubset(self, other):
        return other.start <= self.start and self.end <= other.end

    def issuperset(self, other):
        return other.issubset(self)

    def __eq__(self, other):
        return (self.start, self.end) == (other.start, other.end)

    def __hash__(self, other):
        return hash((self._start, self._end))


def intersection_over_union(arg1, arg2, axis=None):
    """ Intersection over Union
    """
    if type(arg1) != type(arg2):
        raise TypeError("Type of `arg1` and `arg2` must be equivalent.")

    klass = type(arg1)
    if klass is Interval:
        if axis is not None:
            raise ValueError("For Interval, axis must be None.")
        try:
            nominator = Interval.intersection(arg1, arg2).length
        except EmptySet:
            nominator = 0
        denominator = Interval.cover(arg1, arg2).length
        return nominator / denominator
    elif klass is Box:
        if axis == 0:
            # (y, v) case.
            return intersection_over_union(arg1.y_interval, arg2.y_interval)
        elif axis == 1:
            # (x, h) case.
            return intersection_over_union(arg1.x_interval, arg2.x_interval)
        elif axis is None:
            try:
                nominator = Box.intersection(arg1, arg2).area
            except EmptySet:
                nominator = 0
            denominator = arg1.area + arg2.area - nominator
            return nominator / denominator
        else:
            raise TypeError(f"Invalid `axis` specification.") 
    else:
        raise TypeError(f"Class of argument is invalid; `{klass}`") 

iou = intersection_over_union  # alias.

def intersection_over_cover(*args, axis=None):
    """ Intersection over Cover.
    """
    if not args:
        raise EmptySet
    if not isinstance(args[0], (Interval, Box)):
        assert isinstance(args[0], (Sequence, GeneratorType)) 
        if len(args) == 1:
            args = args[0]
        elif len(args) == 2:
            args = args[0]
            axis = args[1]
        else:
            raise TypeError("Arguments format is invalid.")

    if isinstance(args, GeneratorType):
        args = tuple(args)

    klass = type(args[0])
    if not all( klass is type(arg) for arg in args):
        raise TypeError("Type of `args` must be equivalent.")
    if klass is Interval:
        if axis is not None:
            raise ValueError("For Interval, axis must be None.")
        try:
            nominator = Interval.intersection(args).length
        except EmptySet:
            nominator = 0
        denominator = Interval.cover(args).length
        return nominator / denominator
    elif klass is Box:
        if axis == 0:
            # (y, v) case.
            return intersection_over_cover([arg.y_interval for arg in args])
        elif axis == 1:
            # (x, h) case.
            return intersection_over_cover([arg.x_interval for arg in args])
        elif axis is None:
            try:
                nominator = Box.intersection(args).area
            except EmptySet:
                nominator = 0
            denominator = Box.cover(args).area
            return nominator / denominator
        else:
            raise TypeError(f"Invalid `axis` specification.") 
    else:
        raise TypeError(f"Class of argument is invalid; `{klass}`") 

ioc = intersection_over_cover  # alias.


if __name__  == "__main__":
    pass
