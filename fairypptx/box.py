from typing import Sequence, Self, Literal, overload, cast, assert_never, Mapping
from collections.abc import Sequence
from dataclasses import dataclass
from fairypptx.core.types import COMObject


class EmptySet(Exception):
    """Exception for Empty set."""


type X = float
type Y = float


@dataclass(frozen=True)
class Interval:
    start: float
    end: float

    @classmethod
    def from_tuple(cls, t: tuple[float, float]) -> Self:
        """For historical reason."""
        vals = sorted(t)
        return cls(start=vals[0], end=vals[1])

    @property
    def length(self):
        return self.end - self.start

    @property
    def center(self):
        return (self.start + self.end) / 2

    @overload
    @classmethod
    def cover(cls, __intervals: Sequence[Self]) -> Self:
        ...

    @overload
    @classmethod
    def cover(cls, *args: Self) -> Self:
        ...

    @classmethod
    def cover(cls, *args: Self | Sequence[Self]) -> Self:
        if len(args) == 1 and isinstance(args[0], Sequence):
            intervals = cast(Sequence[Self], args[0])
        else:
            intervals = cast(Sequence[Self], list(args))

        if not intervals:
            raise EmptySet()

        start = min(iv.start for iv in intervals)
        end = max(iv.end for iv in intervals)

        return cls(start=start, end=end)

    @overload
    @classmethod
    def intersection(cls, __intervals: Sequence[Self]) -> Self:
        ...

    @overload
    @classmethod
    def intersection(cls, *args: Self) -> Self:
        ...

    @classmethod
    def intersection(cls, *args: Self | Sequence[Self]) -> Self:
        if len(args) == 1 and isinstance(args[0], Sequence):
            intervals = cast(Sequence[Self], args[0])
        else:
            intervals = cast(Sequence[Self], list(args))
        if not intervals:
            raise EmptySet()
        start = max((arg.start for arg in intervals))
        end = min((arg.end for arg in intervals))
        if end < start:
            raise EmptySet()
        return cls(start, end)

    def issubset(self, other: Self) -> bool:
        return other.start <= self.start and self.end <= other.end

    def issuperset(self, other: Self) -> bool:
        return other.issubset(self)

    @classmethod
    def intersection_over_union(cls, interval1: Self, interval2: Self) -> float:
        denominator = cls.cover([interval1, interval2])
        try:
            nominator = cls.intersection([interval1, interval2])
        except EmptySet:
            return 0
        return nominator.length / denominator.length


@dataclass(frozen=True)
class Box:
    # These names comes from the lower strings of COMObject.
    left: float
    top: float
    width: float
    height: float

    @property
    def center(self) -> tuple[Y, X]:
        return (self.top + self.height / 2, self.left + self.width / 2)

    @property
    def right(self) -> X:
        return self.left + self.width

    @property
    def bottom(self) -> Y:
        return self.top + self.height

    @property
    def area(self) -> float:
        return self.height * self.width

    @property
    def y_interval(self) -> Interval:
        return Interval(start=self.top, end=self.top + self.height)

    @property
    def y_length(self) -> float:
        return self.y_interval.length

    @property
    def x_interval(self) -> Interval:
        return Interval(start=self.left, end=self.left + self.width)

    @property
    def x_length(self) -> float:
        return self.x_interval.length

    @classmethod
    def from_intervals(cls, y_interval: Interval, x_interval: Interval) -> Self:
        return cls(
            left=x_interval.start,
            width=x_interval.end - x_interval.start,
            top=y_interval.start,
            height=y_interval.end - y_interval.start,
        )

    @classmethod
    def from_numbers(cls, left: float, top: float, width: float, height: float) -> Self:
        return cls(left=left, top=top, width=width, height=height)

    @classmethod
    def from_tuple(cls, t: tuple[float, float, float, float]) -> Self:
        return cls(left=t[0], top=t[1], width=t[2], height=t[3])

    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        return cls(left=api.Left, top=api.Top, width=api.Width, height=api.Height)

    @classmethod
    def from_dict(cls, t: Mapping[str, float]) -> Self:
        data = {key.lower(): val for key, val in t.items()}
        return cls(left=data["left"], top=data["top"], width=data["width"], height=data["height"])

    @overload
    @classmethod
    def cover(cls, __boxes: Sequence[Self]) -> Self:
        ...

    @overload
    @classmethod
    def cover(cls, *args: Self) -> Self:
        ...

    @classmethod
    def cover(cls, *args: Self | Sequence[Self]) -> Self:
        if len(args) == 1 and isinstance(args[0], Sequence):
            boxes = cast(Sequence[Self], args[0])
        else:
            boxes = cast(Sequence[Self], list(args))
        y_interval = Interval.cover([box.y_interval for box in boxes])
        x_interval = Interval.cover([box.x_interval for box in boxes])
        return cls.from_intervals(y_interval=y_interval, x_interval=x_interval)

    @overload
    @classmethod
    def intersection(cls, __boxes: Sequence[Self]) -> Self:
        ...

    @overload
    @classmethod
    def intersection(cls, *args: Self) -> Self:
        ...

    @classmethod
    def intersection(cls, *args: Self | Sequence[Self]) -> Self:
        if len(args) == 1 and isinstance(args[0], Sequence):
            boxes = cast(Sequence[Self], args[0])
        else:
            boxes = cast(Sequence[Self], list(args))
        y_interval = Interval.intersection([box.y_interval for box in boxes])
        x_interval = Interval.intersection([box.x_interval for box in boxes])
        return cls.from_intervals(y_interval=y_interval, x_interval=x_interval)

    @classmethod
    def intersection_over_union(cls, box1: Self, box2: Self, *, axis: Literal["y", 0, "x", 1] | None = None, instead_cover: bool=False) -> float:
        """
        Calculate IoU. If `axis` is given, IoU of `Interval` is calculated.
        """
        if axis == "y":
            axis = 0
        if axis == "x":
            axis = 1

        match axis:
            case 0:
                return Interval.intersection_over_union(box1.y_interval, box2.y_interval)
            case 1:
                return Interval.intersection_over_union(box1.x_interval, box2.x_interval)
            case None:
                try:
                    nominator = cls.intersection([box1, box2]).area
                except EmptySet:
                    nominator = 0
                if not instead_cover:
                    denominator = denominator = box1.area + box2.area - nominator
                else:
                    denominator = cls.cover(box1, box2).area
                if denominator == 0:
                    return 1.0 if nominator > 0 else 0.0
                return nominator / denominator
            case _ as unreachable:
                assert_never(unreachable)

    @classmethod
    def intersection_over_cover(cls, box1: Self, box2: Self, *, axis: Literal["y", 0, "x", 1] | None = None) -> float:
        return cls.intersection_over_union(box1, box2, axis=axis, instead_cover=True)


    @classmethod
    def center_distance(cls, box1: Self, box2: Self, *, axis: Literal["y", 0, "x", 1] | None = None) -> float:
        if axis == "y":
            axis = 0
        if axis == "x":
            axis = 1

        match axis:
            case 0:
                return abs(box1.center[0] - box2.center[0])
            case 1:
                return abs(box1.center[1] - box2.center[1])
            case None:
                return abs(box1.center[0] - box2.center[0]) + abs(box1.center[1] - box2.center[1])
            case _ as unreachable:
                assert_never(unreachable)
        



