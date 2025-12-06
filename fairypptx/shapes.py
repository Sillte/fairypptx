from typing import Iterator, Self

from fairypptx.constants import msoFalse

from fairypptx._shape.box import Box  # NOQA
from fairypptx.core.application import Application
from fairypptx.core.resolvers import resolve_shapes
from fairypptx.core.types import COMObject
from fairypptx.shape import Shape
from fairypptx.shape_range import ShapeRange

from fairypptx._shape.location import ShapesAdjuster, ShapesAligner, ClusterAligner


class Shapes:
    """Shapes."""

    def __init__(self, arg=None):
        self._api = resolve_shapes(arg)

    @property
    def api(self) -> COMObject:
        return self._api

    def __len__(self) -> int:
        return self.api.Count

    def __iter__(self) -> Iterator[Shape]:
        for index in range(len(self)):
            yield Shape(self.api.Item(index + 1))

    def __getitem__(self, key: int | slice) -> Shape | ShapeRange:
        if isinstance(key, int):
            return list(self)[key]
        elif isinstance(key, slice):
            shapes = list(self)[key]
            return ShapeRange(shapes)

    def add(self, shape_type: int, **kwargs) -> Shape:
        ret_object = self.api.AddShape(shape_type, Left=0, Top=0, Width=100, Height=100)
        return Shape(ret_object)

    @property
    def slide(self) -> "Slide":
        from fairypptx.slide import Slide
        return Slide(self.api.Parent) 

    @property
    def circumscribed_box(self):
        """Return Box which circumscribes `Shapes`.
        """
        boxes = [shape.box for shape in self]
        c_left = min(box.left for box in boxes)
        c_top = min(box.top for box in boxes)
        c_right = max(box.right for box in boxes)
        c_bottom = max(box.bottom for box in boxes)
        c_box = Box(left=c_left, top=c_top, width=c_right - c_left, height=c_bottom - c_top)
        return c_box

    def select(self) -> Self:
        """ Select.
        """
        app = Application()
        app.api.ActiveWindow.Selection.Unselect()
        for shape in self:
            shape.api.Select(msoFalse)
        return self

    def tighten(self):
        for shape in self:
            shape.tighten()
        return self



if __name__ == "__main__":
    pass
