from typing import Iterator, Self, TYPE_CHECKING, overload

from fairypptx.constants import msoFalse, msoShapeNotPrimitive

from fairypptx.box import Box  # NOQA
from fairypptx.core.application import Application
from fairypptx.core.resolvers import resolve_shapes
from fairypptx.core.types import COMObject
from fairypptx.shape import Shape
from fairypptx.shape_range import ShapeRange


if TYPE_CHECKING:
    from fairypptx.slide import Slide


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

    @overload
    def __getitem__(self, key: int) -> Shape:
        ...

    @overload
    def __getitem__(self, key: slice) -> 'ShapeRange':
        ...

    def __getitem__(self, key: int | slice) -> Shape | ShapeRange:
        if isinstance(key, int):
            return list(self)[key]
        elif isinstance(key, slice):
            shapes = list(self)[key]
            return ShapeRange(shapes)

    def add(self, auto_shape_type: int, **kwargs) -> Shape:
        if auto_shape_type == msoShapeNotPrimitive:
            print(f"{auto_shape_type=} is not supported.")
            auto_shape_type = 1
        shape_object = self.api.AddShape(auto_shape_type, Left=0, Top=0, Width=100, Height=100, **kwargs)
        return Shape(shape_object)

    @property
    def slide(self) -> "Slide":
        from fairypptx.slide import Slide
        return Slide(self.api.Parent) 

    @property
    def circumscribed_box(self) -> Box:
        """Return Box which circumscribes `Shapes`.
        """
        boxes = [shape.box for shape in self]
        return Box.cover(boxes)


    def select(self) -> Self:
        """ Select.
        """
        app = Application()
        app.api.ActiveWindow.Selection.Unselect()
        for shape in self:
            shape.api.Select(msoFalse)
        return self

    def tighten(self) -> None:
        for shape in self:
            shape.tighten()



if __name__ == "__main__":
    pass
