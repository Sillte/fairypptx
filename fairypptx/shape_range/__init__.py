from typing import Self, Sequence, Iterator
from collections.abc import Sequence as SeqABC

from fairypptx.core.types import COMObject
from fairypptx.core.application import Application
from fairypptx import constants
from fairypptx.shape import Shape, GroupShape
from fairypptx.object_utils import is_object
from fairypptx.core.resolvers import resolve_shape_range

from fairypptx._shape.location import ShapesAdjuster, ShapesAligner, ClusterAligner

class ShapeRange:
    def __init__(self, arg: COMObject | Sequence[COMObject] |
                       Self | Sequence[Shape] | None = None):
        self._shapes: list[Shape] = self._solve_shapes(arg)

    def __len__(self) -> int:
        return len(self._shapes)

    def __iter__(self) -> Iterator[Shape]:
        yield from self._shapes

    def __getitem__(self, key: int | slice) -> "Shape | ShapeRange":
        if isinstance(key, int):
            return self._shapes[key]
        elif isinstance(key, slice):
            return ShapeRange(self._shapes[key])
        else:
            raise TypeError(f"Invalid key: {key!r}")

    def select(self, append: bool = False) -> Self:
        """ Select.
        """
        app = Application()
        wnd = app.api.ActiveWindow
        if wnd is None:
            raise RuntimeError("No active window to select shapes.")
        if not append:
            wnd.Selection.Unselect()
        for shape in self:
            shape.api.Select(constants.msoFalse)
        return self

    def group(self) -> Shape:
        """
        Side Effect:
            `Selction` changes.
        """
        self.select()
        App = Application()
        wnd = App.api.ActiveWindow
        if wnd is None:
            raise RuntimeError("No active window to select shapes.")
        shape_object = wnd.Selection.ShapeRange.Group()
        return Shape(shape_object)

    @property
    def leafs(self) -> "ShapeRange":
        """Return Shapes. Each shape of the return is not `msoGroup`.
        """
        def _inner(shape: Shape) -> list[Shape]:
            if isinstance(shape, GroupShape):
                return list(shape.children)
            else:
                return [shape]
        shape_list: Sequence[Shape] = sum((_inner(elem) for elem in self), [])
        return ShapeRange(shape_list)

    @property
    def slide(self):
        from fairypptx.slide import Slide
        return Slide(self.api.Item(1).Parent)

    @property
    def api(self) -> COMObject:
        """Reconstruct COM ShapeRange from stored Shape"""
        if not self._shapes:
            msg = "It is impossible to get COMObject for the empty range."
            raise ValueError(msg)
        shapes_api = self._shapes[0].api.Parent  # COM 
        if is_object(shapes_api, "Slide"):
            shapes_api = shapes_api.Shapes
        names = [s.api.Name for s in self._shapes]
        return shapes_api.Range(names)


    def align_cluster(self,
                      axis=None,
                      mode="start",
                      iou_thresh=0.10):
        """ Perform alignment. `align` is applied 
        by the unit of group(cluster).
        """
        return ClusterAligner(axis=axis, mode=mode, iou_thresh=iou_thresh)(self)

    def align(self, axis=None, mode="start"):
        """Align (Make the edge coordination). 
        """
        return ShapesAligner(axis=axis, mode=mode)(self)

    def adjust(self, axis=None):
        """Adjust (keeping the equivalent distance.)
        """
        return ShapesAdjuster(axis=axis)(self)


    def _solve_shapes(self, arg) -> list[Shape]:
        """Normalize input → list[Shape]"""

        if is_object(arg, "ShapeRange"):
            return [Shape(arg.Item(i + 1)) for i in range(arg.Count)]

        # 2) Python list of Shape
        if isinstance(arg, SeqABC) and not isinstance(arg, (str, bytes)):
            if all(isinstance(s, Shape) for s in arg):
                return list(arg)

            if all(is_object(s, "Shape") for s in arg):
                return [Shape(s) for s in arg]

        if isinstance(arg, ShapeRange):
            return list(arg._shapes)

        # 4) Fallback: resolve as COM ShapeRange
        api = resolve_shape_range(arg)
        
        # Just to be safe.
        if is_object(api, "Shape"):
            return [Shape(api)]
        return self._solve_shapes(api)
