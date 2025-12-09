from typing import cast
from collections import UserString
from pywintypes import com_error
from typing import Any, Literal, TYPE_CHECKING, Sequence, Self, Literal
from PIL import Image
from fairypptx import constants
from fairypptx._shape.mixins import LocationMixin
from fairypptx.constants import msoTrue, msoFalse
from fairypptx.registry_utils import BaseModelRegistry

from fairypptx._shape.box import Box
from fairypptx.object_utils import upstream
from fairypptx.core.types import COMObject 

from fairypptx.core.resolvers import resolve_shape 
from fairypptx.core.utils import swap_props 

from fairypptx._shape import FillFormatProperty
from fairypptx._shape import LineFormatProperty
from fairypptx._shape import TextProperty, TextsProperty
from fairypptx._shape import api_functions
from fairypptx import registry_utils

if TYPE_CHECKING:
    from fairypptx import ShapeRange


class Shape(LocationMixin):
    line = LineFormatProperty()
    fill = FillFormatProperty()
    text = TextProperty()
    texts = TextsProperty()

    def __new__(cls, arg: Any = None) -> "Shape":
        # NOTE: For the direction of the dependency, 
        # `Factory` is imported here. 
        klass = ShapeFactory.get_class(arg)
        return super().__new__(klass)

    def __init__(self, arg=None):
        self._api = resolve_shape(arg) 

    @property
    def api(self) -> COMObject:
        return self._api

    @property
    def shapes_api(self) -> COMObject:
        return self.api.Parent.Shapes

    @property
    def box(self):
        return Box.from_api(self.api)


    def select(self, replace=True):
        return self.api.Select(replace)

    def resize(self, *, fontsize: int | None = None):
        from fairypptx._text.editor import FontResizer 
        if fontsize is not None:
            FontResizer(fontsize=fontsize, mode="min")(self.textrange)
            self.tighten()
            return self
        raise NotImplementedError("Yet, not implemented.")


    @property
    def slide(self) -> "Slide":
        from fairypptx.slide import Slide
        return Slide(upstream(self.api, "Slide"))

    @property
    def textrange(self):
        # Return `TextRange`.
        from fairypptx import TextRange
        return TextRange(self.api.TextFrame.TextRange)

    @textrange.setter
    def textrange(self, value):
        self.text = value

    @classmethod
    def make(cls, arg, **kwargs) -> "Shape":
        return ShapeFactory.make(arg, **kwargs)


    @classmethod
    def make_textbox(cls, arg, **kwargs):
        return ShapeFactory.make_textbox(arg, **kwargs)

    @classmethod
    def make_arrow(cls, arg: Literal["right", "left", "up", "down", "both"] = "right"):
        return ShapeFactory.make_textbox(arg)


    def like(self, style: str):
        from fairypptx.editjson.protocols import EditParamProtocol
        basemodel = BaseModelRegistry.fetch("Shape", style)
        basemodel = cast(EditParamProtocol[Shape], basemodel)
        basemodel.apply(self)

    def register(self, sytle: str, style_type: None | str | type=None):
        from fairypptx.editjson.style_type_registry import ShapeStyleTypeRegistry 
        if not isinstance(style_type, type):
            style_type = ShapeStyleTypeRegistry.fetch(style_type) 
        basemodel = style_type.from_entity(self) 
        BaseModelRegistry.put(basemodel, "Shape", sytle)


    def get_styles(self) -> Sequence[str]:
        """Return available styles.
        """
        return BaseModelRegistry.get_keys("Shape")

    def tighten(self, *, oneline: bool =False) -> None:
        """Tighten the Shape according to Text.

        Args:
            oneline: Modify so that text becomes 1 line.
        """
        api_functions.tighten(self.api, oneline=oneline)


    def swap(self, other: Self):
        attrs = ["Left", "Top"]
        swap_props(self.api, other.api, attrs)


    def to_image(self, mode: Literal["RGBA", "RGB"] ="RGBA") -> Image.Image:
        return api_functions.to_image(self.api, mode)


    def is_child(self):
        """Return whether this is child or not. 
        """
        return self.api.Child == constants.msoTrue

    @property
    def parent(self) -> "Shape":
        assert self.is_child()
        return Shape(self.api.ParentGroup)


class GroupShape(Shape):
    def ungroup(self) -> "ShapeRange":
        from fairypptx.shape_range import ShapeRange
        return ShapeRange(self.api.Ungroup())

    @property
    def children(self) -> "ShapeRange":
        from fairypptx.shape_range import ShapeRange
        return ShapeRange([elem for elem in self.api.GroupItems])

class TableShape(Shape):
    pass


class ShapeFactory:

    @staticmethod
    def get_class(arg: None) -> type[Shape]:
        """This function is intended to generate a class 
        from PPTXObejct or COMObject. 
        """
        api = resolve_shape(arg)
        # For some `arg`, `Type` is not accessible.
        try:
            t = api.Type
        except com_error:
            t = None
        if t == constants.msoGroup:
            return GroupShape
        return Shape

    # Base on the argument given by user, 
    # Factory selects the apt Shape.

    @staticmethod
    def make(arg: Any, **kwargs):
        if isinstance(arg, Image.Image):
            shape = ShapeFactory.make_shape_with_image(arg)
        elif isinstance(arg, (str, UserString)):
            shape = ShapeFactory.make_textbox(arg, **kwargs)
        elif isinstance(arg, int):
            shape = ShapeFactory.make_shape_from_type(arg, **kwargs)
        else:
            raise ValueError(f"`{type(arg)}`, `{arg}` is not interpretted.")

        from fairypptx._shape.location import ShapesLocator
        ShapesLocator(mode="center")(shape)
        return shape

    @staticmethod
    def make_textbox(arg: str | UserString) -> Shape:
        shape = ShapeFactory.make_shape_from_type(constants.msoShapeRectangle)
        shape.textrange.api.Text = arg
        shape.tighten()
        return shape

    @staticmethod
    def make_shape_from_type(arg: int, **kwargs) -> Shape:
        from fairypptx import Slide
        shapes = Slide().shapes
        return shapes.add(arg, **kwargs)

    @staticmethod
    def make_shape_with_image(arg: Image.Image, **kwargs) -> Shape:
        from fairypptx import Slide
        from fairypptx import Slide
        shapes = Slide().shapes
        with registry_utils.yield_temporary_dump(arg) as path: 
            shape_object = shapes.api.AddPicture(
                path, msoFalse, msoTrue, Left=0, Top=0, Width=arg.size[0], Height=arg.size[1], **kwargs
            )
            shape = Shape(shape_object)
            shape.width = arg.size[0] 
            shape.height = arg.size[1]
        return shape

    @staticmethod
    def make_arrow(arg: Literal["right", "left", "up", "down", "both"] = "right") -> Shape:
        m = {"right": constants.msoShapeRightArrow, "left": constants.msoShapeLeftArrow,
         "up": constants.msoShapeUpArrow, "down": constants.msoShapeDownArrow, 
         "both": constants.msoShapeLeftRightArrow}
        return ShapeFactory.make_shape_from_type(m[arg])


# High-level APIs are loaded here.
#
from fairypptx._shape.replace import replace
from fairypptx._shape.editor import (
        ShapesEncloser, TitleProvider, BoundingResizer, ShapesResizer)
from fairypptx._shape.selector import ShapesSelector as Selector
from fairypptx._shape.maker import PaletteMaker


if __name__ == "__main__":
    pass
