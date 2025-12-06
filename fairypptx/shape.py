from collections.abc import Sequence
from collections import UserString
from pywintypes import com_error
from typing import Any, Self, Literal
from pywintypes import com_error
from PIL import Image
from fairypptx import constants
from fairypptx.constants import msoTrue, msoFalse

from fairypptx._shape.box import Box
from fairypptx.core.application import Application
from fairypptx import object_utils
from fairypptx.object_utils import is_object, upstream, stored
from fairypptx.core.types import COMObject, PPTXObjectProtocol 

from fairypptx.core.resolvers import resolve_shape 


from fairypptx._text import Text
from fairypptx._shape import FillFormat, FillFormatProperty
from fairypptx._shape import LineFormat, LineFormatProperty
from fairypptx._shape import TextProperty, TextsProperty
from fairypptx._shape.stylist import ShapeStylist
from fairypptx._shape.location import ShapesAdjuster, ShapesAligner, ClusterAligner, ShapesArranger, ShapesLocator
from fairypptx import registry_utils


class LocationMixin:
    """This Mixin handles the functionality of geometry information of `Shape`.
    This Mixin must be applicable to all the `Shape` in the domain of COMObject.
    """

    @property
    def left(self: PPTXObjectProtocol) -> float:
        return self.api.Left

    @left.setter
    def left(self: PPTXObjectProtocol, value: float) -> None:
        self.api.Left = value

    @property
    def top(self: PPTXObjectProtocol) -> float:
        return self.api.Top

    @top.setter
    def top(self: PPTXObjectProtocol, value: float) -> None:
        self.api.Top = value

    @property
    def width(self: PPTXObjectProtocol) -> float:
        return self.api.Width

    @width.setter
    def width(self: PPTXObjectProtocol, value: float) -> None:
        self.api.Width = value

    @property
    def height(self: PPTXObjectProtocol) -> None:
        return self.api.Height

    @height.setter
    def height(self: PPTXObjectProtocol, value: float) -> None:
        self.api.Height = value

    @property
    def size(self: PPTXObjectProtocol) -> tuple[float, float]:
        return (self.api.Width, self.api.Height)

    @size.setter
    def size(self: PPTXObjectProtocol, value: tuple[float, float]) -> None:
        self.api.Width, self.api.Height = value

    @property
    def rotation(self: PPTXObjectProtocol) -> float:
        return self.api.Rotation

    @rotation.setter
    def rotation(self: PPTXObjectProtocol, value: float) -> None: 
        self.api.Rotation = value

    def rotate(self: PPTXObjectProtocol, degree: float) -> None:
        self.api.Rotation += degree


class Shape(LocationMixin):
    line = LineFormatProperty()
    fill = FillFormatProperty()
    text = TextProperty()
    texts = TextsProperty()

    def __new__(cls, arg: Any) -> "Shape":
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

    def resize(self, *, fontsize=None):
        from fairypptx.text import FontResizer
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


    def like(self, style):
        if isinstance(style, str):
            stylist = registry_utils.fetch(self.__class__.__name__, style)
            stylist(self)
            return self
        raise TypeError(f"Currently, type {type(style)} is not accepted.")

    def register(self, key, disk=True):
        stylist = ShapeStylist(self)
        registry_utils.register(
            self.__class__.__name__, key, stylist, extension=".pkl", disk=disk
        )

    def get_styles(self):
        """Return available styles.
        """
        return registry_utils.keys(self.__class__.__name__)

    def tighten(self, *, oneline=False):
        """Tighten the Shape according to Text.

        Args:
            oneline: Modify so that text becomes 1 line.
        """
        if self.api.HasTextFrame:
            if oneline is True:
                self.api.TextFrame.TextRange.Text = self.text.replace("\r", "").replace(
                    "\n", ""
                )
            with stored(self.api, ("TextFrame.AutoSize", "TextFrame.WordWrap")):
                self.api.TextFrame.AutoSize = constants.ppAutoSizeShapeToFitText
                self.api.TextFrame.WordWrap = constants.msoFalse
        return self

    def swap(self, other):
        attrs = ["Left", "Top"]
        ps1 = [object_utils.getattr(self, attr) for attr in attrs]
        ps2 = [object_utils.getattr(other, attr) for attr in attrs]
        for attr, p1, p2 in zip(attrs, ps1, ps2):
            object_utils.setattr(self, attr, p2)
            object_utils.setattr(other, attr, p1)
        return self

    def to_image(self, mode="RGBA"):
        with registry_utils.yield_temporary_path(suffix=".png") as path:
            self.api.Export(path, constants.ppShapeFormatPNG)
            image = Image.open(path).copy()
        return image.convert(mode)


    def is_table(self):
        """Return whether this Shape is Table or not.
        """
        return self.api.Type == constants.msoTable

    def is_child(self):
        """Return whether this is child or not. 
        """
        return self.api.Child == constants.msoTrue

    @property
    def parent(self):
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
        from fairypptx import Shapes 
        if isinstance(arg, Image.Image):
            shape = ShapeFactory.make_shape_with_image(arg)
        elif isinstance(arg, (str, UserString)):
            shape = ShapeFactory.make_textbox(arg, **kwargs)
        elif isinstance(arg, int):
            shape = ShapeFactory.make_shape_from_type(arg, **kwargs)
        else:
            raise ValueError(f"`{type(arg)}`, `{arg}` is not interpretted.")
        ShapesLocator(mode="center")(shape)
        return shape

    @staticmethod
    def make_textbox(arg: str | UserString) -> Shape:
        shape = ShapeFactory.make_shape_from_type(constants.msoShapeRectangle)
        shape.text = arg
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
        with registry_utils.yield_temporary_path(arg) as path: 
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
