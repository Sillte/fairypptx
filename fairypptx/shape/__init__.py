from typing import cast, Self, Any
from collections import UserString
from pywintypes import com_error
from typing import Any, Literal, TYPE_CHECKING, Sequence, Self

from PIL import Image
from fairypptx import constants
from fairypptx.shape.mixins import LocationMixin
from fairypptx.registry_utils import BaseModelRegistry

from fairypptx.box import Box
from fairypptx.object_utils import upstream
from fairypptx.core.types import COMObject 

from fairypptx.core.resolvers import resolve_shape 
from fairypptx.core.utils import swap_props 
from fairypptx.apis.shape.api_factory import ShapeApiFactory

from fairypptx.fill_format import FillFormatProperty
from fairypptx.line_format import LineFormatProperty
from fairypptx.apis.shape import api_functions
from fairypptx.text_range import TextRange
from fairypptx.text_frame import TextFrameProperty

if TYPE_CHECKING:
    from fairypptx import ShapeRange
    from fairypptx import Slide


class Shape(LocationMixin):
    line = LineFormatProperty()
    fill = FillFormatProperty()
    text_frame = TextFrameProperty()


    def __new__(cls, arg: Any = None) -> "Shape":
        api = resolve_shape(arg)
        # For some `arg`, `Type` is not accessible.
        try:
            t = api.Type
        except com_error:
            t = None
        match t:
            case constants.msoGroup:
                klass = GroupShape
            case _:
                klass = cls
        return object.__new__(klass)

    def __init__(self, arg=None):
        self._api = resolve_shape(arg) 

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, Shape):
            return NotImplemented
        return self.api.Id == other.api.Id

    def __hash__(self):
        return hash(self.api.Id)


    @property
    def api(self) -> COMObject:
        return self._api

    @property
    def shapes_api(self) -> COMObject:
        return self.api.Parent.Shapes

    @property
    def box(self):
        return Box.from_api(self.api)

    @box.setter
    def box(self, box: Box):
        self.left, self.top, self.width, self.height = box.left, box.top, box.width, box.height

    def select(self, replace_: bool=True):
        return self.api.Select(replace_)

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
    def text_range(self) -> TextRange:
        return self.text_frame.text_range

    @text_range.setter
    def text_range(self, value: TextRange):
        self.text_frame.text_range = value

    # For backward compatibility, 
    @property
    def textrange(self) -> TextRange:
        return self.text_frame.text_range

    @textrange.setter
    def textrange(self, value: TextRange):
        self.text_frame.text_range = value

    @property
    def text(self) -> str:
        return self.text_range.text

    @text.setter
    def text(self, value: str | TextRange):
        self.text_frame.text_range = value

    @classmethod
    def from_box(cls, box: Box):
        return ShapeFactory.from_box(box)

    @classmethod
    def make(cls, arg, **kwargs) -> "Shape":
        return ShapeFactory.make(arg, **kwargs)


    @classmethod
    def make_textbox(cls, arg, **kwargs):
        return ShapeFactory.make_textbox(arg, **kwargs)

    @classmethod
    def make_arrow(cls, arg: Literal["right", "left", "up", "down", "both"] = "right"):
        return ShapeFactory.make_arrow(arg)


    def like(self, style: str):
        from fairypptx.styles.protocols import StyleModelProtocol
        basemodel = BaseModelRegistry.fetch("Shape", style)
        basemodel = cast(StyleModelProtocol[Shape], basemodel)
        basemodel.apply(self)

    def register(self, sytle: str, style_type: None | str | type=None):
        from fairypptx.styles.style_type_registry import ShapeStyleTypeRegistry 
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

    def is_tight(self) -> bool:
        return api_functions.is_tight(self.api)


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
    """High-level factory for creating Shape wrappers.

    This factory accepts various input types (Image, str, int) and returns
    Shape wrappers ready for use. Internally, it uses ShapeApiFactory to
    create the low-level COM objects, then wraps them using Shape.from_api().
    """

    @staticmethod
    def make(arg: Any, **kwargs) -> Shape:
        """Create a Shape from an image, text string, or shape type constant.

        Args:
            arg: PIL.Image, str/UserString (text), or int (shape type constant).
            **kwargs: Additional arguments for shape positioning/sizing.

        Returns:
            A Shape wrapper (or GroupShape if appropriate).
        """
        if isinstance(arg, Image.Image):
            shape_api = ShapeApiFactory.add_picture(arg, **kwargs)
        elif isinstance(arg, (str, UserString)):
            shape_api = ShapeApiFactory.add_textbox(str(arg), **kwargs)
        elif isinstance(arg, int):
            shape_api = ShapeApiFactory.add_shape_from_type(arg, **kwargs)
        else:
            raise ValueError(f"Unsupported arg type: {type(arg)}, value: {arg}")

        shape = Shape(shape_api)
        y, x = shape.slide.box.center
        shape.top = y - shape.width  / 2
        shape.left = x - shape.height / 2
        return shape

    @staticmethod
    def make_textbox(text: str, **kwargs) -> Shape:
        """Create a textbox with the given text, auto-sized to content.

        Args:
            text: Text content for the textbox.
            **kwargs: Additional shape arguments.

        Returns:
            A Shape wrapper with text, tightened to content.
        """
        shape_api = ShapeApiFactory.add_textbox(text, **kwargs)
        shape = Shape(shape_api)
        shape.tighten()
        return shape

    @staticmethod
    def from_box(box: Box) -> Shape:
        shape_api = ShapeApiFactory.add_shape_from_type(1)
        shape = Shape(shape_api)
        shape.box = box
        return shape

    @staticmethod
    def make_shape_from_type(type_: int, **kwargs) -> Shape:
        """Create a shape of the specified type.

        Args:
            type_id: COM shape type constant (e.g., constants.msoShapeRectangle).
            **kwargs: Additional shape arguments.

        Returns:
            A Shape wrapper.
        """
        shape_api = ShapeApiFactory.add_shape_from_type(type_, **kwargs)
        return Shape(shape_api)

    @staticmethod
    def make_shape_with_image(image: Image.Image, **kwargs) -> Shape:
        """Add a picture to the slide.

        Args:
            image: PIL.Image object.
            **kwargs: Additional shape arguments (position, size overrides).

        Returns:
            A Shape wrapper around the picture.
        """
        shape_api = ShapeApiFactory.add_picture(image, **kwargs)
        shape = Shape(shape_api)
        # Set explicit dimensions if not already provided by AddPicture
        shape.width = image.width
        shape.height = image.height
        return shape

    @staticmethod
    def make_arrow(direction: Literal["right", "left", "up", "down", "both"] = "right", **kwargs) -> Shape:
        """Create an arrow shape in the specified direction.

        Args:
            direction: One of "right", "left", "up", "down", "both".
            **kwargs: Additional shape arguments.

        Returns:
            A Shape wrapper for the arrow.
        """
        shape_api = ShapeApiFactory.make_arrow(direction, **kwargs)
        return Shape(shape_api)



if __name__ == "__main__":
    pass
