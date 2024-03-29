from collections.abc import Sequence
from collections import UserString
from pywintypes import com_error
from PIL import Image
from fairypptx import constants
from fairypptx.constants import msoTrue, msoFalse

from fairypptx.color import Color
from fairypptx.box import Box, intersection_over_cover
from fairypptx.application import Application
from fairypptx.slide import Slide
from fairypptx import object_utils
from fairypptx.object_utils import is_object, upstream, stored

from fairypptx._text import Text
from fairypptx._shape import FillFormat, FillFormatProperty
from fairypptx._shape import LineFormat, LineFormatProperty
from fairypptx._shape import TextProperty, TextsProperty
from fairypptx._shape.stylist import ShapeStylist
from fairypptx._shape.location import ShapesAdjuster, ShapesAligner, ClusterAligner, ShapesArranger, ShapesLocator
from fairypptx import registry_utils


class Shapes:
    """Shapes.
    It accepts a subset of Slide.Shapes Object. 

    Note
    ---------------------
    * `Add` / `Delete` operations may break the indices of this class.

    """

    def __init__(self, arg=None, *, app=None):
        if app is None:
            self.app = Application()
        else:
            self.app = app
        self._api, self._object_list = self._construct(arg)
        assert object_utils.get_type(self._api) == "Shapes"

        # Sorting mechanism desirable.

    @property
    def api(self):
        return self._api

    def __len__(self):
        return len(self._object_list)

    def __iter__(self):
        for index in range(len(self)):
            yield self[index]

    def __getitem__(self, key):
        if isinstance(key, int):
            return Shape(self._object_list[key])
        elif isinstance(key, slice):
            shape_objects = self._object_list[key]
            shapes = [Shape(shape_object) for shape_object in shape_objects]
            return Shapes(shapes)

        raise KeyError(f"`key`: {key}")

    def add(self, shape_type, **kwargs):
        ret_object = self.api.AddShape(shape_type, Left=0, Top=0, Width=100, Height=100)
        return Shape(ret_object)

    def group(self):
        """
        Side Effect:
            `Selction` changes.
        """
        self.select()
        shape_object = self.app.api.ActiveWindow.Selection.ShapeRange.Group()
        return Shape(shape_object)

    @property
    def leafs(self):
        """Return Shapes. Each shape of the return is not `msoGroup`.
        """

        def _inner(shape):
            if shape.api.Type == constants.msoGroup:
                return sum((_inner(Shape(elem)) for elem in shape.api.GroupItems), [])
            else:
                return [shape]

        shape_list = sum((_inner(elem) for elem in self), [])
        result = Shapes(shape_list)
        assert len(result) == len(shape_list)

        return result

    @property
    def slide(self):
        slide_objects = [elem.api.Parent for elem in self]
        assert set(elem.SlideId for elem in slide_objects)
        return Slide(slide_objects[0])

    @property
    def circumscribed_box(self):
        """Return Box which circumscribes `Shapes`.
        """
        boxes = [shape.box for shape in self]
        c_left = min(box.left for box in boxes)
        c_top = min(box.top for box in boxes)
        c_right = max(box.right for box in boxes)
        c_bottom = max(box.bottom for box in boxes)
        c_box = Box(c_left, c_top, c_right - c_left, c_bottom - c_top)
        return c_box

    def select(self):
        """ Select.
        """
        self.app.api.ActiveWindow.Selection.Unselect()
        for shape in self:
            shape.api.Select(msoFalse)
        return self

    def tighten(self):
        for shape in self:
            shape.tighten()
        return self

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

    def __getattr__(self, name):
        if "_api" not in self.__dict__:
            raise AttributeError
        if name.startswith("_"):
            return object.__getattr___(self, name)
        try:
            return getattr(self.__dict__["_api"], name)
        except AttributeError:
            pass

        if self:
            try:
                values = [getattr(shape, name) for shape in self]
            except AttributeError:
                pass
            else:
                return values
        raise AttributeError(f"Cannot find the attribute `{name}`.")

    def __setattr__(self, name, value):
        if "_api" not in self.__dict__:
            object.__setattr__(self, name, value)
            return

        # Especially for ``_indices``.
        if name.startswith("_"):
            object.__setattr__(self, name, value)
            return

        if name in self.__dict__ or name in type(self).__dict__:
            object.__setattr__(self, name, value)
            return
        if hasattr(self.api, name):
            setattr(self.api, name, value)
            return

        if self:
            for shape in self:
                setattr(shape, name, value)
            return

        raise ValueError(f"Cannot find an appropriate attribute by `{name}.`")

    def _construct(self, arg):
        if is_object(arg, "ShapeRange"):
            shape_objects = [elem for elem in arg]
            shapes_objects = [
                shape_object.Parent.Shapes for shape_object in shape_objects
            ]
            assert len(set(elem.Parent.SlideID for elem in shapes_objects)) == 1, "All the shapes must belong to the same slide."
            return shapes_objects[0], shape_objects
        elif is_object(arg, "Shapes"):
            object_list = [arg.Item(index + 1) for index in range(arg.Count)]
            return arg, object_list

        elif is_object(arg, "Slide"):
            object_list = [arg.Item(index + 1) for index in range(arg.Shapes.Count)]
            return arg.Shapes, object_list
        elif isinstance(arg, Shapes):
            return arg.api, arg._object_list
        elif isinstance(arg, Sequence):
            def _to_object(instance):
                if is_object(instance):
                    return instance
                return instance.api

            assert arg, "Empty Shapes is not currently allowed."
            shape_objects = [_to_object(elem) for elem in arg]
            slide_ids = set(elem.Parent.SlideID for elem in shape_objects)
            assert len(slide_ids) <= 1, "All the shapes must belong to the same slide."
            shape_ids = set(elem.Id for elem in shape_objects)
            slide_object = shape_objects[0].Parent
            return slide_object.Shapes, shape_objects

        if arg is None:
            App = self.app.api
            try:
                Selection = App.ActiveWindow.Selection
            except com_error as e:
                # May be `ActiveWindow` does not exist. (esp at an empty file.)
                pass
            else:
                if Selection.Type == constants.ppSelectionShapes:
                    if not Selection.HasChildShapeRange:
                        shape_objects = [shape for shape in Selection.ShapeRange]
                    else:
                        shape_objects = [shape for shape in Selection.ChildShapeRange]
                    shapes_objects = [
                        shape_object.Parent.Shapes for shape_object in shape_objects
                    ]
                    assert len(set(elem.Parent.SlideID for elem in shapes_objects)) == 1, "All the shapes must belong to the same slide."
                    shapes_object = shapes_objects[0]
                    return shapes_object, shape_objects
                elif Selection.Type == constants.ppSelectionText:
                    # Even if Seleciton.Type is ppSelectionText, `Selection.ShapeRange` return ``Shape``.
                    shape_object = Selection.ShapeRange(1)
                    shapes_object = shape_object.Parent.Shapes
                    return shapes_object, [shape_object]
            slide = Slide()
            shape_objects = [elem for elem in slide.api.Shapes]
            return slide.api.Shapes, shape_objects
        raise ValueError(f"Cannot interpret `arg`", arg.__class__)


class Shape:
    line = LineFormatProperty()
    fill = FillFormatProperty()
    text = TextProperty()
    texts = TextsProperty()

    def __init__(self, arg=None, *, app=None):
        if app is None:
            self.app = Application()
        else:
            self.app = app

        self._api = self._fetch_api(arg)

    @property
    def api(self):
        return self._api

    @property
    def box(self):
        return Box(self.api)

    @property
    def left(self):
        return self.api.Left

    @left.setter
    def left(self, value):
        self.api.Left = value

    @property
    def top(self):
        return self.api.Top

    @top.setter
    def top(self, value):
        self.api.Top = value

    @property
    def width(self):
        return self.api.Width

    @width.setter
    def width(self, value):
        self.api.Width = value

    @property
    def height(self):
        return self.api.Height

    @height.setter
    def height(self, value):
        self.api.Height = value

    @property
    def rotation(self):
        return self.api.Rotation

    @rotation.setter
    def rotation(self, value): 
        self.api.Rotation = value

    def rotate(self, degree):
        self.api.Rotation += degree

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
    def size(self):
        return (self.api.Width, self.api.Height)

    @property
    def slide(self):
        return Slide(upstream(self.api, "Slide"), app=self.app)

    @property
    def textrange(self):
        # Return `TextRange`.
        from fairypptx import TextRange

        return TextRange(self.api.TextFrame.TextRange)

    @textrange.setter
    def textrange(self, value):
        self.text = value

    @classmethod
    def make(cls, arg, **kwargs):
        shapes = Shapes()
        if isinstance(arg, Image.Image):
            with registry_utils.yield_temporary_path(arg) as path: 
                shape_object = shapes.api.AddPicture(
                    path, msoFalse, msoTrue, Left=0, Top=0, Width=arg.size[0], Height=arg.size[1],
                )
                shape = Shape(shape_object)
                shape.width = arg.size[0] 
                shape.height = arg.size[1]
        elif isinstance(arg, (str, UserString)):
            shape = cls.make_textbox(arg, **kwargs)
            # TODO: Idetally, interpret of `str` is necessary.
        elif isinstance(arg, int):
            shape = Shape(shapes.add(arg, **kwargs))
        else:
            raise ValueError(f"`{type(arg)}`, `{arg}` is not interpretted.")
        assert isinstance(shape, Shape)
        ShapesLocator(mode="center")(shape)
        return shape


    @classmethod
    def make_textbox(cls, arg, **kwargs):
        assert isinstance(arg, (str, UserString))
        shape = cls.make(constants.msoShapeRectangle)
        shape.text = arg
        shape.tighten()
        return shape

    @classmethod
    def make_arrow(cls, arg=None, direction="right", **kwargs):
        assert arg is None, "Current"
        direction = direction.lower()
        if direction == "right":
            shape = cls.make(constants.msoShapeRightArrow)
        elif direction == "left":
            shape = cls.make(constants.msoShapeLeftArrow)
        elif direction == "up":
            shape = cls.make(constants.msoShapeUpArrow)
        elif direction == "down":
            shape = cls.make(constants.msoShapeDownArrow)
        elif direction == "both":
            shape = cls.make(constants.msoShapeLeftRightArrow)
        else:
            raise ValueError(f"Invalid direction.")
        return shape

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

    def is_leaf(self):
        """Return whether this is NOT Grouped Object.
        """
        return self.api.Type != constants.msoGroup

    def is_child(self):
        """Return whether this is child or not. 
        """
        return self.api.Child == constants.msoTrue


    @property
    def parent(self):
        assert self.is_child()
        return Shape(self.api.ParentGroup)

    @property
    def children(self):
        assert not self.is_leaf()
        return Shapes([elem for elem in self.api.GroupItems])

    def ungroup(self):
        return Shapes(self.api.Ungroup())


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

    def _fetch_api(self, arg):
        if is_object(arg, "Shape"):
            return arg
        elif isinstance(arg, Shape):
            return arg.api
        elif arg is None:
            shapes = Shapes()
            if not shapes:
                raise ValueError("No Shapes.")
            return shapes[0].api
        raise ValueError(f"Cannot interpret `arg`; {arg}.")


# High-level APIs are loaded here.
#
from fairypptx._shape.replace import replace
from fairypptx._shape.editor import (
        ShapesEncloser, TitleProvider, BoundingResizer, ShapesResizer)
from fairypptx._shape.selector import ShapesSelector as Selector
from fairypptx._shape.maker import PaletteMaker


if __name__ == "__main__":
    pass
