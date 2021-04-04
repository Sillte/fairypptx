from _ctypes import COMError
from collections.abc import Sequence
from collections import UserString
from PIL import Image
from fairypptx import constants
from fairypptx.constants import msoTrue, msoFalse

from fairypptx.color import Color
from fairypptx.box import Box, intersection_over_cover
from fairypptx.application import Application
from fairypptx.slide import Slide
from fairypptx.inner import storage
from fairypptx import object_utils
from fairypptx.object_utils import is_object, upstream, stored

from fairypptx._text import Text
from fairypptx._shape import FillFormat, FillFormatProperty
from fairypptx._shape import LineFormat, LineFormatProperty
from fairypptx._shape import TextProperty, TextsProperty
from fairypptx._shape.stylist import ShapeStylist
from fairypptx._shape import LocationAdjuster
from fairypptx._shape.location import ShapesAdjuster, ShapesAligner
from fairypptx._shape.location import ShapesAdjuster, ShapesAligner, ClusterAligner
from fairypptx import registory_utils


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
    def decomposed(self):
        """Return Shapes. Each shape of the return is not `msoGroup`.
        """

        def _inner(shape):
            print(shape.api.Type)
            if shape.api.Type == constants.msoGroup:
                return sum((_inner(Shape(elem)) for elem in shape.api.GroupItems), [])
            else:
                return [shape]

        shape_list = sum((_inner(elem) for elem in self), [])
        result = Shapes(shape_list)
        assert len(result) == len(shape_list)

        return result

    def select(self):
        """ Select.
        """
        self.app.api.ActiveWindow.Selection.Unselect()
        for shape in self:
            shape.api.Select(msoFalse)
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
                #if all(value == values[0] for value in values):
                #    return value
                #else:
                #    raise ValueError(
                #        (
                #            "Non-equivalent values over the Shapes.",
                #            f"Maybe `{name}` returns Object?",
                #            f"values=`{values}`",
                #        )
                #    )
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
            assert len(set(shapes_objects)) == 1, "All the shapes must belong to the same slide."
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
            except COMError as e:
                # May be `ActiveWindow` does not exist. (esp at an empty file.)
                pass
            else:
                if Selection.Type == constants.ppSelectionShapes:
                    shape_objects = [shape for shape in Selection.ShapeRange]
                    shapes_objects = [
                        shape_object.Parent.Shapes for shape_object in shape_objects
                    ]
                    assert len(set(shapes_objects)) == 1
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
        raise ValueError(f"Cannot interpret `arg`; {arg}.")


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
            path = storage.get_path(".png")
            arg.save(path)
            shape_object = shapes.api.AddPicture(
                path, msoFalse, msoTrue, Left=0, Top=0, Width=100, Height=100
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
        l_adjuster = LocationAdjuster(shape)
        l_adjuster.center()
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
        else:
            raise ValueError(f"Invalid direction.")
        return shape

    def like(self, style):
        if isinstance(style, str):
            stylist = registory_utils.fetch(self.__class__.__name__, style)
            stylist(self)
            return self
        raise TypeError(f"Currently, type {type(style)} is not accepted.")

    def register(self, key, disk=True):
        stylist = ShapeStylist(self)
        registory_utils.register(
            self.__class__.__name__, key, stylist, extension=".pkl", disk=disk
        )

    def get_styles(self):
        """Return available styles.
        """
        return registory_utils.keys(self.__class__.__name__)

    def tighten(self, *, oneline=False):
        """Tighten the Shape according to Text.

        Args:
            oneline: Modify so that text becomes 1 line.
        """
        if self.api.HasTextframe:
            if oneline is True:
                self.api.TextFrame.TextRange.Text = self.text.replace("\r", "").replace(
                    "\n", ""
                )
            with stored(self.api, ("TextFrame.AutoSize", "TextFrame.WordWrap")):
                self.api.TextFrame.Autosize = constants.ppAutoSizeShapeToFitText
                self.api.TextFrame.WordWrap = constants.msoFalse
        return self

    def swap(self, other):
        attrs = ["Left", "Top"]
        ps1 = [object_utils.getattr(self, attr) for attr in attrs]
        ps2 = [object_utils.getattr(other, attr) for attr in attrs]
        for attr, p1, p2 in zip(attr, ps1, ps2):
            object_utils.getattr(self, attr, p2)
            object_utils.getattr(other, attr, p1)
        return self

    def to_image(self, mode="RGBA"):
        path = storage.get_path(".png")
        self.api.Export(path, constants.ppShapeFormatPNG)
        image = Image.open(path)
        return image.convert(mode)

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


from fairypptx._shape.replace import replace

if __name__ == "__main__":
    pass
