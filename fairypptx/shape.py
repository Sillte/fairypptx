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

        self._api, self._indices = self._construct(arg)

        # Sorting mechanism desirable.
    @property
    def api(self):
        return self._api

    def __len__(self):
        return len(self._indices)
   
    def __iter__(self):
        for index in range(len(self)):
            yield self[index]

    def __getitem__(self, key):
        if isinstance(key, int):
            index = self._indices[key]
            return Shape(self.api.Item(index + 1))
        elif isinstance(key, slice):
            indices = self._indices[key]
            shapes = [Shape(self.api.Item(index + 1)) for index in indices]
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

    def select(self):
        """ Select.
        """
        self.app.api.ActiveWindow.Selection.Unselect()
        for shape in self:
            shape.api.Select(msoFalse)
        return self

    def align(self, axis=None, mode="center"):
        """
        Side Effect:
            `Selection` changes.
        """
        self.select()
        if axis is None:
            boxes = [shape.box for shape in self]
            y_ratio = intersection_over_cover(boxes, axis=0)
            x_ratio= intersection_over_cover(boxes, axis=1)
            print("x_ratio", x_ratio)
            print("y_ratio", y_ratio)
            if y_ratio < x_ratio: 
                align_cmd = constants.msoAlignCenters
            else:
                align_cmd = constants.msoAlignMiddles

        shape_object = self.app.api.ActiveWindow.Selection.ShapeRange.Align(align_cmd, msoFalse)

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
                if all(value == values[0] for value in values):
                    return value
                else:
                    raise ValueError(("Non-equivalent values over the Shapes.", 
                                      f"Maybe `{name}` returns Object?", 
                                      f"values=`{values}`"))
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
        """
        """
        if is_object(arg, "ShapeRange"):
            raise NotImplementedError()
            slide_objects = [arg.Item(index + 1) for index in range(arg.Count)]
            slides_objects = [slide.Parent.Slides for slide in slide_objects]
            assert len(set(map(id, slides_objects))) == 1, "Slide must be"
            slides_object = slides_objects[0]
            indices = [elem.SlideIndex - 1 for elem in slide_objects]
            return slides_object, indices
        elif is_object(arg, "Shapes"):
            return arg, tuple(range(arg.Count))
        elif is_object(arg, "Slide"):
            return arg.Shapes, tuple(range(arg.Shapes.Count))
        elif isinstance(arg, Shapes):
            return arg.api, arg.indices
        elif isinstance(arg, Sequence):
            def _to_object(instance):
                if is_object(instance):
                    return instance 
                return instance.api 
            shape_objects = [_to_object(elem) for elem in arg]
            slide_ids = set(elem.Parent.SlideID for elem in shape_objects)
            assert len(slide_ids) == 1, "All the shapes must belong to the same slide."
            shape_ids = set(elem.Id for elem in shape_objects)
            slide_object = shape_objects[0].Parent
            indices = [index for index, elem in enumerate(slide_object.Shapes) 
                       if elem.Id in shape_ids]
            return slide_object.Shapes, indices
                
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
                    shape_ids = set(shape.Id for shape in shape_objects)
                    shapes_objects = [shape_object.Parent.Shapes for shape_object in shape_objects]
                    assert len(set(shapes_objects)) == 1
                    shapes_object = shapes_objects[0]
                    indices = [index for index, elem in enumerate(shapes_object) if elem.Id in shape_ids]
                    return shapes_object, indices
                elif Selection.Type == constants.ppSelectionText:
                    # Even if Seleciton.Type is ppSelectionText, `Selection.ShapeRange` return ``Shape``.
                    shape_object = Selection.ShapeRange(1)
                    shapes_object = shape_object.Parent.Shapes
                    indices = [index for index, elem in enumerate(shapes_object) if elem.Id == shape_object.Id]
                    assert len(indices) == 1
                    return shapes_object, indices
            slide = Slide()
            return slide.api.Shapes, range(slide.api.Shapes.Count)
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
            shape_object = shapes.api.AddPicture(path, msoFalse, msoTrue, Left=0, Top=0, Width=100, Height=100)
            shape = Shape(shape_object)
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
    def make_arrow(cls, arg=None, **kwargs):
        assert arg is None, "Current"
        shape = cls.make(constants.msoShapeRightArrow)
        return shape

    def like(self, style):
        if isinstance(style, str):
            stylist = registory_utils.fetch(self.__class__.__name__, style)
            stylist(self)
            return self

    def register(self, key, disk=True):
        stylist = ShapeStylist(self)
        registory_utils.register(self.__class__.__name__,
                                 key,
                                 stylist,
                                 extension=".pkl",
                                 disk=disk)

    def tighten(self, *, oneline=False):
        """Tighten the Shape according to Text.

        Args:
            oneline: Modify so that text becomes 1 line.
        """
        if self.api.HasTextframe:
            if oneline is True:
                self.api.TextFrame.TextRange.Text = self.text.replace("\r", "").replace("\n", "")
            with stored(self.api,("TextFrame.AutoSize", "TextFrame.WordWrap")):
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
