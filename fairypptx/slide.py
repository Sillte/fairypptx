from pywintypes import com_error
from PIL import Image

from fairypptx.presentation import Presentation
from fairypptx.core.resolvers import resolve_slide
from fairypptx import constants

from fairypptx._shape.box import Box
from fairypptx.registry_utils import yield_temporary_path
from fairypptx.object_utils import is_object, upstream


class Slide:
    def __init__(self, arg=None):
        self._api = resolve_slide(arg)

    @property
    def api(self):
        return self._api

    @property
    def shapes(self):
        from fairypptx.shapes import Shapes

        return Shapes(self.api.Shapes)

    @property
    def leaf_shapes(self):
        """Return Shapes, but grouped shape is decomposed.
        """
        from fairypptx.shape import Shape
        from fairypptx.shape_range import ShapeRange

        def _inner(shape):
            if shape.api.Type == constants.msoGroup:
                return sum((_inner(Shape(elem)) for elem in shape.api.GroupItems), [])
            else:
                return [shape]

        shape_list = sum((_inner(elem) for elem in self.shapes), [])
        return ShapeRange(shape_list)

    @property
    def presentation(self):
        return Presentation(upstream(self.api, "Presentation"))

    @property
    def size(self):
        """Return the size of the slide (Width, Height).
        """
        pres = self.presentation
        return (pres.api.PageSetup.SlideWidth, pres.api.PageSetup.SlideHeight)

    @property
    def box(self) -> Box:
        width, height = self.size
        return Box(left=0, top=0, width=width, height=height)

    @property
    def width(self):
        return self.size[0]

    @property
    def height(self):
        return self.size[1]

    @property
    def image(self):
        return self.to_image()

    def to_image(self, box=None, *, mode="RGBA"):
        """ To PIL.Image.

        Arg:
            box(Box, shape): Specify the range of cropping.
        """
        from fairypptx import Shape  # For dependency.
        if isinstance(box, Shape):
            box = box.box

        with yield_temporary_path(suffix=".png") as path:
            self.api.Export(path, "PNG")
            image = Image.open(path).convert(mode).copy()

        if box is not None:
            # Since the size differs, calibration is required.
            ratios = (image.size[0] / self.size[0],
                      image.size[1] / self.size[1])
            left, right = map(lambda x: round(x * ratios[0]), (box.left, box.right))
            top, bottom = map(lambda y: round(y * ratios[1]), (box.top, box.bottom))
            left, right = max(0, left), min(image.size[0], right)
            top, bottom = max(0, top), min(image.size[1], bottom)
            image = image.crop((left, top, right, bottom))
        return image

    def select(self) -> None:
        self.api.Select()

from fairypptx._slide.grid_handler import GridHandler
