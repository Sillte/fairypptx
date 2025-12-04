from pywintypes import com_error
from PIL import Image

from fairypptx.presentation import Presentation
from fairypptx.core.application import Application
from fairypptx import constants

from fairypptx.color import Color
from fairypptx.box import Box
from fairypptx.registry_utils import yield_temporary_path
from fairypptx.object_utils import get_type, is_object, upstream


class Slides:
    """Slides.
    However, this may not behave like all of the slides 
    is handled this class.   
    It accepts a subset of Slides Object. 

    Note
    ---------------------
    * `Add` / `Delete` operations may break this class.

    """

    def __init__(self, arg=None):
        self.app = Application()

        self._api, self._indices = self._construct(arg)

    @property
    def api(self):
        return self._api

    def add(self, index=None, layout=None):
        # Add and AddSlide exist.
        # You may consider the branch processing.

        if index is None:
            index = len(self)
        assert 0 <= index <= len(self)
        if layout is None:
            # Mechanisum of customization is preferred.
            layout = constants.ppLayoutBlank
        return Slide(self.api.Add(index + 1, layout))

    def __len__(self):
        return len(self._indices)

    def __getitem__(self, key):
        if isinstance(key, int):
            index = self._indices[key]
            return Slide(self.api.Item(index + 1))
        elif isinstance(key, slice):
            indices = self._indices[key]
            slides = [Slide(self.api.Item(index + 1)) for index in indices]
            return Slides(slides)

    def __iter__(self):
        for i, index in enumerate(self._indices):
            yield self[i]

    def __len__(self):
        return self.api.Count

    def _construct(self, arg):
        """
        [TODO] When `arg` is None, what kind of specification is desirable?
        """
        if is_object(arg, "SlideRange"):
            slide_objects = [arg.Item(index + 1) for index in range(arg.Count)]
            slides_objects = [slide.Parent.Slides for slide in slide_objects]
            assert len(set(map(id, slides_objects))) == 1, "Slide must be"
            slides_object = slides_objects[0]
            indices = [elem.SlideIndex - 1 for elem in slide_objects]
            return slides_object, indices
        elif is_object(arg, "Slides"):
            return arg, range(arg.Count)
        elif isinstance(arg, Slides):
            return arg.api, arg._indices
        if arg is None:
            App = self.app.api
            try:
                if App.ActiveWindow.Selection.SlideRange:
                    return self._construct(App.ActiveWindow.Selection.SlideRange)
            except com_error:
                pass
            slides = Presentation().slides
            return slides.api, slides._indices
        raise ValueError(f"Cannot interpret `arg`; {arg}.")


class Slide:
    def __init__(self, arg=None):
        self.app = Application()
        self._api = self._fetch_api(arg)

    @property
    def api(self):
        return self._api

    @property
    def shapes(self):
        from fairypptx.shape import Shapes

        return Shapes(self.api.Shapes)

    @property
    def leaf_shapes(self):
        """Return Shapes, but grouped shape is decomposed.
        """
        from fairypptx.shape import Shapes, Shape

        def _inner(shape):
            if shape.api.Type == constants.msoGroup:
                return sum((_inner(Shape(elem)) for elem in shape.api.GroupItems), [])
            else:
                return [shape]

        shape_list = sum((_inner(elem) for elem in self.shapes), [])
        return Shapes(shape_list)

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
    def box(self):
        width, height = self.size
        d = {"Left": 0, "Top": 0, "Width": width, "Height": height}
        return Box(d)

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

    def select(self):
        self.api.Select()


    @property
    def colors(self):
        """Return list of `fairy. 
        """
        colors = set()
        for shape in self.shapes:
            color = shape.fill.color
            if color is not None:
                colors.add(shape.fill.color)
        return list(colors)

    def _fetch_api(self, arg):
        if is_object(arg, "Slide"):
            return arg
        elif isinstance(arg, Slide):
            return arg.api
        pres = Presentation()

        App = self.app.api
        if arg is None:
            try:
                if App.ActiveWindow.ViewType != constants.ppViewNormal:
                    App.ActiveWindow.ViewType = constants.ppViewNormal
            except com_error:
                pass
            try:
                if App.ActiveWindow.Selection.SlideRange:
                    return App.ActiveWindow.Selection.SlideRange.Item(1)
            except com_error:
                pass

            pres = Presentation()
            if pres.slides:
                return pres.slides[-1].api
            else:
                return pres.api.Slides.Add(1, constants.ppLayoutBlank)
            raise ValueError(f"Cannot find an appropriate `slide`.")

        raise ValueError(f"Cannot interpret `arg`; {arg}.")

# To solve the order of `import`.
from fairypptx._slide import GridHandler
