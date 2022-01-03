"""Parts.

I wonder what is appropriate name for these concepts....
"""

from fairypptx import TextRange, Shape, Shapes
from fairypptx import Color
from fairypptx import Shape
from fairypptx.inner import storage
from fairypptx.constants import msoTrue, msoFalse
from fairypptx import constants

from fairypptx.object_utils import is_object, upstream
from fairypptx.inner.MSO import constants
from pywintypes import com_error

from fairyimage import from_latex


class Latex:
    """Latex Parts.

    Here,

    `script`: `Latex` snippet.
    `image`: The output image of
    Note
    -------
    (2022-01-03) : How should we do when multiple `Latex` images are grouped.
    """

    TEXT_COLOR = (117, 117, 117, 0)

    def __init__(self, arg=None):
        self.shape = self._to_root_shape(arg)  # The grouped Shape.
        assert not self.shape.is_leaf()

    @classmethod
    def make(cls, text, **kwargs):
        image = from_latex(text, **kwargs)
        image_shape = Shape.make(image)
        script_shape = Shape.make(text)
        script_shape.textrange.font.color = cls.TEXT_COLOR
        script_shape.textrange.font.size = 1
        script_shape.fill = None
        script_shape.api.Width = 0
        script_shape.api.Height = 0
        script_shape.api.Top = image_shape.top
        script_shape.api.Left = image_shape.left
        shape = Shapes([image_shape, script_shape]).group()
        return Latex(shape)

    def _to_root_shape(self, arg):
        if isinstance(arg, Shape):
            if arg.is_child():
                return arg.parent
            else:
                if not arg.is_leaf():
                    return arg
                else:
                    raise ValueError("Invalid Shape", arg, arg.text)
        elif isinstance(arg, Latex):
            return arg.shape

        elif arg is None:
            for shape in Shapes():
                if self.is_latex(shape):
                    return self._to_root_shape(shape)
            for shape in Slide():
                if self.is_latex(shape):
                    return self._to_root_shape(shape)
        raise ValueError("Not implemented.", arg.__class__)

    @property
    def script(self):
        # Not yet complete.
        # For example, multiple Latex are included,
        # then what should you do?

        def _is_text_shape(shape):
            try:
                shape.text
            except Exception as e:
                return False
            return True

        script = ""
        for child in self.shape.children:
            if _is_text_shape(child):
                script += child.text
        return script

    def compile(self, text, *args, **kwargs):
        # Memorandum...
        # I have to consider when multiple objects exist.
        script_shapes = [
            shape for shape in self.shape.children if self.is_script_shape(shape)
        ]
        assert len(script_shapes) == 1
        script_shape = script_shapes[0]
        image_shapes = [
            shape for shape in self.shape.children if self.is_image_shape(shape)
        ]
        assert len(image_shapes) == 1
        image_shape = image_shapes[0]

        image = from_latex(text, *args, **kwargs)
        path = storage.get_path(".png")
        image.save(path)
        left = image_shape.left
        top = image_shape.top
        width = image_shape.width
        height = image_shape.height

        script_shape.textrange.text = text

        # Perform change of `image`
        # and write the state of `self`
        shapes_api = upstream(self.shape.api, "Slide").Shapes
        output_image_shape = shapes_api.AddPicture(
            path, msoFalse, msoTrue, Left=left, Top=top, Width=width, Height=height
        )
        self.shape.ungroup()
        output_shape = Shapes([script_shape, output_image_shape]).group()
        # `self` should be rewritten.
        image_shape.api.Delete()
        self.shape = output_shape
        return self

    @classmethod
    def is_script_shape(cls, child: Shape):
        """This is a logic to judge whether `shape` represents `script`.

        Color's check and content...
        """
        try:
            child.text
        except Exception as com_error:
            return False
        return True

    @classmethod
    def is_image_shape(cls, child: Shape):
        """This is a logic to judge whether `shape` represents `image`."""
        return child.api.Type == constants.msoPicture

    @classmethod
    def is_latex(cls, shape: Shape):
        """Return whether `shape` is considered to be `Latex` or not."""
        if shape.is_child():
            shape = shape.parent
        if not shape.is_leaf():
            script_shapes = [
                shape for shape in shape.children if cls.is_script_shape(shape)
            ]
            image_shapes = [
                shape for shape in shape.children if cls.is_image_shape(shape)
            ]
            if len(script_shapes) == 1 and len(image_shapes) == 1:
                return True
        return False


if __name__ == "__main__":
    TEXT = r"""
    \begin{align*}
    \frac{x + y}{x - y} = 12
    \end{align*}
    """
    latex = Latex.make(TEXT)

    TEXT = r"""
    \begin{align*}
    \frac{x + z}{x - z} = 11
    \end{align*}
    """

    Latex.script
