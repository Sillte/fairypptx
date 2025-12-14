"""Editor related to `Shapes`.

As you can easily assumes, `editor` is a high-level api, so
* This sub-module can call other more premitive api freely.  
* On contrary, the more premitive sub-modules should not call this.  
"""

import numpy as np 
import _ctypes
from pywintypes import com_error
from fairypptx import constants
from fairypptx.shape import Shape, TableShape
from fairypptx.shape import Box
from fairypptx.table import Table

from fairypptx.object_utils import is_object
from typing import Sequence


def _to_shapes(arg):
    """Convert to `Shapes`."""
    from fairypptx import Shapes

    if isinstance(arg, Shapes):
        return arg
    elif isinstance(arg, Sequence):
        return Shapes(arg)
    elif isinstance(arg, Shape):
        return Shapes([arg])
    elif is_object(arg, "Shapes"):
        return Shapes(arg)
    elif is_object(arg, "Shape"):
        return Shapes(arg)
    raise ValueError(f"Cannot interpret `{arg}`.")


class ShapesEncloser:
    """Enclose the specified shapes.
    """
    def __init__(self,
                 line=3,
                 fill=None,
                 linecolor=(0, 0, 0),
                 *,
                 margin=0.10,
                 left_margin=None,
                 top_margin=None,
                 right_margin=None,
                 bottom_margin=None,
                 ):
        self.line = line
        self.fill = fill
        self.linecolor = linecolor

        self.margin = margin
        self.left_margin = left_margin
        self.top_margin = top_margin
        self.right_margin = right_margin
        self.bottom_margin = bottom_margin

    def _margin_solver(self, c_box):
        """Solves the margin of 
        it returns the actual pixel(float) of margin. (i.e. not ratio)
        (left_margin, top_margin, right_margin, bottom_margin).
        """
        def _to_pixel(margin, length):
            if isinstance(margin, float) and abs(margin) < 1.0:
                return length * margin
            else:
                return margin

        def _solve_margin(first_value, length):
            value = first_value
            if value is None:
                value = self.margin
            assert value is not None
            return _to_pixel(value, length)

        left_margin = _solve_margin(self.left_margin, c_box.x_length)
        top_margin = _solve_margin(self.top_margin, c_box.y_length)
        right_margin = _solve_margin(self.right_margin, c_box.x_length)
        bottom_margin = _solve_margin(self.bottom_margin, c_box.y_length)
        return (left_margin, top_margin, right_margin, bottom_margin)
 
    def __call__(self, shapes):
        if not shapes:
            return None
        shapes = _to_shapes(shapes)
        c_box = shapes.circumscribed_box
        left_margin, top_margin, right_margin, bottom_margin = self._margin_solver(c_box)

        width = c_box.width + (left_margin + right_margin)
        height = c_box.height + (top_margin + bottom_margin)
        shape = Shape.make(1)
        shape.api.Top = c_box.top -  top_margin
        shape.api.Left = c_box.left - left_margin
        shape.api.Width = width
        shape.api.Height = height
        shape.line = self.line
        shape.fill = self.fill
        if self.linecolor:
            shape.line = self.linecolor 
        shape.api.Zorder(constants.msoSendToBack)
        from fairypptx import ShapeRange
        return ShapeRange(list(shapes) + [shape])


class TitleProvider:
    def __init__(self,
                 title,
                 fontsize=None,
                 fontcolor=(0, 0, 0),
                 fill=None,
                 line=None,
                 bold=True,
                 underline=False,
                 ):
        self.title = title
        self.fontsize = fontsize
        self.fontcolor = fontcolor
        self.fill = fill
        self.line = line
        self.bold = bold
        self.underline = underline

    def __call__(self, shapes):
        shapes = _to_shapes(shapes)
        c_box = shapes.circumscribed_box
        shape = Shape.make(1)
        shape.fill = self.fill
        shape.line = self.line
        shape.textrange.text = self.title
        shape.textrange.font.bold = self.bold
        shape.textrange.font.underline = self.underline
        shape.textrange.font.size = self._yield_fontsize(self.fontsize, shapes)
        shape.textrange.font.color = self.fontcolor
        shape.tighten()
        shape.api.Top = c_box.top - shape.height
        shape.api.Left = c_box.left 
        return shape

    def _yield_fontsize(self, fontsize, shapes):
        if fontsize is not None:
            return fontsize
        fontsizes =[]
        for shape in shapes:
            fontsizes.append(shape.textrange.font.size)
        if fontsizes:
            return max(fontsizes)
        else:
            return 18

class ShapesResizer:
    """Shapes Resizer.

    This class resize the given shapes to the equivalent size.


    Related Class.
    -----------
    `shapes.BoundingResizer`: the bounding box of the shapes is resized.

    """

    def __init__(self, size="max"):
        self.size = size

    def _yield_size(self, shapes):
        """Determine the return of size,
        based on the given parameters.
        """
        size = self.size

        if isinstance(size, (list, tuple)):
            width, height = size
        elif isinstance(size, Shape):
            shape = size
            width, height = shape.width, shape.height
        elif isinstance(size, str):
            if size == "max":
                width = max(shape.width for shape in shapes)
                height = max(shape.height for shape in shapes)
            else:
                raise NotImplementedError("This error message must be displayed in `__init``. ") 
        return width, height

    def __call__(self, shapes):
        width, height = self._yield_size(shapes)
        for shape in shapes:
            shape.width = width
            shape.height = height
        return shapes


class BoundingResizer:
    """Resize the bounding box of `Shapes`.

    Args:
        size: 2-tuple. (width, height).
            The expected width and height. 
        fontsize: (float)
            The fontsize of the expected minimum over the shapes.
    """

    def __init__(self, size=None, *, fontsize=None):
        self.size = size
        self.fontsize = fontsize

    def _to_minimum_fontsize(self, textrange):
        fontsizes = set()
        for run in textrange.runs:
            if run.text:
                fontsizes.add(run.font.size)
        if fontsizes:
            return min(fontsizes)
        else:
            return None

    def _get_minimum_fontsize(self, shapes):
        fontsizes = set()
        for shape in shapes:
            if shape.is_table():
                table = Table(shape)
                for row in table.rows:
                    for cell in row:
                        textrange = cell.shape.textrange
                        fontsize = self._to_minimum_fontsize(textrange)
                        if fontsize:
                            fontsizes.add(fontsize)
            else:
                try:
                    fontsize = self._to_minimum_fontsize(shape.textrange)
                except com_error as e:
                    pass
                else:
                    if fontsize:
                        fontsizes.add(fontsize)
        if fontsizes:
            return min(fontsizes)
        else:
            return None

    def _set_fontsize(self, textrange, ratio):
        for run in textrange.runs:
            run.api.Font.Size = round(run.font.size * ratio)

    def _yield_size(self, shapes):
        """Determine the the return of `size`.

        * Priority
        1. `fontsize`
        2. `size`.
        """
        size = self.size
        fontsize = self.fontsize
        
        # For fallback.
        if size is None and fontsize is None:
            fontsize = self._get_minimum_fontsize(shapes.slide.shapes)
            if fontsize is None:
                fontsize = 12

        if fontsize is not None:
            c_box = shapes.circumscribed_box
            c_fontsize = self._get_minimum_fontsize(shapes)
            ratio = fontsize / c_fontsize
            size = ratio 


        if isinstance(size, (int , float)):
            c_box = shapes.circumscribed_box
            c_width = c_box.x_length
            c_height = c_box.y_length
            n_width = c_width * size
            n_height = c_height * size
        elif isinstance(size, (list, tuple)):
            n_width, n_height = size
        else:
            raise ValueError("Invalid size.", size)

        return n_width, n_height


    def __call__(self, shapes):
        """Perform `resize` for all the shapes.  

        Not only it changes the size of `Shape`, 
        but also changes the size of `Font` proportionaly. 

        Note:
        It works only for shapes whose rotation is 0.
        """

        # If the given is `Shape`, then, `Shape` is returned.
        if isinstance(shapes, Shape):
            is_shape = True
        else:
            is_shape = False
        if not shapes:
            return shapes
        shapes = _to_shapes(shapes)


        n_width, n_height = self._yield_size(shapes)
        c_box = shapes.circumscribed_box
        width, height = c_box.x_length, c_box.y_length

        pivot = (c_box.top, c_box.left)  # [y_min, x_min]
        ratios = (n_height / height, n_width / width)
        ratio = np.mean(ratios)
        for shape in shapes:
            # Processings for all the shapes.
            shape.api.Left = (shape.api.Left - pivot[1]) * ratios[1]  + pivot[1]
            shape.api.Width = shape.api.Width * ratios[1]
            shape.api.Top = (shape.api.Top - pivot[0]) * ratios[0]  + pivot[0]
            shape.api.Height = shape.api.Height * ratios[0]

            # For Table.
            if isinstance(shape, TableShape):
                table = Table(shape)
                for row in table.rows:
                    for cell in row:
                        self._set_fontsize(cell.shape.textrange, ratio)
            else:
                try:
                    self._set_fontsize(shape.textrange, ratio)
                except com_error as e:
                    pass

        if not is_shape:
            return Shapes(shapes)
        else:
            return shapes[0]


if __name__ == "__main__":
    pass

