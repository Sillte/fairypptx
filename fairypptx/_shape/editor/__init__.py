"""Editor related to `Shapes`.

As you can easily assumes, `editor` is a high-level api, so
* This sub-module can call other more premitive api freely.  
* On contrary, the more premitive sub-modules should not call this.  
"""


from fairypptx import constants
from fairypptx.shape import Shape, Shapes
from fairypptx.shape import Box


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
        c_box = self.circumscribed_box(shapes)
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
        return Shapes(list(shapes) + [shape])

    @classmethod
    def circumscribed_box(cls, shapes):
        boxes = [shape.box for shape in shapes]
        c_left = min(box.left for box in boxes)
        c_top = min(box.top for box in boxes)
        c_right = max(box.right for box in boxes)
        c_bottom = max(box.bottom for box in boxes)
        c_box = Box(c_left, c_top, c_right - c_left, c_bottom - c_top)
        return c_box


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
        c_box = ShapesEncloser.circumscribed_box(shapes)
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


if __name__ == "__main__":
    pass

