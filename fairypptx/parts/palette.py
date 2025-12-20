from fairypptx.color import Color, make_hue_circle
from fairypptx.shape import Shape
from fairypptx.shapes import Shapes


class PaletteMaker:
    """Make a Color Pallete.
    """

    default_color = Color((25, 239, 198))
    
    def __init__(self,
                 fontsize=18,
                 line=3,
                 colors=None):
        self.fontsize = fontsize
        self.line = line
        self._prepared_colors = colors

    def __call__(self, contents=None):
        if contents is None:
            contents = self._gen_default_content()
        contents = self._to_dict(contents)
        shapes = []
        for key, color in contents.items():
            shape = Shape.make_textbox(key)
            shape.textrange.font.size = self.fontsize
            shape.tighten()
            shape.fill = color
            shape.line = self.line
            shapes.append(shape)
        return Shapes(shapes)

    def _to_dict(self, contents):
        def _to_color(arg):
            try:
                color = Color(arg)
            except Exception as e:
                raise ValueError(f"Cannot decipher `arg`, `{arg}`.") 
            return color
        def _to_pair(elem, index):
            arg = contents[index]
            if isinstance(arg, str):
                colors = self.prepare_colors(len(contents), override=False)
                return arg, colors[index % len(colors)]

            color = _to_color(arg)
            key = str(color.rgba)
            value = color
            return key, value

        from typing import Sequence, Mapping
        if isinstance(contents, Sequence):
            return dict(_to_pair(elem, index) for index, elem in enumerate(contents))
        elif isinstance(contents, Mapping):
            return {str(key) : _to_color(value) for key, value in contents.items()}
        raise ValueError()

    @property
    def prepared_colors(self):
        """Used for choosing the color
        outside the given `contents`.
        """
        if self._prepared_colors:
            return self._prepared_colors
        raise NotImplementedError("`_prepare_colors` must be called priorly.")

    def prepare_colors(self, n_color, override=False):
        if self._prepared_colors and override is False:
            return self._prepared_colors
        colors = make_hue_circle(self.default_color, n_color)
        self._prepared_colors = colors
        return colors

    def _gen_default_content(self):
        slide = Slide()
        colors = []
        for shape in slide.shapes:
            color = shape.fill.color
            if color:
                colors.append(color)
        return colors
