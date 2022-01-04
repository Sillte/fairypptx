"""Markdown

Markdown. 

(2020-04-26)
As you can easily notice, this code is yet a collection 
of Sillte-kun's practice codes,  


(2022/01/04)
Though it is a draft, the concept `Parts` are introduced. 
`Markdown` should be regarded as one of them.  
Therefore, the position of file is changed. 

Currently, this contrains a lot of problems. 

* Copy of `html` to `TextRange` does not work correctly.  
* How to use tags in `Markdown`? 
"""
from pathlib import Path
from typing import Sequence
from fairypptx import Slide, Shapes, Shape, TextRange, Application, Text, Table
from fairypptx import constants

from fairypptx._text.textrange_stylist import ParagraphTextRangeStylist
from fairypptx.parts._markdown import write, interpret
from fairypptx.parts._markdown import toml_config
from fairypptx.parts._markdown import pandoc, html_clipboard


class Markdown:
    """

    Note (murmurs)
    ------

    `Markdown` may contains `Multiple` Shapes...
    """
    def __init__(self, arg=None, **kwargs):
        self.shapes = self._to_shapes(arg)

    def _to_shapes(self, arg):
        if isinstance(arg, Shapes):
            return arg
        elif isinstance(arg, Shape):
            return Shapes([arg])
        elif isinstance(arg, Markdown):
            return arg.shapes
        elif isinstance(arg, Sequence):
            return Shapes(arg)
        elif arg is None:
            return self._to_shapes(Shape())
        raise TypeError("Invalid arg", arg)

    @property
    def shape(self):
        return self.shapes[0]

    @classmethod
    def make(cls, arg, slide=None):
        if slide is None:
            slide = Slide()
        # Necessary to prevent deadlock.
        selection = Application().api.ActiveWindow.Selection
        if selection.Type == constants.ppSelectionText:
            selection.Unselect()

        content = cls._to_content(arg)
        # [TODO] Assume that content is markdown.
        content, config = toml_config.separate(content)
        css = config.get("css", None)
        css_folder = _get_default_css_folder()

        html = pandoc.to_html(content, css=css, css_folder=css_folder)

        # Path("./degub.html").write_text(html)
        html_clipboard.push(html, is_path=None)

        shapes = Shapes(slide.api.Shapes.Paste())
        for shape in shapes:
            if shape.api.Type == constants.msoTable:
                Table(shape).tighten()
            elif hasattr(shape, "textrange"):
                _compensate_textrange(shape)
                shape.tighten()

        # Adjustment geometrically
        # [TODO] It this strategy is all right?  
        if 1 < len(shapes):
            shapes = sorted(shapes, key=lambda shape: (shape.box.top, shape.box.left))
            c_x = shapes[0].api.Left  
            c_y = shapes[0].api.Top
            for shape in shapes:
                shape.api.Left = c_x
                shape.api.Top = c_y
                c_x += shape.api.Width
        return Markdown(shapes)

    @classmethod
    def _to_content(cls, arg):
        try:
            path = Path(arg)
            if path.exists():
                return path.read_text(encoding="utf8")
        except OSError:
            pass
        return arg

    # Since `Markdown` belong to `Part`,   
    # I have to prepare these interfaces.

    @property
    def script(self):
        """
        Note: I know, this is far from complete.
        """
        return self.shape.text


    def compile(self, text, *args, **kwargs):
        # Currently, generate the next `Markdown` and 
        # Change the position and delete the old one. 

        new_markdown = type(self).make(text, *args, **kwargs)

        left = self.shapes[0].left
        top = self.shapes[0].top

        for n_shape in new_markdown.shapes:
            n_shape.left = left
            n_shape.top = top

        for shape in self.shapes:
            shape.api.Delete()
        self = new_markdown
        return  self



def _get_default_css_folder():
    folder = Path().home() / ".fairypptx" / "css"
    if folder.exists():
        return folder
    folder.mkdir()
    # I'd like to put one typical example. 
    sample_css = folder / "sample.css"
    if not sample_css.exists():
        sample_css.write_text(_css_sample, encoding="utf8")
    return folder

def _compensate_textrange(shape):
    """Some properties cannot handle `text_range` appropriately.
    Here, minimum compensation is performed.
    """
    # This value is derived experimentarly.
    # I am not sure this is a good strategy...
    # Ref: https://www.relief.jp/docs/powerpoint-vba-setting-indent.html
    tr = shape.textrange
    for para in shape.textrange.paragraphs:
        para.api2.ParagraphFormat.FirstLineIndent = -22.5


_css_sample = """
body {
    font-family: Meiryo;
    font-size: 18px
}

h1 {
    font-size: 32px;
    font-weight:bold;
}

h2 {
    font-size: 28px;
    font-weight:bold;
}

h3 {
    font-size: 24px;
    font-weight:bold;
    text-decoration: underline; 
}

h4 {
    font-size: 18px;
    text-decoration: underline; 
}

table, th, td {
  border-collapse: collapse;
  border: 3px solid #ccc;
  line-height: 3;
}
""".strip()
            
