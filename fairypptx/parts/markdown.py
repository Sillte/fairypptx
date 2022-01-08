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

(2022/01/06) -> `jsonast` mechanism can help the problem above?  
"""
from pathlib import Path
from typing import Sequence
from fairypptx import Slide, Shapes, Shape, TextRange, Application, Text, Table
from fairypptx import constants

from fairypptx._text.textrange_stylist import ParagraphTextRangeStylist
from fairypptx.parts._markdown import toml_config
from fairypptx.parts._markdown.html import from_html
from fairypptx.parts._markdown.jsonast import from_jsonast


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
    def make(cls, arg, slide=None, engine="jsonast"):
        if slide is None:
            slide = Slide()
        engine = str(engine).lower()

        # Necessary to prevent deadlock.
        selection = Application().api.ActiveWindow.Selection
        if selection.Type == constants.ppSelectionText:
            selection.Unselect()

        # [TODO] As of 2022-01-06
        # generation interface is not compatible.
        # You should consider this.

        if engine == "html":
            return from_html(arg, slide=slide)
        elif engine == "jsonast":
            return from_jsonast(arg)
        else:
            raise NotImplementedError("Engine is not implemented.") 


    @property
    def script(self):
        """
        Note: I know, this is far from complete.
        """
        from fairypptx.parts._markdown.jsonast import to_script
        return to_script(self.shape.textrange)


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



            
