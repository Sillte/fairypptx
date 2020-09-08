"""
Note
----
This module is assumed to freely import all of the Object Model.
On contrary, Object Model classes must import this class in Runtime.

(2020-04-26)
As you can easily notice, this code is yet a collection 
of Sillte-kun's practice codes,  

"""
from pathlib import Path
from fairypptx import Slide, Shapes, Shape, TextRange, Application, Text, Table
from fairypptx import constants

from fairypptx._text.textrange_stylist import ParagraphTextRangeStylist
from fairypptx._markdown import write, interpret
from fairypptx._markdown import toml_config
from fairypptx._markdown import pandoc, html_clipboard


class Markdown:
    """Handle Markdown.

    Wonder
    ----------------------------
    Derive UserString and make it behave like a string.
    """

    def __init__(self, arg, **kwargs):
        if isinstance(arg, (str, Path)):
            self.textrange = self.make(arg).textrange
        elif isinstance(arg, Shape):
            self.textrange = arg.textrange
        elif isinstance(arg, TextRange):
            self.textrange = arg
        elif isinstance(arg, Markdown):
            self.textrange = arg.textrange
        elif isinstance(arg, Text):
            raise NotImplementedError("Currently...")
        else:
            raise TypeError(f"Invalid Argument: `{arg}`", type(arg))

    def __str__(self):
        return interpret(self.textrange.shape)

    @property
    def shape(self):
        if self.textrange:
            return self.textrange.shape
        else:
            # This path is related to derivation of UserString of this class.
            # See the Wonder section of the class.
            raise ValueError("This markdown does not belong to `TextRange/Shape`")

    @classmethod
    def make(cls, arg, slide=None):
        """
        Ideally, the return shape is `Markdown`.
        However, when `Table` is generated,
        multiple shapes are required,
        so this function returns `Markdown` , if possible.
        Otherwise, return `Shapes`.
        """
        if slide is None:
            slide = Slide()

        selection = Application().api.ActiveWindow.Selection
        # Necessary to prevent deadlock.
        if selection.Type == constants.ppSelectionText:
            selection.UnSelect()

        content = _to_content(arg)
        # [TODO] Assume that content is markdown.
        content, config = toml_config.separate(content)
        css = config.get("css", None)
        css_folder = _get_default_css_folder()

        # One strategy for making markdown.
        # However, it seems good to `copy mechanism of html`.
        # shape = Shape()
        # shape = Shape(write(shape, arg))

        # Change of the alignment.
        # shape.textrange.api.ParagraphFormat.Alignment = constants.ppAlignLeft
        # shape.textrange.paragraphformat = {"Alignment": constants.ppAlignLeft}
        # shape.tighten()
        # return Markdown(shape.textrange)

        # This strategy uses HTML copies.
        html = pandoc.to_html(content, css=css, css_folder=css_folder)
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
            c_x = shapes[0].left  
            c_y = shapes[0].top
            for shape in shapes:
                shape.api.Left = c_x
                shape.api.Top = c_y
                c_x += shape.api.Width

        if len(shapes) == 1:
            try:
                return Markdown(shape.textrange)
            except:
                pass
        return Shapes(shapes)


def _to_content(arg):
    try:
        path = Path(arg)
        if path.exists():
            return path.read_text(encoding="utf8")
    except OSError:
        pass
    return arg

def _compensate_textrange(shape):
    """Some properties cannot handle `text_range` appropriately.
    Here, minimum compensation is performed.
    """
    # This value is derived experimentarly.
    # I am not sure this is a good strategy...
    # Ref: https://www.relief.jp/docs/powerpoint-vba-setting-indent.html
    tr = shape.textrange
    for para in shape.textrange.paragraphs:
        para.api.Parent.Parent.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent = (
            -22.5
        )

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

if __name__ == "__main__":
    pass
