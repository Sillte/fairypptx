"""Via `html` generate `Markdown`.  
"""

from pathlib import Path
from fairypptx import constants
from fairypptx import Shape, Shapes, Slide, Application, Table
from fairypptx.shape_range import ShapeRange


class Converter:
    def __init__(self, ):
        pass

    def parse(self, content, slide=None):

        from fairypptx.parts._markdown.html import pandoc, html_clipboard  # For dependency hierarchy. 
        from fairypptx.parts.markdown import toml_config
        if slide is None:
            slide = Slide()
        # Necessary to prevent deadlock.
        selection = Application().api.ActiveWindow.Selection
        if selection.Type == constants.ppSelectionText:
            selection.Unselect()

        content = self._to_content(content)
        # [TODO] Assume that content is markdown.
        content, config = toml_config.separate(content)
        css = config.get("css", None)
        css_folder = _get_default_css_folder()

        html = pandoc.to_html(content, css=css, css_folder=css_folder)

        # Path("./degub.html").write_text(html)
        html_clipboard.push(html, is_path=None)

        shapes = ShapeRange(slide.api.Shapes.Paste())
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
        from fairypptx import Markdown   # For hierarchy dependency. 
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



def from_html(content, slide=None) -> "Markdown":
    """Public API. 
    """
    converter = Converter()
    return converter.parse(content, slide=slide)


