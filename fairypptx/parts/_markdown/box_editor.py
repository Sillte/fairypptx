""" Requires installment of ``Pandoc``.
Using json fileformat of Pandoc, Convert markdown to Textbox in Powerpoint.  

### Reference  
* [http://hackage.haskell.org/package/pandoc-types-1.17.5.1/docs/Text-Pandoc-Definition.html](http://hackage.haskell.org/package/pandoc-types-1.17.5.1/docs/Text-Pandoc-Definition.html)


Box Editor's Rule.
--------------------------

At the end of ``to_textbox`` function, virtual cursor exists at the start of the paragraph. 

"""
import os
import subprocess
from contextlib import contextmanager
import json  
from pprint import pprint  
from fairypptx import constants
from fairypptx.parts._markdown import fonttag_parser


def write(shape, text, default_fontsize=18):
    """ Write ``text`` into ``shape``.
    """
    parser = _construct_parser(text)
    box_editor = BoxEditor(shape, default_fontsize=default_fontsize)
    parser.to_textbox(box_editor)
    box_editor.arrange()

    return shape


def _to_jsonast(input_arg):
    if os.path.exists(input_arg):
        result = subprocess.run(f"pandoc {os.path.abspath(input_arg)} -t json", stdout=subprocess.PIPE)
        return json.loads(result.stdout.decode(), encoding="utf8")
    elif isinstance(input_arg, str):
        result = subprocess.run(f"pandoc -t json", stdout=subprocess.PIPE, input=input_arg.encode("utf8"))
        return json.loads(result.stdout.decode(), encoding="utf8")


def _factory(object_dict):
    try:
        class_object = globals()[object_dict["t"]]
    except KeyError as e:
        print("KeyError", e)
        return None
    else:
        if "c" in object_dict:
            instance = class_object(object_dict["c"])
        else:
            instance = class_object()
    return instance

def _list_class_factory(list_object_dict):
    instances = [instance for instance in map(_factory, list_object_dict) if instance]
    return instances


class Inline:
    pass

class Str(Inline):
    def __init__(self, string):
        self.string = string

    def to_textbox(self, box_editor):
        box_editor.add_text(self.string)
        pass

class Space(Inline):
    def to_textbox(self, box_editor):
        box_editor.add_text(" ")

class SoftBreak(Inline):
    def to_textbox(self, box_editor):
        box_editor.add_softbreak()

class LineBreak(Inline):
    def to_textbox(self, box_editor):
        box_editor.add_text("\013") # Vertical Tab.

class Strong(Inline):
    def __init__(self, inlines):
        self.inlines = _list_class_factory(inlines)

    def to_textbox(self, box_editor):
        box_editor.is_strong = True
        for inline in self.inlines:
            inline.to_textbox(box_editor)
        box_editor.is_strong = False

class Emph(Inline):
    def __init__(self, inlines):
        self.inlines = _list_class_factory(inlines)

    def to_textbox(self, box_editor):
        box_editor.is_italic = True
        for inline in self.inlines:
            inline.to_textbox(box_editor)
        box_editor.is_italic = False

class Span(Inline):
    def __init__(self, c):
        attr, inlines = c
        self.attr = attr
        self.attributes = self.attr[2]
        self.inlines = [instance for instance in map(_factory, inlines) if instance]

    def to_textbox(self, box_editor):
        #print("SPAN, attributes", self.attributes) 
        with box_editor.start_span(self.attributes):
            for inline in self.inlines:
                inline.to_textbox(box_editor)

class Link(Inline):
    def __init__(self, c):
        self.attr, inlines, self.target = c
        self.inlines = _list_class_factory(inlines)
        print("Link", self.attr, self.target, self.inlines)

    def to_textbox(self, box_editor):
        string = self.inlines[0].string
        path = self.target[0]
        box_editor.add_hyperlink(string, path)

        """
        # Currently, multiple attributes are not correctly handled.
        for inline in self.inlines:
            inline.to_textbox(box_editor)
        """

class RawInline(Inline):
    def __init__(self, c):
        self.format, self.string = c
        assert self.format == "html", "Currently, only html format is handled."

    def to_textbox(self, box_editor):
        box_editor.html_inline(self.string)


class Block:
    pass

class Header(Block):
    def __init__(self, c):
        self.level, self.attr, inlines = c
        self.inlines = [instance for instance in map(_factory, inlines) if instance]

    def to_textbox(self, box_editor):
        box_editor.set_format_specifier(f"h{self.level}")
        for inline in self.inlines:
            inline.to_textbox(box_editor)
        box_editor.set_format_specifier(None)
        box_editor.add_text("\r")


class Para(Block):
    def __init__(self, inlines):
        self.inlines = [instance for instance in map(_factory, inlines) if instance]

    def to_textbox(self, box_editor):
        for inline in self.inlines:
            inline.to_textbox(box_editor)
        box_editor.add_text("\r")

class CodeBlock(Block):
    def __init__(self, c):
        self.attr, self.string = c

    def to_textbox(self, box_editor):
        box_editor.add_text(self.string)


class Plain(Block):
    def __init__(self, inlines):
        self.inlines = [instance for instance in map(_factory, inlines) if instance]

    def to_textbox(self, box_editor):
        for inline in self.inlines:
            inline.to_textbox(box_editor)


class BulletList(Block):
    def __init__(self, blocks_list):
        self.parsers = [Parser(blocks) for blocks in blocks_list]

    def to_textbox(self, box_editor):
        with box_editor.start_unordered_list():
            for index, parser in enumerate(self.parsers):
                parser.to_textbox(box_editor)
                box_editor.add_text("\r")

class OrderedList(Block):
    def __init__(self, c):
        self.list_attributes, blocks_list = c
        self.parsers = [Parser(blocks) for blocks in blocks_list]

    def to_textbox(self, box_editor):
        with box_editor.start_ordered_list():
            for index, parser in enumerate(self.parsers):
                parser.to_textbox(box_editor)
                box_editor.add_text("\r")

class Parser:
    """ This is equivalent to blocks.
    """
    def __init__(self, blocks):
        self.blocks = _list_class_factory(blocks)

    def to_textbox(self, box_editor):
        for block in self.blocks:
            block.to_textbox(box_editor)

class BoxEditor:
    """ This class is responsible to create ``Textbox`` from ``Jsonast``.

    As for specifier,
    -------------------
    `h{\number}`, which specifies Header level.

    """
    def __init__(self, shape, default_fontsize=18):
        self.texts = list()
        self.shape = shape
        self.specifier = None
        
        self.indent_level = 1
        self._itemization_mode = None

        self.is_strong = False
        self.is_italic = False
        self.is_underline = False

        self.default_fontsize = default_fontsize
        self.default_fontcolor = (0, 0, 0)
        self.format_handler = HeaderFormatHandler(self.default_fontsize)

        # If an attribute is specified, attribute is pushed,
        # at the end of tag, stack is popped.
        # (color, size)
        self.fonttag_stack = list() 

    @property
    def fontcolor(self):
        if self.fonttag_stack:
            color, _ = self.fonttag_stack[-1]
            return color
        return self.default_fontcolor

    @property
    def fontsize(self):
        if self.fonttag_stack:
            _, font_size = self.fonttag_stack[-1]
            return font_size
        return self.default_fontsize


    def set_format_specifier(self, specifier):
        self.specifier = specifier

    def _set_format(self, textrange):
        """ Depending on the state of the **box_editor**,  
        """

        if self.specifier:
            self.format_handler.set_format(textrange, self.specifier, self.indent_level)
            if self.fonttag_stack: # If Attributes as for tag is specified, they applies. 
                set_property(textrange, fontsize=self.fontsize, color=self.fontcolor)
        else:
            set_property(textrange, fontsize=self.fontsize, color=self.fontcolor)


        """
        _h_pattern = re.compile("h(\d+)")
        if isinstance(self.specifier, int):
            text_functions.set_paragraph(textrange, self.specifier)

        if isinstance(self.specifier, str):
            ret = _h_pattern.findall(self.specifier)
            if ret:
                level = int(ret[0])
                text_functions.set_paragraph(textrange, level)

        """

        textrange.IndentLevel = (self.indent_level)


        #print("ITEMIZATION ", self._itemization_mode, self.indent_level)
        if self._itemization_mode == "unordered":
            textrange.ParagraphFormat.Bullet.Visible = True
            textrange.ParagraphFormat.Bullet.Type = constants.ppBulletUnnumbered
        elif self._itemization_mode == "ordered":
            textrange.ParagraphFormat.Bullet.Visible = True
            textrange.ParagraphFormat.Bullet.Type = constants.ppBulletNumbered
        else:
            textrange.ParagraphFormat.Bullet.Visible = False

        #print("is_strong", self.is_strong)
        if self.is_strong is not None:
            set_property(textrange, is_bold=self.is_strong)

        if self.is_italic is not None:
            set_property(textrange, is_italic=self.is_italic)

        if self.is_underline is not None:
            set_property(textrange, is_underline=self.is_underline)


    def is_last_empty_line(self):
        """ 2018-07-16
        This function is written as a trial to revise that the empty last paragprah is to be solved. 
        Currently, last paragraph remains empty. 
        Trial Strategy is to check if the last line is empty, then stop the addtion of "\r".
        However, this strategy does not work. This is because the after of "\r" is not regarded as one paragraph.
        As one of (untried strategy), cache of 
        """
        runs = list(self.shape.TextFrame.TextRange.Runs())
        if not runs:
            return True
        print("is_last_empty_line", runs[-1].Text.strip())
        return runs[-1].Text.strip() == ""

    @contextmanager
    def start_ordered_list(self):
        yield from self._start_itemization("ordered")

    @contextmanager
    def start_unordered_list(self):
        yield from self._start_itemization("unordered")

    
    def _start_itemization(self, mode_str="ordered"):
        prev_mode =  self._itemization_mode
        is_prev_itemization = (prev_mode != None)
        if is_prev_itemization: 
            self.add_text("\r")
            self.indent_level += 1 

        self._itemization_mode = mode_str
        yield 
        self._itemization_mode = prev_mode

        if is_prev_itemization: 
            self.indent_level -= 1 

    @contextmanager
    def start_span(self, attributes):
        color_arg, size_arg = fonttag_parser.parse_span_color_size(attributes)
        if color_arg or size_arg:
            color_arg = color_arg or self.fontcolor
            size_arg = size_arg or self.fontsize
            self.fonttag_stack.append((color_arg, size_arg))
            yield
            self.fonttag_stack.pop()
        else:  # If nothing is specified, no changed. 
            print("Failed to iterpret Attibute of SPAN", attributes)
            yield

    def add_hyperlink(self, text, path):
        textrange = self.shape.TextFrame.TextRange.InsertAfter(text)
        hyperlink = textrange.ActionSettings(constants.ppMouseClick)
        hyperlink.Action = constants.ppActionHyperlink
        hyperlink.Hyperlink.Address = path

    def add_softbreak(self):
        """ This is different from Normal Markdown Rule, 
        Usually, SoftBreak is not treated as ``LF``.  
        """
        if self._itemization_mode is None:
            self.add_text("\n")


    def add_text(self, text):
        textrange = self.shape.TextFrame.TextRange.InsertAfter(text)
        self._set_format(textrange)
        self.texts.append(text) 

    def html_inline(self, text):
        if text.lower() == "<u>":
            self.is_underline = True
        if text.lower() == "</u>":
            self.is_underline = False
 
        if text.lower() == "<br>" or text.lower() == "<br />":
            self.add_text("\n")

        if text.replace(" ", "").lower().startswith("<font"):
            color_arg, size_arg = fonttag_parser.parse_color_size(text)
            color_arg = color_arg or self.fontcolor
            size_arg = size_arg or self.fontsize
            self.fonttag_stack.append((color_arg, size_arg))

        if text.replace(" ", "").lower().startswith("</font"):
            self.fonttag_stack.pop()


    def arrange(self):
        """ Format the Textbox.

        Problem.
        When the last paragraph is emtpy, eliminate the final (empty) Paragraph.
        """
        # Empty itemization paragraph is deleted.
        from fairypptx import TextRange
        paragraphs = TextRange(self.shape).paragraphs
        #paragraphs = text_functions.get_paragraphs(self.shape)
        for index, paragraph in enumerate(reversed(paragraphs)):
            if paragraph.ParagraphFormat.Bullet.Visible and len(paragraph.Text) == 1:
                paragraph.Delete()

        if self.shape.TextFrame.TextRange.Text[-1] == "\r":
            """ Survey is required.
            If the last Paragraph ends with "\r",
            then Only "\r" is to be removed.
            (However, currently this is not achieved.)
            """
            lines = list(self.shape.TextFrame.TextRange.Lines())
            lines[-1].Text = lines[-1].Text[:-1]
            


def _construct_parser(text):
    json_ast = _to_jsonast(text)
    blocks = json_ast["blocks"]
    # Necessary for debug purpose.
    pprint(blocks)
    parser = Parser(blocks)
    return parser

class _Formatter:
    def __init__(self,fontsize=None,
                      is_bold=None,
                      is_italic=None,
                      is_underline=None):
        self.fontsize = fontsize
        self.is_bold = is_bold
        self.is_italic = is_italic
        self.is_underline = is_underline
        pass

    def set_format(self, textrange):
        set_property(textrange,
                    fontsize=self.fontsize,
                    is_bold=self.is_bold,
                    is_italic=self.is_italic,
                    is_underline=self.is_underline)

class HeaderFormatHandler:
    """ If the intended specifier exists, then 
    Format is set. 
    This Format Handler ignores **indent_level**.
    """
    def __init__(self, default_fontsize):
        general_ratio = {"h1":2.0, "h2":1.5, "h3":1.2, "h4":1.0}
        hlevel_to_fontsize = {key: value * default_fontsize for key, value in general_ratio.items()}
        self.hlevel_to_formatter = {key: _Formatter(fontsize=value) for key, value in hlevel_to_fontsize.items()}


    def set_format(self, textrange, specifier, indent_level):
        print("HEADER SPECIFIER", specifier)
        if specifier in self.hlevel_to_formatter:
            self.hlevel_to_formatter[specifier].set_format(textrange)


def set_property(textrange,
                 fontsize=None,
                 color=None,
                 is_bold=None, 
                 is_italic=None,
                 is_underline=None):
    from fairypptx import Color
    # textrange = _assure_textrange(textrange)
    def _boolean_converter(arg):
        if arg in {constants.msoTrue, constants.msoFalse , constants.msoTriStateMixed}:
            return arg
        if arg is True:
            return constants.msoTrue
        elif arg is False:
            return constants.msoFalse
        raise ValueError("Invalid argment", arg)

    if fontsize is not None:
        textrange.Font.Size = fontsize
    if color is not None:
        textrange.Font.Color.RGB = Color(color).as_int()
    if is_bold is not None:
        textrange.Font.Bold = _boolean_converter(is_bold)
    if is_italic is not None:
        textrange.Font.Italic = _boolean_converter(is_italic)
    if is_underline is not None:
        textrange.Font.Underline = _boolean_converter(is_underline)
    return textrange



text = """
<span style="color: #FF00FF; font-size:10pt">This is a sample of..</span>
It this really OK?
"""

def _survey(shape):
    from fairypptx import TextRange
    paragraphs = TextRange(shape).paragraphs
    for index, paragraph in enumerate(paragraphs):
        print("p_index", index, paragraph.IndentLevel)
        for r_index, run in enumerate(paragraph.Runs()):
            print("r_index", r_index, run.Text.encode())


if __name__ == "__main__":
    pass
