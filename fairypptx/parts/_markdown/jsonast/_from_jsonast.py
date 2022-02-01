"""

Policy:
-----------

`converter.config` will contain the required extensions for expansion. 


Memeorandum
-----------

When you want to `newline` in markdown,  
please consider the usage of `&nbsp;`.

Reference
------
# https://hackage.haskell.org/package/pandoc-types-1.22.1/docs/Text-Pandoc-Definition.html
"""


import subprocess
from pathlib import Path
from fairypptx import Shape, TextRange
from fairypptx import Table as PPTXTable 
from fairypptx import constants 
from fairypptx import text
from contextlib import contextmanager  
import json 


# [NOTE] 
# This is temporary solution.  
# I feel it is better to configure
# these fontsizes...

DEFAULT_FONTSIZE = 16


class Formatter:
    """Formatter is called once it requires the 
    `Format` should be changed.
    """
    def __call__(self, textrange):
        # Default settings are determined here..  
        #
        textrange.font.bold = False
        textrange.font.underline = False
        textrange.font.size = DEFAULT_FONTSIZE

        textrange.api.ParagraphFormat.Bullet.Visible = False

class HeaderFormatter:
    """Formatter for `Header`.  
    """
    def __init__(self, level, converter):
        # If you want to configure the behavior. 
        # you should use `converter.config`.
        self.level = level
        self.converter = converter

        self.level_to_fontsize = {1: 36, 2:24, 3:20, 4:18, 5:16}

    def __call__(self, textrange):
        fontsize = self.level_to_fontsize.get(self.level, DEFAULT_FONTSIZE)

        if self.level <= 3:
            textrange.font.underline = True

        textrange.font.bold = True
        textrange.font.size = fontsize

        textrange.api.ParagraphFormat.Bullet.Visible = False



class Converter:
    """Convert `JsonAst` to `fairypptx.Markdown`. 

    * Tag's interface.
    """
    elements = dict()

    def __init__(self, config=None):
        self._formatters = [Formatter()]
        self.markdown = None

        if config is None:
            config = {}
        self.config = config
        self._indent_level = 0

    # Register the `Element`. 
    @staticmethod
    def element(cls):
        assert hasattr(cls, "from_tag")
        name = cls.__name__
        Converter.elements[name] = cls

    def to_cls(self, tag):
        name = tag["t"]
        element = Converter.elements[name]
        return element

    @property
    def formatter(self):
        return self._formatters[-1]

    @property
    def indent_level(self):
        """Return level of indent.
        """
        if self._indent_level == 0:
            return 1
        return self._indent_level

    def insert(self, text):
        """Insert the `text` 
        with the corrent formatter.

        Micellaneous specifications I cannot understand, but can guess.  
        Inside this function, counter-act to these specification.   

        [TODO]: To acceessing of the last paragraph is not efficient yet.
        You should be take it consider later.  
        """

        # I do now why, but empty text seems illegal, which may differ from MSDN...? 
        # https://docs.microsoft.com/ja-jp/office/vba/api/powerpoint.textrange.insertafter
        if text == "":
            return None

        last_textrange = self.markdown.shape.textrange

        # I do now why, but when the `IndentLevel` is changed,  
        # The previous paragraphs of `Indent` also may change, unintentionally.
        
        if not last_textrange.text:
            last_indent_level = None
        else:
            last_indent_level = last_textrange.paragraphs[-1].api.IndentLevel
        last_n_paragraph = len(last_textrange.paragraphs)
        is_prev_paragraph = bool(last_textrange.paragraphs)

        textrange_api = last_textrange.api
        nt_api = textrange_api.InsertAfter(text)
        textrange = TextRange(nt_api)
        textrange.api.IndentLevel = self.indent_level

        self.formatter(textrange)

        # Here, the paragraph's is reset.  
        paragraphs = self.markdown.shape.textrange.paragraphs

        # If we have to revise the paragraph of the second to last.  
        # we perform these processings, here. 
        is_inc_paragraph = (last_n_paragraph < len(paragraphs))


        if is_inc_paragraph and is_prev_paragraph:
            assert 2 <= len(paragraphs)
            paragraphs[-2].api.IndentLevel = last_indent_level

        return textrange


    def set_tail_cr(self, n_cr=1):
        """This function assures the number of the 
        tail of `Text`'s `carriage return` . 
        """
        paragraphs = self.markdown.shape.textrange.paragraphs
        if not paragraphs:
            self.insert("\r" * n_cr)
            return 

        from itertools import takewhile
        text = paragraphs[-1].text
        n_tail_cr = len(list(takewhile(lambda t: t == "\r", reversed(text))))
        # I do not why, but `len(text)-` seems required.  
        stem = text[:len(text) - n_tail_cr]
        if n_tail_cr == n_cr:
            return 

        # If you set `paragraphs[-1].text`,  
        # then `format` may cnahge, 
        # We would like to prevent these situations as much as possible.
        # Hence, `insert` is used. 
        if n_tail_cr < n_cr:
            self.insert("\r" * (n_cr - n_tail_cr))
        else:
            # [TODO]: `Delete` is more appropriate? 
            # https://docs.microsoft.com/ja-jp/office/vba/api/powerpoint.textrange.delete
            paragraphs[-1].text = stem + "\r" * n_cr


    @contextmanager
    def formatter_scope(self, formatter):
        self._formatters.append(formatter)
        yield formatter
        self._formatters.pop()
    
    @contextmanager
    def inc_indent(self):
        """Increase `indent_level` one.
        """
        self._indent_level += 1
        yield 
        self._indent_level -= 1


    def parse(self, json_ast):
        if self._is_json(json_ast):
            json_ast = json.loads(json_ast)
        elif isinstance(json_ast, (str, Path)):
            json_ast = self._from_str_or_path(json_ast)

        assert isinstance(json_ast, dict)

        blocks = json_ast["blocks"]
        from pprint import pprint
        print("INPUT")
        pprint(blocks)
        shape = Shape.make(1)   # Temporary.  

        from fairypptx import Markdown  # For dependency hierarchy
        markdown = Markdown(shape)
        self.markdown = markdown   # Set `self.markdown`. 

        from pprint import pprint
        for block in blocks:
            pprint(block)
            cls = self.elements[block["t"]]
            cls.from_tag(block, markdown, self)
        markdown.shape.tighten()
        markdown.shape.textrange.paragraphformat.api.Alignment = constants.ppAlignLeft
        return markdown


    @classmethod
    def _from_str_or_path(self, content):
        if self._is_existent_path(content):
            content = Path(content).read_text("utf8")
        ret = subprocess.run("pandoc -t json",
                              universal_newlines=True, 
                              stdout=subprocess.PIPE, 
                              input=content, encoding="utf8")
        assert ret.returncode == 0
        return json.loads(ret.stdout)

    @classmethod
    def _is_json(self, json_ast): 
        try:
            json.loads(json_ast)
        except Exception as e:
            return False
        return True

    @classmethod
    def _is_existent_path(self, content):
        try:
            return Path(content).exists()
        except OSError:
            return False

class Element: 
    def __init_subclass__(cls, **kwargs):
        Converter.element(cls)

    @classmethod
    def from_tag(cls, tag, markdown, converter):
        raise NotImplementedError("")

    @classmethod
    def delegate_inlines(cls, inlines, markdown, converter):
        """ Delegate the `inlines`'s handling.n_tail_cr
        """
        for inline in inlines:
            cls = converter.to_cls(inline)
            element = cls.from_tag(inline, markdown, converter)


class Para(Element):
    @classmethod
    def from_tag(cls, tag, markdown, converter):
        inlines = tag["c"]
        cls.delegate_inlines(inlines, markdown, converter)
        converter.insert("\r")

class Plain(Element):
    @classmethod
    def from_tag(cls, tag, markdown, converter):
        inlines = tag["c"]
        cls.delegate_inlines(inlines, markdown, converter)
        

class Str(Element):
    @classmethod
    def from_tag(cls, tag, markdown, converter):
        string = tag["c"]
        converter.insert(string)

class Space(Element):
    @classmethod
    def from_tag(cls, tag, markdown, converter):
        converter.insert(" ")

class LineBreak(Element):
    @classmethod
    def from_tag(cls, tag, markdown, converter):
        converter.insert("\013")  # vertical tab.

class SoftBreak(Element):
    @classmethod
    def from_tag(cls, tag, markdown, converter):
        pass
        # converter.insert(" ")  

class Header(Element):
    @classmethod
    def from_tag(cls, tag, markdown, converter):
        level, attrs, inlines = tag["c"]
        formatter = HeaderFormatter(level, converter)

        with converter.formatter_scope(formatter):
            # Performs setting of `Format`.  
            cls.delegate_inlines(inlines, markdown, converter)
        converter.set_tail_cr(2)

class Strong(Element):
    @classmethod
    def from_tag(cls, tag, markdown, converter):
        inlines = tag["c"]

        def emphasize(textrange):
            textrange.font.bold = True
        
        with converter.formatter_scope(emphasize):
            # Performs setting of `Format`.  
            cls.delegate_inlines(inlines, markdown, converter)


class CodeBlock(Element):
    @classmethod
    def from_tag(cls, tag, markdown, converter):
        attrs, string = tag["c"]
        converter.insert(string)

class Code(Element):
    @classmethod
    def from_tag(cls, tag, markdown, converter):
        attrs, string = tag["c"]
        converter.insert(string)


def _change_bullet_type(paragraph, bullet_type):
    """
    # [BUG] / [UNSOLVED] I do not why, however, 
    # in some cases, the change of `ParagraphFormat.BulletType` 
    # does not applied when `IndentLevel` is the same as the previous ones...

    # Experimentally, I can guess that 
    # once you changes `IndentLevel` and set `Bullet.Type`, 
    # then, this problem does not seem occur. 
    """
    indent_level = paragraph.api.IndentLevel
    assert 1 <= indent_level <= 5,  "BUG."  
    paragraph.api.IndentLevel = 5
    paragraph.api.ParagraphFormat.Bullet.Type = bullet_type
    paragraph.api.IndentLevel = indent_level

class BulletList(Element):
    @classmethod
    def from_tag(cls, tag, markdown, converter):
        blocks = tag["c"]

        def bullet_list(textrange):
            textrange.api.ParagraphFormat.Bullet.Visible = True
            _change_bullet_type(textrange, constants.ppBulletUnnumbered)

            textrange.font.bold = False
            textrange.font.underline = False
            textrange.font.size = DEFAULT_FONTSIZE


        with converter.formatter_scope(bullet_list), converter.inc_indent():
            for block in blocks:
                for inlines in block:
                    cls = converter.to_cls(inlines)
                    cls.from_tag(inlines, markdown, converter)
                    converter.set_tail_cr(1)

                # For survey.
                #n_length = len(markdown.shape.textrange.text)
                #sub_api = markdown.shape.textrange.api.Characters(n_length - 1, 1)
                #print("sub_api", sub_api.Text)
                #TextRange(sub_api).font.bold = False
                #print(TextRange(sub_api).font.bold)


class OrderedList(Element):
    @classmethod
    def from_tag(cls, tag, markdown, converter):
        list_attributes, blocks = tag["c"]

        def bullet_list(textrange):
            textrange.api.ParagraphFormat.Bullet.Visible = constants.msoTrue
            textrange.api.ParagraphFormat.Bullet.Type = constants.ppBulletNumbered
            textrange.font.size = DEFAULT_FONTSIZE

        with converter.formatter_scope(bullet_list), converter.inc_indent():
            for block in blocks:
                for inlines in block:
                    cls = converter.to_cls(inlines)
                    cls.from_tag(inlines, markdown, converter)
                    converter.set_tail_cr(1)

class Link(Element):                
    @classmethod
    def from_tag(cls, tag, markdown, converter):
        #print("Link", tag); assert False
        attrs, inlines, targets = tag["c"]
        string = "".join([str(inline.get("c", "")) for inline in inlines])
        path = targets[0]
        textrange = converter.insert(string)
        hyperlink = textrange.ActionSettings(constants.ppMouseClick)
        hyperlink.Action = constants.ppActionHyperlink
        hyperlink.Hyperlink.Address = path

class HorizontalRule(Element):                
    @classmethod
    def from_tag(cls, tag, markdown, converter):
        pass

# Below elements are incomplete. 
#

class Table(Element):
    @classmethod
    def from_tag(cls, tag, markdown, converter):
        #print("Link", tag); assert False
        #inlines, alignment, number, columns, rows =  tag["c"]
        # Table Attr Caption [ColSpec] TableHead [TableBody] TableFoot
        attr, caption, colspec, table_head, table_body, table_foot = tag["c"]
        print("Currently `Table` cannot be handled")
        import numpy as np
        values = np.array([["" for _ in range(2)] for _ in range(2)])
        table = PPTXTable.make(values)

        # Currently, (2020/01/06) `markdown`'s cannot handles
        # the multiple shapes. 
        # So, I orphanage the generated Table.
        # markdown._shapes.append(table)


class RawInline(Element):
    """Handling Html...
    """
    @classmethod
    def from_tag(cls, tag, markdown, converter):
       format_, string  = tag["c"]


def from_jsonast(content, config=None):
    converter = Converter(config)
    return converter.parse(content)


if __name__ == "__main__":
    pass
    """
    target = Shape().textrange.paragraphs[-1]
    print(target.text)
    target.api.IndentLevel = 2
    s = target.api.InsertAfter("\r\nHOGEHOIGE")
    s.IndentLevel = 1
    print(target.api.Text, "check")
    print(Shape().textrange.paragraphs[-1].text)

    print(Shape().textrange.api.IndentLevel); exit(0)
    """
    #add(textrange, "\r")
    #exit(0)
    #print(Shape().textrange.text); exit(0)
    #TextRange().api.IndentLevel = 2; exit(0)
    #print(Shape().textrange.text); exit(0)
 
    sample = """
    {"blocks": [{"t": "Para", "c": [{"t": "Str", "c": "Three"}]}]}
    """
    SCRIPT = """ 
* ITEM1 
    1. ITEM1-1
    2. ITEM1-2
        * ITEM1-3
        * ITEM1-4
    * hgoe
    * fgerafva
    1. fgerafva
    3. fafe
    """.strip()

    conv = Converter()
    markdown =  conv.parse(SCRIPT)
    print(markdown.shape.text)

    exit(0)


