""" Convert **TextBox** to markdown format.
"""

import os
import subprocess
from contextlib import contextmanager
from collections import defaultdict
import re
import json  
from pprint import pprint  
from fairypptx import constants
from fairypptx.color import Color

def interpret(shape):
    """ Interpret shape's textrange as markdown.
    """
    result_buffer = ""
    paragraphs = list(shape.api.TextFrame.TextRange.Paragraphs())
    converter = _Converter(paragraphs)
    for paragraph in paragraphs:
        result_buffer += converter.start_paragraph(paragraph)
        runs = paragraph.Runs()
        for run in runs:
            result_buffer += converter.convert_run(run)
        result_buffer += converter.end_paragraph(paragraph)

    return result_buffer


def _infer_default_fontsize(paragraphs):
    fontsize_list = list()
    for paragraph in paragraphs:
        size = paragraph.Font.Size
        if 0 < size:
            fontsize_list.append(size)

    # To compensate the small font exists, modification is required.
    result_fontsize = min(fontsize_list, default=18)
    #print("Inferred FontSize", result_fontsize)
    return result_fontsize

def _infer_default_fontcolor(paragraphs):
    color_counter = defaultdict(int)

    for paragraph in paragraphs:
        for run in paragraph.Runs(): 
            rgb_int = run.Font.Color.RGB
            print("rgb_int", rgb_int)
            if 0 <= rgb_int:
                color_counter[rgb_int] += 1
    if not color_counter:
        return Color((0, 0, 0))
    result_fontcolor = max(color_counter.keys(), key=lambda k: color_counter[k])
    result_fontcolor = Color(result_fontcolor)
    print("Inferred Font Color", result_fontcolor); 
    return result_fontcolor

class _Converter:
    def __init__(self, paragraphs):
        self._itemization_mode = None
        self._indent_level = 1
        self.inferred_default_fontsize = _infer_default_fontsize(paragraphs)
        self.inferred_default_fontcolor = _infer_default_fontcolor(paragraphs)

        pass

    def start_paragraph(self, paragraph):
        # Update of indent_level and itemization_mode.
        prev_indent_level, prev_itemization_mode = self._indent_level, self._itemization_mode
        indent_level = paragraph.IndentLevel
        self._indent_level = indent_level
        is_bullet_visible = paragraph.ParagraphFormat.Bullet.Visible
        if is_bullet_visible:
            bullet_type = paragraph.ParagraphFormat.Bullet.Type
            if bullet_type == constants.ppBulletNumbered: 
                self._itemization_mode = "ordered"
            else:
                self._itemization_mode = "unordered"
        else:
            self._itemization_mode = None

        print("paragraph", paragraph.Text, "index_level", self._indent_level, "itemization", self._itemization_mode )

        ret_text = ""
        
        # Solve Itemization Separation.
        if (prev_indent_level, prev_itemization_mode) != (self._indent_level, self._itemization_mode):
            ret_text = "\n"

        # Indent Solver
        ret_text += " " * 4 * (self._indent_level - 1)

        # Itemization
        if self._itemization_mode == "ordered":
            ret_text += "1. "
            return ret_text
        elif self._itemization_mode == "unordered":
            ret_text += "* "
            return ret_text

        # Inferrence of Header
        """ Currently, Fontsize which is smaller than the normal cannot be handled.
        """ 
        if indent_level == 1:
            size = paragraph.Font.Size
            if 0 < size:
                ratio = size / self.inferred_default_fontsize
                if 2.0 <= ratio:
                    ret_text += "# "
                elif 1.5 <= ratio < 2.0:
                    ret_text += "## "
                elif 1.2 <= ratio < 1.5:
                    ret_text += "### "


        self._itemization_mode = None
        return ret_text

    def end_paragraph(self, paragraph):
        return "\n"


    def convert_run(self, run):
        # Link
        if run.ActionSettings(constants.ppMouseClick).Action == constants.ppActionHyperlink:
            link = run.ActionSettings(constants.ppMouseClick).Hyperlink.Address 
            text = f"[{run.Text}]({link})"
            return text

        text = run.Text
        text = text.replace("\r", "")
        text = text.replace("\013", "  \n")
        
        # String Formatter
        is_bold = (run.Font.Bold == constants.msoTrue)
        is_italic  = (run.Font.Italic == constants.msoTrue)
        is_underline  = (run.Font.Underline == constants.msoTrue)
        if is_bold and is_italic:
            text = f"***{text}***" 
        elif is_bold and (not is_italic):
            text = f"**{text}**" 
        elif (not is_bold) and is_italic:
            text = f"*{text}*" 

        if is_underline:
            text = f"<u>{text}</u>"

        # Handle Color Format.
        if self.inferred_default_fontcolor.as_int() != run.Font.Color.RGB:
            hex_code = Color(run.Font.Color.RGB).as_hex()
            text = f'<span style="color:{hex_code}">{text}</span>'
    
        return text


input_text = """
# Sample Document.
This is a sample.
Sample sentence is as follows.  

*  **hogehgoe** fafa
* [Sample](https://twitter.com/)
* ***e34***
    1. *hoge*
    2. <u>hoge2</u>   
    Happy Ever After
        * A
        * B
        * C
* EEE

### <span style="color:#00FB00">Type2</span>

* B
* C
"""

simple_text = """
<span style="color:#000001">ABDCDAFA</span>

* hoge 
* hoge2

<span style="color:#0000FF">text</span>
"""


if __name__ == "__main__":
    from fairypptx import markdown_functions
    shape = markdown_functions.write(handler, input_text.strip())
    #shape = markdown_functions.write(handler, simple_text.strip())
    #shapes = handler.get_selected_shapes()
    result = interpret(shape)


    with open("output.txt", "w", encoding="utf8") as fp:
        fp.write(result)
    print(result)

    #shape = maker.make_textbox(handler, "", fillcolor=(255, 255, 0), fontcolor=(0, 0, 0))
    #shape = markdown_functions.write(handler, result, shape=shape)

    #markdown_functions.write(handler, result.strip())
    pass
