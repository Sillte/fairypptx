from pathlib import Path
from fairypptx import Shape
from fairypptx import Table as PPTXTable 
from fairypptx import constants 

from fairypptx.parts.markdown import Markdown 
import json 
from fairypptx.parts._markdown.jsonast.context import Context
from fairypptx.parts._markdown.jsonast.editors import get_default_block_editors, get_default_inline_editors
from fairypptx.parts._markdown.jsonast.utils import to_jsonast
from fairypptx.parts._markdown.jsonast.models import PandocJsonAst
from fairypptx.paragraph_format import ParagraphFormat
from fairypptx.font import Font

from fairypptx.apis.font import FontApiModel


from typing import Mapping


class MarkdownFactory:
    def __init__(self, block_editors: Mapping | None = None, inline_editors: Mapping | None = None):
        self.block_editors = block_editors or get_default_block_editors()
        self.inline_editors = inline_editors or get_default_inline_editors()

    def from_document(self, document: str | Path) -> Markdown:
        json_ast_data = to_jsonast(document)
        pandoc_json_ast = PandocJsonAst.model_validate(json_ast_data)
        blocks = pandoc_json_ast.blocks
        #from pprint import pprint
        #pprint(blocks)
        shape = Shape.make(1) 
        context = Context(text_range=shape.text_range,
                          block_editors=self.block_editors,
                          inline_editors=self.inline_editors)
        for block in blocks:
            context.apply_block(block)
        return Markdown(shape)


if __name__ == "__main__":
    pass
    
    SCRIPT = """ 
### HEllo  

こんにちは、これから箇条書きを始めます

* AAA
    - fafa
* BBBB
    1. jkjdfafa
        - "faafaef"
    2. jkjdfafa
    3. jkjdfafa
* CCC
""".strip()
    markdown = MarkdownFactory().from_document(SCRIPT)
    markdown.shape.tighten()

