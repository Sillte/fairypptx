"""
Ref: 
# https://hackage.haskell.org/package/pandoc-types-1.22.1/docs/Text-Pandoc-Definition.html
"""

import json 
from fairypptx.parts._markdown.jsonast.utils import to_jsonast


from pydantic import BaseModel, Field
from typing import Annotated, Literal, Sequence, Mapping, Any

from fairypptx.parts._markdown.jsonast.models.inlines import *

from fairypptx.parts._markdown.jsonast.models.types import Attr
from fairypptx.parts._markdown.jsonast.models.types import Attr, Alignment, ColSpec


class BaseBlock(BaseModel):
    ...

class ParaBlock(BaseBlock):
    t: Annotated[Literal["Para"], Field(description="Tag")] = "Para"
    c: Annotated[Sequence[InlineElement], Field(description="inlines")]

    @property
    def inlines(self) -> Sequence[InlineElement]:
        return self.c

class PlainBlock(BaseBlock):
    t: Annotated[Literal["Plain"], Field(description="Tag")] = "Plain"
    c: Annotated[Sequence[InlineElement], Field(description="inlines")]

    @property
    def inlines(self) -> Sequence[InlineElement]:
        return self.c

class LineBlock(BaseBlock):
    t: Annotated[Literal["LineBlock"], Field(description="Tag")]
    c: Annotated[Sequence[Sequence[InlineElement]], Field(description="inlines")]

class CodeBlock(BaseBlock):
    t: Literal["CodeBlock"] = "CodeBlock"
    c: tuple[Attr, str]

class RawBlock(BaseBlock):
    t: Literal["RawBlock"] = "RawBlock"
    c: tuple[str, str]  # (Format, Text)

class BlockQuote(BaseBlock):
    t: Literal["BlockQuote"] = "BlockQuote"
    c: Sequence["BlockElement"] # 再帰

class OrderedList(BaseBlock):
    t: Literal["OrderedList"] = "OrderedList"
    c: tuple[tuple[int, Any, Any], Sequence[Sequence["BlockElement"]]]

    @property
    def blocks_list(self) -> Sequence[Sequence["BlockElement"]]:
        return self.c[-1]

class BulletList(BaseBlock):
    t: Literal["BulletList"] = "BulletList"
    c: Sequence[Sequence["BlockElement"]]

    @property
    def blocks_list(self) -> Sequence[Sequence["BlockElement"]]:
        return self.c

class DefinitionList(BaseBlock):
    t: Literal["DefinitionList"] = "DefinitionList"
    c: Sequence[tuple[Sequence[InlineElement], Sequence[Sequence["BlockElement"]]]]

class Header(BaseBlock):
    t: Literal["Header"] = "Header"
    c: tuple[int, Attr, Sequence[InlineElement]]

    @property
    def inlines(self) -> Sequence[InlineElement]:
        return self.c[-1]

class HorizontalRule(BaseBlock):
    t: Literal["HorizontalRule"] = "HorizontalRule"

class Div(BaseBlock):
    t: Literal["Div"] = "Div"
    c: tuple[Attr, Sequence["BlockElement"]]

class NullBlock(BaseBlock):
    t: Literal["Null"] = "Null"


class FallbackBlock(BaseBlock):
    """This is a fallback when `BlockElement` is not defined.
    """
    t: Annotated[str, Field(description="Tag")]



Cell = tuple[Attr, Alignment, int, int, Sequence["BlockElement"]]
Row = tuple[Attr, Sequence[Cell]]

type Caption = Annotated[tuple[Any, Sequence["BlockElement"]], Field("Caption")]
type TableHead = tuple[Attr, Sequence[Row]]
type TableBody = tuple[Attr, int, Sequence[Row], Sequence[Row]]
type TableFoot = tuple[Attr, Sequence[Row]]

Caption = tuple[Any, Sequence["BlockElement"]]
TableHead = tuple[Attr, Sequence[Row]]
TableBody = tuple[Attr, int, Sequence[Row], Sequence[Row]] # RowConf, HeadRows, IntermediateRows
TableFoot = tuple[Attr, Sequence[Row]]

class TableBlock(BaseBlock):
    t: Literal["Table"] = "Table"
    c: Annotated[
        tuple[
            Attr, 
            Caption, 
            Sequence[ColSpec], 
            TableHead, 
            Sequence[TableBody], 
            TableFoot
        ],
        Field(description="Detailed Pandoc Table structure")
    ]

    @property
    def head(self) -> TableHead:
        return self.c[3]

    @property
    def bodies(self) -> Sequence[TableBody]:
        return self.c[4]


ValidBlockElement = Annotated[
    ParaBlock | PlainBlock | LineBlock | CodeBlock | RawBlock |
    BlockQuote | OrderedList | BulletList | DefinitionList |
    Header | HorizontalRule | Div | TableBlock | NullBlock,
    Field(description="Valid BlockElement", discriminator="t")
]

BlockElement = Annotated[ValidBlockElement | FallbackBlock, Field(description="BlockElement")]


class PandocJsonAst(BaseModel):
    model_config = {
        "populate_by_name": True
    }
    pandoc_api_version: Annotated[Sequence[int], Field(alias="pandoc-api-version",
                                                description="Pandoc API version")] = [1, 23, 1]
    meta: Annotated[Mapping[str, Any],Field(description="Meta data")] = {}
    blocks: Sequence[BlockElement]

##PandocJsonAst.model_rebuild() # It seems that it is not necessary.

class InlineTester(BaseModel):
    inline: InlineElement



if __name__ == "__main__":
    pass
    sample = """
    {"blocks": [{"t": "Para", "c": [{"t": "Str", "c": "Three"}]}]}
    """
    model = PandocJsonAst(blocks=json.loads(sample)["blocks"])

    SCRIPT = """ 
| Name | Age |
|------|-----|
| Alice| 25  |
| Bob  | 30  |
    """.strip()
    jsonast = to_jsonast(SCRIPT)
    from pprint import pprint
    pprint(jsonast)
    #model = PandocJsonAst.model_validate(jsonast)
    #print("Check")
    #pprint(model.bodies)
