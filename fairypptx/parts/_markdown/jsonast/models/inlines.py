"""
Reference
------
# https://hackage.haskell.org/package/pandoc-types-1.22.1/docs/Text-Pandoc-Definition.html
"""

from pydantic import BaseModel, Field
from typing import Annotated, Literal, Sequence, Any

from fairypptx.parts._markdown.jsonast.models.types import Attr


class DoubleQuoteType(BaseModel):
    t: Annotated[Literal["DoubleQuote"], Field(description="Tag")] = "DoubleQuote"

class SingleQuoteType(BaseModel):
    t: Annotated[Literal["SingleQuote"], Field(description="Tag")] = "SingleQuote"
QuoteType = DoubleQuoteType | SingleQuoteType

class BaseInlineModel(BaseModel):
    ...

class StrInline(BaseInlineModel):
    t: Annotated[Literal["Str"], Field(description="Tag")] = "Str"
    c: Annotated[str, Field(description="text")]

class FallbackInline(BaseInlineModel):
    """This is a fallback when  
    """
    t: Annotated[str, Field(description="Tag")]
    #c: Annotated[Mapping[str, Any] | Sequence[Any], Field(description="Any Data")]

class UnderlineInline(BaseInlineModel):
    t: Annotated[Literal["Underline"], Field(description="Tag")] = "Underline"
    c: Annotated[Sequence["InlineElement"], Field(description="Inlines")]

    @property
    def inlines(self) -> Sequence["InlineElement"]:
        return self.c

class StrongInline(BaseInlineModel):
    t: Annotated[Literal["Strong"], Field(description="Strong")] = "Strong"
    c: Annotated[Sequence["InlineElement"], Field(description="Inlines")]
    
    @property
    def inlines(self) -> Sequence["InlineElement"]:
        return self.c


class StrikeoutInline(BaseInlineModel):
    t: Annotated[Literal["Strikeout"], Field(description="Strikeout")] = "Strikeout"
    c: Annotated[Sequence["InlineElement"], Field(description="Inlines")]

    @property
    def inlines(self) -> Sequence["InlineElement"]:
        return self.c

class SuperscriptInline(BaseInlineModel):
    t: Annotated[Literal["SuperScript"], Field(description="SuperScript")] = "SuperScript"
    c: Annotated[Sequence["InlineElement"], Field(description="Inlines")]

    @property
    def inlines(self) -> Sequence["InlineElement"]:
        return self.c

# Ignore: SmallCaps

class QuotedInline(BaseInlineModel):
    t: Annotated[Literal["Quoted"], Field(description="Quoted")] = "Quoted" 
    c: Annotated[tuple[QuoteType, Sequence["InlineElement"]], Field(description="Quote type and Inlines")]


# Ignore Cite.

class CiteInline(BaseInlineModel):
    t: Annotated[Literal["Cite"], Field(description="Cite")] = "Cite"
    c: Annotated[tuple[Any, Sequence["InlineElement"]], Field(description="Cite")]

class CodeInline(BaseInlineModel):
    t: Annotated[Literal["Code"], Field(description="Code")] = "Code"
    c: Annotated[tuple[Any, str], Field(description="Attr and text")]

class SpaceInline(BaseInlineModel):
    t: Annotated[Literal["Space"], Field(description="Space")]  = "Space"

class SoftBreakInline(BaseInlineModel):
    t: Annotated[Literal["SoftBreak"], Field(description="SoftBreak")] = "SoftBreak"

class LineBreakInline(BaseInlineModel):
    t: Annotated[Literal["LineBreak"], Field(description="LineBreak")] = "LineBreak"

class MathInline(BaseInlineModel):
    t: Annotated[Literal["Math"], Field(description="Math")] = "Math"
    c: Annotated[tuple[Any, str], Field(description="MathType and text")]

class RawInline(BaseInlineModel):
    t: Annotated[Literal["RawInline"], Field(description="RawInline")] = "RawInline"
    c: Annotated[tuple[Any, str], Field(description="Format and text")]

class LinkInline(BaseInlineModel):
    t: Annotated[Literal["Link"], Field(description="Link")] = "Link"
    c: Annotated[tuple[Any, Sequence["InlineElement"], tuple[str, str]], Field(description="Attr, inlines, and target")]

    @property
    def inlines(self) -> Sequence["InlineElement"]:
        return self.c[1]

class ImageInline(BaseInlineModel):
    t: Annotated[Literal["Image"], Field(description="Image")] = "Image"
    c: Annotated[tuple[Any, Sequence["InlineElement"], Any], Field(description="Attr, inlines, and target")]


class SpanInline(BaseInlineModel):
    t: Annotated[Literal["Span"], Field(description="Span")] = "Span"
    c: Annotated[tuple[Any, Sequence["InlineElement"]], Field(description="Attr, inlines")]




ValidInlineElement = Annotated[ StrInline | UnderlineInline | StrongInline | StrikeoutInline |
                                SuperscriptInline | QuotedInline |  CiteInline | CodeInline |
                                SpaceInline | SoftBreakInline | LineBreakInline |  MathInline | 
                                RawInline | LinkInline |  ImageInline | SpaceInline
                                ,Field(description="Valid InlineElement", discriminator="t")]
InlineElement = Annotated[ValidInlineElement | FallbackInline, Field(description="InlineElement")]

