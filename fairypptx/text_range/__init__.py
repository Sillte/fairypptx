from typing import Sequence, TYPE_CHECKING, Literal, cast

from collections.abc import Sequence
from fairypptx.registry_utils import BaseModelRegistry

from fairypptx.object_utils import upstream
from fairypptx import constants
from fairypptx.font import FontProperty
from fairypptx.paragraph_format import ParagraphFormatProperty

from fairypptx.core.resolvers import resolve_text_range
from fairypptx.core.types import COMObject, PPTXObjectProtocol
from fairypptx.apis.text_range.applicator import TextRangeApplicator

if TYPE_CHECKING:
    from fairypptx.shape import Shape

class TextRange:
    font = FontProperty()
    paragraph_format =  ParagraphFormatProperty()

    def __init__(self, arg: PPTXObjectProtocol | COMObject | None = None) -> None:
        self._api = resolve_text_range(arg)

    @property
    def api(self) -> COMObject:
        return self._api

    @property
    def api2(self) -> COMObject:
        start = self._api.Start
        length = self._api.Length
        shape_api = upstream(self._api, "Shape")
        return shape_api.TextFrame2.TextRange.GetCharacters(start, length)

    @property
    def shape(self) -> "Shape":
        from fairypptx.shape import Shape
        return Shape(upstream(self.api, "Shape"))

    @property
    def characters(self) -> Sequence["TextRange"]:
        return [TextRange(elem) for elem in self.api.Characters()]

    @property
    def words(self) -> Sequence["TextRange"]:
        return [TextRange(elem) for elem in self.api.Words()]

    @property
    def lines(self) -> Sequence["TextRange"]:
        return [TextRange(elem) for elem in self.api.Lines()]

    @property
    def sentences(self) -> Sequence["TextRange"]:
        return [TextRange(elem) for elem in self.api.Sentences()]

    @property
    def paragraphs(self) -> Sequence["TextRange"]:
        return [TextRange(elem) for elem in self.api.Paragraphs()]

    @property
    def runs(self) -> Sequence["TextRange"]:
        # (2022/02/08): Experimentally, I feel it is better that `runs` are separated at `paragraphs` 
        # Since the modification of `run` affects unintuitive. 
        # This phenomena was seen when revising `FontResizer`.
        return [TextRange(elem) for para in self.paragraphs for elem in para.api.Runs()]

    @property
    def root(self) -> "TextRange":
        """Return the entire `TextRange`.
        """
        textframe_api = upstream(self.api, "TextFrame")
        return TextRange(textframe_api.TextRange)

    @property
    def paragraph_index(self) -> int:
        """Return the index of `Paragraph`.
        where  the `Start` of `self` is included.
        """
        pivot = self.api.Start
        root = self.root
        for index, para in enumerate(root.paragraphs):
            start, length = para.api.Start, para.api.Length
            if start <= pivot < start + length:
                return index
        raise RuntimeError("Implementation Bug.")
    
    @property
    def editor(self) -> "TextRangeEditor":
        return TextRangeEditor(self)

    def insert(self, text: str, mode: Literal["after", "before"]="after") -> "TextRange":
        return self.editor.insert(text, mode)

    @property
    def text(self) -> str:
        return str(self.api.Text)

    @text.setter
    def text(self, text: str):
        self.api.Text = text

    def itemize(self) -> None:
        for elem in self.paragraphs:
            elem.api.ParagraphFormat.Bullet.Visible = constants.msoTrue
            elem.api.ParagraphFormat.Bullet.Type = constants.ppBulletUnnumbered

    def find(self, target : str) -> Sequence["TextRange"]:
        """Return `list` of `TextRange` whose text is `target`. 
        """
        result = []
        start = 0
        while True:
            cand = self.api.Find(target, start, constants.msoTrue)
            if cand is None:
                break
            start = cand.Start + cand.Length - 1 
            result.append(TextRange(cand))
        return result


    def register(self, style: str, style_type: str | None | type = None) -> None:
        from fairypptx.editjson.style_type_registry import TextRangeStyleTypeRegistry
        if not isinstance(style_type, type): 
            style_type = TextRangeStyleTypeRegistry.fetch(style_type)
        edit_param = style_type.from_entity(self)
        BaseModelRegistry.put(edit_param, "TextRange", style)


    def like(self, style: str) -> None:
        from fairypptx.editjson.protocols import EditParamProtocol
        edit_param = BaseModelRegistry.fetch("TextRange", style)
        edit_param = cast(EditParamProtocol, edit_param)
        edit_param.apply(self)

    @classmethod
    def make(cls, arg: str | Sequence[str]) -> "TextRange":
        return TextRangeFactory.make(arg)

    @classmethod
    def make_itemization(cls, texts: Sequence[str]) -> "TextRange":
        return TextRangeFactory.from_texts(texts)


class TextRangeProperty:
    def __get__(self, parent: PPTXObjectProtocol, objtype=None):
        return TextRange(parent.api.TextRange)

    def __set__(self, parent: PPTXObjectProtocol, value: str) -> None:
        TextRangeApplicator.apply(parent.api.TextRange, value)


# For structure hierarchy.
from fairypptx.text_range.editor import TextRangeEditor    
from fairypptx.text_range.factory import TextRangeFactory    

from fairypptx.text_range.formatters import DefaultFormatter   #NOQA
from fairypptx.text_range.formatters import FontResizer  #NOQA 
