from pydantic import BaseModel
from fairypptx import constants

from typing import Any, Mapping, Self, Sequence
from fairypptx.text_range import TextRange
from fairypptx.styles.font import NaiveFontEditParam
from fairypptx.styles.paragraph_format import NaiveParagraphFormatStyle
from dataclasses import dataclass 

@dataclass(frozen=True)
class ParagraphKey:
    paragraph_type: int  # Type of `Paragraph`.
    indent_level: int   # the level of indent.
    line_number: int  # the line number.


type FormatMapping = Mapping[ParagraphKey, tuple[NaiveFontEditParam, NaiveParagraphFormatStyle]]

class NaiveTextRangeParagraphStyle(BaseModel):
    """A naive TextRange serializer / applier.

    This class provides an example of an EditParam matching
    the `EditParamProtocol` (has `from_entity` and `apply`).
    
    This class only changes the styles of Textrange, not text contents itself.

    This style captures the `ParagraphFormat` for each position and **infer** the appropriate `ParagraphFormat` and `Font`. 
    
    """
    paragraph_keys: Sequence[ParagraphKey]
    paragraph_settings: Sequence[tuple[NaiveFontEditParam, NaiveParagraphFormatStyle]]

    @classmethod
    def from_format_mapping(cls, format_mapping: FormatMapping) -> tuple[Sequence[ParagraphKey], Sequence[tuple[NaiveFontEditParam, NaiveParagraphFormatStyle]]]:
        if not format_mapping:
            return [], []
        keys, settings = zip(*format_mapping.items(), strict=True) 
        return list(keys), list(settings)

    def to_format_mapping(self) -> FormatMapping:
        return {key: value for key, value in zip(self.paragraph_keys, self.paragraph_settings)}

    @classmethod
    def from_entity(cls, entity: TextRange) -> Self:
        """Create a JSON-serializable representation from a `TextRange`.

        This reads the raw `.Text`, and converts `Font` and
        `ParagraphFormat` to plain mappings.
        """
        tr = entity
        data = dict()
        for line_number, para in enumerate(tr.paragraphs):
            p_type = _to_paragraph_type(para)
            key = ParagraphKey(paragraph_type=p_type, indent_level=para.api.IndentLevel, line_number=line_number)
            font_param = NaiveFontEditParam.from_entity(para.font)
            paragraphformat_param = NaiveParagraphFormatStyle.from_entity(para.paragraph_format)
            data[key] = (font_param, paragraphformat_param)
        paragraph_keys, paragraph_settings = cls.from_format_mapping(data)
        return cls(paragraph_keys=paragraph_keys, paragraph_settings=paragraph_settings)

    def apply(self, entity: TextRange) -> TextRange:
        """Apply stored parameters onto `entity` and return it.

        - Sets `Text` content
        - Applies font mapping to `entity.api.Font`
        - Applies paragraphformat mapping to `entity.api.ParagraphFormat`
        """

        format_mapping = self.to_format_mapping()
        def _pick(paragraph_type: int, indent_level: int, line_number: int) -> tuple[NaiveFontEditParam, NaiveParagraphFormatStyle] | None:
            def _dist(key: ParagraphKey):
                return (key.paragraph_type != paragraph_type,
                        abs(key.indent_level - indent_level), 
                        abs(key.line_number - line_number))
            key = min(format_mapping.keys(),  key=_dist, default=None)
            return format_mapping[key] if key else  None
 
        tr = entity

        for line_number, para in enumerate(tr.paragraphs):
            p_type = _to_paragraph_type(para)
            indent_level = para.api.IndentLevel
            if picked :=_pick(p_type, indent_level, line_number):
                font_param, format_param = picked
                format_param.apply(para.paragraph_format)
                font_param.apply(para.font)
        return entity


def _to_paragraph_type(para: TextRange) -> int:
    f_api = para.paragraph_format.api
    assert f_api
    if f_api.Bullet.Visible == constants.msoTrue:
        return f_api.Bullet.Type
    else:
        return constants.ppBulletNone

if __name__ == "__main__":
    from fairypptx import Shape  
    shape = Shape()
    target = NaiveTextRangeParagraphStyle.from_entity(shape.textrange)
    target.apply(shape.textrange)
    data = target.model_dump()
    print(data)
    #import time 
    #for _ in range(20):
    #    print(_)
    #    time.sleep(2)
