"""There are codes related to modify `TextRange`."""

from fairypptx.text_range import TextRange
from typing import Literal, Sequence
from fairypptx import constants


class TextRangeFactory:
    @classmethod
    def make(cls, arg: str | Sequence[str]) -> TextRange:
        if isinstance(arg, str):
            return cls.from_text(arg)
        elif isinstance(arg, Sequence):
            return cls.from_texts(arg)

        msg = f"TextRange cannot be created from `{arg}`."
        raise TypeError(msg)

    @classmethod
    def from_text(cls, text: str) -> TextRange:
        from fairypptx.shape import Shape

        shape = Shape.make(constants.msoShapeRectangle)
        textrange = shape.textrange
        textrange.text = text
        return textrange

    @classmethod
    def from_texts(
        cls, texts: Sequence[str], itemization: Literal["numbered", "unnumbered"] | None = "unnumbered"
    ) -> TextRange:
        from fairypptx import Shape

        assert all(isinstance(text, str) for text in texts), "Current Implementation"
        shape = Shape.make(constants.msoShapeRectangle)
        shape.api.TextFrame.TextRange.Text = "\r".join(texts)
        tr = TextRange(shape.api.TextFrame.TextRange)

        if itemization == "unnumbered":
            tr.api.ParagraphFormat.Bullet.Visible = True
            tr.api.ParagraphFormat.Bullet.Type = constants.ppBulletUnnumbered
        elif itemization == "numbered":
            tr.api.ParagraphFormat.Bullet.Visible = True
            tr.api.ParagraphFormat.Bullet.Type = constants.ppBulletNumbered
        else:
            tr.api.ParagraphFormat.Bullet.Visible = False
        # Itemization's normal display.
        tr.api.ParagraphFormat.Alignment = constants.ppAlignLeft
        return tr
