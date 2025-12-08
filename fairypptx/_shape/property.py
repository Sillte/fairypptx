"""Descriptor of Shapes / Shape for this package.

Note
-----------------------

"""
from typing import TYPE_CHECKING, Sequence

if TYPE_CHECKING:
    from fairypptx.shape import Shape


class TextProperty:
    def __get__(self, shape: "Shape", klass=None) -> str:
        if shape is None:
            raise AttributeError("Cannot accept.")
        return str(shape.api.TextFrame.TextRange.Text)

    def __set__(self, shape: "Shape", value: str):
        from fairypptx.text_range import TextRange
        TextRange(shape).api.Text = value

class TextsProperty:
    """
    Note
    --------------------------------------
    `texts` corresponds to `textrange.runs`.
    """
    def __get__(self, shape: "Shape", klass=None) -> Sequence[str]:
        return [elem.Text for elem in shape.textrange.api.Runs()]

    def __set__(self, shape: "Shape", value: Sequence[str]) -> None:
        from fairypptx.text_range import TextRange
        tr = TextRange(shape.api.TextFrame.TextRange)
        if not value:
            tr.text = ""
            return
        if isinstance(value, Text):
            raise TypeError(f"Type of `{value}` is `Text`. Use `.text`, not `.texts`.")

        tr.text = ""  # Reset.
        for index, elem in enumerate(value):
            tr = tr.insert(elem)
