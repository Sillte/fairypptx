"""Descriptor of Shapes / Shape for this package.

Note
-----------------------

"""
from fairypptx._text import Text

class TextProperty:
    def __get__(self, shape, klass=None):
        if shape is None:
            raise AttributeError("Cannot accept.")
        return Text(shape.api.TextFrame.TextRange)

    def __set__(self, shape, value):
        from fairypptx.text_range import TextRange
        tr = TextRange(shape)
        if isinstance(value, str):
            tr.api.Text = value
        else:
            assert False


class TextsProperty:
    """
    Note
    --------------------------------------
    `texts` corresponds to `textrange.runs`.
    """
    def __get__(self, shape, klass=None):
        if shape is None:
            raise AttributeError()
        return [Text(elem) for elem in shape.textrange.api.Runs()]

    def __set__(self, shape, value):
        from fairypptx import TextRange
        tr = TextRange(shape.api.TextFrame.TextRange)
        if not value:
            tr.text = ""
            return
        if isinstance(value, Text):
            raise TypeError(f"Type of `{value}` is `Text`. Use `.text`, not `.texts`.")

        tr.text = ""  # Reset.
        for index, elem in enumerate(value):
            tr = tr.insert(elem)
