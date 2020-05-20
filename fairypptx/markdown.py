"""
Note
----
This module is assumed to freely import all of the Object Model.
On contrary, Object Model classes must import this class in Runtime.

(2020-04-26)
As you can easily notice, this code is yet a collection 
of Sillte-kun's practice codes,  

"""
from pathlib import Path
from fairypptx import Shape, TextRange, Application, Text
from fairypptx import constants

from fairypptx._markdown import write, interpret

class Markdown:
    """Handle Markdown.

    Wonder
    ----------------------------
    Derive UserString and make it behave like a string.
    """
    def __init__(self, arg, **kwargs):
        if isinstance(arg, (str, Path)):
            self.textrange = self.make(arg).textrange
        elif isinstance(arg, Shape):
            self.textrange = arg.textrange
        elif isinstance(arg, TextRange):
            self.textrange = arg
        elif isinstance(arg, Markdown):
            self.textrange = arg.textrange
        elif isinstance(arg, Text):
            raise NotImplementedError("Currently...")
        else:
            raise TypeError(f"Invalid Argument: `{arg}`", type(arg))
    def __str__(self):
        return interpret(self.textrange.shape)


    @classmethod
    def make(cls, arg, shape=None):
        if shape is None:
            shape = Shape.make("")
        selection = Application().api.ActiveWindow.Selection
        # Necessary to prevent deadlock.
        if selection.Type == constants.ppSelectionText:
            selection.UnSelect()

        arg = _to_content(arg)
        shape = Shape(write(shape, arg))
        shape.tighten()
        return Markdown(shape.textrange)


def _to_content(arg):
    try:
        path = Path(arg)
        if path.exists():
            return path.read_text(encoding="utf8")
    except OSError:
        pass
    return arg


if __name__ == "__main__":
    pass

