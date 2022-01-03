# Classes used over the entire package should be imported first. 
from fairypptx.color import Color  # NOQA
from fairypptx.constants import constants  # NOQA

print(constants)


# Policy of the order of imports: "Ancestors should exist without the decendants."
from fairypptx.application import Application  # NOQA
from fairypptx.presentation import Presentation  # NOQA
from fairypptx.slide import Slides, Slide  # NOQA
from fairypptx.shape import Shapes, Shape # NOQA
from fairypptx.text import TextFrame, TextRange, Text # NOQA
from fairypptx.table import Table, DFTable # NOQA

from fairypptx.parts.markdown import Markdown
from fairypptx.parts.latex import Latex  # NOQA


if __name__ == "__main__":
    shapes = Shapes()
    shape = shapes[0]
    shape.line = 10
    shape.line = 0
