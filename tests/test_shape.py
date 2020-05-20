import pytest
from PIL import Image
from fairypptx import Shape
from fairypptx import Color
from fairypptx import constants
from fairypptx import Application
from fairypptx import constants

def test_line_setter():
    shape = Shape.make(1)

    # Line Weight
    shape.line = 5
    assert shape.api.Line.Weight == 5
    shape.line = 1
    assert shape.api.Line.Weight == 1

    # Line Color
    shape.line = 0
    assert shape.api.Line.ForeColor.RGB == 0

    color_tuple = (10, 34, 43)
    color = Color(color_tuple)
    shape.line = color
    assert shape.api.Line.ForeColor.RGB == color.as_int()

    color_tuple = (10, 34, 13)
    shape.line = color_tuple
    assert shape.api.Line.ForeColor.RGB == Color(color_tuple).as_int()

    shape.line = None
    assert shape.api.Line.Visible == constants.msoFalse


def test_fill_setter():
    shape = Shape.make(4)
    shape.fill = None
    assert shape.api.Fill.Visible == constants.msoFalse

    shape.fill = 0
    assert shape.api.Fill.Visible == constants.msoTrue

    color = (255, 243, 132, 72)
    shape.fill = color
    assert shape.api.Fill.Visible == constants.msoTrue
    assert shape.api.Fill.ForeColor.RGB == Color(color).as_int()
    transparency = 1 - Color(color).alpha
    assert abs(shape.api.Fill.Transparency - transparency) < 1.e-4

def test_text():
    shape = Shape.make(1)
    target =  "Happy?"
    shape.text = target
    assert shape.api.TextFrame.TextRange.Text == target
    assert target == shape.text


def test_image():
    image = Image.new("RGBA", size=(255, 255), color=(255, 0, 0))
    shape = Shape.make(image)
    image = shape.to_image()
    assert isinstance(image, Image.Image)

def test_select():
    shape = Shape.make(1)
    shape.select(False)
    App = Application()
    assert App.ActiveWindow.Selection.Type == constants.ppSelectionShapes

def test_tighten():
    shape = Shape.make(1)
    shape.text = "This is a test of tighten."
    width, height = shape.Width, shape.Height
    shape.tighten()
    assert width != shape.Width
    assert height != shape.Height


if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])

