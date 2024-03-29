import pytest
from fairypptx import constants
from fairypptx import Shape, Shapes
from fairypptx import Slides
from fairypptx import Color

def test_getitem():
    slide = Slides().add(layout=constants.ppLayoutBlank)
    shapes = slide.shapes
    assert len(shapes) == 0, "Assumption about the test"
    count = 10
    for index in range(count):
        shape = shapes.add(1)
        shape.Left = index * 50
        shape.Top = 50
        shape.text = f"index{str(index)}"
        shape.text.font["Size"] = 12
        shape.Width = 50
        shape.Height = 50
    assert len(shapes) == 0, "Change of shapes is not transmitted to ``faritypptx.Shapes``."
    assert len(slide.shapes) == count, "Added shapes."

    color = Color((255, 0, 0))
    shape = slide.shapes[count // 2]
    assert isinstance(shape, Shape)
    slide.shapes[count // 2].fill = color
    assert shape.api.Fill.ForeColor.RGB == color.as_int()

    # Check of Slice.
    shapes = slide.shapes[::2]
    assert len(shapes) == count // 2
    for shape in shapes:
        shape.line = 5

def test_leafs():
    slide = Slides().add(layout=constants.ppLayoutBlank)
    s1 = Shape.make("S1")
    s2 = Shape.make("S2")
    sg = Shapes([s1, s2])
    grouped = sg.group()
    shapes = Shapes([grouped]).leafs
    assert len(shapes) == 2
    texts = {str(shape.text) for shape in shapes}
    assert texts == {"S1", "S2"}

 
if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])

