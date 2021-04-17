import pytest
from PIL import Image
from fairypptx import constants
from fairypptx import Slide, Shape, Shapes, Slides


def test_to_image():
    slide = Slide()
    image = slide.to_image()
    assert isinstance(image, Image.Image)


def test_leaf_shapes():
    slide = Slides().add(layout=constants.ppLayoutBlank)
    slide.select()
    s1 = slide.shapes.add(1)
    s1.text = "S1"
    s2 = slide.shapes.add(1)
    s2.text = "S2"
    sg = Shapes([s1, s2])
    grouped = sg.group()
    shapes = slide.leaf_shapes
    assert {shape.text for shape in shapes} == {"S1", "S2"}


if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])

