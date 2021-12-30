import pytest
from PIL import Image
from fairypptx import Shape
from fairypptx import Color
from fairypptx import constants
from fairypptx import Application
from fairypptx import constants

def test_gradation():
    """Test about Gradation copy. 
    """
    in_shape = Shape.make(1)
    in_shape.fill.api.OneColorGradient(1, 1, 1)
    # in_shape.fill.api.TwoColorGradient(1, 1)

    out_shape = Shape.make(2)
    assert in_shape.fill != out_shape.fill
    out_shape.fill = in_shape.fill
    assert in_shape.fill == out_shape.fill

def test_solid():
    """Test about Gradation copy. 
    """
    in_shape = Shape.make(1)
    in_shape.fill = Color((255, 0, 0))
    # in_shape.fill.api.TwoColorGradient(1, 1)

    out_shape = Shape.make(2)
    out_shape.fill = 43432
    assert in_shape.fill != out_shape.fill
    out_shape.fill = in_shape.fill
    assert in_shape.fill == out_shape.fill


if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])
