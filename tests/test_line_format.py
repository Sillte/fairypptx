import pytest
from fairypptx import Shape
from fairypptx import Color


def test_basic():
    """Test about substitution and `eq`.
    """
    in_shape = Shape.make(1)
    in_shape.line = 4 
    in_shape.line = Color(43 ,434, 1424)
    out_shape = Shape.make(2)
    assert in_shape.line != out_shape.line
    in_shape.line = out_shape.line
    assert in_shape.line == out_shape.line



if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])
