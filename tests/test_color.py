import pytest
import numpy as np
import random

from fairypptx import Color

def test_format():
    color = Color("#FFFFFF")
    assert color.rgb == (255, 255, 255)
    assert color.alpha == 1
    assert color.as_int() == 16777215
    assert color.as_hex() == "#FFFFFF"

    values = [random.randint(0, 255) for _ in range(3)]
    color = Color(values)
    assert color.rgb == tuple(values)
    assert color.alpha == 1.0


if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])
