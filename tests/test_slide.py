import pytest
from PIL import Image
from fairypptx import Slide


def test_to_image():
    slide = Slide()
    image = slide.to_image()
    assert isinstance(image, Image.Image)


if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])

