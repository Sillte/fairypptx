""" Utility functions as for `PIL.image`.

"""
import numpy as np
from fairypptx.color import Color
from PIL import Image

def pad(image, color=0, width=1):
    color = Color(color)
    mode = image.mode
    array = np.array(image.convert("RGBA"))
    array = np.insert(array, (0, -1), color.rgba, axis=0)
    array = np.insert(array, (0, -1), color.rgba, axis=1)
    image = Image.fromarray(array).convert(mode)
    return image

def concatenate(images, axis=0):
    arrays = [np.array(image) for image in images]
    array = np.concatenate(arrays, axis=axis)
    return Image.fromarray(array)


if __name__ == "__main__":
    image = Image.open("1.png")
    image = pad(image)
