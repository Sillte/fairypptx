""" Color Class.

As for channel values, there are two way of ranges.  
* [0, 1] 
* [0, 255]
Fairies attempt to guess which format you try to specify, though handling of 1 may become wrong.
Internally, ``alpha`` is kept in [0, 1] and RGB are kept in [0, 255].

"""

import os
from collections.abc import Sequence 
import json
import colorsys
_this_folder = os.path.dirname(os.path.abspath(__file__))

# _config_folder = os.path.join(_this_folder, "config")
# _str_json_path = os.path.join(_config_folder, "fairy_color.json")

class Color:
    def __init__(self, arg):
        if isinstance(arg, str):
            if arg.startswith("#"):
                color_tuple = _hex_to_color(arg)
            else:
                raise NotImplementedError
        elif isinstance(arg, int):
            color_tuple = _int_to_color(arg)
        elif isinstance(arg, Sequence):
            color_tuple = arg
        else:
            try:
                color_tuple = arg.rgba
            except AttributeError:
                raise ValueError(f"Cannot decipher `{arg}` ", type(arg))


        if len(color_tuple) == 3:
            self._rgb, self._alpha = _normalize(color_tuple, 1.0)
        elif len(color_tuple) == 4:
            self._rgb, self._alpha = _normalize(color_tuple[:3], color_tuple[3])
        else:
            raise ValueError(f"Cannot handle, `{color_tuple}`")
        
    @property
    def rgb(self):
        """3-length tuple.
        Range of channel values is [0, 255].
        """
        return self._rgb

    @property
    def rgba(self):
        """4-length tuple.
        Range of channel values is [0, 255].
        """
        return self._rgb + (round(self._alpha * 255), )

    @property
    def alpha(self):
        """ Value of alpha channel.
        Range is [0, 1]
        """
        return self._alpha

    def as_int(self):
        """ Return RGB as int value.
        """
        return sum(self.rgb[index] << (8*index) for index in range(3))

    def as_hex(self, with_alpha=False):
        """ Return Hex code.
        """
        code = "#" + "".join(map(lambda v: format(v, "X"), self.rgb))
        if with_alpha:
            code += format(int(self.alpha * 255), "X")
        return code.upper()

    def __eq__(self, other):
        return self.rgba == Color(other).rgba

    def __hash__(self):
        return hash(self.rgba)


    def __str__(self):
        if self.alpha == 1:
            return f"Color({self.rgb})"
        else:
            return f"Color({self.rgba})"


def _normalize(color_tuple, alpha):
    """ Normalize the range of values.
    """
    is_floats = [0 <= elem <= 1 for elem in color_tuple]
    if all(is_floats):
        color_tuple = tuple(map(lambda elem: int(elem * 255), color_tuple))
    if 1 < alpha:
        alpha = alpha / 255
    return color_tuple, alpha

def _hex_to_color(color_str):
    color_str = color_str.strip("#")
    assert len(color_str) in {6, 8}, "Invalid Hex Color Code"
    strings = map(lambda pair: "".join(pair), zip(*[iter(color_str)] * 2))
    return tuple(map(lambda s: int(s, 16), strings))

def _int_to_color(color_int):
    rgb_tuple = tuple((color_int >> (index * 8)) & (2**8 -1) for index in range(3))
    return rgb_tuple


def make_hue_circle(seed_color, n_color=5):
    """ Make ``Hue Circle`` with ``seed``.
    Args:
        seed_color: 3-length tuple. (R, G, B) 
        n_color: the number of color.

    Returns:
        ``list`` of colors.
    """
    seed_color = Color(seed_color)
    r, g, b = map(lambda x: x / 255, seed_color.rgb)
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    hs = [(h + d / n_color) % 1.0 for d in range(n_color)]
    colors = [colorsys.hls_to_rgb(h, l, s) for h in hs]
    colors = [Color(color) for color in colors]
    return colors


if __name__ == "__main__":
    fc = Color(4343)
    number= fc.rgb_as_int()
    fc = Color("red")
    color_tuple= fc.rgb_as_tuple()
    print(color_tuple)
