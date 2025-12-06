""" Color Class.

As for channel values, there are two way of ranges.  
* [0, 1] 
* [0, 255]
Fairies attempt to guess which format you try to specify, though handling of 1 may become wrong.
Internally, ``alpha`` is kept in [0, 1] and RGB are kept in [0, 255].

"""

from typing import Any, Self
import colorsys

from dataclasses import dataclass
from typing import Iterable, Tuple, Union
import colorsys

# Shortened here for readability — can include all 140 if needed
CSS_COLORS = {
    "black": "#000000",
    "white": "#FFFFFF",
    "red":   "#FF0000",
    "green": "#008000",
    "blue":  "#0000FF",
    "yellow": "#FFFF00",
    "cyan":   "#00FFFF",
    "magenta": "#FF00FF",
    "gray": "#808080",
    "silver": "#C0C0C0",
    "maroon": "#800000",
    "purple": "#800080",
    "olive": "#808000",
    "navy": "#000080",
    "teal": "#008080",
    "lime": "#00FF00",
    "orange": "#FFA500",
    "pink": "#FFC0CB",
    "brown": "#A52A2A",
    "rebeccapurple": "#663399",
}

@dataclass
class _Color:
    """Pythonic representation of `Color`.
    """

    r: int
    g: int
    b: int
    a: float = 1.0  # [0,1]

    @classmethod
    def from_any(cls, arg: Any) -> "_Color":
        """Smart constructor: accepts RGB tuples, hex, int, CSS name, objects."""
        if isinstance(arg, _Color):
            return arg
        # 1) CSS color name or hext coce.
        if isinstance(arg, str):
            if arg.startswith("#"):
                return cls._from_hex(arg)
            elif arg.lower() in CSS_COLORS:
                arg = CSS_COLORS[arg.lower()]
                return cls._from_hex(arg)
            else:
                raise ValueError(f"Unknown CSS color name or invalid hex code : {arg}")

        # 2) Integer like 0xRRGGBB
        if isinstance(arg, int):
            return cls._from_int(arg)

        # 3) Iterable RGB/RGBA
        if isinstance(arg, Iterable):
            tup = tuple(arg)
            if all(isinstance(elem, (int, float)) for elem in tup):
                if len(tup) == 3:
                    return cls._from_rgb_tuple(tup)  # type:[ignore]
                elif len(tup) == 4:
                    return cls._from_rgba_tuple(tup) # type:[ignore]
            raise ValueError(f"Invalid color tuple: {arg}")

        # 4) Object with `.rgba`
        if hasattr(arg, "rgba"):
            r, g, b, a = arg.rgba
            return cls(r, g, b, a / 255)

        raise ValueError(f"Cannot parse color: {arg}")

    @classmethod
    def _from_hex(cls, s: str) -> Self:
        s = s.lstrip("#")
        if len(s) not in (6, 8):
            raise ValueError(f"Hex must be 6 or 8 chars: {s}")

        r = int(s[0:2], 16)
        g = int(s[2:4], 16)
        b = int(s[4:6], 16)
        a = int(s[6:8], 16) / 255 if len(s) == 8 else 1.0
        return cls(r, g, b, a)

    @classmethod
    def _from_int(cls, value: int) -> Self:
        r = (value >> 16) & 0xFF
        g = (value >> 8) & 0xFF
        b = value & 0xFF
        return cls(r, g, b)

    @classmethod
    def _from_rgb_tuple(cls, arg: Tuple[int | float, int | float, int | float]) -> Self:
        if all(0 <= v <= 1 for v in arg):
            t = tuple(int(v * 255) for v in arg)
        else:
            t = tuple(int(v) for v in arg)
        return cls(*t)

    @classmethod
    def _from_rgba_tuple(cls, t: Tuple[int, int, int, Union[int, float]]) -> "_Color":
        rgb = t[:3]
        alpha = t[3]
        if all(0 <= v <= 1 for v in rgb):
            rgb = tuple(int(v * 255) for v in rgb)
        if alpha > 1:
            alpha = alpha / 255
        return cls(rgb[0], rgb[1], rgb[2], alpha)


# ================================================================
#                      Color class (complete)
# ================================================================

class Color:
    """RGBA color class.
    Internally:
        R,G,B ∈ [0,255]
        A ∈ [0,1]  
    """
    
    def __init__(self, *args, **kwargs):
        if len(args) == 1 and (not kwargs):
            if isinstance(args[0], Color):
                impl = args[0].impl # [NOTE] maybe copy...?
            elif  isinstance(args[0], _Color):
                impl = args[0]
            else:
                impl = _Color.from_any(args[0])
        else:
            # Delete to the internal implementation.
            impl = _Color(*args, **kwargs)
        self._impl: _Color = impl

    @property
    def impl(self) -> _Color:
        return self._impl

    @property
    def r(self) -> int:
        return self._impl.r

    @property
    def g(self) -> int:
        return self._impl.g

    @property
    def b(self) -> int:
        return self._impl.b

    @property
    def a(self) -> float:
        return self._impl.a

    def __eq__(self, other: Self):
        return self.impl == other.impl


    # ============================================================
    #                     Factory methods
    # ============================================================

    @classmethod
    def from_any(cls, arg: Any) -> Self:
        """Smart constructor: accepts RGB tuples, hex, int, CSS name, objects."""
        # 1) CSS color name or hext coce.
        return cls(_Color.from_any(arg))


    # ============================================================
    #                       Representations
    # ============================================================
    @property
    def rgb(self) -> tuple[int, int, int]:
        return (self.r, self.g, self.b)

    @property
    def rgba(self) -> tuple[int, int, int, int]:
        return (self.r, self.g, self.b, int(round(self.a * 255)))

    @property
    def alpha(self) -> float:
        return self.a

    def as_hex(self, with_alpha: bool = False) -> str:
        h = "#{:02X}{:02X}{:02X}".format(self.r, self.g, self.b)
        if with_alpha:
            h += "{:02X}".format(int(round(self.a * 255)))
        return h

    def as_int(self) -> int:
        return (self.r << 16) | (self.g << 8) | self.b

    # ============================================================
    #                     Color transforms
    # ============================================================
    def with_alpha(self, a: float) -> "Color":
        return Color(self.r, self.g, self.b, a)

    def with_rgb(self, r: int | None = None, g: int | None = None, b: int | None =None) -> "Color":
        return Color(r or self.r, g or self.g, b or self.b, self.a)

    def lighten(self, amount: float) -> "Color":
        h, l, s = colorsys.rgb_to_hls(self.r / 255, self.g / 255, self.b / 255)
        l = min(1, l + amount)
        r, g, b = colorsys.hls_to_rgb(h, l, s)
        return Color(int(r * 255), int(g * 255), int(b * 255), self.a)

    def darken(self, amount: float) -> "Color":
        return self.lighten(-amount)

    def hue_rotate(self, deg: float) -> "Color":
        h, l, s = colorsys.rgb_to_hls(self.r / 255, self.g / 255, self.b / 255)
        h = (h + deg / 360) % 1.0
        r, g, b = colorsys.hls_to_rgb(h, l, s)
        return Color(int(r * 255), int(g * 255), int(b * 255), self.a)

    # ============================================================
    #                      Utility
    # ============================================================
    def __str__(self):
        return f"Color({self.as_hex(with_alpha=(self.a < 1))})"


def make_hue_circle(seed_color, n_color=5):
    """ Make ``Hue Circle`` with ``seed``.
    Args:
        seed_color: Color-like object.
        n_color: the number of color.

    Returns:
        ``list`` of colors.
    """
    seed_color = Color(seed_color)
    alpha = seed_color.alpha
    r, g, b = map(lambda x: x / 255, seed_color.rgb)
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    hs = [(h + d / n_color) % 1.0 for d in range(n_color)]
    colors = [colorsys.hls_to_rgb(h, l, s) for h in hs]
    colors = [Color((*list(map(lambda elem: round(elem * 255), color)), alpha)) for color in colors]
    return colors


if __name__ == "__main__":
    make_hue_circle((43,43 ,43))
    fc = Color(4343)
    fc = Color("red")
    fc = Color("#FF0000")
    print(fc)
