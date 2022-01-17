from collections.abc import Sequence 
from fairypptx import constants
from fairypptx.object_utils import ObjectDictMixin, getattr, setattr
from fairypptx.color import Color

class LineFormat(ObjectDictMixin):
    """LineFormat.

    Note
    ------------------------------------
    Insufficient Implementation (2020-04-19).
    Especially, for `arrow`s.

    """
    data = dict()
    data["Style"] = constants.msoLineSingle
    data["ForeColor.RGB"] = 0
    data["Visible"] = constants.msoTrue
    data["Transparency"] = 0

    common_keys = [
        "BackColor.RGB",
        "DashStyle",
        "ForeColor.RGB",
        "InsetPen",
        "Pattern",
        "Transparency",
        "Visible",
        "Weight",
        "Style",
    ]

    def to_dict(self, api_object):
        # Minimum specification
        if getattr(api_object, "Visible") != constants.msoTrue:
            return {"Visible": constants.msoFalse }

        keys = self.common_keys

        if getattr(api_object, "BeginArrowheadStyle") != constants.msoArrowheadNone:
            keys += ["BeginArrowheadStyle", "BeginArrowheadLength", "BeginArrowheadWidth"]
        if getattr(api_object, "EndArrowheadStyle") != constants.msoArrowheadNone:
            keys += ["EndArrowheadStyle", "EndArrowheadLength", "EndArrowheadWidth"]
        d = {key: getattr(api_object, key) for key in keys}

        # (2021/05/19)
        # For some keys, invalid values are initially set.
        # `setattr` of invalid values raises `ValueError`.
        # If these values are stored as `dict`, 
        # failure transpire in `apply` of `ObjectDictMixin`. 
        # Hence, the following procedure is performed.
        #
        # If it takes times, I'd  like to consider
        # to apply check with the limited sub-group of `keys`.
        remove_keys = set()
        for key, value in d.items():
            try:
                setattr(api_object, key, value)
            except ValueError:
                remove_keys.add(key)
        d = {key: value for key, value in d.items() if key not in remove_keys}

        # (2021/05/19) I feel the following procedure becomes unnecessary, 
        #  by `remove_keys`
        # Invalid (not supported) values are over-written.
        #if d["DashStyle"] == constants.msoLineDashStyleMixed:
        #    d["DashStyle"] = constants.msoLineSolid

        return d

    @property
    def color(self): 
        int_rgb = self.api.ForeColor.RGB
        color = Color(int_rgb)
        alpha = 1 - self.api.Transparency 
        return Color((*color.rgb, alpha))

    @color.setter
    def color(self, value): 
        color = Color(value)
        rgb, alpha = color.as_int(), color.alpha
        self.api.ForeColor.RGB = rgb
        self.api.Transparency = 1 - alpha 
        self.api.Visible = True
        return color

    @property
    def weight(self):
        if self.api.Visible:
            return self.api.Weight
        else:
            return None

    @weight.setter
    def weight(self, value):
        self.api.Visible = True
        self.api.Weight = value


class LineFormatProperty:
    def __get__(self, shape, objtype=None):
        try:
            return LineFormat(shape.api.Line)
        except AttributeError as e:
            """ Catch of AttributeError is mandatory.
            """
            raise NotImplementedError("Not-correctly implemented.") from e

    def __set__(self, shape, value):
        Line = shape.api.Line
        if value is None:
            Line.Visible = False
        elif isinstance(value, LineFormat):
            value.apply(Line)
        elif isinstance(value, int):
            if 1 <= value <= 50:
                # Line Weight.
                Line.Visible = True
                # Margin of discussion.
                Line.Style = constants.msoLineSingle
                Line.DashStyle = constants.msoLineSolid
                Line.Weight = value
            else:
                Line.Visible = True
                Line.ForeColor.RGB = value
        elif isinstance(value, Sequence):
            if len(value) == 2:
                weight, color = value
                self.__set__(shape, weight)
                self.__set__(shape, color)
            elif len(value) in {3, 4}:
                color = Color(value)
                self.__set__(shape, color)
            else:
                raise ValueError(f"Given Sequence cannot be handled at `{self.__class__.__name__}`, `{value}`")
        elif isinstance(value, Color):
            Line.ForeColor.RGB = value.as_int() 
            Line.Transparency = 1 - value.alpha
        else:
            raise ValueError(f"`{value}` cannot be set at `{self.__class__.__name__}`.")



