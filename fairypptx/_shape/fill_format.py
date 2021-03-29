from collections import defaultdict
from fairypptx import constants
from fairypptx.color import Color
from fairypptx.object_utils import ObjectDictMixin, getattr


class FillFormat(ObjectDictMixin):
    """Fill Format.

    (2020-04-19) Currently, it is far from perfect.
    Only ``Patterned`` / ``Solid`` is handled.
    """

    data = dict()
    data["Type"] = constants.msoFillSolid
    data["ForeColor.RGB"] = 0
    data["Visible"] = constants.msoFalse
    data["Transparency"] = 0

    readonly = ["Type", "Pattern"]  # readonly parameters.

    common_keys = ["Type", "Visible"]
    type_to_keys = defaultdict(list)
    type_to_keys[constants.msoFillSolid] = ["ForeColor.RGB", "Visible", "Transparency"]
    type_to_keys[constants.msoFillPatterned] = [
        "Pattern",
        "ForeColor.RGB",
        "BackColor.RGB",
    ]

    def __init__(self, arg=None, **kwargs):
        super().__init__(arg, **kwargs)
        assert "Type" in self.data

    def to_dict(self, api_object):
        type_value = getattr(api_object, "Type")
        keys = self.common_keys + self.type_to_keys[type_value]
        return {key: getattr(api_object, key) for key in keys}

    def apply(self, api_object):
        type_value = self.data["Type"]
        if type_value == constants.msoFillSolid:
            api_object.Solid()
        elif type_value == constants.msoFillPatterned:
            api_object.Patterned(self.data["Pattern"])
        else:
            raise ValueError(f"Currently `type_value`={type_value} cannot be handled.")

        super().apply(api_object)

    @property
    def color(self):
        rgb_value = self.get("ForeColor.RGB", None)
        # Currently, `ForeColor.RGB` is required.
        if rgb_value:
            return Color(rgb_value)
        else:
            return None

class FillFormatProperty:
    def __get__(self, shape, objtype=None):
        return FillFormat(shape.api.Fill)

    def __set__(self, shape, value):
        Fill = shape.api.Fill
        print(type(value))
        if value is None:
            Fill.Visible = constants.msoFalse
        elif isinstance(value, FillFormat):
            value.apply(Fill)
        else:
            Fill.Visible = constants.msoTrue
            color = Color(value)
            Fill.ForeColor.RGB = color.as_int()
            Fill.Transparency = 1.0 - color.alpha
            Fill.Solid()
