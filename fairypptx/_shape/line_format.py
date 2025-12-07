from collections.abc import Sequence 
from fairypptx import constants
from fairypptx.object_utils import ObjectDictMixin, getattr, setattr
from fairypptx.color import Color

from fairypptx import constants
from fairypptx.color import Color

from typing import Any, Mapping, Literal, ClassVar, Sequence, Self, Annotated, cast
from pprint import pprint
from fairypptx.enums import MsoFillType
from fairypptx.core.models import ApiBridgeBaseModel
from fairypptx.core.utils import crude_api_read, crude_api_write, remove_invalidity
from fairypptx.core.types import COMObject


class LineFormatApiBridge(ApiBridgeBaseModel):
    api_data: Mapping[str, Any]

    _common_keys: ClassVar[Sequence[str]] = [
            "BackColor.RGB",
            "DashStyle",
            "ForeColor.RGB",
            "InsetPen",
            "Pattern",
            "Transparency",
            "Visible",
            "Weight",
            "Style"]

    @classmethod
    def from_api(cls, api) -> Self:
        data = dict()
        data["Style"] = constants.msoLineSingle
        data["ForeColor.RGB"] = 0
        data["Visible"] = constants.msoTrue
        data["Transparency"] = 0

        keys = list(cls._common_keys)

        if getattr(api, "BeginArrowheadStyle") != constants.msoArrowheadNone:
            keys += ["BeginArrowheadStyle", "BeginArrowheadLength", "BeginArrowheadWidth"]
        if getattr(api, "EndArrowheadStyle") != constants.msoArrowheadNone:
            keys += ["EndArrowheadStyle", "EndArrowheadLength", "EndArrowheadWidth"]
        data.update(crude_api_read(api, keys))

        data = remove_invalidity(api, data)
        return cls(api_data=data)
    
    def apply_api(self, api):
        crude_api_write(api, self.api_data)
        return api


class LineFormat:
    def __init__(self, api):
        self._api = api
        
    @property
    def api(self) -> COMObject:
        return self._api 


    @property
    def color(self) -> Color: 
        int_rgb = self.api.ForeColor.RGB
        color = Color(int_rgb)
        alpha = 1 - self.api.Transparency 
        return Color((*color.rgb, alpha))

    @color.setter
    def color(self, value: Color): 
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
        
    def __eq__(self, other: object) -> bool:
        if not isinstance(other, LineFormat):
            return False
        api_bridge = LineFormatApiBridge.from_api(self.api)
        api_bridge1 = LineFormatApiBridge.from_api(other.api)
        return api_bridge.model_dump(exclude_defaults=True)  == api_bridge1.model_dump(exclude_defaults=True)


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
            api_bridge = LineFormatApiBridge.from_api(value.api)
            api_bridge.apply_api(Line)
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
