from collections.abc import Sequence
from typing import Literal, Annotated, cast, Any

from fairypptx import constants
from fairypptx.color import Color
from fairypptx.apis.line_format.bridge import LineFormatApiBridge
from fairypptx.core.types import COMObject, PPTXObjectProtocol
from fairypptx.enums import MsoFillType
from fairypptx.object_utils import is_object


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
        if type(self) is not type(other):
            return NotImplemented
        api_bridge = LineFormatApiBridge.from_api(self.api)
        api_bridge1 = LineFormatApiBridge.from_api(other.api)
        return api_bridge.model_dump(exclude_defaults=True)  == api_bridge1.model_dump(exclude_defaults=True)

class LineFormatApplicator:
    
    @classmethod
    def apply_api(cls, api: COMObject, value: COMObject) -> None:
        api_bridge = LineFormatApiBridge.from_api(value)
        api_bridge.apply_api(api)


    @classmethod
    def apply_any(cls, api: COMObject, value: Any) -> None:
        api.Visible = constants.msoTrue
        color = Color(value)
        api.ForeColor.RGB = color.as_int()
        api.Transparency = 1.0 - color.alpha


    @classmethod
    def apply(cls, api: COMObject, value: Any) -> None:
        if isinstance(value, PPTXObjectProtocol):
            cls.apply_api(api, value.api)
        elif is_object(value):
            cls.apply_api(api, value)
        elif isinstance(value, int):
            if 1 <= value <= 50:
                api.Visible = True
                # Margin of discussion.
                api.Style = constants.msoLineSingle
                api.DashStyle = constants.msoLineSolid
                api.Weight = value
            else:
                api.Visible = True
                api.ForeColor.RGB = value
        elif isinstance(value, (Sequence, Color)):
            color = Color(value)
            api.ForeColor.RGB = color.as_int()
            api.Transparency = 1 - color.alpha
        elif value is None:
            api.Visible = False
        else:
            raise ValueError(f"`{value}` cannot be set at `{cls.__name__}`.")


class LineFormatProperty:
    def __get__(self, shape, objtype=None):
        try:
            return LineFormat(shape.api.Line)
        except AttributeError as e:
            """ Catch of AttributeError is mandatory.
            """
            raise NotImplementedError("Not-correctly implemented.") from e

    def __set__(self, shape, value):
        LineFormatApplicator.apply(shape.api.Line, value)
