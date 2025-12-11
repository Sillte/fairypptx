from typing import Literal, Annotated, cast

from fairypptx.apis.line_format.applicator import LineFormatApplicator
from fairypptx.color import Color
from fairypptx.apis.line_format.api_model import LineFormatApiModel
from fairypptx.core.types import COMObject
from fairypptx.enums import MsoFillType


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
        api_bridge = LineFormatApiModel.from_api(self.api)
        api_bridge1 = LineFormatApiModel.from_api(other.api)
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
        LineFormatApplicator.apply(shape.api.Line, value)
