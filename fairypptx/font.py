from fairypptx import constants 
from typing import Any

from fairypptx import constants
from fairypptx.apis.font.api_model import FontApiModel
from fairypptx.apis.font.applicator import FontApplicator
from fairypptx.object_utils import to_api2


from fairypptx import constants
from fairypptx.color import Color

from typing import Self
from fairypptx.core.types import COMObject, PPTXObjectProtocol   

from fairypptx.object_utils import to_api2, is_object


class Font:
    """Represents the Font Information. 
    """
    def __init__(self, api: COMObject | PPTXObjectProtocol):
        if isinstance(api, Font):
            api = api.api
        self._api = api

    @property
    def api(self) -> COMObject:
        return self._api

    @property
    def bold(self) -> bool:
        return self.api.Bold != constants.msoFalse

    @bold.setter
    def bold(self, value: bool | int):
        if value is True:
            self.api.Bold = constants.msoTrue
        elif value is False:
            self.api.Bold = constants.msoFalse
        else:
            self.api.Bold = value

    @property
    def underline(self):
        return self.api.Underline != constants.msoFalse

    @underline.setter
    def underline(self, value):
        if value is True:
            self.api.Underline = constants.msoTrue
        elif value is False:
            self.api.Underline = constants.msoFalse
        else:
            self.api.Underline = value

    @property
    def color(self):
        return Color(self.api.Color.RGB)

    @color.setter
    def color(self, value):
        value = Color(value)
        self.api.Color.RGB = value.as_int()
        if value.alpha < 1:
            api2 = to_api2(self.api)
            api2.Fill.Transparency = 1 - value.alpha 

    @property
    def size(self):
        return self.api.Size

    @size.setter
    def size(self, value: int):
        self.api.Size = value


class FontProperty:
    def __get__(self, parent: PPTXObjectProtocol, objtype=None):
        return Font(parent.api.Font)

    def __set__(self, parent: PPTXObjectProtocol, value: Any) -> None:
        FontApplicator.apply(parent.api.Font, value)

