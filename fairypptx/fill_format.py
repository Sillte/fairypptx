from fairypptx.apis.fill_format.bridge import FillFormatApiBridge
from fairypptx.apis.fill_format.applicator import FillApiApplicator
from fairypptx.color import Color, ColorLike
from fairypptx.core.types import COMObject

from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from fairypptx.shape import Shape

class FillFormat:
    """Fill Format.

    (2020-04-19) Currently, it is far from perfect.
    Only ``Patterned`` / ``Solid`` is handled.
    """
    def __init__(self, api: COMObject):
        self._api = api
    
    @property
    def api(self) -> COMObject: 
        return self._api

    @property
    def color(self) -> Color | None:
        rgb_value = self.api.ForeColor.RGB
        # Currently, `ForeColor.RGB` is required.
        if rgb_value:
            return Color(rgb_value)
        else:
            return None
        
    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, FillFormat):
            return NotImplemented
        return FillFormatApiBridge.from_api(self.api) == FillFormatApiBridge.from_api(other.api)


class FillFormatProperty:
    def __get__(self, shape: "Shape", objtype: type | None = None) -> FillFormat:
        return FillFormat(shape.api.Fill)

    def __set__(self, shape: "Shape", value: bool | FillFormat | ColorLike | None) -> None:
        FillApiApplicator.apply(shape.api.Fill, value)
