from fairypptx import constants
from fairypptx.color import Color
from fairypptx.core.types import COMObject, PPTXObjectProtocol
from fairypptx.object_utils import is_object 
from typing import Any
from fairypptx.apis.fill_format.bridge import FillFormatApiBridge



class FillApiApplicator:
    @classmethod
    def apply_bool(cls, api: COMObject, value: bool) -> None:
        api.Visible = constants.msoTrue if value else constants.msoFalse

    @classmethod
    def apply_any(cls, api: COMObject, value: Any) -> None:
        api.Visible = constants.msoTrue
        color = Color(value)
        api.ForeColor.RGB = color.as_int()
        api.Transparency = 1.0 - color.alpha
        api.Solid()

    @classmethod
    def apply_api(cls, api: COMObject, value: COMObject) -> None:
        api_bridge = FillFormatApiBridge.from_api(value)
        api_bridge.apply_api(api)

    @classmethod
    def apply(cls, api: COMObject, value: Any) -> None:
        if isinstance(value, PPTXObjectProtocol):
            cls.apply_api(api, value.api)
        elif is_object(value):
            cls.apply_api(api, value)
        elif isinstance(value, bool):
            cls.apply_bool(api, value)
        elif value is None:
            api.Visible = False
        else:
            cls.apply_any(api, value)
