from fairypptx import constants
from fairypptx.color import Color
from fairypptx.core.types import COMObject
from fairypptx.core.models import ApiApplicator
from typing import Any
from fairypptx.apis.fill_format.api_model import FillFormatApiModel


def apply_custom(api: COMObject, value: Any) -> None:
    if isinstance(value, bool):
        api.Visible = constants.msoTrue if value else constants.msoFalse
    elif value is None:
        api.Visible = False
    else:
        api.Visible = constants.msoTrue
        color = Color(value)
        api.ForeColor.RGB = color.as_int()
        api.Transparency = 1.0 - color.alpha
        api.Solid()


FillApiApplicator = ApiApplicator(FillFormatApiModel, apply_custom)
 
