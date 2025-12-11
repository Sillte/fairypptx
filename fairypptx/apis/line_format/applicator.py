from fairypptx import constants
from fairypptx.apis.line_format.api_model import LineFormatApiModel
from fairypptx.color import Color, ColorLike
from fairypptx.core.types import COMObject
from fairypptx.core.models import ApiApplicator


from typing import Any

def apply_custom(api: COMObject, value: int | bool  | Color | ColorLike | None) -> None:
    if isinstance(value, int):
        if 1 <= value <= 50:
            api.Visible = True
            # Margin of discussion.
            api.Style = constants.msoLineSingle
            api.DashStyle = constants.msoLineSolid
            api.Weight = value
        else:
            api.Visible = True
            api.ForeColor.RGB = value
    elif value is None:
        api.Visible = False

    try:
        color = Color(value)
    except Exception as e:
        pass
    else:
        api.ForeColor.RGB = color.as_int()
        api.Transparency = 1 - color.alpha
        return
    raise ValueError(f"`{value}` cannot be set at `{api}`.")

LineFormatApplicator = ApiApplicator(LineFormatApiModel, apply_custom)