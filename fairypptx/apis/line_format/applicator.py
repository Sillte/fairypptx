

"""Line Format API Applicator: Apply LineFormatApiModel to COMObject.

This module applies LineFormatApiModel instances (Pydantic models) to Win32 COM Line objects.

Responsibility:
    - Convert Pydantic models → COMObject mutations
    - Handle domain-level convenience types (int weight, Color) as a transition layer
    - Dispatch on value type: int (weight 1-50) vs Color (RGB) vs None (invisible)

Note on design layers:
    - API layer (here): Pydantic ↔ COMObject conversion + convenience type dispatch
    - Domain layer: LineFormat should ideally convert types before calling applicator
    - COM layer: Direct COMObject access (Win32 constants like msoLineSingle, msoTrue)
"""

from fairypptx import constants
from fairypptx.apis.line_format.api_model import LineFormatApiModel
from fairypptx.color import Color, ColorLike
from fairypptx.core.types import COMObject
from fairypptx.core.models import ApiApplicator


from typing import Any

def apply_custom(api: COMObject, value: int | bool  | Color | ColorLike | None):
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
    else:
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
