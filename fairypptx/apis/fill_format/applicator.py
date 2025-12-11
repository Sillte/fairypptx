"""Fill Format API Applicator: Strategy for applying values to COMObject.

This module applies FillFormatApiModel instances (Pydantic models) to Win32 COM Fill objects.

Responsibility:
  - Convert Pydantic models → COMObject mutations
  - Handle domain-level convenience types (bool, Color, Sequence) as a transition layer
    
Note on design layers:
  - API layer (here): Pydantic ↔ COMObject conversion
  - Domain layer: Should ideally convert (bool, Color, etc.) → COMObject/PPTXObjectProtocol
    before calling apply(). Currently, this applicator also handles these types for convenience.
  - COM layer: Direct COMObject access (Win32 constants like msoTrue, msoFalse)
"""

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
 
