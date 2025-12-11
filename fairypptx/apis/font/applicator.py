"""Font API Applicator: Apply FontApiModel to COMObject.

This module applies FontApiModel instances (Pydantic models) to Win32 COM Font objects.

No custom conversion strategy is needed for Font because:
  - Font API has a single schema (no type variants like FillFormat)
  - Value types (Color, bool) are handled at the domain layer
  - Direct COMObject → COMObject mapping via ApiApplicator base class suffices

Design note:
  - apply_custom is None, so ApiApplicator.apply() uses only Pydantic model conversion
  - Custom domain-level types (bool, Color, etc.) should be converted to FontApiModel
    at the domain layer before calling this applicator
"""

from fairypptx.core.models import ApiApplicator
from fairypptx.apis.font.api_model import FontApiModel

FontApplicator = ApiApplicator(FontApiModel, None)

