"""Font API Schema.

FontApiModel represents a plain mapping of Font COM properties to Pydantic.
Unlike FillFormat (which uses tagged union for multiple subtypes), Font has
a single schema since Win32 Font API does not have type variants.

Responsibilities:
  - from_api(COMObject) → FontApiModel: Read Font properties from COM and build model
  - apply_api(COMObject) ← FontApiModel: Write Font properties back to COM
  
Implementation notes:
  - crude_api_read / crude_api_write handle common key bulk operations
  - Boolean-like properties (Bold, Italic, etc.) are only stored if they match
    MSO tri-state constants (msoTrue/msoFalse/msoCTrue) to avoid sentinel values
  - Some properties may raise com_error (e.g., if unsupported by shape type)
"""

from pywintypes import com_error
from fairypptx import constants
from fairypptx.core.models import BaseApiModel
from fairypptx.core.utils import crude_api_read, crude_api_write
from fairypptx.core.types import COMObject
from fairypptx.object_utils import getattr as f_getattr


from collections.abc import Mapping, Sequence
from typing import Any, ClassVar, Mapping, Self, Sequence



class FontApiModel(BaseApiModel):
    """Naive font edit parameter that owns dict<->Font conversion.

    Responsibilities:
    - `from_entity`: read a `Font` wrapper or COM `Font` and build a plain mapping.
    - `apply`: write the mapping back to a `Font` wrapper or COM `Font`.

    Implementation notes:
    - `crude_api_read` / `crude_api_write` are used for common keys.
    - For boolean-like properties we only keep keys whose values are one of
      the MSO tri-state constants (msoCTrue/msoTrue/msoFalse) to avoid
      storing other sentinel values.
    """

    api_data: Mapping[str, Any]
    _common_keys: ClassVar[Sequence[str]] = ["Size", "Name", "Color.RGB"]
    _only_determined_keys: ClassVar[Sequence[str]] = [
        "Bold",
        "Italic",
        "Shadow",
        "Superscript",
        "Subscript",
        "Underline",
    ]

    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        data: dict[str, Any] = {}
        # read common keys in bulk
        data.update(crude_api_read(api, cls._common_keys))

        # read boolean-like keys only when they look like MSO tri-state values
        for key in cls._only_determined_keys:
            try:
                value = f_getattr(api, key)
            except (com_error, AttributeError):
                continue
            if value in {constants.msoCTrue, constants.msoTrue, constants.msoFalse}:
                data[key] = value
        return cls(api_data=data)

    def apply_api(self, api:COMObject) -> Self:
        crude_api_write(api, self.api_data)
        return api
