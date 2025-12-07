from pydantic import BaseModel
from typing import Mapping, Any, Self, Sequence, ClassVar

from pywintypes import com_error
from fairypptx import constants

from fairypptx._text import Font
from fairypptx.object_utils import getattr as f_getattr, setattr as f_setattr
from fairypptx.editjson.utils import crude_api_write, crude_api_read


class NaiveFontEditParam(BaseModel):
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
    def from_entity(cls, entity: Font) -> Self:
        """Build a plain mapping from a `Font` wrapper or raw COM `Font` object.

        - Accepts either the `Font` wrapper (with `.api`) or the raw COM object.
        - Reads the common numeric/string keys via `crude_api_read`.
        - Reads boolean-like keys individually and keeps them only if they are
          one of the MSO tri-state constants to avoid storing mixed/missing values.
        """
        api = entity.api

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

    def apply(self, entity: Font) -> Font:
        """Apply stored `api_data` onto the given `Font` wrapper or COM object.

        - Skips keys with value `None`.
        - Uses `crude_api_write` for bulk write; falls back to per-key writes when necessary.
        """
        api = entity.api
        crude_api_write(api, self.api_data)
        return entity
