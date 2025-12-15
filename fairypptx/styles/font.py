from pydantic import BaseModel
from typing import Mapping, Any, Self, Sequence, ClassVar

from pywintypes import com_error
from fairypptx import constants

from fairypptx.apis.font.api_model import FontApiModel
from fairypptx.font import Font



class NaiveFontEditParam(BaseModel):
    """Naive font edit parameter that owns dict<->Font conversion.

    """
    api_bridge: FontApiModel


    @classmethod
    def from_entity(cls, entity: Font) -> Self:
        """Build a plain mapping from a `Font` wrapper or raw COM `Font` object.

        - Accepts either the `Font` wrapper (with `.api`) or the raw COM object.
        - Reads the common numeric/string keys via `crude_api_read`.
        - Reads boolean-like keys individually and keeps them only if they are
          one of the MSO tri-state constants to avoid storing mixed/missing values.
        """
        api = entity.api
        api_bridge = FontApiModel.from_api(api)
        return cls(api_bridge=api_bridge)

    def apply(self, entity: Font) -> Font:
        """Apply stored `api_data` onto the given `Font` wrapper or COM object.

        - Skips keys with value `None`.
        - Uses `crude_api_write` for bulk write; falls back to per-key writes when necessary.
        """
        self.api_bridge.apply_api(entity.api)
        return entity
