



"""Line Format API Schema.

LineFormatApiModel represents a mapping of Line COM properties to Pydantic.
Similar to Font, LineFormat uses a single schema (no type variants).

The API includes support for optional arrow properties, which are only read/written
if the line has arrowheads attached (BeginArrowheadStyle != msoArrowheadNone).

Responsibilities:
    - from_api(COMObject) → LineFormatApiModel: Read Line properties from COM
    - apply_api(COMObject) ← LineFormatApiModel: Write Line properties back to COM

Implementation notes:
    - Base keys (Style, Weight, DashStyle, etc.) are always read
    - Arrow keys are conditionally included if the line has non-None arrowheads
    - remove_invalidity() filters out properties that raise com_error during read
"""

from fairypptx import constants
from fairypptx.core.models import BaseApiModel
from fairypptx.core.utils import crude_api_read, crude_api_write, remove_invalidity
from fairypptx.core.types import COMObject
from fairypptx.object_utils import getattr


from collections.abc import Sequence
from typing import Any, ClassVar, Mapping, Self, Sequence


class LineFormatApiModel(BaseApiModel):
    api_data: Mapping[str, Any]

    _common_keys: ClassVar[Sequence[str]] = [
            "BackColor.RGB",
            "DashStyle",
            "ForeColor.RGB",
            "InsetPen",
            "Pattern",
            "Transparency",
            "Visible",
            "Weight",
            "Style"]

    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        data = dict()
        data["Style"] = constants.msoLineSingle
        data["ForeColor.RGB"] = 0
        data["Visible"] = constants.msoTrue
        data["Transparency"] = 0

        keys = list(cls._common_keys)

        if getattr(api, "BeginArrowheadStyle") != constants.msoArrowheadNone:
            keys += ["BeginArrowheadStyle", "BeginArrowheadLength", "BeginArrowheadWidth"]
        if getattr(api, "EndArrowheadStyle") != constants.msoArrowheadNone:
            keys += ["EndArrowheadStyle", "EndArrowheadLength", "EndArrowheadWidth"]
        data.update(crude_api_read(api, keys))

        data = remove_invalidity(api, data)
        return cls(api_data=data)

    def apply_api(self, api: COMObject) -> None:
        crude_api_write(api, self.api_data)
        return api