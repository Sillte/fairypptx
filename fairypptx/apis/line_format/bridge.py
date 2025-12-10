from fairypptx import constants
from fairypptx.core.models import ApiBridgeBaseModel
from fairypptx.core.utils import crude_api_read, crude_api_write, remove_invalidity
from fairypptx.object_utils import getattr


from collections.abc import Sequence
from typing import Any, ClassVar, Mapping, Self, Sequence


class LineFormatApiBridge(ApiBridgeBaseModel):
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
    def from_api(cls, api) -> Self:
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

    def apply_api(self, api):
        crude_api_write(api, self.api_data)
        return api