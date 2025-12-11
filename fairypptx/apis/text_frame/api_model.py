from fairypptx import constants
from fairypptx.core.models import BaseApiModel
from fairypptx.core.utils import crude_api_read, crude_api_write, remove_invalidity
from fairypptx.core.types import COMObject
from fairypptx.object_utils import getattr


from collections.abc import Sequence
from typing import Any, ClassVar, Mapping, Self, Sequence


class TextFrameApiModel(BaseApiModel):
        # [TODO] To be implemented.
    api_data: Mapping[str, Any]

    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        data = dict()
        return cls(api_data=data)

    def apply_api(self, api: COMObject) -> None:
        crude_api_write(api, self.api_data)
        return api
