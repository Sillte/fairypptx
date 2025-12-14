from fairypptx.core.models import BaseApiModel
from fairypptx.core.types import COMObject


from typing import Self
from fairypptx.apis.text_range import TextRangeApiModel


class TextFrameApiModel(BaseApiModel):
    text_range: TextRangeApiModel 

    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        tr = TextRangeApiModel.from_api(api.TextRange)
        return cls(text_range=tr)

    def apply_api(self, api: COMObject) -> None:
        self.text_range.apply_api(api.TextRange)
