from fairypptx import constants
from fairypptx.core.models import BaseApiModel
from fairypptx.core.utils import crude_api_read, crude_api_write, remove_invalidity
from fairypptx.core.types import COMObject
from fairypptx.object_utils import getattr

from fairypptx.apis.paragraph_format.api_model import ParagraphFormatApiModel


from collections.abc import Sequence
from typing import Any, ClassVar, Mapping, Self


class TextRangeApiModel(BaseApiModel):
    text: str
    paragraph_format: ParagraphFormatApiModel

    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        return cls(text=str(api.Text),
                   paragraph_format=ParagraphFormatApiModel.from_api(api.ParagraphFormat))

    def apply_api(self, api: COMObject) -> None:
        api.Text = self.text
        self.paragraph_format.apply_api(api.ParagraphFormat)
        return api
