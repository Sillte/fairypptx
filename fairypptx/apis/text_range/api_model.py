from pydantic import BaseModel
from fairypptx.core.models import BaseApiModel
from fairypptx.core.types import COMObject

from fairypptx.apis.paragraph_format.api_model import ParagraphFormatApiModel
from fairypptx.apis.font.api_model import FontApiModel


from collections.abc import Sequence
from typing import Self, Sequence

class TextRangeRunModel(BaseModel):
    text: str
    paragraph_format: ParagraphFormatApiModel
    font: FontApiModel
    

class TextRangeApiModel(BaseApiModel):
    runs: Sequence[TextRangeRunModel]


    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        runs = []
        for run_api in api.Runs():
            text = run_api.Text
            paragraph_format = ParagraphFormatApiModel.from_api(run_api.ParagraphFormat)
            font = FontApiModel.from_api(run_api.Font)
            runs.append(TextRangeRunModel(text=text, paragraph_format=paragraph_format, font=font))
        return cls(runs=runs)

    def apply_api(self, api: COMObject) -> None:
        api.Text = ""
        for run in self.runs:
            run_api = api.InsertAfter(run.text)
            run.paragraph_format.apply_api(run_api.ParagraphFormat)
            run.font.apply_api(run_api.Font)
        return api
