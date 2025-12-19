import re
from pydantic import BaseModel
from fairypptx.core.models import BaseApiModel
from fairypptx.core.types import COMObject

from fairypptx.apis.paragraph_format.api_model import ParagraphFormatApiModel
from fairypptx.apis.font.api_model import FontApiModel


from collections.abc import Sequence
from typing import Self, Sequence

_normalize_softbreaks_pattern1 = re.compile(r'(?:\n+\r+\n*)')
_normalize_softbreaks_pattern2 = re.compile(r'(?:\n*\r+\n+)')

def normalize_paragraph_breaks(text: str) -> str:
    """
    Normalize paragraph breaks for PowerPoint TextRange.

    - `\r` represents an explicit paragraph break (Enter).
    - Consecutive `\r` (empty paragraphs) are preserved.
    - Any `\n` adjacent to `\r` is considered meaningless and removed.
    - Soft line breaks (`\n`) inside a paragraph are preserved.
    """
    text = _normalize_softbreaks_pattern1.sub("\r", text)
    text = _normalize_softbreaks_pattern2.sub("\r", text)
    return text


class TextRangeRunModel(BaseModel):
    text: str
    font: FontApiModel

class TextRangeParagraphModel(BaseModel):
    runs: Sequence[TextRangeRunModel]
    paragraph_format: ParagraphFormatApiModel

    @property
    def text(self) -> str:
        return sum([run.text for run in self.runs] , "")


class TextRangeApiModel(BaseApiModel):
    paragraphs: Sequence[TextRangeParagraphModel]

    @classmethod
    def _normalize_paragraph_text(cls, text: str) -> str:
        # PowerPoint COM quirks:
        # - '\n\r' or '\r\n' at paragraph boundary
        text = text.replace("\n\r", "\r")
        text = text.replace("\r\n", "\r")
        return text

    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        paragraphs = []
        for paragraph_api in api.Paragraphs():
            paragraph_format = ParagraphFormatApiModel.from_api(paragraph_api.ParagraphFormat)
            runs = []
            for run_api in paragraph_api.Runs():
                text = normalize_paragraph_breaks(paragraph_api.Text)
                if text.endswith("\r"):
                    text = text[:-1]
                font = FontApiModel.from_api(run_api.Font)
                runs.append(TextRangeRunModel(text=text, font=font))
            paragraphs.append(TextRangeParagraphModel(runs=runs, paragraph_format=paragraph_format))
        return cls(paragraphs=paragraphs)

    def apply_api(self, api: COMObject) -> None:
        api.Text = ""
        for i, paragraph in enumerate(self.paragraphs):
            paragraph_text = "".join([run.text for run in paragraph.runs])
            if i == 0:
                if paragraph_text:
                    inserted_api = api.InsertAfter(paragraph_text)
                else:
                    inserted_api = api.InsertAfter("\r")
            else:
                if paragraph_text:
                    inserted_api = api.InsertAfter(f"\r{paragraph_text}")
                else:
                    inserted_api = api.InsertAfter("\r")

            paragraph.paragraph_format.apply_api(inserted_api.ParagraphFormat)

            # 4. Run ごとに Font を適用
            current_insertion_point = inserted_api.Start # 挿入したテキストの先頭位置
            for run in paragraph.runs:
                if not run.text:
                    continue
                # Runのテキスト範囲を計算
                run_length = len(run.text)
                # TextRange(Start, Length) で run_api を取得
                run_api = api.Characters(current_insertion_point, run_length)
                run.font.apply_api(run_api.Font)
                current_insertion_point += run_length

    @property
    def runs(self) -> Sequence[TextRangeRunModel]:
        return sum((list(paragraph.runs) for paragraph in self.paragraphs), [])


