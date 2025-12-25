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
                text = normalize_paragraph_breaks(run_api.Text)
                if text.endswith("\r"):
                    text = text[:-1]
                font = FontApiModel.from_api(run_api.Font)
                runs.append(TextRangeRunModel(text=text, font=font))
            paragraphs.append(TextRangeParagraphModel(runs=runs, paragraph_format=paragraph_format))
        return cls(paragraphs=paragraphs)


    def apply_api(self, api: COMObject) -> None:
        def _get_vba_len(text: str) -> int:
            """VBA(UTF-16)基準での文字数を取得する"""
            return len(text.encode('utf-16-le')) // 2

        api.Text = ""

        for i, paragraph in enumerate(self.paragraphs):
            # 1. 段落テキストの構築
            raw_text = "".join([run.text for run in paragraph.runs])
            
            # 2. 挿入する文字列の決定（改行コードの扱い）
            if i == 0:
                text_to_insert = raw_text if raw_text else "\r"
            else:
                text_to_insert = f"\r{raw_text}" if raw_text else "\r"

            inserted_api = api.InsertAfter(text_to_insert)
            para_api = api.Paragraphs(api.Paragraphs().Count)

            paragraph.paragraph_format.apply_api(para_api.ParagraphFormat)
            #paragraph.paragraph_format.apply_api(inserted_api.ParagraphFormat)

            # 3. 挿入後の「実際のVBA基準の開始位置」を特定
            # InsertAfterの戻り値(Rangeオブジェクト)のStartプロパティを使うのが最も確実です
            para_start_in_vba = inserted_api.Start 
            
            # 4. Run ごとに Font を適用
            # 文頭に \r を付けた場合はその分をスキップ
            if text_to_insert[0] == "\r":
                offset = 1  # `run.text` starts from `1` in this case.
            else:
                offset = 0

            run_cursor = para_start_in_vba + offset

            for run in paragraph.runs:
                if not run.text:
                    continue
                
                run_length_vba = _get_vba_len(run.text)
                run_api = api.Characters(run_cursor, run_length_vba)
                run.font.apply_api(run_api.Font)
                run_cursor += run_length_vba

    @property
    def runs(self) -> Sequence[TextRangeRunModel]:
        return sum((list(paragraph.runs) for paragraph in self.paragraphs), [])


