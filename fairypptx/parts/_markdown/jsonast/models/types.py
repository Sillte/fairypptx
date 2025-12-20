from pydantic import BaseModel 
from typing import Literal, Sequence

class Alignment(BaseModel):
    t: Literal["AlignLeft", "AlignRight", "AlignCenter", "AlignDefault"]

class ColWidthDefault(BaseModel):
    t: Literal["ColWidthDefault"]

Attr = tuple[str, list[str], list[tuple[str, str]]]

ColWidth = float | ColWidthDefault
ColSpec = tuple[Alignment, ColWidth]



