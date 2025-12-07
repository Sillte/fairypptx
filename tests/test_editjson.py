from fairypptx.editjson.utils import get_discriminator_mapping
import pytest

from pydantic import Field
from typing import Annotated
from enum import IntEnum
from pydantic import BaseModel
from typing import Literal


class FormatTypeEnum(IntEnum):
    Solid = 1
    Pattern = 2
    Gradient = 3


class Solid(BaseModel):
    type: Literal[FormatTypeEnum.Solid] = FormatTypeEnum.Solid

class Pattern(BaseModel):
    type: Literal[FormatTypeEnum.Pattern] = FormatTypeEnum.Pattern

class Gradient(BaseModel):
    type: Literal[FormatTypeEnum.Gradient] = FormatTypeEnum.Gradient

# ==== テストケース ===========================================================

def test_tagged_union_mapping():
    NaiveFormatAnnotated = Annotated[
        Solid | Pattern | Gradient,
        Field(discriminator="type")
    ]

    mapping = get_discriminator_mapping(NaiveFormatAnnotated, "type")

    assert mapping == {
        FormatTypeEnum.Solid: Solid,
        FormatTypeEnum.Pattern: Pattern,
        FormatTypeEnum.Gradient: Gradient,
    }

def test_plain_union_mapping():
    NaiveFormatPlain = Solid | Pattern | Gradient

    mapping = get_discriminator_mapping(NaiveFormatPlain, "type")

    assert mapping == {
        FormatTypeEnum.Solid: Solid,
        FormatTypeEnum.Pattern: Pattern,
        FormatTypeEnum.Gradient: Gradient,
    }
